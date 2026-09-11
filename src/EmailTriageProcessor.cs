using System.Collections.Concurrent;
using System.Text.Json;
using Azure.Connectors.Sdk;
using Azure.Connectors.Sdk.Office365;
using Azure.Connectors.Sdk.Office365.Models;
using Azure.Connectors.Sdk.Office365Users;
using Azure.Connectors.Sdk.Teams;
using Azure.Connectors.Sdk.Teams.Models;
using Microsoft.Extensions.Logging;

namespace EmailTriage.AppService;

/// <summary>
/// Triage pipeline for inbound mail. Ported from the Azure Functions e2e sample's
/// <c>ProcessEmail</c> — the logic is identical; only the *host shell* changed.
///
/// In the Functions sample this ran inside an <c>[Function]</c> method bound with
/// <c>[ConnectorTrigger]</c>. On App Service there is no Functions runtime, so
/// <see cref="OnNewEmailEndpoint"/> receives the raw connector callback over plain HTTP,
/// deserializes it into <see cref="Office365OnNewEmailTriggerPayload"/> (a portable
/// <c>Azure.Connectors.Sdk</c> type), and hands it to <see cref="ProcessAsync"/>.
///
/// The Office 365 / Teams / Office 365 Users *clients* are the same SDK clients the
/// Functions sample used — they only need a connection runtime URL + a managed-identity
/// <c>TokenCredential</c>, so they move to App Service with zero changes.
/// </summary>
public sealed class EmailTriageProcessor
{
    private const string PostAsFlowBot = "Flow bot";
    private const string PostInChannel = "Channel";
    private const int SenderHistoryDays = 7;
    private const int SenderHistoryFetchTop = 25;
    private const int SenderProfileCacheTtlMinutes = 10;

    private sealed record SenderProfile(
        string? DisplayName,
        string? JobTitle,
        string? Department,
        string? ManagerDisplayName);

    private static readonly ConcurrentDictionary<string, (SenderProfile? profile, bool notFound, DateTime cachedAt)> SenderProfileCache = new(StringComparer.OrdinalIgnoreCase);

    private sealed record SenderHistory(int TotalRecent, int LastWeek, DateTime? MostRecent)
    {
        public static SenderHistory Empty { get; } = new(0, 0, null);
    }

    private readonly ILogger<EmailTriageProcessor> _logger;
    private readonly TeamsClient _teamsClient;
    private readonly Office365Client _office365Client;
    private readonly Office365UsersClient _office365UsersClient;
    private readonly ImportanceClassifier _classifier;
    private readonly string _teamsTeamId;
    private readonly string _teamsChannelId;
    private readonly IReadOnlyList<string> _internalDomains;

    public EmailTriageProcessor(
        ILogger<EmailTriageProcessor> logger,
        TeamsClient teamsClient,
        Office365Client office365Client,
        Office365UsersClient office365UsersClient,
        ImportanceClassifier classifier)
    {
        _logger = logger;
        _teamsClient = teamsClient;
        _office365Client = office365Client;
        _office365UsersClient = office365UsersClient;
        _classifier = classifier;
        _teamsTeamId = Environment.GetEnvironmentVariable("TEAMS_TEAM_ID") ?? "";
        _teamsChannelId = Environment.GetEnvironmentVariable("TEAMS_CHANNEL_ID") ?? "";
        _internalDomains = (Environment.GetEnvironmentVariable("INTERNAL_DOMAINS") ?? "")
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(d => d.ToLowerInvariant())
            .ToArray();
    }

    // True when no INTERNAL_DOMAINS allowlist is configured (look up every sender) or when
    // the sender's domain matches the allowlist. False means we should skip the API call
    // and treat the sender as external.
    private bool ShouldLookupSender(string senderEmail)
    {
        if (_internalDomains.Count == 0) return true;
        var atIdx = senderEmail.LastIndexOf('@');
        if (atIdx < 0 || atIdx == senderEmail.Length - 1) return false;
        var domain = senderEmail[(atIdx + 1)..].Trim().ToLowerInvariant();
        return _internalDomains.Any(d => domain == d || domain.EndsWith("." + d));
    }

    /// <summary>
    /// Entry point invoked by the App Service webhook once the connector callback body has been
    /// deserialized. Iterates the batch, classifies each email, and for the important ones
    /// enriches + posts to Teams + flags the source message.
    /// </summary>
    public async Task ProcessAsync(Office365OnNewEmailTriggerPayload? payload, CancellationToken cancellationToken = default)
    {
        var emails = payload?.Body?.Value;
        _logger.LogInformation(
            "Trigger callback received. emailCount={Count}",
            emails?.Count ?? -1);

        if (emails is null || emails.Count == 0)
        {
            _logger.LogWarning("Empty trigger payload — nothing to process.");
            return;
        }

        foreach (var email in emails)
        {
            if (email is null) continue;

            var verdict = _classifier.Classify(
                email.From,
                email.Subject,
                email.Body,
                email.BodyPreview,
                email.Importance);

            if (!verdict.IsImportant)
            {
                _logger.LogInformation(
                    "Skipping non-important email. Subject={Subject} From={From} Importance={Importance}",
                    email.Subject, email.From, email.Importance);
                continue;
            }

            _logger.LogInformation(
                "Important email accepted. Subject={Subject} From={From} Reasons={Reasons}",
                email.Subject, email.From, string.Join(" | ", verdict.Reasons));

            var history = await GetSenderHistoryAsync(email.From);
            await PostTriageCardAsync(email, history, verdict);
            await FlagSourceMessageAsync(email);
        }
    }

    /// <summary>
    /// Pulls the sender's recent history from the watched mailbox via
    /// <see cref="Office365Client.GetEmailsAsync"/>, fanning out across Inbox
    /// and Archive (the connector's GetEmails is per-folder). Best-effort: failures
    /// degrade gracefully to "no history".
    /// </summary>
    private async Task<SenderHistory> GetSenderHistoryAsync(string? senderEmail)
    {
        if (string.IsNullOrWhiteSpace(senderEmail))
        {
            return SenderHistory.Empty;
        }

        string[] folders = ["Inbox", "Archive"];
        var perFolderTasks = folders.Select(f => FetchFromFolderAsync(senderEmail, f));
        var perFolderResults = await Task.WhenAll(perFolderTasks);

        var messages = perFolderResults.SelectMany(r => r).ToList();
        if (messages.Count == 0)
        {
            return SenderHistory.Empty;
        }

        var cutoff = DateTime.UtcNow.AddDays(-SenderHistoryDays);
        var lastWeek = messages.Count(m => m.ReceivedTime is DateTime t && t >= cutoff);
        var mostRecent = messages
            .Select(m => m.ReceivedTime)
            .Where(t => t.HasValue)
            .DefaultIfEmpty()
            .Max();

        return new SenderHistory(messages.Count, lastWeek, mostRecent);
    }

    private async Task<IReadOnlyList<GraphClientReceiveMessage>> FetchFromFolderAsync(string senderEmail, string folder)
    {
        try
        {
            var response = await _office365Client.GetEmailsAsync(
                folder: folder,
                to: null,
                cC: null,
                toOrCC: null,
                from: senderEmail,
                importance: null,
                onlyWithAttachments: false,
                subjectFilter: null,
                fetchOnlyUnreadMessages: false,
                originalMailboxAddress: null,
                includeAttachments: false,
                searchQuery: null,
                top: SenderHistoryFetchTop,
                cancellationToken: default);

            return (IReadOnlyList<GraphClientReceiveMessage>?)response?.Value ?? [];
        }
        catch (ConnectorException ex)
        {
            _logger.LogWarning(ex,
                "Office365 GetEmails failed for sender {Sender} in folder {Folder}. ConnectorName={ConnectorName}, ErrorCode={ErrorCode}, ErrorMessage={ErrorMessage} — skipping that folder.",
                senderEmail, folder, ex.ConnectorName, ex.ErrorCode, ex.Message);
            return [];
        }
    }

    /// <summary>
    /// Sets the Outlook follow-up flag on the source message via the Office 365
    /// connector. Best-effort — flag failures don't fail the pipeline.
    /// </summary>
    private async Task FlagSourceMessageAsync(GraphClientReceiveMessage email)
    {
        if (string.IsNullOrEmpty(email.MessageId))
        {
            _logger.LogDebug("No MessageId on payload; skipping flag.");
            return;
        }

        try
        {
            await _office365Client.FlagAsync(
                messageId: email.MessageId,
                input: new UpdateEmailFlag { Flag = new { flagStatus = "flagged" } },
                originalMailboxAddress: null,
                cancellationToken: default);

            _logger.LogInformation("Flagged source email. MessageId={MessageId}", email.MessageId);
        }
        catch (ConnectorException ex)
        {
            _logger.LogWarning(ex,
                "Failed to flag source email. MessageId={MessageId}. ConnectorName={ConnectorName}, ErrorCode={ErrorCode}, ErrorMessage={ErrorMessage}",
                email.MessageId, ex.ConnectorName, ex.ErrorCode, ex.Message);
        }
    }

    /// <summary>
    /// Looks up the sender's M365 user profile via the Office 365 Users connector.
    /// A successful lookup means the sender is in the org; a 404 means external.
    /// Results are cached for 10 minutes to absorb bursty mail volume.
    /// </summary>
    private async Task<(SenderProfile? profile, bool notFound)> GetSenderProfileAsync(string? senderEmail)
    {
        if (string.IsNullOrWhiteSpace(senderEmail))
            return (null, false);

        var normalizedSender = senderEmail.Trim();
        var now = DateTime.UtcNow;

        if (SenderProfileCache.TryGetValue(normalizedSender, out var cached) &&
            now - cached.cachedAt < TimeSpan.FromMinutes(SenderProfileCacheTtlMinutes))
        {
            return (cached.profile, cached.notFound);
        }

        // Domain prefilter: if INTERNAL_DOMAINS is configured and the sender's domain
        // is not in it, skip the API call and treat the sender as external.
        if (!ShouldLookupSender(normalizedSender))
        {
            SenderProfileCache[normalizedSender] = (null, true, now);
            return (null, true);
        }

        try
        {
            var user = await _office365UsersClient.UserProfileAsync(normalizedSender);

            // Fetch manager display name — best-effort, users with no manager return null or throw.
            string? managerName = null;
            try
            {
                var manager = await _office365UsersClient.ManagerAsync(normalizedSender);
                managerName = manager?.DisplayName;
            }
            catch (ConnectorException)
            {
                // No manager record — tolerated.
            }

            var profile = new SenderProfile(user?.DisplayName, user?.JobTitle, user?.Department, managerName);
            SenderProfileCache[normalizedSender] = (profile, false, now);
            return (profile, false);
        }
        catch (ConnectorException ex)
        {
            _logger.LogWarning(ex,
                "Office 365 Users profile lookup failed for sender {Sender}. ConnectorName={ConnectorName}, ErrorCode={ErrorCode}, ErrorMessage={ErrorMessage}",
                normalizedSender, ex.ConnectorName, ex.ErrorCode, ex.Message);
            return (null, false);
        }
    }

    private async Task PostTriageCardAsync(GraphClientReceiveMessage email, SenderHistory history, ImportanceVerdict verdict)
    {
        if (string.IsNullOrEmpty(_teamsTeamId) || string.IsNullOrEmpty(_teamsChannelId))
        {
            _logger.LogWarning("TEAMS_TEAM_ID or TEAMS_CHANNEL_ID not configured. Skipping Teams notification.");
            return;
        }

        var (senderProfile, senderNotFound) = await GetSenderProfileAsync(email.From);

        var badge = senderNotFound
            ? "🔴 <b>EXTERNAL — verify identity before acting</b><br/>"
            : senderProfile is not null
                ? "🟢 <b>IN-ORG</b><br/>"
                : "";

        var profileLine = senderProfile is not null
            ? $"<br/><b>Title:</b> {senderProfile.JobTitle ?? "(not set)"}" +
              $" | <b>Dept:</b> {senderProfile.Department ?? "(not set)"}" +
              (senderProfile.ManagerDisplayName is not null
                  ? $" | <b>Manager:</b> {senderProfile.ManagerDisplayName}"
                  : "")
            : "";

        var historyLine = history.TotalRecent switch
        {
            0 => "<br/><b>Sender history:</b> no prior emails from this sender in Inbox or Archive",
            _ => $"<br/><b>Sender history:</b> {history.TotalRecent} emails from this sender across Inbox + Archive " +
                 $"({history.LastWeek} in last {SenderHistoryDays}d" +
                 (history.MostRecent is DateTime t ? $", most recent {t:yyyy-MM-dd HH:mm} UTC" : "") +
                 ")"
        };

        var reasonsLine = verdict.Reasons.Count > 0
            ? $"<br/><b>Why flagged:</b> {string.Join("; ", verdict.Reasons)}"
            : "";

        var messageBody =
            $"<b>📧 Email triage — review required (via App Service)</b><br/>" +
            $"{badge}" +
            $"<b>From:</b> {email.From}{profileLine}{historyLine}{reasonsLine}<br/>" +
            $"<b>Subject:</b> {email.Subject}<br/>" +
            $"<b>Preview:</b> {email.BodyPreview ?? "(no preview)"}<br/>" +
            $"<i>(source email has been flagged in Outlook)</i>";

        // DynamicPostMessageRequest is a *dynamic-schema* body whose properties are
        // resolved at runtime by the connector's schema discovery endpoint. The connector
        // backend expects camelCase keys, so populate via AdditionalProperties
        // ([JsonExtensionData] on the base class) using literal camelCase keys.
        var request = new DynamicPostMessageRequest();
        request.AdditionalProperties["recipient"] = JsonSerializer.SerializeToElement(
            new
            {
                groupId = _teamsTeamId,
                channelId = _teamsChannelId,
            });
        request.AdditionalProperties["messageBody"] = JsonSerializer.SerializeToElement(messageBody);

        try
        {
            var result = await _teamsClient.PostMessageToConversationAsync(
                PostAsFlowBot,
                PostInChannel,
                request);

            _logger.LogInformation("Triage card posted to Teams. MessageId={MessageId}", result?.MessageId);
        }
        catch (ConnectorException ex)
        {
            _logger.LogError(ex, "Failed to post Teams message. ConnectorName={ConnectorName}, ErrorCode={ErrorCode}, ErrorMessage={ErrorMessage}", ex.ConnectorName, ex.ErrorCode, ex.Message);
        }
    }
}
