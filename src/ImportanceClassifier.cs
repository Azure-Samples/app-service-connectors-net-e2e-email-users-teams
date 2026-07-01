using System.Text.RegularExpressions;
using Microsoft.Extensions.Logging;

namespace EmailTriage.AppService;

/// <summary>
/// Decides whether an inbound email warrants posting to the triage Teams channel.
///
/// An email is treated as IMPORTANT when ANY of the following holds:
///   1. Sender is in the <c>IMPORTANT_SENDERS</c> allowlist (comma-separated
///      mail/UPN addresses, e.g. your manager + skip-level + key stakeholders).
///   2. The mail's native Importance flag is "high" (sender explicitly marked it).
///   3. The subject/body contains urgency / action-required language above
///      <see cref="ContentScoreThreshold"/> (heuristic — see <see cref="UrgencyPatterns"/>).
///
/// Each match contributes a human-readable reason to <see cref="ImportanceVerdict.Reasons"/>
/// so the Teams card can show *why* an email made it through.
///
/// Ported verbatim from the Azure Functions e2e sample — pure in-process logic with no
/// dependency on any host, so it moves to App Service unchanged.
/// </summary>
public sealed class ImportanceClassifier
{
    private const int ContentScoreThreshold = 2;

    // Pattern => weight => human-readable label. Weights are added; threshold is
    // intentionally low (2) so a single strong signal (e.g. "[URGENT]" subject)
    // OR two weaker ones (e.g. "please review" + "by EOD") cross the bar.
    private static readonly (Regex Pattern, int Weight, string Label)[] UrgencyPatterns =
    [
        // Strong markers — single hit is enough
        (new Regex(@"\[\s*(urgent|action(\s+required|\s+needed)?|important|asap|escalat\w*)\s*\]", RegexOptions.IgnoreCase | RegexOptions.Compiled), 2, "bracketed urgency tag"),
        (new Regex(@"\b(p0|p1|sev\s?1|sev\s?2|outage|incident|blocker)\b", RegexOptions.IgnoreCase | RegexOptions.Compiled), 2, "incident/blocker keyword"),
        (new Regex(@"\b(action\s+(required|needed)|decision\s+(required|needed)|approval\s+(required|needed))\b", RegexOptions.IgnoreCase | RegexOptions.Compiled), 2, "explicit action/decision required"),

        // Medium markers
        (new Regex(@"\b(asap|urgent(ly)?|immediately|right\s+away)\b", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1, "urgency adverb"),
        (new Regex(@"\b(escalat\w*)\b", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1, "escalation language"),
        (new Regex(@"\b(deadline|due\s+(by|date)|need(ed)?\s+by|by\s+(eod|cob|end\s+of\s+(day|week)))\b", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1, "deadline language"),
        (new Regex(@"\b(please\s+(review|respond|approve|confirm|sign|reply))\b", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1, "polite ask"),
        (new Regex(@"\b(blocking|blocked\s+on|stuck\s+on|waiting\s+on\s+you)\b", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1, "blocked / waiting on you"),
    ];

    // Subject in ALL CAPS (>= 4 alpha chars, ignoring punctuation/numbers) usually
    // signals the sender is shouting for attention.
    private static readonly Regex AllCapsSubject = new(@"^[^a-z]*[A-Z][^a-z]{3,}$", RegexOptions.Compiled);

    private readonly ILogger<ImportanceClassifier> _logger;
    private readonly HashSet<string> _importantSenders;

    public ImportanceClassifier(ILogger<ImportanceClassifier> logger)
    {
        _logger = logger;

        var raw = Environment.GetEnvironmentVariable("IMPORTANT_SENDERS") ?? "";
        _importantSenders = new HashSet<string>(
            raw.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries),
            StringComparer.OrdinalIgnoreCase);

        _logger.LogInformation(
            "ImportanceClassifier loaded with {Count} important-sender entries.",
            _importantSenders.Count);
    }

    public ImportanceVerdict Classify(
        string? senderEmail,
        string? subject,
        string? body,
        string? bodyPreview,
        string? nativeImportance)
    {
        var reasons = new List<string>();

        // 1) Important-senders allowlist
        if (!string.IsNullOrWhiteSpace(senderEmail) && _importantSenders.Contains(senderEmail))
        {
            reasons.Add("Sender in IMPORTANT_SENDERS allowlist");
        }

        // 2) Native importance flag
        if (string.Equals(nativeImportance, "high", StringComparison.OrdinalIgnoreCase))
        {
            reasons.Add("Sender marked email High importance");
        }

        // 3) Content scoring
        reasons.AddRange(ScoreContent(subject, body ?? bodyPreview));

        return new ImportanceVerdict(reasons.Count > 0, reasons);
    }

    private static List<string> ScoreContent(string? subject, string? body)
    {
        var reasons = new List<string>();
        var haystack = $"{subject ?? ""}\n{body ?? ""}";
        if (string.IsNullOrWhiteSpace(haystack)) return reasons;

        var score = 0;
        var matchedLabels = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var (pattern, weight, label) in UrgencyPatterns)
        {
            if (pattern.IsMatch(haystack))
            {
                score += weight;
                matchedLabels.Add(label);
            }
        }

        // ALL-CAPS subject contributes one point on top of any keyword hits.
        if (!string.IsNullOrWhiteSpace(subject) &&
            subject.Length >= 6 &&
            AllCapsSubject.IsMatch(subject))
        {
            score += 1;
            matchedLabels.Add("ALL-CAPS subject");
        }

        if (score >= ContentScoreThreshold)
        {
            reasons.Add($"Content signals (score {score}): {string.Join(", ", matchedLabels)}");
        }

        return reasons;
    }
}

public sealed record ImportanceVerdict(bool IsImportant, IReadOnlyList<string> Reasons);
