using System.Text.Json;
using Azure.Core;
using Azure.Identity;
using Azure.Monitor.OpenTelemetry.AspNetCore;
using Azure.Connectors.Sdk.Office365;
using Azure.Connectors.Sdk.Office365.Models;
using Azure.Connectors.Sdk.Office365Users;
using Azure.Connectors.Sdk.Teams;
using EmailTriage.AppService;

var builder = WebApplication.CreateBuilder(args);

// DefaultAzureCredential works in both environments:
//   - In Azure App Service: uses the user-assigned managed identity whose client id is in
//     AZURE_CLIENT_ID (the same identity that has access policies on the 3 connections).
//   - Locally: falls back to the developer's `az login` / VS / VS Code credentials.
var credential = new DefaultAzureCredential(new DefaultAzureCredentialOptions
{
    ManagedIdentityClientId = Environment.GetEnvironmentVariable("AZURE_CLIENT_ID")
});
builder.Services.AddSingleton<TokenCredential>(credential);

// OpenTelemetry -> Azure Monitor (App Insights). App Insights is provisioned with
// disableLocalAuth: true, so the exporter authenticates via the same managed identity.
// Locally APPLICATIONINSIGHTS_CONNECTION_STRING is unset and this whole block no-ops.
var appInsightsConnectionString = Environment.GetEnvironmentVariable("APPLICATIONINSIGHTS_CONNECTION_STRING");
if (!string.IsNullOrWhiteSpace(appInsightsConnectionString))
{
    builder.Services.AddOpenTelemetry().UseAzureMonitor(o =>
    {
        o.ConnectionString = appInsightsConnectionString;
        o.Credential = credential;
    });
}

// The three connector clients — identical to the Functions sample's DI registration.
// Each takes (connectionRuntimeUrl, TokenCredential); the runtime URLs come from the
// Connector Namespace connections wired by infra/main.bicep.
builder.Services.AddSingleton(sp => new TeamsClient(
    new Uri(RequireSetting("TEAMS_CONNECTION_RUNTIME_URL")),
    sp.GetRequiredService<TokenCredential>()));

builder.Services.AddSingleton(sp => new Office365Client(
    new Uri(RequireSetting("OFFICE365_CONNECTION_RUNTIME_URL")),
    sp.GetRequiredService<TokenCredential>()));

builder.Services.AddSingleton(sp => new Office365UsersClient(
    new Uri(RequireSetting("OFFICE365USERS_CONNECTION_RUNTIME_URL")),
    sp.GetRequiredService<TokenCredential>()));

builder.Services.AddSingleton<ImportanceClassifier>();
builder.Services.AddSingleton<EmailTriageProcessor>();

// Case-insensitive matching so the connector's camelCase JSON binds to the SDK model.
var serializerOptions = new JsonSerializerOptions(JsonSerializerDefaults.Web)
{
    PropertyNameCaseInsensitive = true,
};

var app = builder.Build();

// Liveness probe (open — kept out of the Easy Auth-protected path only if you exclude it;
// by default Easy Auth guards everything, which is fine for a private triage app).
app.MapGet("/healthz", () => Results.Ok(new { status = "healthy" }));

// Connector trigger callback. This is the App Service replacement for the Functions
// `[ConnectorTrigger]` binding + `/runtime/webhooks/connector` endpoint.
//
// AUTH: there is intentionally no key/secret check here. App Service **Easy Auth**
// (authsettingsV2) validates the Entra bearer token the Connector Namespace attaches
// (authentication.type = ManagedServiceIdentity) at the edge — signature, issuer,
// audience, and that the caller's oid is the trigger UAMI's principalId. Any request
// that reaches this handler has already passed that gate, so the code just does the work.
app.MapPost("/api/onNewEmail", async (
    HttpRequest request,
    EmailTriageProcessor processor,
    ILoggerFactory loggerFactory,
    CancellationToken cancellationToken) =>
{
    var logger = loggerFactory.CreateLogger("OnNewEmailEndpoint");
    logger.LogInformation("OnNewEmail callback received (caller pre-validated by App Service built-in authentication).");

    Office365OnNewEmailTriggerPayload? payload;
    try
    {
        payload = await request.ReadFromJsonAsync<Office365OnNewEmailTriggerPayload>(
            serializerOptions, cancellationToken);
    }
    catch (JsonException ex)
    {
        logger.LogError(ex, "Failed to deserialize connector trigger payload.");
        return Results.BadRequest(new { error = "invalid payload" });
    }

    await processor.ProcessAsync(payload, cancellationToken);

    // The connector runtime expects a timely 2xx to consider the callback delivered.
    return Results.Ok();
});

app.Run();

// Fail loudly at boot if a required connection URL setting is missing — otherwise the
// connector clients silently get an empty BaseAddress and every call throws deep in the
// request pipeline.
static string RequireSetting(string name)
{
    var value = Environment.GetEnvironmentVariable(name);
    if (string.IsNullOrWhiteSpace(value))
    {
        throw new InvalidOperationException(
            $"Required app setting '{name}' is not set. " +
            $"For local development add it to appsettings.Development.json or user-secrets. " +
            $"For Azure deployments it should be wired via infra/main.bicep.");
    }
    return value;
}
