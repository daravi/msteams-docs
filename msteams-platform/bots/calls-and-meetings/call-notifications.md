---
title: Configure Incoming Call Notifications
description: Learn about incoming notification protocol to convert the call from legacy to Graph format, redirects for region affinity and authenticate the callback.
ms.topic: conceptual
ms.localizationpriority: medium
ms.date: 04/02/2019
---

# Incoming call notifications

In [registering a calls and meetings bot for Microsoft Teams](./registering-calling-bot.md#create-new-bot-or-add-calling-capabilities), the Webhook for calling URL is mentioned. This URL is the webhook endpoint for all incoming calls to your bot.

## Protocol determination

The incoming notification is provided in a legacy format for compatibility with the previous [Skype protocol](/azure/bot-service/dotnet/bot-builder-dotnet-real-time-media-concepts?view=azure-bot-service-3.0&preserve-view=true). To convert the call to the Microsoft Graph protocol, your bot must determine whether the notification is in a legacy format and provides the following response:

```http
HTTP/1.1 204 No Content
```

Your bot receives the notification again, but this time in the Microsoft Graph protocol.

In a future release of the Real-time Media Platform, you can configure the protocol your application supports to avoid receiving the initial callback in the legacy format.

The next section provides details on incoming call notifications redirected for region affinity to your deployment.

## Redirects for region affinity

You call your webhook from the data-center hosting the call. The call starts in any data center and doesn't consider region affinities. The notification is sent to your deployment depending on the GeoDNS resolution. If your application determines, by inspecting the initial notification payload or otherwise, that it needs to run in a different deployment, the application provides the following response:

```http
HTTP/1.1 302 Found
Location: your-new-location
```

Enable your bot to answer an incoming call using the [answer](/graph/api/call-answer?view=graph-rest-1.0&preserve-view=true&tabs=http) API. You can specify the `callbackUri` to handle this particular call. This is useful for stateful instances where your call is handled by a particular partition, and you want to embed this information in the `callbackUri` for routing to the right instance.

The next section provides details on authenticating the callback by inspecting the token posted to your webhook.

## Authenticate the callback

Your bot must inspect the token posted to your webhook to validate the request. Every time the API posts to the webhook, the HTTP POST message contains an OAuth token in the Authorization header as a bearer token, with the audience as your application's App ID.

Your application must validate this token before accepting the callback request.

### Recommended: Use the SDK's built-in authentication provider

The Graph Communications SDK includes `DefaultAuthenticationProvider`, which handles token validation automatically. This is the recommended approach — it validates both legacy tokens and Microsoft Entra tokens, so your bot works seamlessly during the ongoing migration to Entra-based authentication.

To set up authentication when building your communications client:

```csharp
using Microsoft.Graph.Communications.Client;
using Microsoft.Graph.Communications.Client.Authentication;
using Microsoft.Identity.Client;

// Create an MSAL confidential client for outbound Graph API calls
var msalApp = ConfidentialClientApplicationBuilder.Create(appId)
    .WithClientSecret(appSecret) // Or .WithCertificate(cert) for certificate-based auth
    .Build();

// Set up the client with built-in authentication
var builder = new CommunicationsClientBuilder(appName, appId);
builder.SetAuthentication(appId, new MsalTokenProvider(msalApp));
builder.SetNotificationUrl(callbackUrl);
builder.SetServiceBaseUrl(serviceUrl);

var client = builder.Build();
```

The `SetAuthentication` method configures the SDK's `DefaultAuthenticationProvider`, which:

* Validates inbound notification tokens (both legacy OID and Microsoft Entra tokens)
* Uses the provided `ITokenProvider` (via `MsalTokenProvider`) for outbound Graph API requests
* Handles OpenID configuration caching and key refresh automatically

> [!NOTE]
> `MsalTokenProvider` supports app secret, certificate, Subject Name/Issuer (SNI), and managed identity — whatever the `ConfidentialClientApplicationBuilder` was configured with. For custom token acquisition (e.g., SEAL, Azure.Identity), implement the `ITokenProvider` interface directly.

### Migrating from SetAuthenticationProvider (deprecated)

If your bot currently uses `SetAuthenticationProvider` with a custom `AuthenticationProvider` class, migrate to `SetAuthentication` for automatic Entra token support:

```csharp
// Before (deprecated — generates compile warning):
var authProvider = new AuthenticationProvider(appName, appId, appSecret, logger);
builder.SetAuthenticationProvider(authProvider);

// After (recommended):
var msalApp = ConfidentialClientApplicationBuilder.Create(appId)
    .WithClientSecret(appSecret) // Or .WithCertificate(cert) for certificate-based auth
    .Build();
builder.SetAuthentication(appId, new MsalTokenProvider(msalApp));
```

**What changes:**

| Aspect | Old (`SetAuthenticationProvider`) | New (`SetAuthentication`) |
|---|---|---|
| Inbound validation | Custom code fetching OpenID keys | Automatic — validates both legacy and Entra tokens |
| Outbound auth | Custom MSAL token acquisition | `MsalTokenProvider` wraps your existing MSAL app |
| Token migration | Manual — you must update validation code | Automatic — SDK handles both token types |
| Code required | ~100 lines custom `AuthenticationProvider` | 3 lines (create MSAL app + call `SetAuthentication`) |

For custom token acquisition beyond MSAL (e.g., SEAL, Azure.Identity, managed identity), implement `ITokenProvider`:

```csharp
public class MyTokenProvider : ITokenProvider
{
    public async Task<string> AcquireTokenAsync(string tenant)
    {
        // Your custom token acquisition logic
        return accessToken;
    }
}

builder.SetAuthentication(appId, new MyTokenProvider());
```

### Understanding the token format

The following sample shows the HTTP POST message sent to your webhook:

```http
POST https://bot.contoso.com/api/calls
Content-Type: application/json
Authentication: Bearer <TOKEN>

"value": [
    "subscriptionId": "2887CEE8344B47C291F1AF628599A93C",
    "subscriptionExpirationDateTime": "2016-11-20T18:23:45.9356913Z",
    "changeType": "updated",
    "resource": "/app/calls/8A934F51F25B4EE19613D4049491857B",
    "resourceData": {
        "@odata.type": "#microsoft.graph.call",
        "state": "Established"
    }
]
```

The OAuth token has the following values:

```json
{
    "aud": "0efc74f7-41c3-47a4-8775-7259bfef4241",
    "iss": "https://api.botframework.com",
    "iat": 1466741440,
    "nbf": 1466741440,
    "exp": 1466745340,
    "tid": "1fdd12d0-4620-44ed-baec-459b611f84b2"
}
```

Each OAuth token value is used as follows:

* `aud` where audience is the App ID URI specified for the application.
* `tid` is the tenant id for Contoso.com.
* `iss` is the token issuer, `https://api.botframework.com`.

> [!IMPORTANT]
> The token issuer and signing mechanism are being migrated from custom OpenID tokens to standard Microsoft Entra tokens. If you use `DefaultAuthenticationProvider` (recommended), this migration is handled automatically. If you implement custom token validation, ensure your code accepts tokens from both `https://api.botframework.com` and Microsoft Entra issuers (`https://login.microsoftonline.com/{tid}/v2.0`).

### Custom token validation (advanced)

If you need to implement custom token validation instead of using `DefaultAuthenticationProvider`, your webhook must validate the token, ensure it hasn't expired, and check whether it has been signed by the published OpenID configuration. You must also check whether `aud` matches your App ID before accepting the callback request.

The OpenID configuration is published at <https://api.aps.skype.com/v1/.well-known/OpenIdConfiguration>.

For more information, see the [EchoBot sample](https://github.com/microsoftgraph/microsoft-graph-comms-samples/blob/master/Samples/PublicSamples/EchoBot/src/EchoBot/Bot/BotService.cs).

## Next step

> [!div class="nextstepaction"]
> [Requirements and considerations for application-hosted media bots](~/bots/calls-and-meetings/requirements-considerations-application-hosted-media-bots.md)

## See also

* [Build bots for Teams](../what-are-bots.md)
* [Calls and online meetings bots](calls-meetings-bots-overview.md)
* [Work with the cloud communications API in Microsoft Graph](/graph/api/resources/communications-api-overview)
* [Authentication flow for bots in Microsoft Teams](../how-to/authentication/auth-flow-bot.md)
* [Set up an auto attendant](/microsoftteams/create-a-phone-system-auto-attendant)
* [Set up auto answer for Microsoft Teams Rooms on Android and Teams video phone devices](/microsoftteams/set-up-auto-answer-on-teams-android)
