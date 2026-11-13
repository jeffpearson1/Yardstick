# Workload identity federation

Yardstick normally signs in to Microsoft Graph with the client secret stored in
Windows Credential Manager. When it runs in a CI/CD pipeline it can use **workload
identity federation** instead: the pipeline platform issues a short-lived ID token,
and Entra ID accepts it in place of a secret. No secret exists anywhere, so there
is nothing to store, rotate or leak.

Yardstick can get the ID token from:

- **GitHub Actions**, from the job's built-in ID token endpoint;
- **Azure DevOps Pipelines**, through a workload identity federation service
  connection;
- **a file**, for any other platform that can write an ID token to disk.

With none of the variables below set, Yardstick signs in with the stored secret
exactly as before.

## How it works

Every time Yardstick needs a Graph token it:

1. gets a fresh ID token from the source;
2. sends it to Entra as `client_assertion` in the usual client-credentials request,
   in place of `client_secret`.

Yardstick renews the Graph token before it expires, and each renewal gets a new
ID token, so a run can last as long as the job allows.

In federated mode Yardstick also skips everything tied to a stored secret: the
Credential Manager lookup, the legacy `preferences.yaml` secret migration, the
secret-expiry lookup and the expiry email. `Test-YardstickSetup.ps1` reports the
federated identity instead of a stored credential and signs in with it.

## Setting it up

### 1. The app registration

Use the same app registration and Graph permissions as a secret-based install (see
[Microsoft Graph permissions](Installation.md#microsoft-graph-permissions)). On the
app registration, open **Certificates & secrets > Federated credentials** and add a
credential that matches the pipeline's ID token:

| Source | Issuer | Subject |
|---|---|---|
| GitHub Actions | `https://token.actions.githubusercontent.com` | Depends on the trigger, e.g. `repo:<owner>/<repo>:ref:refs/heads/main` or `repo:<owner>/<repo>:environment:<name>` |
| Azure DevOps | Shown on the service connection (`https://vstoken.dev.azure.com/<organization id>`) | Shown on the service connection (`sc://<organization>/<project>/<connection>`) |
| File | The issuer of the token you supply | The subject of the token you supply |

The audience is `api://AzureADTokenExchange` in every case. Entra's
**GitHub Actions deploying Azure resources** scenario fills in the GitHub issuer and
builds the subject for you.

Once federated runs work, the app registration needs no client secret.

### 2. The environment variables

| Variable | Value |
|---|---|
| `YARDSTICK_TENANT_ID` | Tenant ID |
| `YARDSTICK_CLIENT_ID` | The app registration's client (application) ID |
| `YARDSTICK_SERVICE_CONNECTION_ID` | Azure DevOps only: the service connection's ID |
| `YARDSTICK_FEDERATED_TOKEN_FILE` | File only: full path to the file holding the ID token |

Setting any of these turns federated mode on. From then on the tenant, the client
and one ID token source are required, and a missing, malformed or conflicting input
stops the run with an error naming it. **Federated mode never falls back to a client
secret.**

The source is chosen in this order:

1. `YARDSTICK_FEDERATED_TOKEN_FILE`, when set;
2. Azure DevOps, when `YARDSTICK_SERVICE_CONNECTION_ID` is set;
3. GitHub Actions, when the job offers an ID token (`ACTIONS_ID_TOKEN_REQUEST_URL`).

Setting both of the first two is an error.

The job also needs a working Yardstick folder: a checkout of this repository,
`Local\preferences.yaml` with your settings, your icons, and the PowerShell modules
from [Requirements](Installation.md#requirements). The examples below assume an
earlier step provides them.

### GitHub Actions

Give the job `id-token: write` and set the tenant and client:

```yaml
jobs:
  yardstick:
    runs-on: windows-latest
    permissions:
      id-token: write
      contents: read
    steps:
      - uses: actions/checkout@v7
      # ... install modules, provide Local\preferences.yaml and icons ...
      - name: Run Yardstick
        shell: pwsh
        env:
          YARDSTICK_TENANT_ID: ${{ vars.YARDSTICK_TENANT_ID }}
          YARDSTICK_CLIENT_ID: ${{ vars.YARDSTICK_CLIENT_ID }}
        run: ./Yardstick.ps1 -All -NoInteractive -NoEmail
```

Yardstick requests each ID token from `ACTIONS_ID_TOKEN_REQUEST_URL` with the
audience `api://AzureADTokenExchange`, so no Azure login action is needed.

### Azure DevOps

Create an **Azure Resource Manager** service connection that uses workload identity
federation, for the same app registration (or let Azure DevOps create one and use
it as Yardstick's app registration).

Azure DevOps only issues an ID token for a service connection that a task in the
same stage references. The simplest way to satisfy that is to run Yardstick inside
`AzurePowerShell@5`, which names the connection. The task signs in to Azure first; Yardstick doesn't use that session.
For that sign-in to work, the agent needs the `Az.Accounts` module and the identity
needs a role (Reader is enough) on the connection's subscription.

Map `SYSTEM_ACCESSTOKEN`, which Yardstick uses to request ID tokens, and pass the
connection's ID:

```yaml
- task: AzurePowerShell@5
  displayName: Run Yardstick
  inputs:
    azureSubscription: <your service connection>
    azurePowerShellVersion: LatestVersion
    pwsh: true
    ScriptType: InlineScript
    Inline: |
      & ./Yardstick.ps1 -All -NoInteractive -NoEmail
      if ($LASTEXITCODE) {
        Write-Host "##vso[task.complete result=Failed;]Yardstick exited $LASTEXITCODE"
        exit $LASTEXITCODE
      }
  env:
    SYSTEM_ACCESSTOKEN: $(System.AccessToken)
    YARDSTICK_SERVICE_CONNECTION_ID: $(serviceConnectionId)
    YARDSTICK_TENANT_ID: $(tenantId)
    YARDSTICK_CLIENT_ID: $(clientId)
```

Use an inline script as shown: with `ScriptType: FilePath` the task drops the
script's exit code, so a failed run would still pass.

### A token file

On any other platform, write an ID token whose audience is
`api://AzureADTokenExchange` to a file and set `YARDSTICK_FEDERATED_TOKEN_FILE` to
its full path. Yardstick reads the file every time it signs in, so a platform that refreshes
the file (such as a Kubernetes projected service account token) keeps a long run
going. A token written once is used until it expires, so the run must finish within
its lifetime.

## Errors and what they mean

Every federated failure starts with `Workload identity federation`.

| Message contains | Cause |
|---|---|
| `is not configured correctly: ... is not set` | A variable is missing. For Azure DevOps, `SYSTEM_OIDCREQUESTURI` is missing when Yardstick isn't running in a pipeline job |
| `is set but blank` | A variable holds only spaces, often an empty pipeline variable |
| `no ID token source is set` | The tenant or client is set, but there is no token file, no service connection and no GitHub Actions ID token |
| `is not an absolute https URI` | The platform's ID token endpoint isn't HTTPS. Plain HTTP is accepted only to this machine |
| `is not a full path` | `YARDSTICK_FEDERATED_TOKEN_FILE` is relative; Yardstick changes directory during a run |
| `give the job 'permissions: id-token: write'` | GitHub Actions isn't offering the job an ID token |
| `is the unresolved pipeline variable '$(...)'` | The Azure DevOps variable named in the macro doesn't exist (usually a typo) |
| `SYSTEM_ACCESSTOKEN is not available` | The step doesn't map `SYSTEM_ACCESSTOKEN: $(System.AccessToken)` |
| `requesting an ID token from Azure DevOps failed (HTTP 401)` | The job access token was rejected |
| `There is no explicit reference to service connection ...` (HTTP 400) | No task in the stage names the service connection; run Yardstick inside `AzurePowerShell@5` as above |
| `requesting an ID token from Azure DevOps failed (HTTP 400/403/404)` | Wrong service connection ID, the connection doesn't use workload identity federation, or the pipeline isn't permitted to use it |
| `requesting an ID token from GitHub Actions failed (HTTP 401/403)` | The job lacks `id-token: write` |
| `AADSTS700213` or `AADSTS70021` | No federated credential on the app registration matches. The message shows the ID token's `iss` and `sub`; compare them with the federated credential |
| `AADSTS700016` | Wrong `YARDSTICK_CLIENT_ID`, or the app isn't in `YARDSTICK_TENANT_ID` |
| `AADSTS700024` | The ID token is outside its valid time range: agent clock skew, or an expired token file |
| `AADSTS700027` / `AADSTS50027` | Entra couldn't validate the token's signature or format |

Tokens are never written to the console, the log or error messages.

## Testing

`Tests/Unit/YardstickFederation.Tests.ps1` and
`Tests/Unit/YardstickFederation.Wire.Tests.ps1` run with the rest of the unit tests.
The wire tests send the real requests to local stand-ins for each endpoint.

`Tests/Federation/Invoke-FederationSmokeTest.ps1` runs the real `Yardstick.ps1`
startup once per source against local stand-ins for the ID token endpoints, Entra
and Graph, and checks that no secret is used or looked for:

```powershell
pwsh -NoProfile -File Tests/Federation/Invoke-FederationSmokeTest.ps1
```
