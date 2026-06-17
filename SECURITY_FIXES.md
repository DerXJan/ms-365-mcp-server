# M365MCP — Security Hardening Instructions (Org-Mode, Azure App Service)

**Audience:** future Claude session (or any engineer) implementing the security
fixes derived from the org-mode security review against
`MCPsecurityGuideline.txt`.

**Deployment context:**

- Microsoft 365 MCP server is run **directly on Azure App Service** (Node
  process, no Docker, no Container Apps).
- `--http` (Streamable HTTP / OAuth 2.1) transport in `--org-mode`.
- TLS is terminated by Azure App Service in front of the Node process.
- A custom Entra app registration is used (no built-in default `clientId`).
- Secrets are pulled from Azure Key Vault via managed identity
  (`MS365_MCP_KEYVAULT_URL`).

**Out of scope:** Docker / containers, personal-MSA stdio mode, the local
`startlocal.sh` test script.

---

## How to use this document

Each item below is a self-contained task with:

1. **Why** — which guideline section it satisfies.
2. **Files to touch** — concrete paths.
3. **Acceptance criteria** — how you know it's done.
4. **Implementation notes** — patterns, code sketches, test ideas.

Work top-to-bottom; the **Must-fix** block is the gate for security approval.
The **Should-fix** block can ship in a follow-up release. Each task is sized to
be doable in one PR.

When you start, create a tracking branch (e.g. `security/orgmode-hardening`)
and one commit per numbered task. Update this file by checking the box at the
top of each task as you finish it.

---

## Conventions

- **Strict mode:** all new validation paths must **fail closed** (deny on
  error). If you can't decide, return `401`/`403`/`503` — never fall through.
- **Logging discipline:** never log `code`, `code_verifier`, `client_secret`,
  `refresh_token`, full `Authorization` headers, or raw request bodies on
  `/token` / `/register`.
- **No new env vars without docs.** If you add an env var, add it to
  [README.md](README.md) under the existing env-var table and to a new
  `docs/security-config.md` deployment doc.
- **Tests:** add unit tests under [src/__tests__/](src/__tests__/) for every
  middleware/validator. Adversarial cases (malformed JWT, wrong tenant, wrong
  audience, missing scopes, oversized body, unknown fields) are mandatory.

---

# MUST-FIX (security-approval gate)

## 1. [ ] Local JWT validation in the bearer middleware

**Why:** Guideline §2 — audience-bound tokens, per-hop re-validation. The
current middleware in [src/lib/microsoft-auth.ts:14-69](src/lib/microsoft-auth.ts#L14-L69)
only checks JWT `exp` and forwards opaque tokens unverified.

**Files:**

- [src/lib/microsoft-auth.ts](src/lib/microsoft-auth.ts) — replace
  `microsoftBearerTokenAuthMiddleware`.
- New file: `src/lib/jwt-validator.ts`.
- New tests: `src/__tests__/jwt-validator.test.ts`.

**Acceptance:**

- Reject any bearer that is not a JWT (3 base64url parts) with `401
  invalid_token`.
- Verify signature using Entra JWKS for the configured tenant (cache JWKS;
  refresh on `kid` miss).
- Verify claims: `iss` matches the configured authority + tenant, `aud` matches
  the configured MCP resource URI (new env var
  `MS365_MCP_EXPECTED_AUDIENCE`), `tid` ∈ allowlist (new env var
  `MS365_MCP_ALLOWED_TENANTS`, comma-separated), `exp` / `nbf` within skew
  (60 s).
- Required scopes (`scp`) or app-roles (`roles`) per tool — see Task 7 for the
  per-tool check; this task only validates *presence* of the claim.
- Caller identity (`oid`, `tid`, `upn` if present, `appid`) is attached to
  `req` for downstream logging (Task 11).

**Implementation notes:**

- Use the `jose` package (`npm i jose`). Pin a fixed version (no `^`).
- JWKS URI: `${authority}/${tenantId}/discovery/v2.0/keys`. Resolve
  `authority` via `getCloudEndpoints` from
  [src/cloud-config.ts](src/cloud-config.ts).
- Cache: `jose.createRemoteJWKSet(new URL(jwksUri))` already caches. Wrap so
  one validator instance is reused across requests; do not create per-request.
- New env vars to add:
  - `MS365_MCP_EXPECTED_AUDIENCE` (required in HTTP mode) — the `aud` value
    your app registration issues for the MCP resource. Typically `api://<app
    GUID>` or the Application ID URI.
  - `MS365_MCP_ALLOWED_TENANTS` (required in production) — comma-separated
    tenant GUIDs. Reject if `tid` not in the set.
  - `MS365_MCP_JWKS_CACHE_MAX_AGE_S` (optional, default 600).
- On any failure, return `401` with `WWW-Authenticate: Bearer
  resource_metadata=…, error="invalid_token", error_description="…"` matching
  the existing helper [src/lib/microsoft-auth.ts:5-10](src/lib/microsoft-auth.ts#L5-L10).
- Keep an "opaque token bypass" only behind an explicit
  `MS365_MCP_ALLOW_OPAQUE_TOKENS=1` flag (default off) for backward compat
  with MSA personal-account testing — but it must *never* be on in
  org-mode/production.

---

## 2. [ ] Refuse `tenantId='common'` and built-in clientId in production

**Why:** Guideline §2 / §11 — separate identity per server instance, no
implicit cross-tenant trust.

**Files:**

- [src/secrets.ts](src/secrets.ts) — add a production-mode check after secrets
  load.
- [src/index.ts](src/index.ts) or a new `src/lib/startup-checks.ts` — wire the
  check into `main()`.

**Acceptance:**

- New env var `MS365_MCP_PRODUCTION=1` (set in App Service config) gates the
  checks.
- When `MS365_MCP_PRODUCTION=1` and `--http`:
  - Refuse to start if `secrets.tenantId` is missing, empty, or `common`.
  - Refuse to start if `secrets.clientId` equals the built-in default
    returned by `getDefaultClientId()` in
    [src/cloud-config.ts](src/cloud-config.ts).
  - Refuse to start if `MS365_MCP_EXPECTED_AUDIENCE` or
    `MS365_MCP_ALLOWED_TENANTS` are unset (Task 1 dependency).
- Failure mode: `console.error` + `process.exit(1)` before `app.listen`. Never
  start in a half-configured state.

**Implementation notes:**

- Centralize all production-mode startup assertions in one function
  (`assertProductionConfig(secrets, options)`), call once from `main()`.
- Log the assertion summary at startup (without secret values) so an
  operator can verify production mode was detected.

---

## 3. [ ] Disable Dynamic Client Registration by default + redirect-URI allowlist

**Why:** Guideline §3 — no dynamic unverified onboarding. Today
[src/cli.ts:188-196](src/cli.ts#L188-L196) forces it on in HTTP mode and
[src/server.ts:275-292](src/server.ts#L275-L292) echoes back any
`redirect_uris`.

**Files:**

- [src/cli.ts](src/cli.ts) — change default behaviour.
- [src/server.ts](src/server.ts) — add allowlist.

**Acceptance:**

- DCR is **off** by default in HTTP mode. Operator must pass
  `--enable-dynamic-registration` (already exists as a flag) to turn it on.
- When DCR is on, every value in `body.redirect_uris` must match the new env
  var `MS365_MCP_ALLOWED_REDIRECT_URIS` (comma-separated). Mismatched URIs
  cause a `400 invalid_redirect_uri`.
- HTTPS-only: refuse `http://` redirect URIs except `http://localhost(:port)`
  for development.
- Reject `redirect_uris` that contain wildcards, fragments, or query strings.
- Reject when `redirect_uris` is empty.

**Implementation notes:**

- Defensive default change in `src/cli.ts`:

  ```ts
  if (options.http) {
    options.enableDynamicRegistration =
      options.dynamicRegistration === true ? true : false;
  }
  ```

  Note: Commander's `--no-dynamic-registration` sets `dynamicRegistration =
  false`. With the new default-off behaviour the `--no-` flag becomes
  redundant; keep it for backward compatibility but document it as a no-op.

- Update the metadata advertiser
  ([src/server.ts:251-256](src/server.ts#L251-L256)) so it doesn't advertise
  `registration_endpoint` when DCR is off.

---

## 4. [ ] Strict Zod schemas — reject unknown fields

**Why:** Guideline §3 — strict schema validation, reject unknown.

**Files:**

- [src/graph-tools.ts](src/graph-tools.ts) — wrap every `paramSchema` object
  in `z.object(...).strict()` before passing to `server.tool(...)`.
- The generated client at [src/generated/client.ts](src/generated/client.ts)
  produces request-body Zod schemas. They must be passed through a
  `.strict()` walker — implement once in
  `src/lib/strict-schema.ts`.
- Tool-call body parsing (the `paramDef.schema.safeParse` paths in
  [src/graph-tools.ts:255-273](src/graph-tools.ts#L255-L273)) must also use
  the strict schema.

**Acceptance:**

- A request with an unknown top-level parameter is rejected with a clear
  error, not silently dropped.
- A request body containing an unknown nested property is rejected.
- Existing test suite still passes (you may need to fix internal tests that
  relied on lax parsing).
- Add adversarial tests that send `{ foo: "bar", body: { unexpected: 1 } }`
  for at least one read tool and one write tool.

**Implementation notes:**

- For deeply nested `z.object(...)` schemas, write a recursive helper:

  ```ts
  export function deepStrict<T extends z.ZodTypeAny>(s: T): T {
    if (s instanceof z.ZodObject)
      return s.strict().extend(
        Object.fromEntries(
          Object.entries(s.shape).map(([k, v]) => [k, deepStrict(v as any)]),
        ),
      ) as any;
    if (s instanceof z.ZodArray) return z.array(deepStrict(s.element)) as any;
    if (s instanceof z.ZodOptional)
      return deepStrict(s.unwrap()).optional() as any;
    if (s instanceof z.ZodNullable)
      return deepStrict(s.unwrap()).nullable() as any;
    if (s instanceof z.ZodUnion)
      return z.union(s.options.map(deepStrict) as any) as any;
    return s;
  }
  ```

- Allow `z.record(...)` (used for arbitrary maps) to pass through unmodified —
  the discriminator is "is it `ZodObject`?".

---

## 5. [ ] Default to `--read-only`; explicit per-tool write opt-in

**Why:** Guideline §5 — least privilege, no god-tools, structured action
control. Today the entire delegated Graph surface is reachable in `--org-mode`
unless `--read-only` is set.

**Files:**

- [src/cli.ts](src/cli.ts) — invert default.
- [src/graph-tools.ts](src/graph-tools.ts) — load and consult policy.
- New file: `src/lib/write-policy.ts`.
- New file: `policy/write-policy.example.json` (committed example, **not**
  the production policy).
- Deployment doc: document where the production policy lives in App Service
  (recommended: a Key Vault secret `ms365-mcp-write-policy` containing JSON,
  or a path on the App Service file system).

**Acceptance:**

- New env var `MS365_MCP_ALLOW_WRITES=1` is required to enable any non-GET
  tool. Without it, the server runs read-only (existing `--read-only` logic
  applies).
- When `MS365_MCP_ALLOW_WRITES=1`, a policy is loaded from
  `MS365_MCP_WRITE_POLICY_PATH` (file path) or `MS365_MCP_WRITE_POLICY_JSON`
  (inline JSON; for Key Vault). The policy is a simple allowlist:

  ```json
  {
    "allowedWriteTools": [
      "send-mail",
      "create-online-meeting"
    ]
  }
  ```

- Tools that are POST/PATCH/DELETE and not in `allowedWriteTools` are skipped
  during registration with a logged INFO line.
- Loading failure (missing file, malformed JSON) → fail closed: server starts
  as if no writes are allowed and logs WARN.
- An `--read-only` flag still wins (forces read-only regardless of policy).

**Implementation notes:**

- Hash the loaded policy JSON and log the SHA-256 at startup so audit can
  detect drift.
- Do not auto-reload the policy at runtime; require a process restart so
  changes go through your normal App Service deploy/restart workflow.

---

## 6. [ ] Human-in-the-loop confirmation for high-risk write tools

**Why:** Guideline §5 — HITL for high-risk actions.

**Files:**

- New file: `src/lib/hitl.ts`.
- [src/graph-tools.ts](src/graph-tools.ts) — wrap `executeGraphTool` for
  flagged tools.
- Endpoints data: extend [src/endpoints.json](src/endpoints.json) schema
  with an optional `requiresApproval: true` field (and the matching
  `EndpointConfig` interface).

**Acceptance:**

- Tools flagged `requiresApproval: true` cannot execute on the first call.
  They must return a structured "approval required" payload containing an
  approval request ID and a one-time challenge.
- The MCP client (or a human via a separate approval channel) calls a
  separate `approve-tool-call` MCP tool with `{ approvalId, decision }`.
- The original call must be re-issued within a TTL (default 5 min) with the
  `approvalToken` returned from `approve-tool-call`.
- The approval store must persist across stateless HTTP requests — use a
  module-level `Map` with TTL eviction, same hygiene pattern as the existing
  PKCE store ([src/server.ts:340-359](src/server.ts#L340-L359)).
- Tools to flag at minimum: any tool whose required scopes include
  `Mail.Send*`, `ChannelMessage.Send`, `Sites.ReadWrite.All`, plus every
  `DELETE` and any `PATCH` that touches `users`, `groups`, `sites`, or
  `teams`.

**Implementation notes:**

- The simplest production-grade approval channel is a Teams adaptive card
  posted to the user via Graph; that's a follow-up. For the first iteration,
  it is acceptable for `approve-tool-call` to be a separate MCP tool that the
  *user* (not the LLM agent) is expected to invoke after reviewing the
  pending action — document this in [README.md](README.md) under "High-risk
  tools".
- Approval store key = SHA-256 of `(toolName, sortedParamsJson, oid, tid)` so
  the same call cannot be approved once and replayed with different
  parameters.

---

## 7. [ ] Per-tool scope/role authorization (PDP)

**Why:** Guideline §2 — authorization tuple `(user, tenant, tool, action,
resource)`.

**Files:**

- New file: `src/lib/pdp.ts`.
- [src/graph-tools.ts](src/graph-tools.ts) — call PDP at the top of
  `executeGraphTool`, before any Graph traffic.

**Acceptance:**

- Every tool call passes through a policy decision point that verifies:
  - The bearer JWT contains all of the `endpoint.scopes` (or
    `endpoint.workScopes` in org-mode) declared in `endpoints.json` for the
    tool's tool name.
  - The `tid` matches the resolved tenant for the request.
  - In multi-account stdio mode, the resolved account's `tid` matches the
    token's `tid`.
- A failed decision returns a tool-call error with a stable code
  (`unauthorized_tool_call`) and is logged at WARN with the caller's `oid`.

**Implementation notes:**

- Pull the validated JWT claims from `req` (set in Task 1) via
  `requestContext` (already used in
  [src/request-context.ts](src/request-context.ts)). Extend `RequestContext`
  to carry the validated claim set.
- Centralize tool-name → required-scope lookup using the existing
  `endpointsData` array; you already have it loaded in
  [src/graph-tools.ts:48-50](src/graph-tools.ts#L48-L50).

---

## 8. [ ] Fail-closed bearer middleware on malformed tokens

**Why:** Guideline §14 — fail-closed; reject unknown tokens, missing identity.

**Files:**

- [src/lib/microsoft-auth.ts](src/lib/microsoft-auth.ts) — replaced as part of
  Task 1, but call this out as a separate acceptance criterion so it isn't
  forgotten.

**Acceptance:**

- A token that is not a parseable JWT (after the `MS365_MCP_ALLOW_OPAQUE_TOKENS`
  bypass is off, which is the production default) → `401`.
- A JWT missing `exp`, `iss`, `aud`, `tid`, or `oid` → `401`.
- Unit tests cover: empty bearer, two-part token, three-part token with
  invalid base64url, valid header but invalid signature, expired, wrong
  audience, wrong tenant.

---

## 9. [ ] Fail-closed sensitivity-label filter

**Why:** Guideline §14 — fail-closed for security controls.

**Files:**

- [src/lib/sensitivity-filter.ts](src/lib/sensitivity-filter.ts) — change
  the catalog-unavailable branch from "log warn + allow all" to "block all".

**Acceptance:**

- When `loadSensitivityFilterConfig().enabled === true` and the label catalog
  cannot be resolved (`fetchLabelCatalog` returns empty), document-context
  tools must return a clear error result (`isError: true`) with the message
  *"Sensitivity filter active but label catalog could not be resolved;
  refusing to return documents."*
- The email-context filter (which checks the `sensitivity` string property
  directly without needing the catalog) is unaffected — keep its current
  fail-closed-by-design behaviour.
- Add unit tests for: catalog fetch failure → list call returns isError;
  catalog returns 0 entries → list call returns isError; catalog OK → normal
  filtering.

**Implementation notes:**

- The current code path is at
  [src/lib/sensitivity-filter.ts:241-255](src/lib/sensitivity-filter.ts#L241-L255).
  Replace the `return { filtered: items, removedCount: 0 };` fallback with a
  thrown sentinel error type that `executeGraphTool` catches and converts to
  the structured isError result. Don't return all items.

---

## 10. [ ] Tighten secret-handling and stop logging request bodies on `/token`

**Why:** Guideline §7 — no secrets in logs.

**Files:**

- [src/server.ts](src/server.ts) — `/token` handler around lines 430-551.
- [src/lib/microsoft-auth.ts](src/lib/microsoft-auth.ts) — error-response
  logging at lines 117-120 / 172-175.

**Acceptance:**

- `/token` error path no longer includes `body` or `req.body` in any log
  call. Replace with a fixed-shape diagnostic that names which field was
  missing/invalid but never echoes any value other than `grant_type`.
- The `client_id` truncated print at startup
  ([src/server.ts:163-168](src/server.ts#L163-L168)) is removed (or moved to
  DEBUG level only).
- Microsoft error responses are logged with the body **redacted** beyond the
  first 200 characters and with any obvious token/code/refresh substrings
  filtered (regex strip of `(access_token|refresh_token|code|code_verifier|
  client_secret)=[^&\s"]+`).
- Add a unit test: feed a fake error body containing
  `refresh_token=abc.def.ghi` to the redactor, assert the value is gone.

---

## 11. [ ] Caller identity in every log line + structured logger

**Why:** Guideline §9 — log identity chain, tool calls, schema version.

**Files:**

- [src/logger.ts](src/logger.ts) — switch to JSON output and require a
  request-context binder.
- [src/request-context.ts](src/request-context.ts) — extend to carry
  `oid`, `tid`, `appid`, `requestId`.
- [src/server.ts](src/server.ts) — generate a per-request `requestId` (UUID),
  bind it into `requestContext` along with the validated claims from Task 1.
- Every existing `logger.info(...)` etc. is replaced with calls that pull
  context automatically.

**Acceptance:**

- Every log line emitted while inside an `/mcp` request includes
  `requestId`, `oid`, `tid`, `tool` (when known), and a SHA-256 prefix
  (first 12 hex chars) of the loaded `endpoints.json` content (compute once
  at startup, store in module scope).
- Log output is JSON with stable field names: `ts`, `level`, `msg`,
  `requestId`, `oid`, `tid`, `tool`, `endpointsHash`, plus any extra
  structured fields the call site provides.
- No PII beyond `oid`/`tid`/`upn` is logged automatically.

**Implementation notes:**

- Use `winston.format.json()` instead of the current `printf`.
- Wrap `logger` with a small helper:

  ```ts
  function ctxLog(level: string, msg: string, extra?: Record<string, unknown>) {
    const ctx = requestContext.getStore();
    logger.log(level, msg, { ...ctx, ...extra });
  }
  ```

- Where call sites pass `error` objects, serialize them with `name`,
  `message`, `stack`, and any `code` field — but never the raw object (it
  may contain captured request data).

---

## 12. [ ] Route logs to Azure Application Insights (App Service)

**Why:** Guideline §9 — centralized, immutable logs and tracing.

**Files:**

- [src/logger.ts](src/logger.ts) — add an optional Application Insights
  transport.
- [README.md](README.md) — document the App Service configuration.

**Acceptance:**

- When `APPLICATIONINSIGHTS_CONNECTION_STRING` is set (App Service standard
  env var), the logger also ships records to App Insights via
  `applicationinsights` SDK. Local file transports remain for local dev.
- Tool-call telemetry is sent as a custom event `tool_call` with properties
  `tool`, `oid`, `tid`, `requestId`, `durationMs`, `outcome` (`success` |
  `error` | `denied` | `approval_required`).
- Server starts without App Insights when the env var is missing (no hard
  failure for local dev).

**Implementation notes:**

- `npm i applicationinsights`. Pin the version. Initialize once in
  `src/index.ts` before any other module that creates a logger:

  ```ts
  if (process.env.APPLICATIONINSIGHTS_CONNECTION_STRING) {
    const appInsights = await import('applicationinsights');
    appInsights.setup().setAutoCollectConsole(false).start();
  }
  ```

- Add a simple Winston transport that calls
  `client.trackTrace`/`client.trackEvent` rather than depending on
  auto-collect — explicit is easier to reason about.

---

## 13. [ ] OpenTelemetry tracing across MCP → MSAL → Graph

**Why:** Guideline §9 — distributed tracing.

**Files:**

- New file: `src/lib/tracing.ts`.
- [src/index.ts](src/index.ts) — initialize tracing first.
- [src/graph-client.ts](src/graph-client.ts),
  [src/graph-tools.ts](src/graph-tools.ts) — wrap `executeGraphTool` and
  `performRequest` in spans.
- [src/server.ts](src/server.ts) — wrap `/mcp` and `/token` routes with
  HTTP instrumentation.

**Acceptance:**

- Spans exported via OTLP HTTP to the endpoint configured by
  `OTEL_EXPORTER_OTLP_ENDPOINT` (App Service standard).
- Span names: `mcp.request`, `mcp.tool_call` (with attribute `tool`),
  `graph.request` (with attribute `graph.endpoint`).
- Trace context is propagated into the App Insights events from Task 12 so
  a request can be followed end-to-end.

**Implementation notes:**

- Use `@azure/monitor-opentelemetry` for the easiest App Service
  integration; it wires both OpenTelemetry and App Insights automatically
  when `APPLICATIONINSIGHTS_CONNECTION_STRING` is set.

---

## 14. [ ] Rate limiting, quotas, circuit breakers

**Why:** Guideline §10 — economic controls.

**Files:**

- New file: `src/lib/rate-limit.ts`.
- [src/server.ts](src/server.ts) — apply middleware on `/mcp`, `/token`,
  `/register`, `/authorize`.

**Acceptance:**

- Per-`oid` rate limit: default 60 tool calls/minute (configurable via
  `MS365_MCP_RATE_LIMIT_PER_USER_PER_MIN`).
- Per-tenant rate limit: default 600 tool calls/minute
  (`MS365_MCP_RATE_LIMIT_PER_TENANT_PER_MIN`).
- Per-IP rate limit on unauthenticated endpoints (`/register`, `/authorize`,
  `/token`): default 30/min (`MS365_MCP_RATE_LIMIT_PER_IP_PER_MIN`).
- Circuit breaker: if `>10%` of Graph calls in the last 60 s return `429` or
  `5xx`, the server returns `503 service_temporarily_unavailable` for the
  next 30 s instead of forwarding new tool calls. Tunable via env vars.
- Limit responses include `Retry-After` and a stable error code.

**Implementation notes:**

- Use `express-rate-limit` for HTTP endpoints. Pin the version.
- Implement the per-`oid`/per-tenant tool-call limit as a module-level
  sliding-window counter; call it from `executeGraphTool` before any Graph
  traffic. Don't share with the HTTP-rate-limit middleware.
- App Service can run multiple instances (scale-out). The simple in-memory
  limiter will under-count globally. Document this in
  `docs/security-config.md`: for production at scale, configure
  `MS365_MCP_RATE_LIMIT_BACKEND=redis` (future task) and an Azure Cache for
  Redis. For the first iteration, in-memory per-instance is acceptable as
  long as App Service is configured with a small instance count and the
  limit per instance is set so total ≤ desired global cap.

---

## 15. [ ] Refuse to bind without TLS unless explicitly allowed

**Why:** Guideline §6 — TLS ≥1.2.

**Files:**

- [src/server.ts](src/server.ts) — early in `start()` for the HTTP branch.
- [src/cli.ts](src/cli.ts) — add `--insecure-http` flag for local dev only.

**Acceptance:**

- When `--http` and `MS365_MCP_PRODUCTION=1` (Task 2), the server requires:
  - either `WEBSITE_HOSTNAME` is set (App Service marker) **and** TLS is
    terminated upstream by App Service (we trust App Service to do this; no
    code-side TLS), **or**
  - `--insecure-http` is explicitly passed.
- Outside production mode the server still binds plain HTTP (current
  behaviour) so local dev is unaffected.
- If `--insecure-http` is passed in production mode, log a WARN line and
  continue, but record the fact in App Insights so it shows up on a
  dashboard.

**Implementation notes:**

- App Service always presents the request as HTTPS to clients, but the
  Node process receives HTTP. The `trust proxy` setting at
  [src/server.ts:178](src/server.ts#L178) is correct. Just verify
  `req.secure` after the proxy hop in any flow that constructs an absolute
  URL.

---

## 16. [ ] Move token cache out of the package directory by default

**Why:** Guideline §7 — secrets storage hygiene; prevent token loss across
deploys, prevent storage in deploy artefacts.

**Files:**

- [src/auth.ts](src/auth.ts) — change defaults.

**Acceptance:**

- When `WEBSITE_HOSTNAME` is set (App Service), the default token cache
  paths use `process.env.HOME/.ms-365-mcp-server/.token-cache.json` and
  `…/.selected-account.json` instead of the package-relative fallbacks.
- A required env var `MS365_MCP_TOKEN_CACHE_PATH` is still respected and
  takes precedence.
- When `MS365_MCP_PRODUCTION=1` and **HTTP mode**, refuse to start unless
  one of:
  - `MS365_MCP_TOKEN_CACHE_PATH` is set, **or**
  - keytar is available, **or**
  - `MS365_MCP_OAUTH_TOKEN` is set (BYOT — token cache not used).

  In HTTP/OAuth mode the token cache is barely used (tokens come from the
  bearer flow), so this gate mostly catches misconfigurations in stdio-based
  hybrid setups.

---

## 17. [ ] Tighten CORS allowlist

**Why:** Guideline §6 — CORS hygiene.

**Files:**

- [src/server.ts](src/server.ts) — CORS middleware (lines 183-199).

**Acceptance:**

- `MS365_MCP_CORS_ORIGIN` accepts a comma-separated list (existing single
  value still works).
- Reject `*` outright in production mode (Task 2). Log ERROR and exit.
- Reject any non-`https://` origin in production mode unless it is
  `http://localhost(:port)` (dev convenience).
- The CORS handler echoes back only the *exact* matching origin in
  `Access-Control-Allow-Origin`; never echoes the request's `Origin` if it
  is not in the allowlist.
- Add unit tests for: missing Origin, allowed Origin, disallowed Origin,
  wildcard rejected in production.

---

## 18. [ ] Reject malformed `redirect_uris` and forwarded auth params

**Why:** Defense-in-depth around the OAuth bridge.

**Files:**

- [src/server.ts](src/server.ts) — `/authorize` handler (lines 296-427) and
  `/register` handler (Task 3).
- [src/oauth-provider.ts](src/oauth-provider.ts) — `authorize()` override.

**Acceptance:**

- `/authorize` validates the incoming `redirect_uri`:
  - must parse as a `URL`,
  - protocol ∈ `{https:, http:}` (http only for `localhost` per Task 17),
  - host is in `MS365_MCP_ALLOWED_REDIRECT_URIS` (Task 3),
  - no fragment, no userinfo.
- `MS365_MCP_AddToAuthURL` (current behaviour at
  [src/server.ts:416-422](src/server.ts#L416-L422) and
  [src/oauth-provider.ts:102-108](src/oauth-provider.ts#L102-L108)) is
  validated against an allowlist of permissible param names — at minimum
  forbid `client_id`, `redirect_uri`, `code_challenge`,
  `code_challenge_method`, `state`, `response_type`, `scope`, `nonce` —
  basically anything that affects auth security.

**Implementation notes:**

- Move the allowlist into a constant at the top of `src/server.ts` so it is
  reviewable in one place.

---

## 19. [ ] Drop the `accounts[0]` fallback in `getCurrentAccount`

**Why:** Guideline §11 — multi-tenancy isolation. The code at
[src/auth.ts:418-426](src/auth.ts#L418-L426) flags itself as unsafe.

**Files:**

- [src/auth.ts](src/auth.ts) — `getCurrentAccount()` and any caller.

**Acceptance:**

- In multi-account mode (more than one cached account) and no
  `selectedAccountId`, `getCurrentAccount()` throws instead of returning
  `accounts[0]`.
- Single-account mode (`accounts.length === 1`) continues to auto-resolve.
- Audit every caller of `getCurrentAccount()` to ensure the caller catches
  the new error path. The main consumers are `getToken()`
  ([src/auth.ts:381-401](src/auth.ts#L381-L401)) — surface the error to the
  tool call as `unauthorized_tool_call` (Task 7).

---

## 20. [ ] Sign or hash-pin `endpoints.json`

**Why:** Guideline §3 — schema integrity.

**Files:**

- New file: `src/lib/endpoints-integrity.ts`.
- [src/auth.ts](src/auth.ts) and
  [src/graph-tools.ts](src/graph-tools.ts) — replace direct `readFileSync`
  with the verified loader.
- New file: `endpoints.json.sha256` (committed; updated by a build step).
- New script: `bin/seal-endpoints.mjs` that recomputes the hash file when
  `endpoints.json` changes.

**Acceptance:**

- At startup, the loader computes SHA-256 of `endpoints.json` content and
  compares it to `endpoints.json.sha256`.
- Mismatch → fail-closed: log ERROR, exit non-zero. Operators must rerun
  `node bin/seal-endpoints.mjs` and review the diff before deploying.
- The hash is also included in every log line (Task 11).
- Add a CI job (GitHub Actions) that runs `bin/seal-endpoints.mjs` and
  fails if the working tree changes — i.e. the committed hash and content
  must always be in sync.

---

# SHOULD-FIX (post-approval, follow-up release)

## 21. [ ] RFC 8693 token exchange (OBO) so Graph token ≠ MCP token

**Why:** Guideline §2 — never pass user tokens downstream.

**Sketch:** When the bearer middleware accepts a token at the MCP
boundary, exchange it for a narrower Graph access token using the
On-Behalf-Of flow. The Entra app must be configured for OBO with a
`client_secret` from Key Vault. Cache the OBO result with a short TTL.
Touchpoints:
[src/graph-client.ts:88-89](src/graph-client.ts#L88-L89),
[src/oauth-provider.ts:31-49](src/oauth-provider.ts#L31-L49),
[src/request-context.ts](src/request-context.ts).

## 22. [ ] DPoP-bound tokens

**Why:** Guideline §2 — proof of possession. Implement DPoP per RFC 9449
on the MCP boundary. Requires Entra app changes to issue DPoP-bound
tokens and a `DPoP` header validator in middleware. Larger change; pair
with Task 21.

## 23. [ ] Kill switch + admin endpoints

**Why:** Guideline §15.

**Sketch:** Add `POST /admin/quarantine` (admin-token authenticated, env
var `MS365_MCP_ADMIN_TOKEN`) that flips an in-memory flag. While the flag
is on, every tool call returns `503 server_quarantined`. Add `POST
/admin/unquarantine` to clear. Log every transition at WARN with the admin
caller identity. Document the runbook for App Service: set the env var via
App Service config, restart the app to clear if needed.

## 24. [ ] Egress restriction documentation

**Why:** Guideline §8.

**Sketch:** Add a `docs/network-egress.md` listing the exact hostnames per
Microsoft cloud (extract from
[src/cloud-config.ts](src/cloud-config.ts)) so an App Service VNet
integration can configure egress allowlist. No code change.

## 25. [ ] Prompt-injection scrubber on Graph response content

**Why:** Guideline §4 — separation of data vs instructions.

**Sketch:** New `src/lib/content-scrub.ts` that runs over Teams/email/
chat/calendar bodies before they're handed to the model, neutralizing
patterns like `<system>...`, "ignore previous instructions",
fenced-code-block instructions, and known jailbreak markers. Document
that this is best-effort, not a guarantee.

## 26. [ ] SBOM, lockfile-only installs, image/artifact signing

**Why:** Guideline §12.

**Sketch (Azure App Service, no Docker):**

- CI: run `npm ci` (not `npm i`) and `npm audit --audit-level=high
  --omit=dev` as a gate.
- Generate a CycloneDX SBOM with `cyclonedx-npm` and publish as a build
  artefact and a release asset.
- Pin top-level dependencies (drop `^` ranges in `package.json`); rely on
  Renovate / Dependabot for upgrades.
- App Service deploy: prefer ZIP-deploy from a signed artifact in an
  Azure Storage account with immutability policy. Use `az webapp config`
  to record the artifact SHA-256 in the App Service tags so audit can
  verify what is running.

## 27. [ ] Rate-limit storage in Redis for multi-instance correctness

**Why:** Tail of Task 14. When App Service scales out, in-memory
counters under-count global usage. Add an optional Redis-backed counter
behind `MS365_MCP_RATE_LIMIT_BACKEND=redis`.

## 28. [ ] Adversarial test suite

**Why:** Catch regressions for all of the above.

**Sketch:** Add `src/__tests__/adversarial.test.ts` covering: malformed
JWT, valid JWT wrong tenant, valid JWT wrong audience, expired JWT,
missing scopes, oversized request body, unknown body fields, wrong
redirect_uri at `/register`, policy file missing, sensitivity catalog
unavailable, rate-limit boundary, circuit-breaker boundary,
quarantine-on while tool called, malformed `redirect_uri` at
`/authorize`. Each test asserts a deterministic error code and that no
secret-shaped string appears in `logger`'s output.

---

# Acceptance & sign-off

When the **Must-fix** tasks are complete:

1. Re-run `npm run verify` (`generate && lint && format:check && build && test`).
2. Run the new adversarial suite.
3. Hand the security team:
   - This file with all Must-fix boxes ticked.
   - The new `docs/security-config.md` deployment doc.
   - The signed `endpoints.json.sha256`.
   - A sample App Service configuration (env var list, no values) checked
     in to `docs/appservice-env.example.txt`.
4. Demonstrate, against a deployed App Service test slot:
   - A foreign-tenant token is rejected with `401`.
   - A wrong-audience token is rejected with `401`.
   - An unknown-field request body is rejected with `400`.
   - `send-mail` cannot be invoked without explicit policy + approval.
   - A poisoned mail body's `<system>` content is neutralized (Task 25,
     when shipped).
   - The label-catalog-down case returns isError instead of leaking
     documents.
   - Rate-limit / circuit-breaker behaviour under synthetic load.

That demo plus the ticked checklist is the gate for production rollout.
