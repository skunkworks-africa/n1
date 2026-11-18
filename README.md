# N1 project workspace

A public GitHub Pages interface for a private **basic Microsoft Planner plan**. The initial release is in setup mode: no plan, tenant, application or membership is inferred. No project tasks, meetings, participant lists or credentials belong in this repository.

## Features

Microsoft sign-in using MSAL Browser 5 and authorization code with PKCE; responsive light/dark board and list; search and task filters; task creation, editing and deletion; progress, priority, dates, bucket, assignments, existing plan labels, description, checklist and links. Comments and file uploads open in Microsoft Planner. Tasks refresh every 60 seconds while the visible board is open and can be refreshed manually; this is polling, not push synchronization.

## Security boundary

- GitHub serves only static public assets. Microsoft Graph validates delegated tokens and enforces access to the plan on each request. The public interface and its configuration are not a security boundary.
- A dedicated private Microsoft 365 group must contain only approved N1 members and explicitly invited guests. Do not use a public group or a broad company group unless every member is approved for this project.
- The application additionally verifies the configured plan's owning group, private group visibility and the signed-in user's direct membership before loading or changing tasks. Task IDs cannot be used through the client to switch to another plan. These client checks supplement Microsoft's server authorization; changing browser code does not grant Microsoft permissions.
- No application permissions, shared service identity, client secret or backend task cache. Delegated `Tasks.ReadWrite` is **not an app-specific single-plan scope**: it can cover other tasks the user is authorized to access. If the application itself must be restricted to a single plan even under modified client code, use a separately hosted API with delegated on-behalf-of access and a server-side plan allowlist; that is a different deployment.
- Tokens and task data are held in memory, not local/session storage. Sign-out and page departure clear rendered private content. Static assets may be cached by the host; Graph calls use `cache: no-store`. No telemetry, service worker or third-party content CDN.
- Task fields and details use their own original ETags. Conflict responses are shown without blind retries. Multi-request saves can partially succeed; the UI reports this and requires reopening the task. A created task is remembered before saving details so details failure cannot cause an immediate duplicate create.
- Public HTML uses a restrictive CSP, text-only rendering for task content and a no-referrer policy. GitHub Pages does not support arbitrary response headers; a CSP meta tag cannot implement `frame-ancestors`. No iframe protection or server security headers are claimed.

## Local development

Node 24 or a version supported by Vite 8 is required.

```sh
npm ci
npm test
npm run test:browser
npm run build
npm run dev
```

Browser checks use installed Microsoft Edge and synthetic responses for both MSAL and Graph. They do not prove real Microsoft authentication, consent or guest permissions. Production files never include those mocks. Screenshots and test reports are ignored.

## Connect the real plan

1. Confirm the owning Microsoft organization and an existing basic plan, or explicitly approve a new empty plan in a named private Microsoft 365 group. Confirm that the group's actual membership is the approved project roster. This site does not create a group, plan, invitations, assignments, dates or meeting tasks automatically.
2. Register a dedicated **single-tenant SPA** in the owning organization. Use its tenant-specific authority, including for B2B guests. Configure the exact SPA redirect URL `https://n1.skunkworks.africa/redirect.html`. For a GitHub Pages preview, separately register that site's exact `redirect.html` URL. Do not use wildcards, implicit grants or a client secret. MSAL 5 uses the separate redirect bridge page included in the build.
3. Configure delegated Microsoft Graph permissions `User.Read`, `Tasks.ReadWrite`, and `GroupMember.Read.All`. The third permission supplies group visibility and the member picker; it does not modify groups or invite users. Obtain required consent from the owning organization's administrator. Current Microsoft documentation also lists `GroupMember.ReadBasic.All` for membership enumeration; evaluate reducing the scope only after testing the group visibility and user display-name responses in the actual tenant.
4. Require assignment on the enterprise application, and assign only the approved members or approved project group according to tenant licensing. App assignment is an additional restriction, not a replacement for plan membership. Guests require accepted invitations and membership in the plan's group; no invitations are sent by the app.
5. Populate `public/config.json` with the tenant, public application ID, plan ID, group ID and exact redirect URL. These identifiers are not secrets, but disclose only identifiers approved for public configuration. Set `enabled: true` only for the approved environment. A missing or mismatched configuration leaves sign-in disabled.
6. Verify live login, plan loading, create/update/read-back/delete of an explicitly approved test task, conflict behavior, sign-out and access denial for a signed-in nonmember. Repeat with an explicitly approved guest. Check revoked membership and expired sessions. Do not claim personal Microsoft account support based on successful authentication alone.

Direct consumer Microsoft accounts are not supported by the Planner Graph endpoints. An external work account can use an invited B2B guest identity subject to tenant settings. A personal account invited as a guest is a distinct scenario and remains unverified until tested in the host tenant. Premium plans are not available through this Planner API.

## Deploy

Build output alone is published from the repository's `gh-pages` branch. Source stays on `main`. Before publishing, inspect `dist/config.json` and ensure the output contains no private project content. Add `.nojekyll` to the output. Do not copy local `.private`, test reports, fixtures, environment files or Microsoft exports into the published branch.

For the requested custom domain, first claim `n1.skunkworks.africa` in this repository's Pages settings, then add the DNS CNAME `n1` → `skunkworks-africa.github.io` at the authoritative DNS provider. Verify DNS and the Pages certificate, then enforce HTTPS. Configure the matching SPA redirect URL. A CNAME in source alone does not establish domain ownership or prove working DNS/TLS.

## Primary references

- [Planner API: plan authorization, basic plans and version conflicts](https://learn.microsoft.com/en-us/graph/api/resources/planner-overview)
- [Update task: delegated permissions and ETags](https://learn.microsoft.com/en-us/graph/api/plannertask-update?view=graph-rest-1.0)
- [Update task details](https://learn.microsoft.com/en-us/graph/api/plannertaskdetails-update?view=graph-rest-1.0)
- [Guest access and group membership](https://support.microsoft.com/en-us/planner/guest-access-in-microsoft-planner)
- [MSAL Browser](https://learn.microsoft.com/en-us/entra/msal/javascript/browser/about-msal-browser)
- [MSAL 5 redirect bridge](https://learn.microsoft.com/en-us/entra/msal/javascript/browser/redirect-bridge)
- [GitHub Pages custom domain setup](https://docs.github.com/en/pages/configuring-a-custom-domain-for-your-github-pages-site/managing-a-custom-domain-for-your-github-pages-site)
