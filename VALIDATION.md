# Validation — 18 November 2026

The public site is a **setup release**, not an activated live project plan.

- 13 service/client tests passed: configuration binding, authorization denial, private-group requirement, membership requirement, plan ownership, task ownership, pagination destination protection, original task/detail ETags, conflict handling, validation and access revocation before writes.
- 6 Microsoft Edge browser scenarios passed with synthetic Microsoft authentication and Graph responses: disabled public entry, denied access, list/search/edit/save/sign-out, conflict without retry, partial-create recovery and board accessibility. Public entry was checked at desktop and 390px mobile width.
- Production build succeeded. The deployed HTTPS page returned 200, loaded without browser errors or failed responses, kept sign-in disabled and the private workspace hidden, and exposed only empty configuration identifiers.
- Automated axe WCAG A/AA checks reported no violations on the deployed public page and synthetic board. This is not a manual screen-reader conformance claim.
- GitHub Pages build/deployment and the default CodeQL scan succeeded. The code-scanning alert list was empty when checked.
- `npm audit --omit=dev` reported no runtime dependency vulnerabilities. Full audit reports four high-severity development-tool findings originating from the same `picomatch` regular-expression denial-of-service advisory in Vite's dependency tree. The current installed Vite release still includes the affected package. Vite is used only for local development/build, binds to loopback, and is absent from the published runtime. Do not expose the development server or process untrusted glob patterns; recheck the advisory before build-tool upgrades. These findings have not been represented as resolved.

No real Microsoft sign-in, guest session, consent, plan read or task mutation was tested through this application. Before enabling it, complete the deployment steps in README and verify approved member, denied nonmember, guest and task read-back scenarios against the confirmed plan. No sample tasks are included in the deployed site.
