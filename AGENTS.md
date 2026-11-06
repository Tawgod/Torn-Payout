# ROB / Torn City Projects — mandatory development instructions

Applies to ALL development, debugging, code review, code generation, configuration changes and deployment work for ROB, Chain Watch, the Companion, Banking, faction accounting, Armory, OC, war tools, and Torn-Payout.

## Before making a change
1. Read the canonical [Torn API Integration Policy](https://github.com/Tawgod/ironsides-faction-tools/blob/main/TORN_API_POLICY.md) and [Torn API Usage Registry](https://github.com/Tawgod/ironsides-faction-tools/blob/main/API_USAGE.md). Read them from `main` even when operating on a production, staging or feature branch, or on Torn-Payout. Do not assume branch-local copies are current.
2. Determine whether the task has direct or indirect Torn API impact (endpoints, keys, scopes, feature grants, privacy/storage/sharing, polling, retry/backfill, caching, API call comments, reused datasets or budgets). Record `No Torn API impact` when none; do not modify the registry unnecessarily.
3. For an API-impacting change, follow both documents and current official Torn API documentation. Use minimum scopes, official endpoints, efficient shared cached data, correctly labeled requests and responsible rate limits. Preserve the attack warehouse's completeness and Banking's prompt verification / duplicate-payment safety. Never automate prohibited gameplay.
4. Update the canonical `API_USAGE.md` on `main` in the same task with endpoints, feature owners, frequency, triggers, keys, retries, caching and API-budget effects. If keys or data-handling behavior change, update `TORN_API_POLICY.md` and all relevant user-facing disclosures.
5. Check source branch vs Railway live branch and test vs production. Validate changes and deployment/log evidence; use LIVE, TESTING, IN PROGRESS, BUG / VERIFY, PLANNED, BACKLOG or DEFERRED accurately. Discussion or a GitHub commit alone is **not** a deployment.

Security: never expose or log raw API keys; honor faction separation and member permissions. Never let a developer view override destructive Torn/API permissions or membership/ownership checks for contributed keys.

These requirements apply independently of which Torn project chat initiated the work. The API Management chat coordinates usage; it does not exempt other chats from following these rules.

**Limit of this file:** this is repository guidance for agents/tools that read it; it does not automatically load into every new ChatGPT conversation. Keep the Torn City Projects ChatGPT Project instructions aligned with this rule.
