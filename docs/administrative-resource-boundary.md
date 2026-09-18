# Administrative resource boundary

The current product decision is **single-user binding** when `settings.user_id` is configured. For every operation whose registry metadata requires `user_id`, a call carrying a different user is rejected before argument dispatch, secret lookup, credential creation, or Graph client construction. The rejection is deliberately value-free: tenant, client, and user identifiers are never returned by preflight or boundary errors.

This plugin does not infer tenant-wide access from application permissions. Planner operations are the explicit exception to the *user identifier requirement* because their endpoint metadata does not require `user_id`; that does not grant tenant-wide authorization, consent, or execution.

Preflight derives `tenant_id`, `client_id`, and `user_id` requirements from selected operation metadata. Application mode checks the scoped application secret; delegated mode does not read it. Delegated authentication remains not implemented, and all write operations remain non-executable.
