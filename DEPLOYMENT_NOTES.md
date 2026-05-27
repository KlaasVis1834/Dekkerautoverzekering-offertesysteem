# Deployment Notes

## Render

- A push to GitHub triggers a Render deploy.
- New code is live only after Render shows `Deploy finished successfully`.
- If the portal still shows old behavior before the deploy is finished, wait for the active deploy to complete and refresh the page.

## Cloudflare

Add a Cache Rule:

- If `Hostname equals portaal.klaasvis.nl`
- Then `Bypass cache`
- Place the rule at `First`

Disable any Page Rules that use `Cache Everything` for the portal hostname.

After changing these rules, run `Purge Everything` once. After that, normal GitHub to Render deploys should become visible without repeated Cloudflare purges.
