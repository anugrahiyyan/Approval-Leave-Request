export default {
    async fetch(request, env) {
        const url = new URL(request.url);

        // Maintenance Mode Check
        // If MAINTENANCE_MODE is set to "true" in wrangler.toml or Cloudflare Dashboard,
        // redirect everything except the maintenance page itself and static assets.
        if (env.MAINTENANCE_MODE === "true") {
            const isStatic = /\.(css|js|gif|jpg|jpeg|png|svg|ico|woff|woff2|ttf|eot|mp4|webm)$/i.test(url.pathname);
            const isMaintenancePage = url.pathname === "/down/maintenance.html";
            const isExempt = url.pathname === "/privacy.html" || url.pathname === "/terms.html";

            if (!isStatic && !isMaintenancePage && !isExempt) {
                return Response.redirect(`${url.origin}/down/maintenance.html`, 302);
            }
        }

        // Otherwise, serve from the static bucket (Pages)
        return env.ASSETS.fetch(request);
    },
};
