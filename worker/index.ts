import { betterAuth } from "better-auth";

export interface Env {
  DB: D1Database;
  ASSETS: Fetcher;
  BETTER_AUTH_SECRET?: string;
  BETTER_AUTH_URL?: string;
}

export default {
  async fetch(request: Request, env: Env): Promise<Response> {
    const url = new URL(request.url);

    if (url.pathname === "/api/health") {
      return Response.json({
        ok: true,
        service: "budgetwise-api",
      }, {
        headers: { "Cache-Control": "no-store" },
      });
    }

    if (url.pathname.startsWith("/api/auth/")) {
      if (!env.BETTER_AUTH_SECRET) {
        return Response.json(
          { error: "Authentication is not configured yet." },
          { status: 503, headers: { "Cache-Control": "no-store" } },
        );
      }

      const auth = betterAuth({
        database: env.DB,
        secret: env.BETTER_AUTH_SECRET,
        baseURL: env.BETTER_AUTH_URL || url.origin,
        trustedOrigins: [env.BETTER_AUTH_URL || url.origin],
        emailAndPassword: {
          enabled: true,
        },
        rateLimit: {
          enabled: true,
        },
      });

      return auth.handler(request);
    }

    return env.ASSETS.fetch(request);
  },
} satisfies ExportedHandler<Env>;
