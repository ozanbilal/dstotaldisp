/**
 * GeoDisp Auth Gatekeeper Server
 *
 * Express server that protects all static files behind SSO authentication.
 * Replaces nginx static serving with authenticated access.
 */

const express = require("express");
const crypto = require("crypto");
const path = require("path");
const https = require("https");
const http = require("http");
const jwt = require("jsonwebtoken");

const app = express();
const PORT = process.env.PORT || 8080;

// ── Configuration ──
const GEOPROJE_API = process.env.GEOPROJE_API || "https://api.geoproje.com.tr";
const GEOPROJE_APP = process.env.GEOPROJE_APP || "https://app.geoproje.com.tr";
const JWT_SECRET = process.env.JWT_SECRET || process.env.SECRET_KEY;
const SERVICE_SECRET_KEY = process.env.SERVICE_SECRET_KEY || "";
const SERVICE_SLUG = "geodisp";
const IS_DEV = process.env.NODE_ENV !== "production";

// Session cache
const authCache = new Map();
const CACHE_TTL_MS = 5 * 60 * 1000;

setInterval(() => {
  const now = Date.now();
  for (const [key, entry] of authCache) {
    if (now > entry.expires) authCache.delete(key);
  }
}, 10 * 60 * 1000);

/**
 * Validate SSO cookie — local JWT decode first, API fallback.
 */
async function validateAuth(cookieHeader) {
  if (!cookieHeader) return { valid: false, claims: null };

  const cached = authCache.get(cookieHeader);
  if (cached && Date.now() < cached.expires) {
    return { valid: cached.valid, claims: cached.claims };
  }

  if (JWT_SECRET) {
    try {
      const cookies = Object.fromEntries(
        cookieHeader.split(";").map((c) => {
          const [k, ...v] = c.trim().split("=");
          return [k, v.join("=")];
        })
      );
      const token = cookies["access_token"];
      if (token) {
        const decoded = jwt.verify(token, JWT_SECRET);
        const result = {
          valid: true,
          claims: {
            user_id: parseInt(decoded.sub, 10),
            email: decoded.email || "",
            plan: decoded.plan || "free",
            tier: decoded.tier || 0,
            ops_left: decoded.ops_left != null ? decoded.ops_left : 3,
          },
        };
        authCache.set(cookieHeader, { ...result, expires: Date.now() + CACHE_TTL_MS });
        return result;
      }
    } catch { /* fall through to API */ }
  }

  // API fallback
  try {
    const result = await new Promise((resolve) => {
      const url = new URL(`${GEOPROJE_API}/api/auth/me`);
      const transport = url.protocol === "https:" ? https : http;
      const req = transport.request(
        {
          hostname: url.hostname,
          port: url.port || (url.protocol === "https:" ? 443 : 80),
          path: url.pathname,
          method: "GET",
          headers: { Cookie: cookieHeader },
          timeout: 5000,
        },
        (res) => {
          let body = "";
          res.on("data", (chunk) => (body += chunk));
          res.on("end", () => {
            if (res.statusCode === 200) {
              try {
                const user = JSON.parse(body);
                resolve({ valid: true, claims: { user_id: user.id, email: user.email, plan: "free", tier: 0 } });
              } catch { resolve({ valid: true, claims: null }); }
            } else {
              resolve({ valid: false, claims: null });
            }
          });
        }
      );
      req.on("error", () => resolve({ valid: false, claims: null }));
      req.on("timeout", () => { req.destroy(); resolve({ valid: false, claims: null }); });
      req.end();
    });
    authCache.set(cookieHeader, { ...result, expires: Date.now() + CACHE_TTL_MS });
    return result;
  } catch {
    return { valid: false, claims: null };
  }
}

// ── Auth Middleware ──
app.use(async (req, res, next) => {
  if (IS_DEV) return next();
  if (req.path === "/health") return res.json({ status: "ok", service: SERVICE_SLUG });

  const cookieHeader = req.headers.cookie;
  const { valid, claims } = await validateAuth(cookieHeader);

  if (valid) {
    req._claims = claims;
    return next();
  }

  const returnUrl = encodeURIComponent(`${req.protocol}://${req.get("host")}${req.originalUrl}`);
  res.redirect(`${GEOPROJE_APP}/login?redirect=${returnUrl}`);
});

// ── Tier info endpoint ──
app.get("/api/tier", (req, res) => {
  const claims = req._claims || {};
  res.json({ tier: claims.tier || 0, plan: claims.plan || "free" });
});

function sendPythonFile(res, filePath) {
  res.type("text/x-python");
  res.sendFile(filePath, (error) => {
    if (error && !res.headersSent) {
      res.status(error.statusCode || 404).type("text/plain").send("Python module not found");
    }
  });
}

app.get(["/disp_core.py", "/web-ui/disp_core.py"], (req, res) => {
  sendPythonFile(res, path.join(__dirname, "disp_core.py"));
});

// ── Static File Serving ──
app.use(express.static(__dirname, {
  maxAge: IS_DEV ? 0 : "1h",
  etag: true,
  index: "index.html",
}));
app.use("/web-ui", express.static(__dirname, {
  maxAge: IS_DEV ? 0 : "1h",
  etag: true,
  index: "index.html",
}));

// Serve Python/WASM files
app.use("/py", express.static(path.join(__dirname, "py"), {
  maxAge: IS_DEV ? 0 : "7d",
  etag: true,
}));
app.use("/vendor", express.static(path.join(__dirname, "vendor"), {
  maxAge: IS_DEV ? 0 : "7d",
  etag: true,
}));

// SPA fallback
app.get("*", (req, res) => {
  if (path.extname(req.path)) {
    return res.status(404).type("text/plain").send("Not found");
  }
  res.sendFile(path.join(__dirname, "index.html"));
});

app.listen(PORT, () => {
  console.log(`[GEODISP] Auth gatekeeper running on port ${PORT}`);
  console.log(`[GEODISP] Mode: ${IS_DEV ? "DEVELOPMENT (auth bypassed)" : "PRODUCTION"}`);
});
