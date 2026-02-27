import express from "express";
import jwt from "jsonwebtoken";
import bcrypt from "bcryptjs";

const router = express.Router();
const COOKIE_NAME = "tn_token";

function getJwtSecret() {
  const secret = process.env.JWT_SECRET;
  if (!secret) throw new Error("JWT_SECRET is missing in environment variables");
  return secret;
}

function getAdminEmail() {
  return (process.env.ADMIN_EMAIL || "admin@troyer.local").trim().toLowerCase();
}

async function verifyPassword(plain) {
  const hash = process.env.ADMIN_PASSWORD_HASH;
  const fallback = process.env.ADMIN_PASSWORD;

  if (hash) return bcrypt.compare(String(plain || ""), hash);
  return String(plain || "") === String(fallback || "");
}

function cookieOptions() {
  const isProd = process.env.NODE_ENV === "production";
  return {
    httpOnly: true,
    secure: isProd,
    sameSite: "lax",
    path: "/",
    maxAge: 7 * 24 * 60 * 60 * 1000,
  };
}

function signToken(email) {
  return jwt.sign({ sub: "admin", email, role: "admin" }, getJwtSecret(), {
    expiresIn: "7d",
  });
}

router.post("/login", async (req, res) => {
  try {
    const { email, password } = req.body || {};
    const e = String(email || "").trim().toLowerCase();
    const p = String(password || "");

    if (!e || !p) {
      return res.status(400).json({
        ok: false,
        code: "MISSING_FIELDS",
        error: "Email and password are required",
      });
    }

    if (e !== getAdminEmail()) {
      return res.status(401).json({
        ok: false,
        code: "EMAIL_WRONG",
        error: "Email is wrong",
      });
    }

    const ok = await verifyPassword(p);
    if (!ok) {
      return res.status(401).json({
        ok: false,
        code: "PASSWORD_WRONG",
        error: "Password is wrong",
      });
    }

    const token = signToken(e);
    res.cookie(COOKIE_NAME, token, cookieOptions());

    return res.json({ ok: true, user: { email: e, role: "admin" } });
  } catch (e) {
    return res.status(500).json({
      ok: false,
      code: "SERVER_ERROR",
      error: "Server error",
    });
  }
});

router.post("/logout", (req, res) => {
  res.clearCookie(COOKIE_NAME, { path: "/" });
  return res.json({ ok: true });
});

router.get("/me", (req, res) => {
  try {
    const token = req.cookies?.[COOKIE_NAME] || "";
    if (!token) {
      return res.status(401).json({
        ok: false,
        code: "UNAUTHORIZED",
        error: "Unauthorized",
      });
    }

    const payload = jwt.verify(token, getJwtSecret());
    return res.json({
      ok: true,
      user: { email: payload.email, role: payload.role },
    });
  } catch {
    return res.status(401).json({
      ok: false,
      code: "UNAUTHORIZED",
      error: "Unauthorized",
    });
  }
});

export function requireAuth(req, res, next) {
  try {
    if (req.method === "OPTIONS") return next();

    const cookieToken = req.cookies?.[COOKIE_NAME];

    const auth = String(req.headers.authorization || "");
    const bearer = auth.startsWith("Bearer ") ? auth.slice(7).trim() : "";

    const token = cookieToken || bearer;
    if (!token) {
      return res.status(401).json({
        ok: false,
        code: "UNAUTHORIZED",
        error: "Unauthorized",
      });
    }

    const payload = jwt.verify(token, getJwtSecret());
    req.user = payload;
    return next();
  } catch {
    return res.status(401).json({
      ok: false,
      code: "UNAUTHORIZED",
      error: "Unauthorized",
    });
  }
}

export default router;