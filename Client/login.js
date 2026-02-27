const API = "/api";

(function ensureToastUI() {
  if (document.getElementById("toast-css")) return;

  const style = document.createElement("style");
  style.id = "toast-css";
  style.textContent = `
    .toast-wrap{
      position:fixed;
      top:14px;
      left:50%;
      transform:translateX(-50%);
      z-index:99999;
      display:flex;
      flex-direction:column;
      gap:10px;
      width:min(720px, calc(100vw - 24px));
      pointer-events:none;
    }
    .toast{
      pointer-events:auto;
      display:flex;
      align-items:center;
      gap:12px;
      padding:12px 14px;
      background:#ffffff;
      border:1px solid rgba(15,23,42,.10);
      border-radius:12px;
      box-shadow:0 12px 30px rgba(15,23,42,.10);
      font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;
      color:#0f172a;
      opacity:0;
      transform:translateY(-8px);
      transition:opacity .16s ease, transform .16s ease;
    }
    .toast.show{ opacity:1; transform:translateY(0); }
    .toast .icon{
      width:26px;
      height:26px;
      border-radius:999px;
      display:grid;
      place-items:center;
      flex:0 0 auto;
    }
    .toast.success .icon{ background:rgba(16,185,129,.14); color:#059669; }
    .toast.error .icon{ background:rgba(244,63,94,.14); color:#e11d48; }
    .toast .msg{
      font-size:15px;
      line-height:1.25;
      flex:1 1 auto;
      white-space:nowrap;
      overflow:hidden;
      text-overflow:ellipsis;
    }
    .toast .x{
      border:none;
      background:transparent;
      color:rgba(15,23,42,.55);
      cursor:pointer;
      font-size:18px;
      line-height:1;
      padding:4px 6px;
      border-radius:8px;
      flex:0 0 auto;
    }
    .toast .x:hover{ background:rgba(15,23,42,.06); color:rgba(15,23,42,.78); }
    @media (max-width:520px){
      .toast .msg{ white-space:normal; }
    }
  `;
  document.head.appendChild(style);
})();

function getToastWrap() {
  let wrap = document.querySelector(".toast-wrap");
  if (!wrap) {
    wrap = document.createElement("div");
    wrap.className = "toast-wrap";
    document.body.appendChild(wrap);
  }
  return wrap;
}

function toast(message, type = "success", duration = 3000) {
  const wrap = getToastWrap();
  const el = document.createElement("div");
  el.className = `toast ${type}`;

  const icon = document.createElement("div");
  icon.className = "icon";
  icon.innerHTML =
    type === "success"
      ? `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden="true">
           <path d="M20 6L9 17l-5-5" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"/>
         </svg>`
      : `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden="true">
           <path d="M12 8v5" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"/>
           <path d="M12 16.8h.01" stroke="currentColor" stroke-width="3.2" stroke-linecap="round"/>
           <path d="M10.3 4.6h3.4L21 18.4c.8 1.5-.3 3.3-2 3.3H5c-1.7 0-2.8-1.8-2-3.3L10.3 4.6z" stroke="currentColor" stroke-width="2" stroke-linejoin="round"/>
         </svg>`;

  const msg = document.createElement("div");
  msg.className = "msg";
  msg.textContent = message || "";

  const close = document.createElement("button");
  close.className = "x";
  close.type = "button";
  close.setAttribute("aria-label", "Close");
  close.textContent = "×";

  el.appendChild(icon);
  el.appendChild(msg);
  el.appendChild(close);
  wrap.appendChild(el);

  const remove = () => {
    el.classList.remove("show");
    setTimeout(() => el.remove(), 180);
  };

  close.addEventListener("click", remove);
  requestAnimationFrame(() => el.classList.add("show"));

  let t = null;
  if (duration > 0) t = setTimeout(remove, duration);

  return {
    remove: () => {
      if (t) clearTimeout(t);
      remove();
    },
  };
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function setLoading(isLoading) {
  const btn = document.querySelector(".login-btn");
  if (!btn) return;
  btn.disabled = isLoading;
  btn.textContent = isLoading ? "Signing in..." : "Login";
}

async function apiFetch(url, options = {}) {
  const headers = new Headers(options.headers || {});
  if (options.body && !headers.has("Content-Type")) {
    headers.set("Content-Type", "application/json");
  }
  return fetch(url, { ...options, headers, credentials: "include" });
}

function mapLoginError(json, fallback) {
  const raw =
    (json && (json.error || json.message || json.msg)) ||
    fallback ||
    "Login failed";

  const s = String(raw).toLowerCase();

  if (s.includes("password")) return "Password is wrong.";
  if (s.includes("email")) return "Email is wrong.";
  if (s.includes("user") && (s.includes("not") || s.includes("no") || s.includes("found")))
    return "Email is wrong.";
  if (s.includes("invalid credentials") || s.includes("invalid")) return "Email or password is wrong.";

  return String(raw);
}

async function apiLogin(email, password) {
  const res = await apiFetch(`${API}/auth/login`, {
    method: "POST",
    body: JSON.stringify({ email, password }),
  });

  const json = await res.json().catch(() => ({}));

  if (!res.ok) {
    throw new Error(mapLoginError(json, "Login failed"));
  }

  return true;
}

document.addEventListener("DOMContentLoaded", () => {
  const form = document.getElementById("loginForm");
  const emailEl = document.getElementById("email");
  const passEl = document.getElementById("password");
  if (!form || !emailEl || !passEl) return;

  form.addEventListener("submit", async (e) => {
    e.preventDefault();

    const email = emailEl.value.trim();
    const password = passEl.value;

    if (!email || !password) {
      toast("Email and password are required.", "error", 3000);
      return;
    }

    if (!isValidEmail(email)) {
      toast("Enter a valid email address.", "error", 3000);
      emailEl.focus();
      return;
    }

    if (password.length < 6) {
      toast("Password must be at least 6 characters.", "error", 3000);
      passEl.focus();
      return;
    }

    if (password.length > 100) {
      toast("Password is too long.", "error", 3000);
      passEl.focus();
      return;
    }

    let pending = null;

    try {
      setLoading(true);
      pending = toast("Signing in...", "success", 0);

      await apiLogin(email, password);

      if (pending) pending.remove();
      toast("Login successful.", "success", 3000);

      setTimeout(() => {
        window.location.href = "/app";
      }, 250);
    } catch (err) {
      if (pending) pending.remove();
      toast(err?.message || "Login failed.", "error", 3000);
    } finally {
      setLoading(false);
    }
  });
});