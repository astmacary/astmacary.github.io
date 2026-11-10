/* Serve Team Board notification controls. */
function createPushClient({ appId, apiBaseUrl, getToken, onChange, isDemo, standalone, ios }) {
  let key = "", enabled = false, busy = false, error = "", ready = false, checked = false;
  const supported = () => window.isSecureContext && "serviceWorker" in navigator && "PushManager" in window && "Notification" in window;
  const escape = s => String(s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const base = apiBaseUrl.replace(/\/$/, "");
  async function request(path, body, authenticated = true) {
    const headers = { "Content-Type": "application/json" };
    if (authenticated) headers.Authorization = "Bearer " + await getToken();
    const res = await fetch(base + path, { method: body ? "POST" : "GET", headers, body: body ? JSON.stringify({ appId, ...body }) : undefined, cache: "no-store", signal: AbortSignal.timeout(20000) });
    if (!res.ok) throw new Error(res.status === 401 ? "Sign in again, then retry notifications." : "The notification service couldn't be reached. Please retry.");
    return res.status === 204 ? null : res.json();
  }
  async function registration() {
    const reg = await navigator.serviceWorker.register("/sw.js", { scope: "/", updateViaCache: "none" });
    if (reg.active) return reg;
    return new Promise((resolve, reject) => {
      const watched = new Set();
      const cleanup = () => {
        clearTimeout(timer); reg.removeEventListener("updatefound", check);
        for (const worker of watched) worker.removeEventListener("statechange", check);
      };
      const check = () => {
        const worker = reg.installing || reg.waiting;
        if (reg.active || worker?.state === "activated") { cleanup(); resolve(reg); return; }
        if (worker && !watched.has(worker)) { watched.add(worker); worker.addEventListener("statechange", check); }
      };
      const timer = setTimeout(() => { cleanup(); reject(new Error("The app is still updating. Close and reopen it, then retry.")); }, 15000);
      reg.addEventListener("updatefound", check); check();
    });
  }
  function publicKey(value) {
    const s = atob(value.replace(/-/g, "+").replace(/_/g, "/") + "=".repeat((4 - value.length % 4) % 4));
    return Uint8Array.from(s, c => c.charCodeAt(0));
  }
  async function init() {
    if (busy || isDemo() || !base || !supported() || (ios() && !standalone())) return;
    busy = true; error = "";   // no re-render here: the panel stays hidden while this first check runs
    try {
      const cfg = await request("/push/config?appId=" + encodeURIComponent(appId), null, false);
      key = cfg.publicKey; if (!key) throw new Error("Notifications aren't configured yet.");
      ready = true;
      const reg = await registration();
      let sub = await reg.pushManager.getSubscription();
      if (sub && Notification.permission === "granted") {
        const currentKey = sub.options.applicationServerKey;
        if (currentKey && btoa(String.fromCharCode(...new Uint8Array(currentKey))).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "") !== key) {
          await sub.unsubscribe();
          sub = null;
        }
      }
      // Notifications may already be granted from elsewhere in the app (the App Badge prompt grants
      // the same browser permission). When so, subscribe quietly here instead of asking again —
      // granting permission a second time isn't needed, so there's nothing to show the person first.
      if (!sub && Notification.permission === "granted") {
        sub = await reg.pushManager.subscribe({ userVisibleOnly: true, applicationServerKey: publicKey(key) });
      }
      if (sub && Notification.permission === "granted") {
        await request("/push/subscriptions", { action: "upsert", subscription: sub.toJSON() });
        enabled = true;
      }
    } catch (e) { error = e.message; enabled = false; }
    finally { busy = false; checked = true; onChange(); }
  }
  async function enable() {
    if (busy || isDemo()) return;
    // Call permission directly in the tap handler, before any network or worker waits.
    if (!ready) { await init(); return; }
    busy = true; error = "";
    try {
      const permission = await Notification.requestPermission();
      if (permission !== "granted") throw new Error(permission === "denied" ? "Notifications are off. Allow them in your device Settings." : "Notifications weren't enabled. You can try again.");
      onChange();
      const reg = await registration();
      const sub = await reg.pushManager.getSubscription() || await reg.pushManager.subscribe({ userVisibleOnly: true, applicationServerKey: publicKey(key) });
      await request("/push/subscriptions", { action: "upsert", subscription: sub.toJSON() });
      enabled = true;
    } catch (e) { error = e.message; enabled = false; }
    finally { busy = false; onChange(); }
  }
  async function disable() {
    if (!supported()) return;
    busy = true; error = ""; onChange();
    try {
      const reg = await navigator.serviceWorker.getRegistration(new URL("/", location.href));
      const sub = reg && await reg.pushManager.getSubscription();
      if (sub) {
        if (base) { try { await request("/push/subscriptions", { action: "remove", endpoint: sub.endpoint }); } catch (_) { /* Unsubscribe still invalidates this phone's endpoint. */ } }
        if (!await sub.unsubscribe() && await reg.pushManager.getSubscription()) throw new Error("Couldn't turn off notifications. Please retry.");
      }
      enabled = false;
    } catch (e) { error = e.message; throw e; }
    finally { busy = false; onChange(); }
  }
  function panel() {
    if (isDemo()) return "";
    const shell = text => `<section class="panel pad"><h2 class="section-label" style="margin:0">Team Board Notifications</h2>${text}</section>`;
    if (ios() && !standalone()) return shell("<p>To receive notifications on iPhone or iPad (iOS 16.4+), add this app to your Home Screen and open it from its icon.</p>");
    if (!supported()) return shell("<p>This browser doesn't support push notifications. Open the app in a supported browser.</p>");
    if (!base) return shell("<p>Lock-screen notifications are awaiting activation by your team administrator.</p>");
    if (Notification.permission === "denied") return shell("<p>Notifications are off. Allow notifications for this app in your device Settings, then reopen it.</p>");
    if (enabled) return shell(`<p class="hint">Team Board alerts are on for this device.</p><button class="btn" data-act="push-off" ${busy ? "disabled" : ""}>Turn Off Alerts</button>`);
    if (!checked) return "";   // still checking this phone: show nothing until we know notifications are off
    return shell(`<p>Get alerts about new RCI Leadership Team Board posts.</p>${error ? `<p role="alert">${escape(error)}</p>` : ""}<button class="btn primary" data-act="push-on" ${busy ? "disabled" : ""}>${busy ? "Please wait…" : ready ? "Enable Notifications" : "Retry Connection"}</button>`);
  }
  return { init, enable, disable, panel };
}

