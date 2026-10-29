// Reservoir service worker. It only shows push notifications and opens the app when one is tapped. It caches nothing.
self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", (e) => e.waitUntil(self.clients.claim()));
self.addEventListener("push", (e) => {
  let d = {};
  try { d = e.data ? e.data.json() : {}; } catch (err) { d = {}; }
  e.waitUntil(self.registration.showNotification(d.title || "Reservoir", {
    body: d.body || "Open Reservoir",
    tag: d.tag || undefined,
    icon: "icon-192.png",
    badge: "icon-192.png",
    data: { open: d.open || "notices" }
  }));
});
self.addEventListener("notificationclick", (e) => {
  e.notification.close();
  const open = e.notification.data && e.notification.data.open || "notices";
  e.waitUntil((async () => {
    const all = await self.clients.matchAll({ type: "window", includeUncontrolled: true });
    for (const c of all) {
      try { await c.focus(); c.postMessage({ type: "push-open", open }); return; } catch (err) {}
    }
    await self.clients.openWindow("./?open=" + open);
  })());
});
