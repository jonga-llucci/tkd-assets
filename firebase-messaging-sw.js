importScripts('https://www.gstatic.com/firebasejs/10.12.0/firebase-app-compat.js');
importScripts('https://www.gstatic.com/firebasejs/10.12.0/firebase-messaging-compat.js');

firebase.initializeApp({
  apiKey: "YOUR_API_KEY",
  authDomain: "YOUR_PROJECT.firebaseapp.com",
  projectId: "YOUR_PROJECT_ID",
  messagingSenderId: "YOUR_SENDER_ID",
  appId: "YOUR_APP_ID"
});

const messaging = firebase.messaging();

// Background notifications (app closed / in background)
messaging.onBackgroundMessage((payload) => {
  const { title, body, icon, data } = payload.notification;

  self.registration.showNotification(title, {
    body,
    icon: icon || '/icons/icon-192.png',
    badge: '/icons/badge-72.png',
    data: data || {},
    actions: [
      { action: 'open', title: '▶ Open App' },
      { action: 'dismiss', title: 'Dismiss' }
    ],
    vibrate: [200, 100, 200],
    tag: data?.tag || 'default',          // Collapse duplicate nudges
    renotify: true
  });
});

// Notification click handler
self.addEventListener('notificationclick', (event) => {
  event.notification.close();

  if (event.action === 'dismiss') return;

  const url = event.notification.data?.url || '/';
  event.waitUntil(
    clients.matchAll({ type: 'window', includeUncontrolled: true }).then(clientList => {
      const existingWindow = clientList.find(c => c.url.includes(self.location.origin));
      if (existingWindow) return existingWindow.focus();
      return clients.openWindow(url);
    })
  );
});
