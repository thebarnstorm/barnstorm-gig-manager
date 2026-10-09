// firebase-messaging-sw.js
// This file MUST live at the root of your site (same level as index.html)
// so its scope covers the whole app. Do not move it into a subfolder.
//
// This handles push notifications that arrive while the app is closed
// or in the background. Foreground notifications (app open) are handled
// separately in index.html.

importScripts('https://www.gstatic.com/firebasejs/10.8.0/firebase-app-compat.js');
importScripts('https://www.gstatic.com/firebasejs/10.8.0/firebase-messaging-compat.js');

// Same config as in index.html - safe to duplicate here, these are public client identifiers, not secrets.
firebase.initializeApp({
    apiKey: "AIzaSyDhYBUzbpBBv40-X6SLASGikxoMUWcCU5Q",
    authDomain: "barnstorm-gig-manager.firebaseapp.com",
    projectId: "barnstorm-gig-manager",
    storageBucket: "barnstorm-gig-manager.firebasestorage.app",
    messagingSenderId: "819468934815",
    appId: "1:819468934815:web:9959740caf4bec736eccaa"
});

const messaging = firebase.messaging();

messaging.onBackgroundMessage((payload) => {
    console.log('Background push received:', payload);

    // We send data-only messages (see PushNotifications.gs) to avoid
    // duplicate notifications, so title/body come from payload.data.
    const title = payload.data?.title || 'The Barnstorm';
    const options = {
        body: payload.data?.body || '',
        icon: 'https://www.dropbox.com/scl/fi/05pdxbaub02h1qa6pk9dp/Barnstorm-Logo.jpg?rlkey=8fghkclk4vvz0nxjkr2dninl5&st=1q1vn1b1&raw=1',
        badge: 'https://www.dropbox.com/scl/fi/05pdxbaub02h1qa6pk9dp/Barnstorm-Logo.jpg?rlkey=8fghkclk4vvz0nxjkr2dninl5&st=1q1vn1b1&raw=1',
        data: payload.data || {}
    };

    self.registration.showNotification(title, options);
});

// When a user taps a notification, bring them to the app
// (and to a specific tab/conversation if the push says so, e.g.
// data.view = 'messages', data.phone = '+15551234567', data.channel = 'sms')
self.addEventListener('notificationclick', (event) => {
    event.notification.close();
    const d = event.notification.data || {};
    const view = d.view || '';
    event.waitUntil(
        clients.matchAll({ type: 'window', includeUncontrolled: true }).then((clientList) => {
            for (const client of clientList) {
                if (client.url.includes('barnstorm-gig-manager') && 'focus' in client) {
                    if (view) client.postMessage({ type: 'openView', view: view, phone: d.phone || '', channel: d.channel || '' });
                    return client.focus();
                }
            }
            if (clients.openWindow) {
                let url = '/barnstorm-gig-manager/';
                if (view) {
                    const q = new URLSearchParams({ view: view });
                    if (d.phone) q.set('phone', d.phone);
                    if (d.channel) q.set('channel', d.channel);
                    url += '?' + q.toString();
                }
                return clients.openWindow(url);
            }
        })
    );
});
