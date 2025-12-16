// Keep-alive ping service
// This script runs periodically to keep the site "active"

const SITE_URL = 'https://campnominal.vercel.app/status.html';
const PING_INTERVAL = 5 * 60 * 1000; // 5 minutes

function pingSite() {
    fetch(SITE_URL)
        .then(response => {
            if (response.ok) {
                console.log('[Keep-Alive] Ping successful:', new Date().toISOString());
            } else {
                console.warn('[Keep-Alive] Ping failed with status:', response.status);
            }
        })
        .catch(error => {
            console.error('[Keep-Alive] Ping error:', error);
        });
}

// Start pinging after page loads
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => {
        setTimeout(() => {
            pingSite();
            setInterval(pingSite, PING_INTERVAL);
        }, 2000);
    });
} else {
    setTimeout(() => {
        pingSite();
        setInterval(pingSite, PING_INTERVAL);
    }, 2000);
}
