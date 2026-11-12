const API_URL = "https://script.google.com/macros/s/AKfycbxGM_bn3jYFC4cEgqLS6hju8CCEcxyYS8Kte5q4aTJOh8LzUAuizHnLDrG2Htdv3yIn/exec";

function getStoredUser() {
  try {
    return JSON.parse(localStorage.getItem('persawi_user') || 'null');
  } catch (e) {
    return null;
  }
}

function getStoredToken() {
  return localStorage.getItem('persawi_token') || '';
}

function saveSession(user, token) {
  localStorage.setItem('persawi_user', JSON.stringify(user));
  localStorage.setItem('persawi_token', token || '');
}

function clearSession() {
  localStorage.removeItem('persawi_user');
  localStorage.removeItem('persawi_token');
}

async function apiGet(action, params) {
  const url = new URL(API_URL);
  url.searchParams.set('action', action);
  Object.entries(params || {}).forEach(([key, value]) => {
    if (value !== undefined && value !== null && value !== '') {
      url.searchParams.set(key, value);
    }
  });
  const res = await fetch(url.toString());
  return res.json();
}

async function apiPost(payload) {
  const res = await fetch(API_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(payload)
  });
  return res.json();
}

async function loadSiteConfig() {
  try {
    const result = await apiGet('getConfig');
    if (result.status === 'success') {
      const config = Object.assign({}, result.config);
      const maxMb = parseInt(config.MAX_FILE_MB, 10);
      config.MAX_FILE_MB = String(Math.min(Number.isFinite(maxMb) && maxMb > 0 ? maxMb : 25, 25));
      return config;
    }
  } catch (e) {
    console.error("Gagal memuat config:", e);
  }
  return {
    EVENT_NAME: "PERSAWI 2026",
    EVENT_SUBTITLE: "Perkemahan Santri Wahdah Islamiyah",
    EVENT_SLOGAN: "Tunas Muda Berjiwa Rabbani Menuju Luwu Timur Juara",
    MAX_FILE_MB: "25"
  };
}

function escapeHtml(value) {
  return String(value == null ? '' : value).replace(/[&<>"']/g, function (char) {
    return {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#39;'
    }[char];
  });
}

function safeUrl(value) {
  const url = String(value || '');
  return /^https:\/\//i.test(url) ? url : '#';
}