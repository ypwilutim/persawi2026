const API_URL = "https://script.google.com/macros/s/AKfycbwEng0rBpT0amPRjiBosCImD2NZrI6kVsInzPwJk-cqMnTURKc-jg4DcIhMZOuZCvCI/exec";

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
    headers: { 'Content-Type': 'text/plain' },
    body: JSON.stringify(payload)
  });
  return res.json();
}

const DEFAULT_SITE_CONFIG = {
  EVENT_NAME: "PERSAWI 2026",
  EVENT_SUBTITLE: "Perkemahan Santri Wahdah Islamiyah",
  EVENT_SLOGAN: "Tunas Muda Berjiwa Rabbani Menuju Luwu Timur Juara",
  SEKOLAH: "SDIT Insan Rabbani Malili",
  MAX_FILE_MB: "25",
  MAX_R2_MB: "2048"
};

async function loadSiteConfig() {
  let config = {};
  try {
    const result = await apiGet('getConfig');
    if (result.status === 'success') {
      config = result.config || {};
    }
  } catch (e) {
    console.error("Gagal memuat config:", e);
  }

  const merged = Object.assign({}, DEFAULT_SITE_CONFIG, config);
  const maxMb = parseInt(merged.MAX_FILE_MB, 10);
  merged.MAX_FILE_MB = String(Math.min(Number.isFinite(maxMb) && maxMb > 0 ? maxMb : 25, 25));
  const maxR2Mb = parseInt(merged.MAX_R2_MB, 10);
  merged.MAX_R2_MB = String(Math.min(Number.isFinite(maxR2Mb) && maxR2Mb > 0 ? maxR2Mb : 2048, 5120));
  return merged;
}

function parseSekolah(value) {
  return String(value == null ? '' : value)
    .split(/[\n,]+/)
    .map(item => item.trim())
    .filter(Boolean);
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

const HARI_ID = ['Min', 'Sen', 'Sel', 'Rab', 'Kam', 'Jum', 'Sab'];
const BULAN_ID = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];

function formatTanggalID(tanggal) {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(tanggal || ''));
  if (!match) return tanggal || '-';
  const date = new Date(String(tanggal) + 'T00:00:00Z');
  return HARI_ID[date.getUTCDay()] + ', ' + Number(match[3]) + ' ' + BULAN_ID[Number(match[2]) - 1] + ' ' + match[1];
}

function formatJam(jam) {
  return jam ? String(jam) + ' WITA' : '-';
}

function isTaskOverdue(tanggal, jam) {
  if (!tanggal || !jam) return false;
  const due = new Date(String(tanggal) + 'T' + String(jam) + ':00');
  return due.getTime() < Date.now();
}