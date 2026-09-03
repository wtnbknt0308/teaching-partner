import React, { useState, useEffect, useRef, useMemo, useCallback } from "react";
import {
  Calendar, CalendarDays, LayoutGrid, BookOpen, Users, ClipboardList,
  Dumbbell, Settings, ChevronLeft, ChevronRight, Plus, Trash2, Check, X,
  Pencil, Clock, MapPin, Eraser, Save, Home, CalendarRange, Columns,
  Upload, Image as ImageIcon, Loader, Printer, Download, ShieldAlert, RotateCcw, FileText, Tag, Search, Maximize2, Minus, CheckSquare, Bell,
} from "lucide-react";
import mammoth from "mammoth";
import * as XLSX from "xlsx";

/* ============================================================
   最強・教員手帳  Teacher's Techo
   時間割 × 日課 × 授業計画 × 授業数集計 × 部活動 × 名簿 × 校務
   端末に自動保存（window.storage）。データはすべて編集可能。
   ============================================================ */

const STORE_KEY = "teacher-techo-v1";
const WD = ["日", "月", "火", "水", "木", "金", "土"]; // JS getDay index
const DAY_LABELS = ["月", "火", "水", "木", "金", "土"]; // timetable columns (idx 0..5)
const DAYMAP = { 月: 0, 火: 1, 水: 2, 木: 3, 金: 4, 土: 5, 日: -1 };

/* ---------- date helpers ---------- */
const pad = (n) => String(n).padStart(2, "0");
const ymd = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const parseYmd = (s) => { const [y, m, d] = s.split("-").map(Number); return new Date(y, m - 1, d); };
const addDays = (d, n) => { const x = new Date(d); x.setDate(x.getDate() + n); return x; };
const sameDay = (a, b) => ymd(a) === ymd(b);
const jsDayToIdx = (jsDay) => (jsDay === 0 ? -1 : jsDay - 1); // 月=0..土=5, 日=-1
const t2m = (t) => { if (!t) return 9999; const [h, m] = t.split(":").map(Number); return h * 60 + m; };
const uid = () => Math.random().toString(36).slice(2, 9);
const startOfWeekMon = (d) => { const x = new Date(d); const j = x.getDay(); const back = j === 0 ? 6 : j - 1; return addDays(x, -back); };

/* ---------- storage (オフライン優先: localStorage → window.storage → メモリ) ---------- */
const NS = "teacher-techo";
const APP_VERSION = "1.4.0";
const SCHEMA = 2;
const memStore = {};
let quotaHit = false; // 容量超過フラグ（Appが監視）
const lsGet = (k) => { try { const v = window.localStorage.getItem(k); return v == null ? undefined : v; } catch (e) { return undefined; } };
const lsSet = (k, v) => {
  try { window.localStorage.setItem(k, v); return true; }
  catch (e) { if (e && (e.name === "QuotaExceededError" || /quota/i.test(String(e)))) { quotaHit = true; try { window.__tpOnQuota && window.__tpOnQuota(); } catch (_) {} } return false; }
};
const lsDel = (k) => { try { window.localStorage.removeItem(k); } catch (e) {} };
const lsKeys = () => { try { return Object.keys(window.localStorage); } catch (e) { return []; } };

async function loadStore(key) {
  const ls = lsGet(key);
  if (ls !== undefined) { try { return JSON.parse(ls); } catch (e) {} }
  try { if (window.storage) { const r = await window.storage.get(key); if (r && r.value) return JSON.parse(r.value); } } catch (e) {}
  if (memStore[key] !== undefined) return memStore[key];
  return null;
}
async function saveStore(key, val) {
  const s = JSON.stringify(val);
  memStore[key] = val;
  lsSet(key, s);
  try { if (window.storage) await window.storage.set(key, s); } catch (e) {}
}
async function delStore(key) {
  delete memStore[key];
  lsDel(key);
  try { if (window.storage) await window.storage.delete(key); } catch (e) {}
}
const dataKeyOf = (name) => `${NS}:data:${name}`;
const MEETING_CATS = [
  { id: "全体", label: "全体連絡", color: "#3E9BC9" },
  { id: "生徒指導", label: "生徒指導", color: "#D9534F" },
  { id: "行事", label: "行事", color: "#E8845B" },
  { id: "研修", label: "研修", color: "#4F9E86" },
  { id: "分掌", label: "分掌・委員会", color: "#8E7CC3" },
  { id: "その他", label: "その他", color: "#8894A0" },
];
const catMeta = (id) => MEETING_CATS.find((c) => c.id === id) || MEETING_CATS[MEETING_CATS.length - 1];
const mtgImgKey = (name, id) => `mtgimg:${name}:${id}`;
/* 画像は IndexedDB（大容量）に保存。無い環境では localStorage にフォールバック。表示用に同期キャッシュを持つ。 */
let _idbP = null;
function idbOpen() {
  if (_idbP) return _idbP;
  _idbP = new Promise((res) => {
    try {
      if (!window.indexedDB) return res(null);
      const req = window.indexedDB.open("teaching-planners", 1);
      req.onupgradeneeded = () => { const db = req.result; if (!db.objectStoreNames.contains("images")) db.createObjectStore("images"); };
      req.onsuccess = () => res(req.result);
      req.onerror = () => res(null);
    } catch (e) { res(null); }
  });
  return _idbP;
}
async function idbPut(key, val) { const db = await idbOpen(); if (!db) return false; return new Promise((res) => { try { const tx = db.transaction("images", "readwrite"); tx.objectStore("images").put(val, key); tx.oncomplete = () => res(true); tx.onerror = () => res(false); } catch (e) { res(false); } }); }
async function idbGet(key) { const db = await idbOpen(); if (!db) return undefined; return new Promise((res) => { try { const tx = db.transaction("images", "readonly"); const r = tx.objectStore("images").get(key); r.onsuccess = () => res(r.result); r.onerror = () => res(undefined); } catch (e) { res(undefined); } }); }
async function idbDelKey(key) { const db = await idbOpen(); if (!db) return; try { const tx = db.transaction("images", "readwrite"); tx.objectStore("images").delete(key); } catch (e) {} }
async function idbAllKeys(prefix) { const db = await idbOpen(); if (!db) return []; return new Promise((res) => { try { const tx = db.transaction("images", "readonly"); const r = tx.objectStore("images").getAllKeys(); r.onsuccess = () => res((r.result || []).map(String).filter((k) => !prefix || k.startsWith(prefix))); r.onerror = () => res([]); } catch (e) { res([]); } }); }

const imgCache = {};
const imgCacheKey = (name, id) => `${name}::${id}`;
async function preloadMtgImgs(name, ids) {
  for (const id of ids) {
    const ck = imgCacheKey(name, id); if (imgCache[ck] !== undefined) continue;
    let v = await idbGet(mtgImgKey(name, id));
    if (v === undefined) { const ls = lsGet(mtgImgKey(name, id)); if (ls !== undefined) { try { v = JSON.parse(ls); } catch (e) { v = ls; } } }
    imgCache[ck] = v || null;
  }
}
const getMtgImg = (name, id) => (imgCache[imgCacheKey(name, id)] || null);
async function saveMtgImg(name, id, dataUrl) { imgCache[imgCacheKey(name, id)] = dataUrl; const ok = await idbPut(mtgImgKey(name, id), dataUrl); if (!ok) lsSet(mtgImgKey(name, id), JSON.stringify(dataUrl)); }
async function delMtgImg(name, id) { delete imgCache[imgCacheKey(name, id)]; await idbDelKey(mtgImgKey(name, id)); lsDel(mtgImgKey(name, id)); }
async function collectMeetingImages(name) {
  const out = {}; const pre = `mtgimg:${name}:`;
  const keys = await idbAllKeys(pre);
  for (const k of keys) { const v = await idbGet(k); if (v !== undefined) out[k.slice(pre.length)] = v; }
  lsKeys().forEach((k) => { if (k.startsWith(pre) && out[k.slice(pre.length)] === undefined) { const v = lsGet(k); if (v !== undefined) { try { out[k.slice(pre.length)] = JSON.parse(v); } catch (e) { out[k.slice(pre.length)] = v; } } } });
  return out;
}
async function restoreMeetingImages(name, imgs) { if (!imgs) return; for (const [id, url] of Object.entries(imgs)) { await saveMtgImg(name, id, url); } }
// 画像を縮小してJPEG dataURLに（容量対策）
function downscaleImage(file, maxDim = 1400, quality = 0.72) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => {
      const img = new Image();
      img.onload = () => {
        let { width: w, height: h } = img;
        if (w > maxDim || h > maxDim) { const s = maxDim / Math.max(w, h); w = Math.round(w * s); h = Math.round(h * s); }
        const c = document.createElement("canvas"); c.width = w; c.height = h;
        const ctx = c.getContext("2d"); ctx.drawImage(img, 0, 0, w, h);
        try { resolve(c.toDataURL("image/jpeg", quality)); } catch (e) { reject(e); }
      };
      img.onerror = reject; img.src = r.result;
    };
    r.onerror = reject; r.readAsDataURL(file);
  });
}
const sketchPrefix = (name) => `sketch:${name}:`;
// あるユーザーの手書きメモを全部集める（書き出し・移行用）
function collectSketches(name) {
  const out = {}; const pre = sketchPrefix(name);
  lsKeys().forEach((k) => { if (k.startsWith(pre)) { const v = lsGet(k); if (v !== undefined) { try { out[k.slice(pre.length)] = JSON.parse(v); } catch (e) { out[k.slice(pre.length)] = v; } } } });
  return out;
}
function restoreSketches(name, sketches) {
  if (!sketches) return; const pre = sketchPrefix(name);
  Object.entries(sketches).forEach(([dk, val]) => { lsSet(pre + dk, JSON.stringify(val)); });
}
// スキーマ移行：不足フィールドを既定値で補完（更新後も壊れない）
function migrate(data) {
  if (!data || typeof data !== "object") return defaultData();
  const base = defaultData();
  const d = { ...base, ...data };
  d.meta = { ...base.meta, ...(data.meta || {}) };
  d.club = { ...base.club, ...(data.club || {}) };
  d.club.days = (data.club && data.club.days) || {};
  d.club.schedule = (data.club && data.club.schedule) || {};
  d.club.specials = (data.club && data.club.specials) || [];
  d.textbook = { ...base.textbook, ...(data.textbook || {}) };
  d.textbook.units = (data.textbook && data.textbook.units) || [];
  ["events", "todos", "duties", "routine", "tests", "subjects", "classes", "periods", "meetings"].forEach((k) => { if (!Array.isArray(d[k])) d[k] = Array.isArray(base[k]) ? base[k] : []; });
  ["timetable", "lessonLog", "rosters", "weeklyManual", "dayMemo", "targets", "testProgress", "roadmap"].forEach((k) => { if (!d[k] || typeof d[k] !== "object") d[k] = {}; });
  if (!Array.isArray(d.terms) || !d.terms.length) { const fy = parseInt(String(d.meta && d.meta.year || "").replace(/\D/g, ""), 10) || new Date().getFullYear(); d.terms = [{ id: uid(), name: "1学期", start: `${fy}-04-01` }, { id: uid(), name: "2学期", start: `${fy}-09-01` }, { id: uid(), name: "3学期", start: `${fy + 1}-01-08` }]; }
  d.todos = (d.todos || []).map((t) => typeof t === "string" ? { id: uid(), text: t, done: false, cat: "その他" } : { cat: "その他", ...t });
  if (!d.meta.theme) d.meta.theme = "light";
  d.__v = SCHEMA;
  return d;
}

/* ---------- seed data ---------- */
function defaultData() {
  const today = new Date();
  const fyStart = today.getMonth() + 1 < 4 ? today.getFullYear() - 1 : today.getFullYear();
  const sat = addDays(startOfWeekMon(today), 5);
  const ev = (offset, title, type, time) => ({ id: uid(), date: ymd(addDays(today, offset)), title, type, time });
  return {
    meta: { teacher: "", year: `${today.getMonth() + 3 < 3 ? today.getFullYear() - 1 : today.getFullYear()}年度`, includeSat: false, fontScale: "M", contrast: false, theme: "light", onboarded: true },
    periods: [
      { label: "1", start: "08:50", end: "09:40" },
      { label: "2", start: "09:50", end: "10:40" },
      { label: "3", start: "10:50", end: "11:40" },
      { label: "4", start: "11:50", end: "12:40" },
      { label: "5", start: "13:30", end: "14:20" },
      { label: "6", start: "14:30", end: "15:20" },
    ],
    routine: [
      { id: uid(), time: "07:50", title: "出勤・教室準備" },
      { id: uid(), time: "08:15", title: "職員朝礼" },
      { id: uid(), time: "08:30", title: "朝の会（SHR）" },
      { id: uid(), time: "12:40", title: "給食" },
      { id: uid(), time: "13:15", title: "清掃" },
      { id: uid(), time: "15:25", title: "終学活（帰りのSHR）" },
    ],
    subjects: [
      { name: "数学", color: "#3E9BC9" },
      { name: "英語", color: "#E8845B" },
      { name: "国語", color: "#B36AC2" },
      { name: "理科", color: "#4F9E86" },
      { name: "社会", color: "#E0A64B" },
      { name: "保健体育", color: "#5B7CE0" },
      { name: "道徳", color: "#8894A0" },
      { name: "学活", color: "#9AA0A6" },
      { name: "総合", color: "#6BAF7A" },
    ],
    classes: ["1-1", "1-2", "2-1", "2-2", "3-1", "3-2"],
    homeroom: "2-1",
    targets: { 数学: 140, 英語: 140, 国語: 140, 理科: 140, 社会: 105, 保健体育: 105, 道徳: 35, 学活: 35, 総合: 70 },
    // timetable[dayIdx][periodIdx] = { subject, klass, room }
    timetable: {
      "0-0": { subject: "数学", klass: "2-1", room: "" },
      "0-1": { subject: "数学", klass: "1-1", room: "" },
      "0-3": { subject: "数学", klass: "3-1", room: "" },
      "0-5": { subject: "学活", klass: "2-1", room: "" },
      "1-0": { subject: "数学", klass: "2-2", room: "" },
      "1-2": { subject: "数学", klass: "2-1", room: "" },
      "1-4": { subject: "数学", klass: "1-2", room: "" },
      "2-1": { subject: "数学", klass: "3-1", room: "" },
      "2-2": { subject: "数学", klass: "1-1", room: "" },
      "2-3": { subject: "道徳", klass: "2-1", room: "" },
      "3-0": { subject: "数学", klass: "1-2", room: "" },
      "3-1": { subject: "数学", klass: "2-2", room: "" },
      "3-4": { subject: "数学", klass: "2-1", room: "" },
      "4-2": { subject: "数学", klass: "1-1", room: "" },
      "4-3": { subject: "数学", klass: "3-1", room: "" },
      "4-5": { subject: "総合", klass: "2-1", room: "" },
    },
    // lessonLog[`${date}-${periodIdx}`] = { done, subject, klass, topic, hw }
    lessonLog: {},
    events: [
      ev(1, "職員会議", "meeting", "16:00"),
      ev(4, "中間テスト①", "exam", ""),
      ev(9, "生徒総会", "school", "13:30"),
    ],
    club: {
      name: "",
      // schedule[dayIdx] = { on, start, end, place, note } … 日毎入力が無い日の基本表示
      schedule: {},
      // days[date] = { kind:'practice'|'match'|'off', session, place:'二中'|'other', placeOther, time, content, matchName, note }
      days: {},
      specials: [],
    },
    duties: [
      { id: uid(), title: "学年会", dayIdxs: [1], time: "15:40", place: "会議室", note: "" },
      { id: uid(), title: "数学科会", dayIdxs: [3], time: "15:40", place: "数学準備室", note: "" },
      { id: uid(), title: "職員会議", dayIdxs: [0], time: "16:00", place: "会議室", note: "第1・3月曜" },
    ],
    rosters: {
      "2-1": [
        { id: uid(), no: 1, name: "", kana: "", memo: "" },
      ],
    },
    todos: [], // {id, date, text, done}
    meetings: [], // {id, date, title, cat, notes, imgs:[imgId]} 職員会議
    weeklyManual: {}, // { [klass]: 週コマ数 } 手入力があれば集計で優先
    dayMemo: {}, // { [date]: テキストメモ（日報取り込み先） }
    // ---- 授業（年間計画） ----
    textbook: {
      grade: "2年",
      units: [
        { id: uid(), program: "PROGRAM 1", title: "A Trip to Finland", pageFrom: 7, pageTo: 16, grammar: ["It is ... for ... to ~", "tell / ask ... to ~"] },
        { id: uid(), program: "PROGRAM 2", title: "Enjoy Sushi", pageFrom: 17, pageTo: 27, grammar: ["have / has + 過去分詞（継続）", "How long ~?"] },
        { id: uid(), program: "PROGRAM 3", title: "Taste of Culture", pageFrom: 29, pageTo: 34, grammar: ["tell ... that ~", "make A B（SVOC）"] },
      ],
    },
    tests: [
      { id: uid(), name: "1学期中間", type: "定期", date: ymd(addDays(today, 21)), pageFrom: 7, pageTo: 16 },
      { id: uid(), name: "1学期期末", type: "定期", date: ymd(addDays(today, 56)), pageFrom: 7, pageTo: 27 },
    ],
    planClass: "2-1",
    testProgress: {}, // { [testId]: { plannedTo, actualTo, status, note } } 教科担当者で共有する進度計画
    roadmap: {}, // { [subjectName]: { units:[{id,name}], phases:[...], start, perWeek } } 教科ごとのロードマップ
    terms: [ // 時数カウントの起点（学期）
      { id: uid(), name: "1学期", start: `${fyStart}-04-01` },
      { id: uid(), name: "2学期", start: `${fyStart}-09-01` },
      { id: uid(), name: "3学期", start: `${fyStart + 1}-01-08` },
    ],
    planStart: ymd(new Date(today.getMonth() >= 3 ? today.getFullYear() : today.getFullYear() - 1, 3, 1)),
  };
}

/* ---------- small utils bound to data ---------- */
const subjColor = (data, name) => {
  const s = data.subjects.find((x) => x.name === name);
  return s ? s.color : "#9AA0A6";
};
const eventMeta = {
  school: { label: "行事", color: "#3E9BC9" },
  exam: { label: "テスト", color: "#E0A64B" },
  meeting: { label: "会議", color: "#8894A0" },
  club: { label: "部活", color: "#E8845B" },
  other: { label: "その他", color: "#4F9E86" },
};

/* カレンダーの表示カテゴリ（各カテゴリを表示/非表示で切替） */
const CAL_CATS = [
  { id: "event", label: "学校行事", color: "#3E9BC9" },
  { id: "duty", label: "校務", color: "#8894A0" },
  { id: "lesson", label: "授業計画", color: "#4F9E86" },
  { id: "club", label: "部活", color: "#E8845B" },
];
const eventCat = (type) => (type === "meeting" ? "duty" : type === "club" ? "club" : "event");
const defaultVis = () => ({ event: true, duty: true, club: true, lesson: false });

const SESS_LABEL = { after: "放課後", am: "午前", pm: "午後" };
const dayKind = (dk) => dk.kind || (dk.practice === false ? "off" : "practice");
/* 指定日の部活予定（日毎入力 > 曜日テンプレート の順で表示） */
function clubDayDisplay(data, dateKey, jsIdx) {
  const dk = data.club.days?.[dateKey];
  if (dk) {
    const kind = dayKind(dk);
    const place = dk.place === "other" ? (dk.placeOther || "その他") : "二中";
    if (kind === "off") return { kind: "off", content: "部活休み", place: "", time: "", note: dk.note || "" };
    if (kind === "match") return { kind: "match", content: `【大会】${dk.matchName || "大会"}`, place, time: dk.time || "", note: dk.note || "" };
    return { kind: "practice", content: `${SESS_LABEL[dk.session] || "放課後"}練習${dk.content ? `（${dk.content}）` : ""}`, place, time: dk.time || "", note: dk.note || "" };
  }
  const idx = jsIdx;
  const cs = idx >= 0 ? data.club.schedule?.[idx] : null;
  if (cs && cs.on) return { kind: "practice", content: "部活動", place: cs.place || "", time: cs.start ? `${cs.start}–${cs.end || ""}` : "", note: cs.note || "", weekly: true };
  return null;
}

/* 指定日で、表示ONのカテゴリの予定項目だけを集約（vis = {event,duty,lesson,club}） */
function calItemsForDate(data, date, vis) {
  const k = ymd(date);
  const idx = jsDayToIdx(date.getDay());
  const out = [];
  data.events.filter((e) => e.date === k).forEach((e) => {
    const c = eventCat(e.type);
    if (vis[c]) out.push({ title: e.title, color: eventMeta[e.type]?.color, cat: c, time: e.time });
  });
  if (vis.club) {
    data.club.specials.filter((s) => s.date === k).forEach((s) => out.push({ title: s.title, color: "#E8845B", cat: "club", club: true, special: true, time: s.start, lines: [s.title, s.place || "", s.start ? `${s.start}–${s.end || ""}` : "", ""] }));
    const cd = clubDayDisplay(data, k, idx);
    if (cd) out.push({ cat: "club", club: true, color: cd.kind === "match" ? "#D9534F" : cd.kind === "off" ? "#8894A0" : "#E8845B", time: cd.time, weekly: cd.weekly, special: cd.kind === "match", title: cd.content, lines: [cd.content, cd.place, cd.time, cd.note] });
  }
  if (vis.duty && idx >= 0) data.duties.filter((d) => d.dayIdxs.includes(idx)).forEach((d) => out.push({ title: d.title, color: "#8894A0", cat: "duty", time: d.time }));
  if (vis.lesson) Object.entries(data.lessonLog).forEach(([lk, l]) => { if (lk.startsWith(k + "-") && l.topic) out.push({ title: (l.klass ? l.klass + " " : "") + l.topic, color: subjColor(data, l.subject), cat: "lesson" }); });
  return out.sort((a, b) => t2m(a.time) - t2m(b.time));
}

/* 期間内の「予定のある日」を集約（印刷用） */
function collectRows(data, fromDate, toDate, vis, includeFaint = false) {
  const rows = [];
  let d = new Date(fromDate);
  while (d <= toDate) {
    const items = calItemsForDate(data, d, vis).filter((it) => includeFaint || !it.faint);
    if (items.length) rows.push({ date: new Date(d), items });
    d = addDays(d, 1);
  }
  return rows;
}

/* カテゴリの表示/非表示トグル */
function CategoryToggles({ vis, onToggle }) {
  return (
    <div className="tp-flick">
      {CAL_CATS.map((c) => {
        const on = !!vis[c.id];
        return (
          <button key={c.id} className={"tp-flick-seg" + (on ? " on" : " off")} onClick={() => onToggle(c.id)}
            style={on ? { background: c.color, borderColor: c.color } : {}}>
            <span className="tp-flick-dot" style={{ background: on ? "#fff" : c.color }} />{c.label}
          </button>
        );
      })}
    </div>
  );
}

/* ============================================================
   画像取り込み（Claude Vision）ヘルパー
   ============================================================ */
const EVENT_TYPES = ["school", "exam", "meeting", "club", "other"];
const sanitizeType = (t) => (EVENT_TYPES.includes(t) ? t : "school");

function repairTruncatedJSON(t) {
  let i = t.lastIndexOf("}"); if (i < 0) return t;
  t = t.slice(0, i + 1);
  const stack = []; let inStr = false, esc = false;
  for (const ch of t) {
    if (esc) { esc = false; continue; }
    if (ch === "\\") { esc = true; continue; }
    if (ch === '"') { inStr = !inStr; continue; }
    if (inStr) continue;
    if (ch === "{" || ch === "[") stack.push(ch);
    else if (ch === "}" || ch === "]") stack.pop();
  }
  let close = ""; for (let k = stack.length - 1; k >= 0; k--) close += stack[k] === "{" ? "}" : "]";
  return t + close;
}
function parseJSONLoose(text) {
  let t = String(text || "").trim().replace(/```json/gi, "").replace(/```/g, "").trim();
  const s = t.search(/[\[{]/);
  if (s > 0) t = t.slice(s);
  const e = Math.max(t.lastIndexOf("}"), t.lastIndexOf("]"));
  const whole = (s >= 0 && e > 0) ? t.slice(0, e + 1) : t;
  try { return JSON.parse(whole); } catch (_) {}
  try { return JSON.parse(repairTruncatedJSON(t)); } catch (_) {}
  throw new Error("読み取り結果を解析できませんでした（内容が長い場合に起きることがあります）。写真を分割する／範囲を絞るとうまくいくことがあります。");
}

function buildImportPrompt(kind, ctx, teacher) {
  const dutyLine = teacher
    ? `また、「日直」「当番」「担当」などの欄に「${teacher}」の名前がある日があれば myDuties に日付を入れてください。`
    : "";
  if (kind === "month") {
    return {
      system: "あなたは日本の学校の行事予定表を正確に読み取るアシスタントです。出力は指定のJSONのみ。前置き・説明・コードフェンスは一切書かないこと。",
      text: `この画像は${ctx.year}年${ctx.month}月の学校の月間行事予定表です。日付ごとの予定をすべて抽出してください。
- date は "YYYY-MM-DD"（${ctx.year}年${ctx.month}月）。
- type は school(行事) / exam(定期考査・テスト) / meeting(会議・職員朝礼など校務) / club(部活・大会) / other から最も近いものを選ぶ。
- time は時刻が読み取れれば "HH:MM"、なければ空文字 ""。
- title は簡潔に。
${dutyLine}
次のJSON形式だけを返す:
{"events":[{"date":"","title":"","type":"","time":""}],"myDuties":[{"date":"","title":"日直"}],"notes":""}`,
    };
  }
  if (kind === "year") {
    return {
      system: "あなたは日本の学校の年間行事予定表を正確に読み取るアシスタントです。出力は指定のJSONのみ。前置き・説明・コードフェンスは書かないこと。",
      text: `この画像は${ctx.startYear}年度（${ctx.startYear}年4月〜${ctx.startYear + 1}年3月）の年間行事予定表です。主な行事を抽出してください。
- 4〜12月は${ctx.startYear}年、1〜3月は${ctx.startYear + 1}年として date を "YYYY-MM-DD" にする。
- type は school / exam / meeting / club / other。
- title は短く。件数が多い場合は主要な行事を優先。
${dutyLine}
次のJSON形式だけを返す:
{"events":[{"date":"","title":"","type":"","time":""}],"myDuties":[{"date":"","title":"日直"}],"notes":""}`,
    };
  }
  if (kind === "roster") {
    return {
      system: "あなたは学級の生徒名簿を正確に読み取るアシスタントです。出力は指定のJSONのみ。前置き・説明・コードフェンスは書かないこと。",
      text: `この画像は「${ctx.klass || ""}」の生徒名簿です。出席番号・氏名・ふりがなを、番号順に抽出してください。
- no は出席番号（数字）。
- name は氏名。
- kana はふりがな（読めれば。無ければ空文字）。
次のJSON形式だけを返す:
{"students":[{"no":1,"name":"","kana":""}]}`,
    };
  }
  if (kind === "timetable") {
    return {
      system: "あなたは学校の週間時間割表を正確に読み取るアシスタントです。出力は指定のJSONのみ。前置き・説明・コードフェンスは書かないこと。",
      text: `この画像は週間時間割です。各コマ（曜日×時限）を抽出してください。
- day は「月」「火」「水」「木」「金」「土」のいずれか。
- period は時限の数字（1〜6）。
- subject は教科。できれば次の表記に正規化：${(ctx.subjects || []).join("、")}。
- klass はクラス（例：${(ctx.classes || []).slice(0, 3).join("、")} の形式）。読めなければ空文字。
- room は教室（あれば）。
- 空きコマ・行事・給食などは含めない。授業コマのみ。
次のJSON形式だけを返す:
{"cells":[{"day":"月","period":1,"subject":"","klass":"","room":""}]}`,
    };
  }
  if (kind === "club") {
    return {
      system: "あなたは部活動の月間予定表（練習・大会・休み）を正確に読み取るアシスタントです。出力は指定のJSONのみ。前置き・説明・コードフェンスは書かないこと。",
      text: `この画像は部活動の月間予定表です。日ごとの予定を抽出してください。
- date は "YYYY-MM-DD"。年の記載が無ければ ${ctx.year} 年、月の記載が無ければ ${ctx.month} 月として妥当に判断する。
- content は内容（例：練習、放課後練習、午前練習、オフ/休み、◯◯大会 など）。
- place は場所（分かれば）、time は "HH:MM–HH:MM" か開始時刻、無ければ空文字。
- 休み・オフの日は content を「休み」にする。
次のJSON形式だけを返す:
{"days":[{"date":"","content":"","place":"","time":""}]}`,
    };
  }
  if (kind === "match") {
    return {
      system: "あなたは部活動の大会・試合予定表を正確に読み取るアシスタントです。出力は指定のJSONのみ。前置き・説明・コードフェンスは書かないこと。",
      text: `この画像は部活動の大会・試合の予定表です。各予定を抽出してください。
- date は "YYYY-MM-DD"。年の記載が無ければ ${ctx.year} 年として、月から妥当に判断する。
- name は大会名・対戦相手など。
- place は会場名（分かれば）。
- time は "HH:MM–HH:MM" か開始時刻、無ければ空文字。
次のJSON形式だけを返す:
{"matches":[{"date":"","name":"","place":"","time":""}],"notes":""}`,
    };
  }
  // daily
  return {
    system: "あなたは日本の学校の学級日誌・日報・日直日誌を読み取るアシスタントです。出力は指定のJSONのみ。前置き・説明・コードフェンスは書かないこと。",
    text: `この画像は学級日誌／日報です。想定日は ${ctx.date}。次を抽出してください。
- date: 日誌に日付があればその "YYYY-MM-DD"、無ければ ${ctx.date}。
- events: その日の予定・連絡（{title,type,time}。type は school/exam/meeting/club/other、time は "HH:MM" か ""）。
- todos: やること・宿題・持ち物などの短い文字列の配列。
- memo: 本文の要点をまとめたテキスト。
- myDuty: この日の日直が「${teacher || "（氏名未設定）"}」なら true、そうでなければ false。
次のJSON形式だけを返す:
{"date":"","events":[{"title":"","type":"","time":""}],"todos":[""],"memo":"","myDuty":false}`,
  };
}

async function extractFromImage(kind, ctx, b64, mediaType, teacher) {
  const isPdf = mediaType === "application/pdf";
  const source = { type: "base64", media_type: isPdf ? "application/pdf" : (mediaType || "image/jpeg"), data: b64 };
  const block = isPdf ? { type: "document", source } : { type: "image", source };
  const { system, text } = buildImportPrompt(kind, ctx, teacher);
  const maxTok = kind === "year" ? 8000 : (kind === "month" || kind === "roster") ? 6000 : 4000;
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ model: "claude-sonnet-4-6", max_tokens: maxTok, system, messages: [{ role: "user", content: [block, { type: "text", text }] }] }),
  });
  if (!res.ok) throw new Error(`接続エラー（${res.status}）。校内ネットワーク(iFilter)ではAPI通信が遮断される場合があります。自宅・個人端末でお試しください。`);
  const j = await res.json();
  const out = (j.content || []).map((b) => (b.type === "text" ? b.text : "")).join("");
  return parseJSONLoose(out);
}

/* 取り込みモーダル：画像→抽出→確認→反映 */
function ImportModal({ open, kind: kind0, ctx, data, setData, onClose }) {
  const [kind, setKind] = useState(kind0);
  const [rosterKlass, setRosterKlass] = useState("");
  const [preview, setPreview] = useState(null);
  const [b64, setB64] = useState(null);
  const [mt, setMt] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [result, setResult] = useState(null);
  const [evSel, setEvSel] = useState([]);
  const [dutySel, setDutySel] = useState([]);
  const [todoSel, setTodoSel] = useState([]);
  const [memoEdit, setMemoEdit] = useState("");
  const [myDuty, setMyDuty] = useState(false);
  const fileRef = useRef(null);

  useEffect(() => { if (open) { setKind(kind0); setRosterKlass((ctx && ctx.klass) || data.homeroom || (data.classes || [])[0] || ""); setPreview(null); setB64(null); setMt(""); setLoading(false); setError(""); setResult(null); setEvSel([]); setDutySel([]); setTodoSel([]); setMemoEdit(""); setMyDuty(false); } }, [open, kind0]);
  // 種別を切り替えたら結果だけリセット（画像は保持して再実行できる）
  useEffect(() => { setResult(null); setError(""); setEvSel([]); setDutySel([]); setTodoSel([]); setMemoEdit(""); setMyDuty(false); }, [kind]);

  const nowY = new Date().getFullYear(); const nowM = new Date().getMonth() + 1;
  const fiscalStart = nowM < 4 ? nowY - 1 : nowY;
  const curCtx = { ...(ctx || {}),
    year: (ctx && ctx.year) || (ctx && ctx.startYear) || nowY,
    month: (ctx && ctx.month) || nowM,
    startYear: (ctx && ctx.startYear) || fiscalStart,
    date: (ctx && ctx.date) || ymd(new Date()),
    klass: rosterKlass || (ctx && ctx.klass) || data.homeroom || (data.classes || [])[0] || "",
  };

  const onFile = (f) => {
    if (!f) return;
    setError(""); setResult(null); setMt(f.type || "image/jpeg");
    const r = new FileReader();
    r.onload = () => { const url = String(r.result); setPreview(url); setB64(url.split(",")[1]); };
    r.readAsDataURL(f);
  };

  const run = async () => {
    if (!b64) { setError("先に画像を選んでください。"); return; }
    setLoading(true); setError("");
    try {
      const r = await extractFromImage(kind, curCtx, b64, mt, data.meta.teacher);
      setResult(r);
      if (kind === "daily") {
        setEvSel((r.events || []).map(() => true));
        setTodoSel((r.todos || []).map(() => true));
        setMemoEdit(r.memo || "");
        setMyDuty(!!r.myDuty);
      } else if (kind === "match") {
        setEvSel((r.matches || []).map(() => true));
      } else if (kind === "roster") {
        setEvSel((r.students || []).map(() => true));
      } else if (kind === "timetable") {
        setEvSel((r.cells || []).map(() => true));
      } else if (kind === "club") {
        setEvSel((r.days || []).map(() => true));
      } else {
        setEvSel((r.events || []).map(() => true));
        setDutySel((r.myDuties || []).map(() => true));
      }
      if (!(r.events || []).length && !(r.myDuties || []).length && !(r.todos || []).length && !(r.matches || []).length && !(r.cells || []).length && !(r.students || []).length && !(r.days || []).length && !r.memo) setError("読み取れませんでした。鮮明な画像で再度お試しください。");
    } catch (e) {
      setError(e.message || "読み取りに失敗しました。");
    } finally { setLoading(false); }
  };

  const commit = () => {
    if (kind === "roster") {
      setData((d) => {
        const students = (result.students || []).filter((_, i) => evSel[i]).map((s) => ({ id: uid(), no: s.no || "", name: s.name || "", kana: s.kana || "", memo: "" }));
        const classes = new Set(d.classes); if (curCtx.klass) classes.add(curCtx.klass);
        return { ...d, rosters: { ...(d.rosters || {}), [curCtx.klass]: students }, classes: Array.from(classes) };
      });
      onClose();
      return;
    }
    if (kind === "timetable") {
      setData((d) => {
        const tt = { ...d.timetable };
        (result.cells || []).forEach((c) => { const di = DAYMAP[(c.day || "")[0]]; const pi = (Number(c.period) || 0) - 1; if (di >= 0 && pi >= 0 && c.subject) tt[`${di}-${pi}`] = { subject: c.subject, klass: c.klass || "", room: c.room || "" }; });
        return { ...d, timetable: tt };
      });
      onClose();
      return;
    }
    if (kind === "club") {
      const rows = parseClubRows((result.days || []).map((x) => ({ 日付: x.date, 内容: x.content, 場所: x.place, 時間: x.time })), curCtx.year);
      setData((d) => { const nd = { ...(d.club.days || {}) }; rows.forEach((r, i) => { if (evSel[i]) nd[r.date] = r.entry; }); return { ...d, club: { ...d.club, days: nd } }; });
      onClose();
      return;
    }
    if (kind === "match") {
      setData((d) => {
        const nd = { ...(d.club.days || {}) };
        (result.matches || []).forEach((mm, i) => { if (evSel[i] && /^\d{4}-\d{2}-\d{2}$/.test(mm.date)) nd[mm.date] = { kind: "match", matchName: mm.name || "大会", place: mm.place ? "other" : "二中", placeOther: mm.place || "", time: mm.time || "", note: "" }; });
        return { ...d, club: { ...d.club, days: nd } };
      });
      onClose();
      return;
    }
    if (kind === "daily") {
      const date = /^\d{4}-\d{2}-\d{2}$/.test(result.date || "") ? result.date : curCtx.date;
      setData((d) => {
        const add = [];
        (result.events || []).forEach((e, i) => { if (evSel[i]) add.push({ id: uid(), date, title: e.title, type: sanitizeType(e.type), time: e.time || "" }); });
        if (myDuty) add.push({ id: uid(), date, title: "日直", type: "meeting", time: "" });
        const newTodos = (result.todos || []).filter((_, i) => todoSel[i]).map((t) => ({ id: uid(), date, text: t, done: false }));
        const dm = { ...(d.dayMemo || {}) };
        if (memoEdit.trim()) dm[date] = (dm[date] ? dm[date] + "\n" : "") + memoEdit.trim();
        return { ...d, events: [...d.events, ...add], todos: [...d.todos, ...newTodos], dayMemo: dm };
      });
    } else {
      setData((d) => {
        const exists = new Set(d.events.map((e) => e.date + "|" + e.title));
        const add = [];
        (result.events || []).forEach((e, i) => { if (evSel[i] && /^\d{4}-\d{2}-\d{2}$/.test(e.date) && !exists.has(e.date + "|" + e.title)) add.push({ id: uid(), date: e.date, title: e.title, type: sanitizeType(e.type), time: e.time || "" }); });
        (result.myDuties || []).forEach((m, i) => { if (dutySel[i] && /^\d{4}-\d{2}-\d{2}$/.test(m.date) && !exists.has(m.date + "|日直")) add.push({ id: uid(), date: m.date, title: "日直", type: "meeting", time: "" }); });
        return { ...d, events: [...d.events, ...add] };
      });
    }
    onClose();
  };

  const titleMap = { month: "月間予定表", year: "年間予定表", daily: "日報", match: "大会予定", timetable: "週間時間割", roster: "名簿（画像）", club: "部活の月間予定" };
  const hintMap = { month: "月ごとの行事予定表の写真から予定を抽出します。", year: "年間行事予定表の写真から予定を抽出します。", daily: "1日分の連絡・予定・やることを抽出します。", match: "大会日程表から日付・大会名・会場を抽出します。", timetable: "週の時間割表から曜日×限のコマを抽出します。", roster: "名簿の写真から番号・氏名・ふりがなを抽出します。", club: "部活の月間予定表から、練習・休み・大会を日ごとに抽出します。" };
  const toggle = (arr, set, i) => { const a = [...arr]; a[i] = !a[i]; set(a); };
  const commitCount = kind === "daily"
    ? evSel.filter(Boolean).length + todoSel.filter(Boolean).length + (myDuty ? 1 : 0) + (memoEdit.trim() ? 1 : 0)
    : kind === "timetable" ? (result?.cells || []).length
    : evSel.filter(Boolean).length + dutySel.filter(Boolean).length;

  return (
    <Modal open={open} title="画像から取り込み" onClose={onClose} wide>
      <div className="tp-field"><span>取り込む種別</span>
        <Seg options={[{ v: "month", label: "行事(月)" }, { v: "year", label: "行事(年間)" }, { v: "timetable", label: "時間割" }, { v: "club", label: "部活" }, { v: "match", label: "大会" }, { v: "roster", label: "名簿" }, { v: "daily", label: "日報" }]} value={kind} onChange={setKind} />
      </div>
      {kind === "roster" && (
        <label className="tp-field"><span>取り込み先のクラス</span>
          <select value={rosterKlass} onChange={(e) => setRosterKlass(e.target.value)}>
            {(data.classes || []).map((c) => <option key={c} value={c}>{c}</option>)}
          </select>
        </label>
      )}
      <p className="tp-hint" style={{ marginTop: 0 }}>{titleMap[kind]}：{hintMap[kind]}</p>
      {kind !== "match" && kind !== "daily" && !data.meta.teacher && <div className="tp-warn">日直の判定には、設定で「氏名」を登録してください。</div>}
      {kind === "daily" && !data.meta.teacher && <div className="tp-warn">日直の判定には、設定で「氏名」を登録してください。</div>}
      {!result && (
        <>
          <div className="tp-uploadbox" onClick={() => fileRef.current?.click()}>
            {preview ? (mt === "application/pdf" ? <div className="tp-pdfmark">PDF を選択済み</div> : <img src={preview} alt="preview" />) : (
              <div className="tp-upload-empty"><Upload size={26} /><span>画像 / PDF を選ぶ</span><small>行事予定表や日報を撮影・保存したファイル</small></div>
            )}
          </div>
          <input ref={fileRef} type="file" accept="image/*,application/pdf" style={{ display: "none" }} onChange={(e) => onFile(e.target.files?.[0])} />
          {error && <div className="tp-error">{error}</div>}
          <button className="tp-primarybtn tp-full" disabled={!b64 || loading} onClick={run}>
            {loading ? <><Loader size={15} className="tp-spin" /> 読み取り中…</> : <><ImageIcon size={15} /> 読み取る</>}
          </button>
          <p className="tp-hint">読み取り後に内容を確認してから反映できます。通信にはネットワーク接続が必要です。</p>
        </>
      )}

      {result && (kind === "month" || kind === "year") && (
        <>
          <div className="tp-result-head">読み取り結果（チェックした項目を追加します）</div>
          {(result.events || []).length > 0 && <div className="tp-result-sub">行事・予定</div>}
          <ul className="tp-result-list">
            {(result.events || []).map((e, i) => (
              <li key={i}><input type="checkbox" checked={!!evSel[i]} onChange={() => toggle(evSel, setEvSel, i)} />
                <span className="tp-cat-dot" style={{ background: eventMeta[sanitizeType(e.type)]?.color }} />
                <span className="tp-res-date">{e.date?.slice(5)}</span><span className="tp-res-title">{e.title}</span>{e.time && <small>{e.time}</small>}</li>
            ))}
          </ul>
          {(result.myDuties || []).length > 0 && <div className="tp-result-sub">日直（あなた）</div>}
          <ul className="tp-result-list">
            {(result.myDuties || []).map((m, i) => (
              <li key={i}><input type="checkbox" checked={!!dutySel[i]} onChange={() => toggle(dutySel, setDutySel, i)} />
                <span className="tp-cat-dot" style={{ background: "#8894A0" }} /><span className="tp-res-date">{m.date?.slice(5)}</span><span className="tp-res-title">日直</span></li>
            ))}
          </ul>
          {result.notes && <p className="tp-hint">補足: {result.notes}</p>}
          <div className="tp-modal-actions"><button className="tp-ghostbtn" onClick={() => setResult(null)}>やり直す</button><button className="tp-primarybtn" disabled={!commitCount} onClick={commit}><Plus size={15} /> {commitCount}件を追加</button></div>
        </>
      )}

      {result && kind === "match" && (
        <>
          <div className="tp-result-head">大会予定の読み取り結果</div>
          <ul className="tp-result-list">
            {(result.matches || []).length === 0 && <li className="tp-empty">大会を読み取れませんでした</li>}
            {(result.matches || []).map((mm, i) => (
              <li key={i}><input type="checkbox" checked={!!evSel[i]} onChange={() => toggle(evSel, setEvSel, i)} />
                <span className="tp-cat-dot" style={{ background: "#D9534F" }} />
                <span className="tp-res-date">{mm.date?.slice(5)}</span>
                <span className="tp-res-title">{mm.name}{(mm.place || mm.time) && <small className="tp-rec-hw">{[mm.place, mm.time].filter(Boolean).join(" / ")}</small>}</span></li>
            ))}
          </ul>
          {result.notes && <p className="tp-hint">補足: {result.notes}</p>}
          <div className="tp-modal-actions"><button className="tp-ghostbtn" onClick={() => setResult(null)}>やり直す</button><button className="tp-primarybtn" disabled={!commitCount} onClick={commit}><Plus size={15} /> {commitCount}件を大会に追加</button></div>
        </>
      )}

      {result && kind === "club" && (
        <>
          <div className="tp-result-head">部活の月間予定（チェックした日を追加します）</div>
          <ul className="tp-result-list">
            {(result.days || []).length === 0 && <li className="tp-empty">読み取れませんでした</li>}
            {(result.days || []).map((x, i) => { const dd = /^\d{4}-\d{2}-\d{2}$/.test(x.date || "") ? parseYmd(x.date) : null; return (
              <li key={i}><input type="checkbox" checked={!!evSel[i]} onChange={() => toggle(evSel, setEvSel, i)} />
                <span className="tp-cat-dot" style={{ background: /休|オフ|off/i.test(x.content || "") ? "#8894A0" : /大会|試合/.test(x.content || "") ? "#D9534F" : "#E8845B" }} />
                <span className="tp-res-date">{dd ? `${dd.getMonth() + 1}/${dd.getDate()}` : x.date}</span>
                <span className="tp-res-title">{x.content}{(x.place || x.time) && <small className="tp-rec-hw">{[x.place, x.time].filter(Boolean).join(" / ")}</small>}</span></li>
            ); })}
          </ul>
          <div className="tp-modal-actions"><button className="tp-ghostbtn" onClick={() => setResult(null)}>やり直す</button><button className="tp-primarybtn" disabled={!evSel.some(Boolean)} onClick={commit}><Plus size={15} /> {evSel.filter(Boolean).length}件を部活予定に追加</button></div>
        </>
      )}

      {result && kind === "roster" && (
        <>
          <div className="tp-result-head">名簿の読み取り結果（{ctx.klass}）</div>
          <ul className="tp-result-list">
            {(result.students || []).length === 0 && <li className="tp-empty">読み取れませんでした</li>}
            {(result.students || []).map((s, i) => (
              <li key={i}><input type="checkbox" checked={!!evSel[i]} onChange={() => toggle(evSel, setEvSel, i)} />
                <span className="tp-res-date">{s.no}</span><span className="tp-res-title">{s.name}{s.kana ? `（${s.kana}）` : ""}</span></li>
            ))}
          </ul>
          <p className="tp-hint">「取り込む」で {ctx.klass} の名簿を置き換えます。</p>
          <div className="tp-modal-actions"><button className="tp-ghostbtn" onClick={() => setResult(null)}>やり直す</button><button className="tp-primarybtn" disabled={!commitCount} onClick={commit}><Plus size={15} /> {commitCount}名を取り込む</button></div>
        </>
      )}

      {result && kind === "timetable" && (() => {
        const cells = result.cells || [];
        const hasSat = cells.some((c) => (c.day || "")[0] === "土");
        const cols = hasSat ? 6 : 5;
        const maxP = Math.max(6, ...cells.map((c) => Number(c.period) || 0));
        const find = (di, pi) => cells.find((c) => DAYMAP[(c.day || "")[0]] === di && (Number(c.period) || 0) === pi + 1);
        return (
          <>
            <div className="tp-result-head">読み取り結果（週間時間割）</div>
            <p className="tp-hint" style={{ marginTop: 0 }}>「反映」で現在の時間割に上書きします（同じコマのみ置き換え）。細部は反映後に週間画面で修正できます。</p>
            <div className="tp-ttpv" style={{ gridTemplateColumns: `26px repeat(${cols}, 1fr)` }}>
              <div className="tp-ttpv-corner">限</div>
              {Array.from({ length: cols }).map((_, di) => <div key={di} className="tp-ttpv-head">{DAY_LABELS[di]}</div>)}
              {Array.from({ length: maxP }).map((_, pi) => (
                <React.Fragment key={pi}>
                  <div className="tp-ttpv-p">{pi + 1}</div>
                  {Array.from({ length: cols }).map((_, di) => { const c = find(di, pi); return <div key={di} className="tp-ttpv-cell">{c ? <><b>{c.subject}</b>{c.klass && <span>{c.klass}</span>}</> : ""}</div>; })}
                </React.Fragment>
              ))}
            </div>
            <div className="tp-modal-actions"><button className="tp-ghostbtn" onClick={() => setResult(null)}>やり直す</button><button className="tp-primarybtn" disabled={!commitCount} onClick={commit}><Check size={15} /> {commitCount}コマを反映</button></div>
          </>
        );
      })()}

      {result && kind === "daily" && (
        <>
          <div className="tp-result-head">「今日」への反映内容（{/^\d{4}-\d{2}-\d{2}$/.test(result.date || "") ? result.date : ctx.date}）</div>
          <label className="tp-check"><input type="checkbox" checked={myDuty} onChange={(e) => setMyDuty(e.target.checked)} /> この日の日直として予定に入れる</label>
          {(result.events || []).length > 0 && <div className="tp-result-sub">予定・連絡</div>}
          <ul className="tp-result-list">
            {(result.events || []).map((e, i) => (
              <li key={i}><input type="checkbox" checked={!!evSel[i]} onChange={() => toggle(evSel, setEvSel, i)} />
                <span className="tp-cat-dot" style={{ background: eventMeta[sanitizeType(e.type)]?.color }} /><span className="tp-res-title">{e.title}</span>{e.time && <small>{e.time}</small>}</li>
            ))}
          </ul>
          {(result.todos || []).length > 0 && <div className="tp-result-sub">やること</div>}
          <ul className="tp-result-list">
            {(result.todos || []).map((t, i) => (
              <li key={i}><input type="checkbox" checked={!!todoSel[i]} onChange={() => toggle(todoSel, setTodoSel, i)} /><span className="tp-res-title">{t}</span></li>
            ))}
          </ul>
          <label className="tp-field"><span>メモ（本文）</span><textarea rows={3} value={memoEdit} onChange={(e) => setMemoEdit(e.target.value)} /></label>
          <div className="tp-modal-actions"><button className="tp-ghostbtn" onClick={() => setResult(null)}>やり直す</button><button className="tp-primarybtn" disabled={!commitCount} onClick={commit}><Plus size={15} /> 反映する</button></div>
        </>
      )}
    </Modal>
  );
}

/* ============================================================
   Excel / CSV 取り込み（オフライン・SheetJS）
   ============================================================ */
const zen2han = (s) => String(s == null ? "" : s).replace(/[０-９]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xFEE0)).replace(/／/g, "/").replace(/．/g, ".").replace(/　/g, " ");
/* ---- ICS (iCalendar) 連携 ---- */
const icsEsc = (s) => String(s || "").replace(/\\/g, "\\\\").replace(/;/g, "\\;").replace(/,/g, "\\,").replace(/\n/g, "\\n");
const icsDate = (ymdStr) => ymdStr.replace(/-/g, "");
function buildICS(events) {
  const lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Teaching Partner for YOU//JP//", "CALSCALE:GREGORIAN"];
  const stamp = new Date().toISOString().replace(/[-:]/g, "").replace(/\.\d+Z$/, "Z");
  events.forEach((e, i) => {
    if (!e.date) return;
    lines.push("BEGIN:VEVENT");
    lines.push(`UID:tp-${icsDate(e.date)}-${i}@teaching-planners`);
    lines.push(`DTSTAMP:${stamp}`);
    const tm = String(e.time || "").match(/(\d{1,2}):(\d{2})/);
    if (tm) {
      const s = `${icsDate(e.date)}T${pad(+tm[1])}${tm[2]}00`;
      const endH = (+tm[1] + 1) % 24;
      lines.push(`DTSTART:${s}`);
      lines.push(`DTEND:${icsDate(e.date)}T${pad(endH)}${tm[2]}00`);
    } else {
      lines.push(`DTSTART;VALUE=DATE:${icsDate(e.date)}`);
    }
    lines.push(`SUMMARY:${icsEsc(e.title)}`);
    lines.push("END:VEVENT");
  });
  lines.push("END:VCALENDAR");
  return lines.join("\r\n");
}
function parseICS(text) {
  const out = []; const blocks = text.split(/BEGIN:VEVENT/i).slice(1);
  for (const b of blocks) {
    const body = b.split(/END:VEVENT/i)[0];
    const sum = body.match(/\nSUMMARY:(.*)/i) || body.match(/^SUMMARY:(.*)/i);
    const dt = body.match(/DTSTART[^:]*:(\d{8})(T(\d{2})(\d{2}))?/i);
    if (!sum || !dt) continue;
    const date = `${dt[1].slice(0, 4)}-${dt[1].slice(4, 6)}-${dt[1].slice(6, 8)}`;
    const title = sum[1].trim().replace(/\\,/g, ",").replace(/\\;/g, ";").replace(/\\n/g, " ").replace(/\\\\/g, "\\").replace(/\r$/, "");
    const time = dt[3] ? `${dt[3]}:${dt[4]}` : "";
    out.push({ date, title, time, type: "school" });
  }
  return out;
}

function toYmd(v, ctxYear) {
  if (v == null || v === "") return null;
  if (v instanceof Date && !isNaN(v)) return ymd(v);
  if (typeof v === "number" && v > 20000 && v < 60000) { const d = new Date(Math.round((v - 25569) * 86400 * 1000)); if (!isNaN(d)) return ymd(new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate())); }
  const s = zen2han(v).trim();
  let m = s.match(/(\d{4})\D{0,2}(\d{1,2})\D{0,2}(\d{1,2})/);
  if (m) return `${m[1]}-${pad(+m[2])}-${pad(+m[3])}`;
  m = s.match(/^(\d{1,2})\D{1,2}(\d{1,2})/); // 月日のみ → 文脈年で補完（4-3月=年度）
  if (m && ctxYear) { const mo = +m[1], dy = +m[2]; const yy = mo >= 4 ? ctxYear : ctxYear + 1; if (mo >= 1 && mo <= 12 && dy >= 1 && dy <= 31) return `${yy}-${pad(mo)}-${pad(dy)}`; }
  return null;
}
const pick = (obj, names) => { for (const n of names) { for (const k of Object.keys(obj)) { if (zen2han(k).replace(/\s/g, "").toLowerCase() === n.toLowerCase()) return obj[k]; } } return ""; };
const placeFields = (place) => { const p = String(place || "").trim(); if (!p) return { place: "二中", placeOther: "" }; return p.includes("二中") ? { place: "二中", placeOther: "" } : { place: "other", placeOther: p }; };

function classifyClubContent(text, place, time, note) {
  const raw = String(text || "").trim();
  const s = raw.replace(/\s/g, "");
  if (!s && !String(note || "").trim()) return null;
  const t = String(time || "").trim();
  if (s.includes("放課後")) return { kind: "practice", session: "after", ...placeFields(place || "二中"), time: t, note: note || "" };
  if (s.includes("午前")) return { kind: "practice", session: "am", ...placeFields(place || "二中"), time: t, note: note || "" };
  if (s.includes("午後")) return { kind: "practice", session: "pm", ...placeFields(place || "二中"), time: t, note: note || "" };
  if (s.includes("練習")) return { kind: "practice", session: "after", ...placeFields(place || "二中"), time: t, note: note || "" };
  if (s.includes("休")) return { kind: "off", note: note || raw.replace(/休み?/, "").trim() };
  if (/大会|選手権|招待|対抗|新人|全国|インドア|フェス|杯|試合|大会|遠征/.test(s)) return { kind: "match", matchName: raw, ...placeFields(place), time: t, note: note || "" };
  if (place) return { kind: "match", matchName: raw, ...placeFields(place), time: t, note: note || "" };
  return { kind: "off", note: raw };
}

function parseEventRows(objs, ctxYear) {
  const out = [];
  for (const o of objs) {
    const date = toYmd(pick(o, ["日付", "date", "日にち", "月日"]), ctxYear);
    const title = String(pick(o, ["行事名", "行事", "内容", "title", "名称", "予定"]) || "").trim();
    if (!date || !title) continue;
    const tRaw = String(pick(o, ["種類", "type", "区分", "分類"]) || "").trim();
    const type = /テスト|考査/.test(tRaw) ? "exam" : /会議|職員/.test(tRaw) ? "meeting" : /部活|大会/.test(tRaw) ? "club" : /その他/.test(tRaw) ? "other" : "school";
    const time = String(pick(o, ["時刻", "時間", "time", "開始"]) || "").trim();
    out.push({ date, title, type, time });
  }
  return out;
}
function parseClubRows(objs, ctxYear) {
  const out = [];
  for (const o of objs) {
    const date = toYmd(pick(o, ["日付", "date", "日にち", "月日"]), ctxYear);
    if (!date) continue;
    const content = String(pick(o, ["内容", "予定", "練習", "activity"]) || "").trim();
    const place = String(pick(o, ["場所", "会場", "place"]) || "").trim();
    const time = String(pick(o, ["時間", "時刻", "time"]) || "").trim();
    const note = String(pick(o, ["備考", "メモ", "note"]) || "").trim();
    const c = classifyClubContent(content, place, time, note);
    if (c) out.push({ date, entry: c });
  }
  return out;
}
function parseRosterRows(objs) {
  const out = [];
  for (const o of objs) {
    const klass = String(pick(o, ["クラス", "組", "class", "学級"]) || "").trim();
    const name = String(pick(o, ["氏名", "名前", "生徒名", "name"]) || "").trim();
    if (!klass || !name) continue;
    const no = String(pick(o, ["番号", "出席番号", "no", "No", "出席"]) || "").trim();
    const kana = String(pick(o, ["ふりがな", "よみ", "かな", "フリガナ", "kana"]) || "").trim();
    const memo = String(pick(o, ["備考", "メモ", "係", "note"]) || "").trim();
    out.push({ klass, no: no || (out.filter((r) => r.klass === klass).length + 1), name, kana, memo });
  }
  return out;
}
function parseTimetableGrid(grid) {
  // 先頭行に曜日、先頭列に「限」。セルは「教科 クラス」
  if (!grid || grid.length < 2) return [];
  const head = grid[0].map((x) => String(x || "").trim());
  const dayCol = {}; // colIndex -> dayIdx
  head.forEach((h, ci) => { const i = DAY_LABELS.indexOf(h.replace(/曜.*/, "")); if (i >= 0) dayCol[ci] = i; });
  const out = [];
  for (let r = 1; r < grid.length; r++) {
    const row = grid[r];
    const plabel = String(row[0] || "").replace(/[^0-9]/g, "");
    const pi = plabel ? Number(plabel) - 1 : r - 1;
    if (pi < 0) continue;
    Object.entries(dayCol).forEach(([ci, di]) => {
      const cell = String(row[ci] || "").trim();
      if (!cell) return;
      const parts = cell.split(/[\s\n　]+/).filter(Boolean);
      const subject = parts[0] || "";
      const klass = parts.slice(1).join(" ");
      if (subject) out.push({ di, pi, subject, klass });
    });
  }
  return out;
}

function downloadTemplate(kind) {
  let ws, name;
  if (kind === "event") { ws = XLSX.utils.json_to_sheet([{ 日付: "2026-04-08", 行事名: "始業式", 種類: "行事", 時刻: "" }, { 日付: "2026-05-19", 行事名: "定期テスト", 種類: "テスト", 時刻: "" }]); name = "行事テンプレート.xlsx"; }
  else if (kind === "club") { ws = XLSX.utils.json_to_sheet([{ 日付: "2026-04-10", 内容: "放課後練習", 場所: "二中", 時間: "16:00-18:00", 備考: "" }, { 日付: "2026-04-18", 内容: "松江・安来中学生大会", 場所: "安来", 時間: "", 備考: "" }, { 日付: "2026-04-12", 内容: "休み", 場所: "", 時間: "", 備考: "" }]); name = "部活テンプレート.xlsx"; }
  else if (kind === "roster") { ws = XLSX.utils.json_to_sheet([{ クラス: "2-1", 番号: 1, 氏名: "山田 太郎", ふりがな: "やまだ たろう", 備考: "" }, { クラス: "2-1", 番号: 2, 氏名: "佐藤 花子", ふりがな: "さとう はなこ", 備考: "図書委員" }]); name = "名簿テンプレート.xlsx"; }
  else { ws = XLSX.utils.aoa_to_sheet([["限", "月", "火", "水", "木", "金"], ["1", "数学 2-1", "英語 1-1", "", "", ""], ["2", "", "数学 2-2", "", "", ""]]); name = "時間割テンプレート.xlsx"; }
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  try { XLSX.writeFile(wb, name); } catch (e) {}
}

/* オフライン：貼り付け/スキャンした文字を予定に変換（iOSの「テキストをスキャン」で写真から取り込み可） */
function normalizeEventObjects(arr, ctxYear) {
  const out = [];
  (Array.isArray(arr) ? arr : []).forEach((o) => {
    if (!o || typeof o !== "object") return;
    const gv = (keys) => { for (const k of Object.keys(o)) { const kk = zen2han(k).replace(/\s/g, "").toLowerCase(); if (keys.includes(kk)) return o[k]; } return ""; };
    let dRaw = String(gv(["date", "日付", "日", "ymd", "day"]) || "").trim();
    const title = String(gv(["title", "件名", "予定", "予定名", "行事", "name", "内容", "event"]) || "").trim();
    let time = String(gv(["time", "時刻", "時間", "start"]) || "").trim();
    const type = String(gv(["type", "種別", "分類", "category", "kind"]) || "行事").trim() || "行事";
    if (!dRaw || !title) return;
    const z = zen2han(dRaw); let date = "";
    let m = z.match(/(20\d{2})[\/\.\-年]\s*(\d{1,2})[\/\.\-月]\s*(\d{1,2})/);
    if (m) date = `${+m[1]}-${pad(+m[2])}-${pad(+m[3])}`;
    else { m = z.match(/(\d{1,2})[\/\.\-月]\s*(\d{1,2})/); if (m) { const mo = +m[1]; date = `${mo >= 4 ? ctxYear : ctxYear + 1}-${pad(mo)}-${pad(+m[2])}`; } }
    if (!date) return;
    const tm = zen2han(time).match(/(\d{1,2})[:：](\d{2})/); time = tm ? `${pad(+tm[1])}:${tm[2]}` : "";
    out.push({ date, title: title.slice(0, 50), time, type });
  });
  return out;
}
function parseImportText(text, ctxYear) {
  const t = String(text || "").trim();
  if (/^[\[{]/.test(t)) {
    try {
      const j = JSON.parse(t);
      const arr = Array.isArray(j) ? j : (j.events || j.予定 || j.items || j.data || null);
      if (Array.isArray(arr)) return { rows: normalizeEventObjects(arr, ctxYear), mode: "json" };
    } catch (e) { /* JSONとして壊れている→行解析にフォールバック */ }
  }
  return { rows: parseScheduleText(text, ctxYear), mode: "text" };
}
function parseScheduleText(text, ctxYear) {
  const out = [];
  String(text || "").split(/\r?\n/).forEach((raw) => {
    const line = zen2han(raw).trim(); if (!line) return;
    const m = line.match(/^(?:(20\d{2})[\/\.\-年]\s*)?(\d{1,2})\s*[\/\.\-月]\s*(\d{1,2})\s*日?\s*(?:[（(][日月火水木金土曜]+[)）])?\s*(.*)$/);
    if (!m) return;
    const mo = +m[2], dy = +m[3]; if (mo < 1 || mo > 12 || dy < 1 || dy > 31) return;
    const y = m[1] ? +m[1] : (mo >= 4 ? ctxYear : ctxYear + 1);
    let rest = (m[4] || "").trim(); let time = "";
    const tm = rest.match(/(\d{1,2})\s*[:：時]\s*(\d{2})?/); if (tm) { time = `${pad(+tm[1])}:${tm[2] || "00"}`; rest = rest.replace(tm[0], "").trim(); }
    rest = rest.replace(/^[\s:：\-–—・.\)]+/, "").trim();
    if (!rest) return;
    out.push({ date: `${y}-${pad(mo)}-${pad(dy)}`, title: rest.slice(0, 50), time, type: "行事" });
  });
  return out;
}
function parseClubText(text, ctxYear) {
  const objs = [];
  String(text || "").split(/\r?\n/).forEach((raw) => {
    const line = zen2han(raw).trim(); if (!line) return;
    const m = line.match(/^(?:(20\d{2})[\/\.\-年]\s*)?(\d{1,2})\s*[\/\.\-月]\s*(\d{1,2})\s*日?\s*(?:[（(][^)）]*[)）])?\s*(.*)$/);
    if (!m) return; const mo = +m[2], dy = +m[3]; if (mo < 1 || mo > 12 || dy < 1 || dy > 31) return;
    let rest = (m[4] || "").trim(); let time = "";
    const tm = rest.match(/(\d{1,2}:\d{2})\s*[-〜~]?\s*(\d{1,2}:\d{2})?/); if (tm) { time = tm[0].replace(/\s/g, ""); rest = rest.replace(tm[0], "").trim(); }
    const dateStr = (m[1] ? m[1] + "/" : "") + mo + "/" + dy;
    objs.push({ 日付: dateStr, 内容: rest, 時間: time, 場所: "" });
  });
  return parseClubRows(objs, ctxYear);
}
function parseClubJSON(arr, ctxYear) { return parseClubRows(Array.isArray(arr) ? arr : [], ctxYear); }
function dayCharToIdx(v) {
  const s = zen2han(String(v || "")).trim().replace(/曜.*/, "");
  const i = DAY_LABELS.indexOf(s); if (i >= 0) return i;
  const n = parseInt(s, 10); if (!isNaN(n)) { if (n >= 1 && n <= 6) return n - 1; if (n === 0) return -1; }
  return -1;
}
function parseTimetableText(text) {
  const out = [];
  String(text || "").split(/\r?\n/).forEach((raw) => {
    const line = zen2han(raw).trim(); if (!line) return;
    const m = line.match(/^([月火水木金土])\s*曜?\s*(\d{1,2})\s*限?\s*[\s:：]\s*(.+)$/) || line.match(/^([月火水木金土])\s*曜?\s*(\d{1,2})\s*限?\s+(.+)$/);
    if (!m) return; const di = DAY_LABELS.indexOf(m[1]); const pi = +m[2] - 1; if (di < 0 || pi < 0) return;
    const parts = (m[3] || "").trim().split(/[\s　]+/).filter(Boolean);
    const subject = parts[0] || ""; if (!subject) return;
    const klass = parts[1] || ""; const room = parts.slice(2).join(" ");
    out.push({ di, pi, subject, klass, room });
  });
  return out;
}
function parseTimetableJSON(arr) {
  if (Array.isArray(arr) && arr.length && Array.isArray(arr[0])) return parseTimetableGrid(arr).map((r) => ({ ...r, room: "" }));
  const out = [];
  (Array.isArray(arr) ? arr : []).forEach((o) => {
    if (!o || typeof o !== "object") return;
    const gv = (keys) => { for (const k of Object.keys(o)) { const kk = zen2han(k).replace(/\s/g, "").toLowerCase(); if (keys.includes(kk)) return o[k]; } return ""; };
    const di = dayCharToIdx(gv(["day", "曜日", "曜", "weekday"]));
    let pi = parseInt(zen2han(String(gv(["period", "限", "時限", "校時", "コマ"]) || "")), 10); pi = isNaN(pi) ? -1 : pi - 1;
    const subject = String(gv(["subject", "教科", "科目", "授業"]) || "").trim();
    const klass = String(gv(["klass", "class", "クラス", "組", "学級"]) || "").trim();
    const room = String(gv(["room", "教室", "場所"]) || "").trim();
    if (di < 0 || pi < 0 || !subject) return;
    out.push({ di, pi, subject, klass, room });
  });
  return out;
}

function OfflineTextImportModal({ open, data, setData, onClose, showToast }) {
  const ctxYear = parseInt(String(data.meta.year || "").replace(/\D/g, ""), 10) || new Date().getFullYear();
  const [kind, setKind] = useState("event"); // event | timetable | club
  const [text, setText] = useState("");
  const [type, setType] = useState("行事");
  const [rows, setRows] = useState(null);
  const [mode, setMode] = useState("text");
  const [sel, setSel] = useState([]);
  useEffect(() => { if (open) { setKind("event"); setText(""); setRows(null); setSel([]); setType("行事"); setMode("text"); } }, [open]);
  useEffect(() => { setRows(null); setSel([]); }, [kind]);
  const run = () => {
    const t = text.trim(); const isJson = /^[\[{]/.test(t); let r = [], m = "text";
    if (isJson) {
      try {
        const j = JSON.parse(t); const arr = Array.isArray(j) ? j : (j.events || j.予定 || j.items || j.data || j.rows || j.timetable || j.club || j.days || null); m = "json";
        if (kind === "event") r = normalizeEventObjects(Array.isArray(arr) ? arr : [], ctxYear);
        else if (kind === "club") r = parseClubJSON(arr, ctxYear);
        else r = parseTimetableJSON(arr);
      } catch (e) { m = "text"; }
    }
    if (m === "text") {
      if (kind === "event") r = parseScheduleText(text, ctxYear);
      else if (kind === "club") r = parseClubText(text, ctxYear);
      else r = parseTimetableText(text);
    }
    setRows(r); setMode(m); setSel(r.map(() => true));
  };
  const imp = () => {
    if (!rows) return; const chosen = rows.filter((_, i) => sel[i]); if (!chosen.length) { onClose(); return; }
    const prev = data;
    if (kind === "event") {
      const items = chosen.map((r) => ({ id: uid(), date: r.date, title: r.title, type: mode === "json" ? (r.type || type) : type, time: r.time || "" }));
      setData((d) => ({ ...d, events: [...(d.events || []), ...items] }));
      showToast && showToast(`${items.length}件の予定を取り込みました`, prev);
    } else if (kind === "club") {
      setData((d) => { const days = { ...(d.club.days || {}) }; chosen.forEach((r) => { days[r.date] = r.entry; }); return { ...d, club: { ...d.club, days } }; });
      showToast && showToast(`部活予定を${chosen.length}件取り込みました`, prev);
    } else {
      setData((d) => { const tt = { ...d.timetable }; chosen.forEach((r) => { tt[`${r.di}-${r.pi}`] = { subject: r.subject, klass: r.klass || "", room: r.room || "" }; }); return { ...d, timetable: tt }; });
      showToast && showToast(`時間割を${chosen.length}コマ取り込みました`, prev);
    }
    onClose();
  };
  const KINDS = [{ v: "event", label: "予定（月間・週間）" }, { v: "timetable", label: "週の時間割" }, { v: "club", label: "部活予定" }];
  const placeholder = kind === "event" ? "例）\n4/8 始業式\n4月9日(木) 入学式 9:30\n2026/4/20 家庭訪問\n\n※JSONも可：[{\"date\":\"2026-04-08\",\"title\":\"始業式\"}]"
    : kind === "club" ? "例）\n4/10 放課後練習 16:00-18:00\n4/18 松江・安来中学生大会 安来\n4/12 休み\n\n※JSONも可：[{\"日付\":\"4/10\",\"内容\":\"放課後練習\",\"時間\":\"16:00-18:00\"}]"
    : "例）\n月1 数学 2-1 理科室\n月2 英語 1-1\n火3 保健体育 2-2\n\n※JSONも可：[{\"曜日\":\"月\",\"限\":1,\"教科\":\"数学\",\"クラス\":\"2-1\"}]";
  const dayLabel = ["月", "火", "水", "木", "金", "土"];
  return (
    <Modal open={open} title="写真/テキストから取り込み（オフライン）" onClose={onClose} wide>
      <div className="tp-field"><span>取り込む種別</span>
        <Seg options={KINDS} value={kind} onChange={setKind} />
      </div>
      <div className="tp-ocr-guide">
        <b>写真から取り込む手順（オフライン・通信なし）</b>
        <ol>
          <li>入力欄をタップ →キーボードの <b>スキャン（📷）</b>、または長押し →「<b>テキストをスキャン</b>」。</li>
          <li>写真にカメラを向けて文字を挿入し、下の形式に整えて「読み取り」。</li>
        </ol>
        <p className="tp-hint">
          {kind === "event" && "1行に「日付＋予定名」（例：4/8 始業式）。月間・週間どちらの予定もここに入れます。"}
          {kind === "club" && "1行に「日付＋内容（＋場所・時間）」。内容の「練習・休み・大会名」で種類を自動判定します。"}
          {kind === "timetable" && "1行に「曜日＋限＋教科（＋クラス・教室）」（例：月1 数学 2-1）。"}
          {" "}<b>JSON</b>を貼り付けても読み取れます。
        </p>
      </div>
      <label className="tp-field"><span>文字（貼り付け／スキャン／JSON）</span>
        <textarea rows={7} value={text} onChange={(e) => setText(e.target.value)} placeholder={placeholder} />
      </label>
      <div className="tp-field-row">
        {kind === "event" && <label className="tp-field"><span>予定の種別</span>
          <select value={type} onChange={(e) => setType(e.target.value)}>{["行事", "校務", "部活", "授業", "その他"].map((t) => <option key={t} value={t}>{t}</option>)}</select>
        </label>}
        <button className="tp-primarybtn" style={{ alignSelf: "flex-end", marginLeft: "auto" }} onClick={run}><Loader size={15} /> 読み取り</button>
      </div>
      {rows && (
        <div className="tp-ocr-result">
          <div className="tp-mtg-grouphead">読み取り結果 {rows.length}件（{mode === "json" ? "JSON" : "テキスト"}）{rows.length === 0 && "（形式に合う行が見つかりませんでした）"}</div>
          {rows.map((r, i) => {
            let body;
            if (kind === "event") { const d = parseYmd(r.date); body = <><span className="tp-ocr-date">{d.getMonth() + 1}/{d.getDate()}（{WD[d.getDay()]}）{r.time && " " + r.time}</span><input className="tp-ocr-title" value={r.title} onChange={(e) => setRows((rs) => rs.map((x, j) => j === i ? { ...x, title: e.target.value } : x))} /></>; }
            else if (kind === "club") { const d = parseYmd(r.date); const e = r.entry || {}; const SES = { after: "放課後練習", am: "午前練習", pm: "午後練習" }; const label = e.kind === "match" ? `大会：${e.matchName || e.content || ""}` : e.kind === "off" ? "休み" : (SES[e.session] || e.content || "練習"); body = <><span className="tp-ocr-date">{d.getMonth() + 1}/{d.getDate()}（{WD[d.getDay()]}）</span><span className="tp-ocr-title">{label}{e.place ? "／" + e.place : ""}{e.time ? "／" + e.time : ""}</span></>; }
            else { body = <><span className="tp-ocr-date">{dayLabel[r.di]}曜 {r.pi + 1}限</span><span className="tp-ocr-title">{r.subject}{r.klass ? " " + r.klass : ""}{r.room ? "（" + r.room + "）" : ""}</span></>; }
            return (
              <div key={i} className="tp-ocr-row">
                <button className={"tp-donebox" + (sel[i] ? " on" : "")} onClick={() => setSel((s) => s.map((x, j) => j === i ? !x : x))}>{sel[i] && <Check size={12} />}</button>
                {body}
              </div>
            );
          })}
        </div>
      )}
      <div className="tp-modal-actions">
        <button className="tp-ghostbtn" onClick={onClose}>キャンセル</button>
        <button className="tp-primarybtn" disabled={!rows || rows.length === 0 || !sel.some(Boolean)} onClick={imp}><Plus size={15} /> {kind === "timetable" ? "時間割に取り込む" : kind === "club" ? "部活予定に取り込む" : "予定に取り込む"}</button>
      </div>
    </Modal>
  );
}

function ExcelImportModal({ open, onClose, data, setData, showToast }) {
  const [kind, setKind] = useState("club");
  const [rows, setRows] = useState(null);
  const [err, setErr] = useState("");
  const [fname, setFname] = useState("");
  const fileRef = useRef(null);
  useEffect(() => { if (open) { setRows(null); setErr(""); setFname(""); } }, [open, kind]);

  const parse = (f) => {
    if (!f) return; setErr(""); setRows(null); setFname(f.name);
    const r = new FileReader();
    r.onload = () => {
      try {
        const wb = XLSX.read(new Uint8Array(r.result), { type: "array", cellDates: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const ctxYear = parseInt(String(data.meta.year || "").replace(/\D/g, ""), 10) || new Date().getFullYear();
        let parsed;
        if (kind === "timetable") parsed = parseTimetableGrid(XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" }));
        else if (kind === "event") parsed = parseEventRows(XLSX.utils.sheet_to_json(ws, { defval: "" }), ctxYear);
        else if (kind === "roster") parsed = parseRosterRows(XLSX.utils.sheet_to_json(ws, { defval: "" }));
        else parsed = parseClubRows(XLSX.utils.sheet_to_json(ws, { defval: "" }), ctxYear);
        setRows(parsed);
        if (!parsed.length) setErr("データを認識できませんでした。テンプレートの列名（見出し）に合わせてください。");
      } catch (e) { setErr("読み込みに失敗しました（" + (e.message || "") + "）。"); }
    };
    r.onerror = () => setErr("ファイルを読み込めませんでした。");
    r.readAsArrayBuffer(f);
  };

  const commit = () => {
    if (!rows || !rows.length) return;
    const prev = data;
    setData((d) => {
      if (kind === "event") { const ex = new Set(d.events.map((e) => e.date + "|" + e.title)); return { ...d, events: [...d.events, ...rows.filter((r) => !ex.has(r.date + "|" + r.title)).map((r) => ({ id: uid(), ...r }))] }; }
      if (kind === "club") { const days = { ...(d.club.days || {}) }; rows.forEach((r) => { days[r.date] = r.entry; }); return { ...d, club: { ...d.club, days } }; }
      if (kind === "roster") {
        const rosters = { ...(d.rosters || {}) }; const classes = new Set(d.classes);
        const grouped = {}; rows.forEach((r) => { (grouped[r.klass] = grouped[r.klass] || []).push({ id: uid(), no: r.no, name: r.name, kana: r.kana, memo: r.memo }); });
        Object.keys(grouped).forEach((c) => { rosters[c] = grouped[c]; classes.add(c); });
        return { ...d, rosters, classes: Array.from(classes) };
      }
      const tt = { ...d.timetable }; rows.forEach((r) => { tt[`${r.di}-${r.pi}`] = { subject: r.subject, klass: r.klass, room: "" }; }); return { ...d, timetable: tt };
    });
    showToast && showToast(`${label[kind]}を${rows.length}件取り込みました`, prev);
    onClose();
  };

  const label = { event: "行事", club: "部活", timetable: "時間割", roster: "名簿" };
  return (
    <Modal open={open} title="Excel / CSV から取り込み" onClose={onClose} wide>
      <div className="tp-field"><span>取り込む種類</span>
        <Seg options={[{ v: "club", label: "部活" }, { v: "event", label: "行事" }, { v: "timetable", label: "時間割" }, { v: "roster", label: "名簿" }]} value={kind} onChange={setKind} />
      </div>
      <p className="tp-hint" style={{ marginTop: 0 }}>
        {kind === "club" && "列：日付／内容／場所／時間／備考。内容は「放課後練習・午前練習・休み・（大会名）」で判定します。"}
        {kind === "event" && "列：日付／行事名／種類（行事・テスト・会議・部活・その他）／時刻。"}
        {kind === "timetable" && "先頭行に曜日（月〜金）、先頭列に限。セルは「教科 クラス」（例：数学 2-1）。"}
        {kind === "roster" && "列：クラス／番号／氏名／ふりがな／備考。クラスごとに名簿へ取り込みます。"}
      </p>
      <div className="tp-toolbar" style={{ justifyContent: "flex-start", gap: 8 }}>
        <button className="tp-ghostbtn" onClick={() => fileRef.current?.click()}><Upload size={14} /> ファイルを選ぶ（.xlsx / .csv）</button>
        <button className="tp-ghostbtn" onClick={() => downloadTemplate(kind)}><Download size={14} /> テンプレート</button>
      </div>
      <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={(e) => parse(e.target.files?.[0])} />
      {fname && <p className="tp-hint">選択：{fname}</p>}
      {err && <div className="tp-error">{err}</div>}
      {rows && rows.length > 0 && (
        <>
          <div className="tp-result-head">{label[kind]}：{rows.length}件を認識</div>
          <ul className="tp-result-list">
            {rows.slice(0, 8).map((r, i) => (
              <li key={i}>
                {kind === "timetable"
                  ? <span className="tp-res-title">{DAY_LABELS[r.di]}曜 {r.pi + 1}限：{r.subject} {r.klass}</span>
                  : kind === "roster"
                  ? <span className="tp-res-title">{r.klass}　{r.no}　{r.name}{r.kana ? `（${r.kana}）` : ""}</span>
                  : <><span className="tp-res-date">{(r.date || "").slice(5)}</span><span className="tp-res-title">{kind === "event" ? r.title : (r.entry.kind === "match" ? "【大会】" + r.entry.matchName : r.entry.kind === "off" ? "休み" : (r.entry.session === "am" ? "午前" : r.entry.session === "pm" ? "午後" : "放課後") + "練習")}</span></>}
              </li>
            ))}
            {rows.length > 8 && <li className="tp-hint">…ほか {rows.length - 8} 件</li>}
          </ul>
          <div className="tp-modal-actions"><button className="tp-ghostbtn" onClick={() => setRows(null)}>やり直す</button><button className="tp-primarybtn" onClick={commit}><Check size={15} /> {rows.length}件を取り込む</button></div>
        </>
      )}
      <p className="tp-hint">※画像からの読み取りはAI処理のため通信が必要です（Claudeアプリ・オンライン時）。オフラインではExcel/CSVをご利用ください。</p>
    </Modal>
  );
}

/* ============================================================
   Modal
   ============================================================ */
function Modal({ open, title, onClose, children, wide }) {
  if (!open) return null;
  return (
    <div className="tp-modal-back" onClick={onClose}>
      <div className={"tp-modal" + (wide ? " wide" : "")} onClick={(e) => e.stopPropagation()}>
        <div className="tp-modal-head">
          <h3>{title}</h3>
          <button className="tp-iconbtn" onClick={onClose} aria-label="閉じる"><X size={18} /></button>
        </div>
        <div className="tp-modal-body">{children}</div>
      </div>
    </div>
  );
}

/* ============================================================
   SketchPad — 手書きメモ
   ============================================================ */
function SketchPad({ dateKey }) {
  const canvasRef = useRef(null);
  const drawing = useRef(false);
  const last = useRef(null);
  const [color, setColor] = useState("#24323C");

  const getCtx = () => canvasRef.current?.getContext("2d");

  useEffect(() => {
    const c = canvasRef.current;
    if (!c) return;
    const ctx = c.getContext("2d");
    ctx.clearRect(0, 0, c.width, c.height);
    let cancelled = false;
    (async () => {
      const r = await loadStore("sketch:" + dateKey);
      if (!cancelled && r && r.value) {
        const img = new Image();
        img.onload = () => ctx.drawImage(img, 0, 0, c.width, c.height);
        img.src = r.value;
      }
    })();
    return () => { cancelled = true; };
  }, [dateKey]);

  const pos = (e) => {
    const c = canvasRef.current;
    const rect = c.getBoundingClientRect();
    return { x: (e.clientX - rect.left) * (c.width / rect.width), y: (e.clientY - rect.top) * (c.height / rect.height) };
  };
  const down = (e) => { e.preventDefault(); drawing.current = true; last.current = pos(e); };
  const move = (e) => {
    if (!drawing.current) return;
    e.preventDefault();
    const ctx = getCtx(); const p = pos(e);
    ctx.strokeStyle = color; ctx.lineWidth = color === "#FFFFFF" ? 18 : 2.4;
    ctx.lineCap = "round"; ctx.lineJoin = "round";
    ctx.beginPath(); ctx.moveTo(last.current.x, last.current.y); ctx.lineTo(p.x, p.y); ctx.stroke();
    last.current = p;
  };
  const up = () => { drawing.current = false; };
  const save = async () => { const c = canvasRef.current; await saveStore("sketch:" + dateKey, c.toDataURL("image/png")); };
  const clear = async () => { const c = canvasRef.current; getCtx().clearRect(0, 0, c.width, c.height); await delStore("sketch:" + dateKey); };

  return (
    <div className="tp-sketch">
      <div className="tp-sketch-tools">
        {["#24323C", "#3E9BC9", "#E8845B", "#4F9E86"].map((cc) => (
          <button key={cc} className={"tp-pen" + (color === cc ? " on" : "")} style={{ background: cc }} onClick={() => setColor(cc)} aria-label="ペン" />
        ))}
        <button className={"tp-pen eraser" + (color === "#FFFFFF" ? " on" : "")} onClick={() => setColor("#FFFFFF")} aria-label="消しゴム"><Eraser size={14} /></button>
        <span className="tp-sketch-spacer" />
        <button className="tp-ghostbtn" onClick={save}><Save size={13} /> 保存</button>
        <button className="tp-ghostbtn" onClick={clear}><Trash2 size={13} /> 消去</button>
      </div>
      <canvas
        ref={canvasRef} width={920} height={240} className="tp-canvas"
        onPointerDown={down} onPointerMove={move} onPointerUp={up} onPointerLeave={up}
      />
      <p className="tp-hint">Apple Pencil や指で手書き。「保存」でこの日付に記録されます。</p>
    </div>
  );
}

/* ============================================================
   Logo — Teaching Partner for YOU
   ============================================================ */
function Logo({ size = 40 }) {
  const gid = useMemo(() => "tpg" + Math.random().toString(36).slice(2, 7), []);
  return (
    <svg width={size} height={size} viewBox="0 0 48 48" fill="none" xmlns="http://www.w3.org/2000/svg" aria-label="Teaching Partner for YOU" style={{ display: "block" }}>
      <defs>
        <linearGradient id={gid} x1="4" y1="4" x2="44" y2="44" gradientUnits="userSpaceOnUse">
          <stop stopColor="#4AA6D0" /><stop offset="1" stopColor="#2C7CA6" />
        </linearGradient>
      </defs>
      <rect x="4" y="4" width="40" height="40" rx="11.5" fill={`url(#${gid})`} />
      <g transform="matrix(0.02781,0,0,-0.02781,9.9995,35.1386)"><path d="M621 846C611 746 592 645 566 552C535 581 490 618 490 618L443 555H411C467 624 512 694 547 760C572 756 582 761 588 772L457 835C444 801 429 765 410 729L355 777L307 713H304V807C331 811 339 821 341 835L194 847V713H67L75 685H194V555H27L35 527H288C262 491 234 457 204 423H73L82 394H178C129 342 75 295 18 254L27 244C112 285 188 337 255 394H356C343 372 327 345 310 323L251 328V235C159 223 83 214 39 210L88 89C99 92 109 101 115 113L251 156V49C251 37 247 32 232 32C212 32 108 39 108 39V25C157 17 178 5 194 -12C209 -28 214 -54 217 -89C345 -77 362 -35 362 43V192C439 219 502 242 553 261L551 275L362 249V291C383 294 393 301 395 315L361 318C401 338 441 361 473 381C493 383 504 385 512 393L414 478L360 423H287C324 456 357 491 388 527H549L560 528C537 450 510 379 481 320L494 312C539 354 579 404 614 461C627 372 645 289 672 215C607 99 507 1 357 -77L363 -87C519 -39 632 31 713 119C754 38 810 -30 884 -82C899 -29 931 2 986 13L989 23C900 65 828 122 771 192C848 306 886 442 904 595H955C969 595 979 600 982 611C941 649 871 705 871 705L808 624H693C712 674 729 727 743 784C767 785 778 794 782 807ZM388 685C364 642 338 598 308 555H304V685ZM708 289C675 350 650 418 632 494C650 525 666 559 681 595H775C767 485 747 382 708 289Z" fill="#ffffff" /></g>
    </svg>
  );
}
/* ============================================================
   TabBar
   ============================================================ */
const TABS = [
  { id: "today", label: "今日", icon: Home },
  { id: "cal", label: "カレンダー", icon: CalendarDays },
  { id: "classes", label: "授業", icon: BookOpen },
  { id: "club", label: "部活", icon: Dumbbell },
  { id: "roster", label: "名簿", icon: Users },
  { id: "duties", label: "校務", icon: ClipboardList },
  { id: "meetings", label: "職員会議", icon: FileText },
];
function TabBar({ active, onChange, onSettings }) {
  return (
    <nav className="tp-tabs">
      <div className="tp-brand"><Logo size={42} /></div>
      {TABS.map((t) => {
        const Icon = t.icon;
        return (
          <button key={t.id} className={"tp-tab" + (active === t.id ? " on" : "")} onClick={() => onChange(t.id)}>
            <Icon size={20} /><span>{t.label}</span>
          </button>
        );
      })}
      <button className="tp-tab settings" onClick={onSettings}><Settings size={20} /><span>設定</span></button>
    </nav>
  );
}

/* ============================================================
   TODAY — 一日の全予定を統合表示
   ============================================================ */
/* 横スワイプでページ送り（縦スクロール・端スワイプ・タップは邪魔しない） */
function useSwipe(onLeft, onRight) {
  const ref = useRef({ x: 0, y: 0, t: 0, active: false });
  const onTouchStart = (e) => {
    const t = e.touches && e.touches[0]; if (!t) return;
    const w = (typeof window !== "undefined" && window.innerWidth) || 400;
    if (t.clientX < 24 || t.clientX > w - 24) { ref.current.active = false; return; } // iOS戻る等と競合回避
    ref.current = { x: t.clientX, y: t.clientY, t: Date.now(), active: true };
  };
  const onTouchEnd = (e) => {
    if (!ref.current.active) return; ref.current.active = false;
    const t = e.changedTouches && e.changedTouches[0]; if (!t) return;
    const dx = t.clientX - ref.current.x, dy = t.clientY - ref.current.y, dt = Date.now() - ref.current.t;
    if (dt > 800) return; if (Math.abs(dx) < 55) return; if (Math.abs(dx) < Math.abs(dy) * 1.4) return;
    if (dx < 0) { onLeft && onLeft(); } else { onRight && onRight(); }
  };
  return { onTouchStart, onTouchEnd };
}

function TodayView({ data, setData, selDate, setSelDate, user }) {
  const idx = jsDayToIdx(selDate.getDay());
  const key = ymd(selDate);
  const isToday = sameDay(selDate, new Date());

  const agenda = useMemo(() => {
    const items = [];
    const termStart = currentTerm(data, selDate).start;
    data.routine.forEach((r) => items.push({ time: r.time, kind: "routine", title: r.title }));
    if (idx >= 0) {
      data.periods.forEach((p, pi) => {
        const cell = data.timetable[`${idx}-${pi}`];
        if (cell && cell.subject) {
          const log = data.lessonLog[`${key}-${pi}`] || {};
          const seq = cell.klass ? lessonOrdinal(data, cell.subject, cell.klass, key, pi, termStart) : null;
          items.push({ time: p.start, kind: "lesson", periodIdx: pi, period: p.label, end: p.end, subject: cell.subject, klass: cell.klass, room: cell.room, done: !!log.done, topic: log.topic || "", seq });
        }
      });
      const cd = clubDayDisplay(data, key, idx);
      if (cd) items.push({ time: cd.time ? cd.time.split("–")[0] : "", kind: "club", title: cd.kind === "off" ? `${data.club.name}・休み` : cd.kind === "match" ? cd.content : `${data.club.name}（${cd.content}）`, sub: cd.place, note: [cd.time, cd.note].filter(Boolean).join(" / ") });
    }
    data.club.specials.filter((s) => s.date === key).forEach((s) => items.push({ time: s.start, kind: "clubsp", title: s.title, sub: s.place }));
    data.duties.filter((d) => d.dayIdxs.includes(idx)).forEach((d) => items.push({ time: d.time, kind: "duty", title: d.title, sub: d.place, note: d.note }));
    data.events.filter((e) => e.date === key).forEach((e) => items.push({ time: e.time, kind: "event", etype: e.type, title: e.title }));
    return items.sort((a, b) => t2m(a.time) - t2m(b.time));
  }, [data, idx, key]);

  const toggleDone = (pi) => {
    const cell = data.timetable[`${idx}-${pi}`];
    setData((d) => {
      const k = `${key}-${pi}`;
      const cur = d.lessonLog[k] || {};
      const done = !cur.done;
      return { ...d, lessonLog: { ...d.lessonLog, [k]: { ...cur, done, subject: cell.subject, klass: cell.klass } } };
    });
  };
  const [editLesson, setEditLesson] = useState(null); // periodIdx
  const [imp, setImp] = useState(false);

  const dayTodos = data.todos.filter((t) => t.date === key);
  const [newTodo, setNewTodo] = useState("");
  const addTodo = () => { if (!newTodo.trim()) return; setData((d) => ({ ...d, todos: [...d.todos, { id: uid(), date: key, text: newTodo.trim(), done: false }] })); setNewTodo(""); };
  const swipe = useSwipe(() => setSelDate(addDays(selDate, 1)), () => setSelDate(addDays(selDate, -1)));

  return (
    <div className="tp-view" {...swipe}>
      <div className="tp-daynav">
        <button className="tp-iconbtn" onClick={() => setSelDate(addDays(selDate, -1))}><ChevronLeft size={20} /></button>
        <div className="tp-daynav-mid">
          <div className={"tp-daynav-date" + (isToday ? " today" : "")}>
            {selDate.getMonth() + 1}月{selDate.getDate()}日<span className={"tp-wd wd" + selDate.getDay()}>（{WD[selDate.getDay()]}）</span>
          </div>
          {!isToday && <button className="tp-ghostbtn sm" onClick={() => setSelDate(new Date())}>今日へ</button>}
        </div>
        <button className="tp-iconbtn" onClick={() => setSelDate(addDays(selDate, 1))}><ChevronRight size={20} /></button>
      </div>

      <div className="tp-toolbar"><button className="tp-ghostbtn" onClick={() => setImp(true)}><Upload size={14} /> 日報を取り込み</button></div>

      <div className="tp-cardrow">
        <section className="tp-card tp-timeline-card">
          <h4 className="tp-card-title"><Clock size={15} /> 今日の予定</h4>
          {agenda.length === 0 && <p className="tp-empty">予定はありません。日課や時間割を設定すると表示されます。</p>}
          <ul className="tp-timeline">
            {agenda.map((it, i) => (
              <li key={i} className={"tp-tl-item k-" + it.kind}>
                <span className="tp-tl-time">{it.time || "—"}</span>
                <span className="tp-tl-bar" style={{ background: it.kind === "lesson" ? subjColor(data, it.subject) : it.kind === "event" ? (eventMeta[it.etype]?.color) : undefined }} />
                <div className="tp-tl-body">
                  {it.kind === "lesson" ? (
                    <>
                      <div className="tp-tl-main">
                        <button className={"tp-donebox" + (it.done ? " on" : "")} onClick={() => toggleDone(it.periodIdx)} aria-label="実施済み">{it.done && <Check size={12} />}</button>
                        <b>{it.period}限 {it.subject}</b>
                        <span className="tp-chip" style={{ background: subjColor(data, it.subject) }}>{it.klass}</span>
                        {it.seq != null && it.klass && <span className="tp-seqbadge" title="学期の起点から数えた登録授業の通し番号（このクラス）">第{it.seq}時</span>}
                        {it.room && <span className="tp-tl-room"><MapPin size={11} />{it.room}</span>}
                        <button className="tp-linkbtn" onClick={() => setEditLesson(it.periodIdx)}><Pencil size={12} /> 授業計画</button>
                      </div>
                      {it.topic && <div className="tp-tl-topic">{it.topic}</div>}
                    </>
                  ) : (
                    <div className="tp-tl-main">
                      <b>{it.title}</b>
                      {it.kind === "event" && <span className="tp-chip sm" style={{ background: eventMeta[it.etype]?.color }}>{eventMeta[it.etype]?.label}</span>}
                      {(it.kind === "club" || it.kind === "clubsp") && <span className="tp-chip sm" style={{ background: "#E8845B" }}>部活</span>}
                      {it.kind === "duty" && <span className="tp-chip sm" style={{ background: "#8894A0" }}>校務</span>}
                      {it.sub && <span className="tp-tl-room"><MapPin size={11} />{it.sub}</span>}
                      {it.note && <span className="tp-tl-note">{it.note}</span>}
                    </div>
                  )}
                </div>
              </li>
            ))}
          </ul>
        </section>

        <section className="tp-card tp-todo-card">
          <h4 className="tp-card-title"><ClipboardList size={15} /> 今日のやること<span className="tp-cardsub">その日の思いつき・突発の仕事</span></h4>
          <div className="tp-todo-add">
            <input value={newTodo} onChange={(e) => setNewTodo(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addTodo()} placeholder="今日やることをメモ…" />
            <button className="tp-addbtn" onClick={addTodo}><Plus size={16} /></button>
          </div>
          <ul className="tp-todo-list">
            {dayTodos.length === 0 && <li className="tp-empty">タスクなし</li>}
            {dayTodos.map((t) => (
              <li key={t.id} className={t.done ? "done" : ""}>
                <button className={"tp-donebox" + (t.done ? " on" : "")} onClick={() => setData((d) => ({ ...d, todos: d.todos.map((x) => x.id === t.id ? { ...x, done: !x.done } : x) }))}>{t.done && <Check size={12} />}</button>
                <span>{t.text}</span>
                <button className="tp-iconbtn tiny" onClick={() => setData((d) => ({ ...d, todos: d.todos.filter((x) => x.id !== t.id) }))}><Trash2 size={13} /></button>
              </li>
            ))}
          </ul>
        </section>
      </div>

      <section className="tp-card">
        <h4 className="tp-card-title"><ClipboardList size={15} /> メモ</h4>
        <textarea className="tp-daymemo" rows={3} value={data.dayMemo?.[key] || ""} placeholder="この日のメモ（日報の取り込み内容もここに入ります）"
          onChange={(e) => setData((d) => ({ ...d, dayMemo: { ...(d.dayMemo || {}), [key]: e.target.value } }))} />
      </section>

      <section className="tp-card">
        <h4 className="tp-card-title"><Pencil size={15} /> 手書きメモ</h4>
        <SketchPad dateKey={(user || "_") + ":" + key} />
      </section>

      <ImportModal open={imp} kind="daily" ctx={{ date: key }} data={data} setData={setData} onClose={() => setImp(false)} />

      <Modal open={editLesson !== null} title={`授業計画（${selDate.getMonth() + 1}/${selDate.getDate()} ${editLesson !== null ? data.periods[editLesson].label + "限" : ""}）`} onClose={() => setEditLesson(null)}>
        {editLesson !== null && (() => {
          const cell = data.timetable[`${idx}-${editLesson}`] || {};
          const lk = `${key}-${editLesson}`;
          const log = data.lessonLog[lk] || {};
          const setLog = (patch) => setData((d) => ({ ...d, lessonLog: { ...d.lessonLog, [lk]: { ...(d.lessonLog[lk] || {}), subject: cell.subject, klass: cell.klass, ...patch } } }));
          return (
            <>
              <div className="tp-lesson-head"><span className="tp-chip" style={{ background: subjColor(data, cell.subject) }}>{cell.klass}</span> {cell.subject}</div>
              <label className="tp-field"><span>学習内容・進度</span><textarea rows={3} value={log.topic || ""} onChange={(e) => setLog({ topic: e.target.value })} placeholder="例）2章 連立方程式（加減法）／教科書 p.52–55" /></label>
              <label className="tp-field"><span>宿題・連絡</span><input value={log.hw || ""} onChange={(e) => setLog({ hw: e.target.value })} placeholder="例）ワーク p.30" /></label>
              <label className="tp-check"><input type="checkbox" checked={!!log.done} onChange={(e) => setLog({ done: e.target.checked })} /> 実施済みにする（授業数に加算）</label>
              <div className="tp-modal-actions"><span /><button className="tp-primarybtn" onClick={() => setEditLesson(null)}><Check size={15} /> 完了</button></div>
            </>
          );
        })()}
      </Modal>
    </div>
  );
}

/* ============================================================
   WEEK — 週間時間割
   ============================================================ */
function WeekView({ data, setData, selDate, setSelDate, vis, toggleVis, onPrint, onPrintWeekplan }) {
  const cols = data.meta.includeSat ? 6 : 5;
  const monday = startOfWeekMon(selDate);
  const [edit, setEdit] = useState(null); // {dayIdx, periodIdx}
  const [bulk, setBulk] = useState(false);
  const [imp, setImp] = useState(false);

  const weekEvents = useMemo(() => {
    const list = [];
    for (let i = 0; i < 7; i++) {
      const date = addDays(monday, i);
      calItemsForDate(data, date, vis).forEach((it) => list.push({ ...it, date: ymd(date) }));
    }
    return list.sort((a, b) => (a.date + (a.time || "")).localeCompare(b.date + (b.time || "")));
  }, [data, monday, vis]);

  const printWeek = (auto) => onPrint(`${monday.getFullYear()}年 ${monday.getMonth() + 1}/${monday.getDate()}〜 週間予定`, collectRows(data, monday, addDays(monday, 6), vis), vis, auto);
  const swipe = useSwipe(() => setSelDate(addDays(selDate, 7)), () => setSelDate(addDays(selDate, -7)));

  return (
    <div className="tp-view" {...swipe}>
      <div className="tp-daynav">
        <button className="tp-iconbtn" onClick={() => setSelDate(addDays(selDate, -7))}><ChevronLeft size={20} /></button>
        <div className="tp-daynav-mid"><div className="tp-daynav-date">{monday.getMonth() + 1}/{monday.getDate()} 〜 {addDays(monday, cols - 1).getMonth() + 1}/{addDays(monday, cols - 1).getDate()} の週</div>
          <button className="tp-ghostbtn sm" onClick={() => setSelDate(new Date())}>今週へ</button></div>
        <button className="tp-iconbtn" onClick={() => setSelDate(addDays(selDate, 7))}><ChevronRight size={20} /></button>
      </div>

      <CategoryToggles vis={vis} onToggle={toggleVis} />
      <div className="tp-toolbar"><button className="tp-ghostbtn" onClick={() => setBulk(true)}><LayoutGrid size={14} /> 教科をまとめて配置</button><button className="tp-ghostbtn" onClick={() => setImp(true)}><Upload size={14} /> 時間割を取り込み</button><button className="tp-ghostbtn" onClick={() => onPrintWeekplan(ymd(monday))}><BookOpen size={14} /> 週案</button><button className="tp-ghostbtn" onClick={() => printWeek(false)}><Printer size={14} /> 印刷</button><button className="tp-ghostbtn" onClick={() => printWeek(true)}><Download size={14} /> PDF保存</button></div>

      <section className="tp-card tp-tt-card">
        <div className="tp-tt" style={{ gridTemplateColumns: `48px repeat(${cols}, 1fr)` }}>
          <div className="tp-tt-corner">限</div>
          {Array.from({ length: cols }).map((_, di) => {
            const date = addDays(monday, di);
            const isT = sameDay(date, new Date());
            return <div key={di} className={"tp-tt-head" + (isT ? " today" : "")}>{DAY_LABELS[di]}<span className="tp-tt-date">{date.getDate()}</span></div>;
          })}
          {data.periods.map((p, pi) => (
            <React.Fragment key={pi}>
              <div className="tp-tt-period"><b>{p.label}</b><span>{p.start}</span></div>
              {Array.from({ length: cols }).map((_, di) => {
                const cell = data.timetable[`${di}-${pi}`];
                const cdate = addDays(monday, di);
                const seq = cell?.subject && cell?.klass ? lessonOrdinal(data, cell.subject, cell.klass, ymd(cdate), pi, currentTerm(data, cdate).start) : null;
                return (
                  <button key={di} className="tp-tt-cell" onClick={() => setEdit({ dayIdx: di, periodIdx: pi })}
                    style={cell?.subject ? { background: subjColor(data, cell.subject) + "1A", borderLeft: `3px solid ${subjColor(data, cell.subject)}` } : {}}>
                    {cell?.subject ? (<><span className="tp-tt-sub" style={{ color: subjColor(data, cell.subject) }}>{cell.subject}</span><span className="tp-tt-klass">{cell.klass}</span>{cell.room && <span className="tp-tt-room">{cell.room}</span>}{seq != null && <span className="tp-tt-seq">第{seq}時</span>}</>) : <span className="tp-tt-plus">+</span>}
                  </button>
                );
              })}
            </React.Fragment>
          ))}
          {/* 放課後：部活動 */}
          <div className="tp-tt-period after"><b>放課後</b></div>
          {Array.from({ length: cols }).map((_, di) => {
            const date = addDays(monday, di);
            const cd = clubDayDisplay(data, ymd(date), di);
            return <div key={di} className="tp-tt-cell after">{cd && cd.kind !== "off" ? <><span className="tp-tt-club" style={cd.kind === "match" ? { color: "#D9534F" } : {}}><Dumbbell size={11} /> {cd.kind === "match" ? "大会" : "部活"}</span><span className="tp-tt-room">{cd.time || cd.place}</span></> : <span className="tp-tt-off">{cd ? "休み" : "—"}</span>}</div>;
          })}
        </div>
      </section>

      <section className="tp-card">
        <h4 className="tp-card-title"><CalendarDays size={15} /> 今週の予定</h4>
        {weekEvents.length === 0 ? <p className="tp-empty">この表示カテゴリの予定はありません。</p> : (
          <ul className="tp-weeklist">
            {weekEvents.map((e, i) => { const d = parseYmd(e.date); return (
              <li key={i}><span className="tp-wl-date">{d.getMonth() + 1}/{d.getDate()}<small>（{WD[d.getDay()]}）</small></span>
                <span className="tp-cat-dot" style={{ background: e.color }} />
                <span className="tp-wl-title">{e.title}</span>{e.time && <span className="tp-wl-time">{e.time}</span>}</li>
            ); })}
          </ul>
        )}
      </section>

      <Modal open={!!edit} title={edit ? `${DAY_LABELS[edit.dayIdx]}曜 ${data.periods[edit.periodIdx].label}限` : ""} onClose={() => setEdit(null)}>
        {edit && (() => {
          const k = `${edit.dayIdx}-${edit.periodIdx}`;
          const cell = data.timetable[k] || { subject: "", klass: "", room: "" };
          const set = (patch) => setData((d) => ({ ...d, timetable: { ...d.timetable, [k]: { ...cell, ...patch } } }));
          return (
            <>
              <label className="tp-field"><span>教科</span>
                <select value={cell.subject} onChange={(e) => set({ subject: e.target.value })}>
                  <option value="">（なし）</option>
                  {data.subjects.map((s) => <option key={s.name} value={s.name}>{s.name}</option>)}
                </select></label>
              <label className="tp-field"><span>クラス</span>
                <select value={cell.klass} onChange={(e) => set({ klass: e.target.value })}>
                  <option value="">—</option>
                  {data.classes.map((c) => <option key={c} value={c}>{c}</option>)}
                </select></label>
              <label className="tp-field"><span>教室</span><input value={cell.room} onChange={(e) => set({ room: e.target.value })} placeholder="例）理科室" /></label>
              <div className="tp-modal-actions">
                <button className="tp-dangerbtn" style={{ marginTop: 0 }} onClick={() => { setData((d) => { const t = { ...d.timetable }; delete t[k]; return { ...d, timetable: t }; }); setEdit(null); }}><Trash2 size={14} /> このコマを空にする</button>
                <button className="tp-primarybtn" onClick={() => setEdit(null)}><Check size={15} /> 完了</button>
              </div>
            </>
          );
        })()}
      </Modal>
      <ImportModal open={imp} kind="timetable" ctx={{ subjects: data.subjects.map((s) => s.name), classes: data.classes }} data={data} setData={setData} onClose={() => setImp(false)} />
      <BulkTimetableModal open={bulk} cols={cols} data={data} setData={setData} onClose={() => setBulk(false)} />
    </div>
  );
}

/* 週間：教科ごとに曜日×時限をまとめて配置（毎週共通の時間割テンプレに反映） */
function BulkTimetableModal({ open, cols, data, setData, onClose }) {
  const [subject, setSubject] = useState((data.subjects[0] || {}).name || "");
  const [klass, setKlass] = useState("");
  const [room, setRoom] = useState("");
  const [picked, setPicked] = useState({}); // "di-pi": true
  useEffect(() => { if (open) { setSubject((data.subjects[0] || {}).name || ""); setKlass(""); setRoom(""); setPicked({}); } }, [open]);
  const toggle = (di, pi) => setPicked((p) => { const k = `${di}-${pi}`; const n = { ...p }; if (n[k]) delete n[k]; else n[k] = true; return n; });
  const keys = Object.keys(picked);
  const apply = () => { if (!keys.length) return; setData((d) => { const tt = { ...d.timetable }; keys.forEach((k) => { tt[k] = { subject, klass, room }; }); return { ...d, timetable: tt }; }); onClose(); };
  const clear = () => { if (!keys.length) return; setData((d) => { const tt = { ...d.timetable }; keys.forEach((k) => delete tt[k]); return { ...d, timetable: tt }; }); onClose(); };
  return (
    <Modal open={open} title="教科をまとめて配置" onClose={onClose} wide>
      <p className="tp-hint" style={{ marginTop: 0 }}>教科（クラス・教室）を選び、下の表で入れたいコマを<b>タップして複数選択</b> →「まとめて設定」。毎週共通の時間割に反映されます。</p>
      <div className="tp-field-row">
        <label className="tp-field"><span>教科</span>
          <select value={subject} onChange={(e) => setSubject(e.target.value)}>{data.subjects.map((s) => <option key={s.name} value={s.name}>{s.name}</option>)}</select></label>
        <label className="tp-field"><span>クラス</span>
          <select value={klass} onChange={(e) => setKlass(e.target.value)}><option value="">—</option>{data.classes.map((c) => <option key={c} value={c}>{c}</option>)}</select></label>
        <label className="tp-field"><span>教室</span><input value={room} onChange={(e) => setRoom(e.target.value)} placeholder="任意" /></label>
      </div>
      <div className="tp-bulk-grid" style={{ gridTemplateColumns: `40px repeat(${cols}, 1fr)` }}>
        <div className="tp-bulk-corner">限</div>
        {Array.from({ length: cols }).map((_, di) => <div key={di} className="tp-bulk-head">{DAY_LABELS[di]}</div>)}
        {data.periods.map((p, pi) => (
          <React.Fragment key={pi}>
            <div className="tp-bulk-period">{p.label}</div>
            {Array.from({ length: cols }).map((_, di) => { const k = `${di}-${pi}`; const cur = data.timetable[k]; const on = !!picked[k]; return (
              <button key={di} className={"tp-bulk-cell" + (on ? " on" : "")} onClick={() => toggle(di, pi)}>
                {on ? <Check size={14} /> : (cur?.subject ? <span className="tp-bulk-cur">{cur.subject}<br />{cur.klass}</span> : "")}
              </button>
            ); })}
          </React.Fragment>
        ))}
      </div>
      <p className="tp-hint">選択中：{keys.length}コマ</p>
      <div className="tp-modal-actions">
        <button className="tp-dangerbtn" style={{ marginTop: 0 }} disabled={!keys.length} onClick={clear}><Trash2 size={14} /> 選択を空にする</button>
        <button className="tp-primarybtn" disabled={!keys.length || !subject} onClick={apply}><Check size={15} /> まとめて設定（{keys.length}）</button>
      </div>
    </Modal>
  );
}

/* カレンダーセルの予定表示（月間・3か月で共通） */
function CalItems({ evs, mini, max }) {
  const cap = max || 3;
  const clubs = evs.filter((e) => e.club);
  const others = evs.filter((e) => !e.club);
  return (
    <div className={"tp-cal-evs" + (mini ? " mini" : "")}>
      {clubs.map((e, j) => (
        <div key={"c" + j} className={"tp-cal-club" + (e.weekly ? " weekly" : "")} style={{ borderColor: e.color }}>
          {e.lines.filter(Boolean).map((ln, li) => <span key={li} className={"tp-cal-cline l" + li}>{ln}</span>)}
        </div>
      ))}
      {others.slice(0, cap).map((e, j) => <span key={"o" + j} className={"tp-cal-ev" + (e.faint ? " faint" : "")} style={{ background: e.color }}>{e.time ? e.time + " " : ""}{e.title}</span>)}
      {others.length > cap && <span className="tp-cal-more">+{others.length - cap}</span>}
    </div>
  );
}

/* ============================================================
   MONTH — 月間カレンダー
   ============================================================ */
function MonthView({ data, setData, selDate, setSelDate, vis, toggleVis, onPrint }) {
  const [cursor, setCursor] = useState(new Date(selDate.getFullYear(), selDate.getMonth(), 1));
  const [addOpen, setAddOpen] = useState(null); // date string
  const [imp, setImp] = useState(false);
  const first = new Date(cursor.getFullYear(), cursor.getMonth(), 1);
  const startPad = first.getDay();
  const daysInMonth = new Date(cursor.getFullYear(), cursor.getMonth() + 1, 0).getDate();
  const cells = [];
  for (let i = 0; i < startPad; i++) cells.push(null);
  for (let d = 1; d <= daysInMonth; d++) cells.push(new Date(cursor.getFullYear(), cursor.getMonth(), d));
  while (cells.length % 7 !== 0) cells.push(null);

  const swipeM = useSwipe(() => setCursor(new Date(cursor.getFullYear(), cursor.getMonth() + 1, 1)), () => setCursor(new Date(cursor.getFullYear(), cursor.getMonth() - 1, 1)));
  return (
    <div className="tp-view" {...swipeM}>
      <div className="tp-daynav">
        <button className="tp-iconbtn" onClick={() => setCursor(new Date(cursor.getFullYear(), cursor.getMonth() - 1, 1))}><ChevronLeft size={20} /></button>
        <div className="tp-daynav-mid"><div className="tp-daynav-date">{cursor.getFullYear()}年 {cursor.getMonth() + 1}月</div>
          <button className="tp-ghostbtn sm" onClick={() => setCursor(new Date())}>今月へ</button></div>
        <button className="tp-iconbtn" onClick={() => setCursor(new Date(cursor.getFullYear(), cursor.getMonth() + 1, 1))}><ChevronRight size={20} /></button>
      </div>
      <CategoryToggles vis={vis} onToggle={toggleVis} />
      <div className="tp-toolbar">
        <button className="tp-ghostbtn" onClick={() => setImp(true)}><Upload size={14} /> 月間予定表を取り込み</button>
        <button className="tp-ghostbtn" onClick={() => onPrint(`${cursor.getFullYear()}年 ${cursor.getMonth() + 1}月 行事予定`, collectRows(data, new Date(cursor.getFullYear(), cursor.getMonth(), 1), new Date(cursor.getFullYear(), cursor.getMonth() + 1, 0), vis), vis, false)}><Printer size={14} /> 印刷</button>
        <button className="tp-ghostbtn" onClick={() => onPrint(`${cursor.getFullYear()}年 ${cursor.getMonth() + 1}月 行事予定`, collectRows(data, new Date(cursor.getFullYear(), cursor.getMonth(), 1), new Date(cursor.getFullYear(), cursor.getMonth() + 1, 0), vis), vis, true)}><Download size={14} /> PDF保存</button>
      </div>
      <section className="tp-card">
        <div className="tp-cal">
          {WD.map((w, i) => <div key={w} className={"tp-cal-wd wd" + i}>{w}</div>)}
          {cells.map((date, i) => {
            if (!date) return <div key={i} className="tp-cal-cell empty" />;
            const isT = sameDay(date, new Date());
            const evs = calItemsForDate(data, date, vis);
            return (
              <div key={i} className={"tp-cal-cell wd" + date.getDay() + (isT ? " today" : "")} onClick={() => setSelDate(date)} onDoubleClick={() => setAddOpen(ymd(date))}>
                <div className="tp-cal-num">{date.getDate()}<button className="tp-cal-add" onClick={(e) => { e.stopPropagation(); setAddOpen(ymd(date)); }}><Plus size={11} /></button></div>
                <CalItems evs={evs} />
              </div>
            );
          })}
        </div>
        <p className="tp-hint">日付をタップで「今日」表示、＋またはダブルタップで予定を追加。上のバーで表示カテゴリを切替（スワイプ可）。</p>
      </section>

      <AddEventModal open={!!addOpen} date={addOpen} onClose={() => setAddOpen(null)} data={data} setData={setData} />
      <ImportModal open={imp} kind="month" ctx={{ year: cursor.getFullYear(), month: cursor.getMonth() + 1 }} data={data} setData={setData} onClose={() => setImp(false)} />
    </div>
  );
}

function AddEventModal({ open, date, onClose, data, setData }) {
  const [title, setTitle] = useState(""); const [type, setType] = useState("school"); const [time, setTime] = useState("");
  useEffect(() => { if (open) { setTitle(""); setType("school"); setTime(""); } }, [open, date]);
  const dayEvents = data.events.filter((e) => e.date === date);
  const add = () => { if (!title.trim()) return; setData((d) => ({ ...d, events: [...d.events, { id: uid(), date, title: title.trim(), type, time }] })); setTitle(""); setTime(""); };
  const d = date ? parseYmd(date) : null;
  return (
    <Modal open={open} title={d ? `${d.getMonth() + 1}月${d.getDate()}日（${WD[d.getDay()]}）の予定` : ""} onClose={onClose}>
      <ul className="tp-mini-list">
        {dayEvents.length === 0 && <li className="tp-empty">まだ予定はありません</li>}
        {dayEvents.map((e) => <li key={e.id}><span className="tp-chip sm" style={{ background: eventMeta[e.type]?.color }}>{eventMeta[e.type]?.label}</span><span>{e.time} {e.title}</span><button className="tp-iconbtn tiny" onClick={() => setData((dd) => ({ ...dd, events: dd.events.filter((x) => x.id !== e.id) }))}><Trash2 size={13} /></button></li>)}
      </ul>
      <div className="tp-divider" />
      <label className="tp-field"><span>内容</span><input value={title} onChange={(e) => setTitle(e.target.value)} placeholder="例）避難訓練" /></label>
      <div className="tp-field-row">
        <label className="tp-field"><span>種類</span><select value={type} onChange={(e) => setType(e.target.value)}>{Object.entries(eventMeta).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}</select></label>
        <label className="tp-field"><span>時刻</span><input value={time} onChange={(e) => setTime(e.target.value)} placeholder="任意 13:30" /></label>
      </div>
      <div className="tp-modal-actions">
        <button className="tp-ghostbtn" onClick={onClose}>閉じる</button>
        <button className="tp-primarybtn" onClick={add}><Plus size={15} /> 追加</button>
      </div>
    </Modal>
  );
}

/* ============================================================
   QUARTER — 3か月間（ミニカレンダー×3）
   ============================================================ */
function MiniMonth({ data, year, month, vis, onPick }) {
  const first = new Date(year, month, 1);
  const startPad = first.getDay();
  const dim = new Date(year, month + 1, 0).getDate();
  const cells = [];
  for (let i = 0; i < startPad; i++) cells.push(null);
  for (let d = 1; d <= dim; d++) cells.push(new Date(year, month, d));
  while (cells.length % 7 !== 0) cells.push(null);
  return (
    <div className="tp-mini">
      <div className="tp-mini-title">{year}年 {month + 1}月</div>
      <div className="tp-mini-grid">
        {WD.map((w, i) => <div key={w} className={"tp-mini-wd wd" + i}>{w}</div>)}
        {cells.map((date, i) => {
          if (!date) return <div key={i} className="tp-mini-cell empty" />;
          const isT = sameDay(date, new Date());
          const items = calItemsForDate(data, date, vis);
          return (
            <button key={i} className={"tp-mini-cell wd" + date.getDay() + (isT ? " today" : "")} onClick={() => onPick(date)}>
              <span className="tp-mini-num">{date.getDate()}</span>
              <CalItems evs={items} mini max={2} />
            </button>
          );
        })}
      </div>
    </div>
  );
}
function QuarterView({ data, selDate, setSelDate, vis, toggleVis, onPrint, onPrintCal, count = 3 }) {
  const [base, setBase] = useState(new Date(selDate.getFullYear(), selDate.getMonth(), 1));
  const months = Array.from({ length: count }, (_, n) => new Date(base.getFullYear(), base.getMonth() + n, 1));
  const last = months[months.length - 1];
  const printQ = (auto) => onPrintCal(`${months[0].getFullYear()}年 ${months[0].getMonth() + 1}月〜${last.getMonth() + 1}月 予定表`, months, vis, auto);
  const swipeQ = useSwipe(() => setBase(new Date(base.getFullYear(), base.getMonth() + count, 1)), () => setBase(new Date(base.getFullYear(), base.getMonth() - count, 1)));
  return (
    <div className="tp-view" {...swipeQ}>
      <div className="tp-daynav">
        <button className="tp-iconbtn" onClick={() => setBase(new Date(base.getFullYear(), base.getMonth() - count, 1))}><ChevronLeft size={20} /></button>
        <div className="tp-daynav-mid"><div className="tp-daynav-date">{months[0].getMonth() + 1}月 〜 {last.getMonth() + 1}月</div>
          <button className="tp-ghostbtn sm" onClick={() => setBase(new Date())}>今月から</button></div>
        <button className="tp-iconbtn" onClick={() => setBase(new Date(base.getFullYear(), base.getMonth() + count, 1))}><ChevronRight size={20} /></button>
      </div>
      <CategoryToggles vis={vis} onToggle={toggleVis} />
      <div className="tp-toolbar"><button className="tp-ghostbtn" onClick={() => printQ(false)}><Printer size={14} /> 印刷</button><button className="tp-ghostbtn" onClick={() => printQ(true)}><Download size={14} /> PDF保存</button></div>
      <div className={"tp-quarter" + (count >= 6 ? " six" : "")}>
        {months.map((m, i) => <section key={i} className="tp-card tp-mini-card"><MiniMonth data={data} year={m.getFullYear()} month={m.getMonth()} vis={vis} onPick={setSelDate} /></section>)}
      </div>
      <p className="tp-hint">{count}か月をまとめて俯瞰。日付タップで「今日」表示。ドットは表示カテゴリの予定です。</p>
    </div>
  );
}

/* ============================================================
   YEAR — 年間行事予定（4月〜3月）
   ============================================================ */
function YearView({ data, setData, setSelDate, setTab, vis, toggleVis, onPrint }) {
  const now = new Date();
  const [startYear, setStartYear] = useState(now.getMonth() >= 3 ? now.getFullYear() : now.getFullYear() - 1);
  const [imp, setImp] = useState(false);
  const months = Array.from({ length: 12 }).map((_, i) => new Date(startYear, 3 + i, 1)); // 4月始まり

  const yearItems = useMemo(() => {
    const map = {};
    months.forEach((mo) => {
      const y = mo.getFullYear(), m = mo.getMonth(), dim = new Date(y, m + 1, 0).getDate(); const list = [];
      for (let d = 1; d <= dim; d++) { const date = new Date(y, m, d); calItemsForDate(data, date, vis).filter((it) => !it.faint && !(it.club && !it.special)).forEach((it) => list.push({ ...it, day: d, wd: date.getDay() })); }
      map[`${y}-${m}`] = list;
    });
    return map;
  }, [data, vis, startYear]);
  const itemsInMonth = (y, m) => yearItems[`${y}-${m}`] || [];
  const printYear = (auto) => onPrint(`${startYear}年度 年間行事予定`, collectRows(data, new Date(startYear, 3, 1), new Date(startYear + 1, 2, 31), vis), vis, auto);
  const swipeY = useSwipe(() => setStartYear(startYear + 1), () => setStartYear(startYear - 1));

  return (
    <div className="tp-view" {...swipeY}>
      <div className="tp-daynav">
        <button className="tp-iconbtn" onClick={() => setStartYear(startYear - 1)}><ChevronLeft size={20} /></button>
        <div className="tp-daynav-mid"><div className="tp-daynav-date">{startYear}年度（{startYear}.4 〜 {startYear + 1}.3）</div>
          <button className="tp-ghostbtn sm" onClick={() => setStartYear(now.getMonth() >= 3 ? now.getFullYear() : now.getFullYear() - 1)}>今年度</button></div>
        <button className="tp-iconbtn" onClick={() => setStartYear(startYear + 1)}><ChevronRight size={20} /></button>
      </div>
      <CategoryToggles vis={vis} onToggle={toggleVis} />
      <div className="tp-toolbar">
        <button className="tp-ghostbtn" onClick={() => setImp(true)}><Upload size={14} /> 年間予定表を取り込み</button>
        <button className="tp-ghostbtn" onClick={() => printYear(false)}><Printer size={14} /> 印刷</button>
        <button className="tp-ghostbtn" onClick={() => printYear(true)}><Download size={14} /> PDF保存</button>
      </div>
      <div className="tp-year">
        {months.map((m, i) => {
          const list = itemsInMonth(m.getFullYear(), m.getMonth());
          return (
            <section key={i} className="tp-card tp-year-month">
              <button className="tp-year-head" onClick={() => { setSelDate(new Date(m.getFullYear(), m.getMonth(), 1)); setTab("month"); }}>
                {m.getMonth() + 1}月<ChevronRight size={14} />
              </button>
              <ul className="tp-year-list">
                {list.length === 0 && <li className="tp-year-empty">—</li>}
                {list.map((it, j) => (
                  <li key={j} onClick={() => { setSelDate(new Date(m.getFullYear(), m.getMonth(), it.day)); setTab("today"); }}>
                    <span className={"tp-year-day wd" + it.wd}>{it.day}</span>
                    <span className="tp-cat-dot" style={{ background: it.color }} />
                    <span className="tp-year-title">{it.title}</span>
                  </li>
                ))}
              </ul>
            </section>
          );
        })}
      </div>
      <p className="tp-hint">年度の行事を一覧。月ラベルで月間へ、行を押すとその日の「今日」へ。表示カテゴリはバーで切替。</p>
      <ImportModal open={imp} kind="year" ctx={{ startYear }} data={data} setData={setData} onClose={() => setImp(false)} />
    </div>
  );
}

/* ============================================================
   授業：教科書パース ＋ 年間計画ジェネレーター
   ============================================================ */
const PHASES = ["新出文法理解", "教科書概要理解", "本文暗唱", "応用"];
const TEST_TYPES = ["定期", "単元", "単語", "文法"];
const testTypeColor = (t) => (t === "定期" ? "#E0A64B" : t === "単元" ? "#3E9BC9" : t === "単語" ? "#8E7CC3" : "#4F9E86");

function parseTextbookText(text) {
  const lines = String(text || "").split(/\r?\n/).map((l) => zen2han(l).trim());
  const units = [];
  let cur = null;
  const rangeRe = /(?:pp?\.?\s*)?([0-9]{1,3})\s*(?:[〜~\-–―ー]|から|to)\s*([0-9]{1,3})\s*(?:ページ|頁|p\.?)?/i;
  const singleRe = /(?:pp?\.?\s*([0-9]{1,3}))|(?:([0-9]{1,3})\s*(?:ページ|頁))/i;
  const takeRange = (s) => { let m = s.match(rangeRe); if (m) return [Number(m[1]), Number(m[2])]; m = s.match(singleRe); if (m) { const p = Number(m[1] || m[2]); return [p, p]; } return null; };
  const stripRange = (s) => { const m = s.match(/[（(]?\s*(?:pp?\.?\s*)?[0-9]{1,3}\s*(?:[〜~\-–―ー]|から|to)\s*[0-9]{1,3}\s*(?:ページ|頁|p\.?)?\s*[)）]?/i) || s.match(/[（(]?\s*(?:pp?\.?\s*[0-9]{1,3}|[0-9]{1,3}\s*(?:ページ|頁))\s*[)）]?/i); return m ? s.replace(m[0], "").trim() : s; };
  const headRe = /^[◆●■◇・\s]*((?:第\s*[0-9]+\s*(?:章|単元|節|課|部|回|時))|(?:(?:PROGRAM|Program|Lesson|LESSON|Unit|UNIT|Chapter|CHAPTER|Part|PART|Section|SECTION)\s+[0-9]+)|(?:[0-9]+\s*(?:章|単元|節|課)))\s*[\t 　:：.．]*\s*(.*)$/;
  for (const ln of lines) {
    if (!ln) continue;
    const mh = ln.match(headRe);
    if (mh) {
      const label = mh[1].replace(/\s+/g, " ").replace(/\s+([章単元節課部回時])/, "$1").trim();
      let rest = (mh[2] || "").trim();
      const r = takeRange(rest); let pf = null, pt = null;
      if (r) { [pf, pt] = r; rest = stripRange(rest); }
      rest = rest.replace(/^[　\s:：.．\-]+/, "").replace(/[（(]\s*[)）]$/, "").replace(/[◆◇■●▶►◯○]/g, "").replace(/\s{2,}/g, " ").trim();
      cur = { id: uid(), program: label, title: rest, pageFrom: pf, pageTo: pt, grammar: [] };
      units.push(cur);
      continue;
    }
    if (!cur) continue;
    if (cur.pageFrom == null) { const r = takeRange(ln); if (r) { cur.pageFrom = r[0]; cur.pageTo = r[1]; } }
    if (/^[◯○・●▶►\-]/.test(ln)) {
      const g = ln.replace(/^[◯○・●▶►\-]\s*/, "").trim();
      if (g && g.length <= 42 && cur.grammar.length < 8) cur.grammar.push(g);
    }
  }
  units.forEach((u) => { if (u.pageFrom == null) { u.pageFrom = 0; u.pageTo = 0; } if (u.pageTo < u.pageFrom) u.pageTo = u.pageFrom; });
  return units.filter((u) => u.pageFrom || u.title || u.program);
}

function weeklyClassCount(data, klass) {
  const cols = data.meta.includeSat ? 6 : 5;
  let n = 0;
  for (let di = 0; di < cols; di++) for (let pi = 0; pi < data.periods.length; pi++) { const c = data.timetable[`${di}-${pi}`]; if (c?.klass === klass) n++; }
  return n;
}
/* --- 時数カウント（学期の起点から・実施済みベース・教科×クラス別） --- */
function currentTerm(data, date) {
  const terms = [...(data.terms || [])].filter((t) => /^\d{4}-\d{2}-\d{2}$/.test(t.start || "")).sort((a, b) => a.start.localeCompare(b.start));
  if (!terms.length) return { name: "年間", start: `${(date.getMonth() + 1 < 4 ? date.getFullYear() - 1 : date.getFullYear())}-04-01` };
  const key = ymd(date); let cur = terms[0];
  for (const t of terms) { if (t.start <= key) cur = t; else break; }
  return cur;
}
function termEndOf(data, term) {
  const terms = [...(data.terms || [])].filter((t) => /^\d{4}-\d{2}-\d{2}$/.test(t.start || "")).sort((a, b) => a.start.localeCompare(b.start));
  const i = terms.findIndex((t) => t.id === term.id || t.start === term.start);
  return (i >= 0 && i < terms.length - 1) ? terms[i + 1].start : "9999-12-31";
}
// 週の時間割テンプレから、学期起点〜終端(inclusive)に「登録された授業コマ数」を数える（済みに関係なく累積）
function scheduledCountUpTo(data, subject, klass, startYmd, endYmd) {
  if (!subject || endYmd < startYmd) return 0;
  const cols = data.meta.includeSat ? 6 : 5;
  const perDay = {};
  for (let di = 0; di < cols; di++) { let c = 0; for (let pi = 0; pi < data.periods.length; pi++) { const cell = data.timetable[`${di}-${pi}`]; if (cell && cell.subject === subject && (klass == null || cell.klass === klass)) c++; } perDay[di] = c; }
  let n = 0, d = parseYmd(startYmd); const end = parseYmd(endYmd);
  while (d <= end) { const di = jsDayToIdx(d.getDay()); if (di >= 0 && perDay[di]) n += perDay[di]; d = addDays(d, 1); }
  return n;
}
// 実施済み（done）の数を数える
function doneCountUpTo(data, subject, klass, startYmd, endYmd) {
  let n = 0;
  for (const lk in (data.lessonLog || {})) { const l = data.lessonLog[lk]; if (!l || !l.done || l.subject !== subject || (klass != null && l.klass !== klass)) continue; const d = lk.slice(0, 10); if (d < startYmd || d > endYmd) continue; n++; }
  return n;
}
// このコマが学期の「登録ベースで何時間目か」（済みに関係なく通し番号）
function lessonOrdinal(data, subject, klass, key, pi, termStart) {
  if (key < termStart) return 1;
  const before = scheduledCountUpTo(data, subject, klass, termStart, ymd(addDays(parseYmd(key), -1)));
  const di = jsDayToIdx(parseYmd(key).getDay());
  let today = 0; for (let p = 0; p < pi; p++) { const cell = data.timetable[`${di}-${p}`]; if (cell && cell.subject === subject && cell.klass === klass) today++; }
  return before + today + 1;
}
// 学期内の実施済み時数を 教科→{クラス:件数} で集計（従来互換）
function termLessonCounts(data, termStart, termEnd) {
  const map = {};
  for (const lk in (data.lessonLog || {})) {
    const l = data.lessonLog[lk]; if (!l || !l.done || !l.subject) continue;
    const d = lk.slice(0, 10); if (d < termStart || d >= termEnd) continue;
    const kl = l.klass || "—"; (map[l.subject] = map[l.subject] || {}); map[l.subject][kl] = (map[l.subject][kl] || 0) + 1;
  }
  return map;
}
// 時間割テンプレに存在する 教科→[クラス...] の一覧
function timetableSubjectKlasses(data) {
  const cols = data.meta.includeSat ? 6 : 5; const map = {};
  for (let di = 0; di < cols; di++) for (let pi = 0; pi < data.periods.length; pi++) { const c = data.timetable[`${di}-${pi}`]; if (c && c.subject && c.klass) { (map[c.subject] = map[c.subject] || new Set()).add(c.klass); } }
  const out = {}; Object.keys(map).forEach((s) => { out[s] = [...map[s]].sort(); }); return out;
}
const mondaysBetween = (from, to) => { const res = []; let d = startOfWeekMon(from); while (d <= to) { res.push(new Date(d)); d = addDays(d, 7); } return res; };

function generateYearPlan(data, klass) {
  const wk = data.weeklyManual?.[klass] ?? (weeklyClassCount(data, klass) || 4);
  const units = data.textbook?.units || [];
  const tests = [...(data.tests || [])].filter((t) => /^\d{4}-\d{2}-\d{2}$/.test(t.date)).sort((a, b) => a.date.localeCompare(b.date));
  const taught = new Set();
  const plans = [];
  let prevDate = null;
  const planStart = data.planStart ? parseYmd(data.planStart) : (tests[0] ? addDays(parseYmd(tests[0].date), -56) : new Date());
  tests.forEach((t) => {
    const td = parseYmd(t.date);
    const deadline = addDays(td, -7); // テスト1週前
    const spanStart = prevDate ? addDays(prevDate, 1) : planStart;
    const scope = units.filter((u) => u.pageTo >= (t.pageFrom || 0) && u.pageFrom <= (t.pageTo || 9999));
    const tasks = [];
    scope.forEach((u) => {
      if (taught.has(u.id)) { tasks.push({ unit: u, phase: "応用・復習" }); }
      else { PHASES.forEach((p) => tasks.push({ unit: u, phase: p })); taught.add(u.id); }
    });
    let weeks = mondaysBetween(spanStart, deadline);
    if (!weeks.length) weeks = [startOfWeekMon(deadline)];
    const perWeek = Math.max(1, Math.ceil(tasks.length / weeks.length));
    const weekPlan = weeks.map((mon, i) => ({ mon, tasks: tasks.slice(i * perWeek, (i + 1) * perWeek) }));
    plans.push({ test: t, scope, weekPlan, deadline, testWeek: startOfWeekMon(td), wk });
    prevDate = td;
  });
  return plans;
}

/* --- 汎用ロードマップ生成（教科非依存：単元×段階をテストまでに配分） --- */
const DEFAULT_PHASES = ["導入", "展開", "定着", "応用・評価"];
function generateRoadmap(data, subject) {
  const rm = (data.roadmap && data.roadmap[subject]) || {};
  const units = (rm.units && rm.units.length) ? rm.units : [];
  const phases = (rm.phases && rm.phases.length) ? rm.phases : DEFAULT_PHASES;
  const tests = [...(data.tests || [])].filter((t) => /^\d{4}-\d{2}-\d{2}$/.test(t.date)).sort((a, b) => a.date.localeCompare(b.date));
  if (!tests.length || !units.length) return { plans: [], units, phases };
  const start = rm.start ? parseYmd(rm.start) : (tests[0] ? addDays(parseYmd(tests[0].date), -56) : new Date());
  const perTest = Math.ceil(units.length / tests.length);
  const plans = []; let prevDate = null; let ui = 0;
  tests.forEach((t, ti) => {
    const td = parseYmd(t.date); const deadline = addDays(td, -7);
    const spanStart = prevDate ? addDays(prevDate, 1) : start;
    const cnt = (ti === tests.length - 1) ? (units.length - ui) : Math.min(perTest, units.length - ui);
    const scope = units.slice(ui, ui + Math.max(0, cnt)); ui += Math.max(0, cnt);
    const tasks = []; scope.forEach((u) => phases.forEach((ph) => tasks.push({ unit: u, phase: ph })));
    let weeks = mondaysBetween(spanStart, deadline); if (!weeks.length) weeks = [startOfWeekMon(deadline)];
    const pw = Math.max(1, Math.ceil(tasks.length / weeks.length));
    const weekPlan = weeks.map((mon, i) => ({ mon, tasks: tasks.slice(i * pw, (i + 1) * pw) }));
    plans.push({ test: t, scope, weekPlan, deadline });
    prevDate = td;
  });
  return { plans, units, phases };
}
function TextbookPanel({ data, setData, onBack }) {
  const tb = data.textbook || { grade: "", units: [] };
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");
  const fileRef = useRef(null);
  const setUnits = (fn) => setData((d) => ({ ...d, textbook: { ...(d.textbook || {}), units: fn((d.textbook?.units) || []) } }));

  const onWord = async (f) => {
    if (!f) return;
    setErr(""); setBusy(true);
    try {
      const arrayBuffer = await f.arrayBuffer();
      const res = await mammoth.extractRawText({ arrayBuffer });
      const parsed = parseTextbookText(res.value);
      if (!parsed.length) { setErr("単元を検出できませんでした。「第◯章／単元／PROGRAM／Unit」などの見出しと、ページ表記（p.12〜24 や 12-24ページ）のあるWordでお試しください。"); }
      else setData((d) => ({ ...d, textbook: { ...(d.textbook || {}), units: parsed } }));
    } catch (e) {
      setErr("Wordの読み込みに失敗しました（" + (e.message || "") + "）。手入力でも追加できます。");
    } finally { setBusy(false); }
  };
  const upd = (id, patch) => setUnits((u) => u.map((x) => x.id === id ? { ...x, ...patch } : x));
  const del = (id) => setUnits((u) => u.filter((x) => x.id !== id));
  const add = () => setUnits((u) => [...u, { id: uid(), program: "第" + (u.length + 1) + "章", title: "", pageFrom: 0, pageTo: 0, grammar: [] }]);

  return (
    <section className="tp-card">
      {onBack && <button className="tp-linkbtn" onClick={onBack}>← ロードマップに戻る</button>}
      <h4 className="tp-card-title"><BookOpen size={15} /> 教科書データ（全教科・任意）</h4>
      <p className="tp-hint">教科書のWordファイルを読み込むと、見出し（第◯章・単元・PROGRAM・Unit など）とページ表記（p.12〜24／12-24ページ）から単元・ページを自動抽出します。英語では「◯」等の項目行も取り込みます。抽出後は各欄を修正できます。</p>
      <div className="tp-toolbar" style={{ justifyContent: "flex-start", gap: 8 }}>
        <button className="tp-ghostbtn" disabled={busy} onClick={() => fileRef.current?.click()}>{busy ? <><Loader size={14} className="tp-spin" /> 読み込み中…</> : <><Upload size={14} /> 教科書Wordを読み込む</>}</button>
        <input value={tb.grade || ""} onChange={(e) => setData((d) => ({ ...d, textbook: { ...(d.textbook || {}), grade: e.target.value } }))} placeholder="学年（例 2年）" style={{ width: 110 }} />
      </div>
      <input ref={fileRef} type="file" accept=".docx" style={{ display: "none" }} onChange={(e) => onWord(e.target.files?.[0])} />
      {err && <div className="tp-error">{err}</div>}
      <div className="tp-unitlist">
        {(tb.units || []).map((u) => (
          <div key={u.id} className="tp-unitrow">
            <input className="tp-u-prog" value={u.program} onChange={(e) => upd(u.id, { program: e.target.value })} placeholder="見出し（第1章 等）" />
            <input className="tp-u-title" value={u.title} onChange={(e) => upd(u.id, { title: e.target.value })} placeholder="単元名" />
            <span className="tp-u-p">p.<input type="number" value={u.pageFrom} onChange={(e) => upd(u.id, { pageFrom: Number(e.target.value) })} />–<input type="number" value={u.pageTo} onChange={(e) => upd(u.id, { pageTo: Number(e.target.value) })} /></span>
            <input className="tp-u-gram" value={(u.grammar || []).join(" / ")} onChange={(e) => upd(u.id, { grammar: e.target.value.split("/").map((s) => s.trim()).filter(Boolean) })} placeholder="項目・ポイント（/区切り）" />
            <button className="tp-iconbtn tiny" onClick={() => del(u.id)}><Trash2 size={13} /></button>
          </div>
        ))}
      </div>
      <button className="tp-primarybtn" onClick={add}><Plus size={15} /> 単元を追加</button>
    </section>
  );
}

/* --- テストパネル --- */
/* 年間行事(events)から「テスト」を抽出して種別を判定 */
function detectTestType(title) {
  const s = zen2han(String(title || ""));
  if (/単語/.test(s)) return "単語";
  if (/文法/.test(s)) return "文法";
  if (/単元/.test(s)) return "単元";
  if (/(定期|中間|期末|学期末|実力|考査)/.test(s)) return "定期";
  return null;
}
function extractTestsFromEvents(data) {
  const existing = new Set((data.tests || []).map((t) => `${t.date}|${(t.name || "").trim()}`));
  const out = [];
  (data.events || []).forEach((e) => {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(e.date || "")) return;
    const type = detectTestType(e.title) || (/テスト|考査/.test(zen2han(e.title || "")) ? "定期" : null);
    if (!type) return;
    const key = `${e.date}|${(e.title || "").trim()}`;
    if (existing.has(key)) return;
    out.push({ date: e.date, name: (e.title || "テスト").trim(), type });
  });
  return out;
}

function TestsPanel({ data, setData }) {
  const tests = [...(data.tests || [])].sort((a, b) => (a.date || "").localeCompare(b.date || ""));
  const [t, setT] = useState({ name: "", type: "定期", date: "", pageFrom: "", pageTo: "" });
  const [found, setFound] = useState(null);
  const [foundSel, setFoundSel] = useState([]);
  const units = (data.textbook && data.textbook.units) || [];
  const add = () => { if (!t.name.trim() || !t.date) return; setData((d) => ({ ...d, tests: [...(d.tests || []), { id: uid(), name: t.name.trim(), type: t.type, date: t.date, pageFrom: Number(t.pageFrom) || 0, pageTo: Number(t.pageTo) || 0 }] })); setT({ name: "", type: "定期", date: "", pageFrom: "", pageTo: "" }); };
  const del = (id) => setData((d) => ({ ...d, tests: d.tests.filter((x) => x.id !== id) }));
  const updTest = (id, patch) => setData((d) => ({ ...d, tests: d.tests.map((x) => x.id === id ? { ...x, ...patch } : x) }));
  const scan = () => { const f = extractTestsFromEvents(data); setFound(f); setFoundSel(f.map(() => true)); };
  const checkedCount = (foundSel || []).filter(Boolean).length;
  const addFound = () => { const items = (found || []).filter((_, i) => foundSel[i]).map((f) => ({ id: uid(), name: f.name, type: f.type, date: f.date, pageFrom: 0, pageTo: 0 })); if (items.length) setData((d) => ({ ...d, tests: [...(d.tests || []), ...items] })); setFound(null); setFoundSel([]); };
  const applyUnit = (id, uid2) => { const u = units.find((x) => x.id === uid2); if (!u) return; updTest(id, { pageFrom: u.pageFrom || 0, pageTo: u.pageTo || 0, unitId: uid2 }); };

  return (
    <section className="tp-card">
      <h4 className="tp-card-title"><ClipboardList size={15} /> テスト（定期・単元・単語・文法）</h4>
      <p className="tp-hint">定期・単元・単語・文法テストの期日と<b>教科書範囲</b>を登録すると、それぞれの1週前までに段階を終える計画に使えます。年間行事から自動抽出もできます。</p>
      <div className="tp-toolbar" style={{ justifyContent: "flex-start", gap: 8 }}>
        <button className="tp-ghostbtn" onClick={scan}><Search size={14} /> 年間行事からテストを抽出</button>
      </div>
      {found && (
        <div className="tp-testfound">
          <div className="tp-mtg-grouphead">抽出 {found.length}件{found.length === 0 ? "（テストらしい行事は見つかりませんでした）" : " — 追加するものにチェック"}</div>
          {found.map((f, i) => { const d = parseYmd(f.date); return (
            <div key={i} className="tp-testfound-row" onClick={() => setFoundSel((s) => s.map((x, j) => j === i ? !x : x))}>
              <button className={"tp-donebox" + (foundSel[i] ? " on" : "")}>{foundSel[i] && <Check size={12} />}</button>
              <span className="tp-chip sm" style={{ background: testTypeColor(f.type) }}>{f.type}</span>
              <span className="tp-test-date">{d.getMonth() + 1}/{d.getDate()}</span>
              <span className="tp-test-name">{f.name}</span>
            </div>
          ); })}
          {found.length > 0 && <div className="tp-modal-actions"><button className="tp-ghostbtn" onClick={() => { setFound(null); setFoundSel([]); }}>キャンセル</button><button className="tp-primarybtn" disabled={checkedCount === 0} onClick={addFound}><Plus size={15} /> チェックした{checkedCount}件をリストに追加</button></div>}
        </div>
      )}
      <ul className="tp-testlist">
        {tests.length === 0 && <li className="tp-empty">テスト未登録</li>}
        {tests.map((x) => { const d = parseYmd(x.date); return (
          <li key={x.id} className="tp-test-item">
            <span className="tp-chip sm" style={{ background: testTypeColor(x.type) }}>{x.type}</span>
            <span className="tp-test-date">{d.getMonth() + 1}/{d.getDate()}</span>
            <span className="tp-test-name">{x.name}</span>
            <span className="tp-u-p tp-test-range">p.<input type="number" value={x.pageFrom || ""} onChange={(e) => updTest(x.id, { pageFrom: Number(e.target.value) || 0 })} placeholder="—" />–<input type="number" value={x.pageTo || ""} onChange={(e) => updTest(x.id, { pageTo: Number(e.target.value) || 0 })} placeholder="—" /></span>
            {units.length > 0 && <select className="tp-test-unit" value={x.unitId || ""} onChange={(e) => applyUnit(x.id, e.target.value)}><option value="">単元から…</option>{units.map((u) => <option key={u.id} value={u.id}>{u.program || u.title || u.name}</option>)}</select>}
            <button className="tp-iconbtn tiny tp-test-del" onClick={() => del(x.id)}><Trash2 size={13} /></button>
          </li>
        ); })}
      </ul>
      <div className="tp-divider" />
      <div className="tp-test-form">
        <input value={t.name} onChange={(e) => setT({ ...t, name: e.target.value })} placeholder="テスト名（例 1学期期末）" />
        <select value={t.type} onChange={(e) => setT({ ...t, type: e.target.value })}>{TEST_TYPES.map((x) => <option key={x} value={x}>{x}</option>)}</select>
        <input type="date" value={t.date} onChange={(e) => setT({ ...t, date: e.target.value })} />
        <span className="tp-u-p">p.<input type="number" value={t.pageFrom} onChange={(e) => setT({ ...t, pageFrom: e.target.value })} placeholder="from" />–<input type="number" value={t.pageTo} onChange={(e) => setT({ ...t, pageTo: e.target.value })} placeholder="to" /></span>
        <button className="tp-addbtn" onClick={add}><Plus size={16} /></button>
      </div>
    </section>
  );
}

/* --- 年間計画パネル --- */
function YearPlanPanel({ data, setData }) {
  const klass = data.planClass || data.classes[0] || "";
  const plans = useMemo(() => generateYearPlan(data, klass), [data.tests, data.textbook, data.weeklyManual, data.planStart, klass, data.timetable]);
  const wk = data.weeklyManual?.[klass] ?? (weeklyClassCount(data, klass) || 4);
  return (
    <section className="tp-card">
      <div className="tp-plan-head">
        <h4 className="tp-card-title" style={{ margin: 0 }}><CalendarRange size={15} /> 年間授業計画</h4>
        <div className="tp-plan-controls">
          <label>対象<select value={klass} onChange={(e) => setData((d) => ({ ...d, planClass: e.target.value }))}>{data.classes.map((c) => <option key={c} value={c}>{c}</option>)}</select></label>
          <label>開始<input type="date" value={data.planStart || ""} onChange={(e) => setData((d) => ({ ...d, planStart: e.target.value }))} /></label>
          <span className="tp-plan-wk">週{wk}コマ</span>
        </div>
      </div>
      <p className="tp-hint">各テストの<b>1週前</b>までに「新出文法理解 → 教科書概要理解 → 本文暗唱 → 応用」を終える計画です。週コマ数（設定の手入力 or 時間割）を基準に、週ごとに配分します。</p>
      {plans.length === 0 && <p className="tp-empty">テストと教科書単元を登録すると計画が表示されます。</p>}
      {plans.map((pl, i) => {
        const td = parseYmd(pl.test.date);
        return (
          <div key={i} className="tp-plan-block">
            <div className="tp-plan-test" style={{ borderColor: testTypeColor(pl.test.type) }}>
              <span className="tp-chip sm" style={{ background: testTypeColor(pl.test.type) }}>{pl.test.type}</span>
              <b>{pl.test.name}</b><span className="tp-plan-testdate">{td.getMonth() + 1}/{td.getDate()}（{WD[td.getDay()]}）</span>
              <span className="tp-plan-scope">範囲 p.{pl.test.pageFrom}–{pl.test.pageTo}／{pl.scope.map((u) => u.program.replace("PROGRAM ", "P")).join("・") || "該当単元なし"}</span>
            </div>
            <div className="tp-plan-weeks">
              {pl.weekPlan.map((w, j) => (
                <div key={j} className="tp-plan-week">
                  <div className="tp-plan-wdate">{w.mon.getMonth() + 1}/{w.mon.getDate()}週</div>
                  <div className="tp-plan-tasks">
                    {w.tasks.length === 0 ? <span className="tp-plan-none">演習・予備</span> : w.tasks.map((tk, k) => (
                      <span key={k} className="tp-plan-task"><b>{tk.unit.program.replace("PROGRAM ", "P")}</b> {tk.phase}<small>p.{tk.unit.pageFrom}–{tk.unit.pageTo}</small></span>
                    ))}
                  </div>
                </div>
              ))}
              <div className="tp-plan-week deadline"><div className="tp-plan-wdate">{addDays(td, -7).getMonth() + 1}/{addDays(td, -7).getDate()}</div><div className="tp-plan-tasks"><span className="tp-plan-dl">← ここまでに4段階完了</span></div></div>
              <div className="tp-plan-week testrow"><div className="tp-plan-wdate">{td.getMonth() + 1}/{td.getDate()}</div><div className="tp-plan-tasks"><span className="tp-plan-testtag">📝 {pl.test.name}・応用/対策</span></div></div>
            </div>
          </div>
        );
      })}
    </section>
  );
}

/* --- ロードマップ（教科汎用：単元×段階→テストまでの週間指導計画） --- */
function RoadmapPanel({ data, setData, onOpenTextbook }) {
  const testSubjects = (data.subjects || []).filter((s) => s.test).map((s) => s.name);
  const [sub, setSub] = useState(testSubjects[0] || "");
  useEffect(() => { if (!sub && testSubjects[0]) setSub(testSubjects[0]); }, [testSubjects.join(",")]);
  const rm = (data.roadmap && data.roadmap[sub]) || {};
  const units = rm.units || [];
  const phases = (rm.phases && rm.phases.length) ? rm.phases : DEFAULT_PHASES;
  const setRM = (patch) => setData((d) => ({ ...d, roadmap: { ...(d.roadmap || {}), [sub]: { ...(d.roadmap?.[sub] || {}), ...patch } } }));
  const setUnitCount = (n) => { n = Math.max(0, Math.min(60, n | 0)); const next = Array.from({ length: n }, (_, i) => units[i] || { id: uid(), name: "単元" + (i + 1) }); setRM({ units: next }); };
  const setUnitName = (i, name) => { const next = units.map((u, j) => j === i ? { ...u, name } : u); setRM({ units: next }); };
  const setPhasesText = (txt) => setRM({ phases: txt.split(/[\/、,]/).map((s) => s.trim()).filter(Boolean) });
  const { plans } = useMemo(() => generateRoadmap(data, sub), [data.tests, data.roadmap, sub]);

  if (testSubjects.length === 0) {
    return (<section className="tp-card"><h4 className="tp-card-title"><CalendarRange size={15} /> ロードマップ（指導計画）</h4>
      <p className="tp-empty">設定 ›「教科」で、テストを実施する教科に<b>「テスト実施」</b>のチェックを入れてください。選んだ教科ごとに、導入→テストのロードマップを作成できます。</p></section>);
  }
  return (
    <>
      <section className="tp-card">
        <div className="tp-plan-head">
          <h4 className="tp-card-title" style={{ margin: 0 }}><CalendarRange size={15} /> ロードマップ（指導計画）</h4>
          <div className="tp-plan-controls">
            <label>教科<select value={sub} onChange={(e) => setSub(e.target.value)}>{testSubjects.map((s) => <option key={s} value={s}>{s}</option>)}</select></label>
            <label>開始<input type="date" value={rm.start || ""} onChange={(e) => setRM({ start: e.target.value })} /></label>
          </div>
        </div>
        <p className="tp-hint" style={{ marginTop: 0 }}>「単元・段階の数」と「指導の段階」を決めると、登録済みのテスト日を区切りに、導入からテストまでの内容の目安を週ごとに配分します。{onOpenTextbook && <button className="tp-linkbtn" onClick={onOpenTextbook}>教科書データ（Word）を使う</button>}</p>
        <div className="tp-rm-cfg">
          <label>単元・段階の数<input type="number" min="0" max="60" value={units.length} onChange={(e) => setUnitCount(Number(e.target.value))} /></label>
          <label className="tp-rm-phases">指導の段階（/区切り）<input value={phases.join(" / ")} onChange={(e) => setPhasesText(e.target.value)} placeholder="導入 / 展開 / 定着 / 応用・評価" /></label>
        </div>
        {units.length > 0 && (
          <div className="tp-rm-units">
            {units.map((u, i) => <div key={u.id} className="tp-rm-unit"><span className="tp-rm-uno">{i + 1}</span><input value={u.name} onChange={(e) => setUnitName(i, e.target.value)} placeholder={"単元" + (i + 1)} /></div>)}
          </div>
        )}
      </section>

      <section className="tp-card">
        <h4 className="tp-card-title" style={{ margin: "0 0 8px" }}>{sub || "教科"} のロードマップ</h4>
        {plans.length === 0 && <p className="tp-empty">「テスト」タブでテスト日を登録し、上で単元数を設定すると、ここに導入→テストの計画が表示されます。</p>}
        {plans.map((pl, i) => { const td = parseYmd(pl.test.date); return (
          <div key={i} className="tp-plan-block">
            <div className="tp-plan-test" style={{ borderColor: testTypeColor(pl.test.type) }}>
              <span className="tp-chip sm" style={{ background: testTypeColor(pl.test.type) }}>{pl.test.type}</span>
              <b>{pl.test.name}</b><span className="tp-plan-testdate">{td.getMonth() + 1}/{td.getDate()}（{WD[td.getDay()]}）</span>
              <span className="tp-plan-scope">このテストまで：{pl.scope.map((u) => u.name).join("・") || "単元なし"}</span>
            </div>
            <div className="tp-plan-weeks">
              {pl.weekPlan.map((w, j) => (
                <div key={j} className="tp-plan-week">
                  <div className="tp-plan-wdate">{w.mon.getMonth() + 1}/{w.mon.getDate()}週</div>
                  <div className="tp-plan-tasks">
                    {w.tasks.length === 0 ? <span className="tp-plan-none">演習・予備</span> : w.tasks.map((tk, k) => (
                      <span key={k} className="tp-plan-task"><b>{tk.unit.name}</b> {tk.phase}</span>
                    ))}
                  </div>
                </div>
              ))}
              <div className="tp-plan-week deadline"><div className="tp-plan-wdate">{addDays(td, -7).getMonth() + 1}/{addDays(td, -7).getDate()}</div><div className="tp-plan-tasks"><span className="tp-plan-dl">← ここまでに指導完了（テスト1週前）</span></div></div>
              <div className="tp-plan-week testrow"><div className="tp-plan-wdate">{td.getMonth() + 1}/{td.getDate()}</div><div className="tp-plan-tasks"><span className="tp-plan-testtag">📝 {pl.test.name}</span></div></div>
            </div>
          </div>
        ); })}
      </section>
    </>
  );
}

/* 実施時数（教科×クラス）：登録ベースの累積と、済み/累積の分数 */
function ClassProgressCard({ data }) {
  const [termId, setTermId] = useState("");
  const terms = [...(data.terms || [])].filter((t) => /^\d{4}-\d{2}-\d{2}$/.test(t.start || "")).sort((a, b) => a.start.localeCompare(b.start));
  const cur = terms.find((t) => t.id === termId) || currentTerm(data, new Date());
  const endExclusive = termEndOf(data, cur);
  const todayY = ymd(new Date());
  // 数える終端：今日（学期内）／過去学期なら学期末前日／未来学期なら起点前日（=0件）
  const lastDay = ymd(addDays(parseYmd(endExclusive), -1));
  const endY = todayY < cur.start ? ymd(addDays(parseYmd(cur.start), -1)) : (todayY <= lastDay ? todayY : lastDay);
  const sk = timetableSubjectKlasses(data);
  const subjects = Object.keys(sk).sort();
  return (
    <section className="tp-card">
      <div className="tp-plan-head">
        <h4 className="tp-card-title" style={{ margin: 0 }}><BookOpen size={15} /> 実施時数（教科×クラス）</h4>
        <div className="tp-plan-controls">
          <label>学期<select value={cur.id || ""} onChange={(e) => setTermId(e.target.value)}>{terms.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}</select></label>
        </div>
      </div>
      <p className="tp-hint" style={{ marginTop: 0 }}>時間割に登録された授業を、学期の起点から今日までで数えた<b>累積時数</b>です。数字は「<b>済み / 累積</b>」。クラス間の進度差を見ながら内容を検討できます。</p>
      {subjects.length === 0 ? <p className="tp-empty">時間割に授業が登録されていません。週間タブで教科・クラスを設定してください。</p> : (
        <div className="tp-prog2">
          {subjects.map((s) => {
            const ks = sk[s];
            const rows = ks.map((k) => ({ k, sched: scheduledCountUpTo(data, s, k, cur.start, endY), done: doneCountUpTo(data, s, k, cur.start, endY) }));
            const scheds = rows.map((r) => r.sched); const max = Math.max(...scheds); const min = Math.min(...scheds);
            return (
              <div key={s} className="tp-prog2-row">
                <div className="tp-prog2-sub"><span className="tp-dot2" style={{ background: subjColor(data, s) }} />{s}{ks.length > 1 && max !== min && <span className="tp-prog2-gap">差{max - min}</span>}</div>
                <div className="tp-prog2-klasses">
                  {rows.map((r) => <span key={r.k} className={"tp-prog2-k" + (r.sched === max && max !== min ? " hi" : "") + (r.sched === min && max !== min ? " lo" : "")}>{r.k}<b>{r.done}/{r.sched}</b></span>)}
                </div>
              </div>
            );
          })}
        </div>
      )}
    </section>
  );
}

/* ============================================================
   CLASSES — 授業数集計 + 授業記録
   ============================================================ */
function ProgressPanel({ data, setData }) {
  const fileRef = useRef(null);
  const [msg, setMsg] = useState("");
  const tests = (data.tests || []).filter((t) => t.type === "定期").slice().sort((a, b) => (a.date || "").localeCompare(b.date || ""));
  const tp = data.testProgress || {};
  const setTP = (id, patch) => setData((d) => ({ ...d, testProgress: { ...(d.testProgress || {}), [id]: { ...(d.testProgress?.[id] || {}), ...patch } } }));
  const STATUS = ["未設定", "前倒し", "順調", "やや遅れ", "遅れ"];

  const exportProgress = () => {
    try {
      const payload = { __progress: true, __app: APP_VERSION, subject: data.textbook?.grade || "", tests: tests.map((t) => ({ name: t.name, date: t.date, pageFrom: t.pageFrom, pageTo: t.pageTo })), testProgress: tests.reduce((o, t) => { o[t.name] = tp[t.id] || {}; return o; }, {}), units: (data.textbook?.units || []).map((u) => ({ program: u.program, title: u.title, pageFrom: u.pageFrom, pageTo: u.pageTo })) };
      const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
      const url = URL.createObjectURL(blob); const a = document.createElement("a"); a.href = url; a.download = `進度計画_${ymd(new Date())}.json`; document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(url), 1000);
      setMsg("進度計画(共有用)を書き出しました。共有フォルダに置いて、担当の先生が「読み込み」できます。");
    } catch (e) { setMsg("書き出しに失敗しました。"); }
  };
  const importProgress = (f) => {
    if (!f) return; const r = new FileReader();
    r.onload = () => {
      try {
        const obj = JSON.parse(String(r.result));
        setData((d) => {
          let tests2 = [...(d.tests || [])];
          (obj.tests || []).forEach((t) => { const ex = tests2.find((x) => x.name === t.name && x.type === "定期"); if (ex) { ex.date = t.date || ex.date; ex.pageFrom = t.pageFrom ?? ex.pageFrom; ex.pageTo = t.pageTo ?? ex.pageTo; } else tests2.push({ id: uid(), name: t.name, type: "定期", date: t.date || "", pageFrom: t.pageFrom || 0, pageTo: t.pageTo || 0 }); });
          const prog = { ...(d.testProgress || {}) };
          Object.entries(obj.testProgress || {}).forEach(([name, v]) => { const t = tests2.find((x) => x.name === name); if (t) prog[t.id] = { ...(prog[t.id] || {}), ...v }; });
          return { ...d, tests: tests2, testProgress: prog };
        });
        setMsg("進度計画(共有用)を読み込みました（テストと進度を統合）。");
      } catch (e) { setMsg("読み込みに失敗しました（形式を確認してください）。"); }
    };
    r.readAsText(f);
  };

  return (
    <>
      <section className="tp-card">
        <h4 className="tp-card-title"><BookOpen size={15} /> 定期テストごとの進度計画</h4>
        <p className="tp-hint" style={{ marginTop: 0 }}>各定期テストまでに「どこまで進めるか（目標）」と「実際の進度」を記録します。教科担当の先生と、下の書き出し／読み込みで共有できます。</p>
        <div className="tp-toolbar" style={{ justifyContent: "flex-start", gap: 8 }}>
          <button className="tp-ghostbtn" onClick={exportProgress}><Download size={14} /> 進度計画(共有用)を書き出し</button>
          <button className="tp-ghostbtn" onClick={() => fileRef.current?.click()}><Upload size={14} /> 進度計画(共有用)を読み込み</button>
          <input ref={fileRef} type="file" accept="application/json,.json" style={{ display: "none" }} onChange={(e) => { importProgress(e.target.files?.[0]); e.target.value = ""; }} />
        </div>
        {msg && <p className="tp-hint" style={{ color: "#2C7CA6", fontWeight: 700 }}>{msg}</p>}
        {tests.length === 0 && <p className="tp-empty">「テスト」タブで定期テストを登録すると、ここに進度計画が並びます。</p>}
        <div className="tp-prog">
          {tests.map((t) => { const p = tp[t.id] || {}; const st = p.status || "未設定"; return (
            <div key={t.id} className="tp-prog-card">
              <div className="tp-prog-head">
                <span className="tp-prog-name">{t.name}</span>
                <span className="tp-prog-date">{t.date ? `${parseYmd(t.date).getMonth() + 1}/${parseYmd(t.date).getDate()}` : "日付未設定"}</span>
                {(t.pageFrom || t.pageTo) ? <span className="tp-prog-range">範囲 p.{t.pageFrom}–{t.pageTo}</span> : null}
                <span className={"tp-prog-status s" + STATUS.indexOf(st)}>{st}</span>
              </div>
              <div className="tp-prog-grid">
                <label><span>このテストまでの目標</span><input value={p.plannedTo || ""} onChange={(e) => setTP(t.id, { plannedTo: e.target.value })} placeholder="例）PROGRAM 3 / p.27 まで" /></label>
                <label><span>実際の進度</span><input value={p.actualTo || ""} onChange={(e) => setTP(t.id, { actualTo: e.target.value })} placeholder="例）p.24 まで" /></label>
                <label><span>状況</span><select value={st} onChange={(e) => setTP(t.id, { status: e.target.value })}>{STATUS.map((s) => <option key={s} value={s}>{s}</option>)}</select></label>
                <label className="tp-prog-note"><span>メモ（共有事項）</span><input value={p.note || ""} onChange={(e) => setTP(t.id, { note: e.target.value })} placeholder="難単元・補習・進度調整など" /></label>
              </div>
            </div>
          ); })}
        </div>
      </section>
    </>
  );
}

function ClassesView({ data, setData }) {
  const cols = data.meta.includeSat ? 6 : 5;
  // 週あたりコマ数（時間割から）
  const weekly = useMemo(() => {
    const bySub = {}, byClass = {};
    for (let di = 0; di < cols; di++) for (let pi = 0; pi < data.periods.length; pi++) {
      const c = data.timetable[`${di}-${pi}`];
      if (c?.subject) { bySub[c.subject] = (bySub[c.subject] || 0) + 1; if (c.klass) byClass[c.klass] = (byClass[c.klass] || 0) + 1; }
    }
    return { bySub, byClass };
  }, [data, cols]);
  // 実施累計（ログのdoneから）
  const done = useMemo(() => {
    const bySub = {}, byClass = {};
    Object.values(data.lessonLog).forEach((l) => { if (l.done) { if (l.subject) bySub[l.subject] = (bySub[l.subject] || 0) + 1; if (l.klass) byClass[l.klass] = (byClass[l.klass] || 0) + 1; } });
    return { bySub, byClass };
  }, [data]);

  const subjectsUsed = data.subjects.filter((s) => weekly.bySub[s.name] || done.bySub[s.name] || data.targets[s.name]);
  const records = useMemo(() => Object.entries(data.lessonLog).filter(([, l]) => l.done || l.topic).map(([k, l]) => { const parts = k.split("-"); const pi = Number(parts[3]); const date = `${parts[0]}-${parts[1]}-${parts[2]}`; return { date, pi, ...l }; }).sort((a, b) => (b.date + b.pi).localeCompare(a.date + a.pi)), [data.lessonLog]);
  const [filter, setFilter] = useState("");
  const [sub, setSub] = useState("agg");
  const SUBTABS = [{ id: "agg", label: "集計・記録" }, { id: "test", label: "テスト" }, { id: "plan", label: "ロードマップ（計画）" }, { id: "prog", label: "進度（実績）" }];
  const goSub = (dir) => { const ids = SUBTABS.map((s) => s.id); const i = ids.indexOf(sub === "book" ? "plan" : sub); const n = (i + dir + ids.length) % ids.length; setSub(ids[n]); };
  const swipe = useSwipe(() => goSub(1), () => goSub(-1));

  return (
    <div className="tp-view" {...swipe}>
      <div className="tp-subtabs">
        {SUBTABS.map((s) => <button key={s.id} className={"tp-subtab" + (sub === s.id ? " on" : "")} onClick={() => setSub(s.id)}>{s.label}</button>)}
      </div>

      {sub === "book" && <TextbookPanel data={data} setData={setData} onBack={() => setSub("plan")} />}
      {sub === "test" && <TestsPanel data={data} setData={setData} />}
      {sub === "prog" && <ProgressPanel data={data} setData={setData} />}
      {sub === "plan" && <RoadmapPanel data={data} setData={setData} onOpenTextbook={() => setSub("book")} />}

      {sub === "agg" && <>
      <ClassProgressCard data={data} />
      <section className="tp-card">
        <h4 className="tp-card-title"><BookOpen size={15} /> 授業数集計（教科別）</h4>
        <p className="tp-hint">週コマ数は時間割から自動集計。実施累計は「今日」画面で済みにしたコマ数です。目標は標準授業時数の目安。</p>
        <div className="tp-agg">
          {subjectsUsed.map((s) => {
            const w = weekly.bySub[s.name] || 0; const d = done.bySub[s.name] || 0; const tg = data.targets[s.name] || 0;
            const pct = tg ? Math.min(100, Math.round((d / tg) * 100)) : 0;
            return (
              <div key={s.name} className="tp-agg-row">
                <span className="tp-agg-name"><span className="tp-dot" style={{ background: s.color }} />{s.name}</span>
                <span className="tp-agg-week">週<b>{w}</b></span>
                <div className="tp-agg-bar"><div className="tp-agg-fill" style={{ width: pct + "%", background: s.color }} /></div>
                <span className="tp-agg-num">{d}<small> / {tg || "—"}</small></span>
                <input className="tp-agg-target" type="number" value={tg || ""} onChange={(e) => setData((dd) => ({ ...dd, targets: { ...dd.targets, [s.name]: Number(e.target.value) } }))} placeholder="目標" />
              </div>
            );
          })}
        </div>
      </section>

      <section className="tp-card">
        <h4 className="tp-card-title"><Users size={15} /> クラス別コマ数</h4>
        <p className="tp-hint">週コマ数は設定の「週コマ数（手入力）」があればそれを、無ければ時間割から自動算出します。</p>
        <div className="tp-classgrid">
          {data.classes.filter((c) => (data.weeklyManual?.[c] ?? weekly.byClass[c]) || done.byClass[c]).map((c) => {
            const wk = data.weeklyManual?.[c] ?? weekly.byClass[c] ?? 0;
            const manual = data.weeklyManual?.[c] != null;
            return (
              <div key={c} className="tp-classcard">
                <b>{c}</b>
                <div><span>週</span>{wk}<small>コマ{manual ? "・手入力" : ""}</small></div>
                <div><span>累計</span>{done.byClass[c] || 0}<small>コマ</small></div>
              </div>
            );
          })}
        </div>
      </section>

      <section className="tp-card">
        <h4 className="tp-card-title"><ClipboardList size={15} /> 授業記録（進度）</h4>
        <div className="tp-filterbar">
          <select value={filter} onChange={(e) => setFilter(e.target.value)}>
            <option value="">すべてのクラス</option>
            {data.classes.map((c) => <option key={c} value={c}>{c}</option>)}
          </select>
        </div>
        <ul className="tp-reclist">
          {records.filter((r) => !filter || r.klass === filter).length === 0 && <li className="tp-empty">記録はまだありません。「今日」画面の授業計画から記録できます。</li>}
          {records.filter((r) => !filter || r.klass === filter).slice(0, 40).map((r, i) => { const d = parseYmd(r.date); return (
            <li key={i} className={r.done ? "done" : ""}>
              <span className="tp-rec-date">{d.getMonth() + 1}/{d.getDate()}<small> {data.periods[r.pi]?.label}限</small></span>
              <span className="tp-chip sm" style={{ background: subjColor(data, r.subject) }}>{r.klass}</span>
              <span className="tp-rec-topic">{r.topic || <em>（内容未入力）</em>}{r.hw && <small className="tp-rec-hw">宿題: {r.hw}</small>}</span>
              {r.done && <Check size={14} className="tp-rec-check" />}
            </li>
          ); })}
        </ul>
      </section>
      </>}
    </div>
  );
}

/* ============================================================
   CLUB — 部活動（日毎スケジュール ＋ 一括入力）
   ============================================================ */
function Seg({ options, value, onChange, color = "#3E9BC9" }) {
  return (
    <div className="tp-seg">
      {options.map((o) => (
        <button key={String(o.v)} className={"tp-seg-btn" + (value === o.v ? " on" : "")} style={value === o.v ? { background: color, borderColor: color } : {}} onClick={() => onChange(o.v)}>{o.label}</button>
      ))}
    </div>
  );
}
function ClubFields({ val, set }) {
  const kind = val.kind || (val.practice === false ? "off" : "practice");
  return (
    <>
      <div className="tp-field"><span>内容</span>
        <Seg options={[{ v: "practice", label: "練習" }, { v: "match", label: "大会" }, { v: "off", label: "休み" }]} value={kind} onChange={(v) => set({ kind: v })} color={kind === "match" ? "#D9534F" : "#E8845B"} />
      </div>
      {kind === "match" && <label className="tp-field"><span>大会名</span><input value={val.matchName || ""} onChange={(e) => set({ matchName: e.target.value })} placeholder="例）市総体・新人戦・練習試合（対〇〇中）" /></label>}
      {kind === "practice" && (
        <div className="tp-field"><span>時間帯</span>
          <Seg options={[{ v: "after", label: "放課後" }, { v: "am", label: "午前" }, { v: "pm", label: "午後" }]} value={val.session || "after"} onChange={(v) => set({ session: v })} color="#E8845B" />
        </div>
      )}
      {(kind === "practice" || kind === "match") && (
        <>
          <div className="tp-field"><span>場所</span>
            <Seg options={[{ v: "二中", label: "二中" }, { v: "other", label: "その他" }]} value={val.place || "二中"} onChange={(v) => set({ place: v })} />
            {val.place === "other" && <input style={{ marginTop: 6 }} value={val.placeOther || ""} onChange={(e) => set({ placeOther: e.target.value })} placeholder="場所を記入" />}
          </div>
          <label className="tp-field"><span>時間</span><input value={val.time || ""} onChange={(e) => set({ time: e.target.value })} placeholder="例）16:00–18:00" /></label>
          {kind === "practice" && <label className="tp-field"><span>内容メモ（任意）</span><input value={val.content || ""} onChange={(e) => set({ content: e.target.value })} placeholder="例）シュート練習・ゲーム形式" /></label>}
        </>
      )}
      <label className="tp-field"><span>備考</span><input value={val.note || ""} onChange={(e) => set({ note: e.target.value })} placeholder="持ち物・連絡など" /></label>
    </>
  );
}

function ClubView({ data, setData, onShare }) {
  const club = data.club;
  const days = club.days || {};
  const setClub = (patch) => setData((d) => ({ ...d, club: { ...d.club, ...patch } }));
  const setSched = (di, patch) => setClub({ schedule: { ...club.schedule, [di]: { ...(club.schedule[di] || {}), ...patch } } });
  const setDay = (dk, patch) => setData((d) => { const nd = { ...(d.club.days || {}) }; nd[dk] = { ...(nd[dk] || { kind: "practice", session: "after", place: "二中" }), ...patch }; return { ...d, club: { ...d.club, days: nd } }; });
  const clearDay = (dk) => setData((d) => { const nd = { ...(d.club.days || {}) }; delete nd[dk]; return { ...d, club: { ...d.club, days: nd } }; });

  const [cursor, setCursor] = useState(new Date());
  const [editDate, setEditDate] = useState(null);
  const [checked, setChecked] = useState({});
  const [bulkOpen, setBulkOpen] = useState(false);
  const [bulkVal, setBulkVal] = useState({ kind: "practice", session: "after", place: "二中", placeOther: "", time: "", content: "", matchName: "", note: "" });
  const [showWeekly, setShowWeekly] = useState(false);
  const [matchImp, setMatchImp] = useState(false);
  const [clubImp, setClubImp] = useState(false);

  const y = cursor.getFullYear(), m = cursor.getMonth();
  const dim = new Date(y, m + 1, 0).getDate();
  const dates = Array.from({ length: dim }, (_, i) => new Date(y, m, i + 1));
  const checkedKeys = Object.keys(checked).filter((k) => checked[k] && k.startsWith(`${y}-${pad(m + 1)}`));
  const toggleCheck = (k) => setChecked((c) => ({ ...c, [k]: !c[k] }));
  const allChecked = dates.length > 0 && dates.every((d) => checked[ymd(d)]);
  const toggleAll = () => { const nc = { ...checked }; dates.forEach((d) => { nc[ymd(d)] = !allChecked; }); setChecked(nc); };
  const pickWeekday = (wd) => { const nc = { ...checked }; dates.forEach((d) => { if (d.getDay() === wd) nc[ymd(d)] = true; }); setChecked(nc); };
  const applyBulk = () => { setData((d) => { const nd = { ...(d.club.days || {}) }; checkedKeys.forEach((k) => { nd[k] = { ...bulkVal }; }); return { ...d, club: { ...d.club, days: nd } }; }); setBulkOpen(false); setChecked({}); };

  // 大会（日毎 kind=match）＋ 旧specials を統合表示
  const matchDays = Object.entries(club.days || {}).filter(([, v]) => (v.kind || "") === "match").map(([k, v]) => ({ key: k, ...v })).sort((a, b) => a.key.localeCompare(b.key));
  const [mt, setMt] = useState({ date: ymd(new Date()), matchName: "", placeOther: "", start: "", end: "" });
  const addMatch = () => { if (!mt.matchName.trim() || !mt.date) return; setDay(mt.date, { kind: "match", place: "other", placeOther: mt.placeOther, matchName: mt.matchName.trim(), time: [mt.start, mt.end].filter(Boolean).join("–") }); setMt({ date: ymd(new Date()), matchName: "", placeOther: "", start: "", end: "" }); };

  return (
    <div className="tp-view">
      <section className="tp-card">
        <div className="tp-clubhead">
          <Dumbbell size={20} />
          <input className="tp-clubname" value={club.name} onChange={(e) => setClub({ name: e.target.value })} placeholder="部活動名" />
        </div>

        <div className="tp-daynav" style={{ marginBottom: 8 }}>
          <button className="tp-iconbtn" onClick={() => setCursor(new Date(y, m - 1, 1))}><ChevronLeft size={20} /></button>
          <div className="tp-daynav-mid"><div className="tp-daynav-date">{y}年 {m + 1}月の予定</div>
            <button className="tp-ghostbtn sm" onClick={() => setCursor(new Date())}>今月</button>
            <button className="tp-ghostbtn sm" onClick={() => { const rows = dates.map((d) => ({ date: d, items: calItemsForDate(data, d, { club: true }) })).filter((r) => r.items.length); onShare && onShare(`${data.club.name || "部活"} ${y}年${m + 1}月 予定`, rows); }}><Printer size={13} /> 共有</button>
          </div>
          <button className="tp-iconbtn" onClick={() => setCursor(new Date(y, m + 1, 1))}><ChevronRight size={20} /></button>
        </div>

        <div className="tp-toolbar" style={{ justifyContent: "flex-start" }}>
          <button className="tp-ghostbtn" onClick={() => setClubImp(true)}><Upload size={14} /> 月間予定表を画像で取り込み</button>
        </div>

        <div className="tp-bulkbar">
          <label className="tp-check sm"><input type="checkbox" checked={allChecked} onChange={toggleAll} /> 全選択</label>
          <span className="tp-bulk-wd">曜日:{DAY_LABELS.map((l, i) => <button key={i} className="tp-wdpick" onClick={() => pickWeekday(i === 5 ? 6 : i + 1)}>{l}</button>)}<button className="tp-wdpick" onClick={() => pickWeekday(0)}>日</button></span>
          <span className="tp-bulk-spacer" />
          <button className="tp-primarybtn sm" disabled={!checkedKeys.length} onClick={() => setBulkOpen(true)}>選択{checkedKeys.length ? `（${checkedKeys.length}日）` : ""}を一括入力</button>
        </div>

        <div className="tp-dayschedule">
          {dates.map((d) => {
            const k = ymd(d);
            const cd = clubDayDisplay(data, k, jsDayToIdx(d.getDay()));
            const has = !!days[k];
            return (
              <div key={k} className={"tp-dayrow" + (checked[k] ? " checked" : "") + (d.getDay() === 0 ? " sun" : d.getDay() === 6 ? " sat" : "")}>
                <input type="checkbox" checked={!!checked[k]} onChange={() => toggleCheck(k)} />
                <button className="tp-dayrow-main" onClick={() => setEditDate(k)}>
                  <span className="tp-dayrow-date">{d.getDate()}<small>{WD[d.getDay()]}</small></span>
                  {cd ? (
                    <span className="tp-dayrow-info">
                      <span className={"tp-dayrow-content" + (cd.kind === "off" ? " off" : cd.kind === "match" ? " match" : "")}>{cd.content}{cd.weekly && <em className="tp-weeklytag">曜日設定</em>}</span>
                      {(cd.place || cd.time) && <span className="tp-dayrow-sub">{[cd.place, cd.time].filter(Boolean).join("・")}</span>}
                      {cd.note && <span className="tp-dayrow-note">{cd.note}</span>}
                    </span>
                  ) : <span className="tp-dayrow-empty">未設定（タップで入力）</span>}
                  <Pencil size={13} className="tp-dayrow-pen" />
                </button>
                {has && <button className="tp-iconbtn tiny" onClick={() => clearDay(k)}><Trash2 size={13} /></button>}
              </div>
            );
          })}
        </div>
        <p className="tp-hint">日付をタップして入力。チェックを付けた日はまとめて「一括入力」できます。入力した日は月間カレンダーに4行（内容・場所・時間・備考）で表示されます。</p>
      </section>

      <section className="tp-card">
        <h4 className="tp-card-title"><Calendar size={15} /> 大会予定</h4>
        <div className="tp-toolbar" style={{ justifyContent: "flex-start" }}>
          <button className="tp-ghostbtn" onClick={() => setMatchImp(true)}><Upload size={14} /> 大会予定を画像で取り込み</button>
        </div>
        <ul className="tp-splist">
          {matchDays.length === 0 && club.specials.length === 0 && <li className="tp-empty">登録なし</li>}
          {matchDays.map((s) => { const d = parseYmd(s.key); return (
            <li key={s.key}>
              <span className="tp-sp-date" style={{ color: "#D9534F" }}>{d.getMonth() + 1}/{d.getDate()}<small>（{WD[d.getDay()]}）</small></span>
              <span className="tp-sp-title">{s.matchName || "大会"}<small>{s.time && ` ${s.time}`}{(s.place === "other" ? s.placeOther : s.place) && ` ／ ${s.place === "other" ? s.placeOther : s.place}`}</small></span>
              <button className="tp-iconbtn tiny" onClick={() => { setEditDate(s.key); }}><Pencil size={13} /></button>
              <button className="tp-iconbtn tiny" onClick={() => clearDay(s.key)}><Trash2 size={14} /></button>
            </li>
          ); })}
          {club.specials.slice().sort((a, b) => a.date.localeCompare(b.date)).map((s) => { const d = parseYmd(s.date); return (
            <li key={s.id}>
              <span className="tp-sp-date" style={{ color: "#D9534F" }}>{d.getMonth() + 1}/{d.getDate()}<small>（{WD[d.getDay()]}）</small></span>
              <span className="tp-sp-title">{s.title}<small>{s.start && ` ${s.start}–${s.end}`}{s.place && ` ／ ${s.place}`}</small></span>
              <button className="tp-iconbtn tiny" onClick={() => setClub({ specials: club.specials.filter((x) => x.id !== s.id) })}><Trash2 size={14} /></button>
            </li>
          ); })}
        </ul>
        <div className="tp-divider" />
        <div className="tp-sp-form">
          <input type="date" value={mt.date} onChange={(e) => setMt({ ...mt, date: e.target.value })} />
          <input value={mt.matchName} onChange={(e) => setMt({ ...mt, matchName: e.target.value })} placeholder="大会名（例 市総体）" />
          <input className="tp-timeinput" value={mt.start} onChange={(e) => setMt({ ...mt, start: e.target.value })} placeholder="開始" />
          <input className="tp-timeinput" value={mt.end} onChange={(e) => setMt({ ...mt, end: e.target.value })} placeholder="終了" />
          <input value={mt.placeOther} onChange={(e) => setMt({ ...mt, placeOther: e.target.value })} placeholder="会場" />
          <button className="tp-addbtn" onClick={addMatch}><Plus size={16} /></button>
        </div>
        <p className="tp-hint">大会は日毎スケジュール・カレンダー(赤)・今日/週間に反映されます。手入力・画像取り込みどちらでも登録できます。</p>
      </section>

      <section className="tp-card">
        <button className="tp-collapse" onClick={() => setShowWeekly((v) => !v)}>基本の練習曜日（日毎入力が無い日の表示に使用）<ChevronRight size={16} className={showWeekly ? "tp-rot90" : ""} /></button>
        {showWeekly && (
          <div className="tp-clubsched">
            {DAY_LABELS.map((lbl, di) => {
              const s = club.schedule[di] || { on: false };
              return (
                <div key={di} className={"tp-clubrow" + (s.on ? " on" : "")}>
                  <label className="tp-clubday"><input type="checkbox" checked={!!s.on} onChange={(e) => setSched(di, { on: e.target.checked })} /><span>{lbl}</span></label>
                  <input className="tp-timeinput" value={s.start || ""} onChange={(e) => setSched(di, { start: e.target.value })} placeholder="開始" disabled={!s.on} />
                  <span className="tp-tilde">–</span>
                  <input className="tp-timeinput" value={s.end || ""} onChange={(e) => setSched(di, { end: e.target.value })} placeholder="終了" disabled={!s.on} />
                  <input className="tp-placeinput" value={s.place || ""} onChange={(e) => setSched(di, { place: e.target.value })} placeholder="場所" disabled={!s.on} />
                  <input className="tp-noteinput" value={s.note || ""} onChange={(e) => setSched(di, { note: e.target.value })} placeholder="メモ" />
                </div>
              );
            })}
          </div>
        )}
      </section>

      {/* 単日エディタ */}
      <Modal open={!!editDate} title={editDate ? `${parseYmd(editDate).getMonth() + 1}月${parseYmd(editDate).getDate()}日（${WD[parseYmd(editDate).getDay()]}）の部活` : ""} onClose={() => setEditDate(null)}>
        {editDate && (
          <>
            <ClubFields val={days[editDate] || { kind: "practice", session: "after", place: "二中" }} set={(patch) => setDay(editDate, patch)} />
            <div className="tp-modal-actions">
              {days[editDate] ? <button className="tp-dangerbtn" onClick={() => { clearDay(editDate); setEditDate(null); }}><Trash2 size={14} /> 未設定に戻す</button> : <span />}
              <button className="tp-primarybtn" onClick={() => setEditDate(null)}><Check size={15} /> 完了</button>
            </div>
          </>
        )}
      </Modal>

      {/* 一括入力 */}
      <Modal open={bulkOpen} title={`一括入力（${checkedKeys.length}日）`} onClose={() => setBulkOpen(false)}>
        <p className="tp-hint" style={{ marginTop: 0 }}>チェックした{checkedKeys.length}日すべてに、以下の内容を上書きします。</p>
        <ClubFields val={bulkVal} set={(patch) => setBulkVal({ ...bulkVal, ...patch })} />
        <div className="tp-modal-actions"><button className="tp-ghostbtn" onClick={() => setBulkOpen(false)}>キャンセル</button><button className="tp-primarybtn" disabled={!checkedKeys.length} onClick={applyBulk}><Check size={15} /> {checkedKeys.length}日に適用</button></div>
      </Modal>

      <ImportModal open={matchImp} kind="match" ctx={{ year: y }} data={data} setData={setData} onClose={() => setMatchImp(false)} />
      <ImportModal open={clubImp} kind="club" ctx={{ year: y, month: m + 1 }} data={data} setData={setData} onClose={() => setClubImp(false)} />
    </div>
  );
}

/* ============================================================
   ROSTER — 名簿
   ============================================================ */
function RosterView({ data, setData, onPrintSeating, jump, showToast }) {
  const classes = data.classes;
  const [sel, setSel] = useState(data.homeroom || classes[0]);
  useEffect(() => { if (jump && jump.klass && classes.includes(jump.klass)) setSel(jump.klass); }, [jump && jump.at]);
  const [imp, setImp] = useState(false);
  const [paste, setPaste] = useState(false);
  const [seatCols, setSeatCols] = useState(6);
  const list = data.rosters[sel] || [];
  const setList = (fn) => setData((d) => ({ ...d, rosters: { ...d.rosters, [sel]: fn(d.rosters[sel] || []) } }));
  const addRow = () => setList((l) => [...l, { id: uid(), no: l.length + 1, name: "", kana: "", memo: "" }]);
  const upd = (id, patch) => setList((l) => l.map((r) => r.id === id ? { ...r, ...patch } : r));
  const del = (id) => setList((l) => l.filter((r) => r.id !== id));

  return (
    <div className="tp-view">
      <section className="tp-card">
        <div className="tp-roster-tabs">
          {classes.map((c) => <button key={c} className={"tp-rtab" + (sel === c ? " on" : "")} onClick={() => setSel(c)}>{c}{c === data.homeroom && <span className="tp-hr">担任</span>}</button>)}
        </div>
        <div className="tp-toolbar" style={{ justifyContent: "flex-start", gap: 8, flexWrap: "wrap" }}>
          <button className="tp-ghostbtn" onClick={() => setImp(true)}><Upload size={14} /> {sel} を画像から取り込み</button>
          <button className="tp-ghostbtn" onClick={() => setPaste(true)}><ClipboardList size={14} /> 貼り付けで一括入力</button>
          <span className="tp-seat-ctl">座席 列数
            <select value={seatCols} onChange={(e) => setSeatCols(Number(e.target.value))}>{[4, 5, 6, 7, 8].map((n) => <option key={n} value={n}>{n}</option>)}</select>
            <button className="tp-ghostbtn sm" disabled={list.length === 0} onClick={() => onPrintSeating(sel, seatCols)}><Printer size={14} /> 座席表</button>
          </span>
        </div>
        <div className="tp-roster">
          <div className="tp-roster-headrow"><span>No</span><span>氏名</span><span>ふりがな</span><span>メモ</span><span /></div>
          {list.length === 0 && <p className="tp-empty">生徒を追加してください。「画像から取り込み」やExcel取り込み（設定）も使えます。</p>}
          {list.map((r) => (
            <div key={r.id} className="tp-roster-row">
              <input className="tp-no" type="number" value={r.no} onChange={(e) => upd(r.id, { no: e.target.value })} />
              <input value={r.name} onChange={(e) => upd(r.id, { name: e.target.value })} placeholder="氏名" />
              <input value={r.kana} onChange={(e) => upd(r.id, { kana: e.target.value })} placeholder="ふりがな" />
              <input value={r.memo} onChange={(e) => upd(r.id, { memo: e.target.value })} placeholder="係・配慮事項 など" />
              <button className="tp-iconbtn tiny" onClick={() => del(r.id)}><Trash2 size={14} /></button>
            </div>
          ))}
        </div>
        <button className="tp-primarybtn" onClick={addRow}><Plus size={15} /> 生徒を追加</button>
      </section>
      <ImportModal open={imp} kind="roster" ctx={{ klass: sel }} data={data} setData={setData} onClose={() => setImp(false)} />
      <RosterPasteModal open={paste} klass={sel} data={data} setData={setData} onClose={() => setPaste(false)} showToast={showToast} />
    </div>
  );
}

/* 名簿：Excel等からの貼り付けを一括入力（タブ/カンマ/連続スペース区切り） */
function parseRosterPaste(text) {
  const out = [];
  String(text || "").split(/\r?\n/).forEach((raw) => {
    const line = raw.replace(/\u3000/g, " ").trim(); if (!line) return;
    let no = "", name = "", kana = "", memo = "";
    if (line.includes("\t") || line.includes(",")) {
      let cols = line.split(line.includes("\t") ? "\t" : ",").map((c) => c.trim());
      if (/^\d{1,3}$/.test(cols[0])) { no = cols[0]; name = cols[1] || ""; kana = cols[2] || ""; memo = cols.slice(3).join(" "); }
      else { name = cols[0] || ""; kana = cols[1] || ""; memo = cols.slice(2).join(" "); }
    } else {
      // 区切りなし：先頭の番号だけ拾い、残り全部を氏名（姓 名の空白を保持）
      const m = line.match(/^(\d{1,3})[\s.．、]+(.+)$/);
      if (m) { no = m[1]; name = m[2].trim(); } else { name = line; }
    }
    if (!name || /^(no|番号|氏名|名前|生徒)$/i.test(name)) return;
    out.push({ no, name, kana, memo });
  });
  return out;
}
function RosterPasteModal({ open, klass, data, setData, onClose, showToast }) {
  const [text, setText] = useState("");
  const [rows, setRows] = useState(null);
  const [replace, setReplace] = useState(false);
  useEffect(() => { if (open) { setText(""); setRows(null); setReplace(false); } }, [open]);
  const run = () => setRows(parseRosterPaste(text));
  const imp = () => {
    if (!rows || !rows.length) { onClose(); return; }
    const prev = data;
    setData((d) => {
      const cur = d.rosters[klass] || [];
      const added = rows.map((r, i) => ({ id: uid(), no: r.no || (replace ? i + 1 : cur.length + i + 1), name: r.name, kana: r.kana, memo: r.memo }));
      const next = replace ? added : [...cur, ...added];
      return { ...d, rosters: { ...d.rosters, [klass]: next } };
    });
    showToast && showToast(`${klass} に${rows.length}名を${replace ? "置き換え" : "追加"}しました`, prev);
    onClose();
  };
  return (
    <Modal open={open} title={`${klass} に貼り付けで一括入力`} onClose={onClose} wide>
      <p className="tp-hint" style={{ marginTop: 0 }}>Excelの列をコピーして貼り付け、または「氏名（＋ふりがな・メモ）」を1行ずつ。区切りは<b>タブ・カンマ・スペース</b>を自動判定。先頭が数字ならNoとして扱います。</p>
      <label className="tp-field"><span>貼り付け</span>
        <textarea rows={7} value={text} onChange={(e) => setText(e.target.value)} placeholder={"例）\n1\t山田 太郎\tやまだ たろう\t図書委員\n2\t佐藤 花子\tさとう はなこ\n（氏名だけでも可）\n田中 一郎"} />
      </label>
      <div className="tp-field-row" style={{ alignItems: "center" }}>
        <label className="tp-switch" style={{ margin: 0 }}><input type="checkbox" checked={replace} onChange={(e) => setReplace(e.target.checked)} /><span>この名簿を置き換える（既存を消す）</span></label>
        <button className="tp-primarybtn" style={{ marginLeft: "auto" }} onClick={run}><Loader size={15} /> 読み取り</button>
      </div>
      {rows && (
        <div className="tp-ocr-result">
          <div className="tp-mtg-grouphead">読み取り結果 {rows.length}名{rows.length === 0 && "（氏名の行が見つかりませんでした）"}</div>
          {rows.map((r, i) => (
            <div key={i} className="tp-ocr-row">
              <span className="tp-ocr-date">{r.no || i + 1}</span>
              <span className="tp-ocr-title">{r.name}{r.kana ? "（" + r.kana + "）" : ""}{r.memo ? " / " + r.memo : ""}</span>
            </div>
          ))}
        </div>
      )}
      <div className="tp-modal-actions">
        <button className="tp-ghostbtn" onClick={onClose}>キャンセル</button>
        <button className="tp-primarybtn" disabled={!rows || rows.length === 0} onClick={imp}><Plus size={15} /> 名簿に取り込む</button>
      </div>
    </Modal>
  );
}

/* ============================================================
   DUTIES — 校務・日課
   ============================================================ */
function DutiesView({ data, setData }) {
  const [duty, setDuty] = useState({ title: "", dayIdxs: [], time: "", place: "", note: "" });
  const toggleDay = (i) => setDuty((d) => ({ ...d, dayIdxs: d.dayIdxs.includes(i) ? d.dayIdxs.filter((x) => x !== i) : [...d.dayIdxs, i].sort() }));
  const addDuty = () => { if (!duty.title.trim()) return; setData((d) => ({ ...d, duties: [...d.duties, { id: uid(), ...duty }] })); setDuty({ title: "", dayIdxs: [], time: "", place: "", note: "" }); };

  const [rt, setRt] = useState({ time: "", title: "" });
  const addRoutine = () => { if (!rt.title.trim()) return; setData((d) => ({ ...d, routine: [...d.routine, { id: uid(), ...rt }].sort((a, b) => t2m(a.time) - t2m(b.time)) })); setRt({ time: "", title: "" }); };

  return (
    <div className="tp-view">
      <section className="tp-card">
        <h4 className="tp-card-title"><ClipboardList size={15} /> 校務・会議（曜日固定）</h4>
        <ul className="tp-dutylist">
          {data.duties.length === 0 && <li className="tp-empty">登録なし</li>}
          {data.duties.map((d) => (
            <li key={d.id}>
              <span className="tp-duty-days">{d.dayIdxs.length ? d.dayIdxs.map((i) => DAY_LABELS[i]).join("・") : "毎日"}</span>
              <span className="tp-duty-title">{d.title}<small>{d.time && ` ${d.time}`}{d.place && ` ／ ${d.place}`}{d.note && ` ／ ${d.note}`}</small></span>
              <button className="tp-iconbtn tiny" onClick={() => setData((dd) => ({ ...dd, duties: dd.duties.filter((x) => x.id !== d.id) }))}><Trash2 size={14} /></button>
            </li>
          ))}
        </ul>
        <div className="tp-divider" />
        <label className="tp-field"><span>校務・会議名</span><input value={duty.title} onChange={(e) => setDuty({ ...duty, title: e.target.value })} placeholder="例）学年会 / 生徒指導部会" /></label>
        <div className="tp-daypick">{DAY_LABELS.map((l, i) => <button key={i} className={"tp-daychip" + (duty.dayIdxs.includes(i) ? " on" : "")} onClick={() => toggleDay(i)}>{l}</button>)}</div>
        <div className="tp-field-row">
          <label className="tp-field"><span>時刻</span><input value={duty.time} onChange={(e) => setDuty({ ...duty, time: e.target.value })} placeholder="15:40" /></label>
          <label className="tp-field"><span>場所</span><input value={duty.place} onChange={(e) => setDuty({ ...duty, place: e.target.value })} placeholder="会議室" /></label>
        </div>
        <label className="tp-field"><span>メモ</span><input value={duty.note} onChange={(e) => setDuty({ ...duty, note: e.target.value })} placeholder="第1・3週 など" /></label>
        <button className="tp-primarybtn" onClick={addDuty}><Plus size={15} /> 追加</button>
      </section>

      <section className="tp-card">
        <h4 className="tp-card-title"><Clock size={15} /> 毎日の日課</h4>
        <p className="tp-hint">ここで設定した項目は毎日の「今日」画面のタイムラインに並びます。</p>
        <ul className="tp-routinelist">
          {data.routine.map((r) => (
            <li key={r.id}><b>{r.time}</b><span>{r.title}</span><button className="tp-iconbtn tiny" onClick={() => setData((d) => ({ ...d, routine: d.routine.filter((x) => x.id !== r.id) }))}><Trash2 size={13} /></button></li>
          ))}
        </ul>
        <div className="tp-sp-form">
          <input className="tp-timeinput" value={rt.time} onChange={(e) => setRt({ ...rt, time: e.target.value })} placeholder="07:50" />
          <input value={rt.title} onChange={(e) => setRt({ ...rt, title: e.target.value })} placeholder="日課の内容" />
          <button className="tp-addbtn" onClick={addRoutine}><Plus size={16} /></button>
        </div>
      </section>
    </div>
  );
}

/* ============================================================
   SETTINGS
   ============================================================ */
/* ============================================================
   GuideModal — 使い方（チュートリアル）＆ 機能の説明
   ============================================================ */
const TUTORIAL_STEPS = [
  { icon: Home, title: "ようこそ", body: "名前ごとにデータが保存されます。まず「今日」タブで当日の授業・部活・予定・やることを確認しましょう。手書きメモも残せます。" },
  { icon: CalendarDays, title: "予定を入れる", body: "「カレンダー」タブで上部の期間（週/月/3か月/6か月/年）を切り替えて確認。日付やコマをタップして追加、行事は月表示の「＋」から。授業のコマは週表示でタップして教科・クラスを設定します。" },
  { icon: Dumbbell, title: "部活・名簿・授業", body: "「部活」で練習/大会/休みを登録（まとめて入力も可）。「名簿」で生徒と座席表。「授業」でテスト・ロードマップ（計画）・進度（実績）を管理します。" },
  { icon: FileText, title: "職員会議", body: "会議ごとに書き込みと資料写真を保存。資料はタップして直接ペンで書き込めます。件名・書き込み・資料の説明は絞り込みでき、上部の全体検索からその会議へジャンプできます。" },
  { icon: Upload, title: "取り込み", body: "設定＞データ管理から、Excel/CSV・カレンダー(ICS)・画像で予定や名簿を取り込めます。列見出しや日付の表記ゆれも自動で吸収します。" },
  { icon: ShieldAlert, title: "バックアップ（大切）", body: "データは端末内に保存されます。設定＞「バックアップを書き出し」でiCloud/ファイル等へ定期保存を。自動バックアップからの復元もできます。" },
];
const FEATURE_LIST = [
  { icon: Home, name: "今日", desc: "当日の授業・部活・予定・やること（思いつき・突発）・メモ・手書きを一覧。" },
  { icon: CalendarDays, name: "カレンダー", desc: "上部で期間を切替（週／月／3か月／6か月／年）。予定の追加、カテゴリ表示切替、週案・時間割の取り込み、印刷・PDF保存。" },
  { icon: BookOpen, name: "授業", desc: "集計・記録／テスト／ロードマップ（計画）／進度（実績）。教科書（英語）はロードマップ内から任意で。" },
  { icon: Dumbbell, name: "部活", desc: "練習/大会/休みの登録、まとめ入力、月ごとのPDF保存、基本曜日テンプレ。" },
  { icon: Users, name: "名簿", desc: "生徒名簿の編集・画像/Excel取り込み・座席表の印刷。" },
  { icon: ClipboardList, name: "校務", desc: "曜日固定の校務・当番などの登録。" },
  { icon: FileText, name: "職員会議", desc: "会議の記録・資料写真・直接書き込み・一覧の絞り込み・PDF保存・複数まとめてPDF。日付順/カテゴリ順で整理。" },
  { icon: CheckSquare, name: "ToDo（右上）", desc: "計画的な大きいタスク。カテゴリ・締切つき、締切は端末カレンダーへ登録。全画面/サイドバー/最小化を切替。" },
  { icon: Search, name: "全体検索（上部）", desc: "予定・授業・部活・名簿・会議・ToDo・メモを横断検索し、種類で絞り込み、タップでジャンプ。" },
];
function GuideModal({ mode, onClose }) {
  const [step, setStep] = useState(0);
  useEffect(() => { if (mode) setStep(0); }, [mode]);
  if (!mode) return null;
  if (mode === "tutorial") {
    const s = TUTORIAL_STEPS[step]; const Ic = s.icon; const last = step === TUTORIAL_STEPS.length - 1;
    return (
      <div className="tp-modal-back">
        <div className="tp-modal tp-guide" onClick={(e) => e.stopPropagation()}>
          <div className="tp-modal-head"><h3>使い方（{step + 1}/{TUTORIAL_STEPS.length}）</h3><button className="tp-iconbtn" onClick={onClose}><X size={18} /></button></div>
          <div className="tp-guide-step">
            <div className="tp-guide-ic"><Ic size={26} /></div>
            <h4>{s.title}</h4>
            <p>{s.body}</p>
          </div>
          <div className="tp-guide-dots">{TUTORIAL_STEPS.map((_, i) => <span key={i} className={"tp-guide-dot" + (i === step ? " on" : "")} />)}</div>
          <div className="tp-modal-actions">
            <button className="tp-ghostbtn" disabled={step === 0} onClick={() => setStep((v) => Math.max(0, v - 1))}>戻る</button>
            {last ? <button className="tp-primarybtn" onClick={onClose}><Check size={15} /> 始める</button>
              : <button className="tp-primarybtn" onClick={() => setStep((v) => v + 1)}>次へ</button>}
          </div>
        </div>
      </div>
    );
  }
  return (
    <div className="tp-modal-back">
      <div className="tp-modal tp-guide" onClick={(e) => e.stopPropagation()}>
        <div className="tp-modal-head"><h3>機能の説明</h3><button className="tp-iconbtn" onClick={onClose}><X size={18} /></button></div>
        <div className="tp-feat-list">
          {FEATURE_LIST.map((f) => { const Ic = f.icon; return (
            <div key={f.name} className="tp-feat-row"><div className="tp-feat-ic"><Ic size={18} /></div><div><b>{f.name}</b><span>{f.desc}</span></div></div>
          ); })}
          <div className="tp-feat-row"><div className="tp-feat-ic"><Download size={18} /></div><div><b>印刷・PDF / 共有</b><span>各タブの印刷・PDFボタンで配布物を作成。ICSでカレンダー連携も。</span></div></div>
          <div className="tp-feat-row"><div className="tp-feat-ic"><ShieldAlert size={18} /></div><div><b>バックアップ・復元</b><span>設定＞データ管理で書き出し/読み込み、自動バックアップから復元。</span></div></div>
          <div className="tp-feat-row"><div className="tp-feat-ic"><Settings size={18} /></div><div><b>表示・アカウント</b><span>文字サイズ・高コントラスト、名前の切替、最新に更新。</span></div></div>
        </div>
        <div className="tp-modal-actions"><span /><button className="tp-primarybtn" onClick={onClose}><Check size={15} /> 閉じる</button></div>
      </div>
    </div>
  );
}

function SettingsModal({ open, onClose, data, setData, user, onLogout, onExported, pushBackup, showToast }) {
  const [newSub, setNewSub] = useState(""); const [newClass, setNewClass] = useState("");
  const fileRef = useRef(null);
  const icsRef = useRef(null);
  const [impMsg, setImpMsg] = useState("");
  const [xls, setXls] = useState(false);
  const [otx, setOtx] = useState(false);
  const [pending, setPending] = useState(null); // 'timetable' | 'sample'
  const [guide, setGuide] = useState(null); // 'tutorial' | 'features'
  const [backups, setBackups] = useState([]);
  const [restoreIdx, setRestoreIdx] = useState(null);
  useEffect(() => { if (open) { (async () => setBackups((await loadStore(`${NS}:backups:${user}`)) || []))(); } }, [open, user]);
  const doClear = () => {
    const prev = data;
    if (pending === "timetable") { setData((d) => ({ ...d, timetable: {} })); showToast && showToast("時間割を空にしました", prev); }
    else if (pending === "club") { setData((d) => ({ ...d, club: { ...d.club, days: {}, specials: [], schedule: {} } })); showToast && showToast("部活予定を空にしました", prev); }
    else if (pending === "sample") { setData((d) => ({
      ...d, timetable: {}, lessonLog: {}, events: [], todos: [], dayMemo: {}, weeklyManual: {},
      club: { ...d.club, days: {}, specials: [], schedule: {} },
      textbook: { ...(d.textbook || {}), units: [] }, tests: [], testProgress: {}, duties: [], routine: [], rosters: {},
      subjects: [], classes: [], homeroom: "",
    })); showToast && showToast("サンプルを消去しました", prev); }
    setPending(null);
  };
  const setMeta = (patch) => setData((d) => ({ ...d, meta: { ...d.meta, ...patch } }));
  return (
    <Modal open={open} title="設定" onClose={onClose} wide>
      <div className="tp-field-row">
        <label className="tp-field"><span>氏名</span><input value={data.meta.teacher} onChange={(e) => setMeta({ teacher: e.target.value })} placeholder="山田 太郎" /></label>
        <label className="tp-field"><span>年度</span><input value={data.meta.year} onChange={(e) => setMeta({ year: e.target.value })} /></label>
      </div>
      <label className="tp-check"><input type="checkbox" checked={data.meta.includeSat} onChange={(e) => setMeta({ includeSat: e.target.checked })} /> 時間割に土曜を表示する</label>

      <div className="tp-divider" />
      <h4 className="tp-card-title">学期（時数カウントの起点）</h4>
      <p className="tp-hint" style={{ marginTop: 0 }}>各学期の開始日を設定すると、「今日」画面や授業タブで、その学期の実施時数（第◯時）を数えます。</p>
      {(data.terms || []).map((t, i) => (
        <div key={t.id || i} className="tp-term-edit">
          <input className="tp-term-name" value={t.name} onChange={(e) => setData((d) => { const ts = [...d.terms]; ts[i] = { ...ts[i], name: e.target.value }; return { ...d, terms: ts }; })} placeholder="学期名" />
          <input type="date" value={t.start} onChange={(e) => setData((d) => { const ts = [...d.terms]; ts[i] = { ...ts[i], start: e.target.value }; return { ...d, terms: ts }; })} />
          <button className="tp-iconbtn tiny" onClick={() => setData((d) => ({ ...d, terms: d.terms.filter((_, j) => j !== i) }))}><Trash2 size={13} /></button>
        </div>
      ))}
      <button className="tp-ghostbtn" onClick={() => setData((d) => ({ ...d, terms: [...(d.terms || []), { id: uid(), name: `${(d.terms || []).length + 1}学期`, start: "" }] }))}><Plus size={14} /> 学期を追加</button>

      <div className="tp-divider" />
      <h4 className="tp-card-title">授業時間（チャイム）</h4>
      {data.periods.map((p, i) => (
        <div key={i} className="tp-period-edit">
          <input className="tp-plabel" value={p.label} onChange={(e) => setData((d) => { const ps = [...d.periods]; ps[i] = { ...ps[i], label: e.target.value }; return { ...d, periods: ps }; })} />限
          <input className="tp-timeinput" value={p.start} onChange={(e) => setData((d) => { const ps = [...d.periods]; ps[i] = { ...ps[i], start: e.target.value }; return { ...d, periods: ps }; })} />
          <span className="tp-tilde">–</span>
          <input className="tp-timeinput" value={p.end} onChange={(e) => setData((d) => { const ps = [...d.periods]; ps[i] = { ...ps[i], end: e.target.value }; return { ...d, periods: ps }; })} />
        </div>
      ))}

      <div className="tp-divider" />
      <h4 className="tp-card-title">教科と色</h4>
      <div className="tp-subedit">
        {data.subjects.map((s, i) => (
          <div key={i} className="tp-subrow">
            <input type="color" value={s.color} onChange={(e) => setData((d) => { const ss = [...d.subjects]; ss[i] = { ...ss[i], color: e.target.value }; return { ...d, subjects: ss }; })} />
            <input value={s.name} onChange={(e) => setData((d) => { const ss = [...d.subjects]; ss[i] = { ...ss[i], name: e.target.value }; return { ...d, subjects: ss }; })} />
            <label className="tp-sub-testflag" title="テストを実施しロードマップを作る教科"><input type="checkbox" checked={!!s.test} onChange={(e) => setData((d) => { const ss = [...d.subjects]; ss[i] = { ...ss[i], test: e.target.checked }; return { ...d, subjects: ss }; })} />テスト実施</label>
            <button className="tp-iconbtn tiny" onClick={() => setData((d) => ({ ...d, subjects: d.subjects.filter((_, j) => j !== i) }))}><Trash2 size={13} /></button>
          </div>
        ))}
        <div className="tp-sp-form"><input value={newSub} onChange={(e) => setNewSub(e.target.value)} placeholder="教科を追加" /><button className="tp-addbtn" onClick={() => { if (!newSub.trim()) return; setData((d) => ({ ...d, subjects: [...d.subjects, { name: newSub.trim(), color: "#3E9BC9" }] })); setNewSub(""); }}><Plus size={16} /></button></div>
      </div>

      <div className="tp-divider" />
      <h4 className="tp-card-title">クラス</h4>
      <div className="tp-chips-edit">
        {data.classes.map((c) => <span key={c} className="tp-editchip">{c}<button onClick={() => setData((d) => ({ ...d, classes: d.classes.filter((x) => x !== c) }))}><X size={12} /></button></span>)}
      </div>
      <div className="tp-sp-form"><input value={newClass} onChange={(e) => setNewClass(e.target.value)} placeholder="例）1-3" /><button className="tp-addbtn" onClick={() => { if (!newClass.trim()) return; setData((d) => ({ ...d, classes: [...d.classes, newClass.trim()] })); setNewClass(""); }}><Plus size={16} /></button></div>
      <label className="tp-field"><span>担任クラス</span><select value={data.homeroom} onChange={(e) => setData((d) => ({ ...d, homeroom: e.target.value }))}><option value="">なし</option>{data.classes.map((c) => <option key={c} value={c}>{c}</option>)}</select></label>

      <div className="tp-divider" />
      <h4 className="tp-card-title">週コマ数（手入力）</h4>
      <p className="tp-hint">担当クラスごとの週あたり授業コマ数。入力すると「授業」画面の集計と、今後の年間計画の基準になります。空欄なら時間割から自動算出。</p>
      <div className="tp-weekedit">
        {data.classes.map((c) => (
          <label key={c} className="tp-weekedit-row"><span>{c}</span>
            <input type="number" min="0" value={data.weeklyManual?.[c] ?? ""} placeholder="—"
              onChange={(e) => setData((d) => { const wm = { ...(d.weeklyManual || {}) }; if (e.target.value === "") delete wm[c]; else wm[c] = Number(e.target.value); return { ...d, weeklyManual: wm }; })} />
            <small>コマ / 週</small>
          </label>
        ))}
      </div>

      <div className="tp-divider" />
      <h4 className="tp-card-title"><FileText size={15} /> 使い方・ヘルプ</h4>
      <p className="tp-hint" style={{ marginTop: 0 }}>初めての方や機能を確認したいときに。</p>
      <div className="tp-toolbar" style={{ justifyContent: "flex-start", gap: 8 }}>
        <button className="tp-ghostbtn" onClick={() => setGuide("tutorial")}><BookOpen size={14} /> 使い方（チュートリアル）</button>
        <button className="tp-ghostbtn" onClick={() => setGuide("features")}><ClipboardList size={14} /> 機能の説明</button>
      </div>

      <div className="tp-divider" />
      <h4 className="tp-card-title">表示</h4>
      <div className="tp-field"><span>文字サイズ</span><Seg options={[{ v: "S", label: "小" }, { v: "M", label: "標準" }, { v: "L", label: "大" }]} value={data.meta.fontScale || "M"} onChange={(v) => setData((d) => ({ ...d, meta: { ...d.meta, fontScale: v } }))} /></div>
      <label className="tp-switch"><input type="checkbox" checked={!!data.meta.contrast} onChange={(e) => setData((d) => ({ ...d, meta: { ...d.meta, contrast: e.target.checked } }))} /><span>高コントラスト（文字を濃く）</span></label>
      <label className="tp-switch"><input type="checkbox" checked={data.meta.theme === "dark"} onChange={(e) => setData((d) => ({ ...d, meta: { ...d.meta, theme: e.target.checked ? "dark" : "light" } }))} /><span>ダークモード</span></label>

      <div className="tp-divider" />
      <h4 className="tp-card-title">データ管理（バックアップ / 取り込み）</h4>
      <p className="tp-hint" style={{ marginTop: 0 }}>通信不要。JSONファイルで現在の予定を書き出し・読み込みできます。読み込みは今のデータに追加・統合されます。</p>
      <div className="tp-toolbar" style={{ justifyContent: "flex-start", gap: 8 }}>
        <button className="tp-ghostbtn" onClick={async () => {
          try {
            const payload = { ...data, __v: SCHEMA, __app: APP_VERSION, __sketches: collectSketches(user), __meetingImages: await collectMeetingImages(user) };
            const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
            const url = URL.createObjectURL(blob);
            const a = document.createElement("a"); a.href = url; a.download = `TeachingPlanners_${user || "data"}_${ymd(new Date())}.json`; document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(url), 1000);
          } catch (e) {}
          saveStore(`${NS}:lastExport:${user}`, new Date().toISOString());
          onExported && onExported();
          setImpMsg("書き出しました（手書き・会議資料も含みます）。iCloud/ファイルなど端末外にも保存しておくと安心です。");
        }}><Download size={14} /> バックアップを書き出し</button>
        <button className="tp-ghostbtn" onClick={() => fileRef.current?.click()}><Upload size={14} /> バックアップを読み込み</button>
        <button className="tp-ghostbtn" onClick={() => setXls(true)}><Upload size={14} /> Excel/CSVから取り込み</button>
        <button className="tp-ghostbtn" onClick={() => setOtx(true)}><Upload size={14} /> 写真/テキストから取り込み（オフライン）</button>
        <button className="tp-ghostbtn" onClick={() => {
          try {
            const evs = [...(data.events || [])];
            Object.entries(data.club?.days || {}).forEach(([dk, v]) => { if ((v.kind || "") === "match") evs.push({ date: dk, title: `【大会】${v.matchName || "大会"}`, time: v.time || "" }); });
            (data.club?.specials || []).forEach((s) => evs.push({ date: s.date, title: `【大会】${s.title}`, time: s.start || "" }));
            const ics = buildICS(evs);
            const blob = new Blob([ics], { type: "text/calendar" }); const url = URL.createObjectURL(blob);
            const a = document.createElement("a"); a.href = url; a.download = `TeachingPlanners_${ymd(new Date())}.ics`; document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(url), 1000);
            setImpMsg("ICSを書き出しました。Google/Appleカレンダーに読み込めます（行事＋大会）。");
          } catch (e) { setImpMsg("ICS書き出しに失敗しました。"); }
        }}><Download size={14} /> カレンダー(ICS)を書き出し</button>
        <button className="tp-ghostbtn" onClick={() => icsRef.current?.click()}><Upload size={14} /> カレンダー(ICS)を読み込み</button>
        <input ref={icsRef} type="file" accept=".ics,text/calendar" style={{ display: "none" }} onChange={(e) => {
          const f = e.target.files?.[0]; if (!f) return; const r = new FileReader();
          r.onload = () => {
            try {
              const parsed = parseICS(String(r.result)); const prev = data;
              setData((d) => { const ex = new Set(d.events.map((ev) => ev.date + "|" + ev.title)); return { ...d, events: [...d.events, ...parsed.filter((ev) => !ex.has(ev.date + "|" + ev.title)).map((ev) => ({ id: uid(), ...ev }))] }; });
              setImpMsg(`ICSから${parsed.length}件を読み込みました。`); showToast && showToast(`ICSから${parsed.length}件取り込みました`, prev);
            } catch (err) { setImpMsg("ICS読み込みに失敗しました。"); }
            e.target.value = "";
          };
          r.readAsText(f);
        }} />
        <input ref={fileRef} type="file" accept="application/json,.json" style={{ display: "none" }} onChange={(e) => {
          const f = e.target.files?.[0]; if (!f) return;
          const r = new FileReader();
          r.onload = () => {
            try {
              const obj = JSON.parse(String(r.result));
              const prev = data;
              if (obj.__sketches) { try { restoreSketches(user, obj.__sketches); } catch (e) {} }
              if (obj.__meetingImages) { try { restoreMeetingImages(user, obj.__meetingImages); } catch (e) {} }
              const { __sketches, __meetingImages, __v, __app, ...rest } = obj;
              setData((d) => {
                const next = { ...d };
                for (const k of Object.keys(rest)) {
                  if (k === "club") next.club = { ...d.club, ...obj.club, days: { ...(d.club.days || {}), ...(obj.club.days || {}) }, schedule: { ...d.club.schedule, ...(obj.club.schedule || {}) }, specials: obj.club.specials || d.club.specials };
                  else if (k === "events") { const ex = new Set(d.events.map((ev) => ev.date + "|" + ev.title)); next.events = [...d.events, ...(obj.events || []).filter((ev) => !ex.has(ev.date + "|" + ev.title)).map((ev) => ({ id: uid(), ...ev }))]; }
                  else if (k === "timetable") next.timetable = { ...d.timetable, ...obj.timetable };
                  else if (k === "lessonLog") next.lessonLog = { ...d.lessonLog, ...obj.lessonLog };
                  else if (k === "meta") next.meta = { ...d.meta, ...obj.meta, teacher: d.meta.teacher };
                  else next[k] = obj[k];
                }
                return next;
              });
              setImpMsg("読み込みました。");
              showToast && showToast("JSONを読み込みました", prev);
            } catch (err) { setImpMsg("読み込みに失敗しました（JSON形式を確認してください）。"); }
            e.target.value = "";
          };
          r.readAsText(f);
        }} />
      </div>
      {impMsg && <p className="tp-hint" style={{ color: "#2C7CA6", fontWeight: 700 }}>{impMsg}</p>}

      <div className="tp-divider" />
      <h4 className="tp-card-title">自動バックアップ（この端末内）</h4>
      <p className="tp-hint" style={{ marginTop: 0 }}>使用中に自動で控えを保存します（最新5件）。誤って消去・上書きしたときに復元できます。※端末のデータ削除には対応できないため、上の「書き出し」も併用してください。</p>
      <div className="tp-toolbar" style={{ justifyContent: "flex-start", gap: 8 }}>
        <button className="tp-ghostbtn" onClick={async () => { await pushBackup(user, data); setBackups((await loadStore(`${NS}:backups:${user}`)) || []); setImpMsg("バックアップを作成しました。"); }}><ShieldAlert size={14} /> 今すぐバックアップ</button>
      </div>
      <ul className="tp-backup-list">
        {backups.length === 0 && <li className="tp-empty">まだありません</li>}
        {backups.map((b, i) => (
          <li key={i}>
            <span className="tp-backup-at">{new Date(b.at).toLocaleString("ja-JP", { month: "numeric", day: "numeric", hour: "2-digit", minute: "2-digit" })}</span>
            <button className="tp-ghostbtn sm" onClick={() => setRestoreIdx(i)}><RotateCcw size={13} /> この状態に戻す</button>
          </li>
        ))}
      </ul>

      <div className="tp-divider" />
      <h4 className="tp-card-title">初期化（サンプルを消す）</h4>
      <p className="tp-hint" style={{ marginTop: 0 }}>最初から入っているサンプル（数学の授業・時間割・基本の部活曜日・教科/クラスなど）を消せます。取り込んだ予定は消えません（サンプル全消去では消えます）。</p>
      <div className="tp-toolbar" style={{ justifyContent: "flex-start", gap: 8 }}>
        <button className="tp-ghostbtn" onClick={() => setPending("timetable")}><Trash2 size={14} /> 時間割を空にする</button>
        <button className="tp-ghostbtn" onClick={() => setPending("club")}><Trash2 size={14} /> 部活予定を空にする</button>
        <button className="tp-dangerbtn" style={{ marginTop: 0 }} onClick={() => setPending("sample")}><Trash2 size={14} /> サンプルを全消去（まっさら）</button>
      </div>

      <div className="tp-divider" />
      <h4 className="tp-card-title">アカウント</h4>
      <p className="tp-hint" style={{ marginTop: 0 }}>ログイン中：<b>{user}</b>。データはこの名前で端末に保存され、次回ログイン時に引き継がれます。</p>
      <button className="tp-ghostbtn" onClick={onLogout}><Home size={14} /> ログアウト（名前を切り替える）</button>
      <div className="tp-toolbar" style={{ justifyContent: "flex-start", marginTop: 8 }}>
        <button className="tp-ghostbtn" onClick={async () => { try { if (window.caches) { const ks = await caches.keys(); await Promise.all(ks.map((k) => caches.delete(k))); } if (navigator.serviceWorker) { const rs = await navigator.serviceWorker.getRegistrations(); rs.forEach((r) => r.update && r.update()); } } catch (e) {} try { window.location.reload(); } catch (e) {} }}><RotateCcw size={14} /> 最新に更新（再読み込み）</button>
      </div>
      <p className="tp-hint" style={{ marginTop: 12, opacity: .8 }}>Teaching Partner for YOU v{APP_VERSION}（オフライン・端末内保存）</p>
      <ExcelImportModal open={xls} onClose={() => setXls(false)} data={data} setData={setData} showToast={showToast} />
      <OfflineTextImportModal open={otx} onClose={() => setOtx(false)} data={data} setData={setData} showToast={showToast} />
      <GuideModal mode={guide} onClose={() => setGuide(null)} />
      <Modal open={!!pending} title="確認" onClose={() => setPending(null)}>
        <p style={{ fontSize: 14, lineHeight: 1.7 }}>
          {pending === "timetable" ? "時間割をすべて空にします。よろしいですか？" : pending === "club" ? "部活予定（日毎・大会・基本曜日）をすべて空にします。よろしいですか？" : "サンプルの授業・時間割・部活・教科/クラス・名簿・行事などをすべて消去して、まっさらにします。よろしいですか？（元に戻せません）"}
        </p>
        <div className="tp-modal-actions"><button className="tp-ghostbtn" onClick={() => setPending(null)}>キャンセル</button><button className="tp-dangerbtn" style={{ marginTop: 0 }} onClick={doClear}><Trash2 size={14} /> 消去する</button></div>
      </Modal>
      <Modal open={restoreIdx !== null} title="バックアップから復元" onClose={() => setRestoreIdx(null)}>
        <p style={{ fontSize: 14, lineHeight: 1.7 }}>{restoreIdx !== null && backups[restoreIdx] && `${new Date(backups[restoreIdx].at).toLocaleString("ja-JP")} の状態に戻します。現在のデータは置き換わります。よろしいですか？`}</p>
        <div className="tp-modal-actions"><button className="tp-ghostbtn" onClick={() => setRestoreIdx(null)}>キャンセル</button><button className="tp-primarybtn" onClick={() => { const prev = data; setData(backups[restoreIdx].data); showToast && showToast("復元しました", prev); setRestoreIdx(null); setImpMsg("復元しました。"); }}><RotateCcw size={14} /> 復元する</button></div>
      </Modal>
    </Modal>
  );
}

/* ============================================================
   LoginScreen — 名前でログイン（オフライン）
   ============================================================ */
function LoginScreen({ profiles, onLogin, onDelete }) {
  const [name, setName] = useState("");
  const submit = () => { if (name.trim()) onLogin(name.trim()); };
  return (
    <div className="tp-login">
      <div className="tp-login-card">
        <div className="tp-login-brand">
          <Logo size={54} />
          <div className="tp-appname" style={{ fontSize: 26 }}>Teaching Partner for YOU</div>
          <div className="tp-apptag">for teachers</div>
        </div>
        <p className="tp-login-lead">名前でログインすると、その名前で予定が保存され、次回そのまま引き継げます。</p>

        {profiles.length > 0 && (
          <div className="tp-login-profiles">
            <div className="tp-login-sub">保存済みのユーザー</div>
            {profiles.map((p) => (
              <div key={p} className="tp-login-prow">
                <button className="tp-login-pbtn" onClick={() => onLogin(p)}><Users size={15} /> {p}</button>
                <button className="tp-iconbtn tiny" title="削除" onClick={() => { if (window.confirm ? window.confirm(`「${p}」のデータを削除しますか？`) : true) onDelete(p); }}><Trash2 size={14} /></button>
              </div>
            ))}
          </div>
        )}

        <div className="tp-login-sub">{profiles.length > 0 ? "または新しい名前で" : "お名前"}</div>
        <div className="tp-login-form">
          <input value={name} onChange={(e) => setName(e.target.value)} onKeyDown={(e) => e.key === "Enter" && submit()} placeholder="例）山田太郎" autoFocus />
          <button className="tp-primarybtn" onClick={submit}>ログイン</button>
        </div>
        <p className="tp-login-note">通信は使いません。データはこの端末内にのみ保存されます（オフライン）。</p>
      </div>
    </div>
  );
}

/* ============================================================
   MeetingsView — 職員会議（資料PDF化・日付/カテゴリソート・書き込み）
   ============================================================ */
function MeetingsView({ data, setData, user, onPrintMeeting, onPrintMeetings, jump }) {
  const [sortMode, setSortMode] = useState("date"); // 'date' | 'cat'
  const [edit, setEdit] = useState(null); // meeting id being edited, or 'new'
  useEffect(() => { if (jump && jump.id) setEdit(jump.id); }, [jump && jump.at]);
  useEffect(() => { setTaskMsg(""); }, [edit]);
  const [annot, setAnnot] = useState(null); // imgId being annotated
  const [taskMsg, setTaskMsg] = useState("");
  const [proposal, setProposal] = useState(null); // {tasks:[], general:{}}
  const [query, setQuery] = useState("");
  const [summary, setSummary] = useState(null); // {from,to,cat}
  const [, setTick] = useState(0); const bump = () => setTick((t) => t + 1);
  const fileRef = useRef(null);
  const meetings = data.meetings || [];
  const setMeetings = (fn) => setData((d) => ({ ...d, meetings: fn(d.meetings || []) }));

  useEffect(() => { const ids = []; meetings.forEach((m) => (m.imgs || []).forEach((i) => ids.push(i))); if (ids.length) preloadMtgImgs(user, ids).then(bump); }, [meetings.map((m) => (m.imgs || []).join(",")).join("|"), user]);

  const openNew = () => { const id = uid(); setMeetings((l) => [...l, { id, date: ymd(new Date()), title: "", cat: "全体", notes: "", imgs: [] }]); setEdit(id); };
  const cur = meetings.find((m) => m.id === edit);
  const updCur = (patch) => setMeetings((l) => l.map((m) => m.id === edit ? { ...m, ...patch } : m));
  const removeMeeting = (id) => { const m = meetings.find((x) => x.id === id); (m?.imgs || []).forEach((imgId) => delMtgImg(user, imgId)); setMeetings((l) => l.filter((x) => x.id !== id)); setEdit(null); };

  const addImages = async (files) => {
    for (const f of Array.from(files)) {
      if (!/^image\//.test(f.type)) continue;
      try { const url = await downscaleImage(f); const imgId = uid(); await saveMtgImg(user, imgId, url); updCur({ imgs: [...(cur.imgs || []), imgId] }); bump(); } catch (e) {}
    }
    openProposal();
  };
  const openProposal = () => {
    const ctxYear = parseInt(String(data.meta.year || "").replace(/\D/g, ""), 10) || new Date().getFullYear();
    const tasks = extractMeetingTasks(cur, ctxYear);
    const general = { text: `${cur.title || "職員会議"}：資料の作成・提出`, due: cur.date || "", cat: cur.cat || "その他" };
    setProposal({ tasks, general });
  };
  const removeImage = async (imgId) => { await delMtgImg(user, imgId); updCur({ imgs: (cur.imgs || []).filter((x) => x !== imgId) }); bump(); };

  const sorted = meetings.slice().sort((a, b) => (b.date || "").localeCompare(a.date || ""));
  const groups = {};
  if (sortMode === "cat") MEETING_CATS.forEach((c) => { const items = sorted.filter((m) => (m.cat || "その他") === c.id); if (items.length) groups[c.id] = items; });

  const q = query.trim().toLowerCase();
  const hay = (m) => [m.title, m.notes, catMeta(m.cat).label, m.date, ...Object.values(m.caps || {})].join("\u0001").toLowerCase();
  const results = q ? sorted.filter((m) => hay(m).includes(q)) : null;
  const snippet = (text, needle) => {
    if (!text) return null; const low = text.toLowerCase(); const i = low.indexOf(needle);
    if (i < 0) return null; const s = Math.max(0, i - 24), e = Math.min(text.length, i + needle.length + 40);
    return { pre: (s > 0 ? "…" : "") + text.slice(s, i), hit: text.slice(i, i + needle.length), post: text.slice(i + needle.length, e) + (e < text.length ? "…" : "") };
  };
  const matchWhere = (m) => {
    if ((m.title || "").toLowerCase().includes(q)) return { label: "件名", sn: snippet(m.title, q) };
    if ((m.notes || "").toLowerCase().includes(q)) return { label: "書き込み", sn: snippet(m.notes, q) };
    const cap = Object.values(m.caps || {}).find((c) => (c || "").toLowerCase().includes(q));
    if (cap) return { label: "資料", sn: snippet(cap, q) };
    if (catMeta(m.cat).label.toLowerCase().includes(q)) return { label: "カテゴリ", sn: null };
    return { label: "日付", sn: null };
  };

  const MeetingRow = (m) => { const cm = catMeta(m.cat); const d = parseYmd(m.date); return (
    <div key={m.id} className="tp-mtg-item" style={{ borderLeftColor: cm.color }} onClick={() => setEdit(m.id)}>
      <div className="tp-mtg-main">
        <div className="tp-mtg-top"><span className="tp-mtg-date">{d.getFullYear()}/{d.getMonth() + 1}/{d.getDate()}</span><span className="tp-mtg-tag" style={{ background: cm.color }}>{cm.label}</span></div>
        <div className="tp-mtg-title">{m.title || "（無題）"}</div>
        {m.notes && <div className="tp-mtg-preview">{m.notes}</div>}
        {(m.imgs || []).length > 0 && <div className="tp-mtg-att"><ImageIcon size={12} /> 資料 {(m.imgs || []).length} 枚</div>}
      </div>
      <button className="tp-ghostbtn sm" onClick={(e) => { e.stopPropagation(); onPrintMeeting(m.id); }}><Printer size={13} /> PDF保存</button>
    </div>
  ); };

  const ResultRow = (m) => { const cm = catMeta(m.cat); const d = parseYmd(m.date); const w = matchWhere(m); return (
    <div key={m.id} className="tp-mtg-item" style={{ borderLeftColor: cm.color }} onClick={() => setEdit(m.id)}>
      <div className="tp-mtg-main">
        <div className="tp-mtg-top"><span className="tp-mtg-date">{d.getFullYear()}/{d.getMonth() + 1}/{d.getDate()}</span><span className="tp-mtg-tag" style={{ background: cm.color }}>{cm.label}</span><span className="tp-mtg-hitwhere">{w.label}に一致</span></div>
        <div className="tp-mtg-title">{m.title || "（無題）"}</div>
        {w.sn && <div className="tp-mtg-preview">{w.sn.pre}<mark className="tp-hl">{w.sn.hit}</mark>{w.sn.post}</div>}
      </div>
      <span className="tp-mtg-jump">開く ›</span>
    </div>
  ); };

  return (
    <div className="tp-view">
      <section className="tp-card">
        <div className="tp-card-headrow" style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, flexWrap: "wrap" }}>
          <h4 className="tp-card-title" style={{ margin: 0 }}><FileText size={16} /> 職員会議</h4>
          <div className="tp-field" style={{ margin: 0 }}><span style={{ fontSize: 11 }}>並べ替え</span>
            <Seg options={[{ v: "date", label: "日付順" }, { v: "cat", label: "カテゴリ順" }]} value={sortMode} onChange={setSortMode} />
          </div>
        </div>
        <p className="tp-hint" style={{ marginTop: 6 }}>会議ごとに資料（写真）と書き込みを保存。日付順／色タグのカテゴリ順に切替できます。各会議は「PDF保存」で書き出し・印刷できます。</p>
        <div className="tp-mtg-search">
          <Search size={15} />
          <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="この一覧を絞り込む（件名・書き込み・資料の説明）" />
          {query && <button className="tp-iconbtn tiny" onClick={() => setQuery("")}><X size={14} /></button>}
        </div>
        <div className="tp-toolbar" style={{ justifyContent: "flex-start", gap: 8 }}>
          <button className="tp-primarybtn" style={{ marginTop: 0 }} onClick={openNew}><Plus size={15} /> 会議を追加</button>
          <button className="tp-ghostbtn" disabled={meetings.length === 0} onClick={() => setSummary({ from: "", to: "", cat: "all", useSearch: !!q })}><FileText size={14} /> まとめてPDF</button>
        </div>

        {results ? (
          <>
            <div className="tp-mtg-grouphead">絞り込み結果 {results.length}件</div>
            {results.length === 0 ? <p className="tp-empty">一致する会議はありません。</p> : <div className="tp-mtg-list">{results.map(ResultRow)}</div>}
          </>
        ) : (<>
          {meetings.length === 0 && <p className="tp-empty">「会議を追加」から記録を始めましょう。</p>}
          {sortMode === "date" && <div className="tp-mtg-list">{sorted.map(MeetingRow)}</div>}
          {sortMode === "cat" && Object.keys(groups).map((cid) => { const cm = catMeta(cid); return (
            <div key={cid}>
              <div className="tp-mtg-grouphead"><span className="tp-dot2" style={{ background: cm.color }} />{cm.label}（{groups[cid].length}）</div>
              <div className="tp-mtg-list">{groups[cid].map(MeetingRow)}</div>
            </div>
          ); })}
        </>)}
      </section>

      <Modal open={!!cur} title={cur && cur.title ? cur.title : "会議メモ"} onClose={() => setEdit(null)} wide>
        {cur && (
          <>
            <div className="tp-field2" style={{ display: "flex", gap: 10 }}>
              <label className="tp-field" style={{ flex: 1 }}><span>日付</span><input type="date" value={cur.date} onChange={(e) => updCur({ date: e.target.value })} /></label>
            </div>
            <label className="tp-field"><span>件名</span><input value={cur.title} onChange={(e) => updCur({ title: e.target.value })} placeholder="例）11月 職員会議" /></label>
            <div className="tp-field"><span>カテゴリ（色タグ）</span>
              <div className="tp-catchips">{MEETING_CATS.map((c) => <button key={c.id} className={"tp-catchip" + (cur.cat === c.id ? " on" : "")} style={cur.cat === c.id ? { background: c.color } : {}} onClick={() => updCur({ cat: c.id })}>{c.label}</button>)}</div>
            </div>
            <label className="tp-field"><span>書き込み（メモ・決定事項）</span><textarea rows={6} value={cur.notes} onChange={(e) => updCur({ notes: e.target.value })} placeholder="議題・連絡・決定事項・自分のToDo など" /></label>
            <div className="tp-field"><span>資料（写真をアップロード → タップで書き込み）</span>
              <button className="tp-ghostbtn" onClick={() => fileRef.current?.click()}><Upload size={14} /> 画像を追加（複数可）</button>
              <input ref={fileRef} type="file" accept="image/*" multiple style={{ display: "none" }} onChange={(e) => { addImages(e.target.files); e.target.value = ""; }} />
              <div className="tp-mtg-thumbs">{(cur.imgs || []).map((imgId) => { const url = getMtgImg(user, imgId); return (
                <div key={imgId} className="tp-mtg-thumbcol">
                  <div className="tp-mtg-thumb">{url ? <img src={url} alt="" onClick={() => setAnnot(imgId)} /> : <span className="tp-mtg-thumb-load" onClick={() => setAnnot(imgId)} />}<span className="tp-mtg-thumb-pen"><Pencil size={11} /></span><button onClick={() => removeImage(imgId)}><X size={12} /></button></div>
                  <input className="tp-mtg-cap" value={(cur.caps || {})[imgId] || ""} onChange={(e) => updCur({ caps: { ...(cur.caps || {}), [imgId]: e.target.value } })} placeholder="資料の説明（検索対象）" />
                </div>
              ); })}</div>
              <p className="tp-hint" style={{ margin: "4px 0 0" }}>画像は端末内のデータベース(大容量)に保存され、上限に達しにくくなっています。説明を書くと検索できます。</p>
            </div>
            <div className="tp-field"><span>この会議からToDo（締切・作成/提出）</span>
              <button className="tp-ghostbtn" onClick={openProposal}><CheckSquare size={14} /> ToDoを提案・追加</button>
              <p className="tp-hint" style={{ margin: "4px 0 0" }}>書き込みから「締切/提出/○月○日」を読み取り、ポップアップで提案します。資料をアップした時にも自動で表示されます。</p>
            </div>
            <div className="tp-modal-actions">
              <button className="tp-dangerbtn" onClick={() => removeMeeting(cur.id)}><Trash2 size={14} /> この会議を削除</button>
              <div style={{ display: "flex", gap: 8 }}>
                <button className="tp-ghostbtn" onClick={() => onPrintMeeting(cur.id)}><Printer size={14} /> PDF保存</button>
                <button className="tp-primarybtn" onClick={() => setEdit(null)}><Check size={15} /> 完了</button>
              </div>
            </div>
          </>
        )}
      </Modal>
      {annot && <AnnotateModal user={user} imgId={annot} onClose={() => setAnnot(null)} onSaved={bump} />}
      {proposal && <MeetingTaskModal proposal={proposal} onClose={() => setProposal(null)} onAdd={(items) => {
        setData((d) => ({ ...d, todos: [...d.todos, ...items.map((t) => ({ id: uid(), text: t.text, cat: t.cat, due: t.due || "", done: false, date: ymd(new Date()), source: cur && cur.id }))] }));
        setProposal(null);
      }} />}
      <Modal open={!!summary} title="複数会議をまとめてPDF" onClose={() => setSummary(null)}>
        {summary && (() => {
          const inRange = (m) => (!summary.from || m.date >= summary.from) && (!summary.to || m.date <= summary.to);
          const inCat = (m) => summary.cat === "all" || (m.cat || "その他") === summary.cat;
          const base = summary.useSearch && q ? results : sorted;
          const picked = (base || sorted).filter((m) => inRange(m) && inCat(m));
          const ids = picked.slice().sort((a, b) => (a.date || "").localeCompare(b.date || "")).map((m) => m.id);
          const catLabel = summary.cat === "all" ? "全カテゴリ" : catMeta(summary.cat).label;
          return (<>
            {summary.useSearch && q && <p className="tp-hint" style={{ marginTop: 0 }}>検索「{query}」の結果からまとめます。</p>}
            <div className="tp-field2" style={{ display: "flex", gap: 10 }}>
              <label className="tp-field" style={{ flex: 1 }}><span>開始日</span><input type="date" value={summary.from} onChange={(e) => setSummary({ ...summary, from: e.target.value })} /></label>
              <label className="tp-field" style={{ flex: 1 }}><span>終了日</span><input type="date" value={summary.to} onChange={(e) => setSummary({ ...summary, to: e.target.value })} /></label>
            </div>
            <label className="tp-field"><span>カテゴリ</span>
              <select value={summary.cat} onChange={(e) => setSummary({ ...summary, cat: e.target.value })}>
                <option value="all">すべて</option>
                {MEETING_CATS.map((c) => <option key={c.id} value={c.id}>{c.label}</option>)}
              </select>
            </label>
            <p className="tp-hint">対象：<b style={{ color: "#2C7CA6" }}>{ids.length}件</b>（{catLabel}{summary.from || summary.to ? "・期間指定" : ""}）</p>
            <div className="tp-modal-actions">
              <button className="tp-ghostbtn" onClick={() => setSummary(null)}>キャンセル</button>
              <button className="tp-primarybtn" disabled={ids.length === 0} onClick={() => { const t = `職員会議まとめ（${catLabel}${summary.from ? "・" + summary.from : ""}${summary.to ? "〜" + summary.to : ""}）`; onPrintMeetings(ids, t); setSummary(null); }}><Printer size={15} /> {ids.length}件をPDF保存</button>
            </div>
          </>);
        })()}
      </Modal>
    </div>
  );
}

/* 資料アップ時などに出る、締切→ToDo提案ポップアップ */
function MeetingTaskModal({ proposal, onClose, onAdd }) {
  const found = proposal.tasks || [];
  const [rows, setRows] = useState(found.map((t) => ({ ...t, on: true })));
  const [genOn, setGenOn] = useState(false);
  const [genDue, setGenDue] = useState(proposal.general.due || "");
  const setRow = (i, patch) => setRows((r) => r.map((x, j) => j === i ? { ...x, ...patch } : x));
  const add = () => {
    const items = rows.filter((r) => r.on).map((r) => ({ text: r.text, due: r.due, cat: r.cat }));
    if (genOn) items.push({ text: proposal.general.text, due: genDue, cat: proposal.general.cat });
    if (items.length) onAdd(items); else onClose();
  };
  return (
    <div className="tp-modal-back" onClick={onClose}>
      <div className="tp-modal tp-taskprop" onClick={(e) => e.stopPropagation()}>
        <div className="tp-modal-head"><h3>ToDoの提案</h3><button className="tp-iconbtn" onClick={onClose}><X size={18} /></button></div>
        {found.length === 0 ? (
          <p className="tp-taskprop-none">締切の記載はありません。<br /><span>必要なら下の項目をToDoに追加できます。</span></p>
        ) : (
          <>
            <p className="tp-hint" style={{ marginTop: 0 }}>書き込みから次の締切を見つけました。追加するものを選んでください。</p>
            <div className="tp-taskprop-list">
              {rows.map((r, i) => (
                <div key={i} className="tp-taskprop-row">
                  <button className={"tp-donebox" + (r.on ? " on" : "")} onClick={() => setRow(i, { on: !r.on })}>{r.on && <Check size={12} />}</button>
                  <input className="tp-taskprop-text" value={r.text} onChange={(e) => setRow(i, { text: e.target.value })} />
                  <input type="date" value={r.due} onChange={(e) => setRow(i, { due: e.target.value })} />
                </div>
              ))}
            </div>
          </>
        )}
        <div className="tp-taskprop-gen">
          <button className={"tp-donebox" + (genOn ? " on" : "")} onClick={() => setGenOn((v) => !v)}>{genOn && <Check size={12} />}</button>
          <span className="tp-taskprop-gentext">「{proposal.general.text}」も追加</span>
          <input type="date" value={genDue} onChange={(e) => setGenDue(e.target.value)} />
        </div>
        <div className="tp-modal-actions">
          <button className="tp-ghostbtn" onClick={onClose}>閉じる</button>
          <button className="tp-primarybtn" onClick={add}><Plus size={15} /> ToDoに追加</button>
        </div>
      </div>
    </div>
  );
}

/* 資料への直接書き込み（画像の上にペンで描いて保存／端末に保存） */
function AnnotateModal({ user, imgId, onClose, onSaved }) {
  const canvasRef = useRef(null);
  const wrapRef = useRef(null);
  const drawing = useRef(false);
  const undoStack = useRef([]);
  const [color, setColor] = useState("#D9534F");
  const [width, setWidth] = useState(4);

  useEffect(() => { (async () => {
    await preloadMtgImgs(user, [imgId]);
    const url = getMtgImg(user, imgId);
    const cv = canvasRef.current; if (!cv) return;
    const ctx = cv.getContext("2d");
    const draw = (img) => {
      const maxW = Math.min(900, (wrapRef.current?.clientWidth || 900));
      let w = img ? img.width : 900, h = img ? img.height : 600;
      if (w > maxW) { const s = maxW / w; w = Math.round(w * s); h = Math.round(h * s); }
      cv.width = w; cv.height = h;
      ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, w, h);
      if (img) ctx.drawImage(img, 0, 0, w, h);
    };
    if (url) { const img = new Image(); img.onload = () => draw(img); img.onerror = () => draw(null); img.src = url; }
    else draw(null);
  })(); }, [user, imgId]);

  const pos = (e) => { const cv = canvasRef.current; const r = cv.getBoundingClientRect(); const t = e.touches ? e.touches[0] : e; return { x: (t.clientX - r.left) * (cv.width / r.width), y: (t.clientY - r.top) * (cv.height / r.height) }; };
  const start = (e) => { e.preventDefault(); const cv = canvasRef.current; try { undoStack.current.push(cv.toDataURL("image/png")); if (undoStack.current.length > 12) undoStack.current.shift(); } catch (_) {} drawing.current = true; const ctx = cv.getContext("2d"); const p = pos(e); ctx.beginPath(); ctx.moveTo(p.x, p.y); ctx.lineCap = "round"; ctx.lineJoin = "round"; ctx.strokeStyle = color; ctx.lineWidth = width; };
  const move = (e) => { if (!drawing.current) return; e.preventDefault(); const ctx = canvasRef.current.getContext("2d"); const p = pos(e); ctx.lineTo(p.x, p.y); ctx.stroke(); };
  const end = () => { drawing.current = false; };
  const undo = () => { const prev = undoStack.current.pop(); if (!prev) return; const cv = canvasRef.current; const ctx = cv.getContext("2d"); const img = new Image(); img.onload = () => { ctx.clearRect(0, 0, cv.width, cv.height); ctx.drawImage(img, 0, 0); }; img.src = prev; };
  const save = async () => { try { const url = canvasRef.current.toDataURL("image/jpeg", 0.82); await saveMtgImg(user, imgId, url); onSaved && onSaved(); onClose(); } catch (e) { onClose(); } };
  const download = () => { try { const url = canvasRef.current.toDataURL("image/jpeg", 0.9); const a = document.createElement("a"); a.href = url; a.download = `資料_${ymd(new Date())}.jpg`; document.body.appendChild(a); a.click(); a.remove(); } catch (e) {} };

  return (
    <div className="tp-modal-back">
      <div className="tp-modal tp-annot" onClick={(e) => e.stopPropagation()}>
        <div className="tp-modal-head"><h3>資料に書き込み</h3><button className="tp-iconbtn" onClick={onClose}><X size={18} /></button></div>
        <div className="tp-annot-tools">
          {["#D9534F", "#2C7CA6", "#111", "#E0A64B", "#4F9E86"].map((c) => <button key={c} className={"tp-annot-color" + (color === c ? " on" : "")} style={{ background: c }} onClick={() => setColor(c)} />)}
          <span className="tp-annot-sep" />
          {[2, 4, 8].map((w) => <button key={w} className={"tp-annot-w" + (width === w ? " on" : "")} onClick={() => setWidth(w)}><span style={{ width: w + 2, height: w + 2 }} /></button>)}
          <span className="tp-annot-sep" />
          <button className="tp-ghostbtn sm" onClick={undo}><RotateCcw size={13} /> 1つ戻す</button>
        </div>
        <div className="tp-annot-canvaswrap" ref={wrapRef}>
          <canvas ref={canvasRef} className="tp-annot-canvas"
            onMouseDown={start} onMouseMove={move} onMouseUp={end} onMouseLeave={end}
            onTouchStart={start} onTouchMove={move} onTouchEnd={end} />
        </div>
        <div className="tp-modal-actions">
          <button className="tp-ghostbtn" onClick={download}><Download size={14} /> 端末に保存</button>
          <button className="tp-primarybtn" onClick={save}><Check size={15} /> 保存</button>
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   ErrorBoundary — 画面が壊れても全体を落とさない
   ============================================================ */
class ErrorBoundary extends React.Component {
  constructor(props) { super(props); this.state = { err: null }; }
  static getDerivedStateFromError(err) { return { err }; }
  componentDidUpdate(prev) { if (prev.tabKey !== this.props.tabKey && this.state.err) this.setState({ err: null }); }
  render() {
    if (this.state.err) {
      return (
        <div className="tp-view"><section className="tp-card">
          <h4 className="tp-card-title">この画面でエラーが発生しました</h4>
          <p className="tp-hint">他のタブは通常どおり使えます。データは保存されています。設定＞自動バックアップから復元もできます。</p>
          <button className="tp-primarybtn" onClick={() => this.setState({ err: null })}>もう一度表示</button>
        </section></div>
      );
    }
    return this.props.children;
  }
}

/* ============================================================
/* ============================================================
   Onboarding — 初回セットアップ
   ============================================================ */
function OnboardingModal({ data, setData, openImport }) {
  const [step, setStep] = useState(1);
  const setMeta = (patch) => setData((d) => ({ ...d, meta: { ...d.meta, ...patch } }));
  const finish = (mode) => {
    setData((d) => {
      let nd = { ...d, meta: { ...d.meta, onboarded: true } };
      if (mode === "empty") nd = { ...nd, timetable: {}, lessonLog: {}, events: [], todos: [], dayMemo: {}, weeklyManual: {}, club: { ...d.club, days: {}, specials: [], schedule: {} }, textbook: { ...(d.textbook || {}), units: [] }, tests: [], duties: [], routine: [], rosters: {}, subjects: [], classes: [], homeroom: "" };
      return nd;
    });
  };
  return (
    <div className="tp-modal-back">
      <div className="tp-modal tp-onboard">
        <div className="tp-onboard-head"><Logo size={40} /><div><div className="tp-appname" style={{ fontSize: 20 }}>ようこそ、{data.meta.teacher} 先生</div><div className="tp-apptag">かんたん初期設定（{step}/2）</div></div></div>
        {step === 1 && (
          <>
            <label className="tp-field"><span>年度</span><input value={data.meta.year} onChange={(e) => setMeta({ year: e.target.value })} placeholder="例）2026年度" /></label>
            <label className="tp-field"><span>部活動名</span><input value={data.club.name} onChange={(e) => setData((d) => ({ ...d, club: { ...d.club, name: e.target.value } }))} placeholder="例）ソフトテニス部" /></label>
            <div className="tp-field"><span>文字サイズ</span><Seg options={[{ v: "S", label: "小" }, { v: "M", label: "標準" }, { v: "L", label: "大" }]} value={data.meta.fontScale || "M"} onChange={(v) => setMeta({ fontScale: v })} /></div>
            <div className="tp-modal-actions"><span /><button className="tp-primarybtn" onClick={() => setStep(2)}>次へ</button></div>
          </>
        )}
        {step === 2 && (
          <>
            <p className="tp-hint" style={{ marginTop: 0 }}>始め方を選んでください（あとで設定＞初期化・取り込みから変更できます）。</p>
            <div className="tp-onboard-choices">
              <button className="tp-onboard-choice" onClick={() => finish("keep")}><b>サンプル入りで始める</b><span>使い方を見ながら試したい方へ。例の授業・部活が入ります。</span></button>
              <button className="tp-onboard-choice" onClick={() => finish("empty")}><b>まっさらで始める</b><span>自分の予定だけを入れたい方へ。空の状態から始めます。</span></button>
              <button className="tp-onboard-choice" onClick={openImport}><b>取り込みで始める</b><span>Excel/CSV・JSONから予定や名簿を読み込みます。</span></button>
            </div>
            <div className="tp-modal-actions"><button className="tp-ghostbtn" onClick={() => setStep(1)}>戻る</button><span /></div>
          </>
        )}
      </div>
    </div>
  );
}

function todoVEvent(todo) {
  const d = (todo.due || "").replace(/-/g, ""); if (!d) return null;
  const stamp = new Date().toISOString().replace(/[-:]/g, "").replace(/\.\d+Z$/, "Z");
  return ["BEGIN:VEVENT", `UID:tp-todo-${d}-${todo.id}@teaching-planners`, `DTSTAMP:${stamp}`,
    `DTSTART:${d}T080000`, `DTEND:${d}T083000`,
    `SUMMARY:${icsEsc("【ToDo】" + (todo.text || ""))}`,
    "BEGIN:VALARM", "ACTION:DISPLAY", "DESCRIPTION:リマインド", "TRIGGER:-PT0M", "END:VALARM",
    "BEGIN:VALARM", "ACTION:DISPLAY", "DESCRIPTION:前日リマインド", "TRIGGER:-P1D", "END:VALARM",
    "END:VEVENT"].join("\r\n");
}
function wrapICS(events) { return ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Teaching Partner for YOU//JP//", "CALSCALE:GREGORIAN", ...events, "END:VCALENDAR"].join("\r\n"); }
function buildTodoICS(todo) { const e = todoVEvent(todo); return e ? wrapICS([e]) : null; }
function downloadICS(text, name) {
  try { const blob = new Blob([text], { type: "text/calendar" }); const url = URL.createObjectURL(blob); const a = document.createElement("a"); a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(url), 1000); } catch (e) {}
}
function downloadTodoICS(todo) { const ics = buildTodoICS(todo); if (ics) downloadICS(ics, `ToDo_${(todo.text || "").slice(0, 12)}_${todo.due}.ics`); }
function downloadTodosICS(todos) { const events = todos.map(todoVEvent).filter(Boolean); if (!events.length) return; downloadICS(wrapICS(events), `ToDo締切_${ymd(new Date())}.ics`); }

/* ============================================================
   TodoDock — カテゴリ・締切つきToDo（全画面/サイドバー/最小化）
   ============================================================ */
function extractMeetingTasks(m, ctxYear) {
  const notes = m.notes || ""; const found = [];
  const base = /^\d{4}-\d{2}-\d{2}$/.test(m.date || "") ? parseYmd(m.date) : new Date();
  const KW = /(締切|〆切|締め切り|しめ切り|提出|期限|必着|期日|納期|完成|まで(?:に)?)/;
  notes.split(/\r?\n/).forEach((raw) => {
    const line = raw.trim(); if (!line) return;
    if (!KW.test(line)) return;
    const due = parseDueInText(line, ctxYear, base);
    found.push({ text: line.replace(/^[・\-*\s]+/, "").slice(0, 70), due, cat: m.cat || "その他" });
  });
  return found;
}
// 多様な締切表現を日付(YYYY-MM-DD)に。ctxYear=年度開始年、base=基準日(相対・日のみ用)
function weekdayDate(base, dow, scope) {
  const mon = startOfWeekMon(base); const off = (dow === 0 ? 6 : dow - 1);
  if (scope === "thisweek") return addDays(mon, off);
  if (scope === "next") return addDays(mon, 7 + off);
  if (scope === "after2") return addDays(mon, 14 + off);
  const from = scope === "nextocc" ? addDays(base, 1) : base; // coming: 当日以降 / nextocc: 翌日以降
  let d = new Date(from); for (let i = 0; i < 8; i++) { if (d.getDay() === dow) return new Date(d); d = addDays(d, 1); }
  return addDays(mon, off);
}
function parseDueInText(text, ctxYear, base) {
  const s = zen2han(text); const V = (mo, dy) => mo >= 1 && mo <= 12 && dy >= 1 && dy <= 31;
  const y3 = (y, mo, dy) => `${y}-${pad(mo)}-${pad(dy)}`; let m;
  m = s.match(/(?:令和|R)\s*(\d{1,2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日?/); if (m && V(+m[2], +m[3])) return y3(2018 + (+m[1]), +m[2], +m[3]);
  m = s.match(/(20\d{2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日?/); if (m && V(+m[2], +m[3])) return y3(+m[1], +m[2], +m[3]);
  m = s.match(/(20\d{2})\s*[\/\.\-]\s*(\d{1,2})\s*[\/\.\-]\s*(\d{1,2})/); if (m && V(+m[2], +m[3])) return y3(+m[1], +m[2], +m[3]);
  m = s.match(/(\d{1,2})\s*月\s*(\d{1,2})\s*日?/); if (m && V(+m[1], +m[2])) { const mo = +m[1]; return y3(mo >= 4 ? ctxYear : ctxYear + 1, mo, +m[2]); }
  m = s.match(/(\d{1,2})\s*[\/\.]\s*(\d{1,2})(?!\d)/); if (m && V(+m[1], +m[2])) { const mo = +m[1]; return y3(mo >= 4 ? ctxYear : ctxYear + 1, mo, +m[2]); }
  if (base) {
    let mm;
    mm = s.match(/来月\s*(\d{1,2})\s*日/); if (mm) { const dy = +mm[1]; if (dy >= 1 && dy <= 31) return ymd(new Date(base.getFullYear(), base.getMonth() + 1, dy)); }
    mm = s.match(/今月\s*(\d{1,2})\s*日/); if (mm) { const dy = +mm[1]; if (dy >= 1 && dy <= 31) return ymd(new Date(base.getFullYear(), base.getMonth(), dy)); }
    if (/来月末/.test(s)) return ymd(new Date(base.getFullYear(), base.getMonth() + 2, 0));
    if (/(今月末|月末|末日)/.test(s)) return ymd(new Date(base.getFullYear(), base.getMonth() + 1, 0));
    const wm = s.match(/([日月火水木金土])\s*曜(?:日)?/);
    if (wm) {
      const dow = "日月火水木金土".indexOf(wm[1]);
      let scope = "coming";
      if (/再来週/.test(s)) scope = "after2";
      else if (/来週/.test(s)) scope = "next";
      else if (/今週|この週|こんしゅう/.test(s)) scope = "thisweek";
      else if (/次の?|翌/.test(s)) scope = "nextocc";
      return ymd(weekdayDate(base, dow, scope));
    }
    if (/来週末/.test(s)) return ymd(addDays(startOfWeekMon(addDays(base, 7)), 4));
    if (/来週/.test(s)) return ymd(addDays(startOfWeekMon(base), 7));
    if (/今週末|週末/.test(s)) return ymd(addDays(startOfWeekMon(base), 4));
    if (/週明け/.test(s)) return ymd(addDays(startOfWeekMon(base), 7));
    if (/明後日/.test(s)) return ymd(addDays(base, 2));
    if (/明日/.test(s)) return ymd(addDays(base, 1));
    mm = s.match(/(\d{1,2})\s*日/); if (mm) { const dy = +mm[1]; if (dy >= 1 && dy <= 31) return y3(base.getFullYear(), base.getMonth() + 1, dy); }
  }
  return "";
}

function TodoDock({ data, setData, mode, setMode }) {
  const [text, setText] = useState("");
  const [cat, setCat] = useState("その他");
  const [due, setDue] = useState("");
  const [filter, setFilter] = useState("all");
  const todos = data.todos || [];
  const openCount = todos.filter((t) => !t.done).length;
  const today = ymd(new Date());
  const add = () => { if (!text.trim()) return; setData((d) => ({ ...d, todos: [...d.todos, { id: uid(), text: text.trim(), cat, due, done: false, date: today }] })); setText(""); setDue(""); };
  const toggle = (id) => setData((d) => ({ ...d, todos: d.todos.map((t) => t.id === id ? { ...t, done: !t.done } : t) }));
  const del = (id) => setData((d) => ({ ...d, todos: d.todos.filter((t) => t.id !== id) }));
  const shown = todos.filter((t) => filter === "all" || (t.cat || "その他") === filter).slice().sort((a, b) => (a.done - b.done) || ((a.due || "9999") < (b.due || "9999") ? -1 : 1));

  if (mode === "min") {
    return (
      <button className="tp-todo-min" onClick={() => setMode("side")} aria-label="ToDoを開く">
        <CheckSquare size={16} /> ToDo{openCount > 0 && <span className="tp-todo-badge">{openCount}</span>}
      </button>
    );
  }
  const Body = (
    <>
      <div className="tp-todo-dockhead">
        <span className="tp-todo-docktitle"><CheckSquare size={16} /> ToDo<span className="tp-cardsub">計画的な大きいタスク・締切あり</span></span>
        <div className="tp-todo-modes">
          <button className={"tp-iconbtn tiny" + (mode === "full" ? " on" : "")} title="全画面" onClick={() => setMode("full")}><Maximize2 size={15} /></button>
          <button className={"tp-iconbtn tiny" + (mode === "side" ? " on" : "")} title="サイドバー" onClick={() => setMode("side")}><Columns size={15} /></button>
          <button className="tp-iconbtn tiny" title="最小化" onClick={() => setMode("min")}><Minus size={15} /></button>
        </div>
      </div>
      <div className="tp-todo-add2">
        <input value={text} onChange={(e) => setText(e.target.value)} placeholder="やること・締切のあるタスク" onKeyDown={(e) => e.key === "Enter" && add()} />
        <input type="date" value={due} onChange={(e) => setDue(e.target.value)} title="締切" />
        <button className="tp-addbtn" onClick={add}><Plus size={16} /></button>
      </div>
      <div className="tp-todo-catrow">
        <button className={"tp-catchip mini" + (cat === "その他" && false ? "" : "")} style={{ display: "none" }} />
        {MEETING_CATS.map((c) => <button key={c.id} className={"tp-catchip mini" + (cat === c.id ? " on" : "")} style={cat === c.id ? { background: c.color } : {}} onClick={() => setCat(c.id)}>{c.label}</button>)}
      </div>
      <div className="tp-todo-filter">
        <button className={"tp-fchip" + (filter === "all" ? " on" : "")} onClick={() => setFilter("all")}>すべて</button>
        {MEETING_CATS.map((c) => <button key={c.id} className={"tp-fchip" + (filter === c.id ? " on" : "")} onClick={() => setFilter(c.id)}><span className="tp-dot2" style={{ background: c.color }} />{c.label}</button>)}
      </div>
      <div className="tp-todo-hint">🔔＝端末のカレンダー/通知センターに登録（締切のあるタスク）。
        {todos.filter((t) => t.due && !t.done).length > 0 && <button className="tp-todo-bulkics" onClick={() => downloadTodosICS(todos.filter((t) => t.due && !t.done))}><Bell size={12} /> 締切をまとめて登録（{todos.filter((t) => t.due && !t.done).length}）</button>}
      </div>
      <div className="tp-todo-docklist">
        {shown.length === 0 && <p className="tp-empty">タスクはありません。</p>}
        {shown.map((t) => { const cm = catMeta(t.cat); const overdue = t.due && !t.done && t.due < today; return (
          <div key={t.id} className={"tp-todo-drow" + (t.done ? " done" : "")}>
            <button className={"tp-donebox" + (t.done ? " on" : "")} onClick={() => toggle(t.id)}>{t.done && <Check size={12} />}</button>
            <span className="tp-todo-cat" style={{ background: cm.color }} title={cm.label} />
            <span className="tp-todo-dtext">{t.text}</span>
            {t.due && <span className={"tp-todo-due" + (overdue ? " over" : "")}>{t.due.slice(5).replace("-", "/")}{overdue ? " !" : ""}</span>}
            {t.due && <button className="tp-iconbtn tiny" title="端末のカレンダー/通知に追加" onClick={() => downloadTodoICS(t)}><Bell size={13} /></button>}
            <button className="tp-iconbtn tiny" onClick={() => del(t.id)}><Trash2 size={13} /></button>
          </div>
        ); })}
      </div>
    </>
  );
  if (mode === "full") return (<div className="tp-modal-back" onClick={() => setMode("min")}><div className="tp-todo-full" onClick={(e) => e.stopPropagation()}>{Body}</div></div>);
  return <div className="tp-todo-side">{Body}</div>;
}

/* ============================================================
   GlobalSearch — アプリ全体の横断検索
   ============================================================ */
function runGlobalSearch(data, qRaw) {
  const q = qRaw.trim().toLowerCase(); if (!q) return [];
  const has = (s) => String(s == null ? "" : s).toLowerCase().includes(q);
  const R = [];
  (data.events || []).forEach((e) => { if (has(e.title)) R.push({ type: "行事・予定", label: e.title, sub: e.date + (e.time ? " " + e.time : ""), date: e.date, goto: "today" }); });
  Object.entries(data.lessonLog || {}).forEach(([k, v]) => { if (has(v.topic) || has(v.hw)) { const date = k.slice(0, 10); R.push({ type: "授業", label: v.topic || v.hw, sub: `${date} ${v.subject || ""} ${v.klass || ""}`.trim(), date, goto: "today" }); } });
  Object.entries((data.club && data.club.days) || {}).forEach(([date, v]) => { if (has(v.content) || has(v.matchName) || has(v.note) || has(v.place)) R.push({ type: "部活", label: v.matchName || v.content || "部活", sub: `${date}${v.place ? " ・" + v.place : ""}`, date, goto: "today" }); });
  Object.entries(data.rosters || {}).forEach(([klass, list]) => { (list || []).forEach((s) => { if (has(s.name) || has(s.kana) || has(s.memo)) R.push({ type: "名簿", label: s.name || "（未入力）", sub: `${klass}${s.memo ? " ・" + s.memo : ""}`, goto: "roster", klass }); }); });
  (data.meetings || []).forEach((m) => { const capHit = Object.values(m.caps || {}).some(has); if (has(m.title) || has(m.notes) || capHit || has(catMeta(m.cat).label)) R.push({ type: "職員会議", label: m.title || "（無題）", sub: `${m.date} ・${catMeta(m.cat).label}`, goto: "meetings", meetingId: m.id }); });
  (data.tests || []).forEach((t) => { if (has(t.name)) R.push({ type: "テスト", label: t.name, sub: `${t.date || ""} ${t.type || ""}`.trim(), goto: "classes" }); });
  (data.todos || []).forEach((t) => { if (has(t.text) || has(catMeta(t.cat).label)) R.push({ type: "やること", label: t.text, sub: `${t.due ? "〆" + t.due.slice(5).replace("-", "/") + " " : ""}${catMeta(t.cat).label}`, date: t.due || t.date, goto: "today" }); });
  Object.entries(data.dayMemo || {}).forEach(([date, txt]) => { if (has(txt)) R.push({ type: "メモ", label: String(txt).slice(0, 50), sub: date, date, goto: "today" }); });
  return R;
}
function GlobalSearch({ open, data, onClose, onJump }) {
  const [q, setQ] = useState("");
  const [tf, setTf] = useState("all");
  const inputRef = useRef(null);
  useEffect(() => { if (open) { setQ(""); setTf("all"); setTimeout(() => { try { inputRef.current && inputRef.current.focus(); } catch (e) {} }, 50); } }, [open]);
  if (!open) return null;
  const all = runGlobalSearch(data, q);
  const order = ["行事・予定", "授業", "部活", "名簿", "職員会議", "テスト", "やること", "メモ"];
  const counts = {}; all.forEach((r) => { counts[r.type] = (counts[r.type] || 0) + 1; });
  const results = tf === "all" ? all : all.filter((r) => r.type === tf);
  const groups = {}; results.forEach((r) => { (groups[r.type] = groups[r.type] || []).push(r); });
  const hl = (text) => {
    const low = String(text).toLowerCase(); const n = q.trim().toLowerCase(); const i = n ? low.indexOf(n) : -1;
    if (i < 0) return text;
    return (<>{String(text).slice(0, i)}<mark className="tp-hl">{String(text).slice(i, i + n.length)}</mark>{String(text).slice(i + n.length)}</>);
  };
  return (
    <div className="tp-modal-back" onClick={onClose}>
      <div className="tp-modal tp-gsearch" onClick={(e) => e.stopPropagation()}>
        <div className="tp-gsearch-bar">
          <Search size={18} />
          <input ref={inputRef} value={q} onChange={(e) => setQ(e.target.value)} placeholder="予定・授業・部活・名簿・会議・メモを検索" />
          <button className="tp-iconbtn" onClick={onClose}><X size={18} /></button>
        </div>
        <div className="tp-gsearch-body">
          {q.trim() && all.length > 0 && (
            <div className="tp-gsearch-tf">
              <button className={"tp-fchip" + (tf === "all" ? " on" : "")} onClick={() => setTf("all")}>すべて（{all.length}）</button>
              {order.filter((t) => counts[t]).map((t) => <button key={t} className={"tp-fchip" + (tf === t ? " on" : "")} onClick={() => setTf(t)}>{t}（{counts[t]}）</button>)}
            </div>
          )}
          {!q.trim() && <p className="tp-empty">キーワードを入力してください。すべてのメニューを横断して探します。</p>}
          {q.trim() && all.length === 0 && <p className="tp-empty">一致する項目はありません。</p>}
          {order.filter((t) => groups[t]).map((t) => (
            <div key={t} className="tp-gsearch-group">
              <div className="tp-gsearch-gh">{t}（{groups[t].length}）</div>
              {groups[t].slice(0, 20).map((r, i) => (
                <button key={i} className="tp-gsearch-row" onClick={() => onJump(r)}>
                  <span className="tp-gsearch-label">{hl(r.label)}</span>
                  <span className="tp-gsearch-sub">{r.sub}</span>
                  <span className="tp-gsearch-go">›</span>
                </button>
              ))}
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   CalendarView — 週/月/3か月/6か月/年 をまとめた1タブ
   ============================================================ */
function CalendarView({ data, setData, selDate, setSelDate, vis, toggleVis, onPrint, onPrintCal, onPrintWeekplan, goToday, period, setPeriod }) {
  const PERIODS = [{ v: "week", label: "週" }, { v: "month", label: "月" }, { v: "q3", label: "3か月" }, { v: "q6", label: "6か月" }, { v: "year", label: "年" }];
  return (
    <div className="tp-view tp-calview">
      <div className="tp-period-switch">
        <Seg options={PERIODS} value={period} onChange={setPeriod} />
      </div>
      {period === "week" && <WeekView data={data} setData={setData} selDate={selDate} setSelDate={setSelDate} vis={vis} toggleVis={toggleVis} onPrint={onPrint} onPrintWeekplan={onPrintWeekplan} embed />}
      {period === "month" && <MonthView data={data} setData={setData} selDate={selDate} setSelDate={goToday} vis={vis} toggleVis={toggleVis} onPrint={onPrint} embed />}
      {period === "q3" && <QuarterView data={data} selDate={selDate} setSelDate={goToday} vis={vis} toggleVis={toggleVis} onPrint={onPrint} onPrintCal={onPrintCal} count={3} embed />}
      {period === "q6" && <QuarterView data={data} selDate={selDate} setSelDate={goToday} vis={vis} toggleVis={toggleVis} onPrint={onPrint} onPrintCal={onPrintCal} count={6} embed />}
      {period === "year" && <YearView data={data} setData={setData} setSelDate={setSelDate} setTab={() => setPeriod("month")} vis={vis} toggleVis={toggleVis} onPrint={onPrint} embed />}
    </div>
  );
}

/* ============================================================
   APP
   ============================================================ */
export default function App() {
  const [user, setUser] = useState(null);      // ログイン中の名前
  const [profiles, setProfiles] = useState([]); // 保存済みの名前一覧
  const [booting, setBooting] = useState(true);
  const [data, setData] = useState(null);
  const [tab, setTab] = useState("today");
  const [calPeriod, setCalPeriod] = useState("month");
  const [selDate, setSelDate] = useState(new Date());
  const [settings, setSettings] = useState(false);
  const [searchOpen, setSearchOpen] = useState(false);
  const [todoMode, setTodoMode] = useState("min");
  const [rosterJump, setRosterJump] = useState(null); // {klass, at}
  const [meetingJump, setMeetingJump] = useState(null); // {id, at}
  const [calVis, setCalVis] = useState(defaultVis()); // カテゴリ表示ON/OFF（週/月/3か月/年で共有）
  const [printData, setPrintData] = useState(null); // {title, rows, vis, auto}
  const loaded = useRef(false);
  const backupRef = useRef(0);
  const [backupReminder, setBackupReminder] = useState(false);
  const [quotaWarn, setQuotaWarn] = useState(false);
  const [toast, setToast] = useState(null); // {msg, prev, at}
  const toggleVis = (id) => setCalVis((v) => ({ ...v, [id]: !v[id] }));
  const onPrint = (title, rows, vis, auto) => setPrintData({ title, rows, vis, auto, mode: "list" });
  const onPrintCal = (title, months, vis, auto) => setPrintData({ title, months, vis, auto, mode: "calendar" });

  const pushBackup = async (name, snap) => {
    if (!name || !snap) return;
    try {
      const ring = (await loadStore(`${NS}:backups:${name}`)) || [];
      ring.unshift({ at: new Date().toISOString(), data: snap });
      await saveStore(`${NS}:backups:${name}`, ring.slice(0, 5));
    } catch (e) {}
  };

  // 起動：前回ログインしていれば自動で引き継ぎ
  useEffect(() => { (async () => {
    try { window.__tpOnQuota = () => setQuotaWarn(true); } catch (e) {}
    const idx = (await loadStore(`${NS}:index`)) || [];
    setProfiles(idx);
    const last = await loadStore(`${NS}:last`);
    if (last && typeof last === "string") {
      let d = await loadStore(dataKeyOf(last));
      if (!d) { const ring = (await loadStore(`${NS}:backups:${last}`)) || []; if (ring[0]) d = ring[0].data; } // 破損時はバックアップから復旧
      if (d) {
        d = migrate(d);
        setUser(last); setData(d); loaded.current = true; backupRef.current = Date.now();
        pushBackup(last, d);
        const le = await loadStore(`${NS}:lastExport:${last}`);
        const stale = !le || (Date.now() - new Date(le).getTime() > 7 * 864e5);
        if (stale) setBackupReminder(true);
      }
    }
    setBooting(false);
  })(); }, []);

  // データ変更を、ログイン中の名前ごとに自動保存（＋定期バックアップ）
  useEffect(() => {
    if (loaded.current && user && data) {
      saveStore(dataKeyOf(user), data);
      const now = Date.now();
      if (now - backupRef.current > 600000) { backupRef.current = now; pushBackup(user, data); }
    }
  }, [data, user]);

  useEffect(() => { if (!toast) return; const t = setTimeout(() => setToast(null), 6000); return () => clearTimeout(t); }, [toast]);

  const doLogin = async (rawName) => {
    const name = (rawName || "").trim();
    if (!name) return;
    loaded.current = false;
    let d = await loadStore(dataKeyOf(name));
    if (!d) {
      // 新規プロフィール。旧・単一データ or 現在のデータがあれば初回のみ引き継ぐ
      const legacy = (profiles.length === 0) ? await loadStore("teacher-techo-v1") : null;
      d = legacy || defaultData();
      d.meta = { ...d.meta, teacher: name, onboarded: false };
    } else {
      d.meta = { ...d.meta, teacher: d.meta?.teacher || name };
    }
    const idx = (await loadStore(`${NS}:index`)) || [];
    if (!idx.includes(name)) { idx.push(name); await saveStore(`${NS}:index`, idx); }
    d = migrate(d);
    await saveStore(dataKeyOf(name), d);
    await saveStore(`${NS}:last`, name);
    setProfiles(idx.includes(name) ? idx : [...idx, name]);
    setUser(name); setData(d); setTab("today");
    loaded.current = true; backupRef.current = Date.now();
    pushBackup(name, d);
    const le = await loadStore(`${NS}:lastExport:${name}`);
    setBackupReminder(!le || (Date.now() - new Date(le).getTime() > 7 * 864e5));
  };

  const doLogout = async () => { await delStore(`${NS}:last`); setUser(null); setData(null); setSettings(false); loaded.current = false; const idx = (await loadStore(`${NS}:index`)) || []; setProfiles(idx); };

  const deleteProfile = async (name) => { await delStore(dataKeyOf(name)); const idx = ((await loadStore(`${NS}:index`)) || []).filter((n) => n !== name); await saveStore(`${NS}:index`, idx); setProfiles(idx); };

  if (booting) return (<div className="tp-app"><style>{CSS}</style><div className="tp-loading">読み込み中…</div></div>);
  if (!user || !data) return (<div className="tp-app"><style>{CSS}</style><LoginScreen profiles={profiles} onLogin={doLogin} onDelete={deleteProfile} /></div>);

  const appCls = "tp-app fs-" + (data.meta.fontScale || "M").toLowerCase() + (data.meta.contrast ? " contrast" : "") + (data.meta.theme === "dark" ? " dark" : "");
  const showToast = (msg, prev) => { setToast({ msg, prev, at: Date.now() }); };

  return (
    <div className={appCls}>
      <style>{CSS}</style>
      <TabBar active={tab} onChange={setTab} onSettings={() => setSettings(true)} />
      <main className="tp-main">
        <header className="tp-topbar">
          <div className="tp-appbrand">
            <span className="tp-appbrand-logo"><Logo size={34} /></span>
            <div className="tp-appwordmark">
              <span className="tp-appname">Teaching Partner for YOU</span>
              <span className="tp-apptag">for teachers</span>
            </div>
          </div>
          <div className="tp-topbar-right">
            <div className="tp-section">
              <span className="tp-section-name">{TABS.find((t) => t.id === tab)?.label}</span>
              {(data.meta.year || data.meta.teacher) && <span className="tp-topbar-sub">{data.meta.year}{data.meta.teacher && `・${data.meta.teacher}`}</span>}
            </div>
            <button className="tp-set-mobile tp-searchbtn" onClick={() => setSearchOpen(true)} aria-label="全体検索"><Search size={18} /></button>
            <button className="tp-set-mobile" onClick={() => setSettings(true)}><Settings size={18} /></button>
          </div>
        </header>
        {quotaWarn && (
          <div className="tp-backup-banner" style={{ background: "#FDE8E8", borderColor: "#F3C9C9", color: "#9A2323" }}>
            <ShieldAlert size={16} />
            <span>端末の保存容量が上限に近づいています。<b>書き出し</b>して、古い手書きメモの削除をおすすめします。</span>
            <button className="tp-bb-act" style={{ background: "#D9534F" }} onClick={() => setSettings(true)}>書き出す</button>
            <button className="tp-bb-x" onClick={() => setQuotaWarn(false)}><X size={15} /></button>
          </div>
        )}
        {backupReminder && (
          <div className="tp-backup-banner">
            <ShieldAlert size={16} />
            <span>データは端末内のみ。<b>バックアップ（書き出し）</b>をおすすめします。</span>
            <button className="tp-bb-act" onClick={() => { setSettings(true); }}>書き出す</button>
            <button className="tp-bb-x" onClick={() => setBackupReminder(false)}><X size={15} /></button>
          </div>
        )}
        <div className="tp-scroll">
          <ErrorBoundary tabKey={tab}>
          {tab === "today" && <TodayView data={data} setData={setData} selDate={selDate} setSelDate={setSelDate} user={user} />}
          {tab === "cal" && <CalendarView data={data} setData={setData} selDate={selDate} setSelDate={setSelDate} vis={calVis} toggleVis={toggleVis} onPrint={onPrint} onPrintCal={onPrintCal} onPrintWeekplan={(start) => setPrintData({ mode: "weekplan", start, title: "週案" })} goToday={(d) => { setSelDate(d); setTab("today"); }} period={calPeriod} setPeriod={setCalPeriod} />}
          {tab === "classes" && <ClassesView data={data} setData={setData} />}
          {tab === "club" && <ClubView data={data} setData={setData} onShare={(title, rows) => setPrintData({ title, rows, vis: { club: true }, mode: "list" })} />}
          {tab === "roster" && <RosterView data={data} setData={setData} jump={rosterJump} showToast={showToast} onPrintSeating={(klass, cols) => setPrintData({ mode: "seating", klass, cols, title: `${klass} 座席表` })} />}
          {tab === "duties" && <DutiesView data={data} setData={setData} />}
          {tab === "meetings" && <MeetingsView data={data} setData={setData} user={user} jump={meetingJump} onPrintMeeting={(id) => setPrintData({ mode: "meeting", meetingId: id, user, title: "職員会議" })} onPrintMeetings={(ids, title) => setPrintData({ mode: "meetings", meetingIds: ids, user, title })} />}
          </ErrorBoundary>
        </div>
      </main>
      <SettingsModal open={settings} onClose={() => setSettings(false)} data={data} setData={setData} user={user} onLogout={doLogout} onExported={() => setBackupReminder(false)} pushBackup={pushBackup} showToast={showToast} />
      {printData && <PrintRoot data={data} meta={data.meta} printData={printData} onClose={() => setPrintData(null)} />}
      {data.meta.onboarded === false && <OnboardingModal data={data} setData={setData} openImport={() => { setData((d) => ({ ...d, meta: { ...d.meta, onboarded: true } })); setSettings(true); }} />}
      <GlobalSearch open={searchOpen} data={data} onClose={() => setSearchOpen(false)} onJump={(r) => {
        setSearchOpen(false);
        if (r.date) { try { setSelDate(parseYmd(r.date)); } catch (e) {} }
        if (r.goto === "roster" && r.klass) setRosterJump({ klass: r.klass, at: Date.now() });
        if (r.goto === "meetings" && r.meetingId) setMeetingJump({ id: r.meetingId, at: Date.now() });
        setTab(r.goto || "today");
      }} />
      <TodoDock data={data} setData={setData} mode={todoMode} setMode={setTodoMode} />
      {toast && (
        <div className="tp-toast">
          <span>{toast.msg}</span>
          {toast.prev && <button className="tp-toast-undo" onClick={() => { setData(toast.prev); setToast(null); }}><RotateCcw size={14} /> 元に戻す</button>}
          <button className="tp-toast-x" onClick={() => setToast(null)}><X size={15} /></button>
        </div>
      )}
    </div>
  );
}

/* ============================================================
   PrintRoot — 印刷 / PDF保存
   ============================================================ */
function PrintCalendar({ data, months, vis }) {
  return (
    <div className="tp-pcal">
      {months.map((mDate, mi) => {
        const y = mDate.getFullYear(), m = mDate.getMonth();
        const first = new Date(y, m, 1);
        const pad0 = first.getDay();
        const dim = new Date(y, m + 1, 0).getDate();
        const cells = [];
        for (let i = 0; i < pad0; i++) cells.push(null);
        for (let d = 1; d <= dim; d++) cells.push(new Date(y, m, d));
        while (cells.length % 7 !== 0) cells.push(null);
        return (
          <div className="tp-pcal-month" key={mi}>
            <div className="tp-pcal-title">{y}年 {m + 1}月</div>
            <div className="tp-pcal-grid">
              {WD.map((w, i) => <div key={"h" + i} className={"tp-pcal-wd wd" + i}>{w}</div>)}
              {cells.map((date, i) => {
                if (!date) return <div key={i} className="tp-pcal-cell empty" />;
                const items = calItemsForDate(data, date, vis);
                return (
                  <div key={i} className={"tp-pcal-cell wd" + date.getDay()}>
                    <div className="tp-pcal-num">{date.getDate()}</div>
                    <div className="tp-pcal-items">
                      {items.slice(0, 4).map((it, j) => (
                        it.club
                          ? <div key={j} className="tp-pcal-club" style={{ borderColor: it.color }}>{[it.lines[0], it.lines[1], it.lines[2]].filter(Boolean).join(" ")}</div>
                          : <div key={j} className="tp-pcal-ev"><span className="tp-pcal-dot" style={{ background: it.color }} />{it.time ? it.time + " " : ""}{it.title}</div>
                      ))}
                      {items.length > 4 && <div className="tp-pcal-more">+{items.length - 4}</div>}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        );
      })}
    </div>
  );
}

function PrintRoot({ data, meta, printData, onClose }) {
  const { title, rows, months, vis, auto, mode, klass, cols, start, meetingId, user } = printData;
  const isCal = mode === "calendar";
  const isSeat = mode === "seating";
  const isWeekplan = mode === "weekplan";
  const isMeeting = mode === "meeting";
  const isMeetings = mode === "meetings";
  const [, setPtick] = useState(0);
  useEffect(() => { if (isMeeting) { const m = (data.meetings || []).find((x) => x.id === meetingId); if (m) preloadMtgImgs(user, m.imgs || []).then(() => setPtick((t) => t + 1)); } }, [isMeeting, meetingId, user]);
  useEffect(() => { if (isMeetings) { const ids = []; (printData.meetingIds || []).forEach((id) => { const m = (data.meetings || []).find((x) => x.id === id); if (m) (m.imgs || []).forEach((i) => ids.push(i)); }); preloadMtgImgs(user, ids).then(() => setPtick((t) => t + 1)); } }, [isMeetings]);
  useEffect(() => { if (auto) { const t = setTimeout(() => { try { window.print(); } catch (e) {} }, 350); return () => clearTimeout(t); } }, [auto]);
  const legend = !isSeat && !isWeekplan && !isMeeting && !isMeetings ? CAL_CATS.filter((c) => vis[c.id]) : [];
  if (isMeetings) {
    const list = (printData.meetingIds || []).map((id) => (data.meetings || []).find((x) => x.id === id)).filter(Boolean);
    return (
      <div className="tp-print-root">
        <div className="tp-print-bar">
          <span className="tp-print-barinfo">まとめPDFプレビュー（{list.length}件）</span><span className="tp-print-spacer" />
          <button className="tp-ghostbtn" onClick={() => { try { window.print(); } catch (e) {} }}><Printer size={14} /> 印刷</button>
          <button className="tp-ghostbtn" onClick={() => { try { window.print(); } catch (e) {} }}><Download size={14} /> PDF保存</button>
          <button className="tp-iconbtn" onClick={onClose}><X size={18} /></button>
        </div>
        <div className="tp-print-scroll">
          <div className="tp-print-sheet">
            <div className="tp-print-header"><h1>{title || "職員会議まとめ"}</h1><div className="tp-print-meta">{meta.year}{meta.teacher && `　${meta.teacher}`}　全{list.length}件</div></div>
            {list.map((m, idx) => { const cm = catMeta(m.cat); const d = parseYmd(m.date); return (
              <div key={m.id} className={"tp-mtg-sec" + (idx > 0 ? " brk" : "")}>
                <div className="tp-mtg-sechead"><span className="tp-mtg-tag" style={{ background: cm.color }}>{cm.label}</span><span className="tp-mtg-secdate">{d.getFullYear()}/{d.getMonth() + 1}/{d.getDate()}</span><span className="tp-mtg-sectitle">{m.title || "（無題）"}</span></div>
                {m.notes && <div className="tp-mtg-pdfnote">{m.notes}</div>}
                {(m.imgs || []).map((imgId) => { const url = getMtgImg(user, imgId); const cap = (m.caps || {})[imgId]; return url ? <div key={imgId}><img className="tp-mtg-pdfimg" src={url} alt="資料" />{cap && <div className="tp-mtg-pdfcap">{cap}</div>}</div> : null; })}
              </div>
            ); })}
            <div className="tp-print-foot">Teaching Partner for YOU — 職員会議まとめ</div>
          </div>
        </div>
      </div>
    );
  }
  if (isMeeting) {
    const m = (data.meetings || []).find((x) => x.id === meetingId);
    const cm = m ? catMeta(m.cat) : null;
    const d = m ? parseYmd(m.date) : null;
    return (
      <div className="tp-print-root">
        <div className="tp-print-bar">
          <span className="tp-print-barinfo">職員会議PDFプレビュー</span><span className="tp-print-spacer" />
          <button className="tp-ghostbtn" onClick={() => { try { window.print(); } catch (e) {} }}><Printer size={14} /> 印刷</button>
          <button className="tp-ghostbtn" onClick={() => { try { window.print(); } catch (e) {} }}><Download size={14} /> PDF保存</button>
          <button className="tp-iconbtn" onClick={onClose}><X size={18} /></button>
        </div>
        <div className="tp-print-scroll">
          <div className="tp-print-sheet">
            {m ? (<>
              <div className="tp-print-header">
                <h1>{m.title || "職員会議"}</h1>
                <div className="tp-print-meta">{d.getFullYear()}年{d.getMonth() + 1}月{d.getDate()}日　<span className="tp-mtg-tag" style={{ background: cm.color }}>{cm.label}</span>{meta.teacher && `　記録：${meta.teacher}`}</div>
              </div>
              {m.notes && <div className="tp-mtg-pdfnote">{m.notes}</div>}
              {(m.imgs || []).map((imgId) => { const url = getMtgImg(user, imgId); const cap = (m.caps || {})[imgId]; return url ? <div key={imgId}><img className="tp-mtg-pdfimg" src={url} alt="資料" />{cap && <div className="tp-mtg-pdfcap">{cap}</div>}</div> : null; })}
              <div className="tp-print-foot">Teaching Partner for YOU — 職員会議記録</div>
            </>) : <p>会議が見つかりません。</p>}
          </div>
        </div>
      </div>
    );
  }
  if (isWeekplan) {
    const monday = parseYmd(start);
    const cols2 = meta.includeSat ? 6 : 5;
    const days = Array.from({ length: cols2 }, (_, i) => addDays(monday, i));
    const periods = data.periods || [];
    return (
      <div className="tp-print-root">
        <div className="tp-print-bar">
          <span className="tp-print-barinfo">週案プレビュー</span><span className="tp-print-spacer" />
          <button className="tp-ghostbtn" onClick={() => { try { window.print(); } catch (e) {} }}><Printer size={14} /> 印刷</button>
          <button className="tp-ghostbtn" onClick={() => { try { window.print(); } catch (e) {} }}><Download size={14} /> PDF保存</button>
          <button className="tp-iconbtn" onClick={onClose}><X size={18} /></button>
        </div>
        <div className="tp-print-scroll">
          <div className="tp-print-sheet landscape">
            <div className="tp-print-header"><h1>週案　{monday.getFullYear()}年 {monday.getMonth() + 1}/{monday.getDate()}（{WD[monday.getDay()]}）の週</h1><div className="tp-print-meta">{meta.year}{meta.teacher && `　${meta.teacher}`}</div></div>
            <table className="tp-wp-table">
              <thead><tr><th className="tp-wp-corner">限</th>{days.map((d, i) => <th key={i} className={"wd" + d.getDay()}>{d.getMonth() + 1}/{d.getDate()}<br /><small>（{WD[d.getDay()]}）</small></th>)}</tr></thead>
              <tbody>
                {periods.map((p, pi) => (
                  <tr key={pi}>
                    <th className="tp-wp-p">{p.label}<br /><small>{p.start}</small></th>
                    {days.map((d, di) => {
                      const cell = data.timetable[`${di}-${pi}`];
                      const log = data.lessonLog[`${ymd(d)}-${pi}`];
                      return (
                        <td key={di} className="tp-wp-cell">
                          {cell?.subject && <span className="tp-wp-sub" style={{ background: subjColor(data, cell.subject) }}>{cell.subject}{cell.klass ? ` ${cell.klass}` : ""}</span>}
                          {log?.topic && <span className="tp-wp-topic">{log.topic}</span>}
                          {log?.hw && <span className="tp-wp-hw">宿:{log.hw}</span>}
                        </td>
                      );
                    })}
                  </tr>
                ))}
                <tr><th className="tp-wp-p">放課後<br /><small>部活等</small></th>{days.map((d, di) => { const cd = clubDayDisplay(data, ymd(d), di); return <td key={di} className="tp-wp-cell after">{cd && cd.kind !== "off" ? `${cd.content}${cd.time ? " " + cd.time : ""}` : ""}</td>; })}</tr>
              </tbody>
            </table>
            <div className="tp-print-foot">Teaching Partner for YOU — 週案</div>
          </div>
        </div>
      </div>
    );
  }
  if (isSeat) {
    const students = (data.rosters[klass] || []).slice().sort((a, b) => (Number(a.no) || 0) - (Number(b.no) || 0));
    const nCols = cols || 6;
    const nRows = Math.max(1, Math.ceil(students.length / nCols));
    return (
      <div className="tp-print-root">
        <div className="tp-print-bar">
          <span className="tp-print-barinfo">座席表プレビュー</span>
          <span className="tp-print-spacer" />
          <button className="tp-ghostbtn" onClick={() => { try { window.print(); } catch (e) {} }}><Printer size={14} /> 印刷</button>
          <button className="tp-ghostbtn" onClick={() => { try { window.print(); } catch (e) {} }}><Download size={14} /> PDF保存</button>
          <button className="tp-iconbtn" onClick={onClose}><X size={18} /></button>
        </div>
        <div className="tp-print-scroll">
          <div className="tp-print-sheet">
            <div className="tp-print-header"><h1>{title}</h1><div className="tp-print-meta">{meta.year}{meta.teacher && `　${meta.teacher}`}　（{students.length}名）</div></div>
            <div className="tp-seat-board">前（黒板）</div>
            <div className="tp-seat-grid" style={{ gridTemplateColumns: `repeat(${nCols}, 1fr)` }}>
              {Array.from({ length: nRows * nCols }).map((_, i) => { const s = students[i]; return (
                <div key={i} className={"tp-seat" + (s ? "" : " empty")}>{s ? <><span className="tp-seat-no">{s.no}</span><span className="tp-seat-name">{s.name}</span></> : ""}</div>
              ); })}
            </div>
            <div className="tp-print-foot">Teaching Partner for YOU — {new Date().getFullYear()}/{new Date().getMonth() + 1}/{new Date().getDate()} 作成</div>
          </div>
        </div>
      </div>
    );
  }
  return (
    <div className="tp-print-root">
      {isCal && <style>{"@media print{@page{size:A4 landscape;margin:7mm;}}"}</style>}
      <div className="tp-print-bar">
        <span className="tp-print-barinfo">A4プレビュー{isCal ? "（横向き・1枚）" : ""}</span>
        <span className="tp-print-spacer" />
        <button className="tp-ghostbtn" onClick={() => { try { window.print(); } catch (e) {} }}><Printer size={14} /> 印刷</button>
        <button className="tp-ghostbtn" onClick={() => { try { window.print(); } catch (e) {} }}><Download size={14} /> PDF保存</button>
        <button className="tp-iconbtn" onClick={onClose}><X size={18} /></button>
      </div>
      <div className="tp-print-scroll">
        <div className={"tp-print-sheet" + (isCal ? " landscape cal" : "")}>
          <div className="tp-print-header">
            <h1>{title}</h1>
            <div className="tp-print-meta">{meta.year}{meta.teacher && `　${meta.teacher}`}<span className="tp-print-legend inline">{legend.map((c) => <span key={c.id}><span className="tp-cat-dot" style={{ background: c.color }} />{c.label}</span>)}</span></div>
          </div>
          {isCal ? (
            <PrintCalendar data={data} months={months} vis={vis} />
          ) : (
            <>
              <div className="tp-print-legend">
                {legend.map((c) => <span key={c.id}><span className="tp-cat-dot" style={{ background: c.color }} />{c.label}</span>)}
              </div>
              {rows.length === 0 ? <p className="tp-print-empty">表示するカテゴリの予定がありません。</p> : (
                <table className="tp-print-table">
                  <tbody>
                    {rows.map((r, i) => (
                      <tr key={i}>
                        <td className={"tp-print-date wd" + r.date.getDay()}>{r.date.getMonth() + 1}/{r.date.getDate()}<small>（{WD[r.date.getDay()]}）</small></td>
                        <td className="tp-print-items">
                          {r.items.map((it, j) => (
                            it.club && it.lines
                              ? <span key={j} className="tp-print-clubitem"><span className="tp-cat-dot" style={{ background: it.color }} />{it.lines.filter(Boolean).join("／")}</span>
                              : <span key={j} className="tp-print-item"><span className="tp-cat-dot" style={{ background: it.color }} />{it.time && <b>{it.time}</b>}{it.title}</span>
                          ))}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </>
          )}
          <div className="tp-print-foot">Teaching Partner for YOU — {new Date().getFullYear()}/{new Date().getMonth() + 1}/{new Date().getDate()} 作成</div>
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   CSS
   ============================================================ */
const CSS = `
* { box-sizing: border-box; }
.tp-app{
  --paper:#EEF2F5; --card:#FFFFFF; --ink:#24323C; --muted:#7A8894; --line:#E4EAEF;
  --sky:#3E9BC9; --sky-deep:#2C7CA6; --sky-soft:#E7F2F8; --coral:#E8845B; --green:#4F9E86; --amber:#E0A64B;
  font-family:"Hiragino Sans","Hiragino Kaku Gothic ProN","Yu Gothic","Noto Sans JP",Meiryo,sans-serif;
  color:var(--ink); background:var(--paper); display:flex; height:100vh; width:100%; overflow:hidden; -webkit-font-smoothing:antialiased;
}
.tp-loading{margin:auto;color:var(--muted);}
/* ---- appearance: font scale (zoom on content) + high contrast ---- */
.tp-app.fs-s .tp-scroll{ zoom:0.92; } .tp-app.fs-l .tp-scroll{ zoom:1.12; }
.tp-app.contrast{ --ink:#0E1720; --muted:#4A5661; --line:#C9D3DB; }
/* ---- dark mode ---- */
.tp-app.dark{ --paper:#0E151B; --card:#18222B; --ink:#E7EDF2; --muted:#9AA8B4; --line:#2A3742; --sky-soft:#12303E; }
.tp-app.dark input,.tp-app.dark textarea,.tp-app.dark select{ background:#0E151B; color:var(--ink); border-color:var(--line); }
.tp-app.dark .tp-ghostbtn,.tp-app.dark .tp-iconbtn,.tp-app.dark .tp-catchip,.tp-app.dark .tp-annot-w,.tp-app.dark .tp-donebox,.tp-app.dark .tp-seg{ background:#121B22; border-color:var(--line); color:var(--ink); }
.tp-app.dark .tp-ghostbtn{ color:var(--sky); }
.tp-app.dark .tp-seg-btn.on{ background:var(--sky); color:#0E151B; }
.tp-app.dark .tp-modal{ background:var(--card); color:var(--ink); }
.tp-app.dark .tp-modal-back{ background:rgba(0,0,0,.6); }
.tp-app.dark .tp-topbar,.tp-app.dark .tp-tabs{ background:var(--card); border-color:var(--line); }
.tp-app.dark .tp-tab{ color:var(--muted); }
.tp-app.dark .tp-tab.active,.tp-app.dark .tp-tab.on{ color:var(--sky); }
.tp-app.dark .tp-login-card{ background:var(--card); }
.tp-app.dark .tp-card{ background:var(--card); }
.tp-app.dark .tp-set-mobile{ background:#121B22; border-color:var(--line); color:var(--ink); }
.tp-app.dark .tp-mtg-search,.tp-app.dark .tp-gsearch-bar input,.tp-app.dark .tp-gsearch-row:hover{ background:transparent; }
.tp-app.dark .tp-onboard-choice{ background:#121B22; }
.tp-app.dark mark.tp-hl{ background:#7a5f10; color:#fff; }
.tp-app.dark .tp-wp-table thead th{ background:#12252f; } .tp-app.dark .tp-seat-board{ background:#12252f; }
.tp-app.dark .tp-card-title{ color:var(--ink); }
.tp-app.dark .tp-divider{ background:var(--line); }
.tp-app.dark .tp-hint,.tp-app.dark .tp-empty{ color:var(--muted); }
.tp-app.dark .tp-subtab,.tp-app.dark .tp-rtab{ background:#121B22; border-color:var(--line); color:var(--muted); }
.tp-app.dark .tp-subtab.on,.tp-app.dark .tp-rtab.on{ background:var(--sky); color:#0E151B; }
.tp-app.dark .tp-primarybtn{ background:var(--sky); color:#08222e; }
.tp-app.dark .tp-addbtn{ background:var(--sky); color:#08222e; }
.tp-app.dark .tp-cell,.tp-app.dark .tp-mcell,.tp-app.dark .tp-tt-cell,.tp-app.dark .tp-daycell,.tp-app.dark .tp-mini-day{ background:var(--card); border-color:var(--line); }
.tp-app.dark .tp-mtg-item,.tp-app.dark .tp-prog-card,.tp-app.dark .tp-todo-drow{ background:var(--card); }
.tp-app.dark .tp-fchip{ color:var(--muted); border-color:var(--line); }
.tp-app.dark .tp-fchip.on{ background:var(--sky-soft); color:var(--sky); }
.tp-app.dark .tp-todo-due{ background:#12303e; }
.tp-app.dark .tp-print-root,.tp-app.dark .tp-print-sheet{ background:#fff; color:#111; } /* 印刷面は白のまま */
.tp-app.dark .tp-gsearch-gh{ background:var(--card); }
.tp-app.dark .tp-donebox{ background:#0E151B; }
.tp-app.dark .tp-backup-banner{ filter:brightness(.9); }
/* todo hint + search filter + task proposal */
.tp-todo-hint{ font-size:11px; color:var(--muted); margin:2px 0 8px; }
.tp-todo-bulkics{ display:inline-flex; align-items:center; gap:4px; margin-left:6px; background:var(--sky-soft); color:var(--sky-deep); border:1px solid var(--sky); border-radius:14px; padding:3px 9px; font-size:11px; font-weight:800; cursor:pointer; }
.tp-gsearch-tf{ display:flex; gap:5px; flex-wrap:wrap; padding:6px 4px 10px; border-bottom:1px solid var(--line); margin-bottom:6px; }
.tp-taskprop{ max-width:520px; width:92%; }
.tp-taskprop-none{ font-size:15px; font-weight:700; text-align:center; padding:10px 0; }
.tp-taskprop-none span{ font-size:12px; font-weight:400; color:var(--muted); }
.tp-taskprop-list{ display:flex; flex-direction:column; gap:6px; margin:6px 0; }
.tp-taskprop-row{ display:flex; align-items:center; gap:8px; }
.tp-taskprop-text{ flex:1; }
.tp-taskprop-row input[type=date]{ width:130px; }
.tp-taskprop-gen{ display:flex; align-items:center; gap:8px; border-top:1px dashed var(--line); padding-top:10px; margin-top:8px; }
.tp-taskprop-gentext{ flex:1; font-size:13px; color:var(--muted); }
.tp-taskprop-gen input[type=date]{ width:130px; }
/* roadmap config */
.tp-rm-cfg{ display:flex; gap:14px; flex-wrap:wrap; align-items:flex-end; margin:8px 0; }
.tp-rm-cfg label{ display:flex; flex-direction:column; gap:4px; font-size:12px; color:var(--muted); font-weight:700; }
.tp-rm-cfg input[type=number]{ width:90px; }
.tp-rm-phases{ flex:1; min-width:200px; }
.tp-rm-units{ display:flex; flex-direction:column; gap:5px; margin-top:6px; }
.tp-rm-unit{ display:flex; align-items:center; gap:8px; }
.tp-rm-uno{ width:26px; text-align:center; font-weight:800; color:#8894a0; font-size:12px; }
.tp-rm-unit input{ flex:1; }
.tp-sub-testflag{ display:inline-flex; align-items:center; gap:4px; font-size:11px; color:var(--muted); white-space:nowrap; }
.tp-sub-testflag input{ width:auto; }
/* offline text import */
.tp-ocr-guide{ background:var(--sky-soft); border-radius:10px; padding:10px 12px; margin-bottom:10px; font-size:12.5px; }
.tp-ocr-guide ol{ margin:6px 0 4px; padding-left:20px; line-height:1.7; }
.tp-ocr-guide .tp-hint{ margin:4px 0 0; }
.tp-ocr-result{ max-height:38vh; overflow-y:auto; margin-top:6px; border:1px solid var(--line); border-radius:10px; padding:6px; }
.tp-ocr-row{ display:flex; align-items:center; gap:8px; padding:4px 2px; border-bottom:1px solid var(--line); }
.tp-ocr-date{ font-size:12px; color:var(--muted); white-space:nowrap; min-width:96px; }
.tp-ocr-title{ flex:1; }
/* tests panel */
.tp-test-item{ display:flex; align-items:center; justify-content:flex-start; gap:8px; flex-wrap:wrap; padding:8px 0; border-bottom:1px solid var(--line); }
.tp-test-item .tp-test-name{ font-weight:700; }
.tp-test-range{ margin-left:2px; }
.tp-test-range input[type=number]{ width:52px; }
.tp-test-unit{ font-size:12px; max-width:150px; }
.tp-test-del{ margin-left:auto; }
.tp-testfound{ background:var(--sky-soft); border-radius:10px; padding:8px 10px; margin-bottom:8px; }
.tp-testfound-row{ display:flex; align-items:center; gap:8px; padding:4px 2px; font-size:13px; cursor:pointer; }
/* 時数カウンタ */
.tp-seqbadge{ font-size:11px; font-weight:800; color:var(--sky-deep); background:var(--sky-soft); border-radius:10px; padding:1px 8px; white-space:nowrap; }
.tp-prog2{ display:flex; flex-direction:column; gap:8px; }
.tp-prog2-row{ display:flex; flex-direction:column; gap:4px; padding:6px 0; border-bottom:1px solid var(--line); }
.tp-prog2-sub{ display:flex; align-items:center; gap:6px; font-weight:800; font-size:14px; }
.tp-prog2-gap{ font-size:11px; font-weight:700; color:#D9534F; background:#FDE8E8; border-radius:10px; padding:1px 7px; }
.tp-prog2-klasses{ display:flex; flex-wrap:wrap; gap:6px; }
.tp-prog2-k{ display:inline-flex; align-items:center; gap:5px; font-size:12px; color:var(--muted); border:1px solid var(--line); border-radius:8px; padding:3px 8px; }
.tp-prog2-k b{ color:var(--ink); font-size:14px; }
.tp-prog2-k.hi{ border-color:#4F9E86; background:#EAF6F1; } .tp-prog2-k.hi b{ color:#3B7A66; }
.tp-prog2-k.lo{ border-color:#D9A54B; background:#FBF3E2; } .tp-prog2-k.lo b{ color:#B07C1E; }
.tp-term-edit{ display:flex; align-items:center; gap:8px; margin-bottom:6px; }
.tp-term-name{ width:110px; }
/* bulk timetable */
.tp-bulk-grid{ display:grid; gap:4px; margin:8px 0; }
.tp-bulk-corner,.tp-bulk-head,.tp-bulk-period{ font-size:11px; font-weight:800; color:var(--muted); display:flex; align-items:center; justify-content:center; padding:4px; }
.tp-bulk-cell{ min-height:44px; border:1px solid var(--line); border-radius:8px; background:var(--card); cursor:pointer; display:flex; align-items:center; justify-content:center; font-size:10px; color:var(--muted); padding:2px; }
.tp-bulk-cell.on{ background:var(--sky); border-color:var(--sky-deep); color:#fff; }
.tp-bulk-cur{ line-height:1.2; text-align:center; }
.tp-switch{ display:flex; align-items:center; gap:10px; font-size:14px; margin:8px 0; cursor:pointer; }
.tp-switch input{ width:auto; }
/* ---- toast / undo snackbar ---- */
.tp-toast{ position:fixed; left:50%; bottom:22px; transform:translateX(-50%); z-index:300; background:#24323C; color:#fff; border-radius:12px; padding:10px 12px 10px 16px; display:flex; align-items:center; gap:12px; font-size:13.5px; box-shadow:0 8px 30px rgba(0,0,0,.3); max-width:92%; }
.tp-toast-undo{ background:#3E9BC9; color:#fff; border:none; border-radius:8px; padding:6px 12px; font-size:13px; font-weight:800; cursor:pointer; display:inline-flex; align-items:center; gap:5px; }
.tp-toast-x{ background:none; border:none; color:#9fb0bd; cursor:pointer; display:flex; }
/* ---- onboarding ---- */
.tp-onboard{ max-width:460px; }
.tp-onboard-head{ display:flex; align-items:center; gap:12px; margin-bottom:14px; }
.tp-onboard-head svg{ border-radius:11px; }
.tp-onboard-choices{ display:flex; flex-direction:column; gap:10px; margin:6px 0 4px; }
.tp-onboard-choice{ text-align:left; border:1px solid var(--line); background:#fafcfd; border-radius:12px; padding:13px 15px; cursor:pointer; display:flex; flex-direction:column; gap:3px; }
.tp-onboard-choice:hover{ border-color:var(--sky); background:var(--sky-soft); }
.tp-onboard-choice b{ font-size:15px; color:var(--sky-deep); }
.tp-onboard-choice span{ font-size:12px; color:var(--muted); }
/* ---- seating chart ---- */
.tp-seat-ctl{ display:inline-flex; align-items:center; gap:6px; font-size:12px; color:var(--muted); }
.tp-seat-ctl select{ width:auto; padding:6px 8px; }
.tp-seat-board{ text-align:center; font-weight:800; color:#2C7CA6; background:#EEF3F6; border-radius:8px; padding:8px; margin:6px 0 12px; letter-spacing:4px; }
.tp-seat-grid{ display:grid; gap:8px; }
.tp-seat{ border:1.5px solid #cdd8e0; border-radius:10px; min-height:64px; display:flex; flex-direction:column; align-items:center; justify-content:center; gap:3px; padding:6px; }
.tp-seat.empty{ border-style:dashed; background:#fafbfc; }
.tp-seat-no{ font-size:11px; color:#8894a0; font-weight:700; }
.tp-seat-name{ font-size:15px; font-weight:700; }
@media print { .tp-seat{ min-height:20mm; } }
/* ---- test progress plan ---- */
.tp-prog{ display:flex; flex-direction:column; gap:10px; margin-top:6px; }
.tp-prog-card{ border:1px solid var(--line); border-radius:12px; padding:12px; }
.tp-prog-head{ display:flex; align-items:center; gap:10px; flex-wrap:wrap; margin-bottom:10px; }
.tp-prog-name{ font-weight:800; font-size:15px; color:var(--sky-deep); }
.tp-prog-date{ font-size:12px; color:var(--muted); }
.tp-prog-range{ font-size:12px; color:var(--muted); background:var(--sky-soft); padding:2px 8px; border-radius:6px; }
.tp-prog-status{ margin-left:auto; font-size:12px; font-weight:800; padding:3px 10px; border-radius:8px; color:#fff; }
.tp-prog-status.s0{ background:#B7C2CB; } .tp-prog-status.s1{ background:#3E9BC9; } .tp-prog-status.s2{ background:#4F9E86; }
.tp-prog-status.s3{ background:#E0A64B; } .tp-prog-status.s4{ background:#D9534F; }
.tp-prog-grid{ display:grid; grid-template-columns:1fr 1fr; gap:8px 12px; }
.tp-prog-grid label{ display:flex; flex-direction:column; gap:3px; font-size:12px; color:var(--muted); }
.tp-prog-grid label span{ font-weight:700; }
.tp-prog-note{ grid-column:1 / -1; }
@media (max-width:640px){ .tp-prog-grid{ grid-template-columns:1fr; } }
/* ---- 週案 print ---- */
.tp-wp-table{ width:100%; border-collapse:collapse; margin-top:8px; table-layout:fixed; }
.tp-wp-table th,.tp-wp-table td{ border:1px solid #cbd5dd; padding:4px; vertical-align:top; }
.tp-wp-table thead th{ background:#eef3f6; font-size:12px; text-align:center; }
.tp-wp-table thead th.wd0{ color:#D9534F; } .tp-wp-table thead th.wd6{ color:#2C7CA6; }
.tp-wp-corner{ width:44px; }
.tp-wp-p{ width:44px; background:#f6f8fa; text-align:center; font-size:11px; font-weight:700; }
.tp-wp-p small{ font-weight:400; color:#8894a0; }
.tp-wp-cell{ height:58px; }
.tp-wp-cell.after{ height:auto; font-size:11px; }
.tp-wp-sub{ display:inline-block; color:#fff; font-size:10px; font-weight:700; padding:1px 5px; border-radius:4px; margin-bottom:3px; }
.tp-wp-topic{ display:block; font-size:11px; line-height:1.3; }
.tp-wp-hw{ display:block; font-size:10px; color:#8894a0; margin-top:2px; }
/* ---- 職員会議 ---- */
.tp-mtg-list{ display:flex; flex-direction:column; gap:8px; margin-top:8px; }
.tp-mtg-item{ border:1px solid var(--line); border-left-width:5px; border-radius:10px; padding:10px 12px; display:flex; align-items:flex-start; gap:10px; cursor:pointer; }
.tp-mtg-item:hover{ background:var(--sky-soft); }
.tp-mtg-main{ flex:1; min-width:0; }
.tp-mtg-top{ display:flex; align-items:center; gap:8px; flex-wrap:wrap; }
.tp-mtg-date{ font-size:12px; color:var(--muted); font-weight:700; }
.tp-mtg-tag{ font-size:10px; color:#fff; font-weight:800; padding:2px 8px; border-radius:20px; }
.tp-mtg-title{ font-size:15px; font-weight:700; margin-top:2px; }
.tp-mtg-preview{ font-size:12px; color:var(--muted); margin-top:3px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.tp-mtg-att{ font-size:11px; color:var(--sky-deep); margin-top:3px; }
.tp-mtg-gro	{}
.tp-mtg-grouphead{ font-size:12px; font-weight:800; color:var(--muted); margin:12px 0 4px; display:flex; align-items:center; gap:8px; }
.tp-mtg-grouphead .tp-dot2{ width:12px; height:12px; border-radius:4px; }
.tp-catchips{ display:flex; gap:6px; flex-wrap:wrap; }
.tp-catchip{ border:1px solid var(--line); background:#fff; border-radius:20px; padding:6px 12px; font-size:12px; font-weight:700; cursor:pointer; color:var(--muted); }
.tp-catchip.on{ color:#fff; border-color:transparent; }
.tp-mtg-thumbs{ display:flex; gap:8px; flex-wrap:wrap; margin-top:8px; }
.tp-mtg-thumb{ position:relative; width:70px; height:70px; border-radius:8px; overflow:hidden; border:1px solid var(--line); }
.tp-mtg-thumb img{ width:100%; height:100%; object-fit:cover; }
.tp-mtg-thumb button{ position:absolute; top:2px; right:2px; background:rgba(0,0,0,.55); border:none; color:#fff; border-radius:50%; width:20px; height:20px; cursor:pointer; display:flex; align-items:center; justify-content:center; }
.tp-mtg-thumb-load{ display:block; width:100%; height:100%; background:#eef3f6; }
.tp-mtg-thumb-pen{ position:absolute; bottom:2px; left:2px; background:rgba(255,255,255,.9); color:#2C7CA6; border-radius:5px; width:18px; height:18px; display:flex; align-items:center; justify-content:center; pointer-events:none; }
/* annotate modal */
.tp-annot{ max-width:960px; width:96vw; }
.tp-annot-tools{ display:flex; align-items:center; gap:8px; flex-wrap:wrap; margin-bottom:10px; }
.tp-annot-color{ width:26px; height:26px; border-radius:50%; border:2px solid #fff; box-shadow:0 0 0 1px var(--line); cursor:pointer; }
.tp-annot-color.on{ box-shadow:0 0 0 2px var(--sky-deep); }
.tp-annot-w{ width:30px; height:30px; border:1px solid var(--line); border-radius:8px; background:#fff; cursor:pointer; display:flex; align-items:center; justify-content:center; }
.tp-annot-w.on{ border-color:var(--sky-deep); background:var(--sky-soft); }
.tp-annot-w span{ background:#333; border-radius:50%; display:block; }
.tp-annot-sep{ width:1px; height:22px; background:var(--line); }
.tp-annot-canvaswrap{ max-height:62vh; overflow:auto; background:#f2f5f7; border:1px solid var(--line); border-radius:10px; display:flex; justify-content:center; }
.tp-annot-canvas{ touch-action:none; max-width:100%; background:#fff; }
.tp-mtg-pdfimg{ width:100%; margin:8px 0; border:1px solid #ddd; border-radius:6px; }
.tp-mtg-pdfnote{ white-space:pre-wrap; font-size:13px; line-height:1.7; }
/* meeting search + captions */
.tp-mtg-search{ display:flex; align-items:center; gap:8px; border:1px solid var(--line); border-radius:10px; padding:6px 10px; margin:8px 0; color:var(--muted); }
.tp-mtg-search input{ border:none; outline:none; flex:1; font-size:14px; background:transparent; padding:2px 0; }
.tp-mtg-hitwhere{ font-size:10px; color:#fff; background:var(--muted); border-radius:10px; padding:1px 7px; margin-left:auto; }
.tp-hl{ background:#FCE9A6; color:inherit; border-radius:3px; padding:0 1px; }
.tp-mtg-jump{ align-self:center; font-size:12px; font-weight:800; color:var(--sky-deep); white-space:nowrap; }
.tp-mtg-thumbcol{ display:flex; flex-direction:column; gap:4px; width:88px; }
.tp-mtg-thumbcol .tp-mtg-thumb{ width:88px; height:88px; }
.tp-mtg-cap{ width:88px; font-size:11px; padding:4px 6px; }
.tp-mtg-pdfcap{ font-size:12px; color:#555; margin:-4px 0 10px; }
/* global search */
.tp-gsearch{ max-width:560px; width:92%; padding:0; overflow:hidden; }
.tp-gsearch-bar{ display:flex; align-items:center; gap:10px; padding:12px 14px; border-bottom:1px solid var(--line); color:var(--muted); }
.tp-gsearch-bar input{ flex:1; border:none; outline:none; font-size:16px; background:transparent; }
.tp-gsearch-body{ max-height:60vh; overflow-y:auto; padding:8px 10px 14px; }
.tp-gsearch-group{ margin-top:8px; }
.tp-gsearch-gh{ font-size:12px; font-weight:800; color:var(--muted); padding:4px 6px; position:sticky; top:0; background:var(--card); }
.tp-gsearch-row{ display:flex; align-items:center; gap:10px; width:100%; text-align:left; background:none; border:none; border-radius:8px; padding:9px 8px; cursor:pointer; }
.tp-gsearch-row:hover{ background:var(--sky-soft); }
.tp-gsearch-label{ font-size:14px; color:var(--ink); white-space:nowrap; overflow:hidden; text-overflow:ellipsis; max-width:52%; }
.tp-gsearch-sub{ font-size:12px; color:var(--muted); margin-left:auto; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; max-width:38%; }
.tp-gsearch-go{ color:var(--sky-deep); font-weight:800; }
/* ---- ToDo dock ---- */
.tp-todo-min{ position:fixed; right:16px; bottom:80px; z-index:120; display:flex; align-items:center; gap:6px; background:var(--sky-deep); color:#fff; border:none; border-radius:22px; padding:9px 14px; font-size:13px; font-weight:800; box-shadow:0 6px 20px rgba(44,124,166,.35); cursor:pointer; }
.tp-todo-badge{ background:#fff; color:var(--sky-deep); border-radius:12px; font-size:11px; padding:1px 7px; }
.tp-todo-side{ position:fixed; top:0; right:0; width:340px; max-width:88vw; height:100vh; background:var(--card); border-left:1px solid var(--line); box-shadow:-8px 0 30px rgba(0,0,0,.12); z-index:130; display:flex; flex-direction:column; padding:14px; }
.tp-todo-full{ background:var(--card); border-radius:16px; width:92%; max-width:620px; max-height:86vh; padding:16px; display:flex; flex-direction:column; }
.tp-todo-dockhead{ display:flex; align-items:center; justify-content:space-between; margin-bottom:10px; }
.tp-todo-docktitle{ display:flex; align-items:center; gap:7px; font-weight:800; font-size:15px; }
.tp-todo-modes{ display:flex; gap:4px; }
.tp-iconbtn.tiny.on{ background:var(--sky-soft); border-color:var(--sky); color:var(--sky-deep); }
.tp-todo-add2{ display:flex; gap:6px; margin-bottom:8px; }
.tp-todo-add2 input[type=text],.tp-todo-add2 input:not([type=date]){ flex:1; }
.tp-todo-add2 input[type=date]{ width:130px; }
.tp-addbtn{ background:var(--sky); color:#fff; border:none; border-radius:8px; width:38px; display:flex; align-items:center; justify-content:center; cursor:pointer; }
.tp-todo-catrow{ display:flex; gap:5px; flex-wrap:wrap; margin-bottom:8px; }
.tp-catchip.mini{ padding:4px 9px; font-size:11px; }
.tp-todo-filter{ display:flex; gap:5px; flex-wrap:wrap; margin-bottom:8px; }
.tp-fchip{ display:inline-flex; align-items:center; gap:4px; border:1px solid var(--line); background:transparent; border-radius:16px; padding:4px 9px; font-size:11px; color:var(--muted); cursor:pointer; }
.tp-fchip.on{ background:var(--sky-soft); border-color:var(--sky); color:var(--sky-deep); font-weight:700; }
.tp-todo-docklist{ overflow-y:auto; flex:1; display:flex; flex-direction:column; gap:4px; }
.tp-todo-drow{ display:flex; align-items:center; gap:8px; padding:7px 4px; border-bottom:1px solid var(--line); }
.tp-todo-drow.done .tp-todo-dtext{ text-decoration:line-through; color:var(--muted); }
.tp-todo-cat{ width:8px; height:8px; border-radius:50%; flex-shrink:0; }
.tp-todo-dtext{ flex:1; font-size:13.5px; min-width:0; overflow:hidden; text-overflow:ellipsis; }
.tp-todo-due{ font-size:11px; color:var(--muted); background:var(--sky-soft); border-radius:6px; padding:2px 6px; white-space:nowrap; }
.tp-todo-due.over{ background:#FDE8E8; color:#D9534F; font-weight:800; }
@media (max-width:640px){ .tp-todo-min{ bottom:calc(72px + env(safe-area-inset-bottom, 0px)); } }
.tp-searchbtn{ color:var(--sky-deep); }
.tp-cardsub{ font-size:11px; font-weight:400; color:var(--muted); margin-left:8px; }
.tp-linkbtn{ background:none; border:none; color:var(--sky-deep); font-weight:700; font-size:12px; cursor:pointer; padding:0; margin-left:6px; text-decoration:underline; }
.tp-period-switch{ margin-bottom:10px; }
.tp-period-switch .tp-seg{ width:100%; }
.tp-calview .tp-view{ padding:0; }
.tp-quarter.six{ }
@media (min-width:700px){ .tp-quarter.six{ grid-template-columns:repeat(3,1fr); } }
/* guide / help */
.tp-guide{ max-width:460px; }
.tp-guide-step{ text-align:center; padding:10px 6px 4px; }
.tp-guide-ic{ width:56px; height:56px; border-radius:16px; background:var(--sky-soft); color:var(--sky-deep); display:flex; align-items:center; justify-content:center; margin:0 auto 12px; }
.tp-guide-step h4{ margin:0 0 8px; font-size:17px; }
.tp-guide-step p{ margin:0; font-size:14px; line-height:1.75; color:var(--muted); }
.tp-guide-dots{ display:flex; gap:6px; justify-content:center; margin:16px 0 4px; }
.tp-guide-dot{ width:7px; height:7px; border-radius:50%; background:var(--line); }
.tp-guide-dot.on{ background:var(--sky); }
.tp-feat-list{ display:flex; flex-direction:column; gap:2px; max-height:60vh; overflow:auto; }
.tp-feat-row{ display:flex; gap:12px; align-items:flex-start; padding:10px 4px; border-bottom:1px solid var(--line); }
.tp-feat-ic{ flex-shrink:0; width:34px; height:34px; border-radius:9px; background:var(--sky-soft); color:var(--sky-deep); display:flex; align-items:center; justify-content:center; }
.tp-feat-row b{ display:block; font-size:14px; margin-bottom:2px; }
.tp-feat-row span{ font-size:12.5px; color:var(--muted); line-height:1.55; }
.tp-mtg-sec{ margin-bottom:18px; }
.tp-mtg-sec.brk{ break-before:page; page-break-before:always; padding-top:6px; }
.tp-mtg-sechead{ display:flex; align-items:center; gap:10px; border-bottom:2px solid #cbd5dd; padding-bottom:6px; margin-bottom:8px; }
.tp-mtg-secdate{ font-size:12px; color:#667; }
.tp-mtg-sectitle{ font-size:16px; font-weight:800; }
/* tutorial */
.tp-tut-list{ display:flex; flex-direction:column; gap:10px; margin:8px 0; }
.tp-tut-item{ display:flex; gap:12px; align-items:flex-start; }
.tp-tut-icon{ flex:none; width:36px; height:36px; border-radius:9px; background:var(--sky-soft); color:var(--sky-deep); display:flex; align-items:center; justify-content:center; }
.tp-tut-title{ font-weight:800; font-size:14px; color:var(--sky-deep); }
.tp-tut-body{ font-size:13px; color:var(--ink); line-height:1.6; margin-top:2px; }
/* ---- backup banner ---- */
.tp-backup-banner{ display:flex; align-items:center; gap:8px; background:#FFF4E8; border-bottom:1px solid #F1D9BE; color:#9A5B23; font-size:12.5px; padding:8px 14px; }
.tp-backup-banner b{ color:#8a4e18; }
.tp-bb-act{ margin-left:auto; background:#E8845B; color:#fff; border:none; border-radius:8px; padding:5px 12px; font-size:12px; font-weight:800; cursor:pointer; }
.tp-bb-x{ background:none; border:none; color:#b98a56; cursor:pointer; display:flex; padding:2px; }
.tp-backup-list{ list-style:none; margin:8px 0 0; padding:0; display:flex; flex-direction:column; gap:5px; }
.tp-backup-list li{ display:flex; align-items:center; gap:10px; }
.tp-backup-at{ font-size:13px; font-weight:700; color:var(--ink); }
.tp-backup-list .tp-ghostbtn.sm{ margin-left:auto; padding:5px 10px; font-size:12px; }
/* ---- login ---- */
.tp-login{ margin:auto; width:100%; max-width:420px; padding:24px; }
.tp-login-card{ background:#fff; border:1px solid var(--line); border-radius:20px; padding:28px 24px; box-shadow:0 12px 40px rgba(44,124,166,.12); }
.tp-login-brand{ text-align:center; display:flex; flex-direction:column; align-items:center; gap:4px; margin-bottom:14px; }
.tp-login-brand svg{ border-radius:14px; box-shadow:0 6px 16px rgba(44,124,166,.3); margin-bottom:6px; }
.tp-login-lead{ font-size:13px; color:var(--muted); text-align:center; line-height:1.6; margin:0 0 18px; }
.tp-login-sub{ font-size:11.5px; font-weight:700; color:var(--muted); margin:14px 0 8px; }
.tp-login-profiles{ }
.tp-login-prow{ display:flex; align-items:center; gap:6px; margin-bottom:6px; }
.tp-login-pbtn{ flex:1; display:flex; align-items:center; gap:8px; border:1px solid var(--line); background:#fafcfd; border-radius:11px; padding:11px 14px; font-size:14px; font-weight:700; color:var(--ink); cursor:pointer; text-align:left; }
.tp-login-pbtn:hover{ background:var(--sky-soft); border-color:var(--sky); color:var(--sky-deep); }
.tp-login-form{ display:flex; gap:8px; }
.tp-login-form input{ flex:1; font-size:15px; padding:11px 12px; }
.tp-login-form .tp-primarybtn{ white-space:nowrap; }
.tp-login-note{ font-size:11px; color:var(--muted); text-align:center; margin:16px 0 0; }
/* ---- tabs (left on desktop, bottom on mobile) ---- */
.tp-tabs{ width:96px; background:linear-gradient(180deg,#ffffff,#f4f7f9); border-right:1px solid var(--line); display:flex; flex-direction:column; align-items:stretch; padding:14px 8px; gap:4px; flex-shrink:0; z-index:5; overflow-y:auto; }
.tp-brand{ text-align:center; margin-bottom:14px; display:flex; justify-content:center; }
.tp-brand svg{ border-radius:11px; box-shadow:0 4px 12px rgba(44,124,166,.28); }
.tp-tab{ border:none; background:none; padding:11px 4px; border-radius:12px; display:flex; flex-direction:column; align-items:center; gap:4px; color:var(--muted); cursor:pointer; font-size:11.5px; font-weight:600; transition:.15s; }
.tp-tab span{ line-height:1; }
.tp-tab:hover{ background:var(--sky-soft); color:var(--sky-deep); }
.tp-tab.on{ background:var(--sky); color:#fff; box-shadow:0 4px 12px rgba(62,155,201,.35); }
.tp-tab.settings{ margin-top:auto; }
/* ---- main ---- */
.tp-main{ flex:1; display:flex; flex-direction:column; min-width:0; }
.tp-topbar{ padding:14px 24px 10px; display:flex; justify-content:space-between; align-items:center; gap:12px; }
.tp-appbrand{ display:flex; align-items:center; gap:11px; min-width:0; }
.tp-appbrand-logo svg{ border-radius:9px; box-shadow:0 3px 9px rgba(44,124,166,.25); }
.tp-appwordmark{ display:flex; flex-direction:column; line-height:1.02; }
.tp-appname{ font-weight:800; font-size:18px; letter-spacing:.2px; background:linear-gradient(92deg,#2C7CA6,#4AA6D0 60%,#E8845B 190%); -webkit-background-clip:text; background-clip:text; -webkit-text-fill-color:transparent; color:transparent; white-space:nowrap; }
.tp-apptag{ font-size:9px; font-weight:700; letter-spacing:3px; text-transform:uppercase; color:var(--muted); margin-top:2px; }
.tp-topbar-right{ display:flex; align-items:center; gap:12px; flex-shrink:0; }
.tp-section{ text-align:right; display:flex; flex-direction:column; line-height:1.1; }
.tp-section-name{ font-size:16px; font-weight:800; color:var(--ink); }
.tp-topbar-sub{ font-size:11px; font-weight:500; color:var(--muted); }
.tp-set-mobile{ display:none; }
.tp-scroll{ flex:1; overflow-y:auto; padding:4px 24px 40px; }
.tp-view{ display:flex; flex-direction:column; gap:16px; max-width:1040px; }
/* ---- cards ---- */
.tp-card{ background:var(--card); border:1px solid var(--line); border-radius:16px; padding:16px 18px; box-shadow:0 1px 2px rgba(36,50,60,.03); }
.tp-card-title{ margin:0 0 12px; font-size:13.5px; font-weight:700; color:var(--sky-deep); display:flex; align-items:center; gap:6px; }
.tp-cardrow{ display:grid; grid-template-columns:1.7fr 1fr; gap:16px; }
.tp-hint{ font-size:11.5px; color:var(--muted); margin:8px 0 0; line-height:1.5; }
.tp-empty{ color:var(--muted); font-size:13px; padding:10px 2px; }
.tp-divider{ height:1px; background:var(--line); margin:14px 0; }
/* ---- daynav ---- */
.tp-daynav{ display:flex; align-items:center; justify-content:space-between; gap:12px; }
.tp-daynav-mid{ text-align:center; display:flex; flex-direction:column; align-items:center; gap:4px; }
.tp-daynav-date{ font-size:19px; font-weight:800; }
.tp-daynav-date.today{ color:var(--sky-deep); }
.tp-wd{ font-size:14px; margin-left:2px; }
.wd0{ color:#D9534F; } .wd6{ color:var(--sky-deep); }
.tp-iconbtn{ border:1px solid var(--line); background:#fff; border-radius:10px; width:38px; height:38px; display:flex; align-items:center; justify-content:center; color:var(--ink); cursor:pointer; transition:.15s; }
.tp-iconbtn:hover{ background:var(--sky-soft); color:var(--sky-deep); }
.tp-iconbtn.tiny{ width:26px; height:26px; border:none; color:var(--muted); }
.tp-iconbtn.tiny:hover{ color:#D9534F; background:#fbecea; }
.tp-ghostbtn{ border:1px solid var(--line); background:#fff; border-radius:8px; padding:5px 10px; font-size:12px; cursor:pointer; color:var(--sky-deep); font-weight:600; display:inline-flex; align-items:center; gap:4px; }
.tp-ghostbtn.sm{ padding:3px 9px; font-size:11px; }
.tp-ghostbtn:hover{ background:var(--sky-soft); }
/* ---- timeline (today) ---- */
.tp-timeline-card{ }
.tp-timeline{ list-style:none; margin:0; padding:0; }
.tp-tl-item{ display:flex; gap:10px; padding:8px 0; border-bottom:1px dashed var(--line); align-items:flex-start; }
.tp-tl-item:last-child{ border-bottom:none; }
.tp-tl-time{ width:44px; flex-shrink:0; font-variant-numeric:tabular-nums; font-weight:700; font-size:13px; color:var(--muted); padding-top:1px; }
.tp-tl-bar{ width:4px; align-self:stretch; border-radius:3px; background:var(--line); flex-shrink:0; }
.tp-tl-item.k-routine .tp-tl-bar{ background:#DDE4EA; }
.tp-tl-item.k-club .tp-tl-bar,.tp-tl-item.k-clubsp .tp-tl-bar{ background:var(--coral); }
.tp-tl-item.k-duty .tp-tl-bar{ background:#8894A0; }
.tp-tl-body{ flex:1; min-width:0; }
.tp-tl-main{ display:flex; align-items:center; gap:7px; flex-wrap:wrap; font-size:13.5px; }
.tp-tl-item.k-routine .tp-tl-main{ color:var(--muted); font-size:12.5px; }
.tp-tl-room{ font-size:11.5px; color:var(--muted); display:inline-flex; align-items:center; gap:2px; }
.tp-tl-note{ font-size:11.5px; color:var(--coral); }
.tp-tl-topic{ font-size:12px; color:var(--muted); margin-top:3px; padding-left:2px; }
.tp-linkbtn{ border:none; background:none; color:var(--sky-deep); font-size:11.5px; cursor:pointer; display:inline-flex; align-items:center; gap:3px; padding:2px 4px; border-radius:6px; }
.tp-linkbtn:hover{ background:var(--sky-soft); }
.tp-donebox{ width:18px; height:18px; border:1.5px solid var(--muted); border-radius:6px; background:#fff; display:flex; align-items:center; justify-content:center; cursor:pointer; color:#fff; flex-shrink:0; }
.tp-donebox.on{ background:var(--green); border-color:var(--green); }
.tp-chip{ color:#fff; font-size:11px; font-weight:700; padding:2px 8px; border-radius:20px; white-space:nowrap; }
.tp-chip.sm{ font-size:10px; padding:1px 7px; }
/* ---- todo ---- */
.tp-todo-add{ display:flex; gap:6px; margin-bottom:8px; }
.tp-todo-add input{ flex:1; }
.tp-todo-list{ list-style:none; margin:0; padding:0; }
.tp-todo-list li{ display:flex; align-items:center; gap:8px; padding:6px 0; font-size:13px; border-bottom:1px dashed var(--line); }
.tp-todo-list li span{ flex:1; }
.tp-todo-list li.done span{ text-decoration:line-through; color:var(--muted); }
.tp-addbtn{ border:none; background:var(--sky); color:#fff; border-radius:9px; width:34px; height:34px; display:flex; align-items:center; justify-content:center; cursor:pointer; flex-shrink:0; }
.tp-addbtn:hover{ background:var(--sky-deep); }
/* ---- inputs ---- */
input,select,textarea{ font-family:inherit; font-size:13px; color:var(--ink); border:1px solid var(--line); border-radius:9px; padding:8px 10px; background:#fff; outline:none; transition:.15s; }
input:focus,select:focus,textarea:focus{ border-color:var(--sky); box-shadow:0 0 0 3px var(--sky-soft); }
textarea{ resize:vertical; width:100%; }
.tp-field{ display:flex; flex-direction:column; gap:5px; margin-bottom:10px; }
.tp-field>span{ font-size:11.5px; font-weight:600; color:var(--muted); }
.tp-field input,.tp-field select,.tp-field textarea{ width:100%; }
.tp-field-row{ display:flex; gap:10px; }
.tp-field-row .tp-field{ flex:1; }
.tp-check{ display:flex; align-items:center; gap:8px; font-size:13px; margin:6px 0; cursor:pointer; }
.tp-check input{ width:auto; }
.tp-primarybtn{ border:none; background:var(--sky); color:#fff; font-weight:700; font-size:13.5px; padding:10px 16px; border-radius:10px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; }
.tp-primarybtn:hover{ background:var(--sky-deep); }
.tp-dangerbtn{ border:1px solid #f1c9c4; background:#fdf3f2; color:#D9534F; font-weight:600; font-size:12.5px; padding:8px 12px; border-radius:9px; cursor:pointer; display:inline-flex; align-items:center; gap:5px; margin-top:6px; }
/* ---- sketch ---- */
.tp-sketch-tools{ display:flex; align-items:center; gap:6px; margin-bottom:8px; }
.tp-pen{ width:24px; height:24px; border-radius:50%; border:2px solid #fff; box-shadow:0 0 0 1px var(--line); cursor:pointer; display:flex; align-items:center; justify-content:center; color:#fff; }
.tp-pen.on{ box-shadow:0 0 0 2px var(--sky); }
.tp-pen.eraser{ background:#fff; color:var(--muted); }
.tp-sketch-spacer{ flex:1; }
.tp-canvas{ width:100%; height:auto; aspect-ratio:920/240; border:1px solid var(--line); border-radius:12px; background:repeating-linear-gradient(#fff,#fff 31px,#eef3f6 32px); touch-action:none; cursor:crosshair; display:block; }
/* ---- timetable ---- */
.tp-tt-card{ padding:12px; overflow-x:auto; }
.tp-tt{ display:grid; gap:4px; min-width:520px; }
.tp-tt-corner{ font-size:11px; color:var(--muted); display:flex; align-items:center; justify-content:center; }
.tp-tt-head{ text-align:center; font-weight:700; font-size:13px; padding:6px 0; background:var(--sky-soft); border-radius:8px; color:var(--sky-deep); }
.tp-tt-head.today{ background:var(--sky); color:#fff; }
.tp-tt-date{ display:block; font-size:11px; font-weight:600; opacity:.8; }
.tp-tt-period{ display:flex; flex-direction:column; align-items:center; justify-content:center; font-size:11px; color:var(--muted); }
.tp-tt-period b{ font-size:14px; color:var(--ink); }
.tp-tt-period.after b{ font-size:11px; }
.tp-tt-cell{ position:relative; background:#fafcfd; border:1px solid var(--line); border-radius:8px; min-height:56px; padding:5px 6px; cursor:pointer; display:flex; flex-direction:column; gap:2px; align-items:flex-start; justify-content:center; transition:.12s; text-align:left; }
.tp-tt-seq{ position:absolute; top:2px; right:4px; font-size:9px; font-weight:800; color:var(--sky-deep); opacity:.85; line-height:1; }
.tp-tt-cell:hover{ border-color:var(--sky); }
.tp-tt-cell.after{ min-height:40px; cursor:default; }
.tp-tt-sub{ font-weight:700; font-size:13px; }
.tp-tt-klass{ font-size:11px; color:var(--ink); font-weight:600; }
.tp-tt-room{ font-size:10px; color:var(--muted); }
.tp-tt-plus{ color:#c8d2da; font-size:18px; margin:auto; }
.tp-tt-club{ font-size:11px; color:var(--coral); font-weight:700; display:inline-flex; align-items:center; gap:3px; }
.tp-tt-off{ color:#c8d2da; font-size:11px; margin:auto; }
/* week list */
.tp-weeklist,.tp-splist,.tp-dutylist,.tp-mini-list,.tp-reclist{ list-style:none; margin:0; padding:0; }
.tp-weeklist li{ display:flex; align-items:center; gap:10px; padding:8px 0; border-bottom:1px dashed var(--line); font-size:13px; }
.tp-wl-date{ width:70px; font-weight:700; font-size:12px; flex-shrink:0; }
.tp-wl-date small{ color:var(--muted); font-weight:500; }
.tp-wl-title{ flex:1; }
.tp-wl-time{ color:var(--muted); font-size:12px; }
/* ---- calendar ---- */
.tp-cal{ display:grid; grid-template-columns:repeat(7,1fr); gap:4px; }
.tp-cal-wd{ text-align:center; font-size:11px; font-weight:700; padding:4px 0; color:var(--muted); }
.tp-cal-cell{ min-height:80px; border:1px solid var(--line); border-radius:9px; padding:4px; cursor:pointer; background:#fff; transition:.12s; overflow:hidden; }
.tp-cal-cell:hover{ border-color:var(--sky); background:#fbfdfe; }
.tp-cal-cell.empty{ background:transparent; border:none; cursor:default; }
.tp-cal-cell.today{ border-color:var(--sky); box-shadow:0 0 0 2px var(--sky-soft) inset; }
.tp-cal-cell.wd0{ background:#fef7f6; } .tp-cal-cell.wd6{ background:#f6fafc; }
.tp-cal-num{ font-size:12px; font-weight:700; display:flex; justify-content:space-between; align-items:center; }
.tp-cal-add{ opacity:0; border:none; background:var(--sky-soft); color:var(--sky-deep); border-radius:5px; width:18px; height:18px; display:flex; align-items:center; justify-content:center; cursor:pointer; }
.tp-cal-cell:hover .tp-cal-add{ opacity:1; }
.tp-cal-evs{ margin-top:3px; display:flex; flex-direction:column; gap:2px; }
.tp-cal-ev{ font-size:10px; color:#fff; padding:1px 5px; border-radius:4px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; font-weight:600; }
.tp-cal-ev.faint{ opacity:.5; }
.tp-cal-more{ font-size:9.5px; color:var(--muted); }
/* ---- aggregation ---- */
.tp-agg{ display:flex; flex-direction:column; gap:8px; }
.tp-agg-row{ display:grid; grid-template-columns:96px 44px 1fr 66px 62px; gap:8px; align-items:center; font-size:13px; }
.tp-agg-name{ font-weight:700; display:flex; align-items:center; gap:6px; }
.tp-dot{ width:10px; height:10px; border-radius:50%; flex-shrink:0; }
.tp-agg-week{ color:var(--muted); font-size:12px; } .tp-agg-week b{ color:var(--ink); font-size:14px; }
.tp-agg-bar{ height:8px; background:var(--line); border-radius:6px; overflow:hidden; }
.tp-agg-fill{ height:100%; border-radius:6px; transition:width .3s; }
.tp-agg-num{ text-align:right; font-weight:700; font-variant-numeric:tabular-nums; } .tp-agg-num small{ color:var(--muted); font-weight:500; }
.tp-agg-target{ width:100%; padding:5px 6px; font-size:12px; }
.tp-classgrid{ display:grid; grid-template-columns:repeat(auto-fill,minmax(110px,1fr)); gap:10px; }
.tp-classcard{ border:1px solid var(--line); border-radius:11px; padding:10px 12px; }
.tp-classcard b{ font-size:15px; color:var(--sky-deep); }
.tp-classcard div{ display:flex; align-items:baseline; gap:5px; font-size:17px; font-weight:800; margin-top:4px; font-variant-numeric:tabular-nums; }
.tp-classcard div span{ font-size:11px; color:var(--muted); font-weight:600; width:32px; }
.tp-classcard div small{ font-size:10px; color:var(--muted); font-weight:500; }
.tp-filterbar{ margin-bottom:8px; }
.tp-reclist li{ display:flex; align-items:center; gap:9px; padding:7px 0; border-bottom:1px dashed var(--line); font-size:13px; }
.tp-rec-date{ width:62px; flex-shrink:0; font-weight:700; font-size:12px; } .tp-rec-date small{ color:var(--muted); font-weight:500; }
.tp-rec-topic{ flex:1; display:flex; flex-direction:column; } .tp-rec-topic em{ color:#c0cad2; font-style:normal; }
.tp-rec-hw{ color:var(--muted); font-size:11px; }
.tp-rec-check{ color:var(--green); flex-shrink:0; }
.tp-reclist li.done .tp-rec-date{ color:var(--green); }
/* ---- club ---- */
.tp-clubhead{ display:flex; align-items:center; gap:10px; color:var(--coral); margin-bottom:14px; }
.tp-clubname{ font-size:18px; font-weight:800; color:var(--ink); border:none; border-bottom:2px solid var(--line); border-radius:0; padding:4px 2px; flex:1; }
.tp-clubname:focus{ box-shadow:none; border-bottom-color:var(--coral); }
.tp-clubsched{ display:flex; flex-direction:column; gap:6px; }
.tp-clubrow{ display:grid; grid-template-columns:56px 70px 12px 70px 1fr 1.2fr; gap:6px; align-items:center; padding:5px; border-radius:9px; }
.tp-clubrow.on{ background:#fcf5f1; }
.tp-clubday{ display:flex; align-items:center; gap:5px; font-weight:700; font-size:14px; }
.tp-clubday input{ width:auto; }
.tp-timeinput,.tp-placeinput,.tp-noteinput{ font-size:12px; padding:6px 8px; }
.tp-tilde{ text-align:center; color:var(--muted); }
.tp-splist li{ display:flex; align-items:center; gap:10px; padding:8px 0; border-bottom:1px dashed var(--line); }
.tp-sp-date{ width:74px; font-weight:700; font-size:12px; flex-shrink:0; } .tp-sp-date small{ color:var(--muted); font-weight:500; }
.tp-sp-title{ flex:1; font-size:13.5px; font-weight:600; } .tp-sp-title small{ display:block; color:var(--muted); font-weight:500; font-size:11.5px; }
.tp-sp-form{ display:flex; gap:6px; flex-wrap:wrap; align-items:center; }
.tp-sp-form input{ flex:1; min-width:80px; }
.tp-sp-form input[type=date]{ flex:0 0 auto; }
/* ---- roster ---- */
.tp-roster-tabs{ display:flex; gap:6px; flex-wrap:wrap; margin-bottom:14px; }
.tp-rtab{ border:1px solid var(--line); background:#fff; padding:7px 14px; border-radius:20px; font-size:13px; font-weight:700; cursor:pointer; color:var(--muted); position:relative; }
.tp-rtab.on{ background:var(--sky); color:#fff; border-color:var(--sky); }
.tp-hr{ font-size:9px; margin-left:5px; background:var(--coral); color:#fff; padding:1px 5px; border-radius:8px; }
.tp-roster{ display:flex; flex-direction:column; gap:5px; margin-bottom:12px; }
.tp-roster-headrow,.tp-roster-row{ display:grid; grid-template-columns:54px 1.2fr 1fr 1.6fr 30px; gap:8px; align-items:center; }
.tp-roster-headrow{ font-size:11px; color:var(--muted); font-weight:700; padding:0 2px; }
.tp-roster-row input{ width:100%; }
.tp-no{ text-align:center; }
/* ---- duties ---- */
.tp-dutylist li,.tp-routinelist li{ display:flex; align-items:center; gap:10px; padding:8px 0; border-bottom:1px dashed var(--line); }
.tp-duty-days{ width:70px; flex-shrink:0; font-weight:700; font-size:12px; color:var(--sky-deep); }
.tp-duty-title{ flex:1; font-size:13.5px; font-weight:600; } .tp-duty-title small{ color:var(--muted); font-weight:500; }
.tp-daypick,.tp-daychip{ }
.tp-daypick{ display:flex; gap:6px; margin-bottom:10px; }
.tp-daychip{ width:38px; height:38px; border:1px solid var(--line); background:#fff; border-radius:10px; font-weight:700; cursor:pointer; color:var(--muted); }
.tp-daychip.on{ background:var(--sky); color:#fff; border-color:var(--sky); }
.tp-routinelist li b{ width:48px; font-variant-numeric:tabular-nums; }
.tp-routinelist li span{ flex:1; font-size:13.5px; }
/* ---- modal ---- */
.tp-modal-back{ position:fixed; inset:0; background:rgba(36,50,60,.4); display:flex; align-items:center; justify-content:center; padding:20px; z-index:50; }
.tp-modal{ background:#fff; border-radius:18px; width:100%; max-width:440px; max-height:88vh; display:flex; flex-direction:column; box-shadow:0 20px 60px rgba(0,0,0,.25); }
.tp-modal.wide{ max-width:560px; }
.tp-modal-head{ display:flex; align-items:center; justify-content:space-between; padding:16px 18px; border-bottom:1px solid var(--line); }
.tp-modal-head h3{ margin:0; font-size:16px; font-weight:800; }
.tp-modal-body{ padding:16px 18px; overflow-y:auto; }
.tp-lesson-head{ font-size:15px; font-weight:700; margin-bottom:12px; display:flex; align-items:center; gap:8px; }
.tp-mini-list li{ display:flex; align-items:center; gap:8px; font-size:13px; padding:5px 0; }
.tp-mini-list li span:nth-child(2){ flex:1; }
.tp-period-edit{ display:flex; align-items:center; gap:6px; margin-bottom:6px; font-size:12px; color:var(--muted); }
.tp-plabel{ width:40px; text-align:center; }
.tp-subedit{ display:flex; flex-direction:column; gap:6px; }
.tp-subrow{ display:flex; align-items:center; gap:8px; }
.tp-subrow input[type=color]{ width:34px; height:34px; padding:2px; border-radius:8px; }
.tp-subrow input:not([type=color]){ flex:1; }
.tp-chips-edit{ display:flex; flex-wrap:wrap; gap:6px; margin-bottom:8px; }
.tp-editchip{ display:inline-flex; align-items:center; gap:5px; background:var(--sky-soft); color:var(--sky-deep); padding:5px 6px 5px 11px; border-radius:16px; font-size:12.5px; font-weight:700; }
.tp-editchip button{ border:none; background:none; color:var(--sky-deep); cursor:pointer; display:flex; padding:0; }
/* ---- responsive ---- */
/* ---- category flick ---- */
.tp-flick{ display:flex; gap:6px; background:#fff; border:1px solid var(--line); border-radius:12px; padding:5px; overflow-x:auto; }
.tp-flick-seg{ flex:1 0 auto; border:1px solid var(--line); background:#fff; border-radius:8px; padding:8px 12px; font-size:12.5px; font-weight:700; color:var(--muted); cursor:pointer; white-space:nowrap; transition:.15s; display:inline-flex; align-items:center; gap:6px; }
.tp-flick-seg.on{ color:#fff; }
.tp-flick-seg.off{ opacity:.65; }
.tp-flick-dot{ width:9px; height:9px; border-radius:50%; flex-shrink:0; box-shadow:0 0 0 1px rgba(0,0,0,.06); }
.tp-cat-tag{ margin-left:8px; font-size:11px; font-weight:600; color:var(--muted); background:var(--sky-soft); padding:2px 9px; border-radius:10px; }
.tp-cat-dot{ width:9px; height:9px; border-radius:50%; flex-shrink:0; }
/* ---- quarter (3 mini months) ---- */
.tp-quarter{ display:grid; grid-template-columns:repeat(3,1fr); gap:14px; }
.tp-mini-card{ padding:12px; }
.tp-mini-title{ text-align:center; font-weight:800; font-size:14px; margin-bottom:8px; color:var(--sky-deep); }
.tp-mini-grid{ display:grid; grid-template-columns:repeat(7,1fr); gap:2px; }
.tp-mini-wd{ text-align:center; font-size:9.5px; color:var(--muted); font-weight:700; }
.tp-mini-cell{ border:none; background:#fafcfd; border-radius:6px; min-height:42px; display:flex; flex-direction:column; align-items:stretch; justify-content:flex-start; cursor:pointer; padding:2px 3px; gap:1px; text-align:left; overflow:visible; }
.tp-mini-cell.empty{ background:transparent; cursor:default; }
.tp-mini-cell:hover:not(.empty){ background:var(--sky-soft); }
.tp-mini-cell.today{ background:var(--sky-soft); box-shadow:inset 0 0 0 2px var(--sky); }
.tp-mini-num{ font-size:10px; font-weight:700; }
.tp-mini-cell.wd0 .tp-mini-num{ color:#D9534F; } .tp-mini-cell.wd6 .tp-mini-num{ color:var(--sky-deep); }
.tp-mini-dots{ display:flex; gap:2px; height:5px; }
.tp-mini-dots span{ width:5px; height:5px; border-radius:50%; }
/* mini (3か月) の予定テキスト */
.tp-cal-evs.mini{ margin-top:1px; gap:1px; }
.tp-cal-evs.mini .tp-cal-ev{ font-size:8px; padding:0 3px; border-radius:3px; }
.tp-cal-evs.mini .tp-cal-club{ padding:1px 3px; border-left-width:2px; }
.tp-cal-evs.mini .tp-cal-cline{ font-size:7.5px; line-height:1.2; white-space:normal; overflow:visible; text-overflow:clip; word-break:break-word; overflow-wrap:anywhere; }
.tp-cal-evs.mini .tp-cal-more{ font-size:7.5px; }
/* 時間割プレビュー */
.tp-ttpv{ display:grid; gap:3px; margin-bottom:8px; }
.tp-ttpv-corner,.tp-ttpv-head,.tp-ttpv-p{ font-size:11px; font-weight:700; text-align:center; color:var(--sky-deep); display:flex; align-items:center; justify-content:center; }
.tp-ttpv-head{ background:var(--sky-soft); border-radius:6px; padding:4px 0; }
.tp-ttpv-cell{ min-height:34px; border:1px solid var(--line); border-radius:6px; padding:3px; display:flex; flex-direction:column; align-items:center; justify-content:center; text-align:center; }
.tp-ttpv-cell b{ font-size:11px; }
.tp-ttpv-cell span{ font-size:9px; color:var(--muted); }
/* ---- year ---- */
.tp-year{ display:grid; grid-template-columns:repeat(3,1fr); gap:12px; }
.tp-year-month{ padding:12px; }
.tp-year-head{ width:100%; border:none; background:none; text-align:left; font-weight:800; font-size:15px; color:var(--sky-deep); display:flex; align-items:center; justify-content:space-between; cursor:pointer; padding:0 0 6px; border-bottom:1px solid var(--line); margin-bottom:6px; }
.tp-year-list{ list-style:none; margin:0; padding:0; max-height:170px; overflow-y:auto; }
.tp-year-list li{ display:flex; align-items:center; gap:6px; padding:3px 0; font-size:11.5px; cursor:pointer; border-radius:5px; }
.tp-year-list li:hover{ background:var(--sky-soft); }
.tp-year-day{ width:20px; text-align:right; font-weight:700; font-variant-numeric:tabular-nums; flex-shrink:0; }
.tp-year-day.wd0{ color:#D9534F; } .tp-year-day.wd6{ color:var(--sky-deep); }
.tp-year-title{ flex:1; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.tp-year-empty{ color:#c8d2da; font-size:12px; }
/* ---- weekly manual editor ---- */
.tp-weekedit{ display:grid; grid-template-columns:repeat(auto-fill,minmax(150px,1fr)); gap:8px; }
.tp-weekedit-row{ display:flex; align-items:center; gap:6px; font-size:13px; }
.tp-weekedit-row span{ width:44px; font-weight:700; }
.tp-weekedit-row input{ width:56px; text-align:center; }
.tp-weekedit-row small{ color:var(--muted); font-size:11px; }
/* ---- import ---- */
.tp-toolbar{ display:flex; justify-content:flex-end; gap:8px; flex-wrap:wrap; }
.tp-daymemo{ width:100%; }
.tp-uploadbox{ border:2px dashed var(--line); border-radius:14px; min-height:150px; display:flex; align-items:center; justify-content:center; cursor:pointer; background:#fafcfd; overflow:hidden; margin-bottom:12px; }
.tp-uploadbox:hover{ border-color:var(--sky); background:var(--sky-soft); }
.tp-uploadbox img{ max-width:100%; max-height:260px; border-radius:8px; }
.tp-upload-empty{ display:flex; flex-direction:column; align-items:center; gap:6px; color:var(--muted); padding:24px; text-align:center; }
.tp-upload-empty span{ font-weight:700; font-size:14px; color:var(--ink); }
.tp-upload-empty small{ font-size:11.5px; }
.tp-pdfmark{ font-weight:700; color:var(--sky-deep); padding:30px; }
.tp-full{ width:100%; justify-content:center; }
.tp-error{ background:#fdf3f2; color:#c0392b; border:1px solid #f1c9c4; border-radius:9px; padding:9px 12px; font-size:12.5px; margin-bottom:10px; }
.tp-warn{ background:#fff8e9; color:#a97b1e; border:1px solid #f0e0b8; border-radius:9px; padding:8px 12px; font-size:12px; margin-bottom:12px; }
.tp-spin{ animation:tp-rot 1s linear infinite; }
@keyframes tp-rot{ to{ transform:rotate(360deg); } }
.tp-result-head{ font-weight:800; font-size:14px; margin-bottom:8px; }
.tp-result-sub{ font-size:11.5px; font-weight:700; color:var(--sky-deep); margin:10px 0 4px; }
.tp-result-list{ list-style:none; margin:0; padding:0; }
.tp-result-list li{ display:flex; align-items:center; gap:8px; padding:5px 0; font-size:13px; border-bottom:1px dashed var(--line); }
.tp-result-list li input[type=checkbox]{ width:auto; }
.tp-res-date{ font-weight:700; font-size:12px; width:44px; flex-shrink:0; }
.tp-res-title{ flex:1; }
.tp-result-list li small{ color:var(--muted); }
.tp-modal-actions{ display:flex; justify-content:space-between; align-items:center; gap:10px; margin-top:14px; }
.tp-modal-actions .tp-dangerbtn{ margin-top:0; }
.tp-modal-actions > button{ min-height:40px; }
/* ---- classes sub-tabs & plan ---- */
.tp-subtabs{ display:flex; gap:6px; background:#fff; border:1px solid var(--line); border-radius:12px; padding:5px; overflow-x:auto; }
.tp-subtab{ flex:1 0 auto; border:none; background:none; border-radius:8px; padding:9px 14px; font-size:13px; font-weight:700; color:var(--muted); cursor:pointer; white-space:nowrap; }
.tp-subtab.on{ background:var(--sky); color:#fff; }
.tp-unitlist{ display:flex; flex-direction:column; gap:6px; margin:10px 0; }
.tp-unitrow{ display:grid; grid-template-columns:96px 1.2fr auto 1.4fr 28px; gap:6px; align-items:center; }
.tp-unitrow input{ font-size:12px; padding:6px 8px; }
.tp-u-prog{ font-weight:700; }
.tp-u-p{ display:flex; align-items:center; gap:2px; font-size:11px; color:var(--muted); white-space:nowrap; }
.tp-u-p input{ width:52px; text-align:center; padding:5px 4px; }
.tp-testlist{ list-style:none; margin:0; padding:0; }
.tp-testlist li{ display:flex; align-items:center; gap:8px; padding:7px 0; border-bottom:1px dashed var(--line); font-size:13px; }
.tp-test-date{ font-weight:700; width:44px; font-size:12px; }
.tp-test-name{ flex:1; font-weight:600; }
.tp-test-range{ color:var(--muted); font-size:12px; }
.tp-test-form{ display:flex; gap:6px; flex-wrap:wrap; align-items:center; }
.tp-test-form input:first-child{ flex:1; min-width:130px; }
/* year plan */
.tp-plan-head{ display:flex; justify-content:space-between; align-items:center; gap:10px; flex-wrap:wrap; margin-bottom:6px; }
.tp-plan-controls{ display:flex; gap:10px; align-items:center; flex-wrap:wrap; font-size:12px; color:var(--muted); }
.tp-plan-controls select,.tp-plan-controls input{ font-size:12px; padding:5px 7px; margin-left:4px; }
.tp-plan-wk{ background:var(--sky-soft); color:var(--sky-deep); font-weight:700; padding:3px 9px; border-radius:10px; }
.tp-plan-block{ margin-top:14px; }
.tp-plan-test{ display:flex; align-items:center; gap:8px; flex-wrap:wrap; padding:8px 10px; background:#fafcfd; border-left:4px solid; border-radius:8px; margin-bottom:6px; }
.tp-plan-test b{ font-size:14px; }
.tp-plan-testdate{ font-weight:700; color:var(--ink); }
.tp-plan-scope{ font-size:11.5px; color:var(--muted); width:100%; }
.tp-plan-weeks{ display:flex; flex-direction:column; gap:3px; }
.tp-plan-week{ display:grid; grid-template-columns:64px 1fr; gap:8px; align-items:start; padding:5px 4px; border-bottom:1px dashed var(--line); }
.tp-plan-wdate{ font-size:11.5px; font-weight:700; color:var(--sky-deep); padding-top:3px; }
.tp-plan-tasks{ display:flex; flex-wrap:wrap; gap:5px; }
.tp-plan-task{ background:var(--sky-soft); border-radius:7px; padding:3px 8px; font-size:11.5px; display:inline-flex; align-items:center; gap:4px; }
.tp-plan-task b{ color:var(--sky-deep); }
.tp-plan-task small{ color:var(--muted); }
.tp-plan-none{ font-size:11.5px; color:#b7c2cb; }
.tp-plan-week.deadline{ background:#fff7ea; border-radius:7px; }
.tp-plan-dl{ font-size:11.5px; font-weight:700; color:var(--amber); }
.tp-plan-week.testrow{ background:#fdf3f2; border-radius:7px; }
.tp-plan-testtag{ font-size:12px; font-weight:700; color:#c0392b; }
/* ---- calendar club 4-line block ---- */
.tp-cal-cell{ min-height:80px; height:auto; }
.tp-cal-club{ border-left:3px solid #E8845B; background:#fcf5f1; border-radius:5px; padding:2px 4px; display:flex; flex-direction:column; gap:0; margin-bottom:2px; }
.tp-cal-club.weekly{ opacity:.72; }
.tp-cal-cline{ font-size:9.5px; line-height:1.25; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
.tp-cal-cline.l0{ font-weight:700; color:#c0562e; }
.tp-cal-cline.l1,.tp-cal-cline.l2{ color:#6b7a88; }
.tp-cal-cline.l3{ color:#95a2ad; font-style:italic; }
/* ---- club: seg / bulk / day list ---- */
.tp-seg{ display:flex; gap:6px; }
.tp-seg-btn{ flex:1; border:1px solid var(--line); background:#fff; border-radius:9px; padding:9px 6px; font-size:13px; font-weight:700; color:var(--muted); cursor:pointer; }
.tp-seg-btn.on{ color:#fff; }
.tp-bulkbar{ display:flex; align-items:center; gap:8px; flex-wrap:wrap; background:var(--sky-soft); border-radius:10px; padding:8px 10px; margin-bottom:10px; }
.tp-check.sm{ font-size:12px; margin:0; }
.tp-bulk-wd{ display:flex; align-items:center; gap:3px; font-size:11px; color:var(--muted); }
.tp-wdpick{ border:1px solid #cfe0ea; background:#fff; border-radius:6px; width:22px; height:24px; font-size:11px; font-weight:700; color:var(--sky-deep); cursor:pointer; }
.tp-wdpick:hover{ background:var(--sky); color:#fff; }
.tp-bulk-spacer{ flex:1; }
.tp-primarybtn.sm{ padding:7px 12px; font-size:12.5px; }
.tp-dayschedule{ display:flex; flex-direction:column; gap:3px; }
.tp-dayrow{ display:flex; align-items:center; gap:8px; padding:4px 6px; border-radius:9px; border:1px solid transparent; }
.tp-dayrow.checked{ background:var(--sky-soft); border-color:#cfe0ea; }
.tp-dayrow.sun .tp-dayrow-date{ color:#D9534F; } .tp-dayrow.sat .tp-dayrow-date{ color:var(--sky-deep); }
.tp-dayrow>input{ width:auto; flex-shrink:0; }
.tp-dayrow-main{ flex:1; display:flex; align-items:center; gap:10px; background:none; border:none; text-align:left; cursor:pointer; padding:4px 2px; border-bottom:1px dashed var(--line); min-width:0; }
.tp-dayrow-date{ width:34px; flex-shrink:0; font-weight:800; font-size:15px; display:flex; flex-direction:column; align-items:center; line-height:1; }
.tp-dayrow-date small{ font-size:9px; font-weight:600; color:var(--muted); }
.tp-dayrow-info{ flex:1; display:flex; flex-direction:column; gap:1px; min-width:0; }
.tp-dayrow-content{ font-weight:700; font-size:13px; color:#c0562e; }
.tp-dayrow-content.off{ color:var(--muted); }
.tp-dayrow-content.match{ color:#D9534F; }
.tp-weeklytag{ font-size:9px; font-style:normal; background:#e7edf1; color:#8894a0; padding:1px 5px; border-radius:6px; margin-left:6px; }
.tp-dayrow-sub{ font-size:11.5px; color:var(--muted); }
.tp-dayrow-note{ font-size:11px; color:#95a2ad; }
.tp-dayrow-empty{ flex:1; font-size:12.5px; color:#b7c2cb; }
.tp-dayrow-pen{ color:#c8d2da; flex-shrink:0; }
.tp-collapse{ width:100%; border:none; background:none; display:flex; justify-content:space-between; align-items:center; font-size:13px; font-weight:700; color:var(--sky-deep); cursor:pointer; padding:2px 0; }
.tp-rot90{ transform:rotate(90deg); }
.tp-print-clubitem{ display:inline-flex; align-items:center; gap:5px; font-size:12px; }
/* ---- print (A4) ---- */
.tp-print-root{ position:fixed; inset:0; z-index:200; background:#5a6570; display:flex; flex-direction:column; }
/* 3か月カレンダー印刷（A4横・1枚） */
.tp-print-sheet.landscape{ width:297mm; min-height:210mm; }
.tp-print-sheet.cal{ padding:9mm 9mm 6mm; display:flex; flex-direction:column; }
.tp-print-sheet.cal .tp-print-header{ padding-bottom:4px; }
.tp-print-sheet.cal .tp-print-header h1{ font-size:18px; }
.tp-pcal{ display:grid; grid-template-columns:repeat(3,1fr); gap:6mm; margin-top:6px; flex:1; }
.tp-pcal-month{ page-break-inside:avoid; break-inside:avoid; display:flex; flex-direction:column; }
.tp-pcal-title{ font-size:12px; font-weight:800; text-align:center; margin-bottom:3px; color:#2C7CA6; }
.tp-pcal-grid{ display:grid; grid-template-columns:repeat(7,1fr); gap:0.5px; background:#dfe4e8; border:0.5px solid #dfe4e8; }
.tp-pcal-wd{ background:#eef3f6; text-align:center; font-size:7px; font-weight:700; padding:1.5px 0; color:#556; }
.tp-pcal-wd.wd0{ color:#D9534F; } .tp-pcal-wd.wd6{ color:#2C7CA6; }
.tp-pcal-cell{ background:#fff; min-height:12mm; max-height:16mm; overflow:hidden; padding:1px 2px 2px; }
.tp-pcal-cell.empty{ background:#f7f8f9; min-height:12mm; }
.tp-pcal-num{ font-size:8px; font-weight:700; line-height:1.1; }
.tp-pcal-cell.wd0 .tp-pcal-num{ color:#D9534F; } .tp-pcal-cell.wd6 .tp-pcal-num{ color:#2C7CA6; }
.tp-pcal-items{ display:flex; flex-direction:column; gap:0.5px; margin-top:0.5px; }
.tp-pcal-ev{ font-size:6px; line-height:1.2; display:flex; align-items:center; gap:2px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
.tp-pcal-dot{ width:4px; height:4px; border-radius:50%; flex-shrink:0; }
.tp-pcal-club{ font-size:6px; line-height:1.25; border-left:1.5px solid #E8845B; padding-left:2px; color:#b0542e; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
.tp-pcal-more{ font-size:5.5px; color:#8894a0; }
.tp-print-legend.inline{ display:inline-flex; gap:12px; margin-left:14px; vertical-align:middle; }
@media print {
  .tp-print-sheet.landscape{ width:auto !important; min-height:auto !important; }
  .tp-print-sheet.cal{ padding:0 !important; }
  .tp-pcal-cell{ max-height:none; }
}
.tp-print-bar{ display:flex; align-items:center; gap:8px; padding:10px 14px; background:#fff; border-bottom:1px solid var(--line); }
.tp-print-barinfo{ font-weight:800; font-size:13px; color:var(--ink); }
.tp-print-spacer{ flex:1; }
.tp-print-scroll{ flex:1; overflow:auto; padding:20px; display:flex; justify-content:center; }
.tp-print-sheet{ background:#fff; width:210mm; min-height:297mm; padding:16mm 15mm; box-shadow:0 6px 24px rgba(0,0,0,.3); color:#1a1a1a; }
.tp-print-header{ display:flex; justify-content:space-between; align-items:flex-end; border-bottom:2px solid #333; padding-bottom:6px; }
.tp-print-header h1{ margin:0; font-size:22px; font-weight:800; letter-spacing:1px; }
.tp-print-meta{ font-size:13px; }
.tp-print-legend{ display:flex; gap:16px; margin:8px 0 12px; font-size:11px; }
.tp-print-legend span{ display:inline-flex; align-items:center; gap:5px; }
.tp-print-table{ width:100%; border-collapse:collapse; }
.tp-print-table td{ border-bottom:1px solid #ddd; padding:6px 4px; vertical-align:top; }
.tp-print-date{ width:70px; font-weight:700; font-size:12.5px; white-space:nowrap; }
.tp-print-date small{ font-weight:500; color:#666; }
.tp-print-date.wd0{ color:#c0392b; } .tp-print-date.wd6{ color:#2C7CA6; }
.tp-print-items{ display:flex; flex-wrap:wrap; gap:5px 12px; }
.tp-print-item{ display:inline-flex; align-items:center; gap:5px; font-size:12.5px; }
.tp-print-item b{ color:#555; font-weight:700; }
.tp-print-empty{ color:#888; padding:20px 0; }
.tp-print-foot{ margin-top:16px; text-align:right; font-size:10px; color:#999; }
@media print {
  html, body { background:#fff !important; }
  .tp-tabs, .tp-main, .tp-modal-back, .tp-print-bar { display:none !important; }
  .tp-app { display:block !important; height:auto !important; overflow:visible !important; }
  .tp-print-root { position:static !important; background:#fff !important; }
  .tp-print-scroll { overflow:visible !important; padding:0 !important; display:block !important; }
  .tp-print-sheet { width:auto !important; min-height:auto !important; box-shadow:none !important; padding:0 !important; }
  @page { size:A4; margin:14mm; }
}

@media (max-width:820px){
  .tp-app{ flex-direction:column; }
  .tp-tabs{ order:2; width:100%; flex-direction:row; flex-wrap:wrap; justify-content:center; border-right:none; border-top:1px solid var(--line); padding:5px 4px; overflow:visible; flex-shrink:0; }
  .tp-quarter{ grid-template-columns:1fr; }
  .tp-year{ grid-template-columns:1fr 1fr; }
  .tp-flick-seg{ flex:0 0 auto; padding:7px 10px; font-size:11.5px; }
  .tp-brand{ display:none; }
  .tp-tab{ flex:1 1 18%; min-width:52px; max-width:80px; padding:6px 2px; font-size:10px; gap:2px; }
  .tp-tab svg{ width:18px; height:18px; }
  .tp-tab.settings{ margin-top:0; }
  .tp-main{ order:1; flex:1; min-height:0; }
  .tp-scroll{ padding:4px 14px 24px; }
  .tp-topbar{ padding:12px 16px 8px; }
  .tp-appname{ font-size:16px; }
  .tp-apptag{ letter-spacing:2px; }
  .tp-section-name{ font-size:13px; }
  .tp-topbar-sub{ display:none; }
  .tp-set-mobile{ display:flex; border:1px solid var(--line); background:#fff; border-radius:9px; width:36px; height:36px; align-items:center; justify-content:center; color:var(--ink); cursor:pointer; }
  .tp-tab.settings{ display:none; }
  .tp-cardrow{ grid-template-columns:1fr; }
  .tp-print-sheet{ width:auto; padding:8mm; }
  .tp-agg-row{ grid-template-columns:80px 38px 1fr 54px; }
  .tp-agg-target{ display:none; }
  .tp-clubrow{ grid-template-columns:50px 60px 10px 60px; grid-auto-flow:row; }
  .tp-clubrow .tp-placeinput,.tp-clubrow .tp-noteinput{ grid-column:1 / -1; }
  .tp-roster-headrow{ display:none; }
  .tp-roster-row{ grid-template-columns:44px 1fr 26px; grid-auto-flow:row; }
  .tp-roster-row input:nth-child(3),.tp-roster-row input:nth-child(4){ grid-column:2 / -1; }
}
`;
