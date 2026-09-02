import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import { doc, onSnapshot, setDoc } from 'firebase/firestore';
import { Upload, Download, FileSpreadsheet, AlertTriangle, CheckCircle2, EyeOff } from 'lucide-react';
import { db, DASHBOARD_DOC_PATH } from './firebase';
import { SEED_RECORDS, SEED_CANDIDATES } from './seedData';

const MONTHS_ID = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
const RECOMMENDED_LABEL = 'Direkomendasikan';
const CONVERSION_RATIO = 3; // asumsi historis: 1 penempatan berhasil dari setiap 3 kandidat direkomendasikan
const UNFILLED_COLUMNS = [
  { key: 'jabatan', label: 'Posisi' },
  { key: 'unit', label: 'Unit' },
  { key: 'area', label: 'Area' },
  { key: 'ket', label: 'Alasan' },
  { key: 'tmt', label: 'Tgl. efektif' },
  { key: 'status', label: 'Status' },
];

function isFilled(record) {
  return !!(record.kandidat && String(record.kandidat).trim() !== '');
}

function isRecommended(hasil) {
  return hasil != null && String(hasil).trim() === RECOMMENDED_LABEL;
}

function normalizeKey(k) {
  return String(k).trim().toLowerCase();
}

function getField(row, ...candidates) {
  const keys = Object.keys(row);
  for (const cand of candidates) {
    const target = normalizeKey(cand);
    const found = keys.find((k) => normalizeKey(k) === target);
    if (found !== undefined) return row[found];
  }
  return null;
}

function toDateStr(v) {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return null;
    const y = v.getFullYear();
    const m = String(v.getMonth() + 1).padStart(2, '0');
    const d = String(v.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  if (typeof v === 'number') {
    try {
      const parsed = XLSX.SSF.parse_date_code(v);
      if (!parsed) return null;
      const m = String(parsed.m).padStart(2, '0');
      const d = String(parsed.d).padStart(2, '0');
      return `${parsed.y}-${m}-${d}`;
    } catch (e) {
      return null;
    }
  }
  if (typeof v === 'string') {
    const d = new Date(v);
    if (!isNaN(d.getTime())) {
      const y = d.getFullYear();
      const m = String(d.getMonth() + 1).padStart(2, '0');
      const dd = String(d.getDate()).padStart(2, '0');
      return `${y}-${m}-${dd}`;
    }
    return null;
  }
  return null;
}

function formatDateID(iso) {
  if (!iso) return '\u2014';
  const parts = iso.split('-').map(Number);
  const y = parts[0], m = parts[1], d = parts[2];
  if (!y || !m || !d) return '\u2014';
  return `${d} ${MONTHS_ID[m - 1]} ${y}`;
}

function formatDateTimeID(iso) {
  if (!iso) return null;
  const d = new Date(iso);
  if (isNaN(d.getTime())) return null;
  const day = d.getDate();
  const month = MONTHS_ID[d.getMonth()];
  const year = d.getFullYear();
  const hh = String(d.getHours()).padStart(2, '0');
  const mm = String(d.getMinutes()).padStart(2, '0');
  return `${day} ${month} ${year}, ${hh}:${mm}`;
}

function daysStatus(iso) {
  if (!iso) return { label: 'Tanggal tidak diketahui', tone: 'muted', diffDays: null };
  const target = new Date(iso + 'T00:00:00');
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const diffDays = Math.round((today - target) / 86400000);
  if (diffDays > 0) return { label: `Terlambat ${diffDays} hari`, tone: 'danger', diffDays };
  if (diffDays === 0) return { label: 'Jatuh tempo hari ini', tone: 'danger', diffDays };
  return { label: `H-${-diffDays}`, tone: 'upcoming', diffDays };
}

function compareUnfilledRows(a, b, key) {
  if (key === 'tmt') {
    if (!a.tmt && !b.tmt) return 0;
    if (!a.tmt) return 1;
    if (!b.tmt) return -1;
    return new Date(a.tmt) - new Date(b.tmt);
  }
  if (key === 'status') {
    const da = a.status && a.status.diffDays != null ? a.status.diffDays : -Infinity;
    const db_ = b.status && b.status.diffDays != null ? b.status.diffDays : -Infinity;
    return da - db_;
  }
  const va = (a[key] || '').toString().toLowerCase();
  const vb = (b[key] || '').toString().toLowerCase();
  if (va < vb) return -1;
  if (va > vb) return 1;
  return 0;
}

function parseHeadcountSheet(ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  const dataRows = rows.slice(1);
  const parsed = [];
  for (const r of dataRows) {
    if (!Array.isArray(r)) continue;
    const isEmpty = r.every((v) => v === null || v === undefined || String(v).trim() === '');
    if (isEmpty) continue;
    parsed.push({
      no: r[0] ?? null,
      nip: r[1] != null ? String(r[1]) : null,
      nama: r[2] ?? null,
      jabatan: r[3] ?? null,
      unit: r[4] ?? null,
      area: r[5] ?? null,
      ket: r[6] ?? null,
      tmt: toDateStr(r[7]),
      kandidat: r[8] ?? null,
      jabatanFill: r[9] ?? null,
      unitFill: r[10] ?? null,
      tmtFill: toDateStr(r[11]),
      ketFill: r[12] ?? null,
    });
  }
  return parsed;
}

function parseCandidateSheet(ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null });
  const parsed = [];
  for (const row of rows) {
    const nama = getField(row, 'Nama');
    const nip = getField(row, 'NIP');
    const jabatan = getField(row, 'Jabatan');
    const unit = getField(row, 'Unit');
    const area = getField(row, 'Area');
    const proses = getField(row, 'Proses');
    const hasil = getField(row, 'Hasil');
    const kandidat = getField(row, 'Kandidat');
    const isEmpty = !nama && !nip && !jabatan && !area && !hasil;
    if (isEmpty) continue;
    parsed.push({
      nip: nip != null ? String(nip) : null,
      nama: nama ?? null,
      jabatan: jabatan ?? null,
      unit: unit ?? null,
      area: area ?? null,
      proses: proses ?? null,
      hasil: hasil ?? null,
      kandidat: kandidat ?? null,
    });
  }
  return parsed;
}

function parseWorkbook(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
  const kandidatSheetName = wb.SheetNames.find((n) => normalizeKey(n).includes('kandidat'));
  const headcountSheetName = wb.SheetNames.find((n) => n !== kandidatSheetName) || wb.SheetNames[0];

  const headcountRecords = headcountSheetName ? parseHeadcountSheet(wb.Sheets[headcountSheetName]) : [];
  const candidateRecords = kandidatSheetName ? parseCandidateSheet(wb.Sheets[kandidatSheetName]) : [];

  return { headcountRecords, candidateRecords, hasCandidateSheet: !!kandidatSheetName };
}

function ProgressRing({ pct, size = 136, stroke = 11, color = '#0B6E4F', label = 'terisi' }) {
  const radius = (size - stroke) / 2;
  const circumference = 2 * Math.PI * radius;
  const [offset, setOffset] = useState(circumference);

  useEffect(() => {
    const id = requestAnimationFrame(() => {
      setOffset(circumference * (1 - Math.max(0, Math.min(100, pct)) / 100));
    });
    return () => cancelAnimationFrame(id);
  }, [pct, circumference]);

  return (
    <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`} className="hc-ring" role="img" aria-label={`${pct} persen ${label}`}>
      <circle cx={size / 2} cy={size / 2} r={radius} fill="none" stroke="#DCE3DC" strokeWidth={stroke} />
      <circle
        cx={size / 2} cy={size / 2} r={radius} fill="none"
        stroke={color} strokeWidth={stroke} strokeLinecap="round"
        strokeDasharray={circumference}
        strokeDashoffset={offset}
        transform={`rotate(-90 ${size / 2} ${size / 2})`}
        className="hc-ring-progress"
      />
      <text x="50%" y="47%" textAnchor="middle" dominantBaseline="middle" className="hc-ring-pct" style={{ fontSize: Math.round(size * 0.19) }}>{pct}%</text>
      <text x="50%" y="66%" textAnchor="middle" dominantBaseline="middle" className="hc-ring-label" style={{ fontSize: Math.round(size * 0.085) }}>{label}</text>
    </svg>
  );
}

export default function App() {
  const [records, setRecords] = useState(SEED_RECORDS);
  const [candidates, setCandidates] = useState(SEED_CANDIDATES);
  const [hasCandidateSheet, setHasCandidateSheet] = useState(true);
  const [fileName, setFileName] = useState('(contoh bawaan)');
  const [lastUpdated, setLastUpdated] = useState(null);
  const [isConnected, setIsConnected] = useState(false);
  const [toast, setToast] = useState(null);
  const [error, setError] = useState(null);
  const [isUploading, setIsUploading] = useState(false);
  const [isExporting, setIsExporting] = useState(false);
  const [searchText, setSearchText] = useState('');
  const [filterArea, setFilterArea] = useState('');
  const [filterAlasan, setFilterAlasan] = useState('');
  const [sortKey, setSortKey] = useState('tmt');
  const [sortDir, setSortDir] = useState('asc');
  const [visibleColumns, setVisibleColumns] = useState({
    jabatan: true, unit: true, area: true, ket: true, tmt: true, status: true,
  });
  const [showColumnMenu, setShowColumnMenu] = useState(false);
  const columnMenuRef = useRef(null);
  const fileInputRef = useRef(null);
  const dashboardRef = useRef(null);

  // Dengerin perubahan data di Firestore secara real-time. Begitu ada yang
  // upload file baru (dari device mana pun), semua yang lagi buka web ini
  // otomatis ke-update tanpa perlu refresh.
  useEffect(() => {
    const ref = doc(db, ...DASHBOARD_DOC_PATH);
    const unsubscribe = onSnapshot(
      ref,
      (snap) => {
        setIsConnected(true);
        if (!snap.exists()) return; // belum ada yang upload, tetap pakai data contoh
        const data = snap.data();
        if (Array.isArray(data.records)) setRecords(data.records);
        if (Array.isArray(data.candidates)) setCandidates(data.candidates);
        if (data.fileName) setFileName(data.fileName);
        if (data.updatedAt) setLastUpdated(data.updatedAt);
        if (typeof data.hasCandidateSheet === 'boolean') setHasCandidateSheet(data.hasCandidateSheet);
      },
      (err) => {
        console.error(err);
        setError('Gagal konek ke Firestore. Cek koneksi internet, atau config Firebase di src/firebase.js belum bener.');
      }
    );
    return () => unsubscribe();
  }, []);

  useEffect(() => {
    if (!toast) return;
    const t = setTimeout(() => setToast(null), 4000);
    return () => clearTimeout(t);
  }, [toast]);

  useEffect(() => {
    if (!error) return;
    const t = setTimeout(() => setError(null), 7000);
    return () => clearTimeout(t);
  }, [error]);

  async function handleFile(file) {
    setError(null);
    setIsUploading(true);
    try {
      const buf = await file.arrayBuffer();
      const { headcountRecords, candidateRecords, hasCandidateSheet: foundCandidateSheet } = parseWorkbook(buf);
      if (headcountRecords.length === 0) {
        setError('File tidak berisi data Headcount yang bisa dibaca. Pastikan formatnya sama seperti file sebelumnya.');
        return;
      }
      const updatedAt = new Date().toISOString();
      await setDoc(doc(db, ...DASHBOARD_DOC_PATH), {
        records: headcountRecords,
        candidates: candidateRecords,
        fileName: file.name,
        updatedAt,
        hasCandidateSheet: foundCandidateSheet,
      });
      // Tidak perlu setRecords manual di sini -- listener onSnapshot di atas
      // bakal otomatis nangkep perubahan ini dan update tampilan untuk semua orang.
      setSearchText('');
      setFilterArea('');
      setFilterAlasan('');
      setSortKey('tmt');
      setSortDir('asc');
      setToast(
        foundCandidateSheet
          ? `Data diperbarui dari ${file.name} untuk semua orang.`
          : `Data diperbarui dari ${file.name}. Sheet Kandidat tidak ditemukan, jumlah kandidat di-set 0.`
      );
    } catch (e) {
      console.error(e);
      setError('Gagal menyimpan ke Firestore. Cek koneksi internet, atau Firestore rules belum mengizinkan tulis.');
    } finally {
      setIsUploading(false);
    }
  }

  function onInputChange(e) {
    const file = e.target.files && e.target.files[0];
    if (file) handleFile(file);
    e.target.value = '';
  }

  function toggleSort(key) {
    if (sortKey === key) {
      setSortDir((d) => (d === 'asc' ? 'desc' : 'asc'));
    } else {
      setSortKey(key);
      setSortDir('asc');
    }
  }

  function sortIndicator(key) {
    if (sortKey !== key) return null;
    return sortDir === 'asc' ? ' \u25B2' : ' \u25BC';
  }

  function toggleColumn(key) {
    setVisibleColumns((prev) => {
      const next = { ...prev, [key]: !prev[key] };
      const anyVisible = Object.values(next).some(Boolean);
      return anyVisible ? next : prev; // jangan biarin semua kolom ke-hide
    });
  }

  useEffect(() => {
    if (!showColumnMenu) return;
    function handleClickOutside(e) {
      if (columnMenuRef.current && !columnMenuRef.current.contains(e.target)) {
        setShowColumnMenu(false);
      }
    }
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [showColumnMenu]);

  async function handleDownloadImage() {
    const node = dashboardRef.current;
    if (!node || isExporting) return;
    setError(null);
    setIsExporting(true);
    try {
      const clone = node.cloneNode(true);
      clone.querySelectorAll('.hc-header-actions, input[type="file"], .hc-toast, .hc-error, .hc-col-toggle').forEach((el) => el.remove());
      clone.style.margin = '0';

      let maxTableWidth = 0;
      clone.querySelectorAll('.hc-table-scroll').forEach((wrap) => {
        wrap.style.overflow = 'visible';
        const table = wrap.querySelector('table');
        if (table) maxTableWidth = Math.max(maxTableWidth, table.scrollWidth);
      });

      const rect = node.getBoundingClientRect();
      const width = Math.max(Math.ceil(rect.width), maxTableWidth + 120);
      const height = Math.ceil(rect.height);
      clone.style.width = `${width}px`;

      const htmlString = new XMLSerializer().serializeToString(clone);
      const svgString =
        `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}">` +
        `<foreignObject width="100%" height="100%">` +
        `<div xmlns="http://www.w3.org/1999/xhtml">${htmlString}</div>` +
        `</foreignObject></svg>`;

      const svgBlob = new Blob([svgString], { type: 'image/svg+xml;charset=utf-8' });
      const svgUrl = URL.createObjectURL(svgBlob);

      const img = new Image();
      await new Promise((resolve, reject) => {
        img.onload = resolve;
        img.onerror = () => reject(new Error('Gagal memuat pratinjau dashboard.'));
        img.src = svgUrl;
      });

      const scale = 2;
      const canvas = document.createElement('canvas');
      canvas.width = width * scale;
      canvas.height = height * scale;
      const ctx = canvas.getContext('2d');
      ctx.scale(scale, scale);
      ctx.fillStyle = '#EEF1EC';
      ctx.fillRect(0, 0, width, height);
      ctx.drawImage(img, 0, 0, width, height);
      URL.revokeObjectURL(svgUrl);

      const pngUrl = canvas.toDataURL('image/png');
      const link = document.createElement('a');
      const todayStr = new Date().toISOString().slice(0, 10);
      link.href = pngUrl;
      link.download = `progres-headcount-${todayStr}.png`;
      link.click();
      setToast('Gambar dashboard berhasil diunduh.');
    } catch (e) {
      console.error(e);
      setError('Gagal membuat gambar. Kalau masih gagal, coba screenshot manual layar ini.');
    } finally {
      setIsExporting(false);
    }
  }

  const stats = useMemo(() => {
    const total = records.length;
    const filled = records.filter(isFilled).length;
    const unfilled = total - filled;
    const pct = total ? Math.round((filled / total) * 100) : 0;

    const candidateAreaMap = new Map();
    let totalRecommended = 0;
    candidates.forEach((k) => {
      if (!isRecommended(k.hasil)) return;
      totalRecommended += 1;
      const key = k.area && String(k.area).trim() !== '' ? k.area : 'Tidak diketahui';
      candidateAreaMap.set(key, (candidateAreaMap.get(key) || 0) + 1);
    });

    const areaMap = new Map();
    records.forEach((r) => {
      const key = r.area && String(r.area).trim() !== '' ? r.area : 'Tidak diketahui';
      if (!areaMap.has(key)) areaMap.set(key, { area: key, total: 0, filled: 0 });
      const e = areaMap.get(key);
      e.total += 1;
      if (isFilled(r)) e.filled += 1;
    });
    const areaData = Array.from(areaMap.values())
      .map((e) => {
        const unfilledArea = e.total - e.filled;
        const kandidat = candidateAreaMap.get(e.area) || 0;
        const predictedFill = Math.min(unfilledArea, Math.floor(kandidat / CONVERSION_RATIO));
        const predictedFillPct = unfilledArea > 0 ? Math.round((predictedFill / unfilledArea) * 100) : 100;
        return {
          ...e,
          unfilled: unfilledArea,
          pct: e.total ? Math.round((e.filled / e.total) * 100) : 0,
          kandidat,
          predictedFill,
          predictedFillPct,
        };
      })
      .sort((a, b) => b.total - a.total);

    const unfilledList = records
      .filter((r) => !isFilled(r))
      .map((r) => ({ ...r, status: daysStatus(r.tmt) }))
      .sort((a, b) => {
        if (!a.tmt) return 1;
        if (!b.tmt) return -1;
        return new Date(a.tmt) - new Date(b.tmt);
      });

    const totalPredictedFill = Math.min(unfilled, Math.floor(totalRecommended / CONVERSION_RATIO));
    const totalCandidateGap = Math.max(0, unfilled * CONVERSION_RATIO - totalRecommended);
    const predictedFillPct = unfilled > 0 ? Math.round((totalPredictedFill / unfilled) * 100) : 100;

    return {
      total, filled, unfilled, pct, areaData, unfilledList,
      totalRecommended, totalPredictedFill, totalCandidateGap, predictedFillPct,
    };
  }, [records, candidates]);

  const unfilledAreaOptions = useMemo(() => {
    const set = new Set(stats.unfilledList.map((r) => (r.area && String(r.area).trim() !== '' ? r.area : 'Tidak diketahui')));
    return Array.from(set).sort();
  }, [stats.unfilledList]);

  const unfilledAlasanOptions = useMemo(() => {
    const set = new Set(stats.unfilledList.map((r) => (r.ket && String(r.ket).trim() !== '' ? r.ket : 'Tidak diketahui')));
    return Array.from(set).sort();
  }, [stats.unfilledList]);

  const displayedUnfilled = useMemo(() => {
    let list = stats.unfilledList;
    if (filterArea) {
      list = list.filter((r) => (r.area && String(r.area).trim() !== '' ? r.area : 'Tidak diketahui') === filterArea);
    }
    if (filterAlasan) {
      list = list.filter((r) => (r.ket && String(r.ket).trim() !== '' ? r.ket : 'Tidak diketahui') === filterAlasan);
    }
    if (searchText.trim()) {
      const q = searchText.trim().toLowerCase();
      list = list.filter((r) =>
        [r.jabatan, r.unit, r.area, r.ket].some((v) => v && String(v).toLowerCase().includes(q))
      );
    }
    if (sortKey) {
      list = [...list].sort((a, b) => {
        const cmp = compareUnfilledRows(a, b, sortKey);
        return sortDir === 'asc' ? cmp : -cmp;
      });
    }
    return list;
  }, [stats.unfilledList, filterArea, filterAlasan, searchText, sortKey, sortDir]);

  const hasActiveFilter = !!(filterArea || filterAlasan || searchText.trim());
  const lastUpdatedLabel = formatDateTimeID(lastUpdated);

  return (
    <div className="hc-root">
      <div className="hc-shell" ref={dashboardRef}>
        {toast && (
          <div className="hc-toast"><CheckCircle2 size={16} /> {toast}</div>
        )}
        {error && (
          <div className="hc-error"><AlertTriangle size={16} /> {error}</div>
        )}
        {!isConnected && !error && (
          <div className="hc-toast">Menyambungkan ke Firestore...</div>
        )}

        <header className="hc-header">
          <div>
            <h1 className="hc-title">Progres Pemenuhan Headcount</h1>
            <p className="hc-subtitle">Level staff (N6) organik &middot; resign, pensiun, PHK, meninggal, dan cabang baru</p>
          </div>
          <div className="hc-header-actions">
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls"
              style={{ display: 'none' }}
              onChange={onInputChange}
            />
            <button className="hc-btn hc-btn-outline" onClick={handleDownloadImage} disabled={isExporting}>
              <Download size={16} /> {isExporting ? 'Menyiapkan gambar...' : 'Unduh sebagai gambar'}
            </button>
            <button
              className="hc-btn"
              onClick={() => fileInputRef.current && fileInputRef.current.click()}
              disabled={isUploading}
            >
              <Upload size={16} /> {isUploading ? 'Mengunggah...' : 'Upload data baru'}
            </button>
          </div>
        </header>

        <div className="hc-hero">
          <div className="hc-rings">
            <div className="hc-ring-block">
              <p className="hc-ring-title">Kondisi Saat ini</p>
              <ProgressRing pct={stats.pct} color="#0B6E4F" label="terisi" />
            </div>
            <div className="hc-ring-block">
              <p className="hc-ring-title">Potensi setelah asesmen</p>
              <ProgressRing pct={stats.predictedFillPct} color="#C9A227" label="prediksi" size={106} stroke={9} />
            </div>
          </div>
          <div className="hc-chips">
            <div className="hc-chip">
              <span className="hc-chip-value">{stats.total}</span>
              <span className="hc-chip-label">Total Headcount</span>
            </div>
            <div className="hc-chip">
              <span className="hc-chip-value is-primary">{stats.filled}</span>
              <span className="hc-chip-label">Sudah terisi</span>
            </div>
            <div className="hc-chip">
              <span className="hc-chip-value is-warn">{stats.unfilled}</span>
              <span className="hc-chip-label">Belum terisi</span>
            </div>
            <div className="hc-chip">
              <span className="hc-chip-value is-gold">{stats.totalRecommended}</span>
              <span className="hc-chip-label">Kandidat</span>
            </div>
          </div>
        </div>
        <p className="hc-hero-note">
          Ring "prediksi" = perkiraan persentase sisa Headcount yang belum terisi bisa tertutup dari kandidat yang ada sekarang, pakai asumsi rasio konversi historis 1 dari {CONVERSION_RATIO} kandidat berhasil ditempatkan.
          {stats.totalCandidateGap > 0
            ? ` Butuh sekitar ${stats.totalCandidateGap} kandidat lagi untuk menutup sisa ${stats.unfilled} Headcount yang belum terisi.`
            : ' Kandidat yang ada saat ini diperkirakan cukup untuk menutup sisa Headcount yang belum terisi.'}
        </p>

        <section className="hc-panel">
          <div className="hc-panel-header">
            <h2>Progres per area</h2>
            <p className="hc-panel-sub">
              Terisi, belum terisi, jumlah kandidat, dan prediksi pemenuhan per area &middot; diurutkan dari Headcount terbanyak
            </p>
          </div>
          {!hasCandidateSheet && (
            <p className="hc-empty">Sheet Kandidat tidak ditemukan di file terakhir, kolom kandidat menunjukkan 0.</p>
          )}
          <div className="hc-table-scroll">
            <table className="hc-table">
              <thead>
                <tr>
                  <th>Area</th>
                  <th>Total Headcount</th>
                  <th>Terisi</th>
                  <th>Belum terisi</th>
                  <th>% Terisi</th>
                  <th>Kandidat</th>
                  <th>Prediksi terisi</th>
                  <th>% Prediksi terisi</th>
                </tr>
              </thead>
              <tbody>
                {stats.areaData.map((a) => (
                  <tr key={a.area}>
                    <td>{a.area}</td>
                    <td>{a.total}</td>
                    <td>{a.filled}</td>
                    <td>{a.unfilled}</td>
                    <td className="hc-pct-cell">{a.pct}%</td>
                    <td className="hc-kandidat-cell">{a.kandidat}</td>
                    <td className="hc-predict-cell">{a.predictedFill}</td>
                    <td className="hc-predict-cell">{a.predictedFillPct}%</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>

        <section className="hc-panel">
          <div className="hc-panel-header">
            <h2>Belum terisi ({stats.unfilledList.length})</h2>
            <p className="hc-panel-sub">Klik judul kolom untuk sortir &middot; pakai filter di bawah untuk mempersempit</p>
          </div>
          {stats.unfilledList.length === 0 ? (
            <p className="hc-empty">Semua Headcount sudah terisi.</p>
          ) : (
            <>
              <div className="hc-filter-bar">
                <input
                  type="text"
                  className="hc-filter-input"
                  placeholder="Cari posisi, unit, area, atau alasan..."
                  value={searchText}
                  onChange={(e) => setSearchText(e.target.value)}
                />
                <select className="hc-filter-select" value={filterArea} onChange={(e) => setFilterArea(e.target.value)}>
                  <option value="">Semua area</option>
                  {unfilledAreaOptions.map((a) => (
                    <option key={a} value={a}>{a}</option>
                  ))}
                </select>
                <select className="hc-filter-select" value={filterAlasan} onChange={(e) => setFilterAlasan(e.target.value)}>
                  <option value="">Semua alasan</option>
                  {unfilledAlasanOptions.map((a) => (
                    <option key={a} value={a}>{a}</option>
                  ))}
                </select>
                <div className="hc-col-toggle" ref={columnMenuRef}>
                  <button className="hc-btn hc-btn-outline hc-col-btn" onClick={() => setShowColumnMenu((s) => !s)}>
                    <EyeOff size={14} /> Kolom
                  </button>
                  {showColumnMenu && (
                    <div className="hc-col-menu">
                      {UNFILLED_COLUMNS.map((col) => (
                        <label key={col.key} className="hc-col-menu-item">
                          <input
                            type="checkbox"
                            checked={visibleColumns[col.key]}
                            onChange={() => toggleColumn(col.key)}
                          />
                          {col.label}
                        </label>
                      ))}
                    </div>
                  )}
                </div>
                {hasActiveFilter && (
                  <button
                    className="hc-filter-reset"
                    onClick={() => { setSearchText(''); setFilterArea(''); setFilterAlasan(''); }}
                  >
                    Reset filter
                  </button>
                )}
              </div>
              {hasActiveFilter && (
                <p className="hc-filter-count">Menampilkan {displayedUnfilled.length} dari {stats.unfilledList.length} baris</p>
              )}
              {displayedUnfilled.length === 0 ? (
                <p className="hc-empty">Gak ada yang cocok dengan filter ini.</p>
              ) : (
                <div className="hc-table-scroll">
                  <table className="hc-table">
                    <thead>
                      <tr>
                        {visibleColumns.jabatan && <th className="hc-th-sortable" onClick={() => toggleSort('jabatan')}>Posisi{sortIndicator('jabatan')}</th>}
                        {visibleColumns.unit && <th className="hc-th-sortable" onClick={() => toggleSort('unit')}>Unit{sortIndicator('unit')}</th>}
                        {visibleColumns.area && <th className="hc-th-sortable" onClick={() => toggleSort('area')}>Area{sortIndicator('area')}</th>}
                        {visibleColumns.ket && <th className="hc-th-sortable" onClick={() => toggleSort('ket')}>Alasan{sortIndicator('ket')}</th>}
                        {visibleColumns.tmt && <th className="hc-th-sortable" onClick={() => toggleSort('tmt')}>Tgl. efektif{sortIndicator('tmt')}</th>}
                        {visibleColumns.status && <th className="hc-th-sortable" onClick={() => toggleSort('status')}>Status{sortIndicator('status')}</th>}
                      </tr>
                    </thead>
                    <tbody>
                      {displayedUnfilled.map((r, i) => (
                        <tr key={i}>
                          {visibleColumns.jabatan && <td>{r.jabatan || '\u2014'}</td>}
                          {visibleColumns.unit && <td>{r.unit || '\u2014'}</td>}
                          {visibleColumns.area && <td>{r.area || '\u2014'}</td>}
                          {visibleColumns.ket && <td>{r.ket || '\u2014'}</td>}
                          {visibleColumns.tmt && <td>{formatDateID(r.tmt)}</td>}
                          {visibleColumns.status && <td><span className={`hc-status hc-status-${r.status.tone}`}>{r.status.label}</span></td>}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </>
          )}
        </section>

        <footer className="hc-footer">
          <FileSpreadsheet size={13} />
          <span>Sumber: {fileName}</span>
          <span className="hc-dot">&middot;</span>
          <span>{lastUpdatedLabel ? `Diperbarui ${lastUpdatedLabel}` : 'Data awal'}</span>
        </footer>
      </div>
    </div>
  );
}
