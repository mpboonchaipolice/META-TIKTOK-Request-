/* ====================================================================
   ระบบออกหนังสือขอข้อมูล — Application Logic
   ทำงาน 100% ใน Browser · ไม่มีการส่งข้อมูลออกไปยัง server ใดๆ
   ==================================================================== */

(() => {
'use strict';

// ====================================================================
// Configuration
// ====================================================================
const CONFIG = {
  urlsPerLetter: {
    facebook:  4,
    instagram: 4,
    tiktok:    1
  },
  platformNames: {
    facebook:  'Facebook',
    instagram: 'Instagram',
    tiktok:    'TikTok'
  },
  // ไฟล์ XML ใน .docx ที่อาจมีตัวแปรอยู่
  xmlFilesToProcess: [
    'word/document.xml',
    'word/header1.xml',
    'word/header2.xml',
    'word/header3.xml',
    'word/footer1.xml',
    'word/footer2.xml',
    'word/footer3.xml',
    'word/footnotes.xml',
    'word/endnotes.xml'
  ],
  // ตัวแปรที่รองรับใน template
  variables: ['LetterDate', 'LetterNo', 'StartTime', 'URL1', 'URL2', 'URL3', 'URL4'],
  // IndexedDB
  dbName: 'TCSD2_LetterGen',
  dbVersion: 1,
  storeName: 'templates'
};

// ====================================================================
// State
// ====================================================================
const state = {
  templates: {},        // { facebook: ArrayBuffer, instagram: ArrayBuffer, tiktok: ArrayBuffer }
  templateNames: {},    // { facebook: 'filename.docx', ... }
  platform: 'facebook',
  excelData: null,
  excelName: null,
  generatedFiles: [],   // [{ name, blob, urls, letterNo }]
};

// ====================================================================
// DOM refs
// ====================================================================
const $  = (id) => document.getElementById(id);
const $$ = (sel) => document.querySelectorAll(sel);

// ====================================================================
// IndexedDB — เก็บ template ไว้ในเครื่อง
// ====================================================================
function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(CONFIG.dbName, CONFIG.dbVersion);
    req.onerror   = () => reject(req.error);
    req.onsuccess = () => resolve(req.result);
    req.onupgradeneeded = (e) => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains(CONFIG.storeName)) {
        db.createObjectStore(CONFIG.storeName);
      }
    };
  });
}

async function saveTemplateToDb(platform, arrayBuffer, fileName) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(CONFIG.storeName, 'readwrite');
    const store = tx.objectStore(CONFIG.storeName);
    store.put({ data: arrayBuffer, name: fileName }, platform);
    tx.oncomplete = () => resolve();
    tx.onerror    = () => reject(tx.error);
  });
}

async function loadTemplateFromDb(platform) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(CONFIG.storeName, 'readonly');
    const store = tx.objectStore(CONFIG.storeName);
    const req = store.get(platform);
    req.onsuccess = () => resolve(req.result || null);
    req.onerror   = () => reject(req.error);
  });
}

async function clearAllTemplatesFromDb() {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(CONFIG.storeName, 'readwrite');
    tx.objectStore(CONFIG.storeName).clear();
    tx.oncomplete = () => resolve();
    tx.onerror    = () => reject(tx.error);
  });
}

async function restoreTemplatesFromDb() {
  for (const p of ['facebook', 'instagram', 'tiktok']) {
    try {
      const rec = await loadTemplateFromDb(p);
      if (rec && rec.data) {
        state.templates[p]     = rec.data;
        state.templateNames[p] = rec.name;
        markTemplateLoaded(p, rec.name);
      }
    } catch (e) { console.warn('โหลด template ' + p + ' ไม่สำเร็จ:', e); }
  }
  updateTemplateStatus();
}

// ====================================================================
// Template UI
// ====================================================================
function markTemplateLoaded(platform, fileName) {
  const card  = document.querySelector(`.template-card[data-platform="${platform}"]`);
  const state_ = document.querySelector(`[data-state="${platform}"]`);
  if (card) card.classList.add('loaded');
  if (state_) state_.textContent = fileName;
}

function updateTemplateStatus() {
  const loaded = Object.keys(state.templates).filter(k => state.templates[k]).length;
  const status = $('templateStatus');
  if (loaded === 0) {
    status.className = 'status status-empty';
    status.textContent = 'ยังไม่ได้โหลด Template';
  } else if (loaded < 3) {
    status.className = 'status status-partial';
    status.textContent = `โหลดแล้ว ${loaded} / 3 Template`;
  } else {
    status.className = 'status status-ready';
    status.textContent = 'พร้อมใช้งาน · 3 / 3 Template';
  }
}
function setupTemplateUpload() {
  document.querySelectorAll('input[data-tpl]').forEach(input => {
    input.addEventListener('change', async (e) => {
      const file = e.target.files[0];
      if (!file) return;

      const fileName = file.name.toLowerCase();

      if (!fileName.endsWith('.docx')) {
        alert('กรุณาเลือกไฟล์ .docx เท่านั้น');
        return;
      }

      const platform = e.target.dataset.tpl;

      try {
        const buf = await file.arrayBuffer();

        // ตรวจแบบพื้นฐานว่าเป็นไฟล์ ZIP/DOCX หรือไม่
        // ไฟล์ .docx จริง ๆ จะขึ้นต้นด้วยตัวอักษร PK
        const header = new Uint8Array(buf.slice(0, 4));
        const isZipLike = header[0] === 0x50 && header[1] === 0x4B;

        if (!isZipLike) {
          alert(
            'ไฟล์นี้ไม่ใช่ .docx แท้\n\n' +
            'วิธีแก้:\n' +
            '1) เปิดไฟล์ใน Microsoft Word\n' +
            '2) กด File > Save As\n' +
            '3) เลือก Word Document (*.docx)\n' +
            '4) เอาไฟล์ใหม่มาอัปโหลดอีกครั้ง'
          );
          return;
        }

        // ลองอ่านด้วย PizZip แบบไม่บล็อกไฟล์เร็วเกินไป
        try {
          const testZip = new PizZip(buf);

          if (!testZip.file('word/document.xml')) {
            console.warn('ไม่พบ word/document.xml แต่จะยังไม่บล็อกทันที:', file.name);
          }
        } catch (zipErr) {
          console.error('DOCX ZIP WARNING:', zipErr);
          alert(
            'ระบบอ่านโครงสร้าง .docx ไม่สำเร็จ\n\n' +
            'กรุณาลอง Save As ไฟล์ Word ใหม่เป็น .docx อีกครั้ง\n' +
            'ถ้ายังไม่ได้ ให้ส่งไฟล์ Template มาให้ตรวจโครงสร้าง'
          );
          return;
        }

        state.templates[platform] = buf;
        state.templateNames[platform] = file.name;

        await saveTemplateToDb(platform, buf, file.name);

        markTemplateLoaded(platform, file.name);
        updateTemplateStatus();

        alert('โหลด Template สำเร็จ: ' + file.name);

      } catch (err) {
        console.error('โหลด template ไม่สำเร็จ:', err);
        alert('โหลด template ไม่สำเร็จ: ' + err.message);
      }
    });
  });

  $('clearTemplates').addEventListener('click', async () => {
    if (!confirm('ล้าง Template ทั้งหมดออกจากเบราว์เซอร์?')) return;

    await clearAllTemplatesFromDb();

    state.templates = {};
    state.templateNames = {};

    document.querySelectorAll('.template-card').forEach(c => c.classList.remove('loaded'));
    document.querySelectorAll('.tpl-state').forEach(s => s.textContent = 'คลิกเพื่อเลือกไฟล์');
    document.querySelectorAll('input[data-tpl]').forEach(i => i.value = '');

    updateTemplateStatus();
  });

  $('manageTemplates').addEventListener('click', () => {
    $('templateDrawer').scrollIntoView({ behavior: 'smooth', block: 'start' });
  });
}
function setupPlatformSelector() {
  document.querySelectorAll('.seg').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.seg').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      state.platform = btn.dataset.platform;
      const count = CONFIG.urlsPerLetter[state.platform];
      $('platformHint').textContent = `${count} URL ต่อ 1 หนังสือ`;
    });
  });
}

// ====================================================================
// Excel upload & parsing
// ====================================================================
function setupExcelUpload() {
  const dz = $('dropZone');
  const inp = $('excelFile');
  const fileLabel = $('dropFile');

  ['dragenter', 'dragover'].forEach(evt =>
    dz.addEventListener(evt, e => { e.preventDefault(); dz.classList.add('dragover'); })
  );
  ['dragleave', 'drop'].forEach(evt =>
    dz.addEventListener(evt, e => { e.preventDefault(); dz.classList.remove('dragover'); })
  );
  dz.addEventListener('drop', e => {
    const f = e.dataTransfer.files[0];
    if (f) { inp.files = e.dataTransfer.files; handleExcelFile(f); }
  });
  inp.addEventListener('change', () => {
    if (inp.files[0]) handleExcelFile(inp.files[0]);
  });

  function handleExcelFile(file) {
    if (!/\.(xlsx|xls)$/i.test(file.name)) {
      alert('กรุณาเลือกไฟล์ .xlsx หรือ .xls');
      return;
    }
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        const urls = [];
        wb.SheetNames.forEach(name => {
          const sheet = wb.Sheets[name];
          const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: '' });
          rows.forEach(row => {
            row.forEach(cell => {
              const v = String(cell || '').trim();
              if (/^https?:\/\//i.test(v)) urls.push(v);
            });
          });
        });
        if (urls.length === 0) {
          alert('ไม่พบ URL ใน Excel (ต้องขึ้นต้นด้วย http:// หรือ https://)');
          return;
        }
        state.excelData = urls;
        state.excelName = file.name;
        dz.classList.add('has-file');
        fileLabel.textContent = `✓ ${file.name} — พบ ${urls.length} URL`;
      } catch (err) {
        alert('อ่าน Excel ไม่สำเร็จ: ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  }
}

// ====================================================================
// Core: Replace variables in .docx
// ====================================================================
function escapeXml(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * รับ ArrayBuffer ของ template และ object ของตัวแปร
 * คืน Blob ของ .docx ที่ replace ตัวแปรเรียบร้อย
 */
function generateDocxBlob(templateArrayBuffer, variables) {
  // แปลงเป็น Uint8Array เพื่อให้ PizZip อ่าน .docx ได้เสถียรกว่า
  const zip = new PizZip(new Uint8Array(templateArrayBuffer.slice(0)));

  CONFIG.xmlFilesToProcess.forEach(filePath => {
    const file = zip.file(filePath);
    if (!file) return;

    let xml = file.asText();

    // Replace ตัวแปรทุกตัว
    CONFIG.variables.forEach(varName => {
      const value = variables[varName];
      if (value === undefined) return;
      // ใน XML, < และ > ถูกเก็บเป็น &lt; และ &gt;
      const search = `&lt;${varName}&gt;`;
      // split + join เป็นวิธี replaceAll ที่เร็วและ safe (ไม่ต้องสร้าง regex)
      xml = xml.split(search).join(escapeXml(value));
    });

    zip.file(filePath, xml);
  });

  return zip.generate({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    compression: 'DEFLATE'
  });
}

// ====================================================================
// Main: Process letters
// ====================================================================
async function processLetters() {
  // Validate
  if (!state.templates[state.platform]) {
    alert(`ยังไม่ได้โหลด Template ของ ${CONFIG.platformNames[state.platform]} — กรุณาอัพโหลดที่ Section 01`);
    $('templateDrawer').scrollIntoView({ behavior: 'smooth' });
    return;
  }
  const letterDate = $('letterDate').value.trim();
  const letterNoRaw = $('letterNo').value.trim();
  const startTime = $('startTime').value.trim();

  if (!letterDate || !letterNoRaw || !startTime) {
    alert('กรุณากรอกข้อมูลให้ครบ');
    return;
  }
  const startNo = parseInt(letterNoRaw, 10);
  if (isNaN(startNo)) {
    alert('เลขที่หนังสือต้องเป็นตัวเลข');
    return;
  }
  if (!state.excelData || state.excelData.length === 0) {
    alert('กรุณาอัพโหลด Excel ที่มี URL');
    return;
  }

  // จัดกลุ่ม URL
  const perLetter = CONFIG.urlsPerLetter[state.platform];
  const groups = [];
  for (let i = 0; i < state.excelData.length; i += perLetter) {
    groups.push(state.excelData.slice(i, i + perLetter));
  }

  // Show loading
  $('result').classList.add('hidden');
  $('loading').classList.remove('hidden');
  $('loadingTotal').textContent = groups.length;
  $('btnRun').disabled = true;
  state.generatedFiles = [];

  try {
    const platformName = CONFIG.platformNames[state.platform];
    for (let i = 0; i < groups.length; i++) {
      const group = groups[i];
      const currentNo = startNo + i;
      const vars = {
        LetterDate: letterDate,
        LetterNo:   String(currentNo),
        StartTime:  startTime,
        URL1: group[0] || '',
        URL2: group[1] || '',
        URL3: group[2] || '',
        URL4: group[3] || ''
      };

      const blob = generateDocxBlob(state.templates[state.platform], vars);
      const fileName = `หนังสือ_${platformName}_เลขที่_${currentNo}.docx`;
      state.generatedFiles.push({
        name: fileName,
        blob: blob,
        urls: group,
        letterNo: currentNo
      });

      // Update progress
      $('loadingProgress').textContent = i + 1;
      $('.loading-bar-fill') && document.querySelector('.loading-bar-fill').style.setProperty('width', ((i+1)/groups.length*100) + '%');
      const fillEl = document.querySelector('.loading-bar-fill');
      if (fillEl) fillEl.style.width = ((i+1)/groups.length*100) + '%';

      // Yield to browser ให้ render progress ได้
      if (i % 5 === 0) await new Promise(r => setTimeout(r, 0));
    }

    renderResult();
  } catch (err) {
    renderError(err.message || err.toString());
  } finally {
    $('loading').classList.add('hidden');
    $('btnRun').disabled = false;
  }
}

// ====================================================================
// Render Result
// ====================================================================
function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({
    '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
  }[c]));
}

function renderResult() {
  const result = $('result');
  const body   = $('resultBody');
  result.classList.remove('hidden');

  const files = state.generatedFiles;
  const totalUrls = files.reduce((sum, f) => sum + f.urls.length, 0);
  const platformName = CONFIG.platformNames[state.platform];

  let html = `
    <div class="result-summary">
      <div class="stat">
        <div class="stat-value">${files.length}</div>
        <div class="stat-label">หนังสือทั้งหมด</div>
      </div>
      <div class="stat">
        <div class="stat-value">${totalUrls}</div>
        <div class="stat-label">URL ที่ขอ</div>
      </div>
      <div class="stat">
        <div class="stat-value">${platformName}</div>
        <div class="stat-label">แบบฟอร์ม</div>
      </div>
      <div class="stat">
        <div class="stat-value">#${files[0].letterNo}–${files[files.length-1].letterNo}</div>
        <div class="stat-label">ช่วงเลขหนังสือ</div>
      </div>
    </div>

    <div class="batch-bar">
      <button class="batch-btn" id="batchAllDocx">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>
        </svg>
        ดาวน์โหลดทั้งหมด · ZIP
      </button>
      <button class="batch-btn" id="batchOpenAll">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/>
        </svg>
        เปิดทุกไฟล์ในแท็บใหม่
      </button>
    </div>

    <div class="file-list">
      ${files.map((f, i) => `
        <div class="file-row">
          <div class="file-num">№ ${f.letterNo}</div>
          <div class="file-main">
            <div class="file-name">${escapeHtml(f.name)}</div>
            <div class="file-urls">
              ${f.urls.map((u, j) => `<span>URL${j+1}: ${escapeHtml(u)}</span>`).join('')}
            </div>
          </div>
          <button class="dl-btn" data-idx="${i}">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>
            </svg>
            .docx
          </button>
        </div>
      `).join('')}
    </div>
  `;
  body.innerHTML = html;

  // Wire up
  $('batchAllDocx').addEventListener('click', downloadBatchZip);
  $('batchOpenAll').addEventListener('click', () => {
    if (!confirm(`เปิด ${files.length} ไฟล์พร้อมกัน? เบราว์เซอร์อาจเตือนเรื่อง popup`)) return;
    files.forEach(f => {
      const url = URL.createObjectURL(f.blob);
      const a = document.createElement('a');
      a.href = url; a.download = f.name;
      a.click();
      setTimeout(() => URL.revokeObjectURL(url), 5000);
    });
  });
  body.querySelectorAll('.dl-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const idx = parseInt(btn.dataset.idx, 10);
      const f = files[idx];
      saveAs(f.blob, f.name);
    });
  });

  // Scroll
  result.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function renderError(message) {
  $('result').classList.remove('hidden');
  $('resultBody').innerHTML = `
    <div class="error-box">
      <strong>เกิดข้อผิดพลาด</strong>
      ${escapeHtml(message)}
    </div>
  `;
}

async function downloadBatchZip() {
  const btn = $('batchAllDocx');
  const orig = btn.innerHTML;
  btn.disabled = true;
  btn.innerHTML = 'กำลังสร้าง ZIP...';
  try {
    const zip = new JSZip();
    state.generatedFiles.forEach(f => zip.file(f.name, f.blob));
    const zipBlob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE' });
    const ts = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const zipName = `หนังสือ_${CONFIG.platformNames[state.platform]}_${ts}.zip`;
    saveAs(zipBlob, zipName);
  } catch (err) {
    alert('สร้าง ZIP ไม่สำเร็จ: ' + err.message);
  } finally {
    btn.disabled = false;
    btn.innerHTML = orig;
  }
}

// ====================================================================
// Reset
// ====================================================================
function setupReset() {
  $('btnReset').addEventListener('click', () => {
    $('letterDate').value = '';
    $('letterNo').value = '';
    $('startTime').value = '';
    $('excelFile').value = '';
    $('dropZone').classList.remove('has-file');
    $('dropFile').textContent = '';
    state.excelData = null;
    state.excelName = null;
    $('result').classList.add('hidden');
    state.generatedFiles = [];
  });
}

// ====================================================================
// Init
// ====================================================================
document.addEventListener('DOMContentLoaded', async () => {
  setupTemplateUpload();
  setupPlatformSelector();
  setupExcelUpload();
  setupReset();
  $('btnRun').addEventListener('click', processLetters);
  await restoreTemplatesFromDb();
});

})();
