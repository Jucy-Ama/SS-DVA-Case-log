// ==UserScript==
// @name         SS Case Log Auto-Filler v4.8
// @namespace    http://tampermonkey.net/
// @version      4.8
// @description  v4.8: Fix Brand Name + Primary Goal missing from PARSE_MAP & DOM_MAP. Post-fill reminder shows actual Due Date value; adds optimization category reminder.
// @match        https://advertising.amazon.com/case-manager*
// @match        https://advertising.amazon.co.jp/case-manager*
// @match        https://advertising.amazon.co.uk/case-manager*
// @grant        GM_setValue
// @grant        GM_getValue
// @run-at       document-idle
// ==/UserScript==

(function () {
  'use strict';

  // ============================================================
  // 🗺️ PARSE MAP (longest pattern first)
  // ★ v4.8: Added "Brand Name" and standalone "Primary Goal"
  // ============================================================
  const PARSE_MAP = [
    { pattern: "Assignment creation reason", key: "assignmentCreationReason" },
    { pattern: "Submitter email address",    key: "submitterEmail" },
    { pattern: "Optimization marketplace",   key: "optimizationMarketplace" },
    { pattern: "Optimization categories",    key: "optimizationCategories" },
    { pattern: "Primary Goal Consideration", key: "primaryGoalConsideration" },
    { pattern: "Optimization delivery",      key: "optimizationDelivery" },
    { pattern: "Additional information",     key: "additionalInformation" },
    { pattern: "Assignment status",          key: "assignmentstatus" },
    { pattern: "Optimization type",          key: "optimizationType" },
    { pattern: "Primary Goal KPI",           key: "primaryGoalKPI" },
    { pattern: "Primary Goal",              key: "primaryGoalConsideration" },  // ★ v4.8 FIX: Must be AFTER "Primary Goal KPI"
    { pattern: "Case Description",           key: "caseDescription" },
    { pattern: "Advertiser Type",            key: "advertiserType" },
    { pattern: "Account Vertical",           key: "accountVertical" },
    { pattern: "Rodeo order CFID",           key: "rodeoCfId" },
    { pattern: "SF account ID",              key: "sfAccountId" },
    { pattern: "Submitting team",            key: "submittingTeam" },
    { pattern: "Advertiser ID",              key: "advertiserId" },
    { pattern: "Account Name",              key: "accountName" },
    { pattern: "Brand Name",                key: "brandName" },               // ★ v4.8 FIX: NEW
    { pattern: "Requested by",              key: "requestedBy" },
    { pattern: "Requester by",              key: "requestedBy" },
    { pattern: "Submitted by",              key: "submittedBy" },
    { pattern: "Case Status",               key: "caseStatus" },
    { pattern: "Entity ID",                 key: "entityId" },
    { pattern: "Due Date",                  key: "dueDate" },
    { pattern: "Assignee",                  key: "assignee" },
  ];

  const IGNORE_VALUES = [
    "please select", "please input", "keep it blank",
    "tick the levers", "select actual", "(keep it blank)"
  ];

  // ============================================================
  // 🎯 DOM MAP
  // ★ v4.8: Added brandName entry
  // ============================================================
  const DOM_MAP = [
    { key: "assignee",                 selector: '[data-testid="assignee-field"]',                    altSelectors: ['[name="assignee"]', '#assignee'], type: "text" },
    { key: "accountVertical",          selector: '[data-testid="accountVertical-field"]',              altSelectors: [], type: "text" },
    { key: "accountName",              selector: '[data-testid="accountName-field"]',                  altSelectors: ['[name="accountName"]'], type: "text" },
    { key: "brandName",                selector: '[data-testid="brandName-field"]',                    altSelectors: ['[data-testid="brand-name-field"]', '[data-testid="brand.name-field"]', '[name="brandName"]', '[id*="brandName"]', '[id*="brand-name"]'], type: "text" },  // ★ v4.8 FIX: NEW
    { key: "rodeoCfId",                selector: '[data-testid="rodeoAdvertiserCfId-field"]',          altSelectors: ['[name="rodeoAdvertiserCfId"]'], type: "text" },
    { key: "advertiserId",             selector: '[data-testid="advertiser.id-field"]',                altSelectors: ['[name="advertiserId"]'], type: "text" },
    { key: "entityId",                 selector: '[data-testid="entityId-field"]',                     altSelectors: ['[name="entityId"]'], type: "text" },
    { key: "primaryGoalKPI",           selector: '[data-testid="primaryGoalKPI-field"]',               altSelectors: [], type: "text" },
    { key: "primaryGoalConsideration", selector: '[data-testid="primaryGoalConsideration-field"]',     altSelectors: ['[name="primaryGoalConsideration"]'], type: "text" },
    { key: "assignmentCreationReason", selector: '[data-testid="assignmentCreationReason-field"]',     altSelectors: [], type: "text" },
    { key: "submitterEmail",           selector: '[data-testid="submittedByEmailAddress-field"]',      altSelectors: ['[name="submittedByEmailAddress"]'], type: "text" },
    { key: "submittedBy",              selector: '[data-testid="submittedBy-field"]',                  altSelectors: ['[name="submittedBy"]'], type: "text" },
    { key: "requestedBy",              selector: '[data-testid="requestedBy-field"]',                  altSelectors: ['[data-testid="requesterBy-field"]', '[data-testid="requester-field"]', '[name="requestedBy"]', '[name="requesterBy"]'], type: "text" },
    { key: "sfAccountId",              selector: '[data-testid="sfAccountId-field"]',                  altSelectors: [], type: "text" },
    { key: "caseDescription",          selector: '[data-testid="caseDescription-field"]',              altSelectors: ['[data-testid="description-field"]', '#caseDescription', '#description', 'textarea[name*="description"]', 'textarea[name*="Description"]', '[data-testid="caseDescription"] textarea'], type: "textarea" },
    { key: "additionalInformation",    selector: '#additionalInformation',                             altSelectors: ['[data-testid="additionalInformation-field"]', 'textarea[name*="additional"]'], type: "textarea" },
    { key: "caseStatus",               selector: '[data-testid="status-field"]',                       altSelectors: ['[data-testid="caseStatus-field"]'], type: "dropdown" },
    { key: "assignmentstatus",
      selector: '[data-testid="assignmentStatus-field"]',
      altSelectors: [
        '[data-testid="assignmentstatus-field"]',
        '[data-testid*="assignment"][data-testid*="Status"]',
        '[name="assignmentStatus"]',
        '[id*="assignmentStatus"]',
        '[aria-label="Select option"]'
      ],
      type: "dropdown" },
    { key: "advertiserType",           selector: '[data-testid="advertiser.type-field"]',              altSelectors: [], type: "dropdown" },
    { key: "submittingTeam",           selector: '[data-testid="submittedByTeam-field"]',              altSelectors: [], type: "dropdown" },
    { key: "optimizationType",         selector: '[data-testid="optimizationType-field"]',             altSelectors: [], type: "dropdown" },
    { key: "optimizationDelivery",     selector: '[data-testid="optimizationDelivery-field"]',         altSelectors: [], type: "dropdown" },
    { key: "optimizationMarketplace",  selector: '[data-testid="advertiser.marketplaceId-field"]',     altSelectors: [], type: "dropdown" },
    { key: "optimizationCategories",   selector: '[data-testid="optimizationCategories-field"]',       altSelectors: [], type: "multiselect" },
  ];

  const CFG = { FILL_DELAY: 400, DROPDOWN_WAIT: 800, STEP_DELAY: 1500 };
  let parsedData = {};

  // ============================================================
  // 🎨 UI — DRAGGABLE + MINIMIZABLE PANEL
  // ============================================================
  function createPanel() {
    const mini = document.createElement('div');
    mini.id = 'ss-mini';
    mini.innerHTML = '🛠';
    mini.title = 'Open SS Case Filler';
    Object.assign(mini.style, {
      position: 'fixed', top: '10px', right: '10px', zIndex: '99999',
      width: '42px', height: '42px', borderRadius: '50%',
      background: '#ff9900', color: '#000', fontSize: '20px',
      display: 'none', cursor: 'pointer', userSelect: 'none',
      boxShadow: '0 2px 12px rgba(0,0,0,0.4)',
      lineHeight: '42px', textAlign: 'center'
    });
    document.body.appendChild(mini);

    const panel = document.createElement('div');
    panel.id = 'ss-panel';
    const savedPos = JSON.parse(localStorage.getItem('ss-panel-pos') || 'null');
    const startTop = savedPos?.top || '10px';
    const startLeft = savedPos?.left || '';
    const startRight = savedPos ? '' : '10px';
    Object.assign(panel.style, {
      position: 'fixed', zIndex: '99999',
      top: startTop, right: startRight,
      background: '#232f3e', color: '#fff', borderRadius: '12px',
      padding: '0', width: '380px', fontFamily: 'Arial, sans-serif',
      boxShadow: '0 4px 24px rgba(0,0,0,0.5)', fontSize: '13px',
      maxHeight: '90vh', display: 'flex', flexDirection: 'column',
      overflow: 'hidden'
    });
    if (startLeft) panel.style.left = startLeft;

    panel.innerHTML = `
      <div id="ss-header" style="
        background:#1a1a2e; padding:10px 14px; cursor:move; user-select:none;
        display:flex; justify-content:space-between; align-items:center;
        border-bottom:1px solid #333; flex-shrink:0;
      ">
        <div style="display:flex; align-items:center; gap:8px;">
          <span style="font-size:16px;">🛠</span>
          <strong style="font-size:14px;">SS Case Filler</strong>
          <span style="font-size:10px;color:#888;background:#333;padding:1px 6px;border-radius:8px;">v4.8</span>
        </div>
        <div style="display:flex; gap:6px; align-items:center;">
          <span id="ss-btn-minimize" title="Minimize to icon" style="cursor:pointer;font-size:14px;padding:2px 6px;border-radius:4px;background:#333;">━</span>
          <span id="ss-btn-resize" title="Collapse panel" style="cursor:pointer;font-size:14px;padding:2px 6px;border-radius:4px;background:#333;">🔽</span>
        </div>
      </div>
      <div id="ss-body" style="padding:14px; overflow-y:auto; flex:1;">
        <div id="ss-paste-section">
          <div style="background:#ff9900;color:#000;padding:5px 10px;border-radius:6px;font-weight:bold;margin-bottom:8px;font-size:12px;">
            1️⃣ PASTE YOUR SS DVA CASE CREATION INFO
          </div>
          <textarea id="ss-input" placeholder="Paste your case log here...
Supports formats:
  Assignee zhxz
  Assignee: zhxz
  Case Description: some text
  Brand Name: ODODOS
  Requested by: John
  Assignment status: In Progress" style="
            width:100%; height:170px; background:#1a2332; color:#e0e0e0;
            border:1px solid #444; border-radius:6px; padding:8px;
            font-family:monospace; font-size:11px; resize:vertical;
            box-sizing:border-box;
          "></textarea>
          <button id="ss-parse-btn" style="
            width:100%; padding:10px; margin-top:6px; border:none;
            border-radius:6px; background:#ff9900; color:#000;
            font-weight:bold; cursor:pointer; font-size:13px;
          ">📋 Parse & Preview</button>
        </div>
        <div id="ss-preview-section" style="display:none; margin-top:8px;">
          <div style="background:#00a651;color:#fff;padding:5px 10px;border-radius:6px;font-weight:bold;margin-bottom:8px;font-size:12px;">
            ✅ PARSED DATA PREVIEW
          </div>
          <div id="ss-preview" style="
            background:#1a2332; padding:8px; border-radius:6px;
            max-height:180px; overflow-y:auto; font-size:11px; line-height:1.8;
          "></div>
          <button id="ss-edit-btn" style="
            width:100%; padding:6px; margin-top:6px; border:1px solid #555;
            border-radius:6px; background:transparent; color:#aaa;
            cursor:pointer; font-size:11px;
          ">✏️ Back to Edit</button>
        </div>
        <div style="margin-top:10px;">
          <div style="background:#4fc3f7;color:#000;padding:5px 10px;border-radius:6px;font-weight:bold;margin-bottom:8px;font-size:12px;">
            2️⃣ CREATE NEW CASE
          </div>
          <div style="display:flex;gap:4px;margin-bottom:8px;">
            <div id="ss-p0" style="flex:1;text-align:center;padding:3px;border-radius:4px;background:#1a2332;font-size:10px;">New Case</div>
            <div id="ss-p1" style="flex:1;text-align:center;padding:3px;border-radius:4px;background:#1a2332;font-size:10px;">Record Type</div>
            <div id="ss-p2" style="flex:1;text-align:center;padding:3px;border-radius:4px;background:#1a2332;font-size:10px;">Fill Form</div>
          </div>
          <button id="ss-create-btn" style="
            width:100%; padding:12px; border:none; border-radius:6px;
            background:#00a651; color:#fff; font-weight:bold;
            cursor:pointer; font-size:14px;
          ">🚀 Create New Case</button>
          <button id="ss-fill-only-btn" style="
            width:100%; padding:7px; margin-top:6px; border:1px solid #00a651;
            border-radius:6px; background:transparent; color:#00a651;
            cursor:pointer; font-size:11px;
          ">✏️ Fill Only (form already open)</button>
        </div>
        <div id="ss-log" style="
          background:#1a2332; padding:8px; border-radius:6px;
          margin-top:10px; max-height:160px; overflow-y:auto;
          font-size:11px; line-height:1.6; display:none;
        "></div>
      </div>
    `;
    document.body.appendChild(panel);

    // ─── DRAG: Main Panel ───
    let isDragging = false, dragOX = 0, dragOY = 0;
    const header = document.getElementById('ss-header');
    header.addEventListener('mousedown', (e) => {
      if (e.target.id === 'ss-btn-minimize' || e.target.id === 'ss-btn-resize') return;
      isDragging = true;
      const rect = panel.getBoundingClientRect();
      dragOX = e.clientX - rect.left;
      dragOY = e.clientY - rect.top;
      panel.style.right = 'auto';
      panel.style.transition = 'none';
      e.preventDefault();
    });
    document.addEventListener('mousemove', (e) => {
      if (!isDragging) return;
      let x = Math.max(0, Math.min(e.clientX - dragOX, window.innerWidth - panel.offsetWidth));
      let y = Math.max(0, Math.min(e.clientY - dragOY, window.innerHeight - 40));
      panel.style.left = x + 'px';
      panel.style.top = y + 'px';
    });
    document.addEventListener('mouseup', () => {
      if (isDragging) {
        isDragging = false;
        panel.style.transition = '';
        localStorage.setItem('ss-panel-pos', JSON.stringify({ top: panel.style.top, left: panel.style.left }));
      }
    });

    // ─── DRAG: Mini Button ───
    let miniDrag = false, miniOX = 0, miniOY = 0, miniMoved = false;
    mini.addEventListener('mousedown', (e) => {
      miniDrag = true; miniMoved = false;
      miniOX = e.clientX - mini.getBoundingClientRect().left;
      miniOY = e.clientY - mini.getBoundingClientRect().top;
      mini.style.transition = 'none';
      e.preventDefault();
    });
    document.addEventListener('mousemove', (e) => {
      if (!miniDrag) return;
      miniMoved = true;
      mini.style.left = (e.clientX - miniOX) + 'px';
      mini.style.top = (e.clientY - miniOY) + 'px';
      mini.style.right = 'auto';
    });
    document.addEventListener('mouseup', () => {
      if (miniDrag) {
        miniDrag = false;
        mini.style.transition = '';
        if (!miniMoved) { mini.style.display = 'none'; panel.style.display = 'flex'; }
      }
    });

    // ─── MINIMIZE ───
    document.getElementById('ss-btn-minimize').addEventListener('click', () => {
      panel.style.display = 'none';
      mini.style.display = 'block';
      const rect = panel.getBoundingClientRect();
      mini.style.top = rect.top + 'px';
      mini.style.left = rect.left + 'px';
      mini.style.right = 'auto';
    });

    // ─── COLLAPSE ───
    let collapsed = false;
    document.getElementById('ss-btn-resize').addEventListener('click', () => {
      collapsed = !collapsed;
      document.getElementById('ss-body').style.display = collapsed ? 'none' : 'block';
      document.getElementById('ss-btn-resize').textContent = collapsed ? '🔼' : '🔽';
    });

    // ─── BUTTONS ───
    document.getElementById('ss-parse-btn').addEventListener('click', parseInput);
    document.getElementById('ss-edit-btn').addEventListener('click', () => {
      document.getElementById('ss-paste-section').style.display = 'block';
      document.getElementById('ss-preview-section').style.display = 'none';
    });
    document.getElementById('ss-create-btn').addEventListener('click', runFullFlow);
    document.getElementById('ss-fill-only-btn').addEventListener('click', () => stepFillForm());

    const saved = localStorage.getItem('ss-last-input');
    if (saved) document.getElementById('ss-input').value = saved;
  }

  // ============================================================
  // 📋 PARSE — strip colons / full-width colons
  // ============================================================
  function parseInput() {
    const raw = document.getElementById('ss-input').value.trim();
    if (!raw) { alert('Please paste your case info first!'); return; }
    localStorage.setItem('ss-last-input', raw);
    parsedData = {};
    const lines = raw.split('\n');
    for (let line of lines) {
      line = line.trim();
      if (!line || line.startsWith('===')) continue;
      for (const { pattern, key } of PARSE_MAP) {
        if (line.toLowerCase().startsWith(pattern.toLowerCase())) {
          let value = line.substring(pattern.length).replace(/^[\s:：;；]+/, '').trim();
          value = cleanValue(value);
          if (value) {
            if (!parsedData[key]) parsedData[key] = value;
          }
          break;
        }
      }
    }
    showPreview();
  }

  function cleanValue(val) {
    if (!val) return '';
    val = val.replace(/\s*\((?:Order Level ID|Rodeo Account Level ID|14 business days[^)]*|mandatory|keep it blank|Proactive support[^)]*|= Case creation date|Today \\+ 13 calendar days)\)\s*/gi, '').trim();
    const lower = val.toLowerCase();
    for (const ig of IGNORE_VALUES) {
      if (lower.includes(ig.toLowerCase())) return '';
    }
    if (val.includes('  ')) val = val.split(/\s{2,}/)[0].trim();
    return val;
  }

  function showPreview() {
    const el = document.getElementById('ss-preview');
    const count = Object.keys(parsedData).length;
    const shownKeys = new Set();
    let html = \`<div style="color:#69f0ae;margin-bottom:4px;"><strong>✅ \${count} fields parsed</strong></div>\`;
    for (const { pattern, key } of PARSE_MAP) {
      if (shownKeys.has(key)) continue;
      shownKeys.add(key);
      const v = parsedData[key];
      html += v
        ? \`<div><span style="color:#888;">\${pattern}:</span> <span style="color:#69f0ae;">\${v}</span></div>\`
        : \`<div><span style="color:#444;">\${pattern}: (skip)</span></div>\`;
    }
    el.innerHTML = html;
    document.getElementById('ss-paste-section').style.display = 'none';
    document.getElementById('ss-preview-section').style.display = 'block';
  }

  // ============================================================
  // 📊 HELPERS
  // ============================================================
  function log(msg, reset = false) {
    const el = document.getElementById('ss-log');
    el.style.display = 'block';
    if (reset) el.innerHTML = '';
    el.innerHTML += msg + '<br>';
    el.scrollTop = el.scrollHeight;
  }

  function setProgress(n) {
    for (let i = 0; i <= 2; i++) {
      const el = document.getElementById(\`ss-p\${i}\`);
      el.style.background = i < n ? '#00a651' : i === n ? '#ff9900' : '#1a2332';
      el.style.color = i === n ? '#000' : '#fff';
    }
  }

  function delay(ms) { return new Promise(r => setTimeout(r, ms)); }

  function normalize(str) {
    return str.toLowerCase().replace(/[\s_\-]+/g, '');
  }

  // ============================================================
  // 🔍 FIND ELEMENT
  // ============================================================
  function findElement(field) {
    let el = document.querySelector(field.selector);
    if (el) return el;

    if (field.altSelectors && field.altSelectors.length) {
      for (const alt of field.altSelectors) {
        try { el = document.querySelector(alt); } catch (e) { continue; }
        if (el) {
          log(\`  ↪ <span style="color:#4fc3f7;">\${field.key}</span>: found via alt selector\`);
          return el;
        }
      }
    }

    const normalizedKey = normalize(field.key);
    const labels = document.querySelectorAll('label, [class*="label"], [class*="Label"]');
    for (const lbl of labels) {
      const normalizedLabel = normalize(lbl.textContent.trim());
      if (normalizedLabel.includes(normalizedKey) || normalizedKey.includes(normalizedLabel)) {
        const parent = lbl.closest('[class*="field"], [class*="form-group"], [class*="row"]') || lbl.parentElement;
        if (parent) {
          el = parent.querySelector('input, textarea, select, button, [role="combobox"], [role="listbox"]');
          if (el) {
            log(\`  ↪ <span style="color:#4fc3f7;">\${field.key}</span>: found via label search\`);
            return el;
          }
        }
      }
    }

    if (field.type === 'dropdown' || field.type === 'multiselect') {
      const readable = field.key.replace(/([A-Z])/g, ' $1').toLowerCase().trim();
      const words = readable.split(/\s+/).filter(w => w.length > 3);
      const allFields = document.querySelectorAll('[class*="field"], [class*="form-group"], [class*="form-row"]');
      for (const container of allFields) {
        const text = container.textContent.toLowerCase();
        if (words.length >= 2 && words.every(w => text.includes(w))) {
          const trigger = container.querySelector(
            'button, select, [role="combobox"], [role="button"], [role="listbox"], [class*="trigger"], [class*="dropdown"], [class*="select"]'
          );
          if (trigger) {
            log(\`  ↪ <span style="color:#ff9900;">\${field.key}</span>: found via nuclear fallback\`);
            return trigger;
          }
        }
      }
    }

    return null;
  }

  // ============================================================
  // ✏️ FILL TEXT
  // ============================================================
  function fillTextInput(el, value) {
    try {
      if (el.getAttribute('contenteditable') === 'true' || el.isContentEditable) {
        el.focus();
        el.innerHTML = '';
        el.textContent = value;
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        el.dispatchEvent(new Event('blur', { bubbles: true }));
        return true;
      }
      if (!['INPUT', 'TEXTAREA'].includes(el.tagName)) {
        const inner = el.querySelector('input, textarea');
        if (inner) {
          el = inner;
        } else {
          const ce = el.querySelector('[contenteditable="true"]');
          if (ce) {
            ce.focus();
            ce.innerHTML = '';
            ce.textContent = value;
            ce.dispatchEvent(new Event('input', { bubbles: true }));
            ce.dispatchEvent(new Event('change', { bubbles: true }));
            ce.dispatchEvent(new Event('blur', { bubbles: true }));
            return true;
          }
          try {
            el.value = value;
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
            return true;
          } catch (e) { return false; }
        }
      }
      el.focus(); el.click();
      const proto = el.tagName === 'TEXTAREA' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
      const setter = Object.getOwnPropertyDescriptor(proto, 'value')?.set;
      if (setter) setter.call(el, value); else el.value = value;
      el.dispatchEvent(new Event('input', { bubbles: true }));
      el.dispatchEvent(new Event('change', { bubbles: true }));
      el.dispatchEvent(new Event('blur', { bubbles: true }));
      return true;
    } catch (e) {
      console.error('[SS Filler] fillTextInput error:', e);
      return false;
    }
  }

  // ============================================================
  // 🔽 FILL DROPDOWN
  // ============================================================
  async function fillDropdown(buttonEl, value) {
    try {
      let clickTarget = buttonEl;
      if (!['BUTTON', 'SELECT'].includes(buttonEl.tagName)) {
        const inner = buttonEl.querySelector('button, [role="combobox"], [role="button"], [class*="trigger"], [class*="select"]');
        if (inner) clickTarget = inner;
      }
      clickTarget.scrollIntoView({ block: 'center' });
      await delay(200);
      clickTarget.click();
      await delay(CFG.DROPDOWN_WAIT);

      const sels = [
        '[role="option"]', '[role="menuitem"]', '[class*="option"]',
        '[class*="menu-item"]', '[class*="dropdown"] li', '[class*="select"] li',
        '[class*="listbox"] li'
      ];
      let allOpts = [];
      for (const s of sels) {
        document.querySelectorAll(s).forEach(o => {
          const r = o.getBoundingClientRect();
          if (r.width > 0 && r.height > 0) allOpts.push(o);
        });
      }
      const lv = value.toLowerCase().trim();
      let match = allOpts.find(o => o.textContent.trim().toLowerCase() === lv)
        || allOpts.find(o => o.textContent.trim().toLowerCase().startsWith(lv))
        || allOpts.find(o => o.textContent.trim().toLowerCase().includes(lv));
      if (match) { match.click(); await delay(300); return true; }

      document.body.click();
      await delay(500);
      clickTarget.click();
      await delay(CFG.DROPDOWN_WAIT + 200);
      allOpts = [];
      for (const s of sels) {
        document.querySelectorAll(s).forEach(o => {
          const r = o.getBoundingClientRect();
          if (r.width > 0 && r.height > 0) allOpts.push(o);
        });
      }
      match = allOpts.find(o => o.textContent.trim().toLowerCase() === lv)
        || allOpts.find(o => o.textContent.trim().toLowerCase().startsWith(lv))
        || allOpts.find(o => o.textContent.trim().toLowerCase().includes(lv));
      if (match) { match.click(); await delay(300); return true; }

      document.body.click(); await delay(200);
      return false;
    } catch (e) {
      console.error('[SS Filler] fillDropdown error:', e);
      document.body.click();
      return false;
    }
  }

  async function fillMultiSelect(buttonEl, value) {
    try {
      const values = value.split(',').map(v => v.trim()).filter(Boolean);
      if (!values.length) return false;
      let clickTarget = buttonEl;
      if (!['BUTTON', 'SELECT'].includes(buttonEl.tagName)) {
        const inner = buttonEl.querySelector('button, [role="combobox"], [role="button"]');
        if (inner) clickTarget = inner;
      }
      clickTarget.scrollIntoView({ block: 'center' });
      await delay(200); clickTarget.click(); await delay(CFG.DROPDOWN_WAIT);
      let ok = 0;
      for (const v of values) {
        const opts = document.querySelectorAll('[role="option"],[class*="option"]');
        for (const o of opts) {
          const r = o.getBoundingClientRect();
          if (r.width > 0 && r.height > 0 && o.textContent.trim().toLowerCase().includes(v.toLowerCase())) {
            o.click(); ok++; await delay(300); break;
          }
        }
      }
      document.body.click(); await delay(200);
      return ok > 0;
    } catch (e) { document.body.click(); return false; }
  }

  // ============================================================
  // 🚀 STEPS
  // ============================================================
  async function stepNewCase() {
    setProgress(0);
    log('<strong>📌 Opening New Case...</strong>', true);
    if (document.querySelector('[data-testid="form-submit"]')) {
      log('ℹ️ Form already open.'); return true;
    }
    let btn = null;
    document.querySelectorAll('button').forEach(b => {
      if (b.textContent.trim() === 'New Case') btn = b;
    });
    if (!btn) { log('❌ "New Case" not found.'); return false; }
    btn.click(); log('✅ Clicked. Waiting...');
    for (let i = 0; i < 10; i++) {
      await delay(500);
      if (document.querySelector('[data-testid="form-submit"]')) { log('✅ Form loaded.'); return true; }
    }
    log('⏳ Still loading...'); return true;
  }

  async function stepSelectType() {
    setProgress(1);
    log('<strong>📌 Checking record type...</strong>');
    const btn = document.querySelector('#caseRecordTypeDropdown,[data-testid="case-record-type-dropdown"]');
    if (!btn) { log('❌ Not found.'); return false; }
    if (btn.textContent.trim().includes('SSPA ABOS Optimization Case')) {
      log('✅ Already selected.'); return true;
    }
    const ok = await fillDropdown(btn, 'SSPA ABOS Optimization Case');
    if (ok) { log('✅ Selected.'); await delay(CFG.STEP_DELAY); return true; }
    log('❌ Failed.'); return false;
  }

  async function stepFillForm() {
    setProgress(2);
    if (!Object.keys(parsedData).length) { log('❌ No data! Parse first.', true); return; }
    log('<strong>📌 Filling...</strong>');
    log('─'.repeat(30));
    let filled = 0, failed = 0, skipped = 0;
    const failedFields = [];

    for (const f of DOM_MAP) {
      const val = parsedData[f.key];
      if (!val) { skipped++; continue; }
      await delay(CFG.FILL_DELAY);
      let el = findElement(f);
      if (!el) {
        for (let retry = 0; retry < 3; retry++) {
          await delay(1000);
          el = findElement(f);
          if (el) break;
        }
      }
      if (!el) {
        log(\`❌ <span style="color:#ff6b6b">\${f.key}</span>: element not found\`);
        failedFields.push(f.key);
        failed++;
        continue;
      }
      let ok = false;
      if (f.type === 'text' || f.type === 'textarea') ok = fillTextInput(el, val);
      else if (f.type === 'dropdown') ok = await fillDropdown(el, val);
      else if (f.type === 'multiselect') ok = await fillMultiSelect(el, val);

      if (ok) {
        const s = val.length > 30 ? val.substring(0, 30) + '...' : val;
        log(\`✅ <span style="color:#69f0ae">\${f.key}</span> → "\${s}"\`);
        filled++;
      } else {
        log(\`⚠️ <span style="color:#ffd740">\${f.key}</span>: fill failed\`);
        failedFields.push(f.key);
        failed++;
      }
    }

    log('─'.repeat(30));
    log(\`<strong>📊 ✅\${filled} │ ❌\${failed} │ ⏭\${skipped}</strong>\`);
    if (failed) log('<em style="color:#ff9900;">💡 Failed fields → manual input</em>');
    showManualReminder(failedFields);
    setProgress(3);
  }

  // ============================================================
  // ⚠️ POST-FILL MANUAL REMINDER
  // ============================================================
  function showManualReminder(failedFields = []) {
    const manualItems = [];

    const dueDateValue = parsedData.dueDate;
    if (dueDateValue) {
      manualItems.push(\`Due Date — <strong style="color:#ff6b6b;">\${dueDateValue}</strong>\`);
    } else {
      manualItems.push('Due Date');
    }

    const optType = (parsedData.optimizationType || '').trim().toUpperCase();
    if (optType === '' || optType.includes('OPT')) {
      manualItems.push('Update the optimization category. Tick the levers you have suggested.');
    }

    for (const f of failedFields) {
      const readable = f.replace(/([A-Z])/g, ' $1').replace(/^./, s => s.toUpperCase()).trim();
      manualItems.push(\`<strong>\${readable}</strong> (auto-fill failed)\`);
    }

    if (manualItems.length === 0) {
      log(\`
        <div style="
          margin-top:6px; padding:10px 12px;
          background:#00a651; color:#fff;
          border-radius:8px; font-size:12px;
        ">
          <strong>✅ All fields filled successfully!</strong><br>
          <span>👉 Please review and click <strong>Submit</strong>.</span>
        </div>
      \`);
      return;
    }

    const items = manualItems.map(item => \`<li>\${item}</li>\`).join('');

    log(\`
      <div style="
        margin-top:6px; padding:10px 12px;
        background:#ff9900; color:#000;
        border-radius:8px; font-size:12px;
        line-height:1.7;
      ">
        <strong style="font-size:13px;">⚠️ MANUAL ACTION REQUIRED</strong><br>
        <span>Please manually fill/verify the following before submitting:</span>
        <ol style="margin:6px 0 4px 18px; padding:0;">
          \${items}
        </ol>
        <span>👉 After confirming all fields, click <strong>Submit</strong>.</span>
      </div>
    \`);
  }

  async function runFullFlow() {
    if (!Object.keys(parsedData).length) { alert('⚠️ Parse first! (Step 1)'); return; }
    const s0 = await stepNewCase();
    if (!s0) return;
    await delay(CFG.STEP_DELAY);
    const s1 = await stepSelectType();
    if (!s1) return;
    await delay(CFG.STEP_DELAY);
    await stepFillForm();
    log('<strong>🏁 Done! Review the reminder above ☝️</strong>');
  }

  // ============================================================
  // 🚀 INIT
  // ============================================================
  const initCheck = setInterval(() => {
    if (document.body) { clearInterval(initCheck); createPanel(); }
  }, 1000);

})();
