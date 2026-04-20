// ==UserScript==
// @name         SS Case Log Auto-Filler v4.8
// @namespace    http://tampermonkey.net/
// @version      4.9
// @description  v4.8: Fix Brand Name + Primary Goal missing from PARSE_MAP & DOM_MAP.
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
  // PARSE MAP (longest pattern first)
  // v4.8: Added "Brand Name" and standalone "Primary Goal"
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
    { pattern: "Primary Goal",              key: "primaryGoalConsideration" },
    { pattern: "Case Description",           key: "caseDescription" },
    { pattern: "Advertiser Type",            key: "advertiserType" },
    { pattern: "Account Vertical",           key: "accountVertical" },
    { pattern: "Rodeo order CFID",           key: "rodeoCfId" },
    { pattern: "SF account ID",              key: "sfAccountId" },
    { pattern: "Submitting team",            key: "submittingTeam" },
    { pattern: "Advertiser ID",              key: "advertiserId" },
    { pattern: "Account Name",              key: "accountName" },
    { pattern: "Brand Name",                key: "brandName" },
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
  // DOM MAP
  // v4.8: Added brandName entry
  // ============================================================
  const DOM_MAP = [
    { key: "assignee",                 selector: '[data-testid="assignee-field"]',                    altSelectors: ['[name="assignee"]', '#assignee'], type: "text" },
    { key: "accountVertical",          selector: '[data-testid="accountVertical-field"]',              altSelectors: [], type: "text" },
    { key: "accountName",              selector: '[data-testid="accountName-field"]',                  altSelectors: ['[name="accountName"]'], type: "text" },
    { key: "brandName",                selector: '[data-testid="brandName-field"]',                    altSelectors: ['[data-testid="brand-name-field"]', '[data-testid="brand.name-field"]', '[name="brandName"]', '[id*="brandName"]', '[id*="brand-name"]'], type: "text" },
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
  // UI — DRAGGABLE + MINIMIZABLE PANEL
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
    const startTop = savedPos ? savedPos.top : '10px';
    const startLeft = savedPos ? savedPos.left : '';
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

    var panelHTML = '';
    panelHTML += '<div id="ss-header" style="';
    panelHTML += 'background:#1a1a2e; padding:10px 14px; cursor:move; user-select:none;';
    panelHTML += 'display:flex; justify-content:space-between; align-items:center;';
    panelHTML += 'border-bottom:1px solid #333; flex-shrink:0;';
    panelHTML += '">';
    panelHTML += '<div style="display:flex; align-items:center; gap:8px;">';
    panelHTML += '<span style="font-size:16px;">🛠</span>';
    panelHTML += '<strong style="font-size:14px;">SS Case Filler</strong>';
    panelHTML += '<span style="font-size:10px;color:#888;background:#333;padding:1px 6px;border-radius:8px;">v4.8</span>';
    panelHTML += '</div>';
    panelHTML += '<div style="display:flex; gap:6px; align-items:center;">';
    panelHTML += '<span id="ss-btn-minimize" title="Minimize to icon" style="cursor:pointer;font-size:14px;padding:2px 6px;border-radius:4px;background:#333;">━</span>';
    panelHTML += '<span id="ss-btn-resize" title="Collapse panel" style="cursor:pointer;font-size:14px;padding:2px 6px;border-radius:4px;background:#333;">🔽</span>';
    panelHTML += '</div>';
    panelHTML += '</div>';
    panelHTML += '<div id="ss-body" style="padding:14px; overflow-y:auto; flex:1;">';
    panelHTML += '<div id="ss-paste-section">';
    panelHTML += '<div style="background:#ff9900;color:#000;padding:5px 10px;border-radius:6px;font-weight:bold;margin-bottom:8px;font-size:12px;">';
    panelHTML += '1️⃣ PASTE YOUR SS DVA CASE CREATION INFO</div>';
    panelHTML += '<textarea id="ss-input" placeholder="Paste your case log here..." style="';
    panelHTML += 'width:100%; height:170px; background:#1a2332; color:#e0e0e0;';
    panelHTML += 'border:1px solid #444; border-radius:6px; padding:8px;';
    panelHTML += 'font-family:monospace; font-size:11px; resize:vertical;';
    panelHTML += 'box-sizing:border-box;"></textarea>';
    panelHTML += '<button id="ss-parse-btn" style="';
    panelHTML += 'width:100%; padding:10px; margin-top:6px; border:none;';
    panelHTML += 'border-radius:6px; background:#ff9900; color:#000;';
    panelHTML += 'font-weight:bold; cursor:pointer; font-size:13px;';
    panelHTML += '">📋 Parse & Preview</button>';
    panelHTML += '</div>';
    panelHTML += '<div id="ss-preview-section" style="display:none; margin-top:8px;">';
    panelHTML += '<div style="background:#00a651;color:#fff;padding:5px 10px;border-radius:6px;font-weight:bold;margin-bottom:8px;font-size:12px;">';
    panelHTML += '✅ PARSED DATA PREVIEW</div>';
    panelHTML += '<div id="ss-preview" style="';
    panelHTML += 'background:#1a2332; padding:8px; border-radius:6px;';
    panelHTML += 'max-height:180px; overflow-y:auto; font-size:11px; line-height:1.8;';
    panelHTML += '"></div>';
    panelHTML += '<button id="ss-edit-btn" style="';
    panelHTML += 'width:100%; padding:6px; margin-top:6px; border:1px solid #555;';
    panelHTML += 'border-radius:6px; background:transparent; color:#aaa;';
    panelHTML += 'cursor:pointer; font-size:11px;';
    panelHTML += '">✏️ Back to Edit</button>';
    panelHTML += '</div>';
    panelHTML += '<div style="margin-top:10px;">';
    panelHTML += '<div style="background:#4fc3f7;color:#000;padding:5px 10px;border-radius:6px;font-weight:bold;margin-bottom:8px;font-size:12px;">';
    panelHTML += '2️⃣ CREATE NEW CASE</div>';
    panelHTML += '<div style="display:flex;gap:4px;margin-bottom:8px;">';
    panelHTML += '<div id="ss-p0" style="flex:1;text-align:center;padding:3px;border-radius:4px;background:#1a2332;font-size:10px;">New Case</div>';
    panelHTML += '<div id="ss-p1" style="flex:1;text-align:center;padding:3px;border-radius:4px;background:#1a2332;font-size:10px;">Record Type</div>';
    panelHTML += '<div id="ss-p2" style="flex:1;text-align:center;padding:3px;border-radius:4px;background:#1a2332;font-size:10px;">Fill Form</div>';
    panelHTML += '</div>';
    panelHTML += '<button id="ss-create-btn" style="';
    panelHTML += 'width:100%; padding:12px; border:none; border-radius:6px;';
    panelHTML += 'background:#00a651; color:#fff; font-weight:bold;';
    panelHTML += 'cursor:pointer; font-size:14px;';
    panelHTML += '">🚀 Create New Case</button>';
    panelHTML += '<button id="ss-fill-only-btn" style="';
    panelHTML += 'width:100%; padding:7px; margin-top:6px; border:1px solid #00a651;';
    panelHTML += 'border-radius:6px; background:transparent; color:#00a651;';
    panelHTML += 'cursor:pointer; font-size:11px;';
    panelHTML += '">✏️ Fill Only (form already open)</button>';
    panelHTML += '</div>';
    panelHTML += '<div id="ss-log" style="';
    panelHTML += 'background:#1a2332; padding:8px; border-radius:6px;';
    panelHTML += 'margin-top:10px; max-height:160px; overflow-y:auto;';
    panelHTML += 'font-size:11px; line-height:1.6; display:none;';
    panelHTML += '"></div>';
    panelHTML += '</div>';

    panel.innerHTML = panelHTML;
    document.body.appendChild(panel);

    // DRAG: Main Panel
    var isDragging = false, dragOX = 0, dragOY = 0;
    var header = document.getElementById('ss-header');
    header.addEventListener('mousedown', function(e) {
      if (e.target.id === 'ss-btn-minimize' || e.target.id === 'ss-btn-resize') return;
      isDragging = true;
      var rect = panel.getBoundingClientRect();
      dragOX = e.clientX - rect.left;
      dragOY = e.clientY - rect.top;
      panel.style.right = 'auto';
      panel.style.transition = 'none';
      e.preventDefault();
    });
    document.addEventListener('mousemove', function(e) {
      if (!isDragging) return;
      var x = Math.max(0, Math.min(e.clientX - dragOX, window.innerWidth - panel.offsetWidth));
      var y = Math.max(0, Math.min(e.clientY - dragOY, window.innerHeight - 40));
      panel.style.left = x + 'px';
      panel.style.top = y + 'px';
    });
    document.addEventListener('mouseup', function() {
      if (isDragging) {
        isDragging = false;
        panel.style.transition = '';
        localStorage.setItem('ss-panel-pos', JSON.stringify({ top: panel.style.top, left: panel.style.left }));
      }
    });

    // DRAG: Mini Button
    var miniDrag = false, miniOX = 0, miniOY = 0, miniMoved = false;
    mini.addEventListener('mousedown', function(e) {
      miniDrag = true; miniMoved = false;
      miniOX = e.clientX - mini.getBoundingClientRect().left;
      miniOY = e.clientY - mini.getBoundingClientRect().top;
      mini.style.transition = 'none';
      e.preventDefault();
    });
    document.addEventListener('mousemove', function(e) {
      if (!miniDrag) return;
      miniMoved = true;
      mini.style.left = (e.clientX - miniOX) + 'px';
      mini.style.top = (e.clientY - miniOY) + 'px';
      mini.style.right = 'auto';
    });
    document.addEventListener('mouseup', function() {
      if (miniDrag) {
        miniDrag = false;
        mini.style.transition = '';
        if (!miniMoved) { mini.style.display = 'none'; panel.style.display = 'flex'; }
      }
    });

    // MINIMIZE
    document.getElementById('ss-btn-minimize').addEventListener('click', function() {
      panel.style.display = 'none';
      mini.style.display = 'block';
      var rect = panel.getBoundingClientRect();
      mini.style.top = rect.top + 'px';
      mini.style.left = rect.left + 'px';
      mini.style.right = 'auto';
    });

    // COLLAPSE
    var collapsed = false;
    document.getElementById('ss-btn-resize').addEventListener('click', function() {
      collapsed = !collapsed;
      document.getElementById('ss-body').style.display = collapsed ? 'none' : 'block';
      document.getElementById('ss-btn-resize').textContent = collapsed ? '🔼' : '🔽';
    });

    // BUTTONS
    document.getElementById('ss-parse-btn').addEventListener('click', parseInput);
    document.getElementById('ss-edit-btn').addEventListener('click', function() {
      document.getElementById('ss-paste-section').style.display = 'block';
      document.getElementById('ss-preview-section').style.display = 'none';
    });
    document.getElementById('ss-create-btn').addEventListener('click', runFullFlow);
    document.getElementById('ss-fill-only-btn').addEventListener('click', function() { stepFillForm(); });

    var saved = localStorage.getItem('ss-last-input');
    if (saved) document.getElementById('ss-input').value = saved;
  }

  // ============================================================
  // PARSE
  // ============================================================
  function parseInput() {
    var raw = document.getElementById('ss-input').value.trim();
    if (!raw) { alert('Please paste your case info first!'); return; }
    localStorage.setItem('ss-last-input', raw);
    parsedData = {};
    var lines = raw.split('\n');
    for (var li = 0; li < lines.length; li++) {
      var line = lines[li].trim();
      if (!line || line.startsWith('===')) continue;
      for (var pi = 0; pi < PARSE_MAP.length; pi++) {
        var pm = PARSE_MAP[pi];
        if (line.toLowerCase().startsWith(pm.pattern.toLowerCase())) {
          var value = line.substring(pm.pattern.length).replace(/^[\s:：;；]+/, '').trim();
          value = cleanValue(value);
          if (value) {
            if (!parsedData[pm.key]) parsedData[pm.key] = value;
          }
          break;
        }
      }
    }
    showPreview();
  }

  function cleanValue(val) {
    if (!val) return '';
    val = val.replace(/\s*\((?:Order Level ID|Rodeo Account Level ID|14 business days[^)]*|mandatory|keep it blank|Proactive support[^)]*|= Case creation date|Today \+ 13 calendar days)\)\s*/gi, '').trim();
    var lower = val.toLowerCase();
    for (var i = 0; i < IGNORE_VALUES.length; i++) {
      if (lower.indexOf(IGNORE_VALUES[i].toLowerCase()) !== -1) return '';
    }
    if (val.indexOf('  ') !== -1) val = val.split(/\s{2,}/)[0].trim();
    return val;
  }

  function showPreview() {
    var el = document.getElementById('ss-preview');
    var count = Object.keys(parsedData).length;
    var shownKeys = {};
    var html = '<div style="color:#69f0ae;margin-bottom:4px;"><strong>✅ ' + count + ' fields parsed</strong></div>';
    for (var i = 0; i < PARSE_MAP.length; i++) {
      var pm = PARSE_MAP[i];
      if (shownKeys[pm.key]) continue;
      shownKeys[pm.key] = true;
      var v = parsedData[pm.key];
      if (v) {
        html += '<div><span style="color:#888;">' + pm.pattern + ':</span> <span style="color:#69f0ae;">' + v + '</span></div>';
      } else {
        html += '<div><span style="color:#444;">' + pm.pattern + ': (skip)</span></div>';
      }
    }
    el.innerHTML = html;
    document.getElementById('ss-paste-section').style.display = 'none';
    document.getElementById('ss-preview-section').style.display = 'block';
  }

  // ============================================================
  // HELPERS
  // ============================================================
  function log(msg, reset) {
    var el = document.getElementById('ss-log');
    el.style.display = 'block';
    if (reset) el.innerHTML = '';
    el.innerHTML += msg + '<br>';
    el.scrollTop = el.scrollHeight;
  }

  function setProgress(n) {
    for (var i = 0; i <= 2; i++) {
      var el = document.getElementById('ss-p' + i);
      el.style.background = i < n ? '#00a651' : i === n ? '#ff9900' : '#1a2332';
      el.style.color = i === n ? '#000' : '#fff';
    }
  }

  function delay(ms) { return new Promise(function(r) { setTimeout(r, ms); }); }

  function normalize(str) {
    return str.toLowerCase().replace(/[\s_\-]+/g, '');
  }

  // ============================================================
  // FIND ELEMENT
  // ============================================================
  function findElement(field) {
    var el = document.querySelector(field.selector);
    if (el) return el;

    if (field.altSelectors && field.altSelectors.length) {
      for (var a = 0; a < field.altSelectors.length; a++) {
        try { el = document.querySelector(field.altSelectors[a]); } catch (e) { continue; }
        if (el) {
          log('  ↪ <span style="color:#4fc3f7;">' + field.key + '</span>: found via alt selector');
          return el;
        }
      }
    }

    var normalizedKey = normalize(field.key);
    var labels = document.querySelectorAll('label, [class*="label"], [class*="Label"]');
    for (var l = 0; l < labels.length; l++) {
      var lbl = labels[l];
      var normalizedLabel = normalize(lbl.textContent.trim());
      if (normalizedLabel.indexOf(normalizedKey) !== -1 || normalizedKey.indexOf(normalizedLabel) !== -1) {
        var parent = lbl.closest('[class*="field"], [class*="form-group"], [class*="row"]') || lbl.parentElement;
        if (parent) {
          el = parent.querySelector('input, textarea, select, button, [role="combobox"], [role="listbox"]');
          if (el) {
            log('  ↪ <span style="color:#4fc3f7;">' + field.key + '</span>: found via label search');
            return el;
          }
        }
      }
    }

    if (field.type === 'dropdown' || field.type === 'multiselect') {
      var readable = field.key.replace(/([A-Z])/g, ' $1').toLowerCase().trim();
      var words = readable.split(/\s+/).filter(function(w) { return w.length > 3; });
      var allFields = document.querySelectorAll('[class*="field"], [class*="form-group"], [class*="form-row"]');
      for (var f = 0; f < allFields.length; f++) {
        var container = allFields[f];
        var text = container.textContent.toLowerCase();
        if (words.length >= 2 && words.every(function(w) { return text.indexOf(w) !== -1; })) {
          var trigger = container.querySelector(
            'button, select, [role="combobox"], [role="button"], [role="listbox"], [class*="trigger"], [class*="dropdown"], [class*="select"]'
          );
          if (trigger) {
            log('  ↪ <span style="color:#ff9900;">' + field.key + '</span>: found via nuclear fallback');
            return trigger;
          }
        }
      }
    }

    return null;
  }

  // ============================================================
  // FILL TEXT
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
      if (el.tagName !== 'INPUT' && el.tagName !== 'TEXTAREA') {
        var inner = el.querySelector('input, textarea');
        if (inner) {
          el = inner;
        } else {
          var ce = el.querySelector('[contenteditable="true"]');
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
      var proto = el.tagName === 'TEXTAREA' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
      var descriptor = Object.getOwnPropertyDescriptor(proto, 'value');
      var setter = descriptor ? descriptor.set : null;
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
  // FILL DROPDOWN
  // ============================================================
  async function fillDropdown(buttonEl, value) {
    try {
      var clickTarget = buttonEl;
      if (buttonEl.tagName !== 'BUTTON' && buttonEl.tagName !== 'SELECT') {
        var inner = buttonEl.querySelector('button, [role="combobox"], [role="button"], [class*="trigger"], [class*="select"]');
        if (inner) clickTarget = inner;
      }
      clickTarget.scrollIntoView({ block: 'center' });
      await delay(200);
      clickTarget.click();
      await delay(CFG.DROPDOWN_WAIT);

      var sels = [
        '[role="option"]', '[role="menuitem"]', '[class*="option"]',
        '[class*="menu-item"]', '[class*="dropdown"] li', '[class*="select"] li',
        '[class*="listbox"] li'
      ];
      var allOpts = [];
      for (var s = 0; s < sels.length; s++) {
        var opts = document.querySelectorAll(sels[s]);
        for (var o = 0; o < opts.length; o++) {
          var r = opts[o].getBoundingClientRect();
          if (r.width > 0 && r.height > 0) allOpts.push(opts[o]);
        }
      }
      var lv = value.toLowerCase().trim();
      var match = null;
      for (var i = 0; i < allOpts.length; i++) {
        if (allOpts[i].textContent.trim().toLowerCase() === lv) { match = allOpts[i]; break; }
      }
      if (!match) {
        for (var i2 = 0; i2 < allOpts.length; i2++) {
          if (allOpts[i2].textContent.trim().toLowerCase().startsWith(lv)) { match = allOpts[i2]; break; }
        }
      }
      if (!match) {
        for (var i3 = 0; i3 < allOpts.length; i3++) {
          if (allOpts[i3].textContent.trim().toLowerCase().indexOf(lv) !== -1) { match = allOpts[i3]; break; }
        }
      }
      if (match) { match.click(); await delay(300); return true; }

      document.body.click();
      await delay(500);
      clickTarget.click();
      await delay(CFG.DROPDOWN_WAIT + 200);
      allOpts = [];
      for (var s2 = 0; s2 < sels.length; s2++) {
        var opts2 = document.querySelectorAll(sels[s2]);
        for (var o2 = 0; o2 < opts2.length; o2++) {
          var r2 = opts2[o2].getBoundingClientRect();
          if (r2.width > 0 && r2.height > 0) allOpts.push(opts2[o2]);
        }
      }
      match = null;
      for (var j = 0; j < allOpts.length; j++) {
        if (allOpts[j].textContent.trim().toLowerCase() === lv) { match = allOpts[j]; break; }
      }
      if (!match) {
        for (var j2 = 0; j2 < allOpts.length; j2++) {
          if (allOpts[j2].textContent.trim().toLowerCase().startsWith(lv)) { match = allOpts[j2]; break; }
        }
      }
      if (!match) {
        for (var j3 = 0; j3 < allOpts.length; j3++) {
          if (allOpts[j3].textContent.trim().toLowerCase().indexOf(lv) !== -1) { match = allOpts[j3]; break; }
        }
      }
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
      var values = value.split(',').map(function(v) { return v.trim(); }).filter(Boolean);
      if (!values.length) return false;
      var clickTarget = buttonEl;
      if (buttonEl.tagName !== 'BUTTON' && buttonEl.tagName !== 'SELECT') {
        var inner = buttonEl.querySelector('button, [role="combobox"], [role="button"]');
        if (inner) clickTarget = inner;
      }
      clickTarget.scrollIntoView({ block: 'center' });
      await delay(200); clickTarget.click(); await delay(CFG.DROPDOWN_WAIT);
      var ok = 0;
      for (var vi = 0; vi < values.length; vi++) {
        var opts = document.querySelectorAll('[role="option"],[class*="option"]');
        for (var oi = 0; oi < opts.length; oi++) {
          var r = opts[oi].getBoundingClientRect();
          if (r.width > 0 && r.height > 0 && opts[oi].textContent.trim().toLowerCase().indexOf(values[vi].toLowerCase()) !== -1) {
            opts[oi].click(); ok++; await delay(300); break;
          }
        }
      }
      document.body.click(); await delay(200);
      return ok > 0;
    } catch (e) { document.body.click(); return false; }
  }

  // ============================================================
  // STEPS
  // ============================================================
  async function stepNewCase() {
    setProgress(0);
    log('<strong>📌 Opening New Case...</strong>', true);
    if (document.querySelector('[data-testid="form-submit"]')) {
      log('ℹ️ Form already open.'); return true;
    }
    var btn = null;
    var buttons = document.querySelectorAll('button');
    for (var b = 0; b < buttons.length; b++) {
      if (buttons[b].textContent.trim() === 'New Case') { btn = buttons[b]; break; }
    }
    if (!btn) { log('❌ "New Case" not found.'); return false; }
    btn.click(); log('✅ Clicked. Waiting...');
    for (var i = 0; i < 10; i++) {
      await delay(500);
      if (document.querySelector('[data-testid="form-submit"]')) { log('✅ Form loaded.'); return true; }
    }
    log('⏳ Still loading...'); return true;
  }

  async function stepSelectType() {
    setProgress(1);
    log('<strong>📌 Checking record type...</strong>');
    var btn = document.querySelector('#caseRecordTypeDropdown,[data-testid="case-record-type-dropdown"]');
    if (!btn) { log('❌ Not found.'); return false; }
    if (btn.textContent.trim().indexOf('SSPA ABOS Optimization Case') !== -1) {
      log('✅ Already selected.'); return true;
    }
    var ok = await fillDropdown(btn, 'SSPA ABOS Optimization Case');
    if (ok) { log('✅ Selected.'); await delay(CFG.STEP_DELAY); return true; }
    log('❌ Failed.'); return false;
  }

  async function stepFillForm() {
    setProgress(2);
    if (!Object.keys(parsedData).length) { log('❌ No data! Parse first.', true); return; }
    log('<strong>📌 Filling...</strong>');
    log('──────────────────────────────');
    var filled = 0, failed = 0, skipped = 0;
    var failedFields = [];

    for (var fi = 0; fi < DOM_MAP.length; fi++) {
      var f = DOM_MAP[fi];
      var val = parsedData[f.key];
      if (!val) { skipped++; continue; }
      await delay(CFG.FILL_DELAY);
      var el = findElement(f);
      if (!el) {
        for (var retry = 0; retry < 3; retry++) {
          await delay(1000);
          el = findElement(f);
          if (el) break;
        }
      }
      if (!el) {
        log('❌ <span style="color:#ff6b6b">' + f.key + '</span>: element not found');
        failedFields.push(f.key);
        failed++;
        continue;
      }
      var ok = false;
      if (f.type === 'text' || f.type === 'textarea') ok = fillTextInput(el, val);
      else if (f.type === 'dropdown') ok = await fillDropdown(el, val);
      else if (f.type === 'multiselect') ok = await fillMultiSelect(el, val);

      if (ok) {
        var s = val.length > 30 ? val.substring(0, 30) + '...' : val;
        log('✅ <span style="color:#69f0ae">' + f.key + '</span> → "' + s + '"');
        filled++;
      } else {
        log('⚠️ <span style="color:#ffd740">' + f.key + '</span>: fill failed');
        failedFields.push(f.key);
        failed++;
      }
    }

    log('──────────────────────────────');
    log('<strong>📊 ✅' + filled + ' │ ❌' + failed + ' │ ⏭' + skipped + '</strong>');
    if (failed) log('<em style="color:#ff9900;">💡 Failed fields → manual input</em>');
    showManualReminder(failedFields);
    setProgress(3);
  }

  // ============================================================
  // POST-FILL MANUAL REMINDER
  // ============================================================
  function showManualReminder(failedFields) {
    if (!failedFields) failedFields = [];
    var manualItems = [];

    var dueDateValue = parsedData.dueDate;
    if (dueDateValue) {
      manualItems.push('Due Date — <strong style="color:#ff6b6b;">' + dueDateValue + '</strong>');
    } else {
      manualItems.push('Due Date');
    }

    var optType = (parsedData.optimizationType || '').trim().toUpperCase();
    if (optType === '' || optType.indexOf('OPT') !== -1) {
      manualItems.push('Update the optimization category. Tick the levers you have suggested.');
    }

    for (var i = 0; i < failedFields.length; i++) {
      var readable = failedFields[i].replace(/([A-Z])/g, ' $1').replace(/^./, function(s) { return s.toUpperCase(); }).trim();
      manualItems.push('<strong>' + readable + '</strong> (auto-fill failed)');
    }

    if (manualItems.length === 0) {
      log('<div style="margin-top:6px; padding:10px 12px; background:#00a651; color:#fff; border-radius:8px; font-size:12px;">' +
        '<strong>✅ All fields filled successfully!</strong><br>' +
        '<span>👉 Please review and click <strong>Submit</strong>.</span></div>');
      return;
    }

    var items = '';
    for (var j = 0; j < manualItems.length; j++) {
      items += '<li>' + manualItems[j] + '</li>';
    }

    log('<div style="margin-top:6px; padding:10px 12px; background:#ff9900; color:#000; border-radius:8px; font-size:12px; line-height:1.7;">' +
      '<strong style="font-size:13px;">⚠️ MANUAL ACTION REQUIRED</strong><br>' +
      '<span>Please manually fill/verify the following before submitting:</span>' +
      '<ol style="margin:6px 0 4px 18px; padding:0;">' + items + '</ol>' +
      '<span>👉 After confirming all fields, click <strong>Submit</strong>.</span></div>');
  }

  async function runFullFlow() {
    if (!Object.keys(parsedData).length) { alert('⚠️ Parse first! (Step 1)'); return; }
    var s0 = await stepNewCase();
    if (!s0) return;
    await delay(CFG.STEP_DELAY);
    var s1 = await stepSelectType();
    if (!s1) return;
    await delay(CFG.STEP_DELAY);
    await stepFillForm();
    log('<strong>🏁 Done! Review the reminder above ☝️</strong>');
  }

  // ============================================================
  // INIT
  // ============================================================
  var initCheck = setInterval(function() {
    if (document.body) { clearInterval(initCheck); createPanel(); }
  }, 1000);

})();
