// ==UserScript==
// @name         SS Case Log Auto-Filler v5.0
// @namespace    http://tampermonkey.net/
// @version      5.0
// @description  v5.0: Fixed all 4 dropdown fields, lever-to-category mapping, enhanced multiselect, React portal support, debug inspector.
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
  // PARSE MAP
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
    { key: "optimizationType",
      selector: '[data-testid="optimizationType-field"]',
      altSelectors: [
        '[data-testid="optimization-type-field"]',
        '[data-testid="optimizationType"] button',
        '[name="optimizationType"]',
        '[id*="optimizationType"]',
        '[id*="optimization-type"]',
        '[aria-label*="Optimization Type"]',
        '[aria-label*="optimization type"]'
      ],
      type: "dropdown" },
    { key: "optimizationDelivery",
      selector: '[data-testid="optimizationDelivery-field"]',
      altSelectors: [
        '[data-testid="optimization-delivery-field"]',
        '[data-testid="optimizationDelivery"] button',
        '[name="optimizationDelivery"]',
        '[id*="optimizationDelivery"]',
        '[id*="optimization-delivery"]',
        '[aria-label*="Optimization Delivery"]',
        '[aria-label*="optimization delivery"]'
      ],
      type: "dropdown" },
    { key: "optimizationMarketplace",
      selector: '[data-testid="advertiser.marketplaceId-field"]',
      altSelectors: [
        '[data-testid="optimizationMarketplace-field"]',
        '[data-testid="optimization-marketplace-field"]',
        '[data-testid="marketplaceId-field"]',
        '[data-testid="marketplace-field"]',
        '[name="optimizationMarketplace"]',
        '[name="marketplaceId"]',
        '[id*="optimizationMarketplace"]',
        '[id*="marketplace"]',
        '[aria-label*="Optimization Marketplace"]',
        '[aria-label*="Marketplace"]'
      ],
      type: "dropdown" },
    { key: "optimizationCategories",
      selector: '[data-testid="optimizationCategories-field"]',
      altSelectors: [
        '[data-testid="optimization-categories-field"]',
        '[data-testid="optimizationCategories"] button',
        '[name="optimizationCategories"]',
        '[id*="optimizationCategories"]',
        '[id*="optimization-categories"]',
        '[aria-label*="Optimization Categories"]',
        '[aria-label*="optimization categories"]'
      ],
      type: "multiselect" },
  ];

  // ============================================================
  // LEVER TO CATEGORY REVERSE LOOKUP MAP
  // ============================================================
  const CATEGORY_LEVER_MAP = {
    "Frequency Cap / Twitch - Frequency Cap": [
      "Order Frequency Cap", "Line Item Frequency Cap",
      "Order Frequency Group Campaign Association"
    ],
    "Auto-Optimization": [
      "Order Automated Optimization"
    ],
    "Supply Source": [
      "Line Item Supply Sources"
    ],
    "Bid Change": [
      "Line Item Base Supply Bid", "Line Item Max Supply Bid",
      "Line Item Base CPC", "Line Item Sold CPC", "Line Item Sold CPM",
      "Line Item Bid Strategy", "Order Bid Strategy"
    ],
    "Budget Allocation \u2014 Order to Order": [
      "Order Budget", "Order Budget Cap",
      "Order Flight Budget Amount", "Order Flight Budget Rollover"
    ],
    "Domain Targeting": [
      "Line Item Domain Targeting"
    ],
    "Start Date Change": [
      "Order Start Date Time", "Line Item Start Date",
      "Line Item Flight Start Date", "Order Flight Start Time"
    ],
    "Creative": [
      "Line Item Creative Association", "Line Item Creative Status",
      "Line Item Creative Approval Status", "Line Item Creative Sizes",
      "Line Item Placement Position"
    ],
    "Twitch \u2014 Pacing Profile": [
      "Line Item Pacing Profile", "Line Item Catchup Boost"
    ],
    "ASIN / Pixel Updates": [
      "Order ASINs", "Order Conversion Pixels",
      "Order Off Amazon Conversions"
    ],
    "Billing Updates": [
      "Order Agency Fee", "Line Item 3P Fees",
      "Line Item In Market Audience Fees",
      "Line Item Automotive Audience Fees",
      "Line Item Amazon Platform Fee"
    ],
    "Budget Allocation \u2014 Within DSP Order": [
      "Line Item Budget", "Line Item Budget Cap",
      "Line Item Budget Cap Amount", "Line Item Budget Type",
      "Line Item Flight Budget", "Line Item Total Impressions",
      "Line Item Projected Spend", "Line Item Budget Cap Date",
      "Line Item Budget Cap Delivery Profile",
      "Line Item Budget Cap Recurrence"
    ],
    "Day Part": [
      "Line Item Targeting Daypart"
    ],
    "Keyword Targeting": [
      "Line Item Keyword Targeting"
    ],
    "Targeting \u2014 Geo": [
      "Line Item Targeting Location"
    ],
    "Book New Order or Line": [
      "Line Item Name", "Line Item Type", "Line Item Status",
      "Order Name", "Order Status", "Order Type"
    ],
    "Targeting \u2014 Segment": [
      "Line Item Audience Targeting", "Line Item Contextual Targeting",
      "Line Item Targeting String", "Line Item Content Rating",
      "Line Item Content Genres", "Line Item Content Categories",
      "Line Item Mobile OS Targeting", "Line Item Mobile App Targeting",
      "Line Item Mobile Environment Targeting",
      "Line Item Audience Targeting Match Type",
      "Line Item Third Party Custom Segments",
      "Line Item Third Party Custom Predicts",
      "Line Item Third Party Standard Predicts",
      "Line Item Third Party Authentic Attention",
      "Line Item Third Party Context Control Avoidance",
      "Line Item Third Party Context Control Targeting",
      "Line Item Third Party Authentic Brand Safety",
      "Line Item Third Party Custom Contextual",
      "Line Item Site Targeting", "Line Item Site Language",
      "Line Item Unified Language", "Line Item Product Categories"
    ],
    "Pacing Profile": [
      "Line Item Pacing Profile", "Line Item Delivery Priority",
      "Line Item Catchup Boost", "Line Item Optimization Type"
    ],
    "End Date Change": [
      "Order End Date Time", "Line Item End Date",
      "Line Item Flight End Date", "Order Flight End Time",
      "Order Flight"
    ],
    "Other": [
      "Line Item Amazon Viewability",
      "Line Item Third Party Viewability MRC",
      "Line Item Third Party IAS Viewability",
      "Line Item Third Party DV Viewability",
      "Order Goals", "Order General Goals",
      "Line Item Deals", "Line Item Inventory Groups Deals",
      "Order Deals"
    ],
    "Actualization": [
      "Line Item Status", "Order Status"
    ],
    "Twitch \u2014 Book New Order": [
      "Line Item Name", "Line Item Type"
    ],
    "Twitch \u2014 Targeting Update": [
      "Line Item Audience Targeting", "Line Item Contextual Targeting"
    ]
  };

  const LEVER_TO_CATEGORY = {};
  Object.keys(CATEGORY_LEVER_MAP).forEach(function(category) {
    var levers = CATEGORY_LEVER_MAP[category];
    for (var i = 0; i < levers.length; i++) {
      var leverKey = levers[i].toLowerCase().trim();
      if (!LEVER_TO_CATEGORY[leverKey]) {
        LEVER_TO_CATEGORY[leverKey] = [];
      }
      LEVER_TO_CATEGORY[leverKey].push(category);
    }
  });

  // ============================================================
  // CONFIG
  // ============================================================
  const CFG = { FILL_DELAY: 400, DROPDOWN_WAIT: 800, STEP_DELAY: 1500 };
  let parsedData = {};

  // ============================================================
  // UI PANEL
  // ============================================================
  function createPanel() {
    const mini = document.createElement('div');
    mini.id = 'ss-mini';
    mini.textContent = 'SS';
    mini.title = 'Open SS Case Filler';
    Object.assign(mini.style, {
      position: 'fixed', top: '10px', right: '10px', zIndex: '99999',
      width: '42px', height: '42px', borderRadius: '50%',
      background: '#ff9900', color: '#000', fontSize: '14px', fontWeight: 'bold',
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

    var h = '';
    h += '<div id="ss-header" style="';
    h += 'background:#1a1a2e; padding:10px 14px; cursor:move; user-select:none;';
    h += 'display:flex; justify-content:space-between; align-items:center;';
    h += 'border-bottom:1px solid #333; flex-shrink:0;';
    h += '">';
    h += '<div style="display:flex; align-items:center; gap:8px;">';
    h += '<strong style="font-size:14px;">SS Case Filler</strong>';
    h += '<span style="font-size:10px;color:#888;background:#333;padding:1px 6px;border-radius:8px;">v5.0</span>';
    h += '</div>';
    h += '<div style="display:flex; gap:6px; align-items:center;">';
    h += '<span id="ss-btn-minimize" title="Minimize" style="cursor:pointer;font-size:14px;padding:2px 6px;border-radius:4px;background:#333;">-</span>';
    h += '<span id="ss-btn-resize" title="Collapse" style="cursor:pointer;font-size:12px;padding:2px 6px;border-radius:4px;background:#333;">v</span>';
    h += '</div>';
    h += '</div>';
    h += '<div id="ss-body" style="padding:14px; overflow-y:auto; flex:1;">';
    h += '<div id="ss-paste-section">';
    h += '<div style="background:#ff9900;color:#000;padding:5px 10px;border-radius:6px;font-weight:bold;margin-bottom:8px;font-size:12px;">';
    h += '1. PASTE YOUR SS DVA CASE CREATION INFO</div>';
    h += '<textarea id="ss-input" placeholder="Paste your case log here..." style="';
    h += 'width:100%; height:170px; background:#1a2332; color:#e0e0e0;';
    h += 'border:1px solid #444; border-radius:6px; padding:8px;';
    h += 'font-family:monospace; font-size:11px; resize:vertical;';
    h += 'box-sizing:border-box;"></textarea>';
    h += '<button id="ss-parse-btn" style="';
    h += 'width:100%; padding:10px; margin-top:6px; border:none;';
    h += 'border-radius:6px; background:#ff9900; color:#000;';
    h += 'font-weight:bold; cursor:pointer; font-size:13px;';
    h += '">Parse & Preview</button>';
    h += '</div>';
    h += '<div id="ss-preview-section" style="display:none; margin-top:8px;">';
    h += '<div style="background:#00a651;color:#fff;padding:5px 10px;border-radius:6px;font-weight:bold;margin-bottom:8px;font-size:12px;">';
    h += 'PARSED DATA PREVIEW</div>';
    h += '<div id="ss-preview" style="';
    h += 'background:#1a2332; padding:8px; border-radius:6px;';
    h += 'max-height:180px; overflow-y:auto; font-size:11px; line-height:1.8;';
    h += '"></div>';
    h += '<button id="ss-edit-btn" style="';
    h += 'width:100%; padding:6px; margin-top:6px; border:1px solid #555;';
    h += 'border-radius:6px; background:transparent; color:#aaa;';
    h += 'cursor:pointer; font-size:11px;';
    h += '">Back to Edit</button>';
    h += '</div>';
    h += '<div style="margin-top:10px;">';
    h += '<div style="background:#4fc3f7;color:#000;padding:5px 10px;border-radius:6px;font-weight:bold;margin-bottom:8px;font-size:12px;">';
    h += '2. CREATE NEW CASE</div>';
    h += '<div style="display:flex;gap:4px;margin-bottom:8px;">';
    h += '<div id="ss-p0" style="flex:1;text-align:center;padding:3px;border-radius:4px;background:#1a2332;font-size:10px;">New Case</div>';
    h += '<div id="ss-p1" style="flex:1;text-align:center;padding:3px;border-radius:4px;background:#1a2332;font-size:10px;">Record Type</div>';
    h += '<div id="ss-p2" style="flex:1;text-align:center;padding:3px;border-radius:4px;background:#1a2332;font-size:10px;">Fill Form</div>';
    h += '</div>';
    h += '<button id="ss-create-btn" style="';
    h += 'width:100%; padding:12px; border:none; border-radius:6px;';
    h += 'background:#00a651; color:#fff; font-weight:bold;';
    h += 'cursor:pointer; font-size:14px;';
    h += '">Create New Case</button>';
    h += '<button id="ss-fill-only-btn" style="';
    h += 'width:100%; padding:7px; margin-top:6px; border:1px solid #00a651;';
    h += 'border-radius:6px; background:transparent; color:#00a651;';
    h += 'cursor:pointer; font-size:11px;';
    h += '">Fill Only (form already open)</button>';
    h += '<button id="ss-inspect-btn" style="';
    h += 'width:100%; padding:7px; margin-top:6px; border:1px solid #4fc3f7;';
    h += 'border-radius:6px; background:transparent; color:#4fc3f7;';
    h += 'cursor:pointer; font-size:11px;';
    h += '">Inspect Form Fields (Debug)</button>';
    h += '</div>';
    h += '<div id="ss-log" style="';
    h += 'background:#1a2332; padding:8px; border-radius:6px;';
    h += 'margin-top:10px; max-height:160px; overflow-y:auto;';
    h += 'font-size:11px; line-height:1.6; display:none;';
    h += '"></div>';
    h += '</div>';

    panel.innerHTML = h;
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
      document.getElementById('ss-btn-resize').textContent = collapsed ? '^' : 'v';
    });

    // BUTTONS
    document.getElementById('ss-parse-btn').addEventListener('click', parseInput);
    document.getElementById('ss-edit-btn').addEventListener('click', function() {
      document.getElementById('ss-paste-section').style.display = 'block';
      document.getElementById('ss-preview-section').style.display = 'none';
    });
    document.getElementById('ss-create-btn').addEventListener('click', runFullFlow);
    document.getElementById('ss-fill-only-btn').addEventListener('click', function() { stepFillForm(); });

    // DEBUG INSPECTOR
    document.getElementById('ss-inspect-btn').addEventListener('click', function() {
      log('<strong>[INSPECT] Scanning form fields...</strong>', true);
      var allTestIds = document.querySelectorAll('[data-testid]');
      var keywords = ['optim', 'market', 'categ', 'deliv', 'type', 'brand', 'assign', 'status'];
      var found = 0;
      for (var i = 0; i < allTestIds.length; i++) {
        var tid = allTestIds[i].getAttribute('data-testid');
        var tidLower = tid.toLowerCase();
        for (var k = 0; k < keywords.length; k++) {
          if (tidLower.indexOf(keywords[k]) !== -1) {
            var tag = allTestIds[i].tagName.toLowerCase();
            var role = allTestIds[i].getAttribute('role') || '';
            log('[FIELD] <span style="color:#4fc3f7;">' + tid + '</span> &lt;' + tag + '&gt;' + (role ? ' role="' + role + '"' : ''));
            found++;
            break;
          }
        }
      }
      if (!found) log('[WARN] No matching data-testid found.');
      log('<br><strong>[LABELS]</strong>');
      var labels = document.querySelectorAll('label, [class*="label"], [class*="Label"]');
      for (var j = 0; j < labels.length; j++) {
        var txt = labels[j].textContent.trim().toLowerCase();
        for (var k2 = 0; k2 < keywords.length; k2++) {
          if (txt.indexOf(keywords[k2]) !== -1) {
            log('[LABEL] "' + labels[j].textContent.trim() + '"');
            break;
          }
        }
      }
    });

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
          var value = line.substring(pm.pattern.length).replace(/^[\s:;\uFF1A\uFF1B]+/, '').trim();
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
    var html = '<div style="color:#69f0ae;margin-bottom:4px;"><strong>[OK] ' + count + ' fields parsed</strong></div>';
    for (var i = 0; i < PARSE_MAP.length; i++) {
      var pm = PARSE_MAP[i];
      if (shownKeys[pm.key]) continue;
      shownKeys[pm.key] = true;
      var v = parsedData[pm.key];
      if (v) {
        if (pm.key === 'optimizationCategories') {
          var rawVals = smartSplit(v);
          var resolved = resolveToCategories(rawVals);
          html += '<div><span style="color:#888;">' + pm.pattern + ':</span> <span style="color:#69f0ae;">' + v + '</span>';
          if (resolved.length && resolved.join(', ') !== v) {
            html += '<br><span style="color:#4fc3f7;margin-left:12px;">-> Mapped: ' + resolved.join(', ') + '</span>';
          }
          html += '</div>';
        } else {
          html += '<div><span style="color:#888;">' + pm.pattern + ':</span> <span style="color:#69f0ae;">' + v + '</span></div>';
        }
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
  // CATEGORY RESOLUTION
  // ============================================================
  function smartSplit(raw) {
    if (!raw) return [];
    var delimiter = raw.indexOf(';') !== -1 ? ';' : ',';
    var parts = raw.split(delimiter);
    return parts.map(function(v) { return v.trim(); }).filter(function(v) { return v.length > 0; });
  }

  function resolveToCategories(rawValues) {
    var categoryNames = Object.keys(CATEGORY_LEVER_MAP);
    var resolved = {};

    for (var i = 0; i < rawValues.length; i++) {
      var val = rawValues[i].trim();
      var valLower = val.toLowerCase();
      var matched = false;

      // Check 1: Already a category name
      for (var c = 0; c < categoryNames.length; c++) {
        if (categoryNames[c].toLowerCase() === valLower) {
          resolved[categoryNames[c]] = true;
          matched = true;
          break;
        }
      }
      if (matched) continue;

      // Check 2: Exact lever lookup
      if (LEVER_TO_CATEGORY[valLower]) {
        var cats = LEVER_TO_CATEGORY[valLower];
        for (var j = 0; j < cats.length; j++) {
          resolved[cats[j]] = true;
        }
        matched = true;
      }
      if (matched) continue;

      // Check 3: Fuzzy lever match
      var leverKeys = Object.keys(LEVER_TO_CATEGORY);
      for (var k = 0; k < leverKeys.length; k++) {
        if (leverKeys[k].indexOf(valLower) !== -1 || valLower.indexOf(leverKeys[k]) !== -1) {
          var fuzzyCats = LEVER_TO_CATEGORY[leverKeys[k]];
          for (var m = 0; m < fuzzyCats.length; m++) {
            resolved[fuzzyCats[m]] = true;
          }
          matched = true;
          break;
        }
      }
      if (matched) continue;

      // Check 4: Fuzzy category name match
      for (var c2 = 0; c2 < categoryNames.length; c2++) {
        var catLower = categoryNames[c2].toLowerCase();
        if (catLower.indexOf(valLower) !== -1 || valLower.indexOf(catLower) !== -1) {
          resolved[categoryNames[c2]] = true;
          matched = true;
          break;
        }
      }

      if (!matched) {
        resolved[val] = true;
      }
    }

    return Object.keys(resolved);
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
          log('  > <span style="color:#4fc3f7;">' + field.key + '</span>: found via alt selector');
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
            log('  > <span style="color:#4fc3f7;">' + field.key + '</span>: found via label search');
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
            log('  > <span style="color:#ff9900;">' + field.key + '</span>: found via fallback');
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
  // DROPDOWN HELPERS
  // ============================================================
  async function openDropdownMenu(clickTarget) {
    clickTarget.click();
    await delay(CFG.DROPDOWN_WAIT);
    var opts = collectDropdownOptions();
    if (opts.length > 0) return opts;

    clickTarget.focus();
    clickTarget.dispatchEvent(new KeyboardEvent('keydown', {
      key: 'ArrowDown', code: 'ArrowDown', bubbles: true
    }));
    await delay(CFG.DROPDOWN_WAIT);
    opts = collectDropdownOptions();
    if (opts.length > 0) return opts;

    clickTarget.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
    await delay(150);
    clickTarget.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
    await delay(CFG.DROPDOWN_WAIT);
    return collectDropdownOptions();
  }

  function collectDropdownOptions() {
    var sels = [
      '[role="option"]', '[role="menuitem"]', '[role="menuitemradio"]',
      '[role="menuitemcheckbox"]',
      '[class*="option"]:not([class*="option-group"])',
      '[class*="menu-item"]', '[class*="MenuItem"]',
      '[class*="dropdown"] li', '[class*="select"] li',
      '[class*="listbox"] li', '[class*="Dropdown"] li',
      'ul[role="listbox"] > li', '[data-testid*="option"]'
    ];
    var seen = new Set();
    var allOpts = [];
    for (var s = 0; s < sels.length; s++) {
      var opts = document.querySelectorAll(sels[s]);
      for (var o = 0; o < opts.length; o++) {
        if (seen.has(opts[o])) continue;
        seen.add(opts[o]);
        var r = opts[o].getBoundingClientRect();
        if (r.width > 0 && r.height > 0) allOpts.push(opts[o]);
      }
    }
    return allOpts;
  }

  function findDropdownMatch(allOpts, lv) {
    for (var i = 0; i < allOpts.length; i++) {
      if (allOpts[i].textContent.trim().toLowerCase() === lv) return allOpts[i];
    }
    for (var i2 = 0; i2 < allOpts.length; i2++) {
      if (allOpts[i2].textContent.trim().toLowerCase().startsWith(lv)) return allOpts[i2];
    }
    for (var i3 = 0; i3 < allOpts.length; i3++) {
      if (allOpts[i3].textContent.trim().toLowerCase().indexOf(lv) !== -1) return allOpts[i3];
    }
    return null;
  }

  // ============================================================
  // FILL DROPDOWN
  // ============================================================
  async function fillDropdown(buttonEl, value) {
    try {
      var clickTarget = buttonEl;
      if (buttonEl.tagName !== 'BUTTON' && buttonEl.tagName !== 'SELECT') {
        var inner = buttonEl.querySelector(
          'button, [role="combobox"], [role="button"], [class*="trigger"], [class*="select"], [class*="dropdown"]'
        );
        if (inner) clickTarget = inner;
      }

      clickTarget.scrollIntoView({ block: 'center' });
      await delay(200);

      var lv = value.toLowerCase().trim();

      var allOpts = await openDropdownMenu(clickTarget);
      log('  [SCAN] <span style="color:#888;">' + allOpts.length + ' options found</span>');

      var match = findDropdownMatch(allOpts, lv);
      if (match) { match.click(); await delay(300); return true; }

      document.body.click();
      await delay(600);

      allOpts = await openDropdownMenu(clickTarget);
      log('  [RETRY] <span style="color:#888;">' + allOpts.length + ' options found</span>');

      match = findDropdownMatch(allOpts, lv);
      if (match) { match.click(); await delay(300); return true; }

      if (allOpts.length > 0) {
        var available = allOpts.slice(0, 8).map(function(o) {
          return '"' + o.textContent.trim().substring(0, 40) + '"';
        }).join(', ');
        log('  [WARN] <span style="color:#ffd740;">Available: ' + available + '</span>');
        log('  [WARN] <span style="color:#ffd740;">Looking for: "' + value + '"</span>');
      } else {
        log('  [FAIL] <span style="color:#ff6b6b;">No dropdown options detected</span>');
      }

      document.body.click();
      await delay(200);
      return false;
    } catch (e) {
      console.error('[SS Filler] fillDropdown error:', e);
      document.body.click();
      return false;
    }
  }

  // ============================================================
  // MULTISELECT HELPERS
  // ============================================================
  async function openMultiSelectMenu(clickTarget) {
    clickTarget.click();
    await delay(CFG.DROPDOWN_WAIT);
    var opts = collectMultiSelectOptions();
    if (opts.length > 0) return opts;

    clickTarget.focus();
    clickTarget.dispatchEvent(new KeyboardEvent('keydown', {
      key: 'ArrowDown', code: 'ArrowDown', bubbles: true
    }));
    await delay(CFG.DROPDOWN_WAIT);
    opts = collectMultiSelectOptions();
    if (opts.length > 0) return opts;

    clickTarget.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
    await delay(150);
    clickTarget.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
    await delay(CFG.DROPDOWN_WAIT);
    return collectMultiSelectOptions();
  }

  function collectMultiSelectOptions() {
    var sels = [
      '[role="option"]', '[role="menuitem"]', '[role="menuitemcheckbox"]',
      '[role="menuitemradio"]', '[role="checkbox"]',
      '[class*="option"]:not([class*="option-group"]):not([class*="options-container"])',
      '[class*="menu-item"]', '[class*="MenuItem"]',
      '[class*="multi-select"] li', '[class*="multiselect"] li',
      '[class*="MultiSelect"] li', '[class*="checkbox-item"]',
      '[class*="CheckboxItem"]', '[data-testid*="option"]',
      'ul[role="listbox"] > li', '[class*="dropdown"] li',
      '[class*="Dropdown"] li'
    ];
    var seen = new Set();
    var results = [];
    for (var s = 0; s < sels.length; s++) {
      var nodes = document.querySelectorAll(sels[s]);
      for (var n = 0; n < nodes.length; n++) {
        if (seen.has(nodes[n])) continue;
        seen.add(nodes[n]);
        var r = nodes[n].getBoundingClientRect();
        if (r.width > 0 && r.height > 0) {
          results.push({ el: nodes[n], text: nodes[n].textContent.trim() });
        }
      }
    }
    return results;
  }

  function findBestMatch(allOpts, target) {
    var lt = target.toLowerCase().trim();

    for (var i = 0; i < allOpts.length; i++) {
      if (allOpts[i].text.toLowerCase().trim() === lt) return allOpts[i];
    }

    if (lt.length > 3) {
      for (var j = 0; j < allOpts.length; j++) {
        if (allOpts[j].text.toLowerCase().trim().startsWith(lt)) return allOpts[j];
      }
    }

    if (lt.length > 3) {
      var containsMatches = [];
      for (var k = 0; k < allOpts.length; k++) {
        var optText = allOpts[k].text.toLowerCase().trim();
        if (optText.indexOf(lt) !== -1) {
          containsMatches.push({ opt: allOpts[k], score: Math.abs(optText.length - lt.length) });
        }
      }
      if (containsMatches.length > 0) {
        containsMatches.sort(function(a, b) { return a.score - b.score; });
        return containsMatches[0].opt;
      }
    }

    for (var m = 0; m < allOpts.length; m++) {
      var optText2 = allOpts[m].text.toLowerCase().trim();
      if (lt.indexOf(optText2) !== -1 && optText2.length > 3) return allOpts[m];
    }

    return null;
  }

  function isAlreadySelected(el) {
    if (el.getAttribute('aria-selected') === 'true') return true;
    if (el.getAttribute('aria-checked') === 'true') return true;
    var checkbox = el.querySelector('input[type="checkbox"]');
    if (checkbox && checkbox.checked) return true;
    var cls = (el.className || '').toLowerCase();
    if (cls.indexOf('selected') !== -1 || cls.indexOf('checked') !== -1 || cls.indexOf('active') !== -1) return true;
    if (el.getAttribute('data-selected') === 'true' || el.getAttribute('data-checked') === 'true') return true;
    return false;
  }

  function logAvailableOptions(allOpts) {
    if (allOpts.length === 0) return;
    var preview = allOpts.slice(0, 10).map(function(o) {
      return '"' + o.text.substring(0, 50) + '"';
    }).join(', ');
    log('  [LIST] <span style="color:#888;">Available: ' + preview +
      (allOpts.length > 10 ? ' ... +' + (allOpts.length - 10) + ' more' : '') + '</span>');
  }

  // ============================================================
  // FILL MULTISELECT
  // ============================================================
  async function fillMultiSelect(buttonEl, value) {
    try {
      var rawValues = smartSplit(value);
      if (!rawValues.length) {
        log('  [WARN] <span style="color:#ffd740;">multiselect: no valid values</span>');
        return false;
      }

      log('  [INPUT] <span style="color:#888;">Raw: [' +
        rawValues.map(function(v) { return '"' + v + '"'; }).join(', ') + ']</span>');

      var categories = resolveToCategories(rawValues);

      log('  [RESOLVED] <span style="color:#4fc3f7;">Targets: [' +
        categories.map(function(v) { return '"' + v + '"'; }).join(', ') + ']</span>');

      if (!categories.length) {
        log('  [FAIL] <span style="color:#ff6b6b;">No categories resolved</span>');
        return false;
      }

      var clickTarget = buttonEl;
      if (buttonEl.tagName !== 'BUTTON' && buttonEl.tagName !== 'SELECT') {
        var inner = buttonEl.querySelector(
          'button, [role="combobox"], [role="button"], [role="listbox"], ' +
          '[class*="trigger"], [class*="select"], [class*="dropdown"]'
        );
        if (inner) clickTarget = inner;
      }

      clickTarget.scrollIntoView({ block: 'center' });
      await delay(200);

      var allOpts = await openMultiSelectMenu(clickTarget);

      if (allOpts.length === 0) {
        log('  [FAIL] <span style="color:#ff6b6b;">No multiselect options found</span>');
        document.body.click();
        return false;
      }

      log('  [SCAN] <span style="color:#888;">' + allOpts.length + ' options available</span>');

      var ok = 0, missed = [];

      for (var vi = 0; vi < categories.length; vi++) {
        var target = categories[vi];
        var match = findBestMatch(allOpts, target);

        if (match) {
          if (!isAlreadySelected(match.el)) {
            match.el.click();
            await delay(400);
            ok++;
            log('  [OK] <span style="color:#69f0ae;">Selected: "' + match.text + '"</span>');
          } else {
            ok++;
            log('  [OK] <span style="color:#69f0ae;">Already selected: "' + match.text + '"</span>');
          }
        } else {
          missed.push(target);
          log('  [MISS] <span style="color:#ff6b6b;">No match for: "' + target + '"</span>');
        }

        if (vi < categories.length - 1) {
          await delay(200);
          allOpts = collectMultiSelectOptions();
          if (allOpts.length === 0) {
            await delay(300);
            allOpts = await openMultiSelectMenu(clickTarget);
          }
        }
      }

      document.dispatchEvent(new KeyboardEvent('keydown', {
        key: 'Escape', code: 'Escape', bubbles: true
      }));
      await delay(200);
      document.body.click();
      await delay(200);

      if (missed.length > 0) {
        log('  [WARN] <span style="color:#ffd740;">Unmatched: ' +
          missed.map(function(m) { return '"' + m + '"'; }).join(', ') + '</span>');
        logAvailableOptions(allOpts);
      }

      return ok > 0;
    } catch (e) {
      console.error('[SS Filler] fillMultiSelect error:', e);
      document.body.click();
      return false;
    }
  }

  // ============================================================
  // STEPS
  // ============================================================
  async function stepNewCase() {
    setProgress(0);
    log('<strong>[STEP] Opening New Case...</strong>', true);
    if (document.querySelector('[data-testid="form-submit"]')) {
      log('[INFO] Form already open.'); return true;
    }
    var btn = null;
    var buttons = document.querySelectorAll('button');
    for (var b = 0; b < buttons.length; b++) {
      if (buttons[b].textContent.trim() === 'New Case') { btn = buttons[b]; break; }
    }
    if (!btn) { log('[FAIL] "New Case" not found.'); return false; }
    btn.click(); log('[OK] Clicked. Waiting...');
    for (var i = 0; i < 10; i++) {
      await delay(500);
      if (document.querySelector('[data-testid="form-submit"]')) { log('[OK] Form loaded.'); return true; }
    }
    log('[WAIT] Still loading...'); return true;
  }

  async function stepSelectType() {
    setProgress(1);
    log('<strong>[STEP] Checking record type...</strong>');
    var btn = document.querySelector('#caseRecordTypeDropdown,[data-testid="case-record-type-dropdown"]');
    if (!btn) { log('[FAIL] Not found.'); return false; }
    if (btn.textContent.trim().indexOf('SSPA ABOS Optimization Case') !== -1) {
      log('[OK] Already selected.'); return true;
    }
    var ok = await fillDropdown(btn, 'SSPA ABOS Optimization Case');
    if (ok) { log('[OK] Selected.'); await delay(CFG.STEP_DELAY); return true; }
    log('[FAIL] Could not select.'); return false;
  }

  async function stepFillForm() {
    setProgress(2);
    if (!Object.keys(parsedData).length) { log('[FAIL] No data! Parse first.', true); return; }
    log('<strong>[STEP] Filling form...</strong>');
    log('--------------------------------');
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
        log('[FAIL] <span style="color:#ff6b6b">' + f.key + '</span>: element not found');
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
        log('[OK] <span style="color:#69f0ae">' + f.key + '</span> -> "' + s + '"');
        filled++;
      } else {
        log('[WARN] <span style="color:#ffd740">' + f.key + '</span>: fill failed');
        failedFields.push(f.key);
        failed++;
      }
    }

    log('--------------------------------');
    log('<strong>[RESULT] OK:' + filled + ' | FAIL:' + failed + ' | SKIP:' + skipped + '</strong>');
    if (failed) log('<em style="color:#ff9900;">Failed fields need manual input</em>');
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
      manualItems.push('Due Date -- <strong style="color:#ff6b6b;">' + dueDateValue + '</strong>');
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
        '<strong>[OK] All fields filled successfully!</strong><br>' +
        '<span>Please review and click Submit.</span></div>');
      return;
    }

    var items = '';
    for (var j = 0; j < manualItems.length; j++) {
      items += '<li>' + manualItems[j] + '</li>';
    }

    log('<div style="margin-top:6px; padding:10px 12px; background:#ff9900; color:#000; border-radius:8px; font-size:12px; line-height:1.7;">' +
      '<strong style="font-size:13px;">[!] MANUAL ACTION REQUIRED</strong><br>' +
      '<span>Please manually fill/verify the following before submitting:</span>' +
      '<ol style="margin:6px 0 4px 18px; padding:0;">' + items + '</ol>' +
      '<span>After confirming all fields, click <strong>Submit</strong>.</span></div>');
  }

  async function runFullFlow() {
    if (!Object.keys(parsedData).length) { alert('Parse first! (Step 1)'); return; }
    var s0 = await stepNewCase();
    if (!s0) return;
    await delay(CFG.STEP_DELAY);
    var s1 = await stepSelectType();
    if (!s1) return;
    await delay(CFG.STEP_DELAY);
    await stepFillForm();
    log('<strong>[DONE] Review the reminder above.</strong>');
  }

  // ============================================================
  // INIT
  // ============================================================
  var initCheck = setInterval(function() {
    if (document.body) { clearInterval(initCheck); createPanel(); }
  }, 1000);

})();
