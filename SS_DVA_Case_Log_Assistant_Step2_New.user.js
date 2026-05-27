// ==UserScript==
// @name         SS DVA Case Auto-Fill (LCS)
// @namespace    https://amazon.com/lcs/
// @version      1.0
// @description  Auto-fill Salesforce DVA case form with visual feedback
// @author       LCS Team
// @match        https://*.lightning.force.com/*
// @match        https://*.salesforce.com/*
// @grant        GM_registerMenuCommand
// @run-at       document-idle
// ==/UserScript==

(function () {
  'use strict';

  // ================================================================
  // SECTION 1: STYLES
  // ================================================================
  function injectStyles() {
    if (document.getElementById('lcs-autofill-styles')) return;
    const style = document.createElement('style');
    style.id = 'lcs-autofill-styles';
    style.textContent = `
      .lcs-hl { position: relative !important; transition: all 0.3s ease !important; border-radius: 4px !important; }
      .lcs-success { background: rgba(46,204,113,0.12) !important; border-left: 3px solid #2ECC71 !important; }
      .lcs-review  { background: rgba(241,196,15,0.12) !important; border-left: 3px solid #F1C40F !important; animation: lcs-pulse 2s ease-in-out 3; }
      .lcs-failed  { background: rgba(231,76,60,0.12)  !important; border-left: 3px solid #E74C3C !important; }
      .lcs-skipped { background: rgba(149,165,166,0.08)!important; border-left: 3px dashed #95A5A6 !important; }

      .lcs-badge {
        position: absolute !important; top: -8px !important; right: 4px !important;
        font-size: 10px !important; font-weight: 700 !important;
        padding: 2px 8px !important; border-radius: 10px !important;
        z-index: 9999 !important; font-family: Arial, sans-serif !important;
        letter-spacing: 0.5px !important; box-shadow: 0 1px 3px rgba(0,0,0,0.15) !important;
        cursor: help !important; user-select: none !important;
      }
      .lcs-bg-success { background: #2ECC71; color: white; }
      .lcs-bg-review  { background: #F1C40F; color: #5D4E1F; }
      .lcs-bg-failed  { background: #E74C3C; color: white; }
      .lcs-bg-skipped { background: #95A5A6; color: white; }

      .lcs-tip {
        position: absolute !important; bottom: calc(100% + 8px) !important; right: 0 !important;
        background: #232F3E !important; color: #FFF !important;
        padding: 8px 12px !important; border-radius: 6px !important;
        font-size: 11px !important; font-family: 'Courier New', monospace !important;
        white-space: nowrap !important; opacity: 0 !important; pointer-events: none !important;
        transition: opacity 0.2s !important; z-index: 10000 !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.25) !important; max-width: 400px !important;
      }
      .lcs-tip::after {
        content: ''; position: absolute; top: 100%; right: 12px;
        border: 5px solid transparent; border-top-color: #232F3E;
      }
      .lcs-hl:hover .lcs-tip { opacity: 1 !important; }

      @keyframes lcs-pulse {
        0%, 100% { box-shadow: 0 0 0 1px rgba(241,196,15,0.2); }
        50%      { box-shadow: 0 0 0 4px rgba(241,196,15,0.5); }
      }

      #lcs-trigger {
        position: fixed !important; bottom: 20px !important; right: 20px !important;
        background: #FF9900 !important; color: #232F3E !important;
        border: none !important; border-radius: 8px !important;
        padding: 12px 20px !important; font-weight: 700 !important;
        font-size: 13px !important; letter-spacing: 1px !important;
        cursor: pointer !important; z-index: 99998 !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.25) !important;
        font-family: Arial, sans-serif !important;
        transition: transform 0.15s !important;
      }
      #lcs-trigger:hover { transform: translateY(-2px); background: #FFB84D !important; }

      #lcs-modal-bg {
        position: fixed !important; inset: 0 !important;
        background: rgba(0,0,0,0.6) !important; z-index: 99999 !important;
        display: flex !important; align-items: center !important; justify-content: center !important;
      }
      #lcs-modal {
        background: #FFF !important; width: 600px !important; max-width: 90vw !important;
        border-radius: 12px !important; padding: 24px !important;
        font-family: Arial, sans-serif !important;
        box-shadow: 0 8px 32px rgba(0,0,0,0.3) !important;
      }
      #lcs-modal h2 { margin: 0 0 16px 0 !important; color: #232F3E !important; font-size: 18px !important; }
      #lcs-modal p  { margin: 0 0 12px 0 !important; color: #565959 !important; font-size: 13px !important; }
      #lcs-modal textarea {
        width: 100% !important; height: 280px !important;
        font-family: 'Courier New', monospace !important; font-size: 12px !important;
        border: 1px solid #888C8C !important; border-radius: 6px !important;
        padding: 10px !important; box-sizing: border-box !important; resize: vertical !important;
      }
      #lcs-modal .lcs-btnrow {
        display: flex !important; justify-content: flex-end !important;
        gap: 8px !important; margin-top: 16px !important;
      }
      #lcs-modal button {
        padding: 8px 18px !important; border-radius: 6px !important;
        font-weight: 700 !important; font-size: 13px !important; cursor: pointer !important;
        font-family: Arial, sans-serif !important;
      }
      #lcs-modal .lcs-btn-primary { background: #FF9900 !important; color: #232F3E !important; border: none !important; }
      #lcs-modal .lcs-btn-cancel  { background: transparent !important; color: #565959 !important; border: 1px solid #888C8C !important; }

      #lcs-panel {
        position: fixed !important; bottom: 20px !important; right: 20px !important;
        background: #232F3E !important; color: #FFF !important;
        border: 2px solid #FF9900 !important; border-radius: 8px !important;
        padding: 12px 16px !important; z-index: 99999 !important;
        font-family: Arial, sans-serif !important; font-size: 12px !important;
        min-width: 240px !important; box-shadow: 0 4px 20px rgba(0,0,0,0.3) !important;
      }
      #lcs-panel .lcs-pt { color: #FF9900 !important; font-weight: 700 !important; margin-bottom: 8px !important; font-size: 11px !important; letter-spacing: 1px !important; }
      #lcs-panel .lcs-row { display: flex !important; justify-content: space-between !important; margin: 4px 0 !important; }
      #lcs-panel .lcs-cnt { font-weight: 700 !important; font-family: 'Courier New', monospace !important; }
      #lcs-panel button {
        margin-top: 6px !important; padding: 6px 12px !important;
        background: transparent !important; color: #FF9900 !important;
        border: 1px solid #FF9900 !important; border-radius: 4px !important;
        cursor: pointer !important; font-size: 11px !important; font-weight: 700 !important; width: 100% !important;
      }
      #lcs-panel button:hover { background: #FF9900 !important; color: #232F3E !important; }
    `;
    document.head.appendChild(style);
  }

  // ================================================================
  // SECTION 2: UTILITIES
  // ================================================================
  const sleep = ms => new Promise(r => setTimeout(r, ms));

  function setNativeValue(el, value) {
    const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
    setter.call(el, value);
    el.dispatchEvent(new Event('input', { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
  }

  function findFieldWrapper(element) {
    if (!element) return null;
    let el = element;
    for (let i = 0; i < 5; i++) {
      if (!el.parentElement) break;
      const p = el.parentElement;
      const cls = (typeof p.className === 'string') ? p.className : '';
      if (
        cls.includes('cell') || cls.includes('field-wrapper') ||
        cls.includes('form-element') || cls.includes('slds-form-element') ||
        (p.id && p.id.endsWith('-cell'))
      ) return p;
      el = p;
    }
    return element.parentElement || element;
  }

  // ================================================================
  // SECTION 3: HIGHLIGHT SYSTEM
  // ================================================================
  const highlighted = new Map();

  function highlight(fieldEl, status, name, value) {
    if (!fieldEl) return false;
    const wrapper = findFieldWrapper(fieldEl);
    if (!wrapper) return false;

    wrapper.classList.remove('lcs-hl', 'lcs-success', 'lcs-review', 'lcs-failed', 'lcs-skipped');
    wrapper.querySelectorAll('.lcs-badge, .lcs-tip').forEach(el => el.remove());

    const conf = {
      success: { cls: 'lcs-success', label: 'AUTO',    bg: 'lcs-bg-success', tip: 'AUTO-FILLED' },
      review:  { cls: 'lcs-review',  label: 'REVIEW',  bg: 'lcs-bg-review',  tip: 'NEEDS REVIEW' },
      failed:  { cls: 'lcs-failed',  label: 'FAILED',  bg: 'lcs-bg-failed',  tip: 'FAILED - MANUAL ENTRY' },
      skipped: { cls: 'lcs-skipped', label: 'SKIP',    bg: 'lcs-bg-skipped', tip: 'INTENTIONALLY SKIPPED' }
    }[status];
    if (!conf) return false;

    wrapper.classList.add('lcs-hl', conf.cls);

    const badge = document.createElement('span');
    badge.className = `lcs-badge ${conf.bg}`;
    badge.textContent = conf.label;
    wrapper.appendChild(badge);

    const tip = document.createElement('div');
    tip.className = 'lcs-tip';
    const time = new Date().toLocaleTimeString('en-US', { hour12: false });
    const dispVal = value && value.length > 60 ? value.substring(0, 60) + '...' : (value || '(empty)');
    tip.innerHTML = `<div style="font-weight:700;color:#FF9900;margin-bottom:2px">${conf.tip}</div>
                     <div><strong>Field:</strong> ${name}</div>
                     <div><strong>Value:</strong> ${dispVal}</div>
                     <div style="opacity:0.7;font-size:10px;margin-top:2px">${time}</div>`;
    wrapper.appendChild(tip);

    highlighted.set(name, { status, value, time });
    updatePanel();
    return true;
  }

  function clearAllHighlights() {
    document.querySelectorAll('.lcs-hl').forEach(el => {
      el.classList.remove('lcs-hl', 'lcs-success', 'lcs-review', 'lcs-failed', 'lcs-skipped');
    });
    document.querySelectorAll('.lcs-badge, .lcs-tip').forEach(el => el.remove());
    highlighted.clear();
    updatePanel();
    console.log('[LCS] Highlights cleared');
  }

  function showPanel() {
    document.getElementById('lcs-panel')?.remove();
    const panel = document.createElement('div');
    panel.id = 'lcs-panel';
    panel.innerHTML = `
      <div class="lcs-pt">AUTO-FILL STATUS</div>
      <div class="lcs-row"><span>Auto-filled:</span><span class="lcs-cnt" id="lcs-c-success" style="color:#2ECC71">0</span></div>
      <div class="lcs-row"><span>Needs review:</span><span class="lcs-cnt" id="lcs-c-review" style="color:#F1C40F">0</span></div>
      <div class="lcs-row"><span>Failed:</span><span class="lcs-cnt" id="lcs-c-failed" style="color:#E74C3C">0</span></div>
      <div class="lcs-row"><span>Skipped:</span><span class="lcs-cnt" id="lcs-c-skipped" style="color:#95A5A6">0</span></div>
      <button id="lcs-btn-clear">CLEAR HIGHLIGHTS</button>
      <button id="lcs-btn-close">CLOSE PANEL</button>
    `;
    document.body.appendChild(panel);
    document.getElementById('lcs-btn-clear').onclick = clearAllHighlights;
    document.getElementById('lcs-btn-close').onclick = () => panel.remove();
    updatePanel();
  }

  function updatePanel() {
    const c = { success: 0, review: 0, failed: 0, skipped: 0 };
    highlighted.forEach(v => c[v.status]++);
    ['success', 'review', 'failed', 'skipped'].forEach(k => {
      const el = document.getElementById(`lcs-c-${k}`);
      if (el) el.textContent = c[k];
    });
  }

  // ================================================================
  // SECTION 4: PARSE INPUT TEXT
  // ================================================================
  function parseInput(text) {
    const data = {};
    const lines = text.split('\n');
    for (const line of lines) {
      if (line.startsWith('===') || !line.includes(':')) continue;
      const idx = line.indexOf(':');
      const key = line.substring(0, idx).trim();
      const val = line.substring(idx + 1).trim();
      data[key] = val;
    }

    // Normalize key fields
    return {
      assignee: data['Assignee'] || '',
      caseStatus: data['Case Status'] || 'New',
      assignmentStatus: data['Assignment status'] || 'Proposed',
      accountVertical: data['Account Vertical'] || '',
      accountName: data['Account Name'] || '',
      brandName: data['Brand Name'] || '',
      rodeoCfId: data['Rodeo order CFID'] || '',
      advertiserId: data['Advertiser ID'] || '',
      entityId: data['Entity ID'] || '',
      advertiserType: data['Advertiser Type'] || '',
      optimizationMarketplace: data['Optimization marketplace'] || '',
      primaryGoalKPI: data['Primary Goal KPI'] || '',
      primaryGoal: data['Primary Goal'] || '',
      optimizationCategories: (data['Optimization categories'] || '').split(',').map(s => s.trim()).filter(Boolean),
      additionalInfo: data['Additional information'] || '',
      assignmentCreationReason: data['Assignment creation reason'] || '',
      podProcess: data['POD Process'] || '',
      dueDate: (data['Due Date'] || '').split(' ')[0],
      submittingTeam: data['Submitting team'] || 'LCS',
      submitterEmail: data['Submitter email'] || '',
      optimizationDelivery: data['Optimization delivery'] || '',
      requesterBy: data['Requester by'] || '',
      optimizationType: data['Optimization type'] || '',
      sfAccountId: data['SF account ID'] || ''
    };
  }

  // ================================================================
  // SECTION 5: FIELD FILLERS
  // ================================================================
  function shouldSkip(value) {
    if (!value) return true;
    const v = String(value).trim().toLowerCase();
    return v === '' || v === '(blank)' || v === 'n/a' || v === 'null';
  }

  // Lever name aliases (em-dash / hyphen / colon variants)
  const LEVER_ALIAS = {
    'ASIN / Pixel Updates': ['Asin/Pixel', 'ASIN/Pixel Updates', 'ASIN / Pixel'],
    'Targeting — Segment': ['Targeting - Segment', 'Targeting—Segment', 'Targeting Segment'],
    'Twitch — Targeting Update': ['Twitch - Targeting Update', 'Twitch—Targeting Update'],
    'Twitch — Pacing Profile': ['Twitch - Pacing Profile'],
    'Twitch — Book New Order': ['Twitch - Book New Order'],
    'Budget Allocation — Order to Order': ['Budget Allocation: Order to Order', 'Budget Allocation - Order to Order'],
    'Budget Allocation — Within DSP Order': ['Budget Allocation: Within DSP Order']
  };

  function normalizeLeverName(name) {
    return name.toLowerCase().replace(/[—–\-:/\s]+/g, '');
  }

  function findLeverMatch(target, available) {
    if (available.includes(target)) return target;
    for (const alias of (LEVER_ALIAS[target] || [])) {
      if (available.includes(alias)) return alias;
    }
    const norm = normalizeLeverName(target);
    return available.find(opt => normalizeLeverName(opt) === norm);
  }

  function fillTextField(selectors, value, fieldName) {
    if (shouldSkip(value)) return false;
    for (const sel of selectors) {
      const el = document.querySelector(sel);
      if (el) {
        setNativeValue(el, value);
        el.dispatchEvent(new Event('blur', { bubbles: true }));
        console.log(`[OK] ${fieldName} -> "${value}"`);
        highlight(el, 'success', fieldName, value);
        return true;
      }
    }
    console.warn(`[MISS] ${fieldName}`);
    return false;
  }

  async function fillDropdown(triggerSelector, value, fieldName) {
    if (shouldSkip(value)) {
      const el = document.querySelector(triggerSelector);
      if (el) highlight(el, 'skipped', fieldName, '(intentionally blank)');
      return false;
    }
    const trigger = document.querySelector(triggerSelector);
    if (!trigger) {
      console.warn(`[MISS] ${fieldName} trigger not found`);
      return false;
    }
    trigger.click();
    await sleep(300);

    const options = document.querySelectorAll('[role="option"], [role="menuitem"], li[data-value], .slds-listbox__item');
    for (const opt of options) {
      const txt = opt.textContent.trim();
      if (txt === value || txt.includes(value)) {
        opt.click();
        await sleep(150);
        console.log(`[OK] ${fieldName} -> "${value}"`);
        highlight(trigger, 'success', fieldName, value);
        return true;
      }
    }
    console.warn(`[MISS] ${fieldName}: "${value}" not in options`);
    highlight(trigger, 'failed', fieldName, value);
    return false;
  }

  async function fillMultiSelect(triggerSelector, values, fieldName) {
    if (!values || !values.length) return false;
    const trigger = document.querySelector(triggerSelector);
    if (!trigger) {
      console.warn(`[MISS] ${fieldName} trigger not found`);
      return false;
    }
    trigger.click();
    await sleep(400);

    const options = document.querySelectorAll('[role="option"], [role="menuitem"], li[data-value]');
    const available = Array.from(options).map(o => o.textContent.trim());
    const selected = [];
    const missed = [];

    for (const target of values) {
      const matched = findLeverMatch(target, available);
      if (matched) {
        const opt = Array.from(options).find(o => o.textContent.trim() === matched);
        if (opt && !opt.classList.contains('slds-is-selected')) {
          opt.click();
          await sleep(100);
        }
        selected.push(matched);
      } else {
        missed.push(target);
      }
    }

    document.body.click(); // close dropdown
    console.log(`[OK] ${fieldName} -> ${selected.length}/${values.length} matched`);
    if (missed.length) console.warn(`[MISS] Unmatched: ${missed.join(', ')}`);

    highlight(trigger, missed.length ? 'review' : 'success', fieldName, selected.join(', '));
    return missed.length === 0;
  }

  function fillSubmitterEmail(email) {
    if (shouldSkip(email)) return false;
    const selectors = [
      '#submitterEmail-field', '#submitter-email-field',
      '#requesterEmail-field', '#submittedByEmail-field',
      '[data-field="submitterEmail"]',
      'input[aria-label*="Submitter email" i]',
      'input[aria-label*="Submitter Email" i]'
    ];
    for (const sel of selectors) {
      const el = document.querySelector(sel);
      if (el) {
        setNativeValue(el, email);
        el.dispatchEvent(new Event('blur', { bubbles: true }));
        console.log(`[OK] submitterEmail -> "${email}"`);
        highlight(el, 'success', 'Submitter Email', email);
        return true;
      }
    }
    // Fallback scan
    const allInputs = document.querySelectorAll('input[type="text"], input[type="email"], input:not([type])');
    for (const input of allInputs) {
      const meta = `${input.id} ${input.name} ${input.placeholder || ''} ${input.getAttribute('aria-label') || ''}`.toLowerCase();
      if ((meta.includes('email') || meta.includes('submit')) && !input.value) {
        setNativeValue(input, email);
        console.log(`[OK] submitterEmail (fallback) -> "${email}"`);
        highlight(input, 'success', 'Submitter Email', email);
        return true;
      }
    }
    console.warn('[MISS] Submitter email field not found');
    return false;
  }

  async function fillDueDate(targetDate) {
    if (shouldSkip(targetDate)) return false;
    const [year, month, day] = targetDate.split('/').map(Number);
    const dateInput = document.querySelector('#datepicker-start');
    if (!dateInput) {
      console.warn('[MISS] #datepicker-start not found');
      return false;
    }

    dateInput.click();
    dateInput.focus();
    await sleep(500);

    const picker = document.querySelector('[role="dialog"], [class*="datepicker-popup"], [class*="calendar-popup"]');
    if (!picker) {
      // Force-fill fallback
      dateInput.removeAttribute('readonly');
      setNativeValue(dateInput, targetDate);
      dateInput.dispatchEvent(new Event('blur', { bubbles: true }));
      dateInput.setAttribute('readonly', '');
      console.log(`[OK] dueDate (force-fill) -> ${targetDate}`);
      highlight(dateInput, 'review', 'Due Date', targetDate);
      return true;
    }

    const today = new Date();
    const monthsDiff = (year - today.getFullYear()) * 12 + (month - 1 - today.getMonth());
    const navBtn = monthsDiff > 0
      ? picker.querySelector('[aria-label*="Next" i], [class*="next"]')
      : picker.querySelector('[aria-label*="Previous" i], [class*="prev"]');

    for (let i = 0; i < Math.abs(monthsDiff); i++) {
      if (navBtn) { navBtn.click(); await sleep(150); }
    }

    await sleep(200);
    const cells = picker.querySelectorAll('[role="gridcell"], td, button[class*="day"]');
    for (const cell of cells) {
      const text = cell.textContent.trim();
      const disabled = cell.getAttribute('aria-disabled') === 'true'
                    || cell.classList.contains('disabled')
                    || cell.classList.contains('outside-month');
      if (text === String(day) && !disabled) {
        cell.click();
        console.log(`[OK] dueDate -> ${targetDate}`);
        highlight(dateInput, 'review', 'Due Date', targetDate);
        return true;
      }
    }
    console.warn(`[MISS] dueDate cell ${day} not found`);
    highlight(dateInput, 'failed', 'Due Date', targetDate);
    return false;
  }

  // ================================================================
  // SECTION 6: MAIN FILL FLOW
  // ================================================================
  async function autoFillForm(data) {
    console.log('[STEP] Expanding sections...');
    document.querySelectorAll('[aria-expanded="false"]').forEach(el => el.click());
    await sleep(600);

    console.log('[STEP] Filling text fields...');
    fillTextField(['#assignee-field'], data.assignee, 'Assignee');
    fillTextField(['#brandName-field'], data.brandName, 'Brand Name');
    fillTextField(['#primaryGoalKPI-field'], data.primaryGoalKPI, 'Primary Goal KPI');
    fillTextField(['#primaryGoal-field'], data.primaryGoal, 'Primary Goal');
    fillTextField(['#assignmentCreationReason-field'], data.assignmentCreationReason, 'Assignment Creation Reason');
    fillTextField(['#additionalInformation-field', '#additional-information-field'], data.additionalInfo, 'Additional Information');
    fillTextField(['#rodeoCfId-field', '#rodeo-cfid-field'], data.rodeoCfId, 'Rodeo CF ID');
    fillTextField(['#advertiserId-field', '#advertiser-id-field'], data.advertiserId, 'Advertiser ID');
    fillTextField(['#entityId-field', '#entity-id-field'], data.entityId, 'Entity ID');
    fillTextField(['#sfAccountId-field', '#sf-account-id-field'], data.sfAccountId, 'SF Account ID');
    fillTextField(['#accountVertical-field', '#account-vertical-field'], data.accountVertical, 'Account Vertical');
    fillTextField(['#accountName-field', '#account-name-field'], data.accountName, 'Account Name');

    console.log('[STEP] Filling submitter email...');
    fillSubmitterEmail(data.submitterEmail);

    console.log('[STEP] Filling dropdowns...');
    await fillDropdown('#status-field', data.caseStatus, 'Case Status');
    await fillDropdown('#assignmentStatus-field', data.assignmentStatus, 'Assignment Status');
    await fillDropdown('#advertiser\\.type-field', data.advertiserType, 'Advertiser Type');
    await fillDropdown('#advertiser\\.marketplaceId-field', data.optimizationMarketplace, 'Optimization Marketplace');
    await fillDropdown('#optimizationDelivery-field', data.optimizationDelivery, 'Optimization Delivery');

    console.log('[STEP] Handling optimization type...');
    if (shouldSkip(data.optimizationType)) {
      const el = document.querySelector('#optimizationType-field');
      if (el) {
        highlight(el, 'skipped', 'Optimization Type', '(intentionally blank)');
        console.log('[SKIP] optimizationType');
      }
    } else {
      await fillDropdown('#optimizationType-field', data.optimizationType, 'Optimization Type');
    }

    console.log('[STEP] Filling multi-select levers...');
    await fillMultiSelect('#optimizationCategories-field', data.optimizationCategories, 'Optimization Categories');

    console.log('[STEP] Filling due date...');
    await fillDueDate(data.dueDate);

    console.log('================================');
    console.log('[DONE] Auto-fill complete');
    console.log('Verify highlighted fields before submitting');
    console.log('================================');
  }

  // ================================================================
  // SECTION 7: UI - TRIGGER & MODAL
  // ================================================================
  function showInputModal() {
    document.getElementById('lcs-modal-bg')?.remove();

    const bg = document.createElement('div');
    bg.id = 'lcs-modal-bg';
    bg.innerHTML = `
      <div id="lcs-modal">
        <h2>SS DVA Case Auto-Fill</h2>
        <p>Paste the generated case log content from the Case Log Tool below:</p>
        <textarea id="lcs-input" placeholder="=== Case Information ===&#10;Assignee: zhxz&#10;..."></textarea>
        <div class="lcs-btnrow">
          <button class="lcs-btn-cancel" id="lcs-cancel">CANCEL</button>
          <button class="lcs-btn-primary" id="lcs-submit">START AUTO-FILL</button>
        </div>
      </div>
    `;
    document.body.appendChild(bg);

    document.getElementById('lcs-cancel').onclick = () => bg.remove();
    bg.onclick = (e) => { if (e.target === bg) bg.remove(); };

    document.getElementById('lcs-submit').onclick = async () => {
      const text = document.getElementById('lcs-input').value.trim();
      if (!text) { alert('Please paste the case log content'); return; }
      try {
        const data = parseInput(text);
        bg.remove();
        showPanel();
        await autoFillForm(data);
      } catch (e) {
        console.error('[LCS] Error:', e);
        alert('Failed: ' + e.message);
      }
    };
  }

  function createTriggerButton() {
    if (document.getElementById('lcs-trigger')) return;
    const btn = document.createElement('button');
    btn.id = 'lcs-trigger';
    btn.textContent = '[ AUTO-FILL ]';
    btn.onclick = showInputModal;
    document.body.appendChild(btn);
  }

  // ================================================================
  // SECTION 8: INIT
  // ================================================================
  function init() {
    injectStyles();
    createTriggerButton();
    if (typeof GM_registerMenuCommand !== 'undefined') {
      GM_registerMenuCommand('Open Auto-Fill', showInputModal);
      GM_registerMenuCommand('Clear Highlights', clearAllHighlights);
      GM_registerMenuCommand('Show Status Panel', showPanel);
    }
    console.log('[LCS] Auto-Fill loaded - click [ AUTO-FILL ] button');
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }

  // Expose for debugging
  window.LCS_AutoFill = { fill: autoFillForm, parse: parseInput, clear: clearAllHighlights, panel: showPanel };

})();
