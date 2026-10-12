(() => {
  const $ = id => document.getElementById(id);
  // 網頁檔已更新、但背景伺服器還是更新前那個：新功能的 API 會回 404
  const STALE_SERVER_MSG = '伺服器還是舊版：請關閉「門市報表製作」的終端機視窗，再重新雙擊啟動檔。';

  const configCard      = $('configCard');
  const shoppertrakCard = $('shoppertrakCard');
  const progressCard = $('progressCard');
  const downloadCard = $('downloadCard');
  const errorCard    = $('errorCard');

  const shopSelect    = $('shopId');
  const weekEndInput  = $('weekEnd');
  const weekRangeHint = $('weekRangeHint');
  const yoyEndInput   = $('yoyEnd');
  const yoyRangeHint  = $('yoyRangeHint');
  const empCountInput = $('empCount');
  const empCountHint  = $('empCountHint');
  const advPeriods    = $('advPeriods');
  const ADV_FIELDS    = ['prevWkStart','prevWkEnd','wkStart','wkEnd',
                         'moStart','moEnd','lmStart','lmEnd','lyStart','lyEnd'];
  const dssCard       = $('dssCard');
  const stBadge       = $('stBadge');
  const dssBadge      = $('dssBadge');
  const dssForce      = $('dssForce');
  const dssUser       = $('dssUser');
  const dssPass       = $('dssPass');
  const dssSaveBtn    = $('dssSaveBtn');
  const dssClearBtn   = $('dssClearBtn');
  const dssLoginBtn   = $('dssLoginBtn');
  const dssCaptchaGroup = $('dssCaptchaGroup');
  const dssCaptchaImg   = $('dssCaptchaImg');
  const dssCaptchaRefreshBtn = $('dssCaptchaRefreshBtn');
  const dssCaptchaInput = $('dssCaptchaInput');
  const dssCaptchaSubmitBtn = $('dssCaptchaSubmitBtn');
  const dssOtpGroup   = $('dssOtpGroup');
  const dssOtpInput   = $('dssOtpInput');
  const dssOtpSubmitBtn = $('dssOtpSubmitBtn');
  const dssStatus     = $('dssStatus');
  const stUser        = $('stUser');
  const stPass        = $('stPass');
  const stSaveBtn     = $('stSaveBtn');
  const stClearBtn    = $('stClearBtn');
  const stStatus      = $('stStatus');
  const generateBtn   = $('generateBtn');
  const progressBar   = $('progressBar');
  const progressLabel = $('progressLabel');
  const logBox        = $('logBox');
  const downloadBtn   = $('downloadBtn');
  const downloadInfo  = $('downloadInfo');
  const showLogBtn    = $('showLogBtn');
  const downloadLogBox = $('downloadLogBox');
  const resetBtn      = $('resetBtn');
  const retryBtn      = $('retryBtn');
  const errorMsg      = $('errorMsg');

  let pollTimer    = null;
  let currentJobId = null;
  let seenCount    = 0;
  let latestLogMessages = [];

  function fmtDate(d) {
    return `${d.getFullYear()}/${String(d.getMonth()+1).padStart(2,'0')}/${String(d.getDate()).padStart(2,'0')}`;
  }

  function updateHint(val) {
    if (!val) { weekRangeHint.textContent = ''; return; }
    const end   = new Date(val + 'T00:00:00');
    const start = new Date(end); start.setDate(end.getDate() - 6);
    weekRangeHint.textContent = `本週範圍：${fmtDate(start)} ～ ${fmtDate(end)}`;
  }

  function updateYoyHint() {
    // 年對年截止日留空時，視同採用週結束日
    const val = yoyEndInput.value || weekEndInput.value;
    if (!val) { yoyRangeHint.textContent = '留空則沿用週結束日；Speakers 年累積算到本月結束日'; return; }
    const end = new Date(val + 'T00:00:00');
    const y = end.getFullYear();
    const md = `${String(end.getMonth()+1).padStart(2,'0')}/${String(end.getDate()).padStart(2,'0')}`;
    const note = yoyEndInput.value ? '；Speakers 年累積同此截止日'
                                   : '（沿用週結束日；Speakers 年累積算到本月結束日）';
    yoyRangeHint.textContent = `年對年比較：${y-1}/01/01～${y-1}/${md} vs ${y}/01/01～${y}/${md}${note}`;
  }

  async function fetchPeriodDefaults() {
    const wkEnd = weekEndInput.value;
    if (!wkEnd) return;
    try {
      const res = await fetch(`/api/periods?weekEnd=${encodeURIComponent(wkEnd)}`);
      if (!res.ok) return;
      const d = await res.json();
      ADV_FIELDS.forEach(k => { if (d[k]) $(k).value = d[k]; });
    } catch (e) { /* ignore */ }
  }

  advPeriods.addEventListener('toggle', () => {
    // 展開時，若欄位空白則帶入目前週結束日推算的預設值
    if (advPeriods.open && !$('wkEnd').value) fetchPeriodDefaults();
  });

  weekEndInput.addEventListener('change', () => {
    updateHint(weekEndInput.value);
    updateYoyHint();
    if (advPeriods.open) fetchPeriodDefaults();   // 面板開啟時同步刷新預設區間
  });
  yoyEndInput.addEventListener('change', updateYoyHint);

  async function loadStores() {
    try {
      const res = await fetch('/api/stores');
      const data = await res.json();
      const items = data.items || [];
      shopSelect.innerHTML = '';
      for (const it of items) {
        const opt = document.createElement('option');
        opt.value = it.storeId;
        opt.textContent = `${it.storeId}　${it.name}`;
        shopSelect.appendChild(opt);
      }
      if (data.default) shopSelect.value = data.default;
      await loadShopConfig();
      loadMysetup(false);
    } catch (e) {
      console.warn('無法取得門市清單', e);
    }
  }

  async function loadDefaultDate() {
    try {
      const res = await fetch('/api/default-date');
      const { date } = await res.json();
      weekEndInput.value = date;
      updateHint(date);
      updateYoyHint();
    } catch (e) {
      console.warn('無法取得預設日期', e);
    }
  }

  function show(card) {
    [configCard, progressCard, downloadCard, errorCard].forEach(c => {
      c.hidden = (c !== card);
    });
    // 設定卡與來客數／DSS 設定一起顯示／隱藏
    shoppertrakCard.hidden = (card !== configCard);
    dssCard.hidden = (card !== configCard);
    mysetupCard.hidden = (card !== configCard);
    mycoachCard.hidden = (card !== configCard);
  }

  function setBadge(el, text, cls) {
    el.textContent = text;
    el.className = 'status-pill' + (cls ? ' ' + cls : '');
  }

  async function loadShoppertrakStatus() {
    try {
      const res = await fetch('/api/shoppertrak/status');
      const data = await res.json();
      if (!data.available) {
        stStatus.textContent = '（來客數模組未載入，將略過來客數）';
        setBadge(stBadge, '模組未載入', 'err');
        return;
      }
      if (data.username) stUser.value = data.username;
      stStatus.textContent = data.hasCredentials ? '已儲存本機帳密' : '未儲存帳密';
      if (data.hasCredentials) setBadge(stBadge, '✓ 已設定帳密', 'ok');
      else setBadge(stBadge, '未設定帳密', 'warn');
    } catch (e) {
      stStatus.textContent = '';
      setBadge(stBadge, '—', '');
    }
  }

  async function loadShopConfig() {
    const shopId = shopSelect.value;
    if (!shopId) return;
    try {
      const res = await fetch(`/api/config?shopId=${encodeURIComponent(shopId)}`);
      const data = await res.json();
      empCountInput.value = (data.employeeCount != null) ? data.employeeCount : '';
      empCountHint.textContent = data.hasSiteId
        ? '用於人均產值（總營業額 / 編制人數）。會記住上次輸入。'
        : '用於人均產值。注意：此門市無對應 ShopperTrak，將略過來客數。';
    } catch (e) { /* ignore */ }
  }

  shopSelect.addEventListener('change', loadShopConfig);

  stSaveBtn.addEventListener('click', async () => {
    const username = stUser.value.trim();
    const password = stPass.value;
    if (!username || !password) { stStatus.textContent = '請輸入帳號與密碼'; return; }
    stSaveBtn.disabled = true;
    try {
      const res = await fetch('/api/shoppertrak/credentials', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username, password }),
      });
      const data = await res.json();
      stStatus.textContent = (res.ok && data.ok) ? '已儲存本機帳密' : (data.error || '儲存失敗');
      if (res.ok && data.ok) stPass.value = '';
    } catch (e) {
      stStatus.textContent = '儲存失敗：' + e.message;
    } finally {
      stSaveBtn.disabled = false;
    }
  });

  stClearBtn.addEventListener('click', async () => {
    try {
      const res = await fetch('/api/shoppertrak/credentials', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ clear: true }),
      });
      if (res.ok) { stUser.value = ''; stPass.value = ''; stStatus.textContent = '已清除帳密'; }
    } catch (e) { /* ignore */ }
  });

  // ─── DSS（搭售統計）登入流程 ───────────────────────────────────────
  const DSS_STATE_TEXT = {
    idle: '未登入',
    need_captcha: '請輸入圖形驗證碼',
    need_otp: '已要求 DSS 寄送 Email 驗證碼，請收信後輸入',
    logged_in: '✓ 已登入 DSS',
    error: '',
  };

  let dssState = { state: 'idle', force: false };

  function renderDssBadge(data) {
    const forced = data.force ? '強制DSS · ' : '';
    if (!data.available) { setBadge(dssBadge, '模組未載入', 'err'); return; }
    if (data.state === 'logged_in') {
      setBadge(dssBadge, forced + '✓ 已登入', 'ok');
    } else if (data.state === 'need_captcha' || data.state === 'need_otp') {
      setBadge(dssBadge, forced + '登入中…', 'warn');
    } else if (data.force) {
      // 強制 DSS 但未登入 → 醒目錯誤
      setBadge(dssBadge, '強制DSS · ✗ 未登入', 'err');
    } else {
      setBadge(dssBadge, data.hasCredentials ? '未登入（用 EPB）' : '未設定（用 EPB）', '');
    }
  }

  function renderDssState(data) {
    dssState = data;
    if (typeof data.force === 'boolean') dssForce.checked = data.force;
    renderDssBadge(data);
    if (!data.available) {
      dssStatus.textContent = '（DSS 模組未載入，將略過搭售統計）';
      dssCaptchaGroup.hidden = true;
      dssOtpGroup.hidden = true;
      return;
    }
    if (data.username && !dssUser.value) dssUser.value = data.username;
    dssCaptchaGroup.hidden = (data.state !== 'need_captcha');
    dssOtpGroup.hidden = (data.state !== 'need_otp');
    if (data.state === 'need_captcha') {
      dssCaptchaImg.src = `/api/dss/captcha?ts=${Date.now()}`;
      dssCaptchaInput.value = '';
      dssCaptchaInput.focus();
    }
    if (data.state === 'need_otp') {
      dssOtpInput.value = '';
      dssOtpInput.focus();
    }
    let text = DSS_STATE_TEXT[data.state] ?? '';
    if (data.state === 'need_otp' && data.otpSent === false) {
      text = '需要 Email 驗證碼，但未確認寄送成功；請重新登入 DSS 或通知管理員檢查網站流程。';
    }
    if (data.error) text = (text ? text + '　' : '') + data.error;
    if (data.state === 'idle' && !data.hasCredentials) text = '未儲存帳密';
    if (data.force && data.state !== 'logged_in') {
      text = '⚠ 已勾選「強制使用 DSS」但目前未登入，產生報表時搭售統計（Sheet 10/11）會失敗。請先登入或取消勾選。';
    }
    dssStatus.textContent = text;
  }

  async function dssPost(path, body) {
    const res = await fetch(path, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body || {}),
    });
    return res.json();
  }

  async function loadDssStatus() {
    try {
      renderDssState(await (await fetch('/api/dss/status')).json());
    } catch (e) {
      dssStatus.textContent = '';
    }
  }

  dssSaveBtn.addEventListener('click', async () => {
    const username = dssUser.value.trim();
    const password = dssPass.value;
    if (!username || !password) { dssStatus.textContent = '請輸入帳號與密碼'; return; }
    dssSaveBtn.disabled = true;
    try {
      const data = await dssPost('/api/dss/credentials', { username, password });
      dssStatus.textContent = data.ok ? '已儲存本機帳密，可按「登入 DSS」' : (data.error || '儲存失敗');
      if (data.ok) dssPass.value = '';
    } catch (e) {
      dssStatus.textContent = '儲存失敗：' + e.message;
    } finally {
      dssSaveBtn.disabled = false;
    }
  });

  dssClearBtn.addEventListener('click', async () => {
    try {
      const data = await dssPost('/api/dss/credentials', { clear: true });
      if (data.ok) {
        dssUser.value = ''; dssPass.value = '';
        dssCaptchaGroup.hidden = true; dssOtpGroup.hidden = true;
        dssStatus.textContent = '已清除帳密';
      }
    } catch (e) { /* ignore */ }
  });

  dssLoginBtn.addEventListener('click', async () => {
    dssLoginBtn.disabled = true;
    dssStatus.textContent = '連線 DSS 取得驗證碼…';
    try {
      renderDssState(await dssPost('/api/dss/login/start'));
    } catch (e) {
      dssStatus.textContent = '連線失敗：' + e.message;
    } finally {
      dssLoginBtn.disabled = false;
    }
  });

  dssCaptchaRefreshBtn.addEventListener('click', async () => {
    try { renderDssState(await dssPost('/api/dss/login/refresh-captcha')); } catch (e) { /* ignore */ }
  });

  dssCaptchaSubmitBtn.addEventListener('click', async () => {
    const code = dssCaptchaInput.value.trim();
    if (!code) { dssStatus.textContent = '請輸入驗證碼'; return; }
    dssCaptchaSubmitBtn.disabled = true;
    dssStatus.textContent = '登入中…';
    try {
      renderDssState(await dssPost('/api/dss/login/captcha', { code }));
    } catch (e) {
      dssStatus.textContent = '送出失敗：' + e.message;
    } finally {
      dssCaptchaSubmitBtn.disabled = false;
    }
  });

  dssCaptchaInput.addEventListener('keydown', e => {
    if (e.key === 'Enter') dssCaptchaSubmitBtn.click();
  });

  dssOtpSubmitBtn.addEventListener('click', async () => {
    const code = dssOtpInput.value.trim();
    if (!code) { dssStatus.textContent = '請輸入 Email 驗證碼'; return; }
    dssOtpSubmitBtn.disabled = true;
    dssStatus.textContent = '驗證中…';
    try {
      renderDssState(await dssPost('/api/dss/login/otp', { code }));
    } catch (e) {
      dssStatus.textContent = '送出失敗：' + e.message;
    } finally {
      dssOtpSubmitBtn.disabled = false;
    }
  });

  dssOtpInput.addEventListener('keydown', e => {
    if (e.key === 'Enter') dssOtpSubmitBtn.click();
  });

  dssForce.addEventListener('change', async () => {
    const force = dssForce.checked;
    try {
      const data = await dssPost('/api/dss/force', { force });
      if (!data.ok) {
        dssForce.checked = !force;
        dssStatus.textContent = data.error || '設定失敗';
        return;
      }
      renderDssState(Object.assign({}, dssState, { force }));
      if (force && dssState.state !== 'logged_in') {
        dssCard.querySelector('details').open = true;   // 提醒使用者登入
      }
    } catch (e) {
      dssForce.checked = !force;
      dssStatus.textContent = '設定失敗：' + e.message;
    }
  });

  function appendLog(text) {
    const cls = text.includes('✓') ? 'ok'
              : text.includes('✗') ? 'err'
              : 'info';
    const line = document.createElement('span');
    line.className = cls;
    line.textContent = text + '\n';
    logBox.appendChild(line);
    logBox.scrollTop = logBox.scrollHeight;
  }

  function renderDownloadLog() {
    downloadLogBox.innerHTML = '';
    latestLogMessages.forEach(msg => {
      const cls = msg.includes('✓') ? 'ok'
                : msg.includes('✗') ? 'err'
                : 'info';
      const line = document.createElement('span');
      line.className = cls;
      line.textContent = msg + '\n';
      downloadLogBox.appendChild(line);
    });
    downloadLogBox.scrollTop = downloadLogBox.scrollHeight;
  }

  showLogBtn.addEventListener('click', () => {
    const nextHidden = !downloadLogBox.hidden;
    downloadLogBox.hidden = nextHidden;
    showLogBtn.textContent = nextHidden ? '顯示 log' : '隱藏 log';
    if (!nextHidden) renderDownloadLog();
  });

  function stopPoll() {
    if (pollTimer) { clearInterval(pollTimer); pollTimer = null; }
  }

  function estimateProgress(messages) {
    const all = messages.join(' ');
    if (all.includes('儲存 Excel'))   return 95;
    if (all.includes('個人報表'))     return 78;
    if (all.includes('年對年報表'))   return 68;
    if (all.includes('填入 1-2、5-9')) return 60;
    if (all.includes('員工清單'))     return 45;
    if (all.includes('取得 本期'))    return 30;
    if (all.includes('查詢 EPB'))     return 12;
    return 4;
  }


  // ─── MySetup 名稱對照（Personal Setup Excel 的 Staff Name → 員工代碼）──────
  const mysetupCard = $('mysetupCard');
  const mysetupFile = $('mysetupFile');
  const msDetails   = $('msDetails');
  const msBadge     = $('msBadge');
  const msUnmatchedTitle = $('msUnmatchedTitle');
  const msUnmatchedGroup = $('msUnmatchedGroup');
  const msRulesGroup = $('msRulesGroup');
  const msNewName   = $('msNewName');
  const msNewCode   = $('msNewCode');
  const msAddBtn    = $('msAddBtn');
  const msSaveBtn   = $('msSaveBtn');
  const msReloadBtn = $('msReloadBtn');
  const msStatus    = $('msStatus');
  const msNotice    = $('msNotice');
  let ms = { rules: [], unmatched: [], employees: null, skip: '__skip__' };

  function msEsc(t) {
    return String(t ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }

  function msOptions(selected) {
    const opts = [`<option value="">請選擇員工…</option>`];
    const emps = (ms.employees || []).slice();
    if (selected && selected !== ms.skip && !emps.some(e => e.code === selected)) {
      emps.push({ code: selected, name: '', nameLang: '' });   // 規則指到的員工不在近期名單（如已離職）也要能顯示
    }
    emps.forEach(e => {
      const label = [e.code, e.name, e.nameLang].filter(Boolean).join(' ');
      opts.push(`<option value="${msEsc(e.code)}"${e.code === selected ? ' selected' : ''}>${msEsc(label)}</option>`);
    });
    opts.push(`<option value="${msEsc(ms.skip)}"${selected === ms.skip ? ' selected' : ''}>不計入（他店支援／已離職）</option>`);
    return opts.join('');
  }

  function msRender() {
    const pending = ms.unmatched.length;
    if (pending) setBadge(msBadge, `${pending} 個名字待指定`, 'warn');
    else setBadge(msBadge, ms.rules.length ? `✓ ${ms.rules.length} 條規則` : '自動對應', 'ok');

    msUnmatchedTitle.hidden = msUnmatchedGroup.hidden = !pending;
    msUnmatchedGroup.innerHTML = ms.unmatched.map(u => `
      <div class="row" data-name="${msEsc(u.name)}">
        <span class="row-label">${msEsc(u.name)}<span class="row-sub">${u.count} 筆・${msEsc(u.lastSeen || '')}</span></span>
        <div class="row-control"><select>${msOptions('')}</select></div>
      </div>`).join('');

    msRulesGroup.innerHTML = ms.rules.length
      ? ms.rules.map(r => `
      <div class="row" data-name="${msEsc(r.name)}">
        <span class="row-label">${msEsc(r.name)}</span>
        <div class="row-control"><select>${msOptions(r.code)}</select></div>
        <button class="ghost-btn ms-del" type="button" title="刪除這條規則">刪除</button>
      </div>`).join('')
      : `<div class="row"><span class="row-sub flush">尚無規則（目前全部靠自動對應）</span></div>`;
    msNewCode.innerHTML = msOptions('');
  }

  async function loadMysetup(withEmployees) {
    const shopId = shopSelect.value;
    if (!shopId) return;
    if (withEmployees) msStatus.textContent = '載入本店員工名單中（需連 EPB）…';
    try {
      const res = await fetch(`/api/mysetup/aliases?shopId=${encodeURIComponent(shopId)}&employees=${withEmployees ? 1 : 0}`);
      if (res.status === 404) throw new Error(STALE_SERVER_MSG);
      const data = await res.json();
      ms.rules = data.rules || [];
      ms.unmatched = data.unmatched || [];
      ms.skip = data.skip || ms.skip;
      if (withEmployees) {
        ms.employees = data.employees || [];
        msStatus.textContent = data.employeesError
          ? '⚠ 無法載入員工名單（請確認 VPN）：下拉只顯示已設定的員工'
          : '';
      }
      msRender();
    } catch (e) {
      setBadge(msBadge, '—', '');
      msStatus.textContent = '載入失敗：' + e.message;
    }
  }

  // 收合時只抓規則（不連 EPB），展開才載入員工下拉
  msDetails.addEventListener('toggle', () => {
    if (msDetails.open && ms.employees === null) loadMysetup(true);
  });
  shopSelect.addEventListener('change', () => {
    ms.employees = null;
    loadMysetup(msDetails.open);
  });
  msReloadBtn.addEventListener('click', () => loadMysetup(true));

  msRulesGroup.addEventListener('click', e => {
    const btn = e.target.closest('.ms-del');
    if (!btn) return;
    const name = btn.closest('.row').dataset.name;
    ms.rules = ms.rules.filter(r => r.name !== name);
    msRender();
    msStatus.textContent = '已移除，按「儲存對照」後生效';
  });

  msAddBtn.addEventListener('click', () => {
    const name = msNewName.value.trim();
    const code = msNewCode.value;
    if (!name || !code) { msStatus.textContent = '請輸入名字並選擇員工'; return; }
    ms.rules = ms.rules.filter(r => r.name !== name).concat([{ name, code }]);
    msNewName.value = '';
    msRender();
    msStatus.textContent = '已加入，按「儲存對照」後生效';
  });

  msSaveBtn.addEventListener('click', async () => {
    // 已設定規則（含修改過的下拉）＋未對上清單中有選員工的
    const rules = [...msRulesGroup.querySelectorAll('.row[data-name]')].map(row => ({
      name: row.dataset.name, code: row.querySelector('select').value,
    })).filter(r => r.code);
    msUnmatchedGroup.querySelectorAll('.row[data-name]').forEach(row => {
      const code = row.querySelector('select').value;
      if (code) rules.push({ name: row.dataset.name, code });
    });
    msSaveBtn.disabled = true;
    try {
      const res = await fetch('/api/mysetup/aliases', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ shopId: shopSelect.value, rules }),
      });
      if (res.status === 404) throw new Error(STALE_SERVER_MSG);
      const data = await res.json();
      if (!res.ok || !data.ok) throw new Error(data.error || '儲存失敗');
      ms.rules = data.rules || [];
      ms.unmatched = data.unmatched || [];
      msRender();
      msStatus.textContent = `已儲存 ${ms.rules.length} 條規則，下次產報自動套用`;
    } catch (e) {
      msStatus.textContent = '儲存失敗：' + e.message;
    } finally {
      msSaveBtn.disabled = false;
    }
  });

  function readFileAsDataURL(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject(reader.error);
      reader.readAsDataURL(file);
    });
  }


  // ─── MyCoach 登入（充值數）：每次產報前輸入驗證碼登入 ─────────────────
  const mycoachCard = $('mycoachCard');
  const mcBadge = $('mcBadge'), mcUser = $('mcUser'), mcPass = $('mcPass'), mcStatus = $('mcStatus');
  const mcCaptchaGroup = $('mcCaptchaGroup'), mcCaptchaImg = $('mcCaptchaImg'), mcCode = $('mcCode');
  const mcLoginActions = $('mcLoginActions'), mcStartBtn = $('mcStartBtn'), mcLoginBtn = $('mcLoginBtn');
  let mcState = null;

  function renderMycoach(data) {
    mcState = data;
    if (data.username && !mcUser.value) mcUser.value = data.username;
    if (data.hasCredentials) mcPass.placeholder = '已儲存（要更改再輸入）';
    const need = data.state === 'need_captcha';
    mcCaptchaGroup.hidden = mcLoginActions.hidden = !need;
    if (need) { mcCaptchaImg.src = '/api/mycoach/captcha?t=' + Date.now(); }
    if (data.state === 'logged_in') {
      const t = data.loggedInAt ? data.loggedInAt.slice(11, 16) : '';
      setBadge(mcBadge, '✓ 已登入 ' + t, 'ok');
      mcStatus.textContent = '已登入，這次產報會帶入充值數。';
      mcStartBtn.textContent = '重新登入';
    } else {
      setBadge(mcBadge, data.hasCredentials ? '未登入' : '未設定帳密', 'warn');
      mcStatus.textContent = data.error || (need ? '請輸入圖中的驗證碼後按「登入」。' : (data.hasCredentials ? '按「取得驗證碼」開始登入。' : '請先輸入帳號密碼並儲存。'));
      mcStartBtn.textContent = '取得驗證碼';
    }
  }

  async function mcCall(path, body) {
    const res = await fetch(path, body === undefined ? undefined : {
      method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
    if (res.status === 404) throw new Error(STALE_SERVER_MSG);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || '操作失敗');
    renderMycoach(data);
    return data;
  }

  async function loadMycoach(probe) {
    try { await mcCall('/api/mycoach/status' + (probe ? '?probe=1' : '')); }
    catch (e) { setBadge(mcBadge, '—', ''); mcStatus.textContent = e.message; }
  }

  $('mcSaveBtn').addEventListener('click', async () => {
    const username = mcUser.value.trim(), password = mcPass.value;
    if (!username || !password) { mcStatus.textContent = '請輸入帳號與密碼'; return; }
    try { await mcCall('/api/mycoach/credentials', { username, password }); mcPass.value = ''; mcStatus.textContent = '已儲存帳密，按「取得驗證碼」登入。'; }
    catch (e) { mcStatus.textContent = e.message; }
  });
  mcStartBtn.addEventListener('click', async () => {
    mcStartBtn.disabled = true; mcStatus.textContent = '取得驗證碼中…';
    try {
      // 有輸入新密碼就先存，再開始登入
      if (mcUser.value.trim() && mcPass.value) { await mcCall('/api/mycoach/credentials', { username: mcUser.value.trim(), password: mcPass.value }); mcPass.value = ''; }
      await mcCall('/api/mycoach/login/start', {});
      mcCode.value = ''; mcCode.focus();
    } catch (e) { mcStatus.textContent = e.message; } finally { mcStartBtn.disabled = false; }
  });
  $('mcRefreshBtn').addEventListener('click', () => mcCall('/api/mycoach/login/start', {}).catch(e => { mcStatus.textContent = e.message; }));
  async function mcSubmit() {
    const code = mcCode.value.trim();
    if (!code) { mcStatus.textContent = '請輸入驗證碼'; return; }
    mcLoginBtn.disabled = true; mcStatus.textContent = '登入中…';
    try { await mcCall('/api/mycoach/login/code', { code }); mcCode.value = ''; }
    catch (e) { mcStatus.textContent = e.message; } finally { mcLoginBtn.disabled = false; }
  }
  mcLoginBtn.addEventListener('click', mcSubmit);
  mcCode.addEventListener('keydown', e => { if (e.key === 'Enter') mcSubmit(); });

  async function pollStatus() {
    if (!currentJobId) return;
    try {
      const res  = await fetch(`/api/status?jobId=${currentJobId}`);
      const data = await res.json();

      const msgs = data.messages || [];
      latestLogMessages = msgs.slice();
      for (let i = seenCount; i < msgs.length; i++) appendLog(msgs[i]);
      seenCount = msgs.length;

      const pct = estimateProgress(msgs);
      progressBar.style.width = pct + '%';

      if (data.status === 'done') {
        stopPoll();
        progressBar.style.width = '100%';
        progressLabel.textContent = '完成！';
        downloadInfo.textContent = `報表已產生完成（${data.filename || '門市報表.xlsx'}）`;
        downloadBtn.onclick = () => {
          window.location.href = `/api/download?jobId=${currentJobId}`;
        };
        downloadLogBox.hidden = true;
        showLogBtn.textContent = '顯示 log';
        renderDownloadLog();
        const unmatched = data.mysetupUnmatched || 0;
        msNotice.hidden = !unmatched;
        msNotice.textContent = unmatched
          ? `MySetup：${unmatched} 個名字未對上，已列在第 13/14 頁加總列上方。請到「MySetup 名稱對照」指定員工，下次產報自動套用。`
          : '';
        if (unmatched) loadMysetup(false);
        show(downloadCard);
      } else if (data.status === 'error') {
        stopPoll();
        errorMsg.textContent = data.error || '未知錯誤，請確認 VPN 已連線並重試。';
        show(errorCard);
      } else {
        progressLabel.textContent = data.status === 'pending' ? '排隊中…' : '處理中…';
      }
    } catch (e) {
      appendLog('✗ 連線中斷：' + e.message);
    }
  }

  generateBtn.addEventListener('click', async () => {
    const wkEnd = weekEndInput.value;
    const yoyEnd = yoyEndInput.value;   // 留空 → 後端沿用週結束日
    const shopId = shopSelect.value;
    const empCount = empCountInput.value.trim();
    if (!shopId) { alert('請選擇門市'); return; }
    if (!wkEnd)  { alert('請選擇週結束日期'); return; }
    if (yoyEnd && yoyEnd < wkEnd) {
      if (!confirm(`年對年截止日（${yoyEnd}）早於週結束日（${wkEnd}），確定要這樣產生嗎？`)) return;
    }

    // 強制 DSS 模式：未登入直接擋下，避免 Sheet 10/11 必然失敗
    try {
      const dss = await (await fetch('/api/dss/status')).json();
      renderDssState(dss);
      if (dss.force && dss.state !== 'logged_in') {
        dssCard.querySelector('details').open = true;
        alert('已勾選「強制使用 DSS」但 DSS 尚未登入。\n請先在「搭售統計設定」完成登入，或取消強制使用 DSS。');
        return;
      }
    } catch (e) { /* 狀態查詢失敗不擋產生 */ }

    // MyCoach 登入撐不過一天：產報前確認還登入著，沒登入就提醒充值數會留空
    await loadMycoach(true);
    if (!mcState || mcState.state !== 'logged_in') {
      if (!confirm('MyCoach 尚未登入（或登入已過期），第 13/14 頁的充值數會留空。\n確定要直接產生嗎？')) {
        mycoachCard.scrollIntoView({ behavior: 'smooth', block: 'center' });
        return;
      }
    }

    logBox.innerHTML = '';
    downloadLogBox.innerHTML = '';
    downloadLogBox.hidden = true;
    showLogBtn.textContent = '顯示 log';
    latestLogMessages = [];
    seenCount = 0;
    progressBar.style.width = '0%';
    progressLabel.textContent = '準備中…';
    show(progressCard);

    const body = { shopId, weekEnd: wkEnd, yoyEnd: yoyEnd || '', employeeCount: empCount };
    if (mysetupFile.files[0]) {
      try {
        body.mysetupFile = await readFileAsDataURL(mysetupFile.files[0]);
      } catch (e) {
        errorMsg.textContent = 'Personal Setup Excel 讀取失敗：' + e.message;
        show(errorCard);
        return;
      }
    }
    if (advPeriods.open) {
      const periods = {};
      ADV_FIELDS.forEach(k => { const v = $(k).value; if (v) periods[k] = v; });
      if (Object.keys(periods).length) body.periods = periods;
    }

    try {
      const res  = await fetch('/api/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
      });
      const data = await res.json();
      if (!res.ok || data.error) {
        errorMsg.textContent = data.error || '啟動失敗';
        show(errorCard);
        return;
      }
      currentJobId = data.jobId;
      appendLog('工作已啟動，請稍候…');
      pollTimer = setInterval(pollStatus, 2000);
    } catch (e) {
      errorMsg.textContent = '無法連線至本機伺服器：' + e.message;
      show(errorCard);
    }
  });

  resetBtn.addEventListener('click', () => {
    stopPoll();
    currentJobId = null;
    seenCount = 0;
    show(configCard);
  });

  retryBtn.addEventListener('click', () => {
    stopPoll();
    currentJobId = null;
    seenCount = 0;
    show(configCard);
  });

  loadStores();
  loadDefaultDate();
  loadShoppertrakStatus();
  loadDssStatus();
  loadMycoach(false);
})();
