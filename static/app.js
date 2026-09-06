(() => {
  const $ = id => document.getElementById(id);

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
      text = '⚠ 已勾選「強制使用 DSS」但目前未登入，產生報表時搭售統計（Sheet 8/9）會失敗。請先登入或取消勾選。';
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
    if (all.includes('個人 8-11'))    return 78;
    if (all.includes('填入 1-5'))     return 60;
    if (all.includes('員工清單'))     return 45;
    if (all.includes('取得 本期'))    return 30;
    if (all.includes('查詢 EPB'))     return 12;
    return 4;
  }

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

    // 強制 DSS 模式：未登入直接擋下，避免 Sheet 8/9 必然失敗
    try {
      const dss = await (await fetch('/api/dss/status')).json();
      renderDssState(dss);
      if (dss.force && dss.state !== 'logged_in') {
        dssCard.querySelector('details').open = true;
        alert('已勾選「強制使用 DSS」但 DSS 尚未登入。\n請先在「搭售統計設定」完成登入，或取消強制使用 DSS。');
        return;
      }
    } catch (e) { /* 狀態查詢失敗不擋產生 */ }

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
})();
