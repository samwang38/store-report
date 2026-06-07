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
  const resetBtn      = $('resetBtn');
  const retryBtn      = $('retryBtn');
  const errorMsg      = $('errorMsg');

  let pollTimer    = null;
  let currentJobId = null;
  let seenCount    = 0;

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
    if (!val) { yoyRangeHint.textContent = '留空則沿用週結束日'; return; }
    const end = new Date(val + 'T00:00:00');
    const y = end.getFullYear();
    const md = `${String(end.getMonth()+1).padStart(2,'0')}/${String(end.getDate()).padStart(2,'0')}`;
    const note = yoyEndInput.value ? '' : '（沿用週結束日）';
    yoyRangeHint.textContent = `年對年比較：${y-1}/01/01～${y-1}/${md} vs ${y}/01/01～${y}/${md}${note}`;
  }

  weekEndInput.addEventListener('change', () => { updateHint(weekEndInput.value); updateYoyHint(); });
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
    // 設定卡與來客數設定一起顯示／隱藏
    shoppertrakCard.hidden = (card !== configCard);
  }

  async function loadShoppertrakStatus() {
    try {
      const res = await fetch('/api/shoppertrak/status');
      const data = await res.json();
      if (!data.available) {
        stStatus.textContent = '（來客數模組未載入，將略過來客數）';
        return;
      }
      if (data.username) stUser.value = data.username;
      stStatus.textContent = data.hasCredentials ? '已儲存本機帳密' : '未儲存帳密';
    } catch (e) {
      stStatus.textContent = '';
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

  function stopPoll() {
    if (pollTimer) { clearInterval(pollTimer); pollTimer = null; }
  }

  function estimateProgress(messages) {
    const all = messages.join(' ');
    if (all.includes('儲存 Excel'))   return 95;
    if (all.includes('個人 6-9'))     return 78;
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

    logBox.innerHTML = '';
    seenCount = 0;
    progressBar.style.width = '0%';
    progressLabel.textContent = '準備中…';
    show(progressCard);

    try {
      const res  = await fetch('/api/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ shopId, weekEnd: wkEnd, yoyEnd: yoyEnd || '', employeeCount: empCount }),
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
})();
