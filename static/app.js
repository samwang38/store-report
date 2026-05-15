(() => {
  const $ = id => document.getElementById(id);

  const configCard   = $('configCard');
  const progressCard = $('progressCard');
  const downloadCard = $('downloadCard');
  const errorCard    = $('errorCard');

  const shopSelect    = $('shopId');
  const weekEndInput  = $('weekEnd');
  const weekRangeHint = $('weekRangeHint');
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

  weekEndInput.addEventListener('change', () => updateHint(weekEndInput.value));

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
    } catch (e) {
      console.warn('無法取得預設日期', e);
    }
  }

  function show(card) {
    [configCard, progressCard, downloadCard, errorCard].forEach(c => {
      c.hidden = (c !== card);
    });
  }

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
    const shopId = shopSelect.value;
    if (!shopId) { alert('請選擇門市'); return; }
    if (!wkEnd)  { alert('請選擇週結束日期'); return; }

    logBox.innerHTML = '';
    seenCount = 0;
    progressBar.style.width = '0%';
    progressLabel.textContent = '準備中…';
    show(progressCard);

    try {
      const res  = await fetch('/api/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ shopId, weekEnd: wkEnd }),
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
})();
