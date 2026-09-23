window.FLASH_CARDS_CONFIG = {
  apiUrl: 'https://script.google.com/macros/s/AKfycbx-CSqlklOuwsEPyX4-89il2w-GFmZMqOB5wMszAOugUq2R3Q3EePivGNxQSjbPJWtX/exec'
};

const FLASH_APP_VERSION = '20.3';

function installAppUpdateUi_() {
  const title = document.querySelector('.title span');
  if (title) title.textContent = 'v' + FLASH_APP_VERSION;

  const updateCardsBtn = document.getElementById('updateCardsBtn');
  const row = updateCardsBtn ? updateCardsBtn.closest('.row') : null;
  if (row && !document.getElementById('appUpdateBtn')) {
    const btn = document.createElement('button');
    btn.className = 'btn small';
    btn.id = 'appUpdateBtn';
    btn.textContent = '最新版を確認';
    btn.onclick = () => checkForAppUpdate_(true);
    row.appendChild(btn);
  }

  const topics = document.getElementById('topics');
  const topicPanel = topics ? topics.closest('.panel') : null;
  if (topicPanel && !document.getElementById('startBtnTopics')) {
    const btn = document.createElement('button');
    btn.className = 'btn primary';
    btn.id = 'startBtnTopics';
    btn.style.cssText = 'width:100%;font-size:21px;padding:17px;margin:-2px 0 16px';
    btn.textContent = '学習開始';
    btn.onclick = () => startStudy();
    topicPanel.insertAdjacentElement('afterend', btn);
  }
}

async function checkForAppUpdate_(manual) {
  const btn = document.getElementById('appUpdateBtn');
  try {
    if (btn && manual) btn.textContent = '確認中…';
    const r = await fetch('./version.json?t=' + Date.now(), {cache:'no-store'});
    if (!r.ok) throw new Error('version check failed');
    const v = await r.json();
    const latest = String(v.version || '').trim();
    if (latest && latest !== FLASH_APP_VERSION) {
      if (btn) {
        btn.textContent = 'v' + latest + 'へ更新';
        btn.onclick = () => applyAppUpdate_(latest);
      }
      if (manual && typeof setMessage === 'function') setMessage('最新版 v' + latest + ' があります。');
    } else {
      if (btn) {
        btn.textContent = '最新版です';
        setTimeout(() => { if (btn) btn.textContent = '最新版を確認'; }, 1800);
      }
      if (manual && typeof setMessage === 'function') setMessage('現在のアプリは最新版です。');
    }
  } catch (e) {
    if (btn) btn.textContent = '最新版を確認';
    if (manual && typeof setMessage === 'function') setMessage('更新確認に失敗しました。通信状態を確認してください。');
  }
}

let flashReloading_ = false;
async function applyAppUpdate_(latest) {
  const btn = document.getElementById('appUpdateBtn');
  try {
    if (btn) btn.textContent = '更新中…';
    if (!('serviceWorker' in navigator)) {
      location.replace('./?v=' + encodeURIComponent(latest) + '&t=' + Date.now());
      return;
    }
    const reg = await navigator.serviceWorker.register('./service-worker.js', {updateViaCache:'none'});
    await reg.update();
    if (reg.waiting) reg.waiting.postMessage({type:'SKIP_WAITING'});
    await new Promise(resolve => setTimeout(resolve, 600));
    flashReloading_ = true;
    location.replace('./?v=' + encodeURIComponent(latest) + '&t=' + Date.now());
  } catch (e) {
    if (btn) btn.textContent = '更新を再試行';
    if (typeof setMessage === 'function') setMessage('アプリ更新に失敗しました。もう一度押してください。');
  }
}

if ('serviceWorker' in navigator) {
  navigator.serviceWorker.addEventListener('controllerchange', () => {
    if (!flashReloading_) {
      flashReloading_ = true;
      location.reload();
    }
  });
}

// Install the sync patch before index.html's DOMContentLoaded init runs.
document.addEventListener('DOMContentLoaded', () => {
  installAppUpdateUi_();

  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('./service-worker.js', {updateViaCache:'none'})
      .then(reg => reg.update())
      .catch(() => {});
  }

  if (typeof window.syncPending === 'function') {
    window.syncPending = async function syncPendingPatched() {
      if (!navigator.onLine || !token()) return;
      if (syncPromise) return syncPromise;

      syncPromise = (async () => {
        const CONCURRENCY = 6;
        let lastError = '';

        while (navigator.onLine && pendingEvents().length) {
          const batch = pendingEvents().slice(0, CONCURRENCY);
          const results = await Promise.all(batch.map(async ev => {
            try {
              const result = await apiRequest('saveFlashResult', ev);
              return {ok:true, ev, result};
            } catch (e) {
              return {ok:false, ev, error:e && e.message ? e.message : String(e)};
            }
          }));

          const successIds = new Set();
          for (const item of results) {
            if (!item.ok) {
              if (!lastError) lastError = item.error;
              continue;
            }
            successIds.add(item.ev.eventId);
            if (!(item.result && item.result.duplicate)) promoteSyncedEvent(item.ev);
          }

          if (successIds.size) {
            const current = pendingEvents();
            setPending(current.filter(x => !successIds.has(x.eventId)));
          }

          if (results.some(x => !x.ok)) break;
        }

        if (lastError) {
          const n = pendingEvents().length;
          setMessage(`未同期 ${n}件：${lastError}`);
        }
      })();

      try {
        await syncPromise;
      } finally {
        syncPromise = null;
      }

      if (navigator.onLine && token()) await refreshStats();

      const remaining = pendingEvents().length;
      if (!remaining) setMessage('回答履歴を同期しました。');
    };
  }

  setTimeout(() => checkForAppUpdate_(false), 1200);
});

window.addEventListener('load', () => {
  if (typeof window.refreshStats !== 'function') return;

  const startBtn = document.getElementById('startBtn');
  const performanceTitle = document.querySelector('.performance-title');
  const performancePanel = performanceTitle ? performanceTitle.closest('.panel') : null;
  if (startBtn && performancePanel) {
    startBtn.insertAdjacentElement('afterend', performancePanel);
  }

  function rankingTablePatched(keys, dayMap, today) {
    const top = keys.slice(0, 10);
    return '<div class="stats-table-wrap"><table class="stats-table"><thead><tr><th>順位</th><th>日付</th><th>回答</th><th>正答率</th></tr></thead><tbody>' +
      top.map((k, i) => {
        const s = dayMap[k];
        const todayStyle = k === today ? ' style="font-weight:800"' : '';
        return `<tr${todayStyle}><td>${i + 1}</td><td>${formatDayLabel(k)}</td><td>${s.total}</td><td>${pct(s.correct, s.total)}</td></tr>`;
      }).join('') +
      '</tbody></table></div>';
  }

  function todayRankNote(sortedKeys, today, label) {
    const index = sortedKeys.indexOf(today);
    if (index < 10) return '';
    return `<div class="stats-note">今日：${index + 1}位 / ${label}${sortedKeys.length}日</div>`;
  }

  window.renderPerformanceStats = function renderPerformanceStatsPatched() {
    const subjectBox = $('subjectStats');
    const dailyBox = $('dailyStats');
    const rankingBox = $('dailyRankings');
    if (!subjectBox || !dailyBox || !rankingBox) return;

    const subjects = combinedSubjectStats();
    const subjectKeys = Object.keys(subjects).sort(subjectSort);
    if (!subjectKeys.length) {
      subjectBox.innerHTML = '<div class="stats-empty">まだ学習データがありません。</div>';
    } else {
      subjectBox.innerHTML = '<div class="stats-table-wrap"><table class="stats-table"><thead><tr><th>科目</th><th>回答</th><th>正解</th><th>正答率</th></tr></thead><tbody>' +
        subjectKeys.map(k => {
          const s = subjects[k];
          return `<tr><td>${escapeHtml(k)}</td><td>${s.total}</td><td>${s.correct}</td><td>${pct(s.correct, s.total)}</td></tr>`;
        }).join('') +
        '</tbody></table></div>';
    }

    const dayMap = combinedDailyStats();
    const allDayKeys = Object.keys(dayMap).sort().reverse();
    const dayKeys = allDayKeys.slice(0, 30);
    if (!dayKeys.length) {
      dailyBox.innerHTML = '<div class="stats-empty">まだ日別データがありません。</div>';
      rankingBox.innerHTML = '<div class="stats-empty">まだランキング対象データがありません。</div>';
      return;
    }

    dailyBox.innerHTML = '<div class="stats-table-wrap daily"><table class="stats-table"><thead><tr><th>日付</th><th>回答</th><th>正解</th><th>正答率</th></tr></thead><tbody>' +
      dayKeys.map(k => {
        const s = dayMap[k];
        return `<tr><td>${formatDayLabel(k)}</td><td>${s.total}</td><td>${s.correct}</td><td>${pct(s.correct,s.total)}</td></tr>`;
      }).join('') +
      '</tbody></table></div>';

    const today = jstDate();
    const byCount = [...allDayKeys].sort((a, b) =>
      dayMap[b].total - dayMap[a].total ||
      dayMap[b].correct - dayMap[a].correct ||
      b.localeCompare(a)
    );

    const accuracyKeys = allDayKeys.filter(k => Number(dayMap[k].total || 0) >= 50);
    const byAccuracy = [...accuracyKeys].sort((a, b) => {
      const ar = dayMap[a].correct / dayMap[a].total;
      const br = dayMap[b].correct / dayMap[b].total;
      return br - ar || dayMap[b].total - dayMap[a].total || b.localeCompare(a);
    });

    const countNote = todayRankNote(byCount, today, '全');
    let accuracyNote = '';
    if (dayMap[today] && Number(dayMap[today].total || 0) < 50) {
      accuracyNote = '<div class="stats-note">今日は正答率ランキング対象外（50問未満）</div>';
    } else {
      accuracyNote = todayRankNote(byAccuracy, today, '対象');
    }

    rankingBox.innerHTML =
      '<div class="performance-details">' +
        '<div class="muted" style="margin-top:4px">回答数ランキング（過去全日 TOP10）</div>' +
        rankingTablePatched(byCount, dayMap, today) +
        countNote +
        '<div class="muted" style="margin-top:12px">正答率ランキング（50問以上・過去全日 TOP10）</div>' +
        rankingTablePatched(byAccuracy, dayMap, today) +
        accuracyNote +
      '</div>';
  };

  window.refreshStats = async function refreshStatsPatched() {
    if (!navigator.onLine || !token()) return;
    try {
      const today = jstDate();
      const r = await apiRequest('getFlashStatsBundle');
      const b = r.bundle || {};
      const stats = b.stats || {};
      const daily = b.daily || {};
      const serverDate = daily.date || today;
      const serverTodayCount = Number(daily.count || 0);

      remoteStats = {
        total: Number(stats.total || 0),
        correct: Number(stats.correct || 0),
        todayCount: serverTodayCount,
        today: serverDate
      };

      remoteQuestionStats = b.questions && typeof b.questions === 'object' ? b.questions : {};
      remoteBreakdown = {
        subjects: b.subjects && typeof b.subjects === 'object' ? b.subjects : {},
        days: Array.isArray(b.days) ? b.days : []
      };

      saveJson(STATS_KEY, remoteStats);
      saveJson(QUESTION_STATS_KEY, remoteQuestionStats);
      saveJson(BREAKDOWN_KEY, remoteBreakdown);
      updateStatsDisplay();
    } catch (e) {
      console.warn(e);
    }
  };

  updateStatsDisplay();
});
