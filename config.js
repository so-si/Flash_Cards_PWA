window.FLASH_CARDS_CONFIG = {
  apiUrl: 'https://script.google.com/macros/s/AKfycbx-CSqlklOuwsEPyX4-89il2w-GFmZMqOB5wMszAOugUq2R3Q3EePivGNxQSjbPJWtX/exec'
};

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
        return `<tr><td>${formatDayLabel(k)}</td><td>${s.total}</td><td>${s.correct}</td><td>${pct(s.correct, s.total)}</td></tr>`;
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
      const previousTodayCount =
        remoteStats && remoteStats.today === today
          ? Number(remoteStats.todayCount || 0)
          : 0;

      const r = await apiRequest('getFlashStatsBundle');
      const b = r.bundle || {};
      const stats = b.stats || {};
      const daily = b.daily || {};
      const serverDate = daily.date || today;
      const serverTodayCount = Number(daily.count || 0);

      remoteStats = {
        total: Number(stats.total || 0),
        correct: Number(stats.correct || 0),
        todayCount:
          serverDate === today
            ? Math.max(previousTodayCount, serverTodayCount)
            : serverTodayCount,
        today: serverDate
      };

      remoteQuestionStats =
        b.questions && typeof b.questions === 'object'
          ? b.questions
          : {};

      remoteBreakdown = {
        subjects:
          b.subjects && typeof b.subjects === 'object'
            ? b.subjects
            : {},
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