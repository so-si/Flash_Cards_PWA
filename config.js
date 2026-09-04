window.FLASH_CARDS_CONFIG = {
  apiUrl: 'https://script.google.com/macros/s/AKfycbx-CSqlklOuwsEPyX4-89il2w-GFmZMqOB5wMszAOugUq2R3Q3EePivGNxQSjbPJWtX/exec'
};

window.addEventListener('load', () => {
  if (typeof window.refreshStats !== 'function') return;

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

      saveJson(STATS_KEY, remoteStats);
      saveJson(QUESTION_STATS_KEY, remoteQuestionStats);
      updateStatsDisplay();
    } catch (e) {
      console.warn(e);
    }
  };
});
