// Тема до отрисовки, чтобы не мигало: выбранная кнопкой ☀/☾ или как в системе.
try {
  const t = localStorage.getItem('theme');
  if (t === 'light' || t === 'dark') document.documentElement.dataset.theme = t;
} catch { /* как в системе */ }
