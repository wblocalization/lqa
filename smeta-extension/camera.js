// Боковая панель не умеет сама спросить доступ к камере — спрашиваем здесь, в обычной вкладке, один раз.
navigator.mediaDevices.getUserMedia({ video: true, audio: false }).then((s) => {
  s.getTracks().forEach((t) => t.stop());
  document.getElementById('ic').textContent = '✅';
  document.getElementById('t').textContent = 'Готово! Камера разрешена';
  document.getElementById('p').textContent = 'Эту вкладку можно закрыть и снова включить «Жесты» в панели.';
  chrome.storage.local.set({ gesturesGranted: Date.now() });
  setTimeout(() => window.close(), 2500);
}).catch((e) => {
  document.getElementById('ic').textContent = '🙈';
  document.getElementById('t').textContent = 'Камера не разрешена';
  document.getElementById('p').textContent = `Нажмите на значок камеры в адресной строке → «Разрешить», потом обновите страницу. (${e.name})`;
});
