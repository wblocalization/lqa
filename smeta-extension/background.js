// Клик по иконке расширения открывает боковую панель — она не закрывается,
// пока переключаешься между Outlook и ВБ Диском.
chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true }).catch(console.error);
