/**
 * Список фото для сайта ДППиРР (Google Apps Script).
 * Отдаёт JSON: подпапки общей папки «Фото по проектам» и фото в каждой.
 * Развёртывание: script.google.com → Новый проект → вставить этот код →
 * «Начать развёртывание» → «Новое развёртывание» → тип «Веб-приложение» →
 * «Запуск от имени: Я», «Доступ: Все» → «Развернуть» → скопировать URL (…/exec).
 * URL вписывается в index.html: SHEET.photosApi.
 */
const ROOT_ID = '1-rWauGOqjzEZZInRhkTHfn40NPrm_N7S'; // папка «Фото по проектам»

function doGet() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('photos');
  if (hit) return json_(hit);
  const folders = {};
  const subs = DriveApp.getFolderById(ROOT_ID).getFolders();
  while (subs.hasNext()) {
    const f = subs.next();
    const files = [];
    const it = f.getFiles();
    while (it.hasNext()) {
      const x = it.next();
      if (String(x.getMimeType()).indexOf('image/') === 0) {
        files.push({ id: x.getId(), name: x.getName(), t: x.getDateCreated().getTime() });
      }
    }
    files.sort(function (a, b) { return b.t - a.t; });   // новые сверху
    folders[f.getName()] = { id: f.getId(), files: files };
  }
  const out = JSON.stringify({ updated: new Date().toISOString(), folders: folders });
  try { cache.put('photos', out, 300); } catch (e) {}     // кэш 5 минут
  return json_(out);
}

function json_(s) {
  return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON);
}
