// MD5 для версии на планшете (в браузере нет встроенного md5, а Яндекс.Диск сообщает именно md5 файла).
// Принимает Uint8Array, возвращает hex-строку.
function lcMd5Hex(bytes) {
  const add = (x, y) => (x + y) | 0;
  const S = [7, 12, 17, 22, 5, 9, 14, 20, 4, 11, 16, 23, 6, 10, 15, 21];
  const K = [];
  for (let i = 0; i < 64; i++) K[i] = Math.floor(Math.abs(Math.sin(i + 1)) * 4294967296) | 0;
  const len = bytes.length;
  const nBlocks = ((len + 8) >> 6) + 1;
  const total = nBlocks * 64;
  const buf = new Uint8Array(total);
  buf.set(bytes);
  buf[len] = 0x80;
  const dv = new DataView(buf.buffer);
  dv.setUint32(total - 8, (len << 3) >>> 0, true);
  dv.setUint32(total - 4, Math.floor(len / 536870912), true);
  let a0 = 0x67452301, b0 = 0xefcdab89 | 0, c0 = 0x98badcfe | 0, d0 = 0x10325476;
  const M = new Int32Array(16);
  for (let blk = 0; blk < nBlocks; blk++) {
    for (let i = 0; i < 16; i++) M[i] = dv.getInt32(blk * 64 + i * 4, true);
    let A = a0, B = b0, C = c0, D = d0;
    for (let i = 0; i < 64; i++) {
      let F, g;
      if (i < 16) { F = (B & C) | (~B & D); g = i; }
      else if (i < 32) { F = (D & B) | (~D & C); g = (5 * i + 1) % 16; }
      else if (i < 48) { F = B ^ C ^ D; g = (3 * i + 5) % 16; }
      else { F = C ^ (B | ~D); g = (7 * i) % 16; }
      F = add(add(add(F, A), K[i]), M[g]);
      A = D; D = C; C = B;
      const s = S[(i >> 4) * 4 + (i & 3)];
      B = add(B, (F << s) | (F >>> (32 - s)));
    }
    a0 = add(a0, A); b0 = add(b0, B); c0 = add(c0, C); d0 = add(d0, D);
  }
  return [a0, b0, c0, d0].map((v) => [0, 8, 16, 24].map((sh) => ((v >>> sh) & 255).toString(16).padStart(2, "0")).join("")).join("");
}
if (typeof module !== "undefined") module.exports = { lcMd5Hex };

// ============================================================
// Версия для планшета (браузер/PWA): вместо Electron-моста window.litAPI — реализация на
// IndexedDB + fetch. Программа (app.js) та же, что на ноутбуке; отличается только этот слой.
//  • проект хранится в IndexedDB (ключ "project"), теневая копия — "shadow";
//  • «Открыть файл» — системный выбор файла, «Сохранить как» — скачивание файла;
//  • файлы-вложения (презентации и т.п.) на планшет не переносятся — остаются на ноутбуке;
//  • синхронизация с Яндекс.Диском — тот же алгоритм, что в sync.js на ноутбуке (со слиянием версий),
//    но через fetch и вход с возвратом на эту же страницу (токен приходит в адресе после #).
// Важно: сайт «Конструктора по русскому» лежит на том же домене (github.io), поэтому
// база данных здесь называется иначе («literature-constructor»), чтобы данные не смешивались.
// ============================================================
(function () {
  if (window.litAPI) return;
  window.litTabletMode = true;

  const PROJECT_PATH = "Планшет/Мои материалы по литературе.json";
  window.litDefaultProjectPath = PROJECT_PATH;
  const DB_NAME = "literature-constructor";
  const YD_API = "https://cloud-api.yandex.net/v1/disk";
  // Файл на Диске — JS-обёртка вокруг JSON (literature.js); файл русского (lessons.js) лежит рядом и не затрагивается.
  const YD_REMOTE = "app:/literature.js";
  const WRAP_PREFIX = "window.__litPull(";
  const WRAP_SUFFIX = ");\n";
  const wrap = (jsonText) => WRAP_PREFIX + jsonText + WRAP_SUFFIX;
  function unwrap(text) {
    if (!text.startsWith(WRAP_PREFIX)) return text;
    const end = text.lastIndexOf(");");
    return text.slice(WRAP_PREFIX.length, end === -1 ? undefined : end);
  }
  const REDIRECT = location.origin + location.pathname.replace(/index\.html$/, "");
  window.litRedirectUri = REDIRECT;

  // ---------- IndexedDB ----------
  let dbPromise = null;
  const db = () =>
    dbPromise ||
    (dbPromise = new Promise((resolve, reject) => {
      const r = indexedDB.open(DB_NAME, 1);
      r.onupgradeneeded = () => r.result.createObjectStore("kv");
      r.onsuccess = () => resolve(r.result);
      r.onerror = () => reject(r.error);
    }));
  const kvGet = async (k) => {
    const d = await db();
    return new Promise((res, rej) => {
      const t = d.transaction("kv").objectStore("kv").get(k);
      t.onsuccess = () => res(t.result);
      t.onerror = () => rej(t.error);
    });
  };
  const kvSet = async (k, v) => {
    const d = await db();
    return new Promise((res, rej) => {
      const tx = d.transaction("kv", "readwrite");
      tx.objectStore("kv").put(v, k);
      tx.oncomplete = () => res();
      tx.onerror = () => rej(tx.error);
    });
  };
  const kvKeys = async () => {
    const d = await db();
    return new Promise((res, rej) => {
      const r = d.transaction("kv").objectStore("kv").getAllKeys();
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error);
    });
  };
  const kvDel = async (k) => {
    const d = await db();
    return new Promise((res, rej) => {
      const tx = d.transaction("kv", "readwrite");
      tx.objectStore("kv").delete(k);
      tx.oncomplete = () => res();
      tx.onerror = () => rej(tx.error);
    });
  };

  // Один раз перенимаем настройки входа (ID приложения, токен, посредник) у «Конструктора по русскому»:
  // сайт на том же домене, Яндекс-приложение одно — повторно вводить не нужно. Только чтение, чужую базу не трогаем.
  async function readLegacySync() {
    try {
      if (!indexedDB.databases) return null;
      const list = await indexedDB.databases();
      if (!list.some((d) => d.name === "lesson-constructor")) return null;
      const old = await new Promise((resolve, reject) => {
        const r = indexedDB.open("lesson-constructor");
        r.onsuccess = () => resolve(r.result);
        r.onerror = () => reject(r.error);
      });
      try {
        if (!old.objectStoreNames.contains("kv")) return null;
        return await new Promise((res, rej) => {
          const t = old.transaction("kv").objectStore("kv").get("sync");
          t.onsuccess = () => res(t.result || null);
          t.onerror = () => rej(t.error);
        });
      } finally {
        old.close();
      }
    } catch (e) {
      return null;
    }
  }

  // ---------- состояние синхронизации (аналог settings.json на ноутбуке) ----------
  let syncSt = {};
  const saveSync = async (patch) => {
    syncSt = { ...syncSt, ...patch };
    await kvSet("sync", syncSt);
  };
  let authError = null;
  const ready = (async () => {
    try {
      if (navigator.storage && navigator.storage.persist) navigator.storage.persist();
    } catch (e) {}
    syncSt = (await kvGet("sync")) || {};
    if (!syncSt.clientId && !syncSt.token && !syncSt.legacyChecked) {
      const old = await readLegacySync();
      syncSt = { ...syncSt, legacyChecked: true };
      if (old) syncSt = { ...syncSt, clientId: old.clientId || "", token: old.token || null, proxyUrl: old.proxyUrl || "" };
      await kvSet("sync", syncSt);
    }
    // Возврат со страницы входа Яндекса: токен приходит в адресе после #.
    if (location.hash && /access_token=|error=/.test(location.hash)) {
      const p = new URLSearchParams(location.hash.slice(1));
      if (p.get("access_token")) await saveSync({ token: p.get("access_token") });
      else authError = p.get("error_description") || p.get("error");
      history.replaceState(null, "", location.pathname + location.search);
    }
    if (authError) setTimeout(() => alert("Яндекс не выдал доступ: " + authError), 300);
  })();

  // ---------- вспомогательное ----------
  const utf8 = (text) => new TextEncoder().encode(text);
  const md5Of = (text) => lcMd5Hex(utf8(text));
  const sha256Of = async (text) => {
    const buf = await crypto.subtle.digest("SHA-256", utf8(text));
    return Array.from(new Uint8Array(buf)).map((b) => b.toString(16).padStart(2, "0")).join("");
  };
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const stripBom = (t) => (t.charCodeAt(0) === 0xfeff ? t.slice(1) : t);
  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  }
  // Системный выбор файла. input.click() вызывается сразу (в рамках нажатия пользователя).
  function pickTextFile(accept) {
    return new Promise((resolve) => {
      const input = document.createElement("input");
      input.type = "file";
      input.accept = accept;
      input.style.display = "none";
      document.body.appendChild(input);
      let settled = false;
      const finish = (v) => {
        if (settled) return;
        settled = true;
        input.remove();
        resolve(v);
      };
      input.addEventListener("change", () => {
        const f = input.files && input.files[0];
        if (!f) return finish(null);
        const reader = new FileReader();
        reader.onload = () => finish(stripBom(String(reader.result)));
        reader.onerror = () => finish(null);
        reader.readAsText(f, "utf-8");
      });
      input.addEventListener("cancel", () => finish(null));
      input.click();
    });
  }
  const errMsg = (e) => String((e && e.message) || e);
  // Та же защита, что на ноутбуке: «пустой» проект не должен затирать заполненный.
  const looksEmpty = (o) => {
    const isEmpty = (a) => !Array.isArray(a) || a.length === 0;
    return isEmpty(o && o.personalities) && isEmpty(o && o.works) && isEmpty(o && o.theoryHistory) && isEmpty(o && o.tags && o.tags.genre);
  };

  // ---------- проект ----------
  const project = {
    getLastProjectPath: async () => {
      await ready;
      return (await kvGet("project")) ? PROJECT_PATH : null;
    },
    readPath: async () => {
      await ready;
      const text = await kvGet("project");
      return text ? { ok: true, text, path: PROJECT_PATH } : { ok: false, error: "На этом устройстве проект ещё не сохранён" };
    },
    openDialog: async () => {
      const text = await pickTextFile(".json,application/json");
      if (text == null) return { ok: false };
      try {
        JSON.parse(text);
      } catch (e) {
        return { ok: false, error: "Это не файл проекта (не удалось прочитать JSON)" };
      }
      await ready;
      const prev = await kvGet("project");
      if (prev) await kvSet("project-prev", prev); // страховка: предыдущий проект остаётся в хранилище
      await kvSet("project", text);
      return { ok: true, text, path: PROJECT_PATH };
    },
    save: async (_path, text) => {
      try {
        await ready;
        const prev = await kvGet("project");
        if (prev && prev.length > 50000 && text.length < prev.length * 0.05) {
          return { ok: false, error: "Отказ сохранять: новые данные почти пустые, а в хранилище уже есть проект." };
        }
        await kvSet("project", text);
        return { ok: true };
      } catch (e) {
        return { ok: false, error: errMsg(e) };
      }
    },
    saveAsDialog: async (text, suggestedName) => {
      try {
        await ready;
        await kvSet("project", text);
        downloadBlob(new Blob([text], { type: "application/json" }), suggestedName || "Мои материалы по литературе.json");
        return { ok: true, path: PROJECT_PATH };
      } catch (e) {
        return { ok: false, error: errMsg(e) };
      }
    },
    backup: async () => {
      try {
        await ready;
        const text = await kvGet("project");
        if (!text) return { ok: false, error: "Проект ещё не сохранён" };
        const stamp = new Date().toISOString().slice(0, 16).replace(/[:T]/g, "-");
        downloadBlob(new Blob([text], { type: "application/json" }), `Литература — копия ${stamp}.json`);
        return { ok: true, path: "папка «Загрузки» (Downloads)" };
      } catch (e) {
        return { ok: false, error: errMsg(e) };
      }
    },
    // теневая копия — не чаще раза в минуту (проект и так лежит в том же хранилище)
    saveMirror: async (text) => {
      await ready;
      const now = Date.now();
      if (project._lastShadow && now - project._lastShadow < 60000) return { ok: true };
      project._lastShadow = now;
      await kvSet("shadow", text);
      return { ok: true };
    },
    readMirror: async () => {
      await ready;
      const text = await kvGet("shadow");
      return text ? { ok: true, text } : { ok: false };
    },
    // копия проекта перед рискованной операцией (слияние, восстановление) — одна на каждую метку
    snapshot: async (label, text) => {
      await ready;
      await kvSet("snapshot-" + label, text);
      return { ok: true };
    },
    // ежедневные автокопии (последние 14) — в хранилище программы
    dailyBackup: async () => {
      try {
        await ready;
        const text = await kvGet("project");
        if (!text) return { ok: false };
        const d = new Date();
        const p2 = (n) => String(n).padStart(2, "0");
        const key = "daily-" + d.getFullYear() + "-" + p2(d.getMonth() + 1) + "-" + p2(d.getDate());
        const keys = await kvKeys();
        let created = false;
        if (!keys.includes(key)) {
          await kvSet(key, text);
          created = true;
          keys.push(key);
        }
        const daily = keys.filter((k) => String(k).startsWith("daily-")).sort();
        while (daily.length > 14) await kvDel(daily.shift());
        return { ok: true, created };
      } catch (e) {
        return { ok: false, error: errMsg(e) };
      }
    },
    listBackups: async () => {
      try {
        await ready;
        const keys = (await kvKeys()).filter((k) => String(k).startsWith("daily-")).sort().reverse();
        const items = [];
        for (const k of keys) {
          const text = await kvGet(k);
          items.push({ id: k, date: k.slice(6).split("-").reverse().join("."), size: text ? text.length : 0 });
        }
        return { ok: true, items, where: "хранилище программы на этом устройстве" };
      } catch (e) {
        return { ok: false, error: errMsg(e) };
      }
    },
    readBackup: async (_path, id) => {
      await ready;
      const text = await kvGet(id);
      return text ? { ok: true, text } : { ok: false, error: "копия не найдена" };
    },
  };

  // «Отправить задание»: системное меню «Поделиться» (там выбирается мессенджер и файл уходит сразу)
  const share = {
    shareDocument: async ({ fileName, html, text }) => {
      const file = new File(["\ufeff" + html], fileName, { type: "application/msword" });
      try {
        if (navigator.canShare && navigator.canShare({ files: [file] })) {
          await navigator.share({ files: [file], title: fileName.replace(/\.doc$/, ""), text });
          return { ok: true, path: fileName };
        }
      } catch (e) {
        if (e && e.name === "AbortError") return { ok: false, cancelled: true };
      }
      downloadBlob(file, fileName);
      alert("Системное «Поделиться» для файла недоступно — файл сохранён в «Загрузки» (Downloads). Отправьте его в мессенджер оттуда.");
      return { ok: true, path: "Загрузки" };
    },
  };

  // ---------- вложения: на планшете не поддерживаются (файлы остаются на ноутбуке) ----------
  const attachments = {
    pickFiles: async () => {
      alert("Добавлять файлы-вложения (фото, аудио, презентации) можно на ноутбуке — на планшете они не хранятся.");
      return { ok: false };
    },
    saveAttachment: async () => ({ ok: false, error: "На планшете вложения не поддерживаются — добавьте файл на ноутбуке." }),
    saveAttachmentBuffer: async () => ({ ok: false, error: "На планшете вложения не поддерживаются — добавьте файл на ноутбуке." }),
    deleteAttachment: async () => ({ ok: true }),
    missingAttachments: async () => ({ ok: true, missing: [] }),
    restoreAttachment: async () => ({ ok: false }),
    fileUrl: (relPath) => "litfile:///" + encodeURIComponent(relPath),
  };

  // ---------- Яндекс.Диск ----------
  const getToken = () => syncSt.token || null;
  async function ydJson(method, pathAndQuery) {
    const token = getToken();
    if (!token) throw new Error("AUTH");
    const res = await fetch(YD_API + pathAndQuery, { method, headers: { Authorization: "OAuth " + token } });
    if (res.status === 401) throw new Error("AUTH");
    const body = await res.json().catch(() => ({}));
    return { status: res.status, body };
  }
  const ydPathQ = () => `path=${encodeURIComponent(YD_REMOTE)}`;
  async function validateToken() {
    const r = await ydJson("GET", "/resources?path=" + encodeURIComponent("app:/") + "&limit=1");
    if (r.status < 400 || r.status === 404) return;
    throw new Error("Яндекс.Диск ответил " + r.status + ": " + (r.body.description || r.body.message || r.body.error || "отказ в доступе"));
  }
  async function ydMeta() {
    const r = await ydJson("GET", `/resources?${ydPathQ()}&fields=md5,modified,size`);
    if (r.status === 404) return null;
    if (r.status >= 400) throw new Error(r.body.message || "Яндекс.Диск вернул ошибку " + r.status);
    return r.body;
  }
  // Скачивание: подключаем файл с Диска как <script> (без CORS); обёртка передаёт проект в window.__litPull.
  function loadScriptPayload(href) {
    return new Promise((resolve, reject) => {
      let got = null;
      const s = document.createElement("script");
      const cleanup = () => {
        clearTimeout(timer);
        delete window.__litPull;
        s.remove();
      };
      const timer = setTimeout(() => {
        cleanup();
        reject(new Error("Яндекс.Диск слишком долго отдаёт файл (больше 3 минут)"));
      }, 180000);
      window.__litPull = (obj) => {
        got = obj;
      };
      s.onload = () => {
        const obj = got;
        cleanup();
        if (obj == null) reject(new Error("Файл на Диске имеет неожиданный формат"));
        else resolve(JSON.stringify(obj));
      };
      s.onerror = () => {
        cleanup();
        reject(new Error("Не удалось скачать файл проекта с Яндекс.Диска (сервер скачивания не отдал его браузеру)"));
      };
      s.src = href;
      document.head.appendChild(s);
    });
  }
  const proxied = (href) => syncSt.proxyUrl.replace(/\/+$/, "") + "/?u=" + encodeURIComponent(href);
  async function ydDownloadText() {
    const r = await ydJson("GET", `/resources/download?${ydPathQ()}`);
    if (r.status >= 400) throw new Error(r.body.message || "Не удалось скачать файл (" + r.status + ")");
    const href = r.body.href;
    const tried = [];
    // Способ 0: через посредника (если указан в настройках синхронизации).
    if (syncSt.proxyUrl) {
      try {
        const res = await fetch(proxied(href));
        if (res.ok) return unwrap(stripBom(await res.text()));
        tried.push("через посредника: код " + res.status);
      } catch (e) {
        tried.push("через посредника: " + errMsg(e));
      }
    }
    // Способ 1: обычный запрос с заголовком Authorization.
    try {
      const res = await fetch(href, { headers: { Authorization: "OAuth " + getToken() } });
      if (res.ok) return unwrap(stripBom(await res.text()));
      tried.push("запрос с авторизацией: код " + res.status);
    } catch (e) {
      tried.push("запрос с авторизацией: " + errMsg(e));
    }
    // Способ 2: подключить файл как <script> (CORS для скриптов не нужен).
    try {
      return await loadScriptPayload(href);
    } catch (e) {
      tried.push("как скрипт: " + errMsg(e));
    }
    let host = "";
    try {
      host = new URL(href).host;
    } catch (e) {}
    throw new Error("Не удалось скачать файл проекта с Яндекс.Диска (" + host + "). Что пробовали — " + tried.join("; "));
  }
  async function ydUploadText(text) {
    const r = await ydJson("GET", `/resources/upload?${ydPathQ()}&overwrite=true`);
    if (r.status >= 400) throw new Error(r.body.message || "Не удалось начать загрузку (" + r.status + ")");
    let res;
    try {
      const body = new Blob([wrap(text)]);
      try {
        res = await fetch(r.body.href, { method: "PUT", body });
      } catch (e1) {
        if (!syncSt.proxyUrl) throw e1;
        res = await fetch(proxied(r.body.href), { method: "PUT", body }); // напрямую не вышло — через посредника
      }
    } catch (e) {
      throw new Error("Браузер не смог загрузить файл на Яндекс.Диск (" + errMsg(e) + "). Возможно, сервер загрузки не разрешает запросы из браузера.");
    }
    if (!res.ok && res.status !== 201 && res.status !== 202) throw new Error("Не удалось загрузить файл (" + res.status + ")");
  }
  const syncErr = (e) => {
    const isAuth = errMsg(e) === "AUTH";
    return { ok: false, auth: isAuth, error: isAuth ? "Нужно заново подключить Яндекс.Диск" : errMsg(e) };
  };
  const netHint = (e) => (e instanceof TypeError ? new Error("Нет связи с Яндекс.Диском (проверьте интернет): " + errMsg(e)) : e);

  // Имя файла на Диске поменялось: старые отпечатки больше не годятся.
  async function ensureRemoteName() {
    if (syncSt.remoteName !== YD_REMOTE) {
      await saveSync({ remoteName: YD_REMOTE, lastRemoteMd5: null, lastLocalHash: null });
      await kvDel("sync-base");
    }
  }

  async function pushText(localText, force) {
    await ensureRemoteName();
    let parsed;
    try {
      parsed = JSON.parse(localText);
    } catch (e) {
      return { ok: false, error: "Проект не удалось прочитать — загрузка отменена" };
    }
    const localHash = await sha256Of(localText);
    if (!force && syncSt.lastLocalHash === localHash) return { ok: true, status: "nochange" };
    const meta = await ydMeta();
    if (meta && !force && meta.md5 !== syncSt.lastRemoteMd5) return { ok: true, status: "conflict" };
    if (meta && looksEmpty(parsed) && (meta.size || 0) > 200) {
      return { ok: false, error: "Загрузка отменена: проект пустой, а на Диске уже есть данные" };
    }
    await ydUploadText(localText);
    const localMd5 = md5Of(wrap(localText));
    let md5 = localMd5;
    for (let i = 0; i < 5; i++) {
      const m = await ydMeta().catch(() => null);
      if (m && m.md5) {
        md5 = m.md5;
        if (m.md5 === localMd5) break;
      }
      await sleep(700);
    }
    await kvSet("sync-base", localText);
    await saveSync({ lastRemoteMd5: md5, lastLocalHash: localHash, lastSyncAt: Date.now() });
    return { ok: true, status: "pushed", lastSyncAt: syncSt.lastSyncAt };
  }

  const sync = {
    syncInfo: async () => {
      await ready;
      return { clientId: syncSt.clientId || "", connected: !!syncSt.token, lastSyncAt: syncSt.lastSyncAt || null, proxyUrl: syncSt.proxyUrl || "" };
    },
    syncSetProxy: async (url) => {
      await ready;
      await saveSync({ proxyUrl: String(url || "").trim() });
      return { ok: true };
    },
    syncSetClientId: async (id) => {
      await ready;
      await saveSync({ clientId: String(id || "").trim() });
      return { ok: true };
    },
    syncDisconnect: async () => {
      await ready;
      await saveSync({ token: null, lastRemoteMd5: null, lastLocalHash: null });
      await kvDel("sync-base");
      return { ok: true };
    },
    syncSetToken: async (token) => {
      await ready;
      try {
        await saveSync({ token: String(token || "").trim() });
        await validateToken();
        return { ok: true };
      } catch (e) {
        await saveSync({ token: null });
        return syncErr(netHint(e));
      }
    },
    // Вход: уходим на страницу Яндекса и возвращаемся сюда же с токеном в адресе (после #).
    syncConnect: async () => {
      await ready;
      if (!syncSt.clientId) return { ok: false, error: "Сначала укажите ID приложения" };
      location.href = `https://oauth.yandex.ru/authorize?response_type=token&client_id=${encodeURIComponent(syncSt.clientId)}&redirect_uri=${encodeURIComponent(REDIRECT)}`;
      return new Promise(() => {});
    },
    syncPull: async () => {
      await ready;
      try {
        await ensureRemoteName();
        const meta = await ydMeta();
        if (!meta) return { ok: true, status: "missing" };
        if (meta.md5 && meta.md5 === syncSt.lastRemoteMd5) return { ok: true, status: "nochange" };
        const text = await ydDownloadText();
        return { ok: true, status: "remote-newer", text, md5: meta.md5 || md5Of(wrap(text)), modified: meta.modified };
      } catch (e) {
        return syncErr(netHint(e));
      }
    },
    syncIsLocalChanged: async (localText, remoteMd5) => {
      await ready;
      if (syncSt.lastLocalHash && (await sha256Of(localText)) === syncSt.lastLocalHash) return false;
      if (remoteMd5 && md5Of(wrap(localText)) === remoteMd5) return false;
      return true;
    },
    syncGetBase: async () => {
      await ready;
      const text = await kvGet("sync-base");
      return { ok: true, text: text || null };
    },
    syncAcceptRemote: async (remoteMd5, remoteText) => {
      await ready;
      await kvSet("sync-base", remoteText);
      await saveSync({ lastRemoteMd5: remoteMd5, lastLocalHash: null });
      return { ok: true };
    },
    syncMarkSynced: async (remoteMd5, localText) => {
      await ready;
      await kvSet("sync-base", localText);
      await saveSync({ lastRemoteMd5: remoteMd5, lastLocalHash: await sha256Of(localText), lastSyncAt: Date.now() });
      return { ok: true, lastSyncAt: syncSt.lastSyncAt };
    },
    syncPush: async (localText, force) => {
      await ready;
      try {
        return await pushText(localText, !!force);
      } catch (e) {
        return syncErr(netHint(e));
      }
    },
  };

  window.litAPI = { ...project, ...share, ...attachments, ...sync };
})();
