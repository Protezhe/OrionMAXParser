"use strict";

const PROMPT_STORAGE_KEY = "maxChatPromptByChat";
const LEGACY_PROMPT_STORAGE_KEY = "maxChatPromptText";
const PARSE_STORAGE_KEY = "maxChatParseJsonByChat";
const DEFAULT_CHAT_KEY = "__global__";
const DEFAULT_PROMPT = `Меня интересуют следующие аттракционы:

Лунный экспресс / Disko Coaster 40

Астродром / Air Race 6.4

Торнадо / Family Swinger

Вальс часов / Tea Cup 12

Аэротакси / Mini Jet

Авиатор / Sky Roller 24

Для каждого укажи:

Во сколько открылся (время старта) и во сколько закрылся, если это указано

Если работал без остановок - просто укажи работал или не работал

Если был остановлен на ремонт, ТО или по другой причине - укажи время начала останова, время возобновления работы и причину

Если аттракцион не работал весь день - укажи это и причину

Формат вывода: краткая таблица или список по каждому аттракциону.
Не добавляй другие аттракционы, только эти 6.`;

let selectedChatKey = DEFAULT_CHAT_KEY;

const els = {
  status: document.getElementById("status"),
  currentTab: document.getElementById("currentTab"),
  collectBtn: document.getElementById("collectBtn"),
  chatSelect: document.getElementById("chatSelect"),
  refreshChatsBtn: document.getElementById("refreshChatsBtn"),
  dateFilterInput: document.getElementById("dateFilterInput"),
  dateHint: document.getElementById("dateHint"),
  promptInput: document.getElementById("promptInput"),
  dateJsonOutput: document.getElementById("dateJsonOutput"),
  copyResultBtn: document.getElementById("copyResultBtn"),
  saveChatJsonBtn: document.getElementById("saveChatJsonBtn")
};

let promptSaveTimer = null;

function setStatus(text, isError = false) {
  els.status.textContent = text;
  els.status.style.color = isError ? "#b21d1d" : "#1d2533";
}

function pad2(value) {
  return String(value).padStart(2, "0");
}

function formatDisplayDate(dateKey) {
  if (!dateKey) return "";
  const [y, m, d] = dateKey.split("-");
  const months = ["янв", "фев", "мар", "апр", "мая", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"];
  return `${parseInt(d, 10)} ${months[parseInt(m, 10) - 1]} ${y}`;
}

function getPromptText() {
  return typeof els.promptInput.value === "string" ? els.promptInput.value.trim() : "";
}

function parseChatKeyFromUrl(url) {
  const raw = typeof url === "string" ? url.trim() : "";
  if (!raw) return DEFAULT_CHAT_KEY;
  try {
    const parsed = new URL(raw);
    const path = parsed.pathname ? parsed.pathname.trim() : "";
    const hash = parsed.hash ? parsed.hash.trim() : "";
    const search = parsed.search ? parsed.search.trim() : "";
    if (path && path !== "/") return `${parsed.origin}${path}`;
    if (hash) return `${parsed.origin}${hash}`;
    if (search) return `${parsed.origin}${search}`;
    return `${parsed.origin}/`;
  } catch (_e) {
    return raw;
  }
}

function toPromiseChromeCall(callbackStyleFn) {
  return new Promise((resolve, reject) => {
    callbackStyleFn((result) => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
        return;
      }
      resolve(result);
    });
  });
}

async function getActiveTab() {
  const tabs = await toPromiseChromeCall((cb) =>
    chrome.tabs.query({ active: true, currentWindow: true }, cb)
  );
  const [tab] = tabs || [];
  if (!tab || !tab.id) throw new Error("Активная вкладка не найдена.");
  if (!tab.url || !tab.url.startsWith("https://web.max.ru/")) {
    throw new Error("Откройте страницу чата на https://web.max.ru/.");
  }
  return tab;
}

async function getActiveChatContext() {
  const tab = await getActiveTab();
  return {
    tab,
    sourceUrl: tab.url || "",
    chatKey: parseChatKeyFromUrl(tab.url || "")
  };
}

async function injectContentScript(tabId) {
  await toPromiseChromeCall((cb) =>
    chrome.scripting.executeScript({ target: { tabId }, files: ["content.js"] }, cb)
  );
}

async function sendToChat(payload) {
  const tab = await getActiveTab();
  try {
    return await toPromiseChromeCall((cb) => chrome.tabs.sendMessage(tab.id, payload, cb));
  } catch (error) {
    const text = (error && error.message ? error.message : String(error)).toLowerCase();
    if (!text.includes("receiving end does not exist")) throw error;
    await injectContentScript(tab.id);
    return toPromiseChromeCall((cb) => chrome.tabs.sendMessage(tab.id, payload, cb));
  }
}

async function sendToBackground(payload) {
  return toPromiseChromeCall((cb) => chrome.runtime.sendMessage(payload, cb));
}

async function loadPromptText() {
  const context = await getActiveChatContext();
  const data = await toPromiseChromeCall((cb) =>
    chrome.storage.local.get([PROMPT_STORAGE_KEY, LEGACY_PROMPT_STORAGE_KEY], cb)
  );
  const byChat = data[PROMPT_STORAGE_KEY];
  if (byChat && typeof byChat === "object" && typeof byChat[context.chatKey] === "string") {
    return byChat[context.chatKey];
  }
  const legacyPrompt = data[LEGACY_PROMPT_STORAGE_KEY];
  if (typeof legacyPrompt === "string") {
    const migrated = {
      ...(byChat && typeof byChat === "object" ? byChat : {}),
      [DEFAULT_CHAT_KEY]: legacyPrompt
    };
    await toPromiseChromeCall((cb) =>
      chrome.storage.local.set({ [PROMPT_STORAGE_KEY]: migrated }, cb)
    );
    if (context.chatKey === DEFAULT_CHAT_KEY) return legacyPrompt;
  }
  await savePromptText(DEFAULT_PROMPT);
  return DEFAULT_PROMPT;
}

async function savePromptText(promptText) {
  const context = await getActiveChatContext();
  const data = await toPromiseChromeCall((cb) =>
    chrome.storage.local.get(PROMPT_STORAGE_KEY, cb)
  );
  const byChat =
    data[PROMPT_STORAGE_KEY] && typeof data[PROMPT_STORAGE_KEY] === "object"
      ? data[PROMPT_STORAGE_KEY]
      : {};
  await toPromiseChromeCall((cb) =>
    chrome.storage.local.set(
      { [PROMPT_STORAGE_KEY]: { ...byChat, [context.chatKey]: promptText } },
      cb
    )
  );
}

function schedulePromptSave() {
  if (promptSaveTimer) clearTimeout(promptSaveTimer);
  promptSaveTimer = setTimeout(() => {
    void (async () => {
      try {
        await savePromptText(els.promptInput.value || "");
      } catch (_e) { /* ignore while typing */ }
    })();
  }, 300);
}

async function loadChatList() {
  try {
    const data = await toPromiseChromeCall((cb) =>
      chrome.storage.local.get(PARSE_STORAGE_KEY, cb)
    );
    const byChat = data[PARSE_STORAGE_KEY];
    if (!byChat || typeof byChat !== "object" || Array.isArray(byChat)) {
      return [];
    }
    return Object.keys(byChat).filter((k) => k !== DEFAULT_CHAT_KEY || true);
  } catch (_e) {
    return [];
  }
}

async function loadPromptTextForChat(chatKey) {
  try {
    const data = await toPromiseChromeCall((cb) =>
      chrome.storage.local.get([PROMPT_STORAGE_KEY, LEGACY_PROMPT_STORAGE_KEY], cb)
    );
    const byChat = data[PROMPT_STORAGE_KEY];
    if (byChat && typeof byChat === "object" && typeof byChat[chatKey] === "string") {
      return byChat[chatKey];
    }
    return null;
  } catch (_e) {
    return null;
  }
}

function shortLabel(key) {
  try {
    const url = new URL(key);
    const segs = url.pathname.replace(/^\//, "").split("/").filter(Boolean);
    if (segs.length) return segs[segs.length - 1];
    return url.hostname.replace("web.", "");
  } catch (_e) {
    return key.length > 40 ? key.slice(0, 40) + "…" : key;
  }
}

function populateChatSelect(keys, currentKey) {
  const sel = els.chatSelect;
  while (sel.options.length > 0) {
    sel.remove(0);
  }
  if (!keys.length) {
    const opt = document.createElement("option");
    opt.value = "";
    opt.textContent = "— нет данных —";
    opt.disabled = true;
    sel.appendChild(opt);
    return;
  }
  keys.forEach((key) => {
    const opt = document.createElement("option");
    opt.value = key;
    opt.textContent = shortLabel(key);
    sel.appendChild(opt);
  });
  if (currentKey && keys.includes(currentKey)) {
    sel.value = currentKey;
  } else if (keys.length > 0) {
    sel.value = keys[0];
  }
}

async function collectNow() {
  const result = await sendToChat({ type: "COLLECT_NOW" });
  if (!result || !result.ok) {
    throw new Error(result && result.error ? result.error : "Сбор не выполнен.");
  }
  const total = result.total || 0;
  const added = result.result ? result.result.added : 0;
  setStatus(`собрано ${total}, новых ${added}`);
  await syncMessagesToStore();
  await refreshChatList();
  updateDateHint();
  updateOutput();
}

async function syncMessagesToStore() {
  const messages = await (async () => {
    const payload = await sendToChat({ type: "GET_MESSAGES" });
    return payload && payload.messages ? payload.messages : [];
  })();
  const context = await getActiveChatContext();
  const result = await sendToBackground({
    type: "APPEND_PARSE_MESSAGES",
    messages,
    chatKey: context.chatKey,
    sourceUrl: context.sourceUrl
  });
  if (!result || !result.ok) {
    throw new Error(result && result.error ? result.error : "Не удалось сохранить в storage.");
  }
  return result;
}

function filterMessagesByDate(messages, dateKey) {
  if (!dateKey) return messages;
  return messages.filter((msg) => msg.messageDate === dateKey);
}

function buildOutput(dateKey, messages, prompt) {
  const filtered = filterMessagesByDate(messages, dateKey);
  if (!filtered.length) {
    return prompt
      ? `# === PROMPT ===\n${prompt}\n\n# Нет сообщений за выбранную дату.`
      : "# Нет сообщений за выбранную дату.";
  }
  const lines = [];
  if (prompt) {
    lines.push("# === PROMPT ===");
    lines.push(prompt);
    lines.push("");
  }
  lines.push("# === MESSAGES ===");
  for (const msg of filtered) {
    const date = msg.messageDate || "";
    const time = msg.time || "";
    const header = [date, time].filter(Boolean).join(" ");
    lines.push(header ? `[${header}] ${msg.text}` : msg.text);
  }
  return lines.join("\n");
}

async function updateDateHintForChat(chatKey, sourceUrl) {
  try {
    const payload = await sendToBackground({
      type: "GET_PARSE_MESSAGES",
      chatKey,
      sourceUrl
    });
    if (!payload || !payload.ok || !Array.isArray(payload.messages)) {
      els.dateHint.textContent = "";
      return;
    }
    const messages = payload.messages;
    const dates = [...new Set(messages.map((m) => m.messageDate).filter(Boolean))].sort();
    if (!dates.length) {
      els.dateHint.textContent = "";
      return;
    }
    const first = formatDisplayDate(dates[0]);
    const last = formatDisplayDate(dates[dates.length - 1]);
    els.dateHint.textContent = dates.length === 1
      ? `данные за ${first}`
      : `данные с ${first} по ${last}`;
  } catch (_e) {
    els.dateHint.textContent = "";
  }
}

function updateDateHint() {
  const key = els.chatSelect.value || selectedChatKey;
  const url = key === selectedChatKey ? "" : key;
  void updateDateHintForChat(key, url);
}

async function updateOutputForChat(chatKey, sourceUrl) {
  try {
    const payload = await sendToBackground({
      type: "GET_PARSE_MESSAGES",
      chatKey,
      sourceUrl
    });
    if (!payload || !payload.ok || !Array.isArray(payload.messages)) {
      els.dateJsonOutput.value = "";
      return;
    }
    const dateKey = els.dateFilterInput.value || "";
    const prompt = getPromptText();
    els.dateJsonOutput.value = buildOutput(dateKey, payload.messages, prompt);
  } catch (_e) {
    els.dateJsonOutput.value = "";
  }
}

function updateOutput() {
  const key = els.chatSelect.value || selectedChatKey;
  const url = key === selectedChatKey ? "" : key;
  void updateOutputForChat(key, url);
}

function buildSafeFilename(raw) {
  return String(raw || "").replace(/[^a-zA-Zа-яА-ЯёЁ0-9_\-.]/g, "_").slice(0, 80) || "chat";
}

async function saveChatJson() {
  const key = els.chatSelect.value || selectedChatKey;
  const url = key === selectedChatKey ? "" : key;
  const payload = await sendToBackground({
    type: "GET_PARSE_MESSAGES",
    chatKey: key,
    sourceUrl: url
  });
  if (!payload || !payload.ok) {
    throw new Error(payload && payload.error ? payload.error : "Не удалось прочитать данные чата.");
  }
  const messages = Array.isArray(payload.messages) ? payload.messages : [];
  if (!messages.length) {
    throw new Error("В этом чате нет сообщений. Нажмите «Собрать».");
  }
  const label = buildSafeFilename(key);
  const prompt = getPromptText();
  const lines = prompt ? [`# === PROMPT ===\n${prompt}\n# === MESSAGES ===`] : ["# === MESSAGES ==="];
  for (const msg of messages) {
    const date = msg.messageDate || "";
    const time = msg.time || "";
    const header = [date, time].filter(Boolean).join(" ");
    lines.push(header ? `[${header}] ${msg.text}` : msg.text);
  }
  const content = lines.join("\n");
  const result = await sendToBackground({
    type: "DOWNLOAD_FILE",
    mimeType: "text/plain",
    filename: `max-chat-${label}.txt`,
    content
  });
  if (!result || !result.ok) {
    throw new Error(result && result.error ? result.error : "Ошибка сохранения.");
  }
  setStatus("Чат сохранён.");
}

async function refreshChatList() {
  const context = await getActiveChatContext();
  selectedChatKey = context.chatKey;
  els.currentTab.textContent = shortLabel(context.chatKey);
  const keys = await loadChatList();
  populateChatSelect(keys, context.chatKey);
  await updateDateHintForChat(keys.includes(context.chatKey) ? context.chatKey : (keys[0] || context.chatKey), "");
  updateOutput();
}

async function runAction(handler) {
  try {
    setStatus("...");
    await handler();
  } catch (error) {
    setStatus(error.message || String(error), true);
  }
}

els.collectBtn.addEventListener("click", () => runAction(collectNow));

els.chatSelect.addEventListener("change", async () => {
  const key = els.chatSelect.value;
  if (!key) return;
  selectedChatKey = key;
  const savedPrompt = await loadPromptTextForChat(key);
  if (savedPrompt !== null) {
    els.promptInput.value = savedPrompt;
  }
  updateDateHint();
  updateOutput();
});

els.refreshChatsBtn.addEventListener("click", () => runAction(refreshChatList));

els.dateFilterInput.addEventListener("change", updateOutput);

els.promptInput.addEventListener("input", () => {
  schedulePromptSave();
  updateOutput();
});

els.promptInput.addEventListener("blur", () =>
  runAction(async () => {
    await savePromptText(els.promptInput.value || "");
    setStatus("Промт сохранён.");
  })
);

els.copyResultBtn.addEventListener("click", async () => {
  const text = els.dateJsonOutput.value;
  if (!text) return;
  try {
    await navigator.clipboard.writeText(text);
    setStatus("Скопировано!");
  } catch (_e) {
    setStatus("Ошибка копирования", true);
  }
});

els.saveChatJsonBtn.addEventListener("click", () => runAction(saveChatJson));

(async () => {
  try {
    setStatus("...");
    els.promptInput.value = await loadPromptText();
    await refreshChatList();
    setStatus("готово");
  } catch (error) {
    if (!els.promptInput.value) els.promptInput.value = DEFAULT_PROMPT;
    setStatus(error.message || String(error), true);
  }
})();
