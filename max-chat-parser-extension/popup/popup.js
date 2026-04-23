"use strict";

const DEFAULT_SELECTORS = {
  chatRootSelector: "main, [role='main'], [class*='chat']",
  messageSelector: "[data-testid*='message'], [class*='message'], [role='listitem']",
  textSelector: "[data-testid*='text'], [class*='text'], [class*='content'], p, span, div",
  authorSelector: "[data-testid*='author'], [class*='author'], [class*='sender'], [class*='name']",
  timeSelector: "time, [data-testid*='time'], [class*='time'], [class*='timestamp']"
};

const els = {
  count: document.getElementById("count"),
  observeState: document.getElementById("observeState"),
  status: document.getElementById("status"),
  collectBtn: document.getElementById("collectBtn"),
  startObserveBtn: document.getElementById("startObserveBtn"),
  stopObserveBtn: document.getElementById("stopObserveBtn"),
  exportJsonBtn: document.getElementById("exportJsonBtn"),
  exportTxtBtn: document.getElementById("exportTxtBtn"),
  clearBtn: document.getElementById("clearBtn"),
  applySelectorsBtn: document.getElementById("applySelectorsBtn"),
  chatRootSelector: document.getElementById("chatRootSelector"),
  messageSelector: document.getElementById("messageSelector"),
  textSelector: document.getElementById("textSelector"),
  authorSelector: document.getElementById("authorSelector"),
  timeSelector: document.getElementById("timeSelector")
};

function setStatus(text, isError = false) {
  els.status.textContent = text;
  els.status.style.color = isError ? "#b21d1d" : "#1d2533";
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
  if (!tab || !tab.id) {
    throw new Error("Активная вкладка не найдена.");
  }
  if (!tab.url || !tab.url.startsWith("https://web.max.ru/")) {
    throw new Error("Откройте страницу чата на https://web.max.ru/.");
  }
  return tab;
}

async function injectContentScript(tabId) {
  await toPromiseChromeCall((cb) =>
    chrome.scripting.executeScript(
      {
        target: { tabId },
        files: ["content.js"]
      },
      cb
    )
  );
}

async function sendToChat(payload) {
  const tab = await getActiveTab();
  try {
    return await toPromiseChromeCall((cb) => chrome.tabs.sendMessage(tab.id, payload, cb));
  } catch (error) {
    const text = (error && error.message ? error.message : String(error)).toLowerCase();
    if (!text.includes("receiving end does not exist")) {
      throw error;
    }

    await injectContentScript(tab.id);
    return toPromiseChromeCall((cb) => chrome.tabs.sendMessage(tab.id, payload, cb));
  }
}

async function sendToBackground(payload) {
  return toPromiseChromeCall((cb) => chrome.runtime.sendMessage(payload, cb));
}

function setSelectorInputs(settings) {
  const merged = { ...DEFAULT_SELECTORS, ...(settings || {}) };
  els.chatRootSelector.value = merged.chatRootSelector;
  els.messageSelector.value = merged.messageSelector;
  els.textSelector.value = merged.textSelector;
  els.authorSelector.value = merged.authorSelector;
  els.timeSelector.value = merged.timeSelector;
}

function getSelectorInputs() {
  return {
    chatRootSelector: els.chatRootSelector.value.trim(),
    messageSelector: els.messageSelector.value.trim(),
    textSelector: els.textSelector.value.trim(),
    authorSelector: els.authorSelector.value.trim(),
    timeSelector: els.timeSelector.value.trim()
  };
}

function updateStateUi(state) {
  els.count.textContent = String(state.total || 0);
  els.observeState.textContent = state.observing ? "вкл." : "выкл.";
}

async function refreshState() {
  const state = await sendToChat({ type: "GET_STATE" });
  if (!state || !state.ok) {
    throw new Error(state && state.error ? state.error : "Не удалось получить состояние.");
  }
  updateStateUi(state);
  setSelectorInputs(state.settings);
  return state;
}

function toTxt(messages) {
  const lines = [];
  for (const [index, item] of messages.entries()) {
    const partTime = item.time ? `[${item.time}] ` : "";
    const partAuthor = item.author ? `${item.author}: ` : "";
    lines.push(`${index + 1}. ${partTime}${partAuthor}${item.text}`);
  }
  return lines.join("\n");
}

function buildFileName(ext) {
  const stamp = new Date().toISOString().replace(/[:.]/g, "-");
  return `max-chat-${stamp}.${ext}`;
}

async function exportMessages(format) {
  const state = await sendToChat({ type: "GET_STATE" });
  if (!state || !state.ok) {
    throw new Error(state && state.error ? state.error : "Состояние недоступно.");
  }

  if (!state.total) {
    await sendToChat({ type: "COLLECT_NOW" });
  }

  const payload = await sendToChat({ type: "GET_MESSAGES" });
  const messages = payload && payload.messages ? payload.messages : [];
  if (!messages.length) {
    throw new Error("Сообщения не найдены. Проверьте селекторы.");
  }

  if (format === "json") {
    const json = JSON.stringify(
      {
        exportedAt: new Date().toISOString(),
        sourceUrl: state.url || "",
        count: messages.length,
        messages
      },
      null,
      2
    );
    return sendToBackground({
      type: "DOWNLOAD_FILE",
      mimeType: "application/json",
      filename: buildFileName("json"),
      content: json
    });
  }

  return sendToBackground({
    type: "DOWNLOAD_FILE",
    mimeType: "text/plain",
    filename: buildFileName("txt"),
    content: toTxt(messages)
  });
}

async function collectNow() {
  const result = await sendToChat({ type: "COLLECT_NOW" });
  if (!result || !result.ok) {
    throw new Error(result && result.error ? result.error : "Сбор не выполнен.");
  }
  updateStateUi(result);
  const added = result.result ? result.result.added : 0;
  setStatus(`Собрано. Новых: ${added}. Всего: ${result.total}.`);
}

async function startObserver() {
  const state = await sendToChat({ type: "START_OBSERVER" });
  if (!state || !state.ok) {
    throw new Error(state && state.error ? state.error : "Не удалось запустить наблюдение.");
  }
  updateStateUi(state);
  setStatus("Наблюдение включено.");
}

async function stopObserver() {
  const state = await sendToChat({ type: "STOP_OBSERVER" });
  if (!state || !state.ok) {
    throw new Error(state && state.error ? state.error : "Не удалось остановить наблюдение.");
  }
  updateStateUi(state);
  setStatus("Наблюдение выключено.");
}

async function clearMessages() {
  const state = await sendToChat({ type: "CLEAR_MESSAGES" });
  if (!state || !state.ok) {
    throw new Error(state && state.error ? state.error : "Не удалось очистить список.");
  }
  updateStateUi(state);
  setStatus("Список сообщений очищен.");
}

async function applySelectors() {
  const state = await sendToChat({
    type: "SET_SELECTORS",
    settings: getSelectorInputs()
  });
  if (!state || !state.ok) {
    throw new Error(state && state.error ? state.error : "Не удалось применить селекторы.");
  }
  updateStateUi(state);
  setStatus("Селекторы сохранены и применены.");
}

async function runAction(handler) {
  try {
    setStatus("Выполняется...");
    await handler();
  } catch (error) {
    setStatus(error.message || String(error), true);
  }
}

els.collectBtn.addEventListener("click", () => runAction(collectNow));
els.startObserveBtn.addEventListener("click", () => runAction(startObserver));
els.stopObserveBtn.addEventListener("click", () => runAction(stopObserver));
els.clearBtn.addEventListener("click", () => runAction(clearMessages));
els.applySelectorsBtn.addEventListener("click", () => runAction(applySelectors));
els.exportJsonBtn.addEventListener("click", () =>
  runAction(async () => {
    const result = await exportMessages("json");
    if (!result || !result.ok) {
      throw new Error(result && result.error ? result.error : "Ошибка экспорта JSON.");
    }
    setStatus("JSON сохранен.");
  })
);
els.exportTxtBtn.addEventListener("click", () =>
  runAction(async () => {
    const result = await exportMessages("txt");
    if (!result || !result.ok) {
      throw new Error(result && result.error ? result.error : "Ошибка экспорта TXT.");
    }
    setStatus("TXT сохранен.");
  })
);

(async () => {
  try {
    setStatus("Проверка страницы...");
    await refreshState();
    setStatus("Готово.");
  } catch (error) {
    setSelectorInputs(DEFAULT_SELECTORS);
    setStatus(error.message || String(error), true);
  }
})();
