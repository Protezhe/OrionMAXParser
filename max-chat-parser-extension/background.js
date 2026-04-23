"use strict";

const PARSE_STORAGE_KEY = "maxChatParseJsonByChat";
const LEGACY_PARSE_STORAGE_KEY = "maxChatParseJson";
const DEFAULT_CHAT_KEY = "__global__";

function toDataUrl(content, mimeType) {
  const encoded = encodeURIComponent(content);
  return `data:${mimeType};charset=utf-8,${encoded}`;
}

function storageGetLocal(key) {
  return new Promise((resolve, reject) => {
    chrome.storage.local.get(key, (result) => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
        return;
      }
      resolve(result || {});
    });
  });
}

function storageSetLocal(value) {
  return new Promise((resolve, reject) => {
    chrome.storage.local.set(value, () => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
        return;
      }
      resolve();
    });
  });
}

function normalizeText(value) {
  return typeof value === "string" ? value.trim() : "";
}

function parseChatKeyFromUrl(url) {
  const raw = normalizeText(url);
  if (!raw) {
    return DEFAULT_CHAT_KEY;
  }

  try {
    const parsed = new URL(raw);
    const path = normalizeText(parsed.pathname);
    const search = normalizeText(parsed.search);
    const hash = normalizeText(parsed.hash);

    if (path && path !== "/") {
      return `${parsed.origin}${path}`;
    }
    if (hash) {
      return `${parsed.origin}${hash}`;
    }
    if (search) {
      return `${parsed.origin}${search}`;
    }
    return `${parsed.origin}/`;
  } catch (_error) {
    return raw;
  }
}

function normalizeChatKey(chatKey, sourceUrl) {
  const explicit = parseChatKeyFromUrl(chatKey);
  if (explicit !== DEFAULT_CHAT_KEY) {
    return explicit;
  }
  return parseChatKeyFromUrl(sourceUrl);
}

function messageKey(message) {
  const text = normalizeText(message && message.text);
  const messageDate = normalizeText(message && message.messageDate);
  const time = normalizeText(message && message.time);
  return `${messageDate}|${time}|${text}`;
}

async function loadSeedParseJson() {
  try {
    const response = await fetch(chrome.runtime.getURL("parse.json"));
    if (!response.ok) {
      return { updatedAt: null, messages: [] };
    }
    const parsed = await response.json();
    const messages = Array.isArray(parsed && parsed.messages) ? parsed.messages : [];
    return {
      updatedAt: parsed && parsed.updatedAt ? String(parsed.updatedAt) : null,
      messages
    };
  } catch (_error) {
    return { updatedAt: null, messages: [] };
  }
}

async function getParseJsonData() {
  const data = await storageGetLocal([PARSE_STORAGE_KEY, LEGACY_PARSE_STORAGE_KEY]);
  const byChat = data[PARSE_STORAGE_KEY];
  if (byChat && typeof byChat === "object" && !Array.isArray(byChat)) {
    return byChat;
  }

  const legacy = data[LEGACY_PARSE_STORAGE_KEY];
  if (legacy && Array.isArray(legacy.messages)) {
    const migrated = {
      [DEFAULT_CHAT_KEY]: legacy
    };
    await storageSetLocal({ [PARSE_STORAGE_KEY]: migrated });
    return migrated;
  }

  return {};
}

async function getOrCreateParseJsonForChat(chatKey, sourceUrl) {
  const key = normalizeChatKey(chatKey, sourceUrl);
  const byChat = await getParseJsonData();
  const existing = byChat[key];
  if (existing && Array.isArray(existing.messages)) {
    return { key, data: existing, byChat };
  }

  const seeded =
    key === DEFAULT_CHAT_KEY ? await loadSeedParseJson() : { updatedAt: null, messages: [] };
  const nextByChat = {
    ...byChat,
    [key]: seeded
  };
  await storageSetLocal({ [PARSE_STORAGE_KEY]: nextByChat });
  return { key, data: seeded, byChat: nextByChat };
}

async function appendParseMessages(incomingMessages, chatKey, sourceUrl) {
  const incoming = Array.isArray(incomingMessages) ? incomingMessages : [];
  const scoped = await getOrCreateParseJsonForChat(chatKey, sourceUrl);
  const current = scoped.data;
  const existingMessages = Array.isArray(current.messages) ? current.messages : [];
  const seen = new Set(existingMessages.map(messageKey));
  const merged = existingMessages.slice();
  let added = 0;

  for (const message of incoming) {
    const text = normalizeText(message && message.text);
    if (!text) {
      continue;
    }
    const normalized = {
      text,
      messageDate: normalizeText(message && message.messageDate),
      time: normalizeText(message && message.time)
    };
    const key = messageKey(normalized);
    if (seen.has(key)) {
      continue;
    }
    seen.add(key);
    merged.push(normalized);
    added += 1;
  }

  const next = {
    updatedAt: new Date().toISOString(),
    messages: merged
  };
  const nextByChat = {
    ...scoped.byChat,
    [scoped.key]: next
  };
  await storageSetLocal({ [PARSE_STORAGE_KEY]: nextByChat });

  return {
    ok: true,
    chatKey: scoped.key,
    added,
    total: merged.length
  };
}

async function getParseJsonMessages(chatKey, sourceUrl) {
  const scoped = await getOrCreateParseJsonForChat(chatKey, sourceUrl);
  const data = scoped.data;
  return {
    ok: true,
    chatKey: scoped.key,
    updatedAt: data.updatedAt || null,
    messages: Array.isArray(data.messages) ? data.messages : []
  };
}

chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  (async () => {
    if (!message || !message.type) {
      sendResponse({
        ok: false,
        error: "Empty message."
      });
      return;
    }

    if (message.type === "DOWNLOAD_FILE") {
      const filename = message.filename || `max-chat-${Date.now()}.txt`;
      const mimeType = message.mimeType || "text/plain";
      const content = typeof message.content === "string" ? message.content : "";
      const url = toDataUrl(content, mimeType);

      chrome.downloads.download(
        {
          url,
          filename,
          saveAs: true
        },
        (downloadId) => {
          if (chrome.runtime.lastError) {
            sendResponse({
              ok: false,
              error: chrome.runtime.lastError.message
            });
            return;
          }

          sendResponse({
            ok: true,
            downloadId
          });
        }
      );
      return;
    }

    if (message.type === "APPEND_PARSE_MESSAGES") {
      const result = await appendParseMessages(message.messages, message.chatKey, message.sourceUrl);
      sendResponse(result);
      return;
    }

    if (message.type === "GET_PARSE_MESSAGES") {
      const result = await getParseJsonMessages(message.chatKey, message.sourceUrl);
      sendResponse(result);
      return;
    }

    sendResponse({
      ok: false,
      error: `Unknown type: ${message.type}`
    });
  })().catch((error) => {
    sendResponse({
      ok: false,
      error: error && error.message ? error.message : String(error)
    });
  });

  return true;
});
