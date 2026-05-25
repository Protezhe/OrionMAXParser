"use strict";

(() => {
  const STORAGE_KEY = "maxChatParserSettings";

  const DEFAULT_SELECTORS = {
    chatRootSelector: "main, [role='main'], [class*='chat']",
    messageSelector: ".block[role='listitem'], [class*='messageWrapper']",
    textSelector: ".bubble > .text, [data-bubbles-variant] .bubble > .text, [class*='bubble'] > [class*='text']",
    authorSelector: ".header .name .text, [class*='header'] [class*='name'] [class*='text'], [class*='name']",
    timeSelector: ".meta .text, [class*='meta'] [class*='text'], time, [data-testid*='time'], [class*='time'], [class*='timestamp']"
  };

  const LEGACY_SELECTORS = {
    messageSelector: "[data-testid*='message'], [class*='message'], [role='listitem']",
    textSelector: "[data-testid*='text'], [class*='text'], [class*='content'], p, span, div",
    authorSelector: "[data-testid*='author'], [class*='author'], [class*='sender'], [class*='name']",
    timeSelector: "time, [data-testid*='time'], [class*='time'], [class*='timestamp']"
  };

  const state = {
    settings: { ...DEFAULT_SELECTORS },
    messages: [],
    messageIds: new Set(),
    observer: null,
    observing: false,
    collectTimer: null,
    lastCollectAt: null,
    chatKey: ""
  };

  function normalizeText(value) {
    return (value || "").replace(/\s+/g, " ").trim();
  }

  function parseChatKeyFromUrl(url) {
    const raw = normalizeText(url);
    if (!raw) {
      return "__global__";
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

  function clearCollectedMessages() {
    if (state.collectTimer) {
      clearTimeout(state.collectTimer);
      state.collectTimer = null;
    }
    state.messages = [];
    state.messageIds = new Set();
    state.lastCollectAt = null;
  }

  function syncChatContext() {
    const nextChatKey = parseChatKeyFromUrl(location.href);
    if (nextChatKey === state.chatKey) {
      return false;
    }
    state.chatKey = nextChatKey;
    clearCollectedMessages();
    void syncWithStorage();
    return true;
  }

  function messageKey(message) {
    if (!message) {
      return "";
    }
    const text = normalizeText(message.text);
    const date = normalizeText(message.messageDate);
    const time = normalizeText(message.time);
    return `${date}|${time}|${text}`;
  }

  async function syncWithStorage() {
    try {
      const response = await chrome.runtime.sendMessage({
        type: "GET_PARSE_MESSAGES",
        chatKey: state.chatKey,
        sourceUrl: location.href
      });
      if (response && response.ok && Array.isArray(response.messages)) {
        for (const msg of response.messages) {
          const key = messageKey(msg);
          if (key && !state.messageIds.has(key)) {
            state.messageIds.add(key);
            // We don't necessarily need to add to state.messages if we only want to track 'new' ones in state.messages
            // But for consistency with how collectMessages works, let's keep state.messages as "messages found in this session that are NOT in storage"
          }
        }
      }
    } catch (_error) {
      // Background might not be ready or message type unknown
    }
  }

  function uniqueText(items) {
    const out = [];
    const seen = new Set();
    for (const item of items) {
      const value = normalizeText(item);
      if (!value || seen.has(value)) {
        continue;
      }
      seen.add(value);
      out.push(value);
    }
    return out;
  }

  function safeQueryOne(root, selector) {
    if (!selector) {
      return null;
    }
    try {
      return root.querySelector(selector);
    } catch (_error) {
      return null;
    }
  }

  function safeQueryAll(root, selector) {
    if (!selector) {
      return [];
    }
    try {
      return Array.from(root.querySelectorAll(selector));
    } catch (_error) {
      return [];
    }
  }

  function pickFirstText(root, selector) {
    const node = safeQueryOne(root, selector);
    if (!node) {
      return "";
    }
    return normalizeText(node.innerText || node.textContent || "");
  }

  function toIsoIfPossible(value) {
    const normalized = normalizeText(value);
    if (!normalized) {
      return "";
    }
    const parsed = Date.parse(normalized);
    if (Number.isNaN(parsed)) {
      return "";
    }
    return new Date(parsed).toISOString();
  }

  function pad2(value) {
    return String(value).padStart(2, "0");
  }

  function extractTimeOnly(value) {
    const normalized = normalizeText(value);
    if (!normalized) {
      return "";
    }
    const match = normalized.match(/\b([01]?\d|2[0-3]):([0-5]\d)\b/);
    if (!match) {
      return "";
    }
    return `${pad2(match[1])}:${match[2]}`;
  }

  function formatDateKey(date) {
    return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
  }

  function parseDateLabelToKey(label) {
    const normalized = normalizeText(label).toLowerCase();
    if (!normalized) {
      return "";
    }

    const now = new Date();
    if (normalized === "сегодня") {
      return formatDateKey(now);
    }
    if (normalized === "вчера") {
      const day = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
      return formatDateKey(day);
    }

    const ruMonths = {
      "января": 1,
      "февраля": 2,
      "марта": 3,
      "апреля": 4,
      "мая": 5,
      "июня": 6,
      "июля": 7,
      "августа": 8,
      "сентября": 9,
      "октября": 10,
      "ноября": 11,
      "декабря": 12
    };

    const words = normalized.match(/^(\d{1,2})\s+([а-яё]+)(?:\s+(\d{4}))?$/i);
    if (words) {
      const day = Number(words[1]);
      const month = ruMonths[words[2]];
      const year = words[3] ? Number(words[3]) : now.getFullYear();
      if (day >= 1 && day <= 31 && month && year >= 1970) {
        return `${year}-${pad2(month)}-${pad2(day)}`;
      }
    }

    const numeric = normalized.match(/^(\d{1,2})[./](\d{1,2})(?:[./](\d{2,4}))?$/);
    if (numeric) {
      const day = Number(numeric[1]);
      const month = Number(numeric[2]);
      const yearRaw = numeric[3];
      const year = yearRaw ? Number(yearRaw.length === 2 ? `20${yearRaw}` : yearRaw) : now.getFullYear();
      if (day >= 1 && day <= 31 && month >= 1 && month <= 12) {
        return `${year}-${pad2(month)}-${pad2(day)}`;
      }
    }

    return "";
  }

  function combineDateAndTimeToIso(dateKey, time) {
    if (!dateKey || !time) {
      return "";
    }
    const parsed = Date.parse(`${dateKey}T${time}:00`);
    if (Number.isNaN(parsed)) {
      return "";
    }
    return new Date(parsed).toISOString();
  }

  function extractMessageTimestamp(node, settings) {
    const timeNode = safeQueryOne(node, settings.timeSelector);
    if (!timeNode) {
      return {
        time: "",
        messageDate: "",
        messageDateIso: ""
      };
    }

    const rawTime = normalizeText(timeNode.innerText || timeNode.textContent || "");
    const time = extractTimeOnly(rawTime) || rawTime;
    const attrCandidates = [
      "datetime",
      "data-timestamp",
      "data-time",
      "data-date",
      "title",
      "aria-label"
    ];

    let rawDate = "";
    for (const attrName of attrCandidates) {
      const attrValue = timeNode.getAttribute(attrName);
      if (attrValue && normalizeText(attrValue)) {
        rawDate = normalizeText(attrValue);
        break;
      }
    }

    const messageDate = rawDate;
    const messageDateIso = toIsoIfPossible(messageDate);

    return {
      time,
      messageDate,
      messageDateIso
    };
  }

  function extractDateCapsule(node) {
    const capsule = safeQueryOne(node, ".capsule, [class*='capsule']");
    if (!capsule) {
      return "";
    }
    return normalizeText(capsule.innerText || capsule.textContent || "");
  }

  function resolveDateCapsules(root) {
    return safeQueryAll(root, ".capsule, [class*='capsule']")
      .map((node) => {
        const label = normalizeText(node.innerText || node.textContent || "");
        const key = parseDateLabelToKey(label);
        return key ? { node, key } : null;
      })
      .filter(Boolean);
  }

  function isCapsuleBeforeNode(capsuleNode, targetNode) {
    if (!capsuleNode || !targetNode) {
      return false;
    }
    if (capsuleNode === targetNode) {
      return true;
    }
    const relation = capsuleNode.compareDocumentPosition(targetNode);
    return Boolean(
      relation & Node.DOCUMENT_POSITION_FOLLOWING ||
        relation & Node.DOCUMENT_POSITION_CONTAINED_BY
    );
  }

  function findDateKeyForNode(node, capsules) {
    let lastKey = "";
    for (const capsule of capsules) {
      if (isCapsuleBeforeNode(capsule.node, node)) {
        lastKey = capsule.key;
        continue;
      }
      break;
    }
    return lastKey;
  }

  function resolveRoot(settings) {
    const candidates = uniqueText([
      settings.chatRootSelector,
      "main",
      "[role='main']",
      "body"
    ]);

    for (const selector of candidates) {
      const root = safeQueryOne(document, selector);
      if (root) {
        return root;
      }
    }

    return document.body;
  }

  function resolveMessageNodes(settings) {
    const root = resolveRoot(settings);
    const directNodes = safeQueryAll(
      root,
      settings.messageSelector || DEFAULT_SELECTORS.messageSelector
    );
    if (directNodes.length > 0) {
      return directNodes;
    }

    const fallbackSelectors = [
      ".block[role='listitem']",
      "[class*='messageWrapper']",
      "[class*='message']",
      "article"
    ];

    for (const selector of fallbackSelectors) {
      const nodes = safeQueryAll(root, selector);
      if (nodes.length >= 3) {
        return nodes;
      }
    }

    return [];
  }

  function extractMessageText(node, settings, author, time) {
    const parts = safeQueryAll(node, settings.textSelector)
      .map((el) => normalizeText(el.innerText || el.textContent || ""))
      .filter(Boolean);

    const filteredParts = uniqueText(parts).filter(
      (value) => value !== author && value !== time
    );

    if (filteredParts.length > 0) {
      return filteredParts.reduce((longest, current) =>
        current.length > longest.length ? current : longest
      );
    }

    let fallback = normalizeText(node.innerText || node.textContent || "");
    if (author && fallback.startsWith(author)) {
      fallback = normalizeText(fallback.slice(author.length));
    }
    if (time && fallback.endsWith(time)) {
      fallback = normalizeText(fallback.slice(0, -time.length));
    }
    return fallback;
  }

  function buildMessage(node, index, settings, dateFromCapsuleKey) {
    const author = pickFirstText(node, settings.authorSelector);
    const timestamp = extractMessageTimestamp(node, settings);
    const time = timestamp.time;
    const text = extractMessageText(node, settings, author, time);
    const contextDate = dateFromCapsuleKey;
    const messageDate = timestamp.messageDate || contextDate;

    if (!text || text.length < 2) {
      return null;
    }

    return {
      text,
      messageDate,
      time
    };
  }

  function mergeMessages(incoming) {
    let added = 0;
    for (const message of incoming) {
      const key = messageKey(message);
      if (!message || !key || state.messageIds.has(key)) {
        continue;
      }
      state.messageIds.add(key);
      state.messages.push(message);
      added += 1;
    }
    return added;
  }

  function collectMessages() {
    syncChatContext();
    const root = resolveRoot(state.settings);
    const nodes = resolveMessageNodes(state.settings);
    const capsules = resolveDateCapsules(root);
    const parsed = [];
    for (const [index, node] of nodes.entries()) {
      const capsuleText = extractDateCapsule(node);
      const capsuleDate = parseDateLabelToKey(capsuleText);
      if (capsuleDate) {
        continue;
      }
      const effectiveDateKey = findDateKeyForNode(node, capsules);
      const message = buildMessage(node, index, state.settings, effectiveDateKey);
      if (message) {
        parsed.push(message);
      }
    }
    const added = mergeMessages(parsed);
    state.lastCollectAt = new Date().toISOString();

    return {
      foundNodes: nodes.length,
      parsed: parsed.length,
      added,
      total: state.messages.length
    };
  }

  function scheduleCollect(delayMs = 700) {
    if (state.collectTimer) {
      clearTimeout(state.collectTimer);
    }
    state.collectTimer = setTimeout(() => {
      collectMessages();
      state.collectTimer = null;
    }, delayMs);
  }

  function startObserver() {
    if (state.observing) {
      return;
    }
    const root = resolveRoot(state.settings);
    state.observer = new MutationObserver(() => scheduleCollect(500));
    state.observer.observe(root, {
      childList: true,
      subtree: true,
      characterData: true
    });
    state.observing = true;
    collectMessages();
  }

  function stopObserver() {
    if (!state.observer) {
      state.observing = false;
      return;
    }
    state.observer.disconnect();
    state.observer = null;
    state.observing = false;
  }

  function clearMessages() {
    stopObserver();
    clearCollectedMessages();
  }

  function migrateLegacySettings(input) {
    const settings = {
      ...input
    };
    let changed = false;

    if (settings.messageSelector === LEGACY_SELECTORS.messageSelector) {
      settings.messageSelector = DEFAULT_SELECTORS.messageSelector;
      changed = true;
    }
    if (settings.textSelector === LEGACY_SELECTORS.textSelector) {
      settings.textSelector = DEFAULT_SELECTORS.textSelector;
      changed = true;
    }
    if (settings.authorSelector === LEGACY_SELECTORS.authorSelector) {
      settings.authorSelector = DEFAULT_SELECTORS.authorSelector;
      changed = true;
    }
    if (settings.timeSelector === LEGACY_SELECTORS.timeSelector) {
      settings.timeSelector = DEFAULT_SELECTORS.timeSelector;
      changed = true;
    }

    return {
      settings,
      changed
    };
  }

  async function loadSettings() {
    const data = await chrome.storage.sync.get(STORAGE_KEY);
    const saved = data[STORAGE_KEY] || {};
    const merged = {
      ...DEFAULT_SELECTORS,
      ...saved
    };
    const migrated = migrateLegacySettings(merged);
    state.settings = migrated.settings;
    if (migrated.changed) {
      await chrome.storage.sync.set({
        [STORAGE_KEY]: state.settings
      });
    }
  }

  async function saveSettings(nextSettings) {
    state.settings = {
      ...state.settings,
      ...nextSettings
    };
    await chrome.storage.sync.set({
      [STORAGE_KEY]: state.settings
    });
  }

  function getStatePayload() {
    syncChatContext();
    return {
      ok: true,
      url: location.href,
      chatKey: state.chatKey,
      total: state.messages.length,
      observing: state.observing,
      lastCollectAt: state.lastCollectAt,
      settings: state.settings
    };
  }

  chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
    (async () => {
      if (!message || !message.type) {
        sendResponse({ ok: false, error: "Empty message." });
        return;
      }

      if (message.type === "GET_STATE") {
        sendResponse(getStatePayload());
        return;
      }

      if (message.type === "GET_MESSAGES") {
        syncChatContext();
        sendResponse({
          ok: true,
          chatKey: state.chatKey,
          messages: state.messages
        });
        return;
      }

      if (message.type === "COLLECT_NOW") {
        const result = collectMessages();
        sendResponse({
          ...getStatePayload(),
          result
        });
        return;
      }

      if (message.type === "START_OBSERVER") {
        startObserver();
        sendResponse(getStatePayload());
        return;
      }

      if (message.type === "STOP_OBSERVER") {
        stopObserver();
        sendResponse(getStatePayload());
        return;
      }

      if (message.type === "CLEAR_MESSAGES") {
        clearMessages();
        sendResponse(getStatePayload());
        return;
      }

      if (message.type === "SET_SELECTORS") {
        const incoming = message.settings || {};
        await saveSettings(incoming);
        clearMessages();
        const result = collectMessages();
        sendResponse({
          ...getStatePayload(),
          result
        });
        return;
      }

      sendResponse({ ok: false, error: `Unknown type: ${message.type}` });
    })().catch((error) => {
      sendResponse({
        ok: false,
        error: error && error.message ? error.message : String(error)
      });
    });

    return true;
  });

  loadSettings()
    .then(async () => {
      syncChatContext();
      await syncWithStorage();
      collectMessages();
    })
    .catch(async () => {
      syncChatContext();
      await syncWithStorage();
      collectMessages();
    });
})();
