"use strict";

(() => {
  const STORAGE_KEY = "maxChatParserSettings";

  const DEFAULT_SELECTORS = {
    chatRootSelector: "main, [role='main'], [class*='chat']",
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
    lastCollectAt: null
  };

  function normalizeText(value) {
    return (value || "").replace(/\s+/g, " ").trim();
  }

  function stableHash(input) {
    let hash = 0;
    for (let i = 0; i < input.length; i += 1) {
      hash = (hash << 5) - hash + input.charCodeAt(i);
      hash |= 0;
    }
    return `m_${Math.abs(hash)}`;
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
      "[role='listitem']",
      "[class*='message']",
      "article",
      "li"
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

  function buildMessage(node, index, settings) {
    const author = pickFirstText(node, settings.authorSelector);
    const time = pickFirstText(node, settings.timeSelector);
    const text = extractMessageText(node, settings, author, time);

    if (!text || text.length < 2) {
      return null;
    }

    const signature = `${index}|${author}|${time}|${text}`;
    return {
      id: stableHash(signature),
      index,
      author,
      time,
      text,
      sourceUrl: location.href,
      capturedAt: new Date().toISOString()
    };
  }

  function mergeMessages(incoming) {
    let added = 0;
    for (const message of incoming) {
      if (!message || state.messageIds.has(message.id)) {
        continue;
      }
      state.messageIds.add(message.id);
      state.messages.push(message);
      added += 1;
    }
    return added;
  }

  function collectMessages() {
    const nodes = resolveMessageNodes(state.settings);
    const parsed = nodes
      .map((node, index) => buildMessage(node, index, state.settings))
      .filter(Boolean);
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
    state.messages = [];
    state.messageIds = new Set();
    state.lastCollectAt = null;
  }

  async function loadSettings() {
    const data = await chrome.storage.sync.get(STORAGE_KEY);
    const saved = data[STORAGE_KEY] || {};
    state.settings = {
      ...DEFAULT_SELECTORS,
      ...saved
    };
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
    return {
      ok: true,
      url: location.href,
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
        sendResponse({
          ok: true,
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
    .then(() => {
      collectMessages();
    })
    .catch(() => {
      collectMessages();
    });
})();
