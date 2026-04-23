"use strict";

function toDataUrl(content, mimeType) {
  const encoded = encodeURIComponent(content);
  return `data:${mimeType};charset=utf-8,${encoded}`;
}

chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  if (!message || message.type !== "DOWNLOAD_FILE") {
    return;
  }

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

  return true;
});
