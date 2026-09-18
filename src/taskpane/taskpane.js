/* global Office */

(function () {
  "use strict";

  const XAI_API = "https://api.x.ai/v1/chat/completions";
  const MODEL = "grok-3"; // or grok-2, grok-4 depending on availability

  let lastAssistantReply = "";
  let isSending = false;

  // DOM
  const messagesEl = document.getElementById("messages");
  const inputEl = document.getElementById("user-input");
  const sendBtn = document.getElementById("send-btn");
  const settingsBtn = document.getElementById("settings-btn");
  const settingsPanel = document.getElementById("settings-panel");
  const apiKeyInput = document.getElementById("api-key");
  const saveKeyBtn = document.getElementById("save-key");
  const clearKeyBtn = document.getElementById("clear-key");
  const autoContextCb = document.getElementById("auto-context");

  // Init
  Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
      loadSettings();
      bindEvents();
      addSystemMessage("Grok is ready. Select text or type a message.");
    } else {
      addSystemMessage("This add-in only works in Microsoft Word.");
    }
  });

  function loadSettings() {
    const key = localStorage.getItem("xai_api_key") || "";
    apiKeyInput.value = key;
    autoContextCb.checked = localStorage.getItem("auto_context") !== "false";
  }

  function bindEvents() {
    settingsBtn.addEventListener("click", () => {
      settingsPanel.classList.toggle("hidden");
    });

    saveKeyBtn.addEventListener("click", () => {
      const key = apiKeyInput.value.trim();
      if (key) {
        localStorage.setItem("xai_api_key", key);
        addSystemMessage("API key saved.");
        settingsPanel.classList.add("hidden");
      }
    });

    clearKeyBtn.addEventListener("click", () => {
      localStorage.removeItem("xai_api_key");
      apiKeyInput.value = "";
      addSystemMessage("API key cleared.");
    });

    autoContextCb.addEventListener("change", () => {
      localStorage.setItem("auto_context", autoContextCb.checked ? "true" : "false");
    });

    sendBtn.addEventListener("click", sendMessage);
    inputEl.addEventListener("keydown", (e) => {
      if (e.key === "Enter" && !e.shiftKey) {
        e.preventDefault();
        sendMessage();
      }
    });

    document.getElementById("btn-selection").addEventListener("click", () => {
      getSelectionText().then((text) => {
        if (text) {
          inputEl.value = `Regarding this text:\n\n"${text}"\n\n`;
          inputEl.focus();
        } else {
          addSystemMessage("No text selected.");
        }
      });
    });

    document.getElementById("btn-full-doc").addEventListener("click", () => {
      getFullDocumentText().then((text) => {
        if (text) {
          const preview = text.length > 800 ? text.slice(0, 800) + "..." : text;
          inputEl.value = `Here is the current document:\n\n"${preview}"\n\n`;
          inputEl.focus();
        } else {
          addSystemMessage("Could not read document.");
        }
      });
    });

    document.getElementById("btn-rewrite").addEventListener("click", () => {
      getSelectionText().then((text) => {
        if (text) {
          inputEl.value = `Please rewrite and improve the following text while keeping the original meaning:\n\n${text}`;
          sendMessage();
        } else {
          addSystemMessage("Select some text first to rewrite.");
        }
      });
    });

    document.getElementById("btn-insert").addEventListener("click", () => {
      if (lastAssistantReply) {
        insertText(lastAssistantReply);
      } else {
        addSystemMessage("No reply to insert yet.");
      }
    });
  }

  async function sendMessage() {
    const text = inputEl.value.trim();
    if (!text || isSending) return;

    const apiKey = localStorage.getItem("xai_api_key");
    if (!apiKey) {
      addErrorMessage("Please set your xAI API key in settings first.");
      settingsPanel.classList.remove("hidden");
      return;
    }

    isSending = true;
    sendBtn.disabled = true;
    inputEl.value = "";
    addUserMessage(text);

    // Optional auto context
    let context = "";
    if (autoContextCb.checked) {
      try {
        const sel = await getSelectionText();
        if (sel && sel.length > 10) {
          context = `\n\n[Current selection in document]:\n${sel}`;
        }
      } catch (e) {}
    }

    const messages = [
      {
        role: "system",
        content: "You are Grok, a helpful AI assistant built by xAI, integrated into Microsoft Word. Be concise, practical, and writing-focused. When asked to rewrite or improve text, return only the improved version unless the user asks for explanation."
      },
      {
        role: "user",
        content: text + context
      }
    ];

    try {
      const response = await fetch(XAI_API, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Authorization": `Bearer ${apiKey}`
        },
        body: JSON.stringify({
          model: MODEL,
          messages: messages,
          temperature: 0.7,
          stream: false
        })
      });

      if (!response.ok) {
        const err = await response.text();
        throw new Error(`API error ${response.status}: ${err}`);
      }

      const data = await response.json();
      const reply = data.choices?.[0]?.message?.content || "No response.";
      lastAssistantReply = reply;
      addAssistantMessage(reply);
    } catch (err) {
      addErrorMessage(err.message || "Failed to reach Grok.");
    } finally {
      isSending = false;
      sendBtn.disabled = false;
      inputEl.focus();
    }
  }

  // Office.js helpers
  function getSelectionText() {
    return new Promise((resolve) => {
      Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();
        resolve(selection.text || "");
      }).catch(() => resolve(""));
    });
  }

  function getFullDocumentText() {
    return new Promise((resolve) => {
      Word.run(async (context) => {
        const body = context.document.body;
        body.load("text");
        await context.sync();
        resolve(body.text || "");
      }).catch(() => resolve(""));
    });
  }

  function insertText(text) {
    Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertText(text, Word.InsertLocation.replace);
      await context.sync();
      addSystemMessage("Inserted into document.");
    }).catch((err) => {
      addErrorMessage("Insert failed: " + err.message);
    });
  }

  // UI helpers
  function addUserMessage(text) {
    const div = document.createElement("div");
    div.className = "message user";
    div.textContent = text;
    messagesEl.appendChild(div);
    scrollToBottom();
  }

  function addAssistantMessage(text) {
    const div = document.createElement("div");
    div.className = "message assistant";
    div.textContent = text;
    messagesEl.appendChild(div);
    scrollToBottom();
  }

  function addSystemMessage(text) {
    const div = document.createElement("div");
    div.className = "message system";
    div.textContent = text;
    messagesEl.appendChild(div);
    scrollToBottom();
  }

  function addErrorMessage(text) {
    const div = document.createElement("div");
    div.className = "message error";
    div.textContent = text;
    messagesEl.appendChild(div);
    scrollToBottom();
  }

  function scrollToBottom() {
    const container = document.getElementById("chat-container");
    container.scrollTop = container.scrollHeight;
  }
})();
