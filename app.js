const form = document.querySelector("#generator-form");
const promptInput = document.querySelector("#prompt");
const charCount = document.querySelector("#char-count");
const button = document.querySelector("#generate-button");
const status = document.querySelector("#status");
const downloads = document.querySelector("#downloads");
const apiInput = document.querySelector("#api-url");
const saveApiButton = document.querySelector("#save-api");
const apiStorageKey = "fillerbunny.apiUrl";

function getApiBase() {
  const saved = localStorage.getItem(apiStorageKey);
  const defaultBase = window.location.hostname.endsWith(".github.io")
    ? ""
    : window.location.origin;
  return (saved || defaultBase).replace(/\/+$/, "");
}

function showStatus(message, kind = "") {
  status.hidden = !message;
  status.textContent = message;
  status.className = `status ${kind}`.trim();
}

function renderDownloads(result) {
  const base = getApiBase();
  const pptx = result.files.find((file) => file.name.endsWith(".pptx"));
  const mp4 = result.files.find((file) => file.name.endsWith(".mp4"));
  if (!pptx || !mp4) {
    throw new Error("The generator did not return both expected files.");
  }
  document.querySelector("#pptx-download").href = new URL(pptx.url, `${base}/`).href;
  document.querySelector("#mp4-download").href = new URL(mp4.url, `${base}/`).href;
  document.querySelector("#action-count").textContent =
    `${result.actions} ${result.actions === 1 ? "ACTION" : "ACTIONS"}`;
  downloads.hidden = false;
}

promptInput.addEventListener("input", () => {
  charCount.textContent = `${promptInput.value.length.toLocaleString()} / 2,000`;
});

apiInput.value = localStorage.getItem(apiStorageKey) || "";
saveApiButton.addEventListener("click", () => {
  const value = apiInput.value.trim();
  let origin = "";
  if (value) {
    try {
      const parsed = new URL(value);
      if (!["https:", "http:"].includes(parsed.protocol) ||
          parsed.pathname !== "/" || parsed.search || parsed.hash ||
          parsed.username || parsed.password) {
        throw new Error("invalid URL");
      }
      origin = parsed.origin;
    } catch {
      showStatus("Enter a backend origin such as https://your-service.onrender.com.", "error");
      return;
    }
  }
  if (origin && new URL(origin).protocol !== "https:" &&
      !["localhost", "127.0.0.1"].includes(new URL(origin).hostname)) {
    showStatus("Enter a backend URL such as https://your-service.onrender.com.", "error");
    return;
  }
  if (origin) {
    localStorage.setItem(apiStorageKey, origin);
  } else {
    localStorage.removeItem(apiStorageKey);
  }
  showStatus("Generator connection saved in this browser.");
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  downloads.hidden = true;
  const prompt = promptInput.value.trim();
  if (prompt.length < 5) {
    showStatus("Please enter a prompt with at least five characters.", "error");
    return;
  }
  if (!getApiBase()) {
    showStatus("Set the backend service URL above before generating files.", "error");
    return;
  }

  button.disabled = true;
  button.querySelector(".button-label").textContent = "Creating your files…";
  showStatus("Building your presentation and video preview. This may take a little while.", "loading");
  try {
    const response = await fetch(`${getApiBase()}/api/generate`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ prompt }),
    });
    const result = await response.json();
    if (!response.ok) {
      throw new Error(result.detail || `The generator returned HTTP ${response.status}.`);
    }
    renderDownloads(result);
    showStatus("Both files are ready. Download them below.");
  } catch (error) {
    const message = error instanceof TypeError
      ? "Could not reach the generator. Check the backend service URL and try again."
      : error.message;
    showStatus(message, "error");
  } finally {
    button.disabled = false;
    button.querySelector(".button-label").textContent = "Create my files";
  }
});
