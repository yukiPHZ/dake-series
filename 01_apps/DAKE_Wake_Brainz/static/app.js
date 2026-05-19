const statusBadge = document.querySelector("#statusBadge");
const statusText = document.querySelector("#statusText");
const targetIp = document.querySelector("#targetIp");
const message = document.querySelector("#message");
const wakeButton = document.querySelector("#wakeButton");
const refreshButton = document.querySelector("#refreshButton");

function setBusy(isBusy) {
  wakeButton.disabled = isBusy;
  refreshButton.disabled = isBusy;
}

function setMessage(text, isError = false) {
  message.textContent = text;
  message.classList.toggle("error", isError);
}

function paintStatus(status) {
  const normalized = status === "ONLINE" ? "ONLINE" : "OFFLINE";
  statusText.textContent = normalized;
  statusBadge.textContent = normalized;
  statusText.classList.toggle("online", normalized === "ONLINE");
  statusText.classList.toggle("offline", normalized !== "ONLINE");
  statusBadge.classList.toggle("online", normalized === "ONLINE");
  statusBadge.classList.toggle("offline", normalized !== "ONLINE");
  statusBadge.classList.remove("checking");
}

function paintChecking() {
  statusBadge.textContent = "CHECKING";
  statusBadge.classList.add("checking");
}

async function refreshStatus({ silent = false } = {}) {
  if (!silent) {
    paintChecking();
    setMessage("Checking status...");
  }

  const response = await fetch("/api/status", { cache: "no-store" });
  const data = await response.json();
  paintStatus(data.status);
  targetIp.textContent = data.target_ip || "未設定";

  if (data.config_error) {
    setMessage(data.config_error, true);
  } else if (!data.config_exists) {
    setMessage("config.json 未設定。config.example.json を元に作成してください。", true);
  } else if (!data.status_ready) {
    setMessage("target_ip 未設定。状態確認には target_ip が必要です。", true);
  } else if (!data.wake_ready) {
    setMessage(`${data.checked_at} / ${data.status}。target_mac 未設定のため Wake は無効です。`);
  } else {
    setMessage(`${data.checked_at} / ${data.status}`);
  }
}

async function wakeTarget() {
  setBusy(true);
  setMessage("Sending magic packet...");
  try {
    const response = await fetch("/api/wake", { method: "POST" });
    const data = await response.json();
    if (!response.ok || !data.ok) {
      throw new Error(data.message || "Wake failed.");
    }
    setMessage(`${data.message} Refreshing status...`);
    setTimeout(() => refreshStatus({ silent: true }).catch(() => {}), 2200);
    setTimeout(() => refreshStatus({ silent: true }).catch(() => {}), 8000);
  } catch (error) {
    setMessage(error.message || String(error), true);
  } finally {
    setBusy(false);
  }
}

wakeButton.addEventListener("click", wakeTarget);
refreshButton.addEventListener("click", () => {
  setBusy(true);
  refreshStatus()
    .catch((error) => {
      paintStatus("OFFLINE");
      setMessage(error.message || String(error), true);
    })
    .finally(() => setBusy(false));
});

refreshStatus({ silent: true }).catch(() => {
  paintStatus("OFFLINE");
});
