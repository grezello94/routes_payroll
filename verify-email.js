(function initVerifyEmailPage() {
  const submitBtn = document.getElementById("verifyEmailSubmitBtn");
  const messageEl = document.getElementById("verifyEmailMessage");
  const token = new URLSearchParams(window.location.search).get("token") || "";

  function setMessage(text, tone) {
    messageEl.classList.remove("status-info", "status-success", "status-error");
    if (tone === "success") messageEl.classList.add("status-success");
    else if (tone === "error") messageEl.classList.add("status-error");
    else messageEl.classList.add("status-info");
    messageEl.textContent = text || "";
  }

  async function request(path, body) {
    const response = await fetch(path, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(payload.error || "Request failed.");
    }
    return payload;
  }

  if (!token) {
    setMessage("Invalid verification link. Please request a new one from Settings.", "error");
    submitBtn.disabled = true;
    return;
  }

  async function verifyNow() {
    submitBtn.disabled = true;
    setMessage("Verifying email...", "info");
    try {
      const response = await request("/api/auth/verify-email-token", { token });
      setMessage(
        response.message || "Email verified successfully. You can go back to your system or log in to your system.",
        "success"
      );
    } catch (error) {
      setMessage(error.message, "error");
    } finally {
      submitBtn.disabled = false;
    }
  }

  submitBtn.addEventListener("click", verifyNow);
  verifyNow();
})();
