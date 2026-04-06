(function initResetPage() {
  const submitBtn = document.getElementById("resetSubmitBtn");
  const newPasswordEl = document.getElementById("resetNewPassword");
  const confirmEl = document.getElementById("resetConfirmPassword");
  const messageEl = document.getElementById("resetMessage");
  const token = new URLSearchParams(window.location.search).get("token") || "";

  function wirePasswordToggles() {
    const toggleButtons = Array.from(document.querySelectorAll(".password-toggle"));
    for (const button of toggleButtons) {
      button.addEventListener("click", () => {
        const targetId = button.getAttribute("data-password-target") || "";
        const input = document.getElementById(targetId);
        if (!input) return;

        const showing = input.type === "text";
        input.type = showing ? "password" : "text";
        button.setAttribute("aria-pressed", showing ? "false" : "true");
        button.setAttribute("aria-label", showing ? "Show password" : "Hide password");
        button.textContent = "👁";
      });
    }
  }

  function setMessage(text, tone) {
    messageEl.classList.remove("status-info", "status-success", "status-error");
    if (tone === "success") messageEl.classList.add("status-success");
    else if (tone === "error") messageEl.classList.add("status-error");
    else messageEl.classList.add("status-info");
    messageEl.textContent = text || "";
  }

  function isStrongPassword(value) {
    const raw = String(value || "");
    if (raw.length < 8 || raw.length > 80) return false;
    if (!/[a-z]/.test(raw)) return false;
    if (!/[A-Z]/.test(raw)) return false;
    if (!/[0-9]/.test(raw)) return false;
    return true;
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
    setMessage("Invalid reset link. Please request a new one.", "error");
    submitBtn.disabled = true;
    return;
  }

  wirePasswordToggles();

  submitBtn.addEventListener("click", async () => {
    const newPassword = newPasswordEl.value || "";
    const confirm = confirmEl.value || "";

    if (!isStrongPassword(newPassword)) {
      setMessage("Password must be 8+ chars and include uppercase, lowercase, and number.", "error");
      return;
    }
    if (newPassword !== confirm) {
      setMessage("New password and confirm password do not match.", "error");
      return;
    }

    submitBtn.disabled = true;
    setMessage("Updating password...", "info");

    try {
      await request("/api/auth/reset-password-with-token", {
        token,
        newPassword,
      });
      setMessage("Password updated successfully. You can now sign in.", "success");
      newPasswordEl.value = "";
      confirmEl.value = "";
    } catch (error) {
      setMessage(error.message, "error");
    } finally {
      submitBtn.disabled = false;
    }
  });
})();
