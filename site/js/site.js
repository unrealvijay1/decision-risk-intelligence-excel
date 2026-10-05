(() => {
  "use strict";
  const config = window.TELIVU_CONFIG || {};
  const script = document.currentScript;
  const base = new URL("../", script.src);
  const https = value => { try { const u = new URL(value); return u.protocol === "https:" && !u.username && !u.password ? u.href : ""; } catch { return ""; } };
  const download = config.downloadEnabled === true ? https(config.downloadUrl) : "";
  document.querySelectorAll("[data-download]").forEach(link => {
    link.href = download || new URL("index.html#download", base).href;
    link.textContent = download ? "Download Telivu Free" : "Download unavailable";
  });
  document.querySelectorAll("[data-download-status]").forEach(el => { el.textContent = download ? "Windows installer · " + (config.releaseStatus || "Public release") : "Download unavailable"; });
  document.querySelectorAll("[data-version]").forEach(el => { el.textContent = config.productVersion || ""; });
  document.querySelectorAll("[data-installer-size]").forEach(el => { if (download && config.installerSize) { el.hidden = false; el.textContent = "Installer size: " + config.installerSize; } });
  let contact = https(config.contactUrl);
  if (/^mailto:[^\s?]+@[^\s?]+$/.test(config.contactUrl || "")) contact = config.contactUrl;
  if (contact) {
    document.querySelectorAll("[data-contact]").forEach(el => { el.href = contact; el.textContent = "Contact Us"; });
    document.querySelectorAll(".contact-status").forEach(el => { el.textContent = "Contact us for product and business enquiries."; });
  }
  const demo = https(config.demoUrl);
  if (demo) document.querySelectorAll("[data-demo]").forEach(el => { el.href = demo; });
  document.querySelectorAll("[data-release-notes]").forEach(el => {
    const external = https(config.releaseNotesUrl);
    el.href = external || new URL("resources/release-notes/", base).href;
  });
  const toggle = document.querySelector(".menu-toggle");
  const navigation = document.getElementById("navigation");
  const mobile = window.matchMedia("(max-width: 800px)");
  const setOpen = open => { toggle.setAttribute("aria-expanded", String(open)); navigation.toggleAttribute("data-mobile-hidden", mobile.matches && !open); };
  const sync = () => { toggle.hidden = !mobile.matches; setOpen(false); };
  sync(); mobile.addEventListener("change", sync);
  toggle.addEventListener("click", () => setOpen(toggle.getAttribute("aria-expanded") !== "true"));
  navigation.addEventListener("click", event => { if (event.target.closest("a") && mobile.matches) setOpen(false); });
  document.addEventListener("keydown", event => { if (event.key === "Escape" && mobile.matches && toggle.getAttribute("aria-expanded") === "true") { setOpen(false); toggle.focus(); } });
  // Safe static defaults and content remain usable without JS. No tracker is loaded.
})();
