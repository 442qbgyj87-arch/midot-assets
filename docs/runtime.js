/* MIDOT Auto Signature - Outlook event-based add-in
 * Pilot: only explicitly listed users are modified.
 * Host this file over HTTPS at the same origin as runtime.html.
 */
(function () {
  "use strict";

  // ALLOWLIST: no signature is inserted for anyone not listed here.
  // Template values: "ambos", "reporte", "mindsafe".
  // For each new user, add an entry with their email and desired template.
  var USERS = {
    "adrian.torres@midot.com": { template: "ambos", name: "Adrian Torres", title: "" }
  };

  var BANNERS = {
    reporte: {
      image: "https://raw.githubusercontent.com/442qbgyj87-arch/midot-assets/main/firma%20reporte%20midot%202026.png",
      url: "https://midot.com/latam/landpage/descarga-2/",
      alt: "Reporte MIDOT 2026"
    },
    mindsafe: {
      image: "https://raw.githubusercontent.com/442qbgyj87-arch/midot-assets/main/Perfiles%20y%20comportamiento%20humano_%20seguridad%20industrial.png",
      url: "https://mindsafemidot.com/landpage/descarga-reporte-1/",
      alt: "MindSafe - Seguridad industrial"
    }
  };

  function escapeHtml(value) {
    return String(value == null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function renderBanner(banner) {
    return '<table role="presentation" border="0" cellpadding="0" cellspacing="0" width="520">' +
      '<tr><td><a href="' + escapeHtml(banner.url) + '">' +
      '<img src="' + escapeHtml(banner.image) + '" alt="' + escapeHtml(banner.alt) +
      '" width="520" border="0"></a></td></tr></table>';
  }

  function renderSignature(email, user) {
    var profile = Office.context.mailbox.userProfile;
    var name = user.name || (profile && profile.displayName) || email;
    var html = '<table role="presentation" border="0" cellpadding="0" cellspacing="0">' +
      '<tr><td><font face="Arial" color="#242424" size="3"><b>' + escapeHtml(name) +
      '</b></font></td></tr>' +
      (user.title ? '<tr><td><font face="Arial" size="2">' + escapeHtml(user.title) + '</font></td></tr>' : '') +
      '<tr><td><font face="Arial" size="2">MIDOT</font></td></tr>' +
      '<tr><td><font face="Arial" size="2"><a href="mailto:' + escapeHtml(email) + '">' +
      escapeHtml(email) + '</a></font></td></tr></table>';

    if (user.template === "ambos" || user.template === "reporte") {
      html += renderBanner(BANNERS.reporte);
    }
    if (user.template === "ambos" || user.template === "mindsafe") {
      html += renderBanner(BANNERS.mindsafe);
    }
    return html;
  }

  function setMidotSignature(event) {
    var completed = false;
    function finish() {
      if (!completed) {
        completed = true;
        event.completed();
      }
    }
    try {
      if (!Office.context || !Office.context.mailbox || !Office.context.mailbox.item) {
        finish();
        return;
      }
      if (!Office.context.requirements.isSetSupported("Mailbox", "1.10")) {
        finish();
        return;
      }
      var profile = Office.context.mailbox.userProfile;
      var email = (profile && profile.emailAddress ? profile.emailAddress : "").trim().toLowerCase();
      var user = USERS[email];
      if (!user || ["ambos", "reporte", "mindsafe"].indexOf(user.template) < 0) {
        // Important: no changes for users not assigned to pilot/rollout.
        finish();
        return;
      }
      var item = Office.context.mailbox.item;
      var html = renderSignature(email, user);
      if (html.length > 30000) {
        console.error("MIDOT signature exceeds Outlook's 30,000-character limit.");
        finish();
        return;
      }
      // Avoid changing a text-only message: setSignatureAsync(HTML) would fail.
      item.body.getTypeAsync(function (typeResult) {
        if (typeResult.status !== Office.AsyncResultStatus.Succeeded ||
            typeResult.value !== Office.CoercionType.Html) {
          console.warn("MIDOT signature not inserted: the message body is not HTML.");
          finish();
          return;
        }
        try {
          // This API adds a signature or replaces a recognized existing signature.
          // It does NOT change the tenant roaming-signature configuration.
          item.body.setSignatureAsync(
            html,
            { coercionType: Office.CoercionType.Html },
            function (result) {
              if (result.status !== Office.AsyncResultStatus.Succeeded) {
                console.error("MIDOT setSignatureAsync error:", result.error && result.error.message);
              }
              finish();
            }
          );
        } catch (err) {
          console.error("MIDOT signature operation failed:", err);
          finish();
        }
      });
    } catch (err) {
      console.error("MIDOT event handler failed:", err);
      finish();
    }
  }

  // Required by Outlook event-based activation; do not wrap in Office.onReady().
  Office.actions.associate("setMidotSignature", setMidotSignature);
})();
