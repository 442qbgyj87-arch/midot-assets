/* MIDOT Auto Signature - Outlook event-based add-in
 * Pilot: only explicitly listed users are modified.
 * Host this file over HTTPS at the same origin as runtime.html.
 */
(function () {
  "use strict";

  Office.onReady();
  console.info("[MIDOT] Runtime cargado");

  // ALLOWLIST: no signature is inserted for anyone not listed here.
  // Template values: "ambos", "reporte", "mindsafe".
  // For each new user, add an entry with their email and desired template.
  var USERS = {
  // Fase siguiente: estos perfiles se configuran aqui, pero requieren asignacion en M365.
  // template: null = inactivo hasta verificar identidad/correo corporativo.
  "leticia.cabrera@midot.com.mx": {
    template: null,
    name: "Leticia Cabrera",
    title: "KEY ACCOUNT MANAGER",
    company: "MIDOT",
    region: "MÉXICO",
    phone: "3314212737"
  },

  "susana.lechuga@midot.com.mx": {
    template: null,
    name: "Susana Lechuga",
    title: "KEY ACCOUNT MANAGER",
    company: "MIDOT",
    region: "MÉXICO",
    phone: "8112665578"
  },

  "alejandra.santamaria@midot.com.mx": {
    template: "reporte",
    name: "Alejandra Santamaría",
    title: "KEY ACCOUNT MANAGER",
    company: "MIDOT",
    region: "MÉXICO",
    phone: "5524352496"
  },

  "victor.lopez@midot.com.mx": {
    template: "reporte",
    name: "Víctor López",
    title: "ACCOUNT MANAGER",
    company: "MIDOT",
    region: "MÉXICO",
    phone: "5538996467"
  },

  "jessica.gonzalez@midot.com.mx": {
    template: "reporte",
    name: "Jessica González",
    title: "ACCOUNT MANAGER PYME",
    company: "MIDOT",
    region: "MÉXICO",
    phone: "8122150360"
  },

  "norma.escudero@midot.com.mx": {
    template: "reporte",
    name: "Norma Escudero",
    title: "ACCOUNT MANAGER PYME",
    company: "MIDOT",
    region: "MÉXICO",
    phone: "4422647042"
  },

  "paula.mesa@midot.com": {
    template: "reporte",
    name: "Paula Mesa",
    title: "Key Account Manager",
    company: "MIDOT",
    region: "COLOMBIA",
    phone: "57 3183410684"
  },

  "juliana.gomez@midot.com": {
    template: "reporte",
    name: "Juliana Gómez",
    title: "Key Account Manager Latam",
    company: "MIDOT",
    region: "LATAM",
    phone: "57 3123923924"
  },

  "jheymi@midot.com": {
    template: null,
    name: "Jheymi Fernández",
    title: "KEY ACCOUNT MANAGER MINDSAFE",
    company: "MIDOT-MindSafe",
    region: "PERÚ",
    phone: "51 908853185"
  },

  "natasha.antequera@midot.com": {
    template: "ambos",
    name: "Natasha Antequera",
    title: "CUSTOMER ENABLEMENT LEAD",
    company: "MIDOT-MindSafe",
    region: "LATAM",
    phone: "569 86763518"
  },

  "jesus.nava@midot.com.mx": {
    template: "ambos",
    name: "Jesús Nava",
    title: "HEAD OF FINANCE",
    company: "MIDOT-MindSafe",
    region: "LATAM",
    phone: "5513754419"
  },

  "fernanda.rodriguez@midot.com.mx": {
    template: "ambos",
    name: "Fernanda Rodríguez",
    title: "CUSTOMER SUCCESS MANAGER",
    company: "MIDOT-MindSafe",
    region: "LATAM",
    phone: "5523880742"
  },

  "maria.reynoso@midot.com.mx": {
    template: "ambos",
    name: "María Reynoso",
    title: "TRAINING SPECIALIST",
    company: "MIDOT-MindSafe",
    region: "LATAM",
    phone: "5518939379"
  },

  "andrea.montano@midot.com.mx": {
    template: "ambos",
    name: "Andrea Montaño",
    title: "TELEMEETING REPRESENTATIVE",
    company: "MIDOT-MindSafe",
    region: "MÉXICO",
    phone: "5554192299"
  },

  "alfonso.valdivia@midot.com": {
    template: "ambos",
    name: "Alfonso Valdivia",
    title: "PROJECT MANAGER",
    company: "MIDOT-MindSafe",
    region: "LATAM",
    phone: "51 984185404"
  },

  "cuentasporpagar@midot.com.mx": {
    template: null,
    name: "Sofía de la Rosa",
    title: "BECARIA FINANZAS",
    company: "MIDOT-MindSafe",
    region: "MÉXICO"
  },

  "marco.miranda@midot.com.mx": {
    template: "mindsafe",
    name: "Marco Miranda",
    title: "MARKETING SPECIALIST MINDSAFE",
    company: "MindSafe",
    region: "LATAM",
    phone: "5543451804"
  },

  "irais.ortiz@midot.com.mx": {
    template: "mindsafe",
    name: "Irais Ortiz",
    title: "INDUSTRIAL SAFETY SR CONSULTANT",
    company: "MindSafe",
    region: "MÉXICO",
    phone: "5619938508"
  },

  "fabian.agudelo@midot.com": {
    template: "mindsafe",
    name: "Fabián Agudelo",
    title: "Key Account Manager Mindsafe",
    company: "MindSafe",
    region: "COLOMBIA",
    phone: "57 3102090320"
  },

  "jorge.bustillos@midot.com": {
    template: "mindsafe",
    name: "Jorge Bustillos",
    title: "INDUSTRIAL SAFETY SR CONSULTANT",
    company: "MindSafe",
    region: "LATAM",
    phone: "51 959894575"
  },

  "cesar.perez@midot.com.mx": {
    template: "mindsafe",
    name: "César Pérez",
    title: "INDUSTRIAL SAFETY SR CONSULTANT",
    company: "MindSafe",
    region: "MÉXICO"
  },

  "javier.ortiz@midot.com": {
    template: "mindsafe",
    name: "Javier Ortiz",
    title: "HEAD OF SAFETY CONSULTING",
    company: "MindSafe",
    region: "LATAM"
  },

   "ariana.sanchez@midot.com.mx": {
    template: "reporte",
    name: "Ariana Sánchez",
    title: "KEY ACCOUNT MANAGER",
    company: "MIDOT",
    region: "MÉXICO",
    phone: "5541902351"
  },

   "pricila.alcala@midot.com.mx": {
    template: "reporte",
    name: "Pricila Alcalá",
    title: "KEY ACCOUNT MANAGER",
    company: "MIDOT",
    region: "MÉXICO",
    phone: "5560527645"
  },

   "zully.sanchez@midot.com.mx": {
    template: "reporte",
    name: "Zully Sánchez",
    title: "KEY ACCOUNT MANAGER",
    company: "MIDOT",
    region: "MÉXICO",
    phone: "9991753267"
  },

   "adrian.torres@midot.com": {
    template: "ambos",
    name: "Adrian Torres",
    title: "IT LATAM",
    company: "MIDOT-MindSafe",
    region: "LATAM",
    phone: "5554315428"
  },

    
   "jimena.sandoval@midot.com.mx": {
    template: "ambos",
    name: "Jimena Ramírez Sandoval",
    title: "SUPPORT SPECIALIST",
    company: "MIDOT-MindSafe",
    region: "LATAM",
    phone: "5539503816"
  },


    
   "catalina.pemberty@midot.com": {
    template: "ambos",
    name: "Catalina Pemberty",
    title: "Support Specialist",
    company: "MIDOT",
    region: "COLOMBIA",
    phone: "57 3183434714"
  },

   // Eliana: BU sin definir; no insertar firma hasta confirmar plantilla.
   "eliana.velez@midot.com": {
    template: "ambos",
    name: "Eliana Vélez",
    title: "Administrative Assistant",
    company: "MIDOT",
    region: "COLOMBIA",
    phone: "57 3103776729"
  },

   "jackelyne.leon@midot.com": {
    template: "mindsafe",
    name: "Jackelyne León Palomino",
    title: "SALES MANAGER MINDSAFE",
    company: "MindSafe",
    region: "PERÚ",
    phone: "51 959895539"
  },

 
   "marelynduarte@midot.com.mx": {
    template: "ambos",
    name: "Marelyn González",
    title: "TRAINING SPECIALIST",
    company: "MIDOT-MindSafe",
    region: "LATAM",
    phone: "5579337935"
  },

   "cristina.carreon@midot.com.mx": {
    template: "ambos",
    name: "Cristina Carreón",
    title: "TRAINING SPECIALIST",
    company: "MIDOT-MindSafe",
    region: "LATAM",
    phone: "5545751823"
  },

   "gabriel.hernandez@midot.com.mx": {
    template: "reporte",
    name: "Gabriel Hernandez",
    title: "ACCOUNT MANAGER",
    company: "MIDOT",
    region: "MÉXICO",
    phone: "4421425417"
  },

   "carlos.giraldo@midot.com": {
    template: "reporte",
    name: "Carlos Giraldo",
    title: "SALES MANAGER ENTERPRISE",
    company: "MIDOT",
    region: "PERÚ",
    phone: "51 993832159"
  },
  
   "malory.zuluaga@midot.com": {
    template: "reporte",
    name: "Malory Zuluaga",
    title: "KEY ACCOUNT MANAGER LATAM",
    company: "MIDOT",
    region: "LATAM",
    phone: "+57 3182623959"
  },

  "giorgio.dacol@midot.com": {
    template: "reporte",
    name: "Giorgio Da Col",
    title: "MARKETING SPECIALIST ENTERPRISE",
    company: "MIDOT",
    region: "LATAM",
    phone: "5512566353"
  },

  "paola.uribe@midot.com": {
    template: "reporte",
    name: "Paola Uribe Gómez",
    title: "HEAD OF SALES ENTERPRISE",
    phone: "(+52) 5580223418",
    web: "www.midot.com",
    contactStyle: "english"
  },

  "lorena.mercado@midot.com.mx": {
    template: "mindsafe",
    name: "Lorena Mercado",
    title: "HEAD OF SALES MINDSAFE LATAM",
    company: "MindSafe",
    region: "LATAM",
    phone: "5515255027"
  }
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
  var name = user.name ||
    (profile && profile.displayName) || email;

  var html =
    '<table role="presentation" border="0" cellpadding="0" cellspacing="0">' +

    '<tr><td><font face="Arial" color="#242424" size="3"><b>' +
    escapeHtml(name) +
    '</b></font></td></tr>' +

    (user.title ?
      '<tr><td><font face="Arial" size="2">' +
      escapeHtml(user.title) +
      '</font></td></tr>' : '') +

    (user.company ?
      '<tr><td><font face="Arial" size="2">' +
      escapeHtml(user.company) +
      '</font></td></tr>' : '') +

    (user.region ?
      '<tr><td><font face="Arial" size="2">' +
      escapeHtml(user.region) +
      '</font></td></tr>' : '') +

    
(user.phone ?
  '<tr><td><font face="Arial" size="2">' +
  (user.contactStyle === "english" ? "Mobile: " : "Cel.: ") +
  escapeHtml(user.phone) +
  '</font></td></tr>' : '') +


   
'<tr><td><font face="Arial" size="2">' +
(user.contactStyle === "english" ? "E-mail: " : "") +
'<a href="mailto:' + escapeHtml(email) + '">' +
escapeHtml(email) +
'</a></font></td></tr>' +

(user.web ?
  '<tr><td><font face="Arial" size="2">Web: ' +
  '<a href="https://' + escapeHtml(user.web) + '">' +
  escapeHtml(user.web) +
  '</a></font></td></tr>' : '') +

'</table>';


  if (user.template === "ambos" ||
      user.template === "reporte") {
    html += renderBanner(BANNERS.reporte);
  }

  if (user.template === "ambos" ||
      user.template === "mindsafe") {
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
