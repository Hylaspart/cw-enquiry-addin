/* CW Enquiry Assistant - Outlook task pane logic */
(function () {
  "use strict";

  var CFG = window.CW_CONFIG || {};
  var KEY_SETTING = "cwAccessKey";
  var LOGGED_PROP = "cwHubspotLogged";

  var state = {
    item: null,
    email: null,      // { subject, fromName, fromEmail, receivedAt, body }
    analysis: null
  };

  var FIELDS = ["first_name", "last_name", "email", "phone", "nationality",
    "current_location", "matter_type", "key_dates", "summary"];

  function $(id) { return document.getElementById(id); }
  function show(el) { el.classList.remove("hidden"); }
  function hide(el) { el.classList.add("hidden"); }

  /* ---------- status line ---------- */
  function setStatus(msg, kind) {
    var el = $("status");
    if (!msg) { hide(el); el.textContent = ""; return; }
    el.className = "status" + (kind ? " " + kind : "");
    el.textContent = msg;
  }

  /* ---------- settings (roaming = stored in the user's own mailbox) ---------- */
  function getKey() {
    try { return Office.context.roamingSettings.get(KEY_SETTING) || ""; } catch (e) { return ""; }
  }
  function saveKey(value, done) {
    Office.context.roamingSettings.set(KEY_SETTING, value);
    Office.context.roamingSettings.saveAsync(function (r) {
      done(r.status === Office.AsyncResultStatus.Succeeded);
    });
  }
  function openSettings() {
    $("accessKey").value = getKey();
    show($("settingsView")); hide($("mainView"));
  }
  function closeSettings() {
    hide($("settingsView")); show($("mainView"));
  }

  /* ---------- HTML sanitising for AI-generated reply ---------- */
  var ALLOWED_TAGS = { P: 1, BR: 1, STRONG: 1, B: 1, EM: 1, I: 1, U: 1, UL: 1, OL: 1, LI: 1, A: 1, SPAN: 1, DIV: 1 };
  function sanitizeHtml(html) {
    var doc = new DOMParser().parseFromString("<div>" + (html || "") + "</div>", "text/html");
    var root = doc.body.firstChild;
    (function walk(node) {
      Array.prototype.slice.call(node.childNodes).forEach(function (child) {
        if (child.nodeType === 1) {
          if (!ALLOWED_TAGS[child.tagName]) {
            if (/^(SCRIPT|STYLE|IFRAME|OBJECT|EMBED|LINK|META|FORM|INPUT|BUTTON|TEXTAREA|SELECT|IMG|SVG|MATH)$/.test(child.tagName)) {
              child.remove(); return;
            }
            // unwrap unknown but harmless tags
            while (child.firstChild) node.insertBefore(child.firstChild, child);
            child.remove(); return;
          }
          Array.prototype.slice.call(child.attributes).forEach(function (a) {
            var name = a.name.toLowerCase();
            var keep = child.tagName === "A" && name === "href" && /^(https?:|mailto:|tel:)/i.test(a.value.trim());
            if (!keep) child.removeAttribute(a.name);
          });
          walk(child);
        } else if (child.nodeType !== 3) {
          child.remove();
        }
      });
    })(root);
    return root.innerHTML;
  }

  /* ---------- network ---------- */
  // Sent as text/plain with no custom headers, so the browser does not need a CORS preflight.
  function postJSON(url, payload) {
    var ctrl = new AbortController();
    var timer = setTimeout(function () { ctrl.abort(); }, 120000);
    return fetch(url, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=UTF-8" },
      body: JSON.stringify(payload),
      signal: ctrl.signal
    }).then(function (res) {
      clearTimeout(timer);
      return res.text().then(function (text) {
        var data = null;
        try { data = JSON.parse(text); } catch (e) { /* not JSON */ }
        if (!res.ok || !data || data.ok === false) {
          var msg = (data && (data.error || data.message)) || ("Server returned " + res.status);
          if (msg === "Error in workflow") {
            msg = "The server could not process this request. Check the access key in settings (gear icon). " +
              "If it is correct, ask the n8n admin to check the failed execution.";
          }
          throw new Error(msg);
        }
        return data;
      });
    }, function (err) {
      clearTimeout(timer);
      if (err && err.name === "AbortError") throw new Error("The request timed out. Try again.");
      throw new Error("Could not reach the server. Check your connection, or ask IT whether n8n is running.");
    });
  }

  /* ---------- reading the open email ---------- */
  function loadItem() {
    var item = Office.context.mailbox.item;
    state.item = item;
    state.email = null;
    state.analysis = null;
    hide($("results")); hide($("logged")); setStatus("");
    $("analyseBtn").disabled = true;

    if (!item) {
      $("emailSubject").textContent = "Open an email to start.";
      $("emailFrom").textContent = "";
      return;
    }

    var from = item.from || item.sender || {};
    $("emailSubject").textContent = item.subject || "(no subject)";
    $("emailFrom").textContent = (from.displayName || "") + (from.emailAddress ? " <" + from.emailAddress + ">" : "");

    item.body.getAsync(Office.CoercionType.Text, function (r) {
      if (r.status !== Office.AsyncResultStatus.Succeeded) {
        setStatus("Could not read the email body: " + (r.error && r.error.message), "error");
        return;
      }
      var body = r.value || "";
      if (body.length > (CFG.MAX_BODY_CHARS || 20000)) body = body.slice(0, CFG.MAX_BODY_CHARS || 20000);
      state.email = {
        subject: item.subject || "",
        fromName: from.displayName || "",
        fromEmail: from.emailAddress || "",
        receivedAt: item.dateTimeCreated ? new Date(item.dateTimeCreated).toISOString() : "",
        body: body
      };
      $("analyseBtn").disabled = false;
    });

    // Show if someone already logged this email to HubSpot
    if (item.loadCustomPropertiesAsync) {
      item.loadCustomPropertiesAsync(function (r) {
        if (r.status !== Office.AsyncResultStatus.Succeeded) return;
        var info = r.value.get(LOGGED_PROP);
        if (info) {
          $("logged").textContent = "Already logged to HubSpot" + (info.by ? " by " + info.by : "") +
            (info.at ? " on " + new Date(info.at).toLocaleString("en-GB") : "") + ".";
          show($("logged"));
        }
      });
    }
  }

  function staff() {
    var p = Office.context.mailbox.userProfile || {};
    return { name: p.displayName || "", email: p.emailAddress || "" };
  }

  /* ---------- analyse ---------- */
  function analyse() {
    var key = getKey();
    if (!key) { openSettings(); return; }
    if (!state.email) return;

    $("analyseBtn").disabled = true; $("regenBtn").disabled = true;
    setStatus("Reading the enquiry and drafting a reply. This usually takes 10 to 30 seconds.", "busy");

    postJSON(CFG.ANALYSE_URL, {
      key: key,
      email: state.email,
      staff: staff()
    }).then(function (data) {
      state.analysis = data;
      renderResults(data);
      setStatus("");
    }).catch(function (err) {
      setStatus(err.message, "error");
    }).then(function () {
      $("analyseBtn").disabled = false; $("regenBtn").disabled = false;
    });
  }

  function renderResults(data) {
    var e = data.enquiry || {};
    FIELDS.forEach(function (f) { $("f_" + f).value = e[f] || ""; });
    $("f_urgency").value = ["High", "Normal", "Low"].indexOf(e.urgency) >= 0 ? e.urgency : "Normal";
    $("f_language").value = e.language || "";

    var missing = e.missing_info || [];
    var ul = $("missingList"); ul.innerHTML = "";
    missing.forEach(function (m) { var li = document.createElement("li"); li.textContent = m; ul.appendChild(li); });
    if (missing.length) show($("missingWrap")); else hide($("missingWrap"));

    if (e.is_new_enquiry === false) show($("notEnquiry")); else hide($("notEnquiry"));

    var r = data.reply || {};
    $("r_subject").value = r.subject || ("Re: " + (state.email.subject || ""));
    $("replyPreview").innerHTML = sanitizeHtml(r.html || "");

    var enquirerEmail = (e.email || "").trim().toLowerCase();
    var senderEmail = (state.email.fromEmail || "").trim().toLowerCase();
    if (enquirerEmail && enquirerEmail !== senderEmail) show($("newMsgBtn")); else hide($("newMsgBtn"));

    var name = [e.first_name, e.last_name].filter(Boolean).join(" ") || state.email.fromName || "New enquiry";
    $("h_deal_name").value = name + (e.matter_type ? " - " + e.matter_type : "");
    hide($("hubspotResult"));
    $("hubspotBtn").disabled = false;

    show($("results"));
  }

  /* ---------- reply ---------- */
  function currentReplyHtml() { return sanitizeHtml($("replyPreview").innerHTML); }

  function openReply() {
    var html = currentReplyHtml();
    try {
      state.item.displayReplyForm({ htmlBody: html });
    } catch (e) {
      setStatus("Could not open a reply window: " + e.message, "error");
    }
  }

  function openNewMessage() {
    var to = $("f_email").value.trim();
    try {
      Office.context.mailbox.displayNewMessageForm({
        toRecipients: to ? [to] : [],
        subject: $("r_subject").value,
        htmlBody: currentReplyHtml()
      });
    } catch (e) {
      setStatus("Could not open a new email: " + e.message, "error");
    }
  }

  function copyReply() {
    var text = $("replyPreview").innerText;
    function fallback() {
      var ta = document.createElement("textarea");
      ta.value = text; document.body.appendChild(ta); ta.select();
      try { document.execCommand("copy"); setStatus("Copied."); } catch (e) { setStatus("Copy failed. Select the text and copy it manually.", "error"); }
      ta.remove();
    }
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(function () { setStatus("Copied."); }, fallback);
    } else { fallback(); }
    setTimeout(function () { if ($("status").textContent === "Copied.") setStatus(""); }, 2000);
  }

  /* ---------- HubSpot ---------- */
  function logToHubspot() {
    var key = getKey();
    if (!key) { openSettings(); return; }
    var email = $("f_email").value.trim();
    if (!email || email.indexOf("@") < 1) {
      setStatus("An email address is needed to create the HubSpot contact.", "error");
      $("f_email").focus();
      return;
    }
    if (!$("logged").classList.contains("hidden") &&
        !window.confirm("This email was already logged to HubSpot. Log it again?")) return;

    var enquiry = {};
    FIELDS.forEach(function (f) { enquiry[f] = $("f_" + f).value.trim(); });
    enquiry.urgency = $("f_urgency").value;
    enquiry.language = $("f_language").value.trim();
    enquiry.missing_info = Array.prototype.map.call($("missingList").children, function (li) { return li.textContent; });

    $("hubspotBtn").disabled = true;
    setStatus("Saving to HubSpot...", "busy");

    postJSON(CFG.HUBSPOT_URL, {
      key: key,
      enquiry: enquiry,
      deal_name: $("h_deal_name").value.trim(),
      email: { subject: state.email.subject, fromEmail: state.email.fromEmail, receivedAt: state.email.receivedAt },
      staff: staff()
    }).then(function (data) {
      setStatus("");
      var box = $("hubspotResult");
      box.innerHTML = "";
      box.appendChild(document.createTextNode((data.contact_created ? "New contact created" : "Existing contact updated") + " and deal added. "));
      if (data.contact_url) {
        var a = document.createElement("a");
        a.href = data.contact_url; a.target = "_blank"; a.rel = "noopener"; a.textContent = "Open in HubSpot";
        box.appendChild(a);
      }
      show(box);
      markLogged();
    }).catch(function (err) {
      $("hubspotBtn").disabled = false;
      setStatus("HubSpot: " + err.message, "error");
    });
  }

  function markLogged() {
    var item = state.item;
    if (!item || !item.loadCustomPropertiesAsync) return;
    item.loadCustomPropertiesAsync(function (r) {
      if (r.status !== Office.AsyncResultStatus.Succeeded) return;
      var props = r.value;
      props.set(LOGGED_PROP, { by: staff().name, at: new Date().toISOString() });
      props.saveAsync(function () { /* best effort */ });
    });
  }

  /* ---------- start ---------- */
  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Outlook) return;

    $("settingsBtn").onclick = openSettings;
    $("closeSettingsBtn").onclick = closeSettings;
    $("saveSettingsBtn").onclick = function () {
      saveKey($("accessKey").value.trim(), function (ok) {
        closeSettings();
        setStatus(ok ? "Settings saved." : "Could not save settings.", ok ? "" : "error");
        if (ok) setTimeout(function () { setStatus(""); }, 2000);
      });
    };
    $("analyseBtn").onclick = analyse;
    $("regenBtn").onclick = analyse;
    $("replyBtn").onclick = openReply;
    $("newMsgBtn").onclick = openNewMessage;
    $("copyBtn").onclick = copyReply;
    $("hubspotBtn").onclick = logToHubspot;

    // Pinned pane: refresh when the user opens a different email
    if (Office.context.mailbox.addHandlerAsync) {
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadItem);
    }

    loadItem();
    if (!getKey()) openSettings();
  });

  // exposed for tests only
  window.__cw = { sanitizeHtml: sanitizeHtml };
})();
