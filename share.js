/* A file someone shared from My Files. Anyone signed in to the Launchpad can open it;
   anyone can report it to the Launchpad team. */
(function () {
  var $ = function (id) { return document.getElementById(id); };
  var esc = Office.esc, tok = (new URLSearchParams(location.search).get("t") || "").replace(/[^a-f0-9]/gi, "");
  Office.themeToggle($("themeBtn"));
  function msg(h) { $("main").innerHTML = '<div class="empty">' + h + "</div>"; }
  function when(s) { return new Date(s).toLocaleDateString("en-GB", { day: "numeric", month: "long", year: "numeric" }); }
  Office.init("doc").then(function (st) {
    if (!st.user) return msg("<h1>🔗 A shared file</h1><p>Sign in on the Launchpad to open it.</p><a class=\"of-btn\" href=\"https://civilinitiatives.org/\">Sign in on the Launchpad</a>");
    if (tok.length < 32) return msg("<h1>That link isn't complete</h1><p>Ask whoever sent it to copy it again.</p>");
    var sb = Office.sb();
    return sb.rpc("upload_shared", { tok: tok }).then(function (r) {
      if (r.error) throw r.error;
      var f = r.data;
      if (!f) return msg("<h1>This file isn't shared any more</h1><p>Its owner stopped sharing it, or deleted it.</p>");
      return sb.storage.from("uploads").createSignedUrl(f.path, 3600).then(function (s) {
        if (s.error) throw s.error;
        show(sb, f, s.data.signedUrl);
      });
    });
  }).catch(function (e) { msg("Couldn't open this file: " + esc(e.message || e)); });

  function show(sb, f, url) {
    var dl = url + (url.indexOf("?") >= 0 ? "&" : "?") + "download=" + encodeURIComponent(f.name), v;
    if (f.mime.indexOf("image/") === 0) v = '<img src="' + esc(url) + '" alt="' + esc(f.name) + '">';
    else if (f.mime.indexOf("audio/") === 0) v = '<audio controls preload="metadata" src="' + esc(url) + '"></audio>';
    else if (f.mime === "text/plain") v = '<pre id="txt">Loading…</pre>';
    else v = '<div class="big" aria-hidden="true">' + (f.kind === "doc" ? "📄" : "📦") + "</div>";
    $("main").innerHTML = "<h1>" + esc(f.name) + '</h1><p class="meta">Shared by ' + esc(f.owner) + " · " + Office.size(f.bytes) + " · uploaded " + when(f.at) + "</p>" +
      '<div class="view">' + v + "</div>" +
      '<div class="acts"><a class="of-btn" href="' + esc(dl) + '" download>⬇ Download</a>' +
      (f.mime === "application/pdf" ? '<a class="of-btn ghost" href="' + esc(url) + '" target="_blank" rel="noopener">Open the PDF</a>' : "") +
      (f.mine ? '<span class="note" style="margin:0">This is your file - you can stop sharing it in <a href="/">My Files</a>.</span>' : '<button class="of-btn ghost" type="button" id="rep">🚩 Report</button>') + "</div>" +
      '<p class="note">Files people upload aren\'t checked before they\'re shared. If something here isn\'t OK, report it and the Launchpad team will look at it.</p>';
    if (f.mime === "text/plain") fetch(url).then(function (r) { return r.text(); }).then(function (t) {
      $("txt").textContent = t.length > 200000 ? t.slice(0, 200000) + "\n\n… (download it to see the rest)" : t;
    }).catch(function () { $("txt").textContent = "Couldn't show this - download it instead."; });
    var rep = $("rep");
    if (rep) rep.onclick = function () {
      var why = prompt("What's wrong with this file? (The Launchpad team will see this.)", "");
      if (why == null) return;
      rep.disabled = true;
      sb.rpc("upload_report", { tok: tok, why: why.trim() || "Not OK" }).then(function (r) {
        if (r.error) throw r.error;
        rep.textContent = "Reported - thank you"; Office.toast("Reported. The Launchpad team will look at it.");
      }).catch(function (e) { rep.disabled = false; Office.toast(String(e.message || e).replace(/^.*?: /, ""), true); });
    };
  }
})();
