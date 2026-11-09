/* My Files: everything you keep in the Launchpad Suite, in one place -
   what each file costs, unlocking, deleting, your videos, and buying
   more storage with Launchpad Pounds. */
(function () {
  var $ = function (id) { return document.getElementById(id); };
  var esc = Office.esc, SUITE = "https://suite.civilinitiatives.org", HOST = "https://hosting.civilinitiatives.org";
  var KINDS = { doc: ["📘", "Write", "/write/"], sheet: ["📗", "Grid", "/grid/"], deck: ["📙", "Present", "/present/"], note: ["📓", "Notebook", "/notebook/"],
    draw: ["🎨", "Draw", "/draw/"], design: ["🖌", "Design", "/design/"], form: ["📋", "Forms", "/forms/"] };
  var sb = null, me = null, rows = [], vids = [], info = null, bill = null, tab = "all", q = "";
  Office.themeToggle($("themeBtn"));

  function size(n) { return Office.size(n || 0); }
  function when(s) { var d = new Date(s); return d.toLocaleDateString("en-GB", { day: "numeric", month: "short" }) + " " + d.toTimeString().slice(0, 5); }

  Office.init("doc").then(function (st) {
    sb = Office.sb(); me = st.user;
    if (!me) { $("main").innerHTML = '<div class="signin"><h1>🗂 My Files</h1><p class="lead">Sign in on the Launchpad to see your Suite files and videos.</p><a class="of-btn" href="https://civilinitiatives.org/">Sign in on the Launchpad</a></div>'; return; }
    if (!st.cloud) { $("main").innerHTML = '<div class="signin"><h1>🗂 My Files</h1><p class="lead">Saving files to your account isn\'t switched on yet - ask the Launchpad team.</p></div>'; return; }
    return reload();
  });
  function reload() {
    $("status").textContent = "Loading…";
    return Promise.all([
      sb.from("office_files").select("id,kind,title,updated_at,user_id,bytes,locked_at").order("updated_at", { ascending: false }).limit(1000),
      Office.storage(),
      Office.billNow(),
      sb.rpc("list_videos", { mine: true, lim: 200 }).then(function (r) { return r.error ? [] : r.data || []; }, function () { return []; })
    ]).then(function (res) {
      if (res[0].error) throw res[0].error;
      rows = res[0].data || []; info = res[1]; bill = res[2]; vids = res[3];
      $("status").textContent = ""; draw();
    }).catch(function (e) { $("main").innerHTML = '<p class="empty">Couldn\'t load your files: ' + esc(e.message || e) + "</p>"; });
  }
  function draw() {
    var mine = rows.filter(function (r) { return r.user_id === me.id; }), shared = rows.filter(function (r) { return r.user_id !== me.id; });
    var perDay = bill && !bill.exempt ? (bill.files || []).filter(function (f) { return f.state === "paying"; }).length * bill.cost : 0;
    var h = "<h1>🗂 My Files</h1><p class=\"lead\">Everything you keep in the Launchpad Suite. Open a file to carry on with it.</p>";
    // storage and paying for it
    if (info) {
      var pct = Math.min(100, info.mine / Math.max(1, info.person_limit) * 100);
      h += '<section class="card2"><div style="display:flex;justify-content:space-between;flex-wrap:wrap;gap:6px"><b>☁️ Your storage: ' + size(info.mine) + " of " + size(info.person_limit) + "</b>" +
        '<span style="color:var(--soft);font-size:.86rem">' + info.files + " files · videos " + size(info.videos || 0) + (info.extra_gb ? " · " + info.extra_gb + " GB bought" : "") + "</span></div>" +
        '<div style="height:10px;border-radius:999px;background:var(--tool);overflow:hidden;margin:8px 0"><div style="height:100%;width:' + pct.toFixed(1) + "%;background:" + (pct > 90 ? "#dc2626" : pct > 70 ? "#f59e0b" : "var(--a1)") + '"></div></div>' +
        '<p style="color:var(--soft);font-size:.84rem;margin:0">The whole Launchpad has ' + size(info.limit) + " (1 GB for each of the " + info.people + " people, plus storage people have bought), and " + size(info.used) + " is used.</p>" +
        (bill ? '<p style="font-size:.9rem;margin:10px 0 0">' + (bill.exempt ? "💰 Pay as you go: your files are free (you're a Launchpad admin)."
          : "💰 Pay as you go: a file is free for its first 48 hours, then has a grace period (a day and a half, or 2 days at the weekend), then costs " + bill.cost + "p a day. School holidays are free" + (bill["break"] ? " - it's " + esc(bill["break"]) + " now" : "") + ".<br>Your files cost <b>" + perDay + "p a day</b> and you have <b>" + bill.pence + "p</b>. A file that can't be paid for is locked until it's paid for - it can't be opened or downloaded, but nothing is deleted.") + "</p>" : "") +
        '<div class="buy"><button type="button" class="of-btn" id="buyBtn"' + ((info.extra_gb || 0) >= 20 ? " disabled" : "") + ">➕ Buy 1 GB more - 25 Launchpad Pounds</button>" +
        '<span style="color:var(--soft);font-size:.86rem">You have ' + (info.pounds || 0) + " Pounds. It's yours for good (up to 20 GB extra).</span></div></section>";
    }
    h += '<div class="tabs3" role="toolbar" aria-label="Show">' + [["all", "All"]].concat(Object.keys(KINDS).filter(function (k) { return k !== "form"; }).map(function (k) { return [k, KINDS[k][0] + " " + KINDS[k][1]]; }))
      .concat([["shared", "👥 Shared with me"], ["videos", "🎬 Videos"]]).map(function (t) { return '<button type="button" data-tab="' + t[0] + '" aria-pressed="' + (tab === t[0]) + '">' + t[1] + "</button>"; }).join("") + "</div>" +
      '<input class="search" id="q" type="search" placeholder="Find a file by its name" value="' + esc(q) + '" aria-label="Find a file">';
    if (tab === "videos") h += vidRows();
    else {
      var list = (tab === "shared" ? shared : tab === "all" ? mine : mine.filter(function (r) { return r.kind === tab; }))
        .filter(function (r) { return !q || (r.title || "").toLowerCase().indexOf(q.toLowerCase()) >= 0; });
      h += '<div class="rows">' + (list.map(fileRow).join("") || '<p class="empty">' + (tab === "shared" ? "Nothing's been shared with you yet." : "No files here yet - make one in the <a href=\"" + SUITE + "/\">Suite</a>.") + "</p>") + "</div>";
    }
    $("main").innerHTML = h;
    var qi = $("q"); qi.oninput = function () { q = qi.value; var at = qi.selectionStart; draw(); var n = $("q"); n.focus(); n.setSelectionRange(at, at); };
  }
  function fileRow(r) {
    var k = KINDS[r.kind] || ["📄", r.kind, "/"], b = bill && bill.byId && bill.byId[r.id], shared = r.user_id !== me.id;
    var st = shared ? '<span class="st shared">👥 Shared with you</span>' : b && !bill.exempt ? '<span class="st ' + b.state + '">' + esc(Office.billLabel(b)) + "</span>" : "";
    var locked = !!r.locked_at;
    return '<div class="row"><span class="ic">' + k[0] + '</span><span class="t">' + (locked ? "<b>" + esc(r.title || "Untitled") + "</b>" : '<a href="' + SUITE + k[2] + "?id=" + encodeURIComponent(r.id) + '">' + esc(r.title || "Untitled") + "</a>") +
      "<small>" + k[1] + " · " + when(r.updated_at) + (r.bytes ? " · " + size(r.bytes) : "") + "</small></span>" + st + '<span class="acts">' +
      (locked ? '<button type="button" class="pay" data-unlock="' + esc(r.id) + '">🔓 Unlock' + (b && b.owed ? " " + b.owed + "p" : "") + "</button>" : '<a href="' + SUITE + k[2] + "?id=" + encodeURIComponent(r.id) + '">Open</a>') +
      (!shared && !locked ? '<button type="button" data-ren="' + esc(r.id) + '">Rename</button>' : "") +
      (shared ? '<button type="button" data-leave="' + esc(r.id) + '">Remove</button>' : '<button type="button" data-del="' + esc(r.id) + '" aria-label="Delete ' + esc(r.title || "Untitled") + '">🗑</button>') + "</span></div>";
  }
  function vidRows() {
    var list = vids.filter(function (v) { return !q || (v.title || "").toLowerCase().indexOf(q.toLowerCase()) >= 0; });
    return '<p style="color:var(--soft);font-size:.88rem;margin:0 0 10px">Videos you\'ve posted or kept from Vids. They count towards your storage. <a href="' + HOST + '/">See everyone\'s videos →</a></p><div class="rows">' +
      (list.map(function (v) {
        var st = v.hidden ? '<span class="st locked">Hidden - reported</span>' : v.shown ? '<span class="st free">Everyone can watch</span>' : '<span class="st paying">Just for you</span>';
        return '<div class="row"><span class="ic">🎬</span><span class="t"><a href="' + HOST + "/watch/?v=" + encodeURIComponent(v.id) + '">' + esc(v.title) + "</a><small>" + when(v.at) + " · " + size(v.bytes) + " · " + v.views + " views · " + v.likes + " ❤️</small></span>" + st +
          '<span class="acts"><button type="button" data-vshow="' + esc(v.id) + '">' + (v.shown ? "Make it just for me" : "Let everyone watch") + '</button><button type="button" data-vdel="' + esc(v.id) + '" aria-label="Delete ' + esc(v.title) + '">🗑</button></span></div>';
      }).join("") || '<p class="empty">No videos yet. Make one in <a href="' + SUITE + '/vids/">Vids</a> and post it.</p>') + "</div>";
  }
  $("main").addEventListener("click", function (ev) {
    var b = ev.target.closest("button"); if (!b) return;
    var d = b.dataset;
    if (d.tab) { tab = d.tab; draw(); return; }
    if (b.id === "buyBtn") {
      if (b.dataset.sure !== "1") { b.dataset.sure = "1"; b.textContent = "Spend 25 Pounds on 1 GB more? Tap again"; return; }
      b.disabled = true;
      sb.rpc("buy_storage").then(function (r) { if (r.error) throw r.error; Office.toast("You've got " + r.data.extra_gb + " GB extra now. 🎉"); reload(); })
        .catch(function (e) { b.disabled = false; b.dataset.sure = ""; b.textContent = "➕ Buy 1 GB more - 25 Launchpad Pounds"; Office.toast((e.message || String(e)).replace(/^.*?: /, ""), true); });
      return;
    }
    if (d.unlock) {
      b.disabled = true;
      Office.unlock(d.unlock).then(function () { Office.toast("Unlocked."); reload(); }).catch(function (e) { b.disabled = false; Office.toast((e.message || String(e)).replace(/^.*?: /, ""), true); });
      return;
    }
    if (d.ren) {
      var r = rows.filter(function (x) { return x.id === d.ren; })[0], t = prompt("New name", r.title || "");
      if (t == null || !t.trim()) return;
      sb.from("office_files").update({ title: t.trim().slice(0, 120) }).eq("id", d.ren).then(function (x) { if (x.error) Office.toast(x.error.message, true); else reload(); });
      return;
    }
    if (d.del) {
      if (d.sure !== "1") { d.sure = "1"; b.textContent = "Delete?"; return; }
      sb.from("office_files").delete().eq("id", d.del).then(function (x) { if (x.error) Office.toast(x.error.message, true); else { Office.toast("Deleted."); reload(); } });
      return;
    }
    if (d.leave) {
      sb.rpc("unshare_file", { fid: d.leave, who: me.id }).then(function (x) { if (x.error) Office.toast(x.error.message, true); else reload(); });
      return;
    }
    if (d.vshow) {
      var v = vids.filter(function (x) { return x.id === d.vshow; })[0];
      sb.rpc("edit_video", { vid: v.id, t: v.title, a: v.about, show: !v.shown }).then(function (x) { if (x.error) Office.toast(x.error.message, true); else reload(); });
      return;
    }
    if (d.vdel) {
      if (d.sure !== "1") { d.sure = "1"; b.textContent = "Delete?"; return; }
      sb.rpc("delete_video", { vid: d.vdel }).then(function (x) {
        if (x.error) throw x.error;
        return x.data ? sb.storage.from("videos").remove([x.data]) : null;
      }).then(function () { Office.toast("Deleted."); reload(); }).catch(function (e) { Office.toast(e.message || e, true); });
    }
  });
})();
