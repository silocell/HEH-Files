/* My Files: everything you keep in the Launchpad Suite, in one place -
   what each file costs, unlocking, deleting, your videos, and buying
   more storage with Launchpad Pounds. */
(function () {
  var $ = function (id) { return document.getElementById(id); };
  var esc = Office.esc, SUITE = "https://suite.civilinitiatives.org", HOST = "https://hosting.civilinitiatives.org";
  var KINDS = { doc: ["📘", "Write", "/write/"], sheet: ["📗", "Grid", "/grid/"], deck: ["📙", "Present", "/present/"], note: ["📓", "Notebook", "/notebook/"],
    draw: ["🎨", "Draw", "/draw/"], design: ["🖌", "Design", "/design/"], form: ["📋", "Forms", "/forms/"] };
  var sb = null, me = null, rows = [], vids = [], ups = null, urls = {}, info = null, bill = null, tab = "all", q = "", busy = 0;
  var UP_MAX = 25 * 1024 * 1024, SHARE = "https://files.civilinitiatives.org/s/?t=";
  var CODE = /\.(html?|xhtml|xml|svg|js|mjs|cjs|ts|php|css|json|csv|py|java|c|cpp|cs|rb|go|rs|sh|lua|sql|ya?ml|ini|toml)$/i;
  var PROGRAM = /\.(exe|msi|msix|appx|bat|cmd|com|scr|pif|cpl|msc|ps1|psm1|vbs|vbe|jse|wsf|wsh|hta|jar|apk|aab|xapk|ipa|dmg|pkg|deb|rpm|run|bin|lnk|reg|dll|sys|iso|img|vhdx?)$/i;
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
      sb.rpc("list_videos", { mine: true, lim: 200 }).then(function (r) { return r.error ? [] : r.data || []; }, function () { return []; }),
      sb.from("file_uploads").select("id,name,mime,kind,bytes,path,status,share_token,created_at").eq("user_id", me.id).eq("status", "ready").order("created_at", { ascending: false }).limit(1000)
        .then(function (r) { return r.error ? null : r.data || []; }, function () { return null; })
    ]).then(function (res) {
      if (res[0].error) throw res[0].error;
      rows = res[0].data || []; info = res[1]; bill = res[2]; vids = res[3]; ups = res[4];
      return signAll();
    }).then(function () {
      $("status").textContent = busy ? $("status").textContent : ""; draw();
    }).catch(function (e) { $("main").innerHTML = '<p class="empty">Couldn\'t load your files: ' + esc(e.message || e) + "</p>"; });
  }
  // private links to your uploads (they last an hour; the page gets new ones each time it loads)
  function signAll() {
    urls = {};
    var paths = (ups || []).map(function (u) { return u.path; });
    if (!paths.length) return Promise.resolve();
    return sb.storage.from("uploads").createSignedUrls(paths, 3600).then(function (r) {
      (r.data || []).forEach(function (x) { if (x && x.signedUrl && x.path) urls[x.path] = x.signedUrl; });
    }, function () {});
  }
  function draw() {
    var mine = rows.filter(function (r) { return r.user_id === me.id; }), shared = rows.filter(function (r) { return r.user_id !== me.id; });
    var perDay = bill && !bill.exempt ? (bill.files || []).filter(function (f) { return f.state === "paying"; }).length * bill.cost : 0;
    var h = "<h1>🗂 My Files</h1><p class=\"lead\">Everything you keep in the Launchpad Suite, and the files you upload. Open one to carry on with it.</p>";
    // storage and paying for it
    if (info) {
      var pct = Math.min(100, info.mine / Math.max(1, info.person_limit) * 100);
      h += '<section class="card2"><div style="display:flex;justify-content:space-between;flex-wrap:wrap;gap:6px"><b>☁️ Your storage: ' + size(info.mine) + " of " + size(info.person_limit) + "</b>" +
        '<span style="color:var(--soft);font-size:.86rem">' + info.files + " files" + (info.extra_gb ? " · " + info.extra_gb + " GB bought" : "") + "</span></div>" +
        '<div style="height:10px;border-radius:999px;background:var(--tool);overflow:hidden;margin:8px 0"><div style="height:100%;width:' + pct.toFixed(1) + "%;background:" + (pct > 90 ? "#dc2626" : pct > 70 ? "#f59e0b" : "var(--a1)") + '"></div></div>' +
        '<p style="color:var(--soft);font-size:.84rem;margin:0">The whole Launchpad has ' + size(info.limit) + " (1 GB for each of the " + info.people + " people, plus storage people have bought), and " + size(info.used) + " is used.</p>" +
        (info.video_share ? '<p style="font-size:.86rem;margin:8px 0 0">🎬 Videos have their own space: yours use <b>' + size(info.videos || 0) + "</b> of your " + size(info.video_share) + " share, out of " + size(info.video_limit) + " for the whole Launchpad (" + size(info.video_used || 0) + " used). Videos are up to 30 minutes, in 720p.</p>" : "") +
        (bill ? '<p style="font-size:.9rem;margin:10px 0 0">' + (bill.exempt ? "💰 Pay as you go: your files are free (you're a Launchpad admin)."
          : "💰 Pay as you go: a file is free for its first 48 hours, then has a grace period (a day and a half, or 2 days at the weekend), then costs " + bill.cost + "p a day. School holidays are free" + (bill["break"] ? " - it's " + esc(bill["break"]) + " now" : "") + ".<br>Your files cost <b>" + perDay + "p a day</b> and you have <b>" + bill.pence + "p</b>. A file that can't be paid for is locked until it's paid for - it can't be opened or downloaded, but nothing is deleted.") + "</p>" : "") +
        '<div class="buy"><button type="button" class="of-btn" id="buyBtn"' + ((info.extra_gb || 0) >= (info.extra_max != null ? info.extra_max : 20) ? " disabled" : "") + ">➕ Buy 1 GB more - 25 Launchpad Pounds</button>" +
        '<span style="color:var(--soft);font-size:.86rem">You have ' + (info.pounds || 0) + " Pounds. It's yours for good (up to " + (info.extra_max != null ? info.extra_max : 20) + " GB extra" + ((info.extra_gb || 0) >= (info.extra_max != null ? info.extra_max : 20) ? " - you've got the most anyone can have" : "") + ").</span></div></section>";
    }
    h += '<div class="upbar"><button type="button" class="of-btn" id="upBtn">⬆ Upload files</button><span>Pictures, documents, audio, code and more - up to 25 MB each. Only you can see them, unless you make a share link.</span></div>' +
      '<input type="file" id="upIn" multiple hidden>';
    h += '<div class="tabs3" role="toolbar" aria-label="Show">' + [["all", "All"]].concat(Object.keys(KINDS).filter(function (k) { return k !== "form"; }).map(function (k) { return [k, KINDS[k][0] + " " + KINDS[k][1]]; }))
      .concat([["up", "📤 Uploads"], ["shared", "👥 Shared with me"], ["videos", "🎬 Videos"]]).map(function (t) { return '<button type="button" data-tab="' + t[0] + '" aria-pressed="' + (tab === t[0]) + '">' + t[1] + "</button>"; }).join("") + "</div>" +
      '<input class="search" id="q" type="search" placeholder="Find a file by its name" value="' + esc(q) + '" aria-label="Find a file">';
    if (tab === "videos") h += vidRows();
    else if (tab === "up") h += upRows();
    else {
      var list = (tab === "shared" ? shared : tab === "all" ? mine : mine.filter(function (r) { return r.kind === tab; }))
        .filter(function (r) { return !q || (r.title || "").toLowerCase().indexOf(q.toLowerCase()) >= 0; });
      h += '<div class="rows">' + (list.map(fileRow).join("") || '<p class="empty">' + (tab === "shared" ? "Nothing's been shared with you yet." : "No files here yet - make one in the <a href=\"" + SUITE + "/\">Suite</a>.") + "</p>") + "</div>";
    }
    $("main").innerHTML = h;
    $("upIn").onchange = function () { var f = [].slice.call(this.files || []); this.value = ""; upload(f); };
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
  function upIcon(u) { return u.kind === "image" ? "🖼" : u.kind === "audio" ? "🎵" : CODE.test(u.name) ? "💻" : u.kind === "doc" ? "📄" : "📦"; }
  function dl(u) { var s = urls[u.path]; return s ? s + (s.indexOf("?") >= 0 ? "&" : "?") + "download=" + encodeURIComponent(u.name) : ""; }
  function upRows() {
    if (ups === null) return '<p class="empty">Uploading files isn\'t switched on yet - ask the Launchpad team.</p>';
    var list = ups.filter(function (u) { return !q || u.name.toLowerCase().indexOf(q.toLowerCase()) >= 0; });
    return '<div class="drop" id="drop">Drop files here to upload them</div><div class="rows">' + (list.map(function (u) {
      var s = urls[u.path], img = u.mime.indexOf("image/") === 0 && s;
      // pictures, PDFs, audio and plain text open in a new tab; anything else downloads
      var view = s && /^(image\/|audio\/|application\/pdf|text\/plain)/.test(u.mime) ? s : dl(u);
      return '<div class="row"><span class="ic">' + (img ? '<img class="thumb" src="' + esc(s) + '" alt="" loading="lazy">' : upIcon(u)) + '</span><span class="t">' +
        (view ? '<a href="' + esc(view) + '" target="_blank" rel="noopener">' + esc(u.name) + "</a>" : "<b>" + esc(u.name) + "</b>") +
        "<small>Uploaded " + when(u.created_at) + " · " + size(u.bytes) + "</small></span>" +
        (u.share_token ? '<span class="st shared">🔗 Shared by link</span>' : "") + '<span class="acts">' +
        (dl(u) ? '<a href="' + esc(dl(u)) + '" download>Download</a>' : "") +
        (u.share_token ? '<button type="button" data-ucopy="' + esc(u.id) + '">Copy link</button><button type="button" data-uunshare="' + esc(u.id) + '">Stop sharing</button>'
          : '<button type="button" data-ushare="' + esc(u.id) + '">🔗 Share</button>') +
        '<button type="button" data-uren="' + esc(u.id) + '">Rename</button>' +
        '<button type="button" data-udel="' + esc(u.id) + '" aria-label="Delete ' + esc(u.name) + '">🗑</button></span></div>';
    }).join("") || '<p class="empty">' + (q ? "No uploads match that." : "Nothing uploaded yet - tap <b>Upload files</b>, or drop files here.") + "</p>") + "</div>";
  }
  function err(e) { return String((e && (e.message || e.error)) || e).replace(/^.*?: /, ""); }
  function upload(files) {
    if (!files.length) return;
    if (ups === null) { Office.toast("Uploading files isn't switched on yet.", true); return; }
    tab = "up"; draw();
    var bad = files.filter(function (f) { return PROGRAM.test(f.name); }), big = files.filter(function (f) { return f.size > UP_MAX; });
    if (bad.length) Office.toast("Programs and installers can't be uploaded (" + bad.map(function (f) { return f.name; }).join(", ") + ").", true);
    else if (big.length) Office.toast(big[0].name + " is too big - files can be up to 25 MB.", true);
    files = files.filter(function (f) { return bad.indexOf(f) < 0 && big.indexOf(f) < 0 && f.size > 0; });
    var done = 0, failed = 0, i = 0;
    busy++;
    function next() {
      if (i >= files.length) {
        busy--; $("status").textContent = "";
        if (done) Office.toast(done === 1 ? "Uploaded." : "Uploaded " + done + " files.");
        return reload();
      }
      var f = files[i++], slot = null;
      $("status").textContent = "Uploading " + f.name + (files.length > 1 ? " (" + i + " of " + files.length + ")" : "") + "…";
      return sb.rpc("upload_start", { fname: f.name, fsize: f.size }).then(function (r) {
        if (r.error) throw r.error; slot = r.data;
        // sent as raw bytes, so it's stored with the type the Launchpad chose (a web page is kept as plain text), not the one the browser guessed
        return f.arrayBuffer();
      }).then(function (bytes) {
        bytes = noHidden(bytes);
        return sb.storage.from("uploads").upload(slot.path, bytes, { contentType: slot.mime, upsert: false, cacheControl: "3600" });
      }).then(function (r) {
        if (r.error) throw r.error;
        return sb.rpc("upload_done", { fid: slot.id });
      }).then(function (r) {
        if (r.error) throw r.error; done++;
      }).catch(function (e) {
        failed++; Office.toast(f.name + ": " + err(e), true);
        if (slot) sb.rpc("upload_delete", { fid: slot.id }).then(function () {}, function () {});
      }).then(next);
    }
    return next();
  }
  // take the hidden camera data (like where a photo was taken) out of JPEG and PNG pictures
  function noHidden(buf) {
    try {
      var u = new Uint8Array(buf), out = [], i, n;
      if (u[0] === 0xFF && u[1] === 0xD8) {
        out.push(u.subarray(0, 2)); i = 2;
        while (i + 4 <= u.length && u[i] === 0xFF) {
          var mk = u[i + 1];
          if (mk === 0xDA || mk === 0xD9) break;                       // the picture itself starts here
          n = (u[i + 2] << 8 | u[i + 3]) + 2;
          if (mk !== 0xE1 && mk !== 0xED) out.push(u.subarray(i, i + n)); // drop EXIF/XMP and Photoshop data
          i += n;
        }
        out.push(u.subarray(i));
      } else if (u[0] === 0x89 && u[1] === 0x50 && u[2] === 0x4E && u[3] === 0x47) {
        out.push(u.subarray(0, 8)); i = 8;
        while (i + 12 <= u.length) {
          n = ((u[i] << 24) >>> 0) + (u[i + 1] << 16) + (u[i + 2] << 8) + u[i + 3] + 12;
          var ty = String.fromCharCode(u[i + 4], u[i + 5], u[i + 6], u[i + 7]);
          if (ty !== "eXIf" && ty !== "tEXt" && ty !== "iTXt" && ty !== "zTXt") out.push(u.subarray(i, i + n));
          i += n;
          if (ty === "IEND") break;
        }
      } else return buf;
      var len = out.reduce(function (a, x) { return a + x.length; }, 0), res = new Uint8Array(len), at = 0;
      out.forEach(function (x) { res.set(x, at); at += x.length; });
      return res.buffer;
    } catch (e) { return buf; }
  }
  function findUp(id) { return (ups || []).filter(function (u) { return u.id === id; })[0]; }
  function copy(t) {
    return (navigator.clipboard ? navigator.clipboard.writeText(t) : Promise.reject()).then(function () { Office.toast("Link copied - anyone signed in to the Launchpad can open it."); },
      function () { prompt("Copy this link", t); });
  }
  // dropping files anywhere on the page uploads them
  document.addEventListener("dragover", function (e) { if (ups !== null && e.dataTransfer && [].indexOf.call(e.dataTransfer.types || [], "Files") >= 0) { e.preventDefault(); document.body.classList.add("dragging"); } });
  document.addEventListener("dragleave", function (e) { if (!e.relatedTarget) document.body.classList.remove("dragging"); });
  document.addEventListener("drop", function (e) {
    document.body.classList.remove("dragging");
    if (ups === null || !e.dataTransfer || !e.dataTransfer.files || !e.dataTransfer.files.length) return;
    e.preventDefault(); upload([].slice.call(e.dataTransfer.files));
  });
  function vidRows() {
    var list = vids.filter(function (v) { return !q || (v.title || "").toLowerCase().indexOf(q.toLowerCase()) >= 0; });
    return '<p style="color:var(--soft);font-size:.88rem;margin:0 0 10px">Videos you\'ve posted or kept from Vids. They use the video space, not your file storage. <a href="' + HOST + '/">See everyone\'s videos →</a></p><div class="rows">' +
      (list.map(function (v) {
        var st = v.hidden ? '<span class="st locked">Hidden - reported</span>' : v.shown ? '<span class="st free">Everyone can watch</span>' : '<span class="st paying">Just for you</span>';
        return '<div class="row"><span class="ic">' + (v.short ? "📱" : "🎬") + '</span><span class="t"><a href="' + HOST + (v.short ? "/shorts/?v=" : "/watch/?v=") + encodeURIComponent(v.id) + '">' + esc(v.title) + "</a><small>" + when(v.at) + " · " + size(v.bytes) + " · " + v.views + " views · " + v.likes + " ❤️</small></span>" + st +
          '<span class="acts"><button type="button" data-vshow="' + esc(v.id) + '">' + (v.shown ? "Make it just for me" : "Let everyone watch") + '</button><button type="button" data-vdel="' + esc(v.id) + '" aria-label="Delete ' + esc(v.title) + '">🗑</button></span></div>';
      }).join("") || '<p class="empty">No videos yet. Make one in <a href="' + SUITE + '/vids/">Vids</a> and post it.</p>') + "</div>";
  }
  $("main").addEventListener("click", function (ev) {
    var b = ev.target.closest("button"); if (!b) return;
    var d = b.dataset;
    if (d.tab) { tab = d.tab; draw(); return; }
    if (b.id === "upBtn") { $("upIn").click(); return; }
    if (d.ushare || d.uunshare) {
      b.disabled = true;
      sb.rpc("upload_share", { fid: d.ushare || d.uunshare, on_: !!d.ushare }).then(function (r) {
        if (r.error) throw r.error;
        var u = findUp(d.ushare || d.uunshare); if (u) u.share_token = r.data.share_token;
        draw();
        if (r.data.share_token) return copy(SHARE + r.data.share_token); else Office.toast("Stopped sharing - the old link doesn't work any more.");
      }).catch(function (e) { b.disabled = false; Office.toast(err(e), true); });
      return;
    }
    if (d.ucopy) { var cu = findUp(d.ucopy); if (cu && cu.share_token) copy(SHARE + cu.share_token); return; }
    if (d.uren) {
      var ru = findUp(d.uren), nn = prompt("New name", ru.name);
      if (nn == null || !nn.trim()) return;
      sb.rpc("upload_rename", { fid: d.uren, fname: nn.trim() }).then(function (r) { if (r.error) Office.toast(err(r.error), true); else reload(); });
      return;
    }
    if (d.udel) {
      if (d.sure !== "1") { d.sure = "1"; b.textContent = "Delete?"; return; }
      b.disabled = true;
      sb.rpc("upload_delete", { fid: d.udel }).then(function (r) {
        if (r.error) throw r.error;
        return sb.storage.from("uploads").remove([r.data]);
      }).then(function () { Office.toast("Deleted."); reload(); }).catch(function (e) { b.disabled = false; Office.toast(err(e), true); });
      return;
    }
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
