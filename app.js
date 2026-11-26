(function () {
  var $ = function (id) { return document.getElementById(id); };
  var FIELDS = ["so", "kyhieu", "trichyeu", "coquan", "ngay"];
  var pending = null;

  // ---------- Khởi tạo ----------
  Office.onReady(function () {
    pdfjsLib.GlobalWorkerOptions.workerSrc =
      "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
    var drop = $("drop"), fileInput = $("file");
    drop.addEventListener("click", function () { fileInput.click(); });
    fileInput.addEventListener("change", function () {
      if (fileInput.files[0]) handleFile(fileInput.files[0]);
      fileInput.value = "";
    });
    ["dragenter", "dragover"].forEach(function (e) {
      drop.addEventListener(e, function (ev) { ev.preventDefault(); drop.classList.add("over"); });
    });
    ["dragleave", "drop"].forEach(function (e) {
      drop.addEventListener(e, function (ev) { ev.preventDefault(); drop.classList.remove("over"); });
    });
    drop.addEventListener("drop", function (ev) {
      var f = ev.dataTransfer && ev.dataTransfer.files[0];
      if (f) handleFile(f);
    });
    // Chặn việc thả file ra ngoài vùng làm Excel mở file
    document.addEventListener("dragover", function (e) { e.preventDefault(); });
    document.addEventListener("drop", function (e) { e.preventDefault(); });

    $("btnCancel").addEventListener("click", function () { clearForm(); setStatus("", ""); });
    $("btnSave").addEventListener("click", onSave);
    $("btnBack").addEventListener("click", function () { stopWaiting(); });
  });

  function setStatus(msg, cls) { var s = $("status"); s.textContent = msg; s.className = cls || ""; }
  function clearForm() { FIELDS.forEach(function (f) { $(f).value = ""; }); }

  // ---------- Đọc PDF ----------
  async function handleFile(file) {
    if (!/\.pdf$/i.test(file.name) && file.type !== "application/pdf") {
      setStatus("Chỉ hỗ trợ file PDF.", "err"); return;
    }
    setStatus("Đang đọc file...", "");
    try {
      var buf = await file.arrayBuffer();
      var pdf = await pdfjsLib.getDocument({ data: buf }).promise;
      var page = await pdf.getPage(1);
      var width = page.getViewport({ scale: 1 }).width;
      var tc = await page.getTextContent();
      var items = tc.items
        .filter(function (i) { return i.str && i.str.trim(); })
        .map(function (i) {
          return { s: i.str.normalize("NFC"), x: i.transform[4], y: i.transform[5], w: i.width };
        });
      if (!items.length) {
        clearForm();
        setStatus("Không đọc được chữ trong file (có thể là PDF scan). Hãy điền tay.", "warn");
        return;
      }
      var r = parse(items, width);
      FIELDS.forEach(function (f) { $(f).value = r[f] || ""; });
      var empty = FIELDS.filter(function (f) { return !r[f]; }).length;
      setStatus(empty ? "Đã điền tự động, còn " + empty + " ô trống - vui lòng kiểm tra." :
        "Đã điền đủ các ô - vui lòng kiểm tra lại trước khi lưu.", empty ? "warn" : "ok");
    } catch (e) {
      setStatus("Lỗi đọc file: " + (e && e.message ? e.message : e), "err");
    }
  }

  // ---------- Nhận diện trường thông tin ----------
  function toLines(items) {
    var sorted = items.slice().sort(function (a, b) { return b.y - a.y || a.x - b.x; });
    var lines = [], cur = null;
    sorted.forEach(function (it) {
      if (!cur || Math.abs(cur.y - it.y) > 3) { cur = { y: it.y, parts: [] }; lines.push(cur); }
      cur.parts.push(it);
    });
    return lines.map(function (l) {
      l.parts.sort(function (a, b) { return a.x - b.x; });
      var t = "", prev = null;
      l.parts.forEach(function (p) {
        if (prev && !/\s$/.test(prev.s) && !/^\s/.test(p.s) && p.x - (prev.x + prev.w) > 1) t += " ";
        t += p.s; prev = p;
      });
      return t.replace(/\s+/g, " ").trim();
    }).filter(Boolean);
  }

  function parse(items, width) {
    var full = toLines(items);
    var left = toLines(items.filter(function (i) { return i.x < width * 0.48; }));
    var out = { so: "", kyhieu: "", trichyeu: "", coquan: "", ngay: "" };

    // Số / ký hiệu (dùng cột trái để khỏi lẫn với địa danh bên phải)
    var soIdx = -1;
    for (var i = 0; i < left.length; i++) {
      var t = left[i].replace(/\s*-\s*/g, "-").replace(/\s*\/\s*/g, "/");
      var m = t.match(/S[ốo]\s*[:：]?\s*(\d+)\s*\/\s*(\S+)/i);
      if (m) {
        out.so = m[1];
        out.kyhieu = cleanName(m[2].replace(/[.,;]+$/, "").replace(/^\d{4}\//, ""));
        soIdx = i; break;
      }
    }

    // Ngày ban hành
    for (var j = 0; j < full.length; j++) {
      var d = full[j].match(/ng[àa]y\s+(\d{1,2})\s+th[áa]ng\s+(\d{1,2})\s+n[ăa]m\s+(\d{4})/i);
      if (d) { out.ngay = pad(d[1]) + "/" + pad(d[2]) + "/" + d[3]; break; }
    }

    // Cơ quan ban hành: các dòng chữ HOA ở cột trái, phía trên dòng "Số:"
    if (soIdx > 0) {
      var names = left.slice(0, soIdx).filter(function (l) {
        if (/^[-–—_.\s]+$/.test(l)) return false;
        if (/C[ỘO]NG H[ÒO]A|Đ[ỘO]C L[ẬA]P/i.test(l)) return false;
        var letters = l.replace(/[^A-Za-zÀ-ỹĐđ]/g, "");
        if (letters.length < 2) return false;
        var up = letters.replace(/[^A-ZÀ-ỸĐ]/g, "").length;
        return up / letters.length > 0.6;
      });
      out.coquan = names.length ? names[names.length - 1] : "";
    }

    // Trích yếu: dòng "V/v ..." hoặc "Về việc ..."
    for (var k = 0; k < full.length; k++) {
      var tm = full[k].match(/(?:^|\s)(?:V\s*\/\s*v|Về việc)\s*[:.]?\s*(.*)$/i);
      if (tm && (/^(V\s*\/\s*v|Về việc)/i.test(full[k]) || k < 25)) {
        var text = tm[1].trim();
        for (var n = k + 1; n < full.length && n <= k + 3; n++) {
          if (/^(Kính gửi|Căn cứ|Nơi nhận|Điều|Số\s*:|\d+\.|-)/i.test(full[n])) break;
          text += " " + full[n];
        }
        out.trichyeu = cleanName(text.replace(/\s+/g, " ").trim());
        break;
      }
    }
    return out;
  }

  // Ký tự không cho phép đặt tên file -> đổi thành dấu "."
  function cleanName(t) { return (t || "").replace(/[\\\/:*?"<>|]/g, "."); }

  function pad(n) { return ("0" + n).slice(-2); }

  // ---------- Lưu ----------
  function onSave() {
    pending = {};
    FIELDS.forEach(function (f) { pending[f] = $(f).value.trim(); });
    pending.kyhieu = cleanName(pending.kyhieu);
    pending.trichyeu = cleanName(pending.trichyeu);
    $("formView").style.display = "none";
    $("wait").style.display = "block";
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged, onSelect, function () {});
  }

  function stopWaiting() {
    pending = null;
    Office.context.document.removeHandlerAsync(
      Office.EventType.DocumentSelectionChanged, { handler: onSelect }, function () {});
    $("wait").style.display = "none";
    $("formView").style.display = "block";
  }

  function dateSerial(s) {
    var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (!m) return null;
    var dt = new Date(Date.UTC(+m[3], +m[2] - 1, +m[1]));
    if (dt.getUTCMonth() !== +m[2] - 1) return null;
    return dt.getTime() / 86400000 + 25569;
  }

  async function onSelect() {
    if (!pending) return;
    var p = pending;
    pending = null; // tránh ghi 2 lần
    try {
      await Excel.run(async function (ctx) {
        var sel = ctx.workbook.getSelectedRange();
        sel.load("rowIndex,columnIndex");
        var sh = ctx.workbook.worksheets.getActiveWorksheet();
        await ctx.sync();
        var target = sh.getRangeByIndexes(sel.rowIndex, sel.columnIndex, 1, 5);

        var soIsNum = /^[1-9]\d*$/.test(p.so);
        var ser = p.ngay ? dateSerial(p.ngay) : null;
        // Thứ tự: số, ký hiệu, ngày, cơ quan, trích yếu
        target.numberFormat = [[
          soIsNum ? "General" : "@", "@",
          ser !== null ? "dd/mm/yyyy" : "@", "@", "@"
        ]];
        target.values = [[
          p.so === "" ? "" : (soIsNum ? Number(p.so) : p.so),
          p.kyhieu,
          p.ngay === "" ? "" : (ser !== null ? ser : p.ngay),
          p.coquan,
          p.trichyeu
        ]];
        await ctx.sync();
        sh.getRangeByIndexes(sel.rowIndex, sel.columnIndex, 1, 1).select();
        await ctx.sync();
      });
      Office.context.document.removeHandlerAsync(
        Office.EventType.DocumentSelectionChanged, { handler: onSelect }, function () {});
      $("wait").style.display = "none";
      $("formView").style.display = "block";
      clearForm();
      setStatus("Đã lưu. Có thể chọn file công văn tiếp theo.", "ok");
    } catch (e) {
      pending = p;
      setStatus("Lỗi khi ghi vào Excel: " + (e && e.message ? e.message : e), "err");
    }
  }
})();
