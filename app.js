(function () {
  // ===== Cấu hình =====
  var CONFIG = {
    agencyTitleCase: true,   // đổi tên cơ quan IN HOA -> Viết Hoa Chữ Cái Đầu
    includeLoaiVanBan: false // true: thêm tên loại (THÔNG BÁO...) vào đầu trích yếu khi văn bản không có "V/v"
  };

  var $ = function (id) { return document.getElementById(id); };
  var FIELDS = ["so", "kyhieu", "ngay", "coquan", "trichyeu"];
  var pending = null;

  // ---------- Khởi tạo ----------
  if (typeof Office !== "undefined") {
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
      document.addEventListener("dragover", function (e) { e.preventDefault(); });
      document.addEventListener("drop", function (e) { e.preventDefault(); });
      $("btnCancel").addEventListener("click", function () { clearForm(); setStatus("", ""); $("dbg").textContent = ""; });
      $("btnSave").addEventListener("click", onSave);
      $("btnBack").addEventListener("click", function () { stopWaiting(); });
    });
  }

  function setStatus(msg, cls) { var s = $("status"); s.textContent = msg; s.className = cls || ""; }
  function clearForm() { FIELDS.forEach(function (f) { $(f).value = ""; }); }

  // ---------- Đọc PDF ----------
  async function getBoldMap(page, tc) {
    var map = {};
    try {
      await page.getOperatorList();
      Object.keys(tc.styles || {}).forEach(function (id) {
        try {
          var f = page.commonObjs.get(id);
          var name = (f && (f.name || f.fallbackName)) || "";
          map[id] = !!(f && (f.bold || f.black)) || /bold|black|heavy|semibold|[-,]b$/i.test(name);
        } catch (e) { /* font chưa nạp */ }
      });
    } catch (e) { /* bỏ qua */ }
    return map;
  }

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
      var boldMap = await getBoldMap(page, tc);
      var items = tc.items
        .filter(function (i) { return i.str && i.str.trim(); })
        .map(function (i) {
          return { s: i.str, x: i.transform[4], y: i.transform[5], w: i.width, h: i.height, b: !!boldMap[i.fontName] };
        });
      if (!items.length) {
        clearForm();
        setStatus("Không đọc được chữ trong file (có thể là PDF scan). Hãy điền tay.", "warn");
        return;
      }
      var r = parse(items, width);
      FIELDS.forEach(function (f) { $(f).value = r[f] || ""; });
      $("dbg").textContent = r.debug;
      var empty = FIELDS.filter(function (f) { return !r[f]; }).length;
      setStatus(empty ? "Đã điền tự động, còn " + empty + " ô trống - vui lòng kiểm tra." :
        "Đã điền đủ các ô - vui lòng kiểm tra lại trước khi lưu.", empty ? "warn" : "ok");
    } catch (e) {
      setStatus("Lỗi đọc file: " + (e && e.message ? e.message : e), "err");
    }
  }

  // ---------- Hàm hỗ trợ nhận diện ----------
  function squeeze(t) { return t.replace(/\s+/g, ""); }
  // bỏ khoảng trắng + bỏ dấu tiếng Việt + chữ thường  (để so khớp "chịu lỗi")
  function plain(t) {
    return squeeze(t).normalize("NFD").replace(/[\u0300-\u036f]/g, "")
      .replace(/đ/g, "d").replace(/Đ/g, "D").toLowerCase();
  }
  function pad(n) { return ("0" + n).slice(-2); }
  // Ký tự không cho phép đặt tên file -> dấu "."
  function cleanName(t) { return (t || "").replace(/[\\\/:*?"<>|]/g, "."); }

  function toLines(items) {
    var sorted = items.slice().sort(function (a, b) { return b.y - a.y || a.x - b.x; });
    var clusters = [], cur = null;
    sorted.forEach(function (it) {
      if (!cur || Math.abs(cur.y - it.y) > 3) { cur = { y: it.y, parts: [] }; clusters.push(cur); }
      cur.parts.push(it);
    });
    return clusters.map(function (l) {
      l.parts.sort(function (a, b) { return a.x - b.x; });
      var t = "", prev = null, total = 0, boldN = 0;
      l.parts.forEach(function (p) {
        var s = p.s;
        if (prev) {
          var gap = p.x - (prev.x + prev.w);
          var combining = /^[\u0300-\u036f]/.test(s);
          if (!combining && !/\s$/.test(prev.s) && !/^\s/.test(s) && gap > Math.max(1, 0.1 * (p.h || 0))) t += " ";
        }
        t += s; prev = p;
        var n = s.replace(/\s/g, "").length; total += n; if (p.b) boldN += n;
      });
      return {
        text: t.replace(/\s+/g, " ").trim().normalize("NFC"),
        bold: total > 0 && boldN / total > 0.5,
        y: l.y, x: l.parts[0].x
      };
    }).filter(function (l) { return l.text; });
  }

  function isStop(text) {
    var p = plain(text);
    return /^(kinhgui|cancu|noinhan|kinhchuyen|dieu\d|\d+[.)]|-)/.test(p);
  }

  function sentenceCase(s) {
    if (s === s.toUpperCase() && /[A-ZÀ-Ỹ]/.test(s)) {
      s = s.toLowerCase();
      return s.charAt(0).toUpperCase() + s.slice(1);
    }
    return s;
  }

  var ACRONYMS = ["UBND", "HĐND", "MTTQ", "UBMTTQ", "TNHH", "CP", "TP", "TW", "BQL", "VP", "TCT", "NHNN", "BHXH", "VNPT", "EVN", "PCCC", "HTX"];
  var SMALL = { "ty": 1, "của": 1, "và": 1, "về": 1, "cho": 1, "tại": 1, "trên": 1, "trong": 1, "với": 1 };
  function titleCaseVN(s) {
    if (s !== s.toUpperCase()) return s; // đã có chữ thường -> giữ nguyên
    return s.toLowerCase().split(" ").map(function (w, i) {
      if (ACRONYMS.indexOf(w.toUpperCase()) >= 0) return w.toUpperCase();
      if (i > 0 && SMALL[w]) return w;
      return w.charAt(0).toUpperCase() + w.slice(1);
    }).join(" ");
  }

  // ---------- Vùng đầu trang ----------
  var LOAI = ["thongbao", "quyetdinh", "baocao", "kehoach", "totrinh", "chithi", "nghiquyet", "huongdan",
    "giaymoi", "thumoi", "congdien", "chuongtrinh", "dean", "phuongan", "quyche", "quydinh", "bienban",
    "hopdong", "thoathuan", "thongcao", "ketluan", "thongtri", "giaygioithieu", "thungo", "nghidinh"];

  function isVv(p) { return /^(v\/v|veviec)/.test(p); }
  function isLoai(p) { return LOAI.indexOf(p.replace(/[:.]+$/, "")) >= 0; }

  function isAllUpper(t) {
    var letters = t.replace(/[^A-Za-zÀ-ỹĐđ]/g, "");
    return letters.length >= 3 && letters === letters.toUpperCase();
  }

  // Phần đầu trang kết thúc tại dòng đầu tiên là: tên loại (QUYẾT ĐỊNH...), "V/v", "Kính gửi" hoặc "Căn cứ".
  // Số, ngày, cơ quan CHỈ được tìm phía trên vị trí này -> không bao giờ lấy nhầm từ thân văn bản.
  function headerEndY(full) {
    var lim = Math.min(full.length, 30);
    for (var k = 0; k < lim; k++) {
      var p = plain(full[k].text);
      if (/^(kinhgui|cancu)/.test(p) || isVv(p) || isLoai(p)) return full[k].y;
    }
    return full.length > 12 ? full[12].y : -1e9;
  }
  function zone(lines, endY) { return lines.filter(function (l) { return l.y > endY + 1; }); }

  // ---------- Số / ký hiệu ----------
  function fixKy(s) { return s.replace(/[.,;:]+$/, "").replace(/^\d{4}\//, ""); }

  // Cho phép số bị bỏ trống: "Số: /QĐ-HVTP"
  function findSo(left, full) {
    var i, t, m;
    for (i = 0; i < left.length; i++) {
      if (/^so[:;.]?[.…_]*\d*[.…_]*\//.test(plain(left[i].text))) {
        var sq = squeeze(left[i].text).normalize("NFC");
        m = sq.match(/^S[^\d\/]{0,3}?[:：]?[.…_]*(\d*)[.…_]*\/(.+)$/i);
        if (m) return { so: m[1], kyhieu: fixKy(m[2]), y: left[i].y };
      }
    }
    for (i = 0; i < full.length; i++) {
      t = full[i].text.replace(/\s*-\s*/g, "-").replace(/\s*\/\s*/g, "/");
      m = t.match(/^S[^\s\d\/]{0,3}\s*[:：]?\s*[.…_]*(\d*)[.…_]*\s*\/([^\s,]+)/i);
      if (m) return { so: m[1], kyhieu: fixKy(m[2]), y: full[i].y };
    }
    return null;
  }

  // ---------- Ngày ban hành ----------
  function validDate(d, m) { return +d >= 1 && +d <= 31 && +m >= 1 && +m <= 12; }

  // Chỉ lấy khi có đủ ngày + tháng + năm; thiếu thì để trống
  function findDate(full, right) {
    var i, m, p;
    for (i = 0; i < full.length; i++) {
      p = plain(full[i].text);
      m = p.match(/ngay(\d{1,2})thang(\d{1,2})nam(\d{4})/);
      if (m && validDate(m[1], m[2])) return pad(m[1]) + "/" + pad(m[2]) + "/" + m[3];
    }
    for (i = 0; i < right.length; i++) {
      m = right[i].text.match(/(\d{1,2})\D{1,12}?(\d{1,2})\D{1,12}?((?:19|20)\d{2})/);
      if (m && validDate(m[1], m[2])) return pad(m[1]) + "/" + pad(m[2]) + "/" + m[3];
    }
    for (i = 0; i < full.length; i++) {
      m = full[i].text.match(/(\d{1,2})[\/.\-](\d{1,2})[\/.\-]((?:19|20)\d{2})/);
      if (m && validDate(m[1], m[2])) return pad(m[1]) + "/" + pad(m[2]) + "/" + m[3];
    }
    return "";
  }

  // ---------- Cơ quan ban hành ----------
  var PARENT = /^(BỘ|UBND|ỦY BAN|CHÍNH PHỦ|QUỐC HỘI|TỈNH|THÀNH PHỐ|THÀNH ỦY|TỈNH ỦY|HUYỆN ỦY|QUẬN ỦY|HUYỆN|QUẬN|THỊ XÃ|SỞ|TỔNG LIÊN ĐOÀN)(\s|$)/;

  function findAgency(left, soY) {
    var cand = left.filter(function (l) {
      if (soY !== null && l.y <= soY + 3) return false;
      if (/^[-–—_.\s]+$/.test(l.text)) return false;
      if (/C[ỘO]NG H[ÒO]A|Đ[ỘO]C L[ẬA]P/i.test(l.text)) return false;
      if (/^so[:;.]?[.…_]*\d*[.…_]*\//.test(plain(l.text))) return false;
      if (soY === null && !isAllUpper(l.text)) return false;
      return l.text.replace(/[^A-Za-zÀ-ỹĐđ]/g, "").length >= 2;
    });
    if (!cand.length) return "";
    var hasBold = cand.some(function (l) { return l.bold; });
    var hasPlain = cand.some(function (l) { return !l.bold; });
    var chosen;
    if (cand.length >= 2 && hasBold && hasPlain) {
      chosen = cand.filter(function (l) { return l.bold; });
    } else if (cand.length >= 2 && PARENT.test(cand[0].text.toUpperCase())) {
      chosen = cand.slice(1);
    } else {
      chosen = cand;
    }
    var name = chosen.map(function (l) { return l.text; }).join(" ").replace(/\s+/g, " ").trim();
    return CONFIG.agencyTitleCase ? titleCaseVN(name) : name;
  }

  // ---------- Trích yếu ----------
  // mode "vv": dừng ở dòng IN HOA (vd "GIÁM ĐỐC HỌC VIỆN...") hoặc "Căn cứ/Kính gửi"
  // mode "loai": dòng đầu được IN HOA, nhưng dừng khi chuyển từ chữ thường sang IN HOA
  function collect(full, from, max, mode) {
    var out = [], prevUpper = null;
    for (var n = from; n < full.length && out.length < max; n++) {
      var t = full[n].text;
      if (isStop(t)) break;
      var up = isAllUpper(t);
      if (mode === "vv" && up) break;
      if (mode === "loai" && up && prevUpper === false) break;
      out.push(t); prevUpper = up;
    }
    return out;
  }

  function capFirst(s) { return s ? s.charAt(0).toUpperCase() + s.slice(1) : s; }

  function takeVv(full, k) {
    var orig = full[k].text;
    var text = orig.replace(/^\s*(V\s*\/\s*v|Về\s*việc)\s*[:.\-]?\s*/i, "");
    if (text === orig) text = orig.replace(/^.{3,8}?[:.]\s*/, "");
    var rest = collect(full, k + 1, 5, "vv");
    return capFirst((text + " " + rest.join(" ")).replace(/\s+/g, " ").trim());
  }

  function findTrichYeu(full) {
    var lim = Math.min(full.length, 30);
    for (var k = 0; k < lim; k++) {
      var p = plain(full[k].text);
      if (isVv(p)) return takeVv(full, k);
      if (isLoai(p)) {
        if (k + 1 < full.length && isVv(plain(full[k + 1].text))) return takeVv(full, k + 1);
        var lines = collect(full, k + 1, 6, "loai").map(sentenceCase);
        if (!lines.length) return "";
        var res = lines.join(" ");
        if (CONFIG.includeLoaiVanBan) res = sentenceCase(full[k].text) + " " + res.charAt(0).toLowerCase() + res.slice(1);
        return res.replace(/\s+/g, " ").trim();
      }
      if (/^(kinhgui|cancu)/.test(p)) return "";
    }
    return "";
  }

  // ---------- Hàm chính ----------
  function parse(items, width) {
    var full = toLines(items);
    var left = toLines(items.filter(function (i) { return i.x < width * 0.48; }));
    var right = toLines(items.filter(function (i) { return i.x >= width * 0.45; }));
    var endY = headerEndY(full);
    var fz = zone(full, endY), lz = zone(left, endY), rz = zone(right, endY);
    var out = { so: "", kyhieu: "", ngay: "", coquan: "", trichyeu: "" };

    var so = findSo(lz, fz);
    if (so) { out.so = so.so; out.kyhieu = cleanName(so.kyhieu); }
    out.ngay = findDate(fz, rz);
    out.coquan = findAgency(lz, so ? so.y : null);
    out.trichyeu = cleanName(findTrichYeu(full));

    out.debug = "Văn bản đọc được (B = chữ đậm, ---- = hết phần đầu trang):\n" + full.slice(0, 30).map(function (l, i) {
      var mark = (l.y > endY + 1) ? "" : "";
      return (i + 1) + (l.bold ? " [B] " : "     ") + l.text + (l.y === endY ? "   <---- hết phần đầu" : "");
    }).join("\n");
    return out;
  }

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
    pending = null;
    try {
      await Excel.run(async function (ctx) {
        var sel = ctx.workbook.getSelectedRange();
        sel.load("rowIndex,columnIndex");
        var sh = ctx.workbook.worksheets.getActiveWorksheet();
        await ctx.sync();
        var target = sh.getRangeByIndexes(sel.rowIndex, sel.columnIndex, 1, 5);
        var soIsNum = /^[1-9]\d*$/.test(p.so);
        var ser = p.ngay ? dateSerial(p.ngay) : null;
        // Thứ tự ghi: số, ký hiệu, ngày, cơ quan, trích yếu
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
      $("dbg").textContent = "";
      setStatus("Đã lưu. Có thể chọn file công văn tiếp theo.", "ok");
    } catch (e) {
      pending = p;
      setStatus("Lỗi khi ghi vào Excel: " + (e && e.message ? e.message : e), "err");
    }
  }

  if (typeof module !== "undefined") module.exports = { parse: parse };
})();
