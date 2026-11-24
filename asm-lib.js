/* THƯ VIỆN – KHO TỪ EXCEL (Đợt 1, theo mô hình KH-01): chọn khu -> tải file mẫu / up file -> XEM TRƯỚC (chưa ghi) -> lưu lô CHỜ DUYỆT -> DUYỆT -> ĐANG DÙNG.
 * Lô: Bổ sung (thêm dòng) hoặc Thay thế (lô đang dùng cùng khu chuyển "Đã thay"). Sửa tại ô, xóa dòng, ngừng/kích hoạt lại/xóa lô, tải lô về Excel. Mọi ghi qua hàm RPC chỉ ASM gọi được.
 * Đọc Excel bằng SheetJS ngay trên trình duyệt (lib-parse.js); máy chủ kiểm lại toàn bộ khi lưu. Dữ liệu thử là dữ liệu giả/mẫu. */
var ASMLIB = (function () {
  var P = window.LIBP, LS = { sel: '', parsed: null, tenFile: '', cheDo: '', lots: null, loading: false, lot: null, lotRows: null, busy: false };
  var TRANG = { CHO_DUYET: ['t-late', 'Chờ duyệt'], DANG_DUNG: ['t-ok', 'Đang dùng'], DA_THAY: ['t-new', 'Đã thay'], NGUNG_DUNG: ['t-bad', 'Ngừng dùng'], THIEU: ['t-bad', 'Thiếu'] };
  var SAU = [['Phác đồ vaccine', 'Anh gửi dữ liệu phác đồ sau, nút nạp sẽ làm như các khu khác (menu Phác đồ đã giữ chỗ).'], ['Mẫu tin quy trình', 'Các bước D, D+1, D+3, D+7 theo nhóm Lao/6in1. Làm ở đợt sau.'], ['Combo ưu đãi', 'Gói combo dạng bảng. Làm ở đợt sau.'], ['Ảnh, tờ rơi, voucher', 'Kho file/ảnh theo thư mục. Làm ở đợt sau.']];

  function xlsx(cb) { if (window.XLSX) return cb(); var s = document.createElement('script'); s.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'; s.onload = cb; s.onerror = function () { toast('Không tải được bộ đọc Excel. Kiểm tra mạng rồi thử lại.', 0, 1) }; document.head.appendChild(s) }
  function tag(c, t) { return '<span class="tag ' + c + '">' + esc(t) + '</span>' }
  function cut(v, n) { var s = v == null ? '' : String(v); return s.length > n ? s.slice(0, n) + '…' : s }
  function fmt(v) { return typeof v === 'number' ? v.toLocaleString('vi-VN') : v }
  function dt(s) { var d = loc(s); return d ? dstr(d) : '' }
  function cfg() { return P.CONFIG[LS.sel] }
  function loadLots() { if (LS.loading) return; LS.loading = true; return rest('lib_lots?select=*&order=thoi_gian.desc&limit=300').then(function (a) { LS.lots = a; LS.loading = false; if (LS.lot) { var n = a.filter(function (x) { return x.ma_lo === LS.lot.ma_lo })[0]; LS.lot = n || null; if (!n) LS.lotRows = null } render() }, function (e) { LS.loading = false; LS.lots = LS.lots || []; fail(e) }) }
  function openLot(ma) { var l = (LS.lots || []).filter(function (x) { return x.ma_lo === ma })[0]; if (!l) return; LS.lot = l; LS.lotRows = null; render(); rest('lib_rows?select=*&ma_lo=eq.' + encodeURIComponent(ma) + '&order=stt.asc&limit=1000').then(function (a) { LS.lotRows = a; render() }, fail) }
  function reload(keepLot) { LS.loading = false; return loadLots().then(function () { if (keepLot && LS.lot) openLot(LS.lot.ma_lo) }) }

  /* ---------- giao diện ---------- */
  function page() {
    if (LS.lots === null) return '<div class="card empty">Đang tải kho…</div>';
    var h = '<div class="card"><h2>1. Chọn khu vực <span class="pill">Mọi file đều qua Xem trước → Lưu chờ duyệt → Duyệt. Chưa duyệt thì nhân viên chưa thấy gì.</span></h2><div class="form2" style="grid-template-columns:repeat(auto-fill,minmax(210px,1fr))">';
    h += P.ORDER.map(function (k) {
      var c = P.CONFIG[k], dung = 0, nlo = 0; LS.lots.forEach(function (l) { if (l.loai === k) { nlo++; if (l.trang_thai === 'DANG_DUNG') dung += l.so_dong - l.so_dong_thieu } });
      return '<button class="sw" data-lbtile="' + k + '" style="display:block;text-align:left;' + (LS.sel === k ? 'border-color:var(--teal);box-shadow:0 0 0 3px rgba(29,135,135,.15)' : '') + '"><b>' + esc(c.ten) + '</b><small style="margin-top:4px">' + esc(c.moTa.slice(0, 90)) + (c.moTa.length > 90 ? '…' : '') + '</small><div style="margin-top:8px">' + tag(dung ? 't-ok' : 't-new', dung + ' dòng đang dùng') + tag('t-new', nlo + ' lô') + '</div></button>'
    }).join('');
    h += SAU.map(function (s) { return '<div class="sw" style="display:block;opacity:.6;cursor:default"><b>' + esc(s[0]) + '</b><small style="margin-top:4px">' + esc(s[1]) + '</small><div style="margin-top:8px">' + tag('t-new', 'Chưa làm') + '</div></div>' }).join('');
    h += '</div></div>';
    if (LS.sel) h += uploadCard();
    if (LS.parsed) h += previewCard();
    h += lotsCard();
    if (LS.lot) h += lotCard();
    return h
  }
  function uploadCard() {
    var c = cfg(), mode = LS.cheDo || c.cheDo;
    return '<div class="card"><h2>2. Tải tài liệu · ' + esc(c.ten) + '</h2><div class="meta" style="margin-bottom:10px">' + esc(c.moTa) + '</div><div class="row" style="flex-wrap:wrap"><button class="btn b3" id="lb-tpl">' + ic('download') + 'Tải file mẫu (.xlsx)</button><label class="btn b2" style="cursor:pointer">' + ic('upload') + 'Chọn file Excel/CSV<input type="file" id="lb-file" accept=".xlsx,.xls,.csv" hidden></label><span class="meta">Chế độ:</span><select id="lb-mode"><option value="BO_SUNG"' + (mode === 'BO_SUNG' ? ' selected' : '') + '>Bổ sung</option><option value="THAY_THE"' + (mode === 'THAY_THE' ? ' selected' : '') + '>Thay thế</option></select></div>' +
      '<div class="meta" style="margin-top:8px">Anh có thể up thẳng file gốc (hệ thống tự tìm dòng tiêu đề). <b>Bổ sung</b>: thêm dòng mới, giữ nguyên dòng đang dùng. <b>Thay thế</b>: khi Duyệt, lô đang dùng cùng khu chuyển sang "Đã thay" (vẫn lưu, không xóa). Biến trong tin tự đổi về dạng <code>{{TEN_BIEN}}</code>.</div></div>'
  }
  function previewCard() {
    var p = LS.parsed, c = cfg(); if (p.loi) return '<div class="card"><h2>3. Xem trước</h2><div class="err">' + esc(p.loi) + '</div><div class="acts"><button class="btn b3" id="lb-drop">Bỏ file này</button></div></div>';
    var warn = p.rows.filter(function (r) { return r._canh.length }).length;
    var h = '<div class="card"><h2>3. Xem trước và xác nhận <span class="pill">' + esc(LS.tenFile) + (p.sheet ? ' · sheet ' + esc(p.sheet) : '') + '</span></h2><div class="stats stats4"><div class="card st"><div class="l">Dòng đọc được</div><div class="v">' + p.tong + '</div></div><div class="card st"><div class="l">Hợp lệ</div><div class="v" style="color:var(--ok)">' + p.hopLe + '</div></div><div class="card st"><div class="l">Thiếu / sai</div><div class="v" style="color:' + (p.thieu ? 'var(--bad)' : 'inherit') + '">' + p.thieu + '</div></div><div class="card st"><div class="l">Có cảnh báo</div><div class="v" style="color:' + (warn ? 'var(--warn)' : 'inherit') + '">' + warn + '</div></div></div>';
    if (p.ky && p.ky.tu) h += '<div class="alert"><i>📅</i><div><b>Kỳ áp dụng đọc từ đầu file: ' + esc(p.ky.tu) + ' → ' + esc(p.ky.den) + '</b>' + (p.ky.ten ? '<small>' + esc(p.ky.ten) + '</small>' : '') + '</div></div>';
    h += '<div class="meta" style="margin:6px 0">Dòng tiêu đề tìm thấy ở dòng ' + p.dongTieuDe + '. Dòng thiếu thông tin vẫn được lưu nhưng KHÔNG đưa vào dùng khi Duyệt, sửa tại ô sau khi lưu.</div>';
    if (p.boQua.length) h += '<details class="alert a-warn" style="display:block"><summary><b>Hệ thống không đưa vào ' + p.boQua.length + ' dòng</b></summary><div class="meta" style="margin-top:6px">' + p.boQua.slice(0, 20).map(function (b) { return 'Dòng ' + b.dong + ': ' + esc(b.ly_do) }).join('<br>') + '</div></details>';
    h += '<div class="tw"><table><thead><tr><th>Dòng</th><th>Kết quả</th>' + c.cot.map(function (x) { return '<th>' + esc(x.nhan) + '</th>' }).join('') + '</tr></thead><tbody>' + p.rows.slice(0, 50).map(function (r) {
      return '<tr><td>' + r._dong + '</td><td>' + (r._thieu.length ? tag('t-bad', r._thieu.join('; ')) : r._canh.length ? tag('t-late', r._canh.join('; ')) : tag('t-ok', 'Hợp lệ')) + '</td>' + c.cot.map(function (x) { return '<td title="' + esc(r[x.k]) + '">' + esc(cut(fmt(r[x.k]), 60)) + '</td>' }).join('') + '</tr>'
    }).join('') + '</tbody></table></div>' + (p.rows.length > 50 ? '<div class="meta" style="padding:6px">Hiện 50 dòng đầu trong ' + p.rows.length + ' dòng.</div>' : '');
    h += '<div class="acts" style="margin-top:10px"><button class="btn b1" id="lb-save"' + (p.tong ? '' : ' disabled') + '>Lưu ' + p.tong + ' dòng vào lô chờ duyệt</button><button class="btn b3" id="lb-drop">Bỏ file này</button></div></div>';
    return h
  }
  function lotsCard() {
    return '<div class="card" style="padding:6px 8px"><div class="row2" style="padding:10px 10px 0"><h2 style="margin:0">Các lô đã tải</h2><button class="btn b3" id="lb-refresh" style="min-height:38px">Làm mới</button></div><div class="tw"><table><thead><tr><th>Mã lô</th><th>Khu</th><th>File</th><th>Chế độ</th><th>Dòng</th><th>Thiếu</th><th>Trạng thái</th><th>Tải lúc</th><th></th></tr></thead><tbody>' + LS.lots.map(function (l) {
      var st = TRANG[l.trang_thai] || ['t-new', l.trang_thai]; return '<tr' + (LS.lot && LS.lot.ma_lo === l.ma_lo ? ' style="background:var(--tint)"' : '') + '><td><b>' + esc(l.ma_lo) + '</b></td><td>' + esc((P.CONFIG[l.loai] || {}).ten || l.loai) + '</td><td>' + esc(cut(l.ten_file, 40)) + '</td><td>' + (l.che_do === 'THAY_THE' ? 'Thay thế' : 'Bổ sung') + '</td><td>' + l.so_dong + '</td><td>' + (l.so_dong_thieu ? tag('t-bad', l.so_dong_thieu) : '0') + '</td><td>' + tag(st[0], st[1]) + '</td><td>' + esc(dt(l.thoi_gian)) + '</td><td><button class="btn b3" data-lblot="' + esc(l.ma_lo) + '" style="min-height:36px;font-size:13px">Xem</button></td></tr>'
    }).join('') + (LS.lots.length ? '' : '<tr><td colspan="9" class="empty">Chưa có lô nào. Chọn một khu ở trên, tải file mẫu rồi up file.</td></tr>') + '</tbody></table></div></div>'
  }
  function lotCard() {
    var l = LS.lot, c = P.CONFIG[l.loai], st = TRANG[l.trang_thai] || ['t-new', l.trang_thai], canEdit = l.trang_thai === 'CHO_DUYET' || l.trang_thai === 'DANG_DUNG';
    var btn = function (k, t, cls) { return '<button class="btn ' + (cls || 'b3') + '" data-lbact="' + k + '" style="min-height:40px">' + t + '</button>' };
    var h = '<div class="card"><div class="row2"><div><h2 style="margin:0">Lô ' + esc(l.ma_lo) + ' · ' + esc(c.ten) + ' ' + tag(st[0], st[1]) + '</h2><div class="meta">' + esc(l.ten_file || '') + ' · ' + (l.che_do === 'THAY_THE' ? 'Thay thế' : 'Bổ sung') + ' · ' + l.so_dong + ' dòng, thiếu ' + l.so_dong_thieu + (l.thoi_gian_duyet ? ' · duyệt ' + esc(dt(l.thoi_gian_duyet)) : '') + (l.ky && l.ky.tu ? ' · kỳ ' + esc(l.ky.tu) + ' → ' + esc(l.ky.den) : '') + '</div></div><button class="btn b3" id="lb-close" style="min-height:38px">Đóng</button></div>';
    h += '<div class="acts" style="margin:10px 0;flex-wrap:wrap">' + (l.trang_thai === 'CHO_DUYET' ? btn('approve', 'Duyệt', 'b1') : '') + (canEdit ? btn('stop', 'Ngừng dùng') : '') + (l.trang_thai === 'NGUNG_DUNG' ? btn('react', 'Kích hoạt lại') + btn('del', 'Xóa lô') : '') + btn('xl', ic('download') + 'Tải về Excel (.xlsx)') + '</div>';
    if (l.trang_thai === 'CHO_DUYET') h += '<div class="ro">Lô chờ duyệt: nhân viên chưa thấy. ' + (l.che_do === 'THAY_THE' ? 'Khi Duyệt, các lô đang dùng cùng khu sẽ chuyển "Đã thay".' : 'Khi Duyệt, các dòng hợp lệ được thêm vào dùng, dòng đang dùng giữ nguyên.') + '</div>';
    if (canEdit) h += '<div class="meta" style="margin-bottom:6px">Bấm vào một ô để sửa (Enter = lưu, Esc = hủy). ✕ = xóa dòng.</div>';
    if (LS.lotRows === null) return h + '<div class="empty">Đang tải dòng…</div></div>';
    h += '<div class="tw" style="max-height:480px;overflow:auto"><table><thead><tr><th>#</th><th>Trạng thái</th>' + c.cot.map(function (x) { return '<th>' + esc(x.nhan) + '</th>' }).join('') + '<th>Cảnh báo</th>' + (canEdit ? '<th></th>' : '') + '</tr></thead><tbody>' + LS.lotRows.slice(0, 300).map(function (r) {
      var s = TRANG[r.trang_thai] || ['t-new', r.trang_thai];
      return '<tr' + (r.trang_thai === 'THIEU' ? ' style="background:var(--bad-bg)"' : '') + '><td>' + r.stt + '</td><td>' + tag(s[0], s[1]) + '</td>' + c.cot.map(function (x) { var v = r.data[x.k]; return '<td' + (canEdit ? ' class="sua" data-id="' + r.id + '" data-col="' + x.k + '" style="cursor:text"' : '') + ' title="' + esc(v) + '">' + esc(cut(fmt(v), 70)) + '</td>' }).join('') + '<td class="meta">' + esc(r.canh_bao || '') + '</td>' + (canEdit ? '<td><button class="btn b3" data-lbdel="' + r.id + '" style="min-height:32px;padding:0 10px" aria-label="Xóa dòng">✕</button></td>' : '') + '</tr>'
    }).join('') + '</tbody></table></div>' + (LS.lotRows.length > 300 ? '<div class="meta" style="padding:6px">Hiện 300 dòng đầu trong ' + LS.lotRows.length + ' dòng (tải Excel để xem đủ).</div>' : '') + '</div>';
    return h
  }

  /* ---------- hành động ---------- */
  function handle(name, buf) {
    xlsx(function () {
      try { var wb = XLSX.read(buf, { type: 'array', cellDates: true, codepage: 65001 }); LS.parsed = P.fromWorkbook(XLSX, wb, LS.sel); LS.tenFile = name; render() } catch (e) { toast('Không đọc được file: ' + e.message, 0, 1) }
    })
  }
  function template() { xlsx(function () { var t = P.templateAoa(LS.sel), wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(t.du_lieu), 'DU_LIEU'); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(t.huong_dan), 'HUONG_DAN'); XLSX.writeFile(wb, 'FILE_MAU_' + LS.sel + '.xlsx') }) }
  function save() {
    var p = LS.parsed; if (!p || !p.rows || LS.busy) return; LS.busy = true; var b = document.getElementById('lb-save'); if (b) b.disabled = true;
    var rows = p.rows.map(function (r) { var o = {}; Object.keys(r).forEach(function (k) { if (k.charAt(0) !== '_') o[k] = r[k] }); return o });
    rpc('u_lib_save_lot', { p_loai: LS.sel, p_ten_file: LS.tenFile, p_che_do: LS.cheDo || cfg().cheDo, p_ky: p.ky, p_rows: rows, p_req: rid() }).then(function (r) { LS.busy = false; LS.parsed = null; toast('Đã lưu lô ' + r.ma_lo + ': ' + r.tong + ' dòng, thiếu ' + r.thieu + '. Đang chờ duyệt.'); return loadLots().then(function () { openLot(r.ma_lo) }) }, function (e) { LS.busy = false; if (b) b.disabled = false; fail(e) })
  }
  function act(k) {
    var l = LS.lot; if (!l) return; var done = function (msg, closeLot) { return function () { toast(msg); if (closeLot) { LS.lot = null; LS.lotRows = null } return reload(!closeLot) } };
    if (k === 'xl') return exportLot();
    if (k === 'approve') { if (!confirm('Duyệt lô ' + l.ma_lo + '?' + (l.che_do === 'THAY_THE' ? '\nCác lô đang dùng cùng khu sẽ chuyển sang "Đã thay".' : '') + '\nCác dòng hợp lệ sẽ hiện cho nhân viên.')) return; rpc('u_lib_approve', { p_ma_lo: l.ma_lo, p_req: rid() }).then(function (r) { return done('Đã duyệt: ' + r.dung + ' dòng đưa vào sử dụng' + (r.thieu ? ', còn ' + r.thieu + ' dòng thiếu chờ sửa' : ''))() }, fail) }
    if (k === 'stop') { if (!confirm('Ngừng dùng lô ' + l.ma_lo + '? Nhân viên sẽ không thấy các dòng này nữa (dữ liệu vẫn giữ).')) return; rpc('u_lib_stop', { p_ma_lo: l.ma_lo, p_req: rid() }).then(done('Đã ngừng dùng lô'), fail) }
    if (k === 'react') rpc('u_lib_reactivate', { p_ma_lo: l.ma_lo, p_req: rid() }).then(done('Đã đưa lô về chờ duyệt'), fail);
    if (k === 'del') { if (!confirm('Xóa hẳn lô ' + l.ma_lo + ' và ' + l.so_dong + ' dòng? Không hoàn tác được (nhật ký vẫn giữ).')) return; rpc('u_lib_delete', { p_ma_lo: l.ma_lo, p_req: rid() }).then(done('Đã xóa lô', true), fail) }
  }
  function exportLot() {
    var l = LS.lot, c = P.CONFIG[l.loai]; xlsx(function () { var aoa = [c.cot.map(function (x) { return x.nhan })].concat((LS.lotRows || []).map(function (r) { return c.cot.map(function (x) { var v = r.data[x.k]; return v == null ? '' : v }) })); var wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), 'DU_LIEU'); XLSX.writeFile(wb, l.ma_lo + '_' + l.loai + '.xlsx') })
  }
  function editCell(td) {
    if (td.querySelector('input,textarea')) return; var id = Number(td.dataset.id), col = td.dataset.col, r = LS.lotRows.filter(function (x) { return x.id === id })[0]; if (!r) return;
    var cur = r.data[col] == null ? '' : String(r.data[col]), long = cur.length > 50 || /\n/.test(cur), el = document.createElement(long ? 'textarea' : 'input'); el.value = cur; el.style.cssText = 'width:100%;min-width:' + (long ? 280 : 120) + 'px;min-height:' + (long ? 120 : 36) + 'px;border:1px solid var(--teal);border-radius:8px;padding:6px;background:var(--card);color:var(--ink)';
    td.innerHTML = ''; td.appendChild(el); el.focus(); var finished = false;
    function end(save) { if (finished) return; finished = true; if (!save || el.value === cur) { render(); return } rpc('u_lib_edit_row', { p_id: id, p_col: col, p_val: el.value, p_req: rid() }).then(function () { toast('Đã sửa ô'); return reload(true) }, function (e) { fail(e); render() }) }
    el.onkeydown = function (e) { if (e.key === 'Escape') end(false); else if (e.key === 'Enter' && !(long && e.shiftKey)) { e.preventDefault(); end(true) } }; el.onblur = function () { end(true) }
  }
  function bind(on) {
    xlsx(function () { }); if (LS.lots === null && !LS.loading) loadLots();
    [].slice.call(document.querySelectorAll('[data-lbtile]')).forEach(function (b) { b.onclick = function () { LS.sel = b.dataset.lbtile; LS.parsed = null; LS.cheDo = P.CONFIG[LS.sel].cheDo; render() } });
    on('#lb-tpl', 'onclick', template); on('#lb-mode', 'onchange', function (x) { LS.cheDo = x.value });
    on('#lb-file', 'onchange', function (x) { var f = x.files[0]; if (!f) return; var rd = new FileReader(); rd.onload = function () { handle(f.name, rd.result) }; rd.readAsArrayBuffer(f) });
    on('#lb-save', 'onclick', save); on('#lb-drop', 'onclick', function () { LS.parsed = null; render() }); on('#lb-refresh', 'onclick', function () { reload(true) }); on('#lb-close', 'onclick', function () { LS.lot = null; LS.lotRows = null; render() });
    [].slice.call(document.querySelectorAll('[data-lblot]')).forEach(function (b) { b.onclick = function () { openLot(b.dataset.lblot) } });
    [].slice.call(document.querySelectorAll('[data-lbact]')).forEach(function (b) { b.onclick = function () { act(b.dataset.lbact) } });
    [].slice.call(document.querySelectorAll('td.sua')).forEach(function (td) { td.onclick = function () { editCell(td) } });
    [].slice.call(document.querySelectorAll('[data-lbdel]')).forEach(function (b) { b.onclick = function () { if (!confirm('Xóa hẳn dòng này khỏi lô?')) return; rpc('u_lib_del_row', { p_id: Number(b.dataset.lbdel), p_req: rid() }).then(function () { toast('Đã xóa dòng'); return reload(true) }, fail) } })
  }
  return { page: page, bind: bind, handle: handle, state: LS, select: function (k) { LS.sel = k; LS.cheDo = P.CONFIG[k].cheDo; LS.parsed = null } }
})();
