/* GỢI Ý TIN NHẮN + LỜI NÓI KHI GỌI – dùng chung cho Nhân viên và Quản lý (Quản lý chỉ soạn cho lượt của chính mình).
 * Nội dung lấy từ THƯ VIỆN (msg_templates, promos) do ASM duyệt; "vaccine đã đến tuổi" lấy từ PHÁC ĐỒ (vaccine_rules) theo ngày sinh.
 * Ba nút gửi, mỗi nút làm đúng việc của nó rồi tự ghi 1 lần "Nhắn" (hàm u_msg_log):
 *   Zalo     = chép tin + mở zalo.me/<số> (Zalo không cho web điền sẵn tin hay tự kết bạn: nhân viên dán tin rồi gửi/kết bạn)
 *   SMS      = mở app tin nhắn tới số, nội dung điền sẵn (đồng thời chép tin phòng khi máy tính không có app SMS)
 *   Sao chép = chỉ chép tin
 * Dữ liệu và câu chữ mẫu hiện là DỮ LIỆU GIẢ để thử. */
var MSG = (function () {
  var CSS = '.person{display:grid;grid-template-columns:1fr auto;gap:10px;align-items:center;padding:12px 0;border-top:1px solid var(--line)}.person:first-child{border-top:0}.person .phone{font-size:20px}.btn.bo{border:1.5px solid var(--teal);color:var(--teal-d);background:var(--card);min-height:44px}' +
    '.modal.mm{align-items:end}.msheet{width:100%;max-width:560px;max-height:94vh;display:flex;flex-direction:column;background:var(--bg);border-radius:24px 24px 0 0;box-shadow:0 -10px 40px rgba(8,42,62,.25)}' +
    '@media(min-width:960px){.modal.mm{align-items:center}.msheet{border-radius:24px;max-height:90vh}}' +
    '.msh{display:flex;align-items:center;gap:10px;padding:14px 16px 8px}.msh .x{margin-left:auto;border:1px solid var(--line);background:var(--card);border-radius:12px;width:40px;height:40px;display:grid;place-items:center}' +
    '.msb{overflow:auto;padding:4px 16px 12px;display:grid;gap:14px}.msf{padding:10px 16px calc(12px + env(safe-area-inset-bottom));border-top:1px solid var(--line);background:var(--card);display:grid;gap:10px;border-radius:0 0 24px 24px}' +
    '.msheet h3{margin:0 0 8px;font-size:12.5px;letter-spacing:.05em;text-transform:uppercase;color:var(--soft)}.msheet .chips{margin-bottom:0}.chip.star:after{content:" ★";color:var(--orange)}.chip.on.star:after{color:#ffd29a}' +
    '.seg{display:inline-flex;background:var(--new-bg);border-radius:12px;padding:3px}.seg button{border:0;background:none;padding:8px 14px;border-radius:10px;font-weight:600;min-height:38px}.seg button.on{background:var(--card);box-shadow:var(--shadow)}' +
    '.msheet textarea{min-height:200px;line-height:1.55;resize:vertical}.badge{font-size:12.5px;font-weight:600;border-radius:99px;padding:3px 10px}.b-ok{background:var(--ok-bg);color:var(--ok)}.b-ed{background:var(--warn-bg);color:var(--warn)}' +
    '.alt{border:1px solid var(--line);background:var(--card);border-radius:14px;padding:10px 12px;text-align:left;font-size:14px;color:var(--muted);width:100%}.alt b{display:block;color:var(--ink)}.msheet select{border:1px solid var(--line);border-radius:12px;padding:8px 10px;background:var(--card);min-height:40px}' +
    '.vxl{display:grid;gap:6px;margin-top:8px}.vxl label{display:flex;gap:10px;align-items:flex-start;cursor:pointer;font-size:14.5px}.vxl input{width:20px;height:20px;accent-color:var(--teal);flex:none;margin-top:1px}.vxl small{display:block;color:var(--muted)}' +
    '.sendrow{display:grid;grid-template-columns:1.1fr 1fr 1.1fr;gap:8px}.sendrow .btn{min-height:56px;flex-direction:column;gap:0;line-height:1.2;padding:0 6px}.sendrow .btn span{white-space:nowrap}.sendrow small{font-weight:500;opacity:.85;font-size:12px}';
  (function () { try { var s = document.createElement('style'); s.textContent = CSS; document.head.appendChild(s) } catch (e) {} })();
  IC.x = '<path d="M6 6l12 12M18 6 6 18"/>';

  var VARS = ['XUNG_HO', 'XUNG_HO_HOA', 'TEN_NGUOI_LH', 'TEN_BE', 'TUOI', 'TEN_NV', 'TEN_CO_SO', 'DIA_CHI', 'GIO_MO_CUA', 'HOTLINE'];
  var VARDESC = { XUNG_HO: 'mẹ / bố / anh chị…', XUNG_HO_HOA: 'Mẹ / Bố… (đầu câu)', TEN_NGUOI_LH: 'tên người liên hệ', TEN_BE: 'tên bé', TUOI: 'tuổi bé (ngày/tháng/năm)', TEN_NV: 'tên nhân viên', TEN_CO_SO: 'tên Shop', DIA_CHI: 'địa chỉ Shop', GIO_MO_CUA: 'giờ mở cửa', HOTLINE: 'hotline Shop' };
  var STEPS = [['D', 'D · Khởi tạo'], ['D1', 'D+1 · Thăm dò'], ['D3', 'D+3 · Thúc đẩy'], ['D7', 'D+7 · Cấp thiết'], ['TRE', 'Trễ lịch']];
  var GROUPS = [['KH_MOI', 'KH mới'], ['CSKH_SAU_TIEM', 'CSKH sau tiêm'], ['BAU_NHI', 'Bầu – Nhi'], ['CHUNG', 'Chung']];
  var XH = [['', 'Tự động theo người liên hệ'], ['mẹ', 'mẹ'], ['bố', 'bố'], ['anh', 'anh'], ['chị', 'chị'], ['cô', 'cô'], ['chú', 'chú'], ['bác', 'bác'], ['ông', 'ông'], ['bà', 'bà']];

  function strip(s) { return String(s || '').replace(/\s*\([^)]*\)/g, '').trim() }
  function given(s) { var p = strip(s).split(/\s+/); return p[p.length - 1] || '' }
  function cap(s) { return s ? s.charAt(0).toUpperCase() + s.slice(1) : s }
  function custDays(c) { var b = c && c.ngay_sinh ? new Date(c.ngay_sinh + 'T00:00:00') : null; return b ? Math.max(0, Math.floor((Date.now() - b) / 86400000)) : null }
  function tuoiTxt(c) { var d = custDays(c); if (d === null) return ''; if (d < 60) return d + ' ngày tuổi'; var m = Math.floor(d / 30.4375); return m < 24 ? m + ' tháng tuổi' : Math.floor(m / 12) + ' tuổi' }
  function dayLabel(d) { if (d == null) return 'không giới hạn'; if (d < 60) return d + ' ngày'; if (d < 730) return Math.round(d / 30.4375) + ' tháng'; var y = d / 365.25; return (Math.round(y * 10) / 10) + ' tuổi' }
  function xungOf(link) { var q = String(link && link.quan_he || '').toLowerCase(); return ['mẹ', 'bố', 'ông', 'bà', 'anh', 'chị', 'cô', 'chú', 'bác'].indexOf(q) >= 0 ? q : 'anh/chị' }
  function telOf(l) { var s = String(l && l.sdt || '').replace(/[\s.\-()]/g, ''); if (/^\+84\d{9}$/.test(s)) s = '0' + s.slice(3); else if (/^84\d{9}$/.test(s)) s = '0' + s.slice(2); return s }
  function stepNow(t) { var a = ageDays(t.d_giao); return a < 1 ? 'D' : a < 3 ? 'D1' : a < 7 ? 'D3' : 'D7' }
  function fill(text, cx) { return String(text || '').replace(/\{\{([A-Z_]+)\}\}/g, function (m, k) { return cx[k] != null && cx[k] !== '' ? cx[k] : m }) }
  function shopOf() { return (D.shops && D.shops[ME.id_shop]) || {} }
  function ctx(c, link, xhOverride) {
    var xh = xhOverride || xungOf(link), sh = shopOf();
    return { XUNG_HO: xh, XUNG_HO_HOA: cap(xh), TEN_NGUOI_LH: given(link && link.ho_ten), TEN_BE: given(c && c.ho_ten) || 'bé', TUOI: tuoiTxt(c), TEN_NV: given(ME.ho_ten), TEN_CO_SO: strip(sh.ten_shop) || 'Trung tâm tiêm chủng Long Châu', DIA_CHI: strip(sh.dia_chi), GIO_MO_CUA: sh.gio_mo_cua || '', HOTLINE: sh.hotline || '' }
  }
  function preview(text) { return fill(text, { XUNG_HO: 'mẹ', XUNG_HO_HOA: 'Mẹ', TEN_NGUOI_LH: 'Lan', TEN_BE: 'An', TUOI: '3 tháng tuổi', TEN_NV: 'Hiền', TEN_CO_SO: 'Shop mẫu', DIA_CHI: 'Địa chỉ Shop mẫu', GIO_MO_CUA: '7:30–21:00', HOTLINE: '0900 000 111' }) }
  function activeTpl() { return (D.tpl || []).filter(function (x) { return x.trang_thai === 'DANG_DUNG' }) }
  function tplById(ma) { return (D.tpl || []).filter(function (x) { return x.ma === ma })[0] }
  function tplForStep(step) { return activeTpl().filter(function (x) { return x.nhom === 'KH_MOI' && x.buoc === step }) }
  function dueRules(days, gender) { if (days == null) return []; return (D.rules || []).filter(function (r) { return r.trang_thai === 'DANG_DUNG' && days >= r.tu_ngay && (r.den_ngay == null || days <= r.den_ngay) && (!r.gioi_tinh || r.gioi_tinh === gender) }) }

  /* ---------- hiện trong chi tiết khách ---------- */
  function counters(t) {
    var g = 0, m = 0; D.tt.forEach(function (x) { if (x.ma_luot !== t.ma_luot) return; var c = D.dm[x.loai_tuong_tac]; if (!c) return; if (c.ma === 'GOI') g++; else if (isMsg(x)) m++ });
    return '<span class="tag t-new">Gọi ' + g + '/3</span><span class="tag t-new">Nhắn ' + m + '/2</span>'
  }
  function scriptBlock(t, link) {
    var c = D.cus[t.ma_khach] || {}, st = stepNow(t), g = activeTpl().filter(function (x) { return x.nhom === 'GOI' }), tp = g.filter(function (x) { return x.buoc === st })[0] || g[0], txt, src;
    if (tp) { txt = fill(tp.noi_dung, ctx(c, link)); src = 'Thư viện · ' + tp.ma + ' · ' + esc(tp.tieu_de) }
    else { txt = 'Chào ' + xungOf(link) + ' ' + given(link && link.ho_ten) + ', em là ' + given(ME.ho_ten) + ' ở trung tâm tiêm chủng Long Châu. Em gọi về bé ' + given(c.ho_ten) + ' (' + tuoiTxt(c) + '). Anh/chị tiện nói chuyện một chút được không ạ?'; src = 'Mẫu tạm (Thư viện chưa có lời nói khi gọi)' }
    return '<div class="sec"><h3>Gợi ý nói khi gọi</h3><div class="script">' + esc(txt) + '<small>' + src + '</small></div></div>'
  }

  /* ---------- màn soạn tin ---------- */
  var S = null, host = null;
  function guard() { if (S.text !== S.base && !confirm('Anh/chị đã sửa tay nội dung. Đổi lựa chọn sẽ tạo lại tin theo mẫu và mất phần đã sửa. Tiếp tục?')) return false; return true }
  function open(t, ci) {
    var links = D.links[t.ma_khach] || []; if (!links.length) { toast('Khách này chưa có người liên hệ.', 0, 1); return }
    var c = D.cus[t.ma_khach] || {}, step = stepNow(t), tp = tplForStep(step)[0], vx = {}; dueRules(custDays(c), c.gioi_tinh).forEach(function (r) { vx[r.ma] = true });
    S = { t: t, ci: Math.min(ci || 0, links.length - 1), step: step, tplId: tp ? tp.ma : '', vxOn: true, vx: vx, km: {}, ngc: false, dc: false, xh: '', view: 'main', pg: 'KH_MOI', pq: '', text: '', base: '', sending: false };
    host = document.createElement('div'); host.className = 'modal mm'; host.onclick = function (e) { if (e.target === host) close() }; document.body.appendChild(host);
    regen()
  }
  function close() { if (host) host.remove(); host = null; S = null }
  function link() { return (D.links[S.t.ma_khach] || [])[S.ci] || {} }
  function cust() { return D.cus[S.t.ma_khach] || {} }
  function compose() {
    var c = cust(), cx = ctx(c, link(), S.xh), parts = [], tp = tplById(S.tplId);
    if (tp) parts.push(fill(tp.noi_dung, cx));
    if (S.vxOn) { var due = dueRules(custDays(c), c.gioi_tinh).filter(function (r) { return S.vx[r.ma] }); if (due.length) parts.push('Theo ngày sinh, bé ' + cx.TEN_BE + ' đã đến tuổi tiêm: ' + due.map(function (r) { return r.ten }).join(', ') + '.') }
    var km = (D.promos || []).filter(function (p) { return S.km[p.ma] }); if (km.length) parts.push(km.map(function (p) { return fill(p.dong, cx) }).join(' '));
    if (S.ngc) { var ng = activeTpl().filter(function (x) { return x.nhom === 'NGUY_CO' })[0]; if (ng) parts.push(fill(ng.noi_dung, cx)) }
    if (S.dc) parts.push('Địa chỉ: ' + cx.DIA_CHI + '. Giờ mở cửa: ' + cx.GIO_MO_CUA + '. Hotline: ' + cx.HOTLINE + '.');
    return parts.join('\n\n')
  }
  function regen() { S.base = compose(); S.text = S.base; draw() }
  function msgCount() { var m = 0; D.tt.forEach(function (x) { if (x.ma_luot === S.t.ma_luot && isMsg(x)) m++ }); return m }
  function draw() { S.view === 'pick' ? drawPick() : drawMain() }
  function drawMain() {
    var c = cust(), l = link(), links = D.links[S.t.ma_khach] || [], edited = S.text !== S.base, now = stepNow(S.t), tp = tplById(S.tplId), cx = ctx(c, l, S.xh), tel = telOf(l);
    var due = dueRules(custDays(c), c.gioi_tinh), alts = tplForStep(S.step).filter(function (x) { return x.ma !== S.tplId }).slice(0, 2), promos = D.promos || [], nTpl = activeTpl().filter(function (x) { return ['KH_MOI', 'CSKH_SAU_TIEM', 'BAU_NHI', 'CHUNG'].indexOf(x.nhom) >= 0 }).length;
    var h = '<div class="msheet"><div class="msh"><div><div class="name" style="font-size:18px">Tin cho ' + esc(l.quan_he || 'người liên hệ') + ' ' + esc(given(l.ho_ten)) + '</div><div class="meta">Bé ' + esc(given(c.ho_ten)) + ' · ' + esc(tuoiTxt(c)) + (tel ? ' · ' + esc(tel) : '') + '</div></div><button class="x" id="mx" aria-label="Đóng">' + ic('x') + '</button></div><div class="msb">';
    if (links.length > 1) h += '<div><h3>Gửi cho</h3><div class="seg">' + links.map(function (x, i) { return '<button data-ci="' + i + '" class="' + (i === S.ci ? 'on' : '') + '">' + esc(x.quan_he || 'Liên hệ') + (x.ho_ten ? ' ' + esc(given(x.ho_ten)) : '') + '</button>' }).join('') + '</div></div>';
    h += '<div><h3>Bước chăm sóc</h3><div class="chips">' + STEPS.map(function (s) { return '<button class="chip ' + (S.step === s[0] ? 'on' : '') + (s[0] === now ? ' star' : '') + '" data-s="' + s[0] + '">' + esc(s[1]) + '</button>' }).join('') + '</div><div class="meta" style="margin-top:6px">★ = bước hợp với số ngày kể từ khi giao, đã chọn sẵn. "Trễ lịch" chọn tay khi khách đã lỡ mũi.</div></div>';
    h += '<div><div class="row" style="margin-bottom:8px"><h3 style="margin:0">Nội dung (sửa được trước khi gửi)</h3><span class="sp"></span><span class="badge ' + (edited ? 'b-ed' : 'b-ok') + '" id="bd">' + (edited ? 'Đã sửa tay' : 'Đúng theo mẫu') + '</span></div>' +
      (tp ? '<div class="meta" style="margin-bottom:6px">Mẫu ' + esc(tp.ma) + ' · ' + esc(tp.tieu_de) + '</div>' : '<div class="alert a-warn"><i>⚠</i><div><b>Chưa có mẫu cho bước này</b><small>Bấm "Xem thêm mẫu" để chọn, hoặc tự viết bên dưới.</small></div></div>') +
      '<textarea id="tx">' + esc(S.text) + '</textarea><div class="row" style="margin-top:6px"><button class="btn b3" id="rs" style="min-height:38px">Khôi phục theo mẫu</button><span class="sp"></span><span class="meta" id="cc">' + S.text.length + ' ký tự</span></div></div>';
    h += '<div><h3>Thêm vào tin</h3><div style="display:grid;gap:8px">' +
      '<div class="sw" style="display:block"><label class="row" style="cursor:pointer"><input type="checkbox" id="vxon" ' + (S.vxOn ? 'checked' : '') + '><div><b>Vaccine đã đến tuổi</b><small>' + (due.length ? due.length + ' loại theo ngày sinh của bé' : 'Chưa có vaccine nào tới tuổi theo phác đồ') + '</small></div></label>' +
      (S.vxOn && due.length ? '<div class="vxl">' + due.map(function (r) { return '<label><input type="checkbox" data-v="' + esc(r.ma) + '" ' + (S.vx[r.ma] ? 'checked' : '') + '><span>' + esc(r.ten) + '<small>' + esc(dayLabel(r.tu_ngay)) + (r.den_ngay != null ? ' – ' + esc(dayLabel(r.den_ngay)) : ' trở lên') + '</small></span></label>' }).join('') + '</div>' : '') + '</div>' +
      '<div class="sw" style="display:block"><b>Khuyến mãi hiện hành (' + promos.length + ')</b>' + (promos.length ? '<div class="vxl">' + promos.map(function (p) { return '<label><input type="checkbox" data-k="' + esc(p.ma) + '" ' + (S.km[p.ma] ? 'checked' : '') + '><span>' + esc(p.ten) + '</span></label>' }).join('') + '</div>' : '<div class="meta" style="margin-top:6px">Chưa có khuyến mãi đang chạy.</div>') + '</div>' +
      '<label class="sw"><input type="checkbox" id="swn" ' + (S.ngc ? 'checked' : '') + '><div><b>Kèm nguy cơ</b><small>Thêm đoạn giải thích nguy cơ khi chưa tiêm đủ</small></div></label>' +
      '<label class="sw"><input type="checkbox" id="swd" ' + (S.dc ? 'checked' : '') + '><div><b>Kèm địa chỉ, giờ mở cửa, hotline</b><small>' + esc(cx.TEN_CO_SO) + '</small></div></label></div></div>';
    h += '<div><h3>Mẫu khác cùng bước</h3><div style="display:grid;gap:8px">' + alts.map(function (a) { return '<button class="alt" data-t="' + esc(a.ma) + '"><b>' + esc(a.tieu_de) + '</b>' + esc(fill(a.noi_dung, cx).slice(0, 90)) + '…</button>' }).join('') + '<button class="btn b3" id="more">Xem thêm mẫu (' + nTpl + ')</button></div></div>';
    h += '<div class="row"><span class="meta">Xưng hô:</span><select id="xh">' + XH.map(function (o) { return '<option value="' + esc(o[0]) + '"' + (o[0] === S.xh ? ' selected' : '') + '>' + esc(o[1]) + '</option>' }).join('') + '</select></div>';
    h += '</div><div class="msf"><div class="meta" style="text-align:center">Chạm nút nào thì làm đúng việc của nút đó, rồi app tự ghi 1 lần "Nhắn" (' + (msgCount() + 1) + '/2).</div><div class="sendrow"><button class="btn b1" data-go="ZALO"><span>Zalo</span><small>chép tin + mở Zalo</small></button><button class="btn b2" data-go="SMS"><span>SMS</span><small>tới ' + esc(tel || 'chưa có số') + '</small></button><button class="btn b3" data-go="SAO_CHEP"><span>Sao chép</span><small>chỉ chép tin</small></button></div></div></div>';
    host.innerHTML = h; bindMain()
  }
  function bindMain() {
    var q = function (s) { return host.querySelector(s) }, all = function (s) { return [].slice.call(host.querySelectorAll(s)) };
    q('#mx').onclick = close;
    all('[data-ci]').forEach(function (b) { b.onclick = function () { if (!guard()) return; S.ci = Number(b.dataset.ci); regen() } });
    all('.chip[data-s]').forEach(function (b) { b.onclick = function () { if (!guard()) return; S.step = b.dataset.s; var tp = tplForStep(S.step)[0]; if (tp) S.tplId = tp.ma; regen() } });
    all('.alt[data-t]').forEach(function (b) { b.onclick = function () { if (!guard()) return; S.tplId = b.dataset.t; regen() } });
    var tx = q('#tx'); tx.oninput = function () { S.text = tx.value; var ed = S.text !== S.base, bd = q('#bd'); bd.className = 'badge ' + (ed ? 'b-ed' : 'b-ok'); bd.textContent = ed ? 'Đã sửa tay' : 'Đúng theo mẫu'; q('#cc').textContent = S.text.length + ' ký tự' };
    q('#rs').onclick = function () { S.text = S.base; draw() };
    q('#vxon').onchange = function (e) { if (!guard()) { draw(); return } S.vxOn = e.target.checked; regen() };
    all('[data-v]').forEach(function (b) { b.onchange = function () { if (!guard()) { draw(); return } S.vx[b.dataset.v] = b.checked; regen() } });
    all('[data-k]').forEach(function (b) { b.onchange = function () { if (!guard()) { draw(); return } S.km[b.dataset.k] = b.checked; regen() } });
    q('#swn').onchange = function (e) { if (!guard()) { draw(); return } S.ngc = e.target.checked; regen() };
    q('#swd').onchange = function (e) { if (!guard()) { draw(); return } S.dc = e.target.checked; regen() };
    q('#xh').onchange = function (e) { if (!guard()) { draw(); return } S.xh = e.target.value; regen() };
    q('#more').onclick = function () { S.view = 'pick'; draw() };
    all('[data-go]').forEach(function (b) { b.onclick = function () { send(b.dataset.go) } })
  }
  function drawPick() {
    var c = cust(), cx = ctx(c, link(), S.xh), q = S.pq.trim().toLowerCase();
    var list = activeTpl().filter(function (x) { return x.nhom === S.pg && (!q || (x.tieu_de + ' ' + x.noi_dung).toLowerCase().indexOf(q) >= 0) });
    var cnt = function (g) { return activeTpl().filter(function (x) { return x.nhom === g }).length };
    var h = '<div class="msheet"><div class="msh"><button class="x" id="pb" aria-label="Quay lại" style="margin-left:0">' + ic('back') + '</button><div class="name" style="font-size:18px">Chọn mẫu tin</div><button class="x" id="mx" aria-label="Đóng">' + ic('x') + '</button></div><div class="msb">' +
      '<div class="chips">' + GROUPS.map(function (g) { return '<button class="chip ' + (S.pg === g[0] ? 'on' : '') + '" data-g="' + g[0] + '">' + esc(g[1]) + '<em>' + cnt(g[0]) + '</em></button>' }).join('') + '</div>' +
      '<div class="search" style="max-width:none">' + ic('search') + '<input id="pqi" placeholder="Tìm trong mẫu…" value="' + esc(S.pq) + '"></div>' +
      '<div style="display:grid;gap:8px">' + (list.length ? list.map(function (x) { return '<button class="alt" data-t="' + esc(x.ma) + '"><b>' + esc(x.tieu_de) + (x.buoc ? ' <span class="tag t-new" style="margin:0 0 0 4px">' + esc(x.buoc) + '</span>' : '') + '</b>' + esc(fill(x.noi_dung, cx).slice(0, 140)) + '…</button>' }).join('') : '<div class="card empty">Không có mẫu phù hợp.</div>') + '</div></div></div>';
    host.innerHTML = h;
    host.querySelector('#mx').onclick = close; host.querySelector('#pb').onclick = function () { S.view = 'main'; draw() };
    [].slice.call(host.querySelectorAll('[data-g]')).forEach(function (b) { b.onclick = function () { S.pg = b.dataset.g; draw() } });
    var pi = host.querySelector('#pqi'); pi.oninput = function () { S.pq = pi.value; var pos = pi.selectionStart; draw(); var n = host.querySelector('#pqi'); n.focus(); n.setSelectionRange(pos, pos) };
    [].slice.call(host.querySelectorAll('.alt[data-t]')).forEach(function (b) { b.onclick = function () { if (S.text !== S.base && !confirm('Đổi mẫu sẽ tạo lại tin và mất phần đã sửa tay. Tiếp tục?')) return; S.tplId = b.dataset.t; var tp = tplById(S.tplId); if (tp && tp.buoc) S.step = tp.buoc; S.view = 'main'; regen() } })
  }

  /* ---------- gửi: mỗi nút làm việc của nó, rồi ghi 1 lần "Nhắn" ---------- */
  function send(kind) {
    if (S.sending) return; var l = link(), tel = telOf(l), text = S.text.trim(), t = S.t;
    if (!text) { toast('Tin đang trống, hãy chọn mẫu hoặc tự viết.', 0, 1); return }
    if (kind !== 'SAO_CHEP' && !tel) { toast('Người liên hệ này chưa có số điện thoại.', 0, 1); return }
    S.sending = true; var edited = S.text !== S.base, tpl = S.tplId, step = S.step, nlh = l.ma_nlh;
    copy(text);
    if (kind === 'ZALO') window.open('https://zalo.me/' + tel, '_blank', 'noopener');
    if (kind === 'SMS') { var a = document.createElement('a'); a.href = 'sms:' + tel + (/iPhone|iPad|iPod/.test(navigator.userAgent) ? '&' : '?') + 'body=' + encodeURIComponent(text); document.body.appendChild(a); a.click(); a.remove() }
    var chan = { ZALO: 'ZALO', SMS: 'NHAN_SMS', SAO_CHEP: 'NHAN_CHEP' }[kind], loai = byMa('LOAI_TUONG_TAC', chan), prevSt = t.trang_thai, fake = null;
    if (loai) { fake = { ma_luot: t.ma_luot, ma_tk: ME.ma_tk, loai_tuong_tac: loai.ma_dm, ket_qua_lien_lac: 'DM-KL6', ghi_chu: '', d: new Date(), sdt_lien_he: tel }; D.tt.push(fake); relast(t.ma_luot); if (t.trang_thai === 'DA_GIAO') t.trang_thai = 'DANG_XU_LY' }
    var msg = { ZALO: 'Đã chép tin và mở Zalo tới ' + tel + '. Dán tin vào ô chat để gửi hoặc kết bạn.', SMS: 'Đã mở tin nhắn SMS tới ' + tel + ' (tin cũng đã được chép sẵn).', SAO_CHEP: 'Đã chép tin. Dán vào ứng dụng tin nhắn để gửi.' }[kind];
    close(); toast(msg + ' Đã ghi 1 lần Nhắn.'); render();
    rpc('u_msg_log', { p_ma_luot: t.ma_luot, p_kenh: kind, p_ma_nlh: nlh || '', p_ma_mau: tpl || '', p_buoc: step || '', p_da_sua: edited, p_noi_dung: text, p_req: rid() })
      .then(function () { return load(true) }).then(render, function (e) { if (fake) { var i = D.tt.indexOf(fake); if (i >= 0) D.tt.splice(i, 1); t.trang_thai = prevSt; relast(t.ma_luot); render() } fail(e) })
  }

  return { open: open, counters: counters, scriptBlock: scriptBlock, preview: preview, fill: fill, dueRules: dueRules, custDays: custDays, dayLabel: dayLabel, VARS: VARS, VARDESC: VARDESC, STEPS: STEPS, GROUPS: GROUPS }
})();
