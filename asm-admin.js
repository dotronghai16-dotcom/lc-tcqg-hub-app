/* ASM – NHÂN SỰ & TÀI KHOẢN, THƯ VIỆN (mẫu tin, khuyến mãi), PHÁC ĐỒ (vaccine theo tuổi). Mở rộng ASM trong asm.js.
 * Thư viện = "NÓI GÌ" (câu chữ). Phác đồ = "KHI NÀO" (vaccine nào tới tuổi nào). Hai bên gặp nhau ở màn Gợi ý tin nhắn của nhân viên.
 * Mọi ghi qua hàm RPC của cơ sở dữ liệu (chỉ ASM gọi được). Dữ liệu mẫu và câu chữ là DỮ LIỆU GIẢ để thử. */
(function () {
  var X = ASM.X, U = X.u, head = U.head, sel = U.sel, ST = U.ST;
  IC.shield = '<path d="M12 3 4 6v6c0 5 3.5 8 8 9 4.5-1 8-4 8-9V6z"/><path d="m9 12 2 2 4-4"/>';
  var VAI = { QUAN_LY: 'Quản lý', NHAN_VIEN: 'Nhân viên' }, TT = { HOAT_DONG: 'Đang hoạt động', NGUNG: 'Đã khóa' };
  var AS = { tk: { shop: '', st: 'HOAT_DONG', q: '' }, tl: { tab: 'mau', nhom: '', st: 'ACTIVE', q: '' }, pd: { tv: '3', tu: '30', tg: '' } };
  var NHOM = { KH_MOI: 'KH mới', CSKH_SAU_TIEM: 'CSKH sau tiêm', BAU_NHI: 'Bầu – Nhi', CHUNG: 'Chung', NGUY_CO: 'Nguy cơ', GOI: 'Lời nói khi gọi' };
  var BUOC = { D: 'D', D1: 'D+1', D3: 'D+3', D7: 'D+7', TRE: 'Trễ lịch' };
  var TPLST = { DANG_DUNG: ['t-ok', 'Đang dùng'], CHO_DUYET: ['t-late', 'Chờ duyệt'], DA_THAY: ['t-new', 'Đã thay'], NGUNG_DUNG: ['t-bad', 'Ngừng dùng'] };
  var UNIT = { ngay: 1, tuan: 7, thang: 30, nam: 365 };

  function modal() { var d = document.createElement('div'); d.className = 'modal'; document.body.appendChild(d); d.onclick = function (e) { if (e.target === d) d.remove() }; return d }
  function sheet(d, inner, w) { d.innerHTML = '<div class="card" style="width:100%;max-width:' + (w || 560) + 'px;border-radius:22px 22px 0 0;padding:18px;max-height:92vh;overflow:auto">' + inner + '</div>' }
  function q(d, s) { return d.querySelector(s) }
  function done(msg) { toast(msg); return load(true).then(render) }
  function genPw() { var a = 'abcdefghjkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789', o = '', b = new Uint32Array(10); (window.crypto || window.msCrypto).getRandomValues(b); for (var i = 0; i < 10; i++) o += a[b[i] % a.length]; return o + '7' }
  function tag(cls, t) { return '<span class="tag ' + cls + '">' + esc(t) + '</span>' }
  function accs() { return Object.keys(D.acc).map(function (k) { return D.acc[k] }).filter(function (a) { return a.vai_tro !== 'ASM' }) }
  function openN(ma) { return D.turns.filter(function (t) { return t.ma_tk_nhan === ma && open_(t) }).length }
  function qlRow(shopId, except) { return accs().filter(function (a) { return a.vai_tro === 'QUAN_LY' && a.id_shop === shopId && a.trang_thai === 'HOAT_DONG' && a.ma_tk !== except })[0] }

  /* ================= NHÂN SỰ & TÀI KHOẢN ================= */
  function pgTk() {
    var f = AS.tk, SH = U.shops(), ql = f.q.trim().toLowerCase();
    var L = accs().filter(function (a) { return (!f.shop || a.id_shop === f.shop) && (!f.st || a.trang_thai === f.st) && (!ql || (a.ho_ten + ' ' + a.email).toLowerCase().indexOf(ql) >= 0) })
      .sort(function (a, b) { return String(a.id_shop).localeCompare(String(b.id_shop)) || (a.vai_tro === b.vai_tro ? a.ho_ten.localeCompare(b.ho_ten) : a.vai_tro === 'QUAN_LY' ? -1 : 1) });
    var cnt = { '': accs().length, HOAT_DONG: 0, NGUNG: 0 }; accs().forEach(function (a) { cnt[a.trang_thai] = (cnt[a.trang_thai] || 0) + 1 });
    var h = head('Nhân sự & tài khoản', 'Tạo tài khoản, đổi Shop, đặt Quản lý cho từng Shop, khóa hoặc mở, đặt lại mật khẩu. Chỉ ASM làm được.', '<button class="btn b1" id="tk-new">+ Thêm tài khoản</button>');
    h += '<div class="card" style="padding:6px 8px"><h2 style="padding:10px 10px 0">Quản lý của từng Shop</h2><div class="tw"><table><thead><tr><th>Shop</th><th>Quản lý</th><th>Nhân viên</th><th></th></tr></thead><tbody>' + SH.map(function (s) {
      var m = qlRow(s.id_shop), nv = accs().filter(function (a) { return a.id_shop === s.id_shop && a.vai_tro === 'NHAN_VIEN' && a.trang_thai === 'HOAT_DONG' }).length;
      return '<tr><td><b>' + esc(s.ten_shop) + '</b><small>' + esc(s.ma_shop) + '</small></td><td>' + (m ? esc(m.ho_ten) : '<span class="tag t-late">Chưa có Quản lý</span>') + '</td><td>' + nv + '</td><td><button class="btn b3" data-setql="' + esc(s.id_shop) + '" style="min-height:36px;font-size:13px">' + (m ? 'Đổi Quản lý' : 'Chọn Quản lý') + '</button></td></tr>'
    }).join('') + '</tbody></table></div></div>';
    h += '<div class="top" style="flex-wrap:wrap"><div class="search">' + ic('search') + '<input id="tk-q" placeholder="Tìm tên hoặc email…" value="' + esc(f.q) + '"></div>' + sel('tk-shop', [['', 'Tất cả Shop']].concat(SH.map(function (s) { return [s.id_shop, s.ten_shop] })), f.shop) + '</div>';
    h += '<div class="chips">' + [['HOAT_DONG', 'Đang hoạt động'], ['NGUNG', 'Đã khóa'], ['', 'Tất cả']].map(function (x) { return '<button class="chip ' + (f.st === x[0] ? 'on' : '') + '" data-tkst="' + x[0] + '">' + x[1] + '<em>' + (cnt[x[0]] || 0) + '</em></button>' }).join('') + '</div>';
    h += '<div class="card" style="padding:6px 8px"><div class="tw"><table><thead><tr><th>Họ tên</th><th>Vai trò</th><th>Shop</th><th>Đang giữ</th><th>Trạng thái</th><th style="min-width:230px"></th></tr></thead><tbody>' + L.map(function (a) {
      var n = openN(a.ma_tk);
      return '<tr><td><b>' + esc(a.ho_ten) + '</b><small>' + esc(a.email) + '</small></td><td>' + (a.vai_tro === 'QUAN_LY' ? tag('t-today', 'Quản lý') : 'Nhân viên') + '</td><td>' + esc(shopName(a.id_shop)) + '</td><td>' + (n ? n + ' lượt mở' : '–') + '</td><td>' + (a.trang_thai === 'HOAT_DONG' ? tag('t-ok', 'Hoạt động') : tag('t-bad', 'Đã khóa')) + '</td><td><div class="row" style="flex-wrap:wrap"><button class="btn b3" data-tkedit="' + esc(a.ma_tk) + '" style="min-height:36px;font-size:13px">Sửa</button><button class="btn b3" data-tkpw="' + esc(a.ma_tk) + '" style="min-height:36px;font-size:13px">Mật khẩu</button><button class="btn b3" data-tklock="' + esc(a.ma_tk) + '" style="min-height:36px;font-size:13px">' + (a.trang_thai === 'HOAT_DONG' ? 'Khóa' : 'Mở khóa') + '</button></div></td></tr>'
    }).join('') + (L.length ? '' : '<tr><td colspan="6" class="empty">Không có tài khoản phù hợp.</td></tr>') + '</tbody></table></div></div>';
    return h
  }
  function creds(title, email, pw) {
    var d = modal(), url = location.origin + location.pathname, txt = 'Đăng nhập LC TCQG HUB\nLink: ' + url + '\nEmail: ' + email + '\nMật khẩu: ' + pw;
    sheet(d, '<h2 style="font-size:18px;margin:0 0 6px">' + esc(title) + '</h2><div class="meta" style="margin-bottom:10px">Mật khẩu chỉ hiện một lần ở đây. Gửi cho nhân viên và nhắc đổi mật khẩu ở mục Tài khoản.</div><div class="alert" style="display:block;font-size:15px;line-height:1.7"><div>Email: <b>' + esc(email) + '</b></div><div>Mật khẩu: <b style="font-size:18px;letter-spacing:.5px">' + esc(pw) + '</b></div></div><div class="acts" style="margin-top:8px"><button class="btn b1" id="cc">Chép thông tin đăng nhập</button><button class="btn b3" id="cx">Đóng</button></div>', 480);
    q(d, '#cc').onclick = function () { copy(txt); toast('Đã chép thông tin đăng nhập') }; q(d, '#cx').onclick = function () { d.remove() }
  }
  function accModal(a) {
    var isNew = !a, SH = U.shops(), m = { ho_ten: a ? a.ho_ten : '', email: a ? a.email : '', vai_tro: a ? a.vai_tro : 'NHAN_VIEN', id_shop: a ? a.id_shop : (ST.shop || (SH[0] || {}).id_shop || ''), trang_thai: a ? a.trang_thai : 'HOAT_DONG', mk: isNew ? genPw() : '', thay: false }, d = modal();
    function draw() {
      var cf = m.vai_tro === 'QUAN_LY' && m.trang_thai === 'HOAT_DONG' ? qlRow(m.id_shop, a && a.ma_tk) : null;
      sheet(d, '<h2 style="font-size:18px;margin:0 0 10px">' + (isNew ? 'Thêm tài khoản' : 'Sửa tài khoản') + '</h2><div class="form2" style="margin-top:0"><div><label class="meta">Họ tên</label><input id="f-ten" value="' + esc(m.ho_ten) + '"></div><div><label class="meta">Email đăng nhập</label><input id="f-em" type="email" value="' + esc(m.email) + '"' + (isNew ? '' : ' disabled') + '></div><div><label class="meta">Vai trò</label>' + sel('f-vt', [['NHAN_VIEN', 'Nhân viên'], ['QUAN_LY', 'Quản lý']], m.vai_tro) + '</div><div><label class="meta">Shop</label>' + sel('f-sh', SH.map(function (s) { return [s.id_shop, s.ten_shop] }), m.id_shop) + '</div>' +
        (isNew ? '<div style="grid-column:1/-1"><label class="meta">Mật khẩu ban đầu (tối thiểu 8 ký tự)</label><div class="row"><input id="f-mk" value="' + esc(m.mk) + '" autocomplete="off"><button class="btn b3" id="f-gen" style="min-height:44px">Tạo ngẫu nhiên</button></div></div>' : '<div><label class="meta">Trạng thái</label>' + sel('f-tt', [['HOAT_DONG', 'Đang hoạt động'], ['NGUNG', 'Đã khóa']], m.trang_thai) + '</div>') + '</div>' +
        (cf ? '<label class="sw" style="margin-top:10px"><input type="checkbox" id="f-thay" ' + (m.thay ? 'checked' : '') + '><div><b>Thay Quản lý hiện tại</b><small>Shop này đang có Quản lý ' + esc(cf.ho_ten) + '. Chọn để chuyển người đó về Nhân viên và đặt người này làm Quản lý. Mỗi Shop chỉ có một Quản lý.</small></div></label>' : '') +
        (!isNew && a.vai_tro !== m.vai_tro || !isNew && a.id_shop !== m.id_shop ? '<div class="ro" style="margin-top:10px">Lượt đang giữ vẫn giữ tên người này và vẫn ghi Shop lúc giao. Quản lý cũ của Shop sẽ không còn xem được nếu người này chuyển Shop.</div>' : '') +
        '<div class="acts" style="margin-top:12px"><button class="btn b3" id="f-x" style="flex:1">Hủy</button><button class="btn b1" id="f-ok" style="flex:1.4"' + (cf && !m.thay ? ' disabled' : '') + '>' + (isNew ? 'Tạo tài khoản' : 'Lưu') + '</button></div>');
      q(d, '#f-ten').oninput = function (e) { m.ho_ten = e.target.value }; var em = q(d, '#f-em'); em.oninput = function (e) { m.email = e.target.value };
      q(d, '#f-vt').onchange = function (e) { m.vai_tro = e.target.value; m.thay = false; m.ho_ten = q(d, '#f-ten').value; m.email = em.value; draw() }; q(d, '#f-sh').onchange = function (e) { m.id_shop = e.target.value; m.thay = false; m.ho_ten = q(d, '#f-ten').value; m.email = em.value; draw() };
      var tt = q(d, '#f-tt'); if (tt) tt.onchange = function (e) { m.trang_thai = e.target.value; m.ho_ten = q(d, '#f-ten').value; draw() };
      var mk = q(d, '#f-mk'); if (mk) { mk.oninput = function (e) { m.mk = e.target.value }; q(d, '#f-gen').onclick = function () { m.mk = genPw(); mk.value = m.mk } }
      var th = q(d, '#f-thay'); if (th) th.onchange = function (e) { m.thay = e.target.checked; m.ho_ten = q(d, '#f-ten').value; if (mk) m.mk = mk.value; draw() };
      q(d, '#f-x').onclick = function () { d.remove() };
      q(d, '#f-ok').onclick = function () {
        m.ho_ten = q(d, '#f-ten').value; m.email = em.value; if (mk) m.mk = mk.value; var b = q(d, '#f-ok'); b.disabled = true;
        var p = isNew ? rpc('u_asm_account_create', { p_ho_ten: m.ho_ten, p_email: m.email, p_vai_tro: m.vai_tro, p_id_shop: m.id_shop, p_mat_khau: m.mk, p_thay_ql: m.thay, p_req: rid() }) : rpc('u_asm_account_update', { p_ma_tk: a.ma_tk, p_ho_ten: m.ho_ten, p_vai_tro: m.vai_tro, p_id_shop: m.id_shop, p_trang_thai: m.trang_thai, p_thay_ql: m.thay, p_req: rid() });
        p.then(function (r) { d.remove(); if (isNew) { load(true).then(render); creds('Đã tạo tài khoản', r.email, m.mk) } else done('Đã lưu' + (m.trang_thai === 'NGUNG' && r.luot_dang_giu ? '. Người này còn giữ ' + r.luot_dang_giu + ' lượt mở, hãy thu hồi hoặc giao lại ở mục Phân công.' : '')) }, function (e) { b.disabled = false; fail(e) })
      }
    }
    draw()
  }
  function pwAsm(a) {
    var d = modal(), pw = genPw();
    sheet(d, '<h2 style="font-size:18px;margin:0 0 6px">Đặt lại mật khẩu</h2><div class="meta" style="margin-bottom:10px">' + esc(a.ho_ten) + ' · ' + esc(a.email) + '. Mật khẩu cũ sẽ không dùng được nữa.</div><label class="field">Mật khẩu mới (tối thiểu 8 ký tự)<div class="row"><input id="p-pw" value="' + esc(pw) + '" autocomplete="off" style="flex:1"><button class="btn b3" id="p-gen" style="min-height:44px">Tạo ngẫu nhiên</button></div></label><div class="acts"><button class="btn b3" id="p-x" style="flex:1">Hủy</button><button class="btn b1" id="p-ok" style="flex:1.4">Đặt lại</button></div>', 480);
    q(d, '#p-gen').onclick = function () { q(d, '#p-pw').value = genPw() }; q(d, '#p-x').onclick = function () { d.remove() };
    q(d, '#p-ok').onclick = function () { var v = q(d, '#p-pw').value, b = q(d, '#p-ok'); b.disabled = true; rpc('u_asm_account_reset_pw', { p_ma_tk: a.ma_tk, p_mat_khau: v, p_req: rid() }).then(function () { d.remove(); creds('Đã đặt lại mật khẩu', a.email, v) }, function (e) { b.disabled = false; fail(e) }) }
  }
  function lockAsm(a) {
    var on = a.trang_thai === 'HOAT_DONG', n = openN(a.ma_tk), d = modal();
    sheet(d, '<h2 style="font-size:18px;margin:0 0 6px">' + (on ? 'Khóa tài khoản' : 'Mở khóa tài khoản') + '</h2><div class="meta" style="margin-bottom:10px">' + esc(a.ho_ten) + ' · ' + esc(a.email) + '</div>' + (on ? '<div class="ro">Người này sẽ không đăng nhập được nữa. ' + (n ? '<b>Đang giữ ' + n + ' lượt mở</b>: các lượt vẫn mang tên người này, hãy thu hồi hoặc giao lại ở mục Phân công.' : 'Hiện không giữ lượt mở nào.') + ' Lịch sử luôn được giữ.</div>' : '<div class="ro">Người này sẽ đăng nhập lại được với mật khẩu hiện tại.</div>') + '<div class="acts" style="margin-top:10px"><button class="btn b3" id="l-x" style="flex:1">Hủy</button><button class="btn b1" id="l-ok" style="flex:1.4">' + (on ? 'Xác nhận khóa' : 'Mở khóa') + '</button></div>', 480);
    q(d, '#l-x').onclick = function () { d.remove() };
    q(d, '#l-ok').onclick = function () { var b = q(d, '#l-ok'); b.disabled = true; rpc('u_asm_account_update', { p_ma_tk: a.ma_tk, p_ho_ten: a.ho_ten, p_vai_tro: a.vai_tro, p_id_shop: a.id_shop, p_trang_thai: on ? 'NGUNG' : 'HOAT_DONG', p_thay_ql: false, p_req: rid() }).then(function () { d.remove(); done(on ? 'Đã khóa ' + a.ho_ten : 'Đã mở khóa ' + a.ho_ten) }, function (e) { b.disabled = false; fail(e) }) }
  }
  function setQl(shopId) {
    var cur = qlRow(shopId), cands = accs().filter(function (a) { return a.id_shop === shopId && a.trang_thai === 'HOAT_DONG' && (!cur || a.ma_tk !== cur.ma_tk) }), d = modal();
    sheet(d, '<h2 style="font-size:18px;margin:0 0 6px">' + (cur ? 'Đổi Quản lý' : 'Chọn Quản lý') + ' · ' + esc(shopName(shopId)) + '</h2>' + (cur ? '<div class="meta" style="margin-bottom:10px">Quản lý hiện tại: <b>' + esc(cur.ho_ten) + '</b> sẽ chuyển về Nhân viên.</div>' : '') + (cands.length ? '<label class="field">Người làm Quản lý mới' + sel('q-who', cands.map(function (a) { return [a.ma_tk, a.ho_ten + ' (' + VAI[a.vai_tro] + ')'] }), cands[0].ma_tk) + '</label>' : '<div class="ro">Shop này chưa có nhân viên nào khác. Hãy thêm tài khoản trước (vai trò Quản lý).</div>') + '<div class="acts"><button class="btn b3" id="q-x" style="flex:1">Hủy</button>' + (cands.length ? '<button class="btn b1" id="q-ok" style="flex:1.4">Xác nhận</button>' : '') + '</div>', 480);
    q(d, '#q-x').onclick = function () { d.remove() };
    var ok = q(d, '#q-ok'); if (ok) ok.onclick = function () { var a = D.acc[q(d, '#q-who').value]; ok.disabled = true; rpc('u_asm_account_update', { p_ma_tk: a.ma_tk, p_ho_ten: a.ho_ten, p_vai_tro: 'QUAN_LY', p_id_shop: a.id_shop, p_trang_thai: 'HOAT_DONG', p_thay_ql: true, p_req: rid() }).then(function () { d.remove(); done(a.ho_ten + ' là Quản lý ' + shopName(shopId)) }, function (e) { ok.disabled = false; fail(e) }) }
  }

  /* ================= THƯ VIỆN ================= */
  function used(ma) { return (D.mlog || []).filter(function (x) { return x.ma_mau === ma }).length }
  function pgTl() {
    var f = AS.tl, h = head('Thư viện', 'Câu chữ nhân viên dùng khi nhắn và gọi: mẫu tin, khuyến mãi. Mẫu phải được duyệt mới hiện cho nhân viên. Phác đồ (khi nào tiêm gì) ở mục riêng.', f.tab === 'mau' ? '<button class="btn b1" id="tl-new">+ Thêm mẫu</button>' : '<button class="btn b1" id="km-new">+ Thêm khuyến mãi</button>');
    h += '<div class="chips"><button class="chip ' + (f.tab === 'mau' ? 'on' : '') + '" data-tltab="mau">Mẫu tin nhắn &amp; lời gọi<em>' + D.tpl.length + '</em></button><button class="chip ' + (f.tab === 'km' ? 'on' : '') + '" data-tltab="km">Khuyến mãi<em>' + D.promos.length + '</em></button></div>';
    return h + (f.tab === 'mau' ? tplList() : promoList())
  }
  function tplList() {
    var f = AS.tl, ql = f.q.trim().toLowerCase(), all = D.tpl;
    var L = all.filter(function (x) { return (!f.nhom || x.nhom === f.nhom) && (f.st === '' || (f.st === 'ACTIVE' ? (x.trang_thai === 'DANG_DUNG' || x.trang_thai === 'CHO_DUYET') : x.trang_thai === f.st)) && (!ql || (x.tieu_de + ' ' + x.noi_dung).toLowerCase().indexOf(ql) >= 0) });
    var h = '<div class="top" style="flex-wrap:wrap"><div class="search">' + ic('search') + '<input id="tl-q" placeholder="Tìm trong mẫu…" value="' + esc(f.q) + '"></div>' + sel('tl-st', [['ACTIVE', 'Đang dùng và chờ duyệt'], ['DANG_DUNG', 'Đang dùng'], ['CHO_DUYET', 'Chờ duyệt'], ['DA_THAY', 'Đã thay'], ['NGUNG_DUNG', 'Ngừng dùng'], ['', 'Tất cả']], f.st) + '</div>';
    h += '<div class="chips">' + [['', 'Tất cả']].concat(Object.keys(NHOM).map(function (k) { return [k, NHOM[k]] })).map(function (x) { return '<button class="chip ' + (f.nhom === x[0] ? 'on' : '') + '" data-tlg="' + x[0] + '">' + x[1] + '<em>' + all.filter(function (t) { return !x[0] || t.nhom === x[0] }).length + '</em></button>' }).join('') + '</div>';
    h += '<div class="card" style="padding:6px 8px"><div class="tw"><table><thead><tr><th>Mẫu</th><th>Nhóm</th><th>Bước</th><th>Trạng thái</th><th>Đã dùng</th><th style="min-width:190px"></th></tr></thead><tbody>' + L.map(function (x) {
      var st = TPLST[x.trang_thai] || ['t-new', x.trang_thai], btn = function (k, t) { return '<button class="btn b3" data-tp' + k + '="' + esc(x.ma) + '" style="min-height:36px;font-size:13px">' + t + '</button>' };
      return '<tr><td style="max-width:420px"><b>' + esc(x.tieu_de) + '</b><small>' + esc(x.ma) + (x.phien_ban > 1 ? ' · bản ' + x.phien_ban : '') + (x.thay_cho ? ' · thay ' + esc(x.thay_cho) : '') + '</small><small style="white-space:normal">' + esc(MSG.preview(x.noi_dung).slice(0, 140)) + (x.noi_dung.length > 140 ? '…' : '') + '</small></td><td>' + esc(NHOM[x.nhom] || x.nhom) + '</td><td>' + esc(BUOC[x.buoc] || '–') + '</td><td>' + tag(st[0], st[1]) + '</td><td>' + used(x.ma) + '</td><td><div class="row" style="flex-wrap:wrap">' + (x.trang_thai === 'CHO_DUYET' ? btn('edit', 'Sửa') + btn('ok', 'Duyệt dùng') : x.trang_thai === 'DANG_DUNG' ? btn('edit', 'Sửa') + btn('off', 'Ngừng') : x.trang_thai === 'NGUNG_DUNG' ? btn('ok', 'Bật lại') : '') + '</div></td></tr>'
    }).join('') + (L.length ? '' : '<tr><td colspan="6" class="empty">Không có mẫu phù hợp.</td></tr>') + '</tbody></table></div></div>';
    return h
  }
  function tplModal(x) {
    var d = modal(), m = { nhom: x ? x.nhom : 'KH_MOI', buoc: x ? (x.buoc || '') : '', tieu: x ? x.tieu_de : '', nd: x ? x.noi_dung : '' };
    function bad() { var o = [], re = /\{\{([^}]*)\}\}/g, r; while ((r = re.exec(m.nd))) if (MSG.VARS.indexOf(r[1]) < 0) o.push(r[0]); return o }
    function draw() {
      var b = bad();
      sheet(d, '<h2 style="font-size:18px;margin:0 0 4px">' + (x ? 'Sửa mẫu ' + esc(x.ma) : 'Thêm mẫu') + '</h2>' + (x && x.trang_thai === 'DANG_DUNG' ? '<div class="ro">Mẫu đang dùng không bị sửa trực tiếp: hệ thống tạo bản mới, bản cũ chuyển "Đã thay" khi bản mới được duyệt dùng.</div>' : '') +
        '<div class="form2" style="margin-top:0"><div><label class="meta">Nhóm</label>' + sel('t-nh', Object.keys(NHOM).map(function (k) { return [k, NHOM[k]] }), m.nhom) + '</div><div><label class="meta">Bước chăm sóc</label>' + sel('t-bu', [['', '— không theo bước —']].concat(Object.keys(BUOC).map(function (k) { return [k, BUOC[k]] })), m.buoc) + '</div><div style="grid-column:1/-1"><label class="meta">Tiêu đề (nhân viên thấy khi chọn mẫu)</label><input id="t-ti" value="' + esc(m.tieu) + '"></div></div>' +
        '<label class="meta" style="display:block;margin-top:10px">Nội dung</label><textarea id="t-nd" style="min-height:130px">' + esc(m.nd) + '</textarea><div class="row" style="margin:6px 0;flex-wrap:wrap"><span class="meta">Chèn:</span>' + MSG.VARS.map(function (v) { return '<button class="chip" data-var="' + v + '" style="min-height:32px;padding:4px 10px;font-size:12.5px" title="' + esc(MSG.VARDESC[v]) + '">' + v + '</button>' }).join('') + '</div>' +
        (b.length ? '<div class="err">Biến không có trong danh sách: ' + esc(b.join(', ')) + '</div>' : '') + '<div class="script" style="margin:8px 0"><small style="margin:0 0 6px">Xem trước (dữ liệu mẫu)</small>' + esc(MSG.preview(m.nd)) + '</div><div class="meta">' + m.nd.length + '/2000 ký tự</div>' +
        '<div class="acts" style="margin-top:12px;flex-wrap:wrap"><button class="btn b3" id="t-x" style="flex:1">Hủy</button><button class="btn b2" id="t-draft" style="flex:1.2"' + (b.length ? ' disabled' : '') + '>Lưu nháp (chờ duyệt)</button><button class="btn b1" id="t-ok" style="flex:1.4"' + (b.length ? ' disabled' : '') + '>Duyệt và dùng ngay</button></div>');
      var sync = function () { m.nhom = q(d, '#t-nh').value; m.buoc = q(d, '#t-bu').value; m.tieu = q(d, '#t-ti').value; m.nd = q(d, '#t-nd').value };
      q(d, '#t-nh').onchange = function () { sync(); draw() }; q(d, '#t-bu').onchange = sync; q(d, '#t-ti').oninput = sync;
      q(d, '#t-nd').oninput = function () { sync(); var pos = this.selectionStart; draw(); var n = q(d, '#t-nd'); n.focus(); n.setSelectionRange(pos, pos) };
      [].slice.call(d.querySelectorAll('[data-var]')).forEach(function (c) { c.onclick = function () { sync(); var ta = q(d, '#t-nd'), p = ta.selectionStart == null ? m.nd.length : ta.selectionStart, ins = '{{' + c.dataset.var + '}}'; m.nd = m.nd.slice(0, p) + ins + m.nd.slice(ta.selectionEnd == null ? p : ta.selectionEnd); draw(); var n = q(d, '#t-nd'); n.focus(); n.setSelectionRange(p + ins.length, p + ins.length) } });
      q(d, '#t-x').onclick = function () { d.remove() };
      var save = function (st) { sync(); if (!m.tieu.trim() || !m.nd.trim()) { toast('Cần có tiêu đề và nội dung.', 0, 1); return } rpc('u_asm_tpl_save', { p_ma: x ? x.ma : '', p_nhom: m.nhom, p_buoc: m.buoc, p_tieu_de: m.tieu, p_noi_dung: m.nd, p_trang_thai: st, p_req: rid() }).then(function (r) { d.remove(); done(st === 'DANG_DUNG' ? 'Đã duyệt và dùng mẫu ' + r.ma : 'Đã lưu nháp ' + r.ma + ', chờ duyệt') }, fail) };
      q(d, '#t-draft').onclick = function () { save('CHO_DUYET') }; q(d, '#t-ok').onclick = function () { save('DANG_DUNG') }
    }
    draw()
  }
  function tplState(ma, st, msg) { rpc('u_asm_tpl_state', { p_ma: ma, p_trang_thai: st, p_req: rid() }).then(function () { return done(msg) }, fail) }
  function promoList() {
    var h = '<div class="card" style="padding:6px 8px"><div class="tw"><table><thead><tr><th>Khuyến mãi</th><th>Câu dùng trong tin nhắn</th><th>Hiệu lực</th><th>Trạng thái</th><th></th></tr></thead><tbody>' + D.promos.map(function (p) {
      var t0 = new Date().toISOString().slice(0, 10), live = p.trang_thai === 'DANG_DUNG' && p.tu_ngay <= t0 && (!p.den_ngay || p.den_ngay >= t0);
      return '<tr><td><b>' + esc(p.ten) + '</b><small>' + esc(p.ma) + '</small></td><td style="max-width:360px;white-space:normal">' + esc(p.dong) + '</td><td>' + esc(p.tu_ngay) + ' → ' + esc(p.den_ngay || 'không hạn') + '</td><td>' + (live ? tag('t-ok', 'Đang chạy') : p.trang_thai === 'NGUNG_DUNG' ? tag('t-bad', 'Ngừng') : tag('t-new', 'Ngoài hiệu lực')) + '</td><td><button class="btn b3" data-kmedit="' + esc(p.ma) + '" style="min-height:36px;font-size:13px">Sửa</button></td></tr>'
    }).join('') + (D.promos.length ? '' : '<tr><td colspan="5" class="empty">Chưa có khuyến mãi.</td></tr>') + '</tbody></table></div></div>';
    return h
  }
  function promoModal(p) {
    var d = modal(), t0 = new Date().toISOString().slice(0, 10);
    sheet(d, '<h2 style="font-size:18px;margin:0 0 10px">' + (p ? 'Sửa khuyến mãi ' + esc(p.ma) : 'Thêm khuyến mãi') + '</h2><label class="field">Tên (nhân viên thấy khi chọn)<input id="k-ten" value="' + esc(p ? p.ten : '') + '"></label><label class="field">Câu dùng trong tin nhắn<textarea id="k-dong" style="min-height:80px">' + esc(p ? p.dong : '') + '</textarea></label><div class="form2" style="margin-top:0"><div><label class="meta">Từ ngày</label><input id="k-tu" type="date" value="' + esc(p ? p.tu_ngay : t0) + '"></div><div><label class="meta">Đến ngày (trống = không hạn)</label><input id="k-den" type="date" value="' + esc(p && p.den_ngay ? p.den_ngay : '') + '"></div><div><label class="meta">Trạng thái</label>' + sel('k-tt', [['DANG_DUNG', 'Đang dùng'], ['NGUNG_DUNG', 'Ngừng']], p ? p.trang_thai : 'DANG_DUNG') + '</div></div><div class="acts" style="margin-top:12px"><button class="btn b3" id="k-x" style="flex:1">Hủy</button><button class="btn b1" id="k-ok" style="flex:1.4">Lưu</button></div>', 520);
    q(d, '#k-x').onclick = function () { d.remove() };
    q(d, '#k-ok').onclick = function () { var b = q(d, '#k-ok'); b.disabled = true; rpc('u_asm_promo_save', { p_ma: p ? p.ma : '', p_ten: q(d, '#k-ten').value, p_dong: q(d, '#k-dong').value, p_tu: q(d, '#k-tu').value || null, p_den: q(d, '#k-den').value || null, p_trang_thai: q(d, '#k-tt').value, p_req: rid() }).then(function () { d.remove(); done('Đã lưu khuyến mãi') }, function (e) { b.disabled = false; fail(e) }) }
  }

  /* ================= PHÁC ĐỒ ================= */
  function pgPd() {
    var t = AS.pd, days = Math.round(Number(t.tv || 0) * Number(t.tu)), due = MSG.dueRules(days, t.tg);
    var h = head('Phác đồ vaccine', 'Quy tắc "khi nào": vaccine nào đến tuổi nào, tính từ ngày sinh. App dùng để gợi ý "vaccine đã đến tuổi" trong tin nhắn của nhân viên. Dòng chỉ hiện cho nhân viên khi đang dùng.', '<button class="btn b1" id="pd-new">+ Thêm dòng phác đồ</button>');
    h += '<div class="ro" style="background:var(--warn-bg);color:var(--warn)"><b>Dữ liệu mẫu minh họa.</b> Các độ tuổi bên dưới chưa phải phác đồ chuẩn đã duyệt. Bác sĩ cần nhập và duyệt phác đồ thật trước khi dùng cho khách thật; dòng "Chờ duyệt" không hiện cho nhân viên.</div>';
    h += '<div class="card"><h2>Thử nhanh: bé ở tuổi này thì gợi ý những vaccine nào?</h2><div class="row" style="flex-wrap:wrap"><input id="pd-tv" type="number" min="0" value="' + esc(t.tv) + '" style="width:90px;border:1px solid var(--line);border-radius:12px;padding:10px;min-height:44px;background:var(--card)">' + sel('pd-tu', [['1', 'ngày'], ['7', 'tuần'], ['30', 'tháng'], ['365', 'năm']], t.tu) + sel('pd-tg', [['', 'Cả nam và nữ'], ['M', 'Nam'], ['F', 'Nữ']], t.tg) + '<span class="meta">= ' + days + ' ngày tuổi</span></div><div style="margin-top:10px">' + (due.length ? due.map(function (r) { return tag('t-today', r.ten) }).join('') : '<span class="meta">Chưa có vaccine nào đang dùng ở tuổi này.</span>') + '</div></div>';
    h += '<div class="card" style="padding:6px 8px"><div class="tw"><table><thead><tr><th>Vaccine</th><th>Từ tuổi</th><th>Đến tuổi</th><th>Giới tính</th><th>Phác đồ (số mũi, khoảng cách)</th><th>Ghi chú cho nhân viên</th><th>Trạng thái</th><th style="min-width:170px"></th></tr></thead><tbody>' + D.rules.map(function (r) {
      var b = function (k, tx) { return '<button class="btn b3" data-pd' + k + '="' + esc(r.ma) + '" style="min-height:36px;font-size:13px">' + tx + '</button>' };
      return '<tr><td><b>' + esc(r.ten) + '</b><small>' + esc(r.ma_vaccine) + ' · ' + esc(r.ma) + '</small></td><td>' + esc(MSG.dayLabel(r.tu_ngay)) + '<small>' + r.tu_ngay + ' ngày</small></td><td>' + (r.den_ngay == null ? 'Không giới hạn' : esc(MSG.dayLabel(r.den_ngay)) + '<small>' + r.den_ngay + ' ngày</small>') + '</td><td>' + (r.gioi_tinh === 'M' ? 'Nam' : r.gioi_tinh === 'F' ? 'Nữ' : 'Cả hai') + '</td><td style="max-width:220px;white-space:normal">' + esc(r.phac_do || '–') + '</td><td style="max-width:220px;white-space:normal">' + esc(r.cau_goi_y || '–') + '</td><td>' + (r.trang_thai === 'DANG_DUNG' ? tag('t-ok', 'Đang dùng') : r.trang_thai === 'CHO_DUYET' ? tag('t-late', 'Chờ duyệt') : tag('t-bad', 'Ngừng')) + '</td><td><div class="row" style="flex-wrap:wrap">' + b('edit', 'Sửa') + (r.trang_thai === 'CHO_DUYET' ? b('ok', 'Duyệt') : r.trang_thai === 'DANG_DUNG' ? b('off', 'Ngừng') : b('ok', 'Bật lại')) + '</div></td></tr>'
    }).join('') + '</tbody></table></div></div>';
    return h
  }
  function ruleSave(r, patch) { var o = { p_ma: r ? r.ma : '', p_ma_vaccine: r.ma_vaccine, p_ten: r.ten, p_tu_ngay: r.tu_ngay, p_den_ngay: r.den_ngay, p_gioi_tinh: r.gioi_tinh || '', p_phac_do: r.phac_do || '', p_cau: r.cau_goi_y || '', p_trang_thai: r.trang_thai, p_req: rid() }; for (var k in patch) o[k] = patch[k]; return rpc('u_asm_rule_save', o) }
  function ruleModal(r) {
    var d = modal(), pick = function (v) { return v == null ? ['', '1'] : [String(v), '1'] };
    sheet(d, '<h2 style="font-size:18px;margin:0 0 10px">' + (r ? 'Sửa phác đồ ' + esc(r.ma) : 'Thêm dòng phác đồ') + '</h2><div class="form2" style="margin-top:0"><div><label class="meta">Mã vaccine (ví dụ 6IN1)</label><input id="r-mv" value="' + esc(r ? r.ma_vaccine : '') + '"></div><div><label class="meta">Tên hiển thị</label><input id="r-ten" value="' + esc(r ? r.ten : '') + '"></div>' +
      '<div><label class="meta">Từ tuổi</label><div class="row"><input id="r-tu" type="number" min="0" value="' + esc(r ? r.tu_ngay : 0) + '" style="flex:1">' + sel('r-tuu', [['1', 'ngày'], ['7', 'tuần'], ['30', 'tháng'], ['365', 'năm']], '1') + '</div></div><div><label class="meta">Đến tuổi (trống = không giới hạn)</label><div class="row"><input id="r-den" type="number" min="0" value="' + esc(r && r.den_ngay != null ? r.den_ngay : '') + '" style="flex:1">' + sel('r-denu', [['1', 'ngày'], ['7', 'tuần'], ['30', 'tháng'], ['365', 'năm']], '1') + '</div></div>' +
      '<div><label class="meta">Giới tính</label>' + sel('r-gt', [['', 'Cả nam và nữ'], ['M', 'Chỉ nam'], ['F', 'Chỉ nữ']], r ? (r.gioi_tinh || '') : '') + '</div><div><label class="meta">Trạng thái</label>' + sel('r-tt', [['CHO_DUYET', 'Chờ duyệt'], ['DANG_DUNG', 'Đang dùng'], ['NGUNG_DUNG', 'Ngừng']], r ? r.trang_thai : 'CHO_DUYET') + '</div>' +
      '<div style="grid-column:1/-1"><label class="meta">Phác đồ: số mũi, khoảng cách (bác sĩ nhập)</label><textarea id="r-pd" style="min-height:70px">' + esc(r ? r.phac_do || '' : '') + '</textarea></div><div style="grid-column:1/-1"><label class="meta">Ghi chú cho nhân viên (hiện cạnh vaccine trong gợi ý)</label><input id="r-ct" value="' + esc(r ? r.cau_goi_y || '' : '') + '"></div></div><div class="meta" style="margin-top:6px">Tuổi nhập bằng ngày, tuần (7 ngày), tháng (30 ngày) hoặc năm (365 ngày); hệ thống lưu theo ngày.</div><div class="acts" style="margin-top:12px"><button class="btn b3" id="r-x" style="flex:1">Hủy</button><button class="btn b1" id="r-ok" style="flex:1.4">Lưu</button></div>', 600);
    q(d, '#r-x').onclick = function () { d.remove() };
    q(d, '#r-ok').onclick = function () {
      var tu = q(d, '#r-tu').value, den = q(d, '#r-den').value, b = q(d, '#r-ok'); if (tu === '') { toast('Cần nhập "Từ tuổi".', 0, 1); return }
      b.disabled = true; ruleSave({ ma: r ? r.ma : '', ma_vaccine: q(d, '#r-mv').value, ten: q(d, '#r-ten').value, tu_ngay: Math.round(Number(tu) * Number(q(d, '#r-tuu').value)), den_ngay: den === '' ? null : Math.round(Number(den) * Number(q(d, '#r-denu').value)), gioi_tinh: q(d, '#r-gt').value, phac_do: q(d, '#r-pd').value, cau_goi_y: q(d, '#r-ct').value, trang_thai: q(d, '#r-tt').value }, {}).then(function () { d.remove(); done('Đã lưu phác đồ') }, function (e) { b.disabled = false; fail(e) })
    }
  }

  /* ================= đăng ký trang và sự kiện ================= */
  X.pages.a_tk = pgTk; X.pages.a_tl = pgTl; X.pages.a_pd = pgPd;
  X.binders.push(function (on) {
    var all = function (s, fn) { [].slice.call(document.querySelectorAll(s)).forEach(fn) }, byAcc = function (ma) { return D.acc[ma] };
    on('#tk-new', 'onclick', function () { accModal(null) });
    on('#tk-q', 'oninput', function (x) { AS.tk.q = x.value; var p = x.selectionStart; render(); var n = $('#tk-q'); n.focus(); n.setSelectionRange(p, p) });
    on('#tk-shop', 'onchange', function (x) { AS.tk.shop = x.value; render() });
    all('[data-tkst]', function (b) { b.onclick = function () { AS.tk.st = b.dataset.tkst; render() } });
    all('[data-tkedit]', function (b) { b.onclick = function () { accModal(byAcc(b.dataset.tkedit)) } });
    all('[data-tkpw]', function (b) { b.onclick = function () { pwAsm(byAcc(b.dataset.tkpw)) } });
    all('[data-tklock]', function (b) { b.onclick = function () { lockAsm(byAcc(b.dataset.tklock)) } });
    all('[data-setql]', function (b) { b.onclick = function () { setQl(b.dataset.setql) } });
    all('[data-tltab]', function (b) { b.onclick = function () { AS.tl.tab = b.dataset.tltab; render() } });
    all('[data-tlg]', function (b) { b.onclick = function () { AS.tl.nhom = b.dataset.tlg; render() } });
    on('#tl-st', 'onchange', function (x) { AS.tl.st = x.value; render() });
    on('#tl-q', 'oninput', function (x) { AS.tl.q = x.value; var p = x.selectionStart; render(); var n = $('#tl-q'); n.focus(); n.setSelectionRange(p, p) });
    on('#tl-new', 'onclick', function () { tplModal(null) });
    all('[data-tpedit]', function (b) { b.onclick = function () { tplModal(D.tpl.filter(function (x) { return x.ma === b.dataset.tpedit })[0]) } });
    all('[data-tpok]', function (b) { b.onclick = function () { tplState(b.dataset.tpok, 'DANG_DUNG', 'Mẫu ' + b.dataset.tpok + ' đang dùng') } });
    all('[data-tpoff]', function (b) { b.onclick = function () { tplState(b.dataset.tpoff, 'NGUNG_DUNG', 'Đã ngừng mẫu ' + b.dataset.tpoff) } });
    on('#km-new', 'onclick', function () { promoModal(null) });
    all('[data-kmedit]', function (b) { b.onclick = function () { promoModal(D.promos.filter(function (x) { return x.ma === b.dataset.kmedit })[0]) } });
    on('#pd-new', 'onclick', function () { ruleModal(null) });
    all('[data-pdedit]', function (b) { b.onclick = function () { ruleModal(D.rules.filter(function (x) { return x.ma === b.dataset.pdedit })[0]) } });
    all('[data-pdok]', function (b) { b.onclick = function () { var r = D.rules.filter(function (x) { return x.ma === b.dataset.pdok })[0]; ruleSave(r, { p_trang_thai: 'DANG_DUNG' }).then(function () { return done('Phác đồ ' + r.ten + ' đang dùng') }, fail) } });
    all('[data-pdoff]', function (b) { b.onclick = function () { var r = D.rules.filter(function (x) { return x.ma === b.dataset.pdoff })[0]; ruleSave(r, { p_trang_thai: 'NGUNG_DUNG' }).then(function () { return done('Đã ngừng ' + r.ten) }, fail) } });
    on('#pd-tv', 'oninput', function (x) { AS.pd.tv = x.value; var p = x.selectionStart; render(); var n = $('#pd-tv'); n.focus(); n.setSelectionRange(p, p) });
    on('#pd-tu', 'onchange', function (x) { AS.pd.tu = x.value; render() }); on('#pd-tg', 'onchange', function (x) { AS.pd.tg = x.value; render() })
  })
})();
