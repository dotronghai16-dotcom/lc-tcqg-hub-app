/* MÀN ASM (bản thử) – dùng chung khung và dữ liệu với index.html (D, ME, S, rest, rpc, load, esc, ic, toast, ...).
 * ASM = giám đốc kinh doanh khu vực: xem toàn bộ, giao Data (đích danh hoặc cả Shop), bật quyền cho Quản lý chia, thu hồi, nạp khách, xuất dữ liệu.
 * Mọi thao tác ghi qua hàm RPC của cơ sở dữ liệu (chỉ vai trò ASM gọi được). Bản thử: dữ liệu giả, tính tổng hợp ngay trên máy (đủ cho vài trăm dòng). */
var ASM = (function () {
  IC.upload = '<path d="M12 16V4m0 0-4 4m4-4 4 4M5 20h14"/>'; IC.download = '<path d="M12 4v12m0 0-4-4m4 4 4-4M5 20h14"/>'; IC.swap = '<path d="M7 7h12l-3-3M17 17H5l3 3"/>'; IC.book = '<path d="M5 4h11a3 3 0 0 1 3 3v13H8a3 3 0 0 1-3-3zM5 17a3 3 0 0 1 3-3h11"/>'; IC.shop = '<path d="M4 9 5.5 4h13L20 9M4 9h16v11H4zM9 20v-6h6v6"/>';
  IC.more = '<circle cx="5" cy="12" r="1.5"/><circle cx="12" cy="12" r="1.5"/><circle cx="19" cy="12" r="1.5"/>'; IC.shield = '<path d="M12 3 4 6v6c0 5 3.5 8 8 9 4.5-1 8-4 8-9V6z"/><path d="m9 12 2 2 4-4"/>';
  var NAV = [['a_tq', 'Tổng quan', 'home'], ['a_kho', 'Kho Data', 'list'], ['a_pc', 'Phân công', 'swap'], ['a_nap', 'Nạp Data', 'upload'], ['a_bc', 'Kết quả & xuất', 'chart'], ['a_tk', 'Nhân sự & tài khoản', 'team'], ['a_tl', 'Thư viện', 'book'], ['a_pd', 'Phác đồ', 'shield'], ['acc', 'Của tôi', 'user']];
  var MOBILE = ['a_tq', 'a_kho', 'a_pc', 'a_tk'];
  var X = { pages: {}, binders: [], u: {} };
  var ST = { shop: '', band: '', f: '', q: '', pg: 1, sel: {}, selAll: false, imp: null, bcShop: '', bcSt: '' };
  var PAGE = 25;
  function today0() { var d = new Date(); d.setHours(0, 0, 0, 0); return d }
  function overdueDays(t) { if (!t.han_lien_he || !open_(t)) return 0; var h = new Date(t.han_lien_he + 'T00:00:00'), n = Math.floor((today0() - h) / 86400000); return n > 0 ? n : 0 }
  function ageMonths(c) { var b = c && c.ngay_sinh ? new Date(c.ngay_sinh + 'T00:00:00') : null; if (!b) return null; var n = new Date(); return (n.getFullYear() - b.getFullYear()) * 12 + n.getMonth() - b.getMonth() - (n.getDate() < b.getDate() ? 1 : 0) }
  function ageTxt(c) { var m = ageMonths(c); if (m === null) return ''; return m < 24 ? Math.max(0, m) + ' tháng' : Math.floor(m / 12) + ' tuổi' }
  var BANDS = [['', 'Mọi độ tuổi'], ['0', 'Dưới 2 tháng'], ['1', '2 – 6 tháng'], ['2', '6 – 12 tháng'], ['3', '1 – 2 tuổi'], ['4', '2 – 5 tuổi'], ['5', 'Trên 5 tuổi']];
  function bandOf(c) { var m = ageMonths(c); if (m === null) return ''; return m < 2 ? '0' : m < 6 ? '1' : m < 12 ? '2' : m < 24 ? '3' : m < 60 ? '4' : '5' }
  function shops() { return Object.keys(D.shops).map(function (k) { return D.shops[k] }).filter(function (s) { return s.trang_thai === 'HOAT_DONG' }).sort(function (a, b) { return a.ma_shop.localeCompare(b.ma_shop) }) }
  function staffOf(id) { return Object.keys(D.acc).map(function (k) { return D.acc[k] }).filter(function (a) { return a.id_shop === id && a.trang_thai === 'HOAT_DONG' && (a.vai_tro === 'NHAN_VIEN' || a.vai_tro === 'QUAN_LY') }) }
  function qlOf(id) { var q = staffOf(id).filter(function (a) { return a.vai_tro === 'QUAN_LY' })[0]; return q ? q.ho_ten : '–' }
  function turnsByCust() { var m = {}; D.turns.forEach(function (t) { (m[t.ma_khach] = m[t.ma_khach] || []).push(t) }); return m }
  function custState(ts) { if (!ts || !ts.length) return 'kho'; return ts.some(open_) ? 'dang' : 'xong' }
  function yes(t) { var c = cat(t.ma_ket_qua); return t.trang_thai === 'HOAN_THANH' && c && c.ma === 'DONG_Y' }
  function pct(a, b) { return b ? Math.round(a / b * 100) : 0 }
  function sel(id, opts, val) { return '<select id="' + id + '">' + opts.map(function (o) { return '<option value="' + esc(o[0]) + '"' + (String(o[0]) === String(val) ? ' selected' : '') + '>' + esc(o[1]) + '</option>' }).join('') + '</select>' }

  /* ---------- khung ---------- */
  function shell() {
    function b(x) { return '<button class="nav ' + (S.page === x[0] ? 'on' : '') + '" data-p="' + x[0] + '">' + ic(x[2]) + x[1] + '</button>' }
    $('#side').innerHTML = '<div class="brand">' + LOGO + '</div><div class="grp">Điều phối</div>' + NAV.slice(0, 3).map(b).join('') + '<div class="grp">Dữ liệu</div>' + NAV.slice(3, 5).map(b).join('') + '<div class="grp">Nhân sự &amp; nội dung</div>' + NAV.slice(5, 8).map(b).join('') + '<div class="grp">Hệ thống</div>' + NAV.slice(8).map(b).join('') +
      '<div class="me"><div class="ava">' + esc(ini(ME.ho_ten)) + '</div><div><b>' + esc(ME.ho_ten) + '</b><small>ASM · Giám đốc kinh doanh KV</small></div></div>';
    var mob = NAV.filter(function (x) { return MOBILE.indexOf(x[0]) >= 0 }), inMore = MOBILE.indexOf(S.page) < 0;
    $('#tabs').style.setProperty('--n', mob.length + 1);
    $('#tabs').innerHTML = mob.map(function (x) { return '<button class="tab ' + (S.page === x[0] ? 'on' : '') + '" data-p="' + x[0] + '">' + ic(x[2]) + x[1].replace('Nhân sự & tài khoản', 'Nhân sự') + '</button>' }).join('') + '<button class="tab ' + (inMore ? 'on' : '') + '" id="tmore">' + ic('more') + 'Thêm</button>';
    document.querySelectorAll('[data-p]').forEach(function (e) { e.onclick = function () { S.page = e.dataset.p; ST.pg = 1; ST.sel = {}; render() } });
    var tm = $('#tmore'); if (tm) tm.onclick = function () {
      var d = document.createElement('div'); d.className = 'modal'; d.innerHTML = '<div class="card" style="width:100%;max-width:520px;border-radius:22px 22px 0 0;padding:14px 14px calc(14px + env(safe-area-inset-bottom))"><h2 style="font-size:17px;margin:2px 4px 8px">Tất cả mục</h2>' + NAV.map(function (x) { return '<button class="nav ' + (S.page === x[0] ? 'on' : '') + '" data-mp="' + x[0] + '">' + ic(x[2]) + x[1] + '</button>' }).join('') + '</div>';
      document.body.appendChild(d); d.onclick = function (e) { if (e.target === d) d.remove() };
      [].slice.call(d.querySelectorAll('[data-mp]')).forEach(function (b) { b.onclick = function () { S.page = b.dataset.mp; ST.pg = 1; ST.sel = {}; d.remove(); render() } })
    }
  }
  function head(t, p, extra) { return mtop() + '<div class="row2"><div class="greet"><h1>' + t + '</h1>' + (p ? '<p>' + p + '</p>' : '') + '</div>' + (extra || '') + '</div>' }

  /* ---------- Tổng quan ---------- */
  function statsFor(shopId) {
    var T = D.turns.filter(function (t) { return !shopId || t.id_shop === shopId }), o = { luot: T.length, mo: 0, xong: 0, dy: 0, qh: 0, chia: 0 };
    T.forEach(function (t) { if (open_(t)) o.mo++; if (t.trang_thai === 'HOAN_THANH') o.xong++; if (yes(t)) o.dy++; if (overdueDays(t) > 0) o.qh++; if (t.recipient_type === 'SHOP' && t.trang_thai === 'DA_GIAO') o.chia++ });
    return o
  }
  function pgTq() {
    var tc = turnsByCust(), kho = 0, tot = Object.keys(D.cus).length; Object.keys(D.cus).forEach(function (k) { if (custState(tc[k]) === 'kho') kho++ });
    var all = statsFor(ST.shop), first = ME.ho_ten.trim().split(/\s+/).pop(), SH = shops();
    var h = head('Chào anh ' + esc(first) + ', khu vực có <span style="color:var(--teal)">' + SH.length + ' Shop</span>', 'Tổng hợp toàn khu vực · ' + new Date().toLocaleDateString('vi-VN', { weekday: 'long', day: 'numeric', month: 'long' }),
      '<div style="display:flex;gap:8px;align-items:center"><span class="meta">Shop:</span>' + sel('fshop', [['', 'Tất cả Shop']].concat(SH.map(function (s) { return [s.id_shop, s.ten_shop] })), ST.shop) + '</div>');
    h += '<div class="stats" style="grid-template-columns:repeat(6,1fr)">' + [['Khách trong kho', tot], ['Chưa giao', kho], ['Đang giữ', all.mo], ['Hoàn thành', all.xong], ['Đồng ý tiêm', all.dy], ['Quá hạn gọi', all.qh]].map(function (x, i) { return '<div class="card st"><div class="l">' + x[0] + '</div><div class="v"' + (i === 5 && x[1] ? ' style="color:var(--bad)"' : '') + '>' + x[1] + '</div></div>' }).join('') + '</div>';
    var al = '';
    if (kho) al += '<div class="alert"><i>▦</i><div><b>' + kho + ' khách chưa được giao</b><small>Đang nằm trong kho</small></div><button class="btn b2" data-go="a_kho" style="min-height:38px;margin-left:auto">Phân công</button></div>';
    if (all.chia) al += '<div class="alert a-warn"><i>⇄</i><div><b>' + all.chia + ' lượt giao cho Shop chờ chia</b><small>Quản lý chia khi bạn bật quyền cho lần giao đó</small></div><button class="btn b2" data-go="a_pc" style="min-height:38px;margin-left:auto">Xem lần giao</button></div>';
    if (all.qh) al += '<div class="alert a-bad"><i>⚠</i><div><b>' + all.qh + ' lượt quá hạn liên hệ</b><small>Tính từ "Ngày cần liên hệ" của từng lượt</small></div></div>';
    h += '<div class="card"><h2>Cần chú ý</h2>' + (al || '<div class="alert"><i>✓</i><div><b>Chưa có gì cần chú ý</b></div></div>') + '</div>';
    h += '<div class="card" style="padding:6px 8px"><h2 style="padding:10px 10px 0">So sánh các Shop <span class="pill">bấm một Shop để xem nhân viên</span></h2><div class="tw"><table><thead><tr><th>Shop</th><th>Quản lý</th><th>NV</th><th>Được giao</th><th>Đang mở</th><th>Hoàn thành</th><th>Đồng ý tiêm</th><th style="min-width:120px">Tiến độ</th><th>Quá hạn</th></tr></thead><tbody>' +
      SH.map(function (s) { var x = statsFor(s.id_shop); return '<tr class="click" data-shop="' + esc(s.id_shop) + '"' + (ST.shop === s.id_shop ? ' style="background:var(--tint)"' : '') + '><td><b>' + esc(s.ten_shop) + '</b><small>' + esc(s.ma_shop) + '</small></td><td>' + esc(qlOf(s.id_shop)) + '</td><td>' + staffOf(s.id_shop).length + '</td><td>' + x.luot + '</td><td>' + x.mo + '</td><td>' + x.xong + '</td><td>' + x.dy + '</td><td><div class="bar" style="height:6px;border-radius:99px;background:var(--bg);overflow:hidden"><i style="display:block;height:100%;width:' + pct(x.xong, x.luot) + '%;background:var(--teal)"></i></div><small>' + pct(x.xong, x.luot) + '%</small></td><td>' + (x.qh ? '<span class="tag t-late">' + x.qh + '</span>' : '0') + '</td></tr>' }).join('') + '</tbody></table></div></div>';
    if (ST.shop) {
      var staff = staffOf(ST.shop);
      h += '<div class="card" style="padding:6px 8px"><h2 style="padding:10px 10px 0">Nhân viên · ' + esc(shopName(ST.shop)) + '</h2><div class="tw"><table><thead><tr><th>Nhân viên</th><th>Chức danh</th><th>Được giao</th><th>Đang mở</th><th>Hoàn thành</th><th>Đồng ý tiêm</th><th>Quá hạn</th></tr></thead><tbody>' + staff.map(function (a) {
        var T = D.turns.filter(function (t) { return t.recipient_type === 'EMPLOYEE' && t.ma_tk_nhan === a.ma_tk }), mo = T.filter(open_).length, xx = T.filter(function (t) { return t.trang_thai === 'HOAN_THANH' }).length, dy = T.filter(yes).length, qh = T.filter(function (t) { return overdueDays(t) > 0 }).length;
        return '<tr><td><b>' + esc(a.ho_ten) + '</b></td><td>' + (a.vai_tro === 'QUAN_LY' ? 'Quản lý' : 'Nhân viên') + '</td><td>' + T.length + '</td><td>' + mo + '</td><td>' + xx + '</td><td>' + dy + '</td><td>' + (qh ? '<span class="tag t-late">' + qh + '</span>' : '0') + '</td></tr>' }).join('') + '</tbody></table></div></div>';
    }
    return h
  }

  /* ---------- Kho Data ---------- */
  function khoRows() {
    var tc = turnsByCust(), q = ST.q.trim().toLowerCase();
    return Object.keys(D.cus).map(function (k) { var c = D.cus[k], ts = tc[k] || [], l = (D.links[k] || [])[0] || {}; return { c: c, ts: ts, st: custState(ts), l: l } }).filter(function (r) {
      if (ST.f && r.st !== ST.f) return false; if (ST.band && bandOf(r.c) !== ST.band) return false;
      if (q && (r.c.ma_khach + ' ' + r.c.ma_tcqg + ' ' + r.c.ho_ten + ' ' + (r.l.sdt || '') + ' ' + (r.l.ho_ten || '')).toLowerCase().indexOf(q) < 0) return false; return true
    }).sort(function (a, b) { return a.c.ma_khach.localeCompare(b.c.ma_khach) })
  }
  function pgKho() {
    var all = khoRows(), base = Object.keys(D.cus).map(function (k) { return custState(turnsByCust()[k]) }), cnt = { '': base.length, kho: 0, dang: 0, xong: 0 }; base.forEach(function (s) { cnt[s]++ });
    var pages = Math.max(1, Math.ceil(all.length / PAGE)); ST.pg = Math.min(ST.pg, pages); var rows = all.slice((ST.pg - 1) * PAGE, ST.pg * PAGE), n = Object.keys(ST.sel).length;
    var h = head('Kho Data', 'Toàn bộ khách. Chọn khách rồi bấm Phân công.', '');
    h += '<div class="top"><div class="search">' + ic('search') + '<input id="kq" placeholder="Tìm tên, mã TCQG, số điện thoại…" value="' + esc(ST.q) + '"></div>' + sel('fband', BANDS, ST.band) + '</div>';
    h += '<div class="chips">' + [['', 'Tất cả'], ['kho', 'Chưa giao'], ['dang', 'Đang giao'], ['xong', 'Đã xong']].map(function (x) { return '<button class="chip ' + (ST.f === x[0] ? 'on' : '') + '" data-f="' + x[0] + '">' + x[1] + '<em>' + cnt[x[0]] + '</em></button>' }).join('') + '</div>';
    h += '<div class="card" style="padding:6px 8px"><div class="tw"><table><thead><tr><th style="width:36px"><input type="checkbox" id="selpg" ' + (rows.length && rows.every(function (r) { return ST.sel[r.c.ma_khach] }) ? 'checked' : '') + ' style="width:20px;height:20px"></th><th>Khách</th><th>Tuổi</th><th>Phụ huynh</th><th>Trạng thái</th><th>Đang giữ</th></tr></thead><tbody>' +
      rows.map(function (r) { var op = r.ts.filter(open_)[0], who = op ? (op.recipient_type === 'SHOP' ? shopName(op.id_shop) + ' (Shop)' : nm(op.ma_tk_nhan)) : '';
        return '<tr><td><input type="checkbox" data-k="' + esc(r.c.ma_khach) + '" ' + (ST.sel[r.c.ma_khach] ? 'checked' : '') + ' style="width:20px;height:20px"></td><td><b>' + esc(r.c.ho_ten) + '</b><small>' + esc(r.c.ma_tcqg || r.c.ma_khach) + '</small></td><td>' + esc(ageTxt(r.c)) + '</td><td>' + esc(r.l.ho_ten || '') + '<small>' + esc(r.l.sdt || '') + '</small></td><td>' + ({ kho: '<span class="tag t-tint">Chưa giao</span>', dang: '<span class="tag t-late">Đang giao</span>', xong: '<span class="tag t-ok">Đã xong</span>' })[r.st] + '</td><td>' + esc(who) + '</td></tr>' }).join('') +
      (rows.length ? '' : '<tr><td colspan="6" class="empty">Không có khách phù hợp.</td></tr>') + '</tbody></table></div>' +
      '<div class="row2" style="padding:10px"><span class="meta">' + all.length + ' khách · trang ' + ST.pg + '/' + pages + '</span><div style="display:flex;gap:8px"><button class="btn b3" data-pg="-1" style="min-height:38px">‹</button><button class="btn b3" data-pg="1" style="min-height:38px">›</button></div></div></div>';
    h += '<div class="bulk" style="' + (n ? '' : 'display:none') + '"><b>Đã chọn ' + n + ' khách</b><span class="sp"></span>' + (all.length > n ? '<button class="btn b3" id="seltop" style="min-height:42px">Chọn cả ' + all.length + ' khách đang lọc</button>' : '') + '<button class="btn b3" id="selclr" style="min-height:42px">Bỏ chọn</button><button class="btn b1" id="asg" style="min-height:42px">Phân công</button></div>';
    return h
  }
  function assignModal() {
    var ks = Object.keys(ST.sel); if (!ks.length) return; var tc = turnsByCust(), openN = ks.filter(function (k) { return custState(tc[k]) === 'dang' }).length;
    var topics = D.dm ? Object.keys(D.dm).map(function (k) { return D.dm[k] }).filter(function (c) { return c.loai === 'CHU_DE' && c.dang_dung }) : [];
    var m = { kieu: 'NHAN_VIEN', who: {}, shopSel: {}, ql: false, giaolai: false };
    var d = document.createElement('div'); d.className = 'modal';
    function draw() {
      var SH = shops(), nWho = Object.keys(m.who).length, nShop = Object.keys(m.shopSel).length, n = m.kieu === 'SHOP' ? nShop : nWho, per = n ? Math.ceil((ks.length - (m.giaolai ? 0 : openN)) / n) : 0;
      d.innerHTML = '<div class="card" style="width:100%;max-width:640px;border-radius:22px 22px 0 0;padding:18px;max-height:92vh;overflow:auto"><h2 style="font-size:19px;margin:0 0 4px">Phân công ' + ks.length + ' khách</h2><div class="meta" style="margin-bottom:12px">Chia đều cho những người hoặc Shop bạn chọn. Mỗi khách một lượt giao riêng, giữ nguyên mã khách.</div>' +
        '<div class="chips" style="margin-bottom:10px"><button class="chip ' + (m.kieu === 'NHAN_VIEN' ? 'on' : '') + '" data-kieu="NHAN_VIEN">Đích danh nhân viên</button><button class="chip ' + (m.kieu === 'SHOP' ? 'on' : '') + '" data-kieu="SHOP">Cả Shop</button></div>' +
        (m.kieu === 'NHAN_VIEN' ? SH.map(function (s) { var st = staffOf(s.id_shop); return st.length ? '<div class="sec"><h3>' + esc(s.ten_shop) + ' <a href="#" data-allshop="' + esc(s.id_shop) + '" style="text-transform:none;letter-spacing:0;font-weight:600;color:var(--teal-d)">chọn cả Shop</a></h3><div style="display:grid;gap:6px">' + st.map(function (a) { return '<label class="row" style="cursor:pointer"><input type="checkbox" data-w="' + esc(a.ma_tk) + '" ' + (m.who[a.ma_tk] ? 'checked' : '') + ' style="width:20px;height:20px;accent-color:var(--teal)"><span>' + esc(a.ho_ten) + ' <span class="meta">· ' + (a.vai_tro === 'QUAN_LY' ? 'Quản lý' : 'Nhân viên') + '</span></span></label>' }).join('') + '</div></div>' : '' }).join('')
          : '<div class="sec"><h3>Chọn Shop nhận</h3><div style="display:grid;gap:6px">' + SH.map(function (s) { return '<label class="row" style="cursor:pointer"><input type="checkbox" data-sh="' + esc(s.id_shop) + '" ' + (m.shopSel[s.id_shop] ? 'checked' : '') + ' style="width:20px;height:20px;accent-color:var(--teal)"><span>' + esc(s.ten_shop) + ' <span class="meta">· Quản lý ' + esc(qlOf(s.id_shop)) + '</span></span></label>' }).join('') + '</div></div>' +
            '<label class="sw"><input type="checkbox" id="mql" ' + (m.ql ? 'checked' : '') + '><div><b>Cho phép Quản lý chia lại cho nhân viên</b><small>Chỉ áp cho lần giao này. Bạn tắt được bất cứ lúc nào ở mục Phân công.</small></div></label>') +
        '<div class="form2"><div><label class="meta">Chủ đề</label>' + sel('mcd', [['', '— chọn —']].concat(topics.map(function (c) { return [c.ma_dm, c.ten] })), '') + '</div><div><label class="meta">Hạn liên hệ</label><input type="date" id="mhan" value="' + new Date().toISOString().slice(0, 10) + '"></div></div>' +
        (openN ? '<label class="sw" style="margin-top:8px"><input type="checkbox" id="mgl" ' + (m.giaolai ? 'checked' : '') + '><div><b>Vẫn giao ' + openN + ' khách đang có lượt mở</b><small>Mặc định bỏ qua để tránh nhiều người cùng gọi một khách. Bật là giao ngoại lệ.</small></div></label>' : '') +
        '<div class="alert" style="margin-top:10px"><i>Σ</i><div><b>' + (n ? 'Giao ' + (ks.length - (m.giaolai ? 0 : openN)) + ' khách cho ' + n + (m.kieu === 'SHOP' ? ' Shop' : ' người') + ' · khoảng ' + per + ' khách mỗi nơi' : 'Chưa chọn người nhận') + '</b></div></div>' +
        '<div class="acts" style="margin-top:6px"><button class="btn b3" id="mx" style="flex:1">Hủy</button><button class="btn b1" id="mok" style="flex:1.4" ' + (n ? '' : 'disabled') + '>Xác nhận giao</button></div></div>';
      d.querySelectorAll('[data-kieu]').forEach(function (b) { b.onclick = function () { m.kieu = b.dataset.kieu; draw() } });
      d.querySelectorAll('[data-w]').forEach(function (b) { b.onchange = function () { b.checked ? m.who[b.dataset.w] = 1 : delete m.who[b.dataset.w]; draw() } });
      d.querySelectorAll('[data-sh]').forEach(function (b) { b.onchange = function () { b.checked ? m.shopSel[b.dataset.sh] = 1 : delete m.shopSel[b.dataset.sh]; draw() } });
      d.querySelectorAll('[data-allshop]').forEach(function (a) { a.onclick = function (e) { e.preventDefault(); staffOf(a.dataset.allshop).forEach(function (x) { m.who[x.ma_tk] = 1 }); draw() } });
      var q = d.querySelector('#mql'); if (q) q.onchange = function () { m.ql = q.checked; draw() }; var g = d.querySelector('#mgl'); if (g) g.onchange = function () { m.giaolai = g.checked; draw() };
      d.querySelector('#mx').onclick = function () { d.remove() };
      d.querySelector('#mok').onclick = function () {
        var cd = d.querySelector('#mcd').value, han = d.querySelector('#mhan').value, btn = d.querySelector('#mok'); btn.disabled = true; btn.textContent = 'Đang giao…';
        rpc('u_asm_assign', { p_ma_khachs: ks, p_kieu: m.kieu, p_nguoi_nhan: Object.keys(m.kieu === 'SHOP' ? m.shopSel : m.who), p_chu_de: cd, p_muc_dich: cd && D.dm[cd] ? D.dm[cd].ten : 'Giao Data', p_han: han || null, p_cho_ql_chia: m.ql, p_giao_lai: m.giaolai, p_req: rid() })
          .then(function (r) { d.remove(); ST.sel = {}; toast('Đã giao ' + r.so_luot + ' khách (' + r.ma_lan_pc + ')' + (r.bo_qua ? ', bỏ qua ' + r.bo_qua + ' khách đang có lượt mở' : '')); return load(true) }).then(render, function (e) { btn.disabled = false; btn.textContent = 'Xác nhận giao'; fail(e) })
      }
    }
    draw(); document.body.appendChild(d); d.onclick = function (e) { if (e.target === d) d.remove() }
  }

  /* ---------- Phân công (các lần giao) ---------- */
  function runsInfo() {
    var m = {}; D.turns.forEach(function (t) { var r = m[t.ma_lan_pc] = m[t.ma_lan_pc] || { ma: t.ma_lan_pc, n: 0, mo: 0, xong: 0, shopN: 0, empN: 0, chua: 0 }; r.n++; if (open_(t)) r.mo++; if (t.trang_thai === 'HOAN_THANH') r.xong++; if (t.recipient_type === 'SHOP') { r.shopN++; if (t.trang_thai === 'DA_GIAO') r.chua++ } else r.empN++ });
    return Object.keys(m).map(function (k) { var r = m[k], run = D.runs[k] || {}; r.run = run; return r }).sort(function (a, b) { return String(b.run.thoi_gian || '').localeCompare(String(a.run.thoi_gian || '')) })
  }
  function pgPc() {
    var R = runsInfo(), h = head('Phân công', 'Các lần giao Data. Bật quyền cho Quản lý chia với lần giao cho cả Shop.', '');
    h += '<div class="card" style="padding:6px 8px"><div class="tw"><table><thead><tr><th>Lần giao</th><th>Chủ đề</th><th>Ngày giao</th><th>Kiểu</th><th>Lượt</th><th>Đang mở</th><th>Hoàn thành</th><th>Quản lý được chia</th><th></th></tr></thead><tbody>' + R.map(function (r) {
      var run = r.run, d = loc(run.thoi_gian), kieu = r.shopN && r.empN ? 'Đích danh + Shop' : r.shopN ? 'Cả Shop' : 'Đích danh';
      return '<tr><td><b>' + esc(r.ma) + '</b></td><td>' + esc((D.dm[run.ma_chu_de] || {}).ten || run.muc_dich || '') + '</td><td>' + dstr(d) + '</td><td>' + kieu + '</td><td>' + r.n + '</td><td>' + r.mo + '</td><td>' + r.xong + '</td><td>' + (r.shopN ? '<label class="row" style="cursor:pointer"><input type="checkbox" data-ql="' + esc(r.ma) + '" ' + (run.cho_ql_chia ? 'checked' : '') + ' style="width:22px;height:22px;accent-color:var(--teal)"><span class="meta">' + (run.cho_ql_chia ? 'Đang bật' : 'Tắt') + (r.chua ? ' · ' + r.chua + ' chờ chia' : '') + '</span></label>' : '<span class="meta">–</span>') + '</td><td>' + (r.mo ? '<button class="btn b3" data-rev="' + esc(r.ma) + '" style="min-height:36px;font-size:13px">Thu hồi lượt mở</button>' : '') + '</td></tr>' }).join('') + '</tbody></table></div></div>';
    return h
  }
  function revoke(ma) {
    var T = D.turns.filter(function (t) { return t.ma_lan_pc === ma && open_(t) && t.trang_thai === 'DA_GIAO' }); if (!T.length) { toast('Không có lượt nào chưa xử lý để thu hồi.', 0, 1); return }
    var d = document.createElement('div'); d.className = 'modal'; d.innerHTML = '<div class="card" style="width:100%;max-width:480px;border-radius:22px 22px 0 0;padding:18px"><h2 style="font-size:18px;margin:0 0 6px">Thu hồi ' + T.length + ' lượt chưa xử lý</h2><div class="meta" style="margin-bottom:10px">Lượt đã bắt đầu xử lý giữ nguyên. Lịch sử luôn được giữ. Người đang giữ lượt sẽ không ghi thêm được.</div><label class="field">Lý do (không bắt buộc)<input id="rl" placeholder="Ví dụ: giao lại cho Shop khác"></label><div class="acts"><button class="btn b3" id="rx" style="flex:1">Hủy</button><button class="btn b1" id="rok" style="flex:1.4">Xác nhận thu hồi</button></div></div>';
    document.body.appendChild(d); d.onclick = function (e) { if (e.target === d) d.remove() }; d.querySelector('#rx').onclick = function () { d.remove() };
    d.querySelector('#rok').onclick = function () { var b = d.querySelector('#rok'); b.disabled = true; rpc('u_asm_revoke', { p_ma_luots: T.map(function (t) { return t.ma_luot }), p_ly_do: d.querySelector('#rl').value, p_req: rid() }).then(function (r) { d.remove(); toast('Đã thu hồi ' + r.so_luot + ' lượt'); return load(true) }).then(render, function (e) { b.disabled = false; fail(e) }) }
  }

  /* ---------- Nạp Data ---------- */
  var TPL = 'ma_tcqg,ho_ten,ngay_sinh,gioi_tinh,ten_lien_he,quan_he,sdt\nTC900001,Nguyễn Văn An,15/03/2026,M,Nguyễn Thị Lan,Mẹ,0900123456\nTC900002,Trần Bảo Ngọc,2026-01-20,F,Trần Văn Hùng,Bố,0900654321\n';
  function toISO(s) { s = String(s || '').trim(); var m; if (/^\d{4}-\d{2}-\d{2}$/.test(s)) m = s; else if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/))) m = m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2); else return null; var d = new Date(m + 'T00:00:00'); if (isNaN(d) || d > new Date()) return null; var pp = m.split('-'); return (d.getFullYear() === +pp[0] && d.getMonth() + 1 === +pp[1] && d.getDate() === +pp[2]) ? m : null }
  function parseCSV(text) {
    var lines = text.replace(/\r/g, '').split('\n').filter(function (l) { return l.trim() }); if (lines.length < 2) return [];
    var del = (lines[0].indexOf('\t') >= 0) ? '\t' : (lines[0].split(';').length > lines[0].split(',').length ? ';' : ','), hdr = split(lines[0]).map(function (x) { return x.trim().toLowerCase() });
    function split(l) { var out = [], cur = '', q = false; for (var i = 0; i < l.length; i++) { var ch = l[i]; if (ch === '"') { if (q && l[i + 1] === '"') { cur += '"'; i++ } else q = !q } else if (ch === del && !q) { out.push(cur); cur = '' } else cur += ch } out.push(cur); return out }
    return lines.slice(1).map(function (l, i) { var v = split(l), o = { _dong: i + 2 }; hdr.forEach(function (h, j) { o[h] = (v[j] || '').trim() }); return o })
  }
  function checkRows(rows) {
    var seen = {}, have = {}; Object.keys(D.cus).forEach(function (k) { if (D.cus[k].ma_tcqg) have[D.cus[k].ma_tcqg] = 1 });
    return rows.map(function (r) { var e = ''; r.iso = toISO(r.ngay_sinh);
      if (!r.ho_ten) e = 'Thiếu họ tên'; else if (!r.iso) e = 'Ngày sinh không hợp lệ'; else if (r.ma_tcqg && have[r.ma_tcqg]) e = 'Đã có trong kho (bỏ qua)'; else if (r.ma_tcqg && seen[r.ma_tcqg]) e = 'Trùng Mã TCQG trong file'; if (r.ma_tcqg) seen[r.ma_tcqg] = 1; r.err = e; return r })
  }
  function pgNap() {
    var h = head('Nạp Data', 'Dán nội dung file CSV (hoặc chọn file). App kiểm tra, xem trước, rồi mới nạp. Khách trùng Mã TCQG sẽ được bỏ qua.', '');
    h += '<div class="card"><div class="acts" style="margin-bottom:10px"><button class="btn b3" id="tpl">' + ic('download') + 'Tải file mẫu</button><label class="btn b2" style="cursor:pointer">' + ic('upload') + 'Chọn file CSV<input type="file" id="nfile" accept=".csv,.txt" hidden></label></div><textarea id="ntx" style="min-height:140px;font-family:ui-monospace,Menlo,monospace;font-size:13px" placeholder="ma_tcqg,ho_ten,ngay_sinh,gioi_tinh,ten_lien_he,quan_he,sdt">' + esc(ST.impText || '') + '</textarea><div class="acts" style="margin-top:10px"><button class="btn b1" id="nchk">Kiểm tra</button></div></div>';
    if (ST.imp) {
      var ok = ST.imp.filter(function (r) { return !r.err }), dup = ST.imp.filter(function (r) { return /bỏ qua/.test(r.err) }).length, bad = ST.imp.length - ok.length - dup;
      h += '<div class="stats" style="grid-template-columns:repeat(4,1fr)"><div class="card st"><div class="l">Dòng trong file</div><div class="v">' + ST.imp.length + '</div></div><div class="card st"><div class="l">Hợp lệ</div><div class="v" style="color:var(--ok)">' + ok.length + '</div></div><div class="card st"><div class="l">Đã có (bỏ qua)</div><div class="v">' + dup + '</div></div><div class="card st"><div class="l">Lỗi</div><div class="v" style="color:var(--bad)">' + bad + '</div></div></div>' +
        '<div class="card" style="padding:6px 8px"><div class="tw"><table><thead><tr><th>Dòng</th><th>Mã TCQG</th><th>Họ tên</th><th>Ngày sinh</th><th>Liên hệ</th><th>Kết quả kiểm tra</th></tr></thead><tbody>' + ST.imp.slice(0, 30).map(function (r) { return '<tr><td>' + r._dong + '</td><td>' + esc(r.ma_tcqg) + '</td><td>' + esc(r.ho_ten) + '</td><td>' + esc(r.ngay_sinh) + '</td><td>' + esc(r.ten_lien_he) + '<small>' + esc(r.sdt) + '</small></td><td>' + (r.err ? '<span class="tag ' + (/bỏ qua/.test(r.err) ? 't-new' : 't-late') + '">' + esc(r.err) + '</span>' : '<span class="tag t-ok">Hợp lệ</span>') + '</td></tr>' }).join('') + '</tbody></table></div></div>' +
        '<div class="acts"><button class="btn b1" id="nok" ' + (ok.length ? '' : 'disabled') + '>Nạp ' + ok.length + ' khách hợp lệ</button></div>';
    }
    if (ST.impRes) h += '<div class="alert"><i>✓</i><div><b>Đã nạp xong</b><small>Thêm mới ' + ST.impRes.them_moi + ' · bỏ qua do trùng ' + ST.impRes.trung_bo_qua + ' · lỗi ' + ST.impRes.loi.length + '</small></div><button class="btn b2" data-go="a_kho" style="min-height:38px;margin-left:auto">Mở Kho Data</button></div>';
    return h
  }
  function download(name, text) { var a = document.createElement('a'); a.href = URL.createObjectURL(new Blob(['﻿' + text], { type: 'text/csv;charset=utf-8' })); a.download = name; document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove() }, 500) }
  function csvCell(v) { v = String(v == null ? '' : v); return /[",\n;]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v }

  /* ---------- Kết quả & xuất ---------- */
  function bcTurns() { return D.turns.filter(function (t) { return (!ST.bcShop || t.id_shop === ST.bcShop) && (!ST.bcSt || (ST.bcSt === 'MO' ? open_(t) : t.trang_thai === ST.bcSt)) }) }
  function pgBc() {
    var T = bcTurns(), SH = shops(), by = {};
    T.forEach(function (t) { var k = t.recipient_type === 'EMPLOYEE' ? t.ma_tk_nhan : 'SHOP:' + t.id_shop, r = by[k] = by[k] || { k: k, t: t, n: 0, mo: 0, xong: 0, dy: 0 }; r.n++; if (open_(t)) r.mo++; if (t.trang_thai === 'HOAN_THANH') r.xong++; if (yes(t)) r.dy++ });
    var h = head('Kết quả & xuất', 'Xem kết quả theo nhân viên và xuất dữ liệu ra file CSV (mở được bằng Excel).', '');
    h += '<div class="top" style="flex-wrap:wrap">' + sel('bshop', [['', 'Tất cả Shop']].concat(SH.map(function (s) { return [s.id_shop, s.ten_shop] })), ST.bcShop) + sel('bst', [['', 'Mọi trạng thái'], ['MO', 'Đang mở'], ['HOAN_THANH', 'Hoàn thành'], ['THU_HOI', 'Đã thu hồi']], ST.bcSt) + '<span class="sp"></span><button class="btn b1" id="xl">' + ic('download') + 'Xuất theo lượt (' + T.length + ')</button><button class="btn b3" id="xn">Xuất theo nhân viên</button></div>';
    h += '<div class="card" style="padding:6px 8px"><div class="tw"><table><thead><tr><th>Người nhận</th><th>Shop</th><th>Được giao</th><th>Đang mở</th><th>Hoàn thành</th><th>Đồng ý tiêm</th><th>Tỷ lệ đồng ý / hoàn thành</th></tr></thead><tbody>' + Object.keys(by).map(function (k) { return by[k] }).sort(function (a, b) { return b.n - a.n }).map(function (r) {
      return '<tr><td><b>' + esc(r.k.indexOf('SHOP:') === 0 ? 'Cả Shop (chưa chia)' : nm(r.k)) + '</b></td><td>' + esc(shopName(r.t.id_shop)) + '</td><td>' + r.n + '</td><td>' + r.mo + '</td><td>' + r.xong + '</td><td>' + r.dy + '</td><td>' + (r.xong ? pct(r.dy, r.xong) + '%' : '–') + '</td></tr>' }).join('') + '</tbody></table></div></div>';
    return h
  }
  function exportTurns() {
    var T = bcTurns(), head = ['ma_luot', 'ma_khach', 'ma_tcqg', 'ho_ten', 'shop', 'nguoi_nhan', 'chu_de', 'trang_thai', 'ket_qua', 'ngay_giao', 'han_lien_he', 'hoan_thanh_luc'];
    var rows = T.map(function (t) { var c = D.cus[t.ma_khach] || {}; return [t.ma_luot, t.ma_khach, c.ma_tcqg, c.ho_ten, shopName(t.id_shop), t.recipient_type === 'SHOP' ? 'Cả Shop' : nm(t.ma_tk_nhan), (cat(t.ma_chu_de) || {}).ten, t.trang_thai, (cat(t.ma_ket_qua) || {}).ten || '', t.thoi_gian_giao_ban_dau, t.han_lien_he, t.thoi_gian_hoan_thanh || ''].map(csvCell).join(',') });
    download('ket-qua-theo-luot-' + new Date().toISOString().slice(0, 10) + '.csv', head.join(',') + '\n' + rows.join('\n')); toast('Đã xuất ' + T.length + ' lượt')
  }
  function exportStaff() {
    var by = {}; bcTurns().forEach(function (t) { var k = t.recipient_type === 'EMPLOYEE' ? t.ma_tk_nhan : 'SHOP:' + t.id_shop, r = by[k] = by[k] || { nm: k.indexOf('SHOP:') === 0 ? 'Cả Shop (chưa chia)' : nm(k), shop: shopName(t.id_shop), n: 0, mo: 0, xong: 0, dy: 0 }; r.n++; if (open_(t)) r.mo++; if (t.trang_thai === 'HOAN_THANH') r.xong++; if (yes(t)) r.dy++ });
    download('ket-qua-theo-nhan-vien-' + new Date().toISOString().slice(0, 10) + '.csv', 'nguoi_nhan,shop,duoc_giao,dang_mo,hoan_thanh,dong_y_tiem\n' + Object.keys(by).map(function (k) { var r = by[k]; return [r.nm, r.shop, r.n, r.mo, r.xong, r.dy].map(csvCell).join(',') }).join('\n')); toast('Đã xuất theo nhân viên')
  }

  /* ---------- Thư viện, Tài khoản ---------- */
  function pgAcc() { return head('Tài khoản của tôi', '', '') + '<div class="card"><div class="alert"><div class="ava">' + esc(ini(ME.ho_ten)) + '</div><div><b>' + esc(ME.ho_ten) + '</b><small>' + esc(ME.email) + ' · ASM</small></div></div><button class="btn b3" id="cpw" style="width:100%;margin-bottom:8px">Đổi mật khẩu</button><button class="btn b3" id="out" style="width:100%">Đăng xuất</button></div>' }
  X.u = { head: head, sel: sel, shops: shops, staffOf: staffOf, qlOf: qlOf, pct: pct, ST: ST };

  /* ---------- vẽ và gắn sự kiện ---------- */
  function render() {
    shell(); var m = $('#main'), p = S.page; m.innerHTML = ({ a_tq: pgTq, a_kho: pgKho, a_pc: pgPc, a_nap: pgNap, a_bc: pgBc, acc: pgAcc }[p] || X.pages[p] || pgTq)();
    var x, on = function (sel, ev, fn) { var el = $(sel); if (el) el[ev] = function (e) { return fn.call(el, el, e) } };
    document.querySelectorAll('[data-go]').forEach(function (b) { b.onclick = function () { S.page = b.dataset.go; render() } });
    on('#fshop', 'onchange', function (x) { ST.shop = x.value; render() });
    document.querySelectorAll('tr[data-shop]').forEach(function (r) { r.onclick = function () { ST.shop = ST.shop === r.dataset.shop ? '' : r.dataset.shop; render() } });
    on('#kq', 'oninput', function (x) { ST.q = x.value; ST.pg = 1; var pos = x.selectionStart; render(); var n = $('#kq'); n.focus(); n.setSelectionRange(pos, pos) });
    on('#fband', 'onchange', function (x) { ST.band = x.value; ST.pg = 1; render() });
    document.querySelectorAll('.chip[data-f]').forEach(function (b) { b.onclick = function () { ST.f = b.dataset.f; ST.pg = 1; render() } });
    document.querySelectorAll('[data-pg]').forEach(function (b) { b.onclick = function () { ST.pg = Math.max(1, ST.pg + Number(b.dataset.pg)); render() } });
    document.querySelectorAll('input[data-k]').forEach(function (b) { b.onchange = function () { b.checked ? ST.sel[b.dataset.k] = 1 : delete ST.sel[b.dataset.k]; render() } });
    on('#selpg', 'onchange', function (x) { var all = khoRows(), rows = all.slice((ST.pg - 1) * PAGE, ST.pg * PAGE); rows.forEach(function (r) { x.checked ? ST.sel[r.c.ma_khach] = 1 : delete ST.sel[r.c.ma_khach] }); render() });
    on('#seltop', 'onclick', function () { khoRows().forEach(function (r) { ST.sel[r.c.ma_khach] = 1 }); render() });
    on('#selclr', 'onclick', function () { ST.sel = {}; render() }); on('#asg', 'onclick', assignModal);
    document.querySelectorAll('input[data-ql]').forEach(function (b) { b.onchange = function () { b.disabled = true; rpc('u_asm_set_ql_chia', { p_ma_lan_pc: b.dataset.ql, p_value: b.checked }).then(function () { toast(b.checked ? 'Đã cho phép Quản lý chia lần giao ' + b.dataset.ql : 'Đã tắt quyền chia'); return load(true) }).then(render, function (e) { fail(e); load(true).then(render) }) } });
    document.querySelectorAll('[data-rev]').forEach(function (b) { b.onclick = function () { revoke(b.dataset.rev) } });
    on('#tpl', 'onclick', function () { download('mau-nap-khach.csv', TPL) });
    on('#nfile', 'onchange', function (x) { var f = x.files[0]; if (!f) return; var rd = new FileReader(); rd.onload = function () { ST.impText = rd.result; ST.imp = checkRows(parseCSV(rd.result)); ST.impRes = null; render() }; rd.readAsText(f, 'utf-8') });
    on('#nchk', 'onclick', function () { ST.impText = $('#ntx').value; ST.imp = checkRows(parseCSV(ST.impText)); ST.impRes = null; if (!ST.imp.length) toast('Chưa có dòng dữ liệu. Dòng đầu phải là tên cột.', 0, 1); render() });
    on('#nok', 'onclick', function (x) { x.disabled = true; var rows = ST.imp.filter(function (r) { return !r.err }).map(function (r) { return { ma_tcqg: r.ma_tcqg, ho_ten: r.ho_ten, ngay_sinh: r.iso, gioi_tinh: r.gioi_tinh, ten_lh: r.ten_lien_he, quan_he: r.quan_he, sdt: r.sdt } });
      rpc('u_asm_import', { p_rows: rows, p_req: rid() }).then(function (r) { ST.impRes = r; ST.imp = null; ST.impText = ''; toast('Đã nạp ' + r.them_moi + ' khách'); return load(true) }).then(render, function (e) { x.disabled = false; fail(e) }) });
    on('#bshop', 'onchange', function (x) { ST.bcShop = x.value; render() }); on('#bst', 'onchange', function (x) { ST.bcSt = x.value; render() });
    on('#xl', 'onclick', exportTurns); on('#xn', 'onclick', exportStaff); on('#out', 'onclick', logout); on('#cpw', 'onclick', pwModal);
    X.binders.forEach(function (f) { f(on) });
  }
  return { render: render, X: X };
})();
