/* LC TCQG HUB: gộp từ các module pager.js, lib-parse.js, asm.js, asm-lib.js, scan-engine.js, asm-scan.js, asm-hr.js, asm-img.js, asm-admin.js, msg.js (build 20261124a) */

/* ===== pager.js ===== */
/* PHÂN TRANG CHUNG: mọi bảng và danh sách thẻ trong màn chính đều tự có "số dòng mỗi trang" (10 / 15 / 20) và nút trang trước/sau.
 * Chạy sau mỗi lần vẽ màn hình (PGR.run). Không đổi dữ liệu: chỉ ẩn bớt dòng ngoài trang hiện tại. Bảng có thuộc tính data-nopg tự phân trang riêng (Kho Data).
 * Số dòng chọn được nhớ trên máy (localStorage). Trang hiện tại nhớ theo từng bảng; khi số dòng đổi (lọc, tìm) thì về trang 1. */
var PGR = (function () {
  var KEY = 'lc_pgsz', SIZES = [10, 15, 20], st = {};
  function size() { var v = 15; try { v = Number(localStorage.getItem(KEY)) || 15 } catch (e) { } return SIZES.indexOf(v) >= 0 ? v : 15 }
  function setSize(n) { try { localStorage.setItem(KEY, String(n)) } catch (e) { } }
  function opts() { var s = size(); return SIZES.map(function (n) { return '<option value="' + n + '"' + (n === s ? ' selected' : '') + '>' + n + ' dòng</option>' }).join('') }
  function key(el, i) { var ths = [].slice.call(el.querySelectorAll('thead th')).map(function (t) { return t.textContent.trim() }).join(',').slice(0, 60); return (window.S ? S.page : '') + '|' + (ths || 'list' + i) }
  function run() {
    var root = document.getElementById('main'); if (!root) return;
    [].slice.call(root.querySelectorAll('.pgbar')).forEach(function (b) { b.remove() });
    var n = size(), items = [];
    [].slice.call(root.querySelectorAll('table:not([data-nopg])')).forEach(function (tb) { if (tb.closest('.pgskip')) return; var body = tb.tBodies[0]; if (!body) return; var rows = [].slice.call(body.rows).filter(function (r) { return !r.querySelector('td.empty') }); items.push({ el: tb.closest('.tw') || tb, rows: rows, k: key(tb, items.length) }) });
    [].slice.call(root.querySelectorAll('.list:not([data-nopg])')).forEach(function (l) { if (l.closest('.pgskip')) return; var rows = [].slice.call(l.children).filter(function (c) { return !c.classList.contains('empty') && !c.querySelector(':scope > .empty') }); items.push({ el: l, rows: rows, k: key(l, items.length) }) });
    items.forEach(function (it) {
      var total = it.rows.length, s = st[it.k] || (st[it.k] = { p: 1, n: total }); if (s.n !== total) { s.p = 1; s.n = total }
      var pages = Math.max(1, Math.ceil(total / n)); if (s.p > pages) s.p = pages;
      it.rows.forEach(function (r, i) { r.hidden = !(i >= (s.p - 1) * n && i < s.p * n) });
      if (total <= SIZES[0]) return;
      var a = (s.p - 1) * n + 1, b = Math.min(total, s.p * n), bar = document.createElement('div'); bar.className = 'pgbar'; bar.dataset.k = it.k;
      bar.innerHTML = '<span>Dòng ' + a + '–' + b + ' / ' + total + '</span><span class="sp"></span><label>Hiện <select data-pgsize>' + opts() + '</select></label><button class="btn b3" data-pgd="-1"' + (s.p <= 1 ? ' disabled' : '') + ' aria-label="Trang trước">‹</button><b>' + s.p + '/' + pages + '</b><button class="btn b3" data-pgd="1"' + (s.p >= pages ? ' disabled' : '') + ' aria-label="Trang sau">›</button>';
      it.el.insertAdjacentElement('afterend', bar)
    })
  }
  document.addEventListener('click', function (e) {
    var b = e.target.closest && e.target.closest('[data-pgd]'); if (!b) return; var bar = b.closest('.pgbar'), s = st[bar.dataset.k]; if (!s) return; s.p = Math.max(1, s.p + Number(b.dataset.pgd)); run()
  });
  document.addEventListener('change', function (e) {
    var t = e.target; if (!t || !t.hasAttribute || !t.hasAttribute('data-pgsize')) return; setSize(Number(t.value)); if (window.ST && window.S && S.page === 'a_kho') ST.pg = 1; if (typeof render === 'function') render()
  });
  var css = '.pgbar{display:flex;align-items:center;gap:8px;flex-wrap:wrap;padding:8px 8px 4px;font-size:13px;color:var(--muted)}.pgbar .sp{flex:1}.pgbar select{border:1px solid var(--line);border-radius:10px;padding:6px 8px;background:var(--card);min-height:36px;font-size:13px}.pgbar .btn{min-height:36px;min-width:40px;padding:0 10px;font-size:16px}.pgbar b{min-width:44px;text-align:center;color:var(--ink)}';
  try { var s = document.createElement('style'); s.textContent = css; document.head.appendChild(s) } catch (e) { }
  return { run: run, size: size }
})();

;

/* ===== lib-parse.js ===== */
/* ĐỌC FILE EXCEL/CSV CHO THƯ VIỆN – cấu hình cột 7 khu + tìm dòng tiêu đề + chuẩn hóa + kiểm tra dòng.
 * Làm theo mô hình của KH-01 (đặc tả 02_DAC-TA-DUNG-LAI): tìm dòng tiêu đề trong 40 dòng đầu, ghép cột theo tên/alias (bỏ dấu, thường hóa),
 * đọc "kỳ áp dụng" ở đầu file, gộp ô merge, ngày Date -> yyyy-mm-dd (cộng 12 giờ), số tiền/%, chuẩn hóa biến về MỘT dạng {{TEN_BIEN}}.
 * Chạy được cả trong trình duyệt (window.LIBP) và Node (để kiểm thử bằng file thật). Server vẫn kiểm lại toàn bộ khi lưu. */
(function (root) {
  function bo(s) { return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim() }
  var VARS = ['XUNG_HO', 'XUNG_HO_HOA', 'TEN_NGUOI_LH', 'TEN_BE', 'TUOI', 'TEN_NV', 'TEN_CO_SO', 'DIA_CHI', 'GIO_MO_CUA', 'HOTLINE'];
  /* biến cũ của KH-01 -> biến chuẩn (so sau khi bỏ dấu) */
  var VAR_MAP = { 'ten kh': 'TEN_NGUOI_LH', 'ten me': 'TEN_NGUOI_LH', 'ten khach': 'TEN_NGUOI_LH', 'ten khach hang': 'TEN_NGUOI_LH', 'ten phu huynh': 'TEN_NGUOI_LH', 'ten nguoi lien he': 'TEN_NGUOI_LH', 'ten khach hang chung': 'TEN_NGUOI_LH',
    'ten be': 'TEN_BE', 'ten nguoi tiem': 'TEN_BE', 'ten nv': 'TEN_NV', 'ten nhan vien': 'TEN_NV', 'ten tu van vien': 'TEN_NV', 'dia chi shop': 'DIA_CHI', 'dia chi': 'DIA_CHI', 'dia chi diem nhan dien': 'DIA_CHI',
    'hotline': 'HOTLINE', 'hotline bac si': 'HOTLINE', 'ten co so': 'TEN_CO_SO', 'ten shop': 'TEN_CO_SO' };

  /* ---- cấu hình 7 khu (k = tên cột lưu, nhan = tiêu đề file mẫu, alias = tiêu đề khác, bb = bắt buộc, kieu so|ngay, bien = ô nội dung tin, tuKy = ngày có thể lấy từ kỳ áp dụng) ---- */
  var CONFIG = {
    BANG_GIA: { ten: 'Bảng giá', moTa: 'Bảng giá vaccine mũi lẻ. Up thẳng file bảng giá gốc, hệ thống tự tìm dòng tiêu đề và kỳ áp dụng. Nguồn số cho tin nhắn.', cheDo: 'THAY_THE', cot: [
      { k: 'phong_benh', nhan: 'Phòng bệnh', alias: ['loai vaccine'] }, { k: 'ma_sp', nhan: 'Mã sản phẩm', alias: ['ma sp', 'ma hang'] },
      { k: 'ten_vaccine', nhan: 'Tên vaccine', alias: ['ten san pham', 'vaccine', 'ten hang hoa'], bb: 1 }, { k: 'hang', nhan: 'Hãng', alias: ['nha san xuat', 'hang san xuat', 'nuoc san xuat'] },
      { k: 'gia', nhan: 'Giá niêm yết', alias: ['gia', 'gia niem yet mui tiem', 'gia ban', 'don gia', 'gia mui le'], bb: 1, kieu: 'so' },
      { k: 'tu_ngay', nhan: 'Từ ngày', alias: ['ngay bat dau', 'tu ngay ap dung'], kieu: 'ngay', tuKy: 1 }, { k: 'den_ngay', nhan: 'Đến ngày', alias: ['ngay ket thuc', 'den ngay ap dung'], kieu: 'ngay', tuKy: 1 }, { k: 'ghi_chu', nhan: 'Ghi chú' }] },
    CTKM: { ten: 'CTKM tháng', moTa: 'Chương trình khuyến mại theo tháng. Up thẳng file CTKM gốc: hệ thống đọc phần "KM theo sản phẩm" (mũi lẻ, combo, gói). Cần có kỳ áp dụng ở đầu file hoặc cột ngày.', cheDo: 'BO_SUNG', ctkm: 1, cot: [
      { k: 'phan_loai', nhan: 'Phân loại', alias: [] }, { k: 'loai_vaccine', nhan: 'Loại vaccine', alias: [] }, { k: 'noi_dung_ct', nhan: 'Nội dung chương trình', alias: ['noi dung ct'] },
      { k: 'ma_sp', nhan: 'Mã sản phẩm', alias: ['ma sp'] }, { k: 'ten_sp', nhan: 'Tên sản phẩm', alias: ['ten vaccine'], bb: 1 }, { k: 'nha_sx', nhan: 'Nhà sản xuất', alias: ['hang'] },
      { k: 'gia_niem_yet', nhan: 'Giá niêm yết/mũi tiêm', alias: ['gia niem yet', 'gia niem yet mui tiem'], kieu: 'so' }, { k: 'sl_mui', nhan: 'SL mũi/gói', alias: ['sl mui', 'so luong mui'], kieu: 'so' },
      { k: 'gia_goc', nhan: 'Giá gốc', kieu: 'so' }, { k: 'khuyen_mai', nhan: 'Khuyến mãi/giảm giá', alias: ['khuyen mai', 'giam gia'], kieu: 'so', pc: 1 }, { k: 'gia_sau_km', nhan: 'Giá sau khuyến mãi', alias: ['gia sau km'], kieu: 'so' },
      { k: 'gia_1_mui', nhan: 'Giá 1 mũi (khi mua gói)', alias: ['gia 1 mui'], kieu: 'so' }, { k: 'phi_luu_tru', nhan: 'Phí lưu trữ vaccine', alias: ['phi luu tru'], kieu: 'so' }, { k: 'gia_thanh_toan', nhan: 'Giá thanh toán', kieu: 'so' },
      { k: 'pham_vi', nhan: 'Phạm vi áp dụng', alias: ['pham vi'] }, { k: 'tu_ngay', nhan: 'Ngày bắt đầu', alias: ['tu ngay'], kieu: 'ngay', tuKy: 1 }, { k: 'den_ngay', nhan: 'Ngày kết thúc', alias: ['den ngay'], kieu: 'ngay', tuKy: 1 }, { k: 'ghi_chu', nhan: 'Ghi chú' }, { k: 'nhom', nhan: 'Nhóm CT', alias: ['nhom ct'] }, { k: 'ten_chuong_trinh', nhan: 'Tên chương trình', alias: ['ten ct'] }, { k: 'km_kieu', nhan: 'Kiểu KM', alias: [] }] },
    MAU_TUOI: { ten: 'Mẫu Bầu và Nhi', moTa: 'Mẫu tin cho nhóm Bầu và Nhi theo giai đoạn và bước chăm sóc.', cheDo: 'BO_SUNG', cot: [
      { k: 'nhom_tuoi', nhan: 'Nhóm tuổi', bb: 1 }, { k: 'buoc', nhan: 'Bước chăm sóc', alias: ['buoc'] }, { k: 'kenh', nhan: 'Kênh' }, { k: 'ten_mau', nhan: 'Tên mẫu', bb: 1 }, { k: 'noi_dung', nhan: 'Nội dung', bb: 1, bien: 1 }, { k: 'ghi_chu', nhan: 'Ghi chú' }] },
    MAU_CHU_DE: { ten: 'Mẫu theo vaccine và độ tuổi', moTa: 'Mẫu tin theo từng vaccine/chủ đề, chia theo nhóm tuổi/đối tượng.', cheDo: 'BO_SUNG', cot: [
      { k: 'chu_de', nhan: 'Chủ đề', bb: 1 }, { k: 'nhom_ap_dung', nhan: 'Nhóm áp dụng' }, { k: 'ten_mau', nhan: 'Tên mẫu', bb: 1 }, { k: 'noi_dung', nhan: 'Nội dung', bb: 1, bien: 1 }, { k: 'thu_tu', nhan: 'Thứ tự', kieu: 'so' }, { k: 'ghi_chu', nhan: 'Ghi chú' }] },
    MAU_SAU_TIEM: { ten: 'Mẫu CSKH sau tiêm', moTa: 'Mẫu tin chăm sóc sau tiêm (hỏi thăm, dặn dò theo vaccine).', cheDo: 'BO_SUNG', cot: [
      { k: 'ma_mau', nhan: 'Mã mẫu' }, { k: 'nhom_mau', nhan: 'Nhóm mẫu' }, { k: 'vaccine', nhan: 'Vaccine' }, { k: 'thoi_diem', nhan: 'Thời điểm' }, { k: 'doi_tuong', nhan: 'Đối tượng' }, { k: 'kenh', nhan: 'Kênh' }, { k: 'dieu_kien', nhan: 'Điều kiện sử dụng', alias: ['dieu kien'] },
      { k: 'noi_dung', nhan: 'Nội dung tin nhắn', alias: ['noi dung', 'noi dung tin'], bb: 1, bien: 1 }, { k: 'ten_mau', nhan: 'Tên mẫu', bb: 1 }, { k: 'ghi_chu', nhan: 'Ghi chú' }] },
    MAU_CHUNG: { ten: 'Mẫu tin chung', moTa: 'Kho tin nhắn mẫu chung (kết bạn, chăm sóc sau tiêm, vaccine…) dùng tự do.', cheDo: 'THAY_THE', cot: [
      { k: 'ma_tin', nhan: 'Mã tin', alias: ['ma'], bb: 1 }, { k: 'ten_ngan', nhan: 'Tên ngắn' }, { k: 'chu_de', nhan: 'Chủ đề' }, { k: 'ap_dung', nhan: 'Áp dụng cho', alias: ['ap dung'] }, { k: 'noi_dung', nhan: 'Nội dung', alias: ['tin nhan'], bb: 1, bien: 1 },
      { k: 'ma_anh', nhan: 'Mã ảnh' }, { k: 'anh_salechat', nhan: 'Ảnh Salechat', alias: ['duong dan anh'] }, { k: 'thu_tu', nhan: 'Thứ tự', kieu: 'so' }, { k: 'ghi_chu', nhan: 'Ghi chú' }] },
    NGUY_CO: { ten: 'Nguy cơ & đường lây', moTa: 'Nguy cơ, đường lây và điểm cần nhớ theo từng bệnh. Nhân viên bật "Kèm nguy cơ" để chèn vào tin tư vấn.', cheDo: 'THAY_THE', cot: [
      { k: 'benh', nhan: 'Bệnh', bb: 1 }, { k: 'vaccine_ap_dung', nhan: 'Vaccine áp dụng', alias: ['vaccine'] }, { k: 'nguy_co', nhan: 'Nguy cơ', bb: 1, bien: 1 }, { k: 'lay_qua', nhan: 'Lây qua' }, { k: 'diem_can_nho', nhan: 'Điểm cần nhớ' },
      { k: 'tac_nhan', nhan: 'Tác nhân' }, { k: 'trieu_chung', nhan: 'Triệu chứng điển hình', alias: ['trieu chung'] }, { k: 'mua_vung', nhan: 'Mùa cao điểm & vùng miền', alias: ['mua cao diem'] }, { k: 'ghi_chu', nhan: 'Ghi chú' }] }
  };
  var ORDER = ['BANG_GIA', 'CTKM', 'MAU_TUOI', 'MAU_CHU_DE', 'MAU_SAU_TIEM', 'MAU_CHUNG', 'NGUY_CO'];

  function p2(n) { return n < 10 ? '0' + n : '' + n }
  function isoDate(d) { var t = new Date(d.getTime() + 12 * 3600e3); return t.getFullYear() + '-' + p2(t.getMonth() + 1) + '-' + p2(t.getDate()) }
  function cs(v) { if (v == null) return ''; if (v instanceof Date) return isoDate(v); return String(v).trim() }
  function toNum(v) {
    if (v == null || v === '') return null; if (typeof v === 'number') return isFinite(v) ? v : null;
    var s = String(v).replace(/[₫đĐ*\s]|VND|vnd/g, ''); if (!s) return null; var pct = /%$/.test(s); s = s.replace('%', '');
    if (/^\d{1,3}([.,]\d{3})+$/.test(s)) s = s.replace(/[.,]/g, ''); else if (/^\d+,\d{1,2}$/.test(s)) s = s.replace(',', '.'); else if (/,/.test(s) && /\./.test(s)) s = s.replace(/,/g, '');
    var n = Number(s); if (!isFinite(n)) return NaN; return pct ? n / 100 : n
  }
  function toDate(v) {
    if (v == null || v === '') return null; if (v instanceof Date) return isoDate(v); var s = String(v).trim(), m;
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
    if ((m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})$/))) return m[3] + '-' + p2(+m[2]) + '-' + p2(+m[1]);
    return s
  }
  /* chuẩn hóa biến về {{TEN_BIEN}}; trả {t, lat:[biến lạ]} */
  function normVars(t) {
    t = String(t == null ? '' : t);
    function rep(m, inner) { var k = VAR_MAP[bo(inner)]; return k ? '{{' + k + '}}' : m }
    t = t.replace(/\{\{([^{}\n]{1,40})\}\}/g, function (m, i) { var u = i.trim().toUpperCase(); if (VARS.indexOf(u) >= 0) return '{{' + u + '}}'; var k = VAR_MAP[bo(i)] || (u === 'TEN_KHACH' ? 'TEN_NGUOI_LH' : u === 'TEN_NGUOI_TIEM' ? 'TEN_BE' : u === 'TEN_NHAN_VIEN' ? 'TEN_NV' : u === 'DIA_CHI_DIEM_NHAN_DIEN' ? 'DIA_CHI' : u === 'HOTLINE_BAC_SI' ? 'HOTLINE' : ''); return k ? '{{' + k + '}}' : m });
    t = t.replace(/\[([^\[\]\n]{1,40})\]/g, rep).replace(/<([^<>\n]{1,40})>/g, rep);
    return t
  }
  function lat(t) {
    var o = [], r, re = /\{\{([^{}\n]*)\}\}/g; while ((r = re.exec(t))) if (VARS.indexOf(r[1]) < 0) o.push('{{' + r[1] + '}}');
    re = /\[([A-ZÀ-Ỹ][A-ZÀ-Ỹ0-9 _]{2,30})\]/g; while ((r = re.exec(t))) o.push('[' + r[1] + ']'); return o
  }
  /* kiểm tra 1 dòng (bản sao của hàm kiểm trên máy chủ) -> {thieu:[], canh:[]} */
  function check(loai, row) {
    var cfg = CONFIG[loai], th = [], ca = [];
    cfg.cot.forEach(function (c) {
      var v = row[c.k], s = v == null ? '' : String(v).trim();
      if (c.bb && s === '') th.push('Thiếu ' + c.nhan);
      if (s !== '') {
        if (c.kieu === 'so' && !/^-?\d+(\.\d+)?$/.test(s)) th.push(c.nhan + ' không phải số');
        if (c.kieu === 'ngay' && !/^\d{4}-\d{2}-\d{2}$/.test(s)) th.push(c.nhan + ' không phải ngày');
        if (s.length > 5000) th.push(c.nhan + ' quá 5000 ký tự');
        var L = lat(s); if (L.length) ca.push('Biến lạ: ' + L.slice(0, 3).join(', '));
        if (/\.{3,}|…/.test(s) && c.bien) ca.push('Còn chỗ trống "…" cần điền tay');
      }
    });
    return { thieu: th, canh: ca.filter(function (x, i) { return ca.indexOf(x) === i }) }
  }

  /* ---- tìm dòng tiêu đề ---- */
  function colSets(cols) { return cols.map(function (c) { var set = {}; set[bo(c.nhan)] = 1; set[bo(c.k)] = 1; (c.alias || []).forEach(function (a) { set[bo(a)] = 1 }); return { c: c, set: set } }) }
  function findHeader(aoa, cols) {
    var g = colSets(cols), best = null;
    for (var r = 0; r < Math.min(aoa.length, 40); r++) {
      var map = {}, taken = {}, n = 0, row = aoa[r] || [];
      g.forEach(function (x) { for (var j = 0; j < row.length; j++) { if (!taken[j] && row[j] != null && x.set[bo(row[j])]) { map[x.c.k] = j; taken[j] = 1; n++; break } } });
      if (!best || n > best.n) best = { r: r, map: map, n: n };
    }
    return best || { r: 0, map: {}, n: 0 }
  }
  function kyOf(aoa, upto) {
    var out = { ten: '', tu: null, den: null };
    for (var r = 0; r < Math.min(upto, 30); r++) (aoa[r] || []).forEach(function (v) {
      var s = cs(v); if (!s) return; var m;
      if (!out.ten && (m = s.match(/Tên chương trình\s*:?\s*(.+)/i))) out.ten = m[1].trim();
      if (!out.tu && (m = s.match(/(\d{1,2})\s*\/\s*(\d{1,2})(?:\s*\/\s*(\d{4}))?\s*(?:đến|->|→|-|–)\s*(\d{1,2})\s*\/\s*(\d{1,2})\s*\/\s*(\d{4})/i))) {
        var y2 = +m[6], y1 = m[3] ? +m[3] : y2; out.tu = y1 + '-' + p2(+m[2]) + '-' + p2(+m[1]); out.den = y2 + '-' + p2(+m[5]) + '-' + p2(+m[4]);
      }
    });
    return out
  }
  function emptyRow(r) { return !(r || []).some(function (v) { return v != null && String(v).trim() !== '' }) }

  /* ---- đọc chung ---- */
  function parse(loai, aoa) {
    var cfg = CONFIG[loai], cols = cfg.cot, hd = findHeader(aoa, cols);
    var thieu = cols.filter(function (c) { return c.bb && hd.map[c.k] === undefined && !c.tuKy });
    if (!hd.n || thieu.length) return { loi: 'Không tìm thấy cột bắt buộc: ' + thieu.map(function (c) { return c.nhan }).join(', ') + ' (khớp ' + hd.n + ' cột ở dòng ' + (hd.r + 1) + '). Các cột đọc được ở dòng đó: ' + (aoa[hd.r] || []).filter(function (v) { return v != null && String(v).trim() !== '' }).slice(0, 12).map(cs).join(' | ') };
    var ky = kyOf(aoa, hd.r), rows = [], boQua = [], cur = {}, prog = 0, curProg = null, sec = null, secCond = '';
    for (var r = hd.r + 1; r < aoa.length; r++) {
      var src = aoa[r] || []; if (emptyRow(src)) continue;
      if (cfg.ctkm) {
        var first = cs(src.filter(function (v) { return v != null && String(v).trim() !== '' })[0]);
        if (/^\d+\s*\/\s*\S/.test(first) && r > hd.r + 1) { sec = { nhom: 'N' + p2(++prog), title: first.replace(/^\d+\s*\/\s*/, '').trim() }; secCond = ''; continue }
        if (sec) {   /* các mục sau "2/ …", "3/ …": quà tặng, KOL, vòng quay → nhóm "Ưu đãi chung" (đọc gần đúng) */
          var cells = src.map(cs), nz = cells.filter(Boolean).length, hw = cells.filter(function (x) { return /thời gian áp dụng|tên sản phẩm|tên sp|tên quà|thể lệ|giá trị bill|ghi chú|điều kiện áp dụng|phạm vi/i.test(x) }).length;
          if (hw >= 2) { if (/^quà tặng/i.test(cells[2])) secCond = cells[2]; continue }
          if (cells[2]) secCond = cells[2];
          var name = ''; [4, 5, 3].some(function (i) { var x = cells[i]; if (x && !/^'?\d{5,}/.test(x) && !/^(từ|đến|áp dụng)\b/i.test(x)) { name = x; return true } return false });
          if (!name && nz <= 3) name = cells[2];
          if (!name) { boQua.push({ dong: r + 1, ly_do: 'Mục ' + sec.title + ': dòng không có tên quà' }); continue }
          var g = { nhom: sec.nhom, phan_loai: 'Ưu đãi chung', loai_vaccine: sec.title, noi_dung_ct: secCond, ten_sp: name, ghi_chu: cells.filter(function (x) { return x && x !== name && x !== secCond }).join(' · ').slice(0, 300), tu_ngay: ky.tu || undefined, den_ngay: ky.den || undefined, _dong: r + 1, _uudai: 1 };
          if (/^'?\d{5,}/.test(cells[3] || '')) g.ma_sp = cells[3].replace(/'/g, ''); if (ky.ten) g.ten_chuong_trinh = ky.ten;
          rows.push(g); continue
        }
      }
      var o = {};
      cols.forEach(function (c) {
        if (hd.map[c.k] === undefined) return; var v = src[hd.map[c.k]], val;
        if (c.kieu === 'so') { val = toNum(v); if (val !== null && c.pc && val > 0 && val < 1) o.km_kieu = 'PHAN_TRAM' } else if (c.kieu === 'ngay') val = toDate(v); else val = cs(v);
        if (c.bien && typeof val === 'string') val = normVars(val);
        if (val !== null && val !== '' && !(typeof val === 'number' && isNaN(val))) o[c.k] = val; else if (typeof val === 'number' && isNaN(val)) o[c.k] = String(cs(v));
      });
      if (cfg.ctkm) {
        if (o.phan_loai) { prog++; curProg = { nhom: 'N' + p2(prog), phan_loai: o.phan_loai, loai_vaccine: o.loai_vaccine || '', noi_dung_ct: o.noi_dung_ct || '' }; }
        if (!curProg) { boQua.push({ dong: r + 1, ly_do: 'Chưa thuộc chương trình nào' }); continue }
        if (!o.ten_sp) { boQua.push({ dong: r + 1, ly_do: 'Không có tên sản phẩm' }); continue }
        o.nhom = curProg.nhom; o.phan_loai = curProg.phan_loai; if (!o.loai_vaccine) o.loai_vaccine = curProg.loai_vaccine; if (!o.noi_dung_ct) o.noi_dung_ct = curProg.noi_dung_ct;
        if (ky.ten) o.ten_chuong_trinh = ky.ten; if (!o.km_kieu && o.khuyen_mai > 0 && o.khuyen_mai < 1) o.km_kieu = 'PHAN_TRAM';
      } else {
        cols.forEach(function (c) { if (c.fill === 'tiep') { if (o[c.k] === undefined && cur[c.k] !== undefined) o[c.k] = cur[c.k]; else if (o[c.k] !== undefined) cur[c.k] = o[c.k] } });
        var hasBB = cols.some(function (c) { return c.bb && o[c.k] !== undefined }), nonEmpty = src.filter(function (v) { return v != null && String(v).trim() !== '' }).length;
        if (!hasBB && nonEmpty <= 2) { boQua.push({ dong: r + 1, ly_do: 'Dòng ghi chú, không có dữ liệu chính' }); continue }
      }
      cols.forEach(function (c) { if (c.tuKy && o[c.k] === undefined && ky[c.k === 'tu_ngay' ? 'tu' : 'den']) o[c.k] = ky[c.k === 'tu_ngay' ? 'tu' : 'den'] });
      o._dong = r + 1; rows.push(o);
    }
    var tong = rows.length, bad = 0;
    rows.forEach(function (o) { var ck = check(loai, o); o._thieu = ck.thieu; o._canh = ck.canh; if (o._uudai) o._canh.push('Ưu đãi chung: đọc gần đúng, hãy xem lại'); if (ck.thieu.length) bad++ });
    return { rows: rows, ky: ky, boQua: boQua, tong: tong, thieu: bad, hopLe: tong - bad, dongTieuDe: hd.r + 1 }
  }

  /* ---- chọn sheet phù hợp nhất trong workbook (XLSX = thư viện SheetJS) ---- */
  function sheetAoa(XLSX, ws, merge) {
    var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
    if (merge && ws['!merges']) ws['!merges'].forEach(function (m) {
      var v = (aoa[m.s.r] || [])[m.s.c]; if (v == null) return;
      for (var r = m.s.r; r <= m.e.r; r++) { aoa[r] = aoa[r] || []; for (var c = m.s.c; c <= m.e.c; c++) if (aoa[r][c] == null) aoa[r][c] = v }
    });
    return aoa
  }
  function fromWorkbook(XLSX, wb, loai) {
    var cfg = CONFIG[loai], best = null;
    wb.SheetNames.forEach(function (n) {
      var aoa = sheetAoa(XLSX, wb.Sheets[n], !cfg.ctkm), hd = findHeader(aoa, cfg.cot);
      if (!best || hd.n > best.n) best = { n: hd.n, name: n, aoa: aoa }
    });
    if (!best) return { loi: 'File không có sheet nào.' };
    var p = parse(loai, best.aoa); p.sheet = best.name; return p
  }
  /* file mẫu: sheet DU_LIEU (chỉ dòng tiêu đề) + HUONG_DAN (cột, bắt buộc, kiểu, ghi chú) */
  function templateAoa(loai) {
    var cfg = CONFIG[loai];
    return { du_lieu: [cfg.cot.map(function (c) { return c.nhan })], huong_dan: [['Cột', 'Bắt buộc', 'Kiểu', 'Ghi chú']].concat(cfg.cot.map(function (c) { return [c.nhan, c.bb ? 'Có' : '', c.kieu === 'so' ? 'Số' : c.kieu === 'ngay' ? 'Ngày' : 'Chữ', c.bien ? 'Biến trong tin dùng dạng {{TEN_BIEN}}: ' + VARS.join(', ') : (c.tuKy ? 'Có thể bỏ trống nếu file có kỳ áp dụng ở đầu' : '')] })) }
  }
  root.LIBP = { bo: bo, CONFIG: CONFIG, ORDER: ORDER, VARS: VARS, parse: parse, check: check, normVars: normVars, fromWorkbook: fromWorkbook, templateAoa: templateAoa, findHeader: findHeader };
  if (typeof module !== 'undefined') module.exports = root.LIBP;
})(typeof window !== 'undefined' ? window : globalThis);

;

/* ===== asm.js ===== */
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
  function PAGEN() { return PGR.size() }
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
    h += '<div class="stats stats6">' + [['Khách trong kho', tot], ['Chưa giao', kho], ['Đang giữ', all.mo], ['Hoàn thành', all.xong], ['Đồng ý tiêm', all.dy], ['Quá hạn gọi', all.qh]].map(function (x, i) { return '<div class="card st"><div class="l">' + x[0] + '</div><div class="v"' + (i === 5 && x[1] ? ' style="color:var(--bad)"' : '') + '>' + x[1] + '</div></div>' }).join('') + '</div>';
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
    var pages = Math.max(1, Math.ceil(all.length / PAGEN())); ST.pg = Math.min(ST.pg, pages); var rows = all.slice((ST.pg - 1) * PAGEN(), ST.pg * PAGEN()), n = Object.keys(ST.sel).length;
    var h = head('Kho Data', 'Toàn bộ khách. Chọn khách rồi bấm Phân công.', '');
    h += '<div class="top"><div class="search">' + ic('search') + '<input id="kq" placeholder="Tìm tên, mã TCQG, số điện thoại…" value="' + esc(ST.q) + '"></div>' + sel('fband', BANDS, ST.band) + '</div>';
    h += '<div class="chips">' + [['', 'Tất cả'], ['kho', 'Chưa giao'], ['dang', 'Đang giao'], ['xong', 'Đã xong']].map(function (x) { return '<button class="chip ' + (ST.f === x[0] ? 'on' : '') + '" data-f="' + x[0] + '">' + x[1] + '<em>' + cnt[x[0]] + '</em></button>' }).join('') + '</div>';
    h += '<div class="card" style="padding:6px 8px"><div class="tw"><table data-nopg><thead><tr><th style="width:36px"><input type="checkbox" id="selpg" ' + (rows.length && rows.every(function (r) { return ST.sel[r.c.ma_khach] }) ? 'checked' : '') + ' style="width:20px;height:20px"></th><th>Khách</th><th>Tuổi</th><th>Phụ huynh</th><th>Trạng thái</th><th>Đang giữ</th></tr></thead><tbody>' +
      rows.map(function (r) { var op = r.ts.filter(open_)[0], who = op ? (op.recipient_type === 'SHOP' ? shopName(op.id_shop) + ' (Shop)' : nm(op.ma_tk_nhan)) : '';
        return '<tr><td><input type="checkbox" data-k="' + esc(r.c.ma_khach) + '" ' + (ST.sel[r.c.ma_khach] ? 'checked' : '') + ' style="width:20px;height:20px"></td><td><b>' + esc(r.c.ho_ten) + '</b><small>' + esc(r.c.ma_tcqg || r.c.ma_khach) + '</small></td><td>' + esc(ageTxt(r.c)) + '</td><td>' + esc(r.l.ho_ten || '') + '<small>' + esc(r.l.sdt || '') + '</small></td><td>' + ({ kho: '<span class="tag t-tint">Chưa giao</span>', dang: '<span class="tag t-late">Đang giao</span>', xong: '<span class="tag t-ok">Đã xong</span>' })[r.st] + '</td><td>' + esc(who) + '</td></tr>' }).join('') +
      (rows.length ? '' : '<tr><td colspan="6" class="empty">Không có khách phù hợp.</td></tr>') + '</tbody></table></div>' +
      '<div class="row2" style="padding:10px"><span class="meta">' + all.length + ' khách · trang ' + ST.pg + '/' + pages + '</span><div style="display:flex;gap:8px;align-items:center"><select data-pgsize style="border:1px solid var(--line);border-radius:10px;padding:6px 8px;background:var(--card);min-height:38px">' + [10, 15, 20].map(function (n) { return '<option value="' + n + '"' + (n === PAGEN() ? ' selected' : '') + '>' + n + ' dòng</option>' }).join('') + '</select><button class="btn b3" data-pg="-1" style="min-height:38px">‹</button><button class="btn b3" data-pg="1" style="min-height:38px">›</button></div></div></div>';
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
      h += '<div class="stats stats4"><div class="card st"><div class="l">Dòng trong file</div><div class="v">' + ST.imp.length + '</div></div><div class="card st"><div class="l">Hợp lệ</div><div class="v" style="color:var(--ok)">' + ok.length + '</div></div><div class="card st"><div class="l">Đã có (bỏ qua)</div><div class="v">' + dup + '</div></div><div class="card st"><div class="l">Lỗi</div><div class="v" style="color:var(--bad)">' + bad + '</div></div></div>' +
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
    on('#selpg', 'onchange', function (x) { var all = khoRows(), rows = all.slice((ST.pg - 1) * PAGEN(), ST.pg * PAGEN()); rows.forEach(function (r) { x.checked ? ST.sel[r.c.ma_khach] = 1 : delete ST.sel[r.c.ma_khach] }); render() });
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

;

/* ===== asm-lib.js ===== */
/* THƯ VIỆN – KHO TỪ EXCEL (Đợt 1, theo mô hình KH-01): chọn khu -> tải file mẫu / up file -> XEM TRƯỚC (chưa ghi) -> lưu lô CHỜ DUYỆT -> DUYỆT -> ĐANG DÙNG.
 * Lô: Bổ sung (thêm dòng) hoặc Thay thế (lô đang dùng cùng khu chuyển "Đã thay"). Sửa tại ô, xóa dòng, ngừng/kích hoạt lại/xóa lô, tải lô về Excel. Mọi ghi qua hàm RPC chỉ ASM gọi được.
 * Đọc Excel bằng SheetJS ngay trên trình duyệt (lib-parse.js); máy chủ kiểm lại toàn bộ khi lưu. Dữ liệu thử là dữ liệu giả/mẫu. */
var ASMLIB = (function () {
  var P = window.LIBP, LS = { sel: '', parsed: null, tenFile: '', cheDo: '', lots: null, loading: false, lot: null, lotRows: null, busy: false };
  var TRANG = { CHO_DUYET: ['t-late', 'Chờ duyệt'], DANG_DUNG: ['t-ok', 'Đang dùng'], DA_THAY: ['t-new', 'Đã thay'], NGUNG_DUNG: ['t-bad', 'Ngừng dùng'], THIEU: ['t-bad', 'Thiếu'] };
  var SAU = [['Phác đồ vaccine', 'Anh gửi dữ liệu phác đồ sau, nút nạp sẽ làm như các khu khác (menu Phác đồ đã giữ chỗ).'], ['Mẫu tin quy trình', 'Các bước D, D+1, D+3, D+7 theo nhóm Lao/6in1. Làm ở đợt sau.'], ['Combo ưu đãi', 'Gói combo dạng bảng. Làm ở đợt sau.']];

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

;

/* ===== scan-engine.js ===== */
/* QUÉT GIÁ & CTKM (Đợt 2): đối chiếu số tiền trong các mẫu tin với Bảng giá + CTKM đang hiệu lực, đề xuất thay tự động, báo CTKM chưa được nhắc.
 * Hàm thuần (không gọi mạng) để kiểm thử bằng file thật. Đầu vào: dòng thư viện {id, loai, trang_thai, data, canh_bao}. Đầu ra: danh sách mẫu kèm cảnh báo và bản đề xuất.
 * Nguyên tắc an toàn: chỉ đề xuất; thay thật do ASM bấm (máy chủ so lại nội dung cũ trước khi ghi). Cảnh báo không chắc chắn (nhiều giá, số lạ) chỉ cho "Sửa tay". */
(function (root) {
  var MSG_TYPES = ['MAU_TUOI', 'MAU_CHU_DE', 'MAU_SAU_TIEM', 'MAU_CHUNG'];
  var STOP = { ml: 1, inj: 1, vial: 1, ong: 1, hap: 1, phu: 1, viet: 1, nam: 1, vaccine: 1, vacxin: 1, dung: 1, dich: 1, tiem: 1, lieu: 1, hop: 1, chai: 1, bom: 1, pfs: 1, prefilled: 1, syringe: 1, ngua: 1, ung: 1, thu: 1, mui: 1, vac: 1, xin: 1 };
  var PROMO = ['giam', 'uu dai', 'ho tro', 'tang ', 'khuyen mai', 'combo', 'voucher', 'mien phi', 'qua tang', 'chuong trinh'];
  function fold(s) { var o = ''; s = String(s == null ? '' : s); for (var i = 0; i < s.length; i++) { var c = s.charAt(i), b = c.normalize('NFD').charAt(0).toLowerCase(); if (c === 'đ' || c === 'Đ') b = 'd'; o += b } return o }
  function fmt(n) { return String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, '.') }
  function keysOf(name) {
    var toks = fold(name).split(/[^a-z0-9]+/).filter(Boolean), out = [];
    toks.forEach(function (t, i) {
      if (t.length < 4 || STOP[t] || /^\d+$/.test(t) || !/[a-z]/.test(t)) return; var nx = toks[i + 1];
      if (nx && /^\d{1,3}$/.test(nx) && toks[i + 2] !== 'ml') out.unshift(t + ' ' + nx); out.push(t)
    });
    return out.filter(function (k, i) { return out.indexOf(k) === i })
  }
  function phraseKey(lv) { var p = fold(lv).replace(/[^a-z0-9 ]+/g, ' ').replace(/\s+/g, ' ').trim().replace(/^ngua /, ''); return p.length >= 3 && p.length <= 24 ? p : '' }
  function amountsOf(text) {
    var re = /\d{1,3}(?:[.,]\d{3})+(?:\s?(?:đ|₫|vnđ|vnd|đồng))?/gi, m, out = [];
    while ((m = re.exec(text))) {
      var pre = text.charAt(m.index - 1), nxt = text.charAt(m.index + m[0].length);
      if (/\d/.test(pre || '') || /\d/.test(nxt || '')) continue; var v = Number(m[0].replace(/[^\d]/g, '').replace(/^0+/, '') || 0); if (v < 1000) continue;
      out.push({ text: m[0], val: v, start: m.index, end: m.index + m[0].length, unit: /(đ|₫|vnđ|vnd|đồng)$/i.test(m[0]) })
    }
    return out
  }
  function today() { var d = new Date(); return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2) }
  function title(r) { var d = r.data || {}; return r.loai === 'MAU_CHUNG' ? ((d.ma_tin || '') + (d.ten_ngan ? ' – ' + d.ten_ngan : '')) : (d.ten_mau || d.chu_de || ('Dòng ' + r.stt)) }

  /* gom nguồn số liệu: bảng giá, CTKM, tập "số tiền hợp lệ" */
  function sources(rows, incPending, day) {
    var ok = function (r) { return r.trang_thai === 'DANG_DUNG' || (incPending && r.trang_thai === 'CHO_DUYET') };
    var vals = {}, idx = {}, progs = {}, nBG = 0, nCK = 0;
    function addVal(v, src, label) { if (!v || v < 1000) return; if (!vals[v] || (src === 'BG' && vals[v].src !== 'BG')) vals[v] = { src: src, label: label } }
    rows.forEach(function (r) {
      if (!ok(r)) return; var d = r.data || {};
      if (r.loai === 'BANG_GIA') {
        if (d.den_ngay && d.den_ngay < day) return; if (d.tu_ngay && d.tu_ngay > day) return; nBG++;
        var g = Number(d.gia); addVal(g, 'BG', d.ten_vaccine);
        keysOf(d.ten_vaccine).forEach(function (k) { (idx[k] = idx[k] || {})[g] = d.ten_vaccine })
      } else if (r.loai === 'CTKM') {
        if (d.den_ngay && d.den_ngay < day) return; if (d.tu_ngay && d.tu_ngay > day && d.phan_loai !== 'Ưu đãi chung') return; nCK++;
        var p = progs[d.nhom] || (progs[d.nhom] = { nhom: d.nhom, phan_loai: d.phan_loai, loai_vaccine: d.loai_vaccine, noi_dung_ct: d.noi_dung_ct, rows: [], vals: {}, keys: [], uudai: d.phan_loai === 'Ưu đãi chung' });
        p.rows.push(d); var lab = d.ten_sp + ' (' + d.nhom + ')';
        ['gia_niem_yet', 'gia_goc', 'gia_sau_km', 'gia_1_mui', 'phi_luu_tru', 'gia_thanh_toan'].forEach(function (k) { var v = Number(d[k]); if (v >= 1000) { addVal(v, 'CK', lab); if (k !== 'gia_niem_yet' && k !== 'gia_goc') p.vals[v] = 1 } });
        if (d.km_kieu !== 'PHAN_TRAM' && Number(d.khuyen_mai) >= 1000) { addVal(Number(d.khuyen_mai), 'CK', lab); p.vals[Number(d.khuyen_mai)] = 1 }
        amountsOf(d.noi_dung_ct || '').forEach(function (a) { addVal(a.val, 'CK', d.nhom + ': ' + String(d.noi_dung_ct).slice(0, 30)); p.vals[a.val] = 1 });
        keysOf(d.ten_sp).forEach(function (k) { if (p.keys.indexOf(k) < 0) p.keys.push(k) });
        var pk = phraseKey(d.loai_vaccine); if (pk && p.keys.indexOf(pk) < 0) { p.keys.push(pk); (p.phr = p.phr || []).push(pk) }
      }
    });
    return { vals: vals, idx: idx, progs: progs, nBG: nBG, nCK: nCK }
  }
  function occurrences(folded, keys) {
    var occ = []; keys.sort(function (a, b) { return b.length - a.length }).forEach(function (k) {
      var from = 0, i; while ((i = folded.indexOf(k, from)) >= 0) {
        var e = i + k.length, wb = !/[a-z0-9]/.test(folded.charAt(i - 1) || ' ') && !/[a-z0-9]/.test(folded.charAt(e) || ' ');
        if (wb && !occ.some(function (o) { return i < o.end && e > o.start })) occ.push({ key: k, start: i, end: e }); from = i + 1
      }
    });
    return occ
  }
  function nearest(occ, a, text) {
    var bef = null, aft = null;
    occ.forEach(function (o) {
      if (o.end <= a.start) { if (/\d{1,3}(?:[.,]\d{3})+/.test(text.slice(o.end, a.start))) return; var d = a.start - o.end; if (d <= 70 && (!bef || d < bef.d)) bef = { o: o, d: d } }
      else if (o.start >= a.end) { var d2 = o.start - a.end; if (d2 <= 25 && (!aft || d2 < aft.d)) aft = { o: o, d: d2 } }
    });
    var best = bef || aft; if (best) best.o.dist = best.d; return best && best.o
  }

  function scan(rows, opt) {
    opt = opt || {}; var day = opt.day || today(), S = sources(rows, opt.incPending !== false, day), msgs = [], okState = function (r) { return r.trang_thai === 'DANG_DUNG' || (opt.incPending !== false && r.trang_thai === 'CHO_DUYET') };
    var allText = '', progMentioned = {}, bgKeys = Object.keys(S.idx);
    var progKeyList = []; Object.keys(S.progs).forEach(function (n) { var p = S.progs[n]; if (!p.uudai) p.keys.forEach(function (k) { progKeyList.push({ k: k, n: n, phrase: (p.phr || []).indexOf(k) >= 0 }) }) });
    rows.forEach(function (r) {
      if (MSG_TYPES.indexOf(r.loai) < 0 || !okState(r)) return; var text = (r.data && r.data.noi_dung) || ''; if (!text) return;
      var f = fold(text), amts = amountsOf(text), occ = occurrences(f, bgKeys.concat(progKeyList.map(function (x) { return x.k }))), warns = [], edits = [];
      amts.forEach(function (a) {
        var hit = S.vals[a.val]; a.ok = !!hit; a.label = hit ? (hit.src === 'BG' ? 'Bảng giá: ' : 'CTKM: ') + hit.label : '';
        if (a.ok) return; var o = nearest(occ, a, text), vk = o && o.key, P = vk && S.idx[vk] ? Object.keys(S.idx[vk]).map(Number) : [];
        if (vk) a.vk = vk;
        if (P.length === 1 && o.dist > 45) warns.push({ type: 'gia_nhieu', level: 'tay', msg: 'Giá ' + a.text + ' có thể khác giá hiện hành của ' + S.idx[vk][P[0]] + ' (' + fmt(P[0]) + 'đ), nhưng tên vaccine nằm khá xa số tiền nên cần anh xem lại rồi sửa tay.', a: a });
        else if (P.length === 1) { a.sug = P[0]; warns.push({ type: 'gia', level: 'auto', msg: 'Giá ' + a.text + ' khác giá hiện hành của ' + S.idx[vk][P[0]] + ': ' + fmt(P[0]) + 'đ', a: a }); edits.push({ start: a.start, end: a.end, rep: fmt(P[0]) + (a.unit ? 'đ' : '') }) }
        else if (P.length > 1) warns.push({ type: 'gia_nhieu', level: 'tay', msg: 'Giá ' + a.text + ' gần "' + vk + '" nhưng có nhiều giá hiện hành (' + P.map(fmt).join(', ') + '). Chọn giá đúng rồi sửa tay.', a: a });
        else warns.push({ type: 'so_la', level: 'tay', msg: 'Số tiền ' + a.text + ' chưa thấy trong Bảng giá/CTKM hiện hành.', a: a })
      });
      /* CTKM chưa nêu */
      var seenProg = {}, hasPromoWord = PROMO.some(function (w) { return f.indexOf(w) >= 0 });
      var brand = {}; occ.forEach(function (o) { progKeyList.forEach(function (x) { if (x.k === o.key) { if (!seenProg[x.n]) seenProg[x.n] = 1; if (!x.phrase) brand[x.n] = 1 } }) });
      var matched = Object.keys(seenProg);
      matched.forEach(function (n) { var p = S.progs[n]; if (p.uudai) return; if (amts.some(function (a) { return p.vals[a.val] })) { progMentioned[n] = 1; seenProg[n] = 2 } else if (f.indexOf(fold(p.noi_dung_ct || '').slice(0, 28)) >= 0 && (p.noi_dung_ct || '').length > 8) { progMentioned[n] = 1; seenProg[n] = 2 } });
      var doneW = /da tiem|tiem du|tiem roi|da duoc tiem/.test(f), anyM = matched.some(function (n) { return seenProg[n] === 2 });
      var miss = anyM ? [] : matched.filter(function (n) { return seenProg[n] === 1 && !S.progs[n].uudai && (brand[n] || !doneW) });
      if (miss.length && !hasPromoWord) { var p = S.progs[miss[0]], sent = '🎁 ' + String(p.noi_dung_ct || '').replace(/\s+/g, ' ').trim(); warns.push({ type: 'ctkm', level: 'auto', msg: 'Tin có nhắc ' + (p.loai_vaccine || p.rows[0].ten_sp) + ' nhưng chưa nêu CTKM ' + p.nhom + ' (' + String(p.noi_dung_ct || '').slice(0, 70) + ').', sent: sent, prog: p.nhom }); edits.push({ start: text.length, end: text.length, rep: '\n\n' + sent }) }
      (r.canh_bao || '').split('; ').filter(Boolean).forEach(function (c) { warns.push({ type: 'dinh_dang', level: 'tay', msg: c }) });
      var proposed = text; edits.sort(function (a, b) { return b.start - a.start }).forEach(function (e) { proposed = proposed.slice(0, e.start) + e.rep + proposed.slice(e.end) });
      msgs.push({ id: r.id, ma_lo: r.ma_lo, loai: r.loai, trang_thai: r.trang_thai, title: title(r), text: text, amts: amts, warns: warns, proposed: edits.length ? proposed : null, canApply: edits.length > 0 });
      allText += '\n' + f
    });
    var unm = Object.keys(S.progs).map(function (n) { return S.progs[n] }).filter(function (p) { return !p.uudai && !progMentioned[p.nhom] && !Object.keys(p.vals).some(function (v) { return allText.indexOf(String(v).replace(/\B(?=(\d{3})+(?!\d))/g, '.')) >= 0 }) });
    var nWarn = msgs.filter(function (m) { return m.warns.some(function (w) { return w.type !== 'dinh_dang' }) }).length;
    return { msgs: msgs, unmentioned: unm.map(function (p) { return { nhom: p.nhom, phan_loai: p.phan_loai, loai_vaccine: p.loai_vaccine, noi_dung_ct: p.noi_dung_ct, sent: '🎁 ' + String(p.noi_dung_ct || '').replace(/\s+/g, ' ').trim() } }), stats: { msgs: msgs.length, bg: S.nBG, ck: S.nCK, warn: nWarn, unm: unm.length, day: day } }
  }
  root.LIBSCAN = { scan: scan, amountsOf: amountsOf, fmt: fmt, fold: fold, keysOf: keysOf };
  if (typeof module !== 'undefined') module.exports = root.LIBSCAN;
})(typeof window !== 'undefined' ? window : globalThis);

;

/* ===== asm-scan.js ===== */
/* THƯ VIỆN – QUÉT GIÁ & CTKM (Đợt 2): rà các mẫu tin so với Bảng giá + CTKM hiện hành, đề xuất thay tự động, báo CTKM chưa mẫu nào nhắc, cho sửa tay.
 * Quét chỉ ĐỌC. Chỉ khi ASM bấm "Áp đề xuất" / "Thay các mẫu đã tích" / "Lưu" (sửa tay) mới ghi, qua hàm u_lib_replace_text (máy chủ so lại nội dung cũ, ghi nhật ký cũ/mới).
 * Logic đối chiếu nằm ở scan-engine.js. Dữ liệu thử là dữ liệu mẫu. */
var ASMSCAN = (function () {
  var SCN = { inc: true, rows: null, res: null, busy: false, filter: 'warn', sel: {}, skip: {}, open: {}, applying: false };
  var TYPES = { MAU_TUOI: 'Bầu và Nhi', MAU_CHU_DE: 'Theo vaccine', MAU_SAU_TIEM: 'CSKH sau tiêm', MAU_CHUNG: 'Mẫu chung' };
  var WT = { gia: ['t-late', 'Giá lệch'], gia_nhieu: ['t-late', 'Nhiều giá'], so_la: ['t-new', 'Số chưa rõ'], ctkm: ['t-today', 'Chưa nêu CTKM'], dinh_dang: ['t-new', 'Định dạng'] };
  try { var st = document.createElement('style'); st.textContent = '.mk-ok{background:#dff3e7;border-radius:4px;padding:0 2px}.mk-bad{background:#ffd9d6;border-radius:4px;padding:0 2px}.mk-new{background:#fff0a8;border-radius:4px;padding:0 2px}.scn pre{white-space:pre-wrap;word-break:break-word;font:inherit;margin:0}details.scn>summary{cursor:pointer;list-style:none;display:flex;gap:10px;align-items:center;flex-wrap:wrap}details.scn>summary::-webkit-details-marker{display:none}'; document.head.appendChild(st) } catch (e) { }

  function mark(text, amts, cls) { var h = '', pos = 0; amts.slice().sort(function (a, b) { return a.start - b.start }).forEach(function (a) { if (a.start < pos) return; h += esc(text.slice(pos, a.start)) + '<span class="' + cls(a) + '">' + esc(text.slice(a.start, a.end)) + '</span>'; pos = a.end }); return h + esc(text.slice(pos)) }
  function scan() {
    SCN.busy = true; render();
    return rest('lib_rows?select=id,ma_lo,loai,stt,data,trang_thai,canh_bao&loai=in.(BANG_GIA,CTKM,MAU_TUOI,MAU_CHU_DE,MAU_SAU_TIEM,MAU_CHUNG)&trang_thai=in.(DANG_DUNG,CHO_DUYET)&limit=5000').then(function (rows) { SCN.rows = rows; SCN.res = LIBSCAN.scan(rows, { incPending: SCN.inc }); SCN.busy = false; SCN.sel = {}; render() }, function (e) { SCN.busy = false; render(); fail(e) })
  }
  function shown() { var r = SCN.res; if (!r) return []; return r.msgs.filter(function (m) { if (SCN.skip[m.id]) return false; if (SCN.filter === 'warn') return m.warns.some(function (w) { return w.type !== 'dinh_dang' }); if (SCN.filter === 'auto') return m.canApply; return true }) }
  function page() {
    var r = SCN.res, h = '<div class="card"><h2>Quét giá &amp; CTKM so với các mẫu tin <span class="pill">chỉ đọc cho tới khi anh bấm thay</span></h2><div class="meta" style="margin-bottom:10px">Đối chiếu Bảng giá và CTKM đang hiệu lực hôm nay với số tiền trong mọi mẫu tin (Bầu–Nhi, theo vaccine, CSKH sau tiêm, mẫu chung). Báo tin ghi giá khác hiện hành (kèm đề xuất thay), tin nhắc vaccine nhưng chưa nêu CTKM, và CTKM chưa mẫu nào nhắc tới. Việc nào hệ thống không chắc thì để anh <b>sửa tay</b>. Đây là gợi ý để anh xem xét.</div><div class="row" style="flex-wrap:wrap"><label class="row" style="cursor:pointer"><input type="checkbox" id="sc-inc" ' + (SCN.inc ? 'checked' : '') + ' style="width:20px;height:20px;accent-color:var(--teal)"><span>Tính cả lô Chờ duyệt (xem thử)</span></label><span class="sp"></span><button class="btn b1" id="sc-run"' + (SCN.busy ? ' disabled' : '') + '>' + (SCN.busy ? 'Đang quét…' : r ? 'Quét lại' : 'Quét') + '</button></div></div>';
    if (!r) return h + '<div class="card empty">Bấm "Quét" để bắt đầu. Cần có Bảng giá/CTKM và mẫu tin trong Kho từ Excel (đã duyệt, hoặc tick "tính cả lô chờ duyệt").</div>';
    var st = r.stats, list = shown(), nSel = Object.keys(SCN.sel).length, nAuto = r.msgs.filter(function (m) { return m.canApply && !SCN.skip[m.id] }).length;
    h += '<div class="alert"><i>📅</i><div><b>Hôm nay ' + esc(st.day) + ' · đã rà ' + st.msgs + ' mẫu · đang tính ' + st.bg + ' dòng giá, ' + st.ck + ' dòng CTKM</b><small>' + st.warn + ' mẫu có cảnh báo · ' + st.unm + ' CTKM chưa mẫu nào nhắc tới' + (st.bg === 0 ? ' · <b>chưa có Bảng giá nào đang dùng</b>, nên chưa đề xuất được việc thay giá' : '') + '</small></div></div>';
    h += '<div class="chips">' + [['warn', 'Có cảnh báo', r.msgs.filter(function (m) { return !SCN.skip[m.id] && m.warns.some(function (w) { return w.type !== 'dinh_dang' }) }).length], ['auto', 'Có thể tự thay', nAuto], ['all', 'Tất cả', r.msgs.length]].map(function (x) { return '<button class="chip ' + (SCN.filter === x[0] ? 'on' : '') + '" data-scf="' + x[0] + '">' + x[1] + '<em>' + x[2] + '</em></button>' }).join('') + '</div>';
    h += '<div class="row" style="flex-wrap:wrap"><button class="btn b3" id="sc-all" style="min-height:40px">Chọn tất cả mẫu có đề xuất</button><button class="btn b3" id="sc-none" style="min-height:40px">Bỏ chọn</button><span class="sp"></span><button class="btn b1" id="sc-apply" style="min-height:40px"' + (nSel && !SCN.applying ? '' : ' disabled') + '>Thay các mẫu đã tích (' + nSel + ')</button></div>';
    h += '<div class="list" style="margin-top:10px">' + (list.length ? list.map(card).join('') : '<div class="card empty">Không có mẫu nào trong mục này.</div>') + '</div>';
    if (st.unm) h += '<div class="card" style="padding:6px 8px"><h2 style="padding:10px 10px 0">CTKM chưa mẫu nào nhắc tới <span class="pill">' + st.unm + '</span></h2><div class="tw"><table><thead><tr><th>Chương trình</th><th>Loại</th><th>Nội dung</th><th>Câu gợi ý đưa vào mẫu</th><th></th></tr></thead><tbody>' + r.unmentioned.map(function (u, i) { return '<tr><td><b>' + esc(u.nhom) + '</b></td><td>' + esc(u.phan_loai || '') + '<small>' + esc(u.loai_vaccine || '') + '</small></td><td style="white-space:normal;max-width:260px">' + esc(u.noi_dung_ct || '') + '</td><td style="white-space:normal;max-width:300px">' + esc(u.sent) + '</td><td><button class="btn b3" data-sccp="' + i + '" style="min-height:34px;font-size:13px">Chép câu</button></td></tr>' }).join('') + '</tbody></table></div></div>';
    return h
  }
  function card(m) {
    var oth = {}; m.amts.forEach(function (a) { oth[a.val] = 1 });
    var warnN = m.warns.filter(function (w) { return w.type !== 'dinh_dang' }).length, fmtN = m.warns.length - warnN;
    var h = '<details class="card scn" data-scid="' + m.id + '"' + (SCN.open[m.id] ? ' open' : '') + ' style="padding:12px"><summary>' + (m.canApply ? '<input type="checkbox" data-scsel="' + m.id + '" ' + (SCN.sel[m.id] ? 'checked' : '') + ' style="width:20px;height:20px;accent-color:var(--teal)">' : '') + '<b style="flex:1;min-width:160px">' + esc(m.title) + '</b>' + tag('t-new', TYPES[m.loai] || m.loai) + (m.trang_thai === 'CHO_DUYET' ? tag('t-late', 'chờ duyệt') : '') + (warnN ? tag('t-late', warnN + ' cảnh báo') : tag('t-ok', 'ổn')) + (fmtN ? tag('t-new', fmtN + ' định dạng') : '') + '</summary><div style="margin-top:10px">';
    h += '<div class="script" style="margin:0 0 8px"><pre>' + mark(m.text, m.amts, function (a) { return a.ok ? 'mk-ok' : 'mk-bad' }) + '</pre></div>';
    if (m.amts.length) h += '<div class="meta" style="margin-bottom:6px">Số tiền trong tin: ' + m.amts.map(function (a) { return '<span class="tag ' + (a.ok ? 't-ok' : 't-bad') + '">' + esc(a.text) + (a.ok ? ' ✓ ' + esc(a.label) : ' ✗') + '</span>' }).join('') + '</div>';
    h += m.warns.map(function (w) { var t = WT[w.type] || ['t-new', w.type]; return '<div class="alert ' + (w.level === 'auto' ? 'a-warn' : '') + '" style="margin-bottom:6px"><div>' + tag(t[0], t[1]) + ' ' + esc(w.msg) + (w.level === 'auto' ? '<small>Có đề xuất tự thay</small>' : '') + '</div></div>' }).join('');
    if (m.proposed) h += '<div class="meta" style="margin:8px 0 4px"><b>Bản đề xuất</b> (phần vàng là chỗ thay/thêm)</div><div class="script" style="margin:0 0 8px;background:var(--ok-bg)"><pre>' + mark(m.proposed, LIBSCAN.amountsOf(m.proposed).filter(function (a) { return !oth[a.val] }), function () { return 'mk-new' }).replace(/(\n\n)(🎁[^]*)$/, '$1<span class="mk-new">$2</span>') + '</pre></div>';
    h += '<div class="acts" style="flex-wrap:wrap">' + (m.canApply ? '<button class="btn b1" data-scone="' + m.id + '" style="min-height:40px"' + (SCN.applying ? ' disabled' : '') + '>Áp đề xuất</button>' : '') + '<button class="btn b3" data-scman="' + m.id + '" style="min-height:40px">Sửa tay</button><button class="btn b3" data-scskip="' + m.id + '" style="min-height:40px">Bỏ qua</button></div></div></details>';
    return h
  }
  function tag(c, t) { return '<span class="tag ' + c + '">' + esc(t) + '</span>' }
  function byId(id) { return SCN.res.msgs.filter(function (m) { return m.id === Number(id) })[0] }
  function replace(m, text, why) { return rpc('u_lib_replace_text', { p_id: m.id, p_old: m.text, p_new: text, p_ly_do: why, p_req: rid() }) }
  function applyOne(id) { var m = byId(id); if (!m || !m.proposed) return; SCN.applying = true; render(); replace(m, m.proposed, 'Thay theo quét giá & CTKM').then(function () { SCN.applying = false; toast('Đã thay nội dung mẫu "' + m.title + '"'); return scan() }, function (e) { SCN.applying = false; render(); fail(e) }) }
  function applySel() {
    var ids = Object.keys(SCN.sel); if (!ids.length) return; if (!confirm('Thay nội dung ' + ids.length + ' mẫu theo đề xuất? Nội dung cũ được giữ trong nhật ký.')) return; SCN.applying = true; render(); var ok = 0, bad = [];
    ids.reduce(function (p, id) { return p.then(function () { var m = byId(id); return replace(m, m.proposed, 'Thay theo quét giá & CTKM').then(function () { ok++ }, function (e) { bad.push(m.title + ': ' + (e.message || e.code)) }) }) }, Promise.resolve()).then(function () { SCN.applying = false; toast('Đã thay ' + ok + ' mẫu' + (bad.length ? ', ' + bad.length + ' mẫu không thay được: ' + bad[0] : ''), 0, bad.length ? 1 : 0); return scan() })
  }
  function manual(id) {
    var m = byId(id); if (!m) return; var d = document.createElement('div'); d.className = 'modal';
    d.innerHTML = '<div class="card" style="width:100%;max-width:640px;border-radius:22px 22px 0 0;padding:18px;max-height:92vh;overflow:auto"><h2 style="font-size:18px;margin:0 0 4px">Sửa tay · ' + esc(m.title) + '</h2><div class="meta" style="margin-bottom:8px">Sửa nội dung rồi Lưu. Máy chủ kiểm lại: nếu mẫu vừa bị đổi ở nơi khác thì không ghi đè.</div>' + (m.warns.length ? '<div style="margin-bottom:8px">' + m.warns.slice(0, 4).map(function (w) { return '<div class="meta">• ' + esc(w.msg) + '</div>' }).join('') + '</div>' : '') + '<textarea id="sc-t" style="min-height:220px">' + esc(m.text) + '</textarea><div class="acts" style="margin-top:10px"><button class="btn b3" id="sc-x" style="flex:1">Hủy</button>' + (m.proposed ? '<button class="btn b2" id="sc-p" style="flex:1">Dùng bản đề xuất</button>' : '') + '<button class="btn b1" id="sc-s" style="flex:1.4">Lưu</button></div></div>';
    document.body.appendChild(d); d.onclick = function (e) { if (e.target === d) d.remove() }; d.querySelector('#sc-x').onclick = function () { d.remove() };
    var p = d.querySelector('#sc-p'); if (p) p.onclick = function () { d.querySelector('#sc-t').value = m.proposed };
    d.querySelector('#sc-s').onclick = function () { var v = d.querySelector('#sc-t').value, b = d.querySelector('#sc-s'); if (v === m.text) { d.remove(); return } b.disabled = true; replace(m, v, 'Sửa tay sau quét giá & CTKM').then(function () { d.remove(); toast('Đã lưu nội dung mẫu'); return scan() }, function (e) { b.disabled = false; fail(e) }) }
  }
  function bind(on) {
    on('#sc-inc', 'onchange', function (x) { SCN.inc = x.checked }); on('#sc-run', 'onclick', scan);
    [].slice.call(document.querySelectorAll('[data-scf]')).forEach(function (b) { b.onclick = function () { SCN.filter = b.dataset.scf; render() } });
    on('#sc-all', 'onclick', function () { SCN.res.msgs.forEach(function (m) { if (m.canApply && !SCN.skip[m.id]) SCN.sel[m.id] = 1 }); render() }); on('#sc-none', 'onclick', function () { SCN.sel = {}; render() }); on('#sc-apply', 'onclick', applySel);
    [].slice.call(document.querySelectorAll('[data-scsel]')).forEach(function (c) { c.onclick = function (e) { e.stopPropagation() }; c.onchange = function () { c.checked ? SCN.sel[c.dataset.scsel] = 1 : delete SCN.sel[c.dataset.scsel]; var b = document.getElementById('sc-apply'); if (b) { var n = Object.keys(SCN.sel).length; b.disabled = !n; b.textContent = 'Thay các mẫu đã tích (' + n + ')' } } });
    [].slice.call(document.querySelectorAll('details[data-scid]')).forEach(function (d) { d.ontoggle = function () { d.open ? SCN.open[d.dataset.scid] = 1 : delete SCN.open[d.dataset.scid] } });
    [].slice.call(document.querySelectorAll('[data-scone]')).forEach(function (b) { b.onclick = function () { applyOne(b.dataset.scone) } });
    [].slice.call(document.querySelectorAll('[data-scman]')).forEach(function (b) { b.onclick = function () { manual(b.dataset.scman) } });
    [].slice.call(document.querySelectorAll('[data-scskip]')).forEach(function (b) { b.onclick = function () { SCN.skip[b.dataset.scskip] = 1; render() } });
    [].slice.call(document.querySelectorAll('[data-sccp]')).forEach(function (b) { b.onclick = function () { copy(SCN.res.unmentioned[Number(b.dataset.sccp)].sent); toast('Đã chép câu gợi ý') } })
  }
  return { page: page, bind: bind, state: SCN, scan: scan }
})();

;

/* ===== asm-hr.js ===== */
/* NHÂN SỰ – NẠP DANH SÁCH TỪ FILE TOÀN BỘ NHÂN SỰ: kéo thả file Excel/CSV, app tự tìm dòng tiêu đề, nhận ra cột Họ tên / Email / Chức danh / Shop,
 * so với tài khoản đang có rồi cho xem trước: Tạo mới / Cập nhật / Không đổi / Lỗi. Bấm Áp dụng mới ghi (từng người, qua hàm tạo/sửa tài khoản chỉ ASM gọi được).
 * Tài khoản mới có mật khẩu ngẫu nhiên, hiện một lần và tải được thành file CSV để gửi nhân viên. Không tự khóa người vắng trong file. */
var ASMHR = (function () {
  var P = window.LIBP, H = { rows: null, name: '', info: null, thay: false, res: null, busy: false, drag: false };
  var COLS = [
    { k: 'ho_ten', nhan: 'Họ tên', alias: ['ho va ten', 'ho ten', 'ten nhan vien', 'nhan vien', 'ten', 'ho ten nhan vien', 'full name', 'name'], bb: 1 },
    { k: 'email', nhan: 'Email', alias: ['mail', 'dia chi email', 'email cong ty', 'email nhan vien', 'email lien he'], bb: 1 },
    { k: 'chuc_danh', nhan: 'Chức danh', alias: ['vai tro', 'chuc vu', 'vi tri', 'chuc vu cong viec', 'position', 'role'] },
    { k: 'shop', nhan: 'Shop', alias: ['ma shop', 'ten shop', 'co so', 'trung tam', 'cua hang', 'dia diem lam viec', 'dia diem', 'don vi', 'noi lam viec', 'ma co so', 'ten co so', 'ten trung tam'] },
    { k: 'sdt', nhan: 'Số điện thoại', alias: ['sdt', 'dien thoai', 'so dt', 'phone'] }, { k: 'ma_nv', nhan: 'Mã nhân viên', alias: ['ma nv', 'ma so nhan vien', 'ma nhan vien'] }];
  var KQ = { moi: ['t-ok', 'Tạo mới'], sua: ['t-today', 'Cập nhật'], giu: ['t-new', 'Không đổi'], loi: ['t-bad', 'Lỗi'], bo: ['t-new', 'Bỏ qua'] };

  function xlsx(cb) { if (window.XLSX) return cb(); var s = document.createElement('script'); s.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'; s.onload = cb; s.onerror = function () { toast('Không tải được bộ đọc Excel. Kiểm tra mạng.', 0, 1) }; document.head.appendChild(s) }
  function genPw() { var a = 'abcdefghjkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789', o = '', b = new Uint32Array(10); (window.crypto || window.msCrypto).getRandomValues(b); for (var i = 0; i < 10; i++) o += a[b[i] % a.length]; return o + '7' }
  function tag(c, t) { return '<span class="tag ' + c + '">' + esc(t) + '</span>' }
  function accs() { return Object.keys(D.acc).map(function (k) { return D.acc[k] }) }
  function shopsAct() { return Object.keys(D.shops).map(function (k) { return D.shops[k] }).filter(function (s) { return s.trang_thai === 'HOAT_DONG' }) }
  function core(s) { return P.bo(String(s || '').replace(/\([^)]*\)/g, '')).replace(/^(shop|trung tam|nha thuoc|co so)\s+/, '').trim() }
  function findShop(v) {
    var t = P.bo(v); if (!t) return { err: 'Thiếu Shop' }; var all = shopsAct(), hit = all.filter(function (s) { return P.bo(s.ma_shop) === t || P.bo(s.id_shop) === t });
    if (!hit.length) { var tc = core(v); hit = all.filter(function (s) { var c = core(s.ten_shop); return c && (c === tc || (tc.length >= 3 && (c.indexOf(tc) >= 0 || tc.indexOf(c) >= 0))) }) }
    if (!hit.length) hit = all.filter(function (s) { return s.ma_shop && t.replace(/ /g, '').indexOf(String(s.ma_shop)) >= 0 });
    return hit.length === 1 ? { s: hit[0] } : { err: hit.length ? 'Nhiều Shop khớp "' + v + '"' : 'Không thấy Shop "' + v + '"' }
  }
  function role(v) { var t = P.bo(v); if (!t) return 'NHAN_VIEN'; if (/(^| )(asm|giam doc)( |$)/.test(t)) return 'ASM'; return /quan ly|(^| )ql( |$)|cua hang truong|truong (shop|trung tam|nha thuoc)/.test(t) ? 'QUAN_LY' : 'NHAN_VIEN' }
  var VT = { QUAN_LY: 'Quản lý', NHAN_VIEN: 'Nhân viên' };

  /* tìm dòng tiêu đề: khớp đúng tên/alias trước, không có thì khớp chứa cụm (ví dụ "Cơ sở làm việc" chứa "cơ sở") */
  function findHdr(aoa) {
    var best = null;
    for (var r = 0; r < Math.min(aoa.length, 40); r++) {
      var row = aoa[r] || [], map = {}, taken = {}, n = 0;
      COLS.forEach(function (c) {
        var names = [c.nhan, c.k].concat(c.alias || []).map(function (a) { return P.bo(a) }), pick = -1, j, b;
        for (j = 0; j < row.length && pick < 0; j++) if (!taken[j] && row[j] != null && names.indexOf(P.bo(row[j])) >= 0) pick = j;
        for (j = 0; j < row.length && pick < 0; j++) { if (taken[j] || row[j] == null) continue; b = ' ' + P.bo(row[j]) + ' '; if (b.length > 3 && names.some(function (a) { return a.length >= 4 && b.indexOf(' ' + a + ' ') >= 0 })) pick = j }
        if (pick >= 0) { map[c.k] = pick; taken[pick] = 1; n++ }
      });
      if (!best || n > best.n) best = { r: r, map: map, n: n }
    }
    return best || { r: 0, map: {}, n: 0 }
  }
  function analyse(aoa) {
    var hd = findHdr(aoa), miss = COLS.filter(function (c) { return c.bb && hd.map[c.k] === undefined });
    if (!hd.n || miss.length) return { loi: 'Không nhận ra cột bắt buộc: ' + miss.map(function (c) { return c.nhan }).join(', ') + '. Các cột đọc được ở dòng ' + (hd.r + 1) + ': ' + (aoa[hd.r] || []).filter(function (v) { return v != null && String(v).trim() }).slice(0, 12).join(' | ') };
    var byEmail = {}; accs().forEach(function (a) { byEmail[String(a.email).toLowerCase()] = a }); var seen = {}, rows = [], cell = function (r, k) { var j = hd.map[k]; return j === undefined ? '' : String(r[j] == null ? '' : r[j]).trim() };
    for (var i = hd.r + 1; i < aoa.length; i++) {
      var r = aoa[i] || []; if (!r.some(function (v) { return v != null && String(v).trim() })) continue;
      var o = { dong: i + 1, ho_ten: cell(r, 'ho_ten'), email: cell(r, 'email').toLowerCase().replace(/\s+/g, ''), chuc_danh: cell(r, 'chuc_danh'), shopTxt: cell(r, 'shop'), vai_tro: role(cell(r, 'chuc_danh')) };
      if (!o.ho_ten && !o.email) continue;
      if (!o.ho_ten) o.kq = 'loi', o.ghi = 'Thiếu họ tên';
      else if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(o.email)) o.kq = 'loi', o.ghi = 'Email không hợp lệ';
      else if (seen[o.email]) o.kq = 'loi', o.ghi = 'Trùng email với dòng ' + seen[o.email];
      else if (o.vai_tro === 'ASM') o.kq = 'bo', o.ghi = 'Chức danh ASM, không tạo ở đây';
      else {
        seen[o.email] = o.dong; var ex = byEmail[o.email], sh = findShop(o.shopTxt);
        if (ex && ex.vai_tro === 'ASM') o.kq = 'bo', o.ghi = 'Là tài khoản ASM';
        else if (sh.err && !(ex && !o.shopTxt)) o.kq = 'loi', o.ghi = sh.err;
        else {
          o.id_shop = sh.s ? sh.s.id_shop : ex.id_shop;
          if (!ex) { o.kq = 'moi'; o.ghi = 'Tài khoản mới · ' + VT[o.vai_tro] + ' · ' + shopName(o.id_shop); o.mk = genPw() }
          else { o.ma_tk = ex.ma_tk; var ch = []; if (ex.ho_ten !== o.ho_ten) ch.push('tên'); if (ex.id_shop !== o.id_shop) ch.push('Shop ' + shopName(ex.id_shop) + ' → ' + shopName(o.id_shop)); if (cell(r, 'chuc_danh') && ex.vai_tro !== o.vai_tro) ch.push(VT[ex.vai_tro] + ' → ' + VT[o.vai_tro]); if (!cell(r, 'chuc_danh')) o.vai_tro = ex.vai_tro; o.kq = ch.length ? 'sua' : 'giu'; o.ghi = ch.length ? 'Đổi: ' + ch.join(', ') : 'Giống tài khoản hiện có' + (ex.trang_thai === 'NGUNG' ? ' (đang bị khóa)' : '') }
        }
      }
      rows.push(o)
    }
    var inFile = {}; rows.forEach(function (o) { inFile[o.email] = 1 }); var vang = accs().filter(function (a) { return a.vai_tro !== 'ASM' && a.trang_thai === 'HOAT_DONG' && !inFile[String(a.email).toLowerCase()] }).length;
    return { rows: rows, vang: vang, sheet: '', dongTieuDe: hd.r + 1 }
  }
  function handle(name, buf) {
    xlsx(function () {
      try {
        var wb = XLSX.read(buf, { type: 'array', cellDates: true, codepage: 65001 }), best = null;
        wb.SheetNames.forEach(function (n) { var aoa = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null, blankrows: true }), hd = findHdr(aoa); if (!best || hd.n > best.n) best = { n: hd.n, aoa: aoa, name: n } });
        H.info = analyse(best.aoa); H.info.sheet = best.name; H.name = name; H.res = null; render()
      } catch (e) { toast('Không đọc được file: ' + e.message, 0, 1) }
    })
  }
  function card() {
    var h = '<div class="card"><h2>Nạp danh sách nhân sự từ file <span class="pill">kéo file toàn bộ nhân sự vào đây</span></h2>';
    h += '<div id="hr-drop" style="border:2px dashed ' + (H.drag ? 'var(--teal)' : 'var(--line)') + ';border-radius:16px;padding:18px;text-align:center;background:' + (H.drag ? 'var(--tint)' : 'var(--bg)') + '"><div style="font-weight:600">Kéo thả file Excel/CSV vào đây</div><div class="meta" style="margin:4px 0 10px">App tự nhận cột Họ tên, Email, Chức danh, Shop (tên cột khác một chút vẫn nhận). Xem trước rồi mới tạo.</div><label class="btn b2" style="cursor:pointer">' + ic('upload') + 'Hoặc chọn file<input type="file" id="hr-file" accept=".xlsx,.xls,.csv" hidden></label></div>';
    var f = H.info; if (f && f.loi) h += '<div class="err" style="margin-top:10px">' + esc(f.loi) + '</div>';
    if (f && f.rows) {
      var c = { moi: 0, sua: 0, giu: 0, loi: 0, bo: 0 }; f.rows.forEach(function (o) { c[o.kq]++ }); var todo = c.moi + c.sua;
      h += '<div class="meta" style="margin:10px 0 6px">' + esc(H.name) + (f.sheet ? ' · sheet ' + esc(f.sheet) : '') + ' · dòng tiêu đề ' + f.dongTieuDe + ' · ' + f.rows.length + ' dòng</div><div class="stats stats4" style="grid-template-columns:repeat(5,1fr)">' + [['Tạo mới', c.moi, 'var(--ok)'], ['Cập nhật', c.sua, 'var(--teal-d)'], ['Không đổi', c.giu, 'inherit'], ['Lỗi', c.loi, c.loi ? 'var(--bad)' : 'inherit'], ['Bỏ qua', c.bo, 'inherit']].map(function (x) { return '<div class="card st"><div class="l">' + x[0] + '</div><div class="v" style="color:' + x[2] + '">' + x[1] + '</div></div>' }).join('') + '</div>';
      if (f.vang) h += '<div class="alert"><i>ℹ</i><div><b>' + f.vang + ' tài khoản đang hoạt động không có trong file</b><small>App không tự khóa họ. Nếu họ đã nghỉ, khóa ở bảng bên dưới.</small></div></div>';
      h += '<label class="sw" style="margin:8px 0"><input type="checkbox" id="hr-thay" ' + (H.thay ? 'checked' : '') + '><div><b>Cho phép thay Quản lý cũ</b><small>Nếu file ghi một Quản lý mới cho Shop đã có Quản lý khác, người cũ chuyển về Nhân viên. Mỗi Shop chỉ có một Quản lý. Tắt thì dòng đó báo lỗi.</small></div></label>';
      h += '<div class="tw"><table><thead><tr><th>Dòng</th><th>Kết quả</th><th>Họ tên</th><th>Email</th><th>Chức danh</th><th>Shop</th><th>Ghi chú</th></tr></thead><tbody>' + f.rows.map(function (o) { var k = KQ[o.kq]; return '<tr><td>' + o.dong + '</td><td>' + tag(k[0], k[1]) + '</td><td><b>' + esc(o.ho_ten) + '</b></td><td>' + esc(o.email) + '</td><td>' + esc(o.id_shop ? VT[o.vai_tro] : o.chuc_danh) + '</td><td>' + esc(o.id_shop ? shopName(o.id_shop) : o.shopTxt) + '</td><td class="meta" style="white-space:normal">' + esc(o.ghi || '') + '</td></tr>' }).join('') + '</tbody></table></div>';
      h += '<div class="acts" style="margin-top:10px"><button class="btn b1" id="hr-go"' + (todo && !H.busy ? '' : ' disabled') + '>' + (H.busy ? 'Đang áp dụng…' : 'Áp dụng ' + todo + ' thay đổi (' + c.moi + ' mới, ' + c.sua + ' cập nhật)') + '</button><button class="btn b3" id="hr-drop-x">Bỏ file này</button></div>'
    }
    if (H.res) {
      var ok = H.res.filter(function (x) { return x.ok }), bad = H.res.filter(function (x) { return !x.ok }), made = ok.filter(function (x) { return x.mk });
      h += '<div class="alert" style="margin-top:10px"><i>✓</i><div><b>Đã xong: ' + ok.length + ' thành công' + (bad.length ? ', ' + bad.length + ' lỗi' : '') + '</b><small>' + made.length + ' tài khoản mới. Mật khẩu ban đầu chỉ hiện ở đây, tải file đăng nhập để gửi nhân viên và nhắc họ đổi mật khẩu.</small></div>' + (made.length ? '<button class="btn b1" id="hr-csv" style="min-height:40px;margin-left:auto">' + ic('download') + 'Tải file đăng nhập (CSV)</button>' : '') + '</div>';
      if (bad.length) h += '<div class="err">' + bad.slice(0, 8).map(function (x) { return 'Dòng ' + x.dong + ' · ' + esc(x.ho_ten) + ': ' + esc(x.msg) }).join('<br>') + '</div>'
    }
    return h + '</div>'
  }
  function apply() {
    var todo = H.info.rows.filter(function (o) { return o.kq === 'moi' || o.kq === 'sua' }); if (!todo.length || H.busy) return; if (!confirm('Áp dụng ' + todo.length + ' thay đổi nhân sự?')) return; H.busy = true; H.res = []; render();
    todo.reduce(function (p, o) {
      return p.then(function () {
        var req = o.kq === 'moi' ? rpc('u_asm_account_create', { p_ho_ten: o.ho_ten, p_email: o.email, p_vai_tro: o.vai_tro, p_id_shop: o.id_shop, p_mat_khau: o.mk, p_thay_ql: H.thay, p_req: rid() }) : rpc('u_asm_account_update', { p_ma_tk: o.ma_tk, p_ho_ten: o.ho_ten, p_vai_tro: o.vai_tro, p_id_shop: o.id_shop, p_trang_thai: (D.acc[o.ma_tk] || {}).trang_thai || 'HOAT_DONG', p_thay_ql: H.thay, p_req: rid() });
        return req.then(function () { H.res.push({ ok: true, dong: o.dong, ho_ten: o.ho_ten, email: o.email, mk: o.kq === 'moi' ? o.mk : '', vai_tro: o.vai_tro, id_shop: o.id_shop }) }, function (e) { H.res.push({ ok: false, dong: o.dong, ho_ten: o.ho_ten, msg: (e && e.message) || (e && e.code) || 'Lỗi' }) })
      })
    }, Promise.resolve()).then(function () { H.busy = false; H.info = null; toast('Đã áp dụng danh sách nhân sự'); return load(true).then(render) })
  }
  function csv() {
    var rows = H.res.filter(function (x) { return x.ok && x.mk }), t = 'ho_ten,email,mat_khau_ban_dau,vai_tro,shop\n' + rows.map(function (x) { return [x.ho_ten, x.email, x.mk, VT[x.vai_tro], shopName(x.id_shop)].map(function (v) { v = String(v == null ? '' : v); return /[",\n;]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v }).join(',') }).join('\n');
    var a = document.createElement('a'); a.href = URL.createObjectURL(new Blob(['﻿' + t], { type: 'text/csv;charset=utf-8' })); a.download = 'dang-nhap-nhan-su-' + new Date().toISOString().slice(0, 10) + '.csv'; document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove() }, 500)
  }
  function read(f) { var rd = new FileReader(); rd.onload = function () { handle(f.name, rd.result) }; rd.readAsArrayBuffer(f) }
  function bind(on) {
    xlsx(function () { });
    var dz = document.getElementById('hr-drop');
    if (dz) { ['dragenter', 'dragover'].forEach(function (ev) { dz.addEventListener(ev, function (e) { e.preventDefault(); if (!H.drag) { H.drag = true; dz.style.borderColor = 'var(--teal)'; dz.style.background = 'var(--tint)' } }) }); ['dragleave', 'drop'].forEach(function (ev) { dz.addEventListener(ev, function (e) { e.preventDefault(); H.drag = false; dz.style.borderColor = 'var(--line)'; dz.style.background = 'var(--bg)' }) }); dz.addEventListener('drop', function (e) { var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]; if (f) read(f) }) }
    on('#hr-file', 'onchange', function (x) { var f = x.files[0]; if (f) read(f) }); on('#hr-thay', 'onchange', function (x) { H.thay = x.checked }); on('#hr-go', 'onclick', apply); on('#hr-csv', 'onclick', csv); on('#hr-drop-x', 'onclick', function () { H.info = null; H.res = null; render() })
  }
  window.addEventListener('dragover', function (e) { e.preventDefault() }); window.addEventListener('drop', function (e) { e.preventDefault() });
  return { card: card, bind: bind, handle: handle, analyse: analyse, state: H }
})();

;

/* ===== asm-img.js ===== */
/* ẢNH, TỜ RƠI, VOUCHER: kéo thả ảnh/PDF (hoặc cả thư mục) vào, app TỰ NHẬN DIỆN thư mục, chủ đề vaccine, nhóm đối tượng từ tên file/thư mục
 * (anh sửa được trước khi tải lên), kiểm trùng bằng SHA-256, tạo ảnh nhỏ, tải lên kho lưu trữ RIÊNG TƯ của Supabase, ghi thông tin qua hàm chỉ ASM gọi được.
 * Nhân viên dùng ảnh ở màn Gợi ý tin nhắn (msg.js): app gợi ý ảnh hợp với khách và vaccine, nhân viên bấm gửi (chia sẻ sang Zalo) hoặc tải về.
 * Dữ liệu và ảnh thử là dữ liệu mẫu. */
function fileBlob(f) { return tok().then(function (t) { return fetch(URL_ + '/storage/v1/object/authenticated/lib-files/' + f.duong_dan, { headers: { apikey: KEY, Authorization: 'Bearer ' + t } }) }).then(function (r) { if (!r.ok) throw { code: 'LOAD', message: 'Không mở được file (' + r.status + ').' }; return r.blob() }) }
var ASMIMG = (function () {
  var IM = { files: null, loading: false, q: '', thu: '', hidden: false, queue: [], busy: false, prog: '', n: 0 };
  var THU = { VACCINE_DOI_TUONG: 'Vaccine theo đối tượng', CAM_NANG_BAU: 'Cẩm nang chăm sóc bầu', CTKM_VOUCHER: 'Khuyến mãi & voucher', KHAC: 'Khác' };
  var NHOM = { '': '—', BAU: 'Mẹ bầu', NHI: 'Nhi (dưới 9 tuổi)', TRE_LON: 'Trẻ lớn / thanh thiếu niên', NGUOI_LON: 'Người lớn', TAT_CA: 'Tất cả' };
  var VOCAB = [['6 trong 1', ['6in1', '6 in 1', '6 trong 1', 'hexaxim', 'infanrix']], ['Phế cầu', ['phe cau', 'prevenar', 'synflorix', 'vaxneuvance', 'pcv']], ['Rota', ['rota', 'rotarix', 'rotateq', 'rotavin']], ['HPV', ['hpv', 'gardasil']], ['Cúm', ['cum', 'influvac', 'vaxigrip']],
    ['Sởi – Quai bị – Rubella', ['soi', 'mmr', 'rubella', 'quai bi', 'priorix']], ['Thủy đậu', ['thuy dau', 'varivax', 'varilrix']], ['Viêm gan', ['viem gan', 'vgb', 'vga', 'gan b', 'gan a', 'twinrix']], ['Não mô cầu', ['nao mo cau', 'bexsero', 'nimenrix', 'menactra', 'menquadfi', 'vega']], ['Uốn ván – Bạch hầu – Ho gà', ['uon van', 'bach hau', 'ho ga', 'adacel', 'boostrix', '3in1', 'tdap']],
    ['RSV', ['rsv', 'abrysvo', 'beyfortus']], ['Sốt xuất huyết', ['sot xuat huyet', 'sxh', 'qdenga']], ['Zona', ['zona', 'shingrix']], ['Lao', ['lao', 'bcg']], ['Não Nhật Bản', ['nhat ban', 'jevax', 'imojev']], ['Tả – Thương hàn', ['thuong han', 'typhim', 'vi ta']], ['Dại', ['dai', 'verorab', 'speeda']], ['Bại liệt', ['bai liet', 'ipv', 'opv']], ['Viêm não', ['viem nao']]];
  var ALLOWED = { 'image/jpeg': 1, 'image/png': 1, 'image/webp': 1, 'application/pdf': 1 }, EXT = { jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png', webp: 'image/webp', pdf: 'application/pdf' };
  try { var st = document.createElement('style'); st.textContent = '.gal{grid-template-columns:repeat(auto-fill,minmax(150px,1fr))!important}.gal .card{padding:8px;display:grid;gap:6px;align-content:start}.gal img,.gal .pdf{width:100%;height:110px;object-fit:cover;border-radius:10px;background:var(--bg)}.gal .pdf{display:grid;place-items:center;font-weight:700;color:var(--soft)}.qthumb{width:54px;height:54px;object-fit:cover;border-radius:8px;background:var(--bg)}.qt input,.qt select{border:1px solid var(--line);border-radius:8px;padding:6px;background:var(--card);min-height:36px;width:100%;min-width:110px}'; document.head.appendChild(st) } catch (e) { }

  function fold(s) { return window.LIBSCAN ? LIBSCAN.fold(s) : String(s).toLowerCase() }
  function tag(c, t) { return '<span class="tag ' + c + '">' + esc(t) + '</span>' }
  function detect(name, rel) {
    var t = ' ' + fold((rel || '') + ' ' + name).replace(/[^a-z0-9]+/g, ' ') + ' ', has = function (arr) { return arr.some(function (k) { return t.indexOf(' ' + k + ' ') >= 0 || (k.length >= 5 && t.indexOf(k) >= 0) }) };
    var thu = has(['voucher', 'to roi', 'ctkm', 'khuyen mai', 'uu dai', 'combo', 'goi tiem', 'qua tang', 'giam gia', 'sale']) ? 'CTKM_VOUCHER' : has(['bau', 'thai', 'cam nang', 'sau sinh', 'me va be', 'me bau']) ? 'CAM_NANG_BAU' : 'VACCINE_DOI_TUONG';
    var cd = ''; VOCAB.some(function (v) { if (has(v[1])) { cd = v[0]; return true } return false });
    var nhom = has(['bau', 'thai', 'me bau']) ? 'BAU' : has(['tre lon', 'thanh thieu nien', '9 tuoi', 'vi thanh nien']) ? 'TRE_LON' : has(['nguoi lon', 'cao tuoi', 'truong thanh']) ? 'NGUOI_LON' : has(['nhi', 'tre em', 'so sinh', 'be ']) ? 'NHI' : 'TAT_CA';
    return { thu: thu, chu_de: cd, nhom: nhom, ten: String(name).replace(/\.[^.]+$/, '').replace(/[_]+/g, ' ').replace(/\s+/g, ' ').trim() }
  }
  function sha(buf) { return crypto.subtle.digest('SHA-256', buf).then(function (h) { return [].slice.call(new Uint8Array(h)).map(function (b) { return ('0' + b.toString(16)).slice(-2) }).join('') }) }
  function thumb(file) {
    return new Promise(function (res) {
      if (file.type === 'application/pdf') return res(null); var url = URL.createObjectURL(file), im = new Image();
      im.onload = function () { try { var m = 150, k = Math.min(1, m / Math.max(im.width, im.height)), c = document.createElement('canvas'); c.width = Math.max(1, Math.round(im.width * k)); c.height = Math.max(1, Math.round(im.height * k)); var g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height); g.drawImage(im, 0, 0, c.width, c.height); var q = .7, d = c.toDataURL('image/jpeg', q); while (d.length > 38000 && q > .3) { q -= .1; d = c.toDataURL('image/jpeg', q) } URL.revokeObjectURL(url); res(d.length <= 40000 ? d : null) } catch (e) { res(null) } };
      im.onerror = function () { URL.revokeObjectURL(url); res(null) }; im.src = url
    })
  }
  function loadFiles() { if (IM.loading) return; IM.loading = true; return rest('lib_files?select=*&order=thoi_gian.desc&limit=500').then(function (a) { IM.files = a; IM.loading = false; render() }, function (e) { IM.loading = false; IM.files = IM.files || []; fail(e); render() }) }
  function mimeOf(f) { return ALLOWED[f.type] ? f.type : EXT[(f.name.split('.').pop() || '').toLowerCase()] || '' }

  /* ---------- nhận file ---------- */
  function addFiles(items) {
    var list = items.filter(function (x) { return x.file }); if (!list.length) return; IM.busy = true; var total = list.length, done = 0;
    function next() {
      if (!list.length) { IM.busy = false; IM.prog = ''; render(); return }
      var it = list.shift(), f = it.file, mime = mimeOf(f), q = { id: 'q' + (++IM.n), file: f, rel: it.rel || '', mime: mime, size: f.size, inc: true, err: '', dup: '', thumb: null }; IM.prog = 'Đang đọc ' + (++done) + '/' + total + '…'; var d = detect(f.name, q.rel); q.ten = d.ten; q.thu = d.thu; q.chu_de = d.chu_de; q.nhom = d.nhom;
      if (!mime) { q.err = 'Định dạng không hỗ trợ (chỉ JPG, PNG, WEBP, PDF)'; q.inc = false; IM.queue.push(q); return next() }
      if (f.size > 10485760) { q.err = 'Quá 10 MB'; q.inc = false; IM.queue.push(q); return next() }
      f.arrayBuffer().then(sha).then(function (h) { q.sha = h; var ex = (IM.files || []).filter(function (x) { return x.sha256 === h && x.dang_dung === 'CO' })[0] || null; if (ex) q.dup = 'Trùng với "' + ex.ten + '" đã có trong kho'; else if (IM.queue.some(function (o) { return o.sha === h })) q.dup = 'Trùng với file khác trong danh sách này'; if (q.dup) q.inc = false; return thumb(f) }).then(function (t) { q.thumb = t; IM.queue.push(q); if (done % 5 === 0) render(); next() }, function () { q.err = 'Không đọc được file'; q.inc = false; IM.queue.push(q); next() })
    }
    next(); render()
  }
  function walk(entry, path, out) {
    return new Promise(function (res) {
      if (entry.isFile) entry.file(function (f) { out.push({ file: f, rel: path }); res() }, function () { res() });
      else { var rd = entry.createReader(), all = []; (function more() { rd.readEntries(function (es) { if (!es.length) { Promise.all(all.map(function (e) { return walk(e, path + '/' + e.name, out) })).then(res); return } all = all.concat([].slice.call(es)); more() }, function () { res() }) })() }
    })
  }
  function onDrop(e) {
    var dt = e.dataTransfer, out = [], ps = []; if (dt && dt.items && dt.items.length && dt.items[0].webkitGetAsEntry) { [].slice.call(dt.items).forEach(function (it) { var en = it.webkitGetAsEntry && it.webkitGetAsEntry(); if (en) ps.push(walk(en, en.name, out)) }); Promise.all(ps).then(function () { addFiles(out) }) }
    else if (dt && dt.files) addFiles([].slice.call(dt.files).map(function (f) { return { file: f, rel: '' } }))
  }

  /* ---------- tải lên ---------- */
  function upload() {
    var todo = IM.queue.filter(function (q) { return q.inc && !q.err }); if (!todo.length || IM.busy) return; IM.busy = true; var ok = 0, bad = [], n = 0;
    todo.reduce(function (p, q) {
      return p.then(function () {
        IM.prog = 'Đang tải lên ' + (++n) + '/' + todo.length + '…'; render(); var safe = fold(q.file.name).replace(/[^a-z0-9._-]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 60) || 'file', path = 'files/' + Date.now().toString(36) + (Math.random().toString(36).slice(2, 6)) + '-' + safe;
        return tok().then(function (t) { return fetch(URL_ + '/storage/v1/object/lib-files/' + path, { method: 'POST', headers: { apikey: KEY, Authorization: 'Bearer ' + t, 'Content-Type': q.mime, 'cache-control': '3600' }, body: q.file }) }).then(function (r) { if (!r.ok) return r.text().then(function (x) { throw { message: 'Kho lưu trữ từ chối (' + r.status + ')' } }) })
          .then(function () { return rpc('u_file_add', { p_ten: q.ten, p_ten_file: q.file.name, p_thu_muc: q.thu, p_chu_de: q.chu_de, p_nhom: q.nhom, p_mo_ta: q.rel || '', p_mime: q.mime, p_kich_thuoc: q.size, p_sha: q.sha, p_duong_dan: path, p_anh_nho: q.thumb, p_cho_trung: false, p_req: rid() }).catch(function (e) { tok().then(function (t) { fetch(URL_ + '/storage/v1/object/lib-files/' + path, { method: 'DELETE', headers: { apikey: KEY, Authorization: 'Bearer ' + t } }) }); throw e }) })
          .then(function () { ok++; q.done = true }, function (e) { bad.push(q.ten + ': ' + ((e && e.message) || (e && e.code) || 'lỗi')); q.err = (e && e.message) || 'Lỗi'; q.inc = false })
      })
    }, Promise.resolve()).then(function () { IM.busy = false; IM.prog = ''; IM.queue = IM.queue.filter(function (q) { return !q.done }); toast('Đã tải lên ' + ok + ' file' + (bad.length ? ', ' + bad.length + ' file lỗi: ' + bad[0] : ''), 0, bad.length ? 1 : 0); return loadFiles() })
  }

  /* ---------- giao diện ---------- */
  function select(name, opts, val, attr) { return '<select ' + attr + '>' + Object.keys(opts).map(function (k) { return '<option value="' + k + '"' + (k === val ? ' selected' : '') + '>' + esc(opts[k]) + '</option>' }).join('') + '</select>' }
  function page() {
    if (IM.files === null) return '<div class="card empty">Đang tải kho ảnh…</div>';
    var h = '<div class="card"><h2>Tải ảnh, tờ rơi, voucher <span class="pill">tự nhận diện từ tên file và tên thư mục</span></h2><div id="im-drop" style="border:2px dashed var(--line);border-radius:16px;padding:18px;text-align:center;background:var(--bg)"><div style="font-weight:600">Kéo thả ảnh/PDF hoặc cả thư mục vào đây</div><div class="meta" style="margin:4px 0 10px">JPG, PNG, WEBP, PDF, tối đa 10 MB mỗi file. App đoán thư mục (Vaccine theo đối tượng / Cẩm nang bầu / Khuyến mãi & voucher), chủ đề vaccine và nhóm đối tượng; anh xem và sửa trước khi tải lên. File trùng hệt file đã có sẽ bị bỏ qua.</div><div class="row" style="justify-content:center;flex-wrap:wrap"><label class="btn b2" style="cursor:pointer">' + ic('upload') + 'Chọn ảnh/PDF<input type="file" id="im-file" multiple accept=".jpg,.jpeg,.png,.webp,.pdf,image/*,application/pdf" hidden></label><label class="btn b3" style="cursor:pointer">Chọn cả thư mục<input type="file" id="im-dir" webkitdirectory multiple hidden></label></div>' + (IM.prog ? '<div class="meta" style="margin-top:8px">' + esc(IM.prog) + '</div>' : '') + '</div></div>';
    if (IM.queue.length) {
      var inc = IM.queue.filter(function (q) { return q.inc && !q.err }).length, THs = {}; Object.keys(THU).forEach(function (k) { THs[k] = THU[k] });
      h += '<div class="card"><h2>Xem trước và xác nhận <span class="pill">' + IM.queue.length + ' file · ' + inc + ' sẽ tải lên</span></h2><datalist id="im-cd">' + VOCAB.map(function (v) { return '<option value="' + esc(v[0]) + '">' }).join('') + '</datalist><div class="tw qt"><table><thead><tr><th></th><th>Ảnh</th><th>Tên hiển thị</th><th>Thư mục</th><th>Chủ đề / vaccine</th><th>Nhóm đối tượng</th><th>Kết quả</th></tr></thead><tbody>' + IM.queue.map(function (q) {
        return '<tr><td><input type="checkbox" data-iq="' + q.id + '" data-f="inc" ' + (q.inc && !q.err ? 'checked' : '') + (q.err ? ' disabled' : '') + ' style="width:20px;height:20px"></td><td>' + (q.thumb ? '<img class="qthumb" src="' + q.thumb + '">' : '<div class="qthumb" style="display:grid;place-items:center;font-size:11px;color:var(--soft)">' + (q.mime === 'application/pdf' ? 'PDF' : '—') + '</div>') + '</td><td><input data-iq="' + q.id + '" data-f="ten" value="' + esc(q.ten) + '"><small>' + esc(q.rel || q.file.name) + '</small></td><td>' + select('', THU, q.thu, 'data-iq="' + q.id + '" data-f="thu"') + '</td><td><input list="im-cd" data-iq="' + q.id + '" data-f="chu_de" value="' + esc(q.chu_de) + '" placeholder="chưa nhận ra"></td><td>' + select('', NHOM, q.nhom, 'data-iq="' + q.id + '" data-f="nhom"') + '</td><td>' + (q.err ? tag('t-bad', q.err) : q.dup ? tag('t-late', q.dup) : tag('t-ok', 'Sẵn sàng')) + '</td></tr>'
      }).join('') + '</tbody></table></div><div class="acts" style="margin-top:10px"><button class="btn b1" id="im-up"' + (inc && !IM.busy ? '' : ' disabled') + '>Tải lên ' + inc + ' file</button><button class="btn b3" id="im-clear"' + (IM.busy ? ' disabled' : '') + '>Xóa danh sách</button></div></div>'
    }
    var all = IM.files, cnt = function (k) { return all.filter(function (f) { return (!k || f.thu_muc === k) && (IM.hidden || f.dang_dung === 'CO') }).length }, q = fold(IM.q);
    var L = all.filter(function (f) { return (!IM.thu || f.thu_muc === IM.thu) && (IM.hidden || f.dang_dung === 'CO') && (!q || fold(f.ten + ' ' + (f.chu_de || '') + ' ' + (f.ten_file || '')).indexOf(q) >= 0) });
    h += '<div class="top" style="flex-wrap:wrap"><div class="search">' + ic('search') + '<input id="im-q" placeholder="Tìm theo tên, chủ đề…" value="' + esc(IM.q) + '"></div><label class="row" style="cursor:pointer"><input type="checkbox" id="im-hid" ' + (IM.hidden ? 'checked' : '') + ' style="width:20px;height:20px;accent-color:var(--teal)"><span class="meta">Hiện cả ảnh đã ẩn</span></label></div>';
    h += '<div class="chips">' + [['', 'Tất cả']].concat(Object.keys(THU).map(function (k) { return [k, THU[k]] })).map(function (x) { return '<button class="chip ' + (IM.thu === x[0] ? 'on' : '') + '" data-imt="' + x[0] + '">' + esc(x[1]) + '<em>' + cnt(x[0]) + '</em></button>' }).join('') + '</div>';
    h += '<div class="list gal">' + (L.length ? L.map(function (f) { return '<div class="card"' + (f.dang_dung === 'KHONG' ? ' style="opacity:.55"' : '') + '>' + (f.anh_nho ? '<img src="' + f.anh_nho + '" alt="">' : '<div class="pdf">' + (f.mime === 'application/pdf' ? 'PDF' : 'FILE') + '</div>') + '<b style="font-size:13.5px;line-height:1.3">' + esc(f.ten) + '</b><div>' + tag('t-new', THU[f.thu_muc] || f.thu_muc) + (f.chu_de ? tag('t-today', f.chu_de) : '') + (f.nhom ? tag('t-new', NHOM[f.nhom]) : '') + (f.dang_dung === 'KHONG' ? tag('t-bad', 'đã ẩn') : '') + '</div><div class="row" style="flex-wrap:wrap"><button class="btn b3" data-imv="' + f.id + '" style="min-height:34px;font-size:13px;padding:0 10px">Xem</button><button class="btn b3" data-ime="' + f.id + '" style="min-height:34px;font-size:13px;padding:0 10px">Sửa</button><button class="btn b3" data-imh="' + f.id + '" style="min-height:34px;font-size:13px;padding:0 10px">' + (f.dang_dung === 'CO' ? 'Ẩn' : 'Hiện') + '</button><button class="btn b3" data-imd="' + f.id + '" style="min-height:34px;font-size:13px;padding:0 10px" aria-label="Xóa">✕</button></div></div>' }).join('') : '<div class="card empty" style="grid-column:1/-1">Chưa có ảnh nào. Kéo ảnh vào khung phía trên.</div>') + '</div>';
    return h
  }
  function byId(id) { return (IM.files || []).filter(function (f) { return f.id === id })[0] }
  function view(f) {
    var d = document.createElement('div'); d.className = 'modal'; d.innerHTML = '<div class="card" style="width:100%;max-width:720px;border-radius:22px 22px 0 0;padding:16px;max-height:94vh;overflow:auto;text-align:center"><div class="row2"><h2 style="margin:0;font-size:17px">' + esc(f.ten) + '</h2><button class="btn b3" id="v-x" style="min-height:38px">Đóng</button></div><div id="v-b" class="meta" style="padding:30px">Đang mở…</div></div>'; document.body.appendChild(d); d.onclick = function (e) { if (e.target === d) d.remove() }; d.querySelector('#v-x').onclick = function () { d.remove() };
    fileBlob(f).then(function (b) { var u = URL.createObjectURL(b), box = d.querySelector('#v-b'); box.style.padding = '10px 0'; box.innerHTML = f.mime === 'application/pdf' ? '<a class="btn b1" href="' + u + '" target="_blank" rel="noopener">Mở PDF</a>' : '<img src="' + u + '" style="max-width:100%;max-height:70vh;border-radius:12px">' }, function (e) { d.querySelector('#v-b').textContent = (e && e.message) || 'Không mở được file.' })
  }
  function edit(f) {
    var d = document.createElement('div'); d.className = 'modal'; d.innerHTML = '<div class="card" style="width:100%;max-width:520px;border-radius:22px 22px 0 0;padding:18px"><h2 style="font-size:18px;margin:0 0 10px">Sửa thông tin</h2><label class="field">Tên hiển thị<input id="e-ten" value="' + esc(f.ten) + '"></label><div class="form2" style="margin-top:0"><div><label class="meta">Thư mục</label>' + select('', THU, f.thu_muc, 'id="e-thu"') + '</div><div><label class="meta">Nhóm đối tượng</label>' + select('', NHOM, f.nhom || '', 'id="e-nhom"') + '</div><div style="grid-column:1/-1"><label class="meta">Chủ đề / vaccine</label><input id="e-cd" list="im-cd2" value="' + esc(f.chu_de || '') + '"><datalist id="im-cd2">' + VOCAB.map(function (v) { return '<option value="' + esc(v[0]) + '">' }).join('') + '</datalist></div><div style="grid-column:1/-1"><label class="meta">Ghi chú</label><input id="e-mt" value="' + esc(f.mo_ta || '') + '"></div></div><div class="acts" style="margin-top:12px"><button class="btn b3" id="e-x" style="flex:1">Hủy</button><button class="btn b1" id="e-ok" style="flex:1.4">Lưu</button></div></div>'; document.body.appendChild(d); d.onclick = function (e) { if (e.target === d) d.remove() }; d.querySelector('#e-x').onclick = function () { d.remove() };
    d.querySelector('#e-ok').onclick = function () { var b = d.querySelector('#e-ok'); b.disabled = true; save(f, { p_ten: d.querySelector('#e-ten').value, p_thu_muc: d.querySelector('#e-thu').value, p_chu_de: d.querySelector('#e-cd').value, p_nhom: d.querySelector('#e-nhom').value, p_mo_ta: d.querySelector('#e-mt').value }).then(function () { d.remove(); toast('Đã lưu') }, function (e) { b.disabled = false; fail(e) }) }
  }
  function save(f, patch) { var o = { p_id: f.id, p_ten: f.ten, p_thu_muc: f.thu_muc, p_chu_de: f.chu_de || '', p_nhom: f.nhom || '', p_mo_ta: f.mo_ta || '', p_dang_dung: f.dang_dung, p_req: rid() }; for (var k in patch) o[k] = patch[k]; return rpc('u_file_update', o).then(function () { return loadFiles() }) }
  function del(f) {
    if (!confirm('Xóa hẳn "' + f.ten + '"? Không hoàn tác được. (Nếu chỉ muốn tạm không dùng, bấm Ẩn.)')) return;
    rpc('u_file_delete', { p_id: f.id, p_req: rid() }).then(function (r) { tok().then(function (t) { fetch(URL_ + '/storage/v1/object/lib-files/' + r.duong_dan, { method: 'DELETE', headers: { apikey: KEY, Authorization: 'Bearer ' + t } }) }); toast('Đã xóa'); return loadFiles() }, fail)
  }
  function bind(on) {
    if (IM.files === null && !IM.loading) loadFiles();
    var dz = document.getElementById('im-drop'); if (dz) { ['dragenter', 'dragover'].forEach(function (ev) { dz.addEventListener(ev, function (e) { e.preventDefault(); dz.style.borderColor = 'var(--teal)'; dz.style.background = 'var(--tint)' }) }); ['dragleave', 'drop'].forEach(function (ev) { dz.addEventListener(ev, function (e) { e.preventDefault(); dz.style.borderColor = 'var(--line)'; dz.style.background = 'var(--bg)' }) }); dz.addEventListener('drop', onDrop) }
    on('#im-file', 'onchange', function (x) { addFiles([].slice.call(x.files).map(function (f) { return { file: f, rel: '' } })) });
    on('#im-dir', 'onchange', function (x) { addFiles([].slice.call(x.files).map(function (f) { return { file: f, rel: f.webkitRelativePath || '' } })) });
    on('#im-up', 'onclick', upload); on('#im-clear', 'onclick', function () { IM.queue = []; render() });
    [].slice.call(document.querySelectorAll('[data-iq]')).forEach(function (el) { el.onchange = function () { var q = IM.queue.filter(function (x) { return x.id === el.dataset.iq })[0]; if (!q) return; var k = el.dataset.f; if (k === 'inc') { q.inc = el.checked; var b = document.getElementById('im-up'); if (b) { var n = IM.queue.filter(function (x) { return x.inc && !x.err }).length; b.disabled = !n; b.textContent = 'Tải lên ' + n + ' file' } } else q[k] = el.value } });
    on('#im-q', 'oninput', function (x) { IM.q = x.value; var p = x.selectionStart; render(); var n = document.getElementById('im-q'); if (n) { n.focus(); n.setSelectionRange(p, p) } }); on('#im-hid', 'onchange', function (x) { IM.hidden = x.checked; render() });
    [].slice.call(document.querySelectorAll('[data-imt]')).forEach(function (b) { b.onclick = function () { IM.thu = b.dataset.imt; render() } });
    [].slice.call(document.querySelectorAll('[data-imv]')).forEach(function (b) { b.onclick = function () { view(byId(b.dataset.imv)) } });
    [].slice.call(document.querySelectorAll('[data-ime]')).forEach(function (b) { b.onclick = function () { edit(byId(b.dataset.ime)) } });
    [].slice.call(document.querySelectorAll('[data-imh]')).forEach(function (b) { b.onclick = function () { var f = byId(b.dataset.imh); save(f, { p_dang_dung: f.dang_dung === 'CO' ? 'KHONG' : 'CO' }).then(function () { toast(f.dang_dung === 'CO' ? 'Đã ẩn' : 'Đã hiện') }, fail) } });
    [].slice.call(document.querySelectorAll('[data-imd]')).forEach(function (b) { b.onclick = function () { del(byId(b.dataset.imd)) } })
  }
  return { page: page, bind: bind, state: IM, detect: detect, addFiles: addFiles }
})();

;

/* ===== asm-admin.js ===== */
/* ASM – NHÂN SỰ & TÀI KHOẢN, THƯ VIỆN (mẫu tin, khuyến mãi), PHÁC ĐỒ (vaccine theo tuổi). Mở rộng ASM trong asm.js.
 * Thư viện = "NÓI GÌ" (câu chữ). Phác đồ = "KHI NÀO" (vaccine nào tới tuổi nào). Hai bên gặp nhau ở màn Gợi ý tin nhắn của nhân viên.
 * Mọi ghi qua hàm RPC của cơ sở dữ liệu (chỉ ASM gọi được). Dữ liệu mẫu và câu chữ là DỮ LIỆU GIẢ để thử. */
(function () {
  var X = ASM.X, U = X.u, head = U.head, sel = U.sel, ST = U.ST;
  IC.shield = '<path d="M12 3 4 6v6c0 5 3.5 8 8 9 4.5-1 8-4 8-9V6z"/><path d="m9 12 2 2 4-4"/>';
  var VAI = { QUAN_LY: 'Quản lý', NHAN_VIEN: 'Nhân viên' }, TT = { HOAT_DONG: 'Đang hoạt động', NGUNG: 'Đã khóa' };
  var AS = { tk: { shop: '', st: 'HOAT_DONG', q: '' }, tl: { tab: 'kho', nhom: '', st: 'ACTIVE', q: '' }, pd: { tv: '3', tu: '30', tg: '' } };
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
    h += window.ASMHR ? ASMHR.card() : '';
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
    var f = AS.tl, h = head('Thư viện', f.tab === 'kho' ? 'Nội dung nhân viên dùng khi nhắn và gọi. Nạp từ file Excel của công ty: xem trước, lưu chờ duyệt, duyệt thì nhân viên mới thấy. Phác đồ (khi nào tiêm gì) ở mục riêng, sẽ bổ sung sau.' : f.tab === 'anh' ? 'Ảnh, tờ rơi, voucher: kéo ảnh vào, app tự nhận diện để nhân viên dùng ở nút Gợi ý tin nhắn.' : f.tab === 'quet' ? 'Rà số tiền và CTKM trong các mẫu tin so với Bảng giá và CTKM hiện hành, thay tự động hoặc sửa tay.' : 'Mẫu soạn tay: bản tạm, đang dùng cho màn Gợi ý tin nhắn cho tới khi nối kho Excel ở Đợt 3.', f.tab === 'mau' ? '<button class="btn b1" id="tl-new">+ Thêm mẫu</button>' : f.tab === 'km' ? '<button class="btn b1" id="km-new">+ Thêm khuyến mãi</button>' : '');
    h += '<div class="chips"><button class="chip ' + (f.tab === 'kho' ? 'on' : '') + '" data-tltab="kho">Kho từ Excel</button><button class="chip ' + (f.tab === 'quet' ? 'on' : '') + '" data-tltab="quet">Quét giá &amp; CTKM</button><button class="chip ' + (f.tab === 'anh' ? 'on' : '') + '" data-tltab="anh">Ảnh, tờ rơi, voucher</button><button class="chip ' + (f.tab === 'mau' ? 'on' : '') + '" data-tltab="mau">Mẫu soạn tay (tạm)<em>' + D.tpl.length + '</em></button><button class="chip ' + (f.tab === 'km' ? 'on' : '') + '" data-tltab="km">Khuyến mãi soạn tay (tạm)<em>' + D.promos.length + '</em></button></div>';
    return h + (f.tab === 'kho' ? ASMLIB.page() : f.tab === 'quet' ? ASMSCAN.page() : f.tab === 'anh' ? ASMIMG.page() : f.tab === 'mau' ? tplList() : promoList())
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
    h += '<div class="ro"><b>Menu giữ chỗ.</b> Anh sẽ gửi dữ liệu phác đồ chuẩn, khi đó tôi làm nút nạp Excel cho khu này giống các khu trong Thư viện (có xem trước, chờ duyệt, bác sĩ duyệt).</div><div class="ro" style="background:var(--warn-bg);color:var(--warn)"><b>Dữ liệu mẫu minh họa.</b> Các độ tuổi bên dưới chưa phải phác đồ chuẩn đã duyệt. Bác sĩ cần nhập và duyệt phác đồ thật trước khi dùng cho khách thật; dòng "Chờ duyệt" không hiện cho nhân viên.</div>';
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
    if (AS.tl.tab === 'kho' && S.page === 'a_tl' && window.ASMLIB) ASMLIB.bind(on);
    if (AS.tl.tab === 'quet' && S.page === 'a_tl' && window.ASMSCAN) ASMSCAN.bind(on);
    if (AS.tl.tab === 'anh' && S.page === 'a_tl' && window.ASMIMG) ASMIMG.bind(on);
    if (S.page === 'a_tk' && window.ASMHR) ASMHR.bind(on);
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

;

/* ===== msg.js ===== */
/* GỢI Ý TIN NHẮN + LỜI NÓI KHI GỌI – dùng chung cho Nhân viên và Quản lý (Quản lý chỉ soạn cho lượt của chính mình).
 * Nội dung lấy từ THƯ VIỆN (msg_templates, promos) do ASM duyệt; "vaccine đã đến tuổi" lấy từ PHÁC ĐỒ (vaccine_rules) theo ngày sinh.
 * Ba nút gửi, mỗi nút làm đúng việc của nó rồi tự ghi 1 lần "Nhắn" (hàm u_msg_log):
 *   Zalo     = chép tin + mở zalo.me/<số> (Zalo không cho web điền sẵn tin hay tự kết bạn: nhân viên dán tin rồi gửi/kết bạn)
 *   SMS      = mở app tin nhắn tới số, nội dung điền sẵn (đồng thời chép tin phòng khi máy tính không có app SMS)
 *   Sao chép = chỉ chép tin
 * Dữ liệu và câu chữ mẫu hiện là DỮ LIỆU GIẢ để thử. */
var MSG = (function () {
  var CSS = '.imgrow,.imggrid{display:flex;gap:8px;overflow-x:auto;padding-bottom:4px}.imggrid{display:grid;grid-template-columns:repeat(3,1fr);overflow:visible}.ithumb{flex:none;position:relative;width:96px;border:2px solid var(--line);border-radius:12px;background:var(--card);padding:4px;font-size:11px;text-align:left;color:var(--ink)}.imggrid .ithumb{width:auto}.ithumb img{width:100%;height:64px;object-fit:cover;border-radius:8px;display:block}.ithumb.on{border-color:var(--teal);background:var(--tint)}.ithumb span{display:block;margin-top:3px;line-height:1.2;max-height:2.4em;overflow:hidden}.ick{position:absolute;top:6px;right:6px;background:var(--teal);color:#fff;border-radius:50%;width:20px;height:20px;display:grid;place-items:center;font-size:12px}' + '.person{display:grid;grid-template-columns:1fr auto;gap:10px;align-items:center;padding:12px 0;border-top:1px solid var(--line)}.person:first-child{border-top:0}.person .phone{font-size:20px}.btn.bo{border:1.5px solid var(--teal);color:var(--teal-d);background:var(--card);min-height:44px}' +
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
    S = { t: t, ci: Math.min(ci || 0, links.length - 1), step: step, tplId: tp ? tp.ma : '', vxOn: true, vx: vx, km: {}, ngc: false, dc: false, xh: '', view: 'main', pg: 'KH_MOI', pq: '', img: {}, iq: '', it: '', text: '', base: '', sending: false };
    host = document.createElement('div'); host.className = 'modal mm'; host.onclick = function (e) { if (e.target === host) close() }; document.body.appendChild(host);
    loadFiles();
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
  function draw() { S.view === 'pick' ? drawPick() : S.view === 'img' ? drawImgs() : drawMain() }
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
    h += imgBlock(due);
    h += '<div><h3>Mẫu khác cùng bước</h3><div style="display:grid;gap:8px">' + alts.map(function (a) { return '<button class="alt" data-t="' + esc(a.ma) + '"><b>' + esc(a.tieu_de) + '</b>' + esc(fill(a.noi_dung, cx).slice(0, 90)) + '…</button>' }).join('') + '<button class="btn b3" id="more">Xem thêm mẫu (' + nTpl + ')</button></div></div>';
    h += '<div class="row"><span class="meta">Xưng hô:</span><select id="xh">' + XH.map(function (o) { return '<option value="' + esc(o[0]) + '"' + (o[0] === S.xh ? ' selected' : '') + '>' + esc(o[1]) + '</option>' }).join('') + '</select></div>';
    h += '</div><div class="msf">' + (Object.keys(S.img).length ? '<button class="btn bo" id="sendimg" style="width:100%">Gửi kèm ' + Object.keys(S.img).length + ' ảnh (chia sẻ sang Zalo/ứng dụng khác)</button>' : '') + '<div class="meta" style="text-align:center">Chạm nút nào thì làm đúng việc của nút đó, rồi app tự ghi 1 lần "Nhắn" (' + (msgCount() + 1) + '/2).</div><div class="sendrow"><button class="btn b1" data-go="ZALO"><span>Zalo</span><small>chép tin + mở Zalo</small></button><button class="btn b2" data-go="SMS"><span>SMS</span><small>tới ' + esc(tel || 'chưa có số') + '</small></button><button class="btn b3" data-go="SAO_CHEP"><span>Sao chép</span><small>chỉ chép tin</small></button></div></div></div>';
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
    all('[data-img]').forEach(function (b) { b.onclick = function () { var k = b.dataset.img; S.img[k] ? delete S.img[k] : S.img[k] = 1; draw() } });
    var im = q('#imore'); if (im) im.onclick = function () { S.view = 'img'; draw() }; var si = q('#sendimg'); if (si) si.onclick = sendImages;
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

  /* ---------- ảnh đính kèm: gợi ý theo khách và vaccine, nhân viên chọn rồi chia sẻ ---------- */
  function loadFiles() { if (D.files !== undefined || D.filesLoading) return; D.filesLoading = true; rest('lib_files?select=id,ten,thu_muc,chu_de,nhom,mime,duong_dan,anh_nho&dang_dung=eq.CO&order=thoi_gian.desc&limit=300').then(function (a) { D.files = a; D.filesLoading = false; if (S && host) draw() }, function () { D.files = []; D.filesLoading = false }) }
  function fk(x) { return window.LIBSCAN ? LIBSCAN.fold(x || '') : String(x || '').toLowerCase() }
  function group() { var c = cust(), d = custDays(c); if (d == null) return ''; return d < 9 * 365 ? 'NHI' : d < 18 * 365 ? 'TRE_LON' : 'NGUOI_LON' }
  function imgScore(f, names, grp) {
    var sc = 0, fc = fk(f.chu_de); if (fc && names.some(function (n) { return n.indexOf(fc) >= 0 || fc.indexOf(n) >= 0 || n.split(/\s+/).some(function (t) { return t.length >= 3 && fc.split(/\s+/).indexOf(t) >= 0 }) })) sc += 3;
    if (f.nhom === grp) sc += 2; else if (!f.nhom || f.nhom === 'TAT_CA') sc += 1; if (f.thu_muc === 'CTKM_VOUCHER' && Object.keys(S.km).some(function (k) { return S.km[k] })) sc += 2; return sc
  }
  function imgBlock(due) {
    var F = D.files; if (!F || !F.length) return ''; var names = due.filter(function (r) { return S.vx[r.ma] }).map(function (r) { return fk(r.ten) }), tp = tplById(S.tplId); if (tp && tp.tieu_de) names.push(fk(tp.tieu_de));
    var grp = group(), L = F.map(function (f) { return { f: f, s: imgScore(f, names, grp) } }).filter(function (x) { return x.s > 0 }).sort(function (a, b) { return b.s - a.s }).slice(0, 8);
    return '<div><h3>Ảnh đính kèm (gợi ý theo khách và vaccine)</h3>' + (L.length ? '<div class="imgrow">' + L.map(function (x) { return thumbBtn(x.f) }).join('') + '</div>' : '<div class="meta">Chưa có ảnh nào hợp. Bấm "Xem tất cả ảnh".</div>') + '<button class="btn b3" id="imore" style="width:100%;margin-top:8px">Xem tất cả ảnh (' + F.length + ')</button></div>'
  }
  function thumbBtn(f) { return '<button class="ithumb ' + (S.img[f.id] ? 'on' : '') + '" data-img="' + esc(f.id) + '">' + (f.anh_nho ? '<img src="' + f.anh_nho + '" alt="">' : '<div style="height:64px;display:grid;place-items:center;color:var(--soft)">PDF</div>') + '<span>' + esc(f.ten.slice(0, 24)) + '</span>' + (S.img[f.id] ? '<b class="ick">✓</b>' : '') + '</button>' }
  function drawImgs() {
    var F = D.files || [], q = fk(S.iq), L = F.filter(function (f) { return (!S.it || f.thu_muc === S.it) && (!q || fk(f.ten + ' ' + f.chu_de).indexOf(q) >= 0) }), T = { '': 'Tất cả', VACCINE_DOI_TUONG: 'Vaccine', CAM_NANG_BAU: 'Cẩm nang bầu', CTKM_VOUCHER: 'Khuyến mãi' };
    host.innerHTML = '<div class="msheet"><div class="msh"><button class="x" id="pb" aria-label="Quay lại" style="margin-left:0">' + ic('back') + '</button><div class="name" style="font-size:18px">Chọn ảnh (' + Object.keys(S.img).length + ')</div><button class="x" id="mx" aria-label="Đóng">' + ic('x') + '</button></div><div class="msb"><div class="chips">' + Object.keys(T).map(function (k) { return '<button class="chip ' + (S.it === k ? 'on' : '') + '" data-it="' + k + '">' + T[k] + '</button>' }).join('') + '</div><div class="search" style="max-width:none">' + ic('search') + '<input id="iqi" placeholder="Tìm ảnh…" value="' + esc(S.iq) + '"></div><div class="imggrid">' + (L.length ? L.map(thumbBtn).join('') : '<div class="card empty">Không có ảnh phù hợp.</div>') + '</div></div></div>';
    host.querySelector('#mx').onclick = close; host.querySelector('#pb').onclick = function () { S.view = 'main'; draw() };
    [].slice.call(host.querySelectorAll('[data-it]')).forEach(function (b) { b.onclick = function () { S.it = b.dataset.it; draw() } });
    var qi = host.querySelector('#iqi'); qi.oninput = function () { S.iq = qi.value; var p = qi.selectionStart; draw(); var n = host.querySelector('#iqi'); n.focus(); n.setSelectionRange(p, p) };
    [].slice.call(host.querySelectorAll('[data-img]')).forEach(function (b) { b.onclick = function () { var k = b.dataset.img; S.img[k] ? delete S.img[k] : S.img[k] = 1; draw() } })
  }
  function sendImages() {
    var ids = Object.keys(S.img), F = (D.files || []).filter(function (f) { return S.img[f.id] }), text = S.text; if (!F.length) return; toast('Đang chuẩn bị ' + F.length + ' ảnh…');
    Promise.all(F.map(function (f) { return fileBlob(f).then(function (b) { var ext = f.mime === 'image/png' ? 'png' : f.mime === 'application/pdf' ? 'pdf' : f.mime === 'image/webp' ? 'webp' : 'jpg'; return new File([b], (f.ten.replace(/[^\w\- ]+/g, '').slice(0, 40) || 'anh') + '.' + ext, { type: f.mime }) }) })).then(function (files) {
      if (navigator.canShare && navigator.canShare({ files: files })) return navigator.share({ files: files, text: text }).then(function () { toast('Đã mở chia sẻ ảnh. Chọn Zalo để gửi.') }, function (e) { if (e && e.name !== 'AbortError') toast('Không chia sẻ được, đã tải ảnh về máy.', 0, 1) });
      files.forEach(function (f) { var a = document.createElement('a'); a.href = URL.createObjectURL(f); a.download = f.name; document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove() }, 800) }); toast('Máy này không chia sẻ trực tiếp được, đã tải ' + files.length + ' ảnh về để anh/chị đính kèm.')
    }, function (e) { fail(e) })
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

;
