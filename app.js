/* LC TCQG HUB: gộp từ các module pager.js, lib-parse.js, asm.js, asm-lib.js, scan-engine.js, asm-scan.js, asm-hr.js, asm-img.js, asm-admin.js, msg.js, icons.js (build 20261126s) */

/* ===== pager.js ===== */
/* PHÂN TRANG CHUNG: mọi bảng và danh sách thẻ trong màn chính đều tự có "số dòng mỗi trang" (10 / 15 / 20) và nút trang trước/sau.
 * Chạy sau mỗi lần vẽ màn hình (PGR.run). Không đổi dữ liệu: chỉ ẩn bớt dòng ngoài trang hiện tại. Bảng có thuộc tính data-nopg tự phân trang riêng (Kho Data).
 * Số dòng chọn được nhớ trên máy (localStorage). Trang hiện tại nhớ theo từng bảng; khi số dòng đổi (lọc, tìm) thì về trang 1. */
var PGR = (function () {
  var KEY = 'lc_pgsz', SIZES = [10, 15, 20, 30, 50, 100], st = {};
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
  /* SDT_NV = SĐT nhân viên (tự điền từ tài khoản); TEN_VACCINE = vaccine nhân viên đã chọn (tự điền); 5 biến còn lại là chỗ nhân viên điền tay trước khi gửi. */
  var VARS = ['XUNG_HO', 'XUNG_HO_HOA', 'TEN_NGUOI_LH', 'TEN_BE', 'TUOI', 'TEN_NV', 'TEN_CO_SO', 'DIA_CHI', 'GIO_MO_CUA', 'HOTLINE', 'SDT_NV', 'TEN_VACCINE', 'NGAY_TIEM', 'SO_MUI', 'NGAY_HEN', 'LINK_GOOGLE_MAPS', 'NOI_DUNG_PHAN_HOI'];
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
  /* PHÁC ĐỒ & TỪ ĐIỂN VACCINE (Hải 24/11): mỗi dòng 1 vaccine (tên thương mại). Theo mẫu file KH-01 "UP_PHAC_DO" (ten_vaccine, phong_benh, tuoi_bat_dau_ngay, tuoi_ket_thuc_ngay, so_mui, khoang_cach_ngay, noi_dung_tu_van, luu_y, ghi_chu)
   * và thêm cột cho từ điển tra cứu (nước SX, độ tuổi chỉ định, chống chỉ định, phản ứng sau tiêm, nhóm vaccine chuẩn).
   * File cũ ghi tuổi + nước SX chung trong "Lưu ý" ("6 tuần - 24 tháng · SX: Pháp"): app tự tách ra; tuổi kết thúc để trống thì tự suy từ chữ độ tuổi (có ghi chú để bác sĩ xác nhận). */
  function ageEnd(t) {
    var s = String(t || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd'), re = /(?:[-–]|<|duoi)\s*(?:tron\s*)?(?:<\s*)?(\d+)\s*(tuan|thang|tuoi)/g, m, last = null;
    while ((m = re.exec(s))) last = m; if (!last) return null;
    var n = Number(last[1]), u = last[2]; return u === 'tuan' ? n * 7 : u === 'thang' ? Math.round(n * 30.4375) : Math.round(n * 365.25)
  }
  CONFIG.PHAC_DO = { ten: 'Phác đồ & từ điển vaccine', moTa: 'Mỗi dòng một vaccine (tên thương mại): nhóm bệnh, tuổi, phác đồ, nước sản xuất, chống chỉ định, phản ứng sau tiêm. App dùng để gợi ý "vaccine đã đến tuổi", làm nút Từ điển vaccine cho nhân viên và nối đúng tư vấn, nguy cơ, khuyến mãi theo vaccine.', cheDo: 'THAY_THE',
    hau: function (o) {
      /* File từ NotebookLM: tên đánh số "12. Bexsero" → bỏ số thứ tự; ký tự Hán lọt vào chữ ("Viêm脑") → "não" */
      Object.keys(o).forEach(function (k) { if (typeof o[k] === 'string') o[k] = o[k].replace(/\s*脑/g, " não") });
      if (o.ten_vaccine) o.ten_vaccine = String(o.ten_vaccine).replace(/^\s*\d+\s*[.)]\s+/, '').trim();
      var lu = String(o.luu_y || ''), m = /\s*[·•]\s*SX:\s*([^·\n]+?)\s*$/i.exec(lu);
      if (m) { if (!o.nuoc_sx) o.nuoc_sx = m[1].trim(); lu = lu.slice(0, m.index).trim() }
      lu = lu.replace(/^\*\s*/, '').trim(); if (lu && !o.do_tuoi) o.do_tuoi = lu;
      if (o.tuoi_ket_thuc_ngay === undefined || o.tuoi_ket_thuc_ngay === null || o.tuoi_ket_thuc_ngay === '') { var e = ageEnd(o.do_tuoi || lu); if (e != null && e > (Number(o.tuoi_bat_dau_ngay) || 0)) { o.tuoi_ket_thuc_ngay = e; o.ghi_chu = (o.ghi_chu ? o.ghi_chu + ' · ' : '') + 'Tuổi kết thúc tự suy từ độ tuổi chỉ định, cần bác sĩ xác nhận' } }
    },
    cot: [
      { k: 'ten_vaccine', nhan: 'Tên vaccine', alias: ['vaccine', 'ten thuong mai', 'ten vac xin', 'stt ten vaccine', 'stt ten thuong mai'], bb: 1 }, { k: 'phong_benh', nhan: 'Phòng bệnh', alias: ['nhom benh', 'benh phong ngua', 'benh'], bb: 1 },
      { k: 'nhom_vaccine', nhan: 'Nhóm vaccine chuẩn', alias: ['nhom vaccine', 'nhom chuan'] },
      { k: 'tuoi_bat_dau_ngay', nhan: 'Tuổi bắt đầu (ngày)', alias: ['tuoi bat dau', 'tuoi toi thieu ngay', 'tu ngay tuoi'], bb: 1, kieu: 'so' }, { k: 'tuoi_ket_thuc_ngay', nhan: 'Tuổi kết thúc (ngày)', alias: ['tuoi ket thuc', 'tuoi toi da ngay'], kieu: 'so' },
      { k: 'so_mui', nhan: 'Số mũi', alias: ['so mui theo phac do'] }, { k: 'khoang_cach_ngay', nhan: 'Khoảng cách giữa các mũi (ngày)', alias: ['khoang cach', 'khoang cach giua cac mui', 'lich tiem'] },
      { k: 'noi_dung_tu_van', nhan: 'Nội dung tư vấn / phác đồ', alias: ['noi dung tu van', 'phac do', 'phac do tham khao', 'tu van'], bien: 1 },
      { k: 'luu_y', nhan: 'Lưu ý', alias: ['luu y chuyen mon'] }, { k: 'nuoc_sx', nhan: 'Nước sản xuất', alias: ['xuat xu', 'nuoc san xuat', 'nsx'] }, { k: 'do_tuoi', nhan: 'Độ tuổi chỉ định', alias: ['do tuoi', 'chi dinh', 'do tuoi chi dinh', 'doi tuong'] },
      { k: 'chong_chi_dinh', nhan: 'Chống chỉ định / thận trọng', alias: ['chong chi dinh', 'than trong'] }, { k: 'phan_ung', nhan: 'Phản ứng sau tiêm thường gặp', alias: ['phan ung sau tiem', 'tac dung phu', 'phan ung'] }, { k: 'ghi_chu', nhan: 'Ghi chú', alias: ['ghi chu nguon nguon chua neu can y kien bs', 'ghi chu nguon', 'nguon'] }] };
  /* HỎI ĐÁP & XỬ LÝ TỪ CHỐI (Hải 25/11): mỗi dòng một tình huống khách hỏi hoặc từ chối + cách trả lời. "Nhóm vaccine" ghi tên vaccine, nhóm (6 trong 1, Phế cầu…) hoặc "Chung" (áp dụng mọi vaccine). Hiện trong Từ điển vaccine của nhân viên. */
  CONFIG.HOI_DAP = { ten: 'Hỏi đáp & xử lý từ chối', moTa: 'Mỗi dòng một tình huống: khách hỏi hoặc từ chối, kèm cách trả lời. Cột Nhóm vaccine ghi tên vaccine hoặc nhóm (6 trong 1, Phế cầu…), ghi "Chung" nếu áp dụng mọi vaccine. Nhân viên xem trong Từ điển vaccine.', cheDo: 'THAY_THE',
    cot: [
      { k: 'nhom_vaccine', nhan: 'Nhóm vaccine / Vaccine', alias: ['nhom vaccine', 'vaccine', 'ten vaccine', 'nhom vaccine chuan', 'nhom'] }, { k: 'loai', nhan: 'Loại (Hỏi đáp / Từ chối)', alias: ['loai', 'phan loai', 'loai tinh huong'] },
      { k: 'tinh_huong', nhan: 'Khách hỏi / khách nói', alias: ['cau hoi', 'khach hoi', 'khach noi', 'tinh huong', 'ly do tu choi', 'van de'], bb: 1 }, { k: 'tra_loi', nhan: 'Cách trả lời', alias: ['cau tra loi', 'tra loi', 'cach xu ly', 'xu ly', 'goi y tra loi', 'kich ban'], bb: 1, bien: 1 },
      { k: 'luu_y', nhan: 'Lưu ý', alias: ['luu y chuyen mon'] }, { k: 'nguon', nhan: 'Nguồn', alias: ['nguon tai lieu', 'tai lieu'] }, { k: 'ghi_chu', nhan: 'Ghi chú' }] };
  var ORDER = ['BANG_GIA', 'CTKM', 'MAU_TUOI', 'MAU_CHU_DE', 'MAU_SAU_TIEM', 'MAU_CHUNG', 'NGUY_CO', 'PHAC_DO', 'HOI_DAP'];

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
      /* file từ NotebookLM đôi khi lọt ký tự Hán "脑" (não) vào chữ Việt: sửa cho mọi khu */
      Object.keys(o).forEach(function (k) { if (typeof o[k] === 'string' && o[k].indexOf('脑') >= 0) o[k] = o[k].replace(/\s*脑/g, ' não') });
      if (cfg.hau) cfg.hau(o);
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
  var NAV = [['a_tq', 'Tổng quan', 'home'], ['a_kho', 'Kho Data', 'list'], ['a_pc', 'Phân công', 'swap'], ['a_nap', 'Nạp Data', 'upload'], ['a_bc', 'Báo cáo', 'chart'], ['a_tk', 'Nhân sự & tài khoản', 'team'], ['a_tl', 'Thư viện', 'book'], ['a_pd', 'Phác đồ', 'shield'], ['a_nv', 'Kiểm tra gọi & nhắn', 'shield'], ['a_cf', 'Cấu hình ghi nhận', 'shield'], ['acc', 'Của tôi', 'user']];
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
  function qlOf(id) { var q = staffOf(id).filter(function (a) { return a.vai_tro === 'QUAN_LY' }).map(function (a) { return a.ho_ten }); return q.length ? q.join(', ') : '–' }
  function turnsByCust() { var m = {}; D.turns.forEach(function (t) { (m[t.ma_khach] = m[t.ma_khach] || []).push(t) }); return m }
  function custState(ts) { if (!ts || !ts.length) return 'kho'; return ts.some(open_) ? 'dang' : 'xong' }
  function yes(t) { return tdt(t) }
  function pct(a, b) { return b ? Math.round(a / b * 100) : 0 }
  function sel(id, opts, val) { return '<select id="' + id + '">' + opts.map(function (o) { return '<option value="' + esc(o[0]) + '"' + (String(o[0]) === String(val) ? ' selected' : '') + '>' + esc(o[1]) + '</option>' }).join('') + '</select>' }

  /* ---------- khung ---------- */
  function shell() {
    function b(x) { return '<button class="nav ' + (S.page === x[0] ? 'on' : '') + '" data-p="' + x[0] + '">' + ic(x[2]) + x[1] + '</button>' }
    $('#side').innerHTML = '<div class="brand">' + LOGO + '</div><div class="grp">Điều phối</div>' + NAV.slice(0, 3).map(b).join('') + '<div class="grp">Dữ liệu</div>' + NAV.slice(3, 5).map(b).join('') + '<div class="grp">Nhân sự &amp; nội dung</div>' + NAV.slice(5, 8).map(b).join('') + '<div class="grp">Hệ thống</div>' + NAV.slice(8).map(b).join('') +
      '<div class="grp">Công cụ</div><button class="nav" id="lkside">' + ic('book') + 'Từ điển vaccine</button><div class="me"><div class="ava">' + esc(ini(ME.ho_ten)) + '</div><div><b>' + esc(ME.ho_ten) + '</b><small>ASM · Giám đốc kinh doanh KV</small></div></div><div class="vx" style="padding:8px 4px 0;font-size:11px">' + BRAND.line() + '</div>';
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
  /* Tổng quan (26/11/2026): 3 tầng. Tầng 1 chỉ số có thứ bậc, tầng 2 việc cần xử lý ngay + xu hướng 14 ngày, tầng 3 so sánh Shop có tìm kiếm, lọc, sắp xếp. */
  function dkey(d) { return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2) }
  function series(shopId, n) {
    var days = [], i, ix = {}, d0 = today0();
    for (i = n - 1; i >= 0; i--) { var d = new Date(d0.getTime() - i * 864e5); ix[dkey(d)] = days.length; days.push({ d: d, giao: 0, xong: 0 }) }
    D.turns.forEach(function (t) {
      if (shopId && t.id_shop !== shopId) return;
      if (t.d_giao && !isNaN(t.d_giao)) { var k = dkey(t.d_giao); if (ix[k] != null) days[ix[k]].giao++ }
      if (t.trang_thai === 'HOAN_THANH' && t.d_xong && !isNaN(t.d_xong)) { var k2 = dkey(t.d_xong); if (ix[k2] != null) days[ix[k2]].xong++ }
    });
    return days
  }
  /* Biểu đồ vẽ theo bề rộng thật của khung (W = số điểm ảnh) để chữ trục luôn đủ lớn: điện thoại 7 ngày, máy tính 14 ngày */
  /* Ghim cụm đầu trang (tiêu đề + ô tìm + bộ lọc + chip) ở mọi menu: gom các phần tử đầu trang vào một khung dính */
  function stick(m) {
    var kids = [].slice.call(m.children), grab = [], i;
    for (i = 0; i < kids.length; i++) { var k = kids[i]; if ((i === 0 && k.classList.contains('row2')) || (i > 0 && k.matches && k.matches('.kf,.chips,.top,.tbar2,.seg2'))) grab.push(k); else break }
    if (!grab.length) return; var w = document.createElement('div'); w.className = 'stk'; m.insertBefore(w, grab[0]); grab.forEach(function (k) { w.appendChild(k) });
    /* dòng mô tả trang cuộn đi bình thường, chỉ tiêu đề + công cụ được ghim */
    var pp = w.querySelector('.greet p'); if (pp) { var d = document.createElement('div'); d.className = 'pgd'; d.innerHTML = pp.innerHTML; pp.parentNode.removeChild(pp); w.parentNode.insertBefore(d, w.nextSibling) }
  }
  var LASTDAYS = null;
  function fitChart() {
    var w = $('#chw'); if (!w || !LASTDAYS) return; var cw = Math.floor(w.clientWidth);
    if (cw >= 200 && Math.abs(cw - Number(w.getAttribute('data-w') || 0)) > 6) { w.setAttribute('data-w', cw); w.innerHTML = chartSvg(LASTDAYS, Math.min(cw, 760)) }
  }
  function chartSvg(days, W) {
    W = W || 560; var H = W < 400 ? 150 : 170, L = 30, B = 24, T = 10, R = 6, n = days.length, mx = 1; days.forEach(function (x) { mx = Math.max(mx, x.giao, x.xong) });
    var step = Math.max(1, Math.ceil(mx / 4)); mx = step * 4; var cw = (W - L - R) / n, bw = Math.min(14, cw * .36), g = '', b = '', lb = '', i, lstep = cw >= 36 ? 1 : cw >= 17 ? 2 : 3;
    for (i = 0; i <= 4; i++) { var y = T + (H - T - B) * (1 - i / 4); g += '<line x1="' + L + '" x2="' + (W - R) + '" y1="' + y.toFixed(1) + '" y2="' + y.toFixed(1) + '" class="gl"/><text x="' + (L - 6) + '" y="' + (y + 3.5).toFixed(1) + '" class="gt" text-anchor="end">' + (step * i) + '</text>' }
    days.forEach(function (x, j) {
      var cx = L + cw * j + cw / 2, hg = (H - T - B) * x.giao / mx, hx = (H - T - B) * x.xong / mx, lab = ('0' + x.d.getDate()).slice(-2) + '/' + ('0' + (x.d.getMonth() + 1)).slice(-2);
      b += '<rect x="' + (cx - bw - 1).toFixed(1) + '" y="' + (H - B - hg).toFixed(1) + '" width="' + bw.toFixed(1) + '" height="' + Math.max(0, hg).toFixed(1) + '" rx="3" class="bg"><title>' + lab + ': giao ' + x.giao + '</title></rect><rect x="' + (cx + 1).toFixed(1) + '" y="' + (H - B - hx).toFixed(1) + '" width="' + bw.toFixed(1) + '" height="' + Math.max(0, hx).toFixed(1) + '" rx="3" class="bx"><title>' + lab + ': hoàn thành ' + x.xong + '</title></rect>';
      if ((n - 1 - j) % lstep === 0) lb += '<text x="' + cx.toFixed(1) + '" y="' + (H - 6) + '" class="gt" text-anchor="middle">' + lab + '</text>'
    });
    return '<svg viewBox="0 0 ' + W + ' ' + H + '" class="chart" role="img" aria-label="Biểu đồ giao và hoàn thành ' + n + ' ngày gần nhất">' + g + b + lb + '</svg>'
  }
  function shopRows() {
    var SH = shops(), rows = SH.map(function (s) { var x = statsFor(s.id_shop); return { s: s, x: x, ptr: x.luot ? Math.round(x.xong / x.luot * 100) : -1, ql: qlOf(s.id_shop), nv: staffOf(s.id_shop).length } });
    var live = rows.filter(function (r) { return r.x.luot > 0 }), best = null, worst = null;
    if (live.length >= 3) { live.forEach(function (r) { if (!best || r.ptr > best.ptr) best = r; if (!worst || r.ptr < worst.ptr) worst = r }); if (best === worst) best = worst = null }
    rows.forEach(function (r) { r.tag = r === best ? 'Dẫn đầu' : r === worst ? 'Cần hỗ trợ' : '' });
    var q = (ST.sq || '').trim().toLowerCase(); if (q) rows = rows.filter(function (r) { return (r.s.ten_shop + ' ' + r.s.ma_shop + ' ' + r.ql).toLowerCase().indexOf(q) >= 0 });
    if (ST.late) rows = rows.filter(function (r) { return r.x.qh > 0 });
    var k = ST.ssort || 'qh', dir = ST.sdir || -1, val = function (r) { return k === 'ten' ? r.s.ten_shop : k === 'luot' ? r.x.luot : k === 'mo' ? r.x.mo : k === 'ptr' ? r.ptr : k === 'dy' ? r.x.dy : r.x.qh };
    rows.sort(function (a, b) { var A = val(a), B = val(b), c = typeof A === 'string' ? A.localeCompare(B) : A - B; return (c * dir) || a.s.ma_shop.localeCompare(b.s.ma_shop) });
    return rows
  }
  function sortTh(key, label, num) { var on = (ST.ssort || 'qh') === key, dir = ST.sdir || -1; return '<th class="' + (num ? 'num ' : '') + (on ? 'on' : '') + '"><button class="srt" data-sort="' + key + '" aria-label="Sắp xếp theo ' + esc(label) + '">' + esc(label) + '<i>' + (on ? (dir < 0 ? '▼' : '▲') : '') + '</i></button></th>' }
  /* Danh sách đằng sau mỗi ô số liệu ở Tổng quan: hiện tại chỗ (không nhảy màn hình), có nút qua lại */
  var DT_N = 5, DT_TT = { mo: 'Đang xử lý', dy: 'Hẹn / tiêm', qh: 'Quá hạn gọi', chia: 'Lượt chờ chia', kho: 'Khách chưa giao' };
  function dtItems(k) {
    var T = D.turns.filter(function (t) { return !ST.shop || t.id_shop === ST.shop });
    function who(t) { return t.recipient_type === 'SHOP' ? 'Cả Shop (chưa chia)' : nm(t.ma_tk_nhan) }
    function cus(t) { return (D.cus[t.ma_khach] || {}).ho_ten || t.ma_khach }
    if (k === 'kho') { var tc = turnsByCust(); return Object.keys(D.cus).filter(function (m) { return custState(tc[m]) === 'kho' }).sort().map(function (m) { var c = D.cus[m]; return { a: c.ho_ten, b: [ageTxt(c), c.ma_tcqg || c.ma_khach].filter(Boolean).join(' · ') } }) }
    var L = k === 'mo' ? T.filter(open_) : k === 'dy' ? T.filter(yes) : k === 'qh' ? T.filter(function (t) { return overdueDays(t) > 0 }).sort(function (a, b) { return overdueDays(b) - overdueDays(a) }) : T.filter(function (t) { return t.recipient_type === 'SHOP' && t.trang_thai === 'DA_GIAO' });
    return L.map(function (t) { var d = overdueDays(t); return { a: cus(t), b: [shopName(t.id_shop), who(t), d ? 'quá hạn ' + d + ' ngày' : ''].filter(Boolean).join(' · ') } })
  }
  function dtOn(k, src) { return ST.dt === k && ST.dtSrc === src ? ' on' : '' }
  function dtPanel(src) {
    var k = ST.dt; if (!k || ST.dtSrc !== src) return '';
    var L = dtItems(k), pages = Math.max(1, Math.ceil(L.length / DT_N)); ST.dtPg = Math.min(Math.max(1, ST.dtPg || 1), pages);
    var rows = L.slice((ST.dtPg - 1) * DT_N, ST.dtPg * DT_N), go = k === 'kho' ? 'a_kho' : k === 'chia' ? 'a_pc' : 'late';
    return '<div class="dtp"><div class="dth"><b>' + DT_TT[k] + '</b><span class="meta">' + L.length + ' mục</span><span class="sp"></span><button class="ibtn dtx" data-dtx="1" aria-label="Đóng danh sách">' + ic('x') + '</button></div>' +
      (rows.length ? rows.map(function (r) { return '<div class="dti"><b>' + esc(r.a) + '</b><small>' + esc(r.b) + '</small></div>' }).join('') : '<div class="meta" style="padding:8px 2px">Không có mục nào.</div>') +
      '<div class="dtf"><button class="btn b3 dtn" data-dtp="-1"' + (ST.dtPg <= 1 ? ' disabled' : '') + ' aria-label="Trang trước">' + ic('back') + '</button><span class="meta">' + ST.dtPg + ' / ' + pages + '</span><button class="btn b3 dtn" data-dtp="1"' + (ST.dtPg >= pages ? ' disabled' : '') + ' aria-label="Trang sau">' + ic('chev') + '</button><span class="sp"></span><button class="btn b3" data-dtgo="' + go + '">' + (k === 'kho' ? 'Mở Kho Data' : k === 'chia' ? 'Mở Phân công' : 'Xem theo Shop') + '</button></div></div>'
  }
  function pgTq() {
    var tc = turnsByCust(), kho = 0, tot = Object.keys(D.cus).length; Object.keys(D.cus).forEach(function (k) { if (custState(tc[k]) === 'kho') kho++ });
    var all = statsFor(ST.shop), first = ME.ho_ten.trim().split(/\s+/).pop(), SH = shops(), ptr = pct(all.xong, all.luot);
    var h = head('Chào anh ' + esc(first), 'Tổng hợp ' + (ST.shop ? esc(shopName(ST.shop)) : 'toàn khu vực (' + SH.length + ' Shop)') + ' · ' + new Date().toLocaleDateString('vi-VN', { weekday: 'long', day: 'numeric', month: 'long' }),
      '<div style="display:flex;gap:8px;align-items:center"><span class="meta">Shop:</span>' + sel('fshop', [['', 'Tất cả Shop']].concat(SH.map(function (s) { return [s.id_shop, s.ten_shop] })), ST.shop) + '</div>');
    /* Tầng 1: tổng quan, 1 chỉ số chính + 3 chỉ số quan trọng + 2 chỉ số phụ */
    h += '<section class="tier tov"><h2 class="tl">Tổng quan</h2><div class="kp"><div class="card kp-hero"><div class="l">Tiến độ xử lý</div><div class="big">' + ptr + '<small>%</small></div><div class="pb" role="progressbar" aria-valuenow="' + ptr + '" aria-valuemin="0" aria-valuemax="100"><i style="width:' + ptr + '%"></i></div><div class="meta">' + all.xong + ' / ' + all.luot + ' lượt đã hoàn thành</div></div>' +
      '<button class="card kp-m' + dtOn('mo', 'kp') + '" data-dt="mo" data-src="kp"><span class="l">Đang xử lý</span><span class="v">' + all.mo + '</span></button><button class="card kp-m' + dtOn('dy', 'kp') + '" data-dt="dy" data-src="kp"><span class="l">Hẹn / tiêm</span><span class="v ok">' + all.dy + '</span></button><button class="card kp-m' + (all.qh ? ' bad' : '') + dtOn('qh', 'kp') + '" data-dt="qh" data-src="kp"><span class="l">Quá hạn gọi</span><span class="v">' + all.qh + '</span></button></div>' + dtPanel('kp') +
      '<div class="kp2"><button data-go="a_kho"><b>' + tot + '</b> trong kho</button><button data-go="a_kho"><b>' + kho + '</b> chưa giao</button><button data-go="a_pc"><b>' + all.chia + '</b> chờ chia</button></div></section>';
    /* Tầng 2: điều hành */
    var sv = '', tl = [];
    if (all.qh) tl.push('<button class="sv sv1' + dtOn('qh', 'al') + '" data-dt="qh" data-src="al" title="Tính từ Ngày cần liên hệ của từng lượt"><span class="sl">Khẩn</span><span class="sn">' + all.qh + '</span><span class="sb">quá hạn gọi</span></button>');
    if (all.chia) tl.push('<button class="sv sv2' + dtOn('chia', 'al') + '" data-dt="chia" data-src="al" title="Quản lý chia khi bạn bật quyền cho lần giao đó"><span class="sl">Cần làm</span><span class="sn">' + all.chia + '</span><span class="sb">lượt chờ chia</span></button>');
    if (kho) tl.push('<button class="sv sv3' + dtOn('kho', 'al') + '" data-dt="kho" data-src="al" title="Khách đang nằm trong kho"><span class="sl">Nên làm</span><span class="sn">' + kho + '</span><span class="sb">chưa giao</span></button>');
    if (tl.length) sv = '<div class="svg3" style="--n:' + tl.length + '">' + tl.join('') + '</div>' + dtPanel('al');
    var days = series(ST.shop, 14), w1 = 0, w0 = 0; days.forEach(function (x, i) { if (i >= 7) w1 += x.xong; else w0 += x.xong }); var dl = w1 - w0;
    var narrow = window.innerWidth < 960, shown = narrow ? days.slice(-7) : days; LASTDAYS = shown;
    h += '<section class="tier tops"><h2 class="tl">Điều hành</h2><div class="t2"><div class="card cal"><h2>Cần xử lý ngay</h2>' + (sv || '<div class="alert"><i>✓</i><div><b>Chưa có gì cần xử lý</b></div></div>') + '</div>' +
      '<div class="card cch"><div class="row2"><h2 style="margin:0">Tiến độ ' + shown.length + ' ngày gần nhất</h2><span class="trend ' + (dl > 0 ? 'up' : dl < 0 ? 'down' : '') + '">' + (dl > 0 ? '▲ +' + dl : dl < 0 ? '▼ ' + dl : '–') + ' so với 7 ngày trước</span></div><div class="chw" id="chw" data-w="0">' + chartSvg(shown, narrow ? 250 : 560) + '</div><div class="lg"><span><i class="bg"></i>Được giao</span><span><i class="bx"></i>Hoàn thành</span><span class="meta">7 ngày gần: ' + w1 + ' hoàn thành</span></div></div></div></section>';
    /* Tầng 3: phân tích */
    var rows = shopRows();
    var tb = rows.map(function (r) { var x = r.x; return '<tr class="click' + (x.luot ? '' : ' idle') + '" data-shop="' + esc(r.s.id_shop) + '"' + (ST.shop === r.s.id_shop ? ' style="background:var(--tint)"' : '') + '><td class="nm"><b>' + esc(r.s.ten_shop) + '</b> ' + (r.tag ? '<span class="tag ' + (r.tag === 'Dẫn đầu' ? 't-ok' : 't-late') + '">' + r.tag + '</span>' : '') + '<small>' + esc(r.s.ma_shop) + ' · QL ' + esc(r.ql) + ' · ' + r.nv + ' NV</small></td><td class="num">' + x.luot + '</td><td class="num">' + x.mo + '</td><td class="pc"><div class="bar2"><i style="width:' + Math.max(0, r.ptr) + '%"></i></div><span>' + (r.ptr < 0 ? '–' : r.ptr + '%') + '</span></td><td class="num">' + x.dy + '</td><td class="num">' + (x.qh ? '<span class="tag t-late">' + x.qh + '</span>' : '0') + '</td></tr>' }).join('');
    h += '<section class="tier"><h2 class="tl">Phân tích</h2><div class="card" style="padding:12px 12px 6px"><div class="row2" style="margin-bottom:8px"><h2 style="margin:0">So sánh các Shop <span class="meta">' + rows.length + '/' + SH.length + '</span></h2><div class="shtool"><div class="search">' + ic('search') + '<input id="sq" placeholder="Tìm Shop, quản lý…" aria-label="Tìm Shop" value="' + esc(ST.sq || '') + '"></div><button class="chip' + (ST.late ? ' on' : '') + '" id="slate">Chỉ Shop quá hạn</button></div></div>' +
      '<div class="tw"><table class="shtb"><thead><tr>' + sortTh('ten', 'Shop', 0) + sortTh('luot', 'Được giao', 1) + sortTh('mo', 'Đang mở', 1) + sortTh('ptr', 'Hoàn thành', 0) + sortTh('dy', 'Hẹn / tiêm', 1) + sortTh('qh', 'Quá hạn', 1) + '</tr></thead><tbody>' + (tb || '<tr><td colspan="6" class="empty">Không có Shop phù hợp.</td></tr>') + '</tbody></table></div>' +
      '<div class="meta" style="padding:8px 4px">Bấm một Shop để xem nhân viên bên dưới. Bấm tiêu đề cột để sắp xếp.</div></div></section>';
    if (ST.shop) {
      var staff = staffOf(ST.shop);
      h += '<div class="card" style="padding:6px 8px"><h2 style="padding:10px 10px 0">Nhân viên · ' + esc(shopName(ST.shop)) + '</h2><div class="tw"><table><thead><tr><th>Nhân viên</th><th>Chức danh</th><th>Được giao</th><th>Đang mở</th><th>Hoàn thành</th><th>Hẹn / đã tiêm</th><th>Quá hạn</th></tr></thead><tbody>' + staff.map(function (a) {
        var T = D.turns.filter(function (t) { return t.recipient_type === 'EMPLOYEE' && t.ma_tk_nhan === a.ma_tk }), mo = T.filter(open_).length, xx = T.filter(function (t) { return t.trang_thai === 'HOAN_THANH' }).length, dy = T.filter(yes).length, qh = T.filter(function (t) { return overdueDays(t) > 0 }).length;
        return '<tr><td><b>' + esc(a.ho_ten) + '</b></td><td>' + (a.vai_tro === 'QUAN_LY' ? 'Quản lý' : 'Nhân viên') + '</td><td>' + T.length + '</td><td>' + mo + '</td><td>' + xx + '</td><td>' + dy + '</td><td>' + (qh ? '<span class="tag t-late">' + qh + '</span>' : '0') + '</td></tr>' }).join('') + '</tbody></table></div></div>';
    }
    return h
  }

  /* ---------- Kho Data ---------- */
  function khoRows() {
    var tc = turnsByCust(), q = ST.q.trim().toLowerCase(), f0 = ST.kFrom ? new Date(ST.kFrom + 'T00:00:00') : null, f1 = ST.kTo ? new Date(ST.kTo + 'T23:59:59') : null;
    return Object.keys(D.cus).map(function (k) { var c = D.cus[k], ts = tc[k] || [], l = (D.links[k] || [])[0] || {}; return { c: c, ts: ts, st: custState(ts), l: l } }).filter(function (r) {
      if (ST.f && r.st !== ST.f) return false; if (ST.band && bandOf(r.c) !== ST.band) return false;
      if (ST.kShop && !r.ts.some(function (t) { return t.id_shop === ST.kShop })) return false;
      if (ST.kNv && !r.ts.some(function (t) { return t.recipient_type === 'EMPLOYEE' && t.ma_tk_nhan === ST.kNv })) return false;
      if ((f0 || f1) && !r.ts.some(function (t) { var d = t.d_giao; return d && !isNaN(d) && (!f0 || d >= f0) && (!f1 || d <= f1) })) return false;
      if (q && (r.c.ma_khach + ' ' + r.c.ma_tcqg + ' ' + r.c.ho_ten + ' ' + (r.l.sdt || '') + ' ' + (r.l.ho_ten || '')).toLowerCase().indexOf(q) < 0) return false; return true
    }).sort(function (a, b) { return a.c.ma_khach.localeCompare(b.c.ma_khach) })
  }
  function pgKho() {
    var all = khoRows(), base = Object.keys(D.cus).map(function (k) { return custState(turnsByCust()[k]) }), cnt = { '': base.length, kho: 0, dang: 0, xong: 0 }; base.forEach(function (s) { cnt[s]++ });
    var pages = Math.max(1, Math.ceil(all.length / PAGEN())); ST.pg = Math.min(ST.pg, pages); var rows = all.slice((ST.pg - 1) * PAGEN(), ST.pg * PAGEN()), n = Object.keys(ST.sel).length;
    var h = head('Kho Data', 'Toàn bộ khách. Lọc, chọn nhiều khách rồi bấm Phân công.', '');
    var SHK = shops(), staffK = ST.kShop ? staffOf(ST.kShop) : [], act = !!(ST.q || ST.band || ST.f || ST.kShop || ST.kNv || ST.kFrom || ST.kTo);
    var nAct = [ST.kShop, ST.kNv, ST.kFrom, ST.kTo, ST.band].filter(Boolean).length, kopen = !!ST.kOpen;
    h += '<div class="kf"><div class="search">' + ic('search') + '<input id="kq" placeholder="Tìm tên, mã khách, SĐT…" aria-label="Tìm khách" value="' + esc(ST.q) + '"></div><details class="kfa" id="kfa"' + (kopen ? ' open' : '') + '><summary>Bộ lọc' + (nAct ? ' · ' + nAct + ' đang bật' : '') + '</summary><div class="kfb">' +
      sel('kshop', [['', 'Mọi Shop']].concat(SHK.map(function (x) { return [x.id_shop, x.ten_shop] })), ST.kShop) +
      '<select id="knv" aria-label="Nhân viên"' + (ST.kShop ? '' : ' disabled') + '><option value="">' + (ST.kShop ? 'Mọi nhân viên' : 'Chọn Shop trước') + '</option>' + staffK.map(function (a) { return '<option value="' + esc(a.ma_tk) + '"' + (ST.kNv === a.ma_tk ? ' selected' : '') + '>' + esc(a.ho_ten) + '</option>' }).join('') + '</select>' +
      '<label class="kd"><span>Giao từ</span><input type="date" id="kfrom" value="' + esc(ST.kFrom || '') + '"></label><label class="kd"><span>đến</span><input type="date" id="kto" value="' + esc(ST.kTo || '') + '"></label>' +
      sel('fband', BANDS, ST.band) + (act ? '<button class="btn b3" id="kclr" style="min-height:42px">Xóa bộ lọc</button>' : '') + '</div></details></div>';
    h += '<div class="chips">' + [['', 'Tất cả'], ['kho', 'Chưa giao'], ['dang', 'Đang giao'], ['xong', 'Đã xong']].map(function (x) { return '<button class="chip ' + (ST.f === x[0] ? 'on' : '') + '" data-f="' + x[0] + '">' + x[1] + '<em>' + cnt[x[0]] + '</em></button>' }).join('') + '</div>';
    var tagK = { kho: '<span class="tag t-new">Chưa giao</span>', dang: '<span class="tag t-prog">Đang giao</span>', xong: '<span class="tag t-ok">Đã xong</span>' };
    var whoOf = function (r) { var op = r.ts.filter(open_)[0]; return op ? (op.recipient_type === 'SHOP' ? shopName(op.id_shop) + ' (Shop)' : nm(op.ma_tk_nhan)) : '' };
    h += '<div class="card" style="padding:6px 8px"><div class="tw khotw only-d"><table data-nopg><thead><tr><th style="width:36px"><input type="checkbox" id="selpg" aria-label="Chọn cả trang" ' + (rows.length && rows.every(function (r) { return ST.sel[r.c.ma_khach] }) ? 'checked' : '') + ' style="width:20px;height:20px"></th><th>Khách</th><th>Tuổi</th><th>Phụ huynh</th><th>Trạng thái</th><th>Đang giữ</th></tr></thead><tbody>' +
      rows.map(function (r) { return '<tr' + (ST.sel[r.c.ma_khach] ? ' class="picked"' : '') + '><td><input type="checkbox" data-k="' + esc(r.c.ma_khach) + '" aria-label="Chọn ' + esc(r.c.ho_ten) + '" ' + (ST.sel[r.c.ma_khach] ? 'checked' : '') + ' style="width:20px;height:20px"></td><td><b>' + esc(r.c.ho_ten) + '</b><small>' + esc(r.c.ma_tcqg || r.c.ma_khach) + '</small></td><td>' + esc(ageTxt(r.c)) + '</td><td>' + esc(r.l.ho_ten || '') + '<small>' + esc(r.l.sdt || '') + '</small></td><td>' + tagK[r.st] + '</td><td>' + esc(whoOf(r)) + '</td></tr>' }).join('') +
      (rows.length ? '' : '<tr><td colspan="6" class="empty">Không có khách phù hợp.</td></tr>') + '</tbody></table></div>' +
      '<div class="only-m kcards"><button class="btn b3" data-selpage style="min-height:40px;justify-self:start">Chọn cả trang</button>' + rows.map(function (r) { return '<label class="kc' + (ST.sel[r.c.ma_khach] ? ' picked' : '') + '"><input type="checkbox" data-k="' + esc(r.c.ma_khach) + '" aria-label="Chọn ' + esc(r.c.ho_ten) + '" ' + (ST.sel[r.c.ma_khach] ? 'checked' : '') + '><span class="kb"><span class="kt"><b>' + esc(r.c.ho_ten) + '</b>' + tagK[r.st] + '</span><small>' + esc(ageTxt(r.c)) + (r.l.ho_ten ? ' · ' + esc(r.l.ho_ten) : '') + '</small>' + (r.l.sdt ? '<small>' + esc(r.l.sdt) + '</small>' : '') + (whoOf(r) ? '<small>Đang giữ: ' + esc(whoOf(r)) + '</small>' : '') + '</span></label>' }).join('') + (rows.length ? '' : '<div class="empty">Không có khách phù hợp.</div>') + '</div>' +
      '<div class="row2" style="padding:10px"><span class="meta">' + all.length + ' khách · trang ' + ST.pg + '/' + pages + '</span><div style="display:flex;gap:8px;align-items:center"><select data-pgsize aria-label="Số dòng mỗi trang" style="border:1px solid var(--line);border-radius:10px;padding:6px 8px;background:var(--card);min-height:38px">' + [10, 15, 20, 30, 50, 100].map(function (n) { return '<option value="' + n + '"' + (n === PAGEN() ? ' selected' : '') + '>' + n + ' dòng</option>' }).join('') + '</select><button class="btn b3" data-pg="-1" style="min-height:38px" aria-label="Trang trước">‹</button><button class="btn b3" data-pg="1" style="min-height:38px" aria-label="Trang sau">›</button></div></div></div>';
    h += '<div class="bulk" style="' + (n ? '' : 'display:none') + '"><b>Đã chọn ' + n + ' khách</b><span class="sp"></span>' + (all.length > n ? '<button class="btn b3" id="seltop" style="min-height:42px" title="Chọn cả ' + all.length + ' khách đang lọc">Chọn cả ' + all.length + '</button>' : '') + '<button class="btn b3" id="selclr" style="min-height:42px">Bỏ chọn</button><button class="btn b1" id="asg" style="min-height:42px">Phân công</button></div>';
    return h
  }
  function assignModal() {
    var ks = Object.keys(ST.sel); if (!ks.length) return; var tc = turnsByCust(), openN = ks.filter(function (k) { return custState(tc[k]) === 'dang' }).length;
    var topics = D.dm ? Object.keys(D.dm).map(function (k) { return D.dm[k] }).filter(function (c) { return c.loai === 'CHU_DE' && c.dang_dung }) : [];
    var m = { kieu: 'NHAN_VIEN', who: {}, shopSel: {}, ql: false, giaolai: false };
    var d = document.createElement('div'); d.className = 'modal';
    function draw() {
      var oldCard = d.querySelector('.card'), keepTop = oldCard ? oldCard.scrollTop : 0;
      var SH = shops(), nWho = Object.keys(m.who).length, nShop = Object.keys(m.shopSel).length, n = m.kieu === 'SHOP' ? nShop : nWho, per = n ? Math.ceil((ks.length - (m.giaolai ? 0 : openN)) / n) : 0;
      d.innerHTML = '<div class="card" style="width:100%;max-width:640px;border-radius:22px 22px 0 0;padding:18px;max-height:92vh;overflow:auto"><h2 style="font-size:19px;margin:0 0 4px">Phân công ' + ks.length + ' khách</h2><div class="meta" style="margin-bottom:12px">Chia đều cho những người hoặc Shop bạn chọn. Mỗi khách một lượt giao riêng, giữ nguyên mã khách.</div>' +
        '<div class="chips" style="margin-bottom:10px"><button class="chip ' + (m.kieu === 'NHAN_VIEN' ? 'on' : '') + '" data-kieu="NHAN_VIEN">Đích danh nhân viên</button><button class="chip ' + (m.kieu === 'SHOP' ? 'on' : '') + '" data-kieu="SHOP">Cả Shop</button></div>' +
        (m.kieu === 'NHAN_VIEN' ? SH.map(function (s) { var st = staffOf(s.id_shop); return st.length ? '<div class="sec"><h3>' + esc(s.ten_shop) + ' <a href="#" data-allshop="' + esc(s.id_shop) + '" style="text-transform:none;letter-spacing:0;font-weight:600;color:var(--teal-d)">' + (st.every(function (a) { return m.who[a.ma_tk] }) ? 'bỏ chọn cả Shop' : 'chọn cả Shop') + ' (' + st.filter(function (a) { return m.who[a.ma_tk] }).length + '/' + st.length + ')</a></h3><div style="display:grid;gap:6px">' + st.map(function (a) { return '<label class="row" style="cursor:pointer"><input type="checkbox" data-w="' + esc(a.ma_tk) + '" ' + (m.who[a.ma_tk] ? 'checked' : '') + ' style="width:20px;height:20px;accent-color:var(--teal)"><span>' + esc(a.ho_ten) + ' <span class="meta">· ' + (a.vai_tro === 'QUAN_LY' ? 'Quản lý' : 'Nhân viên') + '</span></span></label>' }).join('') + '</div></div>' : '' }).join('')
          : '<div class="sec"><h3>Chọn Shop nhận</h3><div style="display:grid;gap:6px">' + SH.map(function (s) { return '<label class="row" style="cursor:pointer"><input type="checkbox" data-sh="' + esc(s.id_shop) + '" ' + (m.shopSel[s.id_shop] ? 'checked' : '') + ' style="width:20px;height:20px;accent-color:var(--teal)"><span>' + esc(s.ten_shop) + ' <span class="meta">· Quản lý ' + esc(qlOf(s.id_shop)) + '</span></span></label>' }).join('') + '</div></div>' +
            '<label class="sw"><input type="checkbox" id="mql" ' + (m.ql ? 'checked' : '') + '><div><b>Cho phép Quản lý chia lại cho nhân viên</b><small>Chỉ áp cho lần giao này. Bạn tắt được bất cứ lúc nào ở mục Phân công.</small></div></label>') +
        '<div class="form2"><div><label class="meta">Chủ đề</label>' + sel('mcd', [['', '— chọn —']].concat(topics.map(function (c) { return [c.ma_dm, c.ten] })), '') + '</div><div><label class="meta">Hạn liên hệ</label><input type="date" id="mhan" value="' + new Date().toISOString().slice(0, 10) + '"></div></div>' +
        (openN ? '<label class="sw" style="margin-top:8px"><input type="checkbox" id="mgl" ' + (m.giaolai ? 'checked' : '') + '><div><b>Vẫn giao ' + openN + ' khách đang có lượt mở</b><small>Mặc định bỏ qua để tránh nhiều người cùng gọi một khách. Bật là giao ngoại lệ.</small></div></label>' : '') +
        '<div class="alert" style="margin-top:10px"><i>Σ</i><div><b>' + (n ? 'Giao ' + (ks.length - (m.giaolai ? 0 : openN)) + ' khách cho ' + n + (m.kieu === 'SHOP' ? ' Shop' : ' người') + ' · khoảng ' + per + ' khách mỗi nơi' : 'Chưa chọn người nhận') + '</b></div></div>' +
        '<div class="acts" style="margin-top:6px"><button class="btn b3" id="mx" style="flex:1">Hủy</button><button class="btn b1" id="mok" style="flex:1.4" ' + (n ? '' : 'disabled') + '>Xác nhận giao</button></div></div>';
      var newCard = d.querySelector('.card'); if (newCard) newCard.scrollTop = keepTop;
      d.querySelectorAll('[data-kieu]').forEach(function (b) { b.onclick = function () { m.kieu = b.dataset.kieu; draw() } });
      d.querySelectorAll('[data-w]').forEach(function (b) { b.onchange = function () { b.checked ? m.who[b.dataset.w] = 1 : delete m.who[b.dataset.w]; draw() } });
      d.querySelectorAll('[data-sh]').forEach(function (b) { b.onchange = function () { b.checked ? m.shopSel[b.dataset.sh] = 1 : delete m.shopSel[b.dataset.sh]; draw() } });
      d.querySelectorAll('[data-allshop]').forEach(function (a) { a.onclick = function (e) { e.preventDefault(); var st = staffOf(a.dataset.allshop), all = st.every(function (x) { return m.who[x.ma_tk] }); st.forEach(function (x) { if (all) delete m.who[x.ma_tk]; else m.who[x.ma_tk] = 1 }); draw() } });
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
    var h = head('Nạp Data', 'Dán nội dung file CSV (hoặc chọn file). App kiểm tra, xem trước, rồi mới nạp. Khách trùng Mã hồ sơ sẽ được bỏ qua.', '');
    h += '<div class="card"><div class="acts" style="margin-bottom:10px"><button class="btn b3" id="tpl">' + ic('download') + 'Tải file mẫu</button><label class="btn b2" style="cursor:pointer">' + ic('upload') + 'Chọn file CSV<input type="file" id="nfile" accept=".csv,.txt" hidden></label></div><textarea id="ntx" style="min-height:140px;font-family:ui-monospace,Menlo,monospace;font-size:13px" placeholder="ma_tcqg,ho_ten,ngay_sinh,gioi_tinh,ten_lien_he,quan_he,sdt">' + esc(ST.impText || '') + '</textarea><div class="acts" style="margin-top:10px"><button class="btn b1" id="nchk">Kiểm tra</button></div></div>';
    if (ST.imp) {
      var ok = ST.imp.filter(function (r) { return !r.err }), dup = ST.imp.filter(function (r) { return /bỏ qua/.test(r.err) }).length, bad = ST.imp.length - ok.length - dup;
      h += '<div class="stats stats4"><div class="card st"><div class="l">Dòng trong file</div><div class="v">' + ST.imp.length + '</div></div><div class="card st"><div class="l">Hợp lệ</div><div class="v" style="color:var(--ok)">' + ok.length + '</div></div><div class="card st"><div class="l">Đã có (bỏ qua)</div><div class="v">' + dup + '</div></div><div class="card st"><div class="l">Lỗi</div><div class="v" style="color:var(--bad)">' + bad + '</div></div></div>' +
        '<div class="card" style="padding:6px 8px"><div class="tw"><table><thead><tr><th>Dòng</th><th>Mã hồ sơ</th><th>Họ tên</th><th>Ngày sinh</th><th>Liên hệ</th><th>Kết quả kiểm tra</th></tr></thead><tbody>' + ST.imp.slice(0, 30).map(function (r) { return '<tr><td>' + r._dong + '</td><td>' + esc(r.ma_tcqg) + '</td><td>' + esc(r.ho_ten) + '</td><td>' + esc(r.ngay_sinh) + '</td><td>' + esc(r.ten_lien_he) + '<small>' + esc(r.sdt) + '</small></td><td>' + (r.err ? '<span class="tag ' + (/bỏ qua/.test(r.err) ? 't-new' : 't-late') + '">' + esc(r.err) + '</span>' : '<span class="tag t-ok">Hợp lệ</span>') + '</td></tr>' }).join('') + '</tbody></table></div></div>' +
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
    var T = bcTurns(), SH = shops(), by = {}, kq = {}, nKq = 0, nMo = 0, nXong = 0, nDy = 0;
    T.forEach(function (t) {
      var k = t.recipient_type === 'EMPLOYEE' ? t.ma_tk_nhan : 'SHOP:' + t.id_shop, r = by[k] = by[k] || { k: k, t: t, n: 0, mo: 0, xong: 0, dy: 0 }; r.n++;
      if (open_(t)) { r.mo++; nMo++ } if (t.trang_thai === 'HOAN_THANH') { r.xong++; nXong++ } if (yes(t)) { r.dy++; nDy++ }
      if (t.ma_ket_qua) { var nm0 = (cat(t.ma_ket_qua) || {}).ten || t.ma_ket_qua; kq[nm0] = (kq[nm0] || 0) + 1; nKq++ }
    });
    var h = head('Báo cáo', 'Kết quả chăm sóc theo Shop và nhân viên. Xuất file CSV mở được bằng Excel.', '');
    h += '<div class="tbar2">' + sel('bshop', [['', 'Tất cả Shop']].concat(SH.map(function (s) { return [s.id_shop, s.ten_shop] })), ST.bcShop) + sel('bst', [['', 'Mọi trạng thái'], ['MO', 'Đang mở'], ['HOAN_THANH', 'Hoàn thành'], ['THU_HOI', 'Đã thu hồi']], ST.bcSt) +
      '<details class="xp"><summary class="btn b1">' + ic('download') + 'Xuất file</summary><div class="xpm"><button class="btn b3" id="xl">Theo từng lượt (' + T.length + ')</button><button class="btn b3" id="xn">Theo nhân viên</button></div></details></div>';
    h += '<div class="kp3"><div class="card kp-m"><div class="l">Được giao</div><div class="v">' + T.length + '</div></div><div class="card kp-m"><div class="l">Hoàn thành</div><div class="v">' + nXong + '<small> ' + (T.length ? pct(nXong, T.length) : 0) + '%</small></div></div><div class="card kp-m"><div class="l">Hẹn / tiêm</div><div class="v ok">' + nDy + '<small> ' + (nXong ? pct(nDy, nXong) : 0) + '%</small></div></div></div>';
    var ks = Object.keys(kq).sort(function (a, b) { return kq[b] - kq[a] }), none = T.length - nKq, mx = Math.max(1, ks.length ? kq[ks[0]] : 1, none);
    h += '<div class="card"><div class="row2"><h2 style="margin:0">Kết quả ghi nhận</h2><span class="meta">' + nKq + '/' + T.length + ' lượt đã có kết quả</span></div>' + (ks.length ? '<div class="kqs">' + ks.map(function (k) { return '<div class="kqr"><span class="kqn">' + esc(k) + '</span><span class="kqb"><i style="width:' + Math.max(4, Math.round(kq[k] / mx * 100)) + '%"></i></span><b>' + kq[k] + '</b></div>' }).join('') + (none > 0 ? '<div class="kqr kqx"><span class="kqn">Chưa ghi kết quả</span><span class="kqb"><i style="width:' + Math.max(4, Math.round(none / mx * 100)) + '%"></i></span><b>' + none + '</b></div>' : '') + '</div>' : '<div class="meta" style="margin-top:8px">Chưa có lượt nào ghi kết quả trong phạm vi lọc.</div>') + '</div>';
    h += '<div class="card" style="padding:6px 8px"><h2 style="padding:10px 10px 0">Theo người nhận</h2><div class="tw"><table class="rt rtb" style="--c:5"><thead><tr><th>Người nhận</th><th>Được giao</th><th>Đang mở</th><th>Hoàn thành</th><th>Hẹn / đã tiêm</th><th>Tỷ lệ hẹn / hoàn thành</th></tr></thead><tbody>' + Object.keys(by).map(function (k) { return by[k] }).sort(function (a, b) { return b.n - a.n }).map(function (r) {
      return '<tr><td class="rh"><b>' + esc(r.k.indexOf('SHOP:') === 0 ? 'Cả Shop (chưa chia)' : nm(r.k)) + '</b><small>' + esc(shopName(r.t.id_shop)) + '</small></td><td data-label="Giao">' + r.n + '</td><td data-label="Mở">' + r.mo + '</td><td data-label="Xong">' + r.xong + '</td><td data-label="Hẹn">' + r.dy + '</td><td data-label="Tỷ lệ">' + (r.xong ? pct(r.dy, r.xong) + '%' : '–') + '</td></tr>' }).join('') + (T.length ? '' : '<tr><td colspan="6" class="empty">Không có lượt nào trong phạm vi lọc.</td></tr>') + '</tbody></table></div></div>';
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
  function pgAcc() { return head('Tài khoản của tôi', '', '') + '<div class="card"><div class="alert"><div class="ava">' + esc(ini(ME.ho_ten)) + '</div><div><b>' + esc(ME.ho_ten) + '</b><small>' + esc(ME.email) + ' · ASM</small></div></div><button class="btn b3" id="cpw" style="width:100%;margin-bottom:8px">Đổi mật khẩu</button><button class="btn b3" id="out" style="width:100%">Đăng xuất</button><div class="vxc" id="vxc" role="button" tabindex="0" title="Chạm để sao chép phiên bản"><b>' + BRAND.app + '</b>' + BRAND.mota + '<br><em>' + BRAND.nghia + '</em><br>Biên soạn: ' + BRAND.tacgia + ' · ' + BRAND.kv + '<br>' + BRAND.vtext() + ' · bản dựng ' + BRAND.build() + '</div></div>' }
  X.u = { head: head, sel: sel, shops: shops, staffOf: staffOf, qlOf: qlOf, pct: pct, ST: ST };

  /* ---------- vẽ và gắn sự kiện ---------- */
  function render() {
    document.body.setAttribute('data-role', 'asm'); shell(); var m = $('#main'), p = S.page; m.innerHTML = ({ a_tq: pgTq, a_kho: pgKho, a_pc: pgPc, a_nap: pgNap, a_bc: pgBc, acc: pgAcc }[p] || X.pages[p] || pgTq)(); stick(m); fitChart();
    var x, on = function (sel, ev, fn) { var el = $(sel); if (el) el[ev] = function (e) { return fn.call(el, el, e) } };
    document.querySelectorAll('[data-go]').forEach(function (b) { b.onclick = function () { S.page = b.dataset.go; render() } });
    on('#fshop', 'onchange', function (x) { ST.shop = x.value; ST.dtPg = 1; render() });
    document.querySelectorAll('[data-shop]').forEach(function (r) { r.onclick = function () { ST.shop = ST.shop === r.dataset.shop ? '' : r.dataset.shop; render() } });
    document.querySelectorAll('[data-sort]').forEach(function (b) { b.onclick = function (e) { e.stopPropagation(); var k = b.dataset.sort; if ((ST.ssort || 'qh') === k) ST.sdir = -(ST.sdir || -1); else { ST.ssort = k; ST.sdir = k === 'ten' ? 1 : -1 } render() } });
    on('#sq', 'oninput', function (x) { ST.sq = x.value; var pos = x.selectionStart; render(); var n = $('#sq'); n.focus(); n.setSelectionRange(pos, pos) });
    on('#slate', 'onclick', function () { ST.late = !ST.late; render() });
    document.querySelectorAll('[data-dt]').forEach(function (b) { b.onclick = function () { var y = window.scrollY, k = b.dataset.dt, sc = b.dataset.src; if (ST.dt === k && ST.dtSrc === sc) { ST.dt = ''; ST.dtSrc = '' } else { ST.dt = k; ST.dtSrc = sc; ST.dtPg = 1 } render(); window.scrollTo(0, y) } });
    document.querySelectorAll('[data-dtp]').forEach(function (b) { b.onclick = function () { var y = window.scrollY; ST.dtPg = Math.max(1, (ST.dtPg || 1) + Number(b.dataset.dtp)); render(); window.scrollTo(0, y) } });
    document.querySelectorAll('[data-dtx]').forEach(function (b) { b.onclick = function () { var y = window.scrollY; ST.dt = ''; ST.dtSrc = ''; render(); window.scrollTo(0, y) } });
    document.querySelectorAll('[data-dtgo]').forEach(function (b) { b.onclick = function () { var g = b.dataset.dtgo; if (g === 'late') { ST.late = true; render(); var t = $('.shtb'); if (t && t.scrollIntoView) t.scrollIntoView({ behavior: 'smooth', block: 'start' }) } else { S.page = g; render() } } });
    document.querySelectorAll('[data-late]').forEach(function (b) { b.onclick = function () { ST.late = true; render(); var t = $('.shtb'); if (t && t.scrollIntoView) t.scrollIntoView({ behavior: 'smooth', block: 'start' }) } });
    on('#kq', 'oninput', function (x) { ST.q = x.value; ST.pg = 1; var pos = x.selectionStart; render(); var n = $('#kq'); n.focus(); n.setSelectionRange(pos, pos) });
    on('#fband', 'onchange', function (x) { ST.band = x.value; ST.pg = 1; render() });
    on('#kshop', 'onchange', function (x) { ST.kShop = x.value; ST.kNv = ''; ST.pg = 1; render() }); on('#knv', 'onchange', function (x) { ST.kNv = x.value; ST.pg = 1; render() });
    on('#kfrom', 'onchange', function (x) { ST.kFrom = x.value; ST.pg = 1; render() }); on('#kto', 'onchange', function (x) { ST.kTo = x.value; ST.pg = 1; render() });
    on('#kfa', 'ontoggle', function (x) { ST.kOpen = x.open });
    on('#kclr', 'onclick', function () { ST.q = ST.band = ST.f = ST.kShop = ST.kNv = ST.kFrom = ST.kTo = ''; ST.pg = 1; render() });
    document.querySelectorAll('[data-selpage]').forEach(function (b) { b.onclick = function () { var all = khoRows(), rs = all.slice((ST.pg - 1) * PAGEN(), ST.pg * PAGEN()), every = rs.every(function (r) { return ST.sel[r.c.ma_khach] }); rs.forEach(function (r) { every ? delete ST.sel[r.c.ma_khach] : ST.sel[r.c.ma_khach] = 1 }); render() } });
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
    on('#lkbtn', 'onclick', function () { if (window.MSG) MSG.lookup() });
    on('#vxc', 'onclick', function () { copy(BRAND.app + ' ' + BRAND.vtext() + ' · bản dựng ' + BRAND.build()); toast('Đã sao chép phiên bản') });
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
  var SAU = [['Mẫu tin quy trình', 'Các bước D, D+1, D+3, D+7 theo nhóm Lao/6in1. Làm ở đợt sau.'], ['Combo ưu đãi', 'Gói combo dạng bảng. Làm ở đợt sau.']];

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
/* NHÂN SỰ – NẠP DANH SÁCH TỪ FILE TOÀN BỘ NHÂN SỰ (bản xuất của hệ thống nhân sự): kéo thả file Excel/CSV, app tự tìm dòng tiêu đề (kể cả tiêu đề 2 tầng), nhận cột Họ tên / Email / Chức danh /
 * Mã shop / Cửa hàng / Trạng thái (nhiều cột giống tên thì chọn cột CÓ DỮ LIỆU), so với tài khoản và Shop đang có rồi cho xem trước: Tạo mới / Cập nhật / Không đổi / Lỗi / Bỏ qua.
 * Anh đặt "Chức danh → vai trò" (Nhân viên / Quản lý / Bỏ qua) ngay trên màn xem trước. Shop chưa có được tạo từ file (Mã shop + tên cửa hàng).
 * Bấm Áp dụng mới ghi, từng người, qua hàm chỉ ASM gọi được. Tài khoản mới có mật khẩu ngẫu nhiên hiện một lần, tải được CSV. Không tự khóa người vắng trong file. */
var ASMHR = (function () {
  var P = window.LIBP, H = { aoa: null, name: '', info: null, updRole: true, getSdt: false, createShops: true, onlyActive: true, rules: {}, res: null, busy: false, drag: false };
  var COLS = [
    { k: 'ho_ten', nhan: 'Họ tên', alias: ['ho va ten', 'ho ten', 'ten nhan vien', 'nhan vien', 'ho ten nhan vien', 'full name', 'name', 'ten'], bb: 1 },
    { k: 'email', nhan: 'Email', alias: ['mail', 'dia chi email', 'email cong ty', 'email nhan vien', 'email lien he'], bb: 1 },
    { k: 'chuc_danh', nhan: 'Chức danh', alias: ['vai tro', 'chuc vu cong viec', 'vi tri', 'position', 'role', 'chuc vu'] },
    { k: 'ma_shop', nhan: 'Mã shop', alias: ['ma co so', 'ma cua hang', 'ma nha thuoc', 'ma don vi', 'shop id'] },
    { k: 'ten_shop', nhan: 'Cửa hàng', alias: ['ten shop', 'ten co so', 'co so lam viec', 'co so', 'nha thuoc', 'ten cua hang', 'shop', 'dia diem lam viec', 'noi lam viec', 'trung tam'] },
    { k: 'trang_thai', nhan: 'Trạng thái', alias: ['tinh trang', 'trang thai lam viec'] },
    { k: 'sdt', nhan: 'Số điện thoại', alias: ['sdt', 'dien thoai', 'so dt', 'phone'] }, { k: 'ma_nv', nhan: 'Mã nhân viên', alias: ['ma nv', 'ma so nhan vien'] }];
  var KQ = { moi: ['t-ok', 'Tạo mới'], sua: ['t-today', 'Cập nhật'], giu: ['t-new', 'Không đổi'], loi: ['t-bad', 'Lỗi'], bo: ['t-new', 'Bỏ qua'] };
  var VT = { QUAN_LY: 'Quản lý', NHAN_VIEN: 'Nhân viên', BO: 'Bỏ qua' };

  function xlsx(cb) { if (window.XLSX) return cb(); var s = document.createElement('script'); s.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'; s.onload = cb; s.onerror = function () { toast('Không tải được bộ đọc Excel. Kiểm tra mạng.', 0, 1) }; document.head.appendChild(s) }
  function genPw() { var a = 'abcdefghjkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789', o = '', b = new Uint32Array(10); (window.crypto || window.msCrypto).getRandomValues(b); for (var i = 0; i < 10; i++) o += a[b[i] % a.length]; return o + '7' }
  function tag(c, t) { return '<span class="tag ' + c + '">' + esc(t) + '</span>' }
  /* SĐT trong file Excel hay mất số 0 đầu (9 chữ số) -> thêm lại; không hợp lệ -> bỏ trống */
  function normPhone(v) { var s = String(v == null ? '' : v).replace(/[^0-9+]/g, ''); if (s.indexOf('+84') === 0) s = '0' + s.slice(3); else if (/^84\d{9}$/.test(s)) s = '0' + s.slice(2); else if (/^[1-9]\d{8}$/.test(s)) s = '0' + s; return /^0\d{9}$/.test(s) ? s : '' }
  function core(s) { return P.bo(String(s || '').replace(/\([^)]*\)/g, '')).replace(/^(shop|trung tam|nha thuoc|co so|lc|vx)\s+/, '').replace(/^(hni|hn)\s+/, '').trim() }
  /* quy tắc mặc định theo chức danh (anh sửa được trên màn xem trước) */
  function defaultRule(t) { var b = P.bo(t); if (!b) return 'NHAN_VIEN'; if (/giam doc|(^| )asm( |$)/.test(b)) return 'BO'; if (/(^| )bac si( |$)|bao ve/.test(b)) return 'BO'; if (/truong quan ly|pho quan ly|cua hang truong|quan ly (cua hang|shop)|truong ca/.test(b)) return 'QUAN_LY'; return 'NHAN_VIEN' }
  /* Một Shop có thể có NHIỀU Quản lý (Hải 24/11): mọi người mang chức danh được đặt là Quản lý đều là Quản lý. */
  function describeEx(row) { var ex = row.ex, ch = []; if (ex.ho_ten !== row.ho_ten) ch.push('tên'); if (ex.id_shop !== row.id_shop) ch.push('Shop'); if (row.hasCd && ex.vai_tro !== row.vai_tro) ch.push(VT[ex.vai_tro] + ' → ' + VT[row.vai_tro]); row.kq = ch.length ? 'sua' : 'giu'; row.ghi = ch.length ? 'Đổi: ' + ch.join(', ') : 'Giống tài khoản hiện có' + (ex.trang_thai === 'NGUNG' ? ' (đang bị khóa)' : '') }

  /* tìm dòng tiêu đề; mỗi cột chọn trong các cột cùng tên cột CÓ NHIỀU DỮ LIỆU nhất */
  function findHdr(aoa) {
    var best = null;
    for (var r = 0; r < Math.min(aoa.length, 40); r++) {
      var row = aoa[r] || [], map = {}, taken = {}, n = 0, fill = function (j) { var c = 0; for (var i = r + 1; i < Math.min(aoa.length, r + 60); i++) { var v = (aoa[i] || [])[j]; if (v != null && String(v).trim() !== '') c++ } return c };
      COLS.forEach(function (c) {
        var names = [c.nhan, c.k].concat(c.alias || []).map(function (a) { return P.bo(a) }), cand = [], j, b;
        for (j = 0; j < row.length; j++) { if (taken[j] || row[j] == null) continue; var h = P.bo(row[j]); if (!h) continue; var ix = names.indexOf(h); if (ix >= 0) cand.push({ j: j, p: ix }); else { b = ' ' + h + ' '; var k2 = -1; names.some(function (a, i) { if (a.length >= 4 && b.indexOf(' ' + a + ' ') >= 0) { k2 = i; return true } return false }); if (k2 >= 0) cand.push({ j: j, p: 100 + k2 }) } }
        if (!cand.length) return; cand.forEach(function (x) { x.f = fill(x.j) }); cand.sort(function (a, b2) { return (b2.f > 0) - (a.f > 0) || a.p - b2.p || b2.f - a.f }); map[c.k] = cand[0].j; taken[cand[0].j] = 1; n++
      });
      if (!best || n > best.n) best = { r: r, map: map, n: n }
    }
    return best || { r: 0, map: {}, n: 0 }
  }
  function findShop(code, name, shops) {
    var c = String(code || '').trim(), tc = core(name), hit = [];
    if (c) hit = shops.filter(function (s) { return String(s.ma_shop) === c || s.id_shop === 'S' + c });
    if (!hit.length && tc) hit = shops.filter(function (s) { var x = core(s.ten_shop); return x && (x === tc || (tc.length >= 6 && (x.indexOf(tc) >= 0 || tc.indexOf(x) >= 0))) });
    return hit.length === 1 ? { s: hit[0] } : { none: true, n: hit.length }
  }
  /* PHÂN TÍCH (hàm thuần, kiểm thử được): aoa + {shops:[], accs:[]} + tùy chọn */
  function parse(aoa, ctx, o) {
    var hd = findHdr(aoa), miss = COLS.filter(function (c) { return c.bb && hd.map[c.k] === undefined });
    if (!hd.n || miss.length) return { loi: 'Không nhận ra cột bắt buộc: ' + miss.map(function (c) { return c.nhan }).join(', ') + '. Các cột đọc được ở dòng ' + (hd.r + 1) + ': ' + (aoa[hd.r] || []).filter(function (v) { return v != null && String(v).trim() }).slice(0, 14).join(' | ') };
    var cell = function (r, k) { var j = hd.map[k]; return j === undefined ? '' : String(r[j] == null ? '' : r[j]).trim() }, byEmail = {}, titles = {}, seen = {}, rows = [], newShops = {}, shops = ctx.shops.slice();
    ctx.accs.forEach(function (a) { byEmail[String(a.email).toLowerCase()] = a });
    var raw = []; for (var i = hd.r + 1; i < aoa.length; i++) { var r = aoa[i] || []; if (!r.some(function (v) { return v != null && String(v).trim() })) continue; var ho = cell(r, 'ho_ten'), em = cell(r, 'email').toLowerCase().replace(/\s+/g, ''); if (!ho && !em) continue; raw.push({ dong: i + 1, r: r, ho: ho, em: em, cd: cell(r, 'chuc_danh') }); var t = raw[raw.length - 1].cd; if (t) titles[t] = (titles[t] || 0) + 1 }
    var rules = {}; Object.keys(titles).forEach(function (t) { rules[t] = (o.rules && o.rules[t]) || defaultRule(t) });
    raw.forEach(function (x) {
      var r = x.r, rule = x.cd ? rules[x.cd] : 'NHAN_VIEN', st = cell(r, 'trang_thai'), code = cell(r, 'ma_shop'), nm = cell(r, 'ten_shop');
      var row = { dong: x.dong, ho_ten: x.ho, email: x.em, chuc_danh: x.cd, shopTxt: nm || code, ma_shop: code, ten_shop: nm, vai_tro: rule === 'QUAN_LY' ? 'QUAN_LY' : 'NHAN_VIEN', sdt: o.getSdt ? normPhone(cell(r, 'sdt')) : '' }; rows.push(row);
      if (!x.ho) { row.kq = 'loi'; row.ghi = 'Thiếu họ tên'; return }
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x.em)) { row.kq = 'loi'; row.ghi = 'Email không hợp lệ'; return }
      if (seen[x.em]) { row.kq = 'loi'; row.ghi = 'Trùng email với dòng ' + seen[x.em]; return } seen[x.em] = x.dong;
      if (rule === 'BO') { row.kq = 'bo'; row.ghi = 'Chức danh "' + x.cd + '" đặt là Bỏ qua'; return }
      if (o.onlyActive && st && !/dang lam|hoat dong|^lam viec$/.test(P.bo(st))) { row.kq = 'bo'; row.ghi = 'Không đang làm (' + st + ')'; return }
      var ex = byEmail[x.em]; if (ex && ex.vai_tro === 'ASM') { row.kq = 'bo'; row.ghi = 'Là tài khoản ASM'; return }
      var sh = findShop(code, nm, shops), id = sh.s ? sh.s.id_shop : null;
      if (!id) { if (code && nm && o.createShops && /^[A-Za-z0-9]{3,20}$/.test(code)) { id = 'S' + code; var ns = newShops[id] || (newShops[id] = { ma: code, ten: nm, n: 0 }); ns.n++; row.shopMoi = true } else if (ex && !code && !nm) id = ex.id_shop; else { row.kq = 'loi'; row.ghi = !code && !nm ? 'Thiếu Shop' : sh.n > 1 ? 'Nhiều Shop khớp' : 'Chưa có Shop "' + (nm || code) + '" (bật "Tạo Shop chưa có")'; return } }
      row.id_shop = id; var shName = sh.s ? sh.s.ten_shop : (newShops[id] ? newShops[id].ten : id);
      if (!ex) { row.kq = 'moi'; row.ghi = VT[row.vai_tro] + ' · ' + shName + (row.shopMoi ? ' (Shop mới)' : ''); row.mk = genPw() }
      else { row.ma_tk = ex.ma_tk; row.ex = ex; row.hasCd = !!x.cd && o.updRole !== false; if (!row.hasCd) row.vai_tro = ex.vai_tro; describeEx(row); if (row.sdt && !ex.sdt) { row.addSdt = true; if (row.kq === 'giu') { row.kq = 'sua'; row.ghi = 'Bổ sung SĐT' } else row.ghi += ', bổ sung SĐT' } }
    });
    /* chỉ để báo: Shop trong file mà chưa có Quản lý nào (cả trong file lẫn trong hệ thống) */
    var shopsIn = {}, hasQL = {}; rows.forEach(function (o) { if (o.id_shop && (o.kq === 'moi' || o.kq === 'sua' || o.kq === 'giu')) { shopsIn[o.id_shop] = 1; if (o.vai_tro === 'QUAN_LY') hasQL[o.id_shop] = 1 } });
    var noQL = Object.keys(shopsIn).filter(function (id) { return !hasQL[id] && !ctx.accs.some(function (a) { return a.vai_tro === 'QUAN_LY' && a.id_shop === id && a.trang_thai === 'HOAT_DONG' }) }).length;
    var inFile = {}; rows.forEach(function (x) { inFile[x.email] = 1 }); var vangL = ctx.accs.filter(function (a) { return a.vai_tro !== 'ASM' && a.trang_thai === 'HOAT_DONG' && !inFile[String(a.email).toLowerCase()] });
    return { rows: rows, vang: vangL.length, vangList: vangL.map(function (a) { return { ho_ten: a.ho_ten, vai_tro: a.vai_tro, id_shop: a.id_shop } }), noQL: noQL, titles: Object.keys(titles).map(function (t) { return { t: t, n: titles[t], rule: rules[t] } }).sort(function (a, b) { return b.n - a.n }), newShops: Object.keys(newShops).map(function (k) { return newShops[k] }), dongTieuDe: hd.r + 1, map: hd.map }
  }
  function accs() { return Object.keys(D.acc).map(function (k) { return D.acc[k] }) }
  function shopsAct() { return Object.keys(D.shops).map(function (k) { return D.shops[k] }).filter(function (s) { return s.trang_thai === 'HOAT_DONG' }) }
  function redo() { if (!H.aoa) return; var keep = H.info && H.info.sheet; H.info = parse(H.aoa, { shops: shopsAct(), accs: accs() }, { rules: H.rules, onlyActive: H.onlyActive, createShops: H.createShops, updRole: H.updRole, getSdt: H.getSdt }); if (H.info.rows) H.info.sheet = keep }
  function handle(name, buf) {
    xlsx(function () {
      try {
        var wb = XLSX.read(buf, { type: 'array', cellDates: true, codepage: 65001 }), best = null;
        wb.SheetNames.forEach(function (n) { var aoa = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null, blankrows: true }), hd = findHdr(aoa); if (!best || hd.n > best.n) best = { n: hd.n, aoa: aoa, name: n } });
        H.aoa = best.aoa; H.name = name; H.res = null; H.rules = {}; redo(); if (H.info.rows) H.info.sheet = best.name; render()
      } catch (e) { toast('Không đọc được file: ' + e.message, 0, 1) }
    })
  }
  function card() {
    var h = '<div class="card"><h2>Nạp danh sách nhân sự từ file <span class="pill hrpill">kéo file toàn bộ nhân sự vào đây</span></h2>';
    h += '<div id="hr-drop" style="border:2px dashed ' + (H.drag ? 'var(--teal)' : 'var(--line)') + ';border-radius:16px;padding:var(--dzp,18px);text-align:center;background:' + (H.drag ? 'var(--tint)' : 'var(--bg)') + '"><div style="font-weight:600"><span class="only-d">Kéo thả file Excel/CSV vào đây</span><span class="only-m">Tải file nhân sự (Excel/CSV)</span></div><details class="hrh"' + (window.innerWidth >= 960 ? ' open' : '') + '><summary>File cần có cột nào?</summary><div class="meta" style="margin:4px 0 10px">App tự nhận cột Họ tên, Email, Chức danh, Mã shop, Cửa hàng, Trạng thái (tên cột khác một chút vẫn nhận). Xem trước rồi mới tạo.</div></details><label class="btn b2" style="cursor:pointer">' + ic('upload') + 'Hoặc chọn file<input type="file" id="hr-file" accept=".xlsx,.xls,.csv" hidden></label></div>';
    var f = H.info; if (f && f.loi) h += '<div class="err" style="margin-top:10px">' + esc(f.loi) + '</div>';
    if (f && f.rows) {
      var c = { moi: 0, sua: 0, giu: 0, loi: 0, bo: 0 }; f.rows.forEach(function (o) { c[o.kq]++ }); var todo = c.moi + c.sua;
      h += '<div class="meta" style="margin:10px 0 6px">' + esc(H.name) + (f.sheet ? ' · sheet ' + esc(f.sheet) : '') + ' · dòng tiêu đề ' + f.dongTieuDe + ' · ' + f.rows.length + ' dòng</div><div class="stats stats4" style="grid-template-columns:repeat(5,1fr)">' + [['Tạo mới', c.moi, 'var(--ok)'], ['Cập nhật', c.sua, 'var(--teal-d)'], ['Không đổi', c.giu, 'inherit'], ['Lỗi', c.loi, c.loi ? 'var(--bad)' : 'inherit'], ['Bỏ qua', c.bo, 'inherit']].map(function (x) { return '<div class="card st"><div class="l">' + x[0] + '</div><div class="v" style="color:' + x[2] + '">' + x[1] + '</div></div>' }).join('') + '</div>';
      h += '<div class="card pgskip" style="margin:8px 0"><h2>Chức danh → vai trò trong app <span class="pill">anh sửa được</span></h2><div class="tw"><table><thead><tr><th>Chức danh trong file</th><th>Số người</th><th>Vai trò</th></tr></thead><tbody>' + f.titles.map(function (t, i) { return '<tr><td>' + esc(t.t) + '</td><td>' + t.n + '</td><td><select data-hrt="' + i + '" style="border:1px solid var(--line);border-radius:10px;padding:6px;background:var(--card);min-height:38px">' + ['NHAN_VIEN', 'QUAN_LY', 'BO'].map(function (k) { return '<option value="' + k + '"' + (t.rule === k ? ' selected' : '') + '>' + VT[k] + '</option>' }).join('') + '</select></td></tr>' }).join('') + '</tbody></table></div><div class="meta" style="padding:6px">Một Shop có thể có nhiều Quản lý. "Bỏ qua" = không tạo tài khoản cho chức danh đó.</div></div>';
      h += '<div style="display:grid;gap:8px"><label class="sw"><input type="checkbox" id="hr-act" ' + (H.onlyActive ? 'checked' : '') + '><div><b>Chỉ lấy người đang làm</b><small>Bỏ qua người nghỉ sinh, nghỉ việc… theo cột Trạng thái.</small></div></label><label class="sw"><input type="checkbox" id="hr-shop" ' + (H.createShops ? 'checked' : '') + '><div><b>Tạo Shop chưa có trong hệ thống (' + f.newShops.length + ' Shop)</b><small>' + (f.newShops.length ? f.newShops.slice(0, 5).map(function (s) { return esc(s.ma) + ' · ' + esc(s.ten.slice(0, 40)) }).join('<br>') + (f.newShops.length > 5 ? '<br>…' : '') : 'Tất cả Shop trong file đã có.') + '</small></div></label><label class="sw"><input type="checkbox" id="hr-role" ' + (H.updRole ? 'checked' : '') + '><div><b>Cập nhật vai trò theo chức danh cho người đã có tài khoản</b><small>Bật: người đã có tài khoản mà vai trò khác bảng trên sẽ được đổi. Tắt: người đã có giữ nguyên vai trò, file chỉ để thêm người mới và sửa tên/Shop. Người đúng tên, đúng Shop, đúng vai trò luôn được giữ nguyên.</small></div></label><label class="sw"><input type="checkbox" id="hr-sdt" ' + (H.getSdt ? 'checked' : '') + '><div><b>Lấy SĐT từ file (chỉ điền cho người chưa có SĐT)</b><small>Mặc định TẮT vì SĐT trong file nhân sự thường là số cá nhân. SĐT của nhân viên sẽ hiện trong chữ ký mọi tin nhắn gửi khách: để nhân viên tự nhập ở mục Tài khoản, hoặc anh nhập tay ở bảng Tài khoản.</small></div></label></div>';
      if (f.noQL) h += '<div class="alert" style="margin-top:8px"><i>ℹ</i><div><b>' + f.noQL + ' Shop chưa có Quản lý</b><small>Chưa có ai mang chức danh Quản lý cho các Shop này. Anh chọn Quản lý sau bằng nút "Chọn Quản lý" ở bảng Shop bên dưới.</small></div></div>';
      if (f.vang) h += '<div class="alert" style="margin-top:8px"><i>ℹ</i><div><b>' + f.vang + ' tài khoản đang hoạt động không có trong file</b><small>App không tự khóa hay xóa họ. Anh tự xử lý ở bảng Tài khoản bên dưới: Khóa nếu đã có lịch sử làm việc, Xóa nếu nhập sai hoặc chưa làm gì.</small><details style="margin-top:6px"><summary style="cursor:pointer;font-size:13px">Xem danh sách</summary><div style="font-size:13px;line-height:1.6;margin-top:4px">' + f.vangList.slice(0, 80).map(function (a) { return esc(a.ho_ten) + ' · ' + esc(VT[a.vai_tro] || a.vai_tro) + ' · ' + esc(shopName(a.id_shop)) }).join('<br>') + (f.vangList.length > 80 ? '<br>…' : '') + '</div></details></div></div>';
      h += '<div class="tw" style="margin-top:8px"><table><thead><tr><th>Dòng</th><th>Kết quả</th><th>Họ tên</th><th>Email</th><th>Chức danh</th><th>Shop</th><th>Ghi chú</th></tr></thead><tbody>' + f.rows.map(function (o) { var k = KQ[o.kq]; return '<tr><td>' + o.dong + '</td><td>' + tag(k[0], k[1]) + '</td><td><b>' + esc(o.ho_ten) + '</b></td><td>' + esc(o.email) + '</td><td>' + esc(o.chuc_danh) + '</td><td>' + esc((o.ten_shop || o.shopTxt || '').slice(0, 44)) + '</td><td class="meta" style="white-space:normal">' + esc(o.ghi || '') + '</td></tr>' }).join('') + '</tbody></table></div>';
      h += '<div class="acts" style="margin-top:10px"><button class="btn b1" id="hr-go"' + (todo && !H.busy ? '' : ' disabled') + '>' + (H.busy ? 'Đang áp dụng…' : 'Áp dụng ' + todo + ' thay đổi (' + c.moi + ' mới, ' + c.sua + ' cập nhật' + (f.newShops.length && H.createShops ? ', ' + f.newShops.length + ' Shop mới' : '') + ')') + '</button><button class="btn b3" id="hr-drop-x">Bỏ file này</button></div>'
    }
    if (H.res) {
      var ok = H.res.filter(function (x) { return x.ok }), bad = H.res.filter(function (x) { return !x.ok }), made = ok.filter(function (x) { return x.mk });
      h += '<div class="alert" style="margin-top:10px"><i>✓</i><div><b>Đã xong: ' + ok.length + ' thành công' + (bad.length ? ', ' + bad.length + ' lỗi' : '') + '</b><small>' + made.length + ' tài khoản mới. Mật khẩu ban đầu chỉ hiện ở đây, tải file đăng nhập để gửi nhân viên và nhắc họ đổi mật khẩu.</small></div>' + (made.length ? '<button class="btn b1" id="hr-csv" style="min-height:40px;margin-left:auto">' + ic('download') + 'Tải file đăng nhập (CSV)</button>' : '') + '</div>';
      if (bad.length) h += '<div class="err">' + bad.slice(0, 8).map(function (x) { return 'Dòng ' + x.dong + ' · ' + esc(x.ho_ten) + ': ' + esc(x.msg) }).join('<br>') + '</div>'
    }
    return h + '</div>'
  }
  function apply() {
    var f = H.info, todo = f.rows.filter(function (o) { return o.kq === 'moi' || o.kq === 'sua' }); if (!todo.length || H.busy) return; if (!confirm('Áp dụng ' + todo.length + ' thay đổi nhân sự' + (f.newShops.length && H.createShops ? ' và tạo ' + f.newShops.length + ' Shop mới' : '') + '?')) return; H.busy = true; H.res = []; render();
    var shopsToMake = H.createShops ? f.newShops.filter(function (s) { return todo.some(function (o) { return o.shopMoi && o.id_shop === 'S' + s.ma }) }) : [];
    shopsToMake.reduce(function (p, s) { return p.then(function () { return rpc('u_asm_shop_save', { p_ma_shop: s.ma, p_ten_shop: s.ten, p_req: rid() }) }) }, Promise.resolve()).then(function () {
      return todo.reduce(function (p, o) {
        return p.then(function () {
          var req = o.kq === 'moi' ? rpc('u_asm_account_create', { p_ho_ten: o.ho_ten, p_email: o.email, p_vai_tro: o.vai_tro, p_id_shop: o.id_shop, p_mat_khau: o.mk, p_thay_ql: false, p_req: rid(), p_sdt: o.sdt || null }) : rpc('u_asm_account_update', { p_ma_tk: o.ma_tk, p_ho_ten: o.ho_ten, p_vai_tro: o.vai_tro, p_id_shop: o.id_shop, p_trang_thai: (D.acc[o.ma_tk] || {}).trang_thai || 'HOAT_DONG', p_thay_ql: false, p_req: rid(), p_sdt: o.addSdt ? o.sdt : null });
          return req.then(function () { H.res.push({ ok: true, dong: o.dong, ho_ten: o.ho_ten, email: o.email, mk: o.kq === 'moi' ? o.mk : '', vai_tro: o.vai_tro, id_shop: o.id_shop }) }, function (e) { H.res.push({ ok: false, dong: o.dong, ho_ten: o.ho_ten, msg: (e && e.message) || (e && e.code) || 'Lỗi' }) })
        })
      }, Promise.resolve())
    }).then(function () { H.busy = false; H.info = null; H.aoa = null; toast('Đã áp dụng danh sách nhân sự'); return load(true).then(render) }, function (e) { H.busy = false; render(); fail(e) })
  }
  function csv() {
    var shopN = function (id) { return (D.shops[id] || {}).ten_shop || id }, rows = H.res.filter(function (x) { return x.ok && x.mk }), t = 'ho_ten,email,mat_khau_ban_dau,vai_tro,shop\n' + rows.map(function (x) { return [x.ho_ten, x.email, x.mk, VT[x.vai_tro], shopN(x.id_shop)].map(function (v) { v = String(v == null ? '' : v); return /[",\n;]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v }).join(',') }).join('\n');
    var a = document.createElement('a'); a.href = URL.createObjectURL(new Blob(['﻿' + t], { type: 'text/csv;charset=utf-8' })); a.download = 'dang-nhap-nhan-su-' + new Date().toISOString().slice(0, 10) + '.csv'; document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove() }, 500)
  }
  function read(f) { var rd = new FileReader(); rd.onload = function () { handle(f.name, rd.result) }; rd.readAsArrayBuffer(f) }
  function bind(on) {
    xlsx(function () { });
    var dz = document.getElementById('hr-drop');
    if (dz) { ['dragenter', 'dragover'].forEach(function (ev) { dz.addEventListener(ev, function (e) { e.preventDefault(); if (!H.drag) { H.drag = true; dz.style.borderColor = 'var(--teal)'; dz.style.background = 'var(--tint)' } }) }); ['dragleave', 'drop'].forEach(function (ev) { dz.addEventListener(ev, function (e) { e.preventDefault(); H.drag = false; dz.style.borderColor = 'var(--line)'; dz.style.background = 'var(--bg)' }) }); dz.addEventListener('drop', function (e) { var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]; if (f) read(f) }) }
    on('#hr-file', 'onchange', function (x) { var f = x.files[0]; if (f) read(f) });
    on('#hr-role', 'onchange', function (x) { H.updRole = x.checked; redo(); render() }); on('#hr-sdt', 'onchange', function (x) { H.getSdt = x.checked; redo(); render() }); on('#hr-act', 'onchange', function (x) { H.onlyActive = x.checked; redo(); render() }); on('#hr-shop', 'onchange', function (x) { H.createShops = x.checked; redo(); render() });
    [].slice.call(document.querySelectorAll('[data-hrt]')).forEach(function (s) { s.onchange = function () { var t = H.info.titles[Number(s.dataset.hrt)]; H.rules[t.t] = s.value; redo(); render() } });
    on('#hr-go', 'onclick', apply); on('#hr-csv', 'onclick', csv); on('#hr-drop-x', 'onclick', function () { H.info = null; H.aoa = null; H.res = null; render() })
  }
  if (typeof window !== 'undefined' && window.addEventListener) { window.addEventListener('dragover', function (e) { e.preventDefault() }); window.addEventListener('drop', function (e) { e.preventDefault() }) }
  return { card: card, bind: bind, handle: handle, parse: parse, findHdr: findHdr, state: H }
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
    h += '<div class="top tbar" style="flex-wrap:wrap"><div class="search">' + ic('search') + '<input id="im-q" placeholder="Tìm theo tên, chủ đề…" value="' + esc(IM.q) + '"></div><label class="row" style="cursor:pointer"><input type="checkbox" id="im-hid" ' + (IM.hidden ? 'checked' : '') + ' style="width:20px;height:20px;accent-color:var(--teal)"><span class="meta">Hiện cả ảnh đã ẩn</span></label></div>';
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
  var AS = { tk: { shop: '', st: 'HOAT_DONG', q: '', nosdt: false }, tl: { tab: 'kho', nhom: '', st: 'ACTIVE', q: '' }, pd: { tv: '3', tu: '30', tg: '' } };
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
  function fph(s) { s = String(s || ''); return /^0\d{9}$/.test(s) ? s.slice(0, 4) + ' ' + s.slice(4, 7) + ' ' + s.slice(7) : s }
  function qlList(shopId) { return accs().filter(function (a) { return a.vai_tro === 'QUAN_LY' && a.id_shop === shopId && a.trang_thai === 'HOAT_DONG' }) }
  function fdt(x) { if (!x) return ''; var d = new Date(x); return isNaN(d) ? '' : d.toLocaleDateString('vi-VN') + ' ' + d.toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit' }) }

  /* ================= NHÂN SỰ & TÀI KHOẢN ================= */
  function tkRow(a) {
    var n = openN(a.ma_tk), bs = 'style="min-height:36px;font-size:13px"';
    return '<tr><td class="rh"><b>' + esc(a.ho_ten) + '</b><small>' + esc(a.email) + '</small></td><td class="nl">' + (a.vai_tro === 'QUAN_LY' ? tag('t-today', 'Quản lý') : 'Nhân viên') + '</td><td class="rp" data-label="SĐT">' + (a.sdt ? esc(fph(a.sdt)) : '<span class="tag t-late">Chưa có</span>') + '</td><td data-label="Đang giữ">' + (n ? n + ' lượt mở' : '–') + '</td><td class="nl">' + (a.trang_thai === 'HOAT_DONG' ? tag('t-ok', 'Hoạt động') : tag('t-bad', 'Đã khóa')) + '</td><td class="nl mk">' + (a.mk_doi_luc ? tag('t-ok', 'Đã đổi') + '<small>' + esc(fdt(a.mk_doi_luc)) + '</small>' : '<span class="tag t-late">Chưa đổi</span>') + '</td><td data-label="Đăng nhập gần nhất">' + (a.dang_nhap_luc ? esc(fdt(a.dang_nhap_luc)) : '<span class="meta">Chưa đăng nhập</span>') + '</td><td class="ra"><div class="row" style="flex-wrap:wrap"><button class="btn b3" data-tkedit="' + esc(a.ma_tk) + '" ' + bs + '>Sửa</button><button class="btn b3" data-tkpw="' + esc(a.ma_tk) + '" ' + bs + '>Mật khẩu</button><button class="btn b3" data-tklock="' + esc(a.ma_tk) + '" ' + bs + '>' + (a.trang_thai === 'HOAT_DONG' ? 'Khóa' : 'Mở khóa') + '</button><button class="btn b3" data-tkdel="' + esc(a.ma_tk) + '" style="min-height:36px;font-size:13px;color:var(--bad)">Xóa</button></div></td></tr>'
  }
  function pgTk() {
    var f = AS.tk, SH = U.shops(), ql = f.q.trim().toLowerCase(); f.open = f.open || {};
    var L = accs().filter(function (a) { return (!f.shop || a.id_shop === f.shop) && (!f.st || a.trang_thai === f.st) && (!f.nosdt || !a.sdt) && (!ql || (a.ho_ten + ' ' + a.email).toLowerCase().indexOf(ql) >= 0) })
      .sort(function (a, b) { return (a.vai_tro === b.vai_tro ? a.ho_ten.localeCompare(b.ho_ten) : a.vai_tro === 'QUAN_LY' ? -1 : 1) });
    var cnt = { '': accs().length, HOAT_DONG: 0, NGUNG: 0 }; accs().forEach(function (a) { cnt[a.trang_thai] = (cnt[a.trang_thai] || 0) + 1 });
    var h = head('Nhân sự & tài khoản', 'Tạo tài khoản, đổi Shop, chọn Quản lý cho từng Shop (một Shop có thể có nhiều Quản lý), khóa hoặc mở, xóa, đặt lại mật khẩu. Chỉ ASM làm được.', '<button class="btn b1" id="tk-new" aria-label="Thêm tài khoản">+<span class="bt"> Thêm tài khoản</span></button>');
    h += '<div class="tbar2"><div class="search">' + ic('search') + '<input id="tk-q" placeholder="Tìm tên, email…" value="' + esc(f.q) + '"></div>' + sel('tk-shop', [['', 'Tất cả Shop']].concat(SH.map(function (s) { return [s.id_shop, s.ten_shop] })), f.shop) + '</div>';
    var noS = accs().filter(function (a) { return a.trang_thai === 'HOAT_DONG' && !a.sdt }).length;
    h += '<div class="chips">' + [['HOAT_DONG', 'Đang hoạt động'], ['NGUNG', 'Đã khóa'], ['', 'Tất cả']].map(function (x) { return '<button class="chip ' + (f.st === x[0] ? 'on' : '') + '" data-tkst="' + x[0] + '">' + x[1] + '<em>' + (cnt[x[0]] || 0) + '</em></button>' }).join('') + '<button class="chip ' + (f.nosdt ? 'on' : '') + '" data-tknosdt="1" title="Người chưa có SĐT thì tin nhắn của họ thiếu số liên hệ">Chưa có SĐT<em>' + noS + '</em></button></div>';
    var hc = window.ASMHR ? ASMHR.card() : '';
    h += '<details class="hrbox" id="hrbox"' + (f.hrOpen || /id="hr-go"/.test(hc) ? ' open' : '') + '><summary>' + ic('upload') + 'Nạp danh sách nhân sự từ file Excel/CSV</summary>' + hc + '</details>';
    var byShop = {}; L.forEach(function (a) { (byShop[a.id_shop] = byShop[a.id_shop] || []).push(a) });
    var filt = !!(f.q || f.nosdt), shown = SH.filter(function (s) { return (!f.shop || s.id_shop === f.shop) && (byShop[s.id_shop] || (!f.q && !f.nosdt)) });
    h += '<div class="shgs">' + shown.map(function (s) {
      var ms = qlList(s.id_shop), ps = byShop[s.id_shop] || [], o = f.open[s.id_shop], open = o != null ? o : ((filt && ps.length > 0) || f.shop === s.id_shop);
      return '<details class="shg" data-s="' + esc(s.id_shop) + '"' + (open ? ' open' : '') + '><summary><span class="shn1"><b>' + esc(s.ten_shop) + '</b><small>' + (ms.length ? 'Quản lý: ' + ms.map(function (m) { return esc(m.ho_ten) }).join(', ') : '<span class="tag t-late">Chưa có Quản lý</span>') + '</small></span><span class="shcnt">' + ps.length + ' người</span></summary><div class="shb"><div class="shbt"><button class="btn b3" data-setql="' + esc(s.id_shop) + '" style="min-height:36px;font-size:13px">Chọn Quản lý</button></div>' +
        (ps.length ? '<div class="tw"><table class="rt rta" data-nopg><thead><tr><th>Họ tên</th><th>Vai trò</th><th>SĐT</th><th>Đang giữ</th><th>Trạng thái</th><th>Mật khẩu</th><th>Đăng nhập gần nhất</th><th style="min-width:270px"></th></tr></thead><tbody>' + ps.map(tkRow).join('') + '</tbody></table></div>' : '<div class="meta" style="padding:6px 2px 10px">Không có tài khoản phù hợp với bộ lọc.</div>') + '</div></details>'
    }).join('') + (shown.length ? '' : '<div class="card"><div class="meta">Không có Shop hoặc tài khoản phù hợp.</div></div>') + '</div>';
    return h
  }
  function creds(title, email, pw) {
    var d = modal(), url = location.origin + location.pathname, txt = 'Đăng nhập ' + BRAND.app + '\nLink: ' + url + '\nEmail: ' + email + '\nMật khẩu: ' + pw;
    sheet(d, '<h2 style="font-size:18px;margin:0 0 6px">' + esc(title) + '</h2><div class="meta" style="margin-bottom:10px">Mật khẩu chỉ hiện một lần ở đây. Gửi cho nhân viên và nhắc đổi mật khẩu ở mục Tài khoản.</div><div class="alert" style="display:block;font-size:15px;line-height:1.7"><div>Email: <b>' + esc(email) + '</b></div><div>Mật khẩu: <b style="font-size:18px;letter-spacing:.5px">' + esc(pw) + '</b></div></div><div class="acts" style="margin-top:8px"><button class="btn b1" id="cc">Chép thông tin đăng nhập</button><button class="btn b3" id="cx">Đóng</button></div>', 480);
    q(d, '#cc').onclick = function () { copy(txt); toast('Đã chép thông tin đăng nhập') }; q(d, '#cx').onclick = function () { d.remove() }
  }
  function accModal(a) {
    var isNew = !a, SH = U.shops(), m = { ho_ten: a ? a.ho_ten : '', email: a ? a.email : '', vai_tro: a ? a.vai_tro : 'NHAN_VIEN', id_shop: a ? a.id_shop : (ST.shop || (SH[0] || {}).id_shop || ''), trang_thai: a ? a.trang_thai : 'HOAT_DONG', mk: isNew ? genPw() : '', thay: false, sdt: a ? a.sdt || '' : '' }, d = modal();
    function draw() {
      sheet(d, '<h2 style="font-size:18px;margin:0 0 10px">' + (isNew ? 'Thêm tài khoản' : 'Sửa tài khoản') + '</h2><div class="form2" style="margin-top:0"><div><label class="meta">Họ tên</label><input id="f-ten" value="' + esc(m.ho_ten) + '"></div><div><label class="meta">Email đăng nhập</label><input id="f-em" type="email" value="' + esc(m.email) + '"' + (isNew ? '' : ' disabled') + '></div><div><label class="meta">Vai trò</label>' + sel('f-vt', [['NHAN_VIEN', 'Nhân viên'], ['QUAN_LY', 'Quản lý']], m.vai_tro) + '</div><div><label class="meta">Shop</label>' + sel('f-sh', SH.map(function (s) { return [s.id_shop, s.ten_shop] }), m.id_shop) + '</div><div style="grid-column:1/-1"><label class="meta">SĐT nhân viên (hiện trong chữ ký mọi tin nhắn gửi khách, 10 số)</label><input id="f-sdt" type="tel" inputmode="tel" placeholder="Ví dụ 0912345678" value="' + esc(m.sdt) + '" autocomplete="off"></div>' +
        (isNew ? '<div style="grid-column:1/-1"><label class="meta">Mật khẩu ban đầu (tối thiểu 8 ký tự)</label><div class="row"><input id="f-mk" value="' + esc(m.mk) + '" autocomplete="off"><button class="btn b3" id="f-gen" style="min-height:44px">Tạo ngẫu nhiên</button></div></div>' : '<div><label class="meta">Trạng thái</label>' + sel('f-tt', [['HOAT_DONG', 'Đang hoạt động'], ['NGUNG', 'Đã khóa']], m.trang_thai) + '</div>') + '</div>' +
        (!isNew && a.vai_tro !== m.vai_tro || !isNew && a.id_shop !== m.id_shop ? '<div class="ro" style="margin-top:10px">Lượt đang giữ vẫn giữ tên người này và vẫn ghi Shop lúc giao. Quản lý cũ của Shop sẽ không còn xem được nếu người này chuyển Shop.</div>' : '') +
        '<div class="acts" style="margin-top:12px"><button class="btn b3" id="f-x" style="flex:1">Hủy</button><button class="btn b1" id="f-ok" style="flex:1.4">' + (isNew ? 'Tạo tài khoản' : 'Lưu') + '</button></div>');
      q(d, '#f-ten').oninput = function (e) { m.ho_ten = e.target.value }; q(d, '#f-sdt').oninput = function (e) { m.sdt = e.target.value }; var em = q(d, '#f-em'); em.oninput = function (e) { m.email = e.target.value };
      q(d, '#f-vt').onchange = function (e) { m.vai_tro = e.target.value; m.thay = false; m.ho_ten = q(d, '#f-ten').value; m.email = em.value; draw() }; q(d, '#f-sh').onchange = function (e) { m.id_shop = e.target.value; m.thay = false; m.ho_ten = q(d, '#f-ten').value; m.email = em.value; draw() };
      var tt = q(d, '#f-tt'); if (tt) tt.onchange = function (e) { m.trang_thai = e.target.value; m.ho_ten = q(d, '#f-ten').value; draw() };
      var mk = q(d, '#f-mk'); if (mk) { mk.oninput = function (e) { m.mk = e.target.value }; q(d, '#f-gen').onclick = function () { m.mk = genPw(); mk.value = m.mk } }
      q(d, '#f-x').onclick = function () { d.remove() };
      q(d, '#f-ok').onclick = function () {
        m.ho_ten = q(d, '#f-ten').value; m.email = em.value; if (mk) m.mk = mk.value; var b = q(d, '#f-ok'); b.disabled = true;
        var p = isNew ? rpc('u_asm_account_create', { p_ho_ten: m.ho_ten, p_email: m.email, p_vai_tro: m.vai_tro, p_id_shop: m.id_shop, p_mat_khau: m.mk, p_thay_ql: m.thay, p_req: rid(), p_sdt: m.sdt || null }) : rpc('u_asm_account_update', { p_ma_tk: a.ma_tk, p_ho_ten: m.ho_ten, p_vai_tro: m.vai_tro, p_id_shop: m.id_shop, p_trang_thai: m.trang_thai, p_thay_ql: m.thay, p_req: rid(), p_sdt: m.sdt || '' });
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
    var people = accs().filter(function (a) { return a.id_shop === shopId && a.trang_thai === 'HOAT_DONG' }).sort(function (a, b) { return a.vai_tro === b.vai_tro ? a.ho_ten.localeCompare(b.ho_ten) : a.vai_tro === 'QUAN_LY' ? -1 : 1 }), d = modal();
    sheet(d, '<h2 style="font-size:18px;margin:0 0 6px">Chọn Quản lý · ' + esc(shopName(shopId)) + '</h2><div class="meta" style="margin-bottom:10px">Tick những người làm Quản lý của Shop này. Một Shop có thể có nhiều Quản lý. Bỏ tick = chuyển người đó về Nhân viên.</div>' + (people.length ? '<div style="display:grid;gap:6px;margin-bottom:10px">' + people.map(function (a) { return '<label class="row" style="cursor:pointer"><input type="checkbox" data-qw="' + esc(a.ma_tk) + '"' + (a.vai_tro === 'QUAN_LY' ? ' checked' : '') + '><span><b>' + esc(a.ho_ten) + '</b><small style="display:block">' + esc(a.email) + '</small></span></label>' }).join('') + '</div>' : '<div class="ro">Shop này chưa có ai. Hãy thêm tài khoản trước.</div>') + '<div class="acts"><button class="btn b3" id="q-x" style="flex:1">Hủy</button>' + (people.length ? '<button class="btn b1" id="q-ok" style="flex:1.4">Lưu</button>' : '') + '</div>', 480);
    q(d, '#q-x').onclick = function () { d.remove() };
    var ok = q(d, '#q-ok'); if (ok) ok.onclick = function () {
      var chg = people.map(function (a) { return { a: a, on: q(d, '[data-qw="' + a.ma_tk + '"]').checked } }).filter(function (x) { return x.on !== (x.a.vai_tro === 'QUAN_LY') });
      if (!chg.length) { d.remove(); return } ok.disabled = true;
      chg.reduce(function (p, x) { return p.then(function () { return rpc('u_asm_account_update', { p_ma_tk: x.a.ma_tk, p_ho_ten: x.a.ho_ten, p_vai_tro: x.on ? 'QUAN_LY' : 'NHAN_VIEN', p_id_shop: x.a.id_shop, p_trang_thai: 'HOAT_DONG', p_thay_ql: false, p_req: rid() }) }) }, Promise.resolve())
        .then(function () { d.remove(); done('Đã cập nhật Quản lý của ' + shopName(shopId)) }, function (e) { ok.disabled = false; fail(e) })
    }
  }
  function delAsm(a) {
    var n = openN(a.ma_tk), d = modal();
    sheet(d, '<h2 style="font-size:18px;margin:0 0 6px">Xóa tài khoản</h2><div class="meta" style="margin-bottom:10px">' + esc(a.ho_ten) + ' · ' + esc(a.email) + '</div><div class="ro">Xóa hẳn tài khoản này, <b>không hoàn tác được</b>. Chỉ xóa được khi người này chưa có lịch sử làm việc (chưa được giao lượt, chưa liên hệ khách) — ví dụ nhập sai hoặc nghỉ trước khi làm gì. Nếu đã có lịch sử, hệ thống từ chối xóa và anh dùng <b>Khóa</b> để giữ lịch sử chăm sóc khách.' + (n ? ' <b>Người này đang giữ ' + n + ' lượt mở.</b>' : '') + '</div><div class="acts" style="margin-top:10px"><button class="btn b3" id="x-x" style="flex:1">Hủy</button><button class="btn b1" id="x-ok" style="flex:1.4;background:var(--bad)">Xóa hẳn</button></div>', 480);
    q(d, '#x-x').onclick = function () { d.remove() };
    q(d, '#x-ok').onclick = function () { var b = q(d, '#x-ok'); b.disabled = true; rpc('u_asm_account_delete', { p_ma_tk: a.ma_tk, p_req: rid() }).then(function () { d.remove(); done('Đã xóa ' + a.ho_ten) }, function (e) { b.disabled = false; fail(e) }) }
  }

  /* ================= THƯ VIỆN ================= */
  function used(ma) { return (D.mlog || []).filter(function (x) { return x.ma_mau === ma }).length }
  function pgTl() {
    var f = AS.tl, h = head('Thư viện', f.tab === 'kho' ? 'Nội dung nhân viên dùng khi nhắn và gọi. Nạp từ file Excel của công ty: xem trước, lưu chờ duyệt, duyệt thì nhân viên mới thấy. Phác đồ (khi nào tiêm gì) ở mục riêng, sẽ bổ sung sau.' : f.tab === 'anh' ? 'Ảnh, tờ rơi, voucher: kéo ảnh vào, app tự nhận diện để nhân viên dùng ở nút Gợi ý tin nhắn.' : f.tab === 'quet' ? 'Rà số tiền và CTKM trong các mẫu tin so với Bảng giá và CTKM hiện hành, thay tự động hoặc sửa tay.' : 'Mẫu soạn tay (tạm): hiện chỉ còn dùng cho chế độ "Theo quy trình" (nhóm KH mới, 4 mốc D, D+1, D+3, D+7 và Trễ lịch) và lời nói khi gọi. Các nhóm khác (CSKH sau tiêm, Bầu–Nhi, Chung, Nguy cơ) đã chuyển sang Kho Excel, nhân viên không còn thấy ở đây.', f.tab === 'mau' ? '<button class="btn b1" id="tl-new">+ Thêm mẫu</button>' : f.tab === 'km' ? '<button class="btn b1" id="km-new">+ Thêm khuyến mãi</button>' : '');
    h += '<div class="chips"><button class="chip ' + (f.tab === 'kho' ? 'on' : '') + '" data-tltab="kho">Kho từ Excel</button><button class="chip ' + (f.tab === 'quet' ? 'on' : '') + '" data-tltab="quet">Quét giá &amp; CTKM</button><button class="chip ' + (f.tab === 'anh' ? 'on' : '') + '" data-tltab="anh">Ảnh, tờ rơi, voucher</button><button class="chip ' + (f.tab === 'mau' ? 'on' : '') + '" data-tltab="mau">Mẫu soạn tay (tạm)<em>' + D.tpl.length + '</em></button><button class="chip ' + (f.tab === 'km' ? 'on' : '') + '" data-tltab="km">Khuyến mãi soạn tay (tạm)<em>' + D.promos.length + '</em></button></div>';
    return h + (f.tab === 'kho' ? ASMLIB.page() : f.tab === 'quet' ? ASMSCAN.page() : f.tab === 'anh' ? ASMIMG.page() : f.tab === 'mau' ? tplList() : promoList())
  }
  function tplList() {
    var f = AS.tl, ql = f.q.trim().toLowerCase(), all = D.tpl;
    var L = all.filter(function (x) { return (!f.nhom || x.nhom === f.nhom) && (f.st === '' || (f.st === 'ACTIVE' ? (x.trang_thai === 'DANG_DUNG' || x.trang_thai === 'CHO_DUYET') : x.trang_thai === f.st)) && (!ql || (x.tieu_de + ' ' + x.noi_dung).toLowerCase().indexOf(ql) >= 0) });
    var h = '<div class="top tbar" style="flex-wrap:wrap"><div class="search">' + ic('search') + '<input id="tl-q" placeholder="Tìm trong mẫu…" value="' + esc(f.q) + '"></div>' + sel('tl-st', [['ACTIVE', 'Đang dùng và chờ duyệt'], ['DANG_DUNG', 'Đang dùng'], ['CHO_DUYET', 'Chờ duyệt'], ['DA_THAY', 'Đã thay'], ['NGUNG_DUNG', 'Ngừng dùng'], ['', 'Tất cả']], f.st) + '</div>';
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
    var LB = MSG._t.LIB; if (!LB.rows && !LB.err && !AS.pdAsk) { AS.pdAsk = 1; MSG.ensureLib(function () { if (S.page === 'a_pd') render() }) }
    var pi = MSG.pdInfo(), t = AS.pd, days = Math.round(Number(t.tv || 0) * Number(t.tu)), due = MSG.dueRules(days, t.tg);
    var h = head('Phác đồ vaccine', 'Quy tắc "khi nào": vaccine nào đến tuổi nào, tính từ ngày sinh. App dùng để gợi ý "vaccine đã đến tuổi" trong tin nhắn của nhân viên. Dòng chỉ hiện cho nhân viên khi đang dùng.', '<button class="btn b1" id="pd-new">+ Thêm dòng phác đồ</button>');
    h += '<div class="ro"><b>Phác đồ đầy đủ nạp ở Thư viện → "Phác đồ &amp; từ điển vaccine"</b> (kéo file Excel: xem trước, chờ duyệt, bác sĩ duyệt). Từ đó nhân viên có nút <b>Từ điển vaccine</b> và khối "vaccine bé đến tuổi" trong tin nhắn. Bảng nhập tay bên dưới là phác đồ mẫu cũ, chỉ dùng khi chưa nạp phác đồ Excel.<div style="margin-top:10px"><button class="btn b1" id="pd-up">Nạp phác đồ đầy đủ từ Excel →</button></div></div>' + (pi.on ? '<div class="ro" style="background:var(--ok-bg);color:var(--ok)"><b>Đang dùng phác đồ từ Excel</b> (' + pi.n + ' vaccine, ' + pi.groups.length + ' nhóm). Mục "vaccine bé đến tuổi" trong tin nhắn lấy từ đây. Nhóm không tự gợi ý theo tuổi (vẫn tra cứu được trong Từ điển vaccine): <b>' + esc(pi.groups.filter(function (g) { return !g.auto }).map(function (g) { return g.name }).join(', ') || 'không có') + '</b>. Bảng nhập tay bên dưới là phác đồ mẫu cũ, không còn được dùng.</div>' : '<div class="ro" style="background:var(--warn-bg);color:var(--warn)"><b>Dữ liệu mẫu minh họa.</b> Các độ tuổi bên dưới chưa phải phác đồ chuẩn đã duyệt. Bác sĩ cần nhập và duyệt phác đồ thật trước khi dùng cho khách thật; dòng "Chờ duyệt" không hiện cho nhân viên.</div>') + '';
    h += '<div class="card"><h2>Thử nhanh: bé ở tuổi này thì gợi ý những vaccine nào?</h2><div class="row" style="flex-wrap:wrap"><input id="pd-tv" type="number" min="0" value="' + esc(t.tv) + '" style="width:90px;border:1px solid var(--line);border-radius:12px;padding:10px;min-height:44px;background:var(--card)">' + sel('pd-tu', [['1', 'ngày'], ['7', 'tuần'], ['30', 'tháng'], ['365', 'năm']], t.tu) + sel('pd-tg', [['', 'Cả nam và nữ'], ['M', 'Nam'], ['F', 'Nữ']], t.tg) + '<span class="meta">= ' + days + ' ngày tuổi</span></div><div style="margin-top:10px">' + (due.length ? due.map(function (r) { return tag('t-today', r.ten) }).join('') : '<span class="meta">Chưa có vaccine nào đang dùng ở tuổi này.</span>') + '</div></div>';
    h += '<div class="card" style="padding:6px 8px"><div class="tw"><table><thead><tr><th>Vaccine</th><th>Từ tuổi</th><th>Đến tuổi</th><th>Giới tính</th><th>Phác đồ (số mũi, khoảng cách)</th><th>Ghi chú cho nhân viên</th><th>Trạng thái</th><th style="min-width:170px"></th></tr></thead><tbody>' + D.rules.map(function (r) {
      var b = function (k, tx) { return '<button class="btn b3" data-pd' + k + '="' + esc(r.ma) + '" style="min-height:36px;font-size:13px">' + tx + '</button>' };
      return '<tr><td><b>' + esc(r.ten) + '</b><small>' + esc(r.ma_vaccine) + ' · ' + esc(r.ma) + (r.tu_khoa ? ' · từ khóa: ' + esc(r.tu_khoa) : '') + '</small></td><td>' + esc(MSG.dayLabel(r.tu_ngay)) + '<small>' + r.tu_ngay + ' ngày</small></td><td>' + (r.den_ngay == null ? 'Không giới hạn' : esc(MSG.dayLabel(r.den_ngay)) + '<small>' + r.den_ngay + ' ngày</small>') + '</td><td>' + (r.gioi_tinh === 'M' ? 'Nam' : r.gioi_tinh === 'F' ? 'Nữ' : 'Cả hai') + '</td><td style="max-width:220px;white-space:normal">' + esc(r.phac_do || '–') + '</td><td style="max-width:220px;white-space:normal">' + esc(r.cau_goi_y || '–') + '</td><td>' + (r.trang_thai === 'DANG_DUNG' ? tag('t-ok', 'Đang dùng') : r.trang_thai === 'CHO_DUYET' ? tag('t-late', 'Chờ duyệt') : tag('t-bad', 'Ngừng')) + '</td><td><div class="row" style="flex-wrap:wrap">' + b('edit', 'Sửa') + (r.trang_thai === 'CHO_DUYET' ? b('ok', 'Duyệt') : r.trang_thai === 'DANG_DUNG' ? b('off', 'Ngừng') : b('ok', 'Bật lại')) + '</div></td></tr>'
    }).join('') + '</tbody></table></div></div>';
    h += linkCard();
    return h
  }
  /* Kiểm tra liên kết: với từng vaccine trong phác đồ, kho đang có bao nhiêu mẫu để app TỰ CHỌN SẴN khi nhân viên chọn vaccine đó */
  function linkCard() {
    var lib = MSG._t.LIB; if (!lib.rows) { if (!AS.pd.ask) { AS.pd.ask = true; MSG.ensureLib(render) } return '<div class="card"><h2>Liên kết với kho</h2><div class="meta">Đang tải kho mẫu…</div></div>' } MSG.ensureLib();
    var rep = MSG.linkReport(), z = function (n) { return n ? '<span class="tag t-ok">' + n + '</span>' : '<span class="tag t-bad">chưa có</span>' }, bad = rep.filter(function (x) { return !x.tv || !x.nc || !x.km }).length;
    return '<div class="card" style="padding:6px 8px"><h2 style="padding:10px 10px 0">Liên kết với kho <span class="pill">app tự chọn sẵn khi nhân viên chọn vaccine</span></h2><div class="meta" style="padding:0 10px 8px">Với mỗi vaccine trong phác đồ, app tìm trong kho (đang dùng) câu tư vấn, nguy cơ & đường lây, khuyến mãi còn hạn và mẫu sau tiêm của đúng vaccine đó. ' + (bad ? '<b>' + bad + ' vaccine còn thiếu ít nhất một mục</b>: thêm từ khóa ở dòng phác đồ, hoặc bổ sung mẫu vào kho.' : 'Tất cả vaccine đều đã có đủ.') + '</div><div class="tw"><table><thead><tr><th>Vaccine</th><th>Nhóm trong Bảng giá</th><th>Tư vấn</th><th>Nguy cơ</th><th>Khuyến mãi</th><th>Sau tiêm</th></tr></thead><tbody>' + rep.map(function (x) {
      return '<tr><td><b>' + esc(x.name) + '</b><small>' + esc(x.rule.ma) + (x.rule.trang_thai !== 'DANG_DUNG' ? ' · ' + esc(x.rule.trang_thai === 'CHO_DUYET' ? 'chờ duyệt' : 'ngừng') : '') + '</small></td><td>' + (x.groups.length ? esc(x.groups.join('; ').slice(0, 90)) + '<small>' + esc(x.keys.slice(0, 6).join(', ')) + '</small>' : '<span class="tag t-late">Không khớp nhóm nào</span>') + '</td><td>' + z(x.tv) + (x.tvT[0] ? '<small>' + esc(x.tvT[0].slice(0, 50)) + '</small>' : '') + '</td><td>' + z(x.nc) + (x.ncT.length ? '<small>' + esc(x.ncT.slice(0, 3).join(', ')) + '</small>' : '') + '</td><td>' + z(x.km) + '</td><td>' + z(x.st) + (x.st ? '' : '<small>sẽ dùng tin chung</small>') + '</td></tr>'
    }).join('') + '</tbody></table></div></div>'
  }
  function ruleSave(r, patch) { var o = { p_ma: r ? r.ma : '', p_ma_vaccine: r.ma_vaccine, p_ten: r.ten, p_tu_ngay: r.tu_ngay, p_den_ngay: r.den_ngay, p_gioi_tinh: r.gioi_tinh || '', p_phac_do: r.phac_do || '', p_cau: r.cau_goi_y || '', p_trang_thai: r.trang_thai, p_req: rid(), p_tu_khoa: r.tu_khoa == null ? null : r.tu_khoa }; for (var k in patch) o[k] = patch[k]; return rpc('u_asm_rule_save', o) }
  function ruleModal(r) {
    var d = modal(), pick = function (v) { return v == null ? ['', '1'] : [String(v), '1'] };
    sheet(d, '<h2 style="font-size:18px;margin:0 0 10px">' + (r ? 'Sửa phác đồ ' + esc(r.ma) : 'Thêm dòng phác đồ') + '</h2><div class="form2" style="margin-top:0"><div><label class="meta">Mã vaccine (ví dụ 6IN1)</label><input id="r-mv" value="' + esc(r ? r.ma_vaccine : '') + '"></div><div><label class="meta">Tên hiển thị</label><input id="r-ten" value="' + esc(r ? r.ten : '') + '"></div>' +
      '<div><label class="meta">Từ tuổi</label><div class="row"><input id="r-tu" type="number" min="0" value="' + esc(r ? r.tu_ngay : 0) + '" style="flex:1">' + sel('r-tuu', [['1', 'ngày'], ['7', 'tuần'], ['30', 'tháng'], ['365', 'năm']], '1') + '</div></div><div><label class="meta">Đến tuổi (trống = không giới hạn)</label><div class="row"><input id="r-den" type="number" min="0" value="' + esc(r && r.den_ngay != null ? r.den_ngay : '') + '" style="flex:1">' + sel('r-denu', [['1', 'ngày'], ['7', 'tuần'], ['30', 'tháng'], ['365', 'năm']], '1') + '</div></div>' +
      '<div><label class="meta">Giới tính</label>' + sel('r-gt', [['', 'Cả nam và nữ'], ['M', 'Chỉ nam'], ['F', 'Chỉ nữ']], r ? (r.gioi_tinh || '') : '') + '</div><div><label class="meta">Trạng thái</label>' + sel('r-tt', [['CHO_DUYET', 'Chờ duyệt'], ['DANG_DUNG', 'Đang dùng'], ['NGUNG_DUNG', 'Ngừng']], r ? r.trang_thai : 'CHO_DUYET') + '</div>' +
      '<div style="grid-column:1/-1"><label class="meta">Phác đồ: số mũi, khoảng cách (bác sĩ nhập)</label><textarea id="r-pd" style="min-height:70px">' + esc(r ? r.phac_do || '' : '') + '</textarea></div><div style="grid-column:1/-1"><label class="meta">Ghi chú cho nhân viên (hiện cạnh vaccine trong gợi ý)</label><input id="r-ct" value="' + esc(r ? r.cau_goi_y || '' : '') + '"></div><div style="grid-column:1/-1"><label class="meta">Từ khóa nhận diện vaccine (tùy chọn, cách nhau bằng dấu phẩy)</label><input id="r-tk" placeholder="Ví dụ: hexaxim, infanrix hexa" value="' + esc(r ? r.tu_khoa || '' : '') + '"><div class="meta" style="margin-top:4px">App tự nối vaccine này với tư vấn, nguy cơ, khuyến mãi, sau tiêm trong kho theo tên vaccine và Bảng giá. Thêm từ khóa (tên thương mại, cách gọi khác) nếu bảng "Liên kết với kho" bên dưới báo thiếu.</div></div></div><div class="meta" style="margin-top:6px">Tuổi nhập bằng ngày, tuần (7 ngày), tháng (30 ngày) hoặc năm (365 ngày); hệ thống lưu theo ngày.</div><div class="acts" style="margin-top:12px"><button class="btn b3" id="r-x" style="flex:1">Hủy</button><button class="btn b1" id="r-ok" style="flex:1.4">Lưu</button></div>', 600);
    q(d, '#r-x').onclick = function () { d.remove() };
    q(d, '#r-ok').onclick = function () {
      var tu = q(d, '#r-tu').value, den = q(d, '#r-den').value, b = q(d, '#r-ok'); if (tu === '') { toast('Cần nhập "Từ tuổi".', 0, 1); return }
      b.disabled = true; ruleSave({ ma: r ? r.ma : '', ma_vaccine: q(d, '#r-mv').value, ten: q(d, '#r-ten').value, tu_ngay: Math.round(Number(tu) * Number(q(d, '#r-tuu').value)), den_ngay: den === '' ? null : Math.round(Number(den) * Number(q(d, '#r-denu').value)), gioi_tinh: q(d, '#r-gt').value, phac_do: q(d, '#r-pd').value, cau_goi_y: q(d, '#r-ct').value, tu_khoa: q(d, '#r-tk').value, trang_thai: q(d, '#r-tt').value }, {}).then(function () { d.remove(); done('Đã lưu phác đồ') }, function (e) { b.disabled = false; fail(e) })
    }
  }


  /* ================= KIỂM TRA GỌI & NHẮN (mức 1: dấu hiệu để quản lý kiểm tra, KHÔNG phải bằng chứng) ================= */
  AS.nv = { tu: '', den: '', rows: null, loading: false, err: '' };
  function isoD(d) { return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2) }
  function nvScore(r) { var n = 0; ['nhanh', 'roi_ngan', 'dot_bien', 'trung_ghi_chu'].forEach(function (k) { if (Number(r[k]) > 0) n += Number(r[k]) }); if (Number(r.ty_le_khong_nghe) >= 70 && r.tong >= 10) n += 1; return n }
  function nvTags(r) {
    var o = []; [['Lưu nhanh', r.nhanh], ['Rời app ngắn', r.roi_ngan], ['Dồn dập', r.dot_bien], ['Trùng ghi chú', r.trung_ghi_chu]].forEach(function (x) { if (Number(x[1]) > 0) o.push(tag('t-late', x[0] + ' ×' + x[1])) });
    if (Number(r.ty_le_khong_nghe) >= 70 && r.tong >= 10) o.push(tag('t-late', 'Không nghe ' + r.ty_le_khong_nghe + '%')); return o.join('')
  }
  function pgNv() {
    var t = AS.nv, today = isoD(new Date()); if (!t.tu) { t.den = today; t.tu = isoD(new Date(Date.now() - 6 * 864e5)) }
    if (!t.rows && !t.loading && !t.err) { t.loading = true; rpc('u_asm_nghi_van', { p_tu: t.tu, p_den: t.den }).then(function (r) { t.rows = r || []; t.loading = false; render() }, function (e) { t.loading = false; t.err = (e && e.message) || 'Không tải được báo cáo.'; render() }) }
    var h = head('Kiểm tra gọi & nhắn', 'Dấu hiệu bất thường trong cách nhân viên ghi nhận kết quả. Công cụ để quản lý kiểm tra, không phải bằng chứng gian lận.', '');
    var cur = function (n) { return t.den === today && t.tu === isoD(new Date(Date.now() - (n - 1) * 864e5)) }, preset = cur(7) || cur(14) || cur(30);
    h += '<div class="chips ov">' + [7, 14, 30].map(function (n) { return '<button class="chip ' + (cur(n) ? 'on' : '') + '" data-nvr="' + n + '">' + n + ' ngày</button>' }).join('') +
      '<details class="xp"><summary class="chip ' + (preset ? '' : 'on') + '" aria-label="Chọn khoảng ngày">' + ic('clock') + (preset ? '' : ' ' + t.tu.slice(8) + '/' + t.tu.slice(5, 7) + '–' + t.den.slice(8) + '/' + t.den.slice(5, 7)) + '</summary><div class="xpm nvr"><label class="field" style="margin:0">Từ ngày<input type="date" id="nv-tu" value="' + esc(t.tu) + '"></label><label class="field" style="margin:0">Đến ngày<input type="date" id="nv-den" value="' + esc(t.den) + '"></label><button class="btn b1" id="nv-go">Xem báo cáo</button></div></details></div>';
    if (t.err) return h + '<div class="ro" style="background:var(--bad-bg);color:var(--bad)">' + esc(t.err) + '</div>';
    if (t.loading || !t.rows) return h + '<div class="card"><div class="meta">Đang tải…</div></div>';
    var rows = t.rows, flagged = rows.filter(function (r) { return nvScore(r) > 0 }).sort(function (a, b) { return nvScore(b) - nvScore(a) || b.tong - a.tong }), sumTong = 0, sumFlag = 0;
    rows.forEach(function (r) { sumTong += Number(r.tong) || 0; sumFlag += nvScore(r) });
    h += '<div class="kp3"><div class="card kp-m' + (flagged.length ? ' bad' : '') + '"><div class="l">Cần xem</div><div class="v">' + flagged.length + '<small> /' + rows.length + '</small></div></div><div class="card kp-m"><div class="l">Lượt ghi</div><div class="v">' + sumTong + '</div></div><div class="card kp-m"><div class="l">Dấu hiệu</div><div class="v">' + sumFlag + '</div></div></div>';
    h += '<div class="card"><h2>Nên xem trước</h2>' + (flagged.length ? '<div class="nvl">' + flagged.slice(0, 5).map(function (r) { return '<div class="nvr"><div class="nvh"><b>' + esc(r.ho_ten) + '</b><span class="meta">' + r.tong + ' lần ghi</span></div><div class="nvt">' + nvTags(r) + '</div></div>' }).join('') + '</div>' + (flagged.length > 5 ? '<div class="meta" style="margin-top:6px">Còn ' + (flagged.length - 5) + ' nhân viên khác trong bảng chi tiết bên dưới.</div>' : '') : '<div class="alert"><i>✓</i><div><b>Chưa thấy dấu hiệu bất thường</b><small>Trong khoảng ngày đang chọn.</small></div></div>') + '</div>';
    var c = function (n, bad) { return Number(n) > 0 ? tag(bad ? 't-late' : 't-new', n) : '<span class="meta">0</span>' }, open = t.open != null ? t.open : window.innerWidth >= 960;
    h += '<details class="card nvd"' + (open ? ' open' : '') + '><summary>Bảng chi tiết từng nhân viên <span class="meta">(' + rows.length + ')</span></summary><div class="tw"><table class="rt rtb" style="--c:4"><thead><tr><th>Nhân viên</th><th>Số lần ghi</th><th>Lưu nhanh<small>dưới 30 giây sau bấm Gọi</small></th><th>Rời app ngắn<small>dưới 5 giây</small></th><th>Dồn dập<small>từ 6 kết quả trong 10 phút</small></th><th>Trùng ghi chú<small>nội dung lặp từ 3 lần</small></th><th>Không nghe, bận<small>% các lần ghi</small></th><th>Ghi cách cũ<small>không có dấu bấm Gọi</small></th></tr></thead><tbody>' +
      (rows.length ? rows.map(function (r) { return '<tr><td class="rh"><b>' + esc(r.ho_ten) + '</b><small>' + esc(r.ma_tk) + '</small></td><td data-label="Lần ghi">' + r.tong + '</td><td data-label="Nhanh">' + c(r.nhanh, 1) + '</td><td data-label="Rời app">' + c(r.roi_ngan, 1) + '</td><td data-label="Dồn dập">' + c(r.dot_bien, 1) + '</td><td data-label="Trùng">' + c(r.trung_ghi_chu, 1) + '</td><td data-label="Không nghe">' + (Number(r.ty_le_khong_nghe) >= 70 && r.tong >= 10 ? tag('t-late', r.ty_le_khong_nghe + '%') : r.ty_le_khong_nghe + '%') + '</td><td data-label="Cách cũ">' + c(r.cach_cu, 0) + '</td></tr>' }).join('') : '<tr><td colspan="8" class="empty meta">Chưa có kết quả nào trong khoảng ngày này.</td></tr>') + '</tbody></table></div></details>';
    h += '<details class="ro nvh"><summary>Cách đọc báo cáo này</summary>Một dòng có nhiều ô cam chưa nói lên điều gì chắc chắn, chỉ nên là lý do để quản lý gọi lại vài khách của nhân viên đó, hoặc đối chiếu với dữ liệu tiêm thực tế. App chỉ biết nhân viên đã bấm nút Gọi hoặc Nhắn, không biết cuộc gọi hay tin nhắn có thật sự xảy ra. Có thể nhân viên gọi bằng máy khác hoặc nhắn Zalo trên máy tính.</details>';
    return h
  }
  X.pages.a_nv = pgNv;
  X.binders.push(function (on) {
    on('#nv-go', 'onclick', function () { var t = AS.nv; t.tu = $('#nv-tu').value; t.den = $('#nv-den').value; if (!t.tu || !t.den || t.tu > t.den) { toast('Chọn khoảng ngày hợp lệ.', 0, 1); return } t.rows = null; t.err = ''; t.loading = false; render() })
  });


  /* ================= CẤU HÌNH GHI NHẬN KẾT QUẢ (ASM tự chỉnh; máy chủ kiểm khoảng giá trị) ================= */
  var CF_KEYS = [
    ['max_vong', 'Số lần xử lý tối đa cho một khách', 'lần, từ 1 đến 10', 1, 10],
    ['max_vong_ngay', 'Số lần xử lý tối đa trong một ngày', 'lần, 0 = không giới hạn', 0, 10],
    ['min_giay_goi', 'Thời gian tối thiểu từ lúc bấm Gọi đến lúc lưu', 'giây, từ 0 đến 600', 0, 600],
    ['bat_buoc_bam_goi', 'Bắt buộc bấm Gọi trong app trước khi lưu', '1 = bắt buộc, 0 = không', 0, 1],
    ['huy_giay', 'Thời gian được hủy lần ghi vừa lưu', 'giây, từ 0 đến 3600', 0, 3600]
  ];
  var CF_FLAGS = [['mien_nhan', 'Miễn nhắn'], ['ghi_chu_bb', 'Bắt buộc ghi chú'], ['can_hen', 'Cần ngày hẹn'], ['can_vaccine', 'Cần vaccine'], ['can_ngay_tiem', 'Cần ngày tiêm'], ['can_noi_tiem', 'Cần nơi tiêm'], ['can_ly_do', 'Cần lý do']];
  var CF_GRP = { chua: 'Chưa liên lạc được', duoc: 'Đã liên lạc được', cuoi: 'Kết thúc lượt' };
  function pgCf() {
    var h = head('Cấu hình ghi nhận', 'Quy tắc khi nhân viên ghi kết quả gọi và nhắn. Thay đổi có hiệu lực ngay cho mọi nhân viên (nhân viên cần tải lại trang).', '');
    h += '<div class="card"><h2>Quy tắc chung</h2>' + CF_KEYS.map(function (k) {
      return '<div class="row" style="gap:10px;align-items:center;flex-wrap:wrap;padding:8px 0;border-bottom:1px solid var(--line)"><div style="flex:1;min-width:220px"><b>' + k[1] + '</b><div class="meta">' + k[2] + '</div></div><input type="number" class="cf-n" data-k="' + k[0] + '" min="' + k[3] + '" max="' + k[4] + '" value="' + cfgN(k[0], 0) + '" style="width:90px"><button class="btn b2 cf-ns" data-k="' + k[0] + '">Lưu</button></div>'
    }).join('') + '</div>';
    ['chua', 'duoc', 'cuoi'].forEach(function (g) {
      var L = D.rr.filter(function (r) { return r.nhom === g }).sort(function (a, b) { return a.thu_tu - b.thu_tu });
      if (!L.length) return;
      h += '<div class="card"><h2>' + CF_GRP[g] + '</h2><div class="meta" style="margin-bottom:6px">Mã, nhóm và việc có kết thúc lượt hay không là cố định, không sửa ở đây.</div>' + L.map(function (r) {
        return '<div class="cf-r" data-ma="' + esc(r.ma) + '" style="padding:10px 0;border-top:1px solid var(--line)"><div class="row" style="gap:8px;align-items:center;flex-wrap:wrap"><input class="cf-ten" value="' + esc(r.ten) + '" maxlength="60" style="flex:1;min-width:180px">' + (r.la_cuoi ? tag('t-today', 'Kết thúc lượt') : '') + '<label class="meta"><input type="checkbox" class="cf-dd"' + (r.dang_dung ? ' checked' : '') + '> Đang dùng</label></div>' +
          '<div class="row" style="gap:6px 14px;flex-wrap:wrap;margin:8px 0">' + CF_FLAGS.map(function (f) { return '<label class="meta"><input type="checkbox" class="cf-f" data-f="' + f[0] + '"' + (r[f[0]] ? ' checked' : '') + '> ' + f[1] + '</label>' }).join('') + '</div>' +
          '<div class="row" style="gap:8px;align-items:center"><input class="cf-gy" placeholder="Gợi ý ghi chú cho nhân viên" maxlength="200" value="' + esc(r.goi_y_ghi_chu || '') + '" style="flex:1"><button class="btn b1 cf-rs">Lưu</button></div></div>'
      }).join('') + '</div>'
    });
    return h
  }
  X.pages.a_cf = pgCf;
  X.binders.push(function () {
    [].forEach.call(document.querySelectorAll('.cf-ns'), function (b) {
      b.onclick = function () {
        var k = b.getAttribute('data-k'), v = document.querySelector('.cf-n[data-k="' + k + '"]').value;
        if (v === '' || isNaN(Number(v))) { toast('Nhập một số.', 0, 1); return }
        b.disabled = true;
        rpc('u_asm_cfg_save', { p_khoa: k, p_gia_tri: parseInt(v, 10) }).then(function () { done('Đã lưu') }, function (e) { b.disabled = false; toast((e && e.message) || 'Không lưu được.', 0, 1) })
      }
    });
    [].forEach.call(document.querySelectorAll('.cf-rs'), function (b) {
      b.onclick = function () {
        var box = b.closest('.cf-r'), f = function (n) { var e = box.querySelector('.cf-f[data-f="' + n + '"]'); return !!(e && e.checked) };
        b.disabled = true;
        rpc('u_asm_rr_save', { p_ma: box.getAttribute('data-ma'), p_ten: box.querySelector('.cf-ten').value, p_mien_nhan: f('mien_nhan'), p_ghi_chu_bb: f('ghi_chu_bb'), p_can_hen: f('can_hen'), p_can_vaccine: f('can_vaccine'), p_can_ngay_tiem: f('can_ngay_tiem'), p_can_noi_tiem: f('can_noi_tiem'), p_can_ly_do: f('can_ly_do'), p_goi_y: box.querySelector('.cf-gy').value, p_dang_dung: box.querySelector('.cf-dd').checked })
          .then(function () { done('Đã lưu') }, function (e) { b.disabled = false; toast((e && e.message) || 'Không lưu được.', 0, 1) })
      }
    })
  });

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
    all('[data-tknosdt]', function (b) { b.onclick = function () { AS.tk.nosdt = !AS.tk.nosdt; render() } });
    all('[data-tkedit]', function (b) { b.onclick = function () { accModal(byAcc(b.dataset.tkedit)) } });
    all('[data-tkpw]', function (b) { b.onclick = function () { pwAsm(byAcc(b.dataset.tkpw)) } });
    all('[data-tklock]', function (b) { b.onclick = function () { lockAsm(byAcc(b.dataset.tklock)) } });
    all('[data-tkdel]', function (b) { b.onclick = function () { delAsm(byAcc(b.dataset.tkdel)) } });
    all('[data-setql]', function (b) { b.onclick = function () { setQl(b.dataset.setql) } });
    /* nhớ nhóm Shop nào đang mở để vẽ lại không bị gập */
    all('.shg > summary', function (sm) { sm.addEventListener('click', function () { var d = sm.parentNode; setTimeout(function () { AS.tk.open[d.getAttribute('data-s')] = d.open }, 0) }) });
    all('#hrbox > summary', function (sm) { sm.addEventListener('click', function () { var d = sm.parentNode; setTimeout(function () { AS.tk.hrOpen = d.open }, 0) }) });
    all('.nvd > summary', function (sm) { sm.addEventListener('click', function () { var d = sm.parentNode; setTimeout(function () { AS.nv.open = d.open }, 0) }) });
    all('[data-nvr]', function (b) { b.onclick = function () { var t = AS.nv, n = Number(b.dataset.nvr); t.den = isoD(new Date()); t.tu = isoD(new Date(Date.now() - (n - 1) * 864e5)); t.rows = null; t.err = ''; t.loading = false; render() } });
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
    on('#pd-up', 'onclick', function () { AS.tl.tab = 'kho'; if (window.ASMLIB) ASMLIB.select('PHAC_DO'); S.page = 'a_tl'; render(); var tries = 0, tm = setInterval(function () { var h = [].slice.call(document.querySelectorAll('.card h2')).filter(function (x) { return /^2\./.test(x.textContent.trim()) })[0]; if (h) { clearInterval(tm); h.scrollIntoView({ behavior: 'smooth', block: 'start' }) } else if (++tries > 20) clearInterval(tm) }, 300) });
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
 * Hai chế độ riêng biệt: "Theo quy trình" (MẶC ĐỊNH; 4 mốc D, D+1, D+3, D+7 + Trễ lịch; mẫu tạm từ msg_templates nhóm KH_MOI) và "Theo vaccine".
 * THEO VACCINE (Hải 24/11 tối): app liệt kê vaccine bé đã đến tuổi theo PHÁC ĐỒ; nhân viên chọn 1 vaccine thì app TỰ TÌM trong kho Excel và CHỌN SẴN câu tư vấn,
 *   nguy cơ & đường lây, khuyến mãi (CTKM còn hạn) hoặc CSKH sau tiêm của ĐÚNG vaccine đó (nối qua Bảng giá: bệnh ↔ tên thương mại; xem hàm profile()).
 *   Nhân viên chỉ bật/tắt hoặc đổi, không phải đi tìm trong cả danh sách. Chế độ "Theo quy trình" cũng gắn được vaccine để tự thêm nguy cơ/khuyến mãi.
 * MỌI tin có chữ ký tự động "Người chăm sóc: <tên NV> · <Shop> / SĐT/Zalo: <SĐT NV>" (SĐT lấy từ tài khoản; chưa có thì nhắc nhập ngay); tắt được bằng công tắc.
 * Chỗ trống trong mẫu ({{NGAY_TIEM}}, {{SO_MUI}}, …) hiện ô để nhân viên điền; còn trống mà gửi thì app hỏi lại.
 * Nhóm khách chỉ theo TUỔI (từ ngày sinh); hiện CHƯA có nhóm "bầu" nên mẫu thuộc nhóm bầu được ẩn (AN_BAU).
 * Giá/CTKM trong mẫu đã được ASM quét và thay khi duyệt Bảng giá/CTKM: nhân viên KHÔNG thay giá tại chỗ.
 * Nội dung kho tải lại mỗi khi mở màn soạn tin nếu đã quá 30 giây (ASM duyệt lô mới thì nhân viên thấy ngay lần mở sau).
 * Ba nút gửi, mỗi nút làm đúng việc của nó rồi tự ghi 1 lần "Nhắn" (hàm u_msg_log):
 *   Zalo     = chép tin + mở zalo.me/<số> (Zalo không cho web điền sẵn tin hay tự kết bạn: nhân viên dán tin rồi gửi/kết bạn)
 *   SMS      = mở app tin nhắn tới số, nội dung điền sẵn (đồng thời chép tin phòng khi máy tính không có app SMS)
 *   Sao chép = chỉ chép tin */
var MSG = (function () {
  var CSS = `
.imgrow,.imggrid{display:flex;gap:8px;overflow-x:auto;padding-bottom:4px}.imggrid{display:grid;grid-template-columns:repeat(3,1fr);overflow:visible}
.ithumb{flex:none;position:relative;width:96px;border:2px solid var(--line);border-radius:14px;background:var(--card);padding:4px;font-size:11px;text-align:left;color:var(--ink);transition:border-color .15s,transform .12s}
.ithumb:active{transform:scale(.97)}.imggrid .ithumb{width:auto}.ithumb img{width:100%;height:64px;object-fit:cover;border-radius:10px;display:block}
.ithumb.on{border-color:var(--teal);background:var(--tint)}.ithumb span{display:block;margin-top:3px;line-height:1.2;max-height:2.4em;overflow:hidden}
.ick{position:absolute;top:6px;right:6px;background:var(--teal);color:#fff;border-radius:50%;width:20px;height:20px;display:grid;place-items:center;font-size:12px}
.person{display:grid;grid-template-columns:1fr auto;gap:10px;align-items:center;padding:12px 0;border-top:1px solid var(--line)}.person:first-child{border-top:0}.person .phone{font-size:20px}
.btn.bo{border:1.5px solid var(--teal);color:var(--teal-d);background:var(--card);min-height:44px}

/* ---- khung màn soạn tin: kính mờ, tông teal/navy của Long Châu ---- */
.modal.mm{align-items:end;background:rgba(6,24,36,.55);overscroll-behavior:none;touch-action:none}.modal.mm .msheet{touch-action:pan-y}
.modal.mm.first{animation:mslide .2s ease-out}
@keyframes mslide{from{transform:translateY(18px)}to{transform:none}}
@media (prefers-reduced-motion:reduce){.modal.mm.first{animation:none}.abt,.ithumb,.mcard,.mrow{transition:none}}
.msheet{width:100%;max-width:580px;max-height:94vh;max-height:94dvh;display:flex;flex-direction:column;background:var(--bg);border-radius:26px 26px 0 0;overflow:hidden;box-shadow:0 -14px 48px rgba(6,24,36,.38)}
@media(min-width:960px){.modal.mm{align-items:center}.msheet{border-radius:26px;max-height:90vh}}
.mhd{position:relative;flex:none;padding:14px 16px 14px;color:#fff;background:radial-gradient(120% 150% at 0% 0%,rgba(0,187,186,.5),transparent 55%),radial-gradient(90% 130% at 100% 0%,rgba(253,130,6,.2),transparent 52%),linear-gradient(160deg,#0c3b50,#082A3E 62%,#0a2233)}
.mtop{display:flex;align-items:center;gap:12px}.mav{flex:none;width:44px;height:44px;border-radius:14px;display:grid;place-items:center;font-weight:700;font-size:18px;background:rgba(255,255,255,.14);border:1px solid rgba(255,255,255,.3);-webkit-backdrop-filter:blur(8px);backdrop-filter:blur(8px)}
.mtt{min-width:0;flex:1}.mtt .t1{font-weight:700;font-size:17px;line-height:1.25;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.mtt .t2{font-size:13px;color:rgba(255,255,255,.76);margin-top:2px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.mclose{flex:none;width:40px;height:40px;border-radius:13px;border:1px solid rgba(255,255,255,.28);background:rgba(255,255,255,.12);color:#fff;display:grid;place-items:center;transition:background .15s}.mclose:active{background:rgba(255,255,255,.24)}
.mseg{display:grid;grid-template-columns:1fr 1fr;gap:4px;margin-top:12px;padding:4px;border-radius:15px;background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);-webkit-backdrop-filter:blur(10px);backdrop-filter:blur(10px)}
.mseg button{border:0;background:none;color:rgba(255,255,255,.84);font-weight:600;font-size:14.5px;min-height:42px;border-radius:11px;display:flex;align-items:center;justify-content:center;gap:7px;transition:background .18s,color .18s,box-shadow .18s}
.mseg button.on{background:#fff;color:var(--navy);box-shadow:0 2px 10px rgba(0,0,0,.2)}.mseg button svg{width:18px;height:18px}
.mbody{flex:1;min-height:0;overflow:auto;overscroll-behavior:contain;padding:14px 16px 18px;display:flex;flex-direction:column;gap:16px}.mbody>*{flex:none;min-width:0;max-width:100%}
.mbody section{display:grid;gap:8px;align-content:start;grid-template-columns:minmax(0,1fr)}
.mh{display:flex;align-items:center;gap:8px;margin:0;font-size:12px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:var(--soft)}.mh2{margin-top:6px}
.mnote{display:flex;align-items:center;gap:8px;color:var(--muted);font-size:13px;line-height:1.4}
.sg2{flex:none;width:9px;height:9px;border-radius:50%;background:var(--orange);display:inline-block}
.mft{flex:none;padding:10px 16px calc(12px + env(safe-area-inset-bottom));background:rgba(255,255,255,.84);-webkit-backdrop-filter:blur(16px) saturate(1.4);backdrop-filter:blur(16px) saturate(1.4);border-top:1px solid var(--line);display:grid;gap:8px}
@media (prefers-color-scheme:dark){.mft{background:rgba(17,37,52,.86)}}
.mst{display:flex;align-items:center;gap:8px;color:var(--muted);font-size:12.5px;line-height:1.35}.cnt{flex:none;font-weight:700;font-size:12px;border-radius:99px;padding:3px 9px;background:var(--tint);color:var(--teal-d)}
.act3{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:8px}
.act3 .abt{flex-direction:column;gap:3px;min-height:64px;padding:8px 4px;text-align:center}.act3 .abt span{justify-items:center;min-width:0;max-width:100%}.act3 .abt small{max-width:100%;overflow:hidden;text-overflow:ellipsis}
.abt{border:0;border-radius:16px;min-height:56px;padding:0 8px;font-weight:700;display:flex;align-items:center;justify-content:center;gap:8px;text-align:left;transition:transform .12s,box-shadow .2s,opacity .2s}
.abt:active{transform:scale(.97)}.abt:disabled{opacity:.45;box-shadow:none}.abt svg{width:22px;height:22px}.abt span{display:grid;line-height:1.15}.abt b{font-size:15px}.abt small{font-size:11px;font-weight:500;opacity:.85;white-space:nowrap}
.abt.p{background:linear-gradient(135deg,#1f9494,#126868);color:#fff;box-shadow:0 8px 18px rgba(29,135,135,.32)}.abt.s{background:var(--navy);color:#fff}.abt.g{background:var(--card);border:1.5px solid var(--line);color:var(--ink)}
.abt.o{border:1.5px solid var(--teal);background:var(--card);color:var(--teal-d);min-height:46px;width:100%}
.mhd{color:var(--ink);background:var(--card);border-bottom:1px solid var(--line);padding:12px 16px 10px}
.mav{background:var(--tint);border:0;color:var(--teal-d);-webkit-backdrop-filter:none;backdrop-filter:none}
.mtt .t2{color:var(--muted)}.mclose{border:0;background:var(--bg);color:var(--ink)}.mclose:active{background:var(--line)}
.mseg{background:var(--bg);border:0;margin-top:10px;-webkit-backdrop-filter:none;backdrop-filter:none}
.mseg button{color:var(--muted);min-height:38px;font-size:13.5px}.mseg button.on{background:var(--teal);color:#fff;box-shadow:0 4px 12px rgba(29,135,135,.3)}
.abt.s{background:var(--card);border:1.5px solid var(--teal);color:var(--teal-d)}
.act3{grid-template-columns:1.35fr 1fr 1.15fr}.act3 .abt{flex-direction:row;min-height:52px;gap:7px}.act3 .abt small{display:none}.act3 .abt svg{width:20px;height:20px}
.mfull{position:absolute;inset:0;z-index:6;display:grid;align-items:end;justify-items:center;background:rgba(6,24,36,.55)}.mfs{width:100%;max-width:580px;max-height:92dvh;max-height:92vh;display:flex;flex-direction:column;background:var(--card);border-radius:22px 22px 0 0;overflow:hidden}.mfh{display:flex;align-items:center;gap:8px;padding:14px 16px;border-bottom:1px solid var(--line)}.mfh b{flex:1;font-size:16px}.mfm{padding:8px 16px;color:var(--muted);font-size:13px}.mfw{color:var(--warn);font-weight:600}.mfb{flex:1;min-height:0;overflow:auto;overscroll-behavior:contain;padding:6px 16px 14px;font-size:15.5px;line-height:1.6;white-space:normal;user-select:text;-webkit-user-select:text}.mff{display:flex;gap:8px;padding:10px 16px calc(12px + env(safe-area-inset-bottom));border-top:1px solid var(--line)}.mff .abt{flex:1;min-height:48px;justify-content:center}.mpv{flex:none;padding:10px 14px 0;background:var(--bg)}.mpv .ed textarea{min-height:0;height:96px;resize:none;font-size:14px;line-height:1.5;padding:8px 12px}
.mpv .ed .eh,.mpv .ed .ef{padding:6px 12px}.mpv .ed .eh b,.mpv .badge,.mpv .cnt{white-space:nowrap}.mpv .badge{font-size:11.5px;padding:3px 9px}.cnt.warn{background:var(--warn-bg);color:var(--warn)}.act3 .abt b{white-space:nowrap;font-size:14.5px}.mpv .ed{transition:border-color .3s,box-shadow .3s}.mpv.flash .ed{animation:vflash .9s cubic-bezier(.2,.8,.2,1)}
@keyframes vflash{0%{border-color:var(--teal);box-shadow:0 0 0 4px rgba(0,187,186,.28);background:var(--tint)}100%{border-color:var(--line);box-shadow:var(--shadow)}}.abt.mbig{width:100%;min-height:50px;margin-top:6px}
.stp{display:grid;grid-template-columns:repeat(5,1fr);position:relative;padding-top:2px}.stp:before{content:"";position:absolute;left:10%;right:10%;top:20px;height:2px;background:var(--line)}
.stp button{border:0;background:none;display:grid;justify-items:center;gap:6px;position:relative;padding:0 2px;font-weight:600;font-size:12px;color:var(--muted);min-height:68px;align-content:start}
.stp i{position:relative;width:38px;height:38px;border-radius:50%;display:grid;place-items:center;background:var(--card);border:2px solid var(--line);font-style:normal;font-size:12px;font-weight:700;color:var(--muted);transition:background .18s,border-color .18s,box-shadow .18s,color .18s}
.stp i svg{width:17px;height:17px}.stp button.on i{background:var(--teal);border-color:var(--teal);color:#fff;box-shadow:0 0 0 5px rgba(29,135,135,.17)}.stp button.on{color:var(--ink)}
.stp .sg{position:absolute;top:-3px;right:-3px;width:12px;height:12px;border-radius:50%;background:var(--orange);border:2px solid var(--bg)}
.tcard{position:relative;background:var(--card);border:1px solid var(--line);border-radius:18px;padding:13px 14px 12px 18px;box-shadow:var(--shadow);overflow:hidden}
.tcard:before{content:"";position:absolute;left:0;top:0;bottom:0;width:4px;background:linear-gradient(var(--teal),#00BBBA)}
.tctop{display:flex;align-items:flex-start;gap:10px}.tcard .tt{flex:1;font-weight:700;font-size:15.5px;line-height:1.3}.tcard .pv{color:var(--muted);font-size:14px;line-height:1.5;margin-top:8px;display:-webkit-box;-webkit-line-clamp:3;-webkit-box-orient:vertical;overflow:hidden}
.mtags{display:flex;flex-wrap:wrap;gap:6px;margin-top:8px}.mtags:empty{display:none}
.mp{display:inline-flex;align-items:center;gap:4px;font-size:11.5px;font-weight:700;padding:3px 9px;border-radius:99px;background:var(--tint);color:var(--teal-d);white-space:nowrap}.mp.ro{background:var(--new-bg);color:var(--new)}.mp.gr{background:var(--ok-bg);color:var(--ok)}.mp.or{background:var(--warn-bg);color:var(--warn)}
.mlink{border:0;background:none;color:var(--teal-d);font-weight:700;font-size:13.5px;padding:6px 4px;min-height:36px;white-space:nowrap}.mlink:active{opacity:.6}
.mlist{display:grid;gap:8px;grid-template-columns:minmax(0,1fr)}
.mrow,.mcard{width:100%;text-align:left;background:var(--card);border:1px solid var(--line);border-radius:16px;padding:11px 14px;display:flex;align-items:center;gap:10px;color:inherit;transition:border-color .15s,transform .12s,background .15s}
.mrow:active,.mcard:active{transform:scale(.99);border-color:var(--teal)}.mrow>div{flex:1;min-width:0}.mrow b{display:block;font-size:14.5px;margin-bottom:2px}.mrow span{display:block;color:var(--muted);font-size:13px;line-height:1.4;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.mrow svg{flex:none;color:var(--soft);width:18px;height:18px}
.mcard{display:grid;gap:6px;grid-template-columns:minmax(0,1fr)}.mcard.on{border-color:var(--teal);background:var(--tint)}.mcard .mc1{display:flex;gap:8px;align-items:flex-start}.mcard .mc1 b{flex:1;font-size:14.5px;line-height:1.3}.mcard .pv2{color:var(--muted);font-size:13.5px;line-height:1.45;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}
.mck{flex:none;width:24px;height:24px;border-radius:50%;border:2px solid var(--line);display:grid;place-items:center;color:transparent;transition:background .15s,border-color .15s;margin-top:1px}.mck svg{width:14px;height:14px;stroke-width:3}.mcard.on .mck{background:var(--teal);border-color:var(--teal);color:#fff}
.mwarn,.mempty{background:var(--card);border:1.5px dashed var(--line);border-radius:18px;padding:16px;color:var(--muted);font-size:14px;line-height:1.5}.mempty{display:grid;gap:2px;text-align:center}.mempty b{color:var(--ink);font-size:15.5px}
.mwarn{background:var(--warn-bg);border-color:transparent;color:var(--warn)}
.ed{background:var(--card);border:1px solid var(--line);border-radius:18px;box-shadow:var(--shadow);overflow:hidden}
.ed .eh,.ed .ef{display:flex;align-items:center;gap:8px;padding:9px 12px}.ed .eh{border-bottom:1px solid var(--line)}.ed .ef{border-top:1px solid var(--line);color:var(--muted);font-size:13px}.ed .eh b{font-size:14.5px}
.ed textarea{display:block;border:0;border-radius:0;min-height:190px;padding:12px 14px;line-height:1.6;font-size:15.5px;background:transparent;resize:vertical;outline:0}
.badge{font-size:12px;font-weight:700;border-radius:99px;padding:3px 10px}.b-ok{background:var(--ok-bg);color:var(--ok)}.b-ed{background:var(--warn-bg);color:var(--warn)}.over{color:var(--bad);font-weight:700}.longw{color:var(--warn);font-weight:700}
.ls{background:var(--card);border:1px solid var(--line);border-radius:18px;box-shadow:var(--shadow);overflow:hidden}
.li{display:flex;align-items:center;gap:12px;padding:11px 14px;min-height:58px;width:100%;text-align:left;background:none;border:0;border-top:1px solid var(--line);color:inherit;cursor:pointer}.ls>.li:first-child{border-top:0}
.li .ico{flex:none;width:38px;height:38px;border-radius:12px;display:grid;place-items:center;background:var(--tint);color:var(--teal-d)}.li .ico svg{width:19px;height:19px}.li .lt{flex:1;min-width:0}.li .lt b{display:block;font-size:14.5px}.li .lt small{display:block;color:var(--muted);font-size:13px;line-height:1.35}.li .chev{flex:none;color:var(--soft);width:18px;height:18px}
.li:disabled{opacity:.6}.li select{border:1px solid var(--line);border-radius:12px;padding:8px 10px;background:var(--bg);min-height:40px;max-width:55%}
.tg{position:relative;flex:none;width:48px;height:28px}.tg input{position:absolute;inset:0;width:100%;height:100%;margin:0;opacity:0;cursor:pointer;z-index:1}.tg i{position:absolute;inset:0;border-radius:99px;background:#c9d3dc;transition:background .2s}
.tg i:after{content:"";position:absolute;left:3px;top:3px;width:22px;height:22px;border-radius:50%;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,.28);transition:transform .2s}.tg input:checked+i{background:var(--teal)}.tg input:checked+i:after{transform:translateX(20px)}.tg input:disabled+i{opacity:.5}.tg input:focus-visible+i{outline:2px solid var(--teal);outline-offset:2px}
.vxl2{display:grid;gap:2px;padding:2px 14px 10px 64px;border-top:0}.vxl2 label{display:flex;gap:10px;align-items:flex-start;padding:6px 0;font-size:14px;cursor:pointer}.vxl2 input{width:20px;height:20px;accent-color:var(--teal);flex:none;margin-top:1px}.vxl2 small{display:block;color:var(--muted)}
.sg3{display:flex;gap:6px;flex-wrap:wrap}.sg3 button{border:1.5px solid var(--line);background:var(--card);border-radius:13px;padding:8px 14px;font-weight:600;min-height:42px;color:var(--ink);transition:border-color .15s,background .15s}.sg3 button.on{border-color:var(--teal);background:var(--tint);color:var(--teal-d)}
.mtabs{display:flex;gap:6px;overflow-x:auto;padding-bottom:2px;scrollbar-width:none}.mtabs::-webkit-scrollbar{display:none}.mtabs button{flex:none;border:0;background:var(--new-bg);color:var(--new);border-radius:99px;padding:0 14px;font-weight:700;font-size:13.5px;min-height:40px;display:flex;align-items:center;gap:6px;transition:background .15s,color .15s}
.mtabs button.on{background:var(--navy);color:#fff}.mtabs em{font-style:normal;font-size:11.5px;opacity:.8}
.mchips{display:flex;gap:6px;overflow-x:auto;padding-bottom:2px;scrollbar-width:none}.mchips::-webkit-scrollbar{display:none}.mchips button{flex:none;border:1px solid var(--line);background:var(--card);border-radius:99px;padding:0 12px;font-weight:600;font-size:13px;min-height:36px;color:var(--ink)}.mchips button.on{border-color:var(--teal);background:var(--tint);color:var(--teal-d)}
.msearch{display:flex;align-items:center;gap:10px;background:var(--card);border:1px solid var(--line);border-radius:14px;padding:0 14px;min-height:46px;color:var(--soft)}.msearch input{border:0;outline:0;background:none;flex:1;min-width:0;color:var(--ink);font:inherit}
.sk{pointer-events:none}.sk i{display:block;height:12px;border-radius:6px;background:linear-gradient(90deg,var(--new-bg),var(--line),var(--new-bg));background-size:200% 100%;animation:msh 1.2s linear infinite}.sk i+i{margin-top:8px;width:70%}@keyframes msh{to{background-position:-200% 0}}
@media (prefers-reduced-motion:reduce){.sk i{animation:none}}
@media (prefers-color-scheme:dark){.mlink,.abt.o{color:#7ee0df}.mp,.cnt,.li .ico{color:#7ee0df}.mp.ro{color:var(--muted)}.abt.s{background:#1b4b63}.tg i{background:#3a5163}.mcard.on{background:#12343f}}
.lbtn{background:none;border:0;text-align:left;color:inherit;padding:0;cursor:pointer;font:inherit}
.sdtbox{display:flex;gap:8px;padding:2px 14px 12px 64px}.sdtbox input,.fvi{border:1px solid var(--line);border-radius:12px;padding:10px 12px;background:var(--bg);min-height:44px;font:inherit;color:var(--ink)}.sdtbox input{flex:1;min-width:0}.fvi{display:block;width:100%;margin-top:6px}
.mp.zero{background:var(--warn-bg);color:var(--warn)}.mp.lkw{white-space:normal;line-height:1.35;border-radius:12px;text-align:left;max-width:100%}
.lkit{display:flex;gap:10px;align-items:flex-start;padding:8px 0}.lkit input{margin-top:3px}.lkit span{display:grid;gap:1px}.lkds{color:var(--teal-d);font-weight:600;font-size:12.5px}.lkds:empty{display:none}.lkpin{color:var(--teal-d);margin-top:4px}.lkdp{margin:2px 0 10px;padding:10px;border:1px solid var(--line);border-radius:12px;background:var(--bg)}.lkdi{font-size:13px;line-height:1.45;margin-bottom:8px}.lkch{display:flex;flex-wrap:wrap;gap:6px;align-items:center;margin-bottom:8px}.lkchip{display:inline-flex;align-items:center;gap:6px;border:1px solid var(--line);border-radius:99px;padding:8px 12px;font-weight:600;font-size:14px;background:var(--card)}.lkchip:has(input:checked){background:var(--teal);border-color:var(--teal);color:#fff}.lkdt{display:grid;gap:4px;margin-bottom:8px}.lkdt input{width:100%}.lkres2{font-size:13.5px;line-height:1.5;color:var(--ink)}.lkt{background:var(--card);border:1px solid var(--line);border-radius:14px;margin-bottom:8px;overflow:hidden}.lkt summary{list-style:none;cursor:pointer;padding:11px 13px;display:grid;gap:3px}.lkt summary::-webkit-details-marker{display:none}.lkt summary b{font-size:14.5px;line-height:1.35}.lkt summary small{color:var(--muted);font-size:12.5px;line-height:1.4;overflow:hidden;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical}.lkt[open] summary small{display:none}.lkx{padding:0 13px 6px;font-size:14px;line-height:1.55}.lkt .mlink{margin:2px 13px 11px}
.lkn{display:flex;gap:6px;overflow-x:auto;padding:2px 0 8px;scrollbar-width:none}.lkn::-webkit-scrollbar{display:none}.lkn button{flex:0 0 auto;border:1px solid var(--line);background:var(--card);color:var(--ink);border-radius:99px;padding:7px 12px;font-size:13px;font-weight:600;min-height:36px}
.lkage{display:grid;grid-template-columns:1fr 1.3fr;gap:8px;margin-bottom:6px}.lkage select,.lkage input{border:1px solid var(--line);border-radius:12px;padding:10px;background:var(--card);min-height:44px;color:var(--ink);min-width:0;font-size:15px}.mcard.vac .mtags{margin-top:0}.mcard.vac .mc1{align-items:center}
.psw{display:flex;gap:6px;margin-top:10px}.psw button{flex:1;border:1.5px solid var(--line);background:var(--card);border-radius:13px;min-height:42px;font-weight:600;font-size:13.5px;color:var(--ink);transition:border-color .15s,background .15s}.psw button.on{border-color:var(--teal);background:var(--tint);color:var(--teal-d)}
.ibtn{flex:none;width:42px;height:42px;border-radius:14px;border:1px solid var(--line);background:var(--card);display:grid;place-items:center;color:var(--teal-d);margin-right:8px}.ibtn svg{width:20px;height:20px}
.lkf{display:grid;grid-template-columns:1fr 1fr;gap:8px}.lkf select{border:1px solid var(--line);border-radius:12px;padding:10px;background:var(--card);min-height:44px;color:var(--ink);min-width:0}
.lkp{padding:12px 14px;display:grid;gap:4px;line-height:1.55;font-size:14.5px}.lkb{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:11px 13px;margin-bottom:8px;font-size:14px;line-height:1.5}.lkb b{display:block;margin-bottom:2px;font-size:14.5px}.lkb div b{display:inline;font-size:inherit}.lkr{display:flex;justify-content:space-between;gap:10px;padding:6px 0;border-bottom:1px solid var(--line)}
.act2{display:grid;grid-template-columns:repeat(auto-fit,minmax(0,1fr));gap:8px}.act2 .abt{flex-direction:column;gap:3px;min-height:60px;text-align:center}.act2 .abt span{justify-items:center}
.mrow2{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:6px}.mrow2 .abt{min-height:48px;font-size:14px}
@media (prefers-color-scheme:dark){.psw button.on{color:#7ee0df}}
`;
  (function () { try { var s = document.createElement('style'); s.id = 'msg-css'; s.textContent = CSS; document.head.appendChild(s) } catch (e) { } })();
  var ICN = {
    x: '<path d="M6 6l12 12M18 6 6 18"/>', chat: '<path d="M21 12a8 8 0 0 1-11.6 7.1L4 20l1-4.6A8 8 0 1 1 21 12z"/>', sms: '<path d="M4 5h16v11H9l-5 4z"/><path d="M8.5 9.5h7M8.5 12.5h4"/>',
    gift: '<path d="M4 11h16v9H4zM3 7.5h18V11H3zM12 7.5V20"/><path d="M12 7.5C10 7.5 8 6.7 8 5.2S9.6 3 12 7.5zM12 7.5c2 0 4-.8 4-2.3S14.4 3 12 7.5z"/>',
    shield: '<path d="M12 3 4 6v6c0 5 3.5 8 8 9 4.5-1 8-4 8-9V6z"/><path d="M12 9v4M12 16h.01"/>', spark: '<path d="M12 3l2 5 5 2-5 2-2 5-2-5-5-2 5-2z"/>',
    pin: '<path d="M12 21s7-6 7-11a7 7 0 1 0-14 0c0 5 7 11 7 11z"/><circle cx="12" cy="10" r="2.5"/>', image: '<rect x="3" y="4" width="18" height="16" rx="3"/><circle cx="9" cy="10" r="1.8"/><path d="M21 16l-5-5-8 8"/>',
    chev: '<path d="M9 6l6 6-6 6"/>', check: '<path d="M5 12l4.5 4.5L19 7"/>', layers: '<path d="M12 3l9 5-9 5-9-5zM3 13l9 5 9-5"/>',
    book: '<path d="M5 4h11a3 3 0 0 1 3 3v13H8a3 3 0 0 1-3-3zM5 17a3 3 0 0 1 3-3h11"/>', route: '<circle cx="6" cy="6" r="2.5"/><circle cx="18" cy="18" r="2.5"/><path d="M8.5 6H15a3 3 0 0 1 0 6H9a3 3 0 0 0 0 6h6.5"/>', reset: '<path d="M4 4v6h6M20 20v-6h-6M20 10A8 8 0 0 0 6 6L4 10M4 14a8 8 0 0 0 14 4l2-4"/>'
  };
  Object.keys(ICN).forEach(function (k) { if (!IC[k]) IC[k] = ICN[k] });
  var AN_BAU = true;   /* hiện chưa có nhóm khách "bầu" (Hải 24/11): ẩn mẫu thuộc nhóm bầu khỏi nhân viên. Đặt false để hiện lại. */
  var VARS = ['XUNG_HO', 'XUNG_HO_HOA', 'TEN_NGUOI_LH', 'TEN_BE', 'TUOI', 'TEN_NV', 'TEN_CO_SO', 'DIA_CHI', 'GIO_MO_CUA', 'HOTLINE', 'SDT_NV', 'TEN_VACCINE', 'NGAY_TIEM', 'SO_MUI', 'NGAY_HEN', 'LINK_GOOGLE_MAPS', 'NOI_DUNG_PHAN_HOI'];
  var VARDESC = { XUNG_HO: 'mẹ / bố / anh chị…', XUNG_HO_HOA: 'Mẹ / Bố… (đầu câu)', TEN_NGUOI_LH: 'tên người liên hệ', TEN_BE: 'tên bé', TUOI: 'tuổi bé (ngày/tháng/năm)', TEN_NV: 'tên nhân viên', TEN_CO_SO: 'tên Shop', DIA_CHI: 'địa chỉ Shop', GIO_MO_CUA: 'giờ mở cửa', HOTLINE: 'hotline Shop', SDT_NV: 'SĐT nhân viên (tự điền từ tài khoản)', TEN_VACCINE: 'vaccine nhân viên đã chọn (tự điền)', NGAY_TIEM: 'ngày tiêm (nhân viên điền)', SO_MUI: 'mũi thứ mấy (nhân viên điền)', NGAY_HEN: 'ngày hẹn mũi tiếp (nhân viên điền)', LINK_GOOGLE_MAPS: 'link bản đồ Shop (nhân viên điền)', NOI_DUNG_PHAN_HOI: 'nội dung phản hồi (nhân viên điền)' };
  var PHL = { NGAY_TIEM: 'Ngày tiêm (ví dụ 24/11)', SO_MUI: 'Mũi thứ mấy (ví dụ 2)', NGAY_HEN: 'Ngày hẹn mũi tiếp', LINK_GOOGLE_MAPS: 'Link bản đồ Shop', NOI_DUNG_PHAN_HOI: 'Nội dung phản hồi', TEN_VACCINE: 'Tên vaccine', DIA_CHI: 'Địa chỉ Shop', HOTLINE: 'Hotline Shop', GIO_MO_CUA: 'Giờ mở cửa', SDT_NV: 'SĐT của anh/chị' };
  var STEPS = [['D', 'D · Khởi tạo'], ['D1', 'D+1 · Thăm dò'], ['D3', 'D+3 · Thúc đẩy'], ['D7', 'D+7 · Cấp thiết'], ['TRE', 'Trễ lịch']];
  var STEP_SUB = { D: 'Khởi tạo', D1: 'Thăm dò', D3: 'Thúc đẩy', D7: 'Cấp thiết', TRE: 'Trễ lịch' }, STEP_DOT = { D: 'D', D1: 'D+1', D3: 'D+3', D7: 'D+7', TRE: '' };
  var GROUPS = [['KH_MOI', 'KH mới'], ['CSKH_SAU_TIEM', 'CSKH sau tiêm'], ['BAU_NHI', 'Bầu – Nhi'], ['CHUNG', 'Chung']];
  var XH = [['', 'Tự động theo người liên hệ'], ['mẹ', 'mẹ'], ['bố', 'bố'], ['anh', 'anh'], ['chị', 'chị'], ['cô', 'cô'], ['chú', 'chú'], ['bác', 'bác'], ['ông', 'ông'], ['bà', 'bà']];
  var TABS = [['CD', 'Theo vaccine', 'MAU_CHU_DE'], ['NHI', 'Sơ sinh & Nhi', 'MAU_TUOI'], ['CH', 'Mẫu chung', 'MAU_CHUNG'], ['ST', 'Sau tiêm', 'MAU_SAU_TIEM']];
  var LIBT = ['MAU_CHU_DE', 'MAU_TUOI', 'MAU_SAU_TIEM', 'MAU_CHUNG', 'CTKM', 'NGUY_CO', 'BANG_GIA', 'PHAC_DO', 'HOI_DAP'], MAXLEN = 6000, LONG = 900;   /* LONG: từ ngần này ký tự thì cảnh báo "tin dài"; MAXLEN: giới hạn cứng. Lịch sử chỉ ghi 2000 ký tự đầu (u_msg_log). */

  function strip(s) { return String(s || '').replace(/\s*\([^)]*\)/g, '').trim() }
  function given(s) { var p = strip(s).split(/\s+/); return p[p.length - 1] || '' }
  function cap(s) { return s ? s.charAt(0).toUpperCase() + s.slice(1) : s }
  function pad(n) { return n < 10 ? '0' + n : '' + n }
  function todayIso() { var d = new Date(); return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()) }
  function dmy(iso) { var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso || ''); return m ? m[3] + '/' + m[2] + '/' + m[1] : '' }
  function money(n) { return Number(n).toLocaleString('vi-VN') + 'đ' }
  function fph(s) { s = String(s || ''); return /^0\d{9}$/.test(s) ? s.slice(0, 4) + ' ' + s.slice(4, 7) + ' ' + s.slice(7) : s }
  function custDays(c) { var b = c && c.ngay_sinh ? new Date(c.ngay_sinh + 'T00:00:00') : null; return b ? Math.max(0, Math.floor((Date.now() - b) / 86400000)) : null }
  function tuoiTxt(c) { var d = custDays(c); if (d === null) return ''; if (d < 60) return d + ' ngày tuổi'; var m = Math.floor(d / 30.4375); return m < 24 ? m + ' tháng tuổi' : Math.floor(m / 12) + ' tuổi' }
  function dayLabel(d) { if (d == null) return 'không giới hạn'; if (d < 60) return d + ' ngày'; if (d < 730) return Math.round(d / 30.4375) + ' tháng'; var y = d / 365.25; return (Math.round(y * 10) / 10) + ' tuổi' }
  function xungOf(link) { var q = String(link && link.quan_he || '').toLowerCase(); return ['mẹ', 'bố', 'ông', 'bà', 'anh', 'chị', 'cô', 'chú', 'bác'].indexOf(q) >= 0 ? q : 'anh/chị' }
  function telOf(l) { var s = String(l && l.sdt || '').replace(/[\s.\-()]/g, ''); if (/^\+84\d{9}$/.test(s)) s = '0' + s.slice(3); else if (/^84\d{9}$/.test(s)) s = '0' + s.slice(2); return s }
  function stepNow(t) { var a = ageDays(t.d_giao); return a < 1 ? 'D' : a < 3 ? 'D1' : a < 7 ? 'D3' : 'D7' }
  function fill(text, cx) { return String(text || '').replace(/\{\{([A-Z_]+)\}\}/g, function (m, k) { return cx[k] != null && cx[k] !== '' ? cx[k] : m }) }
  function shopOf() { return (D.shops && D.shops[ME.id_shop]) || {} }
  function fk(x) { return window.LIBSCAN ? LIBSCAN.fold(x || '') : String(x || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd') }
  /* Biến tự điền: tên NV, SĐT NV (từ tài khoản), Shop; extra = vaccine đã chọn + chỗ nhân viên đã điền tay. Biến thiếu giá trị thì GIỮ NGUYÊN {{…}} để app nhắc điền. */
  function ctx(c, link, xhOverride, extra) {
    var xh = xhOverride || xungOf(link), sh = shopOf(), o = { XUNG_HO: xh, XUNG_HO_HOA: cap(xh), TEN_NGUOI_LH: given(link && link.ho_ten), TEN_BE: given(c && c.ho_ten) || 'bé', TUOI: tuoiTxt(c), TEN_NV: given(ME.ho_ten), SDT_NV: fph(ME.sdt), TEN_CO_SO: strip(sh.ten_shop) || 'Trung tâm tiêm chủng Long Châu', DIA_CHI: oneLine(sh.dia_chi) || oneLine(sh.ten_shop), GIO_MO_CUA: sh.gio_mo_cua || '', HOTLINE: sh.hotline || '', LINK_GOOGLE_MAPS: '' };
    if (extra) for (var k in extra) if (extra[k] != null && extra[k] !== '') o[k] = extra[k];
    return o
  }
  function preview(text) { return fill(text, { XUNG_HO: 'mẹ', XUNG_HO_HOA: 'Mẹ', TEN_NGUOI_LH: 'Lan', TEN_BE: 'An', TUOI: '3 tháng tuổi', TEN_NV: 'Hiền', SDT_NV: '0900 000 222', TEN_VACCINE: 'Phế cầu', TEN_CO_SO: 'Shop mẫu', DIA_CHI: 'Địa chỉ Shop mẫu', GIO_MO_CUA: '7:30–21:00', HOTLINE: '0900 000 111' }) }
  function activeTpl() { return (D.tpl || []).filter(function (x) { return x.trang_thai === 'DANG_DUNG' }) }
  function tplById(ma) { return (D.tpl || []).filter(function (x) { return x.ma === ma })[0] }
  function tplForStep(step) { return activeTpl().filter(function (x) { return x.nhom === 'KH_MOI' && x.buoc === step }) }
  /* PHÁC ĐỒ TỪ EXCEL (khu PHAC_DO): khi đã nạp và duyệt thì THAY phác đồ mẫu nhập tay. Mỗi nhóm vaccine chuẩn (6 trong 1, Phế cầu, Rota…) là một mục; nhóm "đến tuổi" khi có ÍT NHẤT MỘT dòng vaccine
   * thuộc nhóm có tuổi bắt đầu ≤ tuổi bé ≤ tuổi kết thúc (để trống = không giới hạn). Nhóm trong PD_OFF (hoặc ô "Tự gợi ý theo tuổi" = Không) không tự gợi ý nhưng vẫn tra cứu và chọn tay được. */
  var PD_OFF = ['dai', 'uon van', 'thuong han', 'ta', 'zona', 'rsv', 'khac'], PDG = { k: '', v: [] };
  function pdOn() { return libRows('PHAC_DO').length > 0 }
  function pdGroups() {
    var k = 'pd|' + LIB.at + '|' + libRows('PHAC_DO').length; if (PDG.k === k) return PDG.v;
    var by = {}, ord = []; libRows('PHAC_DO').forEach(function (r) {
      var d = r.d, nm = oneLine(d.nhom_vaccine) || shortName(d.phong_benh) || 'Khác', key = nv(nm).trim(); if (!by[key]) { by[key] = { key: key, name: nm, rows: [], keys: [] }; ord.push(key) }
      by[key].rows.push(r); var bk = brandKey(d.ten_vaccine); if (bk && by[key].keys.indexOf(bk) < 0) by[key].keys.push(bk)
    });
    var v = ord.map(function (x) { var g = by[x], fl = g.rows.map(function (r) { return fk(r.d.goi_y_tuoi || '').trim() }).filter(Boolean); g.auto = fl.length ? !fl.some(function (f) { return /^(khong|no|0|false)/.test(f) }) : PD_OFF.indexOf(g.key) < 0; return g });
    PDG = { k: k, v: v }; return v
  }
  function pdAge(r) { var tu = Number(r.d.tuoi_bat_dau_ngay), kt = r.d.tuoi_ket_thuc_ngay; kt = (kt == null || kt === '') ? null : Number(kt); return { tu: isNaN(tu) ? 0 : tu, kt: kt !== null && isNaN(kt) ? null : kt } }
  function pdRule(g, rows) {
    var a = rows.map(pdAge), tu = Math.min.apply(null, a.map(function (x) { return x.tu })), open = a.some(function (x) { return x.kt === null }), kt = open ? null : Math.max.apply(null, a.map(function (x) { return x.kt }));
    return { ma: 'pd:' + g.key, ten: g.name, ma_vaccine: '', tu_ngay: tu, den_ngay: kt, gioi_tinh: null, trang_thai: 'DANG_DUNG', phac_do: '', cau_goi_y: '', tu_khoa: '', cap_nhat_luc: String(LIB.at), _keys: g.keys, _brands: rows.map(function (r) { return r.d.ten_vaccine }) }
  }
  function pdInfo() { var on = pdOn(); return { on: on, n: libRows('PHAC_DO').length, groups: on ? pdGroups().map(function (g) { return { name: g.name, auto: g.auto, n: g.rows.length } }) : [] } }
  function allRules() { return pdOn() ? pdGroups().map(function (g) { return pdRule(g, g.rows) }) : (D.rules || []).filter(function (r) { return r.trang_thai === 'DANG_DUNG' }) }
  function dueRules(days, gender) {
    if (days == null) return [];
    if (pdOn()) return pdGroups().filter(function (g) { return g.auto }).map(function (g) { var q = g.rows.filter(function (r) { var a = pdAge(r); return days >= a.tu && (a.kt === null || days <= a.kt) }); return q.length ? pdRule(g, q) : null }).filter(Boolean);
    return (D.rules || []).filter(function (r) { return r.trang_thai === 'DANG_DUNG' && days >= r.tu_ngay && (r.den_ngay == null || days <= r.den_ngay) && (!r.gioi_tinh || r.gioi_tinh === gender) })
  }
  function soonRules(days, gender) {
    if (days == null) return [];
    if (pdOn()) return pdGroups().filter(function (g) { return g.auto }).map(function (g) {
      var now = g.rows.some(function (r) { var a = pdAge(r); return days >= a.tu && (a.kt === null || days <= a.kt) }), nx = g.rows.filter(function (r) { var a = pdAge(r); return a.tu > days && a.tu - days <= 45 }); return !now && nx.length ? pdRule(g, nx) : null
    }).filter(Boolean);
    return (D.rules || []).filter(function (r) { return r.trang_thai === 'DANG_DUNG' && days < r.tu_ngay && r.tu_ngay - days <= 45 && (!r.gioi_tinh || r.gioi_tinh === gender) })
  }

  /* ---------- hiện trong chi tiết khách ---------- */
  function counters(t) {
    var r = D.tt.filter(function (x) { return x.ma_luot === t.ma_luot && x.vong != null && !x.huy_luc }).length, m = 0, mx = (D.cfg && D.cfg.max_vong) || 3;
    D.tt.forEach(function (x) { if (x.ma_luot === t.ma_luot && isMsg(x) && !x.huy_luc) m++ });
    return '<span class="tag t-new">Lần xử lý ' + Math.min(r, mx) + '/' + mx + '</span><span class="tag t-new">Đã nhắn ' + m + '</span>'
  }

  function scriptBlock(t, link) {
    var c = D.cus[t.ma_khach] || {}, st = stepNow(t), g = activeTpl().filter(function (x) { return x.nhom === 'GOI' }), tp = g.filter(function (x) { return x.buoc === st })[0] || g[0], txt, src;
    if (tp) { txt = fill(tp.noi_dung, ctx(c, link)); src = 'Thư viện · ' + tp.ma + ' · ' + esc(tp.tieu_de) }
    else { txt = 'Chào ' + xungOf(link) + ' ' + given(link && link.ho_ten) + ', em là ' + given(ME.ho_ten) + ' ở trung tâm tiêm chủng Long Châu. Em gọi về bé ' + given(c.ho_ten) + ' (' + tuoiTxt(c) + '). Anh/chị tiện nói chuyện một chút được không ạ?'; src = 'Mẫu tạm (Thư viện chưa có lời nói khi gọi)' }
    return '<div class="sec"><h3>Gợi ý nói khi gọi</h3><div class="script">' + esc(txt) + '<small>' + src + '</small></div></div>'
  }

  /* ---------- KHO MẪU EXCEL (lib_rows đang dùng): tải lại nếu quá 30 giây ---------- */
  var LIB = { rows: null, at: 0, loading: false, err: null, cbs: [] };
  function loadLib(force) {
    if (LIB.loading || (!force && LIB.rows && Date.now() - LIB.at < 30000)) return; LIB.loading = true;
    rest('lib_rows?select=id,loai,stt,data&trang_thai=eq.DANG_DUNG&loai=in.(' + LIBT.join(',') + ')&order=loai.asc,stt.asc&limit=2000').then(function (a) {
      var by = {}; a.forEach(function (r) { (by[r.loai] = by[r.loai] || []).push({ id: r.id, stt: r.stt, d: r.data || {} }) }); LIB.rows = by; LIB.at = Date.now(); LIB.loading = false; LIB.err = null; PC = {}; afterLib()
    }, function (e) { LIB.loading = false; LIB.err = e || {}; afterLib() })
  }
  function afterLib() { if (S && host) { autoVac(); draw(true) } if (lkHost) lkDraw(); var c = LIB.cbs.splice(0); c.forEach(function (f) { try { f() } catch (e) { } }) }
  function ensureLib(cb) { if (cb) LIB.cbs.push(cb); loadLib(); if (LIB.rows && !LIB.loading) { var c = LIB.cbs.splice(0); c.forEach(function (f) { try { f() } catch (e) { } }) } }
  function isBau(loai, d) {
    if (!AN_BAU) return false; var f = loai === 'MAU_TUOI' ? (d.nhom_tuoi || '') + ' ' + (d.buoc || '') : loai === 'MAU_CHUNG' ? (d.chu_de || '') + ' ' + (d.ap_dung || '') + ' ' + (d.ten_ngan || '') : loai === 'MAU_CHU_DE' ? (d.chu_de || '') + ' ' + (d.nhom_ap_dung || '') : loai === 'MAU_SAU_TIEM' ? (d.ten_mau || '') : '';
    return /(^|[^a-z])(bau|mang thai|thai ky)([^a-z]|$)/.test(fk(f))
  }
  function libRows(loai) { return ((LIB.rows && LIB.rows[loai]) || []).filter(function (x) { return !isBau(loai, x.d) }) }
  function splitRole(s) { var m = /^\s*\[([^\]]+)\]\s*(.*)$/.exec(s || ''); return m ? { role: m[1], title: m[2] || m[1] } : { role: '', title: String(s || '') } }
  function pretty(s) { s = String(s || '').trim(); return s === s.toUpperCase() && /[A-ZÀ-Ỹ]/.test(s) && !/\d/.test(s) ? cap(s.toLowerCase()) : s }
  function nl(s) { return String(s || '').replace(/\r\n?/g, '\n').trim() }
  function oneLine(s) { return String(s || '').replace(/\s+/g, ' ').trim() }
  /* hợp tuổi bé? 'ALL' (không giới hạn) | 'FIT' | 'NO' — chỉ theo TUỔI từ ngày sinh */
  function audFit(text, days) {
    var t = fk(text), y = 365.25, m = 30.4375, lo = null, hi = null, a;
    if (days == null || !t || /tat ca/.test(t)) return 'ALL';
    if ((a = /(\d+)\s*thang\s*[-–]\s*(\d+)\s*tuoi/.exec(t))) { lo = a[1] * m; hi = a[2] * y }
    else if ((a = /(\d+)\s*[-–]\s*(\d+)\s*tuoi/.exec(t))) { lo = a[1] * y; hi = a[2] * y }
    else if ((a = />\s*(\d+)\s*tuoi/.exec(t))) { lo = a[1] * y }
    if (lo !== null) return days >= lo && (hi === null || days <= hi + y) ? 'FIT' : 'NO';
    var g = days < 9 * y ? 'NHI' : days < 18 * y ? 'TRE_LON' : 'NGUOI_LON', set = {};
    if (/so sinh|(^|[^a-z])nhi([^a-z]|$)|6in1|(^|[^a-z])lao([^a-z]|$)/.test(t)) set.NHI = 1;
    if (/tre em/.test(t)) { set.NHI = 1; set.TRE_LON = 1 } if (/thanh thieu nien|tre lon/.test(t)) set.TRE_LON = 1; if (/nguoi truong thanh|nguoi lon|cao tuoi/.test(t)) set.NGUOI_LON = 1;
    return !Object.keys(set).length ? 'ALL' : set[g] ? 'FIT' : 'NO'
  }

  /* ---------- NỐI VACCINE VỚI KHO (Hải 24/11 tối) ----------
   * Nhân viên chọn 1 vaccine (gợi ý theo tuổi bé từ PHÁC ĐỒ) → app tự tìm trong kho: câu tư vấn, nguy cơ & đường lây, CTKM còn hạn, CSKH sau tiêm của ĐÚNG vaccine đó và chọn sẵn.
   * Cách nối: tên vaccine trong phác đồ (+ "từ khóa" ASM nhập) → nhóm bệnh trong BẢNG GIÁ → các tên thương mại (Hexaxim, Prevenar…) → khớp với kho theo
   *   tư vấn: cột chủ đề/nhóm tuổi · nguy cơ: tên bệnh (hoặc vaccine áp dụng) · CTKM: tên sản phẩm · sau tiêm: tên vaccine ghi trong tiêu đề mẫu. */
  function nv(s) { var t = fk(s).replace(/\bvac\s*xin\b|\bvaccine\b|\bvacxin\b|\bvac\b/g, ' ').replace(/\b(\d)\s*(?:trong|in|\.)\s*1\b/g, '$1in1').replace(/[^a-z0-9]+/g, ' ').trim(); return ' ' + t + ' ' }
  function has(h, p) { return !!p && h.indexOf(' ' + p + ' ') >= 0 }
  function anyHas(h, L) { for (var i = 0; i < L.length; i++) if (has(h, L[i])) return true; return false }
  function phrasesOf(txt) {
    var out = []; String(txt || '').split(/[\/,;+&–—]|\s-\s|\svà\s/i).forEach(function (p) { var n = nv(p).trim(); if (!n) return; var t = n.split(' '); while (t.length > 1 && t[t.length - 1].length <= 1) t.pop(); n = t.join(' '); if (n.length >= 3 && out.indexOf(n) < 0) out.push(n) });
    return out
  }
  var GEN = ['uon', 'van', 'bach', 'hau', 'ga', 'soi', 'quai', 'thuy', 'dau', 'viem', 'gan', 'phe', 'cau', 'nao', 'mo', 'dai', 'ta', 'don'], STOP2 = ['virus', 'benh', 'cac', 'ngua', 'nhom', 'viem', 'phoi', 'khac', 'vac', 'xin', 'cua', 'cho', 'nguoi'];
  function brandKey(name) { var t = nv(name).trim().split(' '), a = t[0] || ''; if (a.length >= 3 && GEN.indexOf(a) < 0 && !/^\d/.test(a)) return a; return t.slice(0, 2).join(' ') }
  var BG = { at: -1, g: [] }, VC = {}, PC = {}, VM = {};
  function shortName(n) { return oneLine(String(n || '').replace(/\s*\([^)]*\)/g, '').replace(/^(ngừa|ngua)\s+/i, '').replace(/^các bệnh do\s+/i, '')) || oneLine(n) }
  function bgIndex() {
    if (BG.at === LIB.at && LIB.rows) return BG;
    var by = {}, ord = []; libRows('BANG_GIA').forEach(function (r) {
      var d = r.d, nm = oneLine(d.phong_benh) || 'Khác', k = nv(nm); if (!by[k]) { by[k] = { id: 'bg:' + k.trim(), name: nm, short: shortName(nm), norm: k, brands: [], bn: [], keys: [] }; ord.push(k) }
      var g = by[k]; g.brands.push(d.ten_vaccine); g.bn.push(nv(d.ten_vaccine)); var bk = brandKey(d.ten_vaccine); if (bk && g.keys.indexOf(bk) < 0) g.keys.push(bk)
    });
    BG = { at: LIB.at, g: ord.map(function (k) { return by[k] }) }; return BG
  }
  function mkVRule(r) {
    var ck = 'r:' + r.ma + '|' + LIB.at + '|' + r.cap_nhat_luc + '|' + (r.tu_khoa || '') + '|' + r.tu_ngay + '|' + r.den_ngay; if (VC[ck]) return VC[ck];
    var phr = phrasesOf(r.ten + ';' + (r.tu_khoa || '')), keys = [], groups = [];
    bgIndex().g.forEach(function (g) { if (phr.some(function (p) { return has(g.norm, p) || g.keys.indexOf(p) >= 0 || g.bn.some(function (b) { return has(b, p) }) })) { groups.push(g); g.keys.forEach(function (k) { if (keys.indexOf(k) < 0) keys.push(k) }) } });
    (r._keys || []).forEach(function (k) { if (keys.indexOf(k) < 0) keys.push(k) });
    return (VC[ck] = { id: 'rule:' + r.ma, name: r.ten, rule: r, phr: phr, keys: keys, groups: groups, from: 'RULE' })
  }
  function mkVbg(g) {
    var ck = 'b:' + g.id + '|' + LIB.at; if (VC[ck]) return VC[ck]; var base = nv(g.short).trim(), phr = [base];
    base.split(' ').forEach(function (w) { if (w.length >= 4 && STOP2.indexOf(w) < 0 && phr.indexOf(w) < 0) phr.push(w) });
    return (VC[ck] = { id: g.id, name: g.short, group: g, phr: phr, keys: g.keys.slice(), groups: [g], from: 'BG' })
  }
  function reg(V) { VM[V.id] = V; return V }
  function dueVacs() { var c = cust(), d = custDays(c); return dueRules(d, c.gioi_tinh).map(mkVRule) }
  function soonVacs() { var c = cust(), d = custDays(c); return soonRules(d, c.gioi_tinh).map(mkVRule) }
  function ageOf(V) { var r = V.rule; return r ? 'từ ' + dayLabel(r.tu_ngay) + (r.den_ngay != null ? ' đến ' + dayLabel(r.den_ngay) : ' trở lên') : 'theo Bảng giá' }

  var TABS_BY = {};
  function mkItem(tab, loai, r) {
    var d = r.d, o = { tab: tab, loai: loai, id: r.id, stt: r.stt, text: nl(d.noi_dung), ord: Number(d.thu_tu) || r.stt, key: tab + ':' + r.id }, s;
    if (loai === 'MAU_CHU_DE') { s = splitRole(d.ten_mau); o.ma = 'CD-' + r.stt; o.role = s.role; o.ten = s.title; o.group = d.chu_de || 'Khác'; o.aud = d.nhom_ap_dung || ''; o.audFit = o.aud; o.topic = d.chu_de || '' }
    else if (loai === 'MAU_TUOI') { s = splitRole(d.ten_mau); o.ma = 'NHI-' + r.stt; o.role = s.role; o.ten = s.title; o.group = pretty(d.nhom_tuoi) || 'Khác'; o.aud = d.buoc || ''; o.audFit = (d.nhom_tuoi || '') + ' ' + (d.buoc || ''); o.topic = d.nhom_tuoi || '' }
    else if (loai === 'MAU_CHUNG') { o.ma = d.ma_tin || 'TN-' + r.stt; o.ten = d.ten_ngan || d.ma_tin || 'Mẫu chung'; o.role = ''; o.group = d.chu_de || 'Chung'; o.aud = d.ap_dung || ''; o.audFit = o.aud; o.topic = (d.ten_ngan || '') + ' ' + (d.chu_de || '') }
    else { var p = String(d.ten_mau || '').split(/\s+-\s+/), lt = /^([A-Z])\.\s*/.exec(p.slice(2).join(' - ')); o.ma = p[0] || 'ST-' + r.stt; o.role = ''; o.vac = (p[1] && !/^tat ca/.test(fk(p[1]))) ? p[1] : ''; o.aud = ''; o.audFit = ''; o.letter = lt ? lt[1] : ''; o.group = (p.slice(2).join(' - ') || 'Sau tiêm').replace(/^[A-Z]\.\s*/, ''); o.ten = o.vac ? shortName(o.vac) + ' · ' + o.ma : o.group + ' · ' + o.ma; o.topic = '' }
    o.fit = audFit(o.audFit, custDays(cust())); return o
  }
  function itemsOf(tab) { var cfg = TABS.filter(function (t) { return t[0] === tab })[0]; return cfg ? libRows(cfg[2]).map(function (r) { return mkItem(tab, cfg[2], r) }) : [] }
  function tabName(tab) { var cfg = TABS.filter(function (t) { return t[0] === tab })[0]; return cfg ? cfg[1] : '' }
  function findItem(key) { var i = key.indexOf(':'), tab = key.slice(0, i), id = Number(key.slice(i + 1)); return itemsOf(tab).filter(function (o) { return o.id === id })[0] }
  function khoCount() { return TABS.reduce(function (n, t) { return n + libRows(t[2]).length }, 0) }
  function pickQT(tp) { return tp ? { src: 'QT', ma: tp.ma, ten: tp.tieu_de, text: tp.noi_dung, tag: 'Mẫu quy trình', role: '', fit: '' } : null }
  function pickK(o) { return { src: 'KHO', ma: o.ma, ten: o.ten, text: o.text, tag: tabName(o.tab), role: o.role, fit: o.fit, aud: o.aud, key: o.key } }

  /* ---------- CTKM hiện hành (kho CTKM, còn hạn theo ngày) và Nguy cơ ---------- */
  function ctkmGroups() {
    var t = todayIso(), by = {}, order = [];
    libRows('CTKM').forEach(function (r) {
      var d = r.d; if ((d.tu_ngay && d.tu_ngay > t) || (d.den_ngay && d.den_ngay < t)) return; var k = d.nhom || ('R' + r.id);
      if (!by[k]) { by[k] = { nhom: k, loai: d.phan_loai || '', chung: d.phan_loai === 'Ưu đãi chung', title: d.noi_dung_ct || d.ten_chuong_trinh || k, items: [], den: d.den_ngay || '', note: '', vac: d.loai_vaccine || '' }; order.push(k) }
      by[k].items.push(d); if (d.den_ngay && d.den_ngay > by[k].den) by[k].den = d.den_ngay
    });
    return order.map(function (k) {
      var g = by[k]; g.note = ''; g.inline = false;
      if (!g.chung) { var ns = []; g.items.forEach(function (d) { var n = oneLine(d.ghi_chu); if (n && ns.indexOf(n) < 0) ns.push(n) }); if (ns.length === 1 && g.items.every(function (d) { return oneLine(d.ghi_chu) === ns[0] })) g.note = ns[0]; else g.inline = ns.length > 0 }
      return g
    }).sort(function (a, b) { return (a.chung - b.chung) || String(a.nhom).localeCompare(String(b.nhom), 'vi', { numeric: true }) })
  }
  function itemLine(d, chung, inline) {
    if (chung) return '- Tặng ' + oneLine(d.ten_sp);
    var pay = d.gia_thanh_toan != null ? d.gia_thanh_toan : d.gia_sau_km != null ? d.gia_sau_km : null, orig = d.gia_niem_yet != null ? d.gia_niem_yet : d.gia_goc, n = inline ? oneLine(d.ghi_chu) : '';
    return '- ' + oneLine(d.ten_sp) + (d.sl_mui > 1 ? ' (' + d.sl_mui + ' mũi)' : '') + (pay != null ? ': ' + (pay === 0 ? 'MIỄN PHÍ' : money(pay)) + (orig && orig > pay ? ' (giá gốc ' + money(orig) + ')' : '') : '') + (n ? ' [' + n + ']' : '')
  }
  function ctkmLines(g) { return g.items.map(function (d) { return itemLine(d, g.chung, g.inline) }).join('\n') + (g.den ? '\n(Áp dụng đến ' + dmy(g.den) + (g.note ? '. ' + g.note : '') + ')' : '') }
  function ctkmText(g) { return (g.chung ? 'Quà tặng theo hóa đơn - ' : 'Ưu đãi: ') + g.title + '\n' + ctkmLines(g) }
  function ncList() { return libRows('NGUY_CO') }
  function ncText(d) { var e = function (s) { return String(s || '').trim().replace(/[.;]+$/, '') }, s = e(d.benh) + ': ' + e(d.nguy_co) + '.'; if (d.lay_qua) s += ' Lây qua: ' + e(d.lay_qua) + '.'; if (d.diem_can_nho) s += ' Cần nhớ: ' + e(d.diem_can_nho) + '.'; return s }

  /* Hồ sơ nối của 1 vaccine: những gì kho có cho đúng vaccine đó */
  function profile(V) {
    var k = V.id + '|' + LIB.at + '|' + (V.rule ? V.rule.cap_nhat_luc : '') + '|' + todayIso() + '|' + (S && S.t ? S.t.ma_khach : ''); if (PC[k]) return PC[k];
    var phr = V.phr, keys = V.keys, P = { tv: [], nc: [], km: [], st: [], stGen: [] };
    [['CD', 0], ['NHI', 1], ['CH', 2]].forEach(function (t) { itemsOf(t[0]).forEach(function (o) { if (o.topic && anyHas(nv(o.topic), phr)) { o.pri = t[1] + (/ra soat/.test(fk(o.topic)) ? 0.5 : 0); P.tv.push(o) } }) });
    if (!P.tv.length) itemsOf('CH').forEach(function (o) { if (anyHas(nv(o.text.slice(0, 300)), phr)) { o.pri = 3; P.tv.push(o) } });
    P.tv.sort(function (a, b) { return a.pri - b.pri || (a.fit === 'NO') - (b.fit === 'NO') || a.ord - b.ord });
    var rows = ncList(), direct = rows.filter(function (r) { return anyHas(nv(r.d.benh), phr) });
    P.nc = direct.length ? direct : rows.filter(function (r) { var n = nv(r.d.vaccine_ap_dung); return anyHas(n, keys) || anyHas(n, phr) });
    P.km = ctkmGroups().filter(function (g) { return !g.chung && g.items.some(function (d) { var n = nv(d.ten_sp); return anyHas(n, keys) || anyHas(n, phr) }) });
    var st = itemsOf('ST'); P.st = st.filter(function (o) { return o.letter === 'C' && o.vac && (anyHas(nv(o.vac), phr) || anyHas(nv(o.vac), keys)) }); P.stGen = st.filter(function (o) { return o.letter === 'A' });
    return (PC[k] = P)
  }
  function linkReport() { return allRules().map(function (r) { var V = mkVRule(r), P = profile(V); return { rule: r, name: r.ten, groups: V.groups.map(function (g) { return g.short }), keys: V.keys, tv: P.tv.length, nc: P.nc.length, km: P.km.length, st: P.st.length, tvT: P.tv.slice(0, 3).map(function (o) { return o.ten }), ncT: P.nc.map(function (x) { return x.d.benh }), stT: P.st.map(function (o) { return o.ten }) } }) }

  /* ---------- màn soạn tin ---------- */
  var S = null, host = null;
  function guard() { if (S.text !== S.base && !confirm('Anh/chị đã sửa tay nội dung. Đổi lựa chọn sẽ tạo lại tin theo mẫu và mất phần đã sửa. Tiếp tục?')) return false; return true }
  function onKey(e) { if (e.key === 'Escape' && host) { if (S && (S.view === 'ctkm' || S.view === 'nc')) leaveCheck(); else if (S && S.view !== 'main') { S.view = 'main'; draw() } else close() } }
  function open(t, ci) {
    var links = D.links[t.ma_khach] || []; if (!links.length) { toast('Khách này chưa có người liên hệ.', 0, 1); return }
    var step = stepNow(t);
    S = { t: t, ci: Math.min(ci || 0, links.length - 1), mode: 'QT', step: step, pq: pickQT(tplForStep(step)[0]), vac: null, purpose: 'TV', pk: null, st: null, km: {}, nc: {}, sig: true, dc: false, xh: '', fv: {}, view: 'main', pt: 'CD', pg: '', pq2: '', lim: 40, pm: true, pctx: 'tv', cf: '', cm: true, nq: '', vq: '', img: {}, iq: '', it: '', text: '', base: '', sending: false, snap: null, dirty: false };
    lkLock(true); host = document.createElement('div'); host.className = 'modal mm first'; host.onclick = function (e) { if (e.target === host) close() }; document.body.appendChild(host);
    document.addEventListener('keydown', onKey); setTimeout(function () { if (host) host.classList.remove('first') }, 450);
    loadFiles(); loadLib(); autoVac(); regen()
  }
  /* Hải 25/11: mở màn soạn tin là có sẵn vaccine đầu tiên bé đã đến tuổi (một chạm để đổi hoặc bỏ). Chỉ bật nguy cơ đầu tiên, KHÔNG tự bật khuyến mãi để tin không quá dài. */
  function autoVac() {
    if (!S || S.vacAuto || S.vac || S.mode !== 'QT' || !LIB.rows || S.text !== S.base) return; S.vacAuto = true;
    var due = dueVacs(), dd = custDays(cust()); if (dd != null && dd > 60) { var nl = due.filter(function (v) { return nv(v.name).trim() !== 'lao' }); if (nl.length) due = nl }   /* Lao tiêm lúc sơ sinh: bé trên 2 tháng thì không tự chọn */
    if (!due.length) return; var V = reg(due[0]); S.vac = V; S.purpose = 'TV'; applyKit(); S.km = {}; var k = Object.keys(S.nc); S.nc = {}; if (k[0]) S.nc[k[0]] = 1;
    S.base = compose(); S.text = S.base
  }
  function close() { document.removeEventListener('keydown', onKey); if (host) { host.remove(); lkLock(false) } host = null; S = null; if (window.__afterMsg) { var f = window.__afterMsg; window.__afterMsg = null; setTimeout(f, 250) } }
  function link() { return (D.links[S.t.ma_khach] || [])[S.ci] || {} }
  function cust() { return D.cus[S && S.t ? S.t.ma_khach : ''] || {} }
  function cxS() { var e = {}; for (var k in S.fv) e[k] = S.fv[k]; if (S.vac) e.TEN_VACCINE = S.vac.name; return ctx(cust(), link(), S.xh, e) }
  function sigText(cx) { return 'Người chăm sóc: ' + strip(ME.ho_ten) + ' · ' + cx.TEN_CO_SO + (cx.SDT_NV ? '\nSĐT/Zalo: ' + cx.SDT_NV : '') }
  function sigNeeded(txt) { var d = String(ME.sdt || ''); return !(d && txt.replace(/\D/g, '').indexOf(d) >= 0) }
  function compose() {
    var cx = cxS(), parts = [], f = function (t) { return fill(nl(t), cx) };
    if (S.compact && canCompact()) {
      compactParts(S.vac).forEach(function (x) { parts.push(f(x)) });
      if (S.dc) parts.push('Địa chỉ: ' + cx.DIA_CHI + (cx.HOTLINE ? '. Hotline: ' + cx.HOTLINE : '') + '.');
      var t2 = parts.join('\n\n'); if (S.sig && sigNeeded(t2)) t2 = t2 + '\n\n' + sigText(cx); return t2
    }
    if (S.mode === 'QT' && S.pq) parts.push(f(S.pq.text));
    if (S.pk) parts.push(f(S.pk.text));
    if (S.st) parts.push(f(S.st.text));
    var nBase = parts.length;
    ncList().forEach(function (r) { if (S.nc[r.id]) parts.push(ncText(r.d)) });
    /* mẫu có sẵn dòng tiêu đề của chương trình (ví dụ "Combo bảo vệ hô hấp: Giảm 5%") thì chỉ nối các dòng giá ngay dưới tiêu đề đó, không lặp tiêu đề */
    ctkmGroups().forEach(function (g) {
      if (!S.km[g.nhom]) return; var tn = fk(g.title), host = -1; if (tn.length >= 8) for (var i = 0; i < nBase; i++) if (fk(parts[i]).indexOf(tn) >= 0) { host = i; break }
      if (host >= 0) parts[host] = parts[host] + '\n' + ctkmLines(g); else parts.push(ctkmText(g))
    });
    if (S.dc) parts.push('Địa chỉ: ' + cx.DIA_CHI + (cx.GIO_MO_CUA ? '. Giờ mở cửa: ' + cx.GIO_MO_CUA : '') + (cx.HOTLINE ? '. Hotline: ' + cx.HOTLINE : '') + '.');
    var txt = parts.join('\n\n'); if (S.sig && sigNeeded(txt)) txt = (txt ? txt + '\n\n' : '') + sigText(cx); return txt
  }

  /* ---- Tin gọn: tự tóm tắt Đường lây - Lợi ích - Nguy cơ từ kho NGUY_CO và PHAC_DO đã duyệt (không thêm kiến thức mới) ---- */
  function clip(t, n) {
    t = oneLine(t).replace(/[.;,\s]+$/, ''); if (t.length <= n) return t;
    var cl = t.match(/[^;,]+[;,]?\s*/g) || [t], acc = '';
    for (var i = 0; i < cl.length; i++) { if (acc && (acc + cl[i]).length > n) break; acc += cl[i] }
    if (acc.length > n * 1.4) { acc = acc.slice(0, Math.floor(n * 1.3)); acc = acc.slice(0, acc.lastIndexOf(' ')) }
    return acc.replace(/[.;,\s]+$/, '')
  }
  function lowF(t) { t = oneLine(t); return t ? t.charAt(0).toLowerCase() + t.slice(1) : t }
  function briefLines(V) {
    var g = V.rule ? pdGroups().filter(function (x) { return V.rule.ma === 'pd:' + x.key })[0] : null, d0 = g && g.rows[0] ? g.rows[0].d : null, P = profile(V), nc = P.nc[0] ? P.nc[0].d : null, L = [];
    if (nc && nc.lay_qua) L.push('- Đường lây: ' + clip(nc.lay_qua, 85));
    var pb = d0 && d0.phong_benh ? clip(d0.phong_benh, 85) : (nc ? clip(nc.benh, 70) : '');
    if (pb) { var mui = d0 && /^\d+$/.test(String(d0.so_mui).trim()) ? ', phác đồ ' + String(d0.so_mui).trim() + ' mũi' : ''; L.push('- Lợi ích: phòng ' + pb + mui) }
    if (nc && nc.nguy_co) L.push('- Nguy cơ nếu không tiêm: ' + lowF(clip(nc.nguy_co, 95)));
    return L
  }
  function compactParts(V) {
    var parts = ['Em chào {{XUNG_HO}} {{TEN_NGUOI_LH}} ạ, em là {{TEN_NV}} bên Tiêm chủng Long Châu. Thông tin nhanh về vaccine {{TEN_VACCINE}} cho bé {{TEN_BE}}:'], L = briefLines(V);
    if (L.length) parts.push(L.join('\n'));
    var km = profile(V).km[0]; if (km) { var tt = oneLine(km.title), m2 = tt.slice(1).indexOf(tt.slice(0, 6)); if (m2 > 10) tt = tt.slice(0, m2 + 1).trim() + ' (và các gói khác)'; if (tt.length <= 80) parts.push('Ưu đãi: ' + tt) }
    parts.push('{{XUNG_HO_HOA}} cần em tư vấn thêm thì nhắn lại em nhé ạ.'); return parts
  }
  function canCompact() { return !!(S && S.mode === 'KHO' && S.vac && S.purpose !== 'ST') }
  function toggleCompact() { if (!guard()) return; S.compact = !S.compact; regen() }
  function openFull() {
    if (!host || !S || host.querySelector('.mfull')) return; var len = S.text.length, segs = Math.ceil(len / 67), d = document.createElement('div'); d.className = 'mfull';
    d.innerHTML = '<div class="mfs"><div class="mfh"><b>Tin hoàn chỉnh sẽ gửi</b><button class="mclose" id="fx" aria-label="Đóng">' + ic('x') + '</button></div><div class="mfm">' + len + ' ký tự · SMS ≈ ' + segs + ' tin · Zalo gửi trọn một tin' + (segs > 3 && canCompact() && !S.compact ? ' <span class="mfw">Tin dài, nên tóm gọn nếu gửi SMS</span>' : '') + '</div><div class="mfb">' + esc(S.text || 'Chưa có nội dung.').replace(/\n/g, '<br>') + '</div><div class="mff">' + (canCompact() ? '<button class="abt s" id="fcp2">' + (S.compact ? 'Về tin đầy đủ' : 'Tóm gọn tin') + '</button>' : '') + '<button class="abt p" id="fok">Xong</button></div></div>';
    host.appendChild(d); var shut = function () { d.remove() }; d.querySelector('#fx').onclick = shut; d.querySelector('#fok').onclick = shut; d.onclick = function (e) { if (e.target === d) shut() };
    var c2 = d.querySelector('#fcp2'); if (c2) c2.onclick = function () { toggleCompact(); if (host && S) openFull() }
  }
  function regen() { S.base = compose(); S.text = S.base; draw() }
  /* chọn tự động theo vaccine + mục đích: tư vấn, nguy cơ, CTKM (hoặc sau tiêm) của ĐÚNG vaccine đó */
  function applyKit() {
    S.km = {}; S.nc = {}; S.pk = null; S.st = null; var V = S.vac; if (!V) return; var P = profile(V);
    if (S.purpose === 'ST') { var s = P.st[0] || P.stGen[0]; if (s) S.st = pickK(s); return }
    if (S.mode === 'KHO' && P.tv[0]) S.pk = pickK(P.tv[0]);
    P.nc.forEach(function (r) { S.nc[r.id] = 1 }); P.km.forEach(function (g) { S.km[g.nhom] = 1 })
  }
  function setVaccine(V) { if (!guard()) return; S.vac = V; if (!V) S.compact = false; S.purpose = 'TV'; S.view = 'main'; applyKit(); regen() }
  function setMode(m) {
    if (m === S.mode) return; if (!guard()) return; S.mode = m;
    if (m === 'KHO') { loadLib(); if (S.vac && S.purpose === 'TV' && !S.pk) { var P = profile(S.vac); if (P.tv[0]) S.pk = pickK(P.tv[0]) } } else { S.pk = null; if (!S.pq) S.pq = pickQT(tplForStep(S.step)[0]) }
    regen()
  }
  function lenInfo(n) { return { txt: n + ' ký tự' + (n > LONG ? ' · SMS ≈ ' + Math.ceil(n / 67) + ' tin' : ''), cls: n > MAXLEN ? 'over' : n > LONG ? 'longw' : '' } }
  function msgCount() { var m = 0; D.tt.forEach(function (x) { if (x.ma_luot === S.t.ma_luot && isMsg(x)) m++ }); return m }
  function openPH() { var m = String(S.text || '').match(/\{\{[A-Z_]+\}\}/g) || [], o = []; m.forEach(function (x) { var n = x.slice(2, -2); if (o.indexOf(n) < 0) o.push(n) }); return o }
  function draw(keep) {
    if (!host || !S) return; var b = host.querySelector('.mbody'), top = keep && b ? b.scrollTop : 0, v = S.view;
    if (v === 'pick') drawPick(); else if (v === 'ctkm') drawCtkm(); else if (v === 'nc') drawNc(); else if (v === 'vac') drawVac(); else if (v === 'img') drawImgs(); else drawMain();
    if (top) { var nb = host.querySelector('.mbody'); if (nb) nb.scrollTop = top }
  }
  function hdrBack(title, sub) { return '<div class="mhd"><div class="mtop"><button class="mclose" id="pb" aria-label="Quay lại">' + ic('back') + '</button><div class="mtt"><div class="t1">' + esc(title) + '</div>' + (sub ? '<div class="t2">' + esc(sub) + '</div>' : '') + '</div><button class="mclose" id="mx" aria-label="Đóng">' + ic('x') + '</button></div></div>' }
  function skeleton() { var h = '<div class="mlist">'; for (var i = 0; i < 3; i++) h += '<div class="mrow sk"><div><i></i><i></i></div></div>'; return h + '</div>' }
  function tcard(p, btn) {
    var cx = cxS();
    return '<div class="tcard"><div class="tctop"><div class="tt">' + esc(p.ten) + '</div>' + (btn ? '<button class="mlink" id="chg">' + btn + '</button>' : '') + '</div><div class="mtags">' + (p.tag ? '<span class="mp">' + esc(p.tag) + '</span>' : '') + (p.role ? '<span class="mp ro">' + esc(p.role) + '</span>' : '') + (p.fit === 'FIT' ? '<span class="mp gr">Hợp tuổi bé</span>' : '') + '</div><div class="pv">' + esc(fill(nl(p.text), cx).replace(/\s+/g, ' ').slice(0, 240)) + '</div></div>'
  }
  function qtBlock() {
    var now = stepNow(S.t), L = tplForStep(S.step), sel = S.pq && S.pq.ma, alts = L.filter(function (x) { return x.ma !== sel }), cx = cxS();
    var h = '<section><div class="mh">Mốc chăm sóc</div><div class="stp">' + STEPS.map(function (s) { var k = s[0]; return '<button class="' + (S.step === k ? 'on' : '') + '" data-s="' + k + '" aria-pressed="' + (S.step === k) + '"><i>' + (k === 'TRE' ? ic('clock') : STEP_DOT[k]) + (k === now ? '<u class="sg" title="Gợi ý"></u>' : '') + '</i><span>' + STEP_SUB[k] + '</span></button>' }).join('') + '</div><div class="mnote"><u class="sg2"></u><span>Gợi ý theo số ngày kể từ khi giao. Chọn tay "Trễ lịch" khi khách đã lỡ mũi.</span></div></section>';
    h += '<section><div class="mh">Mẫu quy trình</div>' + (S.pq ? tcard(S.pq, null) : '<div class="mwarn"><b>Chưa có mẫu cho mốc này.</b> Chuyển sang "Theo vaccine" hoặc tự viết bên dưới.</div>');
    if (alts.length) h += '<div class="mh mh2">Mẫu khác cùng mốc</div><div class="mlist">' + alts.slice(0, 3).map(function (a) { return '<button class="mrow" data-t="' + esc(a.ma) + '"><div><b>' + esc(a.tieu_de) + '</b><span>' + esc(fill(a.noi_dung, cx).replace(/\s+/g, ' ').slice(0, 100)) + '</span></div>' + ic('chev') + '</button>' }).join('') + '</div>';
    return h + '</section>' + vacChips()
  }
  /* QT: dải chọn vaccine gọn (tùy chọn) */
  function vacChips() {
    if (LIB.loading && !LIB.rows) return ''; var due = dueVacs(); if (S.vac && !due.some(function (v) { return v.id === S.vac.id })) due = due.concat([S.vac]);
    return '<section><div class="mh">Gắn vaccine · tự điền nguy cơ, khuyến mãi (tùy chọn)</div><div class="mchips vch">' + due.map(function (v) { reg(v); return '<button class="' + (S.vac && S.vac.id === v.id ? 'on' : '') + '" data-vac="' + esc(v.id) + '">' + esc(v.name) + '</button>' }).join('') + '<button id="vmore">Vaccine khác…</button></div>' + (S.vac ? '<div class="mnote"><span>Đang gắn ' + esc(S.vac.name) + '.</span><button class="mlink" id="vlook">Xem hồ sơ vaccine</button></div>' : '<div class="mnote"><span>Chọn vaccine bé đã đến tuổi theo phác đồ: app tự thêm nguy cơ và khuyến mãi đúng vaccine đó.</span></div>') + '</section>'
  }
  function vacCard(V, dim) {
    reg(V); var P = profile(V), sel = S.vac && S.vac.id === V.id, z = function (n, t) { return '<span class="mp ' + (n ? 'gr' : 'zero') + '">' + t + ' ' + (n || 'chưa có') + '</span>' };
    return '<button class="mcard vac ' + (sel ? 'on' : '') + '" data-vac="' + esc(V.id) + '"' + (dim ? ' style="opacity:.8"' : '') + '><div class="mc1"><b>' + esc(V.name) + '</b><span class="mck">' + ic('check') + '</span></div><div class="mtags"><span class="mp">' + esc((dim ? 'Sắp đến tuổi · ' : '') + (V.rule ? cap(ageOf(V)) : 'Theo Bảng giá')) + '</span></div><div class="mtags">' + z(P.tv.length, 'Tư vấn') + z(P.nc.length, 'Nguy cơ') + z(P.km.length, 'Khuyến mãi') + '</div></button>'
  }
  function vacSummary() {
    var V = S.vac, r = V.rule, mau = r && fk(r.phac_do).indexOf('mau') === 0;
    return '<section><div class="mh">Vaccine đang tư vấn</div><div class="tcard"><div class="tctop"><div class="tt">' + esc(V.name) + '</div><span><button class="mlink" id="vlook">Hồ sơ</button><button class="mlink" id="vchg">Đổi vaccine</button></span></div><div class="mtags"><span class="mp">' + esc(cap(ageOf(V))) + '</span>' + (mau ? '<span class="mp or">Phác đồ mẫu tạm</span>' : '') + '</div><div class="psw"><button data-pur="TV" class="' + (S.purpose === 'TV' ? 'on' : '') + '">Tư vấn tiêm</button><button data-pur="ST" class="' + (S.purpose === 'ST' ? 'on' : '') + '">CSKH sau tiêm</button></div></div></section>'
  }
  function khoBlock() {
    if (LIB.loading && !LIB.rows) return '<section><div class="mh">Vaccine bé đến tuổi tiêm</div>' + skeleton() + '</section>';
    if (LIB.err && !LIB.rows) return '<section><div class="mwarn">Không tải được kho mẫu (kiểm tra mạng). <button class="mlink" id="lretry">Thử lại</button></div></section>';
    var n = khoCount(); if (!n) return '<section><div class="mwarn"><b>Kho chưa có mẫu nào đang dùng.</b> ASM duyệt lô mẫu ở mục Thư viện thì mẫu sẽ hiện ở đây.</div></section>';
    if (S.vac) return vacSummary();
    var due = dueVacs(), soon = soonVacs(), h = '<section><div class="mh">Vaccine bé đến tuổi tiêm · theo phác đồ</div>';
    h += due.length ? '<div class="mlist">' + due.map(function (v) { return vacCard(v) }).join('') + '</div>' : '<div class="mwarn"><b>Chưa có vaccine nào tới tuổi theo phác đồ.</b> ' + (custDays(cust()) == null ? 'Bé chưa có ngày sinh. ' : '') + 'Chọn "Vaccine khác" bên dưới.</div>';
    if (soon.length) h += '<div class="mh mh2">Sắp đến tuổi (trong 45 ngày)</div><div class="mlist">' + soon.map(function (v) { return vacCard(v, true) }).join('') + '</div>';
    h += '<div class="mrow2"><button class="abt o" id="vmore">' + ic('layers') + 'Vaccine khác</button><button class="abt o" id="bro">' + ic('search') + 'Duyệt kho (' + n + ')</button></div></section>';
    if (S.pk) h += '<section><div class="mh">Mẫu đang chọn</div>' + tcard(S.pk, 'Đổi mẫu') + '</section>';
    return h
  }
  /* các dòng "tự điền theo vaccine": mỗi dòng có công tắc, chạm vào chữ để đổi lựa chọn */
  function kitRow(ico, id, title, sub, on, canTg, warn) {
    return '<div class="li"><span class="ico">' + ic(ico) + '</span><button class="lt lbtn" data-kit="' + id + '"><b>' + esc(title) + '</b><small' + (warn ? ' style="color:var(--warn)"' : '') + '>' + sub + '</small></button><span class="tg"><input type="checkbox" data-tg="' + id + '" aria-label="' + esc(title) + '" ' + (on ? 'checked ' : '') + (canTg ? '' : 'disabled') + '><i></i></span></div>'
  }
  function kitList() {
    var V = S.vac, P = V ? profile(V) : null, cx = cxS(), nm = V ? esc(V.name) : '', rows = '';
    var nSel = function (m) { return Object.keys(m).filter(function (k) { return m[k] }).length };
    var tvSub = S.pk ? esc(S.pk.ten) : V ? (P.tv.length ? P.tv.length + ' mẫu cho ' + nm + ' · chạm để chọn' : 'Kho chưa có mẫu tư vấn riêng cho ' + nm) : 'Chọn vaccine ở trên để tự điền, hoặc chạm để chọn thủ công';
    rows += kitRow('spark', 'tv', 'Câu tư vấn', tvSub, !!S.pk, !!(V && P.tv.length) || !!S.pk, V && !P.tv.length && !S.pk);
    var nc = nSel(S.nc), ncSub = nc ? ncList().filter(function (r) { return S.nc[r.id] }).slice(0, 3).map(function (r) { return esc(r.d.benh) }).join(', ') + (nc > 3 ? '…' : '') : V ? (P.nc.length ? P.nc.length + ' mục cho ' + nm : 'Kho Nguy cơ chưa có ' + nm) : 'Chọn vaccine ở trên để tự điền, hoặc chạm để chọn thủ công';
    rows += kitRow('shield', 'nc', 'Nguy cơ & đường lây', ncSub, nc > 0, !!(V && P.nc.length) || nc > 0, V && !P.nc.length && !nc);
    var km = nSel(S.km), kmSub = km ? km + ' nhóm ưu đãi đã chọn' + (V && P.km.length ? ' · ' + nm : '') : V ? (P.km.length ? P.km.length + ' nhóm ưu đãi còn hiệu lực cho ' + nm : 'Chưa có khuyến mãi còn hiệu lực cho ' + nm) : ctkmGroups().length + ' nhóm ưu đãi còn hiệu lực · chạm để chọn';
    rows += kitRow('gift', 'km', 'Khuyến mãi', kmSub, km > 0, !!(V && P.km.length) || km > 0, V && !P.km.length && !km);
    if (S.mode === 'KHO') rows += kitRow('chat', 'st', 'CSKH sau tiêm', S.st ? esc(S.st.ten) : V ? (P.st.length ? P.st.length + ' mẫu riêng cho ' + nm : 'Chưa có mẫu riêng · dùng tin chung sau tiêm') : 'Chọn vaccine ở trên để tự điền', !!S.st, !!V || !!S.st);
    var noPhone = !ME.sdt;
    rows += kitRow('user', 'sig', 'Chữ ký người chăm sóc', S.sig ? (noPhone ? 'Chưa có SĐT của anh/chị: tin sẽ thiếu số liên hệ' : esc(strip(ME.ho_ten)) + ' · ' + esc(cx.TEN_CO_SO) + ' · ' + esc(cx.SDT_NV)) : 'Đang tắt: tin không có tên nhân viên, Shop, SĐT', S.sig, true, S.sig && noPhone);
    if (noPhone) rows += '<div class="sdtbox"><input id="mysdt" type="tel" inputmode="tel" placeholder="Nhập SĐT của anh/chị, ví dụ 0912345678" autocomplete="tel" aria-label="SĐT của anh/chị"><button class="abt p" id="sdtsave" style="min-height:44px;flex:none;padding:0 16px"><b>Lưu SĐT</b></button></div>';
    rows += '<label class="li"><span class="ico">' + ic('pin') + '</span><span class="lt"><b>Địa chỉ Shop</b><small>' + esc(cx.DIA_CHI) + '</small></span><span class="tg"><input type="checkbox" id="swd" ' + (S.dc ? 'checked' : '') + '><i></i></span></label>';
    rows += '<div class="li"><span class="ico">' + ic('user') + '</span><span class="lt"><b>Xưng hô</b></span><select id="xh" aria-label="Xưng hô">' + XH.map(function (o) { return '<option value="' + esc(o[0]) + '"' + (o[0] === S.xh ? ' selected' : '') + '>' + esc(o[1]) + '</option>' }).join('') + '</select></div>';
    return '<section><div class="mh">' + (V ? 'Tự điền theo ' + nm : 'Thêm vào tin') + '</div><div class="ls">' + rows + '</div></section>'
  }
  function phCard() {
    var L = openPH(); if (!L.length) return '';
    return '<section><div class="mh">Còn chỗ trống cần điền (' + L.length + ')</div><div class="ls">' + L.map(function (n) { return '<div class="li"><span class="lt"><b>' + esc(PHL[n] || n) + '</b><input class="fvi" data-fv="' + esc(n) + '" value="' + esc(S.fv[n] || '') + '" placeholder="Nhập rồi chạm ra ngoài hoặc bấm Enter" aria-label="' + esc(PHL[n] || n) + '"></span></div>' }).join('') + '</div><div class="mnote"><span>Chưa điền thì chữ {{…}} vẫn nằm trong tin gửi khách. App sẽ hỏi lại trước khi gửi.</span></div></section>'
  }
  function drawMain() {
    var c = cust(), l = link(), links = D.links[S.t.ma_khach] || [], tel = telOf(l), edited = S.text !== S.base, len = S.text.length, over = len > MAXLEN, nm = msgCount(), nImg = Object.keys(S.img).length;
    var ini = (given(l.ho_ten) || l.quan_he || '?').charAt(0).toUpperCase(), due = dueVacs();
    var h = '<div class="msheet" role="dialog" aria-label="Soạn tin"><div class="mhd"><div class="mtop"><div class="mav">' + esc(ini) + '</div><div class="mtt"><div class="t1">Soạn tin cho ' + esc(cap(l.quan_he || 'người liên hệ')) + ' ' + esc(given(l.ho_ten)) + '</div><div class="t2">Bé ' + esc(given(c.ho_ten)) + (tuoiTxt(c) ? ' · ' + esc(tuoiTxt(c)) : '') + (tel ? ' · ' + esc(tel) : '') + '</div></div><button class="mclose" id="mx" aria-label="Đóng">' + ic('x') + '</button></div>' +
      '<div class="mseg" role="tablist"><button role="tab" aria-selected="' + (S.mode === 'QT') + '" class="' + (S.mode === 'QT' ? 'on' : '') + '" data-mode="QT">' + ic('route') + 'Theo quy trình</button><button role="tab" aria-selected="' + (S.mode === 'KHO') + '" class="' + (S.mode === 'KHO' ? 'on' : '') + '" data-mode="KHO">' + ic('layers') + 'Theo vaccine</button></div></div>' +
      '<div class="mpv"><div class="ed"><div class="eh"><b>Tin sẽ gửi</b><button class="mlink" id="mfull" title="Xem toàn bộ nội dung sẽ gửi">Xem đầy đủ</button><span class="sp"></span><span class="cnt' + (nm >= 2 ? ' warn' : '') + '" title="Số lần đã nhắn khách này">Nhắn ' + nm + '/2</span><span class="badge ' + (edited ? 'b-ed' : 'b-ok') + '" id="bd">' + (edited ? 'Đã sửa tay' : 'Đúng theo mẫu') + '</span></div><textarea id="tx" aria-label="Nội dung tin" placeholder="Chọn vaccine hoặc mẫu, hoặc tự viết nội dung tin…">' + esc(S.text) + '</textarea><div class="ef"><button class="mlink" id="rs">' + ic('reset').replace('class="i"', 'class="i" style="width:16px;height:16px;vertical-align:-3px"') + ' Khôi phục theo mẫu</button><span class="sp"></span>' + (canCompact() ? '<button class="mlink" id="cmp">' + (S.compact ? 'Tin đầy đủ' : (len > 201 ? 'Tóm gọn (tin đang dài)' : 'Tóm gọn')) + '</button>' : '') + '<span id="cc" class="' + lenInfo(len).cls + '">' + lenInfo(len).txt + '</span></div></div></div><div class="mbody">';
    if (links.length > 1) h += '<section><div class="mh">Gửi cho</div><div class="sg3">' + links.map(function (x, i) { return '<button data-ci="' + i + '" class="' + (i === S.ci ? 'on' : '') + '">' + esc(cap(x.quan_he || 'Liên hệ')) + (x.ho_ten ? ' ' + esc(given(x.ho_ten)) : '') + '</button>' }).join('') + '</div></section>';
    h += S.mode === 'QT' ? qtBlock() : khoBlock();
    h += phCard() + kitList() + imgBlock();
    var can = !!S.text.trim() && !over;
    h += '</div><div class="mft">' + (nImg ? '<button class="abt o" id="sendimg">' + ic('image') + 'Gửi kèm ' + nImg + ' ảnh · chia sẻ sang Zalo</button>' : '') + (over || len > LONG ? '<div class="mst"><span>' + (over ? 'Tin quá dài (trên ' + MAXLEN + ' ký tự), hãy rút ngắn để gửi.' : 'Tin dài: Zalo gửi được, SMS sẽ tách nhiều tin. Có thể bấm "Tóm gọn" hoặc tắt bớt phần ở mục "Tự điền" bên dưới.') + '</span></div>' : '') + '<div class="act3">' +
      '<button class="abt p" data-go="ZALO" ' + (can ? '' : 'disabled') + '>' + ic('chat') + '<span><b>Zalo</b><small>chép tin + mở Zalo</small></span></button>' +
      '<button class="abt s" data-go="SMS" ' + (can ? '' : 'disabled') + '>' + ic('sms') + '<span><b>SMS</b><small>' + (tel ? 'tới ' + esc(tel) : 'chưa có số') + '</small></span></button>' +
      '<button class="abt g" data-go="SAO_CHEP" ' + (can ? '' : 'disabled') + '>' + ic('copy') + '<span><b>Sao chép</b><small>chỉ chép tin</small></span></button></div></div></div>';
    host.innerHTML = h; bindMain();
    if (S.lastTxt !== undefined && S.lastTxt !== S.text) { var pv = host.querySelector('.mpv'); if (pv) pv.classList.add('flash') } S.lastTxt = S.text
  }
  function bindMain() {
    var q = function (s) { return host.querySelector(s) }, all = function (s) { return [].slice.call(host.querySelectorAll(s)) };
    q('#mx').onclick = close;
    all('[data-mode]').forEach(function (b) { b.onclick = function () { setMode(b.dataset.mode) } });
    all('[data-ci]').forEach(function (b) { b.onclick = function () { if (!guard()) return; S.ci = Number(b.dataset.ci); regen() } });
    all('.stp [data-s]').forEach(function (b) { b.onclick = function () { if (!guard()) return; S.step = b.dataset.s; S.pq = pickQT(tplForStep(S.step)[0]); regen() } });
    all('.mrow[data-t]').forEach(function (b) { b.onclick = function () { if (!guard()) return; S.pq = pickQT(tplById(b.dataset.t)); regen() } });
    all('[data-vac]').forEach(function (b) { b.onclick = function () { var V = VM[b.dataset.vac]; if (!V) return; if (S.vac && S.vac.id === V.id && b.closest('.vch')) { if (!guard()) return; S.vac = null; applyKit(); regen(); return } setVaccine(V) } });
    var vm = q('#vmore'); if (vm) vm.onclick = function () { S.vq = ''; S.view = 'vac'; draw() };
    var vl = q('#vlook'); if (vl) vl.onclick = function () { lookup({ q: S.vac.name }) }; var vc = q('#vchg'); if (vc) vc.onclick = function () { if (!guard()) return; S.vac = null; applyKit(); regen() };
    all('[data-pur]').forEach(function (b) { b.onclick = function () { if (S.purpose === b.dataset.pur || !guard()) return; S.purpose = b.dataset.pur; applyKit(); regen() } });
    var chg = q('#chg'); if (chg) chg.onclick = function () { openPick('tv') }; var bro = q('#bro'); if (bro) bro.onclick = function () { S.pm = false; openPick('tv') };
    var lr = q('#lretry'); if (lr) lr.onclick = function () { loadLib(true); draw(true) };
    var tx = q('#tx'); tx.oninput = function () { S.text = tx.value; var ed = S.text !== S.base, bd = q('#bd'), over = S.text.length > MAXLEN; bd.className = 'badge ' + (ed ? 'b-ed' : 'b-ok'); bd.textContent = ed ? 'Đã sửa tay' : 'Đúng theo mẫu'; var cc = q('#cc'); var li = lenInfo(S.text.length); cc.textContent = li.txt; cc.className = li.cls; var can = !!S.text.trim() && !over; all('[data-go]').forEach(function (g) { g.disabled = !can }) };
    q('#rs').onclick = function () { S.text = S.base; draw(true) };
    var mf = q('#mfull'); if (mf) mf.onclick = openFull; var cp = q('#cmp'); if (cp) cp.onclick = toggleCompact;
    /* các dòng tự điền theo vaccine */
    all('[data-kit]').forEach(function (b) { b.onclick = function () { var k = b.dataset.kit; if (k === 'tv') openPick('tv'); else if (k === 'st') openPick('st'); else if (k === 'nc') { S.nq = ''; enterCheck('nc') } else if (k === 'km') { S.cm = !!(S.vac && profile(S.vac).km.length); enterCheck('ctkm') } else if (k === 'sig') { if (!guard()) return; S.sig = !S.sig; regen() } } });
    all('[data-tg]').forEach(function (b) { b.onchange = function () { if (!guard()) { draw(true); return } toggleKit(b.dataset.tg, b.checked); regen() } });
    var sw = q('#swd'); if (sw) sw.onchange = function (e) { if (!guard()) { draw(true); return } S.dc = e.target.checked; regen() };
    q('#xh').onchange = function (e) { if (!guard()) { draw(true); return } S.xh = e.target.value; regen() };
    var sv = q('#sdtsave'); if (sv) sv.onclick = saveMyPhone; var si2 = q('#mysdt'); if (si2) si2.onkeydown = function (e) { if (e.key === 'Enter') saveMyPhone() };
    all('input[data-fv]').forEach(function (i) { i.onchange = function () { var n = i.dataset.fv, v = i.value.trim(); if (!v) return; S.fv[n] = v; var tok = '{{' + n + '}}'; S.base = S.base.split(tok).join(v); S.text = S.text.split(tok).join(v); draw(true) }; i.onkeydown = function (e) { if (e.key === 'Enter') i.blur() } });
    all('[data-img]').forEach(function (b) { b.onclick = function () { var k = b.dataset.img; S.img[k] ? delete S.img[k] : S.img[k] = 1; draw(true) } });
    var im = q('#imore'); if (im) im.onclick = function () { S.view = 'img'; draw() }; var si = q('#sendimg'); if (si) si.onclick = sendImages;
    all('[data-go]').forEach(function (b) { b.onclick = function () { send(b.dataset.go) } })
  }
  function toggleKit(k, on) {
    var V = S.vac, P = V ? profile(V) : null;
    if (k === 'tv') { if (on && P && P.tv[0]) S.pk = pickK(P.tv[0]); else S.pk = null }
    else if (k === 'nc') { if (on && P) P.nc.forEach(function (r) { S.nc[r.id] = 1 }); else S.nc = {} }
    else if (k === 'km') { if (on && P) P.km.forEach(function (g) { S.km[g.nhom] = 1 }); else S.km = {} }
    else if (k === 'st') { if (on && P) { var s = P.st[0] || P.stGen[0]; S.st = s ? pickK(s) : null } else S.st = null }
    else if (k === 'sig') S.sig = on
  }
  function saveMyPhone() {
    var i = host.querySelector('#mysdt'), b = host.querySelector('#sdtsave'); if (!i || !i.value.trim()) { toast('Nhập SĐT của anh/chị trước khi lưu.', 0, 1); return } b.disabled = true;
    rpc('u_me_save_phone', { p_sdt: i.value, p_req: rid() }).then(function (r) { ME.sdt = r.sdt || null; if (D.acc && D.acc[ME.ma_tk]) D.acc[ME.ma_tk].sdt = r.sdt || null; toast('Đã lưu SĐT. Số này nằm trong chữ ký mọi tin nhắn.'); if (S && host) { var keepEdited = S.text !== S.base; if (keepEdited) { draw(true) } else regen() } }, function (e) { if (b) b.disabled = false; fail(e) })
  }
  function openPick(ctxName) { S.pctx = ctxName; if (ctxName === 'st') S.pt = 'ST'; else if (S.pt === 'ST') S.pt = 'CD'; S.pq2 = ''; S.pg = ''; S.lim = 40; if (S.vac && ctxName !== 'x' && S.pm !== false) S.pm = true; S.view = 'pick'; draw() }

  /* ---------- chọn mẫu từ kho: tab + tìm + lọc nhóm (+ chỉ vaccine đang chọn) ---------- */
  function drawPick() {
    var ctxT = S.pctx || 'tv', tabs = ctxT === 'st' ? TABS.filter(function (t) { return t[0] === 'ST' }) : TABS.filter(function (t) { return t[0] !== 'ST' }), q = fk(S.pq2.trim());
    if (!tabs.some(function (t) { return t[0] === S.pt })) S.pt = tabs[0][0];
    var P = S.vac ? profile(S.vac) : null, matchSet = {}; if (P) P.tv.concat(P.st, P.stGen).forEach(function (o) { matchSet[o.key] = 1 });
    var all = itemsOf(S.pt), onlyM = !!(P && S.pm !== false), groups = {}, order = [];
    var base = onlyM ? all.filter(function (o) { return matchSet[o.key] && (ctxT === 'st' ? true : true) }) : all;
    if (ctxT === 'st' && onlyM) base = all.filter(function (o) { return P.st.some(function (x) { return x.key === o.key }) || P.stGen.some(function (x) { return x.key === o.key }) });
    base.forEach(function (o) { if (!groups[o.group]) { groups[o.group] = 0; order.push(o.group) } groups[o.group]++ });
    order.sort(function (a, b) { return groups[b] - groups[a] || a.localeCompare(b, 'vi') });
    var list = base.filter(function (o) { return (!S.pg || o.group === S.pg) && (!q || fk(o.ten + ' ' + o.text + ' ' + o.group).indexOf(q) >= 0) });
    list.forEach(function (o) { o.hit = !!matchSet[o.key]; o.rk = (o.hit ? 0 : 1) * 10 + (o.fit === 'NO' ? 5 : 0) });
    list.sort(function (a, b) { return a.rk - b.rk || a.ord - b.ord });
    var more = list.length > S.lim, shown = list.slice(0, S.lim), selKey = (ctxT === 'st' ? S.st : S.pk) && (ctxT === 'st' ? S.st.key : S.pk.key);
    var h = '<div class="msheet">' + hdrBack(ctxT === 'st' ? 'Mẫu CSKH sau tiêm' : 'Kho mẫu tin', (onlyM ? 'Chỉ mẫu của ' + S.vac.name + ' · ' : '') + base.length + ' mẫu') + '<div class="mbody">' +
      (tabs.length > 1 ? '<div class="mtabs" role="tablist">' + tabs.map(function (t) { var n = libRows(t[2]).length; return '<button role="tab" class="' + (S.pt === t[0] ? 'on' : '') + '" data-pt="' + t[0] + '">' + esc(t[1]) + '<em>' + n + '</em></button>' }).join('') + '</div>' : '') +
      (P ? '<div class="mchips"><button class="' + (onlyM ? 'on' : '') + '" data-pm="1">Chỉ ' + esc(S.vac.name) + '</button><button class="' + (!onlyM ? 'on' : '') + '" data-pm="0">Cả kho</button></div>' : '') +
      '<div class="msearch">' + ic('search') + '<input id="pqi" placeholder="Tìm trong mẫu…" value="' + esc(S.pq2) + '" aria-label="Tìm mẫu"></div>' +
      (order.length > 1 ? '<div class="mchips"><button class="' + (!S.pg ? 'on' : '') + '" data-pg="">Tất cả · ' + base.length + '</button>' + order.map(function (g) { return '<button class="' + (S.pg === g ? 'on' : '') + '" data-pg="' + esc(g) + '">' + esc(g) + ' · ' + groups[g] + '</button>' }).join('') + '</div>' : '') +
      (LIB.loading && !LIB.rows ? skeleton() : '<div class="mlist">' + (shown.length ? shown.map(function (o) {
        return '<button class="mcard ' + (selKey === o.ma ? 'on' : '') + '" data-k="' + esc(o.key) + '"><div class="mc1"><b>' + esc(o.ten) + '</b></div><div class="mtags">' + (o.hit ? '<span class="mp gr">Đúng vaccine</span>' : '') + (o.fit === 'FIT' ? '<span class="mp gr">Hợp tuổi bé</span>' : o.fit === 'NO' ? '<span class="mp ro">Khác nhóm tuổi</span>' : '') + (o.role ? '<span class="mp ro">' + esc(o.role) + '</span>' : '') + (o.aud && !/tat ca/.test(fk(o.aud)) ? '<span class="mp">' + esc(o.aud.slice(0, 40)) + '</span>' : '') + '</div><div class="pv2">' + esc(o.text.replace(/\s+/g, ' ').slice(0, 200)) + '</div></button>'
      }).join('') + (more ? '<button class="abt o" id="lmore">Hiện thêm (' + (list.length - S.lim) + ')</button>' : '') : '<div class="mwarn">' + (onlyM ? 'Kho chưa có mẫu cho ' + esc(S.vac.name) + '. Bấm "Cả kho" để chọn mẫu khác.' : 'Không có mẫu phù hợp. Thử bỏ bớt bộ lọc.') + '</div>') + '</div>') + '</div></div>';
    host.innerHTML = h;
    host.querySelector('#mx').onclick = close; host.querySelector('#pb').onclick = function () { S.view = 'main'; draw() };
    [].slice.call(host.querySelectorAll('[data-pt]')).forEach(function (b) { b.onclick = function () { S.pt = b.dataset.pt; S.pg = ''; S.lim = 40; draw() } });
    [].slice.call(host.querySelectorAll('[data-pm]')).forEach(function (b) { b.onclick = function () { S.pm = b.dataset.pm === '1'; S.pg = ''; S.lim = 40; draw() } });
    [].slice.call(host.querySelectorAll('[data-pg]')).forEach(function (b) { b.onclick = function () { S.pg = b.dataset.pg; S.lim = 40; draw() } });
    var pi = host.querySelector('#pqi'); pi.oninput = function () { S.pq2 = pi.value; S.lim = 40; var pos = pi.selectionStart; draw(); var n = host.querySelector('#pqi'); n.focus(); n.setSelectionRange(pos, pos) };
    var lm = host.querySelector('#lmore'); if (lm) lm.onclick = function () { S.lim += 40; draw(true) };
    [].slice.call(host.querySelectorAll('.mcard[data-k]')).forEach(function (b) { b.onclick = function () { var o = findItem(b.dataset.k); if (!o) return; if (!guard()) return; if (ctxT === 'st') { S.st = pickK(o); S.purpose = 'ST' } else { S.pk = pickK(o) } S.view = 'main'; regen() } })
  }
  /* màn chọn vaccine khác (ngoài danh sách gợi ý theo tuổi): phác đồ đầy đủ + Bảng giá */
  function drawVac() {
    var q = fk(S.vq.trim()), rules = allRules().filter(function (r) { return !q || fk(r.ten + ' ' + (r.tu_khoa || '') + ' ' + (r._brands || []).join(' ')).indexOf(q) >= 0 }).map(mkVRule), bg = bgIndex().g.filter(function (g) { return !q || fk(g.name + ' ' + g.brands.join(' ')).indexOf(q) >= 0 }).map(mkVbg);
    var h = '<div class="msheet">' + hdrBack('Chọn vaccine', 'App tự nối nguy cơ, khuyến mãi, tư vấn, sau tiêm') + '<div class="mbody"><div class="msearch">' + ic('search') + '<input id="vqi" placeholder="Tìm vaccine hoặc tên thương mại…" value="' + esc(S.vq) + '" aria-label="Tìm vaccine"></div>' +
      '<section><div class="mh">Theo phác đồ (' + rules.length + ')</div>' + (rules.length ? '<div class="mlist">' + rules.map(function (v) { return vacCard(v) }).join('') + '</div>' : '<div class="mnote"><span>Không có dòng phác đồ phù hợp.</span></div>') + '</section>' +
      '<section><div class="mh">Theo Bảng giá (' + bg.length + ')</div>' + (LIB.loading && !LIB.rows ? skeleton() : bg.length ? '<div class="mlist">' + bg.map(function (v) { return vacCard(v) }).join('') + '</div>' : '<div class="mnote"><span>Không có nhóm vaccine phù hợp.</span></div>') + '</section></div></div>';
    host.innerHTML = h; host.querySelector('#mx').onclick = close; host.querySelector('#pb').onclick = function () { S.view = 'main'; draw() };
    [].slice.call(host.querySelectorAll('[data-vac]')).forEach(function (b) { b.onclick = function () { var V = VM[b.dataset.vac]; if (V) setVaccine(V) } });
    var vi = host.querySelector('#vqi'); vi.oninput = function () { S.vq = vi.value; var pos = vi.selectionStart; draw(); var n = host.querySelector('#vqi'); n.focus(); n.setSelectionRange(pos, pos) }
  }

  /* ---------- chọn nhiều: CTKM hiện hành, Nguy cơ ---------- */
  function checkFoot(n) { return '<div class="mft"><button class="abt p" id="done" style="min-height:52px;justify-content:center">' + ic('check') + '<span><b>Xong' + (n ? ' · đã chọn ' + n : '') + '</b></span></button></div>' }
  function bindCheck(sel, key) {
    host.querySelector('#mx').onclick = close; host.querySelector('#pb').onclick = leaveCheck; host.querySelector('#done').onclick = leaveCheck;
    [].slice.call(host.querySelectorAll(sel)).forEach(function (b) { b.onclick = function () { var k = b.dataset.c; if (S[key][k]) delete S[key][k]; else S[key][k] = 1; S.dirty = true; draw(true) } })
  }
  /* mở màn chọn nhiều: chụp lại lựa chọn cũ để hoàn lại nếu nhân viên không muốn mất phần đã sửa tay */
  function enterCheck(view) { S.snap = { km: JSON.parse(JSON.stringify(S.km)), nc: JSON.parse(JSON.stringify(S.nc)) }; S.dirty = false; S.view = view; draw() }
  function leaveCheck() {
    var dirty = S.dirty; S.dirty = false;
    if (!dirty) { S.view = 'main'; draw(); return }
    if (S.text !== S.base && !confirm('Thêm hoặc bỏ nội dung sẽ tạo lại tin và mất phần đã sửa tay. Tiếp tục?')) { S.km = S.snap.km; S.nc = S.snap.nc; S.view = 'main'; draw(); return }
    S.view = 'main'; regen()
  }
  function drawCtkm() {
    var G = ctkmGroups(), P = S.vac ? profile(S.vac) : null, mset = {}; if (P) P.km.forEach(function (g) { mset[g.nhom] = 1 });
    var onlyM = !!(P && P.km.length && S.cm), cats = {}, order = []; (onlyM ? G.filter(function (g) { return mset[g.nhom] }) : G).forEach(function (g) { var k = g.chung ? 'Quà tặng' : g.loai || 'Khác'; if (!cats[k]) { cats[k] = 0; order.push(k) } cats[k]++ });
    var list = (onlyM ? G.filter(function (g) { return mset[g.nhom] }) : G).filter(function (g) { return !S.cf || (g.chung ? 'Quà tặng' : g.loai || 'Khác') === S.cf }), n = Object.keys(S.km).filter(function (k) { return S.km[k] }).length;
    list.forEach(function (g) { g.hit = !!mset[g.nhom] }); list.sort(function (a, b) { return (b.hit ? 1 : 0) - (a.hit ? 1 : 0) });
    var h = '<div class="msheet">' + hdrBack('Khuyến mãi hiện hành', G.length ? G.length + ' nhóm ưu đãi còn hiệu lực hôm nay' : 'Chưa có chương trình còn hiệu lực') + '<div class="mbody">' +
      (P ? '<div class="mchips"><button class="' + (onlyM ? 'on' : '') + '" data-cm="1">Chỉ ' + esc(S.vac.name) + ' · ' + P.km.length + '</button><button class="' + (!onlyM ? 'on' : '') + '" data-cm="0">Tất cả · ' + G.length + '</button></div>' : '') +
      (order.length > 1 ? '<div class="mchips"><button class="' + (!S.cf ? 'on' : '') + '" data-cf="">Mọi loại</button>' + order.map(function (k) { return '<button class="' + (S.cf === k ? 'on' : '') + '" data-cf="' + esc(k) + '">' + esc(k) + ' · ' + cats[k] + '</button>' }).join('') + '</div>' : '') +
      '<div class="mnote"><span>Giá và ưu đãi đã khớp Bảng giá/CTKM mới nhất do ASM duyệt. Chọn nhóm nào thì tin sẽ kèm nhóm đó.</span></div><div class="mlist">' +
      (list.length ? list.map(function (g) { return '<button class="mcard ' + (S.km[g.nhom] ? 'on' : '') + '" data-c="' + esc(g.nhom) + '"><div class="mc1"><b>' + esc(g.title) + '</b><span class="mck">' + ic('check') + '</span></div><div class="mtags"><span class="mp">' + esc(g.chung ? 'Quà tặng' : g.loai || 'Ưu đãi') + '</span>' + (g.hit ? '<span class="mp gr">Đúng vaccine</span>' : '') + '<span class="mp ro">' + g.items.length + ' mục' + (g.den ? ' · đến ' + esc(dmy(g.den)) : '') + '</span></div><div class="pv2">' + esc(g.items.slice(0, 3).map(function (d) { return oneLine(d.ten_sp) }).join(' · ')) + '</div></button>' }).join('') : '<div class="mwarn">' + (onlyM ? 'Chưa có khuyến mãi nào còn hiệu lực cho ' + esc(S.vac.name) + '.' : 'Hiện chưa có chương trình khuyến mãi nào còn hiệu lực trong kho.') + '</div>') + '</div></div>' + checkFoot(n) + '</div>';
    host.innerHTML = h; bindCheck('.mcard[data-c]', 'km');
    [].slice.call(host.querySelectorAll('[data-cf]')).forEach(function (b) { b.onclick = function () { S.cf = b.dataset.cf; draw(true) } });
    [].slice.call(host.querySelectorAll('[data-cm]')).forEach(function (b) { b.onclick = function () { S.cm = b.dataset.cm === '1'; S.cf = ''; draw(true) } })
  }
  function drawNc() {
    var P = S.vac ? profile(S.vac) : null, mset = {}; if (P) P.nc.forEach(function (r) { mset[r.id] = 1 }); var q = fk(S.nq.trim()), L = ncList();
    var n = Object.keys(S.nc).filter(function (k) { return S.nc[k] }).length;
    L = L.filter(function (r) { return !q || fk(r.d.benh + ' ' + r.d.nguy_co).indexOf(q) >= 0 }).map(function (r) { r.hit = !!mset[r.id]; return r }); L.sort(function (a, b) { return (b.hit ? 1 : 0) - (a.hit ? 1 : 0) });
    var h = '<div class="msheet">' + hdrBack('Nguy cơ & đường lây', ncList().length + ' bệnh trong kho · nội dung do ASM duyệt') + '<div class="mbody"><div class="msearch">' + ic('search') + '<input id="nqi" placeholder="Tìm bệnh…" value="' + esc(S.nq) + '" aria-label="Tìm bệnh"></div><div class="mlist">' +
      (L.length ? L.map(function (r) { return '<button class="mcard ' + (S.nc[r.id] ? 'on' : '') + '" data-c="' + r.id + '"><div class="mc1"><b>' + esc(r.d.benh) + '</b><span class="mck">' + ic('check') + '</span></div>' + (r.hit ? '<div class="mtags"><span class="mp gr">Đúng vaccine</span></div>' : '') + '<div class="pv2">' + esc(String(r.d.nguy_co || '').replace(/\s+/g, ' ').slice(0, 160)) + '</div></button>' }).join('') : '<div class="mwarn">Không có bệnh phù hợp.</div>') + '</div></div>' + checkFoot(n) + '</div>';
    host.innerHTML = h; bindCheck('.mcard[data-c]', 'nc');
    var ni = host.querySelector('#nqi'); ni.oninput = function () { S.nq = ni.value; var pos = ni.selectionStart; draw(); var m = host.querySelector('#nqi'); m.focus(); m.setSelectionRange(pos, pos) }
  }

  /* ---------- ảnh đính kèm: gợi ý theo khách và vaccine, nhân viên chọn rồi chia sẻ ---------- */
  function loadFiles() { if (D.files !== undefined || D.filesLoading) return; D.filesLoading = true; rest('lib_files?select=id,ten,thu_muc,chu_de,nhom,mime,duong_dan,anh_nho&dang_dung=eq.CO&order=thoi_gian.desc&limit=300').then(function (a) { D.files = a; D.filesLoading = false; if (S && host) draw(true) }, function () { D.files = []; D.filesLoading = false }) }
  function group() { var c = cust(), d = custDays(c); if (d == null) return ''; return d < 9 * 365 ? 'NHI' : d < 18 * 365 ? 'TRE_LON' : 'NGUOI_LON' }
  function imgScore(f, names, grp) {
    var sc = 0, fc = fk(f.chu_de); if (fc && names.some(function (n) { return n.indexOf(fc) >= 0 || fc.indexOf(n) >= 0 || n.split(/\s+/).some(function (t) { return t.length >= 3 && fc.split(/\s+/).indexOf(t) >= 0 }) })) sc += 3;
    if (f.nhom === grp) sc += 2; else if (!f.nhom || f.nhom === 'TAT_CA') sc += 1; if (f.thu_muc === 'CTKM_VOUCHER' && Object.keys(S.km).some(function (k) { return S.km[k] })) sc += 2; return sc
  }
  function imgBlock() {
    var F = D.files; if (!F || !F.length) return ''; var names = [], pk = S.mode === 'QT' ? S.pq : S.pk;
    if (S.vac) { S.vac.phr.forEach(function (p) { names.push(p) }); S.vac.keys.forEach(function (k) { names.push(k) }) } else dueVacs().forEach(function (v) { v.phr.forEach(function (p) { names.push(p) }) });
    if (pk && pk.ten) names.push(fk(pk.ten));
    var grp = group(), L = F.map(function (f) { return { f: f, s: imgScore(f, names, grp) } }).filter(function (x) { return x.s > 0 }).sort(function (a, b) { return b.s - a.s }).slice(0, 8);
    return '<section><div class="mh">Ảnh đính kèm · gợi ý theo khách và vaccine</div>' + (L.length ? '<div class="imgrow">' + L.map(function (x) { return thumbBtn(x.f) }).join('') + '</div>' : '<div class="mnote"><span>Chưa có ảnh nào hợp. Bấm "Xem tất cả ảnh".</span></div>') + '<button class="abt o" id="imore">' + ic('image') + 'Xem tất cả ảnh (' + F.length + ')</button></section>'
  }
  function thumbBtn(f) { return '<button class="ithumb ' + (S.img[f.id] ? 'on' : '') + '" data-img="' + esc(f.id) + '">' + (f.anh_nho ? '<img src="' + f.anh_nho + '" alt="" width="88" height="64" loading="lazy" decoding="async">' : '<div style="height:64px;display:grid;place-items:center;color:var(--soft)">PDF</div>') + '<span>' + esc(/^https?\b/i.test(f.ten) ? (f.chu_de || 'Ảnh') : f.ten.slice(0, 24)) + '</span>' + (S.img[f.id] ? '<b class="ick">✓</b>' : '') + '</button>' }
  function drawImgs() {
    var F = D.files || [], q = fk(S.iq), L = F.filter(function (f) { return (!S.it || f.thu_muc === S.it) && (!q || fk(f.ten + ' ' + f.chu_de).indexOf(q) >= 0) }), T = { '': 'Tất cả', VACCINE_DOI_TUONG: 'Vaccine', CAM_NANG_BAU: 'Cẩm nang bầu', CTKM_VOUCHER: 'Khuyến mãi' }, nI = Object.keys(S.img).length;
    host.innerHTML = '<div class="msheet">' + hdrBack('Chọn ảnh đính kèm', nI ? nI + ' ảnh đã chọn' : 'Chạm vào ảnh để chọn') + '<div class="mbody"><div class="mchips">' + Object.keys(T).map(function (k) { return '<button class="' + (S.it === k ? 'on' : '') + '" data-it="' + k + '">' + T[k] + '</button>' }).join('') + '</div><div class="msearch">' + ic('search') + '<input id="iqi" placeholder="Tìm ảnh…" value="' + esc(S.iq) + '" aria-label="Tìm ảnh"></div><div class="imggrid">' + (L.length ? L.map(thumbBtn).join('') : '<div class="mwarn" style="grid-column:1/-1">Không có ảnh phù hợp.</div>') + '</div></div>' + checkFoot(nI) + '</div>';
    host.querySelector('#mx').onclick = close; host.querySelector('#pb').onclick = function () { S.view = 'main'; draw() }; host.querySelector('#done').onclick = function () { S.view = 'main'; draw() };
    [].slice.call(host.querySelectorAll('[data-it]')).forEach(function (b) { b.onclick = function () { S.it = b.dataset.it; draw(true) } });
    var qi = host.querySelector('#iqi'); qi.oninput = function () { S.iq = qi.value; var p = qi.selectionStart; draw(); var n = host.querySelector('#iqi'); n.focus(); n.setSelectionRange(p, p) };
    [].slice.call(host.querySelectorAll('[data-img]')).forEach(function (b) { b.onclick = function () { var k = b.dataset.img; S.img[k] ? delete S.img[k] : S.img[k] = 1; draw(true) } })
  }
  function sendImages() {
    var F = (D.files || []).filter(function (f) { return S.img[f.id] }), text = S.text; if (!F.length) return; toast('Đang chuẩn bị ' + F.length + ' ảnh…');
    Promise.all(F.map(function (f) { return fileBlob(f).then(function (b) { var ext = f.mime === 'image/png' ? 'png' : f.mime === 'application/pdf' ? 'pdf' : f.mime === 'image/webp' ? 'webp' : 'jpg'; return new File([b], (f.ten.replace(/[^\w\- ]+/g, '').slice(0, 40) || 'anh') + '.' + ext, { type: f.mime }) }) })).then(function (files) {
      if (navigator.canShare && navigator.canShare({ files: files })) return navigator.share({ files: files, text: text }).then(function () { toast('Đã mở chia sẻ ảnh. Chọn Zalo để gửi.') }, function (e) { if (e && e.name !== 'AbortError') toast('Không chia sẻ được, đã tải ảnh về máy.', 0, 1) });
      files.forEach(function (f) { var a = document.createElement('a'); a.href = URL.createObjectURL(f); a.download = f.name; document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove() }, 800) }); toast('Máy này không chia sẻ trực tiếp được, đã tải ' + files.length + ' ảnh về để anh/chị đính kèm.')
    }, function (e) { fail(e) })
  }

  /* ---------- gửi: mỗi nút làm việc của nó, rồi ghi 1 lần "Nhắn" ---------- */
  function send(kind) {
    if (S.sending) return; var l = link(), tel = telOf(l), text = S.text.trim(), t = S.t;
    if (!text) { toast('Tin đang trống, hãy chọn vaccine, mẫu hoặc tự viết.', 0, 1); return }
    if (text.length > MAXLEN) { toast('Tin dài quá ' + MAXLEN + ' ký tự, hãy rút ngắn.', 0, 1); return }
    if (kind !== 'SAO_CHEP' && !tel) { toast('Người liên hệ này chưa có số điện thoại.', 0, 1); return }
    var ph = openPH(); if (ph.length && !confirm('Tin còn chỗ trống chưa điền: ' + ph.map(function (n) { return PHL[n] || n }).join(', ') + '.\nChữ {{…}} sẽ gửi nguyên cho khách. Vẫn gửi?')) return;
    if (S.sig && !ME.sdt && !confirm('Anh/chị chưa có SĐT trong tài khoản nên tin thiếu số liên hệ. Vẫn gửi?')) return;
    S.sending = true; var edited = S.text !== S.base, pk = S.mode === 'QT' ? S.pq : (S.pk || S.st), tpl = pk ? pk.ma : '', step = S.mode === 'QT' ? S.step : '', nlh = l.ma_nlh, vtag = S.vac ? S.vac.name : '';
    copy(text);
    if (kind === 'ZALO') window.open('https://zalo.me/' + tel, '_blank', 'noopener');
    if (kind === 'SMS') { var a = document.createElement('a'); a.href = 'sms:' + tel + (/iPhone|iPad|iPod/.test(navigator.userAgent) ? '&' : '?') + 'body=' + encodeURIComponent(text); document.body.appendChild(a); a.click(); a.remove() }
    var chan = { ZALO: 'ZALO', SMS: 'NHAN_SMS', SAO_CHEP: 'NHAN_CHEP' }[kind], loai = byMa('LOAI_TUONG_TAC', chan), prevSt = t.trang_thai, fake = null;
    if (loai) { fake = { ma_luot: t.ma_luot, ma_tk: ME.ma_tk, loai_tuong_tac: loai.ma_dm, ket_qua_lien_lac: 'DM-KL6', ghi_chu: '', d: new Date(), sdt_lien_he: tel }; D.tt.push(fake); relast(t.ma_luot); if (t.trang_thai === 'DA_GIAO') t.trang_thai = 'DANG_XU_LY' }
    var msg = { ZALO: 'Đã chép tin và mở Zalo tới ' + tel + '. Dán tin vào ô chat để gửi hoặc kết bạn.', SMS: 'Đã mở tin nhắn SMS tới ' + tel + ' (tin cũng đã được chép sẵn).', SAO_CHEP: 'Đã chép tin. Dán vào ứng dụng tin nhắn để gửi.' }[kind];
    close(); toast(msg + ' Đã ghi 1 lần Nhắn.'); render();
    rpc('u_msg_log', { p_ma_luot: t.ma_luot, p_kenh: kind, p_ma_nlh: nlh || '', p_ma_mau: tpl || '', p_buoc: step ? (vtag ? step + ' · ' + vtag : step) : (vtag ? 'Vaccine ' + vtag : ''), p_da_sua: edited, p_noi_dung: text.slice(0, 2000), p_req: rid() })
      .then(function () { return load(true) }).then(render, function (e) { if (fake) { var i = D.tt.indexOf(fake); if (i >= 0) D.tt.splice(i, 1); t.trang_thai = prevSt; relast(t.ma_luot); render() } fail(e) })
  }

  /* ---------- TỪ ĐIỂN VACCINE (Hải 24-25/11): MỘT nút tra cứu mọi thứ về vaccine, lấy từ Thư viện đã duyệt ----------
   * Tab "Tìm vaccine": gõ tên, bệnh, nước sản xuất → hồ sơ: phác đồ, chống chỉ định, nguy cơ & đường lây, giá, khuyến mãi hôm nay, mẫu tư vấn, hỏi đáp & xử lý từ chối, tin sau tiêm.
   * Tab "Theo tuổi": nhập năm sinh / ngày sinh / tuổi → những vaccine tiêm được hiện tại (theo tuổi bắt đầu, kết thúc trong Phác đồ) kèm giá và khuyến mãi.
   * Danh sách = khu "Phác đồ & từ điển vaccine" (nếu đã nạp), không thì Bảng giá. Mở từ màn soạn tin thì có nút "Dùng cho tin nhắn" để tự điền đúng vaccine đó. */
  var LK = null, lkHost = null;
  var LK_AGES = [['', 'Mọi lứa tuổi', null, null], ['A', 'Sơ sinh (dưới 1 tháng)', 0, 30], ['B', '1 – 12 tháng', 30, 365], ['C', '1 – 2 tuổi', 365, 730], ['D', '2 – 9 tuổi', 730, 3285], ['E', '9 – 18 tuổi', 3285, 6570], ['F', 'Người lớn 18 – 50 tuổi', 6570, 18263], ['G', 'Trên 50 tuổi', 18263, 40000]];
  var LK_MODES = [['nam', 'Năm sinh', 'ví dụ 2024'], ['ngay', 'Ngày sinh', 'ví dụ 15/03/2024'], ['tuoi', 'Tuổi (năm)', 'ví dụ 35'], ['thang', 'Tháng tuổi', 'ví dụ 8']];
  function lkRows() {
    var pd = libRows('PHAC_DO'); if (pd.length) return pd.map(function (r) { return { src: 'PD', id: 'pd' + r.id, name: r.d.ten_vaccine || '', benh: oneLine(r.d.phong_benh), d: r.d } });
    return libRows('BANG_GIA').map(function (r) { return { src: 'BG', id: 'bg' + r.id, name: r.d.ten_vaccine || '', benh: oneLine(r.d.phong_benh), d: r.d } })
  }
  function lkVac(row) {
    var ck = 'lk:' + row.id + '|' + LIB.at; if (VC[ck]) return VC[ck];
    var key = brandKey(row.name), groups = bgIndex().g.filter(function (g) { return g.keys.indexOf(key) >= 0 || g.bn.some(function (b) { return has(b, key) }) }), phr = phrasesOf(row.benh + ';' + (row.d.nhom_vaccine || ''));
    groups.forEach(function (g) { phrasesOf(g.short).forEach(function (p) { if (phr.indexOf(p) < 0) phr.push(p) }) });
    var htxt = nv(row.benh + ' ; ' + (row.d.nhom_vaccine || '')); ncList().forEach(function (r) { var b = nv(r.d.benh).trim(); if (b.length >= 3 && has(htxt, b) && phr.indexOf(b) < 0) phr.push(b) });   /* tên bệnh trong kho Nguy cơ nằm trong mô tả bệnh của vaccine (ví dụ "Não mô cầu" ⊂ "Viêm màng não mô cầu B+C") */
    return (VC[ck] = { id: 'lk:' + row.id, name: row.name, phr: phr, keys: [key], groups: groups, from: 'LK', group: groups[0] ? groups[0].short : shortName(row.benh) })
  }
  function lkGroup(row) { var V = lkVac(row); return V.group || shortName(row.benh) || 'Khác' }
  function lkPrice(row) { var key = brandKey(row.name); return libRows('BANG_GIA').filter(function (r) { return has(nv(r.d.ten_vaccine), key) }) }
  function lkKm(row, P) { var key = brandKey(row.name), out = []; P.km.forEach(function (g) { g.items.forEach(function (d) { if (has(nv(d.ten_sp), key)) out.push({ g: g, d: d }) }) }); return out }
  function lkAgeOk(row, a) { if (!a || a[2] == null) return true; var d = row.d, tu = Number(d.tuoi_bat_dau_ngay), kt = d.tuoi_ket_thuc_ngay == null ? null : Number(d.tuoi_ket_thuc_ngay); if (isNaN(tu)) return true; return tu < a[3] && (kt == null || kt >= a[2]) }
  /* văn bản mẫu hiển thị trong từ điển: có khách đang mở thì điền sẵn tên, Shop…; không thì biến {{…}} thành [mô tả] để nhân viên tự thay */
  function lkRaw(t) { var x = nl(t); if (S && host && S.t) { try { x = fill(x, cxS()) } catch (e) { } } return x.replace(/\{\{([A-Z0-9_]+)\}\}/g, function (m, k) { return '[' + (VARDESC[k] || k) + ']' }) }
  /* tuổi từ năm sinh / ngày sinh / tuổi → khoảng ngày [nhỏ nhất, lớn nhất] */
  function lkParseAge(mode, v) {
    v = String(v || '').trim(); if (!v) return null; var now = new Date(), y = now.getFullYear(), t0 = Date.UTC(y, now.getMonth(), now.getDate()), D1 = 86400000, m, n;
    if (mode === 'nam') { n = Number(v); if (!/^\d{4}$/.test(v) || n < 1900 || n > y) return { err: 'Năm sinh gồm 4 chữ số, từ 1900 đến ' + y + '.' }; return { a: [Math.max(0, Math.round((t0 - Date.UTC(n, 11, 31)) / D1)), Math.round((t0 - Date.UTC(n, 0, 1)) / D1)], label: 'sinh năm ' + n + ' (' + (y - n) + ' tuổi)', approx: true } }
    if (mode === 'ngay') { var iso = /^(\d{4})-(\d{2})-(\d{2})$/.exec(v); if (iso) v = +iso[3] + '/' + +iso[2] + '/' + iso[1]; m = /^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})$/.exec(v); if (!m) return { err: 'Nhập ngày sinh dạng 15/03/2024.' }; var b = Date.UTC(+m[3], +m[2] - 1, +m[1]); if (isNaN(b) || new Date(b).getUTCDate() !== +m[1] || b > t0) return { err: 'Ngày sinh không hợp lệ.' }; var d = Math.round((t0 - b) / D1); return { a: [d, d], label: 'sinh ' + v + ' · ' + dayLabel(d) } }
    n = Math.floor(Number(v.replace(',', '.'))); if (isNaN(n) || n < 0 || n > 120) return { err: 'Nhập một số từ 0 đến 120.' };
    if (mode === 'tuoi') return { a: [Math.round(n * 365.25), Math.round((n + 1) * 365.25) - 1], label: 'đã tròn ' + n + ' tuổi', approx: true };
    return { a: [Math.round(n * 30.4375), Math.round((n + 1) * 30.4375) - 1], label: 'đã tròn ' + n + ' tháng', approx: true }
  }
  function lkElig(row, a) { var q = pdAge(row); if (a[0] >= q.tu && (q.kt === null || a[1] <= q.kt)) return 2; return a[1] >= q.tu && (q.kt === null || a[0] <= q.kt) ? 1 : 0 }   /* 2 = chắc chắn, 1 = tùy ngày sinh */
  function lkKmNames() { var out = []; ctkmGroups().forEach(function (g) { if (!g.chung) g.items.forEach(function (d) { out.push(nv(d.ten_sp)) }) }); return out }
  function lkPriceMin(row) { var p = lkPrice(row).map(function (r) { return Number(r.d.gia) }).filter(function (n) { return n > 0 }); return p.length ? Math.min.apply(null, p) : null }
  /* hỏi đáp & xử lý từ chối: khu HOI_DAP (nhóm vaccine = tên vaccine/nhóm/bệnh, hoặc "Chung" cho mọi vaccine) */
  function lkFaq(row, V) {
    var out = []; libRows('HOI_DAP').forEach(function (r) {
      var d = r.d, n = nv(d.nhom_vaccine || ''), gen = !String(d.nhom_vaccine || '').trim() || ['chung', 'tat ca', 'moi vaccine', 'moi loai'].some(function (p) { return has(n, p) });
      var hit = !gen && (anyHas(n, V.phr) || anyHas(n, V.keys) || V.phr.some(function (p) { return has(n, p) }) || has(n, nv(row.d.nhom_vaccine || '').trim()) || has(n, nv(row.name).trim()));
      if (gen || hit) out.push({ gen: gen, d: d })
    });
    return out.sort(function (a, b) { return (a.gen ? 1 : 0) - (b.gen ? 1 : 0) })
  }
  function lookup(opts) {
    if (lkHost) return; opts = opts || {}; LK = { tab: opts.tab || 'tim', q: opts.q || '', age: '', benh: '', lim: 30, sel: null, fromSheet: !!(S && host), am: 'nam', av: '', done: {}, doneOpen: false, cp: {} };
    lkLock(true); lkHost = document.createElement('div'); lkHost.className = 'modal mm first'; lkHost.style.zIndex = 60; lkHost.onclick = function (e) { if (e.target === lkHost) lkClose() }; document.body.appendChild(lkHost);
    document.addEventListener('keydown', lkKey); setTimeout(function () { if (lkHost) lkHost.classList.remove('first') }, 450); loadLib(); lkDraw()
  }
  var LKY = 0, LKN = 0;
  function lkLock(on) {
    var b = document.body;
    if (on) { if (LKN++ > 0) return; var sw = Math.max(0, window.innerWidth - document.documentElement.clientWidth); LKY = window.pageYOffset || 0; b.style.position = 'fixed'; b.style.top = (-LKY) + 'px'; b.style.left = '0'; b.style.right = '0'; b.style.width = '100%'; if (sw) b.style.paddingRight = sw + 'px' }
    else { if (LKN === 0 || --LKN > 0) return; b.style.position = ''; b.style.top = ''; b.style.left = ''; b.style.right = ''; b.style.width = ''; b.style.paddingRight = ''; window.scrollTo(0, LKY) }
  }
  function lkClose() { document.removeEventListener('keydown', lkKey); if (lkHost) lkHost.remove(); lkHost = null; LK = null; lkLock(false) }
  function lkKey(e) { if (e.key === 'Escape' && lkHost) { if (LK.sel) { LK.sel = null; lkDraw() } else lkClose() } }
  function lkDraw() {
    if (!lkHost || !LK) return; var b = lkHost.querySelector('.mbody'), top = b ? b.scrollTop : 0; if (LK.sel) lkDetail(); else if (LK.tab === 'tuoi') lkAgeView(); else lkList(); if (top && !LK.sel) { var nb = lkHost.querySelector('.mbody'); if (nb) nb.scrollTop = top }
  }
  function lkHead(sub) {
    return '<div class="mhd"><div class="mtop"><div class="mav">' + ic('book') + '</div><div class="mtt"><div class="t1">Từ điển vaccine</div><div class="t2">' + esc(sub) + '</div></div><button class="mclose" id="lx" aria-label="Đóng">' + ic('x') + '</button></div>' +
      '<div class="mseg" role="tablist"><button role="tab" aria-selected="' + (LK.tab === 'tim') + '" class="' + (LK.tab === 'tim' ? 'on' : '') + '" data-lt="tim">' + ic('search') + 'Tìm vaccine</button><button role="tab" aria-selected="' + (LK.tab === 'tuoi') + '" class="' + (LK.tab === 'tuoi' ? 'on' : '') + '" data-lt="tuoi">' + ic('clock') + 'Theo tuổi</button></div></div>'
  }
  function lkBindHead() {
    lkHost.querySelector('#lx').onclick = lkClose;
    [].slice.call(lkHost.querySelectorAll('[data-lt]')).forEach(function (b) { b.onclick = function () { LK.tab = b.dataset.lt; LK.lim = 30; lkDraw(); var nb = lkHost.querySelector('.mbody'); if (nb) nb.scrollTop = 0 } });
    [].slice.call(lkHost.querySelectorAll('[data-lk]')).forEach(function (b) { b.onclick = function () { LK.sel = b.dataset.lk; lkDraw(); var nb = lkHost.querySelector('.mbody'); if (nb) nb.scrollTop = 0 } })
  }
  function lkSubline(rows) { var pdOn_ = rows.length && rows[0].src === 'PD'; return LIB.rows ? rows.length + ' vaccine' + (pdOn_ ? ' · phác đồ, giá, khuyến mãi, tư vấn' : ' · từ Bảng giá (chưa nạp phác đồ)') : 'Đang tải…' }
  function lkList() {
    var rows = lkRows(), q = fk(LK.q.trim()), qt = q ? q.split(/\s+/) : [], ageSel = LK_AGES.filter(function (a) { return a[0] === LK.age })[0], groups = {}, order = [];
    rows.forEach(function (r) { var g = lkGroup(r); if (!groups[g]) { groups[g] = 0; order.push(g) } groups[g]++ }); order.sort(function (a, b) { return a.localeCompare(b, 'vi') });
    var list = rows.filter(function (r) { return (!LK.benh || lkGroup(r) === LK.benh) && lkAgeOk(r, ageSel) && (!qt.length || (function (hay) { return qt.every(function (t) { return hay.indexOf(t) >= 0 }) })(fk(r.name + ' ' + r.benh + ' ' + (r.d.nuoc_sx || '') + ' ' + (r.d.nhom_vaccine || '') + ' ' + lkGroup(r)))) });
    list.sort(function (a, b) { return lkGroup(a).localeCompare(lkGroup(b), 'vi') || a.name.localeCompare(b.name, 'vi') });
    var more = list.length > LK.lim, shown = list.slice(0, LK.lim);
    var h = '<div class="msheet">' + lkHead(lkSubline(rows)) + '<div class="mbody">' +
      '<div class="msearch">' + ic('search') + '<input id="lq" placeholder="Tên vaccine, bệnh, nước sản xuất… (ví dụ Bexsero, HPV, Bỉ)" value="' + esc(LK.q) + '" aria-label="Tìm vaccine"></div>' +
      '<div class="lkf"><select id="la" aria-label="Lứa tuổi">' + LK_AGES.map(function (a) { return '<option value="' + a[0] + '"' + (a[0] === LK.age ? ' selected' : '') + '>' + esc(a[1]) + '</option>' }).join('') + '</select><select id="lb" aria-label="Nhóm bệnh"><option value="">Mọi nhóm bệnh</option>' + order.map(function (g) { return '<option value="' + esc(g) + '"' + (g === LK.benh ? ' selected' : '') + '>' + esc(g.slice(0, 46)) + ' (' + groups[g] + ')</option>' }).join('') + '</select></div>';
    if (LIB.loading && !LIB.rows) h += skeleton();
    else if (!rows.length) h += '<div class="mwarn"><b>Chưa có dữ liệu để tra cứu.</b> ASM nạp Bảng giá hoặc "Phác đồ & từ điển vaccine" ở mục Thư viện thì danh sách sẽ hiện ở đây.</div>';
    else h += '<div class="mnote"><span>' + list.length + ' kết quả. Chạm vào vaccine để xem hồ sơ.</span></div><div class="mlist">' + (shown.length ? shown.map(function (r) { return lkCard(r) }).join('') + (more ? '<button class="abt o" id="lmore2">Hiện thêm (' + (list.length - LK.lim) + ')</button>' : '') : '<div class="mwarn">Không có vaccine phù hợp bộ lọc.</div>') + '</div>';
    h += '</div></div>'; lkHost.innerHTML = h; lkBindHead();
    var li = lkHost.querySelector('#lq'); li.oninput = function () { LK.q = li.value; LK.lim = 30; var pos = li.selectionStart; lkDraw(); var n = lkHost.querySelector('#lq'); n.focus(); n.setSelectionRange(pos, pos) };
    var sa = lkHost.querySelector('#la'); sa.onchange = function () { LK.age = sa.value; LK.lim = 30; lkDraw() }; var sb = lkHost.querySelector('#lb'); sb.onchange = function () { LK.benh = sb.value; LK.lim = 30; lkDraw() };
    var lm = lkHost.querySelector('#lmore2'); if (lm) lm.onclick = function () { LK.lim += 30; lkDraw() }
  }
  function lkCard(r, extra) {
    var d = r.d, age = d.do_tuoi || '', pm = r.src === 'PD' || r.src === 'BG' ? lkPriceMin(r) : null, tags = (d.nuoc_sx ? '<span class="mp">' + esc(d.nuoc_sx) + '</span>' : '') + (age ? '<span class="mp ro">' + esc(age.slice(0, 40)) + '</span>' : '') + (extra || '');
    return '<button class="mcard" data-lk="' + esc(r.id) + '"><div class="mc1"><b>' + esc(r.name) + '</b>' + ic('chev') + '</div><div class="mnote" style="margin-top:-2px"><span>' + esc(lkGroup(r).slice(0, 60)) + '</span></div>' + (tags ? '<div class="mtags">' + tags + '</div>' : '') + '</button>'
  }
  /* Tab "Theo tuổi": bé/khách ở tuổi này tiêm được những vaccine nào hiện tại */
  function lkAgeView() {
    var rows = lkRows(), pdOn_ = rows.length && rows[0].src === 'PD', pa = lkParseAge(LK.am, LK.av), mode = LK_MODES.filter(function (m) { return m[0] === LK.am })[0];
    var h = '<div class="msheet">' + lkHead(lkSubline(rows)) + '<div class="mbody"><div class="lkage"><select id="lam" aria-label="Cách nhập tuổi">' + LK_MODES.map(function (m) { return '<option value="' + m[0] + '"' + (m[0] === LK.am ? ' selected' : '') + '>' + esc(m[1]) + '</option>' }).join('') + '</select>' + (LK.am === 'ngay' ? '<input id="lav" type="date" min="1900-01-01" max="' + lkIso(new Date()) + '" value="' + esc(LK.av) + '" aria-label="Ngày sinh">' : '<input id="lav" inputmode="numeric" pattern="[0-9]*" enterkeyhint="done" placeholder="' + esc(mode[2]) + '" value="' + esc(LK.av) + '" aria-label="' + esc(mode[1]) + '" autocomplete="off">') + '</div>' + lkDoneBox(rows, pdOn_, pa);
    h += '<div id="lkres">' + lkAgeBody(rows, pdOn_, pa) + '</div></div></div>'; lkHost.innerHTML = h; lkBindHead();
    var am = lkHost.querySelector('#lam'); am.onchange = function () { LK.am = am.value; LK.av = ''; lkDraw(); var n = lkHost.querySelector('#lav'); if (n) n.focus() };
    LK.pa = pa; var rs = lkHost.querySelector('#lkres'); if (rs) rs.onclick = function (e) { var b = e.target.closest ? e.target.closest('[data-lk]') : null; if (b) { LK.sel = b.getAttribute('data-lk'); lkDraw(); var nb = lkHost.querySelector('.mbody'); if (nb) nb.scrollTop = 0 } };
    var av = lkHost.querySelector('#lav'), tm = null, upd = function () { LK.av = av.value; var pa2 = lkParseAge(LK.am, LK.av); LK.pa = pa2; var r = lkHost.querySelector('#lkres'); if (r) r.innerHTML = lkAgeBody(lkRows(), true, pa2); lkDoneRefresh(pa2) };
    av.oninput = function () { clearTimeout(tm); tm = setTimeout(upd, 250) }; av.onchange = function () { clearTimeout(tm); upd() };
    var dn = lkHost.querySelector('#lkdone'); if (dn) { dn.ontoggle = function () { LK.doneOpen = dn.open }; dn.onchange = function (e) { lkDoneChange(e.target, upd) } }
    if (!LK.av && !LK.sel && LK.tab === 'tuoi' && LK.am !== 'ngay') av.focus()
  }
  function lkIso(d) { return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2) }
  function lkTok(benh) { return String(benh || '').split(/[;,\/+\n]|\s-\s/).map(function (x) { return { k: nv(x).trim(), t: x.trim() } }).filter(function (x) { return x.k.length > 2 }) }
  function lkDoneSet(rows) { var o = {}; rows.forEach(function (r) { if (LK.done[r.id]) lkTok(r.benh).forEach(function (x) { o[x.k] = 1 }) }); return o }
  function lkOverlap(r, dset) { var tk = lkTok(r.benh), hit = tk.filter(function (x) { return dset[x.k] }); return hit.length ? { full: hit.length === tk.length, names: hit.map(function (x) { return x.t }) } : null }
  /* ---- số mũi & lịch mũi tiếp theo (đọc từ cột so_mui, khoang_cach_ngay của Phác đồ; chỉ tính ngày khi cách viết rõ ràng) ---- */
  function lkSoMui(v) {
    v = String(v == null ? '' : v).trim(); var m = /^(\d+)(?:\s*-\s*(\d+))?$/.exec(v);
    if (m) return { n: Math.min(6, +(m[2] || m[1])), amb: !!m[2] && m[2] !== m[1] };
    var all = [], re = /(\d+)\s*(?:mũi|liều|\()/g, x; while ((x = re.exec(v))) all.push(+x[1]);
    if (!all.length) { var f = /(\d+)/.exec(v); return { n: Math.min(6, f ? +f[1] : 1), amb: true } }
    return { n: Math.min(6, Math.max.apply(null, all)), amb: all.some(function (a) { return a !== all[0] }) }
  }
  function lkGaps(row, n) {
    var t = oneLine(row.d.khoang_cach_ngay || ''); if (n < 2 || !t) return null;
    var lab = /^(Lịch cơ bản|PĐ tiêu chuẩn|Lịch thường quy)\s*:\s*/i.exec(t); if (lab) t = t.slice(lab[0].length).split(';')[0].trim();
    else if (/[;]|hoặc|\(\s*[<>=]|\d\s*[mt]\b/.test(t)) return null;
    var m = /^(\d+(?:\s*-\s*\d+)+)\s*(tháng|ngày)/.exec(t), nums, i, g = [];
    if (m) {
      nums = m[1].split(/\s*-\s*/).map(Number);
      if (nums[0] === 0) { if (nums.length !== n) return null; for (i = 1; i < nums.length; i++) g.push([nums[i] - nums[i - 1], m[2]]); return g }
      if (nums.length !== n - 1) return null; return nums.map(function (x) { return [x, m[2]] })
    }
    m = /^(?:cách\s*)?(\d+)\s*(ngày|tháng)/.exec(t); if (!m) return null;
    for (i = 1; i < n; i++) g.push([+m[1], m[2]]); return g
  }
  function lkAddGap(d, gp) { var x = new Date(d.getTime()); if (gp[1] === 'tháng') x.setMonth(x.getMonth() + gp[0]); else x.setDate(x.getDate() + gp[0]); return x }
  function lkFmtD(d) { return ('0' + d.getDate()).slice(-2) + '/' + ('0' + (d.getMonth() + 1)).slice(-2) + '/' + d.getFullYear() }
  function lkDoseInfo(row, st) {
    var sm = lkSoMui(row.d.so_mui), n = sm.n, got = Object.keys(st.m).map(Number).filter(function (k) { return k >= 1 && k <= n }), c = got.length;
    var gaps = sm.amb ? null : lkGaps(row, n), txt = '', short = '';
    if (c >= n) { txt = '<b>Đã đủ ' + n + ' mũi</b> theo phác đồ.' + (/nhắc/i.test(row.d.khoang_cach_ngay || '') ? ' ' + esc(oneLine(row.d.khoang_cach_ngay)) : ''); short = 'đã tiêm đủ ' + n + ' mũi' }
    else {
      var nx = 1; while (st.m[nx]) nx++; short = 'đã tiêm ' + c + '/' + n + ' · còn mũi ' + nx;
      txt = '<b>Đã tiêm ' + c + '/' + n + ' mũi. Mũi tiếp theo: mũi ' + nx + '.</b>';
      if (nx === 1) txt += ' Chưa tiêm mũi nào.';
      else if (gaps && gaps[nx - 2]) {
        var gp = gaps[nx - 2], gtxt = gp[0] + ' ' + gp[1];
        if (st.d) { var ld = new Date(st.d + 'T00:00:00'), due = lkAddGap(ld, gp), late = due.getTime() < Date.now(); txt += ' Cách mũi trước ' + esc(gtxt) + ': <b>' + (late ? 'đã đến hạn từ ' : 'từ ') + esc(lkFmtD(due)) + '</b> (dự kiến, bác sĩ xác nhận).'; short += late ? ' · đã đến hạn' : ' · từ ' + lkFmtD(due) }
        else txt += ' Cách mũi trước ' + esc(gtxt) + '. Chọn ngày tiêm mũi gần nhất để tính ngày.'
      } else if (c) txt += ' Khoảng cách tùy phác đồ, xem mô tả phía trên.'
    }
    if (sm.amb) txt += ' <span class="meta">Số mũi tùy độ tuổi hoặc lịch, hãy đối chiếu phác đồ phía trên.</span>';
    return { html: txt, short: short, n: n }
  }
  function lkDosePanel(row) {
    var st = LK.done[row.id], inf = lkDoseInfo(row, st), d = row.d, n = inf.n, chips = '';
    for (var k = 1; k <= n; k++) chips += '<label class="lkchip"><input type="checkbox" data-m="' + k + '"' + (st.m[k] ? ' checked' : '') + '>Mũi ' + k + '</label>';
    return '<div class="lkdp" data-p="' + esc(row.id) + '"><div class="lkdi">' + (d.noi_dung_tu_van ? '<div>' + esc(oneLine(d.noi_dung_tu_van)) + '</div>' : '') + (d.khoang_cach_ngay ? '<div class="meta">Khoảng cách: ' + esc(oneLine(d.khoang_cach_ngay)) + '</div>' : '') + '</div>' +
      '<div class="lkch"><span class="meta">Mũi đã tiêm:</span>' + chips + '</div>' +
      '<label class="lkdt"><span class="meta">Ngày tiêm mũi gần nhất (nếu biết)</span><input type="date" data-dd max="' + lkIso(new Date()) + '" value="' + esc(st.d || '') + '"></label>' +
      '<div class="lkres2">' + inf.html + '</div></div>'
  }
  function lkDoneChange(el, upd) {
    if (!el || !el.getAttribute) return; var dn = lkHost.querySelector('#lkdone'), pn = el.closest ? el.closest('.lkdp') : null;
    if (el.hasAttribute('data-d')) {
      var id = el.getAttribute('data-d'), row = lkRows().filter(function (r) { return r.id === id })[0];
      if (el.checked && row) LK.done[id] = { m: {}, d: '' }; else delete LK.done[id];
      var bx = dn.querySelector('#lkdl'); bx.innerHTML = lkDoneList(lkRows(), LK.pa); bx.scrollTop = 0;
      var n = Object.keys(LK.done).length, sm = dn.querySelector('summary b'); if (sm) sm.textContent = 'Khách đã tiêm vaccine nào?' + (n ? ' (' + n + ')' : ''); upd(); return
    }
    if (pn && (el.hasAttribute('data-m') || el.hasAttribute('data-dd'))) {
      var pid = pn.getAttribute('data-p'), st = LK.done[pid], rw = lkRows().filter(function (r) { return r.id === pid })[0]; if (!st || !rw) return;
      if (el.hasAttribute('data-m')) { var k = +el.getAttribute('data-m'); if (el.checked) st.m[k] = 1; else delete st.m[k] } else st.d = el.value || '';
      var inf2 = lkDoseInfo(rw, st); pn.querySelector('.lkres2').innerHTML = inf2.html; var sd = dn.querySelector('.lkds[data-s="' + pid + '"]'); if (sd) sd.textContent = inf2.short; upd()
    }
  }
  function lkDoneList(rows, pa) {
    var ok = pa && !pa.err ? function (r) { return lkElig(r, pa.a) > 0 } : function () { return true };
    function item(r) { var st = LK.done[r.id]; return '<label class="lkit"><input type="checkbox" data-d="' + esc(r.id) + '"' + (st ? ' checked' : '') + '><span><b>' + esc(r.name) + '</b><small class="lkds" data-s="' + esc(r.id) + '">' + (st ? esc(lkDoseInfo(r, st).short) : '') + '</small></span></label>' + (st ? lkDosePanel(r) : '') }
    function group(list) { var by = {}, ord = []; list.forEach(function (r) { var g = oneLine(r.d.nhom_vaccine) || 'Khác'; if (!by[g]) { by[g] = []; ord.push(g) } by[g].push(r) }); return ord.map(function (g) { return '<div class="mh" style="margin-top:8px">' + esc(g) + '</div>' + by[g].map(item).join('') }).join('') }
    var sel = rows.filter(function (r) { return LK.done[r.id] }), rest = rows.filter(function (r) { return !LK.done[r.id] }), a = rest.filter(ok), b = rest.filter(function (r) { return !ok(r) });
    return (sel.length ? '<div class="mh lkpin">Đã chọn (' + sel.length + ')</div>' + sel.map(item).join('') : '') + group(a) + (b.length ? '<details class="lkt" style="margin-top:10px"><summary><b>Vaccine ngoài độ tuổi hiện tại (' + b.length + ')</b><small>Dùng khi khách đã tiêm hồi nhỏ</small></summary><div style="padding:0 10px 10px">' + group(b) + '</div></details>' : '')
  }
  function lkDoneRefresh(pa) {
    var box = lkHost && lkHost.querySelector('#lkdl'); if (!box) return; var sig = pa && !pa.err ? pa.a.join(',') : ''; if (sig === LK.dsig) return; LK.dsig = sig; box.innerHTML = lkDoneList(lkRows(), pa)
  }
  function lkDoneBox(rows, on, pa) {
    if (!on) return ''; var n = Object.keys(LK.done).length; LK.dsig = pa && !pa.err ? pa.a.join(',') : '';
    return '<details class="lkt" id="lkdone"' + (LK.doneOpen ? ' open' : '') + '><summary><b>Khách đã tiêm vaccine nào?' + (n ? ' (' + n + ')' : '') + '</b><small>Tích vaccine đã tiêm, chọn các mũi đã tiêm để loại trừ và xem mũi tiếp theo</small></summary><div id="lkdl" style="padding:0 10px 10px;max-height:340px;overflow:auto">' + lkDoneList(rows, pa) + '</div></details>'
  }
  function lkAgeBody(rows, pdOn_, pa) {
    var h = '';
    if (LIB.loading && !LIB.rows) h += skeleton();
    else if (!pdOn_) h += '<div class="mwarn"><b>Chưa nạp "Phác đồ & từ điển vaccine".</b> ASM nạp ở mục Thư viện thì mới tra được vaccine theo tuổi.</div>';
    else if (!pa) h += '<div class="mnote"><span>Nhập năm sinh, ngày sinh, tuổi hoặc tháng tuổi để xem vaccine tiêm được hiện tại, kèm giá và khuyến mãi hôm nay.</span></div>';
    else if (pa.err) h += '<div class="mwarn">' + esc(pa.err) + '</div>';
    else {
      var kmN = lkKmNames(), a = pa.a, by = {}, ord = [], offBy = {}, offOrd = [], soon = [], total = 0, offN = 0;
      var gAuto = {}; pdGroups().forEach(function (g) { gAuto[g.key] = g.auto });
      var dset = lkDoneSet(rows), doneL = rows.filter(function (r) { return LK.done[r.id] }), ovL = [];
      rows.forEach(function (r) {
        if (LK.done[r.id]) return;
        var nm = oneLine(r.d.nhom_vaccine) || shortName(r.benh) || 'Khác', key = nv(nm).trim(), auto = gAuto[key] !== false, e = lkElig(r, a);
        var ov = (e && auto) ? lkOverlap(r, dset) : null; if (ov) { ovL.push({ r: r, ov: ov }); return }
        if (e) { var tgt = auto ? by : offBy, o = auto ? ord : offOrd; if (!tgt[nm]) { tgt[nm] = []; o.push(nm) } tgt[nm].push({ r: r, e: e }); if (auto) total++; else offN++ }
        else if (auto) { var q = pdAge(r); if (q.tu > a[1] && q.tu - a[1] <= 45) soon.push(r) }
      });
      var card = function (x) {
        var pm = lkPriceMin(x.r), km = kmN.some(function (n) { return has(n, brandKey(x.r.name)) });
        return lkCard(x.r, (pm ? '<span class="mp gr">từ ' + esc(money(pm)) + '</span>' : '') + (km ? '<span class="mp or">Có khuyến mãi</span>' : '') + (x.e === 1 ? '<span class="mp zero">tùy ngày sinh</span>' : ''))
      };
      h += '<div class="tcard"><div class="tctop"><div class="tt">Khách ' + esc(pa.label) + '</div></div><div class="mtags"><span class="mp gr">' + total + ' vaccine tiêm được hiện tại</span>' + (ord.length ? '<span class="mp">' + ord.length + ' nhóm</span>' : '') + '</div></div>';
      if (pa.approx && LK.am === 'nam') h += '<div class="mnote"><span>Chỉ biết năm sinh nên có vaccine ghi "tùy ngày sinh". Nhập ngày sinh để chính xác.</span></div>';
      h += ord.length ? ord.map(function (nm) { return '<section><div class="mh">' + esc(nm) + ' (' + by[nm].length + ')</div><div class="mlist">' + by[nm].map(card).join('') + '</div></section>' }).join('') : '<div class="mwarn">Chưa có vaccine nào trong phác đồ phù hợp tuổi này.</div>';
      if (doneL.length) h += '<section><div class="mh">Khách đã tiêm (' + doneL.length + ')</div><div class="mlist">' + doneL.map(function (r) { return lkCard(r, '<span class="mp zero">' + esc(lkDoseInfo(r, LK.done[r.id]).short) + '</span>') }).join('') + '</div></section>';
      if (ovL.length) h += '<section><div class="mh">Trùng bệnh với vaccine đã tiêm (' + ovL.length + ')</div><div class="mnote"><span>Cùng phòng bệnh với vaccine khách đã tiêm. App chưa biết khách đã tiêm mấy mũi, nên bác sĩ quyết định tiêm tiếp, tiêm nhắc hay đổi loại.</span></div><div class="mlist">' + ovL.map(function (x) { return lkCard(x.r, '<span class="mp or">' + (x.ov.full ? 'Trùng toàn bộ: ' : 'Trùng một phần: ') + esc(x.ov.names.slice(0, 3).join(', ')) + '</span>') }).join('') + '</div></section>';
      if (soon.length) h += '<section><div class="mh">Sắp đến tuổi (trong 45 ngày)</div><div class="mlist">' + soon.map(function (r) { return lkCard(r) }).join('') + '</div></section>';
      if (offN) h += '<details class="lkt"><summary><b>Tiêm theo chỉ định riêng (' + offN + ')</b><small>Dại, uốn ván, thương hàn, tả, zona, RSV người lớn… không tự gợi ý theo tuổi, nhưng vẫn tra được</small></summary><div style="padding:0 10px 10px"><div class="mlist">' + offOrd.map(function (nm) { return offBy[nm].map(card).join('') }).join('') + '</div></div></details>';
      h += '<div class="mnote"><span>Dựa trên tuổi bắt đầu, kết thúc trong phác đồ đã duyệt. Bác sĩ quyết định chỉ định cuối cùng.</span></div>'
    }
    return h
  }
  function lkSec(title, body, id) { return body ? '<section' + (id ? ' id="' + id + '"' : '') + '><div class="mh">' + esc(title) + '</div>' + body + '</section>' : '' }
  function lkTpl(o, k) { var t = lkRaw(o.text); LK.cp[k] = t; return '<details class="lkt"><summary><b>' + esc(o.ten) + '</b><small>' + esc(oneLine(t).slice(0, 120)) + '…</small></summary><div class="lkx">' + esc(t).replace(/\n/g, '<br>') + '</div><button class="mlink" data-cp="' + k + '">Sao chép nội dung</button></details>' }
  function lkSummary(row, P, price, km) {
    var d = row.d, L = [row.name + ' – ' + (row.benh || '')]; if (d.do_tuoi) L.push('Độ tuổi: ' + d.do_tuoi); if (d.nuoc_sx) L.push('Nước SX: ' + d.nuoc_sx); if (d.noi_dung_tu_van) L.push('Phác đồ: ' + oneLine(d.noi_dung_tu_van));
    if (P.nc[0]) L.push('Nguy cơ: ' + oneLine(P.nc[0].d.nguy_co).slice(0, 200)); if (price.length) L.push('Giá: ' + price.map(function (r) { return money(r.d.gia) }).join(' / ')); if (km.length) L.push('Khuyến mãi: ' + km.slice(0, 3).map(function (x) { return itemLine(x.d, false, false).replace(/^- /, '') }).join('; '));
    return L.join('\n')
  }
  function lkDetail() {
    var row = lkRows().filter(function (r) { return r.id === LK.sel })[0]; if (!row) { LK.sel = null; lkDraw(); return }
    var V = lkVac(row), P = profile(V), d = row.d, price = lkPrice(row), km = lkKm(row, P), faq = lkFaq(row, V), has_ = function (v) { return v != null && String(v).trim() !== '' }; LK.cp = {};
    var tu = has_(d.tuoi_bat_dau_ngay) ? Number(d.tuoi_bat_dau_ngay) : null, kt = has_(d.tuoi_ket_thuc_ngay) ? Number(d.tuoi_ket_thuc_ngay) : null, infer = /tự suy/.test(d.ghi_chu || '');
    var age = '';
    if (d.do_tuoi) age += '<div><b>Độ tuổi chỉ định:</b> ' + esc(d.do_tuoi) + '</div>';
    if (tu != null) age += '<div><b>Từ</b> ' + esc(dayLabel(tu)) + ' tuổi' + (kt != null ? ' <b>đến</b> ' + esc(dayLabel(kt)) : ' <b>trở lên</b>') + (infer && kt != null ? ' <span class="mp or">tuổi kết thúc tự suy, chờ bác sĩ</span>' : '') + '</div>';
    if (d.noi_dung_tu_van) age += '<div style="margin-top:6px"><b>Phác đồ tham khảo:</b><br>' + esc(nl(d.noi_dung_tu_van)).replace(/\n/g, '<br>') + '</div>';
    if (d.so_mui) age += '<div><b>Số mũi:</b> ' + esc(d.so_mui) + '</div>'; if (d.khoang_cach_ngay) age += '<div><b>Khoảng cách mũi:</b> ' + esc(d.khoang_cach_ngay) + '</div>';
    var pro = ''; if (d.chong_chi_dinh) pro += '<div><b>Chống chỉ định / thận trọng:</b> ' + esc(d.chong_chi_dinh) + '</div>'; if (d.phan_ung) pro += '<div><b>Phản ứng sau tiêm thường gặp:</b> ' + esc(d.phan_ung) + '</div>'; if (d.luu_y) pro += '<div><b>Lưu ý khi tiêm:</b> ' + esc(d.luu_y) + '</div>'; if (d.nhom_vaccine) pro += '<div><b>Nhóm vaccine:</b> ' + esc(d.nhom_vaccine) + '</div>';
    var nc = P.nc.map(function (r) { var x = r.d; return '<div class="lkb"><b>' + esc(x.benh) + '</b><div>' + esc(oneLine(x.nguy_co)) + '</div>' + (x.lay_qua ? '<div><b>Lây qua:</b> ' + esc(oneLine(x.lay_qua)) + '</div>' : '') + (x.trieu_chung ? '<div><b>Triệu chứng:</b> ' + esc(oneLine(x.trieu_chung)) + '</div>' : '') + (x.mua_vung ? '<div><b>Mùa & vùng:</b> ' + esc(oneLine(x.mua_vung)) + '</div>' : '') + (x.diem_can_nho ? '<div><b>Cần nhớ:</b> ' + esc(oneLine(x.diem_can_nho)) + '</div>' : '') + '</div>' }).join('');
    var gia = price.map(function (r) { return '<div class="lkr"><span>' + esc(r.d.ten_vaccine) + '</span><b>' + esc(money(r.d.gia)) + '</b></div>' }).join('') + (price[0] && (price[0].d.tu_ngay || price[0].d.den_ngay) ? '<div class="meta">Kỳ áp dụng: ' + esc(dmy(price[0].d.tu_ngay)) + ' – ' + esc(dmy(price[0].d.den_ngay)) + '</div>' : '');
    var kmh = km.map(function (x) { return '<div class="lkb"><b>' + esc(x.g.title) + '</b><div>' + esc(itemLine(x.d, false, false).replace(/^- /, '')) + '</div>' + (x.d.ghi_chu ? '<div class="meta">' + esc(oneLine(x.d.ghi_chu)) + '</div>' : '') + '<div class="meta">Áp dụng đến ' + esc(dmy(x.g.den)) + '</div></div>' }).join('');
    var tv = P.tv.map(function (o, i) { return lkTpl(o, 'tv' + i) }).join(''), st = P.st.map(function (o, i) { return lkTpl(o, 'st' + i) }).join('');
    var refuse = faq.filter(function (x) { return /tu choi|do dap|ngai|khong muon/.test(fk(x.d.loai || '')) }), qa = faq.filter(function (x) { return refuse.indexOf(x) < 0 });
    var fqItem = function (x, k) { var t = lkRaw(x.d.tra_loi) + (x.d.luu_y ? '\n(Lưu ý: ' + nl(x.d.luu_y) + ')' : ''); LK.cp[k] = lkRaw(x.d.tra_loi); return '<details class="lkt"><summary><b>' + esc(oneLine(x.d.tinh_huong)) + '</b><small>' + (x.gen ? 'Dùng chung mọi vaccine · ' : '') + esc(oneLine(lkRaw(x.d.tra_loi)).slice(0, 100)) + '…</small></summary><div class="lkx">' + esc(t).replace(/\n/g, '<br>') + (x.d.nguon ? '<div class="meta" style="margin-top:6px">Nguồn: ' + esc(oneLine(x.d.nguon)) + '</div>' : '') + '</div><button class="mlink" data-cp="' + k + '">Sao chép câu trả lời</button></details>' };
    var hd = (refuse.length ? '<div class="mh mh2">Khi khách từ chối, do dự (' + refuse.length + ')</div>' + refuse.map(function (x, i) { return fqItem(x, 'fr' + i) }).join('') : '') + (qa.length ? '<div class="mh mh2">Khách hay hỏi (' + qa.length + ')</div>' + qa.map(function (x, i) { return fqItem(x, 'fq' + i) }).join('') : '');
    var nav = [['lks-pd', 'Phác đồ'], ['lks-gia', 'Giá & khuyến mãi'], ['lks-tv', 'Tư vấn' + (P.tv.length ? ' (' + P.tv.length + ')' : '')], ['lks-hd', 'Hỏi đáp, từ chối' + (faq.length ? ' (' + faq.length + ')' : '')], ['lks-st', 'Sau tiêm' + (P.st.length ? ' (' + P.st.length + ')' : '')]];
    var h = '<div class="msheet"><div class="mhd"><div class="mtop"><button class="mclose" id="lb2" aria-label="Quay lại">' + ic('back') + '</button><div class="mtt"><div class="t1">' + esc(row.name) + '</div><div class="t2">' + esc(lkGroup(row)) + '</div></div><button class="mclose" id="lx" aria-label="Đóng">' + ic('x') + '</button></div></div><div class="mbody">' +
      '<div class="tcard"><div class="mtags" style="margin-top:0">' + (d.nuoc_sx ? '<span class="mp">Nước SX: ' + esc(d.nuoc_sx) + '</span>' : '') + (row.benh ? '<span class="mp ro lkw">' + esc(row.benh.slice(0, 120)) + '</span>' : '') + '</div></div>' +
      '<div class="lkn">' + nav.map(function (n) { return '<button data-go="' + n[0] + '">' + esc(n[1]) + '</button>' }).join('') + '</div>' +
      lkSec('Độ tuổi & phác đồ', age ? '<div class="ls lkp">' + age + '</div>' : '<div class="mnote"><span>Chưa có tuổi, phác đồ trong kho cho vaccine này.</span></div>', 'lks-pd') + lkSec('Lưu ý chuyên môn', pro ? '<div class="ls lkp">' + pro + '</div>' : '<div class="mnote"><span>Chưa có chống chỉ định, phản ứng sau tiêm trong kho (ASM bổ sung ở Thư viện → Phác đồ & từ điển vaccine).</span></div>') +
      lkSec('Nguy cơ & đường lây', nc || '<div class="mnote"><span>Kho Nguy cơ chưa có bệnh này.</span></div>') + lkSec('Giá niêm yết', gia || '<div class="mnote"><span>Bảng giá chưa có vaccine này.</span></div>', 'lks-gia') + lkSec('Khuyến mãi còn hiệu lực hôm nay', kmh || '<div class="mnote"><span>Hôm nay chưa có khuyến mãi cho vaccine này.</span></div>') +
      lkSec('Mẫu tư vấn (' + P.tv.length + ')', tv || '<div class="mnote"><span>Kho chưa có mẫu tư vấn riêng cho vaccine này.</span></div>', 'lks-tv') +
      lkSec('Hỏi đáp & xử lý từ chối', hd || '<div class="mnote"><span>Kho chưa có hỏi đáp, tình huống từ chối cho vaccine này. ASM nạp ở Thư viện → "Hỏi đáp & xử lý từ chối".</span></div>', 'lks-hd') +
      lkSec('Tin CSKH sau tiêm (' + P.st.length + ')', st || '<div class="mnote"><span>Kho chưa có tin sau tiêm riêng cho vaccine này.</span></div>', 'lks-st') +
      '<div class="mnote"><span>Nội dung do ASM duyệt trong Thư viện. App chỉ tra cứu và gợi ý, không thay chỉ định của bác sĩ.</span></div></div><div class="mft"><div class="act2">' + (LK.fromSheet ? '<button class="abt p" id="luse">' + ic('chat') + '<span><b>Dùng cho tin nhắn</b><small>tự điền đúng vaccine này</small></span></button>' : '') + '<button class="abt g" id="lcopy">' + ic('copy') + '<span><b>Sao chép tóm tắt</b><small>gửi cho đồng nghiệp</small></span></button></div></div></div>';
    lkHost.innerHTML = h; lkHost.querySelector('#lx').onclick = lkClose; lkHost.querySelector('#lb2').onclick = function () { LK.sel = null; lkDraw() };
    var cp = lkHost.querySelector('#lcopy'); if (cp) cp.onclick = function () { copy(lkSummary(row, P, price, km)); toast('Đã chép tóm tắt ' + row.name) };
    [].slice.call(lkHost.querySelectorAll('[data-cp]')).forEach(function (b) { b.onclick = function () { copy(LK.cp[b.dataset.cp] || ''); toast('Đã sao chép') } });
    [].slice.call(lkHost.querySelectorAll('[data-go]')).forEach(function (b) { b.onclick = function () { var el = lkHost.querySelector('#' + b.dataset.go); if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' }) } });
    var us = lkHost.querySelector('#luse'); if (us) us.onclick = function () { lkClose(); if (S && host) setVaccine(lkVac(row)) }
  }

  return { open: open, counters: counters, scriptBlock: scriptBlock, preview: preview, fill: fill, dueRules: dueRules, soonRules: soonRules, pdInfo: pdInfo, custDays: custDays, dayLabel: dayLabel, VARS: VARS, VARDESC: VARDESC, STEPS: STEPS, GROUPS: GROUPS, ensureLib: ensureLib, linkReport: linkReport, lookup: lookup, _t: { briefLines: briefLines, audFit: audFit, ctkmGroups: ctkmGroups, ctkmText: ctkmText, ncText: ncText, itemsOf: itemsOf, libRows: libRows, isBau: isBau, LIB: LIB, nv: nv, phrasesOf: phrasesOf, brandKey: brandKey, mkVRule: mkVRule, mkVbg: mkVbg, profile: profile, bgIndex: bgIndex, linkReport: linkReport } }
})();

;

/* ===== icons.js ===== */
/* Biểu tượng Lucide v1.80.0 (lucide.dev, giấy phép ISC). Ghi đè bộ IC; nạp SAU các module khác. */
(function () {
  var L = {
 "more": "<circle cx=\"12\" cy=\"12\" r=\"1\" /><circle cx=\"19\" cy=\"12\" r=\"1\" /><circle cx=\"5\" cy=\"12\" r=\"1\" />",
 "home": "<path d=\"M15 21v-8a1 1 0 0 0-1-1h-4a1 1 0 0 0-1 1v8\" /><path d=\"M3 10a2 2 0 0 1 .709-1.528l7-6a2 2 0 0 1 2.582 0l7 6A2 2 0 0 1 21 10v9a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z\" />",
 "list": "<path d=\"M3 5h.01\" /><path d=\"M3 12h.01\" /><path d=\"M3 19h.01\" /><path d=\"M8 5h13\" /><path d=\"M8 12h13\" /><path d=\"M8 19h13\" />",
 "clock": "<path d=\"M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8\" /><path d=\"M3 3v5h5\" /><path d=\"M12 7v5l4 2\" />",
 "chart": "<path d=\"M3 3v16a2 2 0 0 0 2 2h16\" /><path d=\"M18 17V9\" /><path d=\"M13 17V5\" /><path d=\"M8 17v-3\" />",
 "user": "<path d=\"M19 21v-2a4 4 0 0 0-4-4H9a4 4 0 0 0-4 4v2\" /><circle cx=\"12\" cy=\"7\" r=\"4\" />",
 "team": "<path d=\"M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2\" /><path d=\"M16 3.128a4 4 0 0 1 0 7.744\" /><path d=\"M22 21v-2a4 4 0 0 0-3-3.87\" /><circle cx=\"9\" cy=\"7\" r=\"4\" />",
 "call": "<path d=\"M13.832 16.568a1 1 0 0 0 1.213-.303l.355-.465A2 2 0 0 1 17 15h3a2 2 0 0 1 2 2v3a2 2 0 0 1-2 2A18 18 0 0 1 2 4a2 2 0 0 1 2-2h3a2 2 0 0 1 2 2v3a2 2 0 0 1-.8 1.6l-.468.351a1 1 0 0 0-.292 1.233 14 14 0 0 0 6.392 6.384\" />",
 "msg": "<path d=\"M22 17a2 2 0 0 1-2 2H6.828a2 2 0 0 0-1.414.586l-2.202 2.202A.71.71 0 0 1 2 21.286V5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2z\" />",
 "back": "<path d=\"m15 18-6-6 6-6\" />",
 "copy": "<rect width=\"14\" height=\"14\" x=\"8\" y=\"8\" rx=\"2\" ry=\"2\" /><path d=\"M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2\" />",
 "search": "<path d=\"m21 21-4.34-4.34\" /><circle cx=\"11\" cy=\"11\" r=\"8\" />",
 "x": "<path d=\"M18 6 6 18\" /><path d=\"m6 6 12 12\" />",
 "chat": "<path d=\"M2.992 16.342a2 2 0 0 1 .094 1.167l-1.065 3.29a1 1 0 0 0 1.236 1.168l3.413-.998a2 2 0 0 1 1.099.092 10 10 0 1 0-4.777-4.719\" />",
 "sms": "<path d=\"M22 17a2 2 0 0 1-2 2H6.828a2 2 0 0 0-1.414.586l-2.202 2.202A.71.71 0 0 1 2 21.286V5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2z\" /><path d=\"M7 11h10\" /><path d=\"M7 15h6\" /><path d=\"M7 7h8\" />",
 "gift": "<path d=\"M12 7v14\" /><path d=\"M20 11v8a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2v-8\" /><path d=\"M7.5 7a1 1 0 0 1 0-5A4.8 8 0 0 1 12 7a4.8 8 0 0 1 4.5-5 1 1 0 0 1 0 5\" /><rect x=\"3\" y=\"7\" width=\"18\" height=\"4\" rx=\"1\" />",
 "shield": "<path d=\"M20 13c0 5-3.5 7.5-7.66 8.95a1 1 0 0 1-.67-.01C7.5 20.5 4 18 4 13V6a1 1 0 0 1 1-1c2 0 4.5-1.2 6.24-2.72a1.17 1.17 0 0 1 1.52 0C14.51 3.81 17 5 19 5a1 1 0 0 1 1 1z\" /><path d=\"m9 12 2 2 4-4\" />",
 "spark": "<path d=\"M11.017 2.814a1 1 0 0 1 1.966 0l1.051 5.558a2 2 0 0 0 1.594 1.594l5.558 1.051a1 1 0 0 1 0 1.966l-5.558 1.051a2 2 0 0 0-1.594 1.594l-1.051 5.558a1 1 0 0 1-1.966 0l-1.051-5.558a2 2 0 0 0-1.594-1.594l-5.558-1.051a1 1 0 0 1 0-1.966l5.558-1.051a2 2 0 0 0 1.594-1.594z\" /><path d=\"M20 2v4\" /><path d=\"M22 4h-4\" /><circle cx=\"4\" cy=\"20\" r=\"2\" />",
 "pin": "<path d=\"M20 10c0 4.993-5.539 10.193-7.399 11.799a1 1 0 0 1-1.202 0C9.539 20.193 4 14.993 4 10a8 8 0 0 1 16 0\" /><circle cx=\"12\" cy=\"10\" r=\"3\" />",
 "image": "<rect width=\"18\" height=\"18\" x=\"3\" y=\"3\" rx=\"2\" ry=\"2\" /><circle cx=\"9\" cy=\"9\" r=\"2\" /><path d=\"m21 15-3.086-3.086a2 2 0 0 0-2.828 0L6 21\" />",
 "chev": "<path d=\"m9 18 6-6-6-6\" />",
 "check": "<path d=\"M20 6 9 17l-5-5\" />",
 "layers": "<path d=\"M12.83 2.18a2 2 0 0 0-1.66 0L2.6 6.08a1 1 0 0 0 0 1.83l8.58 3.91a2 2 0 0 0 1.66 0l8.58-3.9a1 1 0 0 0 0-1.83z\" /><path d=\"M2 12a1 1 0 0 0 .58.91l8.6 3.91a2 2 0 0 0 1.65 0l8.58-3.9A1 1 0 0 0 22 12\" /><path d=\"M2 17a1 1 0 0 0 .58.91l8.6 3.91a2 2 0 0 0 1.65 0l8.58-3.9A1 1 0 0 0 22 17\" />",
 "book": "<path d=\"M12 5v16\" /><path d=\"M20.001 19A2 2 0 0022 17V5a2 2 0 00-1.999-2L16 3.002A5 5 0 0012 5a5 5 0 00-4-2H4a2 2 0 00-2 2v12a2 2 0 001.999 2H8a5 5 0 014 2 5 5 0 014-2z\" />",
 "route": "<path d=\"M9 5h8.5a1 1 0 010 7h-11a1 1 0 000 7H15\" /><circle cx=\"18\" cy=\"19\" r=\"3\" /><circle cx=\"6\" cy=\"5\" r=\"3\" />",
 "reset": "<path d=\"M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8\" /><path d=\"M3 3v5h5\" />",
 "upload": "<path d=\"M12 3v12\" /><path d=\"m17 8-5-5-5 5\" /><path d=\"M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4\" />",
 "download": "<path d=\"M12 15V3\" /><path d=\"M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4\" /><path d=\"m7 10 5 5 5-5\" />",
 "swap": "<path d=\"M8 3 4 7l4 4\" /><path d=\"M4 7h16\" /><path d=\"m16 21 4-4-4-4\" /><path d=\"M20 17H4\" />",
 "shop": "<path d=\"M15 21v-5a1 1 0 0 0-1-1h-4a1 1 0 0 0-1 1v5\" /><path d=\"M17.774 10.31a1.12 1.12 0 0 0-1.549 0 2.5 2.5 0 0 1-3.451 0 1.12 1.12 0 0 0-1.548 0 2.5 2.5 0 0 1-3.452 0 1.12 1.12 0 0 0-1.549 0 2.5 2.5 0 0 1-3.77-3.248l2.889-4.184A2 2 0 0 1 7 2h10a2 2 0 0 1 1.653.873l2.895 4.192a2.5 2.5 0 0 1-3.774 3.244\" /><path d=\"M4 10.95V19a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-8.05\" />"
};
  Object.keys(L).forEach(function (k) { IC[k] = L[k] });
})();

;
