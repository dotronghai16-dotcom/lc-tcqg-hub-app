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
      { k: 'ma_mau', nhan: 'Mã mẫu', bb: 1 }, { k: 'nhom_mau', nhan: 'Nhóm mẫu' }, { k: 'vaccine', nhan: 'Vaccine' }, { k: 'thoi_diem', nhan: 'Thời điểm' }, { k: 'doi_tuong', nhan: 'Đối tượng' }, { k: 'kenh', nhan: 'Kênh' }, { k: 'dieu_kien', nhan: 'Điều kiện sử dụng', alias: ['dieu kien'] },
      { k: 'noi_dung', nhan: 'Nội dung tin nhắn', alias: ['noi dung', 'noi dung tin'], bb: 1, bien: 1 }, { k: 'ten_mau', nhan: 'Tên mẫu' }] },
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
    if (!hd.n || thieu.length) return { loi: 'Không tìm thấy cột bắt buộc: ' + thieu.map(function (c) { return c.nhan }).join(', ') + ' (khớp ' + hd.n + ' cột ở dòng ' + (hd.r + 1) + ')' };
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
