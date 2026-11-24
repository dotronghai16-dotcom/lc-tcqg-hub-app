/* LC TCQG HUB: g·ªôp t·ª´ c√°c module pager.js, lib-parse.js, asm.js, asm-lib.js, scan-engine.js, asm-scan.js, asm-hr.js, asm-img.js, asm-admin.js, msg.js (build 20261124k) */

/* ===== pager.js ===== */
/* PH√ÇN TRANG CHUNG: m·ªçi b·∫£ng v√† danh s√°ch th·∫ª trong m√†n ch√≠nh ƒë·ªÅu t·ª± c√≥ "s·ªë d√≤ng m·ªói trang" (10 / 15 / 20) v√† n√∫t trang tr∆∞·ªõc/sau.
 * Ch·∫°y sau m·ªói l·∫ßn v·∫Ω m√†n h√¨nh (PGR.run). Kh√¥ng ƒë·ªïi d·ªØ li·ªáu: ch·ªâ ·∫©n b·ªõt d√≤ng ngo√†i trang hi·ªán t·∫°i. B·∫£ng c√≥ thu·ªôc t√≠nh data-nopg t·ª± ph√¢n trang ri√™ng (Kho Data).
 * S·ªë d√≤ng ch·ªçn ƒë∆∞·ª£c nh·ªõ tr√™n m√°y (localStorage). Trang hi·ªán t·∫°i nh·ªõ theo t·ª´ng b·∫£ng; khi s·ªë d√≤ng ƒë·ªïi (l·ªçc, t√¨m) th√¨ v·ªÅ trang 1. */
var PGR = (function () {
  var KEY = 'lc_pgsz', SIZES = [10, 15, 20], st = {};
  function size() { var v = 15; try { v = Number(localStorage.getItem(KEY)) || 15 } catch (e) { } return SIZES.indexOf(v) >= 0 ? v : 15 }
  function setSize(n) { try { localStorage.setItem(KEY, String(n)) } catch (e) { } }
  function opts() { var s = size(); return SIZES.map(function (n) { return '<option value="' + n + '"' + (n === s ? ' selected' : '') + '>' + n + ' d√≤ng</option>' }).join('') }
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
      bar.innerHTML = '<span>D√≤ng ' + a + '‚Äì' + b + ' / ' + total + '</span><span class="sp"></span><label>Hi·ªán <select data-pgsize>' + opts() + '</select></label><button class="btn b3" data-pgd="-1"' + (s.p <= 1 ? ' disabled' : '') + ' aria-label="Trang tr∆∞·ªõc">‚Äπ</button><b>' + s.p + '/' + pages + '</b><button class="btn b3" data-pgd="1"' + (s.p >= pages ? ' disabled' : '') + ' aria-label="Trang sau">‚Ä∫</button>';
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
/* ƒê·ªåC FILE EXCEL/CSV CHO TH∆Ø VI·ªÜN ‚Äì c·∫•u h√¨nh c·ªôt 7 khu + t√¨m d√≤ng ti√™u ƒë·ªÅ + chu·∫©n h√≥a + ki·ªÉm tra d√≤ng.
 * L√†m theo m√¥ h√¨nh c·ªßa KH-01 (ƒë·∫∑c t·∫£ 02_DAC-TA-DUNG-LAI): t√¨m d√≤ng ti√™u ƒë·ªÅ trong 40 d√≤ng ƒë·∫ßu, gh√©p c·ªôt theo t√™n/alias (b·ªè d·∫•u, th∆∞·ªùng h√≥a),
 * ƒë·ªçc "k·ª≥ √°p d·ª•ng" ·ªü ƒë·∫ßu file, g·ªôp √¥ merge, ng√†y Date -> yyyy-mm-dd (c·ªông 12 gi·ªù), s·ªë ti·ªÅn/%, chu·∫©n h√≥a bi·∫øn v·ªÅ M·ªòT d·∫°ng {{TEN_BIEN}}.
 * Ch·∫°y ƒë∆∞·ª£c c·∫£ trong tr√¨nh duy·ªát (window.LIBP) v√† Node (ƒë·ªÉ ki·ªÉm th·ª≠ b·∫±ng file th·∫≠t). Server v·∫´n ki·ªÉm l·∫°i to√†n b·ªô khi l∆∞u. */
(function (root) {
  function bo(s) { return String(s == null ? '' : s).normalize('NFD').replace(/[ÃÄ-ÕØ]/g, '').replace(/ƒë/g, 'd').replace(/ƒê/g, 'D').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim() }
  /* SDT_NV = SƒêT nh√¢n vi√™n (t·ª± ƒëi·ªÅn t·ª´ t√†i kho·∫£n); TEN_VACCINE = vaccine nh√¢n vi√™n ƒë√£ ch·ªçn (t·ª± ƒëi·ªÅn); 5 bi·∫øn c√≤n l·∫°i l√† ch·ªó nh√¢n vi√™n ƒëi·ªÅn tay tr∆∞·ªõc khi g·ª≠i. */
  var VARS = ['XUNG_HO', 'XUNG_HO_HOA', 'TEN_NGUOI_LH', 'TEN_BE', 'TUOI', 'TEN_NV', 'TEN_CO_SO', 'DIA_CHI', 'GIO_MO_CUA', 'HOTLINE', 'SDT_NV', 'TEN_VACCINE', 'NGAY_TIEM', 'SO_MUI', 'NGAY_HEN', 'LINK_GOOGLE_MAPS', 'NOI_DUNG_PHAN_HOI'];
  /* bi·∫øn c≈© c·ªßa KH-01 -> bi·∫øn chu·∫©n (so sau khi b·ªè d·∫•u) */
  var VAR_MAP = { 'ten kh': 'TEN_NGUOI_LH', 'ten me': 'TEN_NGUOI_LH', 'ten khach': 'TEN_NGUOI_LH', 'ten khach hang': 'TEN_NGUOI_LH', 'ten phu huynh': 'TEN_NGUOI_LH', 'ten nguoi lien he': 'TEN_NGUOI_LH', 'ten khach hang chung': 'TEN_NGUOI_LH',
    'ten be': 'TEN_BE', 'ten nguoi tiem': 'TEN_BE', 'ten nv': 'TEN_NV', 'ten nhan vien': 'TEN_NV', 'ten tu van vien': 'TEN_NV', 'dia chi shop': 'DIA_CHI', 'dia chi': 'DIA_CHI', 'dia chi diem nhan dien': 'DIA_CHI',
    'hotline': 'HOTLINE', 'hotline bac si': 'HOTLINE', 'ten co so': 'TEN_CO_SO', 'ten shop': 'TEN_CO_SO' };

  /* ---- c·∫•u h√¨nh 7 khu (k = t√™n c·ªôt l∆∞u, nhan = ti√™u ƒë·ªÅ file m·∫´u, alias = ti√™u ƒë·ªÅ kh√°c, bb = b·∫Øt bu·ªôc, kieu so|ngay, bien = √¥ n·ªôi dung tin, tuKy = ng√†y c√≥ th·ªÉ l·∫•y t·ª´ k·ª≥ √°p d·ª•ng) ---- */
  var CONFIG = {
    BANG_GIA: { ten: 'B·∫£ng gi√°', moTa: 'B·∫£ng gi√° vaccine m≈©i l·∫ª. Up th·∫≥ng file b·∫£ng gi√° g·ªëc, h·ªá th·ªëng t·ª± t√¨m d√≤ng ti√™u ƒë·ªÅ v√† k·ª≥ √°p d·ª•ng. Ngu·ªìn s·ªë cho tin nh·∫Øn.', cheDo: 'THAY_THE', cot: [
      { k: 'phong_benh', nhan: 'Ph√≤ng b·ªánh', alias: ['loai vaccine'] }, { k: 'ma_sp', nhan: 'M√£ s·∫£n ph·∫©m', alias: ['ma sp', 'ma hang'] },
      { k: 'ten_vaccine', nhan: 'T√™n vaccine', alias: ['ten san pham', 'vaccine', 'ten hang hoa'], bb: 1 }, { k: 'hang', nhan: 'H√£ng', alias: ['nha san xuat', 'hang san xuat', 'nuoc san xuat'] },
      { k: 'gia', nhan: 'Gi√° ni√™m y·∫øt', alias: ['gia', 'gia niem yet mui tiem', 'gia ban', 'don gia', 'gia mui le'], bb: 1, kieu: 'so' },
      { k: 'tu_ngay', nhan: 'T·ª´ ng√†y', alias: ['ngay bat dau', 'tu ngay ap dung'], kieu: 'ngay', tuKy: 1 }, { k: 'den_ngay', nhan: 'ƒê·∫øn ng√†y', alias: ['ngay ket thuc', 'den ngay ap dung'], kieu: 'ngay', tuKy: 1 }, { k: 'ghi_chu', nhan: 'Ghi ch√∫' }] },
    CTKM: { ten: 'CTKM th√°ng', moTa: 'Ch∆∞∆°ng tr√¨nh khuy·∫øn m·∫°i theo th√°ng. Up th·∫≥ng file CTKM g·ªëc: h·ªá th·ªëng ƒë·ªçc ph·∫ßn "KM theo s·∫£n ph·∫©m" (m≈©i l·∫ª, combo, g√≥i). C·∫ßn c√≥ k·ª≥ √°p d·ª•ng ·ªü ƒë·∫ßu file ho·∫∑c c·ªôt ng√†y.', cheDo: 'BO_SUNG', ctkm: 1, cot: [
      { k: 'phan_loai', nhan: 'Ph√¢n lo·∫°i', alias: [] }, { k: 'loai_vaccine', nhan: 'Lo·∫°i vaccine', alias: [] }, { k: 'noi_dung_ct', nhan: 'N·ªôi dung ch∆∞∆°ng tr√¨nh', alias: ['noi dung ct'] },
      { k: 'ma_sp', nhan: 'M√£ s·∫£n ph·∫©m', alias: ['ma sp'] }, { k: 'ten_sp', nhan: 'T√™n s·∫£n ph·∫©m', alias: ['ten vaccine'], bb: 1 }, { k: 'nha_sx', nhan: 'Nh√† s·∫£n xu·∫•t', alias: ['hang'] },
      { k: 'gia_niem_yet', nhan: 'Gi√° ni√™m y·∫øt/m≈©i ti√™m', alias: ['gia niem yet', 'gia niem yet mui tiem'], kieu: 'so' }, { k: 'sl_mui', nhan: 'SL m≈©i/g√≥i', alias: ['sl mui', 'so luong mui'], kieu: 'so' },
      { k: 'gia_goc', nhan: 'Gi√° g·ªëc', kieu: 'so' }, { k: 'khuyen_mai', nhan: 'Khuy·∫øn m√£i/gi·∫£m gi√°', alias: ['khuyen mai', 'giam gia'], kieu: 'so', pc: 1 }, { k: 'gia_sau_km', nhan: 'Gi√° sau khuy·∫øn m√£i', alias: ['gia sau km'], kieu: 'so' },
      { k: 'gia_1_mui', nhan: 'Gi√° 1 m≈©i (khi mua g√≥i)', alias: ['gia 1 mui'], kieu: 'so' }, { k: 'phi_luu_tru', nhan: 'Ph√≠ l∆∞u tr·ªØ vaccine', alias: ['phi luu tru'], kieu: 'so' }, { k: 'gia_thanh_toan', nhan: 'Gi√° thanh to√°n', kieu: 'so' },
      { k: 'pham_vi', nhan: 'Ph·∫°m vi √°p d·ª•ng', alias: ['pham vi'] }, { k: 'tu_ngay', nhan: 'Ng√†y b·∫Øt ƒë·∫ßu', alias: ['tu ngay'], kieu: 'ngay', tuKy: 1 }, { k: 'den_ngay', nhan: 'Ng√†y k·∫øt th√∫c', alias: ['den ngay'], kieu: 'ngay', tuKy: 1 }, { k: 'ghi_chu', nhan: 'Ghi ch√∫' }, { k: 'nhom', nhan: 'Nh√≥m CT', alias: ['nhom ct'] }, { k: 'ten_chuong_trinh', nhan: 'T√™n ch∆∞∆°ng tr√¨nh', alias: ['ten ct'] }, { k: 'km_kieu', nhan: 'Ki·ªÉu KM', alias: [] }] },
    MAU_TUOI: { ten: 'M·∫´u B·∫ßu v√† Nhi', moTa: 'M·∫´u tin cho nh√≥m B·∫ßu v√† Nhi theo giai ƒëo·∫°n v√† b∆∞·ªõc chƒÉm s√≥c.', cheDo: 'BO_SUNG', cot: [
      { k: 'nhom_tuoi', nhan: 'Nh√≥m tu·ªïi', bb: 1 }, { k: 'buoc', nhan: 'B∆∞·ªõc chƒÉm s√≥c', alias: ['buoc'] }, { k: 'kenh', nhan: 'K√™nh' }, { k: 'ten_mau', nhan: 'T√™n m·∫´u', bb: 1 }, { k: 'noi_dung', nhan: 'N·ªôi dung', bb: 1, bien: 1 }, { k: 'ghi_chu', nhan: 'Ghi ch√∫' }] },
    MAU_CHU_DE: { ten: 'M·∫´u theo vaccine v√† ƒë·ªô tu·ªïi', moTa: 'M·∫´u tin theo t·ª´ng vaccine/ch·ªß ƒë·ªÅ, chia theo nh√≥m tu·ªïi/ƒë·ªëi t∆∞·ª£ng.', cheDo: 'BO_SUNG', cot: [
      { k: 'chu_de', nhan: 'Ch·ªß ƒë·ªÅ', bb: 1 }, { k: 'nhom_ap_dung', nhan: 'Nh√≥m √°p d·ª•ng' }, { k: 'ten_mau', nhan: 'T√™n m·∫´u', bb: 1 }, { k: 'noi_dung', nhan: 'N·ªôi dung', bb: 1, bien: 1 }, { k: 'thu_tu', nhan: 'Th·ª© t·ª±', kieu: 'so' }, { k: 'ghi_chu', nhan: 'Ghi ch√∫' }] },
    MAU_SAU_TIEM: { ten: 'M·∫´u CSKH sau ti√™m', moTa: 'M·∫´u tin chƒÉm s√≥c sau ti√™m (h·ªèi thƒÉm, d·∫∑n d√≤ theo vaccine).', cheDo: 'BO_SUNG', cot: [
      { k: 'ma_mau', nhan: 'M√£ m·∫´u' }, { k: 'nhom_mau', nhan: 'Nh√≥m m·∫´u' }, { k: 'vaccine', nhan: 'Vaccine' }, { k: 'thoi_diem', nhan: 'Th·ªùi ƒëi·ªÉm' }, { k: 'doi_tuong', nhan: 'ƒê·ªëi t∆∞·ª£ng' }, { k: 'kenh', nhan: 'K√™nh' }, { k: 'dieu_kien', nhan: 'ƒêi·ªÅu ki·ªán s·ª≠ d·ª•ng', alias: ['dieu kien'] },
      { k: 'noi_dung', nhan: 'N·ªôi dung tin nh·∫Øn', alias: ['noi dung', 'noi dung tin'], bb: 1, bien: 1 }, { k: 'ten_mau', nhan: 'T√™n m·∫´u', bb: 1 }, { k: 'ghi_chu', nhan: 'Ghi ch√∫' }] },
    MAU_CHUNG: { ten: 'M·∫´u tin chung', moTa: 'Kho tin nh·∫Øn m·∫´u chung (k·∫øt b·∫°n, chƒÉm s√≥c sau ti√™m, vaccine‚Ä¶) d√πng t·ª± do.', cheDo: 'THAY_THE', cot: [
      { k: 'ma_tin', nhan: 'M√£ tin', alias: ['ma'], bb: 1 }, { k: 'ten_ngan', nhan: 'T√™n ng·∫Øn' }, { k: 'chu_de', nhan: 'Ch·ªß ƒë·ªÅ' }, { k: 'ap_dung', nhan: '√Åp d·ª•ng cho', alias: ['ap dung'] }, { k: 'noi_dung', nhan: 'N·ªôi dung', alias: ['tin nhan'], bb: 1, bien: 1 },
      { k: 'ma_anh', nhan: 'M√£ ·∫£nh' }, { k: 'anh_salechat', nhan: '·∫¢nh Salechat', alias: ['duong dan anh'] }, { k: 'thu_tu', nhan: 'Th·ª© t·ª±', kieu: 'so' }, { k: 'ghi_chu', nhan: 'Ghi ch√∫' }] },
    NGUY_CO: { ten: 'Nguy c∆° & ƒë∆∞·ªùng l√¢y', moTa: 'Nguy c∆°, ƒë∆∞·ªùng l√¢y v√† ƒëi·ªÉm c·∫ßn nh·ªõ theo t·ª´ng b·ªánh. Nh√¢n vi√™n b·∫≠t "K√®m nguy c∆°" ƒë·ªÉ ch√®n v√†o tin t∆∞ v·∫•n.', cheDo: 'THAY_THE', cot: [
      { k: 'benh', nhan: 'B·ªánh', bb: 1 }, { k: 'vaccine_ap_dung', nhan: 'Vaccine √°p d·ª•ng', alias: ['vaccine'] }, { k: 'nguy_co', nhan: 'Nguy c∆°', bb: 1, bien: 1 }, { k: 'lay_qua', nhan: 'L√¢y qua' }, { k: 'diem_can_nho', nhan: 'ƒêi·ªÉm c·∫ßn nh·ªõ' },
      { k: 'tac_nhan', nhan: 'T√°c nh√¢n' }, { k: 'trieu_chung', nhan: 'Tri·ªáu ch·ª©ng ƒëi·ªÉn h√¨nh', alias: ['trieu chung'] }, { k: 'mua_vung', nhan: 'M√πa cao ƒëi·ªÉm & v√πng mi·ªÅn', alias: ['mua cao diem'] }, { k: 'ghi_chu', nhan: 'Ghi ch√∫' }] }
  };
  /* PH√ÅC ƒê·ªí & T·ª™ ƒêI·ªÇN VACCINE (H·∫£i 24/11): m·ªói d√≤ng 1 vaccine (t√™n th∆∞∆°ng m·∫°i). Theo m·∫´u file KH-01 "UP_PHAC_DO" (ten_vaccine, phong_benh, tuoi_bat_dau_ngay, tuoi_ket_thuc_ngay, so_mui, khoang_cach_ngay, noi_dung_tu_van, luu_y, ghi_chu)
   * v√† th√™m c·ªôt cho t·ª´ ƒëi·ªÉn tra c·ª©u (n∆∞·ªõc SX, ƒë·ªô tu·ªïi ch·ªâ ƒë·ªãnh, ch·ªëng ch·ªâ ƒë·ªãnh, ph·∫£n ·ª©ng sau ti√™m, nh√≥m vaccine chu·∫©n).
   * File c≈© ghi tu·ªïi + n∆∞·ªõc SX chung trong "L∆∞u √Ω" ("6 tu·∫ßn - 24 th√°ng ¬∑ SX: Ph√°p"): app t·ª± t√°ch ra; tu·ªïi k·∫øt th√∫c ƒë·ªÉ tr·ªëng th√¨ t·ª± suy t·ª´ ch·ªØ ƒë·ªô tu·ªïi (c√≥ ghi ch√∫ ƒë·ªÉ b√°c sƒ© x√°c nh·∫≠n). */
  function ageEnd(t) {
    var s = String(t || '').toLowerCase().normalize('NFD').replace(/[ÃÄ-ÕØ]/g, '').replace(/ƒë/g, 'd'), re = /(?:[-‚Äì]|<|duoi)\s*(?:tron\s*)?(?:<\s*)?(\d+)\s*(tuan|thang|tuoi)/g, m, last = null;
    while ((m = re.exec(s))) last = m; if (!last) return null;
    var n = Number(last[1]), u = last[2]; return u === 'tuan' ? n * 7 : u === 'thang' ? Math.round(n * 30.4375) : Math.round(n * 365.25)
  }
  CONFIG.PHAC_DO = { ten: 'Ph√°c ƒë·ªì & t·ª´ ƒëi·ªÉn vaccine', moTa: 'M·ªói d√≤ng m·ªôt vaccine (t√™n th∆∞∆°ng m·∫°i): nh√≥m b·ªánh, tu·ªïi, ph√°c ƒë·ªì, n∆∞·ªõc s·∫£n xu·∫•t, ch·ªëng ch·ªâ ƒë·ªãnh, ph·∫£n ·ª©ng sau ti√™m. App d√πng ƒë·ªÉ g·ª£i √Ω "vaccine ƒë√£ ƒë·∫øn tu·ªïi", l√†m n√∫t Tra c·ª©u vaccine cho nh√¢n vi√™n v√† n·ªëi ƒë√∫ng t∆∞ v·∫•n, nguy c∆°, khuy·∫øn m√£i theo vaccine.', cheDo: 'THAY_THE',
    hau: function (o) {
      var lu = String(o.luu_y || ''), m = /\s*[¬∑‚Ä¢]\s*SX:\s*([^¬∑\n]+?)\s*$/i.exec(lu);
      if (m) { if (!o.nuoc_sx) o.nuoc_sx = m[1].trim(); lu = lu.slice(0, m.index).trim() }
      lu = lu.replace(/^\*\s*/, '').trim(); if (lu && !o.do_tuoi) o.do_tuoi = lu;
      if (o.tuoi_ket_thuc_ngay === undefined || o.tuoi_ket_thuc_ngay === null || o.tuoi_ket_thuc_ngay === '') { var e = ageEnd(o.do_tuoi || lu); if (e != null && e > (Number(o.tuoi_bat_dau_ngay) || 0)) { o.tuoi_ket_thuc_ngay = e; o.ghi_chu = (o.ghi_chu ? o.ghi_chu + ' ¬∑ ' : '') + 'Tu·ªïi k·∫øt th√∫c t·ª± suy t·ª´ ƒë·ªô tu·ªïi ch·ªâ ƒë·ªãnh, c·∫ßn b√°c sƒ© x√°c nh·∫≠n' } }
    },
    cot: [
      { k: 'ten_vaccine', nhan: 'T√™n vaccine', alias: ['vaccine', 'ten thuong mai', 'ten vac xin'], bb: 1 }, { k: 'phong_benh', nhan: 'Ph√≤ng b·ªánh', alias: ['nhom benh', 'benh phong ngua', 'benh'], bb: 1 },
      { k: 'nhom_vaccine', nhan: 'Nh√≥m vaccine chu·∫©n', alias: ['nhom vaccine', 'nhom chuan'] },
      { k: 'tuoi_bat_dau_ngay', nhan: 'Tu·ªïi b·∫Øt ƒë·∫ßu (ng√†y)', alias: ['tuoi bat dau', 'tuoi toi thieu ngay', 'tu ngay tuoi'], bb: 1, kieu: 'so' }, { k: 'tuoi_ket_thuc_ngay', nhan: 'Tu·ªïi k·∫øt th√∫c (ng√†y)', alias: ['tuoi ket thuc', 'tuoi toi da ngay'], kieu: 'so' },
      { k: 'so_mui', nhan: 'S·ªë m≈©i', alias: ['so mui theo phac do'] }, { k: 'khoang_cach_ngay', nhan: 'Kho·∫£ng c√°ch gi·ªØa c√°c m≈©i (ng√†y)', alias: ['khoang cach', 'khoang cach giua cac mui', 'lich tiem'] },
      { k: 'noi_dung_tu_van', nhan: 'N·ªôi dung t∆∞ v·∫•n / ph√°c ƒë·ªì', alias: ['noi dung tu van', 'phac do', 'phac do tham khao', 'tu van'], bien: 1 },
      { k: 'luu_y', nhan: 'L∆∞u √Ω', alias: ['luu y chuyen mon'] }, { k: 'nuoc_sx', nhan: 'N∆∞·ªõc s·∫£n xu·∫•t', alias: ['xuat xu', 'nuoc san xuat', 'nsx'] }, { k: 'do_tuoi', nhan: 'ƒê·ªô tu·ªïi ch·ªâ ƒë·ªãnh', alias: ['do tuoi', 'chi dinh', 'do tuoi chi dinh', 'doi tuong'] },
      { k: 'chong_chi_dinh', nhan: 'Ch·ªëng ch·ªâ ƒë·ªãnh / th·∫≠n tr·ªçng', alias: ['chong chi dinh', 'than trong'] }, { k: 'phan_ung', nhan: 'Ph·∫£n ·ª©ng sau ti√™m th∆∞·ªùng g·∫∑p', alias: ['phan ung sau tiem', 'tac dung phu', 'phan ung'] }, { k: 'ghi_chu', nhan: 'Ghi ch√∫' }] };
  var ORDER = ['BANG_GIA', 'CTKM', 'MAU_TUOI', 'MAU_CHU_DE', 'MAU_SAU_TIEM', 'MAU_CHUNG', 'NGUY_CO', 'PHAC_DO'];

  function p2(n) { return n < 10 ? '0' + n : '' + n }
  function isoDate(d) { var t = new Date(d.getTime() + 12 * 3600e3); return t.getFullYear() + '-' + p2(t.getMonth() + 1) + '-' + p2(t.getDate()) }
  function cs(v) { if (v == null) return ''; if (v instanceof Date) return isoDate(v); return String(v).trim() }
  function toNum(v) {
    if (v == null || v === '') return null; if (typeof v === 'number') return isFinite(v) ? v : null;
    var s = String(v).replace(/[‚Ç´ƒëƒê*\s]|VND|vnd/g, ''); if (!s) return null; var pct = /%$/.test(s); s = s.replace('%', '');
    if (/^\d{1,3}([.,]\d{3})+$/.test(s)) s = s.replace(/[.,]/g, ''); else if (/^\d+,\d{1,2}$/.test(s)) s = s.replace(',', '.'); else if (/,/.test(s) && /\./.test(s)) s = s.replace(/,/g, '');
    var n = Number(s); if (!isFinite(n)) return NaN; return pct ? n / 100 : n
  }
  function toDate(v) {
    if (v == null || v === '') return null; if (v instanceof Date) return isoDate(v); var s = String(v).trim(), m;
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
    if ((m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})$/))) return m[3] + '-' + p2(+m[2]) + '-' + p2(+m[1]);
    return s
  }
  /* chu·∫©n h√≥a bi·∫øn v·ªÅ {{TEN_BIEN}}; tr·∫£ {t, lat:[bi·∫øn l·∫°]} */
  function normVars(t) {
    t = String(t == null ? '' : t);
    function rep(m, inner) { var k = VAR_MAP[bo(inner)]; return k ? '{{' + k + '}}' : m }
    t = t.replace(/\{\{([^{}\n]{1,40})\}\}/g, function (m, i) { var u = i.trim().toUpperCase(); if (VARS.indexOf(u) >= 0) return '{{' + u + '}}'; var k = VAR_MAP[bo(i)] || (u === 'TEN_KHACH' ? 'TEN_NGUOI_LH' : u === 'TEN_NGUOI_TIEM' ? 'TEN_BE' : u === 'TEN_NHAN_VIEN' ? 'TEN_NV' : u === 'DIA_CHI_DIEM_NHAN_DIEN' ? 'DIA_CHI' : u === 'HOTLINE_BAC_SI' ? 'HOTLINE' : ''); return k ? '{{' + k + '}}' : m });
    t = t.replace(/\[([^\[\]\n]{1,40})\]/g, rep).replace(/<([^<>\n]{1,40})>/g, rep);
    return t
  }
  function lat(t) {
    var o = [], r, re = /\{\{([^{}\n]*)\}\}/g; while ((r = re.exec(t))) if (VARS.indexOf(r[1]) < 0) o.push('{{' + r[1] + '}}');
    re = /\[([A-Z√Ä-·ª∏][A-Z√Ä-·ª∏0-9 _]{2,30})\]/g; while ((r = re.exec(t))) o.push('[' + r[1] + ']'); return o
  }
  /* ki·ªÉm tra 1 d√≤ng (b·∫£n sao c·ªßa h√†m ki·ªÉm tr√™n m√°y ch·ªß) -> {thieu:[], canh:[]} */
  function check(loai, row) {
    var cfg = CONFIG[loai], th = [], ca = [];
    cfg.cot.forEach(function (c) {
      var v = row[c.k], s = v == null ? '' : String(v).trim();
      if (c.bb && s === '') th.push('Thi·∫øu ' + c.nhan);
      if (s !== '') {
        if (c.kieu === 'so' && !/^-?\d+(\.\d+)?$/.test(s)) th.push(c.nhan + ' kh√¥ng ph·∫£i s·ªë');
        if (c.kieu === 'ngay' && !/^\d{4}-\d{2}-\d{2}$/.test(s)) th.push(c.nhan + ' kh√¥ng ph·∫£i ng√†y');
        if (s.length > 5000) th.push(c.nhan + ' qu√° 5000 k√Ω t·ª±');
        var L = lat(s); if (L.length) ca.push('Bi·∫øn l·∫°: ' + L.slice(0, 3).join(', '));
        if (/\.{3,}|‚Ä¶/.test(s) && c.bien) ca.push('C√≤n ch·ªó tr·ªëng "‚Ä¶" c·∫ßn ƒëi·ªÅn tay');
      }
    });
    return { thieu: th, canh: ca.filter(function (x, i) { return ca.indexOf(x) === i }) }
  }

  /* ---- t√¨m d√≤ng ti√™u ƒë·ªÅ ---- */
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
      if (!out.ten && (m = s.match(/T√™n ch∆∞∆°ng tr√¨nh\s*:?\s*(.+)/i))) out.ten = m[1].trim();
      if (!out.tu && (m = s.match(/(\d{1,2})\s*\/\s*(\d{1,2})(?:\s*\/\s*(\d{4}))?\s*(?:ƒë·∫øn|->|‚Üí|-|‚Äì)\s*(\d{1,2})\s*\/\s*(\d{1,2})\s*\/\s*(\d{4})/i))) {
        var y2 = +m[6], y1 = m[3] ? +m[3] : y2; out.tu = y1 + '-' + p2(+m[2]) + '-' + p2(+m[1]); out.den = y2 + '-' + p2(+m[5]) + '-' + p2(+m[4]);
      }
    });
    return out
  }
  function emptyRow(r) { return !(r || []).some(function (v) { return v != null && String(v).trim() !== '' }) }

  /* ---- ƒë·ªçc chung ---- */
  function parse(loai, aoa) {
    var cfg = CONFIG[loai], cols = cfg.cot, hd = findHeader(aoa, cols);
    var thieu = cols.filter(function (c) { return c.bb && hd.map[c.k] === undefined && !c.tuKy });
    if (!hd.n || thieu.length) return { loi: 'Kh√¥ng t√¨m th·∫•y c·ªôt b·∫Øt bu·ªôc: ' + thieu.map(function (c) { return c.nhan }).join(', ') + ' (kh·ªõp ' + hd.n + ' c·ªôt ·ªü d√≤ng ' + (hd.r + 1) + '). C√°c c·ªôt ƒë·ªçc ƒë∆∞·ª£c ·ªü d√≤ng ƒë√≥: ' + (aoa[hd.r] || []).filter(function (v) { return v != null && String(v).trim() !== '' }).slice(0, 12).map(cs).join(' | ') };
    var ky = kyOf(aoa, hd.r), rows = [], boQua = [], cur = {}, prog = 0, curProg = null, sec = null, secCond = '';
    for (var r = hd.r + 1; r < aoa.length; r++) {
      var src = aoa[r] || []; if (emptyRow(src)) continue;
      if (cfg.ctkm) {
        var first = cs(src.filter(function (v) { return v != null && String(v).trim() !== '' })[0]);
        if (/^\d+\s*\/\s*\S/.test(first) && r > hd.r + 1) { sec = { nhom: 'N' + p2(++prog), title: first.replace(/^\d+\s*\/\s*/, '').trim() }; secCond = ''; continue }
        if (sec) {   /* c√°c m·ª•c sau "2/ ‚Ä¶", "3/ ‚Ä¶": qu√† t·∫∑ng, KOL, v√≤ng quay ‚Üí nh√≥m "∆Øu ƒë√£i chung" (ƒë·ªçc g·∫ßn ƒë√∫ng) */
          var cells = src.map(cs), nz = cells.filter(Boolean).length, hw = cells.filter(function (x) { return /th·ªùi gian √°p d·ª•ng|t√™n s·∫£n ph·∫©m|t√™n sp|t√™n qu√†|th·ªÉ l·ªá|gi√° tr·ªã bill|ghi ch√∫|ƒëi·ªÅu ki·ªán √°p d·ª•ng|ph·∫°m vi/i.test(x) }).length;
          if (hw >= 2) { if (/^qu√† t·∫∑ng/i.test(cells[2])) secCond = cells[2]; continue }
          if (cells[2]) secCond = cells[2];
          var name = ''; [4, 5, 3].some(function (i) { var x = cells[i]; if (x && !/^'?\d{5,}/.test(x) && !/^(t·ª´|ƒë·∫øn|√°p d·ª•ng)\b/i.test(x)) { name = x; return true } return false });
          if (!name && nz <= 3) name = cells[2];
          if (!name) { boQua.push({ dong: r + 1, ly_do: 'M·ª•c ' + sec.title + ': d√≤ng kh√¥ng c√≥ t√™n qu√†' }); continue }
          var g = { nhom: sec.nhom, phan_loai: '∆Øu ƒë√£i chung', loai_vaccine: sec.title, noi_dung_ct: secCond, ten_sp: name, ghi_chu: cells.filter(function (x) { return x && x !== name && x !== secCond }).join(' ¬∑ ').slice(0, 300), tu_ngay: ky.tu || undefined, den_ngay: ky.den || undefined, _dong: r + 1, _uudai: 1 };
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
        if (!curProg) { boQua.push({ dong: r + 1, ly_do: 'Ch∆∞a thu·ªôc ch∆∞∆°ng tr√¨nh n√†o' }); continue }
        if (!o.ten_sp) { boQua.push({ dong: r + 1, ly_do: 'Kh√¥ng c√≥ t√™n s·∫£n ph·∫©m' }); continue }
        o.nhom = curProg.nhom; o.phan_loai = curProg.phan_loai; if (!o.loai_vaccine) o.loai_vaccine = curProg.loai_vaccine; if (!o.noi_dung_ct) o.noi_dung_ct = curProg.noi_dung_ct;
        if (ky.ten) o.ten_chuong_trinh = ky.ten; if (!o.km_kieu && o.khuyen_mai > 0 && o.khuyen_mai < 1) o.km_kieu = 'PHAN_TRAM';
      } else {
        cols.forEach(function (c) { if (c.fill === 'tiep') { if (o[c.k] === undefined && cur[c.k] !== undefined) o[c.k] = cur[c.k]; else if (o[c.k] !== undefined) cur[c.k] = o[c.k] } });
        var hasBB = cols.some(function (c) { return c.bb && o[c.k] !== undefined }), nonEmpty = src.filter(function (v) { return v != null && String(v).trim() !== '' }).length;
        if (!hasBB && nonEmpty <= 2) { boQua.push({ dong: r + 1, ly_do: 'D√≤ng ghi ch√∫, kh√¥ng c√≥ d·ªØ li·ªáu ch√≠nh' }); continue }
      }
      cols.forEach(function (c) { if (c.tuKy && o[c.k] === undefined && ky[c.k === 'tu_ngay' ? 'tu' : 'den']) o[c.k] = ky[c.k === 'tu_ngay' ? 'tu' : 'den'] });
      if (cfg.hau) cfg.hau(o);
      o._dong = r + 1; rows.push(o);
    }
    var tong = rows.length, bad = 0;
    rows.forEach(function (o) { var ck = check(loai, o); o._thieu = ck.thieu; o._canh = ck.canh; if (o._uudai) o._canh.push('∆Øu ƒë√£i chung: ƒë·ªçc g·∫ßn ƒë√∫ng, h√£y xem l·∫°i'); if (ck.thieu.length) bad++ });
    return { rows: rows, ky: ky, boQua: boQua, tong: tong, thieu: bad, hopLe: tong - bad, dongTieuDe: hd.r + 1 }
  }

  /* ---- ch·ªçn sheet ph√π h·ª£p nh·∫•t trong workbook (XLSX = th∆∞ vi·ªán SheetJS) ---- */
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
    if (!best) return { loi: 'File kh√¥ng c√≥ sheet n√†o.' };
    var p = parse(loai, best.aoa); p.sheet = best.name; return p
  }
  /* file m·∫´u: sheet DU_LIEU (ch·ªâ d√≤ng ti√™u ƒë·ªÅ) + HUONG_DAN (c·ªôt, b·∫Øt bu·ªôc, ki·ªÉu, ghi ch√∫) */
  function templateAoa(loai) {
    var cfg = CONFIG[loai];
    return { du_lieu: [cfg.cot.map(function (c) { return c.nhan })], huong_dan: [['C·ªôt', 'B·∫Øt bu·ªôc', 'Ki·ªÉu', 'Ghi ch√∫']].concat(cfg.cot.map(function (c) { return [c.nhan, c.bb ? 'C√≥' : '', c.kieu === 'so' ? 'S·ªë' : c.kieu === 'ngay' ? 'Ng√†y' : 'Ch·ªØ', c.bien ? 'Bi·∫øn trong tin d√πng d·∫°ng {{TEN_BIEN}}: ' + VARS.join(', ') : (c.tuKy ? 'C√≥ th·ªÉ b·ªè tr·ªëng n·∫øu file c√≥ k·ª≥ √°p d·ª•ng ·ªü ƒë·∫ßu' : '')] })) }
  }
  root.LIBP = { bo: bo, CONFIG: CONFIG, ORDER: ORDER, VARS: VARS, parse: parse, check: check, normVars: normVars, fromWorkbook: fromWorkbook, templateAoa: templateAoa, findHeader: findHeader };
  if (typeof module !== 'undefined') module.exports = root.LIBP;
})(typeof window !== 'undefined' ? window : globalThis);

;

/* ===== asm.js ===== */
/* M√ÄN ASM (b·∫£n th·ª≠) ‚Äì d√πng chung khung v√† d·ªØ li·ªáu v·ªõi index.html (D, ME, S, rest, rpc, load, esc, ic, toast, ...).
 * ASM = gi√°m ƒë·ªëc kinh doanh khu v·ª±c: xem to√†n b·ªô, giao Data (ƒë√≠ch danh ho·∫∑c c·∫£ Shop), b·∫≠t quy·ªÅn cho Qu·∫£n l√Ω chia, thu h·ªìi, n·∫°p kh√°ch, xu·∫•t d·ªØ li·ªáu.
 * M·ªçi thao t√°c ghi qua h√†m RPC c·ªßa c∆° s·ªü d·ªØ li·ªáu (ch·ªâ vai tr√≤ ASM g·ªçi ƒë∆∞·ª£c). B·∫£n th·ª≠: d·ªØ li·ªáu gi·∫£, t√≠nh t·ªïng h·ª£p ngay tr√™n m√°y (ƒë·ªß cho v√†i trƒÉm d√≤ng). */
var ASM = (function () {
  IC.upload = '<path d="M12 16V4m0 0-4 4m4-4 4 4M5 20h14"/>'; IC.download = '<path d="M12 4v12m0 0-4-4m4 4 4-4M5 20h14"/>'; IC.swap = '<path d="M7 7h12l-3-3M17 17H5l3 3"/>'; IC.book = '<path d="M5 4h11a3 3 0 0 1 3 3v13H8a3 3 0 0 1-3-3zM5 17a3 3 0 0 1 3-3h11"/>'; IC.shop = '<path d="M4 9 5.5 4h13L20 9M4 9h16v11H4zM9 20v-6h6v6"/>';
  IC.more = '<circle cx="5" cy="12" r="1.5"/><circle cx="12" cy="12" r="1.5"/><circle cx="19" cy="12" r="1.5"/>'; IC.shield = '<path d="M12 3 4 6v6c0 5 3.5 8 8 9 4.5-1 8-4 8-9V6z"/><path d="m9 12 2 2 4-4"/>';
  var NAV = [['a_tq', 'T·ªïng quan', 'home'], ['a_kho', 'Kho Data', 'list'], ['a_pc', 'Ph√¢n c√¥ng', 'swap'], ['a_nap', 'N·∫°p Data', 'upload'], ['a_bc', 'K·∫øt qu·∫£ & xu·∫•t', 'chart'], ['a_tk', 'Nh√¢n s·ª± & t√†i kho·∫£n', 'team'], ['a_tl', 'Th∆∞ vi·ªán', 'book'], ['a_pd', 'Ph√°c ƒë·ªì', 'shield'], ['acc', 'C·ªßa t√¥i', 'user']];
  var MOBILE = ['a_tq', 'a_kho', 'a_pc', 'a_tk'];
  var X = { pages: {}, binders: [], u: {} };
  var ST = { shop: '', band: '', f: '', q: '', pg: 1, sel: {}, selAll: false, imp: null, bcShop: '', bcSt: '' };
  function PAGEN() { return PGR.size() }
  function today0() { var d = new Date(); d.setHours(0, 0, 0, 0); return d }
  function overdueDays(t) { if (!t.han_lien_he || !open_(t)) return 0; var h = new Date(t.han_lien_he + 'T00:00:00'), n = Math.floor((today0() - h) / 86400000); return n > 0 ? n : 0 }
  function ageMonths(c) { var b = c && c.ngay_sinh ? new Date(c.ngay_sinh + 'T00:00:00') : null; if (!b) return null; var n = new Date(); return (n.getFullYear() - b.getFullYear()) * 12 + n.getMonth() - b.getMonth() - (n.getDate() < b.getDate() ? 1 : 0) }
  function ageTxt(c) { var m = ageMonths(c); if (m === null) return ''; return m < 24 ? Math.max(0, m) + ' th√°ng' : Math.floor(m / 12) + ' tu·ªïi' }
  var BANDS = [['', 'M·ªçi ƒë·ªô tu·ªïi'], ['0', 'D∆∞·ªõi 2 th√°ng'], ['1', '2 ‚Äì 6 th√°ng'], ['2', '6 ‚Äì 12 th√°ng'], ['3', '1 ‚Äì 2 tu·ªïi'], ['4', '2 ‚Äì 5 tu·ªïi'], ['5', 'Tr√™n 5 tu·ªïi']];
  function bandOf(c) { var m = ageMonths(c); if (m === null) return ''; return m < 2 ? '0' : m < 6 ? '1' : m < 12 ? '2' : m < 24 ? '3' : m < 60 ? '4' : '5' }
  function shops() { return Object.keys(D.shops).map(function (k) { return D.shops[k] }).filter(function (s) { return s.trang_thai === 'HOAT_DONG' }).sort(function (a, b) { return a.ma_shop.localeCompare(b.ma_shop) }) }
  function staffOf(id) { return Object.keys(D.acc).map(function (k) { return D.acc[k] }).filter(function (a) { return a.id_shop === id && a.trang_thai === 'HOAT_DONG' && (a.vai_tro === 'NHAN_VIEN' || a.vai_tro === 'QUAN_LY') }) }
  function qlOf(id) { var q = staffOf(id).filter(function (a) { return a.vai_tro === 'QUAN_LY' }).map(function (a) { return a.ho_ten }); return q.length ? q.join(', ') : '‚Äì' }
  function turnsByCust() { var m = {}; D.turns.forEach(function (t) { (m[t.ma_khach] = m[t.ma_khach] || []).push(t) }); return m }
  function custState(ts) { if (!ts || !ts.length) return 'kho'; return ts.some(open_) ? 'dang' : 'xong' }
  function yes(t) { var c = cat(t.ma_ket_qua); return t.trang_thai === 'HOAN_THANH' && c && c.ma === 'DONG_Y' }
  function pct(a, b) { return b ? Math.round(a / b * 100) : 0 }
  function sel(id, opts, val) { return '<select id="' + id + '">' + opts.map(function (o) { return '<option value="' + esc(o[0]) + '"' + (String(o[0]) === String(val) ? ' selected' : '') + '>' + esc(o[1]) + '</option>' }).join('') + '</select>' }

  /* ---------- khung ---------- */
  function shell() {
    function b(x) { return '<button class="nav ' + (S.page === x[0] ? 'on' : '') + '" data-p="' + x[0] + '">' + ic(x[2]) + x[1] + '</button>' }
    $('#side').innerHTML = '<div class="brand">' + LOGO + '</div><div class="grp">ƒêi·ªÅu ph·ªëi</div>' + NAV.slice(0, 3).map(b).join('') + '<div class="grp">D·ªØ li·ªáu</div>' + NAV.slice(3, 5).map(b).join('') + '<div class="grp">Nh√¢n s·ª± &amp; n·ªôi dung</div>' + NAV.slice(5, 8).map(b).join('') + '<div class="grp">H·ªá th·ªëng</div>' + NAV.slice(8).map(b).join('') +
      '<div class="me"><div class="ava">' + esc(ini(ME.ho_ten)) + '</div><div><b>' + esc(ME.ho_ten) + '</b><small>ASM ¬∑ Gi√°m ƒë·ªëc kinh doanh KV</small></div></div>';
    var mob = NAV.filter(function (x) { return MOBILE.indexOf(x[0]) >= 0 }), inMore = MOBILE.indexOf(S.page) < 0;
    $('#tabs').style.setProperty('--n', mob.length + 1);
    $('#tabs').innerHTML = mob.map(function (x) { return '<button class="tab ' + (S.page === x[0] ? 'on' : '') + '" data-p="' + x[0] + '">' + ic(x[2]) + x[1].replace('Nh√¢n s·ª± & t√†i kho·∫£n', 'Nh√¢n s·ª±') + '</button>' }).join('') + '<button class="tab ' + (inMore ? 'on' : '') + '" id="tmore">' + ic('more') + 'Th√™m</button>';
    document.querySelectorAll('[data-p]').forEach(function (e) { e.onclick = function () { S.page = e.dataset.p; ST.pg = 1; ST.sel = {}; render() } });
    var tm = $('#tmore'); if (tm) tm.onclick = function () {
      var d = document.createElement('div'); d.className = 'modal'; d.innerHTML = '<div class="card" style="width:100%;max-width:520px;border-radius:22px 22px 0 0;padding:14px 14px calc(14px + env(safe-area-inset-bottom))"><h2 style="font-size:17px;margin:2px 4px 8px">T·∫•t c·∫£ m·ª•c</h2>' + NAV.map(function (x) { return '<button class="nav ' + (S.page === x[0] ? 'on' : '') + '" data-mp="' + x[0] + '">' + ic(x[2]) + x[1] + '</button>' }).join('') + '</div>';
      document.body.appendChild(d); d.onclick = function (e) { if (e.target === d) d.remove() };
      [].slice.call(d.querySelectorAll('[data-mp]')).forEach(function (b) { b.onclick = function () { S.page = b.dataset.mp; ST.pg = 1; ST.sel = {}; d.remove(); render() } })
    }
  }
  function head(t, p, extra) { return mtop() + '<div class="row2"><div class="greet"><h1>' + t + '</h1>' + (p ? '<p>' + p + '</p>' : '') + '</div>' + (extra || '') + '</div>' }

  /* ---------- T·ªïng quan ---------- */
  function statsFor(shopId) {
    var T = D.turns.filter(function (t) { return !shopId || t.id_shop === shopId }), o = { luot: T.length, mo: 0, xong: 0, dy: 0, qh: 0, chia: 0 };
    T.forEach(function (t) { if (open_(t)) o.mo++; if (t.trang_thai === 'HOAN_THANH') o.xong++; if (yes(t)) o.dy++; if (overdueDays(t) > 0) o.qh++; if (t.recipient_type === 'SHOP' && t.trang_thai === 'DA_GIAO') o.chia++ });
    return o
  }
  function pgTq() {
    var tc = turnsByCust(), kho = 0, tot = Object.keys(D.cus).length; Object.keys(D.cus).forEach(function (k) { if (custState(tc[k]) === 'kho') kho++ });
    var all = statsFor(ST.shop), first = ME.ho_ten.trim().split(/\s+/).pop(), SH = shops();
    var h = head('Ch√†o anh ' + esc(first) + ', khu v·ª±c c√≥ <span style="color:var(--teal)">' + SH.length + ' Shop</span>', 'T·ªïng h·ª£p to√†n khu v·ª±c ¬∑ ' + new Date().toLocaleDateString('vi-VN', { weekday: 'long', day: 'numeric', month: 'long' }),
      '<div style="display:flex;gap:8px;align-items:center"><span class="meta">Shop:</span>' + sel('fshop', [['', 'T·∫•t c·∫£ Shop']].concat(SH.map(function (s) { return [s.id_shop, s.ten_shop] })), ST.shop) + '</div>');
    h += '<div class="stats stats6">' + [['Kh√°ch trong kho', tot], ['Ch∆∞a giao', kho], ['ƒêang gi·ªØ', all.mo], ['Ho√†n th√†nh', all.xong], ['ƒê·ªìng √Ω ti√™m', all.dy], ['Qu√° h·∫°n g·ªçi', all.qh]].map(function (x, i) { return '<div class="card st"><div class="l">' + x[0] + '</div><div class="v"' + (i === 5 && x[1] ? ' style="color:var(--bad)"' : '') + '>' + x[1] + '</div></div>' }).join('') + '</div>';
    var al = '';
    if (kho) al += '<div class="alert"><i>‚ñ¶</i><div><b>' + kho + ' kh√°ch ch∆∞a ƒë∆∞·ª£c giao</b><small>ƒêang n·∫±m trong kho</small></div><button class="btn b2" data-go="a_kho" style="min-height:38px;margin-left:auto">Ph√¢n c√¥ng</button></div>';
    if (all.chia) al += '<div class="alert a-warn"><i>‚áÑ</i><div><b>' + all.chia + ' l∆∞·ª£t giao cho Shop ch·ªù chia</b><small>Qu·∫£n l√Ω chia khi b·∫°n b·∫≠t quy·ªÅn cho l·∫ßn giao ƒë√≥</small></div><button class="btn b2" data-go="a_pc" style="min-height:38px;margin-left:auto">Xem l·∫ßn giao</button></div>';
    if (all.qh) al += '<div class="alert a-bad"><i>‚ö†</i><div><b>' + all.qh + ' l∆∞·ª£t qu√° h·∫°n li√™n h·ªá</b><small>T√≠nh t·ª´ "Ng√†y c·∫ßn li√™n h·ªá" c·ªßa t·ª´ng l∆∞·ª£t</small></div></div>';
    h += '<div class="card"><h2>C·∫ßn ch√∫ √Ω</h2>' + (al || '<div class="alert"><i>‚úì</i><div><b>Ch∆∞a c√≥ g√¨ c·∫ßn ch√∫ √Ω</b></div></div>') + '</div>';
    h += '<div class="card" style="padding:6px 8px"><h2 style="padding:10px 10px 0">So s√°nh c√°c Shop <span class="pill">b·∫•m m·ªôt Shop ƒë·ªÉ xem nh√¢n vi√™n</span></h2><div class="tw"><table><thead><tr><th>Shop</th><th>Qu·∫£n l√Ω</th><th>NV</th><th>ƒê∆∞·ª£c giao</th><th>ƒêang m·ªü</th><th>Ho√†n th√†nh</th><th>ƒê·ªìng √Ω ti√™m</th><th style="min-width:120px">Ti·∫øn ƒë·ªô</th><th>Qu√° h·∫°n</th></tr></thead><tbody>' +
      SH.map(function (s) { var x = statsFor(s.id_shop); return '<tr class="click" data-shop="' + esc(s.id_shop) + '"' + (ST.shop === s.id_shop ? ' style="background:var(--tint)"' : '') + '><td><b>' + esc(s.ten_shop) + '</b><small>' + esc(s.ma_shop) + '</small></td><td>' + esc(qlOf(s.id_shop)) + '</td><td>' + staffOf(s.id_shop).length + '</td><td>' + x.luot + '</td><td>' + x.mo + '</td><td>' + x.xong + '</td><td>' + x.dy + '</td><td><div class="bar" style="height:6px;border-radius:99px;background:var(--bg);overflow:hidden"><i style="display:block;height:100%;width:' + pct(x.xong, x.luot) + '%;background:var(--teal)"></i></div><small>' + pct(x.xong, x.luot) + '%</small></td><td>' + (x.qh ? '<span class="tag t-late">' + x.qh + '</span>' : '0') + '</td></tr>' }).join('') + '</tbody></table></div></div>';
    if (ST.shop) {
      var staff = staffOf(ST.shop);
      h += '<div class="card" style="padding:6px 8px"><h2 style="padding:10px 10px 0">Nh√¢n vi√™n ¬∑ ' + esc(shopName(ST.shop)) + '</h2><div class="tw"><table><thead><tr><th>Nh√¢n vi√™n</th><th>Ch·ª©c danh</th><th>ƒê∆∞·ª£c giao</th><th>ƒêang m·ªü</th><th>Ho√†n th√†nh</th><th>ƒê·ªìng √Ω ti√™m</th><th>Qu√° h·∫°n</th></tr></thead><tbody>' + staff.map(function (a) {
        var T = D.turns.filter(function (t) { return t.recipient_type === 'EMPLOYEE' && t.ma_tk_nhan === a.ma_tk }), mo = T.filter(open_).length, xx = T.filter(function (t) { return t.trang_thai === 'HOAN_THANH' }).length, dy = T.filter(yes).length, qh = T.filter(function (t) { return overdueDays(t) > 0 }).length;
        return '<tr><td><b>' + esc(a.ho_ten) + '</b></td><td>' + (a.vai_tro === 'QUAN_LY' ? 'Qu·∫£n l√Ω' : 'Nh√¢n vi√™n') + '</td><td>' + T.length + '</td><td>' + mo + '</td><td>' + xx + '</td><td>' + dy + '</td><td>' + (qh ? '<span class="tag t-late">' + qh + '</span>' : '0') + '</td></tr>' }).join('') + '</tbody></table></div></div>';
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
    var h = head('Kho Data', 'To√†n b·ªô kh√°ch. Ch·ªçn kh√°ch r·ªìi b·∫•m Ph√¢n c√¥ng.', '');
    h += '<div class="top"><div class="search">' + ic('search') + '<input id="kq" placeholder="T√¨m t√™n, m√£ TCQG, s·ªë ƒëi·ªán tho·∫°i‚Ä¶" value="' + esc(ST.q) + '"></div>' + sel('fband', BANDS, ST.band) + '</div>';
    h += '<div class="chips">' + [['', 'T·∫•t c·∫£'], ['kho', 'Ch∆∞a giao'], ['dang', 'ƒêang giao'], ['xong', 'ƒê√£ xong']].map(function (x) { return '<button class="chip ' + (ST.f === x[0] ? 'on' : '') + '" data-f="' + x[0] + '">' + x[1] + '<em>' + cnt[x[0]] + '</em></button>' }).join('') + '</div>';
    h += '<div class="card" style="padding:6px 8px"><div class="tw"><table data-nopg><thead><tr><th style="width:36px"><input type="checkbox" id="selpg" ' + (rows.length && rows.every(function (r) { return ST.sel[r.c.ma_khach] }) ? 'checked' : '') + ' style="width:20px;height:20px"></th><th>Kh√°ch</th><th>Tu·ªïi</th><th>Ph·ª• huynh</th><th>Tr·∫°ng th√°i</th><th>ƒêang gi·ªØ</th></tr></thead><tbody>' +
      rows.map(function (r) { var op = r.ts.filter(open_)[0], who = op ? (op.recipient_type === 'SHOP' ? shopName(op.id_shop) + ' (Shop)' : nm(op.ma_tk_nhan)) : '';
        return '<tr><td><input type="checkbox" data-k="' + esc(r.c.ma_khach) + '" ' + (ST.sel[r.c.ma_khach] ? 'checked' : '') + ' style="width:20px;height:20px"></td><td><b>' + esc(r.c.ho_ten) + '</b><small>' + esc(r.c.ma_tcqg || r.c.ma_khach) + '</small></td><td>' + esc(ageTxt(r.c)) + '</td><td>' + esc(r.l.ho_ten || '') + '<small>' + esc(r.l.sdt || '') + '</small></td><td>' + ({ kho: '<span class="tag t-tint">Ch∆∞a giao</span>', dang: '<span class="tag t-late">ƒêang giao</span>', xong: '<span class="tag t-ok">ƒê√£ xong</span>' })[r.st] + '</td><td>' + esc(who) + '</td></tr>' }).join('') +
      (rows.length ? '' : '<tr><td colspan="6" class="empty">Kh√¥ng c√≥ kh√°ch ph√π h·ª£p.</td></tr>') + '</tbody></table></div>' +
      '<div class="row2" style="padding:10px"><span class="meta">' + all.length + ' kh√°ch ¬∑ trang ' + ST.pg + '/' + pages + '</span><div style="display:flex;gap:8px;align-items:center"><select data-pgsize style="border:1px solid var(--line);border-radius:10px;padding:6px 8px;background:var(--card);min-height:38px">' + [10, 15, 20].map(function (n) { return '<option value="' + n + '"' + (n === PAGEN() ? ' selected' : '') + '>' + n + ' d√≤ng</option>' }).join('') + '</select><button class="btn b3" data-pg="-1" style="min-height:38px">‚Äπ</button><button class="btn b3" data-pg="1" style="min-height:38px">‚Ä∫</button></div></div></div>';
    h += '<div class="bulk" style="' + (n ? '' : 'display:none') + '"><b>ƒê√£ ch·ªçn ' + n + ' kh√°ch</b><span class="sp"></span>' + (all.length > n ? '<button class="btn b3" id="seltop" style="min-height:42px">Ch·ªçn c·∫£ ' + all.length + ' kh√°ch ƒëang l·ªçc</button>' : '') + '<button class="btn b3" id="selclr" style="min-height:42px">B·ªè ch·ªçn</button><button class="btn b1" id="asg" style="min-height:42px">Ph√¢n c√¥ng</button></div>';
    return h
  }
  function assignModal() {
    var ks = Object.keys(ST.sel); if (!ks.length) return; var tc = turnsByCust(), openN = ks.filter(function (k) { return custState(tc[k]) === 'dang' }).length;
    var topics = D.dm ? Object.keys(D.dm).map(function (k) { return D.dm[k] }).filter(function (c) { return c.loai === 'CHU_DE' && c.dang_dung }) : [];
    var m = { kieu: 'NHAN_VIEN', who: {}, shopSel: {}, ql: false, giaolai: false };
    var d = document.createElement('div'); d.className = 'modal';
    function draw() {
      var SH = shops(), nWho = Object.keys(m.who).length, nShop = Object.keys(m.shopSel).length, n = m.kieu === 'SHOP' ? nShop : nWho, per = n ? Math.ceil((ks.length - (m.giaolai ? 0 : openN)) / n) : 0;
      d.innerHTML = '<div class="card" style="width:100%;max-width:640px;border-radius:22px 22px 0 0;padding:18px;max-height:92vh;overflow:auto"><h2 style="font-size:19px;margin:0 0 4px">Ph√¢n c√¥ng ' + ks.length + ' kh√°ch</h2><div class="meta" style="margin-bottom:12px">Chia ƒë·ªÅu cho nh·ªØng ng∆∞·ªùi ho·∫∑c Shop b·∫°n ch·ªçn. M·ªói kh√°ch m·ªôt l∆∞·ª£t giao ri√™ng, gi·ªØ nguy√™n m√£ kh√°ch.</div>' +
        '<div class="chips" style="margin-bottom:10px"><button class="chip ' + (m.kieu === 'NHAN_VIEN' ? 'on' : '') + '" data-kieu="NHAN_VIEN">ƒê√≠ch danh nh√¢n vi√™n</button><button class="chip ' + (m.kieu === 'SHOP' ? 'on' : '') + '" data-kieu="SHOP">C·∫£ Shop</button></div>' +
        (m.kieu === 'NHAN_VIEN' ? SH.map(function (s) { var st = staffOf(s.id_shop); return st.length ? '<div class="sec"><h3>' + esc(s.ten_shop) + ' <a href="#" data-allshop="' + esc(s.id_shop) + '" style="text-transform:none;letter-spacing:0;font-weight:600;color:var(--teal-d)">ch·ªçn c·∫£ Shop</a></h3><div style="display:grid;gap:6px">' + st.map(function (a) { return '<label class="row" style="cursor:pointer"><input type="checkbox" data-w="' + esc(a.ma_tk) + '" ' + (m.who[a.ma_tk] ? 'checked' : '') + ' style="width:20px;height:20px;accent-color:var(--teal)"><span>' + esc(a.ho_ten) + ' <span class="meta">¬∑ ' + (a.vai_tro === 'QUAN_LY' ? 'Qu·∫£n l√Ω' : 'Nh√¢n vi√™n') + '</span></span></label>' }).join('') + '</div></div>' : '' }).join('')
          : '<div class="sec"><h3>Ch·ªçn Shop nh·∫≠n</h3><div style="display:grid;gap:6px">' + SH.map(function (s) { return '<label class="row" style="cursor:pointer"><input type="checkbox" data-sh="' + esc(s.id_shop) + '" ' + (m.shopSel[s.id_shop] ? 'checked' : '') + ' style="width:20px;height:20px;accent-color:var(--teal)"><span>' + esc(s.ten_shop) + ' <span class="meta">¬∑ Qu·∫£n l√Ω ' + esc(qlOf(s.id_shop)) + '</span></span></label>' }).join('') + '</div></div>' +
            '<label class="sw"><input type="checkbox" id="mql" ' + (m.ql ? 'checked' : '') + '><div><b>Cho ph√©p Qu·∫£n l√Ω chia l·∫°i cho nh√¢n vi√™n</b><small>Ch·ªâ √°p cho l·∫ßn giao n√†y. B·∫°n t·∫Øt ƒë∆∞·ª£c b·∫•t c·ª© l√∫c n√†o ·ªü m·ª•c Ph√¢n c√¥ng.</small></div></label>') +
        '<div class="form2"><div><label class="meta">Ch·ªß ƒë·ªÅ</label>' + sel('mcd', [['', '‚Äî ch·ªçn ‚Äî']].concat(topics.map(function (c) { return [c.ma_dm, c.ten] })), '') + '</div><div><label class="meta">H·∫°n li√™n h·ªá</label><input type="date" id="mhan" value="' + new Date().toISOString().slice(0, 10) + '"></div></div>' +
        (openN ? '<label class="sw" style="margin-top:8px"><input type="checkbox" id="mgl" ' + (m.giaolai ? 'checked' : '') + '><div><b>V·∫´n giao ' + openN + ' kh√°ch ƒëang c√≥ l∆∞·ª£t m·ªü</b><small>M·∫∑c ƒë·ªãnh b·ªè qua ƒë·ªÉ tr√°nh nhi·ªÅu ng∆∞·ªùi c√πng g·ªçi m·ªôt kh√°ch. B·∫≠t l√† giao ngo·∫°i l·ªá.</small></div></label>' : '') +
        '<div class="alert" style="margin-top:10px"><i>Œ£</i><div><b>' + (n ? 'Giao ' + (ks.length - (m.giaolai ? 0 : openN)) + ' kh√°ch cho ' + n + (m.kieu === 'SHOP' ? ' Shop' : ' ng∆∞·ªùi') + ' ¬∑ kho·∫£ng ' + per + ' kh√°ch m·ªói n∆°i' : 'Ch∆∞a ch·ªçn ng∆∞·ªùi nh·∫≠n') + '</b></div></div>' +
        '<div class="acts" style="margin-top:6px"><button class="btn b3" id="mx" style="flex:1">H·ªßy</button><button class="btn b1" id="mok" style="flex:1.4" ' + (n ? '' : 'disabled') + '>X√°c nh·∫≠n giao</button></div></div>';
      d.querySelectorAll('[data-kieu]').forEach(function (b) { b.onclick = function () { m.kieu = b.dataset.kieu; draw() } });
      d.querySelectorAll('[data-w]').forEach(function (b) { b.onchange = function () { b.checked ? m.who[b.dataset.w] = 1 : delete m.who[b.dataset.w]; draw() } });
      d.querySelectorAll('[data-sh]').forEach(function (b) { b.onchange = function () { b.checked ? m.shopSel[b.dataset.sh] = 1 : delete m.shopSel[b.dataset.sh]; draw() } });
      d.querySelectorAll('[data-allshop]').forEach(function (a) { a.onclick = function (e) { e.preventDefault(); staffOf(a.dataset.allshop).forEach(function (x) { m.who[x.ma_tk] = 1 }); draw() } });
      var q = d.querySelector('#mql'); if (q) q.onchange = function () { m.ql = q.checked; draw() }; var g = d.querySelector('#mgl'); if (g) g.onchange = function () { m.giaolai = g.checked; draw() };
      d.querySelector('#mx').onclick = function () { d.remove() };
      d.querySelector('#mok').onclick = function () {
        var cd = d.querySelector('#mcd').value, han = d.querySelector('#mhan').value, btn = d.querySelector('#mok'); btn.disabled = true; btn.textContent = 'ƒêang giao‚Ä¶';
        rpc('u_asm_assign', { p_ma_khachs: ks, p_kieu: m.kieu, p_nguoi_nhan: Object.keys(m.kieu === 'SHOP' ? m.shopSel : m.who), p_chu_de: cd, p_muc_dich: cd && D.dm[cd] ? D.dm[cd].ten : 'Giao Data', p_han: han || null, p_cho_ql_chia: m.ql, p_giao_lai: m.giaolai, p_req: rid() })
          .then(function (r) { d.remove(); ST.sel = {}; toast('ƒê√£ giao ' + r.so_luot + ' kh√°ch (' + r.ma_lan_pc + ')' + (r.bo_qua ? ', b·ªè qua ' + r.bo_qua + ' kh√°ch ƒëang c√≥ l∆∞·ª£t m·ªü' : '')); return load(true) }).then(render, function (e) { btn.disabled = false; btn.textContent = 'X√°c nh·∫≠n giao'; fail(e) })
      }
    }
    draw(); document.body.appendChild(d); d.onclick = function (e) { if (e.target === d) d.remove() }
  }

  /* ---------- Ph√¢n c√¥ng (c√°c l·∫ßn giao) ---------- */
  function runsInfo() {
    var m = {}; D.turns.forEach(function (t) { var r = m[t.ma_lan_pc] = m[t.ma_lan_pc] || { ma: t.ma_lan_pc, n: 0, mo: 0, xong: 0, shopN: 0, empN: 0, chua: 0 }; r.n++; if (open_(t)) r.mo++; if (t.trang_thai === 'HOAN_THANH') r.xong++; if (t.recipient_type === 'SHOP') { r.shopN++; if (t.trang_thai === 'DA_GIAO') r.chua++ } else r.empN++ });
    return Object.keys(m).map(function (k) { var r = m[k], run = D.runs[k] || {}; r.run = run; return r }).sort(function (a, b) { return String(b.run.thoi_gian || '').localeCompare(String(a.run.thoi_gian || '')) })
  }
  function pgPc() {
    var R = runsInfo(), h = head('Ph√¢n c√¥ng', 'C√°c l·∫ßn giao Data. B·∫≠t quy·ªÅn cho Qu·∫£n l√Ω chia v·ªõi l·∫ßn giao cho c·∫£ Shop.', '');
    h += '<div class="card" style="padding:6px 8px"><div class="tw"><table><thead><tr><th>L·∫ßn giao</th><th>Ch·ªß ƒë·ªÅ</th><th>Ng√†y giao</th><th>Ki·ªÉu</th><th>L∆∞·ª£t</th><th>ƒêang m·ªü</th><th>Ho√†n th√†nh</th><th>Qu·∫£n l√Ω ƒë∆∞·ª£c chia</th><th></th></tr></thead><tbody>' + R.map(function (r) {
      var run = r.run, d = loc(run.thoi_gian), kieu = r.shopN && r.empN ? 'ƒê√≠ch danh + Shop' : r.shopN ? 'C·∫£ Shop' : 'ƒê√≠ch danh';
      return '<tr><td><b>' + esc(r.ma) + '</b></td><td>' + esc((D.dm[run.ma_chu_de] || {}).ten || run.muc_dich || '') + '</td><td>' + dstr(d) + '</td><td>' + kieu + '</td><td>' + r.n + '</td><td>' + r.mo + '</td><td>' + r.xong + '</td><td>' + (r.shopN ? '<label class="row" style="cursor:pointer"><input type="checkbox" data-ql="' + esc(r.ma) + '" ' + (run.cho_ql_chia ? 'checked' : '') + ' style="width:22px;height:22px;accent-color:var(--teal)"><span class="meta">' + (run.cho_ql_chia ? 'ƒêang b·∫≠t' : 'T·∫Øt') + (r.chua ? ' ¬∑ ' + r.chua + ' ch·ªù chia' : '') + '</span></label>' : '<span class="meta">‚Äì</span>') + '</td><td>' + (r.mo ? '<button class="btn b3" data-rev="' + esc(r.ma) + '" style="min-height:36px;font-size:13px">Thu h·ªìi l∆∞·ª£t m·ªü</button>' : '') + '</td></tr>' }).join('') + '</tbody></table></div></div>';
    return h
  }
  function revoke(ma) {
    var T = D.turns.filter(function (t) { return t.ma_lan_pc === ma && open_(t) && t.trang_thai === 'DA_GIAO' }); if (!T.length) { toast('Kh√¥ng c√≥ l∆∞·ª£t n√†o ch∆∞a x·ª≠ l√Ω ƒë·ªÉ thu h·ªìi.', 0, 1); return }
    var d = document.createElement('div'); d.className = 'modal'; d.innerHTML = '<div class="card" style="width:100%;max-width:480px;border-radius:22px 22px 0 0;padding:18px"><h2 style="font-size:18px;margin:0 0 6px">Thu h·ªìi ' + T.length + ' l∆∞·ª£t ch∆∞a x·ª≠ l√Ω</h2><div class="meta" style="margin-bottom:10px">L∆∞·ª£t ƒë√£ b·∫Øt ƒë·∫ßu x·ª≠ l√Ω gi·ªØ nguy√™n. L·ªãch s·ª≠ lu√¥n ƒë∆∞·ª£c gi·ªØ. Ng∆∞·ªùi ƒëang gi·ªØ l∆∞·ª£t s·∫Ω kh√¥ng ghi th√™m ƒë∆∞·ª£c.</div><label class="field">L√Ω do (kh√¥ng b·∫Øt bu·ªôc)<input id="rl" placeholder="V√≠ d·ª•: giao l·∫°i cho Shop kh√°c"></label><div class="acts"><button class="btn b3" id="rx" style="flex:1">H·ªßy</button><button class="btn b1" id="rok" style="flex:1.4">X√°c nh·∫≠n thu h·ªìi</button></div></div>';
    document.body.appendChild(d); d.onclick = function (e) { if (e.target === d) d.remove() }; d.querySelector('#rx').onclick = function () { d.remove() };
    d.querySelector('#rok').onclick = function () { var b = d.querySelector('#rok'); b.disabled = true; rpc('u_asm_revoke', { p_ma_luots: T.map(function (t) { return t.ma_luot }), p_ly_do: d.querySelector('#rl').value, p_req: rid() }).then(function (r) { d.remove(); toast('ƒê√£ thu h·ªìi ' + r.so_luot + ' l∆∞·ª£t'); return load(true) }).then(render, function (e) { b.disabled = false; fail(e) }) }
  }

  /* ---------- N·∫°p Data ---------- */
  var TPL = 'ma_tcqg,ho_ten,ngay_sinh,gioi_tinh,ten_lien_he,quan_he,sdt\nTC900001,Nguy·ªÖn VƒÉn An,15/03/2026,M,Nguy·ªÖn Th·ªã Lan,M·∫π,0900123456\nTC900002,Tr·∫ßn B·∫£o Ng·ªçc,2026-01-20,F,Tr·∫ßn VƒÉn H√πng,B·ªë,0900654321\n';
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
      if (!r.ho_ten) e = 'Thi·∫øu h·ªç t√™n'; else if (!r.iso) e = 'Ng√†y sinh kh√¥ng h·ª£p l·ªá'; else if (r.ma_tcqg && have[r.ma_tcqg]) e = 'ƒê√£ c√≥ trong kho (b·ªè qua)'; else if (r.ma_tcqg && seen[r.ma_tcqg]) e = 'Tr√πng M√£ TCQG trong file'; if (r.ma_tcqg) seen[r.ma_tcqg] = 1; r.err = e; return r })
  }
  function pgNap() {
    var h = head('N·∫°p Data', 'D√°n n·ªôi dung file CSV (ho·∫∑c ch·ªçn file). App ki·ªÉm tra, xem tr∆∞·ªõc, r·ªìi m·ªõi n·∫°p. Kh√°ch tr√πng M√£ TCQG s·∫Ω ƒë∆∞·ª£c b·ªè qua.', '');
    h += '<div class="card"><div class="acts" style="margin-bottom:10px"><button class="btn b3" id="tpl">' + ic('download') + 'T·∫£i file m·∫´u</button><label class="btn b2" style="cursor:pointer">' + ic('upload') + 'Ch·ªçn file CSV<input type="file" id="nfile" accept=".csv,.txt" hidden></label></div><textarea id="ntx" style="min-height:140px;font-family:ui-monospace,Menlo,monospace;font-size:13px" placeholder="ma_tcqg,ho_ten,ngay_sinh,gioi_tinh,ten_lien_he,quan_he,sdt">' + esc(ST.impText || '') + '</textarea><div class="acts" style="margin-top:10px"><button class="btn b1" id="nchk">Ki·ªÉm tra</button></div></div>';
    if (ST.imp) {
      var ok = ST.imp.filter(function (r) { return !r.err }), dup = ST.imp.filter(function (r) { return /b·ªè qua/.test(r.err) }).length, bad = ST.imp.length - ok.length - dup;
      h += '<div class="stats stats4"><div class="card st"><div class="l">D√≤ng trong file</div><div class="v">' + ST.imp.length + '</div></div><div class="card st"><div class="l">H·ª£p l·ªá</div><div class="v" style="color:var(--ok)">' + ok.length + '</div></div><div class="card st"><div class="l">ƒê√£ c√≥ (b·ªè qua)</div><div class="v">' + dup + '</div></div><div class="card st"><div class="l">L·ªói</div><div class="v" style="color:var(--bad)">' + bad + '</div></div></div>' +
        '<div class="card" style="padding:6px 8px"><div class="tw"><table><thead><tr><th>D√≤ng</th><th>M√£ TCQG</th><th>H·ªç t√™n</th><th>Ng√†y sinh</th><th>Li√™n h·ªá</th><th>K·∫øt qu·∫£ ki·ªÉm tra</th></tr></thead><tbody>' + ST.imp.slice(0, 30).map(function (r) { return '<tr><td>' + r._dong + '</td><td>' + esc(r.ma_tcqg) + '</td><td>' + esc(r.ho_ten) + '</td><td>' + esc(r.ngay_sinh) + '</td><td>' + esc(r.ten_lien_he) + '<small>' + esc(r.sdt) + '</small></td><td>' + (r.err ? '<span class="tag ' + (/b·ªè qua/.test(r.err) ? 't-new' : 't-late') + '">' + esc(r.err) + '</span>' : '<span class="tag t-ok">H·ª£p l·ªá</span>') + '</td></tr>' }).join('') + '</tbody></table></div></div>' +
        '<div class="acts"><button class="btn b1" id="nok" ' + (ok.length ? '' : 'disabled') + '>N·∫°p ' + ok.length + ' kh√°ch h·ª£p l·ªá</button></div>';
    }
    if (ST.impRes) h += '<div class="alert"><i>‚úì</i><div><b>ƒê√£ n·∫°p xong</b><small>Th√™m m·ªõi ' + ST.impRes.them_moi + ' ¬∑ b·ªè qua do tr√πng ' + ST.impRes.trung_bo_qua + ' ¬∑ l·ªói ' + ST.impRes.loi.length + '</small></div><button class="btn b2" data-go="a_kho" style="min-height:38px;margin-left:auto">M·ªü Kho Data</button></div>';
    return h
  }
  function download(name, text) { var a = document.createElement('a'); a.href = URL.createObjectURL(new Blob(['Ôªø' + text], { type: 'text/csv;charset=utf-8' })); a.download = name; document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove() }, 500) }
  function csvCell(v) { v = String(v == null ? '' : v); return /[",\n;]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v }

  /* ---------- K·∫øt qu·∫£ & xu·∫•t ---------- */
  function bcTurns() { return D.turns.filter(function (t) { return (!ST.bcShop || t.id_shop === ST.bcShop) && (!ST.bcSt || (ST.bcSt === 'MO' ? open_(t) : t.trang_thai === ST.bcSt)) }) }
  function pgBc() {
    var T = bcTurns(), SH = shops(), by = {};
    T.forEach(function (t) { var k = t.recipient_type === 'EMPLOYEE' ? t.ma_tk_nhan : 'SHOP:' + t.id_shop, r = by[k] = by[k] || { k: k, t: t, n: 0, mo: 0, xong: 0, dy: 0 }; r.n++; if (open_(t)) r.mo++; if (t.trang_thai === 'HOAN_THANH') r.xong++; if (yes(t)) r.dy++ });
    var h = head('K·∫øt qu·∫£ & xu·∫•t', 'Xem k·∫øt qu·∫£ theo nh√¢n vi√™n v√† xu·∫•t d·ªØ li·ªáu ra file CSV (m·ªü ƒë∆∞·ª£c b·∫±ng Excel).', '');
    h += '<div class="top" style="flex-wrap:wrap">' + sel('bshop', [['', 'T·∫•t c·∫£ Shop']].concat(SH.map(function (s) { return [s.id_shop, s.ten_shop] })), ST.bcShop) + sel('bst', [['', 'M·ªçi tr·∫°ng th√°i'], ['MO', 'ƒêang m·ªü'], ['HOAN_THANH', 'Ho√†n th√†nh'], ['THU_HOI', 'ƒê√£ thu h·ªìi']], ST.bcSt) + '<span class="sp"></span><button class="btn b1" id="xl">' + ic('download') + 'Xu·∫•t theo l∆∞·ª£t (' + T.length + ')</button><button class="btn b3" id="xn">Xu·∫•t theo nh√¢n vi√™n</button></div>';
    h += '<div class="card" style="padding:6px 8px"><div class="tw"><table><thead><tr><th>Ng∆∞·ªùi nh·∫≠n</th><th>Shop</th><th>ƒê∆∞·ª£c giao</th><th>ƒêang m·ªü</th><th>Ho√†n th√†nh</th><th>ƒê·ªìng √Ω ti√™m</th><th>T·ª∑ l·ªá ƒë·ªìng √Ω / ho√†n th√†nh</th></tr></thead><tbody>' + Object.keys(by).map(function (k) { return by[k] }).sort(function (a, b) { return b.n - a.n }).map(function (r) {
      return '<tr><td><b>' + esc(r.k.indexOf('SHOP:') === 0 ? 'C·∫£ Shop (ch∆∞a chia)' : nm(r.k)) + '</b></td><td>' + esc(shopName(r.t.id_shop)) + '</td><td>' + r.n + '</td><td>' + r.mo + '</td><td>' + r.xong + '</td><td>' + r.dy + '</td><td>' + (r.xong ? pct(r.dy, r.xong) + '%' : '‚Äì') + '</td></tr>' }).join('') + '</tbody></table></div></div>';
    return h
  }
  function exportTurns() {
    var T = bcTurns(), head = ['ma_luot', 'ma_khach', 'ma_tcqg', 'ho_ten', 'shop', 'nguoi_nhan', 'chu_de', 'trang_thai', 'ket_qua', 'ngay_giao', 'han_lien_he', 'hoan_thanh_luc'];
    var rows = T.map(function (t) { var c = D.cus[t.ma_khach] || {}; return [t.ma_luot, t.ma_khach, c.ma_tcqg, c.ho_ten, shopName(t.id_shop), t.recipient_type === 'SHOP' ? 'C·∫£ Shop' : nm(t.ma_tk_nhan), (cat(t.ma_chu_de) || {}).ten, t.trang_thai, (cat(t.ma_ket_qua) || {}).ten || '', t.thoi_gian_giao_ban_dau, t.han_lien_he, t.thoi_gian_hoan_thanh || ''].map(csvCell).join(',') });
    download('ket-qua-theo-luot-' + new Date().toISOString().slice(0, 10) + '.csv', head.join(',') + '\n' + rows.join('\n')); toast('ƒê√£ xu·∫•t ' + T.length + ' l∆∞·ª£t')
  }
  function exportStaff() {
    var by = {}; bcTurns().forEach(function (t) { var k = t.recipient_type === 'EMPLOYEE' ? t.ma_tk_nhan : 'SHOP:' + t.id_shop, r = by[k] = by[k] || { nm: k.indexOf('SHOP:') === 0 ? 'C·∫£ Shop (ch∆∞a chia)' : nm(k), shop: shopName(t.id_shop), n: 0, mo: 0, xong: 0, dy: 0 }; r.n++; if (open_(t)) r.mo++; if (t.trang_thai === 'HOAN_THANH') r.xong++; if (yes(t)) r.dy++ });
    download('ket-qua-theo-nhan-vien-' + new Date().toISOString().slice(0, 10) + '.csv', 'nguoi_nhan,shop,duoc_giao,dang_mo,hoan_thanh,dong_y_tiem\n' + Object.keys(by).map(function (k) { var r = by[k]; return [r.nm, r.shop, r.n, r.mo, r.xong, r.dy].map(csvCell).join(',') }).join('\n')); toast('ƒê√£ xu·∫•t theo nh√¢n vi√™n')
  }

  /* ---------- Th∆∞ vi·ªán, T√†i kho·∫£n ---------- */
  function pgAcc() { return head('T√†i kho·∫£n c·ªßa t√¥i', '', '') + '<div class="card"><div class="alert"><div class="ava">' + esc(ini(ME.ho_ten)) + '</div><div><b>' + esc(ME.ho_ten) + '</b><small>' + esc(ME.email) + ' ¬∑ ASM</small></div></div><button class="btn b3" id="cpw" style="width:100%;margin-bottom:8px">ƒê·ªïi m·∫≠t kh·∫©u</button><button class="btn b3" id="out" style="width:100%">ƒêƒÉng xu·∫•t</button></div>' }
  X.u = { head: head, sel: sel, shops: shops, staffOf: staffOf, qlOf: qlOf, pct: pct, ST: ST };

  /* ---------- v·∫Ω v√† g·∫Øn s·ª± ki·ªán ---------- */
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
    document.querySelectorAll('input[data-ql]').forEach(function (b) { b.onchange = function () { b.disabled = true; rpc('u_asm_set_ql_chia', { p_ma_lan_pc: b.dataset.ql, p_value: b.checked }).then(function () { toast(b.checked ? 'ƒê√£ cho ph√©p Qu·∫£n l√Ω chia l·∫ßn giao ' + b.dataset.ql : 'ƒê√£ t·∫Øt quy·ªÅn chia'); return load(true) }).then(render, function (e) { fail(e); load(true).then(render) }) } });
    document.querySelectorAll('[data-rev]').forEach(function (b) { b.onclick = function () { revoke(b.dataset.rev) } });
    on('#tpl', 'onclick', function () { download('mau-nap-khach.csv', TPL) });
    on('#nfile', 'onchange', function (x) { var f = x.files[0]; if (!f) return; var rd = new FileReader(); rd.onload = function () { ST.impText = rd.result; ST.imp = checkRows(parseCSV(rd.result)); ST.impRes = null; render() }; rd.readAsText(f, 'utf-8') });
    on('#nchk', 'onclick', function () { ST.impText = $('#ntx').value; ST.imp = checkRows(parseCSV(ST.impText)); ST.impRes = null; if (!ST.imp.length) toast('Ch∆∞a c√≥ d√≤ng d·ªØ li·ªáu. D√≤ng ƒë·∫ßu ph·∫£i l√† t√™n c·ªôt.', 0, 1); render() });
    on('#nok', 'onclick', function (x) { x.disabled = true; var rows = ST.imp.filter(function (r) { return !r.err }).map(function (r) { return { ma_tcqg: r.ma_tcqg, ho_ten: r.ho_ten, ngay_sinh: r.iso, gioi_tinh: r.gioi_tinh, ten_lh: r.ten_lien_he, quan_he: r.quan_he, sdt: r.sdt } });
      rpc('u_asm_import', { p_rows: rows, p_req: rid() }).then(function (r) { ST.impRes = r; ST.imp = null; ST.impText = ''; toast('ƒê√£ n·∫°p ' + r.them_moi + ' kh√°ch'); return load(true) }).then(render, function (e) { x.disabled = false; fail(e) }) });
    on('#bshop', 'onchange', function (x) { ST.bcShop = x.value; render() }); on('#bst', 'onchange', function (x) { ST.bcSt = x.value; render() });
    on('#xl', 'onclick', exportTurns); on('#xn', 'onclick', exportStaff); on('#out', 'onclick', logout); on('#cpw', 'onclick', pwModal);
    on('#lkbtn', 'onclick', function () { if (window.MSG) MSG.lookup() });
    X.binders.forEach(function (f) { f(on) });
  }
  return { render: render, X: X };
})();

;

/* ===== asm-lib.js ===== */
/* TH∆Ø VI·ªÜN ‚Äì KHO T·ª™ EXCEL (ƒê·ª£t 1, theo m√¥ h√¨nh KH-01): ch·ªçn khu -> t·∫£i file m·∫´u / up file -> XEM TR∆Ø·ªöC (ch∆∞a ghi) -> l∆∞u l√¥ CH·ªú DUY·ªÜT -> DUY·ªÜT -> ƒêANG D√ôNG.
 * L√¥: B·ªï sung (th√™m d√≤ng) ho·∫∑c Thay th·∫ø (l√¥ ƒëang d√πng c√πng khu chuy·ªÉn "ƒê√£ thay"). S·ª≠a t·∫°i √¥, x√≥a d√≤ng, ng·ª´ng/k√≠ch ho·∫°t l·∫°i/x√≥a l√¥, t·∫£i l√¥ v·ªÅ Excel. M·ªçi ghi qua h√†m RPC ch·ªâ ASM g·ªçi ƒë∆∞·ª£c.
 * ƒê·ªçc Excel b·∫±ng SheetJS ngay tr√™n tr√¨nh duy·ªát (lib-parse.js); m√°y ch·ªß ki·ªÉm l·∫°i to√†n b·ªô khi l∆∞u. D·ªØ li·ªáu th·ª≠ l√† d·ªØ li·ªáu gi·∫£/m·∫´u. */
var ASMLIB = (function () {
  var P = window.LIBP, LS = { sel: '', parsed: null, tenFile: '', cheDo: '', lots: null, loading: false, lot: null, lotRows: null, busy: false };
  var TRANG = { CHO_DUYET: ['t-late', 'Ch·ªù duy·ªát'], DANG_DUNG: ['t-ok', 'ƒêang d√πng'], DA_THAY: ['t-new', 'ƒê√£ thay'], NGUNG_DUNG: ['t-bad', 'Ng·ª´ng d√πng'], THIEU: ['t-bad', 'Thi·∫øu'] };
  var SAU = [['M·∫´u tin quy tr√¨nh', 'C√°c b∆∞·ªõc D, D+1, D+3, D+7 theo nh√≥m Lao/6in1. L√†m ·ªü ƒë·ª£t sau.'], ['Combo ∆∞u ƒë√£i', 'G√≥i combo d·∫°ng b·∫£ng. L√†m ·ªü ƒë·ª£t sau.']];

  function xlsx(cb) { if (window.XLSX) return cb(); var s = document.createElement('script'); s.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'; s.onload = cb; s.onerror = function () { toast('Kh√¥ng t·∫£i ƒë∆∞·ª£c b·ªô ƒë·ªçc Excel. Ki·ªÉm tra m·∫°ng r·ªìi th·ª≠ l·∫°i.', 0, 1) }; document.head.appendChild(s) }
  function tag(c, t) { return '<span class="tag ' + c + '">' + esc(t) + '</span>' }
  function cut(v, n) { var s = v == null ? '' : String(v); return s.length > n ? s.slice(0, n) + '‚Ä¶' : s }
  function fmt(v) { return typeof v === 'number' ? v.toLocaleString('vi-VN') : v }
  function dt(s) { var d = loc(s); return d ? dstr(d) : '' }
  function cfg() { return P.CONFIG[LS.sel] }
  function loadLots() { if (LS.loading) return; LS.loading = true; return rest('lib_lots?select=*&order=thoi_gian.desc&limit=300').then(function (a) { LS.lots = a; LS.loading = false; if (LS.lot) { var n = a.filter(function (x) { return x.ma_lo === LS.lot.ma_lo })[0]; LS.lot = n || null; if (!n) LS.lotRows = null } render() }, function (e) { LS.loading = false; LS.lots = LS.lots || []; fail(e) }) }
  function openLot(ma) { var l = (LS.lots || []).filter(function (x) { return x.ma_lo === ma })[0]; if (!l) return; LS.lot = l; LS.lotRows = null; render(); rest('lib_rows?select=*&ma_lo=eq.' + encodeURIComponent(ma) + '&order=stt.asc&limit=1000').then(function (a) { LS.lotRows = a; render() }, fail) }
  function reload(keepLot) { LS.loading = false; return loadLots().then(function () { if (keepLot && LS.lot) openLot(LS.lot.ma_lo) }) }

  /* ---------- giao di·ªán ---------- */
  function page() {
    if (LS.lots === null) return '<div class="card empty">ƒêang t·∫£i kho‚Ä¶</div>';
    var h = '<div class="card"><h2>1. Ch·ªçn khu v·ª±c <span class="pill">M·ªçi file ƒë·ªÅu qua Xem tr∆∞·ªõc ‚Üí L∆∞u ch·ªù duy·ªát ‚Üí Duy·ªát. Ch∆∞a duy·ªát th√¨ nh√¢n vi√™n ch∆∞a th·∫•y g√¨.</span></h2><div class="form2" style="grid-template-columns:repeat(auto-fill,minmax(210px,1fr))">';
    h += P.ORDER.map(function (k) {
      var c = P.CONFIG[k], dung = 0, nlo = 0; LS.lots.forEach(function (l) { if (l.loai === k) { nlo++; if (l.trang_thai === 'DANG_DUNG') dung += l.so_dong - l.so_dong_thieu } });
      return '<button class="sw" data-lbtile="' + k + '" style="display:block;text-align:left;' + (LS.sel === k ? 'border-color:var(--teal);box-shadow:0 0 0 3px rgba(29,135,135,.15)' : '') + '"><b>' + esc(c.ten) + '</b><small style="margin-top:4px">' + esc(c.moTa.slice(0, 90)) + (c.moTa.length > 90 ? '‚Ä¶' : '') + '</small><div style="margin-top:8px">' + tag(dung ? 't-ok' : 't-new', dung + ' d√≤ng ƒëang d√πng') + tag('t-new', nlo + ' l√¥') + '</div></button>'
    }).join('');
    h += SAU.map(function (s) { return '<div class="sw" style="display:block;opacity:.6;cursor:default"><b>' + esc(s[0]) + '</b><small style="margin-top:4px">' + esc(s[1]) + '</small><div style="margin-top:8px">' + tag('t-new', 'Ch∆∞a l√†m') + '</div></div>' }).join('');
    h += '</div></div>';
    if (LS.sel) h += uploadCard();
    if (LS.parsed) h += previewCard();
    h += lotsCard();
    if (LS.lot) h += lotCard();
    return h
  }
  function uploadCard() {
    var c = cfg(), mode = LS.cheDo || c.cheDo;
    return '<div class="card"><h2>2. T·∫£i t√†i li·ªáu ¬∑ ' + esc(c.ten) + '</h2><div class="meta" style="margin-bottom:10px">' + esc(c.moTa) + '</div><div class="row" style="flex-wrap:wrap"><button class="btn b3" id="lb-tpl">' + ic('download') + 'T·∫£i file m·∫´u (.xlsx)</button><label class="btn b2" style="cursor:pointer">' + ic('upload') + 'Ch·ªçn file Excel/CSV<input type="file" id="lb-file" accept=".xlsx,.xls,.csv" hidden></label><span class="meta">Ch·∫ø ƒë·ªô:</span><select id="lb-mode"><option value="BO_SUNG"' + (mode === 'BO_SUNG' ? ' selected' : '') + '>B·ªï sung</option><option value="THAY_THE"' + (mode === 'THAY_THE' ? ' selected' : '') + '>Thay th·∫ø</option></select></div>' +
      '<div class="meta" style="margin-top:8px">Anh c√≥ th·ªÉ up th·∫≥ng file g·ªëc (h·ªá th·ªëng t·ª± t√¨m d√≤ng ti√™u ƒë·ªÅ). <b>B·ªï sung</b>: th√™m d√≤ng m·ªõi, gi·ªØ nguy√™n d√≤ng ƒëang d√πng. <b>Thay th·∫ø</b>: khi Duy·ªát, l√¥ ƒëang d√πng c√πng khu chuy·ªÉn sang "ƒê√£ thay" (v·∫´n l∆∞u, kh√¥ng x√≥a). Bi·∫øn trong tin t·ª± ƒë·ªïi v·ªÅ d·∫°ng <code>{{TEN_BIEN}}</code>.</div></div>'
  }
  function previewCard() {
    var p = LS.parsed, c = cfg(); if (p.loi) return '<div class="card"><h2>3. Xem tr∆∞·ªõc</h2><div class="err">' + esc(p.loi) + '</div><div class="acts"><button class="btn b3" id="lb-drop">B·ªè file n√†y</button></div></div>';
    var warn = p.rows.filter(function (r) { return r._canh.length }).length;
    var h = '<div class="card"><h2>3. Xem tr∆∞·ªõc v√† x√°c nh·∫≠n <span class="pill">' + esc(LS.tenFile) + (p.sheet ? ' ¬∑ sheet ' + esc(p.sheet) : '') + '</span></h2><div class="stats stats4"><div class="card st"><div class="l">D√≤ng ƒë·ªçc ƒë∆∞·ª£c</div><div class="v">' + p.tong + '</div></div><div class="card st"><div class="l">H·ª£p l·ªá</div><div class="v" style="color:var(--ok)">' + p.hopLe + '</div></div><div class="card st"><div class="l">Thi·∫øu / sai</div><div class="v" style="color:' + (p.thieu ? 'var(--bad)' : 'inherit') + '">' + p.thieu + '</div></div><div class="card st"><div class="l">C√≥ c·∫£nh b√°o</div><div class="v" style="color:' + (warn ? 'var(--warn)' : 'inherit') + '">' + warn + '</div></div></div>';
    if (p.ky && p.ky.tu) h += '<div class="alert"><i>üìÖ</i><div><b>K·ª≥ √°p d·ª•ng ƒë·ªçc t·ª´ ƒë·∫ßu file: ' + esc(p.ky.tu) + ' ‚Üí ' + esc(p.ky.den) + '</b>' + (p.ky.ten ? '<small>' + esc(p.ky.ten) + '</small>' : '') + '</div></div>';
    h += '<div class="meta" style="margin:6px 0">D√≤ng ti√™u ƒë·ªÅ t√¨m th·∫•y ·ªü d√≤ng ' + p.dongTieuDe + '. D√≤ng thi·∫øu th√¥ng tin v·∫´n ƒë∆∞·ª£c l∆∞u nh∆∞ng KH√îNG ƒë∆∞a v√†o d√πng khi Duy·ªát, s·ª≠a t·∫°i √¥ sau khi l∆∞u.</div>';
    if (p.boQua.length) h += '<details class="alert a-warn" style="display:block"><summary><b>H·ªá th·ªëng kh√¥ng ƒë∆∞a v√†o ' + p.boQua.length + ' d√≤ng</b></summary><div class="meta" style="margin-top:6px">' + p.boQua.slice(0, 20).map(function (b) { return 'D√≤ng ' + b.dong + ': ' + esc(b.ly_do) }).join('<br>') + '</div></details>';
    h += '<div class="tw"><table><thead><tr><th>D√≤ng</th><th>K·∫øt qu·∫£</th>' + c.cot.map(function (x) { return '<th>' + esc(x.nhan) + '</th>' }).join('') + '</tr></thead><tbody>' + p.rows.slice(0, 50).map(function (r) {
      return '<tr><td>' + r._dong + '</td><td>' + (r._thieu.length ? tag('t-bad', r._thieu.join('; ')) : r._canh.length ? tag('t-late', r._canh.join('; ')) : tag('t-ok', 'H·ª£p l·ªá')) + '</td>' + c.cot.map(function (x) { return '<td title="' + esc(r[x.k]) + '">' + esc(cut(fmt(r[x.k]), 60)) + '</td>' }).join('') + '</tr>'
    }).join('') + '</tbody></table></div>' + (p.rows.length > 50 ? '<div class="meta" style="padding:6px">Hi·ªán 50 d√≤ng ƒë·∫ßu trong ' + p.rows.length + ' d√≤ng.</div>' : '');
    h += '<div class="acts" style="margin-top:10px"><button class="btn b1" id="lb-save"' + (p.tong ? '' : ' disabled') + '>L∆∞u ' + p.tong + ' d√≤ng v√†o l√¥ ch·ªù duy·ªát</button><button class="btn b3" id="lb-drop">B·ªè file n√†y</button></div></div>';
    return h
  }
  function lotsCard() {
    return '<div class="card" style="padding:6px 8px"><div class="row2" style="padding:10px 10px 0"><h2 style="margin:0">C√°c l√¥ ƒë√£ t·∫£i</h2><button class="btn b3" id="lb-refresh" style="min-height:38px">L√†m m·ªõi</button></div><div class="tw"><table><thead><tr><th>M√£ l√¥</th><th>Khu</th><th>File</th><th>Ch·∫ø ƒë·ªô</th><th>D√≤ng</th><th>Thi·∫øu</th><th>Tr·∫°ng th√°i</th><th>T·∫£i l√∫c</th><th></th></tr></thead><tbody>' + LS.lots.map(function (l) {
      var st = TRANG[l.trang_thai] || ['t-new', l.trang_thai]; return '<tr' + (LS.lot && LS.lot.ma_lo === l.ma_lo ? ' style="background:var(--tint)"' : '') + '><td><b>' + esc(l.ma_lo) + '</b></td><td>' + esc((P.CONFIG[l.loai] || {}).ten || l.loai) + '</td><td>' + esc(cut(l.ten_file, 40)) + '</td><td>' + (l.che_do === 'THAY_THE' ? 'Thay th·∫ø' : 'B·ªï sung') + '</td><td>' + l.so_dong + '</td><td>' + (l.so_dong_thieu ? tag('t-bad', l.so_dong_thieu) : '0') + '</td><td>' + tag(st[0], st[1]) + '</td><td>' + esc(dt(l.thoi_gian)) + '</td><td><button class="btn b3" data-lblot="' + esc(l.ma_lo) + '" style="min-height:36px;font-size:13px">Xem</button></td></tr>'
    }).join('') + (LS.lots.length ? '' : '<tr><td colspan="9" class="empty">Ch∆∞a c√≥ l√¥ n√†o. Ch·ªçn m·ªôt khu ·ªü tr√™n, t·∫£i file m·∫´u r·ªìi up file.</td></tr>') + '</tbody></table></div></div>'
  }
  function lotCard() {
    var l = LS.lot, c = P.CONFIG[l.loai], st = TRANG[l.trang_thai] || ['t-new', l.trang_thai], canEdit = l.trang_thai === 'CHO_DUYET' || l.trang_thai === 'DANG_DUNG';
    var btn = function (k, t, cls) { return '<button class="btn ' + (cls || 'b3') + '" data-lbact="' + k + '" style="min-height:40px">' + t + '</button>' };
    var h = '<div class="card"><div class="row2"><div><h2 style="margin:0">L√¥ ' + esc(l.ma_lo) + ' ¬∑ ' + esc(c.ten) + ' ' + tag(st[0], st[1]) + '</h2><div class="meta">' + esc(l.ten_file || '') + ' ¬∑ ' + (l.che_do === 'THAY_THE' ? 'Thay th·∫ø' : 'B·ªï sung') + ' ¬∑ ' + l.so_dong + ' d√≤ng, thi·∫øu ' + l.so_dong_thieu + (l.thoi_gian_duyet ? ' ¬∑ duy·ªát ' + esc(dt(l.thoi_gian_duyet)) : '') + (l.ky && l.ky.tu ? ' ¬∑ k·ª≥ ' + esc(l.ky.tu) + ' ‚Üí ' + esc(l.ky.den) : '') + '</div></div><button class="btn b3" id="lb-close" style="min-height:38px">ƒê√≥ng</button></div>';
    h += '<div class="acts" style="margin:10px 0;flex-wrap:wrap">' + (l.trang_thai === 'CHO_DUYET' ? btn('approve', 'Duy·ªát', 'b1') : '') + (canEdit ? btn('stop', 'Ng·ª´ng d√πng') : '') + (l.trang_thai === 'NGUNG_DUNG' ? btn('react', 'K√≠ch ho·∫°t l·∫°i') + btn('del', 'X√≥a l√¥') : '') + btn('xl', ic('download') + 'T·∫£i v·ªÅ Excel (.xlsx)') + '</div>';
    if (l.trang_thai === 'CHO_DUYET') h += '<div class="ro">L√¥ ch·ªù duy·ªát: nh√¢n vi√™n ch∆∞a th·∫•y. ' + (l.che_do === 'THAY_THE' ? 'Khi Duy·ªát, c√°c l√¥ ƒëang d√πng c√πng khu s·∫Ω chuy·ªÉn "ƒê√£ thay".' : 'Khi Duy·ªát, c√°c d√≤ng h·ª£p l·ªá ƒë∆∞·ª£c th√™m v√†o d√πng, d√≤ng ƒëang d√πng gi·ªØ nguy√™n.') + '</div>';
    if (canEdit) h += '<div class="meta" style="margin-bottom:6px">B·∫•m v√†o m·ªôt √¥ ƒë·ªÉ s·ª≠a (Enter = l∆∞u, Esc = h·ªßy). ‚úï = x√≥a d√≤ng.</div>';
    if (LS.lotRows === null) return h + '<div class="empty">ƒêang t·∫£i d√≤ng‚Ä¶</div></div>';
    h += '<div class="tw" style="max-height:480px;overflow:auto"><table><thead><tr><th>#</th><th>Tr·∫°ng th√°i</th>' + c.cot.map(function (x) { return '<th>' + esc(x.nhan) + '</th>' }).join('') + '<th>C·∫£nh b√°o</th>' + (canEdit ? '<th></th>' : '') + '</tr></thead><tbody>' + LS.lotRows.slice(0, 300).map(function (r) {
      var s = TRANG[r.trang_thai] || ['t-new', r.trang_thai];
      return '<tr' + (r.trang_thai === 'THIEU' ? ' style="background:var(--bad-bg)"' : '') + '><td>' + r.stt + '</td><td>' + tag(s[0], s[1]) + '</td>' + c.cot.map(function (x) { var v = r.data[x.k]; return '<td' + (canEdit ? ' class="sua" data-id="' + r.id + '" data-col="' + x.k + '" style="cursor:text"' : '') + ' title="' + esc(v) + '">' + esc(cut(fmt(v), 70)) + '</td>' }).join('') + '<td class="meta">' + esc(r.canh_bao || '') + '</td>' + (canEdit ? '<td><button class="btn b3" data-lbdel="' + r.id + '" style="min-height:32px;padding:0 10px" aria-label="X√≥a d√≤ng">‚úï</button></td>' : '') + '</tr>'
    }).join('') + '</tbody></table></div>' + (LS.lotRows.length > 300 ? '<div class="meta" style="padding:6px">Hi·ªán 300 d√≤ng ƒë·∫ßu trong ' + LS.lotRows.length + ' d√≤ng (t·∫£i Excel ƒë·ªÉ xem ƒë·ªß).</div>' : '') + '</div>';
    return h
  }

  /* ---------- h√†nh ƒë·ªông ---------- */
  function handle(name, buf) {
    xlsx(function () {
      try { var wb = XLSX.read(buf, { type: 'array', cellDates: true, codepage: 65001 }); LS.parsed = P.fromWorkbook(XLSX, wb, LS.sel); LS.tenFile = name; render() } catch (e) { toast('Kh√¥ng ƒë·ªçc ƒë∆∞·ª£c file: ' + e.message, 0, 1) }
    })
  }
  function template() { xlsx(function () { var t = P.templateAoa(LS.sel), wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(t.du_lieu), 'DU_LIEU'); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(t.huong_dan), 'HUONG_DAN'); XLSX.writeFile(wb, 'FILE_MAU_' + LS.sel + '.xlsx') }) }
  function save() {
    var p = LS.parsed; if (!p || !p.rows || LS.busy) return; LS.busy = true; var b = document.getElementById('lb-save'); if (b) b.disabled = true;
    var rows = p.rows.map(function (r) { var o = {}; Object.keys(r).forEach(function (k) { if (k.charAt(0) !== '_') o[k] = r[k] }); return o });
    rpc('u_lib_save_lot', { p_loai: LS.sel, p_ten_file: LS.tenFile, p_che_do: LS.cheDo || cfg().cheDo, p_ky: p.ky, p_rows: rows, p_req: rid() }).then(function (r) { LS.busy = false; LS.parsed = null; toast('ƒê√£ l∆∞u l√¥ ' + r.ma_lo + ': ' + r.tong + ' d√≤ng, thi·∫øu ' + r.thieu + '. ƒêang ch·ªù duy·ªát.'); return loadLots().then(function () { openLot(r.ma_lo) }) }, function (e) { LS.busy = false; if (b) b.disabled = false; fail(e) })
  }
  function act(k) {
    var l = LS.lot; if (!l) return; var done = function (msg, closeLot) { return function () { toast(msg); if (closeLot) { LS.lot = null; LS.lotRows = null } return reload(!closeLot) } };
    if (k === 'xl') return exportLot();
    if (k === 'approve') { if (!confirm('Duy·ªát l√¥ ' + l.ma_lo + '?' + (l.che_do === 'THAY_THE' ? '\nC√°c l√¥ ƒëang d√πng c√πng khu s·∫Ω chuy·ªÉn sang "ƒê√£ thay".' : '') + '\nC√°c d√≤ng h·ª£p l·ªá s·∫Ω hi·ªán cho nh√¢n vi√™n.')) return; rpc('u_lib_approve', { p_ma_lo: l.ma_lo, p_req: rid() }).then(function (r) { return done('ƒê√£ duy·ªát: ' + r.dung + ' d√≤ng ƒë∆∞a v√†o s·ª≠ d·ª•ng' + (r.thieu ? ', c√≤n ' + r.thieu + ' d√≤ng thi·∫øu ch·ªù s·ª≠a' : ''))() }, fail) }
    if (k === 'stop') { if (!confirm('Ng·ª´ng d√πng l√¥ ' + l.ma_lo + '? Nh√¢n vi√™n s·∫Ω kh√¥ng th·∫•y c√°c d√≤ng n√†y n·ªØa (d·ªØ li·ªáu v·∫´n gi·ªØ).')) return; rpc('u_lib_stop', { p_ma_lo: l.ma_lo, p_req: rid() }).then(done('ƒê√£ ng·ª´ng d√πng l√¥'), fail) }
    if (k === 'react') rpc('u_lib_reactivate', { p_ma_lo: l.ma_lo, p_req: rid() }).then(done('ƒê√£ ƒë∆∞a l√¥ v·ªÅ ch·ªù duy·ªát'), fail);
    if (k === 'del') { if (!confirm('X√≥a h·∫≥n l√¥ ' + l.ma_lo + ' v√† ' + l.so_dong + ' d√≤ng? Kh√¥ng ho√†n t√°c ƒë∆∞·ª£c (nh·∫≠t k√Ω v·∫´n gi·ªØ).')) return; rpc('u_lib_delete', { p_ma_lo: l.ma_lo, p_req: rid() }).then(done('ƒê√£ x√≥a l√¥', true), fail) }
  }
  function exportLot() {
    var l = LS.lot, c = P.CONFIG[l.loai]; xlsx(function () { var aoa = [c.cot.map(function (x) { return x.nhan })].concat((LS.lotRows || []).map(function (r) { return c.cot.map(function (x) { var v = r.data[x.k]; return v == null ? '' : v }) })); var wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), 'DU_LIEU'); XLSX.writeFile(wb, l.ma_lo + '_' + l.loai + '.xlsx') })
  }
  function editCell(td) {
    if (td.querySelector('input,textarea')) return; var id = Number(td.dataset.id), col = td.dataset.col, r = LS.lotRows.filter(function (x) { return x.id === id })[0]; if (!r) return;
    var cur = r.data[col] == null ? '' : String(r.data[col]), long = cur.length > 50 || /\n/.test(cur), el = document.createElement(long ? 'textarea' : 'input'); el.value = cur; el.style.cssText = 'width:100%;min-width:' + (long ? 280 : 120) + 'px;min-height:' + (long ? 120 : 36) + 'px;border:1px solid var(--teal);border-radius:8px;padding:6px;background:var(--card);color:var(--ink)';
    td.innerHTML = ''; td.appendChild(el); el.focus(); var finished = false;
    function end(save) { if (finished) return; finished = true; if (!save || el.value === cur) { render(); return } rpc('u_lib_edit_row', { p_id: id, p_col: col, p_val: el.value, p_req: rid() }).then(function () { toast('ƒê√£ s·ª≠a √¥'); return reload(true) }, function (e) { fail(e); render() }) }
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
    [].slice.call(document.querySelectorAll('[data-lbdel]')).forEach(function (b) { b.onclick = function () { if (!confirm('X√≥a h·∫≥n d√≤ng n√†y kh·ªèi l√¥?')) return; rpc('u_lib_del_row', { p_id: Number(b.dataset.lbdel), p_req: rid() }).then(function () { toast('ƒê√£ x√≥a d√≤ng'); return reload(true) }, fail) } })
  }
  return { page: page, bind: bind, handle: handle, state: LS, select: function (k) { LS.sel = k; LS.cheDo = P.CONFIG[k].cheDo; LS.parsed = null } }
})();

;

/* ===== scan-engine.js ===== */
/* QU√âT GI√Å & CTKM (ƒê·ª£t 2): ƒë·ªëi chi·∫øu s·ªë ti·ªÅn trong c√°c m·∫´u tin v·ªõi B·∫£ng gi√° + CTKM ƒëang hi·ªáu l·ª±c, ƒë·ªÅ xu·∫•t thay t·ª± ƒë·ªông, b√°o CTKM ch∆∞a ƒë∆∞·ª£c nh·∫Øc.
 * H√†m thu·∫ßn (kh√¥ng g·ªçi m·∫°ng) ƒë·ªÉ ki·ªÉm th·ª≠ b·∫±ng file th·∫≠t. ƒê·∫ßu v√†o: d√≤ng th∆∞ vi·ªán {id, loai, trang_thai, data, canh_bao}. ƒê·∫ßu ra: danh s√°ch m·∫´u k√®m c·∫£nh b√°o v√† b·∫£n ƒë·ªÅ xu·∫•t.
 * Nguy√™n t·∫Øc an to√†n: ch·ªâ ƒë·ªÅ xu·∫•t; thay th·∫≠t do ASM b·∫•m (m√°y ch·ªß so l·∫°i n·ªôi dung c≈© tr∆∞·ªõc khi ghi). C·∫£nh b√°o kh√¥ng ch·∫Øc ch·∫Øn (nhi·ªÅu gi√°, s·ªë l·∫°) ch·ªâ cho "S·ª≠a tay". */
(function (root) {
  var MSG_TYPES = ['MAU_TUOI', 'MAU_CHU_DE', 'MAU_SAU_TIEM', 'MAU_CHUNG'];
  var STOP = { ml: 1, inj: 1, vial: 1, ong: 1, hap: 1, phu: 1, viet: 1, nam: 1, vaccine: 1, vacxin: 1, dung: 1, dich: 1, tiem: 1, lieu: 1, hop: 1, chai: 1, bom: 1, pfs: 1, prefilled: 1, syringe: 1, ngua: 1, ung: 1, thu: 1, mui: 1, vac: 1, xin: 1 };
  var PROMO = ['giam', 'uu dai', 'ho tro', 'tang ', 'khuyen mai', 'combo', 'voucher', 'mien phi', 'qua tang', 'chuong trinh'];
  function fold(s) { var o = ''; s = String(s == null ? '' : s); for (var i = 0; i < s.length; i++) { var c = s.charAt(i), b = c.normalize('NFD').charAt(0).toLowerCase(); if (c === 'ƒë' || c === 'ƒê') b = 'd'; o += b } return o }
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
    var re = /\d{1,3}(?:[.,]\d{3})+(?:\s?(?:ƒë|‚Ç´|vnƒë|vnd|ƒë·ªìng))?/gi, m, out = [];
    while ((m = re.exec(text))) {
      var pre = text.charAt(m.index - 1), nxt = text.charAt(m.index + m[0].length);
      if (/\d/.test(pre || '') || /\d/.test(nxt || '')) continue; var v = Number(m[0].replace(/[^\d]/g, '').replace(/^0+/, '') || 0); if (v < 1000) continue;
      out.push({ text: m[0], val: v, start: m.index, end: m.index + m[0].length, unit: /(ƒë|‚Ç´|vnƒë|vnd|ƒë·ªìng)$/i.test(m[0]) })
    }
    return out
  }
  function today() { var d = new Date(); return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2) }
  function title(r) { var d = r.data || {}; return r.loai === 'MAU_CHUNG' ? ((d.ma_tin || '') + (d.ten_ngan ? ' ‚Äì ' + d.ten_ngan : '')) : (d.ten_mau || d.chu_de || ('D√≤ng ' + r.stt)) }

  /* gom ngu·ªìn s·ªë li·ªáu: b·∫£ng gi√°, CTKM, t·∫≠p "s·ªë ti·ªÅn h·ª£p l·ªá" */
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
        if (d.den_ngay && d.den_ngay < day) return; if (d.tu_ngay && d.tu_ngay > day && d.phan_loai !== '∆Øu ƒë√£i chung') return; nCK++;
        var p = progs[d.nhom] || (progs[d.nhom] = { nhom: d.nhom, phan_loai: d.phan_loai, loai_vaccine: d.loai_vaccine, noi_dung_ct: d.noi_dung_ct, rows: [], vals: {}, keys: [], uudai: d.phan_loai === '∆Øu ƒë√£i chung' });
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
        var hit = S.vals[a.val]; a.ok = !!hit; a.label = hit ? (hit.src === 'BG' ? 'B·∫£ng gi√°: ' : 'CTKM: ') + hit.label : '';
        if (a.ok) return; var o = nearest(occ, a, text), vk = o && o.key, P = vk && S.idx[vk] ? Object.keys(S.idx[vk]).map(Number) : [];
        if (vk) a.vk = vk;
        if (P.length === 1 && o.dist > 45) warns.push({ type: 'gia_nhieu', level: 'tay', msg: 'Gi√° ' + a.text + ' c√≥ th·ªÉ kh√°c gi√° hi·ªán h√†nh c·ªßa ' + S.idx[vk][P[0]] + ' (' + fmt(P[0]) + 'ƒë), nh∆∞ng t√™n vaccine n·∫±m kh√° xa s·ªë ti·ªÅn n√™n c·∫ßn anh xem l·∫°i r·ªìi s·ª≠a tay.', a: a });
        else if (P.length === 1) { a.sug = P[0]; warns.push({ type: 'gia', level: 'auto', msg: 'Gi√° ' + a.text + ' kh√°c gi√° hi·ªán h√†nh c·ªßa ' + S.idx[vk][P[0]] + ': ' + fmt(P[0]) + 'ƒë', a: a }); edits.push({ start: a.start, end: a.end, rep: fmt(P[0]) + (a.unit ? 'ƒë' : '') }) }
        else if (P.length > 1) warns.push({ type: 'gia_nhieu', level: 'tay', msg: 'Gi√° ' + a.text + ' g·∫ßn "' + vk + '" nh∆∞ng c√≥ nhi·ªÅu gi√° hi·ªán h√†nh (' + P.map(fmt).join(', ') + '). Ch·ªçn gi√° ƒë√∫ng r·ªìi s·ª≠a tay.', a: a });
        else warns.push({ type: 'so_la', level: 'tay', msg: 'S·ªë ti·ªÅn ' + a.text + ' ch∆∞a th·∫•y trong B·∫£ng gi√°/CTKM hi·ªán h√†nh.', a: a })
      });
      /* CTKM ch∆∞a n√™u */
      var seenProg = {}, hasPromoWord = PROMO.some(function (w) { return f.indexOf(w) >= 0 });
      var brand = {}; occ.forEach(function (o) { progKeyList.forEach(function (x) { if (x.k === o.key) { if (!seenProg[x.n]) seenProg[x.n] = 1; if (!x.phrase) brand[x.n] = 1 } }) });
      var matched = Object.keys(seenProg);
      matched.forEach(function (n) { var p = S.progs[n]; if (p.uudai) return; if (amts.some(function (a) { return p.vals[a.val] })) { progMentioned[n] = 1; seenProg[n] = 2 } else if (f.indexOf(fold(p.noi_dung_ct || '').slice(0, 28)) >= 0 && (p.noi_dung_ct || '').length > 8) { progMentioned[n] = 1; seenProg[n] = 2 } });
      var doneW = /da tiem|tiem du|tiem roi|da duoc tiem/.test(f), anyM = matched.some(function (n) { return seenProg[n] === 2 });
      var miss = anyM ? [] : matched.filter(function (n) { return seenProg[n] === 1 && !S.progs[n].uudai && (brand[n] || !doneW) });
      if (miss.length && !hasPromoWord) { var p = S.progs[miss[0]], sent = 'üéÅ ' + String(p.noi_dung_ct || '').replace(/\s+/g, ' ').trim(); warns.push({ type: 'ctkm', level: 'auto', msg: 'Tin c√≥ nh·∫Øc ' + (p.loai_vaccine || p.rows[0].ten_sp) + ' nh∆∞ng ch∆∞a n√™u CTKM ' + p.nhom + ' (' + String(p.noi_dung_ct || '').slice(0, 70) + ').', sent: sent, prog: p.nhom }); edits.push({ start: text.length, end: text.length, rep: '\n\n' + sent }) }
      (r.canh_bao || '').split('; ').filter(Boolean).forEach(function (c) { warns.push({ type: 'dinh_dang', level: 'tay', msg: c }) });
      var proposed = text; edits.sort(function (a, b) { return b.start - a.start }).forEach(function (e) { proposed = proposed.slice(0, e.start) + e.rep + proposed.slice(e.end) });
      msgs.push({ id: r.id, ma_lo: r.ma_lo, loai: r.loai, trang_thai: r.trang_thai, title: title(r), text: text, amts: amts, warns: warns, proposed: edits.length ? proposed : null, canApply: edits.length > 0 });
      allText += '\n' + f
    });
    var unm = Object.keys(S.progs).map(function (n) { return S.progs[n] }).filter(function (p) { return !p.uudai && !progMentioned[p.nhom] && !Object.keys(p.vals).some(function (v) { return allText.indexOf(String(v).replace(/\B(?=(\d{3})+(?!\d))/g, '.')) >= 0 }) });
    var nWarn = msgs.filter(function (m) { return m.warns.some(function (w) { return w.type !== 'dinh_dang' }) }).length;
    return { msgs: msgs, unmentioned: unm.map(function (p) { return { nhom: p.nhom, phan_loai: p.phan_loai, loai_vaccine: p.loai_vaccine, noi_dung_ct: p.noi_dung_ct, sent: 'üéÅ ' + String(p.noi_dung_ct || '').replace(/\s+/g, ' ').trim() } }), stats: { msgs: msgs.length, bg: S.nBG, ck: S.nCK, warn: nWarn, unm: unm.length, day: day } }
  }
  root.LIBSCAN = { scan: scan, amountsOf: amountsOf, fmt: fmt, fold: fold, keysOf: keysOf };
  if (typeof module !== 'undefined') module.exports = root.LIBSCAN;
})(typeof window !== 'undefined' ? window : globalThis);

;

/* ===== asm-scan.js ===== */
/* TH∆Ø VI·ªÜN ‚Äì QU√âT GI√Å & CTKM (ƒê·ª£t 2): r√† c√°c m·∫´u tin so v·ªõi B·∫£ng gi√° + CTKM hi·ªán h√†nh, ƒë·ªÅ xu·∫•t thay t·ª± ƒë·ªông, b√°o CTKM ch∆∞a m·∫´u n√†o nh·∫Øc, cho s·ª≠a tay.
 * Qu√©t ch·ªâ ƒê·ªåC. Ch·ªâ khi ASM b·∫•m "√Åp ƒë·ªÅ xu·∫•t" / "Thay c√°c m·∫´u ƒë√£ t√≠ch" / "L∆∞u" (s·ª≠a tay) m·ªõi ghi, qua h√†m u_lib_replace_text (m√°y ch·ªß so l·∫°i n·ªôi dung c≈©, ghi nh·∫≠t k√Ω c≈©/m·ªõi).
 * Logic ƒë·ªëi chi·∫øu n·∫±m ·ªü scan-engine.js. D·ªØ li·ªáu th·ª≠ l√† d·ªØ li·ªáu m·∫´u. */
var ASMSCAN = (function () {
  var SCN = { inc: true, rows: null, res: null, busy: false, filter: 'warn', sel: {}, skip: {}, open: {}, applying: false };
  var TYPES = { MAU_TUOI: 'B·∫ßu v√† Nhi', MAU_CHU_DE: 'Theo vaccine', MAU_SAU_TIEM: 'CSKH sau ti√™m', MAU_CHUNG: 'M·∫´u chung' };
  var WT = { gia: ['t-late', 'Gi√° l·ªách'], gia_nhieu: ['t-late', 'Nhi·ªÅu gi√°'], so_la: ['t-new', 'S·ªë ch∆∞a r√µ'], ctkm: ['t-today', 'Ch∆∞a n√™u CTKM'], dinh_dang: ['t-new', 'ƒê·ªãnh d·∫°ng'] };
  try { var st = document.createElement('style'); st.textContent = '.mk-ok{background:#dff3e7;border-radius:4px;padding:0 2px}.mk-bad{background:#ffd9d6;border-radius:4px;padding:0 2px}.mk-new{background:#fff0a8;border-radius:4px;padding:0 2px}.scn pre{white-space:pre-wrap;word-break:break-word;font:inherit;margin:0}details.scn>summary{cursor:pointer;list-style:none;display:flex;gap:10px;align-items:center;flex-wrap:wrap}details.scn>summary::-webkit-details-marker{display:none}'; document.head.appendChild(st) } catch (e) { }

  function mark(text, amts, cls) { var h = '', pos = 0; amts.slice().sort(function (a, b) { return a.start - b.start }).forEach(function (a) { if (a.start < pos) return; h += esc(text.slice(pos, a.start)) + '<span class="' + cls(a) + '">' + esc(text.slice(a.start, a.end)) + '</span>'; pos = a.end }); return h + esc(text.slice(pos)) }
  function scan() {
    SCN.busy = true; render();
    return rest('lib_rows?select=id,ma_lo,loai,stt,data,trang_thai,canh_bao&loai=in.(BANG_GIA,CTKM,MAU_TUOI,MAU_CHU_DE,MAU_SAU_TIEM,MAU_CHUNG)&trang_thai=in.(DANG_DUNG,CHO_DUYET)&limit=5000').then(function (rows) { SCN.rows = rows; SCN.res = LIBSCAN.scan(rows, { incPending: SCN.inc }); SCN.busy = false; SCN.sel = {}; render() }, function (e) { SCN.busy = false; render(); fail(e) })
  }
  function shown() { var r = SCN.res; if (!r) return []; return r.msgs.filter(function (m) { if (SCN.skip[m.id]) return false; if (SCN.filter === 'warn') return m.warns.some(function (w) { return w.type !== 'dinh_dang' }); if (SCN.filter === 'auto') return m.canApply; return true }) }
  function page() {
    var r = SCN.res, h = '<div class="card"><h2>Qu√©t gi√° &amp; CTKM so v·ªõi c√°c m·∫´u tin <span class="pill">ch·ªâ ƒë·ªçc cho t·ªõi khi anh b·∫•m thay</span></h2><div class="meta" style="margin-bottom:10px">ƒê·ªëi chi·∫øu B·∫£ng gi√° v√† CTKM ƒëang hi·ªáu l·ª±c h√¥m nay v·ªõi s·ªë ti·ªÅn trong m·ªçi m·∫´u tin (B·∫ßu‚ÄìNhi, theo vaccine, CSKH sau ti√™m, m·∫´u chung). B√°o tin ghi gi√° kh√°c hi·ªán h√†nh (k√®m ƒë·ªÅ xu·∫•t thay), tin nh·∫Øc vaccine nh∆∞ng ch∆∞a n√™u CTKM, v√† CTKM ch∆∞a m·∫´u n√†o nh·∫Øc t·ªõi. Vi·ªác n√†o h·ªá th·ªëng kh√¥ng ch·∫Øc th√¨ ƒë·ªÉ anh <b>s·ª≠a tay</b>. ƒê√¢y l√† g·ª£i √Ω ƒë·ªÉ anh xem x√©t.</div><div class="row" style="flex-wrap:wrap"><label class="row" style="cursor:pointer"><input type="checkbox" id="sc-inc" ' + (SCN.inc ? 'checked' : '') + ' style="width:20px;height:20px;accent-color:var(--teal)"><span>T√≠nh c·∫£ l√¥ Ch·ªù duy·ªát (xem th·ª≠)</span></label><span class="sp"></span><button class="btn b1" id="sc-run"' + (SCN.busy ? ' disabled' : '') + '>' + (SCN.busy ? 'ƒêang qu√©t‚Ä¶' : r ? 'Qu√©t l·∫°i' : 'Qu√©t') + '</button></div></div>';
    if (!r) return h + '<div class="card empty">B·∫•m "Qu√©t" ƒë·ªÉ b·∫Øt ƒë·∫ßu. C·∫ßn c√≥ B·∫£ng gi√°/CTKM v√† m·∫´u tin trong Kho t·ª´ Excel (ƒë√£ duy·ªát, ho·∫∑c tick "t√≠nh c·∫£ l√¥ ch·ªù duy·ªát").</div>';
    var st = r.stats, list = shown(), nSel = Object.keys(SCN.sel).length, nAuto = r.msgs.filter(function (m) { return m.canApply && !SCN.skip[m.id] }).length;
    h += '<div class="alert"><i>üìÖ</i><div><b>H√¥m nay ' + esc(st.day) + ' ¬∑ ƒë√£ r√† ' + st.msgs + ' m·∫´u ¬∑ ƒëang t√≠nh ' + st.bg + ' d√≤ng gi√°, ' + st.ck + ' d√≤ng CTKM</b><small>' + st.warn + ' m·∫´u c√≥ c·∫£nh b√°o ¬∑ ' + st.unm + ' CTKM ch∆∞a m·∫´u n√†o nh·∫Øc t·ªõi' + (st.bg === 0 ? ' ¬∑ <b>ch∆∞a c√≥ B·∫£ng gi√° n√†o ƒëang d√πng</b>, n√™n ch∆∞a ƒë·ªÅ xu·∫•t ƒë∆∞·ª£c vi·ªác thay gi√°' : '') + '</small></div></div>';
    h += '<div class="chips">' + [['warn', 'C√≥ c·∫£nh b√°o', r.msgs.filter(function (m) { return !SCN.skip[m.id] && m.warns.some(function (w) { return w.type !== 'dinh_dang' }) }).length], ['auto', 'C√≥ th·ªÉ t·ª± thay', nAuto], ['all', 'T·∫•t c·∫£', r.msgs.length]].map(function (x) { return '<button class="chip ' + (SCN.filter === x[0] ? 'on' : '') + '" data-scf="' + x[0] + '">' + x[1] + '<em>' + x[2] + '</em></button>' }).join('') + '</div>';
    h += '<div class="row" style="flex-wrap:wrap"><button class="btn b3" id="sc-all" style="min-height:40px">Ch·ªçn t·∫•t c·∫£ m·∫´u c√≥ ƒë·ªÅ xu·∫•t</button><button class="btn b3" id="sc-none" style="min-height:40px">B·ªè ch·ªçn</button><span class="sp"></span><button class="btn b1" id="sc-apply" style="min-height:40px"' + (nSel && !SCN.applying ? '' : ' disabled') + '>Thay c√°c m·∫´u ƒë√£ t√≠ch (' + nSel + ')</button></div>';
    h += '<div class="list" style="margin-top:10px">' + (list.length ? list.map(card).join('') : '<div class="card empty">Kh√¥ng c√≥ m·∫´u n√†o trong m·ª•c n√†y.</div>') + '</div>';
    if (st.unm) h += '<div class="card" style="padding:6px 8px"><h2 style="padding:10px 10px 0">CTKM ch∆∞a m·∫´u n√†o nh·∫Øc t·ªõi <span class="pill">' + st.unm + '</span></h2><div class="tw"><table><thead><tr><th>Ch∆∞∆°ng tr√¨nh</th><th>Lo·∫°i</th><th>N·ªôi dung</th><th>C√¢u g·ª£i √Ω ƒë∆∞a v√†o m·∫´u</th><th></th></tr></thead><tbody>' + r.unmentioned.map(function (u, i) { return '<tr><td><b>' + esc(u.nhom) + '</b></td><td>' + esc(u.phan_loai || '') + '<small>' + esc(u.loai_vaccine || '') + '</small></td><td style="white-space:normal;max-width:260px">' + esc(u.noi_dung_ct || '') + '</td><td style="white-space:normal;max-width:300px">' + esc(u.sent) + '</td><td><button class="btn b3" data-sccp="' + i + '" style="min-height:34px;font-size:13px">Ch√©p c√¢u</button></td></tr>' }).join('') + '</tbody></table></div></div>';
    return h
  }
  function card(m) {
    var oth = {}; m.amts.forEach(function (a) { oth[a.val] = 1 });
    var warnN = m.warns.filter(function (w) { return w.type !== 'dinh_dang' }).length, fmtN = m.warns.length - warnN;
    var h = '<details class="card scn" data-scid="' + m.id + '"' + (SCN.open[m.id] ? ' open' : '') + ' style="padding:12px"><summary>' + (m.canApply ? '<input type="checkbox" data-scsel="' + m.id + '" ' + (SCN.sel[m.id] ? 'checked' : '') + ' style="width:20px;height:20px;accent-color:var(--teal)">' : '') + '<b style="flex:1;min-width:160px">' + esc(m.title) + '</b>' + tag('t-new', TYPES[m.loai] || m.loai) + (m.trang_thai === 'CHO_DUYET' ? tag('t-late', 'ch·ªù duy·ªát') : '') + (warnN ? tag('t-late', warnN + ' c·∫£nh b√°o') : tag('t-ok', '·ªïn')) + (fmtN ? tag('t-new', fmtN + ' ƒë·ªãnh d·∫°ng') : '') + '</summary><div style="margin-top:10px">';
    h += '<div class="script" style="margin:0 0 8px"><pre>' + mark(m.text, m.amts, function (a) { return a.ok ? 'mk-ok' : 'mk-bad' }) + '</pre></div>';
    if (m.amts.length) h += '<div class="meta" style="margin-bottom:6px">S·ªë ti·ªÅn trong tin: ' + m.amts.map(function (a) { return '<span class="tag ' + (a.ok ? 't-ok' : 't-bad') + '">' + esc(a.text) + (a.ok ? ' ‚úì ' + esc(a.label) : ' ‚úó') + '</span>' }).join('') + '</div>';
    h += m.warns.map(function (w) { var t = WT[w.type] || ['t-new', w.type]; return '<div class="alert ' + (w.level === 'auto' ? 'a-warn' : '') + '" style="margin-bottom:6px"><div>' + tag(t[0], t[1]) + ' ' + esc(w.msg) + (w.level === 'auto' ? '<small>C√≥ ƒë·ªÅ xu·∫•t t·ª± thay</small>' : '') + '</div></div>' }).join('');
    if (m.proposed) h += '<div class="meta" style="margin:8px 0 4px"><b>B·∫£n ƒë·ªÅ xu·∫•t</b> (ph·∫ßn v√†ng l√† ch·ªó thay/th√™m)</div><div class="script" style="margin:0 0 8px;background:var(--ok-bg)"><pre>' + mark(m.proposed, LIBSCAN.amountsOf(m.proposed).filter(function (a) { return !oth[a.val] }), function () { return 'mk-new' }).replace(/(\n\n)(üéÅ[^]*)$/, '$1<span class="mk-new">$2</span>') + '</pre></div>';
    h += '<div class="acts" style="flex-wrap:wrap">' + (m.canApply ? '<button class="btn b1" data-scone="' + m.id + '" style="min-height:40px"' + (SCN.applying ? ' disabled' : '') + '>√Åp ƒë·ªÅ xu·∫•t</button>' : '') + '<button class="btn b3" data-scman="' + m.id + '" style="min-height:40px">S·ª≠a tay</button><button class="btn b3" data-scskip="' + m.id + '" style="min-height:40px">B·ªè qua</button></div></div></details>';
    return h
  }
  function tag(c, t) { return '<span class="tag ' + c + '">' + esc(t) + '</span>' }
  function byId(id) { return SCN.res.msgs.filter(function (m) { return m.id === Number(id) })[0] }
  function replace(m, text, why) { return rpc('u_lib_replace_text', { p_id: m.id, p_old: m.text, p_new: text, p_ly_do: why, p_req: rid() }) }
  function applyOne(id) { var m = byId(id); if (!m || !m.proposed) return; SCN.applying = true; render(); replace(m, m.proposed, 'Thay theo qu√©t gi√° & CTKM').then(function () { SCN.applying = false; toast('ƒê√£ thay n·ªôi dung m·∫´u "' + m.title + '"'); return scan() }, function (e) { SCN.applying = false; render(); fail(e) }) }
  function applySel() {
    var ids = Object.keys(SCN.sel); if (!ids.length) return; if (!confirm('Thay n·ªôi dung ' + ids.length + ' m·∫´u theo ƒë·ªÅ xu·∫•t? N·ªôi dung c≈© ƒë∆∞·ª£c gi·ªØ trong nh·∫≠t k√Ω.')) return; SCN.applying = true; render(); var ok = 0, bad = [];
    ids.reduce(function (p, id) { return p.then(function () { var m = byId(id); return replace(m, m.proposed, 'Thay theo qu√©t gi√° & CTKM').then(function () { ok++ }, function (e) { bad.push(m.title + ': ' + (e.message || e.code)) }) }) }, Promise.resolve()).then(function () { SCN.applying = false; toast('ƒê√£ thay ' + ok + ' m·∫´u' + (bad.length ? ', ' + bad.length + ' m·∫´u kh√¥ng thay ƒë∆∞·ª£c: ' + bad[0] : ''), 0, bad.length ? 1 : 0); return scan() })
  }
  function manual(id) {
    var m = byId(id); if (!m) return; var d = document.createElement('div'); d.className = 'modal';
    d.innerHTML = '<div class="card" style="width:100%;max-width:640px;border-radius:22px 22px 0 0;padding:18px;max-height:92vh;overflow:auto"><h2 style="font-size:18px;margin:0 0 4px">S·ª≠a tay ¬∑ ' + esc(m.title) + '</h2><div class="meta" style="margin-bottom:8px">S·ª≠a n·ªôi dung r·ªìi L∆∞u. M√°y ch·ªß ki·ªÉm l·∫°i: n·∫øu m·∫´u v·ª´a b·ªã ƒë·ªïi ·ªü n∆°i kh√°c th√¨ kh√¥ng ghi ƒë√®.</div>' + (m.warns.length ? '<div style="margin-bottom:8px">' + m.warns.slice(0, 4).map(function (w) { return '<div class="meta">‚Ä¢ ' + esc(w.msg) + '</div>' }).join('') + '</div>' : '') + '<textarea id="sc-t" style="min-height:220px">' + esc(m.text) + '</textarea><div class="acts" style="margin-top:10px"><button class="btn b3" id="sc-x" style="flex:1">H·ªßy</button>' + (m.proposed ? '<button class="btn b2" id="sc-p" style="flex:1">D√πng b·∫£n ƒë·ªÅ xu·∫•t</button>' : '') + '<button class="btn b1" id="sc-s" style="flex:1.4">L∆∞u</button></div></div>';
    document.body.appendChild(d); d.onclick = function (e) { if (e.target === d) d.remove() }; d.querySelector('#sc-x').onclick = function () { d.remove() };
    var p = d.querySelector('#sc-p'); if (p) p.onclick = function () { d.querySelector('#sc-t').value = m.proposed };
    d.querySelector('#sc-s').onclick = function () { var v = d.querySelector('#sc-t').value, b = d.querySelector('#sc-s'); if (v === m.text) { d.remove(); return } b.disabled = true; replace(m, v, 'S·ª≠a tay sau qu√©t gi√° & CTKM').then(function () { d.remove(); toast('ƒê√£ l∆∞u n·ªôi dung m·∫´u'); return scan() }, function (e) { b.disabled = false; fail(e) }) }
  }
  function bind(on) {
    on('#sc-inc', 'onchange', function (x) { SCN.inc = x.checked }); on('#sc-run', 'onclick', scan);
    [].slice.call(document.querySelectorAll('[data-scf]')).forEach(function (b) { b.onclick = function () { SCN.filter = b.dataset.scf; render() } });
    on('#sc-all', 'onclick', function () { SCN.res.msgs.forEach(function (m) { if (m.canApply && !SCN.skip[m.id]) SCN.sel[m.id] = 1 }); render() }); on('#sc-none', 'onclick', function () { SCN.sel = {}; render() }); on('#sc-apply', 'onclick', applySel);
    [].slice.call(document.querySelectorAll('[data-scsel]')).forEach(function (c) { c.onclick = function (e) { e.stopPropagation() }; c.onchange = function () { c.checked ? SCN.sel[c.dataset.scsel] = 1 : delete SCN.sel[c.dataset.scsel]; var b = document.getElementById('sc-apply'); if (b) { var n = Object.keys(SCN.sel).length; b.disabled = !n; b.textContent = 'Thay c√°c m·∫´u ƒë√£ t√≠ch (' + n + ')' } } });
    [].slice.call(document.querySelectorAll('details[data-scid]')).forEach(function (d) { d.ontoggle = function () { d.open ? SCN.open[d.dataset.scid] = 1 : delete SCN.open[d.dataset.scid] } });
    [].slice.call(document.querySelectorAll('[data-scone]')).forEach(function (b) { b.onclick = function () { applyOne(b.dataset.scone) } });
    [].slice.call(document.querySelectorAll('[data-scman]')).forEach(function (b) { b.onclick = function () { manual(b.dataset.scman) } });
    [].slice.call(document.querySelectorAll('[data-scskip]')).forEach(function (b) { b.onclick = function () { SCN.skip[b.dataset.scskip] = 1; render() } });
    [].slice.call(document.querySelectorAll('[data-sccp]')).forEach(function (b) { b.onclick = function () { copy(SCN.res.unmentioned[Number(b.dataset.sccp)].sent); toast('ƒê√£ ch√©p c√¢u g·ª£i √Ω') } })
  }
  return { page: page, bind: bind, state: SCN, scan: scan }
})();

;

/* ===== asm-hr.js ===== */
/* NH√ÇN S·ª∞ ‚Äì N·∫†P DANH S√ÅCH T·ª™ FILE TO√ÄN B·ªò NH√ÇN S·ª∞ (b·∫£n xu·∫•t c·ªßa h·ªá th·ªëng nh√¢n s·ª±): k√©o th·∫£ file Excel/CSV, app t·ª± t√¨m d√≤ng ti√™u ƒë·ªÅ (k·ªÉ c·∫£ ti√™u ƒë·ªÅ 2 t·∫ßng), nh·∫≠n c·ªôt H·ªç t√™n / Email / Ch·ª©c danh /
 * M√£ shop / C·ª≠a h√†ng / Tr·∫°ng th√°i (nhi·ªÅu c·ªôt gi·ªëng t√™n th√¨ ch·ªçn c·ªôt C√ì D·ªÆ LI·ªÜU), so v·ªõi t√†i kho·∫£n v√† Shop ƒëang c√≥ r·ªìi cho xem tr∆∞·ªõc: T·∫°o m·ªõi / C·∫≠p nh·∫≠t / Kh√¥ng ƒë·ªïi / L·ªói / B·ªè qua.
 * Anh ƒë·∫∑t "Ch·ª©c danh ‚Üí vai tr√≤" (Nh√¢n vi√™n / Qu·∫£n l√Ω / B·ªè qua) ngay tr√™n m√†n xem tr∆∞·ªõc. Shop ch∆∞a c√≥ ƒë∆∞·ª£c t·∫°o t·ª´ file (M√£ shop + t√™n c·ª≠a h√†ng).
 * B·∫•m √Åp d·ª•ng m·ªõi ghi, t·ª´ng ng∆∞·ªùi, qua h√†m ch·ªâ ASM g·ªçi ƒë∆∞·ª£c. T√†i kho·∫£n m·ªõi c√≥ m·∫≠t kh·∫©u ng·∫´u nhi√™n hi·ªán m·ªôt l·∫ßn, t·∫£i ƒë∆∞·ª£c CSV. Kh√¥ng t·ª± kh√≥a ng∆∞·ªùi v·∫Øng trong file. */
var ASMHR = (function () {
  var P = window.LIBP, H = { aoa: null, name: '', info: null, updRole: true, getSdt: false, createShops: true, onlyActive: true, rules: {}, res: null, busy: false, drag: false };
  var COLS = [
    { k: 'ho_ten', nhan: 'H·ªç t√™n', alias: ['ho va ten', 'ho ten', 'ten nhan vien', 'nhan vien', 'ho ten nhan vien', 'full name', 'name', 'ten'], bb: 1 },
    { k: 'email', nhan: 'Email', alias: ['mail', 'dia chi email', 'email cong ty', 'email nhan vien', 'email lien he'], bb: 1 },
    { k: 'chuc_danh', nhan: 'Ch·ª©c danh', alias: ['vai tro', 'chuc vu cong viec', 'vi tri', 'position', 'role', 'chuc vu'] },
    { k: 'ma_shop', nhan: 'M√£ shop', alias: ['ma co so', 'ma cua hang', 'ma nha thuoc', 'ma don vi', 'shop id'] },
    { k: 'ten_shop', nhan: 'C·ª≠a h√†ng', alias: ['ten shop', 'ten co so', 'co so lam viec', 'co so', 'nha thuoc', 'ten cua hang', 'shop', 'dia diem lam viec', 'noi lam viec', 'trung tam'] },
    { k: 'trang_thai', nhan: 'Tr·∫°ng th√°i', alias: ['tinh trang', 'trang thai lam viec'] },
    { k: 'sdt', nhan: 'S·ªë ƒëi·ªán tho·∫°i', alias: ['sdt', 'dien thoai', 'so dt', 'phone'] }, { k: 'ma_nv', nhan: 'M√£ nh√¢n vi√™n', alias: ['ma nv', 'ma so nhan vien'] }];
  var KQ = { moi: ['t-ok', 'T·∫°o m·ªõi'], sua: ['t-today', 'C·∫≠p nh·∫≠t'], giu: ['t-new', 'Kh√¥ng ƒë·ªïi'], loi: ['t-bad', 'L·ªói'], bo: ['t-new', 'B·ªè qua'] };
  var VT = { QUAN_LY: 'Qu·∫£n l√Ω', NHAN_VIEN: 'Nh√¢n vi√™n', BO: 'B·ªè qua' };

  function xlsx(cb) { if (window.XLSX) return cb(); var s = document.createElement('script'); s.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'; s.onload = cb; s.onerror = function () { toast('Kh√¥ng t·∫£i ƒë∆∞·ª£c b·ªô ƒë·ªçc Excel. Ki·ªÉm tra m·∫°ng.', 0, 1) }; document.head.appendChild(s) }
  function genPw() { var a = 'abcdefghjkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789', o = '', b = new Uint32Array(10); (window.crypto || window.msCrypto).getRandomValues(b); for (var i = 0; i < 10; i++) o += a[b[i] % a.length]; return o + '7' }
  function tag(c, t) { return '<span class="tag ' + c + '">' + esc(t) + '</span>' }
  /* SƒêT trong file Excel hay m·∫•t s·ªë 0 ƒë·∫ßu (9 ch·ªØ s·ªë) -> th√™m l·∫°i; kh√¥ng h·ª£p l·ªá -> b·ªè tr·ªëng */
  function normPhone(v) { var s = String(v == null ? '' : v).replace(/[^0-9+]/g, ''); if (s.indexOf('+84') === 0) s = '0' + s.slice(3); else if (/^84\d{9}$/.test(s)) s = '0' + s.slice(2); else if (/^[1-9]\d{8}$/.test(s)) s = '0' + s; return /^0\d{9}$/.test(s) ? s : '' }
  function core(s) { return P.bo(String(s || '').replace(/\([^)]*\)/g, '')).replace(/^(shop|trung tam|nha thuoc|co so|lc|vx)\s+/, '').replace(/^(hni|hn)\s+/, '').trim() }
  /* quy t·∫Øc m·∫∑c ƒë·ªãnh theo ch·ª©c danh (anh s·ª≠a ƒë∆∞·ª£c tr√™n m√†n xem tr∆∞·ªõc) */
  function defaultRule(t) { var b = P.bo(t); if (!b) return 'NHAN_VIEN'; if (/giam doc|(^| )asm( |$)/.test(b)) return 'BO'; if (/(^| )bac si( |$)|bao ve/.test(b)) return 'BO'; if (/truong quan ly|pho quan ly|cua hang truong|quan ly (cua hang|shop)|truong ca/.test(b)) return 'QUAN_LY'; return 'NHAN_VIEN' }
  /* M·ªôt Shop c√≥ th·ªÉ c√≥ NHI·ªÄU Qu·∫£n l√Ω (H·∫£i 24/11): m·ªçi ng∆∞·ªùi mang ch·ª©c danh ƒë∆∞·ª£c ƒë·∫∑t l√† Qu·∫£n l√Ω ƒë·ªÅu l√† Qu·∫£n l√Ω. */
  function describeEx(row) { var ex = row.ex, ch = []; if (ex.ho_ten !== row.ho_ten) ch.push('t√™n'); if (ex.id_shop !== row.id_shop) ch.push('Shop'); if (row.hasCd && ex.vai_tro !== row.vai_tro) ch.push(VT[ex.vai_tro] + ' ‚Üí ' + VT[row.vai_tro]); row.kq = ch.length ? 'sua' : 'giu'; row.ghi = ch.length ? 'ƒê·ªïi: ' + ch.join(', ') : 'Gi·ªëng t√†i kho·∫£n hi·ªán c√≥' + (ex.trang_thai === 'NGUNG' ? ' (ƒëang b·ªã kh√≥a)' : '') }

  /* t√¨m d√≤ng ti√™u ƒë·ªÅ; m·ªói c·ªôt ch·ªçn trong c√°c c·ªôt c√πng t√™n c·ªôt C√ì NHI·ªÄU D·ªÆ LI·ªÜU nh·∫•t */
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
  /* PH√ÇN T√çCH (h√†m thu·∫ßn, ki·ªÉm th·ª≠ ƒë∆∞·ª£c): aoa + {shops:[], accs:[]} + t√πy ch·ªçn */
  function parse(aoa, ctx, o) {
    var hd = findHdr(aoa), miss = COLS.filter(function (c) { return c.bb && hd.map[c.k] === undefined });
    if (!hd.n || miss.length) return { loi: 'Kh√¥ng nh·∫≠n ra c·ªôt b·∫Øt bu·ªôc: ' + miss.map(function (c) { return c.nhan }).join(', ') + '. C√°c c·ªôt ƒë·ªçc ƒë∆∞·ª£c ·ªü d√≤ng ' + (hd.r + 1) + ': ' + (aoa[hd.r] || []).filter(function (v) { return v != null && String(v).trim() }).slice(0, 14).join(' | ') };
    var cell = function (r, k) { var j = hd.map[k]; return j === undefined ? '' : String(r[j] == null ? '' : r[j]).trim() }, byEmail = {}, titles = {}, seen = {}, rows = [], newShops = {}, shops = ctx.shops.slice();
    ctx.accs.forEach(function (a) { byEmail[String(a.email).toLowerCase()] = a });
    var raw = []; for (var i = hd.r + 1; i < aoa.length; i++) { var r = aoa[i] || []; if (!r.some(function (v) { return v != null && String(v).trim() })) continue; var ho = cell(r, 'ho_ten'), em = cell(r, 'email').toLowerCase().replace(/\s+/g, ''); if (!ho && !em) continue; raw.push({ dong: i + 1, r: r, ho: ho, em: em, cd: cell(r, 'chuc_danh') }); var t = raw[raw.length - 1].cd; if (t) titles[t] = (titles[t] || 0) + 1 }
    var rules = {}; Object.keys(titles).forEach(function (t) { rules[t] = (o.rules && o.rules[t]) || defaultRule(t) });
    raw.forEach(function (x) {
      var r = x.r, rule = x.cd ? rules[x.cd] : 'NHAN_VIEN', st = cell(r, 'trang_thai'), code = cell(r, 'ma_shop'), nm = cell(r, 'ten_shop');
      var row = { dong: x.dong, ho_ten: x.ho, email: x.em, chuc_danh: x.cd, shopTxt: nm || code, ma_shop: code, ten_shop: nm, vai_tro: rule === 'QUAN_LY' ? 'QUAN_LY' : 'NHAN_VIEN', sdt: o.getSdt ? normPhone(cell(r, 'sdt')) : '' }; rows.push(row);
      if (!x.ho) { row.kq = 'loi'; row.ghi = 'Thi·∫øu h·ªç t√™n'; return }
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x.em)) { row.kq = 'loi'; row.ghi = 'Email kh√¥ng h·ª£p l·ªá'; return }
      if (seen[x.em]) { row.kq = 'loi'; row.ghi = 'Tr√πng email v·ªõi d√≤ng ' + seen[x.em]; return } seen[x.em] = x.dong;
      if (rule === 'BO') { row.kq = 'bo'; row.ghi = 'Ch·ª©c danh "' + x.cd + '" ƒë·∫∑t l√† B·ªè qua'; return }
      if (o.onlyActive && st && !/dang lam|hoat dong|^lam viec$/.test(P.bo(st))) { row.kq = 'bo'; row.ghi = 'Kh√¥ng ƒëang l√†m (' + st + ')'; return }
      var ex = byEmail[x.em]; if (ex && ex.vai_tro === 'ASM') { row.kq = 'bo'; row.ghi = 'L√† t√†i kho·∫£n ASM'; return }
      var sh = findShop(code, nm, shops), id = sh.s ? sh.s.id_shop : null;
      if (!id) { if (code && nm && o.createShops && /^[A-Za-z0-9]{3,20}$/.test(code)) { id = 'S' + code; var ns = newShops[id] || (newShops[id] = { ma: code, ten: nm, n: 0 }); ns.n++; row.shopMoi = true } else if (ex && !code && !nm) id = ex.id_shop; else { row.kq = 'loi'; row.ghi = !code && !nm ? 'Thi·∫øu Shop' : sh.n > 1 ? 'Nhi·ªÅu Shop kh·ªõp' : 'Ch∆∞a c√≥ Shop "' + (nm || code) + '" (b·∫≠t "T·∫°o Shop ch∆∞a c√≥")'; return } }
      row.id_shop = id; var shName = sh.s ? sh.s.ten_shop : (newShops[id] ? newShops[id].ten : id);
      if (!ex) { row.kq = 'moi'; row.ghi = VT[row.vai_tro] + ' ¬∑ ' + shName + (row.shopMoi ? ' (Shop m·ªõi)' : ''); row.mk = genPw() }
      else { row.ma_tk = ex.ma_tk; row.ex = ex; row.hasCd = !!x.cd && o.updRole !== false; if (!row.hasCd) row.vai_tro = ex.vai_tro; describeEx(row); if (row.sdt && !ex.sdt) { row.addSdt = true; if (row.kq === 'giu') { row.kq = 'sua'; row.ghi = 'B·ªï sung SƒêT' } else row.ghi += ', b·ªï sung SƒêT' } }
    });
    /* ch·ªâ ƒë·ªÉ b√°o: Shop trong file m√† ch∆∞a c√≥ Qu·∫£n l√Ω n√†o (c·∫£ trong file l·∫´n trong h·ªá th·ªëng) */
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
      } catch (e) { toast('Kh√¥ng ƒë·ªçc ƒë∆∞·ª£c file: ' + e.message, 0, 1) }
    })
  }
  function card() {
    var h = '<div class="card"><h2>N·∫°p danh s√°ch nh√¢n s·ª± t·ª´ file <span class="pill">k√©o file to√†n b·ªô nh√¢n s·ª± v√†o ƒë√¢y</span></h2>';
    h += '<div id="hr-drop" style="border:2px dashed ' + (H.drag ? 'var(--teal)' : 'var(--line)') + ';border-radius:16px;padding:18px;text-align:center;background:' + (H.drag ? 'var(--tint)' : 'var(--bg)') + '"><div style="font-weight:600">K√©o th·∫£ file Excel/CSV v√†o ƒë√¢y</div><div class="meta" style="margin:4px 0 10px">App t·ª± nh·∫≠n c·ªôt H·ªç t√™n, Email, Ch·ª©c danh, M√£ shop, C·ª≠a h√†ng, Tr·∫°ng th√°i (t√™n c·ªôt kh√°c m·ªôt ch√∫t v·∫´n nh·∫≠n). Xem tr∆∞·ªõc r·ªìi m·ªõi t·∫°o.</div><label class="btn b2" style="cursor:pointer">' + ic('upload') + 'Ho·∫∑c ch·ªçn file<input type="file" id="hr-file" accept=".xlsx,.xls,.csv" hidden></label></div>';
    var f = H.info; if (f && f.loi) h += '<div class="err" style="margin-top:10px">' + esc(f.loi) + '</div>';
    if (f && f.rows) {
      var c = { moi: 0, sua: 0, giu: 0, loi: 0, bo: 0 }; f.rows.forEach(function (o) { c[o.kq]++ }); var todo = c.moi + c.sua;
      h += '<div class="meta" style="margin:10px 0 6px">' + esc(H.name) + (f.sheet ? ' ¬∑ sheet ' + esc(f.sheet) : '') + ' ¬∑ d√≤ng ti√™u ƒë·ªÅ ' + f.dongTieuDe + ' ¬∑ ' + f.rows.length + ' d√≤ng</div><div class="stats stats4" style="grid-template-columns:repeat(5,1fr)">' + [['T·∫°o m·ªõi', c.moi, 'var(--ok)'], ['C·∫≠p nh·∫≠t', c.sua, 'var(--teal-d)'], ['Kh√¥ng ƒë·ªïi', c.giu, 'inherit'], ['L·ªói', c.loi, c.loi ? 'var(--bad)' : 'inherit'], ['B·ªè qua', c.bo, 'inherit']].map(function (x) { return '<div class="card st"><div class="l">' + x[0] + '</div><div class="v" style="color:' + x[2] + '">' + x[1] + '</div></div>' }).join('') + '</div>';
      h += '<div class="card pgskip" style="margin:8px 0"><h2>Ch·ª©c danh ‚Üí vai tr√≤ trong app <span class="pill">anh s·ª≠a ƒë∆∞·ª£c</span></h2><div class="tw"><table><thead><tr><th>Ch·ª©c danh trong file</th><th>S·ªë ng∆∞·ªùi</th><th>Vai tr√≤</th></tr></thead><tbody>' + f.titles.map(function (t, i) { return '<tr><td>' + esc(t.t) + '</td><td>' + t.n + '</td><td><select data-hrt="' + i + '" style="border:1px solid var(--line);border-radius:10px;padding:6px;background:var(--card);min-height:38px">' + ['NHAN_VIEN', 'QUAN_LY', 'BO'].map(function (k) { return '<option value="' + k + '"' + (t.rule === k ? ' selected' : '') + '>' + VT[k] + '</option>' }).join('') + '</select></td></tr>' }).join('') + '</tbody></table></div><div class="meta" style="padding:6px">M·ªôt Shop c√≥ th·ªÉ c√≥ nhi·ªÅu Qu·∫£n l√Ω. "B·ªè qua" = kh√¥ng t·∫°o t√†i kho·∫£n cho ch·ª©c danh ƒë√≥.</div></div>';
      h += '<div style="display:grid;gap:8px"><label class="sw"><input type="checkbox" id="hr-act" ' + (H.onlyActive ? 'checked' : '') + '><div><b>Ch·ªâ l·∫•y ng∆∞·ªùi ƒëang l√†m</b><small>B·ªè qua ng∆∞·ªùi ngh·ªâ sinh, ngh·ªâ vi·ªác‚Ä¶ theo c·ªôt Tr·∫°ng th√°i.</small></div></label><label class="sw"><input type="checkbox" id="hr-shop" ' + (H.createShops ? 'checked' : '') + '><div><b>T·∫°o Shop ch∆∞a c√≥ trong h·ªá th·ªëng (' + f.newShops.length + ' Shop)</b><small>' + (f.newShops.length ? f.newShops.slice(0, 5).map(function (s) { return esc(s.ma) + ' ¬∑ ' + esc(s.ten.slice(0, 40)) }).join('<br>') + (f.newShops.length > 5 ? '<br>‚Ä¶' : '') : 'T·∫•t c·∫£ Shop trong file ƒë√£ c√≥.') + '</small></div></label><label class="sw"><input type="checkbox" id="hr-role" ' + (H.updRole ? 'checked' : '') + '><div><b>C·∫≠p nh·∫≠t vai tr√≤ theo ch·ª©c danh cho ng∆∞·ªùi ƒë√£ c√≥ t√†i kho·∫£n</b><small>B·∫≠t: ng∆∞·ªùi ƒë√£ c√≥ t√†i kho·∫£n m√† vai tr√≤ kh√°c b·∫£ng tr√™n s·∫Ω ƒë∆∞·ª£c ƒë·ªïi. T·∫Øt: ng∆∞·ªùi ƒë√£ c√≥ gi·ªØ nguy√™n vai tr√≤, file ch·ªâ ƒë·ªÉ th√™m ng∆∞·ªùi m·ªõi v√† s·ª≠a t√™n/Shop. Ng∆∞·ªùi ƒë√∫ng t√™n, ƒë√∫ng Shop, ƒë√∫ng vai tr√≤ lu√¥n ƒë∆∞·ª£c gi·ªØ nguy√™n.</small></div></label><label class="sw"><input type="checkbox" id="hr-sdt" ' + (H.getSdt ? 'checked' : '') + '><div><b>L·∫•y SƒêT t·ª´ file (ch·ªâ ƒëi·ªÅn cho ng∆∞·ªùi ch∆∞a c√≥ SƒêT)</b><small>M·∫∑c ƒë·ªãnh T·∫ÆT v√¨ SƒêT trong file nh√¢n s·ª± th∆∞·ªùng l√† s·ªë c√° nh√¢n. SƒêT c·ªßa nh√¢n vi√™n s·∫Ω hi·ªán trong ch·ªØ k√Ω m·ªçi tin nh·∫Øn g·ª≠i kh√°ch: ƒë·ªÉ nh√¢n vi√™n t·ª± nh·∫≠p ·ªü m·ª•c T√†i kho·∫£n, ho·∫∑c anh nh·∫≠p tay ·ªü b·∫£ng T√†i kho·∫£n.</small></div></label></div>';
      if (f.noQL) h += '<div class="alert" style="margin-top:8px"><i>‚Ñπ</i><div><b>' + f.noQL + ' Shop ch∆∞a c√≥ Qu·∫£n l√Ω</b><small>Ch∆∞a c√≥ ai mang ch·ª©c danh Qu·∫£n l√Ω cho c√°c Shop n√†y. Anh ch·ªçn Qu·∫£n l√Ω sau b·∫±ng n√∫t "Ch·ªçn Qu·∫£n l√Ω" ·ªü b·∫£ng Shop b√™n d∆∞·ªõi.</small></div></div>';
      if (f.vang) h += '<div class="alert" style="margin-top:8px"><i>‚Ñπ</i><div><b>' + f.vang + ' t√†i kho·∫£n ƒëang ho·∫°t ƒë·ªông kh√¥ng c√≥ trong file</b><small>App kh√¥ng t·ª± kh√≥a hay x√≥a h·ªç. Anh t·ª± x·ª≠ l√Ω ·ªü b·∫£ng T√†i kho·∫£n b√™n d∆∞·ªõi: Kh√≥a n·∫øu ƒë√£ c√≥ l·ªãch s·ª≠ l√†m vi·ªác, X√≥a n·∫øu nh·∫≠p sai ho·∫∑c ch∆∞a l√†m g√¨.</small><details style="margin-top:6px"><summary style="cursor:pointer;font-size:13px">Xem danh s√°ch</summary><div style="font-size:13px;line-height:1.6;margin-top:4px">' + f.vangList.slice(0, 80).map(function (a) { return esc(a.ho_ten) + ' ¬∑ ' + esc(VT[a.vai_tro] || a.vai_tro) + ' ¬∑ ' + esc(shopName(a.id_shop)) }).join('<br>') + (f.vangList.length > 80 ? '<br>‚Ä¶' : '') + '</div></details></div></div>';
      h += '<div class="tw" style="margin-top:8px"><table><thead><tr><th>D√≤ng</th><th>K·∫øt qu·∫£</th><th>H·ªç t√™n</th><th>Email</th><th>Ch·ª©c danh</th><th>Shop</th><th>Ghi ch√∫</th></tr></thead><tbody>' + f.rows.map(function (o) { var k = KQ[o.kq]; return '<tr><td>' + o.dong + '</td><td>' + tag(k[0], k[1]) + '</td><td><b>' + esc(o.ho_ten) + '</b></td><td>' + esc(o.email) + '</td><td>' + esc(o.chuc_danh) + '</td><td>' + esc((o.ten_shop || o.shopTxt || '').slice(0, 44)) + '</td><td class="meta" style="white-space:normal">' + esc(o.ghi || '') + '</td></tr>' }).join('') + '</tbody></table></div>';
      h += '<div class="acts" style="margin-top:10px"><button class="btn b1" id="hr-go"' + (todo && !H.busy ? '' : ' disabled') + '>' + (H.busy ? 'ƒêang √°p d·ª•ng‚Ä¶' : '√Åp d·ª•ng ' + todo + ' thay ƒë·ªïi (' + c.moi + ' m·ªõi, ' + c.sua + ' c·∫≠p nh·∫≠t' + (f.newShops.length && H.createShops ? ', ' + f.newShops.length + ' Shop m·ªõi' : '') + ')') + '</button><button class="btn b3" id="hr-drop-x">B·ªè file n√†y</button></div>'
    }
    if (H.res) {
      var ok = H.res.filter(function (x) { return x.ok }), bad = H.res.filter(function (x) { return !x.ok }), made = ok.filter(function (x) { return x.mk });
      h += '<div class="alert" style="margin-top:10px"><i>‚úì</i><div><b>ƒê√£ xong: ' + ok.length + ' th√†nh c√¥ng' + (bad.length ? ', ' + bad.length + ' l·ªói' : '') + '</b><small>' + made.length + ' t√†i kho·∫£n m·ªõi. M·∫≠t kh·∫©u ban ƒë·∫ßu ch·ªâ hi·ªán ·ªü ƒë√¢y, t·∫£i file ƒëƒÉng nh·∫≠p ƒë·ªÉ g·ª≠i nh√¢n vi√™n v√† nh·∫Øc h·ªç ƒë·ªïi m·∫≠t kh·∫©u.</small></div>' + (made.length ? '<button class="btn b1" id="hr-csv" style="min-height:40px;margin-left:auto">' + ic('download') + 'T·∫£i file ƒëƒÉng nh·∫≠p (CSV)</button>' : '') + '</div>';
      if (bad.length) h += '<div class="err">' + bad.slice(0, 8).map(function (x) { return 'D√≤ng ' + x.dong + ' ¬∑ ' + esc(x.ho_ten) + ': ' + esc(x.msg) }).join('<br>') + '</div>'
    }
    return h + '</div>'
  }
  function apply() {
    var f = H.info, todo = f.rows.filter(function (o) { return o.kq === 'moi' || o.kq === 'sua' }); if (!todo.length || H.busy) return; if (!confirm('√Åp d·ª•ng ' + todo.length + ' thay ƒë·ªïi nh√¢n s·ª±' + (f.newShops.length && H.createShops ? ' v√† t·∫°o ' + f.newShops.length + ' Shop m·ªõi' : '') + '?')) return; H.busy = true; H.res = []; render();
    var shopsToMake = H.createShops ? f.newShops.filter(function (s) { return todo.some(function (o) { return o.shopMoi && o.id_shop === 'S' + s.ma }) }) : [];
    shopsToMake.reduce(function (p, s) { return p.then(function () { return rpc('u_asm_shop_save', { p_ma_shop: s.ma, p_ten_shop: s.ten, p_req: rid() }) }) }, Promise.resolve()).then(function () {
      return todo.reduce(function (p, o) {
        return p.then(function () {
          var req = o.kq === 'moi' ? rpc('u_asm_account_create', { p_ho_ten: o.ho_ten, p_email: o.email, p_vai_tro: o.vai_tro, p_id_shop: o.id_shop, p_mat_khau: o.mk, p_thay_ql: false, p_req: rid(), p_sdt: o.sdt || null }) : rpc('u_asm_account_update', { p_ma_tk: o.ma_tk, p_ho_ten: o.ho_ten, p_vai_tro: o.vai_tro, p_id_shop: o.id_shop, p_trang_thai: (D.acc[o.ma_tk] || {}).trang_thai || 'HOAT_DONG', p_thay_ql: false, p_req: rid(), p_sdt: o.addSdt ? o.sdt : null });
          return req.then(function () { H.res.push({ ok: true, dong: o.dong, ho_ten: o.ho_ten, email: o.email, mk: o.kq === 'moi' ? o.mk : '', vai_tro: o.vai_tro, id_shop: o.id_shop }) }, function (e) { H.res.push({ ok: false, dong: o.dong, ho_ten: o.ho_ten, msg: (e && e.message) || (e && e.code) || 'L·ªói' }) })
        })
      }, Promise.resolve())
    }).then(function () { H.busy = false; H.info = null; H.aoa = null; toast('ƒê√£ √°p d·ª•ng danh s√°ch nh√¢n s·ª±'); return load(true).then(render) }, function (e) { H.busy = false; render(); fail(e) })
  }
  function csv() {
    var shopN = function (id) { return (D.shops[id] || {}).ten_shop || id }, rows = H.res.filter(function (x) { return x.ok && x.mk }), t = 'ho_ten,email,mat_khau_ban_dau,vai_tro,shop\n' + rows.map(function (x) { return [x.ho_ten, x.email, x.mk, VT[x.vai_tro], shopN(x.id_shop)].map(function (v) { v = String(v == null ? '' : v); return /[",\n;]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v }).join(',') }).join('\n');
    var a = document.createElement('a'); a.href = URL.createObjectURL(new Blob(['Ôªø' + t], { type: 'text/csv;charset=utf-8' })); a.download = 'dang-nhap-nhan-su-' + new Date().toISOString().slice(0, 10) + '.csv'; document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove() }, 500)
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
/* ·∫¢NH, T·ªú R∆†I, VOUCHER: k√©o th·∫£ ·∫£nh/PDF (ho·∫∑c c·∫£ th∆∞ m·ª•c) v√†o, app T·ª∞ NH·∫¨N DI·ªÜN th∆∞ m·ª•c, ch·ªß ƒë·ªÅ vaccine, nh√≥m ƒë·ªëi t∆∞·ª£ng t·ª´ t√™n file/th∆∞ m·ª•c
 * (anh s·ª≠a ƒë∆∞·ª£c tr∆∞·ªõc khi t·∫£i l√™n), ki·ªÉm tr√πng b·∫±ng SHA-256, t·∫°o ·∫£nh nh·ªè, t·∫£i l√™n kho l∆∞u tr·ªØ RI√äNG T∆Ø c·ªßa Supabase, ghi th√¥ng tin qua h√†m ch·ªâ ASM g·ªçi ƒë∆∞·ª£c.
 * Nh√¢n vi√™n d√πng ·∫£nh ·ªü m√†n G·ª£i √Ω tin nh·∫Øn (msg.js): app g·ª£i √Ω ·∫£nh h·ª£p v·ªõi kh√°ch v√† vaccine, nh√¢n vi√™n b·∫•m g·ª≠i (chia s·∫ª sang Zalo) ho·∫∑c t·∫£i v·ªÅ.
 * D·ªØ li·ªáu v√† ·∫£nh th·ª≠ l√† d·ªØ li·ªáu m·∫´u. */
function fileBlob(f) { return tok().then(function (t) { return fetch(URL_ + '/storage/v1/object/authenticated/lib-files/' + f.duong_dan, { headers: { apikey: KEY, Authorization: 'Bearer ' + t } }) }).then(function (r) { if (!r.ok) throw { code: 'LOAD', message: 'Kh√¥ng m·ªü ƒë∆∞·ª£c file (' + r.status + ').' }; return r.blob() }) }
var ASMIMG = (function () {
  var IM = { files: null, loading: false, q: '', thu: '', hidden: false, queue: [], busy: false, prog: '', n: 0 };
  var THU = { VACCINE_DOI_TUONG: 'Vaccine theo ƒë·ªëi t∆∞·ª£ng', CAM_NANG_BAU: 'C·∫©m nang chƒÉm s√≥c b·∫ßu', CTKM_VOUCHER: 'Khuy·∫øn m√£i & voucher', KHAC: 'Kh√°c' };
  var NHOM = { '': '‚Äî', BAU: 'M·∫π b·∫ßu', NHI: 'Nhi (d∆∞·ªõi 9 tu·ªïi)', TRE_LON: 'Tr·∫ª l·ªõn / thanh thi·∫øu ni√™n', NGUOI_LON: 'Ng∆∞·ªùi l·ªõn', TAT_CA: 'T·∫•t c·∫£' };
  var VOCAB = [['6 trong 1', ['6in1', '6 in 1', '6 trong 1', 'hexaxim', 'infanrix']], ['Ph·∫ø c·∫ßu', ['phe cau', 'prevenar', 'synflorix', 'vaxneuvance', 'pcv']], ['Rota', ['rota', 'rotarix', 'rotateq', 'rotavin']], ['HPV', ['hpv', 'gardasil']], ['C√∫m', ['cum', 'influvac', 'vaxigrip']],
    ['S·ªüi ‚Äì Quai b·ªã ‚Äì Rubella', ['soi', 'mmr', 'rubella', 'quai bi', 'priorix']], ['Th·ªßy ƒë·∫≠u', ['thuy dau', 'varivax', 'varilrix']], ['Vi√™m gan', ['viem gan', 'vgb', 'vga', 'gan b', 'gan a', 'twinrix']], ['N√£o m√¥ c·∫ßu', ['nao mo cau', 'bexsero', 'nimenrix', 'menactra', 'menquadfi', 'vega']], ['U·ªën v√°n ‚Äì B·∫°ch h·∫ßu ‚Äì Ho g√†', ['uon van', 'bach hau', 'ho ga', 'adacel', 'boostrix', '3in1', 'tdap']],
    ['RSV', ['rsv', 'abrysvo', 'beyfortus']], ['S·ªët xu·∫•t huy·∫øt', ['sot xuat huyet', 'sxh', 'qdenga']], ['Zona', ['zona', 'shingrix']], ['Lao', ['lao', 'bcg']], ['N√£o Nh·∫≠t B·∫£n', ['nhat ban', 'jevax', 'imojev']], ['T·∫£ ‚Äì Th∆∞∆°ng h√†n', ['thuong han', 'typhim', 'vi ta']], ['D·∫°i', ['dai', 'verorab', 'speeda']], ['B·∫°i li·ªát', ['bai liet', 'ipv', 'opv']], ['Vi√™m n√£o', ['viem nao']]];
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

  /* ---------- nh·∫≠n file ---------- */
  function addFiles(items) {
    var list = items.filter(function (x) { return x.file }); if (!list.length) return; IM.busy = true; var total = list.length, done = 0;
    function next() {
      if (!list.length) { IM.busy = false; IM.prog = ''; render(); return }
      var it = list.shift(), f = it.file, mime = mimeOf(f), q = { id: 'q' + (++IM.n), file: f, rel: it.rel || '', mime: mime, size: f.size, inc: true, err: '', dup: '', thumb: null }; IM.prog = 'ƒêang ƒë·ªçc ' + (++done) + '/' + total + '‚Ä¶'; var d = detect(f.name, q.rel); q.ten = d.ten; q.thu = d.thu; q.chu_de = d.chu_de; q.nhom = d.nhom;
      if (!mime) { q.err = 'ƒê·ªãnh d·∫°ng kh√¥ng h·ªó tr·ª£ (ch·ªâ JPG, PNG, WEBP, PDF)'; q.inc = false; IM.queue.push(q); return next() }
      if (f.size > 10485760) { q.err = 'Qu√° 10 MB'; q.inc = false; IM.queue.push(q); return next() }
      f.arrayBuffer().then(sha).then(function (h) { q.sha = h; var ex = (IM.files || []).filter(function (x) { return x.sha256 === h && x.dang_dung === 'CO' })[0] || null; if (ex) q.dup = 'Tr√πng v·ªõi "' + ex.ten + '" ƒë√£ c√≥ trong kho'; else if (IM.queue.some(function (o) { return o.sha === h })) q.dup = 'Tr√πng v·ªõi file kh√°c trong danh s√°ch n√†y'; if (q.dup) q.inc = false; return thumb(f) }).then(function (t) { q.thumb = t; IM.queue.push(q); if (done % 5 === 0) render(); next() }, function () { q.err = 'Kh√¥ng ƒë·ªçc ƒë∆∞·ª£c file'; q.inc = false; IM.queue.push(q); next() })
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

  /* ---------- t·∫£i l√™n ---------- */
  function upload() {
    var todo = IM.queue.filter(function (q) { return q.inc && !q.err }); if (!todo.length || IM.busy) return; IM.busy = true; var ok = 0, bad = [], n = 0;
    todo.reduce(function (p, q) {
      return p.then(function () {
        IM.prog = 'ƒêang t·∫£i l√™n ' + (++n) + '/' + todo.length + '‚Ä¶'; render(); var safe = fold(q.file.name).replace(/[^a-z0-9._-]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 60) || 'file', path = 'files/' + Date.now().toString(36) + (Math.random().toString(36).slice(2, 6)) + '-' + safe;
        return tok().then(function (t) { return fetch(URL_ + '/storage/v1/object/lib-files/' + path, { method: 'POST', headers: { apikey: KEY, Authorization: 'Bearer ' + t, 'Content-Type': q.mime, 'cache-control': '3600' }, body: q.file }) }).then(function (r) { if (!r.ok) return r.text().then(function (x) { throw { message: 'Kho l∆∞u tr·ªØ t·ª´ ch·ªëi (' + r.status + ')' } }) })
          .then(function () { return rpc('u_file_add', { p_ten: q.ten, p_ten_file: q.file.name, p_thu_muc: q.thu, p_chu_de: q.chu_de, p_nhom: q.nhom, p_mo_ta: q.rel || '', p_mime: q.mime, p_kich_thuoc: q.size, p_sha: q.sha, p_duong_dan: path, p_anh_nho: q.thumb, p_cho_trung: false, p_req: rid() }).catch(function (e) { tok().then(function (t) { fetch(URL_ + '/storage/v1/object/lib-files/' + path, { method: 'DELETE', headers: { apikey: KEY, Authorization: 'Bearer ' + t } }) }); throw e }) })
          .then(function () { ok++; q.done = true }, function (e) { bad.push(q.ten + ': ' + ((e && e.message) || (e && e.code) || 'l·ªói')); q.err = (e && e.message) || 'L·ªói'; q.inc = false })
      })
    }, Promise.resolve()).then(function () { IM.busy = false; IM.prog = ''; IM.queue = IM.queue.filter(function (q) { return !q.done }); toast('ƒê√£ t·∫£i l√™n ' + ok + ' file' + (bad.length ? ', ' + bad.length + ' file l·ªói: ' + bad[0] : ''), 0, bad.length ? 1 : 0); return loadFiles() })
  }

  /* ---------- giao di·ªán ---------- */
  function select(name, opts, val, attr) { return '<select ' + attr + '>' + Object.keys(opts).map(function (k) { return '<option value="' + k + '"' + (k === val ? ' selected' : '') + '>' + esc(opts[k]) + '</option>' }).join('') + '</select>' }
  function page() {
    if (IM.files === null) return '<div class="card empty">ƒêang t·∫£i kho ·∫£nh‚Ä¶</div>';
    var h = '<div class="card"><h2>T·∫£i ·∫£nh, t·ªù r∆°i, voucher <span class="pill">t·ª± nh·∫≠n di·ªán t·ª´ t√™n file v√† t√™n th∆∞ m·ª•c</span></h2><div id="im-drop" style="border:2px dashed var(--line);border-radius:16px;padding:18px;text-align:center;background:var(--bg)"><div style="font-weight:600">K√©o th·∫£ ·∫£nh/PDF ho·∫∑c c·∫£ th∆∞ m·ª•c v√†o ƒë√¢y</div><div class="meta" style="margin:4px 0 10px">JPG, PNG, WEBP, PDF, t·ªëi ƒëa 10 MB m·ªói file. App ƒëo√°n th∆∞ m·ª•c (Vaccine theo ƒë·ªëi t∆∞·ª£ng / C·∫©m nang b·∫ßu / Khuy·∫øn m√£i & voucher), ch·ªß ƒë·ªÅ vaccine v√† nh√≥m ƒë·ªëi t∆∞·ª£ng; anh xem v√† s·ª≠a tr∆∞·ªõc khi t·∫£i l√™n. File tr√πng h·ªát file ƒë√£ c√≥ s·∫Ω b·ªã b·ªè qua.</div><div class="row" style="justify-content:center;flex-wrap:wrap"><label class="btn b2" style="cursor:pointer">' + ic('upload') + 'Ch·ªçn ·∫£nh/PDF<input type="file" id="im-file" multiple accept=".jpg,.jpeg,.png,.webp,.pdf,image/*,application/pdf" hidden></label><label class="btn b3" style="cursor:pointer">Ch·ªçn c·∫£ th∆∞ m·ª•c<input type="file" id="im-dir" webkitdirectory multiple hidden></label></div>' + (IM.prog ? '<div class="meta" style="margin-top:8px">' + esc(IM.prog) + '</div>' : '') + '</div></div>';
    if (IM.queue.length) {
      var inc = IM.queue.filter(function (q) { return q.inc && !q.err }).length, THs = {}; Object.keys(THU).forEach(function (k) { THs[k] = THU[k] });
      h += '<div class="card"><h2>Xem tr∆∞·ªõc v√† x√°c nh·∫≠n <span class="pill">' + IM.queue.length + ' file ¬∑ ' + inc + ' s·∫Ω t·∫£i l√™n</span></h2><datalist id="im-cd">' + VOCAB.map(function (v) { return '<option value="' + esc(v[0]) + '">' }).join('') + '</datalist><div class="tw qt"><table><thead><tr><th></th><th>·∫¢nh</th><th>T√™n hi·ªÉn th·ªã</th><th>Th∆∞ m·ª•c</th><th>Ch·ªß ƒë·ªÅ / vaccine</th><th>Nh√≥m ƒë·ªëi t∆∞·ª£ng</th><th>K·∫øt qu·∫£</th></tr></thead><tbody>' + IM.queue.map(function (q) {
        return '<tr><td><input type="checkbox" data-iq="' + q.id + '" data-f="inc" ' + (q.inc && !q.err ? 'checked' : '') + (q.err ? ' disabled' : '') + ' style="width:20px;height:20px"></td><td>' + (q.thumb ? '<img class="qthumb" src="' + q.thumb + '">' : '<div class="qthumb" style="display:grid;place-items:center;font-size:11px;color:var(--soft)">' + (q.mime === 'application/pdf' ? 'PDF' : '‚Äî') + '</div>') + '</td><td><input data-iq="' + q.id + '" data-f="ten" value="' + esc(q.ten) + '"><small>' + esc(q.rel || q.file.name) + '</small></td><td>' + select('', THU, q.thu, 'data-iq="' + q.id + '" data-f="thu"') + '</td><td><input list="im-cd" data-iq="' + q.id + '" data-f="chu_de" value="' + esc(q.chu_de) + '" placeholder="ch∆∞a nh·∫≠n ra"></td><td>' + select('', NHOM, q.nhom, 'data-iq="' + q.id + '" data-f="nhom"') + '</td><td>' + (q.err ? tag('t-bad', q.err) : q.dup ? tag('t-late', q.dup) : tag('t-ok', 'S·∫µn s√†ng')) + '</td></tr>'
      }).join('') + '</tbody></table></div><div class="acts" style="margin-top:10px"><button class="btn b1" id="im-up"' + (inc && !IM.busy ? '' : ' disabled') + '>T·∫£i l√™n ' + inc + ' file</button><button class="btn b3" id="im-clear"' + (IM.busy ? ' disabled' : '') + '>X√≥a danh s√°ch</button></div></div>'
    }
    var all = IM.files, cnt = function (k) { return all.filter(function (f) { return (!k || f.thu_muc === k) && (IM.hidden || f.dang_dung === 'CO') }).length }, q = fold(IM.q);
    var L = all.filter(function (f) { return (!IM.thu || f.thu_muc === IM.thu) && (IM.hidden || f.dang_dung === 'CO') && (!q || fold(f.ten + ' ' + (f.chu_de || '') + ' ' + (f.ten_file || '')).indexOf(q) >= 0) });
    h += '<div class="top" style="flex-wrap:wrap"><div class="search">' + ic('search') + '<input id="im-q" placeholder="T√¨m theo t√™n, ch·ªß ƒë·ªÅ‚Ä¶" value="' + esc(IM.q) + '"></div><label class="row" style="cursor:pointer"><input type="checkbox" id="im-hid" ' + (IM.hidden ? 'checked' : '') + ' style="width:20px;height:20px;accent-color:var(--teal)"><span class="meta">Hi·ªán c·∫£ ·∫£nh ƒë√£ ·∫©n</span></label></div>';
    h += '<div class="chips">' + [['', 'T·∫•t c·∫£']].concat(Object.keys(THU).map(function (k) { return [k, THU[k]] })).map(function (x) { return '<button class="chip ' + (IM.thu === x[0] ? 'on' : '') + '" data-imt="' + x[0] + '">' + esc(x[1]) + '<em>' + cnt(x[0]) + '</em></button>' }).join('') + '</div>';
    h += '<div class="list gal">' + (L.length ? L.map(function (f) { return '<div class="card"' + (f.dang_dung === 'KHONG' ? ' style="opacity:.55"' : '') + '>' + (f.anh_nho ? '<img src="' + f.anh_nho + '" alt="">' : '<div class="pdf">' + (f.mime === 'application/pdf' ? 'PDF' : 'FILE') + '</div>') + '<b style="font-size:13.5px;line-height:1.3">' + esc(f.ten) + '</b><div>' + tag('t-new', THU[f.thu_muc] || f.thu_muc) + (f.chu_de ? tag('t-today', f.chu_de) : '') + (f.nhom ? tag('t-new', NHOM[f.nhom]) : '') + (f.dang_dung === 'KHONG' ? tag('t-bad', 'ƒë√£ ·∫©n') : '') + '</div><div class="row" style="flex-wrap:wrap"><button class="btn b3" data-imv="' + f.id + '" style="min-height:34px;font-size:13px;padding:0 10px">Xem</button><button class="btn b3" data-ime="' + f.id + '" style="min-height:34px;font-size:13px;padding:0 10px">S·ª≠a</button><button class="btn b3" data-imh="' + f.id + '" style="min-height:34px;font-size:13px;padding:0 10px">' + (f.dang_dung === 'CO' ? '·∫®n' : 'Hi·ªán') + '</button><button class="btn b3" data-imd="' + f.id + '" style="min-height:34px;font-size:13px;padding:0 10px" aria-label="X√≥a">‚úï</button></div></div>' }).join('') : '<div class="card empty" style="grid-column:1/-1">Ch∆∞a c√≥ ·∫£nh n√†o. K√©o ·∫£nh v√†o khung ph√≠a tr√™n.</div>') + '</div>';
    return h
  }
  function byId(id) { return (IM.files || []).filter(function (f) { return f.id === id })[0] }
  function view(f) {
    var d = document.createElement('div'); d.className = 'modal'; d.innerHTML = '<div class="card" style="width:100%;max-width:720px;border-radius:22px 22px 0 0;padding:16px;max-height:94vh;overflow:auto;text-align:center"><div class="row2"><h2 style="margin:0;font-size:17px">' + esc(f.ten) + '</h2><button class="btn b3" id="v-x" style="min-height:38px">ƒê√≥ng</button></div><div id="v-b" class="meta" style="padding:30px">ƒêang m·ªü‚Ä¶</div></div>'; document.body.appendChild(d); d.onclick = function (e) { if (e.target === d) d.remove() }; d.querySelector('#v-x').onclick = function () { d.remove() };
    fileBlob(f).then(function (b) { var u = URL.createObjectURL(b), box = d.querySelector('#v-b'); box.style.padding = '10px 0'; box.innerHTML = f.mime === 'application/pdf' ? '<a class="btn b1" href="' + u + '" target="_blank" rel="noopener">M·ªü PDF</a>' : '<img src="' + u + '" style="max-width:100%;max-height:70vh;border-radius:12px">' }, function (e) { d.querySelector('#v-b').textContent = (e && e.message) || 'Kh√¥ng m·ªü ƒë∆∞·ª£c file.' })
  }
  function edit(f) {
    var d = document.createElement('div'); d.className = 'modal'; d.innerHTML = '<div class="card" style="width:100%;max-width:520px;border-radius:22px 22px 0 0;padding:18px"><h2 style="font-size:18px;margin:0 0 10px">S·ª≠a th√¥ng tin</h2><label class="field">T√™n hi·ªÉn th·ªã<input id="e-ten" value="' + esc(f.ten) + '"></label><div class="form2" style="margin-top:0"><div><label class="meta">Th∆∞ m·ª•c</label>' + select('', THU, f.thu_muc, 'id="e-thu"') + '</div><div><label class="meta">Nh√≥m ƒë·ªëi t∆∞·ª£ng</label>' + select('', NHOM, f.nhom || '', 'id="e-nhom"') + '</div><div style="grid-column:1/-1"><label class="meta">Ch·ªß ƒë·ªÅ / vaccine</label><input id="e-cd" list="im-cd2" value="' + esc(f.chu_de || '') + '"><datalist id="im-cd2">' + VOCAB.map(function (v) { return '<option value="' + esc(v[0]) + '">' }).join('') + '</datalist></div><div style="grid-column:1/-1"><label class="meta">Ghi ch√∫</label><input id="e-mt" value="' + esc(f.mo_ta || '') + '"></div></div><div class="acts" style="margin-top:12px"><button class="btn b3" id="e-x" style="flex:1">H·ªßy</button><button class="btn b1" id="e-ok" style="flex:1.4">L∆∞u</button></div></div>'; document.body.appendChild(d); d.onclick = function (e) { if (e.target === d) d.remove() }; d.querySelector('#e-x').onclick = function () { d.remove() };
    d.querySelector('#e-ok').onclick = function () { var b = d.querySelector('#e-ok'); b.disabled = true; save(f, { p_ten: d.querySelector('#e-ten').value, p_thu_muc: d.querySelector('#e-thu').value, p_chu_de: d.querySelector('#e-cd').value, p_nhom: d.querySelector('#e-nhom').value, p_mo_ta: d.querySelector('#e-mt').value }).then(function () { d.remove(); toast('ƒê√£ l∆∞u') }, function (e) { b.disabled = false; fail(e) }) }
  }
  function save(f, patch) { var o = { p_id: f.id, p_ten: f.ten, p_thu_muc: f.thu_muc, p_chu_de: f.chu_de || '', p_nhom: f.nhom || '', p_mo_ta: f.mo_ta || '', p_dang_dung: f.dang_dung, p_req: rid() }; for (var k in patch) o[k] = patch[k]; return rpc('u_file_update', o).then(function () { return loadFiles() }) }
  function del(f) {
    if (!confirm('X√≥a h·∫≥n "' + f.ten + '"? Kh√¥ng ho√†n t√°c ƒë∆∞·ª£c. (N·∫øu ch·ªâ mu·ªën t·∫°m kh√¥ng d√πng, b·∫•m ·∫®n.)')) return;
    rpc('u_file_delete', { p_id: f.id, p_req: rid() }).then(function (r) { tok().then(function (t) { fetch(URL_ + '/storage/v1/object/lib-files/' + r.duong_dan, { method: 'DELETE', headers: { apikey: KEY, Authorization: 'Bearer ' + t } }) }); toast('ƒê√£ x√≥a'); return loadFiles() }, fail)
  }
  function bind(on) {
    if (IM.files === null && !IM.loading) loadFiles();
    var dz = document.getElementById('im-drop'); if (dz) { ['dragenter', 'dragover'].forEach(function (ev) { dz.addEventListener(ev, function (e) { e.preventDefault(); dz.style.borderColor = 'var(--teal)'; dz.style.background = 'var(--tint)' }) }); ['dragleave', 'drop'].forEach(function (ev) { dz.addEventListener(ev, function (e) { e.preventDefault(); dz.style.borderColor = 'var(--line)'; dz.style.background = 'var(--bg)' }) }); dz.addEventListener('drop', onDrop) }
    on('#im-file', 'onchange', function (x) { addFiles([].slice.call(x.files).map(function (f) { return { file: f, rel: '' } })) });
    on('#im-dir', 'onchange', function (x) { addFiles([].slice.call(x.files).map(function (f) { return { file: f, rel: f.webkitRelativePath || '' } })) });
    on('#im-up', 'onclick', upload); on('#im-clear', 'onclick', function () { IM.queue = []; render() });
    [].slice.call(document.querySelectorAll('[data-iq]')).forEach(function (el) { el.onchange = function () { var q = IM.queue.filter(function (x) { return x.id === el.dataset.iq })[0]; if (!q) return; var k = el.dataset.f; if (k === 'inc') { q.inc = el.checked; var b = document.getElementById('im-up'); if (b) { var n = IM.queue.filter(function (x) { return x.inc && !x.err }).length; b.disabled = !n; b.textContent = 'T·∫£i l√™n ' + n + ' file' } } else q[k] = el.value } });
    on('#im-q', 'oninput', function (x) { IM.q = x.value; var p = x.selectionStart; render(); var n = document.getElementById('im-q'); if (n) { n.focus(); n.setSelectionRange(p, p) } }); on('#im-hid', 'onchange', function (x) { IM.hidden = x.checked; render() });
    [].slice.call(document.querySelectorAll('[data-imt]')).forEach(function (b) { b.onclick = function () { IM.thu = b.dataset.imt; render() } });
    [].slice.call(document.querySelectorAll('[data-imv]')).forEach(function (b) { b.onclick = function () { view(byId(b.dataset.imv)) } });
    [].slice.call(document.querySelectorAll('[data-ime]')).forEach(function (b) { b.onclick = function () { edit(byId(b.dataset.ime)) } });
    [].slice.call(document.querySelectorAll('[data-imh]')).forEach(function (b) { b.onclick = function () { var f = byId(b.dataset.imh); save(f, { p_dang_dung: f.dang_dung === 'CO' ? 'KHONG' : 'CO' }).then(function () { toast(f.dang_dung === 'CO' ? 'ƒê√£ ·∫©n' : 'ƒê√£ hi·ªán') }, fail) } });
    [].slice.call(document.querySelectorAll('[data-imd]')).forEach(function (b) { b.onclick = function () { del(byId(b.dataset.imd)) } })
  }
  return { page: page, bind: bind, state: IM, detect: detect, addFiles: addFiles }
})();

;

/* ===== asm-admin.js ===== */
/* ASM ‚Äì NH√ÇN S·ª∞ & T√ÄI KHO·∫¢N, TH∆Ø VI·ªÜN (m·∫´u tin, khuy·∫øn m√£i), PH√ÅC ƒê·ªí (vaccine theo tu·ªïi). M·ªü r·ªông ASM trong asm.js.
 * Th∆∞ vi·ªán = "N√ìI G√å" (c√¢u ch·ªØ). Ph√°c ƒë·ªì = "KHI N√ÄO" (vaccine n√†o t·ªõi tu·ªïi n√†o). Hai b√™n g·∫∑p nhau ·ªü m√†n G·ª£i √Ω tin nh·∫Øn c·ªßa nh√¢n vi√™n.
 * M·ªçi ghi qua h√†m RPC c·ªßa c∆° s·ªü d·ªØ li·ªáu (ch·ªâ ASM g·ªçi ƒë∆∞·ª£c). D·ªØ li·ªáu m·∫´u v√† c√¢u ch·ªØ l√† D·ªÆ LI·ªÜU GI·∫¢ ƒë·ªÉ th·ª≠. */
(function () {
  var X = ASM.X, U = X.u, head = U.head, sel = U.sel, ST = U.ST;
  IC.shield = '<path d="M12 3 4 6v6c0 5 3.5 8 8 9 4.5-1 8-4 8-9V6z"/><path d="m9 12 2 2 4-4"/>';
  var VAI = { QUAN_LY: 'Qu·∫£n l√Ω', NHAN_VIEN: 'Nh√¢n vi√™n' }, TT = { HOAT_DONG: 'ƒêang ho·∫°t ƒë·ªông', NGUNG: 'ƒê√£ kh√≥a' };
  var AS = { tk: { shop: '', st: 'HOAT_DONG', q: '', nosdt: false }, tl: { tab: 'kho', nhom: '', st: 'ACTIVE', q: '' }, pd: { tv: '3', tu: '30', tg: '' } };
  var NHOM = { KH_MOI: 'KH m·ªõi', CSKH_SAU_TIEM: 'CSKH sau ti√™m', BAU_NHI: 'B·∫ßu ‚Äì Nhi', CHUNG: 'Chung', NGUY_CO: 'Nguy c∆°', GOI: 'L·ªùi n√≥i khi g·ªçi' };
  var BUOC = { D: 'D', D1: 'D+1', D3: 'D+3', D7: 'D+7', TRE: 'Tr·ªÖ l·ªãch' };
  var TPLST = { DANG_DUNG: ['t-ok', 'ƒêang d√πng'], CHO_DUYET: ['t-late', 'Ch·ªù duy·ªát'], DA_THAY: ['t-new', 'ƒê√£ thay'], NGUNG_DUNG: ['t-bad', 'Ng·ª´ng d√πng'] };
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

  /* ================= NH√ÇN S·ª∞ & T√ÄI KHO·∫¢N ================= */
  function pgTk() {
    var f = AS.tk, SH = U.shops(), ql = f.q.trim().toLowerCase();
    var L = accs().filter(function (a) { return (!f.shop || a.id_shop === f.shop) && (!f.st || a.trang_thai === f.st) && (!f.nosdt || !a.sdt) && (!ql || (a.ho_ten + ' ' + a.email).toLowerCase().indexOf(ql) >= 0) })
      .sort(function (a, b) { return String(a.id_shop).localeCompare(String(b.id_shop)) || (a.vai_tro === b.vai_tro ? a.ho_ten.localeCompare(b.ho_ten) : a.vai_tro === 'QUAN_LY' ? -1 : 1) });
    var cnt = { '': accs().length, HOAT_DONG: 0, NGUNG: 0 }; accs().forEach(function (a) { cnt[a.trang_thai] = (cnt[a.trang_thai] || 0) + 1 });
    var h = head('Nh√¢n s·ª± & t√†i kho·∫£n', 'T·∫°o t√†i kho·∫£n, ƒë·ªïi Shop, ch·ªçn Qu·∫£n l√Ω cho t·ª´ng Shop (m·ªôt Shop c√≥ th·ªÉ c√≥ nhi·ªÅu Qu·∫£n l√Ω), kh√≥a ho·∫∑c m·ªü, x√≥a, ƒë·∫∑t l·∫°i m·∫≠t kh·∫©u. Ch·ªâ ASM l√†m ƒë∆∞·ª£c.', '<button class="btn b1" id="tk-new">+ Th√™m t√†i kho·∫£n</button>');
    h += window.ASMHR ? ASMHR.card() : '';
    h += '<div class="card" style="padding:6px 8px"><h2 style="padding:10px 10px 0">Qu·∫£n l√Ω c·ªßa t·ª´ng Shop</h2><div class="tw"><table><thead><tr><th>Shop</th><th>Qu·∫£n l√Ω</th><th>Nh√¢n vi√™n</th><th></th></tr></thead><tbody>' + SH.map(function (s) {
      var ms = qlList(s.id_shop), nv = accs().filter(function (a) { return a.id_shop === s.id_shop && a.vai_tro === 'NHAN_VIEN' && a.trang_thai === 'HOAT_DONG' }).length;
      return '<tr><td><b>' + esc(s.ten_shop) + '</b><small>' + esc(s.ma_shop) + '</small></td><td>' + (ms.length ? ms.map(function (m) { return esc(m.ho_ten) }).join(', ') : '<span class="tag t-late">Ch∆∞a c√≥ Qu·∫£n l√Ω</span>') + '</td><td>' + nv + '</td><td><button class="btn b3" data-setql="' + esc(s.id_shop) + '" style="min-height:36px;font-size:13px">' + 'Ch·ªçn Qu·∫£n l√Ω' + '</button></td></tr>'
    }).join('') + '</tbody></table></div></div>';
    h += '<div class="top" style="flex-wrap:wrap"><div class="search">' + ic('search') + '<input id="tk-q" placeholder="T√¨m t√™n ho·∫∑c email‚Ä¶" value="' + esc(f.q) + '"></div>' + sel('tk-shop', [['', 'T·∫•t c·∫£ Shop']].concat(SH.map(function (s) { return [s.id_shop, s.ten_shop] })), f.shop) + '</div>';
    var noS = accs().filter(function (a) { return a.trang_thai === 'HOAT_DONG' && !a.sdt }).length;
    h += '<div class="chips">' + [['HOAT_DONG', 'ƒêang ho·∫°t ƒë·ªông'], ['NGUNG', 'ƒê√£ kh√≥a'], ['', 'T·∫•t c·∫£']].map(function (x) { return '<button class="chip ' + (f.st === x[0] ? 'on' : '') + '" data-tkst="' + x[0] + '">' + x[1] + '<em>' + (cnt[x[0]] || 0) + '</em></button>' }).join('') + '<button class="chip ' + (f.nosdt ? 'on' : '') + '" data-tknosdt="1" title="Ng∆∞·ªùi ch∆∞a c√≥ SƒêT th√¨ tin nh·∫Øn c·ªßa h·ªç thi·∫øu s·ªë li√™n h·ªá">Ch∆∞a c√≥ SƒêT<em>' + noS + '</em></button></div>';
    h += '<div class="card" style="padding:6px 8px"><div class="tw"><table><thead><tr><th>H·ªç t√™n</th><th>Vai tr√≤</th><th>Shop</th><th>SƒêT</th><th>ƒêang gi·ªØ</th><th>Tr·∫°ng th√°i</th><th>M·∫≠t kh·∫©u</th><th>ƒêƒÉng nh·∫≠p g·∫ßn nh·∫•t</th><th style="min-width:270px"></th></tr></thead><tbody>' + L.map(function (a) {
      var n = openN(a.ma_tk);
      return '<tr><td><b>' + esc(a.ho_ten) + '</b><small>' + esc(a.email) + '</small></td><td>' + (a.vai_tro === 'QUAN_LY' ? tag('t-today', 'Qu·∫£n l√Ω') : 'Nh√¢n vi√™n') + '</td><td>' + esc(shopName(a.id_shop)) + '</td><td>' + (a.sdt ? esc(fph(a.sdt)) : '<span class="tag t-late">Ch∆∞a c√≥</span>') + '</td><td>' + (n ? n + ' l∆∞·ª£t m·ªü' : '‚Äì') + '</td><td>' + (a.trang_thai === 'HOAT_DONG' ? tag('t-ok', 'Ho·∫°t ƒë·ªông') : tag('t-bad', 'ƒê√£ kh√≥a')) + '</td><td>' + (a.mk_doi_luc ? tag('t-ok', 'ƒê√£ ƒë·ªïi') + '<small>' + esc(fdt(a.mk_doi_luc)) + '</small>' : '<span class="tag t-late">Ch∆∞a ƒë·ªïi</span>') + '</td><td>' + (a.dang_nhap_luc ? esc(fdt(a.dang_nhap_luc)) : '<span class="meta">Ch∆∞a ƒëƒÉng nh·∫≠p</span>') + '</td><td><div class="row" style="flex-wrap:wrap"><button class="btn b3" data-tkedit="' + esc(a.ma_tk) + '" style="min-height:36px;font-size:13px">S·ª≠a</button><button class="btn b3" data-tkpw="' + esc(a.ma_tk) + '" style="min-height:36px;font-size:13px">M·∫≠t kh·∫©u</button><button class="btn b3" data-tklock="' + esc(a.ma_tk) + '" style="min-height:36px;font-size:13px">' + (a.trang_thai === 'HOAT_DONG' ? 'Kh√≥a' : 'M·ªü kh√≥a') + '</button><button class="btn b3" data-tkdel="' + esc(a.ma_tk) + '" style="min-height:36px;font-size:13px;color:var(--bad)">X√≥a</button></div></td></tr>'
    }).join('') + (L.length ? '' : '<tr><td colspan="9" class="empty">Kh√¥ng c√≥ t√†i kho·∫£n ph√π h·ª£p.</td></tr>') + '</tbody></table></div></div>';
    return h
  }
  function creds(title, email, pw) {
    var d = modal(), url = location.origin + location.pathname, txt = 'ƒêƒÉng nh·∫≠p LC TCQG HUB\nLink: ' + url + '\nEmail: ' + email + '\nM·∫≠t kh·∫©u: ' + pw;
    sheet(d, '<h2 style="font-size:18px;margin:0 0 6px">' + esc(title) + '</h2><div class="meta" style="margin-bottom:10px">M·∫≠t kh·∫©u ch·ªâ hi·ªán m·ªôt l·∫ßn ·ªü ƒë√¢y. G·ª≠i cho nh√¢n vi√™n v√† nh·∫Øc ƒë·ªïi m·∫≠t kh·∫©u ·ªü m·ª•c T√†i kho·∫£n.</div><div class="alert" style="display:block;font-size:15px;line-height:1.7"><div>Email: <b>' + esc(email) + '</b></div><div>M·∫≠t kh·∫©u: <b style="font-size:18px;letter-spacing:.5px">' + esc(pw) + '</b></div></div><div class="acts" style="margin-top:8px"><button class="btn b1" id="cc">Ch√©p th√¥ng tin ƒëƒÉng nh·∫≠p</button><button class="btn b3" id="cx">ƒê√≥ng</button></div>', 480);
    q(d, '#cc').onclick = function () { copy(txt); toast('ƒê√£ ch√©p th√¥ng tin ƒëƒÉng nh·∫≠p') }; q(d, '#cx').onclick = function () { d.remove() }
  }
  function accModal(a) {
    var isNew = !a, SH = U.shops(), m = { ho_ten: a ? a.ho_ten : '', email: a ? a.email : '', vai_tro: a ? a.vai_tro : 'NHAN_VIEN', id_shop: a ? a.id_shop : (ST.shop || (SH[0] || {}).id_shop || ''), trang_thai: a ? a.trang_thai : 'HOAT_DONG', mk: isNew ? genPw() : '', thay: false, sdt: a ? a.sdt || '' : '' }, d = modal();
    function draw() {
      sheet(d, '<h2 style="font-size:18px;margin:0 0 10px">' + (isNew ? 'Th√™m t√†i kho·∫£n' : 'S·ª≠a t√†i kho·∫£n') + '</h2><div class="form2" style="margin-top:0"><div><label class="meta">H·ªç t√™n</label><input id="f-ten" value="' + esc(m.ho_ten) + '"></div><div><label class="meta">Email ƒëƒÉng nh·∫≠p</label><input id="f-em" type="email" value="' + esc(m.email) + '"' + (isNew ? '' : ' disabled') + '></div><div><label class="meta">Vai tr√≤</label>' + sel('f-vt', [['NHAN_VIEN', 'Nh√¢n vi√™n'], ['QUAN_LY', 'Qu·∫£n l√Ω']], m.vai_tro) + '</div><div><label class="meta">Shop</label>' + sel('f-sh', SH.map(function (s) { return [s.id_shop, s.ten_shop] }), m.id_shop) + '</div><div style="grid-column:1/-1"><label class="meta">SƒêT nh√¢n vi√™n (hi·ªán trong ch·ªØ k√Ω m·ªçi tin nh·∫Øn g·ª≠i kh√°ch, 10 s·ªë)</label><input id="f-sdt" type="tel" inputmode="tel" placeholder="V√≠ d·ª• 0912345678" value="' + esc(m.sdt) + '" autocomplete="off"></div>' +
        (isNew ? '<div style="grid-column:1/-1"><label class="meta">M·∫≠t kh·∫©u ban ƒë·∫ßu (t·ªëi thi·ªÉu 8 k√Ω t·ª±)</label><div class="row"><input id="f-mk" value="' + esc(m.mk) + '" autocomplete="off"><button class="btn b3" id="f-gen" style="min-height:44px">T·∫°o ng·∫´u nhi√™n</button></div></div>' : '<div><label class="meta">Tr·∫°ng th√°i</label>' + sel('f-tt', [['HOAT_DONG', 'ƒêang ho·∫°t ƒë·ªông'], ['NGUNG', 'ƒê√£ kh√≥a']], m.trang_thai) + '</div>') + '</div>' +
        (!isNew && a.vai_tro !== m.vai_tro || !isNew && a.id_shop !== m.id_shop ? '<div class="ro" style="margin-top:10px">L∆∞·ª£t ƒëang gi·ªØ v·∫´n gi·ªØ t√™n ng∆∞·ªùi n√†y v√† v·∫´n ghi Shop l√∫c giao. Qu·∫£n l√Ω c≈© c·ªßa Shop s·∫Ω kh√¥ng c√≤n xem ƒë∆∞·ª£c n·∫øu ng∆∞·ªùi n√†y chuy·ªÉn Shop.</div>' : '') +
        '<div class="acts" style="margin-top:12px"><button class="btn b3" id="f-x" style="flex:1">H·ªßy</button><button class="btn b1" id="f-ok" style="flex:1.4">' + (isNew ? 'T·∫°o t√†i kho·∫£n' : 'L∆∞u') + '</button></div>');
      q(d, '#f-ten').oninput = function (e) { m.ho_ten = e.target.value }; q(d, '#f-sdt').oninput = function (e) { m.sdt = e.target.value }; var em = q(d, '#f-em'); em.oninput = function (e) { m.email = e.target.value };
      q(d, '#f-vt').onchange = function (e) { m.vai_tro = e.target.value; m.thay = false; m.ho_ten = q(d, '#f-ten').value; m.email = em.value; draw() }; q(d, '#f-sh').onchange = function (e) { m.id_shop = e.target.value; m.thay = false; m.ho_ten = q(d, '#f-ten').value; m.email = em.value; draw() };
      var tt = q(d, '#f-tt'); if (tt) tt.onchange = function (e) { m.trang_thai = e.target.value; m.ho_ten = q(d, '#f-ten').value; draw() };
      var mk = q(d, '#f-mk'); if (mk) { mk.oninput = function (e) { m.mk = e.target.value }; q(d, '#f-gen').onclick = function () { m.mk = genPw(); mk.value = m.mk } }
      q(d, '#f-x').onclick = function () { d.remove() };
      q(d, '#f-ok').onclick = function () {
        m.ho_ten = q(d, '#f-ten').value; m.email = em.value; if (mk) m.mk = mk.value; var b = q(d, '#f-ok'); b.disabled = true;
        var p = isNew ? rpc('u_asm_account_create', { p_ho_ten: m.ho_ten, p_email: m.email, p_vai_tro: m.vai_tro, p_id_shop: m.id_shop, p_mat_khau: m.mk, p_thay_ql: m.thay, p_req: rid(), p_sdt: m.sdt || null }) : rpc('u_asm_account_update', { p_ma_tk: a.ma_tk, p_ho_ten: m.ho_ten, p_vai_tro: m.vai_tro, p_id_shop: m.id_shop, p_trang_thai: m.trang_thai, p_thay_ql: m.thay, p_req: rid(), p_sdt: m.sdt || '' });
        p.then(function (r) { d.remove(); if (isNew) { load(true).then(render); creds('ƒê√£ t·∫°o t√†i kho·∫£n', r.email, m.mk) } else done('ƒê√£ l∆∞u' + (m.trang_thai === 'NGUNG' && r.luot_dang_giu ? '. Ng∆∞·ªùi n√†y c√≤n gi·ªØ ' + r.luot_dang_giu + ' l∆∞·ª£t m·ªü, h√£y thu h·ªìi ho·∫∑c giao l·∫°i ·ªü m·ª•c Ph√¢n c√¥ng.' : '')) }, function (e) { b.disabled = false; fail(e) })
      }
    }
    draw()
  }
  function pwAsm(a) {
    var d = modal(), pw = genPw();
    sheet(d, '<h2 style="font-size:18px;margin:0 0 6px">ƒê·∫∑t l·∫°i m·∫≠t kh·∫©u</h2><div class="meta" style="margin-bottom:10px">' + esc(a.ho_ten) + ' ¬∑ ' + esc(a.email) + '. M·∫≠t kh·∫©u c≈© s·∫Ω kh√¥ng d√πng ƒë∆∞·ª£c n·ªØa.</div><label class="field">M·∫≠t kh·∫©u m·ªõi (t·ªëi thi·ªÉu 8 k√Ω t·ª±)<div class="row"><input id="p-pw" value="' + esc(pw) + '" autocomplete="off" style="flex:1"><button class="btn b3" id="p-gen" style="min-height:44px">T·∫°o ng·∫´u nhi√™n</button></div></label><div class="acts"><button class="btn b3" id="p-x" style="flex:1">H·ªßy</button><button class="btn b1" id="p-ok" style="flex:1.4">ƒê·∫∑t l·∫°i</button></div>', 480);
    q(d, '#p-gen').onclick = function () { q(d, '#p-pw').value = genPw() }; q(d, '#p-x').onclick = function () { d.remove() };
    q(d, '#p-ok').onclick = function () { var v = q(d, '#p-pw').value, b = q(d, '#p-ok'); b.disabled = true; rpc('u_asm_account_reset_pw', { p_ma_tk: a.ma_tk, p_mat_khau: v, p_req: rid() }).then(function () { d.remove(); creds('ƒê√£ ƒë·∫∑t l·∫°i m·∫≠t kh·∫©u', a.email, v) }, function (e) { b.disabled = false; fail(e) }) }
  }
  function lockAsm(a) {
    var on = a.trang_thai === 'HOAT_DONG', n = openN(a.ma_tk), d = modal();
    sheet(d, '<h2 style="font-size:18px;margin:0 0 6px">' + (on ? 'Kh√≥a t√†i kho·∫£n' : 'M·ªü kh√≥a t√†i kho·∫£n') + '</h2><div class="meta" style="margin-bottom:10px">' + esc(a.ho_ten) + ' ¬∑ ' + esc(a.email) + '</div>' + (on ? '<div class="ro">Ng∆∞·ªùi n√†y s·∫Ω kh√¥ng ƒëƒÉng nh·∫≠p ƒë∆∞·ª£c n·ªØa. ' + (n ? '<b>ƒêang gi·ªØ ' + n + ' l∆∞·ª£t m·ªü</b>: c√°c l∆∞·ª£t v·∫´n mang t√™n ng∆∞·ªùi n√†y, h√£y thu h·ªìi ho·∫∑c giao l·∫°i ·ªü m·ª•c Ph√¢n c√¥ng.' : 'Hi·ªán kh√¥ng gi·ªØ l∆∞·ª£t m·ªü n√†o.') + ' L·ªãch s·ª≠ lu√¥n ƒë∆∞·ª£c gi·ªØ.</div>' : '<div class="ro">Ng∆∞·ªùi n√†y s·∫Ω ƒëƒÉng nh·∫≠p l·∫°i ƒë∆∞·ª£c v·ªõi m·∫≠t kh·∫©u hi·ªán t·∫°i.</div>') + '<div class="acts" style="margin-top:10px"><button class="btn b3" id="l-x" style="flex:1">H·ªßy</button><button class="btn b1" id="l-ok" style="flex:1.4">' + (on ? 'X√°c nh·∫≠n kh√≥a' : 'M·ªü kh√≥a') + '</button></div>', 480);
    q(d, '#l-x').onclick = function () { d.remove() };
    q(d, '#l-ok').onclick = function () { var b = q(d, '#l-ok'); b.disabled = true; rpc('u_asm_account_update', { p_ma_tk: a.ma_tk, p_ho_ten: a.ho_ten, p_vai_tro: a.vai_tro, p_id_shop: a.id_shop, p_trang_thai: on ? 'NGUNG' : 'HOAT_DONG', p_thay_ql: false, p_req: rid() }).then(function () { d.remove(); done(on ? 'ƒê√£ kh√≥a ' + a.ho_ten : 'ƒê√£ m·ªü kh√≥a ' + a.ho_ten) }, function (e) { b.disabled = false; fail(e) }) }
  }
  function setQl(shopId) {
    var people = accs().filter(function (a) { return a.id_shop === shopId && a.trang_thai === 'HOAT_DONG' }).sort(function (a, b) { return a.vai_tro === b.vai_tro ? a.ho_ten.localeCompare(b.ho_ten) : a.vai_tro === 'QUAN_LY' ? -1 : 1 }), d = modal();
    sheet(d, '<h2 style="font-size:18px;margin:0 0 6px">Ch·ªçn Qu·∫£n l√Ω ¬∑ ' + esc(shopName(shopId)) + '</h2><div class="meta" style="margin-bottom:10px">Tick nh·ªØng ng∆∞·ªùi l√†m Qu·∫£n l√Ω c·ªßa Shop n√†y. M·ªôt Shop c√≥ th·ªÉ c√≥ nhi·ªÅu Qu·∫£n l√Ω. B·ªè tick = chuy·ªÉn ng∆∞·ªùi ƒë√≥ v·ªÅ Nh√¢n vi√™n.</div>' + (people.length ? '<div style="display:grid;gap:6px;margin-bottom:10px">' + people.map(function (a) { return '<label class="row" style="cursor:pointer"><input type="checkbox" data-qw="' + esc(a.ma_tk) + '"' + (a.vai_tro === 'QUAN_LY' ? ' checked' : '') + '><span><b>' + esc(a.ho_ten) + '</b><small style="display:block">' + esc(a.email) + '</small></span></label>' }).join('') + '</div>' : '<div class="ro">Shop n√†y ch∆∞a c√≥ ai. H√£y th√™m t√†i kho·∫£n tr∆∞·ªõc.</div>') + '<div class="acts"><button class="btn b3" id="q-x" style="flex:1">H·ªßy</button>' + (people.length ? '<button class="btn b1" id="q-ok" style="flex:1.4">L∆∞u</button>' : '') + '</div>', 480);
    q(d, '#q-x').onclick = function () { d.remove() };
    var ok = q(d, '#q-ok'); if (ok) ok.onclick = function () {
      var chg = people.map(function (a) { return { a: a, on: q(d, '[data-qw="' + a.ma_tk + '"]').checked } }).filter(function (x) { return x.on !== (x.a.vai_tro === 'QUAN_LY') });
      if (!chg.length) { d.remove(); return } ok.disabled = true;
      chg.reduce(function (p, x) { return p.then(function () { return rpc('u_asm_account_update', { p_ma_tk: x.a.ma_tk, p_ho_ten: x.a.ho_ten, p_vai_tro: x.on ? 'QUAN_LY' : 'NHAN_VIEN', p_id_shop: x.a.id_shop, p_trang_thai: 'HOAT_DONG', p_thay_ql: false, p_req: rid() }) }) }, Promise.resolve())
        .then(function () { d.remove(); done('ƒê√£ c·∫≠p nh·∫≠t Qu·∫£n l√Ω c·ªßa ' + shopName(shopId)) }, function (e) { ok.disabled = false; fail(e) })
    }
  }
  function delAsm(a) {
    var n = openN(a.ma_tk), d = modal();
    sheet(d, '<h2 style="font-size:18px;margin:0 0 6px">X√≥a t√†i kho·∫£n</h2><div class="meta" style="margin-bottom:10px">' + esc(a.ho_ten) + ' ¬∑ ' + esc(a.email) + '</div><div class="ro">X√≥a h·∫≥n t√†i kho·∫£n n√†y, <b>kh√¥ng ho√†n t√°c ƒë∆∞·ª£c</b>. Ch·ªâ x√≥a ƒë∆∞·ª£c khi ng∆∞·ªùi n√†y ch∆∞a c√≥ l·ªãch s·ª≠ l√†m vi·ªác (ch∆∞a ƒë∆∞·ª£c giao l∆∞·ª£t, ch∆∞a li√™n h·ªá kh√°ch) ‚Äî v√≠ d·ª• nh·∫≠p sai ho·∫∑c ngh·ªâ tr∆∞·ªõc khi l√†m g√¨. N·∫øu ƒë√£ c√≥ l·ªãch s·ª≠, h·ªá th·ªëng t·ª´ ch·ªëi x√≥a v√† anh d√πng <b>Kh√≥a</b> ƒë·ªÉ gi·ªØ l·ªãch s·ª≠ chƒÉm s√≥c kh√°ch.' + (n ? ' <b>Ng∆∞·ªùi n√†y ƒëang gi·ªØ ' + n + ' l∆∞·ª£t m·ªü.</b>' : '') + '</div><div class="acts" style="margin-top:10px"><button class="btn b3" id="x-x" style="flex:1">H·ªßy</button><button class="btn b1" id="x-ok" style="flex:1.4;background:var(--bad)">X√≥a h·∫≥n</button></div>', 480);
    q(d, '#x-x').onclick = function () { d.remove() };
    q(d, '#x-ok').onclick = function () { var b = q(d, '#x-ok'); b.disabled = true; rpc('u_asm_account_delete', { p_ma_tk: a.ma_tk, p_req: rid() }).then(function () { d.remove(); done('ƒê√£ x√≥a ' + a.ho_ten) }, function (e) { b.disabled = false; fail(e) }) }
  }

  /* ================= TH∆Ø VI·ªÜN ================= */
  function used(ma) { return (D.mlog || []).filter(function (x) { return x.ma_mau === ma }).length }
  function pgTl() {
    var f = AS.tl, h = head('Th∆∞ vi·ªán', f.tab === 'kho' ? 'N·ªôi dung nh√¢n vi√™n d√πng khi nh·∫Øn v√† g·ªçi. N·∫°p t·ª´ file Excel c·ªßa c√¥ng ty: xem tr∆∞·ªõc, l∆∞u ch·ªù duy·ªát, duy·ªát th√¨ nh√¢n vi√™n m·ªõi th·∫•y. Ph√°c ƒë·ªì (khi n√†o ti√™m g√¨) ·ªü m·ª•c ri√™ng, s·∫Ω b·ªï sung sau.' : f.tab === 'anh' ? '·∫¢nh, t·ªù r∆°i, voucher: k√©o ·∫£nh v√†o, app t·ª± nh·∫≠n di·ªán ƒë·ªÉ nh√¢n vi√™n d√πng ·ªü n√∫t G·ª£i √Ω tin nh·∫Øn.' : f.tab === 'quet' ? 'R√† s·ªë ti·ªÅn v√† CTKM trong c√°c m·∫´u tin so v·ªõi B·∫£ng gi√° v√† CTKM hi·ªán h√†nh, thay t·ª± ƒë·ªông ho·∫∑c s·ª≠a tay.' : 'M·∫´u so·∫°n tay (t·∫°m): hi·ªán ch·ªâ c√≤n d√πng cho ch·∫ø ƒë·ªô "Theo quy tr√¨nh" (nh√≥m KH m·ªõi, 4 m·ªëc D, D+1, D+3, D+7 v√† Tr·ªÖ l·ªãch) v√† l·ªùi n√≥i khi g·ªçi. C√°c nh√≥m kh√°c (CSKH sau ti√™m, B·∫ßu‚ÄìNhi, Chung, Nguy c∆°) ƒë√£ chuy·ªÉn sang Kho Excel, nh√¢n vi√™n kh√¥ng c√≤n th·∫•y ·ªü ƒë√¢y.', f.tab === 'mau' ? '<button class="btn b1" id="tl-new">+ Th√™m m·∫´u</button>' : f.tab === 'km' ? '<button class="btn b1" id="km-new">+ Th√™m khuy·∫øn m√£i</button>' : '');
    h += '<div class="chips"><button class="chip ' + (f.tab === 'kho' ? 'on' : '') + '" data-tltab="kho">Kho t·ª´ Excel</button><button class="chip ' + (f.tab === 'quet' ? 'on' : '') + '" data-tltab="quet">Qu√©t gi√° &amp; CTKM</button><button class="chip ' + (f.tab === 'anh' ? 'on' : '') + '" data-tltab="anh">·∫¢nh, t·ªù r∆°i, voucher</button><button class="chip ' + (f.tab === 'mau' ? 'on' : '') + '" data-tltab="mau">M·∫´u so·∫°n tay (t·∫°m)<em>' + D.tpl.length + '</em></button><button class="chip ' + (f.tab === 'km' ? 'on' : '') + '" data-tltab="km">Khuy·∫øn m√£i so·∫°n tay (t·∫°m)<em>' + D.promos.length + '</em></button></div>';
    return h + (f.tab === 'kho' ? ASMLIB.page() : f.tab === 'quet' ? ASMSCAN.page() : f.tab === 'anh' ? ASMIMG.page() : f.tab === 'mau' ? tplList() : promoList())
  }
  function tplList() {
    var f = AS.tl, ql = f.q.trim().toLowerCase(), all = D.tpl;
    var L = all.filter(function (x) { return (!f.nhom || x.nhom === f.nhom) && (f.st === '' || (f.st === 'ACTIVE' ? (x.trang_thai === 'DANG_DUNG' || x.trang_thai === 'CHO_DUYET') : x.trang_thai === f.st)) && (!ql || (x.tieu_de + ' ' + x.noi_dung).toLowerCase().indexOf(ql) >= 0) });
    var h = '<div class="top" style="flex-wrap:wrap"><div class="search">' + ic('search') + '<input id="tl-q" placeholder="T√¨m trong m·∫´u‚Ä¶" value="' + esc(f.q) + '"></div>' + sel('tl-st', [['ACTIVE', 'ƒêang d√πng v√† ch·ªù duy·ªát'], ['DANG_DUNG', 'ƒêang d√πng'], ['CHO_DUYET', 'Ch·ªù duy·ªát'], ['DA_THAY', 'ƒê√£ thay'], ['NGUNG_DUNG', 'Ng·ª´ng d√πng'], ['', 'T·∫•t c·∫£']], f.st) + '</div>';
    h += '<div class="chips">' + [['', 'T·∫•t c·∫£']].concat(Object.keys(NHOM).map(function (k) { return [k, NHOM[k]] })).map(function (x) { return '<button class="chip ' + (f.nhom === x[0] ? 'on' : '') + '" data-tlg="' + x[0] + '">' + x[1] + '<em>' + all.filter(function (t) { return !x[0] || t.nhom === x[0] }).length + '</em></button>' }).join('') + '</div>';
    h += '<div class="card" style="padding:6px 8px"><div class="tw"><table><thead><tr><th>M·∫´u</th><th>Nh√≥m</th><th>B∆∞·ªõc</th><th>Tr·∫°ng th√°i</th><th>ƒê√£ d√πng</th><th style="min-width:190px"></th></tr></thead><tbody>' + L.map(function (x) {
      var st = TPLST[x.trang_thai] || ['t-new', x.trang_thai], btn = function (k, t) { return '<button class="btn b3" data-tp' + k + '="' + esc(x.ma) + '" style="min-height:36px;font-size:13px">' + t + '</button>' };
      return '<tr><td style="max-width:420px"><b>' + esc(x.tieu_de) + '</b><small>' + esc(x.ma) + (x.phien_ban > 1 ? ' ¬∑ b·∫£n ' + x.phien_ban : '') + (x.thay_cho ? ' ¬∑ thay ' + esc(x.thay_cho) : '') + '</small><small style="white-space:normal">' + esc(MSG.preview(x.noi_dung).slice(0, 140)) + (x.noi_dung.length > 140 ? '‚Ä¶' : '') + '</small></td><td>' + esc(NHOM[x.nhom] || x.nhom) + '</td><td>' + esc(BUOC[x.buoc] || '‚Äì') + '</td><td>' + tag(st[0], st[1]) + '</td><td>' + used(x.ma) + '</td><td><div class="row" style="flex-wrap:wrap">' + (x.trang_thai === 'CHO_DUYET' ? btn('edit', 'S·ª≠a') + btn('ok', 'Duy·ªát d√πng') : x.trang_thai === 'DANG_DUNG' ? btn('edit', 'S·ª≠a') + btn('off', 'Ng·ª´ng') : x.trang_thai === 'NGUNG_DUNG' ? btn('ok', 'B·∫≠t l·∫°i') : '') + '</div></td></tr>'
    }).join('') + (L.length ? '' : '<tr><td colspan="6" class="empty">Kh√¥ng c√≥ m·∫´u ph√π h·ª£p.</td></tr>') + '</tbody></table></div></div>';
    return h
  }
  function tplModal(x) {
    var d = modal(), m = { nhom: x ? x.nhom : 'KH_MOI', buoc: x ? (x.buoc || '') : '', tieu: x ? x.tieu_de : '', nd: x ? x.noi_dung : '' };
    function bad() { var o = [], re = /\{\{([^}]*)\}\}/g, r; while ((r = re.exec(m.nd))) if (MSG.VARS.indexOf(r[1]) < 0) o.push(r[0]); return o }
    function draw() {
      var b = bad();
      sheet(d, '<h2 style="font-size:18px;margin:0 0 4px">' + (x ? 'S·ª≠a m·∫´u ' + esc(x.ma) : 'Th√™m m·∫´u') + '</h2>' + (x && x.trang_thai === 'DANG_DUNG' ? '<div class="ro">M·∫´u ƒëang d√πng kh√¥ng b·ªã s·ª≠a tr·ª±c ti·∫øp: h·ªá th·ªëng t·∫°o b·∫£n m·ªõi, b·∫£n c≈© chuy·ªÉn "ƒê√£ thay" khi b·∫£n m·ªõi ƒë∆∞·ª£c duy·ªát d√πng.</div>' : '') +
        '<div class="form2" style="margin-top:0"><div><label class="meta">Nh√≥m</label>' + sel('t-nh', Object.keys(NHOM).map(function (k) { return [k, NHOM[k]] }), m.nhom) + '</div><div><label class="meta">B∆∞·ªõc chƒÉm s√≥c</label>' + sel('t-bu', [['', '‚Äî kh√¥ng theo b∆∞·ªõc ‚Äî']].concat(Object.keys(BUOC).map(function (k) { return [k, BUOC[k]] })), m.buoc) + '</div><div style="grid-column:1/-1"><label class="meta">Ti√™u ƒë·ªÅ (nh√¢n vi√™n th·∫•y khi ch·ªçn m·∫´u)</label><input id="t-ti" value="' + esc(m.tieu) + '"></div></div>' +
        '<label class="meta" style="display:block;margin-top:10px">N·ªôi dung</label><textarea id="t-nd" style="min-height:130px">' + esc(m.nd) + '</textarea><div class="row" style="margin:6px 0;flex-wrap:wrap"><span class="meta">Ch√®n:</span>' + MSG.VARS.map(function (v) { return '<button class="chip" data-var="' + v + '" style="min-height:32px;padding:4px 10px;font-size:12.5px" title="' + esc(MSG.VARDESC[v]) + '">' + v + '</button>' }).join('') + '</div>' +
        (b.length ? '<div class="err">Bi·∫øn kh√¥ng c√≥ trong danh s√°ch: ' + esc(b.join(', ')) + '</div>' : '') + '<div class="script" style="margin:8px 0"><small style="margin:0 0 6px">Xem tr∆∞·ªõc (d·ªØ li·ªáu m·∫´u)</small>' + esc(MSG.preview(m.nd)) + '</div><div class="meta">' + m.nd.length + '/2000 k√Ω t·ª±</div>' +
        '<div class="acts" style="margin-top:12px;flex-wrap:wrap"><button class="btn b3" id="t-x" style="flex:1">H·ªßy</button><button class="btn b2" id="t-draft" style="flex:1.2"' + (b.length ? ' disabled' : '') + '>L∆∞u nh√°p (ch·ªù duy·ªát)</button><button class="btn b1" id="t-ok" style="flex:1.4"' + (b.length ? ' disabled' : '') + '>Duy·ªát v√† d√πng ngay</button></div>');
      var sync = function () { m.nhom = q(d, '#t-nh').value; m.buoc = q(d, '#t-bu').value; m.tieu = q(d, '#t-ti').value; m.nd = q(d, '#t-nd').value };
      q(d, '#t-nh').onchange = function () { sync(); draw() }; q(d, '#t-bu').onchange = sync; q(d, '#t-ti').oninput = sync;
      q(d, '#t-nd').oninput = function () { sync(); var pos = this.selectionStart; draw(); var n = q(d, '#t-nd'); n.focus(); n.setSelectionRange(pos, pos) };
      [].slice.call(d.querySelectorAll('[data-var]')).forEach(function (c) { c.onclick = function () { sync(); var ta = q(d, '#t-nd'), p = ta.selectionStart == null ? m.nd.length : ta.selectionStart, ins = '{{' + c.dataset.var + '}}'; m.nd = m.nd.slice(0, p) + ins + m.nd.slice(ta.selectionEnd == null ? p : ta.selectionEnd); draw(); var n = q(d, '#t-nd'); n.focus(); n.setSelectionRange(p + ins.length, p + ins.length) } });
      q(d, '#t-x').onclick = function () { d.remove() };
      var save = function (st) { sync(); if (!m.tieu.trim() || !m.nd.trim()) { toast('C·∫ßn c√≥ ti√™u ƒë·ªÅ v√† n·ªôi dung.', 0, 1); return } rpc('u_asm_tpl_save', { p_ma: x ? x.ma : '', p_nhom: m.nhom, p_buoc: m.buoc, p_tieu_de: m.tieu, p_noi_dung: m.nd, p_trang_thai: st, p_req: rid() }).then(function (r) { d.remove(); done(st === 'DANG_DUNG' ? 'ƒê√£ duy·ªát v√† d√πng m·∫´u ' + r.ma : 'ƒê√£ l∆∞u nh√°p ' + r.ma + ', ch·ªù duy·ªát') }, fail) };
      q(d, '#t-draft').onclick = function () { save('CHO_DUYET') }; q(d, '#t-ok').onclick = function () { save('DANG_DUNG') }
    }
    draw()
  }
  function tplState(ma, st, msg) { rpc('u_asm_tpl_state', { p_ma: ma, p_trang_thai: st, p_req: rid() }).then(function () { return done(msg) }, fail) }
  function promoList() {
    var h = '<div class="card" style="padding:6px 8px"><div class="tw"><table><thead><tr><th>Khuy·∫øn m√£i</th><th>C√¢u d√πng trong tin nh·∫Øn</th><th>Hi·ªáu l·ª±c</th><th>Tr·∫°ng th√°i</th><th></th></tr></thead><tbody>' + D.promos.map(function (p) {
      var t0 = new Date().toISOString().slice(0, 10), live = p.trang_thai === 'DANG_DUNG' && p.tu_ngay <= t0 && (!p.den_ngay || p.den_ngay >= t0);
      return '<tr><td><b>' + esc(p.ten) + '</b><small>' + esc(p.ma) + '</small></td><td style="max-width:360px;white-space:normal">' + esc(p.dong) + '</td><td>' + esc(p.tu_ngay) + ' ‚Üí ' + esc(p.den_ngay || 'kh√¥ng h·∫°n') + '</td><td>' + (live ? tag('t-ok', 'ƒêang ch·∫°y') : p.trang_thai === 'NGUNG_DUNG' ? tag('t-bad', 'Ng·ª´ng') : tag('t-new', 'Ngo√†i hi·ªáu l·ª±c')) + '</td><td><button class="btn b3" data-kmedit="' + esc(p.ma) + '" style="min-height:36px;font-size:13px">S·ª≠a</button></td></tr>'
    }).join('') + (D.promos.length ? '' : '<tr><td colspan="5" class="empty">Ch∆∞a c√≥ khuy·∫øn m√£i.</td></tr>') + '</tbody></table></div></div>';
    return h
  }
  function promoModal(p) {
    var d = modal(), t0 = new Date().toISOString().slice(0, 10);
    sheet(d, '<h2 style="font-size:18px;margin:0 0 10px">' + (p ? 'S·ª≠a khuy·∫øn m√£i ' + esc(p.ma) : 'Th√™m khuy·∫øn m√£i') + '</h2><label class="field">T√™n (nh√¢n vi√™n th·∫•y khi ch·ªçn)<input id="k-ten" value="' + esc(p ? p.ten : '') + '"></label><label class="field">C√¢u d√πng trong tin nh·∫Øn<textarea id="k-dong" style="min-height:80px">' + esc(p ? p.dong : '') + '</textarea></label><div class="form2" style="margin-top:0"><div><label class="meta">T·ª´ ng√†y</label><input id="k-tu" type="date" value="' + esc(p ? p.tu_ngay : t0) + '"></div><div><label class="meta">ƒê·∫øn ng√†y (tr·ªëng = kh√¥ng h·∫°n)</label><input id="k-den" type="date" value="' + esc(p && p.den_ngay ? p.den_ngay : '') + '"></div><div><label class="meta">Tr·∫°ng th√°i</label>' + sel('k-tt', [['DANG_DUNG', 'ƒêang d√πng'], ['NGUNG_DUNG', 'Ng·ª´ng']], p ? p.trang_thai : 'DANG_DUNG') + '</div></div><div class="acts" style="margin-top:12px"><button class="btn b3" id="k-x" style="flex:1">H·ªßy</button><button class="btn b1" id="k-ok" style="flex:1.4">L∆∞u</button></div>', 520);
    q(d, '#k-x').onclick = function () { d.remove() };
    q(d, '#k-ok').onclick = function () { var b = q(d, '#k-ok'); b.disabled = true; rpc('u_asm_promo_save', { p_ma: p ? p.ma : '', p_ten: q(d, '#k-ten').value, p_dong: q(d, '#k-dong').value, p_tu: q(d, '#k-tu').value || null, p_den: q(d, '#k-den').value || null, p_trang_thai: q(d, '#k-tt').value, p_req: rid() }).then(function () { d.remove(); done('ƒê√£ l∆∞u khuy·∫øn m√£i') }, function (e) { b.disabled = false; fail(e) }) }
  }

  /* ================= PH√ÅC ƒê·ªí ================= */
  function pgPd() {
    var t = AS.pd, days = Math.round(Number(t.tv || 0) * Number(t.tu)), due = MSG.dueRules(days, t.tg);
    var h = head('Ph√°c ƒë·ªì vaccine', 'Quy t·∫Øc "khi n√†o": vaccine n√†o ƒë·∫øn tu·ªïi n√†o, t√≠nh t·ª´ ng√†y sinh. App d√πng ƒë·ªÉ g·ª£i √Ω "vaccine ƒë√£ ƒë·∫øn tu·ªïi" trong tin nh·∫Øn c·ªßa nh√¢n vi√™n. D√≤ng ch·ªâ hi·ªán cho nh√¢n vi√™n khi ƒëang d√πng.', '<button class="btn b1" id="pd-new">+ Th√™m d√≤ng ph√°c ƒë·ªì</button>');
    h += '<div class="ro"><b>Ph√°c ƒë·ªì ƒë·∫ßy ƒë·ªß n·∫°p ·ªü Th∆∞ vi·ªán ‚Üí "Ph√°c ƒë·ªì &amp; t·ª´ ƒëi·ªÉn vaccine"</b> (k√©o file Excel: xem tr∆∞·ªõc, ch·ªù duy·ªát, b√°c sƒ© duy·ªát). T·ª´ ƒë√≥ nh√¢n vi√™n c√≥ n√∫t <b>Tra c·ª©u vaccine</b>. B·∫£ng nh·∫≠p tay b√™n d∆∞·ªõi l√† ph√°c ƒë·ªì m·∫´u ƒëang d√πng cho kh·ªëi "vaccine b√© ƒë·∫øn tu·ªïi" trong tin nh·∫Øn, s·∫Ω ƒë∆∞·ª£c thay b·∫±ng ph√°c ƒë·ªì n·∫°p t·ª´ Excel khi anh duy·ªát.</div><div class="ro" style="background:var(--warn-bg);color:var(--warn)"><b>D·ªØ li·ªáu m·∫´u minh h·ªça.</b> C√°c ƒë·ªô tu·ªïi b√™n d∆∞·ªõi ch∆∞a ph·∫£i ph√°c ƒë·ªì chu·∫©n ƒë√£ duy·ªát. B√°c sƒ© c·∫ßn nh·∫≠p v√† duy·ªát ph√°c ƒë·ªì th·∫≠t tr∆∞·ªõc khi d√πng cho kh√°ch th·∫≠t; d√≤ng "Ch·ªù duy·ªát" kh√¥ng hi·ªán cho nh√¢n vi√™n.</div>';
    h += '<div class="card"><h2>Th·ª≠ nhanh: b√© ·ªü tu·ªïi n√†y th√¨ g·ª£i √Ω nh·ªØng vaccine n√†o?</h2><div class="row" style="flex-wrap:wrap"><input id="pd-tv" type="number" min="0" value="' + esc(t.tv) + '" style="width:90px;border:1px solid var(--line);border-radius:12px;padding:10px;min-height:44px;background:var(--card)">' + sel('pd-tu', [['1', 'ng√†y'], ['7', 'tu·∫ßn'], ['30', 'th√°ng'], ['365', 'nƒÉm']], t.tu) + sel('pd-tg', [['', 'C·∫£ nam v√† n·ªØ'], ['M', 'Nam'], ['F', 'N·ªØ']], t.tg) + '<span class="meta">= ' + days + ' ng√†y tu·ªïi</span></div><div style="margin-top:10px">' + (due.length ? due.map(function (r) { return tag('t-today', r.ten) }).join('') : '<span class="meta">Ch∆∞a c√≥ vaccine n√†o ƒëang d√πng ·ªü tu·ªïi n√†y.</span>') + '</div></div>';
    h += '<div class="card" style="padding:6px 8px"><div class="tw"><table><thead><tr><th>Vaccine</th><th>T·ª´ tu·ªïi</th><th>ƒê·∫øn tu·ªïi</th><th>Gi·ªõi t√≠nh</th><th>Ph√°c ƒë·ªì (s·ªë m≈©i, kho·∫£ng c√°ch)</th><th>Ghi ch√∫ cho nh√¢n vi√™n</th><th>Tr·∫°ng th√°i</th><th style="min-width:170px"></th></tr></thead><tbody>' + D.rules.map(function (r) {
      var b = function (k, tx) { return '<button class="btn b3" data-pd' + k + '="' + esc(r.ma) + '" style="min-height:36px;font-size:13px">' + tx + '</button>' };
      return '<tr><td><b>' + esc(r.ten) + '</b><small>' + esc(r.ma_vaccine) + ' ¬∑ ' + esc(r.ma) + (r.tu_khoa ? ' ¬∑ t·ª´ kh√≥a: ' + esc(r.tu_khoa) : '') + '</small></td><td>' + esc(MSG.dayLabel(r.tu_ngay)) + '<small>' + r.tu_ngay + ' ng√†y</small></td><td>' + (r.den_ngay == null ? 'Kh√¥ng gi·ªõi h·∫°n' : esc(MSG.dayLabel(r.den_ngay)) + '<small>' + r.den_ngay + ' ng√†y</small>') + '</td><td>' + (r.gioi_tinh === 'M' ? 'Nam' : r.gioi_tinh === 'F' ? 'N·ªØ' : 'C·∫£ hai') + '</td><td style="max-width:220px;white-space:normal">' + esc(r.phac_do || '‚Äì') + '</td><td style="max-width:220px;white-space:normal">' + esc(r.cau_goi_y || '‚Äì') + '</td><td>' + (r.trang_thai === 'DANG_DUNG' ? tag('t-ok', 'ƒêang d√πng') : r.trang_thai === 'CHO_DUYET' ? tag('t-late', 'Ch·ªù duy·ªát') : tag('t-bad', 'Ng·ª´ng')) + '</td><td><div class="row" style="flex-wrap:wrap">' + b('edit', 'S·ª≠a') + (r.trang_thai === 'CHO_DUYET' ? b('ok', 'Duy·ªát') : r.trang_thai === 'DANG_DUNG' ? b('off', 'Ng·ª´ng') : b('ok', 'B·∫≠t l·∫°i')) + '</div></td></tr>'
    }).join('') + '</tbody></table></div></div>';
    h += linkCard();
    return h
  }
  /* Ki·ªÉm tra li√™n k·∫øt: v·ªõi t·ª´ng vaccine trong ph√°c ƒë·ªì, kho ƒëang c√≥ bao nhi√™u m·∫´u ƒë·ªÉ app T·ª∞ CH·ªåN S·∫¥N khi nh√¢n vi√™n ch·ªçn vaccine ƒë√≥ */
  function linkCard() {
    var lib = MSG._t.LIB; if (!lib.rows) { if (!AS.pd.ask) { AS.pd.ask = true; MSG.ensureLib(render) } return '<div class="card"><h2>Li√™n k·∫øt v·ªõi kho</h2><div class="meta">ƒêang t·∫£i kho m·∫´u‚Ä¶</div></div>' } MSG.ensureLib();
    var rep = MSG.linkReport(), z = function (n) { return n ? '<span class="tag t-ok">' + n + '</span>' : '<span class="tag t-bad">ch∆∞a c√≥</span>' }, bad = rep.filter(function (x) { return !x.tv || !x.nc || !x.km }).length;
    return '<div class="card" style="padding:6px 8px"><h2 style="padding:10px 10px 0">Li√™n k·∫øt v·ªõi kho <span class="pill">app t·ª± ch·ªçn s·∫µn khi nh√¢n vi√™n ch·ªçn vaccine</span></h2><div class="meta" style="padding:0 10px 8px">V·ªõi m·ªói vaccine trong ph√°c ƒë·ªì, app t√¨m trong kho (ƒëang d√πng) c√¢u t∆∞ v·∫•n, nguy c∆° & ƒë∆∞·ªùng l√¢y, khuy·∫øn m√£i c√≤n h·∫°n v√† m·∫´u sau ti√™m c·ªßa ƒë√∫ng vaccine ƒë√≥. ' + (bad ? '<b>' + bad + ' vaccine c√≤n thi·∫øu √≠t nh·∫•t m·ªôt m·ª•c</b>: th√™m t·ª´ kh√≥a ·ªü d√≤ng ph√°c ƒë·ªì, ho·∫∑c b·ªï sung m·∫´u v√†o kho.' : 'T·∫•t c·∫£ vaccine ƒë·ªÅu ƒë√£ c√≥ ƒë·ªß.') + '</div><div class="tw"><table><thead><tr><th>Vaccine</th><th>Nh√≥m trong B·∫£ng gi√°</th><th>T∆∞ v·∫•n</th><th>Nguy c∆°</th><th>Khuy·∫øn m√£i</th><th>Sau ti√™m</th></tr></thead><tbody>' + rep.map(function (x) {
      return '<tr><td><b>' + esc(x.name) + '</b><small>' + esc(x.rule.ma) + (x.rule.trang_thai !== 'DANG_DUNG' ? ' ¬∑ ' + esc(x.rule.trang_thai === 'CHO_DUYET' ? 'ch·ªù duy·ªát' : 'ng·ª´ng') : '') + '</small></td><td>' + (x.groups.length ? esc(x.groups.join('; ').slice(0, 90)) + '<small>' + esc(x.keys.slice(0, 6).join(', ')) + '</small>' : '<span class="tag t-late">Kh√¥ng kh·ªõp nh√≥m n√†o</span>') + '</td><td>' + z(x.tv) + (x.tvT[0] ? '<small>' + esc(x.tvT[0].slice(0, 50)) + '</small>' : '') + '</td><td>' + z(x.nc) + (x.ncT.length ? '<small>' + esc(x.ncT.slice(0, 3).join(', ')) + '</small>' : '') + '</td><td>' + z(x.km) + '</td><td>' + z(x.st) + (x.st ? '' : '<small>s·∫Ω d√πng tin chung</small>') + '</td></tr>'
    }).join('') + '</tbody></table></div></div>'
  }
  function ruleSave(r, patch) { var o = { p_ma: r ? r.ma : '', p_ma_vaccine: r.ma_vaccine, p_ten: r.ten, p_tu_ngay: r.tu_ngay, p_den_ngay: r.den_ngay, p_gioi_tinh: r.gioi_tinh || '', p_phac_do: r.phac_do || '', p_cau: r.cau_goi_y || '', p_trang_thai: r.trang_thai, p_req: rid(), p_tu_khoa: r.tu_khoa == null ? null : r.tu_khoa }; for (var k in patch) o[k] = patch[k]; return rpc('u_asm_rule_save', o) }
  function ruleModal(r) {
    var d = modal(), pick = function (v) { return v == null ? ['', '1'] : [String(v), '1'] };
    sheet(d, '<h2 style="font-size:18px;margin:0 0 10px">' + (r ? 'S·ª≠a ph√°c ƒë·ªì ' + esc(r.ma) : 'Th√™m d√≤ng ph√°c ƒë·ªì') + '</h2><div class="form2" style="margin-top:0"><div><label class="meta">M√£ vaccine (v√≠ d·ª• 6IN1)</label><input id="r-mv" value="' + esc(r ? r.ma_vaccine : '') + '"></div><div><label class="meta">T√™n hi·ªÉn th·ªã</label><input id="r-ten" value="' + esc(r ? r.ten : '') + '"></div>' +
      '<div><label class="meta">T·ª´ tu·ªïi</label><div class="row"><input id="r-tu" type="number" min="0" value="' + esc(r ? r.tu_ngay : 0) + '" style="flex:1">' + sel('r-tuu', [['1', 'ng√†y'], ['7', 'tu·∫ßn'], ['30', 'th√°ng'], ['365', 'nƒÉm']], '1') + '</div></div><div><label class="meta">ƒê·∫øn tu·ªïi (tr·ªëng = kh√¥ng gi·ªõi h·∫°n)</label><div class="row"><input id="r-den" type="number" min="0" value="' + esc(r && r.den_ngay != null ? r.den_ngay : '') + '" style="flex:1">' + sel('r-denu', [['1', 'ng√†y'], ['7', 'tu·∫ßn'], ['30', 'th√°ng'], ['365', 'nƒÉm']], '1') + '</div></div>' +
      '<div><label class="meta">Gi·ªõi t√≠nh</label>' + sel('r-gt', [['', 'C·∫£ nam v√† n·ªØ'], ['M', 'Ch·ªâ nam'], ['F', 'Ch·ªâ n·ªØ']], r ? (r.gioi_tinh || '') : '') + '</div><div><label class="meta">Tr·∫°ng th√°i</label>' + sel('r-tt', [['CHO_DUYET', 'Ch·ªù duy·ªát'], ['DANG_DUNG', 'ƒêang d√πng'], ['NGUNG_DUNG', 'Ng·ª´ng']], r ? r.trang_thai : 'CHO_DUYET') + '</div>' +
      '<div style="grid-column:1/-1"><label class="meta">Ph√°c ƒë·ªì: s·ªë m≈©i, kho·∫£ng c√°ch (b√°c sƒ© nh·∫≠p)</label><textarea id="r-pd" style="min-height:70px">' + esc(r ? r.phac_do || '' : '') + '</textarea></div><div style="grid-column:1/-1"><label class="meta">Ghi ch√∫ cho nh√¢n vi√™n (hi·ªán c·∫°nh vaccine trong g·ª£i √Ω)</label><input id="r-ct" value="' + esc(r ? r.cau_goi_y || '' : '') + '"></div><div style="grid-column:1/-1"><label class="meta">T·ª´ kh√≥a nh·∫≠n di·ªán vaccine (t√πy ch·ªçn, c√°ch nhau b·∫±ng d·∫•u ph·∫©y)</label><input id="r-tk" placeholder="V√≠ d·ª•: hexaxim, infanrix hexa" value="' + esc(r ? r.tu_khoa || '' : '') + '"><div class="meta" style="margin-top:4px">App t·ª± n·ªëi vaccine n√†y v·ªõi t∆∞ v·∫•n, nguy c∆°, khuy·∫øn m√£i, sau ti√™m trong kho theo t√™n vaccine v√† B·∫£ng gi√°. Th√™m t·ª´ kh√≥a (t√™n th∆∞∆°ng m·∫°i, c√°ch g·ªçi kh√°c) n·∫øu b·∫£ng "Li√™n k·∫øt v·ªõi kho" b√™n d∆∞·ªõi b√°o thi·∫øu.</div></div></div><div class="meta" style="margin-top:6px">Tu·ªïi nh·∫≠p b·∫±ng ng√†y, tu·∫ßn (7 ng√†y), th√°ng (30 ng√†y) ho·∫∑c nƒÉm (365 ng√†y); h·ªá th·ªëng l∆∞u theo ng√†y.</div><div class="acts" style="margin-top:12px"><button class="btn b3" id="r-x" style="flex:1">H·ªßy</button><button class="btn b1" id="r-ok" style="flex:1.4">L∆∞u</button></div>', 600);
    q(d, '#r-x').onclick = function () { d.remove() };
    q(d, '#r-ok').onclick = function () {
      var tu = q(d, '#r-tu').value, den = q(d, '#r-den').value, b = q(d, '#r-ok'); if (tu === '') { toast('C·∫ßn nh·∫≠p "T·ª´ tu·ªïi".', 0, 1); return }
      b.disabled = true; ruleSave({ ma: r ? r.ma : '', ma_vaccine: q(d, '#r-mv').value, ten: q(d, '#r-ten').value, tu_ngay: Math.round(Number(tu) * Number(q(d, '#r-tuu').value)), den_ngay: den === '' ? null : Math.round(Number(den) * Number(q(d, '#r-denu').value)), gioi_tinh: q(d, '#r-gt').value, phac_do: q(d, '#r-pd').value, cau_goi_y: q(d, '#r-ct').value, tu_khoa: q(d, '#r-tk').value, trang_thai: q(d, '#r-tt').value }, {}).then(function () { d.remove(); done('ƒê√£ l∆∞u ph√°c ƒë·ªì') }, function (e) { b.disabled = false; fail(e) })
    }
  }

  /* ================= ƒëƒÉng k√Ω trang v√† s·ª± ki·ªán ================= */
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
    all('[data-tltab]', function (b) { b.onclick = function () { AS.tl.tab = b.dataset.tltab; render() } });
    all('[data-tlg]', function (b) { b.onclick = function () { AS.tl.nhom = b.dataset.tlg; render() } });
    on('#tl-st', 'onchange', function (x) { AS.tl.st = x.value; render() });
    on('#tl-q', 'oninput', function (x) { AS.tl.q = x.value; var p = x.selectionStart; render(); var n = $('#tl-q'); n.focus(); n.setSelectionRange(p, p) });
    on('#tl-new', 'onclick', function () { tplModal(null) });
    all('[data-tpedit]', function (b) { b.onclick = function () { tplModal(D.tpl.filter(function (x) { return x.ma === b.dataset.tpedit })[0]) } });
    all('[data-tpok]', function (b) { b.onclick = function () { tplState(b.dataset.tpok, 'DANG_DUNG', 'M·∫´u ' + b.dataset.tpok + ' ƒëang d√πng') } });
    all('[data-tpoff]', function (b) { b.onclick = function () { tplState(b.dataset.tpoff, 'NGUNG_DUNG', 'ƒê√£ ng·ª´ng m·∫´u ' + b.dataset.tpoff) } });
    on('#km-new', 'onclick', function () { promoModal(null) });
    all('[data-kmedit]', function (b) { b.onclick = function () { promoModal(D.promos.filter(function (x) { return x.ma === b.dataset.kmedit })[0]) } });
    on('#pd-new', 'onclick', function () { ruleModal(null) });
    all('[data-pdedit]', function (b) { b.onclick = function () { ruleModal(D.rules.filter(function (x) { return x.ma === b.dataset.pdedit })[0]) } });
    all('[data-pdok]', function (b) { b.onclick = function () { var r = D.rules.filter(function (x) { return x.ma === b.dataset.pdok })[0]; ruleSave(r, { p_trang_thai: 'DANG_DUNG' }).then(function () { return done('Ph√°c ƒë·ªì ' + r.ten + ' ƒëang d√πng') }, fail) } });
    all('[data-pdoff]', function (b) { b.onclick = function () { var r = D.rules.filter(function (x) { return x.ma === b.dataset.pdoff })[0]; ruleSave(r, { p_trang_thai: 'NGUNG_DUNG' }).then(function () { return done('ƒê√£ ng·ª´ng ' + r.ten) }, fail) } });
    on('#pd-tv', 'oninput', function (x) { AS.pd.tv = x.value; var p = x.selectionStart; render(); var n = $('#pd-tv'); n.focus(); n.setSelectionRange(p, p) });
    on('#pd-tu', 'onchange', function (x) { AS.pd.tu = x.value; render() }); on('#pd-tg', 'onchange', function (x) { AS.pd.tg = x.value; render() })
  })
})();

;

/* ===== msg.js ===== */
/* G·ª¢I √ù TIN NH·∫ÆN + L·ªúI N√ìI KHI G·ªåI ‚Äì d√πng chung cho Nh√¢n vi√™n v√† Qu·∫£n l√Ω (Qu·∫£n l√Ω ch·ªâ so·∫°n cho l∆∞·ª£t c·ªßa ch√≠nh m√¨nh).
 * Hai ch·∫ø ƒë·ªô ri√™ng bi·ªát: "Theo quy tr√¨nh" (M·∫∂C ƒê·ªäNH; 4 m·ªëc D, D+1, D+3, D+7 + Tr·ªÖ l·ªãch; m·∫´u t·∫°m t·ª´ msg_templates nh√≥m KH_MOI) v√† "Theo vaccine".
 * THEO VACCINE (H·∫£i 24/11 t·ªëi): app li·ªát k√™ vaccine b√© ƒë√£ ƒë·∫øn tu·ªïi theo PH√ÅC ƒê·ªí; nh√¢n vi√™n ch·ªçn 1 vaccine th√¨ app T·ª∞ T√åM trong kho Excel v√† CH·ªåN S·∫¥N c√¢u t∆∞ v·∫•n,
 *   nguy c∆° & ƒë∆∞·ªùng l√¢y, khuy·∫øn m√£i (CTKM c√≤n h·∫°n) ho·∫∑c CSKH sau ti√™m c·ªßa ƒê√öNG vaccine ƒë√≥ (n·ªëi qua B·∫£ng gi√°: b·ªánh ‚Üî t√™n th∆∞∆°ng m·∫°i; xem h√†m profile()).
 *   Nh√¢n vi√™n ch·ªâ b·∫≠t/t·∫Øt ho·∫∑c ƒë·ªïi, kh√¥ng ph·∫£i ƒëi t√¨m trong c·∫£ danh s√°ch. Ch·∫ø ƒë·ªô "Theo quy tr√¨nh" c≈©ng g·∫Øn ƒë∆∞·ª£c vaccine ƒë·ªÉ t·ª± th√™m nguy c∆°/khuy·∫øn m√£i.
 * M·ªåI tin c√≥ ch·ªØ k√Ω t·ª± ƒë·ªông "Ng∆∞·ªùi chƒÉm s√≥c: <t√™n NV> ¬∑ <Shop> / SƒêT/Zalo: <SƒêT NV>" (SƒêT l·∫•y t·ª´ t√†i kho·∫£n; ch∆∞a c√≥ th√¨ nh·∫Øc nh·∫≠p ngay); t·∫Øt ƒë∆∞·ª£c b·∫±ng c√¥ng t·∫Øc.
 * Ch·ªó tr·ªëng trong m·∫´u ({{NGAY_TIEM}}, {{SO_MUI}}, ‚Ä¶) hi·ªán √¥ ƒë·ªÉ nh√¢n vi√™n ƒëi·ªÅn; c√≤n tr·ªëng m√† g·ª≠i th√¨ app h·ªèi l·∫°i.
 * Nh√≥m kh√°ch ch·ªâ theo TU·ªîI (t·ª´ ng√†y sinh); hi·ªán CH∆ØA c√≥ nh√≥m "b·∫ßu" n√™n m·∫´u thu·ªôc nh√≥m b·∫ßu ƒë∆∞·ª£c ·∫©n (AN_BAU).
 * Gi√°/CTKM trong m·∫´u ƒë√£ ƒë∆∞·ª£c ASM qu√©t v√† thay khi duy·ªát B·∫£ng gi√°/CTKM: nh√¢n vi√™n KH√îNG thay gi√° t·∫°i ch·ªó.
 * N·ªôi dung kho t·∫£i l·∫°i m·ªói khi m·ªü m√†n so·∫°n tin n·∫øu ƒë√£ qu√° 30 gi√¢y (ASM duy·ªát l√¥ m·ªõi th√¨ nh√¢n vi√™n th·∫•y ngay l·∫ßn m·ªü sau).
 * Ba n√∫t g·ª≠i, m·ªói n√∫t l√†m ƒë√∫ng vi·ªác c·ªßa n√≥ r·ªìi t·ª± ghi 1 l·∫ßn "Nh·∫Øn" (h√†m u_msg_log):
 *   Zalo     = ch√©p tin + m·ªü zalo.me/<s·ªë> (Zalo kh√¥ng cho web ƒëi·ªÅn s·∫µn tin hay t·ª± k·∫øt b·∫°n: nh√¢n vi√™n d√°n tin r·ªìi g·ª≠i/k·∫øt b·∫°n)
 *   SMS      = m·ªü app tin nh·∫Øn t·ªõi s·ªë, n·ªôi dung ƒëi·ªÅn s·∫µn (ƒë·ªìng th·ªùi ch√©p tin ph√≤ng khi m√°y t√≠nh kh√¥ng c√≥ app SMS)
 *   Sao ch√©p = ch·ªâ ch√©p tin */
var MSG = (function () {
  var CSS = `
.imgrow,.imggrid{display:flex;gap:8px;overflow-x:auto;padding-bottom:4px}.imggrid{display:grid;grid-template-columns:repeat(3,1fr);overflow:visible}
.ithumb{flex:none;position:relative;width:96px;border:2px solid var(--line);border-radius:14px;background:var(--card);padding:4px;font-size:11px;text-align:left;color:var(--ink);transition:border-color .15s,transform .12s}
.ithumb:active{transform:scale(.97)}.imggrid .ithumb{width:auto}.ithumb img{width:100%;height:64px;object-fit:cover;border-radius:10px;display:block}
.ithumb.on{border-color:var(--teal);background:var(--tint)}.ithumb span{display:block;margin-top:3px;line-height:1.2;max-height:2.4em;overflow:hidden}
.ick{position:absolute;top:6px;right:6px;background:var(--teal);color:#fff;border-radius:50%;width:20px;height:20px;display:grid;place-items:center;font-size:12px}
.person{display:grid;grid-template-columns:1fr auto;gap:10px;align-items:center;padding:12px 0;border-top:1px solid var(--line)}.person:first-child{border-top:0}.person .phone{font-size:20px}
.btn.bo{border:1.5px solid var(--teal);color:var(--teal-d);background:var(--card);min-height:44px}

/* ---- khung m√†n so·∫°n tin: k√≠nh m·ªù, t√¥ng teal/navy c·ªßa Long Ch√¢u ---- */
.modal.mm{align-items:end;background:rgba(6,24,36,.55);-webkit-backdrop-filter:blur(6px);backdrop-filter:blur(6px)}
.modal.mm.first{animation:mfade .2s ease}.modal.mm.first .msheet{animation:mup .3s cubic-bezier(.2,.8,.2,1)}
@keyframes mfade{from{opacity:0}to{opacity:1}}@keyframes mup{from{transform:translateY(28px);opacity:0}to{transform:none;opacity:1}}
@media (prefers-reduced-motion:reduce){.modal.mm.first,.modal.mm.first .msheet{animation:none}.abt,.ithumb,.mcard,.mrow{transition:none}}
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
.abt.o{border:1.5px solid var(--teal);background:var(--card);color:var(--teal-d);min-height:46px;width:100%}.abt.mbig{width:100%;min-height:50px;margin-top:6px}
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
.mp.zero{background:var(--warn-bg);color:var(--warn)}.mcard.vac .mtags{margin-top:0}.mcard.vac .mc1{align-items:center}
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
  var AN_BAU = true;   /* hi·ªán ch∆∞a c√≥ nh√≥m kh√°ch "b·∫ßu" (H·∫£i 24/11): ·∫©n m·∫´u thu·ªôc nh√≥m b·∫ßu kh·ªèi nh√¢n vi√™n. ƒê·∫∑t false ƒë·ªÉ hi·ªán l·∫°i. */
  var VARS = ['XUNG_HO', 'XUNG_HO_HOA', 'TEN_NGUOI_LH', 'TEN_BE', 'TUOI', 'TEN_NV', 'TEN_CO_SO', 'DIA_CHI', 'GIO_MO_CUA', 'HOTLINE', 'SDT_NV', 'TEN_VACCINE', 'NGAY_TIEM', 'SO_MUI', 'NGAY_HEN', 'LINK_GOOGLE_MAPS', 'NOI_DUNG_PHAN_HOI'];
  var VARDESC = { XUNG_HO: 'm·∫π / b·ªë / anh ch·ªã‚Ä¶', XUNG_HO_HOA: 'M·∫π / B·ªë‚Ä¶ (ƒë·∫ßu c√¢u)', TEN_NGUOI_LH: 't√™n ng∆∞·ªùi li√™n h·ªá', TEN_BE: 't√™n b√©', TUOI: 'tu·ªïi b√© (ng√†y/th√°ng/nƒÉm)', TEN_NV: 't√™n nh√¢n vi√™n', TEN_CO_SO: 't√™n Shop', DIA_CHI: 'ƒë·ªãa ch·ªâ Shop', GIO_MO_CUA: 'gi·ªù m·ªü c·ª≠a', HOTLINE: 'hotline Shop', SDT_NV: 'SƒêT nh√¢n vi√™n (t·ª± ƒëi·ªÅn t·ª´ t√†i kho·∫£n)', TEN_VACCINE: 'vaccine nh√¢n vi√™n ƒë√£ ch·ªçn (t·ª± ƒëi·ªÅn)', NGAY_TIEM: 'ng√†y ti√™m (nh√¢n vi√™n ƒëi·ªÅn)', SO_MUI: 'm≈©i th·ª© m·∫•y (nh√¢n vi√™n ƒëi·ªÅn)', NGAY_HEN: 'ng√†y h·∫πn m≈©i ti·∫øp (nh√¢n vi√™n ƒëi·ªÅn)', LINK_GOOGLE_MAPS: 'link b·∫£n ƒë·ªì Shop (nh√¢n vi√™n ƒëi·ªÅn)', NOI_DUNG_PHAN_HOI: 'n·ªôi dung ph·∫£n h·ªìi (nh√¢n vi√™n ƒëi·ªÅn)' };
  var PHL = { NGAY_TIEM: 'Ng√†y ti√™m (v√≠ d·ª• 24/11)', SO_MUI: 'M≈©i th·ª© m·∫•y (v√≠ d·ª• 2)', NGAY_HEN: 'Ng√†y h·∫πn m≈©i ti·∫øp', LINK_GOOGLE_MAPS: 'Link b·∫£n ƒë·ªì Shop', NOI_DUNG_PHAN_HOI: 'N·ªôi dung ph·∫£n h·ªìi', TEN_VACCINE: 'T√™n vaccine', DIA_CHI: 'ƒê·ªãa ch·ªâ Shop', HOTLINE: 'Hotline Shop', GIO_MO_CUA: 'Gi·ªù m·ªü c·ª≠a', SDT_NV: 'SƒêT c·ªßa anh/ch·ªã' };
  var STEPS = [['D', 'D ¬∑ Kh·ªüi t·∫°o'], ['D1', 'D+1 ¬∑ ThƒÉm d√≤'], ['D3', 'D+3 ¬∑ Th√∫c ƒë·∫©y'], ['D7', 'D+7 ¬∑ C·∫•p thi·∫øt'], ['TRE', 'Tr·ªÖ l·ªãch']];
  var STEP_SUB = { D: 'Kh·ªüi t·∫°o', D1: 'ThƒÉm d√≤', D3: 'Th√∫c ƒë·∫©y', D7: 'C·∫•p thi·∫øt', TRE: 'Tr·ªÖ l·ªãch' }, STEP_DOT = { D: 'D', D1: 'D+1', D3: 'D+3', D7: 'D+7', TRE: '' };
  var GROUPS = [['KH_MOI', 'KH m·ªõi'], ['CSKH_SAU_TIEM', 'CSKH sau ti√™m'], ['BAU_NHI', 'B·∫ßu ‚Äì Nhi'], ['CHUNG', 'Chung']];
  var XH = [['', 'T·ª± ƒë·ªông theo ng∆∞·ªùi li√™n h·ªá'], ['m·∫π', 'm·∫π'], ['b·ªë', 'b·ªë'], ['anh', 'anh'], ['ch·ªã', 'ch·ªã'], ['c√¥', 'c√¥'], ['ch√∫', 'ch√∫'], ['b√°c', 'b√°c'], ['√¥ng', '√¥ng'], ['b√†', 'b√†']];
  var TABS = [['CD', 'Theo vaccine', 'MAU_CHU_DE'], ['NHI', 'S∆° sinh & Nhi', 'MAU_TUOI'], ['CH', 'M·∫´u chung', 'MAU_CHUNG'], ['ST', 'Sau ti√™m', 'MAU_SAU_TIEM']];
  var LIBT = ['MAU_CHU_DE', 'MAU_TUOI', 'MAU_SAU_TIEM', 'MAU_CHUNG', 'CTKM', 'NGUY_CO', 'BANG_GIA', 'PHAC_DO'], MAXLEN = 6000, LONG = 900;   /* LONG: t·ª´ ng·∫ßn n√†y k√Ω t·ª± th√¨ c·∫£nh b√°o "tin d√†i"; MAXLEN: gi·ªõi h·∫°n c·ª©ng. L·ªãch s·ª≠ ch·ªâ ghi 2000 k√Ω t·ª± ƒë·∫ßu (u_msg_log). */

  function strip(s) { return String(s || '').replace(/\s*\([^)]*\)/g, '').trim() }
  function given(s) { var p = strip(s).split(/\s+/); return p[p.length - 1] || '' }
  function cap(s) { return s ? s.charAt(0).toUpperCase() + s.slice(1) : s }
  function pad(n) { return n < 10 ? '0' + n : '' + n }
  function todayIso() { var d = new Date(); return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()) }
  function dmy(iso) { var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso || ''); return m ? m[3] + '/' + m[2] + '/' + m[1] : '' }
  function money(n) { return Number(n).toLocaleString('vi-VN') + 'ƒë' }
  function fph(s) { s = String(s || ''); return /^0\d{9}$/.test(s) ? s.slice(0, 4) + ' ' + s.slice(4, 7) + ' ' + s.slice(7) : s }
  function custDays(c) { var b = c && c.ngay_sinh ? new Date(c.ngay_sinh + 'T00:00:00') : null; return b ? Math.max(0, Math.floor((Date.now() - b) / 86400000)) : null }
  function tuoiTxt(c) { var d = custDays(c); if (d === null) return ''; if (d < 60) return d + ' ng√†y tu·ªïi'; var m = Math.floor(d / 30.4375); return m < 24 ? m + ' th√°ng tu·ªïi' : Math.floor(m / 12) + ' tu·ªïi' }
  function dayLabel(d) { if (d == null) return 'kh√¥ng gi·ªõi h·∫°n'; if (d < 60) return d + ' ng√†y'; if (d < 730) return Math.round(d / 30.4375) + ' th√°ng'; var y = d / 365.25; return (Math.round(y * 10) / 10) + ' tu·ªïi' }
  function xungOf(link) { var q = String(link && link.quan_he || '').toLowerCase(); return ['m·∫π', 'b·ªë', '√¥ng', 'b√†', 'anh', 'ch·ªã', 'c√¥', 'ch√∫', 'b√°c'].indexOf(q) >= 0 ? q : 'anh/ch·ªã' }
  function telOf(l) { var s = String(l && l.sdt || '').replace(/[\s.\-()]/g, ''); if (/^\+84\d{9}$/.test(s)) s = '0' + s.slice(3); else if (/^84\d{9}$/.test(s)) s = '0' + s.slice(2); return s }
  function stepNow(t) { var a = ageDays(t.d_giao); return a < 1 ? 'D' : a < 3 ? 'D1' : a < 7 ? 'D3' : 'D7' }
  function fill(text, cx) { return String(text || '').replace(/\{\{([A-Z_]+)\}\}/g, function (m, k) { return cx[k] != null && cx[k] !== '' ? cx[k] : m }) }
  function shopOf() { return (D.shops && D.shops[ME.id_shop]) || {} }
  function fk(x) { return window.LIBSCAN ? LIBSCAN.fold(x || '') : String(x || '').toLowerCase().normalize('NFD').replace(/[ÃÄ-ÕØ]/g, '').replace(/ƒë/g, 'd') }
  /* Bi·∫øn t·ª± ƒëi·ªÅn: t√™n NV, SƒêT NV (t·ª´ t√†i kho·∫£n), Shop; extra = vaccine ƒë√£ ch·ªçn + ch·ªó nh√¢n vi√™n ƒë√£ ƒëi·ªÅn tay. Bi·∫øn thi·∫øu gi√° tr·ªã th√¨ GI·ªÆ NGUY√äN {{‚Ä¶}} ƒë·ªÉ app nh·∫Øc ƒëi·ªÅn. */
  function ctx(c, link, xhOverride, extra) {
    var xh = xhOverride || xungOf(link), sh = shopOf(), o = { XUNG_HO: xh, XUNG_HO_HOA: cap(xh), TEN_NGUOI_LH: given(link && link.ho_ten), TEN_BE: given(c && c.ho_ten) || 'b√©', TUOI: tuoiTxt(c), TEN_NV: given(ME.ho_ten), SDT_NV: fph(ME.sdt), TEN_CO_SO: strip(sh.ten_shop) || 'Trung t√¢m ti√™m ch·ªßng Long Ch√¢u', DIA_CHI: oneLine(sh.dia_chi) || oneLine(sh.ten_shop), GIO_MO_CUA: sh.gio_mo_cua || '', HOTLINE: sh.hotline || '', LINK_GOOGLE_MAPS: '' };
    if (extra) for (var k in extra) if (extra[k] != null && extra[k] !== '') o[k] = extra[k];
    return o
  }
  function preview(text) { return fill(text, { XUNG_HO: 'm·∫π', XUNG_HO_HOA: 'M·∫π', TEN_NGUOI_LH: 'Lan', TEN_BE: 'An', TUOI: '3 th√°ng tu·ªïi', TEN_NV: 'Hi·ªÅn', SDT_NV: '0900 000 222', TEN_VACCINE: 'Ph·∫ø c·∫ßu', TEN_CO_SO: 'Shop m·∫´u', DIA_CHI: 'ƒê·ªãa ch·ªâ Shop m·∫´u', GIO_MO_CUA: '7:30‚Äì21:00', HOTLINE: '0900 000 111' }) }
  function activeTpl() { return (D.tpl || []).filter(function (x) { return x.trang_thai === 'DANG_DUNG' }) }
  function tplById(ma) { return (D.tpl || []).filter(function (x) { return x.ma === ma })[0] }
  function tplForStep(step) { return activeTpl().filter(function (x) { return x.nhom === 'KH_MOI' && x.buoc === step }) }
  function dueRules(days, gender) { if (days == null) return []; return (D.rules || []).filter(function (r) { return r.trang_thai === 'DANG_DUNG' && days >= r.tu_ngay && (r.den_ngay == null || days <= r.den_ngay) && (!r.gioi_tinh || r.gioi_tinh === gender) }) }

  /* ---------- hi·ªán trong chi ti·∫øt kh√°ch ---------- */
  function counters(t) {
    var g = 0, m = 0; D.tt.forEach(function (x) { if (x.ma_luot !== t.ma_luot) return; var c = D.dm[x.loai_tuong_tac]; if (!c) return; if (c.ma === 'GOI') g++; else if (isMsg(x)) m++ });
    return '<span class="tag t-new">G·ªçi ' + g + '/3</span><span class="tag t-new">Nh·∫Øn ' + m + '/2</span>'
  }
  function scriptBlock(t, link) {
    var c = D.cus[t.ma_khach] || {}, st = stepNow(t), g = activeTpl().filter(function (x) { return x.nhom === 'GOI' }), tp = g.filter(function (x) { return x.buoc === st })[0] || g[0], txt, src;
    if (tp) { txt = fill(tp.noi_dung, ctx(c, link)); src = 'Th∆∞ vi·ªán ¬∑ ' + tp.ma + ' ¬∑ ' + esc(tp.tieu_de) }
    else { txt = 'Ch√†o ' + xungOf(link) + ' ' + given(link && link.ho_ten) + ', em l√† ' + given(ME.ho_ten) + ' ·ªü trung t√¢m ti√™m ch·ªßng Long Ch√¢u. Em g·ªçi v·ªÅ b√© ' + given(c.ho_ten) + ' (' + tuoiTxt(c) + '). Anh/ch·ªã ti·ªán n√≥i chuy·ªán m·ªôt ch√∫t ƒë∆∞·ª£c kh√¥ng ·∫°?'; src = 'M·∫´u t·∫°m (Th∆∞ vi·ªán ch∆∞a c√≥ l·ªùi n√≥i khi g·ªçi)' }
    return '<div class="sec"><h3>G·ª£i √Ω n√≥i khi g·ªçi</h3><div class="script">' + esc(txt) + '<small>' + src + '</small></div></div>'
  }

  /* ---------- KHO M·∫™U EXCEL (lib_rows ƒëang d√πng): t·∫£i l·∫°i n·∫øu qu√° 30 gi√¢y ---------- */
  var LIB = { rows: null, at: 0, loading: false, err: null, cbs: [] };
  function loadLib(force) {
    if (LIB.loading || (!force && LIB.rows && Date.now() - LIB.at < 30000)) return; LIB.loading = true;
    rest('lib_rows?select=id,loai,stt,data&trang_thai=eq.DANG_DUNG&loai=in.(' + LIBT.join(',') + ')&order=loai.asc,stt.asc&limit=2000').then(function (a) {
      var by = {}; a.forEach(function (r) { (by[r.loai] = by[r.loai] || []).push({ id: r.id, stt: r.stt, d: r.data || {} }) }); LIB.rows = by; LIB.at = Date.now(); LIB.loading = false; LIB.err = null; PC = {}; afterLib()
    }, function (e) { LIB.loading = false; LIB.err = e || {}; afterLib() })
  }
  function afterLib() { if (S && host) draw(true); if (lkHost) lkDraw(); var c = LIB.cbs.splice(0); c.forEach(function (f) { try { f() } catch (e) { } }) }
  function ensureLib(cb) { if (cb) LIB.cbs.push(cb); loadLib(); if (LIB.rows && !LIB.loading) { var c = LIB.cbs.splice(0); c.forEach(function (f) { try { f() } catch (e) { } }) } }
  function isBau(loai, d) {
    if (!AN_BAU) return false; var f = loai === 'MAU_TUOI' ? (d.nhom_tuoi || '') + ' ' + (d.buoc || '') : loai === 'MAU_CHUNG' ? (d.chu_de || '') + ' ' + (d.ap_dung || '') + ' ' + (d.ten_ngan || '') : loai === 'MAU_CHU_DE' ? (d.chu_de || '') + ' ' + (d.nhom_ap_dung || '') : loai === 'MAU_SAU_TIEM' ? (d.ten_mau || '') : '';
    return /(^|[^a-z])(bau|mang thai|thai ky)([^a-z]|$)/.test(fk(f))
  }
  function libRows(loai) { return ((LIB.rows && LIB.rows[loai]) || []).filter(function (x) { return !isBau(loai, x.d) }) }
  function splitRole(s) { var m = /^\s*\[([^\]]+)\]\s*(.*)$/.exec(s || ''); return m ? { role: m[1], title: m[2] || m[1] } : { role: '', title: String(s || '') } }
  function pretty(s) { s = String(s || '').trim(); return s === s.toUpperCase() && /[A-Z√Ä-·ª∏]/.test(s) && !/\d/.test(s) ? cap(s.toLowerCase()) : s }
  function nl(s) { return String(s || '').replace(/\r\n?/g, '\n').trim() }
  function oneLine(s) { return String(s || '').replace(/\s+/g, ' ').trim() }
  /* h·ª£p tu·ªïi b√©? 'ALL' (kh√¥ng gi·ªõi h·∫°n) | 'FIT' | 'NO' ‚Äî ch·ªâ theo TU·ªîI t·ª´ ng√†y sinh */
  function audFit(text, days) {
    var t = fk(text), y = 365.25, m = 30.4375, lo = null, hi = null, a;
    if (days == null || !t || /tat ca/.test(t)) return 'ALL';
    if ((a = /(\d+)\s*thang\s*[-‚Äì]\s*(\d+)\s*tuoi/.exec(t))) { lo = a[1] * m; hi = a[2] * y }
    else if ((a = /(\d+)\s*[-‚Äì]\s*(\d+)\s*tuoi/.exec(t))) { lo = a[1] * y; hi = a[2] * y }
    else if ((a = />\s*(\d+)\s*tuoi/.exec(t))) { lo = a[1] * y }
    if (lo !== null) return days >= lo && (hi === null || days <= hi + y) ? 'FIT' : 'NO';
    var g = days < 9 * y ? 'NHI' : days < 18 * y ? 'TRE_LON' : 'NGUOI_LON', set = {};
    if (/so sinh|(^|[^a-z])nhi([^a-z]|$)|6in1|(^|[^a-z])lao([^a-z]|$)/.test(t)) set.NHI = 1;
    if (/tre em/.test(t)) { set.NHI = 1; set.TRE_LON = 1 } if (/thanh thieu nien|tre lon/.test(t)) set.TRE_LON = 1; if (/nguoi truong thanh|nguoi lon|cao tuoi/.test(t)) set.NGUOI_LON = 1;
    return !Object.keys(set).length ? 'ALL' : set[g] ? 'FIT' : 'NO'
  }

  /* ---------- N·ªêI VACCINE V·ªöI KHO (H·∫£i 24/11 t·ªëi) ----------
   * Nh√¢n vi√™n ch·ªçn 1 vaccine (g·ª£i √Ω theo tu·ªïi b√© t·ª´ PH√ÅC ƒê·ªí) ‚Üí app t·ª± t√¨m trong kho: c√¢u t∆∞ v·∫•n, nguy c∆° & ƒë∆∞·ªùng l√¢y, CTKM c√≤n h·∫°n, CSKH sau ti√™m c·ªßa ƒê√öNG vaccine ƒë√≥ v√† ch·ªçn s·∫µn.
   * C√°ch n·ªëi: t√™n vaccine trong ph√°c ƒë·ªì (+ "t·ª´ kh√≥a" ASM nh·∫≠p) ‚Üí nh√≥m b·ªánh trong B·∫¢NG GI√Å ‚Üí c√°c t√™n th∆∞∆°ng m·∫°i (Hexaxim, Prevenar‚Ä¶) ‚Üí kh·ªõp v·ªõi kho theo
   *   t∆∞ v·∫•n: c·ªôt ch·ªß ƒë·ªÅ/nh√≥m tu·ªïi ¬∑ nguy c∆°: t√™n b·ªánh (ho·∫∑c vaccine √°p d·ª•ng) ¬∑ CTKM: t√™n s·∫£n ph·∫©m ¬∑ sau ti√™m: t√™n vaccine ghi trong ti√™u ƒë·ªÅ m·∫´u. */
  function nv(s) { var t = fk(s).replace(/\bvac\s*xin\b|\bvaccine\b|\bvacxin\b|\bvac\b/g, ' ').replace(/\b(\d)\s*(?:trong|in|\.)\s*1\b/g, '$1in1').replace(/[^a-z0-9]+/g, ' ').trim(); return ' ' + t + ' ' }
  function has(h, p) { return !!p && h.indexOf(' ' + p + ' ') >= 0 }
  function anyHas(h, L) { for (var i = 0; i < L.length; i++) if (has(h, L[i])) return true; return false }
  function phrasesOf(txt) {
    var out = []; String(txt || '').split(/[\/,;+&‚Äì‚Äî]|\s-\s|\sv√†\s/i).forEach(function (p) { var n = nv(p).trim(); if (!n) return; var t = n.split(' '); while (t.length > 1 && t[t.length - 1].length <= 1) t.pop(); n = t.join(' '); if (n.length >= 3 && out.indexOf(n) < 0) out.push(n) });
    return out
  }
  var GEN = ['uon', 'van', 'bach', 'hau', 'ga', 'soi', 'quai', 'thuy', 'dau', 'viem', 'gan', 'phe', 'cau', 'nao', 'mo', 'dai', 'ta', 'don'], STOP2 = ['virus', 'benh', 'cac', 'ngua', 'nhom', 'viem', 'phoi', 'khac', 'vac', 'xin', 'cua', 'cho', 'nguoi'];
  function brandKey(name) { var t = nv(name).trim().split(' '), a = t[0] || ''; if (a.length >= 3 && GEN.indexOf(a) < 0 && !/^\d/.test(a)) return a; return t.slice(0, 2).join(' ') }
  var BG = { at: -1, g: [] }, VC = {}, PC = {}, VM = {};
  function shortName(n) { return oneLine(String(n || '').replace(/\s*\([^)]*\)/g, '').replace(/^(ng·ª´a|ngua)\s+/i, '').replace(/^c√°c b·ªánh do\s+/i, '')) || oneLine(n) }
  function bgIndex() {
    if (BG.at === LIB.at && LIB.rows) return BG;
    var by = {}, ord = []; libRows('BANG_GIA').forEach(function (r) {
      var d = r.d, nm = oneLine(d.phong_benh) || 'Kh√°c', k = nv(nm); if (!by[k]) { by[k] = { id: 'bg:' + k.trim(), name: nm, short: shortName(nm), norm: k, brands: [], bn: [], keys: [] }; ord.push(k) }
      var g = by[k]; g.brands.push(d.ten_vaccine); g.bn.push(nv(d.ten_vaccine)); var bk = brandKey(d.ten_vaccine); if (bk && g.keys.indexOf(bk) < 0) g.keys.push(bk)
    });
    BG = { at: LIB.at, g: ord.map(function (k) { return by[k] }) }; return BG
  }
  function mkVRule(r) {
    var ck = 'r:' + r.ma + '|' + LIB.at + '|' + r.cap_nhat_luc + '|' + (r.tu_khoa || ''); if (VC[ck]) return VC[ck];
    var phr = phrasesOf(r.ten + ';' + (r.tu_khoa || '')), keys = [], groups = [];
    bgIndex().g.forEach(function (g) { if (phr.some(function (p) { return has(g.norm, p) || g.keys.indexOf(p) >= 0 || g.bn.some(function (b) { return has(b, p) }) })) { groups.push(g); g.keys.forEach(function (k) { if (keys.indexOf(k) < 0) keys.push(k) }) } });
    return (VC[ck] = { id: 'rule:' + r.ma, name: r.ten, rule: r, phr: phr, keys: keys, groups: groups, from: 'RULE' })
  }
  function mkVbg(g) {
    var ck = 'b:' + g.id + '|' + LIB.at; if (VC[ck]) return VC[ck]; var base = nv(g.short).trim(), phr = [base];
    base.split(' ').forEach(function (w) { if (w.length >= 4 && STOP2.indexOf(w) < 0 && phr.indexOf(w) < 0) phr.push(w) });
    return (VC[ck] = { id: g.id, name: g.short, group: g, phr: phr, keys: g.keys.slice(), groups: [g], from: 'BG' })
  }
  function reg(V) { VM[V.id] = V; return V }
  function dueVacs() { var c = cust(), d = custDays(c); return dueRules(d, c.gioi_tinh).map(mkVRule) }
  function soonVacs() { var c = cust(), d = custDays(c); if (d == null) return []; return (D.rules || []).filter(function (r) { return r.trang_thai === 'DANG_DUNG' && d < r.tu_ngay && r.tu_ngay - d <= 45 && (!r.gioi_tinh || r.gioi_tinh === c.gioi_tinh) }).map(mkVRule) }
  function ageOf(V) { var r = V.rule; return r ? 't·ª´ ' + dayLabel(r.tu_ngay) + (r.den_ngay != null ? ' ƒë·∫øn ' + dayLabel(r.den_ngay) : ' tr·ªü l√™n') : 'theo B·∫£ng gi√°' }

  var TABS_BY = {};
  function mkItem(tab, loai, r) {
    var d = r.d, o = { tab: tab, loai: loai, id: r.id, stt: r.stt, text: nl(d.noi_dung), ord: Number(d.thu_tu) || r.stt, key: tab + ':' + r.id }, s;
    if (loai === 'MAU_CHU_DE') { s = splitRole(d.ten_mau); o.ma = 'CD-' + r.stt; o.role = s.role; o.ten = s.title; o.group = d.chu_de || 'Kh√°c'; o.aud = d.nhom_ap_dung || ''; o.audFit = o.aud; o.topic = d.chu_de || '' }
    else if (loai === 'MAU_TUOI') { s = splitRole(d.ten_mau); o.ma = 'NHI-' + r.stt; o.role = s.role; o.ten = s.title; o.group = pretty(d.nhom_tuoi) || 'Kh√°c'; o.aud = d.buoc || ''; o.audFit = (d.nhom_tuoi || '') + ' ' + (d.buoc || ''); o.topic = d.nhom_tuoi || '' }
    else if (loai === 'MAU_CHUNG') { o.ma = d.ma_tin || 'TN-' + r.stt; o.ten = d.ten_ngan || d.ma_tin || 'M·∫´u chung'; o.role = ''; o.group = d.chu_de || 'Chung'; o.aud = d.ap_dung || ''; o.audFit = o.aud; o.topic = (d.ten_ngan || '') + ' ' + (d.chu_de || '') }
    else { var p = String(d.ten_mau || '').split(/\s+-\s+/), lt = /^([A-Z])\.\s*/.exec(p.slice(2).join(' - ')); o.ma = p[0] || 'ST-' + r.stt; o.role = ''; o.vac = (p[1] && !/^tat ca/.test(fk(p[1]))) ? p[1] : ''; o.aud = ''; o.audFit = ''; o.letter = lt ? lt[1] : ''; o.group = (p.slice(2).join(' - ') || 'Sau ti√™m').replace(/^[A-Z]\.\s*/, ''); o.ten = o.vac ? shortName(o.vac) + ' ¬∑ ' + o.ma : o.group + ' ¬∑ ' + o.ma; o.topic = '' }
    o.fit = audFit(o.audFit, custDays(cust())); return o
  }
  function itemsOf(tab) { var cfg = TABS.filter(function (t) { return t[0] === tab })[0]; return cfg ? libRows(cfg[2]).map(function (r) { return mkItem(tab, cfg[2], r) }) : [] }
  function tabName(tab) { var cfg = TABS.filter(function (t) { return t[0] === tab })[0]; return cfg ? cfg[1] : '' }
  function findItem(key) { var i = key.indexOf(':'), tab = key.slice(0, i), id = Number(key.slice(i + 1)); return itemsOf(tab).filter(function (o) { return o.id === id })[0] }
  function khoCount() { return TABS.reduce(function (n, t) { return n + libRows(t[2]).length }, 0) }
  function pickQT(tp) { return tp ? { src: 'QT', ma: tp.ma, ten: tp.tieu_de, text: tp.noi_dung, tag: 'M·∫´u quy tr√¨nh', role: '', fit: '' } : null }
  function pickK(o) { return { src: 'KHO', ma: o.ma, ten: o.ten, text: o.text, tag: tabName(o.tab), role: o.role, fit: o.fit, aud: o.aud, key: o.key } }

  /* ---------- CTKM hi·ªán h√†nh (kho CTKM, c√≤n h·∫°n theo ng√†y) v√† Nguy c∆° ---------- */
  function ctkmGroups() {
    var t = todayIso(), by = {}, order = [];
    libRows('CTKM').forEach(function (r) {
      var d = r.d; if ((d.tu_ngay && d.tu_ngay > t) || (d.den_ngay && d.den_ngay < t)) return; var k = d.nhom || ('R' + r.id);
      if (!by[k]) { by[k] = { nhom: k, loai: d.phan_loai || '', chung: d.phan_loai === '∆Øu ƒë√£i chung', title: d.noi_dung_ct || d.ten_chuong_trinh || k, items: [], den: d.den_ngay || '', note: '', vac: d.loai_vaccine || '' }; order.push(k) }
      by[k].items.push(d); if (d.den_ngay && d.den_ngay > by[k].den) by[k].den = d.den_ngay
    });
    return order.map(function (k) {
      var g = by[k]; g.note = ''; g.inline = false;
      if (!g.chung) { var ns = []; g.items.forEach(function (d) { var n = oneLine(d.ghi_chu); if (n && ns.indexOf(n) < 0) ns.push(n) }); if (ns.length === 1 && g.items.every(function (d) { return oneLine(d.ghi_chu) === ns[0] })) g.note = ns[0]; else g.inline = ns.length > 0 }
      return g
    }).sort(function (a, b) { return (a.chung - b.chung) || String(a.nhom).localeCompare(String(b.nhom), 'vi', { numeric: true }) })
  }
  function itemLine(d, chung, inline) {
    if (chung) return '- T·∫∑ng ' + oneLine(d.ten_sp);
    var pay = d.gia_thanh_toan != null ? d.gia_thanh_toan : d.gia_sau_km != null ? d.gia_sau_km : null, orig = d.gia_niem_yet != null ? d.gia_niem_yet : d.gia_goc, n = inline ? oneLine(d.ghi_chu) : '';
    return '- ' + oneLine(d.ten_sp) + (d.sl_mui > 1 ? ' (' + d.sl_mui + ' m≈©i)' : '') + (pay != null ? ': ' + (pay === 0 ? 'MI·ªÑN PH√ç' : money(pay)) + (orig && orig > pay ? ' (gi√° g·ªëc ' + money(orig) + ')' : '') : '') + (n ? ' [' + n + ']' : '')
  }
  function ctkmLines(g) { return g.items.map(function (d) { return itemLine(d, g.chung, g.inline) }).join('\n') + (g.den ? '\n(√Åp d·ª•ng ƒë·∫øn ' + dmy(g.den) + (g.note ? '. ' + g.note : '') + ')' : '') }
  function ctkmText(g) { return (g.chung ? 'Qu√† t·∫∑ng theo h√≥a ƒë∆°n - ' : '∆Øu ƒë√£i: ') + g.title + '\n' + ctkmLines(g) }
  function ncList() { return libRows('NGUY_CO') }
  function ncText(d) { var e = function (s) { return String(s || '').trim().replace(/[.;]+$/, '') }, s = e(d.benh) + ': ' + e(d.nguy_co) + '.'; if (d.lay_qua) s += ' L√¢y qua: ' + e(d.lay_qua) + '.'; if (d.diem_can_nho) s += ' C·∫ßn nh·ªõ: ' + e(d.diem_can_nho) + '.'; return s }

  /* H·ªì s∆° n·ªëi c·ªßa 1 vaccine: nh·ªØng g√¨ kho c√≥ cho ƒë√∫ng vaccine ƒë√≥ */
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
  function linkReport() { return (D.rules || []).map(function (r) { var V = mkVRule(r), P = profile(V); return { rule: r, name: r.ten, groups: V.groups.map(function (g) { return g.short }), keys: V.keys, tv: P.tv.length, nc: P.nc.length, km: P.km.length, st: P.st.length, tvT: P.tv.slice(0, 3).map(function (o) { return o.ten }), ncT: P.nc.map(function (x) { return x.d.benh }), stT: P.st.map(function (o) { return o.ten }) } }) }

  /* ---------- m√†n so·∫°n tin ---------- */
  var S = null, host = null;
  function guard() { if (S.text !== S.base && !confirm('Anh/ch·ªã ƒë√£ s·ª≠a tay n·ªôi dung. ƒê·ªïi l·ª±a ch·ªçn s·∫Ω t·∫°o l·∫°i tin theo m·∫´u v√† m·∫•t ph·∫ßn ƒë√£ s·ª≠a. Ti·∫øp t·ª•c?')) return false; return true }
  function onKey(e) { if (e.key === 'Escape' && host) { if (S && (S.view === 'ctkm' || S.view === 'nc')) leaveCheck(); else if (S && S.view !== 'main') { S.view = 'main'; draw() } else close() } }
  function open(t, ci) {
    var links = D.links[t.ma_khach] || []; if (!links.length) { toast('Kh√°ch n√†y ch∆∞a c√≥ ng∆∞·ªùi li√™n h·ªá.', 0, 1); return }
    var step = stepNow(t);
    S = { t: t, ci: Math.min(ci || 0, links.length - 1), mode: 'QT', step: step, pq: pickQT(tplForStep(step)[0]), vac: null, purpose: 'TV', pk: null, st: null, km: {}, nc: {}, sig: true, dc: false, xh: '', fv: {}, view: 'main', pt: 'CD', pg: '', pq2: '', lim: 40, pm: true, pctx: 'tv', cf: '', cm: true, nq: '', vq: '', img: {}, iq: '', it: '', text: '', base: '', sending: false, snap: null, dirty: false };
    host = document.createElement('div'); host.className = 'modal mm first'; host.onclick = function (e) { if (e.target === host) close() }; document.body.appendChild(host);
    document.addEventListener('keydown', onKey); setTimeout(function () { if (host) host.classList.remove('first') }, 450);
    loadFiles(); loadLib(); regen()
  }
  function close() { document.removeEventListener('keydown', onKey); if (host) host.remove(); host = null; S = null }
  function link() { return (D.links[S.t.ma_khach] || [])[S.ci] || {} }
  function cust() { return D.cus[S && S.t ? S.t.ma_khach : ''] || {} }
  function cxS() { var e = {}; for (var k in S.fv) e[k] = S.fv[k]; if (S.vac) e.TEN_VACCINE = S.vac.name; return ctx(cust(), link(), S.xh, e) }
  function sigText(cx) { return 'Ng∆∞·ªùi chƒÉm s√≥c: ' + strip(ME.ho_ten) + ' ¬∑ ' + cx.TEN_CO_SO + (cx.SDT_NV ? '\nSƒêT/Zalo: ' + cx.SDT_NV : '') }
  function sigNeeded(txt) { var d = String(ME.sdt || ''); return !(d && txt.replace(/\D/g, '').indexOf(d) >= 0) }
  function compose() {
    var cx = cxS(), parts = [], f = function (t) { return fill(nl(t), cx) };
    if (S.mode === 'QT' && S.pq) parts.push(f(S.pq.text));
    if (S.pk) parts.push(f(S.pk.text));
    if (S.st) parts.push(f(S.st.text));
    var nBase = parts.length;
    ncList().forEach(function (r) { if (S.nc[r.id]) parts.push(ncText(r.d)) });
    /* m·∫´u c√≥ s·∫µn d√≤ng ti√™u ƒë·ªÅ c·ªßa ch∆∞∆°ng tr√¨nh (v√≠ d·ª• "Combo b·∫£o v·ªá h√¥ h·∫•p: Gi·∫£m 5%") th√¨ ch·ªâ n·ªëi c√°c d√≤ng gi√° ngay d∆∞·ªõi ti√™u ƒë·ªÅ ƒë√≥, kh√¥ng l·∫∑p ti√™u ƒë·ªÅ */
    ctkmGroups().forEach(function (g) {
      if (!S.km[g.nhom]) return; var tn = fk(g.title), host = -1; if (tn.length >= 8) for (var i = 0; i < nBase; i++) if (fk(parts[i]).indexOf(tn) >= 0) { host = i; break }
      if (host >= 0) parts[host] = parts[host] + '\n' + ctkmLines(g); else parts.push(ctkmText(g))
    });
    if (S.dc) parts.push('ƒê·ªãa ch·ªâ: ' + cx.DIA_CHI + (cx.GIO_MO_CUA ? '. Gi·ªù m·ªü c·ª≠a: ' + cx.GIO_MO_CUA : '') + (cx.HOTLINE ? '. Hotline: ' + cx.HOTLINE : '') + '.');
    var txt = parts.join('\n\n'); if (S.sig && sigNeeded(txt)) txt = (txt ? txt + '\n\n' : '') + sigText(cx); return txt
  }
  function regen() { S.base = compose(); S.text = S.base; draw() }
  /* ch·ªçn t·ª± ƒë·ªông theo vaccine + m·ª•c ƒë√≠ch: t∆∞ v·∫•n, nguy c∆°, CTKM (ho·∫∑c sau ti√™m) c·ªßa ƒê√öNG vaccine ƒë√≥ */
  function applyKit() {
    S.km = {}; S.nc = {}; S.pk = null; S.st = null; var V = S.vac; if (!V) return; var P = profile(V);
    if (S.purpose === 'ST') { var s = P.st[0] || P.stGen[0]; if (s) S.st = pickK(s); return }
    if (S.mode === 'KHO' && P.tv[0]) S.pk = pickK(P.tv[0]);
    P.nc.forEach(function (r) { S.nc[r.id] = 1 }); P.km.forEach(function (g) { S.km[g.nhom] = 1 })
  }
  function setVaccine(V) { if (!guard()) return; S.vac = V; S.purpose = 'TV'; S.view = 'main'; applyKit(); regen() }
  function setMode(m) {
    if (m === S.mode) return; if (!guard()) return; S.mode = m;
    if (m === 'KHO') { loadLib(); if (S.vac && S.purpose === 'TV' && !S.pk) { var P = profile(S.vac); if (P.tv[0]) S.pk = pickK(P.tv[0]) } } else { S.pk = null; if (!S.pq) S.pq = pickQT(tplForStep(S.step)[0]) }
    regen()
  }
  function lenInfo(n) { return { txt: n + ' k√Ω t·ª±' + (n > LONG ? ' ¬∑ SMS ‚âà ' + Math.ceil(n / 67) + ' tin' : ''), cls: n > MAXLEN ? 'over' : n > LONG ? 'longw' : '' } }
  function msgCount() { var m = 0; D.tt.forEach(function (x) { if (x.ma_luot === S.t.ma_luot && isMsg(x)) m++ }); return m }
  function openPH() { var m = String(S.text || '').match(/\{\{[A-Z_]+\}\}/g) || [], o = []; m.forEach(function (x) { var n = x.slice(2, -2); if (o.indexOf(n) < 0) o.push(n) }); return o }
  function draw(keep) {
    if (!host || !S) return; var b = host.querySelector('.mbody'), top = keep && b ? b.scrollTop : 0, v = S.view;
    if (v === 'pick') drawPick(); else if (v === 'ctkm') drawCtkm(); else if (v === 'nc') drawNc(); else if (v === 'vac') drawVac(); else if (v === 'img') drawImgs(); else drawMain();
    if (top) { var nb = host.querySelector('.mbody'); if (nb) nb.scrollTop = top }
  }
  function hdrBack(title, sub) { return '<div class="mhd"><div class="mtop"><button class="mclose" id="pb" aria-label="Quay l·∫°i">' + ic('back') + '</button><div class="mtt"><div class="t1">' + esc(title) + '</div>' + (sub ? '<div class="t2">' + esc(sub) + '</div>' : '') + '</div><button class="mclose" id="mx" aria-label="ƒê√≥ng">' + ic('x') + '</button></div></div>' }
  function skeleton() { var h = '<div class="mlist">'; for (var i = 0; i < 3; i++) h += '<div class="mrow sk"><div><i></i><i></i></div></div>'; return h + '</div>' }
  function tcard(p, btn) {
    var cx = cxS();
    return '<div class="tcard"><div class="tctop"><div class="tt">' + esc(p.ten) + '</div>' + (btn ? '<button class="mlink" id="chg">' + btn + '</button>' : '') + '</div><div class="mtags">' + (p.tag ? '<span class="mp">' + esc(p.tag) + '</span>' : '') + (p.role ? '<span class="mp ro">' + esc(p.role) + '</span>' : '') + (p.fit === 'FIT' ? '<span class="mp gr">H·ª£p tu·ªïi b√©</span>' : '') + '</div><div class="pv">' + esc(fill(nl(p.text), cx).replace(/\s+/g, ' ').slice(0, 240)) + '</div></div>'
  }
  function qtBlock() {
    var now = stepNow(S.t), L = tplForStep(S.step), sel = S.pq && S.pq.ma, alts = L.filter(function (x) { return x.ma !== sel }), cx = cxS();
    var h = '<section><div class="mh">M·ªëc chƒÉm s√≥c</div><div class="stp">' + STEPS.map(function (s) { var k = s[0]; return '<button class="' + (S.step === k ? 'on' : '') + '" data-s="' + k + '" aria-pressed="' + (S.step === k) + '"><i>' + (k === 'TRE' ? ic('clock') : STEP_DOT[k]) + (k === now ? '<u class="sg" title="G·ª£i √Ω"></u>' : '') + '</i><span>' + STEP_SUB[k] + '</span></button>' }).join('') + '</div><div class="mnote"><u class="sg2"></u><span>G·ª£i √Ω theo s·ªë ng√†y k·ªÉ t·ª´ khi giao. Ch·ªçn tay "Tr·ªÖ l·ªãch" khi kh√°ch ƒë√£ l·ª° m≈©i.</span></div></section>';
    h += '<section><div class="mh">M·∫´u quy tr√¨nh</div>' + (S.pq ? tcard(S.pq, null) : '<div class="mwarn"><b>Ch∆∞a c√≥ m·∫´u cho m·ªëc n√†y.</b> Chuy·ªÉn sang "Theo vaccine" ho·∫∑c t·ª± vi·∫øt b√™n d∆∞·ªõi.</div>');
    if (alts.length) h += '<div class="mh mh2">M·∫´u kh√°c c√πng m·ªëc</div><div class="mlist">' + alts.slice(0, 3).map(function (a) { return '<button class="mrow" data-t="' + esc(a.ma) + '"><div><b>' + esc(a.tieu_de) + '</b><span>' + esc(fill(a.noi_dung, cx).replace(/\s+/g, ' ').slice(0, 100)) + '</span></div>' + ic('chev') + '</button>' }).join('') + '</div>';
    return h + '</section>' + vacChips()
  }
  /* QT: d·∫£i ch·ªçn vaccine g·ªçn (t√πy ch·ªçn) */
  function vacChips() {
    if (LIB.loading && !LIB.rows) return ''; var due = dueVacs(); if (S.vac && !due.some(function (v) { return v.id === S.vac.id })) due = due.concat([S.vac]);
    return '<section><div class="mh">G·∫Øn vaccine ¬∑ t·ª± ƒëi·ªÅn nguy c∆°, khuy·∫øn m√£i (t√πy ch·ªçn)</div><div class="mchips vch">' + due.map(function (v) { reg(v); return '<button class="' + (S.vac && S.vac.id === v.id ? 'on' : '') + '" data-vac="' + esc(v.id) + '">' + esc(v.name) + '</button>' }).join('') + '<button id="vmore">Vaccine kh√°c‚Ä¶</button></div>' + (S.vac ? '<div class="mnote"><span>ƒêang g·∫Øn ' + esc(S.vac.name) + '.</span><button class="mlink" id="vlook">Xem h·ªì s∆° vaccine</button></div>' : '<div class="mnote"><span>Ch·ªçn vaccine b√© ƒë√£ ƒë·∫øn tu·ªïi theo ph√°c ƒë·ªì: app t·ª± th√™m nguy c∆° v√† khuy·∫øn m√£i ƒë√∫ng vaccine ƒë√≥.</span></div>') + '</section>'
  }
  function vacCard(V, dim) {
    reg(V); var P = profile(V), sel = S.vac && S.vac.id === V.id, z = function (n, t) { return '<span class="mp ' + (n ? 'gr' : 'zero') + '">' + t + ' ' + (n || 'ch∆∞a c√≥') + '</span>' };
    return '<button class="mcard vac ' + (sel ? 'on' : '') + '" data-vac="' + esc(V.id) + '"' + (dim ? ' style="opacity:.8"' : '') + '><div class="mc1"><b>' + esc(V.name) + '</b><span class="mck">' + ic('check') + '</span></div><div class="mtags"><span class="mp">' + esc((dim ? 'S·∫Øp ƒë·∫øn tu·ªïi ¬∑ ' : '') + (V.rule ? cap(ageOf(V)) : 'Theo B·∫£ng gi√°')) + '</span></div><div class="mtags">' + z(P.tv.length, 'T∆∞ v·∫•n') + z(P.nc.length, 'Nguy c∆°') + z(P.km.length, 'Khuy·∫øn m√£i') + '</div></button>'
  }
  function vacSummary() {
    var V = S.vac, r = V.rule, mau = r && fk(r.phac_do).indexOf('mau') === 0;
    return '<section><div class="mh">Vaccine ƒëang t∆∞ v·∫•n</div><div class="tcard"><div class="tctop"><div class="tt">' + esc(V.name) + '</div><span><button class="mlink" id="vlook">H·ªì s∆°</button><button class="mlink" id="vchg">ƒê·ªïi vaccine</button></span></div><div class="mtags"><span class="mp">' + esc(cap(ageOf(V))) + '</span>' + (mau ? '<span class="mp or">Ph√°c ƒë·ªì m·∫´u t·∫°m</span>' : '') + '</div><div class="psw"><button data-pur="TV" class="' + (S.purpose === 'TV' ? 'on' : '') + '">T∆∞ v·∫•n ti√™m</button><button data-pur="ST" class="' + (S.purpose === 'ST' ? 'on' : '') + '">CSKH sau ti√™m</button></div></div></section>'
  }
  function khoBlock() {
    if (LIB.loading && !LIB.rows) return '<section><div class="mh">Vaccine b√© ƒë·∫øn tu·ªïi ti√™m</div>' + skeleton() + '</section>';
    if (LIB.err && !LIB.rows) return '<section><div class="mwarn">Kh√¥ng t·∫£i ƒë∆∞·ª£c kho m·∫´u (ki·ªÉm tra m·∫°ng). <button class="mlink" id="lretry">Th·ª≠ l·∫°i</button></div></section>';
    var n = khoCount(); if (!n) return '<section><div class="mwarn"><b>Kho ch∆∞a c√≥ m·∫´u n√†o ƒëang d√πng.</b> ASM duy·ªát l√¥ m·∫´u ·ªü m·ª•c Th∆∞ vi·ªán th√¨ m·∫´u s·∫Ω hi·ªán ·ªü ƒë√¢y.</div></section>';
    if (S.vac) return vacSummary();
    var due = dueVacs(), soon = soonVacs(), h = '<section><div class="mh">Vaccine b√© ƒë·∫øn tu·ªïi ti√™m ¬∑ theo ph√°c ƒë·ªì</div>';
    h += due.length ? '<div class="mlist">' + due.map(function (v) { return vacCard(v) }).join('') + '</div>' : '<div class="mwarn"><b>Ch∆∞a c√≥ vaccine n√†o t·ªõi tu·ªïi theo ph√°c ƒë·ªì.</b> ' + (custDays(cust()) == null ? 'B√© ch∆∞a c√≥ ng√†y sinh. ' : '') + 'Ch·ªçn "Vaccine kh√°c" b√™n d∆∞·ªõi.</div>';
    if (soon.length) h += '<div class="mh mh2">S·∫Øp ƒë·∫øn tu·ªïi (trong 45 ng√†y)</div><div class="mlist">' + soon.map(function (v) { return vacCard(v, true) }).join('') + '</div>';
    h += '<div class="mrow2"><button class="abt o" id="vmore">' + ic('layers') + 'Vaccine kh√°c</button><button class="abt o" id="bro">' + ic('search') + 'Duy·ªát kho (' + n + ')</button></div></section>';
    if (S.pk) h += '<section><div class="mh">M·∫´u ƒëang ch·ªçn</div>' + tcard(S.pk, 'ƒê·ªïi m·∫´u') + '</section>';
    return h
  }
  /* c√°c d√≤ng "t·ª± ƒëi·ªÅn theo vaccine": m·ªói d√≤ng c√≥ c√¥ng t·∫Øc, ch·∫°m v√†o ch·ªØ ƒë·ªÉ ƒë·ªïi l·ª±a ch·ªçn */
  function kitRow(ico, id, title, sub, on, canTg, warn) {
    return '<div class="li"><span class="ico">' + ic(ico) + '</span><button class="lt lbtn" data-kit="' + id + '"><b>' + esc(title) + '</b><small' + (warn ? ' style="color:var(--warn)"' : '') + '>' + sub + '</small></button><span class="tg"><input type="checkbox" data-tg="' + id + '" aria-label="' + esc(title) + '" ' + (on ? 'checked ' : '') + (canTg ? '' : 'disabled') + '><i></i></span></div>'
  }
  function kitList() {
    var V = S.vac, P = V ? profile(V) : null, cx = cxS(), nm = V ? esc(V.name) : '', rows = '';
    var nSel = function (m) { return Object.keys(m).filter(function (k) { return m[k] }).length };
    var tvSub = S.pk ? esc(S.pk.ten) : V ? (P.tv.length ? P.tv.length + ' m·∫´u cho ' + nm + ' ¬∑ ch·∫°m ƒë·ªÉ ch·ªçn' : 'Kho ch∆∞a c√≥ m·∫´u t∆∞ v·∫•n ri√™ng cho ' + nm) : 'Ch·ªçn vaccine ·ªü tr√™n ƒë·ªÉ t·ª± ƒëi·ªÅn, ho·∫∑c ch·∫°m ƒë·ªÉ ch·ªçn th·ªß c√¥ng';
    rows += kitRow('spark', 'tv', 'C√¢u t∆∞ v·∫•n', tvSub, !!S.pk, !!(V && P.tv.length) || !!S.pk, V && !P.tv.length && !S.pk);
    var nc = nSel(S.nc), ncSub = nc ? ncList().filter(function (r) { return S.nc[r.id] }).slice(0, 3).map(function (r) { return esc(r.d.benh) }).join(', ') + (nc > 3 ? '‚Ä¶' : '') : V ? (P.nc.length ? P.nc.length + ' m·ª•c cho ' + nm : 'Kho Nguy c∆° ch∆∞a c√≥ ' + nm) : 'Ch·ªçn vaccine ·ªü tr√™n ƒë·ªÉ t·ª± ƒëi·ªÅn, ho·∫∑c ch·∫°m ƒë·ªÉ ch·ªçn th·ªß c√¥ng';
    rows += kitRow('shield', 'nc', 'Nguy c∆° & ƒë∆∞·ªùng l√¢y', ncSub, nc > 0, !!(V && P.nc.length) || nc > 0, V && !P.nc.length && !nc);
    var km = nSel(S.km), kmSub = km ? km + ' nh√≥m ∆∞u ƒë√£i ƒë√£ ch·ªçn' + (V && P.km.length ? ' ¬∑ ' + nm : '') : V ? (P.km.length ? P.km.length + ' nh√≥m ∆∞u ƒë√£i c√≤n hi·ªáu l·ª±c cho ' + nm : 'Ch∆∞a c√≥ khuy·∫øn m√£i c√≤n hi·ªáu l·ª±c cho ' + nm) : ctkmGroups().length + ' nh√≥m ∆∞u ƒë√£i c√≤n hi·ªáu l·ª±c ¬∑ ch·∫°m ƒë·ªÉ ch·ªçn';
    rows += kitRow('gift', 'km', 'Khuy·∫øn m√£i', kmSub, km > 0, !!(V && P.km.length) || km > 0, V && !P.km.length && !km);
    if (S.mode === 'KHO') rows += kitRow('chat', 'st', 'CSKH sau ti√™m', S.st ? esc(S.st.ten) : V ? (P.st.length ? P.st.length + ' m·∫´u ri√™ng cho ' + nm : 'Ch∆∞a c√≥ m·∫´u ri√™ng ¬∑ d√πng tin chung sau ti√™m') : 'Ch·ªçn vaccine ·ªü tr√™n ƒë·ªÉ t·ª± ƒëi·ªÅn', !!S.st, !!V || !!S.st);
    var noPhone = !ME.sdt;
    rows += kitRow('user', 'sig', 'Ch·ªØ k√Ω ng∆∞·ªùi chƒÉm s√≥c', S.sig ? (noPhone ? 'Ch∆∞a c√≥ SƒêT c·ªßa anh/ch·ªã: tin s·∫Ω thi·∫øu s·ªë li√™n h·ªá' : esc(strip(ME.ho_ten)) + ' ¬∑ ' + esc(cx.TEN_CO_SO) + ' ¬∑ ' + esc(cx.SDT_NV)) : 'ƒêang t·∫Øt: tin kh√¥ng c√≥ t√™n nh√¢n vi√™n, Shop, SƒêT', S.sig, true, S.sig && noPhone);
    if (noPhone) rows += '<div class="sdtbox"><input id="mysdt" type="tel" inputmode="tel" placeholder="Nh·∫≠p SƒêT c·ªßa anh/ch·ªã, v√≠ d·ª• 0912345678" autocomplete="tel" aria-label="SƒêT c·ªßa anh/ch·ªã"><button class="abt p" id="sdtsave" style="min-height:44px;flex:none;padding:0 16px"><b>L∆∞u SƒêT</b></button></div>';
    rows += '<label class="li"><span class="ico">' + ic('pin') + '</span><span class="lt"><b>ƒê·ªãa ch·ªâ Shop</b><small>' + esc(cx.DIA_CHI) + '</small></span><span class="tg"><input type="checkbox" id="swd" ' + (S.dc ? 'checked' : '') + '><i></i></span></label>';
    rows += '<div class="li"><span class="ico">' + ic('user') + '</span><span class="lt"><b>X∆∞ng h√¥</b></span><select id="xh" aria-label="X∆∞ng h√¥">' + XH.map(function (o) { return '<option value="' + esc(o[0]) + '"' + (o[0] === S.xh ? ' selected' : '') + '>' + esc(o[1]) + '</option>' }).join('') + '</select></div>';
    return '<section><div class="mh">' + (V ? 'T·ª± ƒëi·ªÅn theo ' + nm : 'Th√™m v√†o tin') + '</div><div class="ls">' + rows + '</div></section>'
  }
  function phCard() {
    var L = openPH(); if (!L.length) return '';
    return '<section><div class="mh">C√≤n ch·ªó tr·ªëng c·∫ßn ƒëi·ªÅn (' + L.length + ')</div><div class="ls">' + L.map(function (n) { return '<div class="li"><span class="lt"><b>' + esc(PHL[n] || n) + '</b><input class="fvi" data-fv="' + esc(n) + '" value="' + esc(S.fv[n] || '') + '" placeholder="Nh·∫≠p r·ªìi ch·∫°m ra ngo√†i ho·∫∑c b·∫•m Enter" aria-label="' + esc(PHL[n] || n) + '"></span></div>' }).join('') + '</div><div class="mnote"><span>Ch∆∞a ƒëi·ªÅn th√¨ ch·ªØ {{‚Ä¶}} v·∫´n n·∫±m trong tin g·ª≠i kh√°ch. App s·∫Ω h·ªèi l·∫°i tr∆∞·ªõc khi g·ª≠i.</span></div></section>'
  }
  function drawMain() {
    var c = cust(), l = link(), links = D.links[S.t.ma_khach] || [], tel = telOf(l), edited = S.text !== S.base, len = S.text.length, over = len > MAXLEN, nm = msgCount(), nImg = Object.keys(S.img).length;
    var ini = (given(l.ho_ten) || l.quan_he || '?').charAt(0).toUpperCase(), due = dueVacs();
    var h = '<div class="msheet" role="dialog" aria-label="So·∫°n tin"><div class="mhd"><div class="mtop"><div class="mav">' + esc(ini) + '</div><div class="mtt"><div class="t1">So·∫°n tin cho ' + esc(cap(l.quan_he || 'ng∆∞·ªùi li√™n h·ªá')) + ' ' + esc(given(l.ho_ten)) + '</div><div class="t2">B√© ' + esc(given(c.ho_ten)) + (tuoiTxt(c) ? ' ¬∑ ' + esc(tuoiTxt(c)) : '') + (tel ? ' ¬∑ ' + esc(tel) : '') + '</div></div><button class="mclose" id="mx" aria-label="ƒê√≥ng">' + ic('x') + '</button></div>' +
      '<div class="mseg" role="tablist"><button role="tab" aria-selected="' + (S.mode === 'QT') + '" class="' + (S.mode === 'QT' ? 'on' : '') + '" data-mode="QT">' + ic('route') + 'Theo quy tr√¨nh</button><button role="tab" aria-selected="' + (S.mode === 'KHO') + '" class="' + (S.mode === 'KHO' ? 'on' : '') + '" data-mode="KHO">' + ic('layers') + 'Theo vaccine</button></div></div><div class="mbody">';
    if (links.length > 1) h += '<section><div class="mh">G·ª≠i cho</div><div class="sg3">' + links.map(function (x, i) { return '<button data-ci="' + i + '" class="' + (i === S.ci ? 'on' : '') + '">' + esc(cap(x.quan_he || 'Li√™n h·ªá')) + (x.ho_ten ? ' ' + esc(given(x.ho_ten)) : '') + '</button>' }).join('') + '</div></section>';
    h += S.mode === 'QT' ? qtBlock() : khoBlock();
    h += '<section><div class="mh">N·ªôi dung tin</div><div class="ed"><div class="eh"><b>' + (nm ? 'Tin th·ª© ' + (nm + 1) : 'Tin ƒë·∫ßu ti√™n') + '</b><span class="sp"></span><span class="badge ' + (edited ? 'b-ed' : 'b-ok') + '" id="bd">' + (edited ? 'ƒê√£ s·ª≠a tay' : 'ƒê√∫ng theo m·∫´u') + '</span></div><textarea id="tx" aria-label="N·ªôi dung tin" placeholder="Ch·ªçn vaccine ho·∫∑c m·∫´u, ho·∫∑c t·ª± vi·∫øt n·ªôi dung tin‚Ä¶">' + esc(S.text) + '</textarea><div class="ef"><button class="mlink" id="rs">' + ic('reset').replace('class="i"', 'class="i" style="width:16px;height:16px;vertical-align:-3px"') + ' Kh√¥i ph·ª•c theo m·∫´u</button><span class="sp"></span><span id="cc" class="' + lenInfo(len).cls + '">' + lenInfo(len).txt + '</span></div></div></section>';
    h += phCard() + kitList() + imgBlock();
    var can = !!S.text.trim() && !over;
    h += '</div><div class="mft">' + (nImg ? '<button class="abt o" id="sendimg">' + ic('image') + 'G·ª≠i k√®m ' + nImg + ' ·∫£nh ¬∑ chia s·∫ª sang Zalo</button>' : '') + '<div class="mst"><span class="cnt">Nh·∫Øn ' + nm + '/2</span><span>' + (over ? 'Tin qu√° d√†i (tr√™n ' + MAXLEN + ' k√Ω t·ª±), h√£y r√∫t ng·∫Øn ƒë·ªÉ g·ª≠i.' : len > LONG ? 'Tin d√†i: Zalo g·ª≠i ƒë∆∞·ª£c, SMS s·∫Ω t√°ch nhi·ªÅu tin. C√≥ th·ªÉ t·∫Øt b·ªõt ph·∫ßn ·ªü m·ª•c "T·ª± ƒëi·ªÅn" b√™n d∆∞·ªõi.' : 'Ch·∫°m n√∫t g·ª≠i: app l√†m ƒë√∫ng vi·ªác c·ªßa n√∫t ƒë√≥ v√† t·ª± ghi 1 l·∫ßn "Nh·∫Øn".') + '</span></div><div class="act3">' +
      '<button class="abt p" data-go="ZALO" ' + (can ? '' : 'disabled') + '>' + ic('chat') + '<span><b>Zalo</b><small>ch√©p tin + m·ªü Zalo</small></span></button>' +
      '<button class="abt s" data-go="SMS" ' + (can ? '' : 'disabled') + '>' + ic('sms') + '<span><b>SMS</b><small>' + (tel ? 't·ªõi ' + esc(tel) : 'ch∆∞a c√≥ s·ªë') + '</small></span></button>' +
      '<button class="abt g" data-go="SAO_CHEP" ' + (can ? '' : 'disabled') + '>' + ic('copy') + '<span><b>Sao ch√©p</b><small>ch·ªâ ch√©p tin</small></span></button></div></div></div>';
    host.innerHTML = h; bindMain()
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
    var tx = q('#tx'); tx.oninput = function () { S.text = tx.value; var ed = S.text !== S.base, bd = q('#bd'), over = S.text.length > MAXLEN; bd.className = 'badge ' + (ed ? 'b-ed' : 'b-ok'); bd.textContent = ed ? 'ƒê√£ s·ª≠a tay' : 'ƒê√∫ng theo m·∫´u'; var cc = q('#cc'); var li = lenInfo(S.text.length); cc.textContent = li.txt; cc.className = li.cls; var can = !!S.text.trim() && !over; all('[data-go]').forEach(function (g) { g.disabled = !can }) };
    q('#rs').onclick = function () { S.text = S.base; draw(true) };
    /* c√°c d√≤ng t·ª± ƒëi·ªÅn theo vaccine */
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
    var i = host.querySelector('#mysdt'), b = host.querySelector('#sdtsave'); if (!i || !i.value.trim()) { toast('Nh·∫≠p SƒêT c·ªßa anh/ch·ªã tr∆∞·ªõc khi l∆∞u.', 0, 1); return } b.disabled = true;
    rpc('u_me_save_phone', { p_sdt: i.value, p_req: rid() }).then(function (r) { ME.sdt = r.sdt || null; if (D.acc && D.acc[ME.ma_tk]) D.acc[ME.ma_tk].sdt = r.sdt || null; toast('ƒê√£ l∆∞u SƒêT. S·ªë n√†y n·∫±m trong ch·ªØ k√Ω m·ªçi tin nh·∫Øn.'); if (S && host) { var keepEdited = S.text !== S.base; if (keepEdited) { draw(true) } else regen() } }, function (e) { if (b) b.disabled = false; fail(e) })
  }
  function openPick(ctxName) { S.pctx = ctxName; if (ctxName === 'st') S.pt = 'ST'; else if (S.pt === 'ST') S.pt = 'CD'; S.pq2 = ''; S.pg = ''; S.lim = 40; if (S.vac && ctxName !== 'x' && S.pm !== false) S.pm = true; S.view = 'pick'; draw() }

  /* ---------- ch·ªçn m·∫´u t·ª´ kho: tab + t√¨m + l·ªçc nh√≥m (+ ch·ªâ vaccine ƒëang ch·ªçn) ---------- */
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
    var h = '<div class="msheet">' + hdrBack(ctxT === 'st' ? 'M·∫´u CSKH sau ti√™m' : 'Kho m·∫´u tin', (onlyM ? 'Ch·ªâ m·∫´u c·ªßa ' + S.vac.name + ' ¬∑ ' : '') + base.length + ' m·∫´u') + '<div class="mbody">' +
      (tabs.length > 1 ? '<div class="mtabs" role="tablist">' + tabs.map(function (t) { var n = libRows(t[2]).length; return '<button role="tab" class="' + (S.pt === t[0] ? 'on' : '') + '" data-pt="' + t[0] + '">' + esc(t[1]) + '<em>' + n + '</em></button>' }).join('') + '</div>' : '') +
      (P ? '<div class="mchips"><button class="' + (onlyM ? 'on' : '') + '" data-pm="1">Ch·ªâ ' + esc(S.vac.name) + '</button><button class="' + (!onlyM ? 'on' : '') + '" data-pm="0">C·∫£ kho</button></div>' : '') +
      '<div class="msearch">' + ic('search') + '<input id="pqi" placeholder="T√¨m trong m·∫´u‚Ä¶" value="' + esc(S.pq2) + '" aria-label="T√¨m m·∫´u"></div>' +
      (order.length > 1 ? '<div class="mchips"><button class="' + (!S.pg ? 'on' : '') + '" data-pg="">T·∫•t c·∫£ ¬∑ ' + base.length + '</button>' + order.map(function (g) { return '<button class="' + (S.pg === g ? 'on' : '') + '" data-pg="' + esc(g) + '">' + esc(g) + ' ¬∑ ' + groups[g] + '</button>' }).join('') + '</div>' : '') +
      (LIB.loading && !LIB.rows ? skeleton() : '<div class="mlist">' + (shown.length ? shown.map(function (o) {
        return '<button class="mcard ' + (selKey === o.ma ? 'on' : '') + '" data-k="' + esc(o.key) + '"><div class="mc1"><b>' + esc(o.ten) + '</b></div><div class="mtags">' + (o.hit ? '<span class="mp gr">ƒê√∫ng vaccine</span>' : '') + (o.fit === 'FIT' ? '<span class="mp gr">H·ª£p tu·ªïi b√©</span>' : o.fit === 'NO' ? '<span class="mp ro">Kh√°c nh√≥m tu·ªïi</span>' : '') + (o.role ? '<span class="mp ro">' + esc(o.role) + '</span>' : '') + (o.aud && !/tat ca/.test(fk(o.aud)) ? '<span class="mp">' + esc(o.aud.slice(0, 40)) + '</span>' : '') + '</div><div class="pv2">' + esc(o.text.replace(/\s+/g, ' ').slice(0, 200)) + '</div></button>'
      }).join('') + (more ? '<button class="abt o" id="lmore">Hi·ªán th√™m (' + (list.length - S.lim) + ')</button>' : '') : '<div class="mwarn">' + (onlyM ? 'Kho ch∆∞a c√≥ m·∫´u cho ' + esc(S.vac.name) + '. B·∫•m "C·∫£ kho" ƒë·ªÉ ch·ªçn m·∫´u kh√°c.' : 'Kh√¥ng c√≥ m·∫´u ph√π h·ª£p. Th·ª≠ b·ªè b·ªõt b·ªô l·ªçc.') + '</div>') + '</div>') + '</div></div>';
    host.innerHTML = h;
    host.querySelector('#mx').onclick = close; host.querySelector('#pb').onclick = function () { S.view = 'main'; draw() };
    [].slice.call(host.querySelectorAll('[data-pt]')).forEach(function (b) { b.onclick = function () { S.pt = b.dataset.pt; S.pg = ''; S.lim = 40; draw() } });
    [].slice.call(host.querySelectorAll('[data-pm]')).forEach(function (b) { b.onclick = function () { S.pm = b.dataset.pm === '1'; S.pg = ''; S.lim = 40; draw() } });
    [].slice.call(host.querySelectorAll('[data-pg]')).forEach(function (b) { b.onclick = function () { S.pg = b.dataset.pg; S.lim = 40; draw() } });
    var pi = host.querySelector('#pqi'); pi.oninput = function () { S.pq2 = pi.value; S.lim = 40; var pos = pi.selectionStart; draw(); var n = host.querySelector('#pqi'); n.focus(); n.setSelectionRange(pos, pos) };
    var lm = host.querySelector('#lmore'); if (lm) lm.onclick = function () { S.lim += 40; draw(true) };
    [].slice.call(host.querySelectorAll('.mcard[data-k]')).forEach(function (b) { b.onclick = function () { var o = findItem(b.dataset.k); if (!o) return; if (!guard()) return; if (ctxT === 'st') { S.st = pickK(o); S.purpose = 'ST' } else { S.pk = pickK(o) } S.view = 'main'; regen() } })
  }
  /* m√†n ch·ªçn vaccine kh√°c (ngo√†i danh s√°ch g·ª£i √Ω theo tu·ªïi): ph√°c ƒë·ªì ƒë·∫ßy ƒë·ªß + B·∫£ng gi√° */
  function drawVac() {
    var q = fk(S.vq.trim()), rules = (D.rules || []).filter(function (r) { return r.trang_thai === 'DANG_DUNG' && (!q || fk(r.ten + ' ' + (r.tu_khoa || '')).indexOf(q) >= 0) }).map(mkVRule), bg = bgIndex().g.filter(function (g) { return !q || fk(g.name + ' ' + g.brands.join(' ')).indexOf(q) >= 0 }).map(mkVbg);
    var h = '<div class="msheet">' + hdrBack('Ch·ªçn vaccine', 'App t·ª± n·ªëi nguy c∆°, khuy·∫øn m√£i, t∆∞ v·∫•n, sau ti√™m') + '<div class="mbody"><div class="msearch">' + ic('search') + '<input id="vqi" placeholder="T√¨m vaccine ho·∫∑c t√™n th∆∞∆°ng m·∫°i‚Ä¶" value="' + esc(S.vq) + '" aria-label="T√¨m vaccine"></div>' +
      '<section><div class="mh">Theo ph√°c ƒë·ªì (' + rules.length + ')</div>' + (rules.length ? '<div class="mlist">' + rules.map(function (v) { return vacCard(v) }).join('') + '</div>' : '<div class="mnote"><span>Kh√¥ng c√≥ d√≤ng ph√°c ƒë·ªì ph√π h·ª£p.</span></div>') + '</section>' +
      '<section><div class="mh">Theo B·∫£ng gi√° (' + bg.length + ')</div>' + (LIB.loading && !LIB.rows ? skeleton() : bg.length ? '<div class="mlist">' + bg.map(function (v) { return vacCard(v) }).join('') + '</div>' : '<div class="mnote"><span>Kh√¥ng c√≥ nh√≥m vaccine ph√π h·ª£p.</span></div>') + '</section></div></div>';
    host.innerHTML = h; host.querySelector('#mx').onclick = close; host.querySelector('#pb').onclick = function () { S.view = 'main'; draw() };
    [].slice.call(host.querySelectorAll('[data-vac]')).forEach(function (b) { b.onclick = function () { var V = VM[b.dataset.vac]; if (V) setVaccine(V) } });
    var vi = host.querySelector('#vqi'); vi.oninput = function () { S.vq = vi.value; var pos = vi.selectionStart; draw(); var n = host.querySelector('#vqi'); n.focus(); n.setSelectionRange(pos, pos) }
  }

  /* ---------- ch·ªçn nhi·ªÅu: CTKM hi·ªán h√†nh, Nguy c∆° ---------- */
  function checkFoot(n) { return '<div class="mft"><button class="abt p" id="done" style="min-height:52px;justify-content:center">' + ic('check') + '<span><b>Xong' + (n ? ' ¬∑ ƒë√£ ch·ªçn ' + n : '') + '</b></span></button></div>' }
  function bindCheck(sel, key) {
    host.querySelector('#mx').onclick = close; host.querySelector('#pb').onclick = leaveCheck; host.querySelector('#done').onclick = leaveCheck;
    [].slice.call(host.querySelectorAll(sel)).forEach(function (b) { b.onclick = function () { var k = b.dataset.c; if (S[key][k]) delete S[key][k]; else S[key][k] = 1; S.dirty = true; draw(true) } })
  }
  /* m·ªü m√†n ch·ªçn nhi·ªÅu: ch·ª•p l·∫°i l·ª±a ch·ªçn c≈© ƒë·ªÉ ho√†n l·∫°i n·∫øu nh√¢n vi√™n kh√¥ng mu·ªën m·∫•t ph·∫ßn ƒë√£ s·ª≠a tay */
  function enterCheck(view) { S.snap = { km: JSON.parse(JSON.stringify(S.km)), nc: JSON.parse(JSON.stringify(S.nc)) }; S.dirty = false; S.view = view; draw() }
  function leaveCheck() {
    var dirty = S.dirty; S.dirty = false;
    if (!dirty) { S.view = 'main'; draw(); return }
    if (S.text !== S.base && !confirm('Th√™m ho·∫∑c b·ªè n·ªôi dung s·∫Ω t·∫°o l·∫°i tin v√† m·∫•t ph·∫ßn ƒë√£ s·ª≠a tay. Ti·∫øp t·ª•c?')) { S.km = S.snap.km; S.nc = S.snap.nc; S.view = 'main'; draw(); return }
    S.view = 'main'; regen()
  }
  function drawCtkm() {
    var G = ctkmGroups(), P = S.vac ? profile(S.vac) : null, mset = {}; if (P) P.km.forEach(function (g) { mset[g.nhom] = 1 });
    var onlyM = !!(P && P.km.length && S.cm), cats = {}, order = []; (onlyM ? G.filter(function (g) { return mset[g.nhom] }) : G).forEach(function (g) { var k = g.chung ? 'Qu√† t·∫∑ng' : g.loai || 'Kh√°c'; if (!cats[k]) { cats[k] = 0; order.push(k) } cats[k]++ });
    var list = (onlyM ? G.filter(function (g) { return mset[g.nhom] }) : G).filter(function (g) { return !S.cf || (g.chung ? 'Qu√† t·∫∑ng' : g.loai || 'Kh√°c') === S.cf }), n = Object.keys(S.km).filter(function (k) { return S.km[k] }).length;
    list.forEach(function (g) { g.hit = !!mset[g.nhom] }); list.sort(function (a, b) { return (b.hit ? 1 : 0) - (a.hit ? 1 : 0) });
    var h = '<div class="msheet">' + hdrBack('Khuy·∫øn m√£i hi·ªán h√†nh', G.length ? G.length + ' nh√≥m ∆∞u ƒë√£i c√≤n hi·ªáu l·ª±c h√¥m nay' : 'Ch∆∞a c√≥ ch∆∞∆°ng tr√¨nh c√≤n hi·ªáu l·ª±c') + '<div class="mbody">' +
      (P ? '<div class="mchips"><button class="' + (onlyM ? 'on' : '') + '" data-cm="1">Ch·ªâ ' + esc(S.vac.name) + ' ¬∑ ' + P.km.length + '</button><button class="' + (!onlyM ? 'on' : '') + '" data-cm="0">T·∫•t c·∫£ ¬∑ ' + G.length + '</button></div>' : '') +
      (order.length > 1 ? '<div class="mchips"><button class="' + (!S.cf ? 'on' : '') + '" data-cf="">M·ªçi lo·∫°i</button>' + order.map(function (k) { return '<button class="' + (S.cf === k ? 'on' : '') + '" data-cf="' + esc(k) + '">' + esc(k) + ' ¬∑ ' + cats[k] + '</button>' }).join('') + '</div>' : '') +
      '<div class="mnote"><span>Gi√° v√† ∆∞u ƒë√£i ƒë√£ kh·ªõp B·∫£ng gi√°/CTKM m·ªõi nh·∫•t do ASM duy·ªát. Ch·ªçn nh√≥m n√†o th√¨ tin s·∫Ω k√®m nh√≥m ƒë√≥.</span></div><div class="mlist">' +
      (list.length ? list.map(function (g) { return '<button class="mcard ' + (S.km[g.nhom] ? 'on' : '') + '" data-c="' + esc(g.nhom) + '"><div class="mc1"><b>' + esc(g.title) + '</b><span class="mck">' + ic('check') + '</span></div><div class="mtags"><span class="mp">' + esc(g.chung ? 'Qu√† t·∫∑ng' : g.loai || '∆Øu ƒë√£i') + '</span>' + (g.hit ? '<span class="mp gr">ƒê√∫ng vaccine</span>' : '') + '<span class="mp ro">' + g.items.length + ' m·ª•c' + (g.den ? ' ¬∑ ƒë·∫øn ' + esc(dmy(g.den)) : '') + '</span></div><div class="pv2">' + esc(g.items.slice(0, 3).map(function (d) { return oneLine(d.ten_sp) }).join(' ¬∑ ')) + '</div></button>' }).join('') : '<div class="mwarn">' + (onlyM ? 'Ch∆∞a c√≥ khuy·∫øn m√£i n√†o c√≤n hi·ªáu l·ª±c cho ' + esc(S.vac.name) + '.' : 'Hi·ªán ch∆∞a c√≥ ch∆∞∆°ng tr√¨nh khuy·∫øn m√£i n√†o c√≤n hi·ªáu l·ª±c trong kho.') + '</div>') + '</div></div>' + checkFoot(n) + '</div>';
    host.innerHTML = h; bindCheck('.mcard[data-c]', 'km');
    [].slice.call(host.querySelectorAll('[data-cf]')).forEach(function (b) { b.onclick = function () { S.cf = b.dataset.cf; draw(true) } });
    [].slice.call(host.querySelectorAll('[data-cm]')).forEach(function (b) { b.onclick = function () { S.cm = b.dataset.cm === '1'; S.cf = ''; draw(true) } })
  }
  function drawNc() {
    var P = S.vac ? profile(S.vac) : null, mset = {}; if (P) P.nc.forEach(function (r) { mset[r.id] = 1 }); var q = fk(S.nq.trim()), L = ncList();
    var n = Object.keys(S.nc).filter(function (k) { return S.nc[k] }).length;
    L = L.filter(function (r) { return !q || fk(r.d.benh + ' ' + r.d.nguy_co).indexOf(q) >= 0 }).map(function (r) { r.hit = !!mset[r.id]; return r }); L.sort(function (a, b) { return (b.hit ? 1 : 0) - (a.hit ? 1 : 0) });
    var h = '<div class="msheet">' + hdrBack('Nguy c∆° & ƒë∆∞·ªùng l√¢y', ncList().length + ' b·ªánh trong kho ¬∑ n·ªôi dung do ASM duy·ªát') + '<div class="mbody"><div class="msearch">' + ic('search') + '<input id="nqi" placeholder="T√¨m b·ªánh‚Ä¶" value="' + esc(S.nq) + '" aria-label="T√¨m b·ªánh"></div><div class="mlist">' +
      (L.length ? L.map(function (r) { return '<button class="mcard ' + (S.nc[r.id] ? 'on' : '') + '" data-c="' + r.id + '"><div class="mc1"><b>' + esc(r.d.benh) + '</b><span class="mck">' + ic('check') + '</span></div>' + (r.hit ? '<div class="mtags"><span class="mp gr">ƒê√∫ng vaccine</span></div>' : '') + '<div class="pv2">' + esc(String(r.d.nguy_co || '').replace(/\s+/g, ' ').slice(0, 160)) + '</div></button>' }).join('') : '<div class="mwarn">Kh√¥ng c√≥ b·ªánh ph√π h·ª£p.</div>') + '</div></div>' + checkFoot(n) + '</div>';
    host.innerHTML = h; bindCheck('.mcard[data-c]', 'nc');
    var ni = host.querySelector('#nqi'); ni.oninput = function () { S.nq = ni.value; var pos = ni.selectionStart; draw(); var m = host.querySelector('#nqi'); m.focus(); m.setSelectionRange(pos, pos) }
  }

  /* ---------- ·∫£nh ƒë√≠nh k√®m: g·ª£i √Ω theo kh√°ch v√† vaccine, nh√¢n vi√™n ch·ªçn r·ªìi chia s·∫ª ---------- */
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
    return '<section><div class="mh">·∫¢nh ƒë√≠nh k√®m ¬∑ g·ª£i √Ω theo kh√°ch v√† vaccine</div>' + (L.length ? '<div class="imgrow">' + L.map(function (x) { return thumbBtn(x.f) }).join('') + '</div>' : '<div class="mnote"><span>Ch∆∞a c√≥ ·∫£nh n√†o h·ª£p. B·∫•m "Xem t·∫•t c·∫£ ·∫£nh".</span></div>') + '<button class="abt o" id="imore">' + ic('image') + 'Xem t·∫•t c·∫£ ·∫£nh (' + F.length + ')</button></section>'
  }
  function thumbBtn(f) { return '<button class="ithumb ' + (S.img[f.id] ? 'on' : '') + '" data-img="' + esc(f.id) + '">' + (f.anh_nho ? '<img src="' + f.anh_nho + '" alt="">' : '<div style="height:64px;display:grid;place-items:center;color:var(--soft)">PDF</div>') + '<span>' + esc(f.ten.slice(0, 24)) + '</span>' + (S.img[f.id] ? '<b class="ick">‚úì</b>' : '') + '</button>' }
  function drawImgs() {
    var F = D.files || [], q = fk(S.iq), L = F.filter(function (f) { return (!S.it || f.thu_muc === S.it) && (!q || fk(f.ten + ' ' + f.chu_de).indexOf(q) >= 0) }), T = { '': 'T·∫•t c·∫£', VACCINE_DOI_TUONG: 'Vaccine', CAM_NANG_BAU: 'C·∫©m nang b·∫ßu', CTKM_VOUCHER: 'Khuy·∫øn m√£i' }, nI = Object.keys(S.img).length;
    host.innerHTML = '<div class="msheet">' + hdrBack('Ch·ªçn ·∫£nh ƒë√≠nh k√®m', nI ? nI + ' ·∫£nh ƒë√£ ch·ªçn' : 'Ch·∫°m v√†o ·∫£nh ƒë·ªÉ ch·ªçn') + '<div class="mbody"><div class="mchips">' + Object.keys(T).map(function (k) { return '<button class="' + (S.it === k ? 'on' : '') + '" data-it="' + k + '">' + T[k] + '</button>' }).join('') + '</div><div class="msearch">' + ic('search') + '<input id="iqi" placeholder="T√¨m ·∫£nh‚Ä¶" value="' + esc(S.iq) + '" aria-label="T√¨m ·∫£nh"></div><div class="imggrid">' + (L.length ? L.map(thumbBtn).join('') : '<div class="mwarn" style="grid-column:1/-1">Kh√¥ng c√≥ ·∫£nh ph√π h·ª£p.</div>') + '</div></div>' + checkFoot(nI) + '</div>';
    host.querySelector('#mx').onclick = close; host.querySelector('#pb').onclick = function () { S.view = 'main'; draw() }; host.querySelector('#done').onclick = function () { S.view = 'main'; draw() };
    [].slice.call(host.querySelectorAll('[data-it]')).forEach(function (b) { b.onclick = function () { S.it = b.dataset.it; draw(true) } });
    var qi = host.querySelector('#iqi'); qi.oninput = function () { S.iq = qi.value; var p = qi.selectionStart; draw(); var n = host.querySelector('#iqi'); n.focus(); n.setSelectionRange(p, p) };
    [].slice.call(host.querySelectorAll('[data-img]')).forEach(function (b) { b.onclick = function () { var k = b.dataset.img; S.img[k] ? delete S.img[k] : S.img[k] = 1; draw(true) } })
  }
  function sendImages() {
    var F = (D.files || []).filter(function (f) { return S.img[f.id] }), text = S.text; if (!F.length) return; toast('ƒêang chu·∫©n b·ªã ' + F.length + ' ·∫£nh‚Ä¶');
    Promise.all(F.map(function (f) { return fileBlob(f).then(function (b) { var ext = f.mime === 'image/png' ? 'png' : f.mime === 'application/pdf' ? 'pdf' : f.mime === 'image/webp' ? 'webp' : 'jpg'; return new File([b], (f.ten.replace(/[^\w\- ]+/g, '').slice(0, 40) || 'anh') + '.' + ext, { type: f.mime }) }) })).then(function (files) {
      if (navigator.canShare && navigator.canShare({ files: files })) return navigator.share({ files: files, text: text }).then(function () { toast('ƒê√£ m·ªü chia s·∫ª ·∫£nh. Ch·ªçn Zalo ƒë·ªÉ g·ª≠i.') }, function (e) { if (e && e.name !== 'AbortError') toast('Kh√¥ng chia s·∫ª ƒë∆∞·ª£c, ƒë√£ t·∫£i ·∫£nh v·ªÅ m√°y.', 0, 1) });
      files.forEach(function (f) { var a = document.createElement('a'); a.href = URL.createObjectURL(f); a.download = f.name; document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove() }, 800) }); toast('M√°y n√†y kh√¥ng chia s·∫ª tr·ª±c ti·∫øp ƒë∆∞·ª£c, ƒë√£ t·∫£i ' + files.length + ' ·∫£nh v·ªÅ ƒë·ªÉ anh/ch·ªã ƒë√≠nh k√®m.')
    }, function (e) { fail(e) })
  }

  /* ---------- g·ª≠i: m·ªói n√∫t l√†m vi·ªác c·ªßa n√≥, r·ªìi ghi 1 l·∫ßn "Nh·∫Øn" ---------- */
  function send(kind) {
    if (S.sending) return; var l = link(), tel = telOf(l), text = S.text.trim(), t = S.t;
    if (!text) { toast('Tin ƒëang tr·ªëng, h√£y ch·ªçn vaccine, m·∫´u ho·∫∑c t·ª± vi·∫øt.', 0, 1); return }
    if (text.length > MAXLEN) { toast('Tin d√†i qu√° ' + MAXLEN + ' k√Ω t·ª±, h√£y r√∫t ng·∫Øn.', 0, 1); return }
    if (kind !== 'SAO_CHEP' && !tel) { toast('Ng∆∞·ªùi li√™n h·ªá n√†y ch∆∞a c√≥ s·ªë ƒëi·ªán tho·∫°i.', 0, 1); return }
    var ph = openPH(); if (ph.length && !confirm('Tin c√≤n ch·ªó tr·ªëng ch∆∞a ƒëi·ªÅn: ' + ph.map(function (n) { return PHL[n] || n }).join(', ') + '.\nCh·ªØ {{‚Ä¶}} s·∫Ω g·ª≠i nguy√™n cho kh√°ch. V·∫´n g·ª≠i?')) return;
    if (S.sig && !ME.sdt && !confirm('Anh/ch·ªã ch∆∞a c√≥ SƒêT trong t√†i kho·∫£n n√™n tin thi·∫øu s·ªë li√™n h·ªá. V·∫´n g·ª≠i?')) return;
    S.sending = true; var edited = S.text !== S.base, pk = S.mode === 'QT' ? S.pq : (S.pk || S.st), tpl = pk ? pk.ma : '', step = S.mode === 'QT' ? S.step : '', nlh = l.ma_nlh, vtag = S.vac ? S.vac.name : '';
    copy(text);
    if (kind === 'ZALO') window.open('https://zalo.me/' + tel, '_blank', 'noopener');
    if (kind === 'SMS') { var a = document.createElement('a'); a.href = 'sms:' + tel + (/iPhone|iPad|iPod/.test(navigator.userAgent) ? '&' : '?') + 'body=' + encodeURIComponent(text); document.body.appendChild(a); a.click(); a.remove() }
    var chan = { ZALO: 'ZALO', SMS: 'NHAN_SMS', SAO_CHEP: 'NHAN_CHEP' }[kind], loai = byMa('LOAI_TUONG_TAC', chan), prevSt = t.trang_thai, fake = null;
    if (loai) { fake = { ma_luot: t.ma_luot, ma_tk: ME.ma_tk, loai_tuong_tac: loai.ma_dm, ket_qua_lien_lac: 'DM-KL6', ghi_chu: '', d: new Date(), sdt_lien_he: tel }; D.tt.push(fake); relast(t.ma_luot); if (t.trang_thai === 'DA_GIAO') t.trang_thai = 'DANG_XU_LY' }
    var msg = { ZALO: 'ƒê√£ ch√©p tin v√† m·ªü Zalo t·ªõi ' + tel + '. D√°n tin v√†o √¥ chat ƒë·ªÉ g·ª≠i ho·∫∑c k·∫øt b·∫°n.', SMS: 'ƒê√£ m·ªü tin nh·∫Øn SMS t·ªõi ' + tel + ' (tin c≈©ng ƒë√£ ƒë∆∞·ª£c ch√©p s·∫µn).', SAO_CHEP: 'ƒê√£ ch√©p tin. D√°n v√†o ·ª©ng d·ª•ng tin nh·∫Øn ƒë·ªÉ g·ª≠i.' }[kind];
    close(); toast(msg + ' ƒê√£ ghi 1 l·∫ßn Nh·∫Øn.'); render();
    rpc('u_msg_log', { p_ma_luot: t.ma_luot, p_kenh: kind, p_ma_nlh: nlh || '', p_ma_mau: tpl || '', p_buoc: step ? (vtag ? step + ' ¬∑ ' + vtag : step) : (vtag ? 'Vaccine ' + vtag : ''), p_da_sua: edited, p_noi_dung: text.slice(0, 2000), p_req: rid() })
      .then(function () { return load(true) }).then(render, function (e) { if (fake) { var i = D.tt.indexOf(fake); if (i >= 0) D.tt.splice(i, 1); t.trang_thai = prevSt; relast(t.ma_luot); render() } fail(e) })
  }

  /* ---------- TRA C·ª®U VACCINE (H·∫£i 24/11 ƒë√™m): h·ªì s∆° t·ª´ng vaccine l·∫•y t·ª´ Th∆∞ vi·ªán ƒë√£ duy·ªát ----------
   * Danh s√°ch = khu "Ph√°c ƒë·ªì & t·ª´ ƒëi·ªÉn vaccine" (n·∫øu ƒë√£ n·∫°p), kh√¥ng th√¨ B·∫£ng gi√°. M·ªói th·∫ª: tu·ªïi, ph√°c ƒë·ªì, n∆∞·ªõc SX, ch·ªëng ch·ªâ ƒë·ªãnh, ph·∫£n ·ª©ng sau ti√™m, nguy c∆° & ƒë∆∞·ªùng l√¢y,
   * gi√° + khuy·∫øn m√£i h√¥m nay, c√¢u t∆∞ v·∫•n, tin sau ti√™m. M·ªü t·ª´ m√†n so·∫°n tin th√¨ c√≥ n√∫t "D√πng cho tin nh·∫Øn" ƒë·ªÉ t·ª± ƒëi·ªÅn ƒë√∫ng vaccine ƒë√≥. */
  var LK = null, lkHost = null;
  var LK_AGES = [['', 'M·ªçi l·ª©a tu·ªïi', null, null], ['A', 'S∆° sinh (d∆∞·ªõi 1 th√°ng)', 0, 30], ['B', '1 ‚Äì 12 th√°ng', 30, 365], ['C', '1 ‚Äì 2 tu·ªïi', 365, 730], ['D', '2 ‚Äì 9 tu·ªïi', 730, 3285], ['E', '9 ‚Äì 18 tu·ªïi', 3285, 6570], ['F', 'Ng∆∞·ªùi l·ªõn 18 ‚Äì 50 tu·ªïi', 6570, 18263], ['G', 'Tr√™n 50 tu·ªïi', 18263, 40000]];
  function lkRows() {
    var pd = libRows('PHAC_DO'); if (pd.length) return pd.map(function (r) { return { src: 'PD', id: 'pd' + r.id, name: r.d.ten_vaccine || '', benh: oneLine(r.d.phong_benh), d: r.d } });
    return libRows('BANG_GIA').map(function (r) { return { src: 'BG', id: 'bg' + r.id, name: r.d.ten_vaccine || '', benh: oneLine(r.d.phong_benh), d: r.d } })
  }
  function lkVac(row) {
    var ck = 'lk:' + row.id + '|' + LIB.at; if (VC[ck]) return VC[ck];
    var key = brandKey(row.name), groups = bgIndex().g.filter(function (g) { return g.keys.indexOf(key) >= 0 || g.bn.some(function (b) { return has(b, key) }) }), phr = phrasesOf(row.benh + ';' + (row.d.nhom_vaccine || ''));
    groups.forEach(function (g) { phrasesOf(g.short).forEach(function (p) { if (phr.indexOf(p) < 0) phr.push(p) }) });
    var htxt = nv(row.benh + ' ; ' + (row.d.nhom_vaccine || '')); ncList().forEach(function (r) { var b = nv(r.d.benh).trim(); if (b.length >= 3 && has(htxt, b) && phr.indexOf(b) < 0) phr.push(b) });   /* t√™n b·ªánh trong kho Nguy c∆° n·∫±m trong m√¥ t·∫£ b·ªánh c·ªßa vaccine (v√≠ d·ª• "N√£o m√¥ c·∫ßu" ‚äÇ "Vi√™m m√†ng n√£o m√¥ c·∫ßu B+C") */
    return (VC[ck] = { id: 'lk:' + row.id, name: row.name, phr: phr, keys: [key], groups: groups, from: 'LK', group: groups[0] ? groups[0].short : shortName(row.benh) })
  }
  function lkGroup(row) { var V = lkVac(row); return V.group || shortName(row.benh) || 'Kh√°c' }
  function lkPrice(row) { var key = brandKey(row.name); return libRows('BANG_GIA').filter(function (r) { return has(nv(r.d.ten_vaccine), key) }) }
  function lkKm(row, P) { var key = brandKey(row.name), out = []; P.km.forEach(function (g) { g.items.forEach(function (d) { if (has(nv(d.ten_sp), key)) out.push({ g: g, d: d }) }) }); return out }
  function lkAgeOk(row, a) { if (!a || a[2] == null) return true; var d = row.d, tu = Number(d.tuoi_bat_dau_ngay), kt = d.tuoi_ket_thuc_ngay == null ? null : Number(d.tuoi_ket_thuc_ngay); if (isNaN(tu)) return true; return tu < a[3] && (kt == null || kt >= a[2]) }
  function lookup(opts) {
    if (lkHost) return; opts = opts || {}; LK = { q: opts.q || '', age: '', benh: '', lim: 30, sel: null, fromSheet: !!(S && host), src: null };
    lkHost = document.createElement('div'); lkHost.className = 'modal mm first'; lkHost.style.zIndex = 60; lkHost.onclick = function (e) { if (e.target === lkHost) lkClose() }; document.body.appendChild(lkHost);
    document.addEventListener('keydown', lkKey); setTimeout(function () { if (lkHost) lkHost.classList.remove('first') }, 450); loadLib(); lkDraw()
  }
  function lkClose() { document.removeEventListener('keydown', lkKey); if (lkHost) lkHost.remove(); lkHost = null; LK = null }
  function lkKey(e) { if (e.key === 'Escape' && lkHost) { if (LK.sel) { LK.sel = null; lkDraw() } else lkClose() } }
  function lkDraw() {
    if (!lkHost || !LK) return; var b = lkHost.querySelector('.mbody'), top = b ? b.scrollTop : 0; if (LK.sel) lkDetail(); else lkList(); if (top && !LK.sel) { var nb = lkHost.querySelector('.mbody'); if (nb) nb.scrollTop = top }
  }
  function lkList() {
    var rows = lkRows(), q = fk(LK.q.trim()), qt = q ? q.split(/\s+/) : [], ageSel = LK_AGES.filter(function (a) { return a[0] === LK.age })[0], groups = {}, order = [];
    rows.forEach(function (r) { var g = lkGroup(r); if (!groups[g]) { groups[g] = 0; order.push(g) } groups[g]++ }); order.sort(function (a, b) { return a.localeCompare(b, 'vi') });
    var list = rows.filter(function (r) { return (!LK.benh || lkGroup(r) === LK.benh) && lkAgeOk(r, ageSel) && (!qt.length || (function (hay) { return qt.every(function (t) { return hay.indexOf(t) >= 0 }) })(fk(r.name + ' ' + r.benh + ' ' + (r.d.nuoc_sx || '') + ' ' + (r.d.nhom_vaccine || '') + ' ' + lkGroup(r)))) });
    list.sort(function (a, b) { return lkGroup(a).localeCompare(lkGroup(b), 'vi') || a.name.localeCompare(b.name, 'vi') });
    var more = list.length > LK.lim, shown = list.slice(0, LK.lim), pdOn = rows.length && rows[0].src === 'PD';
    var h = '<div class="msheet"><div class="mhd"><div class="mtop"><div class="mav">' + ic('book') + '</div><div class="mtt"><div class="t1">Tra c·ª©u vaccine</div><div class="t2">' + (LIB.rows ? rows.length + ' vaccine' + (pdOn ? ' ¬∑ t·ª´ Ph√°c ƒë·ªì & t·ª´ ƒëi·ªÉn' : ' ¬∑ t·ª´ B·∫£ng gi√° (ch∆∞a n·∫°p ph√°c ƒë·ªì)') : 'ƒêang t·∫£i‚Ä¶') + '</div></div><button class="mclose" id="lx" aria-label="ƒê√≥ng">' + ic('x') + '</button></div></div><div class="mbody">' +
      '<div class="msearch">' + ic('search') + '<input id="lq" placeholder="T√™n vaccine, b·ªánh, n∆∞·ªõc s·∫£n xu·∫•t‚Ä¶ (v√≠ d·ª• Bexsero, HPV, B·ªâ)" value="' + esc(LK.q) + '" aria-label="T√¨m vaccine"></div>' +
      '<div class="lkf"><select id="la" aria-label="L·ª©a tu·ªïi">' + LK_AGES.map(function (a) { return '<option value="' + a[0] + '"' + (a[0] === LK.age ? ' selected' : '') + '>' + esc(a[1]) + '</option>' }).join('') + '</select><select id="lb" aria-label="Nh√≥m b·ªánh"><option value="">M·ªçi nh√≥m b·ªánh</option>' + order.map(function (g) { return '<option value="' + esc(g) + '"' + (g === LK.benh ? ' selected' : '') + '>' + esc(g.slice(0, 46)) + ' (' + groups[g] + ')</option>' }).join('') + '</select></div>';
    if (LIB.loading && !LIB.rows) h += skeleton();
    else if (!rows.length) h += '<div class="mwarn"><b>Ch∆∞a c√≥ d·ªØ li·ªáu ƒë·ªÉ tra c·ª©u.</b> ASM n·∫°p B·∫£ng gi√° ho·∫∑c "Ph√°c ƒë·ªì & t·ª´ ƒëi·ªÉn vaccine" ·ªü m·ª•c Th∆∞ vi·ªán th√¨ danh s√°ch s·∫Ω hi·ªán ·ªü ƒë√¢y.</div>';
    else h += '<div class="mnote"><span>' + list.length + ' k·∫øt qu·∫£. Ch·∫°m v√†o vaccine ƒë·ªÉ xem h·ªì s∆°.</span></div><div class="mlist">' + (shown.length ? shown.map(function (r) {
      var d = r.d, age = d.do_tuoi || '', tags = (d.nuoc_sx ? '<span class="mp">' + esc(d.nuoc_sx) + '</span>' : '') + (age ? '<span class="mp ro">' + esc(age.slice(0, 40)) + '</span>' : '');
      return '<button class="mcard" data-lk="' + esc(r.id) + '"><div class="mc1"><b>' + esc(r.name) + '</b>' + ic('chev') + '</div><div class="mnote" style="margin-top:-2px"><span>' + esc(lkGroup(r).slice(0, 60)) + '</span></div>' + (tags ? '<div class="mtags">' + tags + '</div>' : '') + '</button>'
    }).join('') + (more ? '<button class="abt o" id="lmore2">Hi·ªán th√™m (' + (list.length - LK.lim) + ')</button>' : '') : '<div class="mwarn">Kh√¥ng c√≥ vaccine ph√π h·ª£p b·ªô l·ªçc.</div>') + '</div>';
    h += '</div></div>'; lkHost.innerHTML = h; lkHost.querySelector('#lx').onclick = lkClose;
    var li = lkHost.querySelector('#lq'); li.oninput = function () { LK.q = li.value; LK.lim = 30; var pos = li.selectionStart; lkDraw(); var n = lkHost.querySelector('#lq'); n.focus(); n.setSelectionRange(pos, pos) };
    var sa = lkHost.querySelector('#la'); sa.onchange = function () { LK.age = sa.value; LK.lim = 30; lkDraw() }; var sb = lkHost.querySelector('#lb'); sb.onchange = function () { LK.benh = sb.value; LK.lim = 30; lkDraw() };
    var lm = lkHost.querySelector('#lmore2'); if (lm) lm.onclick = function () { LK.lim += 30; lkDraw() };
    [].slice.call(lkHost.querySelectorAll('[data-lk]')).forEach(function (b) { b.onclick = function () { LK.sel = b.dataset.lk; lkDraw(); var nb = lkHost.querySelector('.mbody'); if (nb) nb.scrollTop = 0 } })
  }
  function lkSec(title, body) { return body ? '<section><div class="mh">' + esc(title) + '</div>' + body + '</section>' : '' }
  function lkSummary(row, P, price, km) {
    var d = row.d, L = [row.name + ' ‚Äì ' + (row.benh || '')]; if (d.do_tuoi) L.push('ƒê·ªô tu·ªïi: ' + d.do_tuoi); if (d.nuoc_sx) L.push('N∆∞·ªõc SX: ' + d.nuoc_sx); if (d.noi_dung_tu_van) L.push('Ph√°c ƒë·ªì: ' + oneLine(d.noi_dung_tu_van));
    if (P.nc[0]) L.push('Nguy c∆°: ' + oneLine(P.nc[0].d.nguy_co).slice(0, 200)); if (price.length) L.push('Gi√°: ' + price.map(function (r) { return money(r.d.gia) }).join(' / ')); if (km.length) L.push('Khuy·∫øn m√£i: ' + km.slice(0, 3).map(function (x) { return itemLine(x.d, false, false).replace(/^- /, '') }).join('; '));
    return L.join('\n')
  }
  function lkDetail() {
    var row = lkRows().filter(function (r) { return r.id === LK.sel })[0]; if (!row) { LK.sel = null; lkList(); return }
    var V = lkVac(row), P = profile(V), d = row.d, price = lkPrice(row), km = lkKm(row, P), has_ = function (v) { return v != null && String(v).trim() !== '' };
    var tu = has_(d.tuoi_bat_dau_ngay) ? Number(d.tuoi_bat_dau_ngay) : null, kt = has_(d.tuoi_ket_thuc_ngay) ? Number(d.tuoi_ket_thuc_ngay) : null, infer = /t·ª± suy/.test(d.ghi_chu || '');
    var age = '';
    if (d.do_tuoi) age += '<div><b>ƒê·ªô tu·ªïi ch·ªâ ƒë·ªãnh:</b> ' + esc(d.do_tuoi) + '</div>';
    if (tu != null) age += '<div><b>T·ª´</b> ' + esc(dayLabel(tu)) + ' tu·ªïi' + (kt != null ? ' <b>ƒë·∫øn</b> ' + esc(dayLabel(kt)) : ' <b>tr·ªü l√™n</b>') + (infer && kt != null ? ' <span class="mp or">tu·ªïi k·∫øt th√∫c t·ª± suy, ch·ªù b√°c sƒ©</span>' : '') + '</div>';
    if (d.noi_dung_tu_van) age += '<div style="margin-top:6px"><b>Ph√°c ƒë·ªì tham kh·∫£o:</b><br>' + esc(nl(d.noi_dung_tu_van)).replace(/\n/g, '<br>') + '</div>';
    if (d.so_mui) age += '<div><b>S·ªë m≈©i:</b> ' + esc(d.so_mui) + '</div>'; if (d.khoang_cach_ngay) age += '<div><b>Kho·∫£ng c√°ch m≈©i:</b> ' + esc(d.khoang_cach_ngay) + '</div>';
    var pro = ''; if (d.chong_chi_dinh) pro += '<div><b>Ch·ªëng ch·ªâ ƒë·ªãnh / th·∫≠n tr·ªçng:</b> ' + esc(d.chong_chi_dinh) + '</div>'; if (d.phan_ung) pro += '<div><b>Ph·∫£n ·ª©ng sau ti√™m th∆∞·ªùng g·∫∑p:</b> ' + esc(d.phan_ung) + '</div>'; if (d.nhom_vaccine) pro += '<div><b>Nh√≥m vaccine:</b> ' + esc(d.nhom_vaccine) + '</div>';
    var nc = P.nc.map(function (r) { var x = r.d; return '<div class="lkb"><b>' + esc(x.benh) + '</b><div>' + esc(oneLine(x.nguy_co)) + '</div>' + (x.lay_qua ? '<div><b>L√¢y qua:</b> ' + esc(oneLine(x.lay_qua)) + '</div>' : '') + (x.trieu_chung ? '<div><b>Tri·ªáu ch·ª©ng:</b> ' + esc(oneLine(x.trieu_chung)) + '</div>' : '') + (x.mua_vung ? '<div><b>M√πa & v√πng:</b> ' + esc(oneLine(x.mua_vung)) + '</div>' : '') + (x.diem_can_nho ? '<div><b>C·∫ßn nh·ªõ:</b> ' + esc(oneLine(x.diem_can_nho)) + '</div>' : '') + '</div>' }).join('');
    var gia = price.map(function (r) { return '<div class="lkr"><span>' + esc(r.d.ten_vaccine) + '</span><b>' + esc(money(r.d.gia)) + '</b></div>' }).join('') + (price[0] && (price[0].d.tu_ngay || price[0].d.den_ngay) ? '<div class="meta">K·ª≥ √°p d·ª•ng: ' + esc(dmy(price[0].d.tu_ngay)) + ' ‚Äì ' + esc(dmy(price[0].d.den_ngay)) + '</div>' : '');
    var kmh = km.map(function (x) { return '<div class="lkb"><b>' + esc(x.g.title) + '</b><div>' + esc(itemLine(x.d, false, false).replace(/^- /, '')) + '</div>' + (x.d.ghi_chu ? '<div class="meta">' + esc(oneLine(x.d.ghi_chu)) + '</div>' : '') + '<div class="meta">√Åp d·ª•ng ƒë·∫øn ' + esc(dmy(x.g.den)) + '</div></div>' }).join('');
    var tv = P.tv.slice(0, 3).map(function (o) { return '<div class="lkb"><b>' + esc(o.ten) + '</b><div>' + esc(oneLine(o.text).slice(0, 220)) + '‚Ä¶</div></div>' }).join('');
    var st = (P.st.length ? P.st : []).slice(0, 2).map(function (o) { return '<div class="lkb"><b>' + esc(o.ten) + '</b><div>' + esc(oneLine(o.text).slice(0, 200)) + '‚Ä¶</div></div>' }).join('');
    var h = '<div class="msheet"><div class="mhd"><div class="mtop"><button class="mclose" id="lb2" aria-label="Quay l·∫°i">' + ic('back') + '</button><div class="mtt"><div class="t1">' + esc(row.name) + '</div><div class="t2">' + esc(lkGroup(row)) + '</div></div><button class="mclose" id="lx" aria-label="ƒê√≥ng">' + ic('x') + '</button></div></div><div class="mbody">' +
      '<div class="tcard"><div class="mtags" style="margin-top:0">' + (d.nuoc_sx ? '<span class="mp">N∆∞·ªõc SX: ' + esc(d.nuoc_sx) + '</span>' : '') + (row.benh ? '<span class="mp ro">' + esc(row.benh.slice(0, 70)) + '</span>' : '') + '</div></div>' +
      lkSec('ƒê·ªô tu·ªïi & ph√°c ƒë·ªì', age ? '<div class="ls lkp">' + age + '</div>' : '<div class="mnote"><span>Ch∆∞a c√≥ tu·ªïi, ph√°c ƒë·ªì trong kho cho vaccine n√†y.</span></div>') + lkSec('L∆∞u √Ω chuy√™n m√¥n', pro ? '<div class="ls lkp">' + pro + '</div>' : '<div class="mnote"><span>Ch∆∞a c√≥ ch·ªëng ch·ªâ ƒë·ªãnh, ph·∫£n ·ª©ng sau ti√™m trong kho (ASM b·ªï sung ·ªü Th∆∞ vi·ªán ‚Üí Ph√°c ƒë·ªì & t·ª´ ƒëi·ªÉn vaccine).</span></div>') +
      lkSec('Nguy c∆° & ƒë∆∞·ªùng l√¢y', nc || '<div class="mnote"><span>Kho Nguy c∆° ch∆∞a c√≥ b·ªánh n√†y.</span></div>') + lkSec('Gi√° ni√™m y·∫øt', gia) + lkSec('Khuy·∫øn m√£i c√≤n hi·ªáu l·ª±c h√¥m nay', kmh || '<div class="mnote"><span>H√¥m nay ch∆∞a c√≥ khuy·∫øn m√£i cho vaccine n√†y.</span></div>') + lkSec('C√¢u t∆∞ v·∫•n g·ª£i √Ω (' + P.tv.length + ')', tv || '<div class="mnote"><span>Kho ch∆∞a c√≥ m·∫´u t∆∞ v·∫•n ri√™ng.</span></div>') + lkSec('Tin CSKH sau ti√™m (' + P.st.length + ')', st) +
      '<div class="mnote"><span>N·ªôi dung do ASM duy·ªát trong Th∆∞ vi·ªán. App ch·ªâ tra c·ª©u v√† g·ª£i √Ω, kh√¥ng thay ch·ªâ ƒë·ªãnh c·ªßa b√°c sƒ©.</span></div></div><div class="mft"><div class="act2">' + (LK.fromSheet ? '<button class="abt p" id="luse">' + ic('chat') + '<span><b>D√πng cho tin nh·∫Øn</b><small>t·ª± ƒëi·ªÅn ƒë√∫ng vaccine n√†y</small></span></button>' : '') + '<button class="abt g" id="lcopy">' + ic('copy') + '<span><b>Sao ch√©p t√≥m t·∫Øt</b><small>g·ª≠i cho ƒë·ªìng nghi·ªáp</small></span></button></div></div></div>';
    lkHost.innerHTML = h; lkHost.querySelector('#lx').onclick = lkClose; lkHost.querySelector('#lb2').onclick = function () { LK.sel = null; lkDraw() };
    var cp = lkHost.querySelector('#lcopy'); if (cp) cp.onclick = function () { copy(lkSummary(row, P, price, km)); toast('ƒê√£ ch√©p t√≥m t·∫Øt ' + row.name) };
    var us = lkHost.querySelector('#luse'); if (us) us.onclick = function () { lkClose(); if (S && host) setVaccine(lkVac(row)) }
  }

  return { open: open, counters: counters, scriptBlock: scriptBlock, preview: preview, fill: fill, dueRules: dueRules, custDays: custDays, dayLabel: dayLabel, VARS: VARS, VARDESC: VARDESC, STEPS: STEPS, GROUPS: GROUPS, ensureLib: ensureLib, linkReport: linkReport, lookup: lookup, _t: { audFit: audFit, ctkmGroups: ctkmGroups, ctkmText: ctkmText, ncText: ncText, itemsOf: itemsOf, libRows: libRows, isBau: isBau, LIB: LIB, nv: nv, phrasesOf: phrasesOf, brandKey: brandKey, mkVRule: mkVRule, mkVbg: mkVbg, profile: profile, bgIndex: bgIndex, linkReport: linkReport } }
})();

;
