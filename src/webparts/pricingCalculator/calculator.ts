// @ts-nocheck
/* eslint-disable */

export function initCalculator(domElement: HTMLElement, siteUrl: string): void {

  if (domElement.querySelector('#main')) return;

  // Force LTR direction so inputs don't type backwards in RTL SharePoint sites
  domElement.setAttribute('dir', 'ltr');
  domElement.style.direction = 'ltr';

  // ── Fonts ─────────────────────────────────────────────────────────
  if (!document.getElementById('carbon-calc-fonts')) {
    const lnk = document.createElement('link');
    lnk.id  = 'carbon-calc-fonts';
    lnk.rel = 'stylesheet';
    lnk.href = 'https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&display=swap';
    document.head.appendChild(lnk);
  }

  // ── CSS ───────────────────────────────────────────────────────────
  if (!document.getElementById('carbon-calc-css')) {
    const st = document.createElement('style');
    st.id = 'carbon-calc-css';
    st.textContent = [
      '*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }',
      ':root { --g:#4db749;--gl:#e8f5e7;--gd:#3a9c38;--b:#29b6c9;--bl:#e4f6f9;--dk:#2d2d2d;--gy:#f5f5f5;--bd:#e0e0e0;--rd:#d94444;--font:"DM Sans",sans-serif;--serif:Tahoma,"Trebuchet MS",Arial,sans-serif; }',
      'body { background:#fafafa; font-family:var(--font); color:var(--dk); min-height:100vh; direction:ltr; }',
      'input,select,textarea { font-family:var(--font); direction:ltr !important; text-align:left !important; unicode-bidi:isolate; }',
      'button { font-family:var(--font); }',
      'input[type=number]::-webkit-inner-spin-button { opacity:1; }',
      '#sp-status-chip { font-size:11px;font-weight:600;padding:4px 10px;border-radius:20px;border:1px solid #ddd;background:#f0f0f0;color:#888;letter-spacing:.03em;white-space:nowrap;transition:all .4s; }',
      '#header { background:var(--dk);padding:22px 24px; }',
      '#header-inner { max-width:980px;margin:0 auto;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px; }',
      '.logo { display:flex;align-items:center; }',
      '.header-tag { font-size:12px;color:rgba(255,255,255,.5);letter-spacing:.06em;text-transform:uppercase; }',
      '#hero { background:linear-gradient(135deg,var(--g) 0%,var(--gd) 100%);padding:28px 24px; }',
      '#hero-inner { max-width:980px;margin:0 auto; }',
      '#hero h1 { font-family:var(--serif);font-size:28px;font-weight:700;color:#fff; }',
      '#hero p { font-size:14px;color:rgba(255,255,255,.8);margin-top:6px; }',
      '#main { max-width:980px;margin:0 auto;padding:28px 24px 80px; }',
      '.card { background:#fff;border-radius:12px;padding:24px;margin-bottom:20px;border:1px solid var(--bd);box-shadow:0 1px 4px rgba(0,0,0,.04); }',
      '.card-green { border-top:3px solid var(--g); }',
      '.card-blue { border-top:3px solid var(--b); }',
      '.coll { border-radius:12px;margin-bottom:20px;overflow:hidden;border:1px solid var(--bd);box-shadow:0 1px 4px rgba(0,0,0,.04);transition:all .2s; }',
      '.coll.collapsed { background:#f8f8f8;border-color:#e8e8e8;box-shadow:none; }',
      '.coll-header { width:100%;padding:16px 24px;background:transparent;border:none;cursor:pointer;display:flex;justify-content:space-between;align-items:center;text-align:left; }',
      '.coll-header:hover { background:rgba(0,0,0,.02); }',
      '.coll-left { display:flex;align-items:center;gap:12px; }',
      '.coll-chevron { transition:transform .2s;flex-shrink:0; }',
      '.coll.expanded .coll-chevron { transform:rotate(90deg); }',
      '.coll-title-wrap .coll-num { font-size:10px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;display:block; }',
      '.coll-title-wrap h3 { font-family:var(--serif);font-size:16px;font-weight:700;margin-top:2px;transition:color .2s; }',
      '.coll-title-wrap .coll-sub { font-size:11px;color:#bbb;margin-top:2px; }',
      '.coll-active-badge { font-size:9px;font-weight:700;padding:2px 6px;border-radius:4px;letter-spacing:.04em;text-transform:uppercase; }',
      '.coll-right { text-align:right; }',
      '.coll-right .coll-total-label { font-size:10px;color:#bbb;text-transform:uppercase;letter-spacing:.06em; }',
      '.coll-right .coll-total-val { font-size:16px;font-weight:700;font-variant-numeric:tabular-nums;transition:color .2s; }',
      '.coll-body { padding:4px 24px 24px;border-top:1px solid var(--bd); }',
      '.coll-body .coll-body-sub { font-size:12px;color:#999;margin:12px 0 8px; }',
      '.sec-head { display:flex;justify-content:space-between;align-items:flex-end;margin-bottom:16px;padding-bottom:10px; }',
      '.sec-head-left .sec-num { font-size:11px;font-weight:700;letter-spacing:.1em;text-transform:uppercase; }',
      '.sec-head-left h3 { font-family:var(--serif);font-size:18px;font-weight:700;margin-top:4px;color:var(--dk); }',
      '.sec-head-left p { font-size:12px;color:#999;margin-top:2px; }',
      '.sec-head-right { text-align:right; }',
      '.sec-head-right .sec-total-label { font-size:11px;color:#aaa;text-transform:uppercase;letter-spacing:.06em; }',
      '.sec-head-right .sec-total-val { font-size:20px;font-weight:700;font-variant-numeric:tabular-nums; }',
      '.field { margin-bottom:16px; }',
      '.field label { display:block;font-size:11px;font-weight:600;letter-spacing:.08em;text-transform:uppercase;color:#888;margin-bottom:6px; }',
      '.field input,.field select,.field textarea { width:100%;padding:12px 14px;background:#fff;border:1.5px solid var(--bd);border-radius:8px;color:var(--dk);font-size:14px;outline:none;transition:border-color .2s; }',
      '.field input:focus,.field select:focus,.field textarea:focus { border-color:var(--g); }',
      '.field select { appearance:none;-webkit-appearance:none;cursor:pointer;background-image:url("data:image/svg+xml,%3Csvg xmlns=\'http://www.w3.org/2000/svg\' width=\'12\' height=\'8\' fill=\'%23888\'%3E%3Cpath d=\'M6 8L0 0h12z\'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 14px center; }',
      '.grid-2 { display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:16px; }',
      '.grid-2-fixed { display:grid;grid-template-columns:auto 1fr;gap:16px;align-items:start; }',
      '.toggle-group { display:flex;border-radius:8px;overflow:hidden;border:1.5px solid var(--bd); }',
      '.toggle-group button { flex:1;padding:12px 20px;border:none;cursor:pointer;font-size:14px;font-weight:500;transition:all .2s;background:#fff;color:#888; }',
      '.toggle-group button.active { background:var(--g);color:#fff;font-weight:700; }',
      '.toggle-sm button { padding:10px 20px;font-size:13px; }',
      '.toggle-xs button { padding:8px 18px;font-size:14px;font-weight:600; }',
      '.info-box { padding:12px 16px;background:var(--gl);border-radius:8px;margin-top:8px; }',
      '.info-box .calc-line { font-size:13px;color:#555; }',
      '.info-box .calc-line code { font-family:monospace;font-size:13px;color:var(--gd);font-weight:700;background:#fff;padding:2px 8px;border-radius:4px; }',
      '.info-box .calc-detail { margin-top:6px;font-size:11px;color:#888;line-height:1.5; }',
      '.tbl-wrap { overflow-x:auto; }',
      'table.svc { width:100%;border-collapse:collapse;font-size:13px; }',
      'table.svc thead tr { border-bottom:2px solid var(--bd); }',
      'table.svc th { text-align:left;padding:10px 8px;color:#999;font-weight:600;font-size:10px;letter-spacing:.06em;text-transform:uppercase;white-space:nowrap; }',
      'table.svc th.r { text-align:right; } table.svc th.c { text-align:center; }',
      'table.svc tbody tr { transition:background .15s; }',
      'table.svc tbody tr:hover { background:var(--gy); }',
      'table.svc td { padding:8px;border-bottom:1px solid #eee;color:var(--dk); }',
      'table.svc td.grey { color:#999;font-size:12px;white-space:nowrap; }',
      'table.svc td.r { text-align:right; } table.svc td.c { text-align:center; }',
      'table.svc tr.split-row td { border-bottom:1px dashed #f0f0f0; }',
      '.am-badge { margin-left:6px;font-size:10px;font-weight:700;padding:1px 6px;border-radius:4px; }',
      '.variant-sel { font-size:11px;padding:3px 6px;border:1px solid var(--bd);border-radius:4px;color:#666;background:#fafafa;outline:none;max-width:320px;margin-top:4px;width:100%; }',
      '.price-cell { display:flex;align-items:center;gap:4px;justify-content:flex-end; }',
      '.price-btn { padding:5px 10px;border-radius:5px;font-size:12px;font-weight:600;cursor:pointer;white-space:nowrap;transition:all .15s;border:1.5px solid var(--bd);background:#fff;color:#aaa; }',
      '.price-btn.active-green { border:2px solid var(--g);background:var(--g);color:#fff; }',
      '.price-btn.active-blue { border:2px solid var(--b);background:var(--b);color:#fff; }',
      '.price-input-wrap { display:flex;align-items:center;gap:1px;border:1.5px solid var(--bd);border-radius:5px;padding:0 2px 0 6px;background:#fff;transition:border-color .15s; }',
      '.price-input-wrap.custom-green { border:2px solid var(--g); }',
      '.price-input-wrap.custom-blue { border:2px solid var(--b); }',
      '.price-input-wrap.custom-red { border:2px solid var(--rd); }',
      '.price-prefix { font-size:12px;font-weight:600;color:#ccc; }',
      '.price-prefix.green { color:var(--g); } .price-prefix.blue { color:var(--b); } .price-prefix.red { color:var(--rd); }',
      '.price-input-wrap input { width:62px;padding:5px 4px;border:none;outline:none;font-size:12px;font-weight:600;text-align:right;background:transparent;color:#ccc; }',
      '.price-input-wrap input.green { color:var(--g); } .price-input-wrap input.blue { color:var(--b); } .price-input-wrap input.red { color:var(--rd); }',
      '.qty-input { width:52px;padding:6px 8px;background:#fff;border:1.5px solid var(--bd);border-radius:6px;font-size:13px;text-align:center;outline:none; }',
      '.qty-input:focus { border-color:var(--g); }',
      '.add-split-btn { width:24px;height:24px;border-radius:6px;border:1.5px solid var(--bd);background:#fff;cursor:pointer;font-size:16px;font-weight:700;display:inline-flex;align-items:center;justify-content:center;transition:all .15s; }',
      '.add-split-btn:hover { background:var(--g);color:#fff;border-color:var(--g); }',
      '.rem-split-btn { width:24px;height:24px;border-radius:6px;border:1.5px solid var(--bd);background:#fff;color:var(--rd);cursor:pointer;font-size:14px;font-weight:700;display:inline-flex;align-items:center;justify-content:center;transition:all .15s; }',
      '.rem-split-btn:hover { background:var(--rd);color:#fff; }',
      '.split-label { font-size:12px;color:#aaa;font-style:italic; }',
      '.row-total-badge { font-size:10px;font-weight:700;margin-top:2px; }',
      '.search-wrap { display:flex;align-items:center;gap:8px; }',
      '.search-wrap input { flex:1;padding:12px 14px;background:#fff;border:1.5px solid var(--bd);border-radius:8px;color:var(--dk);font-size:14px;outline:none; }',
      '.search-wrap input:focus { border-color:var(--g); }',
      '#search-results { border-top:1px solid var(--bd);padding-top:8px;display:none; }',
      '.search-result { display:flex;justify-content:space-between;align-items:center;width:100%;padding:8px 12px;background:transparent;border:none;cursor:pointer;border-radius:6px;font-size:13px;text-align:left; }',
      '.search-result:hover { background:var(--gy); }',
      '.search-badge { font-size:10px;font-weight:600;letter-spacing:.06em;text-transform:uppercase;padding:2px 8px;border-radius:4px; }',
      '.pkg-grid { display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));gap:16px;margin-top:8px; }',
      '.pkg-card { border:1.5px solid var(--bd);border-radius:12px;padding:20px;cursor:pointer;transition:all .2s;background:#fff;position:relative;overflow:hidden; }',
      '.pkg-top-bar { position:absolute;top:0;left:0;right:0;height:3px;display:none; }',
      '.pkg-card.selected .pkg-top-bar { display:block; }',
      '.pkg-tier { font-size:10px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;margin-bottom:4px; }',
      '.pkg-price { font-size:28px;font-weight:700;font-family:var(--serif);color:var(--dk); }',
      '.pkg-price span { font-size:13px;font-weight:500;color:#999;font-family:var(--font); }',
      '.pkg-annual { font-size:11px;color:#aaa;margin-bottom:12px; }',
      '.pkg-target { font-size:12px;font-weight:600;margin-bottom:4px; }',
      '.pkg-tagline { font-size:11px;color:#888;font-style:italic;margin-bottom:12px; }',
      '.pkg-section-lbl { font-size:10px;font-weight:600;color:#888;text-transform:uppercase;letter-spacing:.06em;margin-bottom:6px; }',
      '.pkg-feature { display:flex;gap:6px;margin-bottom:4px;font-size:12px;color:#555; }',
      '.pkg-excl { font-size:11px;color:#bbb;margin-bottom:2px; }',
      '.pkg-select-btn { text-align:center;padding:8px 16px;border-radius:8px;font-size:13px;font-weight:600;transition:all .2s;margin-top:14px;border:1.5px solid;cursor:pointer;width:100%; }',
      '.pkg-delivery { margin-top:8px;font-size:10px;color:#aaa;text-align:center; }',
      '#grand-total-bar { background:var(--dk);border-radius:12px;padding:28px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:16px; }',
      '#gt-label { font-size:11px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:rgba(255,255,255,.45); }',
      '#gt-strikethrough { font-size:14px;color:rgba(255,255,255,.35);text-decoration:line-through;margin-top:4px;display:none; }',
      '#gt-value { font-size:36px;font-weight:700;color:#fff;font-family:var(--serif);font-variant-numeric:tabular-nums;line-height:1.1;margin-top:4px; }',
      '#gt-value span { font-size:16px;font-weight:500;opacity:.4;font-family:var(--font);margin-left:6px; }',
      '#gt-monthly { font-size:13px;color:rgba(255,255,255,.4);margin-top:4px; }',
      '#reset-btn { padding:10px 24px;background:var(--g);color:#fff;border:none;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;transition:background .2s; }',
      '#reset-btn:hover { background:var(--gd); }',
      '#breakdown { margin-top:20px;padding:16px 20px;background:#fff;border-radius:10px;border:1px solid var(--bd);box-shadow:0 1px 4px rgba(0,0,0,.04); }',
      '#breakdown .bd-title { font-size:11px;font-weight:600;letter-spacing:.08em;text-transform:uppercase;color:#aaa;margin-bottom:12px; }',
      '.bd-row { display:flex;justify-content:space-between;align-items:center;padding:8px 0;border-bottom:1px solid #f0f0f0;font-size:13px; }',
      '.bd-row-left { display:flex;align-items:center;gap:8px; }',
      '.bd-dot { width:8px;height:8px;border-radius:50%; }',
      '.bd-total-row { display:flex;justify-content:space-between;padding:12px 0 0;font-size:16px;font-weight:700; }',
      '.disc-applied { padding:10px 16px;background:#fff5f5;border-radius:8px;border:1px solid #ffdddd;display:flex;align-items:center;gap:12px;flex-wrap:wrap; }',
      '#email-textarea { width:100%;min-height:300px;padding:16px;background:var(--gy);border:1.5px solid var(--bd);border-radius:8px;color:var(--dk);font-size:13px;line-height:1.7;outline:none;resize:vertical; }',
      '#email-textarea:focus { border-color:var(--g); }',
      '.email-btn-row { display:flex;gap:10px;flex-wrap:wrap;margin-bottom:16px; }',
      '.btn-green { padding:12px 24px;background:var(--g);color:#fff;border:none;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;display:flex;align-items:center;gap:8px;transition:background .2s; }',
      '.btn-green:hover { background:var(--gd); } .btn-green:disabled { background:#aaa;cursor:default; }',
      '.btn-blue { padding:12px 24px;background:#0078d4;color:#fff;border:none;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;display:flex;align-items:center;gap:8px;transition:opacity .2s; }',
      '.btn-blue:hover { opacity:.88; }',
      '.btn-outline { padding:6px 14px;background:#fff;border:1px solid var(--bd);border-radius:6px;font-size:11px;font-weight:600;color:#666;cursor:pointer;display:flex;align-items:center;gap:4px; }',
      '.email-actions-overlay { position:absolute;top:10px;right:10px;display:flex;gap:6px; }',
      '.email-wrap { position:relative; }',
      '@keyframes spin { to { transform:rotate(360deg); } }',
      '.spin { animation:spin 1s linear infinite; }',
      '#footer { text-align:center;font-size:11px;color:#bbb;margin-top:24px; }'
    ].join('\n');
    document.head.appendChild(st);
  }

  // ── HTML ──────────────────────────────────────────────────────────
  domElement.innerHTML =
    '<div id="header"><div id="header-inner">' +
    '<div class="logo"><svg height="36" viewBox="0 0 478.89 199.93" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" style="height:36px;width:auto;flex-shrink:0"><defs><style>.lc1{fill:none}.lc14{fill:#fff}</style><clipPath id="lcp1"><path class="lc1" d="M138.28,135.11l-1.39.57-9.36,3.92-85.75,36,49.51,20.82a23.56,23.56,0,0,0,17.79.17l53-21.58-15.51-37.08a5.09,5.09,0,0,0-5-3.39,10.29,10.29,0,0,0-3.2.59"/></clipPath><linearGradient id="lcg1" x1="-842.62" y1="247.92" x2="-840.21" y2="247.92" gradientTransform="translate(-173.97 18088.39) rotate(107.32) scale(20.46)" gradientUnits="userSpaceOnUse"><stop offset="0" stop-color="#7a7b7d"/><stop offset=".23" stop-color="#606062"/><stop offset=".57" stop-color="#3f3f40"/><stop offset=".75" stop-color="#323233"/><stop offset="1" stop-color="#323233"/></linearGradient><clipPath id="lcp2"><path class="lc1" d="M53.24,135.12l18.31,53,20.14,8.42a23,23,0,0,0,17.58.06L162,175l-61.22-24.5c-.75-.2-30.65-12.73-30.65-12.73h0a13.9,13.9,0,0,1-6.82-6.81l-.25-.62Z"/></clipPath><linearGradient id="lcg2" x1="-884.08" y1="251.1" x2="-881.67" y2="251.1" gradientTransform="translate(-23853.6 15558.77) rotate(163.15) scale(31.01)" xlink:href="#lcg1"/><clipPath id="lcp3"><path class="lc1" d="M18,141.62l8.29,20.19a23,23,0,0,0,12.39,12.47l52.57,22.11L63.9,132.13a9.12,9.12,0,0,1-.58-1.12h0l-2.88-7-7.83-19.08a13.9,13.9,0,0,1,0-9.64l.26-.61L42.56,91.17Z"/></clipPath><linearGradient id="lcg3" x1="-904.5" y1="265.3" x2="-902.09" y2="265.3" gradientTransform="translate(-28548.1 -5824.42) rotate(-151.85) scale(31.03)" xlink:href="#lcg1"/><clipPath id="lcp4"><path class="lc1" d="M13,71,4.56,91.12a23,23,0,0,0-.06,17.59l21.64,52.72,26-64.69a11.09,11.09,0,0,1,.38-1.2h0l2.93-7,7.95-19a14,14,0,0,1,6.82-6.82l.62-.25L66,52.7Z"/></clipPath><linearGradient id="lcg4" x1="-908.99" y1="289.58" x2="-906.58" y2="289.58" gradientTransform="translate(-16723.25 -24245.58) rotate(-106.85) scale(31.01)" xlink:href="#lcg1"/><clipPath id="lcp5"><path class="lc1" d="M39.14,25.79A23,23,0,0,0,26.67,38.18l-22,52.57L68.78,63.4c.36-.21.73-.4,1.11-.58h0l7-2.87L96,52.12a14,14,0,0,1,9.64,0l.61.26,3.56-10.32L59.31,17.51Z"/></clipPath><linearGradient id="lcg5" x1="-894.88" y1="310.2" x2="-892.47" y2="310.2" gradientTransform="translate(4655.7 -28926.32) rotate(-61.85) scale(31.01)" xlink:href="#lcg1"/><clipPath id="lcp6"><path class="lc1" d="M100.88,2.28A23.21,23.21,0,0,0,92.29,4L41.66,24.77l86.71,36.82,9.41,3.92,1.38.58c2.2.76,6.41,1.53,8.22-2.8l15.5-37.08-53-22.15a23.21,23.21,0,0,0-8.74-1.78Z"/></clipPath><linearGradient id="lcg6" x1="-885.43" y1="278.19" x2="-883.02" y2="278.19" gradientTransform="matrix(-90.39,-39.99,39.99,-90.39,-91000.33,-10203.04)" xlink:href="#lcg1"/></defs><g clip-path="url(#lcp1)"><polygon fill="url(#lcg1)" points="180.3 140.22 151.34 233.09 23.44 193.2 52.4 100.34 180.3 140.22"/></g><g clip-path="url(#lcp2)"><polygon fill="url(#lcg2)" points="181.05 193.17 62.39 229.13 34.19 136.08 152.85 100.12 181.05 193.17"/></g><g clip-path="url(#lcp3)"><polygon fill="url(#lcg3)" points="74.96 226.86 -25.77 172.98 34.29 60.7 135.01 114.57 74.96 226.86"/></g><g clip-path="url(#lcp4)"><polygon fill="url(#lcg4)" points="7.96 180.48 -28 61.82 65.05 33.62 101.01 152.28 7.96 180.48"/></g><g clip-path="url(#lcp5)"><polygon fill="url(#lcg5)" points="-25.75 74.45 28.1 -26.22 140.29 33.79 86.44 134.46 -25.75 74.45"/></g><g clip-path="url(#lcp6)"><polygon fill="url(#lcg6)" points="143 112.45 17.49 56.91 61.5 -42.56 187.01 12.97 143 112.45"/></g><path class="lc14" d="M233.08,140.48a101.1,101.1,0,0,1-11,.49c-23.32,0-27.86-11.42-27.86-32.81,0-11,.77-26.51,12.67-31.54,4.65-1.93,10.26-2.22,15.19-2.22a106.39,106.39,0,0,1,11,.48c2.42.19,5,.39,5,3.39v3.19c0,3-.29,4.64-3.68,4.64s-7-.48-10.45-.48c-2.9,0-7.93.19-10.55,1.35-6.77,2.91-5.9,15.29-5.9,21.29,0,7.55-.58,19.16,9,21a38.72,38.72,0,0,0,7.45.48c3.48,0,7-.48,10.45-.48,4.74,0,3.68,4.16,3.68,7.83,0,3-2.62,3.2-5,3.39"/><path class="lc14" d="M281.86,139.71h-5.42c-2.23,0-2.81-1.45-2.81-3.48-4.54,2.8-8.9,4.74-14.32,4.74-9.48,0-13.84-5.13-13.84-14.42,0-10.55,6.3-14.12,15.78-15.19a87.05,87.05,0,0,1,12.29-.77c0-3.87.58-9.29-4.84-9.29a61.75,61.75,0,0,0-13.92,2,9.11,9.11,0,0,1-2.32.29c-3.38,0-2.8-3.39-2.8-5.81,0-2.13.58-3.38,2.71-4.06a67.73,67.73,0,0,1,15.86-1.84c4.65,0,10.55.77,13.84,4.45,3.87,4.26,3.68,10.26,3.68,15.68V135.5c0,2.81-1,4.16-3.87,4.16m-8.25-20.38c-1.55-.09-3.19-.19-4.74-.19a44.28,44.28,0,0,0-6.09.29c-4.07.58-5.71,1.94-5.71,6.1,0,2.32.38,6.48,3.58,6.48,4.83,0,9.09-2,13-4.84Z"/><path class="lc14" d="M322.14,103.33c-5.61,0-8.42.87-13.07,3.77v28.35c0,2.81-1,4.26-3.87,4.26h-3.67c-3.39,0-5.13-.39-5.13-4.26V97.52c0-3,1.26-4.16,4.26-4.16h4.54c2.23,0,3.1,1.35,3.58,3.39,3.87-2.32,8.13-4.45,12.68-4.45,4.74,0,4.26,2.9,4.26,6.58,0,2.61-.58,4.45-3.58,4.45"/><path class="lc14" d="M353,141a91.23,91.23,0,0,1-15.48-1.16c-2.61-.49-4.64-1.26-4.64-4.36V77.28c0-3.1,1.25-4.16,4.25-4.16h4.36c2.61,0,3.77,1.74,3.77,4.16v15.6a84.66,84.66,0,0,1,11.9-1c15.19,0,17.42,10.93,17.42,23.9,0,16-4.26,25.16-21.58,25.16m8.15-34.9c-.87-3.58-2.42-3.87-5.9-3.87a50.22,50.22,0,0,0-10,1.26v26.51a54.32,54.32,0,0,0,7.94.87c3.19,0,6.67-.58,7.74-4.07.87-2.9,1.06-8,1.06-11.12a40.08,40.08,0,0,0-.87-9.58"/><path class="lc14" d="M405.14,141c-16.16,0-22.26-9.29-22.26-24.48s6-24.58,22.26-24.58,22.15,9.29,22.15,24.58S421.39,141,405.14,141m7.9-35.69c-1.74-3.19-4.65-3.09-7.93-3.09s-6.2-.1-7.94,3.09-1.84,7.84-1.84,11.23.2,8.22,1.84,11.22,4.74,3,7.94,3,6.19.1,7.93-3,1.74-7.93,1.74-11.22-.1-8.23-1.74-11.23"/><path class="lc14" d="M469.88,139.71c-3.38,0-5.12-.39-5.12-4.26V110.39c0-1.55.09-5.51-.68-6.87a3.12,3.12,0,0,0-3-1.45c-4.74,0-8.32,2.51-11.9,5.32v28.06c0,5.13-3.29,4.26-7.26,4.26-3.39,0-5.13-.39-5.13-4.26V97.52c0-3,1.26-4.16,4.26-4.16h4.36c2.12,0,2.8,1.07,3.28,3,4.26-2.81,8.42-4.45,13.55-4.45,8.8,0,14.9,4.26,14.9,13.45v30.09c0,5-3.38,4.26-7.26,4.26"/></svg></div>' +
    '<span class="header-tag">Partner Pricing Tool</span>' +
    '<span id="sp-status-chip">Loading prices…</span>' +
    '</div></div>' +
    '<div id="hero"><div id="hero-inner">' +
    '<h1>Services Calculator</h1>' +
    '<p>Estimate annual compliance fees for your client engagement</p>' +
    '</div></div>' +
    '<div id="main"></div>';

  // ══════════════════════════════════════════════════════
  //  DATA
  // ══════════════════════════════════════════════════════
  const TIERS = [
    {lower:0,fee:1500},{lower:200000,fee:2000},{lower:400000,fee:2500},
    {lower:600000,fee:3000},{lower:800000,fee:4000},{lower:2000000,fee:8000},
    {lower:5000000,fee:12000}
  ];
  const QUALITY = [
    {label:'Cloud Software (Xero/QBO)',m:1},
    {label:'Excel Spreadsheets',m:1.2},
    {label:'Shoebox / PDF Statements',m:1.4}
  ];
  const TIDY = [
    {label:'Reconciled / Clean',m:1},
    {label:'Minor Clean-up Required',m:1.2},
    {label:'Unreconciled Bank',m:1.3}
  ];
  const ENT_OPTS = [1,2,3,4,5,6,7,8,9,10];
  const sa = (a) => [...a].sort((x,y) => x.name.localeCompare(y.name));

  const YEAR_END = sa([
    {name:'Additional Time – Delay due to Client',unit:'Per Month',good:100,bad:300},
    {name:'BAS Amendment',unit:'Per BAS',good:150,bad:450},
    {name:'Bank Account not Reconciled',unit:'Per Account',good:200,bad:600},
    {name:'CGT – per Property',unit:'Per Property',good:300,bad:1350},
    {name:'CGT – per Share/Unit Holding',unit:'Per Holding',good:150,bad:300},
    {name:'Div 7A Calculation / Loan Documents',unit:'Per Calc/Doc',good:200,bad:950},
    {name:'Dividends Declared / Trust Resolution',unit:'Per Entity',good:150,bad:300},
    {name:'FBT Return',unit:'Per Return',good:300,bad:800},
    {name:'Foreign Currency Translation',unit:'Per Calc/Doc',good:500,bad:950},
    {name:'Payroll not Reconciled',unit:'Per Reconciliation',good:300,bad:1500},
    {name:'Services Agreement / Hire Agreement',unit:'Per Calc/Doc',good:150,bad:1500}
  ]);
  const OTHER = sa([
    {name:'ASIC Secretarial',unit:'Per Company / Lodgement',good:250,bad:350},
    {name:'ATO Payment Plan',unit:'Per Entity',good:100,bad:100},
    {name:'Business Structure Advice + Set Up',unit:'Per Entity',good:500,bad:2000},
    {name:'Monthly / Quarterly Review Meeting',unit:'Per Meeting',good:300,bad:1200},
    {name:'Tax Planning',unit:'Per Entity',good:500,bad:1500},
    {name:'Taxable Payments Annual Report (TPAR)',unit:'Per Year',good:100,bad:200},
    {name:'Xero Setup / Onboarding / Clean Up',unit:'Per File',good:300,bad:700}
  ]);
  const INDIV = sa([
    {name:'Base Rate – Tax Return Preparation',unit:'Per Tax Return',good:250,bad:300},
    {name:'CGT',unit:'Per CGT Event',good:150,bad:300},
    {name:'Depreciation Schedule',unit:'Per Schedule',good:100,bad:150},
    {name:'Dividends / Investment Income',unit:'Per Schedule',good:50,bad:150},
    {name:'Motor Vehicle Schedule',unit:'Per Vehicle',good:50,bad:150},
    {name:'Rental Property Schedule',unit:'Per Property',good:200,bad:500},
    {name:'Sole Trader Schedule (P&L Not Provided)',unit:'Per Schedule',good:1000,bad:2500},
    {name:'Sole Trader Schedule (P&L Provided)',unit:'Per Schedule',good:500,bad:1000}
  ]);
  const BOOK = sa([
    {name:'100 Transactions/month ($220/mo)',unit:'Per BAS',good:220,bad:400},
    {name:'BAS Lodgement (No Bookkeeping)',unit:'Per BAS',good:220,bad:550},
    {name:'IAS Lodgement (PAYGW only)',unit:'Per IAS',good:150,bad:300},
    {name:'Payroll Management',unit:'Per Employee/Month',good:50,bad:80,am:12},
    {name:'Per 100 Transactions Thereafter/mo',unit:'Per BAS',good:200,bad:200}
  ]);
  const NE_COMPANY = [
    {name:'New Trading Company',unit:'Per Company',good:1400},
    {name:'New Trustee Company for Existing Trust',unit:'Per Company',good:1100},
    {name:'Constitution Upgrade',unit:'Per Upgrade',good:350,variants:[{label:'Trustee Company',price:350},{label:'Trading Company',price:450}]},
    {name:'Deregistration of Company',unit:'Per Company',good:540},
    {name:'Change of Company Name',unit:'Per Change',good:810},
    {name:'New Company with Trust Owning Shares',unit:'Per Setup',good:2100}
  ];
  const NE_TRUST = [
    {name:'New Discretionary Trust',unit:'Per Trust',good:900,variants:[{label:'Standard',price:900},{label:'With Corporate Trustee',price:2100}]},
    {name:'New Unit Trust',unit:'Per Trust',good:1200,variants:[{label:'Standard',price:1200},{label:'With Corporate Trustee',price:2100}]},
    {name:'Change of / to Individual Trustee',unit:'Per Change',good:570},
    {name:'Change Trustee',unit:'Per Change',good:570,variants:[{label:'To Existing Company',price:570},{label:'To New Corporate Trustee',price:1600}]}
  ];
  const NE_SMSF = [
    {name:'New SMSF',unit:'Per Fund',good:900,variants:[{label:'With Individual Trustees',price:900},{label:'With Corporate Trustee',price:2000},{label:'With Constitution Upgrade (existing co)',price:1230},{label:'With Corp Trustee & Bare Trust w/ Corp Trustee',price:3600}]},
    {name:'Change of Individual Trustee',unit:'Per Change',good:550,variants:[{label:'Without Deed Upgrade',price:550},{label:'With Deed Upgrade',price:1000}]},
    {name:'Change Trustee to Existing Company',unit:'Per Change',good:550,variants:[{label:'Without Deed Upgrade & without co type change',price:550},{label:'Without Deed Upgrade, with co type change',price:1100},{label:'With Deed Upgrade (without co type change)',price:1100},{label:'With Deed Upgrade (with co type change to SPC)',price:1650}]},
    {name:'Change Trustee to a New Company',unit:'Per Change',good:1490,variants:[{label:'Without Deed Upgrade',price:1490},{label:'With Deed Upgrade',price:1900}]},
    {name:'Deed Upgrade',unit:'Per Upgrade',good:900},
    {name:'Replacement Deed',unit:'Per Deed',good:700},
    {name:'LRBA Trust',unit:'Per Trust',good:750,variants:[{label:'Standard',price:750},{label:'With Corporate Trustee',price:2100}]}
  ];
  const NE_OTHER = [
    {name:'ABN & TFN Registration / Reactivation',unit:'Per Registration',good:180,variants:[{label:'Standard (ABN & TFN)',price:180},{label:'With GST, FTC, PAYG Withholding',price:220}]},
    {name:'GST, FTC & PAYG Withholding Registration / Reactivation',unit:'Per Registration',good:70},
    {name:'Business Name Registration / Renewal',unit:'Per Registration',good:190,variants:[{label:'One Business Name (incl. ASIC fee: $102/3yr or $44/1yr)',price:190},{label:'Two Business Names (incl. ASIC fee: $102/3yr or $44/1yr)',price:210}]}
  ];
  const NE_ANNUAL = [
    {name:'Corporate Secretarial Maintenance',unit:'Per Company',good:360,variants:[{label:'First Company',price:360},{label:'Additional Companies',price:300},{label:'Elston Clients & SMSF Corporate Trustee',price:250}]},
    {name:'Share Transfers, Unit Allotment & Dividend Statement',unit:'Per Service',good:180}
  ];
  let PACKAGES = [
    {id:'core',name:'Core',monthly:220,annual:2640,color:'#6c757d',target:'Micro / Solopreneur',tagline:'Keep me compliant and stress-free.',features:['Annual financial statements','Income tax return prep & lodgement','Quarterly BAS lodgement','ASIC annual company review','Basic Xero/QBO support','Email & phone support'],excluded:['Tax planning','Bookkeeping','FBT','Complex CGT']},
    {id:'growth',name:'Growth',monthly:440,annual:5280,color:'#4db749',target:'Established SME',tagline:'Help me understand my numbers and save tax.',features:['Everything in Core','Tax planning session (annual)','Quarterly BAS review & advice','1x annual strategy call','Dividend / trust resolutions','Up to 250 transactions/month'],excluded:['Bookkeeping (daily)','Complex structuring','R&D tax incentive']},
    {id:'scale',name:'Scale',monthly:770,annual:9240,color:'#29b6c9',target:'Scaling SME',tagline:'Act as my sounding board for growth.',features:['Everything in Growth','Bi-monthly management reporting','Virtual CFO meetings (bi-monthly)','Budgets & cash flow forecasting','Dedicated account manager','Payroll for up to 10 employees','ASIC secretarial & compliance'],excluded:['Daily bookkeeping','Payroll processing (>10)']},
    {id:'enterprise',name:'Enterprise',monthly:1320,annual:15840,color:'#6f42c1',target:'Large / Complex Groups',tagline:'A full finance function, outsourced to experts.',features:['Everything in Scale','Monthly CFO advisory meetings','Unlimited transactions','Unlimited payroll','Custom management reporting','Priority response SLA','Multi-entity consolidation','Board pack preparation'],excluded:['Audit services','R&D tax incentive']}
  ];

  let YEAR_END_LIVE=YEAR_END, OTHER_LIVE=OTHER, INDIV_LIVE=INDIV, BOOK_LIVE=BOOK;
  let NE_COMPANY_LIVE=NE_COMPANY, NE_TRUST_LIVE=NE_TRUST, NE_SMSF_LIVE=NE_SMSF, NE_OTHER_LIVE=NE_OTHER, NE_ANNUAL_LIVE=NE_ANNUAL;

  // ══════════════════════════════════════════════════════
  //  STATE
  // ══════════════════════════════════════════════════════
  function initRows(items) { return items.map(item => ({splits:[{useGood:true,cp:String(item.bad||item.good),qty:0,variantIdx:0}]})); }

  const state = {
    cName:'',cEmail:'',cType:0,turnover:'',qIdx:0,tIdx:0,eIdx:0,
    yeS:initRows(YEAR_END),otS:initRows(OTHER),inS:initRows(INDIV),bkS:initRows(BOOK),
    neCoS:initRows(NE_COMPANY),neTrS:initRows(NE_TRUST),neSmS:initRows(NE_SMSF),neOtS:initRows(NE_OTHER),neAnS:initRows(NE_ANNUAL),
    viewMode:'services',selPkg:null,discType:'percent',discVal:'',discReason:'',
    emailDraft:'',emailLoading:false,hlSec:null,hlIdx:null,_searchQ:'',
    expanded:{yearEnd:true,other:true,book:true,indiv:false,newEntities:false,neCompany:false,neTrust:false,neSmsf:false,neOther:false,neAnnual:false,discount:false}
  };

  // ══════════════════════════════════════════════════════
  //  HELPERS
  // ══════════════════════════════════════════════════════
  const fmt = (v) => '$' + Math.round(v).toLocaleString('en-AU');
  function escHtml(str) { return String(str||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
  function lookupFee(t) { const n=parseFloat(String(t).replace(/[^0-9.]/g,''))||0; let f=TIERS[0].fee; for(const tier of TIERS) if(n>=tier.lower) f=tier.fee; return f; }
  function getVariantPrice(item,sp) { if(item.variants&&item.variants[sp.variantIdx]) return item.variants[sp.variantIdx].price; return item.good; }
  function rowTotal(item,row) { const am=item.am||1; return row.splits.reduce((s,sp)=>{ const bp=getVariantPrice(item,sp); const p=sp.useGood?bp:(parseFloat(sp.cp)||0); return s+p*(sp.qty||0)*am; },0); }
  function secTotal(items,rows) { return items.reduce((s,item,i)=>s+rowTotal(item,rows[i]),0); }
  function hasAct(items,rows) { return rows.some(row=>row.splits.some(sp=>(sp.qty||0)>0)); }

  function calcDerived() {
    const isCorp=state.cType===0, ent=ENT_OPTS[state.eIdx], baseFee=lookupFee(state.turnover);
    const qM=QUALITY[state.qIdx].m, tM=TIDY[state.tIdx].m;
    const baseTotal=Math.round(baseFee*qM*tM*ent);
    const calcStr=fmt(baseFee)+' \xd7 '+qM+' \xd7 '+tM+' \xd7 '+ent;
    const yeT=secTotal(YEAR_END_LIVE,state.yeS), otT=secTotal(OTHER_LIVE,state.otS);
    const inT=secTotal(INDIV_LIVE,state.inS), bkT=secTotal(BOOK_LIVE,state.bkS);
    const neCoT=secTotal(NE_COMPANY_LIVE,state.neCoS), neTrT=secTotal(NE_TRUST_LIVE,state.neTrS);
    const neSmT=secTotal(NE_SMSF_LIVE,state.neSmS), neOtT=secTotal(NE_OTHER_LIVE,state.neOtS);
    const neAnT=secTotal(NE_ANNUAL_LIVE,state.neAnS);
    const neTotal=neCoT+neTrT+neSmT+neOtT+neAnT;
    const pkgObj=PACKAGES.find(p=>p.id===state.selPkg);
    const usePkg=isCorp&&state.viewMode==='packages'&&pkgObj;
    const subtotal=usePkg?baseTotal+pkgObj.annual+neTotal+inT:(isCorp?baseTotal:0)+(yeT+otT+bkT)+inT+neTotal;
    const dv=parseFloat(state.discVal)||0;
    const discAmt=state.discType==='percent'?Math.round(subtotal*dv/100):Math.round(dv);
    const discLabel=state.discType==='percent'?dv+'%':fmt(dv);
    const grandTotal=Math.max(0,subtotal-discAmt), monthly=Math.round(grandTotal/12);
    return {isCorp,ent,baseFee,qM,tM,baseTotal,calcStr,yeT,otT,inT,bkT,neCoT,neTrT,neSmT,neOtT,neAnT,neTotal,pkgObj,usePkg,subtotal,dv,discAmt,discLabel,grandTotal,monthly};
  }

  // ══════════════════════════════════════════════════════
  //  RENDER
  // ══════════════════════════════════════════════════════
  function render() {
    const main=document.getElementById('main');
    if(!main) return;
    const scrollY=window.scrollY||window.pageYOffset||0;
    const aId=document.activeElement&&document.activeElement.id;
    // selectionStart throws for type="email"/"number" in Chrome — must be wrapped
    let aSS=null, aSE=null;
    try { if(aId){ aSS=document.activeElement.selectionStart; aSE=document.activeElement.selectionEnd; } } catch(e){}
    const d=calcDerived();
    main.innerHTML=buildHTML(d);
    attachEvents(d);
    window.scrollTo(0,scrollY);
    if(aId){
      const el=document.getElementById(aId);
      if(el){
        el.focus({preventScroll:true});
        // Only restore cursor for inputs that support selectionRange (i.e. type="text")
        if(aSS!==null){try{el.setSelectionRange(aSS,aSE);}catch(e){}}
      }
    }
  }

  function fld(label,type,id,val,ph) { return '<div class="field"><label>'+label+'</label><input dir="ltr" type="'+type+'" id="f-'+id+'" value="'+escHtml(val)+'" placeholder="'+ph+'"></div>'; }
  function selFld(label,id,val,opts) { return '<div class="field"><label>'+label+'</label><select dir="ltr" id="f-'+id+'">'+opts.map((o,i)=>'<option value="'+i+'"'+(i===val?' selected':'')+'>'+escHtml(o)+'</option>').join('')+'</select></div>'; }

  function buildHTML(d) {
    const {isCorp,ent,baseFee,qM,tM,baseTotal,calcStr,yeT,otT,inT,bkT,neTotal,pkgObj,usePkg,subtotal,discAmt,discLabel,grandTotal,monthly}=d;
    const hasYe=hasAct(YEAR_END_LIVE,state.yeS),hasOt=hasAct(OTHER_LIVE,state.otS),hasBk=hasAct(BOOK_LIVE,state.bkS),hasIn=hasAct(INDIV_LIVE,state.inS);
    const hasNe=hasAct(NE_COMPANY_LIVE,state.neCoS)||hasAct(NE_TRUST_LIVE,state.neTrS)||hasAct(NE_SMSF_LIVE,state.neSmS)||hasAct(NE_OTHER_LIVE,state.neOtS)||hasAct(NE_ANNUAL_LIVE,state.neAnS);
    const corpOrder=[
      {key:'yearEnd',num:'02',title:'Year End Compliance Add-Ons',items:YEAR_END_LIVE,rows:state.yeS,total:yeT,accent:'g',sub:null,hasAct:hasYe},
      {key:'other',num:'03',title:'Other Services',items:OTHER_LIVE,rows:state.otS,total:otT,accent:'g',sub:'Corporate services',hasAct:hasOt},
      {key:'book',num:'04',title:'Bookkeeping',items:BOOK_LIVE,rows:state.bkS,total:bkT,accent:'b',sub:null,hasAct:hasBk},
      {key:'indiv',num:'05',title:'Individuals',items:INDIV_LIVE,rows:state.inS,total:inT,accent:'g',sub:'Individual tax return services (also available for corporate clients)',hasAct:hasIn}
    ];
    const indivOrder=[
      {key:'indiv',num:'01',title:'Individuals',items:INDIV_LIVE,rows:state.inS,total:inT,accent:'g',sub:'Individual tax return services',hasAct:hasIn},
      {key:'yearEnd',num:'02',title:'Year End Compliance Add-Ons',items:YEAR_END_LIVE,rows:state.yeS,total:yeT,accent:'g',sub:'Available for corporate engagements',hasAct:hasYe},
      {key:'other',num:'03',title:'Other Services',items:OTHER_LIVE,rows:state.otS,total:otT,accent:'g',sub:'Available for corporate engagements',hasAct:hasOt},
      {key:'book',num:'04',title:'Bookkeeping',items:BOOK_LIVE,rows:state.bkS,total:bkT,accent:'b',sub:'Available for corporate engagements',hasAct:hasBk}
    ];
    const sectionOrder=isCorp?corpOrder:indivOrder;
    const showYe=state.expanded.yearEnd||hasYe,showOt=state.expanded.other||hasOt,showBk=state.expanded.book||hasBk,showIn=state.expanded.indiv||hasIn,showNe=state.expanded.newEntities||hasNe;
    const bdRows=[];
    if(usePkg){if(isCorp)bdRows.push({label:'Base Compliance Fee',value:baseTotal,c:'#4db749'});bdRows.push({label:pkgObj.name+' Package',value:pkgObj.annual,c:pkgObj.color});}
    else{if(isCorp)bdRows.push({label:'Base Compliance Fee',value:baseTotal,c:'#4db749'});if(showYe)bdRows.push({label:'Year End Add-Ons',value:yeT,c:'#4db749'});if(showOt)bdRows.push({label:'Other Services',value:otT,c:'#4db749'});if(showBk)bdRows.push({label:'Bookkeeping',value:bkT,c:'#29b6c9'});if(showIn)bdRows.push({label:'Individual Services',value:inT,c:'#4db749'});}
    if(showNe)bdRows.push({label:'New Entities & Registrations',value:neTotal,c:'#4db749'});
    let h='';
    h+='<div class="card card-green"><div class="sec-head" style="border-bottom:2.5px solid #4db749"><div class="sec-head-left"><h3>Client Type</h3></div></div><div class="toggle-group"><button id="btn-corp" class="'+(isCorp?'active':'')+'">Corporate</button><button id="btn-indiv" class="'+(!isCorp?'active':'')+'">Individual</button></div></div>';
    if(isCorp){h+='<div class="card card-green"><div class="sec-head" style="border-bottom:2.5px solid #4db749"><div class="sec-head-left"><span class="sec-num" style="color:#4db749">01</span><h3>Base Annual Compliance Fee</h3><p>Turnover, record quality, tidiness, and entity count</p></div><div class="sec-head-right"><div class="sec-total-label">Section Total</div><div class="sec-total-val" style="color:#4db749">'+fmt(baseTotal)+'</div></div></div><div class="grid-2">'+fld('Client Annual Turnover ($)','number','turnover',state.turnover,'e.g. 514192')+selFld('Record Quality','qIdx',state.qIdx,QUALITY.map(x=>x.label))+selFld('Tidiness Level','tIdx',state.tIdx,TIDY.map(x=>x.label))+selFld('Number of Entities','eIdx',state.eIdx,ENT_OPTS.map(n=>n+' entit'+(n>1?'ies':'y')))+'</div><div class="info-box"><div class="calc-line"><strong>Calculation:</strong> <code>'+calcStr+' = '+fmt(baseTotal)+'</code></div><div class="calc-detail">Base '+fmt(baseFee)+' ('+(state.turnover?'$'+Number(state.turnover).toLocaleString():'$0')+' turnover) \xb7 Quality \xd7'+qM+' ('+QUALITY[state.qIdx].label+') \xb7 Tidiness \xd7'+tM+' ('+TIDY[state.tIdx].label+') \xb7 '+ent+' entit'+(ent>1?'ies':'y')+'</div></div></div>';}
    h+='<div class="card card-green">';
    if(isCorp){h+='<div class="toggle-group toggle-sm" style="margin-bottom:16px"><button id="btn-services" class="'+(state.viewMode==='services'?'active':'')+'">Services</button><button id="btn-packages" class="'+(state.viewMode==='packages'?'active':'')+'">Packages</button></div>';}
    if(state.viewMode==='services'){h+='<div class="search-wrap"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#aaa" stroke-width="2" stroke-linecap="round"><circle cx="11" cy="11" r="8"/><path d="M21 21l-4.35-4.35"/></svg><input dir="ltr" type="text" id="search-input" placeholder="Search for a service…" value="'+escHtml(state._searchQ||'')+'"></div><div id="search-results"></div>';}
    else{h+=buildPackagesHTML();}
    h+='</div>';
    if(state.viewMode==='services'){for(const sec of sectionOrder){h+=buildColl(sec.key,sec.num,sec.title,sec.sub,sec.total,sec.accent,sec.items,sec.rows,sec.hasAct,false);}}
    h+=buildNeColl(d);h+=buildDiscColl(d);
    h+='<div id="grand-total-bar"><div><div id="gt-label">Estimated Annual Fee</div><div id="gt-strikethrough" style="display:'+(discAmt>0?'block':'none')+'">'+fmt(subtotal)+'</div><div id="gt-value">'+fmt(grandTotal)+'<span>/ year</span></div><div id="gt-monthly">~'+fmt(monthly)+' / month</div></div><button id="reset-btn">Reset Calculator</button></div>';
    h+='<div id="breakdown"><div class="bd-title">Fee Breakdown</div>'+bdRows.map(r=>'<div class="bd-row"><div class="bd-row-left"><div class="bd-dot" style="background:'+r.c+';opacity:'+(r.value>0?1:.25)+'"></div><span style="color:#666">'+escHtml(r.label)+'</span></div><span style="color:'+(r.value>0?'var(--dk)':'#ccc')+';font-weight:'+(r.value>0?600:400)+';font-variant-numeric:tabular-nums">'+fmt(r.value)+'</span></div>').join('')+(discAmt>0?'<div class="bd-row"><div class="bd-row-left"><div class="bd-dot" style="background:var(--rd)"></div><span style="color:var(--rd)">Discount ('+escHtml(discLabel)+(state.discReason?' – '+escHtml(state.discReason):'')+')'+' </span></div><span style="color:var(--rd);font-weight:600">–'+fmt(discAmt)+'</span></div>':'')+' <div class="bd-total-row"><span>Total</span><span style="color:#4db749">'+fmt(grandTotal)+'</span></div></div>';
    h+=buildEmailSection(d);
    h+='<p id="footer">All prices are estimates in AUD (ex. GST). Final fees subject to engagement review.</p>';
    return h;
  }

  function buildColl(key,num,title,subtitle,total,accentKey,items,rows,hasActivity,noGrey) {
    const expanded=state.expanded[key], ac=accentKey==='b'?'#29b6c9':'#4db749';
    const headerAc=(!expanded&&!noGrey)?'#aaa':ac, borderCol=(!expanded&&!noGrey)?'#d0d0d0':ac;
    const bg=expanded?'#fff':'#f8f8f8', border=(!expanded&&!noGrey)?'#e8e8e8':'var(--bd)';
    const titleCol=(!expanded&&!noGrey)?'#888':'var(--dk)', totalCol=(!expanded&&!noGrey)?'#bbb':(total>0?ac:'#ccc');
    let h='<div id="coll-'+key+'" class="coll '+(expanded?'expanded':'collapsed')+'" style="border-top:3px solid '+borderCol+';border-color:'+border+';background:'+bg+'">';
    h+='<button class="coll-header" data-toggle="'+key+'"><div class="coll-left"><svg class="coll-chevron" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="'+headerAc+'" stroke-width="2.5" stroke-linecap="round"><polyline points="9 18 15 12 9 6"/></svg><div class="coll-title-wrap">'+(num?'<span class="coll-num" style="color:'+headerAc+'">'+num+'</span>':'')+'<h3 style="color:'+titleCol+'">'+escHtml(title)+'</h3>'+(subtitle&&!expanded?'<p class="coll-sub">'+escHtml(subtitle)+'</p>':'')+'</div>'+(!expanded&&hasActivity?'<span class="coll-active-badge" style="color:'+(noGrey?ac:'#888')+';background:'+(noGrey?ac+'20':'#e8e8e8')+'">Active</span>':'')+'</div><div class="coll-right"><div class="coll-total-label">Section Total</div><div class="coll-total-val" style="color:'+totalCol+'">'+fmt(total)+'</div></div></button>';
    if(expanded){h+='<div class="coll-body">'+(subtitle?'<p class="coll-body-sub">'+escHtml(subtitle)+'</p>':'')+buildTable(key,items,rows,ac)+'</div>';}
    h+='</div>'; return h;
  }

  function buildNeColl(d) {
    const {neCoT,neTrT,neSmT,neOtT,neAnT,neTotal}=d, key='newEntities', expanded=state.expanded[key], ac='#4db749';
    const headerAc=expanded?ac:'#aaa', borderCol=expanded?ac:'#d0d0d0', bg=expanded?'#fff':'#f8f8f8';
    const border=expanded?'var(--bd)':'#e8e8e8', titleCol=expanded?'var(--dk)':'#888', totalCol=expanded?(neTotal>0?ac:'#ccc'):'#bbb';
    const hasNe=hasAct(NE_COMPANY_LIVE,state.neCoS)||hasAct(NE_TRUST_LIVE,state.neTrS)||hasAct(NE_SMSF_LIVE,state.neSmS)||hasAct(NE_OTHER_LIVE,state.neOtS)||hasAct(NE_ANNUAL_LIVE,state.neAnS);
    let h='<div id="coll-newEntities" class="coll '+(expanded?'expanded':'collapsed')+'" style="border-top:3px solid '+borderCol+';border-color:'+border+';background:'+bg+'">';
    h+='<button class="coll-header" data-toggle="newEntities"><div class="coll-left"><svg class="coll-chevron" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="'+headerAc+'" stroke-width="2.5" stroke-linecap="round"><polyline points="9 18 15 12 9 6"/></svg><div class="coll-title-wrap"><h3 style="color:'+titleCol+'">New Entities &amp; Registrations</h3>'+(!expanded?'<p class="coll-sub">Company/Trust/SMSF setup, deeds, and other registration services</p>':'')+'</div>'+(!expanded&&hasNe?'<span class="coll-active-badge" style="color:#888;background:#e8e8e8">Active</span>':'')+'</div><div class="coll-right"><div class="coll-total-label">Section Total</div><div class="coll-total-val" style="color:'+totalCol+'">'+fmt(neTotal)+'</div></div></button>';
    if(expanded){h+='<div class="coll-body"><div style="margin-top:12px">'+buildColl('neCompany','','Company',null,neCoT,'g',NE_COMPANY_LIVE,state.neCoS,hasAct(NE_COMPANY_LIVE,state.neCoS),false)+buildColl('neTrust','','Trust',null,neTrT,'g',NE_TRUST_LIVE,state.neTrS,hasAct(NE_TRUST_LIVE,state.neTrS),false)+buildColl('neSmsf','','SMSF Services',null,neSmT,'g',NE_SMSF_LIVE,state.neSmS,hasAct(NE_SMSF_LIVE,state.neSmS),false)+buildColl('neOther','','All Other Services',null,neOtT,'g',NE_OTHER_LIVE,state.neOtS,hasAct(NE_OTHER_LIVE,state.neOtS),false)+buildColl('neAnnual','','Annual Company Reviews',null,neAnT,'g',NE_ANNUAL_LIVE,state.neAnS,hasAct(NE_ANNUAL_LIVE,state.neAnS),false)+'</div></div>';}
    h+='</div>'; return h;
  }

  function buildDiscColl(d) {
    const {discAmt,discLabel}=d, key='discount', expanded=state.expanded[key], ac='#4db749';
    const totalDisplay=discAmt>0?'-'+fmt(discAmt):fmt(0), totalCol=discAmt>0?'var(--rd)':'#ccc';
    let h='<div id="coll-discount" class="coll '+(expanded?'expanded':'collapsed')+'" style="border-top:3px solid '+ac+';border-color:var(--bd);background:'+(expanded?'#fff':'#f8f8f8')+'">';
    h+='<button class="coll-header" data-toggle="discount"><div class="coll-left"><svg class="coll-chevron" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="'+ac+'" stroke-width="2.5" stroke-linecap="round"><polyline points="9 18 15 12 9 6"/></svg><div class="coll-title-wrap"><h3>Discount / Adjustment</h3>'+(!expanded?'<p class="coll-sub" style="color:#aaa">Apply a percentage or fixed discount to the total</p>':'')+'</div>'+(!expanded&&discAmt>0?'<span class="coll-active-badge" style="color:'+ac+';background:'+ac+'20">Active</span>':'')+'</div><div class="coll-right"><div class="coll-total-label">Discount</div><div class="coll-total-val" style="color:'+totalCol+'">'+totalDisplay+'</div></div></button>';
    if(expanded){h+='<div class="coll-body"><div class="grid-2-fixed" style="margin-top:8px"><div class="field"><label>Type</label><div class="toggle-group toggle-xs"><button id="disc-pct" class="'+(state.discType==='percent'?'active':'')+'">%</button><button id="disc-fix" class="'+(state.discType==='fixed'?'active':'')+'">$</button></div></div>'+fld(state.discType==='percent'?'Discount %':'Discount Amount ($)','number','discVal',state.discVal,state.discType==='percent'?'e.g. 10':'e.g. 500')+'</div><div class="field"><label>Reason</label><input dir="ltr" type="text" id="f-discReason" value="'+escHtml(state.discReason)+'" placeholder="Give a reason for discount..."></div>'+(discAmt>0?'<div class="disc-applied"><span style="font-size:14px;color:var(--rd);font-weight:700">-'+fmt(discAmt)+'</span><span style="font-size:12px;color:#999">discount applied</span>'+(state.discReason?'<span style="font-size:12px;color:#777;font-style:italic">— '+escHtml(state.discReason)+'</span>':'')+'</div>':'')+'</div>';}
    h+='</div>'; return h;
  }

  function buildTable(secKey,items,rows,ac) {
    let h='<div class="tbl-wrap"><table class="svc"><thead><tr><th>Service</th><th>Unit</th><th class="r">Base / Tidy-up Price</th><th class="c">Qty</th><th class="r">Subtotal</th><th style="width:36px"></th></tr></thead><tbody>';
    items.forEach((item,ri)=>{
      const row=rows[ri], rt=rowTotal(item,row), tq=row.splits.reduce((s,sp)=>s+(sp.qty||0),0), isBlue=ac==='#29b6c9';
      row.splits.forEach((sp,si)=>{
        const isFirst=si===0, isLast=si===row.splits.length-1, bp=getVariantPrice(item,sp);
        const p=sp.useGood?bp:(parseFloat(sp.cp)||0), am=item.am||1, sub=p*(sp.qty||0)*am;
        const isC=!sp.useGood, red=isC&&(parseFloat(sp.cp)||0)<bp;
        const pfx=isC?(red?'red':(isBlue?'blue':'green')):'', wc=isC?(red?'custom-red':(isBlue?'custom-blue':'custom-green')):'';
        const btnSt=sp.useGood?'border:2px solid '+ac+';background:'+ac+';color:#fff':'border:1.5px solid var(--bd);background:#fff;color:#aaa';
        const isHL=state.hlSec===secKey&&state.hlIdx===ri;
        h+='<tr class="'+(isLast?'':'split-row')+'" style="'+(isHL?'background:#fffbe6':'')+'" data-sec="'+secKey+'" data-ri="'+ri+'">';
        h+='<td style="padding-left:'+(si>0?28:8)+'px">'+(isFirst?escHtml(item.name):'<span class="split-label">↳ '+escHtml(item.name)+' (split '+(si+1)+')</span>');
        if(item.variants){h+='<br><select dir="ltr" class="variant-sel" data-sec="'+secKey+'" data-ri="'+ri+'" data-si="'+si+'">'+item.variants.map((v,vi)=>'<option value="'+vi+'"'+(vi===sp.variantIdx?' selected':'')+'>'+escHtml(v.label)+' \xb7 '+fmt(v.price)+'</option>').join('')+'</select>';}
        h+='</td>';
        h+='<td class="grey">'+(isFirst?(escHtml(item.unit)+(item.am>1?'<span class="am-badge" style="color:'+ac+';background:'+ac+'18">\xd7'+item.am+'</span>':'')):'&nbsp;')+'</td>';
        h+='<td class="r"><div class="price-cell"><button class="price-btn" data-sec="'+secKey+'" data-ri="'+ri+'" data-si="'+si+'" data-action="useGood" style="'+btnSt+'">'+fmt(bp)+'</button><div class="price-input-wrap '+wc+'"><span class="price-prefix '+pfx+'">$</span><input dir="ltr" id="inp-price-'+secKey+'-'+ri+'-'+si+'" class="'+pfx+'" type="number" min="0" value="'+sp.cp+'" data-sec="'+secKey+'" data-ri="'+ri+'" data-si="'+si+'" data-action="setPrice"></div></div></td>';
        h+='<td class="c"><input dir="ltr" id="inp-qty-'+secKey+'-'+ri+'-'+si+'" class="qty-input" type="number" min="0" value="'+(sp.qty||0)+'" data-sec="'+secKey+'" data-ri="'+ri+'" data-si="'+si+'" data-action="setQty"></td>';
        h+='<td class="r" style="color:'+(sub>0?'var(--dk)':'#ccc')+';font-weight:'+(sub>0?600:400)+';font-variant-numeric:tabular-nums">'+(isFirst&&row.splits.length>1?(isLast?fmt(rt):''):fmt(sub))+(isLast&&row.splits.length>1&&!isFirst?'<div class="row-total-badge" style="color:'+ac+'">Row: '+fmt(rt)+'</div>':'')+'</td>';
        h+='<td class="c">'+(isFirst&&tq>=1?'<button class="add-split-btn" data-sec="'+secKey+'" data-ri="'+ri+'" style="color:'+ac+'" title="Add split row">+</button>':'')+(si>0?'<button class="rem-split-btn" data-sec="'+secKey+'" data-ri="'+ri+'" data-si="'+si+'" title="Remove split">\xd7</button>':'')+'</td></tr>';
      });
    });
    h+='</tbody></table></div>'; return h;
  }

  function buildPackagesHTML() {
    let h='<div class="sec-head" style="border-bottom:2.5px solid #4db749"><div class="sec-head-left"><h3>Select a Package</h3><p>Pre-configured compliance &amp; bookkeeping bundles</p></div></div><div class="pkg-grid">';
    for(const pkg of PACKAGES){
      const sel=state.selPkg===pkg.id;
      h+='<div class="pkg-card '+(sel?'selected':'')+'" style="'+(sel?'border-color:'+pkg.color+';background:'+pkg.color+'08':'')+'"><div class="pkg-top-bar" style="background:'+pkg.color+'"></div><div class="pkg-tier" style="color:'+pkg.color+'">'+pkg.name+'</div><div class="pkg-price">'+fmt(pkg.monthly)+'<span>/mo</span></div><div class="pkg-annual">'+fmt(pkg.annual)+'/year</div><div class="pkg-target">'+pkg.target+'</div><div class="pkg-tagline">"'+escHtml(pkg.tagline)+'"</div><div class="pkg-section-lbl">Includes</div>'+pkg.features.map(f=>'<div class="pkg-feature"><span style="color:'+pkg.color+';font-weight:700;flex-shrink:0">✓</span>'+escHtml(f)+'</div>').join('')+'<div class="pkg-section-lbl" style="margin-top:10px">Not included</div>'+pkg.excluded.map(f=>'<div class="pkg-excl">– '+escHtml(f)+'</div>').join('')+'<button class="pkg-select-btn" data-pkg="'+pkg.id+'" style="background:'+(sel?pkg.color:'#fff')+';color:'+(sel?'#fff':pkg.color)+';border-color:'+pkg.color+'">'+(sel?'✓ Selected':'Select Package')+'</button><div class="pkg-delivery">Delivery: '+pkg.team+'</div></div>';
    }
    h+='</div>'; return h;
  }

  function buildEmailSection(d) {
    const {isCorp,baseTotal,calcStr,grandTotal,monthly,subtotal,discAmt,discLabel}=d;
    let h='<div style="margin-top:20px"><div class="card card-green"><div class="sec-head" style="border-bottom:2.5px solid #4db749"><div class="sec-head-left"><h3>Scope of Works &amp; Draft Email</h3><p>Generate a professional engagement document for your client</p></div></div><div class="grid-2">'+fld('Client Name','text','cName',state.cName,'e.g. Smith & Co')+fld('Client Email','email','cEmail',state.cEmail,'client@example.com')+'</div><div class="email-btn-row"><button class="btn-green" id="gen-email-btn" '+(state.emailLoading?'disabled':'')+'>'+(state.emailLoading?'<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2" class="spin"><path d="M12 2v4M12 18v4M4.93 4.93l2.83 2.83M16.24 16.24l2.83 2.83M2 12h4M18 12h4M4.93 19.07l2.83-2.83M16.24 7.76l2.83-2.83"/></svg>Generating…':'<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round"><rect x="2" y="4" width="20" height="16" rx="2"/><path d="M22 4L12 13 2 4"/></svg>Generate Draft Email')+'</button><button class="btn-blue" id="outlook-btn"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round"><path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"/><polyline points="22,6 12,13 2,6"/></svg>Draft in Outlook</button></div>';
    if(state.emailDraft){h+='<div class="email-wrap"><textarea dir="ltr" id="email-textarea">'+escHtml(state.emailDraft)+'</textarea><div class="email-actions-overlay"><button class="btn-outline" id="copy-email-btn"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="#666" stroke-width="2" stroke-linecap="round"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 01-2-2V4a2 2 0 012-2h9a2 2 0 012 2v1"/></svg>Copy</button><button class="btn-blue" id="outlook-btn2" style="padding:6px 14px;font-size:11px">Outlook</button></div></div>';}
    h+='</div></div>'; return h;
  }

  // ══════════════════════════════════════════════════════
  //  ROW STATE
  // ══════════════════════════════════════════════════════
  const SECTION_MAP = {
    yearEnd:{get items(){return YEAR_END_LIVE;},rows:()=>state.yeS,set:v=>state.yeS=v},
    other:{get items(){return OTHER_LIVE;},rows:()=>state.otS,set:v=>state.otS=v},
    indiv:{get items(){return INDIV_LIVE;},rows:()=>state.inS,set:v=>state.inS=v},
    book:{get items(){return BOOK_LIVE;},rows:()=>state.bkS,set:v=>state.bkS=v},
    neCompany:{get items(){return NE_COMPANY_LIVE;},rows:()=>state.neCoS,set:v=>state.neCoS=v},
    neTrust:{get items(){return NE_TRUST_LIVE;},rows:()=>state.neTrS,set:v=>state.neTrS=v},
    neSmsf:{get items(){return NE_SMSF_LIVE;},rows:()=>state.neSmS,set:v=>state.neSmS=v},
    neOther:{get items(){return NE_OTHER_LIVE;},rows:()=>state.neOtS,set:v=>state.neOtS=v},
    neAnnual:{get items(){return NE_ANNUAL_LIVE;},rows:()=>state.neAnS,set:v=>state.neAnS=v}
  };
  function updateSplit(sec,ri,si,patch){const m=SECTION_MAP[sec];m.set(m.rows().map((r,idx)=>idx!==ri?r:{...r,splits:r.splits.map((sp,j)=>j===si?{...sp,...patch}:sp)}));}
  function addSplit(sec,ri){const m=SECTION_MAP[sec],item=m.items[ri];m.set(m.rows().map((r,idx)=>idx!==ri?r:{...r,splits:[...r.splits,{useGood:true,cp:String(item.bad||item.good),qty:1,variantIdx:r.splits[0].variantIdx||0}]}));}
  function removeSplit(sec,ri,si){const m=SECTION_MAP[sec];m.set(m.rows().map((r,idx)=>idx!==ri?r:{...r,splits:r.splits.filter((_,j)=>j!==si)}));}

  // ══════════════════════════════════════════════════════
  //  SEARCH
  // ══════════════════════════════════════════════════════
  let ALL_SEARCH_LIVE = buildSearchIndex();
  function buildSearchIndex(){return[...YEAR_END_LIVE.map((it,idx)=>({...it,sec:'yearEnd',idx,sl:'Year End',sc:'#4db749'})),...OTHER_LIVE.map((it,idx)=>({...it,sec:'other',idx,sl:'Other',sc:'#4db749'})),...BOOK_LIVE.map((it,idx)=>({...it,sec:'book',idx,sl:'Bookkeeping',sc:'#29b6c9'})),...INDIV_LIVE.map((it,idx)=>({...it,sec:'indiv',idx,sl:'Individuals',sc:'#4db749'})),...NE_COMPANY_LIVE.map((it,idx)=>({...it,sec:'neCompany',idx,sl:'New Co.',sc:'#4db749'})),...NE_TRUST_LIVE.map((it,idx)=>({...it,sec:'neTrust',idx,sl:'New Trust',sc:'#4db749'})),...NE_SMSF_LIVE.map((it,idx)=>({...it,sec:'neSmsf',idx,sl:'SMSF',sc:'#4db749'})),...NE_OTHER_LIVE.map((it,idx)=>({...it,sec:'neOther',idx,sl:'Other NE',sc:'#4db749'})),...NE_ANNUAL_LIVE.map((it,idx)=>({...it,sec:'neAnnual',idx,sl:'Annual',sc:'#4db749'}))];}
  function rebuildSearchIndex(){ALL_SEARCH_LIVE=buildSearchIndex();}

  function handleSearch(q){
    state._searchQ=q;
    const rd=document.getElementById('search-results'); if(!q.trim()||!rd)return;
    const matches=ALL_SEARCH_LIVE.filter(s=>s.name.toLowerCase().includes(q.toLowerCase())).slice(0,8);
    if(!matches.length){rd.style.display='none';return;}
    rd.style.display='block';
    rd.innerHTML=matches.map(r=>'<button class="search-result" data-sec="'+r.sec+'" data-idx="'+r.idx+'"><span style="color:var(--dk)">'+escHtml(r.name)+'</span><span class="search-badge" style="background:'+(r.sc==='#29b6c9'?'#e4f6f9':'#e8f5e7')+';color:'+r.sc+'">'+r.sl+'</span></button>').join('');
    rd.querySelectorAll('.search-result').forEach(btn=>{btn.addEventListener('click',()=>{const sec=btn.dataset.sec,idx=parseInt(btn.dataset.idx),neKeys=['neCompany','neTrust','neSmsf','neOther','neAnnual'];state.expanded[sec]=true;if(neKeys.includes(sec))state.expanded.newEntities=true;state.hlSec=sec;state.hlIdx=idx;state._searchQ='';render();setTimeout(()=>{const el=document.getElementById('coll-'+sec);if(el)el.scrollIntoView({behavior:'smooth',block:'start'});},80);setTimeout(()=>{state.hlSec=null;state.hlIdx=null;render();},3000);});});
  }

  // ══════════════════════════════════════════════════════
  //  EMAIL
  // ══════════════════════════════════════════════════════
  function buildActiveItems(items,rows){const out=[];items.forEach((item,i)=>{rows[i].splits.forEach(sp=>{if((sp.qty||0)>0){const bp=getVariantPrice(item,sp),p=sp.useGood?bp:(parseFloat(sp.cp)||0),am=item.am||1,vLabel=item.variants&&item.variants[sp.variantIdx]?' ('+item.variants[sp.variantIdx].label+')':'';out.push({name:item.name+vLabel,unit:item.unit,price:p,qty:sp.qty,mult:am,sub:p*sp.qty*am});}});});return out;}

  function buildDocSections(d){const{isCorp,usePkg,pkgObj}=d,sections=[];if(usePkg){sections.push({title:pkgObj.name+' Package',items:[{name:pkgObj.name+' Package',unit:'Annual',price:pkgObj.annual,qty:1,mult:1,sub:pkgObj.annual}]});}else{if(isCorp){sections.push({title:'Year End Compliance Add-Ons',items:buildActiveItems(YEAR_END_LIVE,state.yeS)});sections.push({title:'Other Services',items:buildActiveItems(OTHER_LIVE,state.otS)});sections.push({title:'Bookkeeping',items:buildActiveItems(BOOK_LIVE,state.bkS)});}else{sections.push({title:'Individual Services',items:buildActiveItems(INDIV_LIVE,state.inS)});sections.push({title:'Year End Compliance Add-Ons',items:buildActiveItems(YEAR_END_LIVE,state.yeS)});sections.push({title:'Other Services',items:buildActiveItems(OTHER_LIVE,state.otS)});sections.push({title:'Bookkeeping',items:buildActiveItems(BOOK_LIVE,state.bkS)});}}const neItems=[...buildActiveItems(NE_COMPANY_LIVE,state.neCoS),...buildActiveItems(NE_TRUST_LIVE,state.neTrS),...buildActiveItems(NE_SMSF_LIVE,state.neSmS),...buildActiveItems(NE_OTHER_LIVE,state.neOtS),...buildActiveItems(NE_ANNUAL_LIVE,state.neAnS)];if(neItems.length)sections.push({title:'New Entities & Registrations',items:neItems});return sections.filter(s=>s.items.length>0);}

  function generateEmail(){
    const d=calcDerived(),{isCorp,baseTotal,calcStr,subtotal,discAmt,discLabel,grandTotal,monthly}=d,sections=buildDocSections(d);
    const client=state.cName||'Valued Client';
    let body='Dear '+client+',\n\n';
    body+='Thank you for the opportunity to work with you. Please find below our Scope of Works and Fee Estimate for your review.\n\n';
    let n=1;
    if(isCorp){
      body+=n+'. BASE ANNUAL COMPLIANCE FEE\n';
      body+='Based on your business profile ('+calcStr+'), your estimated base annual compliance fee is '+fmt(baseTotal)+'.\n\n';
      n++;
    }
    sections.forEach(sec=>{
      body+=n+'. '+sec.title.toUpperCase()+'\n';
      sec.items.forEach(it=>{
        body+='  • '+it.name+' — '+it.qty+(it.qty>1?' units':' unit')+' \xd7 '+fmt(it.price)+(it.mult>1?' \xd7 '+it.mult+' months':'')+' = '+fmt(it.sub)+' p.a.\n';
      });
      body+='\n';
      n++;
    });
    body+=n+'. FEE SUMMARY\n';
    body+='  Subtotal: '+fmt(subtotal)+'\n';
    if(discAmt>0){body+='  Discount ('+discLabel+(state.discReason?' — '+state.discReason:'')+'):  –'+fmt(discAmt)+'\n';}
    body+='  Estimated Annual Total: '+fmt(grandTotal)+' (approx. '+fmt(monthly)+'/month)\n\n';
    body+='Please note all fees are estimates in AUD exclusive of GST. Final fees are subject to engagement review and may vary based on the complexity of your affairs.\n\n';
    body+='Should you have any questions or wish to discuss further, please do not hesitate to reach out. We look forward to working with you.\n\n';
    body+='Kind regards,\nThe Carbon Group Team';
    state.emailDraft=body;
    state.emailLoading=false;
    render();
  }

  function draftInOutlook(){const d=calcDerived(),subject='Carbon Group – Scope of Works & Fee Estimate'+(state.cName?' – '+state.cName:''),body=state.emailDraft||'Dear '+(state.cName||'Valued Client')+',\n\nPlease find below our fee estimate.\n\nEstimated Annual Total: '+fmt(d.grandTotal)+' (~'+fmt(d.monthly)+'/month)\n\nAll prices are estimates in AUD (ex. GST). Final fees are subject to engagement review.\n\nKind regards,\nThe Carbon Group Team';window.open('mailto:'+encodeURIComponent(state.cEmail||'')+'?subject='+encodeURIComponent(subject)+'&body='+encodeURIComponent(body),'_blank');}

  // ══════════════════════════════════════════════════════
  //  EVENTS
  // ══════════════════════════════════════════════════════
  function attachEvents(d) {
    const on=(id,ev,fn)=>{const el=document.getElementById(id);if(el)el.addEventListener(ev,fn);};
    on('btn-corp','click',()=>{state.cType=0;state.expanded={yearEnd:true,other:true,book:true,indiv:false,newEntities:false,neCompany:false,neTrust:false,neSmsf:false,neOther:false,neAnnual:false,discount:false};render();});
    on('btn-indiv','click',()=>{state.cType=1;state.expanded={yearEnd:false,other:false,book:false,indiv:true,newEntities:false,neCompany:false,neTrust:false,neSmsf:false,neOther:false,neAnnual:false,discount:false};render();});
    // number/email inputs: update state on input but only render on blur (avoids cursor-jumping)
    [['turnover',v=>state.turnover=v],['cEmail',v=>state.cEmail=v],['discVal',v=>state.discVal=v]].forEach(([id,fn])=>{
      on('f-'+id,'input',e=>{fn(e.target.value);});
      on('f-'+id,'blur',()=>{render();});
    });
    // text inputs: can render on input since cursor restoration works for type=text
    [['cName',v=>state.cName=v],['discReason',v=>state.discReason=v]].forEach(([id,fn])=>{on('f-'+id,'input',e=>{fn(e.target.value);render();});});
    [['qIdx',v=>state.qIdx=v],['tIdx',v=>state.tIdx=v],['eIdx',v=>state.eIdx=v]].forEach(([id,fn])=>{on('f-'+id,'change',e=>{fn(parseInt(e.target.value));render();});});
    on('btn-services','click',()=>{state.viewMode='services';render();});
    on('btn-packages','click',()=>{state.viewMode='packages';render();});
    on('disc-pct','click',()=>{state.discType='percent';state.discVal='';render();});
    on('disc-fix','click',()=>{state.discType='fixed';state.discVal='';render();});
    document.querySelectorAll('.coll-header[data-toggle]').forEach(btn=>{btn.addEventListener('click',()=>{state.expanded[btn.dataset.toggle]=!state.expanded[btn.dataset.toggle];render();});});
    document.querySelectorAll('.pkg-select-btn[data-pkg]').forEach(btn=>{btn.addEventListener('click',e=>{e.stopPropagation();const id=btn.dataset.pkg;state.selPkg=state.selPkg===id?null:id;render();});});
    document.querySelectorAll('.pkg-card').forEach(card=>{card.addEventListener('click',()=>{const id=card.querySelector('[data-pkg]')?.dataset.pkg;if(id){state.selPkg=state.selPkg===id?null:id;render();}});});
    const si=document.getElementById('search-input');if(si){si.addEventListener('input',e=>handleSearch(e.target.value));if(state._searchQ){si.value=state._searchQ;handleSearch(state._searchQ);}}
    document.querySelectorAll('[data-action="useGood"]').forEach(btn=>{btn.addEventListener('click',()=>{const{sec,ri,si}=btn.dataset;updateSplit(sec,parseInt(ri),parseInt(si),{useGood:true});render();});});
    document.querySelectorAll('[data-action="setPrice"]').forEach(inp=>{
      inp.addEventListener('focus',()=>{const{sec,ri,si}=inp.dataset,m=SECTION_MAP[sec];if(!m)return;const sp=m.rows()[parseInt(ri)]?.splits[parseInt(si)];if(sp&&sp.useGood)updateSplit(sec,parseInt(ri),parseInt(si),{useGood:false});});
      inp.addEventListener('input',()=>{const{sec,ri,si}=inp.dataset;updateSplit(sec,parseInt(ri),parseInt(si),{useGood:false,cp:inp.value});});
      inp.addEventListener('blur',()=>{render();});
    });
    document.querySelectorAll('[data-action="setQty"]').forEach(inp=>{
      inp.addEventListener('input',()=>{const{sec,ri,si}=inp.dataset;updateSplit(sec,parseInt(ri),parseInt(si),{qty:Math.max(0,parseInt(inp.value)||0)});});
      inp.addEventListener('blur',()=>{render();});
    });
    document.querySelectorAll('.add-split-btn').forEach(btn=>{btn.addEventListener('click',()=>{addSplit(btn.dataset.sec,parseInt(btn.dataset.ri));render();});});
    document.querySelectorAll('.rem-split-btn').forEach(btn=>{btn.addEventListener('click',()=>{removeSplit(btn.dataset.sec,parseInt(btn.dataset.ri),parseInt(btn.dataset.si));render();});});
    document.querySelectorAll('.variant-sel').forEach(sel=>{sel.addEventListener('change',()=>{const{sec,ri,si}=sel.dataset,vi=parseInt(sel.value),m=SECTION_MAP[sec];if(!m)return;const item=m.items[parseInt(ri)];updateSplit(sec,parseInt(ri),parseInt(si),{variantIdx:vi,cp:String(item.variants?.[vi]?.price||item.good)});render();});});
    const ta=document.getElementById('email-textarea');if(ta)ta.addEventListener('input',e=>{state.emailDraft=e.target.value;});
    on('gen-email-btn','click',()=>{generateEmail();});
    on('copy-email-btn','click',()=>{navigator.clipboard.writeText(state.emailDraft);});
    on('outlook-btn','click',draftInOutlook);on('outlook-btn2','click',draftInOutlook);
    on('reset-btn','click',()=>{Object.assign(state,{cName:'',cEmail:'',cType:0,turnover:'',qIdx:0,tIdx:0,eIdx:0,yeS:initRows(YEAR_END_LIVE),otS:initRows(OTHER_LIVE),inS:initRows(INDIV_LIVE),bkS:initRows(BOOK_LIVE),neCoS:initRows(NE_COMPANY_LIVE),neTrS:initRows(NE_TRUST_LIVE),neSmS:initRows(NE_SMSF_LIVE),neOtS:initRows(NE_OTHER_LIVE),neAnS:initRows(NE_ANNUAL_LIVE),viewMode:'services',selPkg:null,discType:'percent',discVal:'',discReason:'',emailDraft:'',emailLoading:false,hlSec:null,hlIdx:null,_searchQ:'',expanded:{yearEnd:true,other:true,book:true,indiv:false,newEntities:false,neCompany:false,neTrust:false,neSmsf:false,neOther:false,neAnnual:false,discount:false}});render();});
  }

  // ══════════════════════════════════════════════════════
  //  SHAREPOINT LOADER
  // ══════════════════════════════════════════════════════
  const SP_LIST_NAME='PricingServices', SP_CACHE_KEY='carbonCalcSpCache';
  let spStatus='loading', spErrorMsg='';
  const SP_SECTION_MAP={'year end':'YEAR_END','other services':'OTHER','individuals':'INDIV','bookkeeping':'BOOK','ne company':'NE_COMPANY','ne trust':'NE_TRUST','ne smsf':'NE_SMSF','ne other':'NE_OTHER','ne annual':'NE_ANNUAL'};
  function parseVariants(raw){if(!raw||!String(raw).trim())return undefined;try{const v=JSON.parse(raw);if(Array.isArray(v)&&v.length)return v;}catch(e){}return undefined;}
  function spItemsToGroups(items){
    const groups={};
    for(const item of items){
      const key=SP_SECTION_MAP[(item.Title||'').trim().toLowerCase()];
      if(!key)continue;
      if(!groups[key])groups[key]=[];
      const g=parseFloat(item.GoodRate||item.Good_x0020_Rate||item.goodrate)||0;
      const b=parseFloat(item.BadRate||item.Bad_x0020_Rate||item.badrate)||0;
      const obj={name:item.ServiceType||item.Service_x0020_Type||'',unit:item.Unit||'',good:g,bad:b};
      const am=parseFloat(item.AnnualMultiplier||item.Annual_x0020_Multiplier);
      if(am>1)obj.am=am;
      const variants=parseVariants(item.Variants||item.variants);
      if(variants)obj.variants=variants;
      groups[key].push(obj);
    }
    return groups;
  }
  function applyGroups(groups){if(groups.YEAR_END?.length)YEAR_END_LIVE=sa(groups.YEAR_END);if(groups.OTHER?.length)OTHER_LIVE=sa(groups.OTHER);if(groups.INDIV?.length)INDIV_LIVE=sa(groups.INDIV);if(groups.BOOK?.length)BOOK_LIVE=sa(groups.BOOK);if(groups.NE_COMPANY?.length)NE_COMPANY_LIVE=groups.NE_COMPANY;if(groups.NE_TRUST?.length)NE_TRUST_LIVE=groups.NE_TRUST;if(groups.NE_SMSF?.length)NE_SMSF_LIVE=groups.NE_SMSF;if(groups.NE_OTHER?.length)NE_OTHER_LIVE=groups.NE_OTHER;if(groups.NE_ANNUAL?.length)NE_ANNUAL_LIVE=groups.NE_ANNUAL;state.yeS=initRows(YEAR_END_LIVE);state.otS=initRows(OTHER_LIVE);state.inS=initRows(INDIV_LIVE);state.bkS=initRows(BOOK_LIVE);state.neCoS=initRows(NE_COMPANY_LIVE);state.neTrS=initRows(NE_TRUST_LIVE);state.neSmS=initRows(NE_SMSF_LIVE);state.neOtS=initRows(NE_OTHER_LIVE);state.neAnS=initRows(NE_ANNUAL_LIVE);}
  function saveToCache(items){try{localStorage.setItem(SP_CACHE_KEY,JSON.stringify({ts:Date.now(),items}));}catch(e){}}
  function loadFromCache(){try{const raw=localStorage.getItem(SP_CACHE_KEY);if(!raw)return null;return JSON.parse(raw);}catch(e){return null;}}
  function updateStatusChip(){const chip=document.getElementById('sp-status-chip');if(!chip)return;const styles={live:{text:'✓ Prices loaded from SharePoint',bg:'#e8f5e7',color:'#3a9c38',border:'#4db74960'},cached:{text:'⚠ Offline — showing cached prices',bg:'#fff3cd',color:'#856404',border:'#ffc10760'},fallback:{text:'⚠ SharePoint unreachable — using built-in prices',bg:'#fff3cd',color:'#856404',border:'#ffc10760'},error:{text:'✕ SharePoint error — using built-in prices',bg:'#fde8e8',color:'#c0392b',border:'#e74c3c60'},loading:{text:'Loading prices…',bg:'#f0f0f0',color:'#888',border:'#ddd'}};const s=styles[spStatus]||styles.loading;chip.textContent=s.text;chip.style.background=s.bg;chip.style.color=s.color;chip.style.borderColor=s.border;chip.title=spErrorMsg||'';}
  function showErrorBanner(msg){const ex=document.getElementById('sp-error-banner');if(ex)ex.remove();if(!msg)return;const b=document.createElement('div');b.id='sp-error-banner';b.innerHTML='<span>⚠ <strong>SharePoint unavailable.</strong> '+escHtml(msg)+' — Running with '+(spStatus==='cached'?'cached':'built-in')+' prices.</span><button onclick="this.parentElement.remove()" style="background:none;border:none;cursor:pointer;font-size:16px;color:inherit;padding:0 0 0 12px">\xd7</button>';Object.assign(b.style,{display:'flex',justifyContent:'space-between',alignItems:'center',padding:'12px 20px',background:'#fff8e1',borderLeft:'4px solid #ffc107',borderRadius:'8px',marginBottom:'20px',fontSize:'13px',color:'#5d4037'});const main=document.getElementById('main');if(main)main.prepend(b);}

  async function loadFromSharePoint(){
    let url=(siteUrl||'').replace(/\/$/,'');
    if(!url){try{if(window._spPageContextInfo?.siteAbsoluteUrl)url=window._spPageContextInfo.siteAbsoluteUrl.replace(/\/$/,'');}catch(e){}}
    if(!url){spStatus='fallback';spErrorMsg='SP_SITE_URL not configured.';updateStatusChip();return;}
    const endpoint=url+'/_api/web/lists/getbytitle(\''+SP_LIST_NAME+'\')/items?$top=5000';
    try{const res=await fetch(endpoint,{headers:{'Accept':'application/json;odata=nometadata'},credentials:'same-origin'});if(!res.ok)throw new Error('SharePoint returned HTTP '+res.status);const json=await res.json(),items=json.value||[];if(!items.length)throw new Error('List returned 0 items — check list name and column names');applyGroups(spItemsToGroups(items));saveToCache(items);spStatus='live';spErrorMsg='';console.info('[Services Calculator] Loaded '+items.length+' items.');}
    catch(err){spErrorMsg=err.message;console.warn('[Services Calculator] '+err.message);const cached=loadFromCache();if(cached?.items?.length){applyGroups(spItemsToGroups(cached.items));spStatus='cached';const age=Math.round((Date.now()-(cached.ts||0))/60000);spErrorMsg+=' (cached, '+age+' min ago)';}else{spStatus='error';}}
    updateStatusChip();rebuildSearchIndex();render();
    if(spStatus==='error'||spStatus==='cached'){showErrorBanner(spErrorMsg);}
  }

  async function loadPackagesFromSharePoint(){
    let url=(siteUrl||'').replace(/\/$/,'');
    if(!url){try{if(window._spPageContextInfo?.siteAbsoluteUrl)url=window._spPageContextInfo.siteAbsoluteUrl.replace(/\/$/,'');}catch(e){}}
    if(!url)return;
    const endpoint=url+'/_api/web/lists/getbytitle(\'PricingPackages\')/items?$top=100';
    try{
      const res=await fetch(endpoint,{headers:{'Accept':'application/json;odata=nometadata'},credentials:'same-origin'});
      if(!res.ok)throw new Error('PricingPackages HTTP '+res.status);
      const json=await res.json(),items=(json.value||[]);
      if(!items.length)throw new Error('PricingPackages list is empty');
      const splitPipe=(v)=>(v||'').split('|').map(s=>s.trim()).filter(s=>s.length>0);
      const mapped=items.map(item=>({
        id:(item.Title||'').toLowerCase().replace(/\s+/g,'-'),
        name:item.Title||'',
        monthly:parseFloat(item.Monthly||item.monthly)||0,
        annual:parseFloat(item.Annual||item.annual)||0,
        color:item.Colour||item.Color||item.colour||'#6c757d',
        target:item.Target||item.target||'',
        tagline:item.Tagline||item.tagline||'',
        features:splitPipe(item.Features||item.features),
        excluded:splitPipe(item.Excluded||item.excluded),
        _sort:parseInt(item.SortOrder||item.sortorder)||99
      }));
      mapped.sort((a,b)=>a._sort-b._sort);
      PACKAGES=mapped;
      console.info('[Services Calculator] Loaded '+mapped.length+' packages from SharePoint.');
      render();
    }catch(err){
      console.warn('[Services Calculator] PricingPackages: '+err.message+' — using built-in packages.');
    }
  }

  // ══════════════════════════════════════════════════════
  //  BOOT
  // ══════════════════════════════════════════════════════
  render();
  loadFromSharePoint();
  loadPackagesFromSharePoint();

} // end initCalculator
