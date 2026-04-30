<!DOCTYPE html>
<html lang="ru">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover"/>
  <title>Quality | Хаб Качества 360°</title>
  <link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&display=swap" rel="stylesheet">
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
  <script src="https://cdn.sheetjs.com/xlsx-0.20.2/package/dist/xlsx.full.min.js"></script>
  <style>
    :root {
      --f-indigo: #6366F1;
      --f-indigo-hover: #818CF8;
      --f-indigo-soft: rgba(99,102,241,0.08);
      --f-indigo-glow: rgba(99,102,241,0.25);
      --f-violet: #8B5CF6;
      --f-violet-soft: rgba(139,92,246,0.08);
      --f-cyan: #06B6D4;
      --f-cyan-soft: rgba(6,182,212,0.08);
      --f-emerald: #10B981;
      --f-emerald-soft: rgba(16,185,129,0.08);
      --f-rose: #F43F5E;
      --f-rose-soft: rgba(244,63,94,0.08);
      --f-amber: #F59E0B;
      --f-amber-soft: rgba(245,158,11,0.08);
      --f-grad: linear-gradient(135deg, #6366F1 0%, #8B5CF6 50%, #06B6D4 100%);
      --f-grad-btn: linear-gradient(135deg, #6366F1 0%, #8B5CF6 100%);
      --f-zinc-950: #09090B;
      --f-zinc-900: #18181B;
      --f-zinc-800: #27272A;
      --f-zinc-700: #3F3F46;
      --f-zinc-500: #71717A;
      --f-zinc-400: #A1A1AA;
      --f-zinc-300: #D4D4D8;
      --f-zinc-200: #E4E4E7;
      --f-zinc-100: #F4F4F5;
      --f-zinc-50: #FAFAFA;
      --f-white: #FFFFFF;
      --f-glass-bg: rgba(255,255,255,0.03);
      --f-glass-border: rgba(255,255,255,0.06);
      --f-glass-hover: rgba(255,255,255,0.06);
      --f-light-glass-bg: rgba(255,255,255,0.7);
      --f-light-glass-border: rgba(0,0,0,0.06);
      --f-r-sm: 10px;
      --f-r-md: 14px;
      --f-r-lg: 18px;
      --f-r-xl: 24px;
      --f-r-full: 9999px;
      --f-shadow: 0 1px 3px rgba(0,0,0,0.3), 0 4px 16px rgba(0,0,0,0.2);
      --f-shadow-lg: 0 4px 8px rgba(0,0,0,0.3), 0 8px 32px rgba(0,0,0,0.25);
      --f-shadow-indigo: 0 4px 20px rgba(99,102,241,0.2);
      --f-t: 0.2s cubic-bezier(0.4, 0, 0.2, 1);
      --bee-yellow: #FFD600;
      --bee-yellow-glow: rgba(255,214,0,0.4);
    }
    * { margin:0; padding:0; box-sizing:border-box; }
    body {
      font-family: 'Manrope', sans-serif;
      background: var(--f-zinc-950);
      min-height: 100vh;
      color: var(--f-zinc-100);
      transition: background 0.5s ease, color 0.5s ease;
      font-weight: 500;
      -webkit-font-smoothing: antialiased;
      overflow-x: hidden;
    }
    body:not(.dark) { background: var(--f-zinc-50); color: var(--f-zinc-900); }
    body::after {
      content: ''; position: fixed; inset: 0; z-index: 9998; pointer-events: none; opacity: 0.015;
      background-image: url("data:image/svg+xml,%3Csvg viewBox='0 0 256 256' xmlns='http://www.w3.org/2000/svg'%3E%3Cfilter id='n'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='0.9' numOctaves='4' stitchTiles='stitch'/%3E%3C/filter%3E%3Crect width='100%25' height='100%25' filter='url(%23n)'/%3E%3C/svg%3E");
    }
    body:not(.dark)::after { opacity: 0.02; }
    .mesh-orb {
      position: fixed; border-radius: 50%; pointer-events: none; z-index: 0;
      filter: blur(100px); opacity: 0.4;
      animation: mesh-float 30s infinite ease-in-out;
    }
    .mesh-orb-1 { width: 600px; height: 600px; top: -10%; left: -5%; background: radial-gradient(circle, rgba(99,102,241,0.3), transparent 70%); }
    .mesh-orb-2 { width: 500px; height: 500px; bottom: -10%; right: -5%; background: radial-gradient(circle, rgba(6,182,212,0.25), transparent 70%); animation-delay: -10s; }
    .mesh-orb-3 { width: 400px; height: 400px; top: 40%; left: 50%; background: radial-gradient(circle, rgba(139,92,246,0.2), transparent 70%); animation-delay: -20s; }
    body:not(.dark) .mesh-orb { opacity: 0.15; }
    @keyframes mesh-float {
      0%, 100% { transform: translate(0,0) scale(1); }
      25% { transform: translate(30px,-40px) scale(1.05); }
      50% { transform: translate(-20px,30px) scale(0.95); }
      75% { transform: translate(40px,20px) scale(1.02); }
    }
    .container { max-width: 1400px; margin: 0 auto; padding: 2rem 2rem 3rem; position: relative; z-index: 1; }
    header { position: relative; text-align: center; margin-bottom: 2rem; z-index: 2; }
    /* ===== Beeline Logo ===== */
    .bee-logo-wrap {
      display: flex; justify-content: center; margin-bottom: 1.5rem;
    }
    .bee-logo-wrap svg {
      height: 36px; width: auto;
      filter: drop-shadow(0 2px 8px rgba(255,214,0,0.25));
      transition: filter 0.3s ease;
    }
    .bee-logo-wrap svg:hover {
      filter: drop-shadow(0 4px 16px rgba(255,214,0,0.5));
    }
    body:not(.dark) .bee-logo-wrap svg .bee-pill { fill: #FFD600; }
    body.dark .bee-logo-wrap svg .bee-pill { fill: #FFD600; }
    body:not(.dark) .bee-logo-wrap svg .bee-text { fill: #000000; }
    body.dark .bee-logo-wrap svg .bee-text { fill: #000000; }
    .theme-toggle {
      position: absolute; top: 0; right: 0;
      background: var(--f-glass-bg); backdrop-filter: blur(12px);
      color: var(--f-zinc-400); border-radius: var(--f-r-full);
      padding: 9px 20px; font-size: 0.8rem; font-weight: 600;
      cursor: pointer; transition: var(--f-t);
      display: flex; align-items: center; gap: 8px; font-family: inherit;
      border: 1px solid var(--f-glass-border);
    }
    body:not(.dark) .theme-toggle { background: var(--f-light-glass-bg); border-color: var(--f-light-glass-border); color: var(--f-zinc-500); }
    .theme-toggle:hover { border-color: var(--f-indigo); color: var(--f-indigo); }
    h1 {
      font-size: 3.6rem; font-weight: 800;
      letter-spacing: -0.04em;
      margin-bottom: 0; animation: fadeInUp 0.8s ease;
      line-height: 1;
      background: var(--f-grad);
      -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;
    }
    body:not(.dark) h1 { background: linear-gradient(135deg, #4F46E5 0%, #7C3AED 50%, #0891B2 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text; }
    @keyframes fadeInUp { from { opacity:0; transform:translateY(20px); } to { opacity:1; transform:translateY(0); } }
    @keyframes fadeIn { from { opacity:0; } to { opacity:1; } }
    @keyframes scaleIn { from { opacity:0; transform:scale(0.97); } to { opacity:1; transform:scale(1); } }
    .main-title {
      text-align: center; font-size: 1.5rem; font-weight: 700;
      margin: 2rem 0 2.5rem; position: relative; display: inline-block; width: 100%;
      padding-bottom: 14px;
    }
    .main-title::after {
      content: ''; position: absolute; bottom: 0; left: 50%; transform: translateX(-50%);
      width: 48px; height: 3px; background: var(--f-grad); border-radius: 2px;
    }
    body:not(.dark) .main-title { color: var(--f-zinc-900); }
    .page-title-icon { margin-right: 10px; }
    .buttons-grid {
      display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
      gap: 1.25rem; margin: 2rem 0 3rem;
    }
    .direction-card {
      background: var(--f-glass-bg); backdrop-filter: blur(16px);
      border-radius: var(--f-r-xl); padding: 2.5rem 2rem 2rem; text-align: center;
      cursor: pointer; transition: var(--f-t);
      border: 1px solid var(--f-glass-border);
      position: relative; overflow: hidden;
      display: flex; flex-direction: column; align-items: center; justify-content: center;
      min-height: 210px;
    }
    body:not(.dark) .direction-card { background: var(--f-white); border-color: var(--f-light-glass-border); box-shadow: 0 1px 3px rgba(0,0,0,0.04), 0 4px 16px rgba(0,0,0,0.03); }
    .direction-card::before {
      content: ''; position: absolute; top: 0; left: 0;
      width: 100%; height: 2px; background: var(--f-grad);
      opacity: 0; transition: opacity 0.3s ease;
    }
    .direction-card:hover::before { opacity: 1; }
    .direction-card:hover { border-color: rgba(99,102,241,0.2); transform: translateY(-4px); box-shadow: var(--f-shadow-lg); }
    body:not(.dark) .direction-card:hover { border-color: rgba(99,102,241,0.15); box-shadow: 0 4px 8px rgba(0,0,0,0.06), 0 12px 36px rgba(0,0,0,0.08); }
    .drag-icon, .link-drag-icon {
      position: absolute; bottom: 10px; right: 10px;
      background: none; border: none;
      cursor: grab; z-index: 10; font-size: 0.65rem;
      color: var(--f-zinc-700); opacity: 0.2;
      border-radius: 4px; padding: 2px;
      transition: all 0.2s ease;
    }
    body:not(.dark) .drag-icon, body:not(.dark) .link-drag-icon { color: var(--f-zinc-400); }
    .drag-icon:hover, .link-drag-icon:hover { opacity: 0.8; color: var(--f-indigo); transform: scale(1.1); }
    .admin-icon {
      position: absolute; top: 12px; background: transparent !important; border: none !important;
      cursor: pointer; z-index: 10; font-size: 0.85rem; padding: 4px;
      transition: all 0.2s ease; opacity: 0.3;
    }
    .admin-icon:hover { opacity: 1; transform: scale(1.1); }
    .admin-icon.edit { right: 42px; color: var(--f-indigo); }
    .admin-icon.delete { right: 12px; color: var(--f-rose); }
    .card-icon {
      font-size: 1.6rem; margin-bottom: 0.85rem; display: inline-flex; align-items: center; justify-content: center;
      width: 56px; height: 56px; border-radius: 14px; transition: all 0.3s ease;
    }
    .ci-indigo { background: var(--f-indigo-soft); color: var(--f-indigo); }
    .ci-violet { background: var(--f-violet-soft); color: var(--f-violet); }
    .ci-cyan { background: var(--f-cyan-soft); color: var(--f-cyan); }
    .ci-emerald { background: var(--f-emerald-soft); color: var(--f-emerald); }
    .ci-rose { background: var(--f-rose-soft); color: var(--f-rose); }
    .direction-card:hover .card-icon { transform: scale(1.08); }
    .direction-name { font-size: 1.2rem; font-weight: 700; margin: 0; letter-spacing: 0.01em; }
    body:not(.dark) .direction-name { color: var(--f-zinc-900); }
    .links-container {
      display: grid; grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
      gap: 0.75rem; margin-top: 2rem;
    }
    .link-btn {
      background: var(--f-glass-bg); backdrop-filter: blur(12px);
      border: 1px solid var(--f-glass-border); padding: 1.1rem 1.4rem;
      border-radius: var(--f-r-lg); font-size: 0.88rem; font-weight: 500;
      text-align: left; cursor: pointer; transition: var(--f-t);
      display: flex; align-items: center; gap: 12px; position: relative;
      font-family: inherit; color: var(--f-zinc-200);
    }
    body:not(.dark) .link-btn { background: var(--f-white); color: var(--f-zinc-900); border-color: var(--f-light-glass-border); box-shadow: 0 1px 2px rgba(0,0,0,0.03); }
    .link-btn:hover {
      border-color: rgba(99,102,241,0.3);
      background: var(--f-indigo-soft); color: var(--f-indigo-hover);
      transform: translateX(3px);
    }
    body:not(.dark) .link-btn:hover { background: var(--f-indigo-soft); color: var(--f-indigo); border-color: rgba(99,102,241,0.2); }
    .link-btn i { font-size: 0.9rem; color: var(--f-indigo); transition: var(--f-t); flex-shrink: 0; opacity: 0.7; }
    .link-btn:hover i { opacity: 1; }
    .link-admin-icon {
      position: absolute; top: 10px; background: transparent !important; border: none !important;
      cursor: pointer; z-index: 10; font-size: 0.75rem; padding: 4px 6px;
      transition: all 0.2s ease; opacity: 0.3;
    }
    .link-admin-icon:hover { opacity: 1; transform: scale(1.1); }
    .link-admin-icon.edit { right: 70px; color: var(--f-indigo); }
    .link-admin-icon.delete { right: 36px; color: var(--f-rose); }
    .dragging { opacity: 0.4; transform: scale(0.98); }
    #back-btn {
      position: fixed; top: 2rem; left: 2rem; z-index: 100;
      background: var(--f-glass-bg); backdrop-filter: blur(12px);
      border: 1px solid var(--f-glass-border);
      padding: 9px 20px; border-radius: var(--f-r-full); font-weight: 600;
      cursor: pointer; transition: var(--f-t);
      display: inline-flex; align-items: center; gap: 8px;
      color: var(--f-zinc-400); font-family: inherit; font-size: 0.8rem;
    }
    body:not(.dark) #back-btn { background: var(--f-light-glass-bg); color: var(--f-zinc-600); border-color: var(--f-light-glass-border); }
    #back-btn:hover { border-color: var(--f-indigo); color: var(--f-indigo); transform: translateX(-3px); }
    /* ===== Quality Tooltip ===== */
    .quality-tooltip {
      position: absolute;
      bottom: calc(100% + 12px);
      left: 50%;
      transform: translateX(-50%) translateY(6px);
      min-width: 320px;
      max-width: 420px;
      padding: 14px 18px;
      background: var(--f-zinc-900);
      backdrop-filter: blur(20px);
      border: 1px solid var(--f-indigo);
      border-radius: var(--f-r-md);
      font-size: 0.82rem;
      font-weight: 400;
      line-height: 1.55;
      color: var(--f-zinc-200);
      pointer-events: none;
      opacity: 0;
      visibility: hidden;
      transition: opacity 0.25s ease, transform 0.25s ease, visibility 0.25s ease;
      z-index: 999;
      box-shadow: 0 8px 32px rgba(99,102,241,0.2), 0 2px 8px rgba(0,0,0,0.4);
    }
    body:not(.dark) .quality-tooltip {
      background: var(--f-white);
      color: var(--f-zinc-800);
      border-color: rgba(99,102,241,0.3);
      box-shadow: 0 8px 32px rgba(99,102,241,0.1), 0 2px 8px rgba(0,0,0,0.08);
    }
    .quality-tooltip::after {
      content: '';
      position: absolute;
      top: 100%;
      left: 50%;
      transform: translateX(-50%);
      border: 7px solid transparent;
      border-top-color: var(--f-indigo);
    }
    body:not(.dark) .quality-tooltip::after {
      border-top-color: rgba(99,102,241,0.3);
    }
    .quality-tooltip .tt-icon {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      width: 22px;
      height: 22px;
      border-radius: 6px;
      background: var(--f-indigo-soft);
      color: var(--f-indigo);
      font-size: 0.65rem;
      margin-right: 6px;
      flex-shrink: 0;
      vertical-align: middle;
    }
    .quality-tooltip .tt-header {
      display: flex;
      align-items: center;
      margin-bottom: 8px;
      font-weight: 700;
      font-size: 0.78rem;
      color: var(--f-indigo);
      text-transform: uppercase;
      letter-spacing: 0.04em;
    }
    body:not(.dark) .quality-tooltip .tt-header {
      color: var(--f-indigo);
    }
    .quality-tooltip strong {
      color: var(--f-indigo-hover);
      font-weight: 700;
    }
    body:not(.dark) .quality-tooltip strong {
      color: var(--f-indigo);
    }
    .link-btn:hover .quality-tooltip {
      opacity: 1;
      visibility: visible;
      transform: translateX(-50%) translateY(0);
    }
    /* ===== Feedback ===== */
    .feedback-modern {
      margin-top: 5rem; padding: 1.5rem 1.75rem;
      background: var(--f-glass-bg); backdrop-filter: blur(16px);
      border-radius: var(--f-r-xl); text-align: center;
      border: 1px solid var(--f-glass-border);
      animation: scaleIn 0.4s ease;
      max-width: 480px; margin-left: auto; margin-right: auto;
      position: relative;
    }
    .feedback-modern::before {
      content: ''; position: absolute; top: -24px; left: 50%; transform: translateX(-50%);
      width: 40px; height: 3px; background: var(--f-grad); border-radius: 2px; opacity: 0.4;
    }
    body:not(.dark) .feedback-modern { background: var(--f-white); border-color: var(--f-light-glass-border); box-shadow: 0 1px 3px rgba(0,0,0,0.04), 0 4px 16px rgba(0,0,0,0.03); }
    .feedback-modern h3 { font-size: 0.95rem; margin-bottom: 0.75rem; font-weight: 600; }
    body:not(.dark) .feedback-modern h3 { color: var(--f-zinc-900); }
    .feedback-modern p { font-size: 0.78rem; }
    .stars-modern { display: flex; gap: 12px; justify-content: center; margin: 0.85rem 0; direction: ltr; }
    .stars-modern i { font-size: 1.5rem; cursor: pointer; transition: all 0.15s ease; color: var(--f-zinc-700); }
    body:not(.dark) .stars-modern i { color: var(--f-zinc-300); }
    .stars-modern i:hover, .stars-modern i.active {
      color: var(--bee-yellow) !important;
      transform: scale(1.15);
      filter: drop-shadow(0 0 6px var(--bee-yellow-glow));
    }
    textarea {
      width: 100%; max-width: 460px; padding: 0.75rem 0.9rem;
      border-radius: var(--f-r-md); border: 1px solid var(--f-glass-border);
      background: var(--f-glass-bg); backdrop-filter: blur(8px); font-family: inherit;
      resize: vertical; margin: 0.6rem 0; outline: none;
      transition: var(--f-t); font-size: 0.85rem; color: inherit;
    }
    textarea:focus { border-color: var(--f-indigo); box-shadow: 0 0 0 3px var(--f-indigo-soft); }
    body:not(.dark) textarea { background: var(--f-zinc-50); border-color: var(--f-light-glass-border); color: var(--f-zinc-900); }
    .btn-primary {
      background: var(--f-grad-btn); border: none; padding: 9px 24px;
      border-radius: var(--f-r-full); color: white; font-weight: 600;
      cursor: pointer; transition: var(--f-t);
      font-size: 0.8rem; margin-top: 6px; font-family: inherit;
    }
    .btn-primary:hover { transform: translateY(-1px); box-shadow: var(--f-shadow-indigo); opacity: 0.9; }
    .admin-fab {
      position: fixed; bottom: 24px; right: 24px;
      width: 52px; height: 52px; border-radius: 50%;
      background: var(--f-grad-btn); border: none; color: white;
      font-size: 1.2rem; cursor: pointer;
      box-shadow: var(--f-shadow-indigo);
      transition: var(--f-t); z-index: 1000;
      display: flex; align-items: center; justify-content: center;
    }
    .admin-fab:hover { transform: scale(1.08); box-shadow: 0 8px 28px rgba(99,102,241,0.35); }
    footer { text-align: center; margin-top: 4rem; padding-top: 2rem; font-size: 0.75rem; opacity: 0.35; border-top: 1px solid var(--f-glass-border); }
    body:not(.dark) footer { border-top-color: var(--f-light-glass-border); color: var(--f-zinc-500); }
    .analytics-container { max-width: 1400px; margin: 0 auto; padding: 20px; }
    .tabs {
      display: inline-flex; margin-bottom: 24px; gap: 4px;
      background: var(--f-glass-bg); backdrop-filter: blur(12px);
      padding: 4px; border-radius: var(--f-r-full);
      border: 1px solid var(--f-glass-border);
    }
    body:not(.dark) .tabs { background: var(--f-zinc-100); border-color: var(--f-light-glass-border); }
    .tab-button {
      padding: 8px 18px; font-size: 0.82rem; font-weight: 600;
      background: transparent; border: none;
      border-radius: var(--f-r-full); cursor: pointer; transition: var(--f-t);
      color: var(--f-zinc-500); font-family: inherit;
    }
    .tab-button.active { background: var(--f-indigo); color: white; }
    body:not(.dark) .tab-button { color: var(--f-zinc-500); }
    .tab-content { display: none; }
    .tab-content.active { display: block; animation: fadeIn 0.25s ease; }
    .stats-cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 0.75rem; margin-bottom: 2rem; }
    .stat-card {
      background: var(--f-glass-bg); backdrop-filter: blur(12px);
      border-radius: var(--f-r-lg); padding: 1.1rem; text-align: center;
      border: 1px solid var(--f-glass-border); transition: var(--f-t);
    }
    body:not(.dark) .stat-card { background: var(--f-white); border-color: var(--f-light-glass-border); box-shadow: 0 1px 2px rgba(0,0,0,0.03); }
    .stat-card:hover { border-color: rgba(99,102,241,0.15); }
    .stat-number { font-size: 2rem; font-weight: 800; color: var(--f-indigo); letter-spacing: -0.02em; }
    body:not(.dark) .stat-number { color: var(--f-indigo); }
    .stat-label { font-size: 0.75rem; opacity: 0.5; font-weight: 500; margin-top: 2px; }
    .data-table-wrapper { overflow-x: auto; margin-top: 16px; border-radius: var(--f-r-md); border: 1px solid var(--f-glass-border); width: 100%; }
    body:not(.dark) .data-table-wrapper { border-color: var(--f-light-glass-border); }
    table { width: 100%; border-collapse: collapse; background: var(--f-glass-bg); backdrop-filter: blur(8px); overflow: hidden; table-layout: auto; }
    body:not(.dark) table { background: var(--f-white); }
    th, td { padding: 12px 14px; text-align: left; border-bottom: 1px solid var(--f-glass-border); white-space: normal; word-wrap: break-word; font-size: 0.85rem; }
    body:not(.dark) th, body:not(.dark) td { border-bottom-color: var(--f-light-glass-border); }
    th { background: var(--f-indigo-soft); color: var(--f-indigo-hover); font-weight: 600; text-transform: uppercase; font-size: 0.68rem; letter-spacing: 0.06em; position: sticky; top: 0; z-index: 10; }
    body:not(.dark) th { color: var(--f-indigo); }
    .export-btn {
      background: var(--f-grad-btn); border: none; padding: 9px 20px;
      border-radius: var(--f-r-full); cursor: pointer; font-weight: 600;
      color: white; font-family: inherit; font-size: 0.82rem;
      transition: var(--f-t);
    }
    .export-btn:hover { opacity: 0.9; transform: translateY(-1px); box-shadow: var(--f-shadow-indigo); }
    .filter-group { display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 16px; }
    .filter-group select, .filter-group input {
      padding: 8px 14px; border-radius: var(--f-r-full);
      border: 1px solid var(--f-glass-border); background: var(--f-glass-bg);
      backdrop-filter: blur(8px); font-family: inherit;
      color: inherit; cursor: pointer; font-weight: 500; outline: none;
      transition: var(--f-t); font-size: 0.85rem;
    }
    body:not(.dark) .filter-group select, body:not(.dark) .filter-group input { background: var(--f-white); border-color: var(--f-light-glass-border); color: var(--f-zinc-900); }
    .filter-group select:focus, .filter-group input:focus { border-color: var(--f-indigo); box-shadow: 0 0 0 3px var(--f-indigo-soft); }
    @media (max-width: 768px) {
      .container { padding: 1rem; }
      h1 { font-size: 2.4rem; }
      .buttons-grid { gap: 1rem; }
      #back-btn { top: 1rem; left: 1rem; padding: 7px 14px; font-size: 0.75rem; }
      header { margin-top: 2.5rem; }
      .tabs { gap: 2px; padding: 3px; }
      .tab-button { padding: 6px 12px; font-size: 0.75rem; }
      .feedback-modern { padding: 1.25rem 1rem; margin-top: 4rem; }
      .feedback-modern h3 { font-size: 0.88rem; }
      .stars-modern i { font-size: 1.3rem; }
      .stats-cards { grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); }
      .bee-logo-wrap svg { height: 28px; }
      .quality-tooltip {
        min-width: 260px;
        max-width: 320px;
        left: 0;
        transform: translateX(0) translateY(6px);
        font-size: 0.78rem;
        padding: 12px 14px;
      }
      .link-btn:hover .quality-tooltip {
        transform: translateX(0) translateY(0);
      }
      .quality-tooltip::after {
        left: 24px;
        transform: none;
      }
    }
  </style>
</head>
<body>
  <div class="mesh-orb mesh-orb-1"></div>
  <div class="mesh-orb mesh-orb-2"></div>
  <div class="mesh-orb mesh-orb-3"></div>
  <div class="container">
    <header>
      <button id="theme-toggle" class="theme-toggle"><i class="fas fa-moon"></i> <span>Тема</span></button>
      <!-- Beeline Logo SVG -->
      <div class="bee-logo-wrap">
        <svg viewBox="0 0 320 60" fill="none" xmlns="http://www.w3.org/2000/svg">
          <rect class="bee-pill" x="1" y="1" width="318" height="58" rx="29" fill="#FFD600"/>
          <text class="bee-text" x="160" y="40" text-anchor="middle" font-family="'Manrope','Inter',sans-serif" font-weight="800" font-size="28" fill="#000000" letter-spacing="-0.5">bee&#x2060;line</text>
          <circle cx="235" cy="16" r="3" fill="#000000"/>
        </svg>
      </div>
      <h1>Quality</h1>
    </header>
    <main id="main-content"></main>
    <footer><p>© Quality · Управление качеством</p></footer>
  </div>
  <script>
    const ORDER_KEYS={mainButtons:'main_buttons_order',directionLinks:(dir)=>`direction_${dir}_links_order`};
    const CI=['ci-indigo','ci-violet','ci-cyan','ci-emerald','ci-rose'];
    const defaultMainButtons=[
      {id:'service',text:'Сервис',icon:'fa-headset',href:'?dir=service'},
      {id:'sales',text:'Продажи',icon:'fa-arrow-trend-up',href:'?dir=sales'},
      {id:'retention',text:'Сохранение',icon:'fa-shield-heart',href:'?dir=retention'}
    ];
    const directionsDefaults={
      service:{name:"Сервис",icon:"fa-headset",links:[
        {text:"Регламенты оценки качества / Матрицы оценки / Карта фрода",href:"#",type:"link"},
        {text:"Описание скоринга / Новости / Процесс взаимодействия с учениками",href:"#",type:"link"},
        {text:"Стандарты обслуживания",href:"#",type:"link"},
        {text:"Чат бот",href:"#",type:"link"},
        {text:"Отчет по качеству",href:"#",type:"link"},
        {text:"Нормативы / Цели",href:"#",type:"link"},
        {text:"Контроль качества. Жалобы / Благодарности / Ошибки сотрудников",href:"#",type:"link"}
      ]},
      sales:{name:"Продажи",icon:"fa-arrow-trend-up",links:[
        {text:"Регламенты оценки качества / Матрицы оценки / Карта фрода",href:"#",type:"link"},
        {text:"Описание скоринга / Новости / Процесс взаимодействия с учениками",href:"#",type:"link"},
        {text:"Стандарты обслуживания",href:"#",type:"link"},
        {text:"Отчет по качеству",href:"#",type:"link"},
        {text:"Нормативы / Цели",href:"#",type:"link"},
        {text:"Чат бот",href:"#",type:"link"},
        {text:"Контроль качества. Жалобы / Благодарности / Ошибки сотрудников",href:"#",type:"link"}
      ]},
      retention:{name:"Сохранение",icon:"fa-shield-heart",links:[
        {text:"Регламенты оценки качества / Матрицы оценки / Карта фрода",href:"#",type:"link"},
        {text:"Описание скоринга / Новости / Процесс взаимодействия с учениками",href:"#",type:"link"},
        {text:"Стандарты обслуживания",href:"#",type:"link"},
        {text:"Чат бот",href:"#",type:"link"},
        {text:"Отчет по качеству",href:"#",type:"link"},
        {text:"Нормативы / Цели",href:"#",type:"link"},
        {text:"Контроль качества. Жалобы / Благодарности / Ошибки сотрудников",href:"#",type:"link"}
      ]}
    };
    function fixReportsLinks(dirs){Object.keys(dirs).forEach(k=>{if(dirs[k]&&Array.isArray(dirs[k].links)){dirs[k].links.forEach(link=>{if(link.text&&link.text.includes('Отчет по качеству')&&link.type!=='link'){link.type='link';if(!link.href||link.href==='reports'||link.href==='reports-sales'||link.href==='reports-retention'){link.href='#';}}});}});}
    function saveOrder(k,a){localStorage.setItem(k,JSON.stringify(a));}
    function loadOrder(k,da){const s=localStorage.getItem(k);if(s){try{const o=JSON.parse(s);if(Array.isArray(o)&&o.length===da.length)return o;}catch(e){}}return da.map((_,i)=>i);}
    function reorderArrayByIndices(a,oi){const r=[];oi.forEach(i=>{if(i>=0&&i<a.length)r.push(a[i]);});return r;}
    function createAdminIcons(p,ia,od,oe){if(!ia)return;const ei=document.createElement('i');ei.className='fas fa-pencil-alt admin-icon edit';ei.title='Редактировать';ei.onclick=oe;const di=document.createElement('i');di.className='fas fa-trash-alt admin-icon delete';di.title='Удалить';di.onclick=od;p.appendChild(ei);p.appendChild(di);}
    function createDragIcon(el,isL=false){const d=document.createElement('i');d.className=isL?'fas fa-grip-vertical link-drag-icon':'fas fa-grip-vertical drag-icon';d.title='Перетащить';el.appendChild(d);return d;}
    function saveMainButtons(){localStorage.setItem('mainButtons',JSON.stringify(mainButtons));}
    function addQualityTooltip(btn){const tt=document.createElement('div');tt.className='quality-tooltip';tt.innerHTML='<div class="tt-header"><span class="tt-icon"><i class="fas fa-lock"></i></span> Требуется доступ</div>Доступ предоставляется по заявке <strong>QLIK Stream 02.027_A. Customer Care Qlik Sense. Доступ к Стримам</strong> роль <strong>Пользователь</strong>';btn.appendChild(tt);}
    function renderMainPage(){const m=document.getElementById('main-content');if(feedbackStatsMode==='true'){renderFeedbackStatsPage();return;}if(dirKey&&directions[dirKey]){renderDirectionPage(dirKey);return;}renderMainButtonsPage();}
    function renderMainButtonsPage(){const m=document.getElementById('main-content');m.innerHTML=`<h2 class="main-title">Выбери направление</h2><div class="buttons-grid" id="main-buttons"></div>`;const g=document.getElementById('main-buttons');let cb=[...mainButtons];const so=loadOrder(ORDER_KEYS.mainButtons,cb);const ob=reorderArrayByIndices(cb,so);function rb(buttons){g.innerHTML='';const f=document.createDocumentFragment();buttons.forEach((btn,idx)=>{const c=document.createElement('div');c.className='direction-card animate-card';c.setAttribute('data-id',btn.id);const cc=CI[idx%CI.length];c.innerHTML=`<div class="card-icon ${cc}"><i class="fas ${btn.icon||'fa-star'}"></i></div><div class="direction-name">${btn.text}</div>`;createDragIcon(c);c.onclick=(e)=>{if(e.target.closest('.drag-icon')||e.target.closest('.admin-icon'))return;localStorage.setItem('currentDirection',btn.id);window.location.href=`?dir=${btn.id}`;};if(isAdmin){createAdminIcons(c,isAdmin,()=>{if(confirm('Удалить?')){mainButtons.splice(mainButtons.findIndex(b=>b.id===btn.id),1);saveMainButtons();rb(reorderArrayByIndices(mainButtons,loadOrder(ORDER_KEYS.mainButtons,mainButtons)));}},()=>{const nt=prompt('Название:',btn.text);if(nt)btn.text=nt;const nh=prompt('Ссылка:',btn.href);if(nh)btn.href=nh;saveMainButtons();rb(reorderArrayByIndices(mainButtons,loadOrder(ORDER_KEYS.mainButtons,mainButtons)));});}c.draggable=true;c.addEventListener('dragstart',(e)=>{e.dataTransfer.setData('text/plain',btn.id);c.classList.add('dragging');});c.addEventListener('dragend',()=>c.classList.remove('dragging'));c.addEventListener('dragover',(e)=>e.preventDefault());c.addEventListener('drop',(e)=>{e.preventDefault();const fi=mainButtons.findIndex(b=>b.id===e.dataTransfer.getData('text/plain'));const ti=mainButtons.findIndex(b=>b.id===btn.id);if(fi===ti)return;const mv=mainButtons.splice(fi,1)[0];mainButtons.splice(ti,0,mv);saveMainButtons();saveOrder(ORDER_KEYS.mainButtons,mainButtons.map((_,i)=>i));rb(mainButtons);});f.appendChild(c);});g.appendChild(f);if(isAdmin){const ac=document.createElement('div');ac.className='direction-card';ac.innerHTML='<div class="card-icon ci-indigo"><i class="fas fa-plus"></i></div><div class="direction-name">Добавить</div>';ac.onclick=()=>{const t=prompt('Название:');if(t){const id=t.toLowerCase().replace(/\s/g,'_');const h=prompt('Ссылка:',`?dir=${id}`);mainButtons.push({id,text:t,icon:'fa-folder',href:h||`?dir=${id}`});saveMainButtons();if(!directions[id])directions[id]={name:t,icon:'fa-folder',links:[]};localStorage.setItem('directionsData',JSON.stringify(directions));rb(reorderArrayByIndices(mainButtons,loadOrder(ORDER_KEYS.mainButtons,mainButtons)));}};g.appendChild(ac);}}rb(ob);addAdminFAB();}
    function renderDirectionPage(dk){const dir=directions[dk];const m=document.getElementById('main-content');m.innerHTML=`<button id="back-btn"><i class="fas fa-arrow-left"></i> Назад</button><h2 class="main-title"><i class="fas ${dir.icon||'fa-folder'} page-title-icon"></i> ${dir.name}</h2><div id="links-container" class="links-container"></div>`;const ct=document.getElementById('links-container');const so=loadOrder(ORDER_KEYS.directionLinks(dk),dir.links);const ol=reorderArrayByIndices(dir.links,so);function rl(links){ct.innerHTML='';const f=document.createDocumentFragment();links.forEach((link,idx)=>{const b=document.createElement('button');b.className='link-btn';b.setAttribute('data-idx',idx);b.innerHTML=`<i class="fas fa-arrow-up-right-from-square"></i> ${link.text}`;createDragIcon(b,true);if(link.text&&link.text.includes('Отчет по качеству')){addQualityTooltip(b);}b.onclick=(e)=>{if(e.target.closest('.link-drag-icon')||e.target.closest('.link-admin-icon')||e.target.closest('.quality-tooltip'))return;if(link.href&&link.href!=='#'){const cl=JSON.parse(localStorage.getItem('quality_clicks')||'[]');cl.push({timestamp:new Date().toISOString(),direction:dir.name,linkText:link.text});localStorage.setItem('quality_clicks',JSON.stringify(cl));window.open(link.href,'_blank');}else{alert('Ссылка временно недоступна');}};if(isAdmin){createAdminIcons(b,isAdmin,()=>{if(confirm('Удалить?')){dir.links.splice(dir.links.findIndex(l=>l.text===link.text),1);localStorage.setItem('directionsData',JSON.stringify(directions));rl(reorderArrayByIndices(dir.links,loadOrder(ORDER_KEYS.directionLinks(dk),dir.links)));}},()=>{const nt=prompt('Текст:',link.text);if(nt)link.text=nt;const nh=prompt('URL:',link.href);if(nh)link.href=nh;localStorage.setItem('directionsData',JSON.stringify(directions));rl(reorderArrayByIndices(dir.links,loadOrder(ORDER_KEYS.directionLinks(dk),dir.links)));});}b.draggable=true;b.addEventListener('dragstart',(e)=>{e.dataTransfer.setData('text/plain',idx);b.classList.add('dragging');});b.addEventListener('dragend',()=>b.classList.remove('dragging'));b.addEventListener('dragover',(e)=>e.preventDefault());b.addEventListener('drop',(e)=>{e.preventDefault();const fi=parseInt(e.dataTransfer.getData('text/plain'),10);if(fi===idx)return;const mv=dir.links.splice(fi,1)[0];dir.links.splice(idx,0,mv);localStorage.setItem('directionsData',JSON.stringify(directions));const oi=dir.links.map((_,i)=>i);saveOrder(ORDER_KEYS.directionLinks(dk),oi);rl(reorderArrayByIndices(dir.links,oi));});f.appendChild(b);});ct.appendChild(f);if(isAdmin){const ab=document.createElement('button');ab.className='link-btn';ab.innerHTML='<i class="fas fa-plus"></i> Добавить';ab.onclick=()=>{const t=prompt('Текст:');if(t){const h=prompt('URL:','#');dir.links.push({text:t,href:h||'#',type:'link'});localStorage.setItem('directionsData',JSON.stringify(directions));rl(reorderArrayByIndices(dir.links,loadOrder(ORDER_KEYS.directionLinks(dk),dir.links)));}};ct.appendChild(ab);}}rl(ol);addFeedbackSection(dir.name);document.getElementById('back-btn').onclick=()=>window.location.href=window.location.pathname;}
    function renderFeedbackStatsPage(){const fb=JSON.parse(localStorage.getItem('quality_feedback')||'[]');const m=document.getElementById('main-content');const rc={1:0,2:0,3:0,4:0,5:0};fb.forEach(f=>{if(f.rating>=1&&f.rating<=5)rc[f.rating]++;});const t=fb.length;function prd(ds){if(!ds)return null;const p=ds.split(',')[0].split('.');if(p.length!==3)return null;return new Date(parseInt(p[2],10),parseInt(p[1],10)-1,parseInt(p[0],10));}const mn=["Январь","Февраль","Март","Апрель","Май","Июнь","Июль","Август","Сентябрь","Октябрь","Ноябрь","Декабрь"];const ys=new Set();const ms=new Set();fb.forEach(f=>{const d=prd(f.timestamp);if(d){ys.add(d.getFullYear());ms.add(d.getMonth());}});const sy=Array.from(ys).sort((a,b)=>b-a);const sm=Array.from(ms).sort((a,b)=>a-b);const md={};fb.forEach(f=>{const d=prd(f.timestamp);if(d){const ym=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;if(!md[ym])md[ym]={month:mn[d.getMonth()]+' '+d.getFullYear(),ratings:[0,0,0,0,0],count:0,avgRating:0};md[ym].ratings[f.rating-1]++;md[ym].count++;md[ym].avgRating+=f.rating;}});const sym=Object.keys(md).sort();const cl=sym.map(ym=>md[ym].month);const cd={labels:cl,datasets:[{label:'Средний балл',data:sym.map(ym=>{const d=md[ym];return d.count>0?parseFloat((d.avgRating/d.count).toFixed(2)):0;}),backgroundColor:'#6366F1',borderColor:'#6366F1',borderWidth:2,tension:0.4,yAxisID:'y'},{label:'Количество отзывов',data:sym.map(ym=>md[ym].count),backgroundColor:'rgba(6,182,212,0.2)',borderColor:'#06B6D4',borderWidth:2,type:'bar',yAxisID:'y1'}]};const fs='background:rgba(255,255,255,0.03);backdrop-filter:blur(8px);border-radius:9999px;padding:8px 14px;border:1px solid rgba(255,255,255,0.06);font-family:inherit;font-weight:500;font-size:0.85rem;color:inherit;';m.innerHTML=`<button id="back-btn"><i class="fas fa-arrow-left"></i> Назад</button><h2 class="main-title"><i class="fas fa-chart-column page-title-icon"></i> Статистика обратной связи</h2><div class="analytics-container"><div class="stats-cards"><div class="stat-card"><div class="stat-number">${t}</div><div class="stat-label">Всего отзывов</div></div>${[1,2,3,4,5].map(r=>'<div class="stat-card"><div class="stat-number">'+rc[r]+'</div><div class="stat-label">'+'★'.repeat(r)+'☆'.repeat(5-r)+' ('+(t?((rc[r]/t)*100).toFixed(1):0)+'%)</div></div>').join('')}</div><div class="tabs"><button class="tab-button active" data-tab="reviews"><i class="fas fa-comment-dots" style="margin-right:6px;"></i>Отзывы</button><button class="tab-button" data-tab="rating-chart"><i class="fas fa-chart-pie" style="margin-right:6px;"></i>Распределение</button><button class="tab-button" data-tab="rating-trend"><i class="fas fa-chart-column" style="margin-right:6px;"></i>Динамика</button></div><div id="reviews-tab" class="tab-content active"><div id="ff" style="display:flex;gap:10px;margin-bottom:16px;flex-wrap:wrap;"><select id="yf" style="${fs}"><option value="">Все годы</option>${sy.map(y=>'<option value="'+y+'">'+y+'</option>').join('')}</select><select id="mf" style="${fs}"><option value="">Все месяцы</option>${sm.map(mo=>'<option value="'+mo+'">'+mn[mo]+'</option>').join('')}</select><select id="rf" style="${fs}"><option value="">Оценка</option><option value="5">★★★★★</option><option value="4">★★★★☆</option><option value="3">★★★☆☆</option><option value="2">★★☆☆☆</option><option value="1">★☆☆☆☆</option></select></div><div class="data-table-wrapper"><table id="ft"><thead><th>Дата</th><th>Направление</th><th>Оценка</th><th>Отзыв</th></thead><tbody></tbody></table></div><button class="export-btn" id="efb" style="margin-top:12px;"><i class="fas fa-download"></i> Скачать отзывы</button></div><div id="rating-chart-tab" class="tab-content"><div style="height:400px;"><canvas id="fc"></canvas></div></div><div id="rating-trend-tab" class="tab-content"><div style="height:400px;"><canvas id="rtc"></canvas></div></div></div>`;function rft(){const y=document.getElementById('yf').value;const mo=document.getElementById('mf').value;const ra=document.getElementById('rf').value;let fl=[...fb];if(y)fl=fl.filter(f=>{const d=prd(f.timestamp);return d&&d.getFullYear()==y;});if(mo!=="")fl=fl.filter(f=>{const d=prd(f.timestamp);return d&&d.getMonth()==mo;});if(ra)fl=fl.filter(f=>f.rating===parseInt(ra));const tb=document.querySelector('#ft tbody');if(tb){tb.innerHTML='';fl.forEach(f=>{const d=prd(f.timestamp);let fd='—';if(d)fd=String(d.getDate()).padStart(2,'0')+'.'+String(d.getMonth()+1).padStart(2,'0')+'.'+d.getFullYear();const tr=document.createElement('tr');tr.innerHTML='<td>'+fd+'</td><td>'+f.direction+'</td><td>'+'★'.repeat(f.rating)+'☆'.repeat(5-f.rating)+'</td><td>'+(f.comment||'')+'</td>';tb.appendChild(tr);});}}function rrc(){const ctx=document.getElementById('fc').getContext('2d');if(ctx){if(window.fbC)window.fbC.destroy();window.fbC=new Chart(ctx,{type:'bar',data:{labels:['★☆☆☆☆','★★☆☆☆','★★★☆☆','★★★★☆','★★★★★'],datasets:[{label:'Отзывы',data:[rc[1],rc[2],rc[3],rc[4],rc[5]],backgroundColor:'#6366F1',borderRadius:6}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'top',labels:{color:'#A1A1AA'}}},scales:{y:{beginAtZero:true,ticks:{color:'#71717A'},grid:{color:'rgba(255,255,255,0.04)'}},x:{ticks:{color:'#71717A'},grid:{color:'rgba(255,255,255,0.04)'}}}}});}}function rrtc(){const ctx=document.getElementById('rtc').getContext('2d');if(ctx&&cl.length>0){if(window.fTC)window.fTC.destroy();window.fTC=new Chart(ctx,{type:'line',data:cd,options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'top',labels:{color:'#A1A1AA'}}},scales:{y:{type:'linear',display:true,position:'left',beginAtZero:true,max:5,ticks:{color:'#71717A'},grid:{color:'rgba(255,255,255,0.04)'},title:{display:true,text:'Средний балл',color:'#71717A'}},y1:{type:'linear',display:true,position:'right',beginAtZero:true,ticks:{color:'#71717A'},grid:{color:'rgba(255,255,255,0.04)'},title:{display:true,text:'Отзывы',color:'#71717A'}}}}});}}document.getElementById('yf').addEventListener('change',rft);document.getElementById('mf').addEventListener('change',rft);document.getElementById('rf').addEventListener('change',rft);document.getElementById('efb').addEventListener('click',()=>{const y=document.getElementById('yf').value;const mo=document.getElementById('mf').value;const ra=document.getElementById('rf').value;let fl=[...fb];if(y)fl=fl.filter(f=>{const d=prd(f.timestamp);return d&&d.getFullYear()==y;});if(mo!=="")fl=fl.filter(f=>{const d=prd(f.timestamp);return d&&d.getMonth()==mo;});if(ra)fl=fl.filter(f=>f.rating===parseInt(ra));const ws=XLSX.utils.json_to_sheet(fl.map(f=>({Дата:f.timestamp,Направление:f.direction,Оценка:f.rating+' звезд',Отзыв:f.comment})));const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Отзывы');XLSX.writeFile(wb,'отзывы_'+new Date().toISOString().slice(0,10)+'.xlsx');});rft();rrc();rrtc();const tabs=document.querySelectorAll('.tab-button');tabs.forEach(tab=>{tab.addEventListener('click',()=>{tabs.forEach(x=>x.classList.remove('active'));tab.classList.add('active');document.querySelectorAll('.tab-content').forEach(c=>c.classList.remove('active'));document.getElementById(tab.dataset.tab+'-tab').classList.add('active');if(tab.dataset.tab==='rating-trend')rrtc();else if(tab.dataset.tab==='rating-chart')rrc();});});document.getElementById('back-btn').onclick=()=>window.location.href=window.location.pathname;}
    function addFeedbackSection(dn){const m=document.getElementById('main-content');const ex=document.querySelector('.feedback-modern');if(ex)ex.remove();const d=document.createElement('div');d.className='feedback-modern';d.innerHTML=`<h3><i class="fas fa-sparkles" style="color:var(--bee-yellow)"></i> Обратная связь</h3><p style="margin-bottom:6px;opacity:0.5;font-size:0.78rem;">Твоё мнение помогает становиться лучше</p><div class="stars-modern" id="stars-modern">${[1,2,3,4,5].map(i=>'<i class="far fa-star" data-val="'+i+'"></i>').join('')}</div><textarea id="fb-text" rows="3" placeholder="Комментарий или предложение..."></textarea><button class="btn-primary" id="send-fb"><i class="fas fa-paper-plane"></i> Отправить</button>`;m.appendChild(d);let rating=0;const stars=d.querySelectorAll('.stars-modern i');stars.forEach(s=>{s.addEventListener('click',()=>{rating=parseInt(s.dataset.val);stars.forEach(ss=>{if(parseInt(ss.dataset.val)<=rating)ss.className='fas fa-star active';else ss.className='far fa-star';});});s.addEventListener('mouseenter',()=>{const hv=parseInt(s.dataset.val);stars.forEach(ss=>{if(parseInt(ss.dataset.val)<=hv)ss.style.color='#FFD600';else ss.style.color='var(--f-zinc-700)';});});s.addEventListener('mouseleave',()=>{stars.forEach(ss=>{if(parseInt(ss.dataset.val)<=rating)ss.style.color='#FFD600';else ss.style.color='var(--f-zinc-700)';});});});d.querySelector('#send-fb').onclick=()=>{if(rating===0)return alert('Поставьте оценку');const fb={timestamp:new Date().toLocaleString(),direction:dn,rating:rating,comment:d.querySelector('#fb-text').value||'Без комментария'};const list=JSON.parse(localStorage.getItem('quality_feedback')||'[]');list.push(fb);localStorage.setItem('quality_feedback',JSON.stringify(list));alert('Спасибо за отзыв!');rating=0;stars.forEach(s=>{s.className='far fa-star';s.style.color='var(--f-zinc-700)';});d.querySelector('#fb-text').value='';};}
    function addAdminFAB(){const b=document.createElement('button');b.className='admin-fab';b.innerHTML=isAdmin?'<i class="fas fa-sign-out-alt"></i>':'<i class="fas fa-lock"></i>';b.onclick=()=>{if(isAdmin){localStorage.removeItem('adminActive');location.reload();}else{const p=prompt('Пароль:');if(p==='beeline2025'){localStorage.setItem('adminActive','true');location.reload();}else alert('Неверный пароль');}};document.body.appendChild(b);if(isAdmin){const sb=document.createElement('button');sb.className='admin-fab';sb.style.bottom='90px';sb.style.background='linear-gradient(135deg,#EF4444,#DC2626)';sb.innerHTML='<i class="fas fa-chart-column"></i>';sb.onclick=()=>{window.location.href='?feedback=true';};document.body.appendChild(sb);}}
    const themeToggle=document.getElementById('theme-toggle');if(localStorage.getItem('dark-theme')==='true')document.body.classList.add('dark');themeToggle.addEventListener('click',()=>{document.body.classList.toggle('dark');const isD=document.body.classList.contains('dark');localStorage.setItem('dark-theme',isD);themeToggle.querySelector('i').className=isD?'fas fa-moon':'fas fa-sun';themeToggle.querySelector('span').textContent='Тема';});
    let directions=JSON.parse(JSON.stringify(directionsDefaults));let mainButtons=[...defaultMainButtons];
    const savedData=localStorage.getItem('directionsData');if(savedData){try{const l=JSON.parse(savedData);Object.keys(directions).forEach(k=>{if(l[k]&&Array.isArray(l[k].links))directions[k].links=l[k].links;});}catch(e){}}
    fixReportsLinks(directions);localStorage.setItem('directionsData',JSON.stringify(directions));
    const savedMB=localStorage.getItem('mainButtons');if(savedMB){try{const p=JSON.parse(savedMB);if(Array.isArray(p)&&p.length>0)mainButtons=p;}catch(e){}}
    const isAdmin=localStorage.getItem('adminActive')==='true';const urlParams=new URLSearchParams(window.location.search);const dirKey=urlParams.get('dir');const feedbackStatsMode=urlParams.get('feedback');
    renderMainPage();
  </script>
</body>
</html>
