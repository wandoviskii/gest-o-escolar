<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>EduFlow Pro - Gestão Escolar & Diário de Classe</title>
  <!-- Bibliotecas ExcelJS e FileSaver para gerar .XLSX com estilos no navegador -->
  <script src="https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.3.0/exceljs.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js"></script>

  <style>
    :root {
      --primary: #4f46e5;
      --primary-hover: #4338ca;
      --primary-light: #eef2ff;
      --primary-glow: rgba(79, 70, 229, 0.15);
      
      --secondary: #0ea5e9;
      --secondary-light: #e0f2fe;
      
      --success: #10b981;
      --success-light: #ecfdf5;
      --success-text: #047857;
      
      --warning: #f59e0b;
      --warning-light: #fffbeb;
      --warning-text: #b45309;
      
      --danger: #ef4444;
      --danger-light: #fef2f2;
      --danger-text: #b91c1c;
      
      --bg: #f8fafc;
      --card-bg: rgba(255, 255, 255, 0.95);
      --text: #0f172a;
      --text-muted: #64748b;
      --border: #e2e8f0;
      --border-subtle: #f1f5f9;
      
      --radius-sm: 8px;
      --radius: 14px;
      --radius-lg: 20px;
      
      --shadow-sm: 0 1px 2px 0 rgba(0, 0, 0, 0.05);
      --shadow-md: 0 4px 6px -1px rgba(0, 0, 0, 0.07), 0 2px 4px -2px rgba(0, 0, 0, 0.05);
      --shadow-lg: 0 10px 25px -3px rgba(15, 23, 42, 0.08), 0 4px 6px -4px rgba(15, 23, 42, 0.04);
      --shadow-glow: 0 0 20px rgba(79, 70, 229, 0.15);
    }

    * {
      box-sizing: border-box;
      margin: 0;
      padding: 0;
      font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    }

    body {
      background-color: var(--bg);
      color: var(--text);
      min-height: 100vh;
      display: flex;
      flex-direction: column;
      background-image: radial-gradient(#e2e8f0 1px, transparent 1px);
      background-size: 24px 24px;
    }

    /* HEADER GLASSMORPHISM */
    header {
      background: rgba(255, 255, 255, 0.85);
      backdrop-filter: blur(12px);
      -webkit-backdrop-filter: blur(12px);
      border-bottom: 1px solid var(--border);
      position: sticky;
      top: 0;
      z-index: 50;
      padding: 0.85rem 2rem;
      display: flex;
      justify-content: space-between;
      align-items: center;
      flex-wrap: wrap;
      gap: 1.2rem;
    }

    .school-info-header {
      display: flex;
      align-items: center;
      gap: 1rem;
      cursor: pointer;
      padding: 0.35rem 0.6rem;
      border-radius: var(--radius);
      transition: all 0.2s ease;
    }

    .school-info-header:hover {
      background: var(--primary-light);
    }

    .school-logo-container {
      width: 48px;
      height: 48px;
      border-radius: 12px;
      background: linear-gradient(135deg, var(--primary) 0%, var(--secondary) 100%);
      color: white;
      display: flex;
      align-items: center;
      justify-content: center;
      box-shadow: var(--shadow-sm);
      overflow: hidden;
      flex-shrink: 0;
    }

    .school-logo-container img {
      width: 100%;
      height: 100%;
      object-fit: cover;
    }

    .school-titles h1 {
      font-size: 1.15rem;
      font-weight: 700;
      color: var(--text);
      letter-spacing: -0.3px;
      display: flex;
      align-items: center;
      gap: 0.5rem;
    }

    .school-titles p {
      font-size: 0.8rem;
      color: var(--text-muted);
      font-weight: 500;
    }

    .header-actions {
      display: flex;
      align-items: center;
      gap: 0.8rem;
      flex-wrap: wrap;
    }

    .role-selector-bar, .class-selector-bar {
      display: flex;
      align-items: center;
      gap: 0.5rem;
      background: #f1f5f9;
      padding: 0.35rem 0.85rem;
      border-radius: 9999px;
      border: 1px solid var(--border);
    }

    .role-selector-bar label, .class-selector-bar label {
      font-size: 0.75rem;
      font-weight: 700;
      color: var(--text-muted);
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }

    .role-selector-bar select, .class-selector-bar select {
      background: transparent;
      color: var(--primary);
      border: none;
      font-weight: 700;
      font-size: 0.88rem;
      cursor: pointer;
      outline: none;
    }

    /* NAVIGATION */
    nav {
      background: white;
      border-bottom: 1px solid var(--border);
      padding: 0.4rem 2rem;
      display: flex;
      gap: 0.5rem;
      overflow-x: auto;
    }

    .tab-btn {
      background: transparent;
      border: none;
      padding: 0.65rem 1.2rem;
      border-radius: var(--radius-sm);
      font-weight: 600;
      font-size: 0.88rem;
      color: var(--text-muted);
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      gap: 0.55rem;
      transition: all 0.2s ease;
      white-space: nowrap;
    }

    .tab-btn:hover {
      color: var(--primary);
      background: var(--primary-light);
    }

    .tab-btn.active {
      color: var(--primary);
      background: var(--primary-light);
      font-weight: 700;
    }

    .tab-btn.role-restricted {
      display: none !important;
    }

    /* MAIN CONTENT */
    main {
      flex: 1;
      padding: 2rem;
      max-width: 1400px;
      margin: 0 auto;
      width: 100%;
    }

    .tab-content {
      display: none;
      animation: modernFade 0.25s ease-out;
    }

    .tab-content.active {
      display: block;
    }

    @keyframes modernFade {
      from { opacity: 0; transform: translateY(6px); }
      to { opacity: 1; transform: translateY(0); }
    }

    /* CARDS */
    .card {
      background: var(--card-bg);
      backdrop-filter: blur(8px);
      border-radius: var(--radius);
      border: 1px solid var(--border);
      box-shadow: var(--shadow-sm);
      padding: 1.6rem;
      margin-bottom: 1.6rem;
      transition: box-shadow 0.2s ease;
    }

    .card:hover {
      box-shadow: var(--shadow-md);
    }

    .card-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 1.4rem;
      flex-wrap: wrap;
      gap: 0.8rem;
    }

    .card-title {
      font-size: 1.2rem;
      font-weight: 700;
      letter-spacing: -0.3px;
      color: var(--text);
      display: flex;
      align-items: center;
      gap: 0.5rem;
    }

    .card-subtitle {
      font-size: 0.85rem;
      color: var(--text-muted);
      margin-top: 0.2rem;
    }

    .grid-2 {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(340px, 1fr));
      gap: 1.6rem;
    }

    /* FORMS */
    .form-group {
      margin-bottom: 1rem;
    }

    .form-group label {
      display: block;
      font-size: 0.82rem;
      font-weight: 600;
      color: var(--text-muted);
      margin-bottom: 0.4rem;
      text-transform: uppercase;
      letter-spacing: 0.3px;
    }

    .form-control {
      width: 100%;
      padding: 0.65rem 0.95rem;
      border: 1px solid var(--border);
      border-radius: var(--radius-sm);
      font-size: 0.92rem;
      background: #ffffff;
      color: var(--text);
      transition: all 0.2s;
    }

    .form-control:focus {
      outline: none;
      border-color: var(--primary);
      box-shadow: 0 0 0 3px var(--primary-glow);
    }

    /* BUTTONS */
    .btn {
      padding: 0.6rem 1.1rem;
      border-radius: var(--radius-sm);
      font-weight: 600;
      font-size: 0.88rem;
      cursor: pointer;
      border: none;
      display: inline-flex;
      align-items: center;
      gap: 0.45rem;
      transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
      box-shadow: var(--shadow-sm);
    }

    .btn:active {
      transform: scale(0.98);
    }

    .btn-primary { background: var(--primary); color: white; }
    .btn-primary:hover { background: var(--primary-hover); box-shadow: 0 4px 12px rgba(79, 70, 229, 0.25); }

    .btn-success { background: var(--success); color: white; }
    .btn-success:hover { background: #059669; box-shadow: 0 4px 12px rgba(16, 185, 129, 0.25); }

    .btn-danger { background: var(--danger); color: white; }
    .btn-danger:hover { background: #dc2626; }

    .btn-secondary { background: #f1f5f9; color: var(--text); border: 1px solid var(--border); }
    .btn-secondary:hover { background: #e2e8f0; }

    .btn-sm { padding: 0.35rem 0.7rem; font-size: 0.8rem; border-radius: 6px; }

    /* TABLES */
    .table-container {
      overflow-x: auto;
      border-radius: var(--radius-sm);
      border: 1px solid var(--border);
      background: white;
    }

    table {
      width: 100%;
      border-collapse: separate;
      border-spacing: 0;
      font-size: 0.9rem;
    }

    th {
      background-color: #f8fafc;
      font-weight: 700;
      color: var(--text-muted);
      text-transform: uppercase;
      font-size: 0.74rem;
      letter-spacing: 0.6px;
      padding: 0.85rem 1rem;
      border-bottom: 1px solid var(--border);
    }

    td {
      padding: 0.85rem 1rem;
      border-bottom: 1px solid var(--border-subtle);
      vertical-align: middle;
      color: var(--text);
    }

    tr:last-child td {
      border-bottom: none;
    }

    tr:hover td {
      background-color: #f8fafc;
    }

    /* BADGES */
    .badge {
      display: inline-flex;
      align-items: center;
      gap: 0.3rem;
      padding: 0.3rem 0.75rem;
      border-radius: 9999px;
      font-size: 0.75rem;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.3px;
    }

    .badge-success { background-color: var(--success-light); color: var(--success-text); }
    .badge-danger { background-color: var(--danger-light); color: var(--danger-text); }
    .badge-warning { background-color: var(--warning-light); color: var(--warning-text); }
    .badge-info { background-color: var(--primary-light); color: var(--primary); }

    /* ATTENDANCE CHECKBOXES */
    .checkbox-freq-label {
      display: inline-flex;
      align-items: center;
      gap: 0.5rem;
      font-size: 0.85rem;
      font-weight: 600;
      cursor: pointer;
      user-select: none;
      padding: 0.45rem 0.8rem;
      border-radius: 8px;
      border: 1px solid var(--border);
      background: #ffffff;
      transition: all 0.2s ease;
    }

    .checkbox-freq-label:hover {
      border-color: #cbd5e1;
      background: #f8fafc;
    }

    .checkbox-freq-label.is-falta {
      background: #fef2f2;
      border-color: #fca5a5;
      color: var(--danger-text);
      font-weight: 700;
    }

    .checkbox-freq-label.is-justificada {
      background: #fffbeb;
      border-color: #fde68a;
      color: var(--warning-text);
      font-weight: 700;
    }

    .checkbox-freq {
      width: 18px;
      height: 18px;
      cursor: pointer;
      border-radius: 4px;
    }

    .checkbox-freq.danger-box { accent-color: var(--danger); }
    .checkbox-freq.warning-box { accent-color: var(--warning); }

    .input-nota {
      width: 68px;
      text-align: center;
      padding: 0.4rem;
      border-radius: 8px;
      border: 1px solid var(--border);
      font-weight: 700;
      background: #f8fafc;
      transition: all 0.2s;
    }

    .input-nota:focus {
      background: white;
      border-color: var(--primary);
      box-shadow: 0 0 0 3px var(--primary-glow);
    }

    /* STATS KPI */
    .stat-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
      gap: 1.2rem;
      margin-bottom: 1.8rem;
    }

    .stat-card {
      background: white;
      border-radius: var(--radius);
      padding: 1.25rem 1.4rem;
      border: 1px solid var(--border);
      box-shadow: var(--shadow-sm);
      display: flex;
      align-items: center;
      gap: 1.2rem;
      transition: transform 0.2s, box-shadow 0.2s;
    }

    .stat-card:hover {
      transform: translateY(-2px);
      box-shadow: var(--shadow-md);
    }

    .stat-icon {
      width: 48px;
      height: 48px;
      border-radius: 12px;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 1.4rem;
    }

    .stat-val {
      font-size: 1.65rem;
      font-weight: 800;
      color: var(--text);
      letter-spacing: -0.5px;
      line-height: 1.1;
    }

    .stat-label {
      font-size: 0.78rem;
      color: var(--text-muted);
      text-transform: uppercase;
      font-weight: 700;
      letter-spacing: 0.5px;
    }

    .points-tag {
      font-size: 1.05rem;
      font-weight: 800;
      padding: 0.25rem 0.75rem;
      border-radius: 9999px;
      background: var(--primary-light);
      color: var(--primary);
      display: inline-block;
    }

    /* BIMESTRE TABS */
    .bimestre-tabs {
      display: flex;
      gap: 0.5rem;
      margin-bottom: 1.2rem;
      overflow-x: auto;
      padding-bottom: 0.3rem;
    }

    .bim-btn {
      padding: 0.5rem 1rem;
      border: 1px solid var(--border);
      background: white;
      border-radius: 9999px;
      font-weight: 600;
      cursor: pointer;
      font-size: 0.85rem;
      color: var(--text-muted);
      transition: all 0.2s;
    }

    .bim-btn:hover {
      border-color: var(--primary);
      color: var(--primary);
    }

    .bim-btn.active {
      background: var(--primary);
      color: white;
      border-color: var(--primary);
      box-shadow: 0 2px 8px rgba(79, 70, 229, 0.3);
    }

    /* COMPONENTES DO MURAL DE AVISOS, EVENTOS E CHAT DE RESPONSÁVEIS */
    .notice-card {
      background: white;
      border: 1px solid var(--border);
      border-radius: var(--radius);
      padding: 1.2rem;
      margin-bottom: 1rem;
      box-shadow: var(--shadow-sm);

    }

    .notice-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 0.6rem;
    }

    .notice-title {
      font-size: 1.05rem;
      font-weight: 700;
      color: var(--text);
    }

    .notice-meta {
      font-size: 0.78rem;
      color: var(--text-muted);
    }

    .notice-body {
      font-size: 0.9rem;
      color: #334155;
      line-height: 1.5;
    }

    .chat-box {
      border: 1px solid var(--border);
      border-radius: var(--radius);
      background: white;
      height: 380px;
      display: flex;
      flex-direction: column;
    }

    .chat-messages {
      flex: 1;
      padding: 1rem;
      overflow-y: auto;
      display: flex;
      flex-direction: column;
      gap: 0.8rem;
    }

    .chat-msg {
      max-width: 80%;
      padding: 0.7rem 1rem;
      border-radius: 12px;
      font-size: 0.88rem;
      line-height: 1.4;
    }

    .chat-msg.sent {
      align-self: flex-end;
      background: var(--primary);
      color: white;
      border-bottom-right-radius: 2px;
    }

    .chat-msg.received {
      align-self: flex-start;
      background: #f1f5f9;
      color: var(--text);
      border-bottom-left-radius: 2px;
    }

    .chat-input-area {
      padding: 0.8rem;
      border-top: 1px solid var(--border);
      display: flex;
      gap: 0.6rem;
    }

    /* MODAL */
    .modal-overlay {
      position: fixed;
      top: 0; left: 0; right: 0; bottom: 0;
      background: rgba(15, 23, 42, 0.4);
      backdrop-filter: blur(4px);
      display: none;
      align-items: center;
      justify-content: center;
      z-index: 1000;
      animation: modalFadeIn 0.2s ease-out;
    }

    .modal-overlay.open { display: flex; }

    @keyframes modalFadeIn {
      from { opacity: 0; transform: scale(0.96); }
      to { opacity: 1; transform: scale(1); }
    }

    .modal-card {
      background: white;
      border-radius: var(--radius-lg);
      padding: 2rem;
      max-width: 500px;
      width: 90%;
      box-shadow: var(--shadow-lg);
    }

    /* REGRAS DE IMPRESSÃO & EXPORTAÇÃO PDF */
    @page {
      size: A4 landscape;
      margin: 10mm 12mm 10mm 12mm;
    }

    @media print {
      html, body {
        width: 100% !important;
        height: auto !important;
        margin: 0 !important;
        padding: 0 !important;
        background: #ffffff !important;
        background-image: none !important;
        color: #000000 !important;
        font-size: 10pt !important;
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
      }

      header, nav, .no-print, .btn, input, select, .bimestre-tabs, .stat-grid, .modal-overlay {
        display: none !important;
      }

      main {
        padding: 0 !important;
        margin: 0 !important;
        max-width: 100% !important;
        width: 100% !important;
      }

      .card {
        padding: 0 !important;
        margin: 0 !important;
        box-shadow: none !important;
        border: none !important;
        background: transparent !important;
      }

      .tab-content {
        display: none !important;
      }

      #tab-relatorios {
        display: block !important;
      }

      .print-header {
        display: flex !important;
        align-items: center;
        gap: 1.2rem;
        margin-bottom: 1.2rem;
        border-bottom: 2px solid #0f172a;
        padding-bottom: 0.8rem;
      }

      .print-logo {
        width: 65px;
        height: 65px;
        object-fit: contain;
      }

      .table-container {
        overflow: visible !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 4px !important;
        width: 100% !important;
      }

      table {
        width: 100% !important;
        table-layout: auto !important;
        font-size: 8.5pt !important;
        border-collapse: collapse !important;
      }

      thead {
        display: table-header-group !important;
      }

      tr {
        page-break-inside: avoid !important;
        break-inside: avoid !important;
      }

      th, td {
        padding: 0.45rem 0.4rem !important;
        border: 1px solid #e2e8f0 !important;
        white-space: normal !important;
        word-break: normal !important;
      }

      th {
        background-color: #f1f5f9 !important;
        color: #0f172a !important;
        font-size: 7.5pt !important;
        font-weight: 800 !important;
        text-align: center !important;
      }

      th:nth-child(2), td:nth-child(2) {
        text-align: left !important;
      }

      .badge {
        padding: 0.15rem 0.45rem !important;
        font-size: 7pt !important;
        border: 1px solid #94a3b8 !important;
      }
    }

    .print-header { display: none; }
  </style>
</head>
<body>

  <!-- HEADER -->
  <header>
    <div class="school-info-header" onclick="openSchoolModal()" title="Clique para editar instituição e logo">
      <div class="school-logo-container" id="headerLogoContainer">
        <span id="headerLogoPlaceholder" style="font-size: 1.6rem;">🏫</span>
        <img id="headerLogoImg" style="display: none;" alt="Logo da Escola" />
      </div>
      <div class="school-titles">
        <h1 id="headerSchoolName">Escola Municipal de Ensino Fundamental <span style="font-size: 0.8rem; opacity: 0.6;">✏️</span></h1>
        <p>Diário de Classe &bull; Sistema de Gestão Escolar</p>
      </div>
    </div>

    <div class="header-actions">
      <!-- SELETOR DE PERFIL DE ACESSO -->
      <div class="role-selector-bar">
        <label for="currentRoleSelect">Perfil:</label>
        <select id="currentRoleSelect" onchange="changeActiveRole(this.value)">
          <option value="direcao">👑 Direção / Coordenação</option>
          <option value="professor" selected>👨‍🏫 Professor(a)</option>
          <option value="aluno">👨‍🎓 Aluno(a)</option>
          <option value="pais">👨‍👩‍👧 Pais / Responsáveis</option>
        </select>
      </div>

      <!-- SELETOR DE TURMA -->
      <div class="class-selector-bar">
        <label for="currentTurmaSelect">Turma:</label>
        <select id="currentTurmaSelect" onchange="changeActiveTurma(this.value)"></select>
      </div>

      <button class="btn btn-secondary btn-sm no-print admin-only" onclick="openSchoolModal()">⚙️ Configurar</button>
    </div>
  </header>

  <!-- NAVIGATION TABS -->
  <nav class="no-print">
    <button class="tab-btn active" id="tabNav-turmas" onclick="switchTab('tab-turmas', this)">
      <span>👥</span> Turmas & Alunos
    </button>
    <button class="tab-btn" id="tabNav-frequencia" onclick="switchTab('tab-frequencia', this)">
      <span>📅</span> Frequência / Chamada
    </button>
    <button class="tab-btn" id="tabNav-pontos" onclick="switchTab('tab-pontos', this)">
      <span>⭐</span> Pontos de Atividades
    </button>
    <button class="tab-btn" id="tabNav-notas" onclick="switchTab('tab-notas', this)">
      <span>📝</span> Notas Bimestrais
    </button>
    <button class="tab-btn" id="tabNav-relatorios" onclick="switchTab('tab-relatorios', this)">
      <span>📊</span> Relatório Geral
    </button>
    <button class="tab-btn" id="tabNav-comunicacao" onclick="switchTab('tab-comunicacao', this)">
      <span>📢</span> Mural & Comunicação
    </button>
    <button class="tab-btn" id="tabNav-agenda" onclick="switchTab('tab-agenda', this)">
      <span>🗓️</span> Agenda de Eventos
    </button>
    <button class="tab-btn" id="tabNav-canal" onclick="switchTab('tab-canal', this)">
      <span>💬</span> Canal com Responsáveis
    </button>
  </nav>

  <main>

    <!-- TAB 1: TURMAS E ALUNOS -->
    <section id="tab-turmas" class="tab-content active">
      <div class="grid-2">
        <div class="card">
          <div class="card-header">
            <div>
              <h2 class="card-title">🏫 Turmas Registradas</h2>
              <p class="card-subtitle">Cadastre e alterne entre as turmas ativas</p>
            </div>
          </div>
          <form onsubmit="handleAddTurma(event)" class="admin-teacher-only" style="display: flex; gap: 0.6rem; margin-bottom: 1.4rem;">
            <input type="text" id="newTurmaInput" class="form-control" placeholder="Ex: 6º Ano A, 9º Ano B..." required />
            <button type="submit" class="btn btn-primary">+ Criar</button>
          </form>

          <div class="table-container">
            <table id="tableTurmas">
              <thead>
                <tr>
                  <th>Nome da Turma</th>
                  <th>Alunos</th>
                  <th style="text-align: right;">Ações</th>
                </tr>
              </thead>
              <tbody></tbody>
            </table>
          </div>
        </div>

        <div class="card">
          <div class="card-header">
            <div>
              <h2 class="card-title">👨‍🎓 Alunos da Turma: <span class="current-turma-name" style="color: var(--primary);"></span></h2>
              <p class="card-subtitle">Matricule novos alunos na turma selecionada</p>
            </div>
          </div>
          <form onsubmit="handleAddAluno(event)" class="admin-teacher-only" style="display: flex; gap: 0.6rem; margin-bottom: 1.4rem;">
            <input type="text" id="newAlunoInput" class="form-control" placeholder="Nome completo do aluno" required />
            <button type="submit" class="btn btn-success">+ Matricular</button>
          </form>

          <div class="table-container">
            <table id="tableAlunos">
              <thead>
                <tr>
                  <th style="width: 50px;">Nº</th>
                  <th>Nome do Aluno</th>
                  <th style="text-align: right;">Ações</th>
                </tr>
              </thead>
              <tbody></tbody>
            </table>
          </div>
        </div>
      </div>
    </section>

    <!-- TAB 2: FREQUÊNCIA (CHECKBOXES) -->
    <section id="tab-frequencia" class="tab-content">
      <div class="card">
        <div class="card-header">
          <div>
            <h2 class="card-title">📅 Diário de Chamada - <span class="current-turma-name"></span></h2>
            <p class="card-subtitle">
              ✨ <strong>Padrão: Presença automática</strong> | Marque apenas as caixas de <strong>Falta</strong> ou <strong>Falta Justificada</strong>
            </p>
          </div>
          <div style="display: flex; gap: 0.75rem; align-items: center; flex-wrap: wrap;">
            <input type="date" id="freqDateInput" class="form-control" style="width: auto; font-weight: 600;" onchange="renderFrequencia()" />
            <button class="btn btn-secondary btn-sm admin-teacher-only" onclick="clearDayAttendance()">✨ Redefinir Todos como Presentes</button>
          </div>
        </div>

        <div class="table-container">
          <table id="tableFrequencia">
            <thead>
              <tr>
                <th style="width: 50px;">Nº</th>
                <th>Aluno</th>
                <th style="text-align: center; width: 160px;">Falta (F)</th>
                <th style="text-align: center; width: 200px;">Falta Justificada (FJ)</th>
                <th style="text-align: center; width: 150px;">Status no Dia</th>
                <th style="text-align: right; width: 160px;">Acumulado</th>
              </tr>
            </thead>
            <tbody></tbody>
          </table>
        </div>
      </div>
    </section>

    <!-- TAB 3: PONTOS DE ATIVIDADES -->
    <section id="tab-pontos" class="tab-content">
      <div class="card">
        <div class="card-header">
          <div>
            <h2 class="card-title">⭐ Pontos Extras & Participação</h2>
            <p class="card-subtitle">Incremente ou deduza pontuações de atividades práticas, postura e tarefas</p>
          </div>
        </div>

        <div class="table-container">
          <table id="tablePontos">
            <thead>
              <tr>
                <th style="width: 50px;">Nº</th>
                <th>Aluno</th>
                <th style="text-align: center; width: 140px;">Saldo de Pontos</th>
                <th style="text-align: center; width: 360px;">Ajustar Pontuação</th>
              </tr>
            </thead>
            <tbody></tbody>
          </table>
        </div>
      </div>
    </section>

    <!-- TAB 4: NOTAS BIMESTRAIS -->
    <section id="tab-notas" class="tab-content">
      <div class="card">
        <div class="card-header">
          <div>
            <h2 class="card-title">📝 Notas das Avaliações - <span class="current-turma-name"></span></h2>
            <p class="card-subtitle">Cálculo: <strong>Média Bimestral = (Ativ 1 + Ativ 2 + Prova) / 3</strong> &bull; <strong>Média de Aprovação: 7.0</strong></p>
          </div>
          <div style="display: flex; gap: 0.6rem;">
            <span class="badge badge-success">≥ 7.0 Aprovado</span>
            <span class="badge badge-danger">&lt; 7.0 Recuperação</span>
          </div>
        </div>

        <div class="bimestre-tabs">
          <button class="bim-btn active" id="bimBtn0" onclick="switchBimestre(0)">1º Bimestre</button>
          <button class="bim-btn" id="bimBtn1" onclick="switchBimestre(1)">2º Bimestre</button>
          <button class="bim-btn" id="bimBtn2" onclick="switchBimestre(2)">3º Bimestre</button>
          <button class="bim-btn" id="bimBtn3" onclick="switchBimestre(3)">4º Bimestre</button>
          <button class="bim-btn" id="bimBtnAll" onclick="switchBimestre(-1)">Visão Consolidada Anual</button>
        </div>

        <div class="table-container">
          <table id="tableNotas">
            <thead>
              <tr id="tableNotasHeadRow"></tr>
            </thead>
            <tbody></tbody>
          </table>
        </div>
      </div>
    </section>

    <!-- TAB 5: RELATÓRIO GERAL -->
    <section id="tab-relatorios" class="tab-content">
      <div class="print-header">
        <div class="school-logo-container" style="background: transparent; border: none; box-shadow: none;">
          <span id="printLogoPlaceholder" style="font-size: 2.5rem;">🏫</span>
          <img id="printLogoImg" class="print-logo" style="display: none;" alt="Logo" />
        </div>
        <div>
          <h2 style="font-size: 1.4rem; font-weight: 800; color: #0f172a;" id="printSchoolName">Nome da Escola</h2>
          <p style="font-weight: 700; color: #334155; font-size: 1rem;">Diário de Classe &bull; Relatório Geral Consolidado</p>
          <p style="font-size: 0.85rem; color: #64748b;">Turma: <strong class="current-turma-name" style="color: #0f172a;"></strong> | Média Mínima: <strong>7.0</strong> | Data de Emissão: <strong id="reportDate"></strong></p>
        </div>
      </div>

      <div class="card">
        <div class="card-header no-print">
          <div>
            <h2 class="card-title">📊 Relatório Geral & Desempenho da Turma</h2>
            <p class="card-subtitle">Visão unificada de notas bimestrais, presenças, faltas e pontos extras</p>
          </div>
          <div style="display: flex; gap: 0.6rem; flex-wrap: wrap;">
            <button class="btn btn-primary" onclick="printReport()">🖨️ Imprimir / Salvar PDF</button>
            <button class="btn btn-success" onclick="exportStyledExcel()">📊 Baixar Excel Estilizado (.xlsx)</button>
          </div>
        </div>

        <div class="stat-grid no-print">
          <div class="stat-card">
            <div class="stat-icon" style="background: var(--primary-light); color: var(--primary);">👥</div>
            <div>
              <div class="stat-val" id="statTotalAlunos">0</div>
              <div class="stat-label">Alunos Matriculados</div>
            </div>
          </div>
          <div class="stat-card">
            <div class="stat-icon" style="background: var(--secondary-light); color: var(--secondary);">📅</div>
            <div>
              <div class="stat-val" id="statTotalAulas">0</div>
              <div class="stat-label">Aulas Registradas</div>
            </div>
          </div>
          <div class="stat-card">
            <div class="stat-icon" style="background: var(--success-light); color: var(--success);">📈</div>
            <div>
              <div class="stat-val" id="statMediaGeral">0.0</div>
              <div class="stat-label">Média Geral da Turma</div>
            </div>
          </div>
          <div class="stat-card">
            <div class="stat-icon" style="background: #fdf4ff; color: #c026d3;">🎯</div>
            <div>
              <div class="stat-val" id="statTaxaPresenca">0%</div>
              <div class="stat-label">Taxa de Assiduidade</div>
            </div>
          </div>
        </div>

        <div class="table-container">
          <table id="tableRelatorio">
            <thead>
              <tr>
                <th style="width: 35px;">Nº</th>
                <th>Nome do Aluno</th>
                <th style="text-align: center; width: 60px;">Pres.</th>
                <th style="text-align: center; width: 55px;">Faltas</th>
                <th style="text-align: center; width: 55px;">Justif.</th>
                <th style="text-align: center; width: 65px;">% Pres.</th>
                <th style="text-align: center; width: 70px;">Pontos</th>
                <th style="text-align: center; width: 65px;">1º Bim</th>
                <th style="text-align: center; width: 65px;">2º Bim</th>
                <th style="text-align: center; width: 65px;">3º Bim</th>
                <th style="text-align: center; width: 65px;">4º Bim</th>
                <th style="text-align: center; width: 75px;">Média Final</th>
                <th style="text-align: center; width: 85px;">Situação</th>
              </tr>
            </thead>
            <tbody></tbody>
          </table>
        </div>
      </div>
    </section>

    <!-- TAB 6: MURAL DE AVISOS & COMUNICADOS EM TEMPO REAL -->
    <section id="tab-comunicacao" class="tab-content">
      <div class="grid-2">
        <div class="card admin-teacher-only">
          <div class="card-header">
            <div>
              <h2 class="card-title">📢 Publicar Comunicado / Aviso</h2>
              <p class="card-subtitle">Dispare em tempo real via Push, SMS e WhatsApp para Pais e Alunos</p>
            </div>
          </div>
          <form onsubmit="handleSendAviso(event)">
            <div class="form-group">
              <label>Título do Comunicado:</label>
              <input type="text" id="avisoTituloInput" class="form-control" placeholder="Ex: Reunião de Pais e Mestres" required />
            </div>
            <div class="form-group">
              <label>Público Alvo:</label>
              <select id="avisoAlvoSelect" class="form-control">
                <option value="Toda a Escola">Toda a Escola</option>
                <option value="Apenas Responsáveis">Apenas Pais / Responsáveis</option>
                <option value="Apenas Alunos">Apenas Alunos</option>
              </select>
            </div>
            <div class="form-group">
              <label>Canais de Disparo em Tempo Real:</label>
              <div style="display: flex; gap: 1rem; margin-top: 0.3rem;">
                <label style="text-transform: none; font-weight: 600;"><input type="checkbox" id="chkPush" checked /> 🔔 Push App</label>
                <label style="text-transform: none; font-weight: 600;"><input type="checkbox" id="chkWhats" checked /> 💬 WhatsApp</label>
                <label style="text-transform: none; font-weight: 600;"><input type="checkbox" id="chkSMS" /> 📱 SMS</label>
              </div>
            </div>
            <div class="form-group">
              <label>Mensagem:</label>
              <textarea id="avisoTextoInput" class="form-control" rows="4" placeholder="Escreva a mensagem oficial..." required></textarea>
            </div>
            <button type="submit" class="btn btn-primary" style="width: 100%;">🚀 Transmitir Comunicado</button>
          </form>
        </div>

        <div class="card" style="grid-column: span 1;">
          <div class="card-header">
            <div>
              <h2 class="card-title">📋 Mural Oficial de Avisos</h2>
              <p class="card-subtitle">Comunicados recentes emitidos pela Direção e Professores</p>
            </div>
          </div>
          <div id="muralContainer"></div>
        </div>
      </div>
    </section>

    <!-- TAB 7: AGENDA DE EVENTOS ESCOLARES -->
    <section id="tab-agenda" class="tab-content">
      <div class="grid-2">
        <div class="card admin-teacher-only">
          <div class="card-header">
            <div>
              <h2 class="card-title">🗓️ Cadastrar Novo Evento</h2>
              <p class="card-subtitle">Adicione provas, reuniões, feriados e atividades esportivas</p>
            </div>
          </div>
          <form onsubmit="handleAddEvento(event)">
            <div class="form-group">
              <label>Nome do Evento:</label>
              <input type="text" id="eventoNomeInput" class="form-control" placeholder="Ex: Feira de Ciências 2026" required />
            </div>
            <div class="form-group">
              <label>Data do Evento:</label>
              <input type="date" id="eventoDataInput" class="form-control" required />
            </div>
            <div class="form-group">
              <label>Categoria:</label>
              <select id="eventoCatSelect" class="form-control">
                <option value="Avaliação">📝 Avaliação / Prova</option>
                <option value="Reunião">👥 Reunião de Pais</option>
                <option value="Cultural">🎨 Evento Cultural / Esportivo</option>
                <option value="Feriado">🌴 Feriado / Recesso</option>
              </select>
            </div>
            <button type="submit" class="btn btn-primary" style="width: 100%;">+ Adicionar à Agenda</button>
          </form>
        </div>

        <div class="card">
          <div class="card-header">
            <div>
              <h2 class="card-title">📅 Próximos Eventos Cadastrados</h2>
              <p class="card-subtitle">Calendário acadêmico unificado da instituição</p>
            </div>
          </div>
          <div class="table-container">
            <table id="tableEventos">
              <thead>
                <tr>
                  <th>Data</th>
                  <th>Evento</th>
                  <th style="text-align: center;">Categoria</th>
                  <th style="text-align: right;" class="admin-teacher-only">Ações</th>
                </tr>
              </thead>
              <tbody></tbody>
            </table>
          </div>
        </div>
      </div>
    </section>

    <!-- TAB 8: CANAL DIRETO COM RESPONSÁVEIS -->
    <section id="tab-canal" class="tab-content">
      <div class="card">
        <div class="card-header">
          <div>
            <h2 class="card-title">💬 Canal Direto com Pais & Responsáveis</h2>
            <p class="card-subtitle">Comunicação direta, segura e gravada entre Família e Escola</p>
          </div>
          <div class="form-group" style="margin-bottom: 0; min-width: 260px;">
            <select id="chatAlunoSelect" class="form-control" onchange="switchChatAluno(this.value)"></select>
          </div>
        </div>

        <div class="chat-box">
          <div class="chat-messages" id="chatMessagesContainer"></div>
          <div class="chat-input-area">
            <input type="text" id="chatInputText" class="form-control" placeholder="Digite uma mensagem para o responsável/escola..." onkeypress="if(event.key==='Enter') sendChatMessage()" />
            <button class="btn btn-primary" onclick="sendChatMessage()">Enviar</button>
          </div>
        </div>
      </div>
    </section>

  </main>

  <!-- MODAL CONFIGURAÇÃO ESCOLA -->
  <div class="modal-overlay" id="schoolModal">
    <div class="modal-card">
      <h3 style="font-size: 1.3rem; font-weight: 800; margin-bottom: 0.4rem; color: var(--text);">Configurar Escola</h3>
      <p style="font-size: 0.85rem; color: var(--text-muted); margin-bottom: 1.4rem;">Personalize a identidade da instituição no cabeçalho e relatórios</p>
      
      <div class="form-group">
        <label>Nome da Escola / Instituição:</label>
        <input type="text" id="schoolNameInput" class="form-control" placeholder="Ex: Escola Municipal Darcy Ribeiro" />
      </div>

      <div class="form-group">
        <label>Logotipo da Escola (PNG ou JPG):</label>
        <input type="file" id="schoolLogoFileInput" accept="image/*" class="form-control" onchange="handleLogoUpload(event)" />
      </div>

      <div style="display: flex; justify-content: flex-end; gap: 0.6rem; margin-top: 1.8rem;">
        <button class="btn btn-danger btn-sm" onclick="removeLogo()" style="margin-right: auto;">Remover Logo</button>
        <button class="btn btn-secondary" onclick="closeSchoolModal()">Cancelar</button>
        <button class="btn btn-primary" onclick="saveSchoolConfig()">Salvar Alterações</button>
      </div>
    </div>
  </div>

  <script>
    // --- STORAGE & ESTADO ---
    const STORAGE_KEY = "eduflow_pro_db_v7";
    const MEDIA_APROVACAO = 7.0;
    let currentActiveBimestre = 0; 
    let currentRole = "professor"; // direcao, professor, aluno, pais

    const defaultData = {
      escola: {
        nome: "Escola Municipal de Ensino Fundamental",
        logo: null
      },
      avisos: [
        {
          id: "av_1",
          titulo: "Reunião Geral de Pais e Mestres",
          autor: "Coordenação Pedagógica",
          data: "20/10/2026",
          alvo: "Apenas Responsáveis",
          texto: "Prezados pais, convidamos todos para nossa reunião do 3º trimestre nesta sexta-feira às 18h no auditório principal."
        },
        {
          id: "av_2",
          titulo: "Início dos Jogos Escolares 2026",
          autor: "Prof. Wanderley Junior",
          data: "18/10/2026",
          alvo: "Toda a Escola",
          texto: "Lembramos a todos os alunos que as inscrições para as modalidades de futsal e handebol encerram amanhã!"
        }
      ],
      eventos: [
        { id: "ev_1", nome: "Feira de Ciências e Tecnologia", data: "2026-10-28", cat: "Cultural" },
        { id: "ev_2", nome: "Avaliação Bimestral de Educação Física", data: "2026-11-05", cat: "Avaliação" },
        { id: "ev_3", nome: "Conselho de Classe Participativo", data: "2026-11-12", cat: "Reunião" }
      ],
      mensagensChats: {
        "al_1": [
          { remetente: "Escola", texto: "Olá! Gostaria de confirmar a presença da Ana na feira de ciências.", data: "14/10 10:15" },
          { remetente: "Responsável", texto: "Bom dia! Sim, ela irá apresentar o projeto de robótica.", data: "14/10 10:20" }
        ]
      },
      turmas: [
        {
          id: "turma_1",
          nome: "6º Ano A",
          alunos: [
            {
              id: "al_1",
              nome: "Ana Beatriz Silva",
              pontos: 2.5,
              bimestres: [
                { ativ1: 8.0, ativ2: 7.5, prova: 8.5 },
                { ativ1: 7.0, ativ2: 7.0, prova: 7.0 },
                { ativ1: null, ativ2: null, prova: null },
                { ativ1: null, ativ2: null, prova: null }
              ],
              frequencia: { "2026-08-18": "P", "2026-08-19": "P", "2026-08-20": "P" }
            },
            {
              id: "al_2",
              nome: "Carlos Eduardo Santos",
              pontos: 1.0,
              bimestres: [
                { ativ1: 6.0, ativ2: 5.5, prova: 6.5 },
                { ativ1: 6.5, ativ2: 6.0, prova: 7.0 },
                { ativ1: null, ativ2: null, prova: null },
                { ativ1: null, ativ2: null, prova: null }
              ],
              frequencia: { "2026-08-18": "F", "2026-08-19": "P", "2026-08-20": "P" }
            },
            {
              id: "al_3",
              nome: "Gabriel Martins Pereira",
              pontos: 3.0,
              bimestres: [
                { ativ1: 9.0, ativ2: 9.5, prova: 10.0 },
                { ativ1: 8.5, ativ2: 9.0, prova: 9.5 },
                { ativ1: null, ativ2: null, prova: null },
                { ativ1: null, ativ2: null, prova: null }
              ],
              frequencia: { "2026-08-18": "P", "2026-08-19": "P", "2026-08-20": "P" }
            },
            {
              id: "al_4",
              nome: "Mariana Oliveira Costa",
              pontos: 1.5,
              bimestres: [
                { ativ1: 7.0, ativ2: 7.5, prova: 8.0 },
                { ativ1: null, ativ2: null, prova: null },
                { ativ1: null, ativ2: null, prova: null },
                { ativ1: null, ativ2: null, prova: null }
              ],
              frequencia: { "2026-08-18": "FJ", "2026-08-19": "P", "2026-08-20": "P" }
            }
          ]
        },
        {
          id: "turma_2",
          nome: "9º Ano B",
          alunos: [
            {
              id: "al_5",
              nome: "Larissa Fernandes Souza",
              pontos: 2.0,
              bimestres: [
                { ativ1: 7.5, ativ2: 8.0, prova: 8.5 },
                { ativ1: null, ativ2: null, prova: null },
                { ativ1: null, ativ2: null, prova: null },
                { ativ1: null, ativ2: null, prova: null }
              ],
              frequencia: { "2026-08-20": "P" }
            }
          ]
        }
      ],
      activeTurmaId: "turma_1"
    };

    let appState = loadState();

    function loadState() {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) {
        try { return JSON.parse(saved); } catch (e) { console.error("Erro ao carregar dados", e); }
      }
      return defaultData;
    }

    function saveState() {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(appState));
      renderAll();
    }

    function getActiveTurma() {
      return appState.turmas.find(t => t.id === appState.activeTurmaId) || appState.turmas[0];
    }

    // --- MUDANÇA DE PERMISSÃO DE ACESSO (ROLES) ---
    function changeActiveRole(newRole) {
      currentRole = newRole;
      applyRolePermissions();
      renderAll();
    }

    function applyRolePermissions() {
      // Ocultar/Exibir formulários de edição baseados no perfil
      const isAdmin = currentRole === "direcao";
      const isTeacher = currentRole === "professor";
      const isAdminOrTeacher = isAdmin || isTeacher;

      document.querySelectorAll(".admin-only").forEach(el => {
        el.style.display = isAdmin ? "inline-flex" : "none";
      });

      document.querySelectorAll(".admin-teacher-only").forEach(el => {
        el.style.display = isAdminOrTeacher ? "" : "none";
      });

      // Desabilitar edição de notas/frequência para Alunos e Pais
      const inputs = document.querySelectorAll(".input-nota, .checkbox-freq");
      inputs.forEach(inp => {
        inp.disabled = !isAdminOrTeacher;
      });
    }

    // --- MURAL DE AVISOS & COMUNICADOS ---
    function handleSendAviso(e) {
      e.preventDefault();
      const titulo = document.getElementById("avisoTituloInput").value.trim();
      const alvo = document.getElementById("avisoAlvoSelect").value;
      const texto = document.getElementById("avisoTextoInput").value.trim();

      if (!titulo || !texto) return;

      const newAviso = {
        id: "av_" + Date.now(),
        titulo: titulo,
        autor: currentRole === "direcao" ? "Direção Escolar" : "Prof. (Diário de Classe)",
        data: new Date().toLocaleDateString("pt-BR"),
        alvo: alvo,
        texto: texto
      };

      appState.avisos.unshift(newAviso);
      document.getElementById("avisoTituloInput").value = "";
      document.getElementById("avisoTextoInput").value = "";
      
      alert(`🚀 Comunicado enviado com sucesso via Push/WhatsApp/SMS para: ${alvo}`);
      saveState();
    }

    function deleteAviso(avisoId) {
      if (confirm("Excluir este aviso do mural?")) {
        appState.avisos = appState.avisos.filter(a => a.id !== avisoId);
        saveState();
      }
    }

    function renderMural() {
      const container = document.getElementById("muralContainer");
      container.innerHTML = "";

      if (!appState.avisos || appState.avisos.length === 0) {
        container.innerHTML = `<p style="color: var(--text-muted); font-size: 0.9rem;">Nenhum comunicado publicado no momento.</p>`;
        return;
      }

      appState.avisos.forEach(a => {
        const card = document.createElement("div");
        card.className = "notice-card";
        card.innerHTML = `
          <div class="notice-header">
            <span class="notice-title">${a.titulo}</span>
            <span class="badge badge-info">${a.alvo}</span>
          </div>
          <p class="notice-body">${a.texto}</p>
          <div class="notice-meta" style="margin-top: 0.8rem; display: flex; justify-content: space-between; align-items: center;">
            <span>Por: <strong>${a.autor}</strong> &bull; ${a.data}</span>
            ${currentRole === "direcao" ? `<button class="btn btn-danger btn-sm" onclick="deleteAviso('${a.id}')">Excluir</button>` : ''}
          </div>
        `;
        container.appendChild(card);
      });
    }

    // --- AGENDA DE EVENTOS ---
    function handleAddEvento(e) {
      e.preventDefault();
      const nome = document.getElementById("eventoNomeInput").value.trim();
      const data = document.getElementById("eventoDataInput").value;
      const cat = document.getElementById("eventoCatSelect").value;

      if (!nome || !data) return;

      appState.eventos.push({ id: "ev_" + Date.now(), nome: nome, data: data, cat: cat });
      document.getElementById("eventoNomeInput").value = "";
      saveState();
    }

    function deleteEvento(evId) {
      if (confirm("Remover evento da agenda?")) {
        appState.eventos = appState.eventos.filter(e => e.id !== evId);
        saveState();
      }
    }

    function renderAgenda() {
      const tbody = document.querySelector("#tableEventos tbody");
      tbody.innerHTML = "";

      if (!appState.eventos || appState.eventos.length === 0) {
        tbody.innerHTML = `<tr><td colspan="4" style="text-align: center; color: var(--text-muted); padding: 1.5rem;">Nenhum evento agendado.</td></tr>`;
        return;
      }

      appState.eventos.sort((a,b) => a.data.localeCompare(b.data)).forEach(ev => {
        const tr = document.createElement("tr");
        const dataFmt = new Date(ev.data + "T00:00:00").toLocaleDateString("pt-BR");
        tr.innerHTML = `
          <td><strong>${dataFmt}</strong></td>
          <td>${ev.nome}</td>
          <td style="text-align: center;"><span class="badge badge-info">${ev.cat}</span></td>
          <td style="text-align: right;" class="admin-teacher-only">
            <button class="btn btn-danger btn-sm" onclick="deleteEvento('${ev.id}')">Remover</button>
          </td>
        `;
        tbody.appendChild(tr);
      });
      applyRolePermissions();
    }

    // --- CANAL DIRETO COM RESPONSÁVEIS ---
    let selectedChatAlunoId = null;

    function renderChatSelect() {
      const select = document.getElementById("chatAlunoSelect");
      select.innerHTML = "";
      const turma = getActiveTurma();

      if (!turma || turma.alunos.length === 0) {
        select.innerHTML = `<option value="">Nenhum aluno cadastrado</option>`;
        return;
      }

      turma.alunos.forEach(a => {
        const opt = document.createElement("option");
        opt.value = a.id;
        opt.textContent = `Chat com Responsáveis de: ${a.nome}`;
        select.appendChild(opt);
      });

      if (!selectedChatAlunoId || !turma.alunos.find(a => a.id === selectedChatAlunoId)) {
        selectedChatAlunoId = turma.alunos[0].id;
      }
      select.value = selectedChatAlunoId;
      renderChatMessages();
    }

    function switchChatAluno(alunoId) {
      selectedChatAlunoId = alunoId;
      renderChatMessages();
    }

    function renderChatMessages() {
      const container = document.getElementById("chatMessagesContainer");
      container.innerHTML = "";

      if (!selectedChatAlunoId) return;

      const msgs = appState.mensagensChats[selectedChatAlunoId] || [];

      if (msgs.length === 0) {
        container.innerHTML = `<p style="text-align: center; color: var(--text-muted); font-size: 0.85rem; margin-top: auto;">Inicie uma conversa direta com os responsáveis do aluno.</p>`;
        return;
      }

      msgs.forEach(m => {
        const div = document.createElement("div");
        const isSentByMe = (currentRole === "pais" && m.remetente === "Responsável") || (currentRole !== "pais" && m.remetente === "Escola");
        div.className = `chat-msg ${isSentByMe ? 'sent' : 'received'}`;
        div.innerHTML = `
          <div><strong>${m.remetente}:</strong> ${m.texto}</div>
          <div style="font-size: 0.7rem; opacity: 0.8; text-align: right; margin-top: 0.2rem;">${m.data}</div>
        `;
        container.appendChild(div);
      });

      container.scrollTop = container.scrollHeight;
    }

    function sendChatMessage() {
      const input = document.getElementById("chatInputText");
      const texto = input.value.trim();
      if (!texto || !selectedChatAlunoId) return;

      if (!appState.mensagensChats[selectedChatAlunoId]) {
        appState.mensagensChats[selectedChatAlunoId] = [];
      }

      const senderLabel = currentRole === "pais" ? "Responsável" : (currentRole === "direcao" ? "Direção Escolar" : "Prof. Responsável");
      const now = new Date();
      const timeStr = `${now.getDate()}/${now.getMonth()+1} ${now.getHours()}:${String(now.getMinutes()).padStart(2,'0')}`;

      appState.mensagensChats[selectedChatAlunoId].push({
        remetente: senderLabel,
        texto: texto,
        data: timeStr
      });

      input.value = "";
      saveState();
    }

    // --- MODAL & LOGO ---
    function openSchoolModal() {
      document.getElementById("schoolNameInput").value = appState.escola.nome || "";
      document.getElementById("schoolModal").classList.add("open");
    }

    function closeSchoolModal() {
      document.getElementById("schoolModal").classList.remove("open");
    }

    let tempUploadedLogo = null;
    function handleLogoUpload(event) {
      const file = event.target.files[0];
      if (file) {
        const reader = new FileReader();
        reader.onload = function(e) { tempUploadedLogo = e.target.result; };
        reader.readAsDataURL(file);
      }
    }

    function removeLogo() {
      appState.escola.logo = null;
      tempUploadedLogo = null;
      saveState();
      closeSchoolModal();
    }

    function saveSchoolConfig() {
      const name = document.getElementById("schoolNameInput").value.trim();
      if (name) appState.escola.nome = name;
      if (tempUploadedLogo) {
        appState.escola.logo = tempUploadedLogo;
        tempUploadedLogo = null;
      }
      saveState();
      closeSchoolModal();
    }

    function updateSchoolUI() {
      const nome = appState.escola.nome || "Nome da Escola";
      document.getElementById("headerSchoolName").innerHTML = `${nome} <span style="font-size: 0.8rem; opacity: 0.6;">✏️</span>`;
      document.getElementById("printSchoolName").textContent = nome;

      const headerLogoImg = document.getElementById("headerLogoImg");
      const headerPlaceholder = document.getElementById("headerLogoPlaceholder");
      const printLogoImg = document.getElementById("printLogoImg");
      const printPlaceholder = document.getElementById("printLogoPlaceholder");

      if (appState.escola.logo) {
        headerLogoImg.src = appState.escola.logo;
        headerLogoImg.style.display = "block";
        headerPlaceholder.style.display = "none";

        printLogoImg.src = appState.escola.logo;
        printLogoImg.style.display = "block";
        printPlaceholder.style.display = "none";
      } else {
        headerLogoImg.style.display = "none";
        headerPlaceholder.style.display = "block";

        printLogoImg.style.display = "none";
        printPlaceholder.style.display = "block";
      }
    }

    // --- NAVEGAÇÃO ---
    function switchTab(tabId, btn) {
      document.querySelectorAll(".tab-content").forEach(el => el.classList.remove("active"));
      document.querySelectorAll(".tab-btn").forEach(el => el.classList.remove("active"));

      const target = document.getElementById(tabId);
      if (target) target.classList.add("active");
      if (btn) btn.classList.add("active");

      if (tabId === "tab-relatorios") renderRelatorio();
      if (tabId === "tab-frequencia") renderFrequencia();
      if (tabId === "tab-notas") renderNotas();
      if (tabId === "tab-pontos") renderPontos();
      if (tabId === "tab-comunicacao") renderMural();
      if (tabId === "tab-agenda") renderAgenda();
      if (tabId === "tab-canal") renderChatSelect();
    }

    function changeActiveTurma(turmaId) {
      appState.activeTurmaId = turmaId;
      saveState();
    }

    // --- TURMAS E ALUNOS ---
    function handleAddTurma(e) {
      e.preventDefault();
      const input = document.getElementById("newTurmaInput");
      const nome = input.value.trim();
      if (!nome) return;

      const newTurma = {
        id: "turma_" + Date.now(),
        nome: nome,
        alunos: []
      };

      appState.turmas.push(newTurma);
      appState.activeTurmaId = newTurma.id;
      input.value = "";
      saveState();
    }

    function deleteTurma(turmaId) {
      if (appState.turmas.length <= 1) {
        alert("Você deve manter ao menos uma turma cadastrada.");
        return;
      }
      if (confirm("Deseja realmente excluir esta turma e todos os registros de seus alunos?")) {
        appState.turmas = appState.turmas.filter(t => t.id !== turmaId);
        if (appState.activeTurmaId === turmaId) {
          appState.activeTurmaId = appState.turmas[0].id;
        }
        saveState();
      }
    }

    function handleAddAluno(e) {
      e.preventDefault();
      const input = document.getElementById("newAlunoInput");
      const nome = input.value.trim();
      if (!nome) return;

      const turma = getActiveTurma();
      if (!turma) return;

      turma.alunos.push({
        id: "al_" + Date.now(),
        nome: nome,
        pontos: 0,
        bimestres: [
          { ativ1: null, ativ2: null, prova: null },
          { ativ1: null, ativ2: null, prova: null },
          { ativ1: null, ativ2: null, prova: null },
          { ativ1: null, ativ2: null, prova: null }
        ],
        frequencia: {}
      });

      input.value = "";
      saveState();
    }

    function deleteAluno(alunoId) {
      if (confirm("Tem certeza que deseja remover este aluno?")) {
        const turma = getActiveTurma();
        turma.alunos = turma.alunos.filter(a => a.id !== alunoId);
        saveState();
      }
    }

    // --- FREQUÊNCIA ---
    function handleAttendanceCheckbox(alunoId, type, isChecked) {
      const turma = getActiveTurma();
      const date = document.getElementById("freqDateInput").value;
      const aluno = turma.alunos.find(a => a.id === alunoId);
      if (!aluno || !date) return;

      if (!isChecked) {
        aluno.frequencia[date] = "P";
      } else {
        aluno.frequencia[date] = (type === "F") ? "F" : "FJ";
      }
      saveState();
    }

    function clearDayAttendance() {
      const turma = getActiveTurma();
      const date = document.getElementById("freqDateInput").value;
      if (!date || !turma) return;

      turma.alunos.forEach(a => {
        a.frequencia[date] = "P";
      });
      saveState();
    }

    // --- PONTOS ---
    function adjustPontos(alunoId, delta) {
      const turma = getActiveTurma();
      const aluno = turma.alunos.find(a => a.id === alunoId);
      if (aluno) {
        aluno.pontos = Math.max(0, parseFloat(((aluno.pontos || 0) + delta).toFixed(2)));
        saveState();
      }
    }

    function setCustomPontos(alunoId) {
      const turma = getActiveTurma();
      const aluno = turma.alunos.find(a => a.id === alunoId);
      if (aluno) {
        const val = prompt(`Digite a pontuação total para ${aluno.nome}:`, aluno.pontos || 0);
        if (val !== null && !isNaN(val)) {
          aluno.pontos = Math.max(0, parseFloat(Number(val).toFixed(2)));
          saveState();
        }
      }
    }

    // --- NOTAS & BIMESTRES ---
    function switchBimestre(bimIndex) {
      currentActiveBimestre = bimIndex;
      document.querySelectorAll(".bim-btn").forEach(btn => btn.classList.remove("active"));
      if (bimIndex === -1) {
        document.getElementById("bimBtnAll").classList.add("active");
      } else {
        document.getElementById(`bimBtn${bimIndex}`).classList.add("active");
      }
      renderNotas();
    }

    function updateNotaItem(alunoId, bimIndex, field, value) {
      const turma = getActiveTurma();
      const aluno = turma.alunos.find(a => a.id === alunoId);
      if (aluno && aluno.bimestres[bimIndex]) {
        if (value === "" || value === null) {
          aluno.bimestres[bimIndex][field] = null;
        } else {
          let num = parseFloat(value);
          if (isNaN(num)) num = null;
          else num = Math.min(10, Math.max(0, parseFloat(num.toFixed(1))));
          aluno.bimestres[bimIndex][field] = num;
        }
        localStorage.setItem(STORAGE_KEY, JSON.stringify(appState));
        renderNotas();
      }
    }

    function calcMediaBimestre(bimObj) {
      if (!bimObj) return null;
      const { ativ1, ativ2, prova } = bimObj;
      const notas = [ativ1, ativ2, prova].filter(n => n !== null && n !== undefined && !isNaN(n));
      if (notas.length === 0) return null;
      const soma = notas.reduce((acc, curr) => acc + curr, 0);
      return parseFloat((soma / notas.length).toFixed(1));
    }

    function calcMediaAnual(bimestresArray) {
      if (!bimestresArray) return null;
      const mediasBim = bimestresArray.map(b => calcMediaBimestre(b)).filter(m => m !== null);
      if (mediasBim.length === 0) return null;
      const soma = mediasBim.reduce((acc, curr) => acc + curr, 0);
      return parseFloat((soma / mediasBim.length).toFixed(1));
    }

    // --- IMPRESSÃO FORMATADA ---
    function printReport() {
      switchTab('tab-relatorios', document.querySelectorAll('.tab-btn')[4]);
      renderRelatorio();
      setTimeout(() => {
        window.print();
      }, 150);
    }

    // --- GERADOR EXCEL COM DESIGN IDÊNTICO (USANDO EXCELJS) ---
    async function exportStyledExcel() {
      const turma = getActiveTurma();
      if (!turma || turma.alunos.length === 0) {
        alert("Não há dados cadastrados nesta turma para exportar.");
        return;
      }

      const workbook = new ExcelJS.Workbook();
      workbook.creator = "EduFlow Pro";
      workbook.created = new Date();

      const ws = workbook.addWorksheet("Relatório Consolidado", {
        views: [{ showGridLines: true }]
      });

      const COLOR_PRIMARY = "4F46E5";
      const COLOR_PRIMARY_LIGHT = "EEF2FF";
      const COLOR_DARK_TEXT = "0F172A";
      const COLOR_MUTED = "64748B";
      const COLOR_BORDER = "CBD5E1";
      const COLOR_BORDER_LIGHT = "E2E8F0";
      const COLOR_ZEBRA = "F8FAFC";
      const COLOR_SUCCESS_BG = "ECFDF5";
      const COLOR_SUCCESS_TXT = "047857";
      const COLOR_DANGER_BG = "FEF2F2";
      const COLOR_DANGER_TXT = "B91C1C";

      ws.mergeCells("A1:M1");
      const titleCell = ws.getCell("A1");
      titleCell.value = (appState.escola.nome || "ESCOLA MUNICIPAL").toUpperCase();
      titleCell.font = { name: "Segoe UI", size: 14, bold: true, color: { argb: "FFFFFFFF" } };
      titleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + COLOR_PRIMARY } };
      titleCell.alignment = { horizontal: "left", vertical: "middle", indent: 1 };
      ws.getRow(1).height = 30;

      ws.mergeCells("A2:M2");
      const subCell = ws.getCell("A2");
      subCell.value = "DIÁRIO DE CLASSE • RELATÓRIO CONSOLIDADO DE FREQUÊNCIA, PONTOS E NOTAS";
      subCell.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FFE0E7FF" } };
      subCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + COLOR_PRIMARY } };
      subCell.alignment = { horizontal: "left", vertical: "middle", indent: 1 };
      ws.getRow(2).height = 20;

      ws.getCell("A3").value = `Turma: ${turma.nome}`;
      ws.getCell("A3").font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "FF" + COLOR_DARK_TEXT } };
      ws.getCell("A3").alignment = { vertical: "middle" };

      ws.getCell("D3").value = `Critério de Aprovação: Média ≥ ${MEDIA_APROVACAO.toFixed(1)}`;
      ws.getCell("D3").font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "FF" + COLOR_PRIMARY } };
      ws.getCell("D3").alignment = { vertical: "middle" };

      ws.mergeCells("K3:M3");
      const dateCell = ws.getCell("K3");
      dateCell.value = `Emissão: ${new Date().toLocaleDateString('pt-BR')}`;
      dateCell.font = { name: "Segoe UI", size: 10, italic: true, color: { argb: "FF" + COLOR_MUTED } };
      dateCell.alignment = { horizontal: "right", vertical: "middle" };
      ws.getRow(3).height = 22;

      ws.getRow(4).height = 8;

      const kpis = [
        { tRange: "B5:C5", vRange: "B6:C6", title: "👥 TOTAL DE ALUNOS", formula: `COUNTA(B10:B${9 + turma.alunos.length})`, fmt: "0", bg: COLOR_PRIMARY_LIGHT, txt: COLOR_PRIMARY },
        { tRange: "E5:F5", vRange: "E6:F6", title: "📅 AULAS REGISTRADAS", val: document.getElementById("statTotalAulas").textContent || "0", fmt: "0", bg: COLOR_PRIMARY_LIGHT, txt: COLOR_PRIMARY },
        { tRange: "H5:I5", vRange: "H6:I6", title: "📈 MÉDIA DA TURMA", formula: `AVERAGE(L10:L${9 + turma.alunos.length})`, fmt: "0.0", bg: COLOR_SUCCESS_BG, txt: COLOR_SUCCESS_TXT },
        { tRange: "K5:L5", vRange: "K6:L6", title: "🎯 ASSIDUIDADE GERAL", formula: `AVERAGE(F10:F${9 + turma.alunos.length})`, fmt: "0.0%", bg: "FDF4FF", txt: "C026D3" }
      ];

      kpis.forEach(k => {
        ws.mergeCells(k.tRange);
        ws.mergeCells(k.vRange);
        const tc = ws.getCell(k.tRange.split(":")[0]);
        tc.value = k.title;
        tc.font = { name: "Segoe UI", size: 7.5, bold: true, color: { argb: "FF" + COLOR_MUTED } };
        tc.alignment = { horizontal: "center", vertical: "middle" };
        tc.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };

        const vc = ws.getCell(k.vRange.split(":")[0]);
        if (k.formula) vc.value = { formula: k.formula };
        else vc.value = Number(k.val) || k.val;
        vc.font = { name: "Segoe UI", size: 13, bold: true, color: { argb: "FF" + k.txt } };
        vc.alignment = { horizontal: "center", vertical: "middle" };
        vc.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + k.bg } };
        vc.numFmt = k.fmt;
      });

      ws.getRow(5).height = 16;
      ws.getRow(6).height = 24;
      ws.getRow(7).height = 10;

      const cols = [
        { col: "A", key: "num", header: "Nº", width: 6 },
        { col: "B", key: "nome", header: "Nome do Aluno", width: 28 },
        { col: "C", key: "p", header: "Presenças", width: 12 },
        { col: "D", key: "f", header: "Faltas (F)", width: 12 },
        { col: "E", key: "fj", header: "Justif. (FJ)", width: 12 },
        { col: "F", key: "pct", header: "% Presença", width: 13 },
        { col: "G", key: "pts", header: "Pontos", width: 11 },
        { col: "H", key: "b1", header: "1º Bim", width: 11 },
        { col: "I", key: "b2", header: "2º Bim", width: 11 },
        { col: "J", key: "b3", header: "3º Bim", width: 11 },
        { col: "K", key: "b4", header: "4º Bim", width: 11 },
        { col: "L", key: "med", header: "Média Final", width: 14 },
        { col: "M", key: "st", header: "Situação", width: 15 }
      ];

      cols.forEach(c => {
        const cell = ws.getCell(`${c.col}9`);
        cell.value = c.header;
        cell.font = { name: "Segoe UI", size: 9, bold: true, color: { argb: "FFFFFFFF" } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + COLOR_PRIMARY } };
        cell.alignment = { horizontal: c.col === "B" ? "left" : "center", vertical: "middle", indent: c.col === "B" ? 1 : 0 };
        cell.border = {
          top: { style: "medium", color: { argb: "FF" + COLOR_PRIMARY } },
          bottom: { style: "medium", color: { argb: "FF" + COLOR_PRIMARY } },
          left: { style: "thin", color: { argb: "FF" + COLOR_BORDER } },
          right: { style: "thin", color: { argb: "FF" + COLOR_BORDER } }
        };
        ws.getColumn(c.col).width = c.width;
      });
      ws.getRow(9).height = 25;

      const startRow = 10;
      turma.alunos.forEach((a, idx) => {
        const row = startRow + idx;
        let p = 0, f = 0, fj = 0;
        Object.values(a.frequencia || {}).forEach(st => {
          if (st === 'P') p++;
          if (st === 'F') f++;
          if (st === 'FJ') fj++;
        });

        const m1 = calcMediaBimestre(a.bimestres[0]);
        const m2 = calcMediaBimestre(a.bimestres[1]);
        const m3 = calcMediaBimestre(a.bimestres[2]);
        const m4 = calcMediaBimestre(a.bimestres[3]);
        const medReal = calcMediaAnual(a.bimestres);

        ws.getCell(`A${row}`).value = idx + 1;
        ws.getCell(`B${row}`).value = a.nome;
        ws.getCell(`C${row}`).value = p;
        ws.getCell(`D${row}`).value = f;
        ws.getCell(`E${row}`).value = fj;
        ws.getCell(`F${row}`).value = { formula: `C${row}/MAX(1, C${row}+D${row}+E${row})` };
        ws.getCell(`G${row}`).value = a.pontos || 0;
        ws.getCell(`H${row}`).value = m1 !== null ? m1 : "";
        ws.getCell(`I${row}`).value = m2 !== null ? m2 : "";
        ws.getCell(`J${row}`).value = m3 !== null ? m3 : "";
        ws.getCell(`K${row}`).value = m4 !== null ? m4 : "";
        ws.getCell(`L${row}`).value = { formula: `AVERAGE(H${row}:K${row})` };
        ws.getCell(`M${row}`).value = { formula: `IF(ISBLANK(L${row}),"PENDENTE",IF(L${row}>=${MEDIA_APROVACAO},"APROVADO","RECUPERAÇÃO"))` };

        ws.getCell(`F${row}`).numFmt = "0.0%";
        ws.getCell(`G${row}`).numFmt = '0.0" pts"';
        ["H", "I", "J", "K", "L"].forEach(col => { ws.getCell(`${col}${row}`).numFmt = "0.0"; });

        const isZebra = (idx % 2 === 1);
        const rowBg = isZebra ? COLOR_ZEBRA : "FFFFFF";

        cols.forEach(c => {
          const cell = ws.getCell(`${c.col}${row}`);
          cell.font = { name: "Segoe UI", size: 9.5, color: { argb: "FF" + COLOR_DARK_TEXT } };
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + rowBg } };
          cell.border = {
            top: { style: "thin", color: { argb: "FF" + COLOR_BORDER_LIGHT } },
            bottom: { style: "thin", color: { argb: "FF" + COLOR_BORDER_LIGHT } },
            left: { style: "thin", color: { argb: "FF" + COLOR_BORDER_LIGHT } },
            right: { style: "thin", color: { argb: "FF" + COLOR_BORDER_LIGHT } }
          };
          cell.alignment = { horizontal: c.col === "B" ? "left" : "center", vertical: "middle", indent: c.col === "B" ? 1 : 0 };
        });

        ws.getCell(`B${row}`).font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF" + COLOR_DARK_TEXT } };
        ws.getCell(`C${row}`).font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF" + COLOR_SUCCESS_TXT } };
        ws.getCell(`D${row}`).font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF" + COLOR_DANGER_TXT } };
        ws.getCell(`L${row}`).font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "FF" + COLOR_PRIMARY } };

        const isAprov = (medReal !== null && medReal >= MEDIA_APROVACAO);
        const situacaoCell = ws.getCell(`M${row}`);
        situacaoCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + (isAprov ? COLOR_SUCCESS_BG : COLOR_DANGER_BG) } };
        situacaoCell.font = { name: "Segoe UI", size: 8.5, bold: true, color: { argb: "FF" + (isAprov ? COLOR_SUCCESS_TXT : COLOR_DANGER_TXT) } };

        ws.getRow(row).height = 22;
      });

      const totRow = startRow + turma.alunos.length;
      ws.getCell(`B${totRow}`).value = "MÉDIA GERAL DA TURMA";
      ws.getCell(`C${totRow}`).value = { formula: `SUM(C${startRow}:C${totRow - 1})` };
      ws.getCell(`D${totRow}`).value = { formula: `SUM(D${startRow}:D${totRow - 1})` };
      ws.getCell(`E${totRow}`).value = { formula: `SUM(E${startRow}:E${totRow - 1})` };
      ws.getCell(`F${totRow}`).value = { formula: `AVERAGE(F${startRow}:F${totRow - 1})` };
      ws.getCell(`G${totRow}`).value = { formula: `AVERAGE(G${startRow}:G${totRow - 1})` };
      ws.getCell(`H${totRow}`).value = { formula: `AVERAGE(H${startRow}:H${totRow - 1})` };
      ws.getCell(`I${totRow}`).value = { formula: `AVERAGE(I${startRow}:I${totRow - 1})` };
      ws.getCell(`J${totRow}`).value = { formula: `AVERAGE(J${startRow}:J${totRow - 1})` };
      ws.getCell(`K${totRow}`).value = { formula: `AVERAGE(K${startRow}:K${totRow - 1})` };
      ws.getCell(`L${totRow}`).value = { formula: `AVERAGE(L${startRow}:L${totRow - 1})` };

      ws.getCell(`F${totRow}`).numFmt = "0.0%";
      ws.getCell(`G${totRow}`).numFmt = '0.0" pts"';
      ["H", "I", "J", "K", "L"].forEach(col => { ws.getCell(`${col}${totRow}`).numFmt = "0.0"; });

      cols.forEach(c => {
        const cell = ws.getCell(`${c.col}${totRow}`);
        cell.font = { name: "Segoe UI", size: 9, bold: true, color: { argb: "FF" + COLOR_DARK_TEXT } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF1F5F9" } };
        cell.border = {
          top: { style: "medium", color: { argb: "FF" + COLOR_BORDER } },
          bottom: { style: "double", color: { argb: "FF" + COLOR_BORDER } }
        };
        cell.alignment = { horizontal: c.col === "B" ? "left" : "center", vertical: "middle", indent: c.col === "B" ? 1 : 0 };
      });
      ws.getRow(totRow).height = 24;

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      saveAs(blob, `Relatorio_${turma.nome.replace(/\s+/g, '_')}.xlsx`);
    }

    // --- RENDERIZAÇÃO GERAL ---
    function renderAll() {
      updateSchoolUI();
      renderSelectTurmas();
      renderTurmasTable();
      renderAlunosTable();
      renderFrequencia();
      renderPontos();
      renderNotas();
      renderRelatorio();
      renderMural();
      renderAgenda();
      renderChatSelect();
      applyRolePermissions();
    }

    function renderSelectTurmas() {
      const select = document.getElementById("currentTurmaSelect");
      select.innerHTML = "";
      appState.turmas.forEach(t => {
        const opt = document.createElement("option");
        opt.value = t.id;
        opt.textContent = t.nome;
        opt.selected = t.id === appState.activeTurmaId;
        select.appendChild(opt);
      });

      const activeTurma = getActiveTurma();
      document.querySelectorAll(".current-turma-name").forEach(el => {
        el.textContent = activeTurma ? activeTurma.nome : "";
      });
    }

    function renderTurmasTable() {
      const tbody = document.querySelector("#tableTurmas tbody");
      tbody.innerHTML = "";
      appState.turmas.forEach(t => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td><strong>${t.nome}</strong></td>
          <td><span class="badge badge-info">${t.alunos.length} alunos</span></td>
          <td style="text-align: right;">
            <button class="btn btn-secondary btn-sm" onclick="changeActiveTurma('${t.id}')">Selecionar</button>
            <button class="btn btn-danger btn-sm admin-only" onclick="deleteTurma('${t.id}')">Excluir</button>
          </td>
        `;
        tbody.appendChild(tr);
      });
    }

    function renderAlunosTable() {
      const tbody = document.querySelector("#tableAlunos tbody");
      tbody.innerHTML = "";
      const turma = getActiveTurma();
      if (!turma || turma.alunos.length === 0) {
        tbody.innerHTML = `<tr><td colspan="3" style="text-align: center; color: var(--text-muted); padding: 1.5rem;">Nenhum aluno matriculado nesta turma.</td></tr>`;
        return;
      }

      turma.alunos.forEach((a, idx) => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td><span style="font-weight: 700; color: var(--text-muted);">${idx + 1}</span></td>
          <td><strong>${a.nome}</strong></td>
          <td style="text-align: right;" class="admin-teacher-only">
            <button class="btn btn-danger btn-sm" onclick="deleteAluno('${a.id}')">Remover</button>
          </td>
        `;
        tbody.appendChild(tr);
      });
    }

    function renderFrequencia() {
      const dateInput = document.getElementById("freqDateInput");
      if (!dateInput.value) {
        const today = new Date().toISOString().split("T")[0];
        dateInput.value = today;
      }
      const currentDate = dateInput.value;
      const turma = getActiveTurma();
      const tbody = document.querySelector("#tableFrequencia tbody");
      tbody.innerHTML = "";

      if (!turma || turma.alunos.length === 0) {
        tbody.innerHTML = `<tr><td colspan="6" style="text-align: center; color: var(--text-muted); padding: 1.5rem;">Nenhum aluno nesta turma.</td></tr>`;
        return;
      }

      turma.alunos.forEach((a, idx) => {
        const currentStatus = a.frequencia[currentDate] || "P";
        const isFalta = currentStatus === "F";
        const isJustificada = currentStatus === "FJ";

        let presencas = 0, faltas = 0, justificadas = 0;
        Object.values(a.frequencia).forEach(st => {
          if (st === 'P') presencas++;
          if (st === 'F') faltas++;
          if (st === 'FJ') justificadas++;
        });

        let statusBadge = `<span class="badge badge-success">✓ Presente</span>`;
        if (isFalta) statusBadge = `<span class="badge badge-danger">✗ Falta</span>`;
        if (isJustificada) statusBadge = `<span class="badge badge-warning">⚠ Justificada</span>`;

        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td><span style="font-weight: 700; color: var(--text-muted);">${idx + 1}</span></td>
          <td><strong>${a.nome}</strong></td>
          
          <td style="text-align: center;">
            <label class="checkbox-freq-label ${isFalta ? 'is-falta' : ''}">
              <input type="checkbox" class="checkbox-freq danger-box" 
                ${isFalta ? 'checked' : ''} 
                onchange="handleAttendanceCheckbox('${a.id}', 'F', this.checked)" />
              <span>Falta</span>
            </label>
          </td>

          <td style="text-align: center;">
            <label class="checkbox-freq-label ${isJustificada ? 'is-justificada' : ''}">
              <input type="checkbox" class="checkbox-freq warning-box" 
                ${isJustificada ? 'checked' : ''} 
                onchange="handleAttendanceCheckbox('${a.id}', 'FJ', this.checked)" />
              <span>Justificada</span>
            </label>
          </td>

          <td style="text-align: center;">${statusBadge}</td>

          <td style="text-align: right; font-size: 0.85rem;">
            <span style="color: var(--success); font-weight: 700;">${presencas} P</span> / 
            <span style="color: var(--danger); font-weight: 700;">${faltas} F</span>
            ${justificadas > 0 ? `/ <span style="color: var(--warning); font-weight: 700;">${justificadas} FJ</span>` : ''}
          </td>
        `;
        tbody.appendChild(tr);
      });
      applyRolePermissions();
    }

    function renderPontos() {
      const turma = getActiveTurma();
      const tbody = document.querySelector("#tablePontos tbody");
      tbody.innerHTML = "";

      if (!turma || turma.alunos.length === 0) {
        tbody.innerHTML = `<tr><td colspan="4" style="text-align: center; color: var(--text-muted); padding: 1.5rem;">Nenhum aluno nesta turma.</td></tr>`;
        return;
      }

      turma.alunos.forEach((a, idx) => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td><span style="font-weight: 700; color: var(--text-muted);">${idx + 1}</span></td>
          <td><strong>${a.nome}</strong></td>
          <td style="text-align: center;">
            <span class="points-tag">${(a.pontos || 0).toFixed(1)} pts</span>
          </td>
          <td style="text-align: center; display: flex; gap: 0.45rem; justify-content: center;" class="admin-teacher-only">
            <button class="btn btn-secondary btn-sm" onclick="adjustPontos('${a.id}', -1.0)">-1.0</button>
            <button class="btn btn-secondary btn-sm" onclick="adjustPontos('${a.id}', -0.5)">-0.5</button>
            <button class="btn btn-success btn-sm" onclick="adjustPontos('${a.id}', 0.5)">+0.5</button>
            <button class="btn btn-success btn-sm" onclick="adjustPontos('${a.id}', 1.0)">+1.0</button>
            <button class="btn btn-primary btn-sm" onclick="setCustomPontos('${a.id}')">✏️ Editar</button>
          </td>
        `;
        tbody.appendChild(tr);
      });
      applyRolePermissions();
    }

    function renderNotas() {
      const turma = getActiveTurma();
      const theadRow = document.getElementById("tableNotasHeadRow");
      const tbody = document.querySelector("#tableNotas tbody");
      tbody.innerHTML = "";

      if (!turma || turma.alunos.length === 0) {
        theadRow.innerHTML = `<th>Aluno</th>`;
        tbody.innerHTML = `<tr><td style="text-align: center; color: var(--text-muted); padding: 1.5rem;">Nenhum aluno nesta turma.</td></tr>`;
        return;
      }

      if (currentActiveBimestre >= 0) {
        const bimNum = currentActiveBimestre + 1;
        theadRow.innerHTML = `
          <th style="width: 50px;">Nº</th>
          <th>Aluno</th>
          <th style="text-align: center; width: 110px;">Atividade 1</th>
          <th style="text-align: center; width: 110px;">Atividade 2</th>
          <th style="text-align: center; width: 110px;">Prova Bimestral</th>
          <th style="text-align: center; width: 120px;">Média ${bimNum}º Bim</th>
          <th style="text-align: center; width: 110px;">Média Anual</th>
          <th style="text-align: center; width: 120px;">Status</th>
        `;

        turma.alunos.forEach((a, idx) => {
          const bimData = (a.bimestres && a.bimestres[currentActiveBimestre]) || { ativ1: null, ativ2: null, prova: null };
          const mediaBim = calcMediaBimestre(bimData);
          const mediaAnual = calcMediaAnual(a.bimestres);

          let statusBadge = `<span class="badge badge-info">Pendente</span>`;
          if (mediaAnual !== null) {
            statusBadge = mediaAnual >= MEDIA_APROVACAO 
              ? `<span class="badge badge-success">Aprovado</span>` 
              : `<span class="badge badge-danger">Recuperação</span>`;
          }

          const tr = document.createElement("tr");
          tr.innerHTML = `
            <td><span style="font-weight: 700; color: var(--text-muted);">${idx + 1}</span></td>
            <td><strong>${a.nome}</strong></td>
            <td style="text-align: center;">
              <input type="number" step="0.1" min="0" max="10" class="input-nota" 
                value="${bimData.ativ1 !== null ? bimData.ativ1 : ''}" 
                onchange="updateNotaItem('${a.id}', ${currentActiveBimestre}, 'ativ1', this.value)" />
            </td>
            <td style="text-align: center;">
              <input type="number" step="0.1" min="0" max="10" class="input-nota" 
                value="${bimData.ativ2 !== null ? bimData.ativ2 : ''}" 
                onchange="updateNotaItem('${a.id}', ${currentActiveBimestre}, 'ativ2', this.value)" />
            </td>
            <td style="text-align: center;">
              <input type="number" step="0.1" min="0" max="10" class="input-nota" 
                value="${bimData.prova !== null ? bimData.prova : ''}" 
                onchange="updateNotaItem('${a.id}', ${currentActiveBimestre}, 'prova', this.value)" />
            </td>
            <td style="text-align: center; font-weight: 800; font-size: 0.95rem; color: ${mediaBim >= MEDIA_APROVACAO ? 'var(--success)' : (mediaBim !== null ? 'var(--danger)' : 'inherit')};">
              ${mediaBim !== null ? mediaBim.toFixed(1) : '-'}
            </td>
            <td style="text-align: center; font-weight: 800; font-size: 1.05rem; color: ${mediaAnual >= MEDIA_APROVACAO ? 'var(--success)' : (mediaAnual !== null ? 'var(--danger)' : 'inherit')};">
              ${mediaAnual !== null ? mediaAnual.toFixed(1) : '-'}
            </td>
            <td style="text-align: center;">${statusBadge}</td>
          `;
          tbody.appendChild(tr);
        });
      } else {
        theadRow.innerHTML = `
          <th style="width: 50px;">Nº</th>
          <th>Aluno</th>
          <th style="text-align: center; width: 100px;">Média 1º B</th>
          <th style="text-align: center; width: 100px;">Média 2º B</th>
          <th style="text-align: center; width: 100px;">Média 3º B</th>
          <th style="text-align: center; width: 100px;">Média 4º B</th>
          <th style="text-align: center; width: 110px;">Média Anual</th>
          <th style="text-align: center; width: 120px;">Situação</th>
        `;

        turma.alunos.forEach((a, idx) => {
          const m1 = calcMediaBimestre(a.bimestres[0]);
          const m2 = calcMediaBimestre(a.bimestres[1]);
          const m3 = calcMediaBimestre(a.bimestres[2]);
          const m4 = calcMediaBimestre(a.bimestres[3]);
          const mediaAnual = calcMediaAnual(a.bimestres);

          let statusBadge = `<span class="badge badge-info">Pendente</span>`;
          if (mediaAnual !== null) {
            statusBadge = mediaAnual >= MEDIA_APROVACAO 
              ? `<span class="badge badge-success">Aprovado</span>` 
              : `<span class="badge badge-danger">Recuperação</span>`;
          }

          const tr = document.createElement("tr");
          tr.innerHTML = `
            <td><span style="font-weight: 700; color: var(--text-muted);">${idx + 1}</span></td>
            <td><strong>${a.nome}</strong></td>
            <td style="text-align: center; font-weight: 600;">${m1 !== null ? m1.toFixed(1) : '-'}</td>
            <td style="text-align: center; font-weight: 600;">${m2 !== null ? m2.toFixed(1) : '-'}</td>
            <td style="text-align: center; font-weight: 600;">${m3 !== null ? m3.toFixed(1) : '-'}</td>
            <td style="text-align: center; font-weight: 600;">${m4 !== null ? m4.toFixed(1) : '-'}</td>
            <td style="text-align: center; font-weight: 800; font-size: 1.05rem; color: ${mediaAnual >= MEDIA_APROVACAO ? 'var(--success)' : (mediaAnual !== null ? 'var(--danger)' : 'inherit')};">
              ${mediaAnual !== null ? mediaAnual.toFixed(1) : '-'}
            </td>
            <td style="text-align: center;">${statusBadge}</td>
          `;
          tbody.appendChild(tr);
        });
      }
      applyRolePermissions();
    }

    function renderRelatorio() {
      const turma = getActiveTurma();
      const tbody = document.querySelector("#tableRelatorio tbody");
      tbody.innerHTML = "";

      const reportDate = document.getElementById("reportDate");
      if (reportDate) reportDate.textContent = new Date().toLocaleDateString("pt-BR");

      if (!turma || turma.alunos.length === 0) {
        tbody.innerHTML = `<tr><td colspan="13" style="text-align: center; color: var(--text-muted); padding: 2rem;">Nenhum dado disponível para exibição.</td></tr>`;
        document.getElementById("statTotalAlunos").textContent = "0";
        document.getElementById("statTotalAulas").textContent = "0";
        document.getElementById("statMediaGeral").textContent = "0.0";
        document.getElementById("statTaxaPresenca").textContent = "0%";
        return;
      }

      const allDates = new Set();
      turma.alunos.forEach(a => {
        Object.keys(a.frequencia).forEach(d => allDates.add(d));
      });
      const totalAulas = allDates.size;

      let somaMediasTurma = 0;
      let countAlunosComNota = 0;
      let totalPresencasGeral = 0;
      let totalChamadasGeral = 0;

      turma.alunos.forEach((a, idx) => {
        let p = 0, f = 0, fj = 0;
        Object.values(a.frequencia).forEach(st => {
          if (st === 'P') p++;
          if (st === 'F') f++;
          if (st === 'FJ') fj++;
        });

        const totalMarcado = p + f + fj;
        const pctPresenca = totalMarcado > 0 ? ((p / totalMarcado) * 100).toFixed(0) : 100;
        
        totalPresencasGeral += p;
        totalChamadasGeral += totalMarcado;

        const m1 = calcMediaBimestre(a.bimestres[0]);
        const m2 = calcMediaBimestre(a.bimestres[1]);
        const m3 = calcMediaBimestre(a.bimestres[2]);
        const m4 = calcMediaBimestre(a.bimestres[3]);
        const mediaAnual = calcMediaAnual(a.bimestres);

        if (mediaAnual !== null) {
          somaMediasTurma += mediaAnual;
          countAlunosComNota++;
        }

        let statusText = mediaAnual === null ? 'Pendente' : (mediaAnual >= MEDIA_APROVACAO ? 'Aprovado' : 'Recuperação');
        let statusColor = mediaAnual === null ? 'badge-info' : (mediaAnual >= MEDIA_APROVACAO ? 'badge-success' : 'badge-danger');

        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td><span style="font-weight: 700; color: var(--text-muted);">${idx + 1}</span></td>
          <td><strong>${a.nome}</strong></td>
          <td style="text-align: center; color: var(--success); font-weight: 800;">${p}</td>
          <td style="text-align: center; color: var(--danger); font-weight: 800;">${f}</td>
          <td style="text-align: center; color: var(--warning); font-weight: 800;">${fj}</td>
          <td style="text-align: center;"><strong>${pctPresenca}%</strong></td>
          <td style="text-align: center;"><span class="badge badge-info">${(a.pontos || 0).toFixed(1)} pts</span></td>
          <td style="text-align: center; font-weight: 600;">${m1 !== null ? m1.toFixed(1) : '-'}</td>
          <td style="text-align: center; font-weight: 600;">${m2 !== null ? m2.toFixed(1) : '-'}</td>
          <td style="text-align: center; font-weight: 600;">${m3 !== null ? m3.toFixed(1) : '-'}</td>
          <td style="text-align: center; font-weight: 600;">${m4 !== null ? m4.toFixed(1) : '-'}</td>
          <td style="text-align: center; font-weight: 800; font-size: 1rem; color: ${mediaAnual >= MEDIA_APROVACAO ? 'var(--success)' : (mediaAnual !== null ? 'var(--danger)' : 'inherit')};">
            ${mediaAnual !== null ? mediaAnual.toFixed(1) : '-'}
          </td>
          <td style="text-align: center;"><span class="badge ${statusColor}">${statusText}</span></td>
        `;
        tbody.appendChild(tr);
      });

      document.getElementById("statTotalAlunos").textContent = turma.alunos.length;
      document.getElementById("statTotalAulas").textContent = totalAulas;
      document.getElementById("statMediaGeral").textContent = countAlunosComNota > 0 ? (somaMediasTurma / countAlunosComNota).toFixed(1) : "0.0";
      document.getElementById("statTaxaPresenca").textContent = totalChamadasGeral > 0 ? `${((totalPresencasGeral / totalChamadasGeral) * 100).toFixed(0)}%` : "100%";
    }

    // Inicialização
    window.addEventListener("DOMContentLoaded", () => {
      renderAll();
    });
  </script>
</body>
</html>
