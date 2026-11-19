import os
import json
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# =====================================================================
# CONFIGURACIÓN DE RUTAS Y NOMBRES DE ARCHIVO
# =====================================================================
EXCEL_ENTRADA = "reporte_sucio.xlsx"  
EXCEL_SALIDA = "Reporte_Limpiado_Visual.xlsx"
DASHBOARD_SALIDA = "index.html"  

# Paleta de Colores Oficial HAVAS (Blanco, Negro y Rojo)
COLOR_HAVAS_BLACK = "111111"  # Negro corporativo para cabeceras
COLOR_HAVAS_RED = "E31B23"    # Rojo Havas
COLOR_LIGHT_BG = "F8FAFC"     # Gris muy claro para filas alternas
COLOR_WHITE = "FFFFFF"

# =====================================================================
# PLANTILLA HTML INTEGRADA (Identidad Corporativa HAVAS)
# =====================================================================
HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Havas - Performance Dashboard: GCLID vs LOGS</title>
    <style>
        :root {
            --primary: #111111; /* Havas Black */
            --havas-red: #e31b23; /* Havas Red */
            --secondary: #00a896; /* Naturgy Teal */
            --accent: #ff6b6b;
            --bg-main: #f4f7f6;
            --bg-card: #ffffff;
            --text-dark: #111111;
            --text-light: #718096;
            --border: #e2e8f0;
            --success: #48bb78;
            --warning: #ed8936;
            --danger: #e53e3e;
        }
        * { box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }
        body { background-color: var(--bg-main); color: var(--text-dark); line-height: 1.5; padding-bottom: 50px; }
        
        /* HEADER CORPORATIVO HAVAS: FONDO BLANCO, LOGO ROJO, LETRAS NEGRAS */
        header { 
            background-color: #ffffff; 
            border-bottom: 2px solid var(--border); 
            color: #000000; 
            padding: 15px 40px; 
            display: flex; 
            justify-content: space-between; 
            align-items: center; 
            box-shadow: 0 2px 8px rgba(0,0,0,0.04); 
        }
        .logo-area { display: flex; align-items: center; gap: 15px; }
        .logo-havas { 
            font-weight: 900; 
            font-size: 1.6rem; 
            letter-spacing: 3px; 
            color: var(--havas-red); 
            border-right: 2px solid var(--border); 
            padding-right: 15px; 
        }
        .project-title { 
            font-size: 1.1rem; 
            font-weight: 700; 
            color: #000000; 
        }
        .badge-client { 
            background-color: #f7fafc; 
            color: var(--text-dark); 
            border: 1px solid var(--border);
            font-weight: bold; 
            padding: 6px 14px; 
            border-radius: 20px; 
            font-size: 0.85rem; 
        }

        .container { max-width: 1400px; margin: 30px auto; padding: 0 20px; }
        .kpi-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 20px; margin-bottom: 30px; }
        .kpi-card { background-color: var(--bg-card); border-radius: 12px; padding: 24px; box-shadow: 0 4px 6px rgba(0,0,0,0.02); border: 1px solid var(--border); position: relative; overflow: hidden; transition: transform 0.2s; }
        .kpi-card:hover { transform: translateY(-2px); }
        .kpi-card::before { content: ''; position: absolute; top: 0; left: 0; width: 4px; height: 100%; background-color: var(--havas-red); }
        .kpi-card.alert::before { background-color: var(--accent); }
        .kpi-label { font-size: 0.85rem; color: var(--text-light); text-transform: uppercase; font-weight: 600; letter-spacing: 0.5px; }
        .kpi-value { font-size: 2.2rem; font-weight: 700; color: var(--primary); margin: 10px 0; }
        .kpi-desc { font-size: 0.8rem; color: var(--text-light); }
        .dashboard-grid { display: grid; grid-template-columns: 1fr; gap: 30px; }
        @media (min-width: 1024px) { .dashboard-grid { grid-template-columns: 350px 1fr; } }
        .sidebar-card { background-color: var(--bg-card); border-radius: 12px; padding: 24px; box-shadow: 0 4px 6px rgba(0,0,0,0.02); border: 1px solid var(--border); height: fit-content; }
        .filter-group { margin-bottom: 20px; }
        .filter-label { display: block; font-size: 0.9rem; font-weight: 600; margin-bottom: 10px; color: var(--primary); }
        .select-input { width: 100%; padding: 10px 12px; border-radius: 8px; border: 1px solid var(--border); background-color: #fff; color: var(--text-dark); font-size: 0.95rem; outline: none; cursor: pointer; }
        .main-card { background-color: var(--bg-card); border-radius: 12px; padding: 24px; box-shadow: 0 4px 6px rgba(0,0,0,0.02); border: 1px solid var(--border); }
        .card-header-flex { display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid var(--border); padding-bottom: 15px; margin-bottom: 20px; flex-wrap: wrap; gap: 15px; }
        .card-title { font-size: 1.25rem; font-weight: 700; color: var(--primary); }
        .visual-chart-container { margin-bottom: 30px; background: #fafbfc; border-radius: 8px; padding: 20px; border: 1px dashed var(--border); }
        .coverage-bar-group { margin-bottom: 15px; }
        .coverage-labels { display: flex; justify-content: space-between; font-size: 0.85rem; margin-bottom: 5px; }
        .campaign-name-label { font-weight: 600; color: var(--primary); max-width: 70%; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
        .bar-track { background-color: #edf2f7; height: 24px; border-radius: 12px; position: relative; overflow: hidden; }
        .bar-fill-gclid { background: linear-gradient(90deg, #111111 0%, var(--havas-red) 100%); height: 100%; border-radius: 12px 0 0 12px; position: absolute; top: 0; left: 0; transition: width 0.5s ease; }
        .bar-marker-logs { width: 4px; background-color: var(--primary); height: 100%; position: absolute; top: 0; z-index: 10; transition: left 0.5s ease; }
        .bar-marker-logs::after { content: 'LOGS'; position: absolute; top: -16px; transform: translateX(-50%); font-size: 0.65rem; background: var(--primary); color: white; padding: 1px 4px; border-radius: 4px; font-weight: bold; }
        .table-responsive { overflow-x: auto; }
        table { width: 100%; border-collapse: collapse; text-align: left; font-size: 0.9rem; }
        th { background-color: #111111; color: #ffffff; padding: 12px 16px; font-weight: 700; border-bottom: 2px solid var(--border); text-transform: uppercase; font-size: 0.75rem; letter-spacing: 0.5px; }
        td { padding: 14px 16px; border-bottom: 1px solid var(--border); color: var(--text-dark); }
        tr:hover td { background-color: #fcfdfe; }
        .badge-coverage { display: inline-block; padding: 3px 8px; border-radius: 12px; font-size: 0.8rem; font-weight: 600; }
        .badge-good { background-color: rgba(72, 187, 120, 0.15); color: var(--success); }
        .badge-warning { background-color: rgba(237, 137, 54, 0.15); color: var(--warning); }
        .badge-danger { background-color: rgba(229, 62, 98, 0.15); color: var(--danger); }
        .info-banner { background-color: #f7fafc; border-left: 4px solid var(--havas-red); padding: 15px 20px; border-radius: 0 8px 8px 0; font-size: 0.85rem; color: var(--text-dark); margin-bottom: 25px; display: flex; align-items: center; gap: 12px; }
        .search-input { padding: 8px 12px; border: 1px solid var(--border); border-radius: 6px; font-size: 0.85rem; width: 220px; outline: none; }
        .btn-action { background-color: var(--primary); color: white; border: none; padding: 8px 16px; border-radius: 6px; font-weight: 600; font-size: 0.85rem; cursor: pointer; display: inline-flex; align-items: center; gap: 8px; text-decoration: none; }
        .tab-btn { background-color: transparent; border: none; font-size: 0.95rem; font-weight: 600; padding: 8px 16px; cursor: pointer; color: var(--text-light); border-bottom: 2px solid transparent; }
        .tab-btn.active { color: var(--primary); border-bottom-color: var(--havas-red); }
        .tab-nav { display: flex; gap: 15px; margin-bottom: 25px; border-bottom: 1px solid var(--border); }
        .python-sync-box { background: #edf2f7; border-radius: 8px; padding: 15px; font-size: 0.8rem; margin-top: 25px; border: 1px solid #cbd5e0; }
    </style>
</head>
<body>
    <header>
        <div class="logo-area">
            <span class="logo-havas">HAVAS</span>
            <span class="project-title">Auditor de Conversiones (GCLID vs LOGS)</span>
        </div>
        <div><span class="badge-client">Cliente: Naturgy Energía</span></div>
    </header>
    <div class="container">
        <div class="info-banner">
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="16" x2="12" y2="12"></line><line x1="12" y1="8" x2="12.01" y2="8"></line></svg>
            <div>Este reporte audita la discrepancia entre las conversiones registradas en el link con GCLID y las confirmadas en los LOGS internos de Naturgy.</div>
        </div>
        <div class="kpi-row">
            <div class="kpi-card">
                <div class="kpi-label">LINK con GCLID Totales</div>
                <div class="kpi-value" id="kpi-link">0</div>
                <div class="kpi-desc">Conversiones asociadas a clicks calificados</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-label">LOGS Confirmados Totales</div>
                <div class="kpi-value" id="kpi-logs">0</div>
                <div class="kpi-desc">Conversiones reportadas en sistemas internos</div>
            </div>
            <div class="kpi-card" id="kpi-cobertura-card">
                <div class="kpi-label">% Cobertura General</div>
                <div class="kpi-value" id="kpi-cobertura">0%</div>
                <div class="kpi-desc">Ratio promedio LOGS vs LINK</div>
            </div>
            <div class="kpi-card alert" id="kpi-alerts-card">
                <div class="kpi-label">Alertas de Desviación</div>
                <div class="kpi-value" id="kpi-alerts">0</div>
                <div class="kpi-desc">Contratos fuera del margen óptimo (90% - 105%)</div>
            </div>
        </div>
        
        <!-- Pestañas corregidas con IDs para evitar fallos de comillas -->
        <div class="tab-nav">
            <button id="btn-tab-normal" class="tab-btn active" onclick="switchTab('normal')">Vista General (Contratos)</button>
            <button id="btn-tab-campana" class="tab-btn" onclick="switchTab('campana')">Desglose por Campaña</button>
        </div>
        
        <div class="dashboard-grid">
            <div class="sidebar-card">
                <h3 style="font-size: 1.05rem; margin-bottom: 20px; color: var(--primary);">Filtros de Análisis</h3>
                <div class="filter-group">
                    <label class="filter-label" for="contract-filter">Contrato o Producto:</label>
                    <select id="contract-filter" class="select-input" onchange="applyFilters()">
                        <option value="all">-- Todos los Contratos --</option>
                    </select>
                </div>
                <div class="filter-group" id="campaign-filter-wrapper" style="display: none;">
                    <label class="filter-label" for="campaign-search">Buscar Campaña:</label>
                    <input type="text" id="campaign-search" class="search-input" style="width: 100%;" placeholder="Escribe el nombre de la campaña..." onkeyup="applyFilters()">
                </div>
                <div class="python-sync-box">
                    <h4 style="margin-bottom: 8px;">🔄 Sincronizado por Python</h4>
                    <p>Los datos han sido leídos y procesados del excel original automáticamente.</p>
                </div>
            </div>
            <div class="main-card">
                <div class="card-header-flex">
                    <div class="card-title" id="table-title">Análisis de Cobertura Global</div>
                    <div><button class="btn-action" onclick="exportCSV()">Exportar CSV Limpio</button></div>
                </div>
                <div class="visual-chart-container" id="visual-chart-box">
                    <h4 style="font-size: 0.9rem; margin-bottom: 15px; color: var(--text-light); text-transform: uppercase;">Comparación Visual: GCLID vs LOGS</h4>
                    <div id="chart-list"></div>
                </div>
                <div class="table-responsive">
                    <table id="data-table">
                        <thead><tr id="table-headers"></tr></thead>
                        <tbody id="table-body"></tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
    <script>
        // DATOS INYECTADOS POR PYTHON
        // INSERT_DATA_RESULTADOS_HERE
        // INSERT_DATA_CAMPANAS_HERE

        let currentTab = 'normal';

        function initFilters() {
            const select = document.getElementById('contract-filter');
            select.innerHTML = '<option value="all">-- Todos los Contratos --</option>';
            const uniqueContracts = [...new Set(dataResultados.map(item => item.contrato))];
            uniqueContracts.sort().forEach(contract => {
                const opt = document.createElement('option');
                opt.value = contract;
                opt.textContent = contract;
                select.appendChild(opt);
            });
        }

        function switchTab(tab) {
            currentTab = tab;
            document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
            if(tab === 'normal') {
                document.getElementById('btn-tab-normal').classList.add('active');
                document.getElementById('table-title').textContent = 'Análisis de Cobertura Global (Contratos)';
                document.getElementById('campaign-filter-wrapper').style.display = 'none';
            } else {
                document.getElementById('btn-tab-campana').classList.add('active');
                document.getElementById('table-title').textContent = 'Análisis de Cobertura Detallado (Campañas)';
                document.getElementById('campaign-filter-wrapper').style.display = 'block';
            }
            applyFilters();
        }

        function applyFilters() {
            const selectedContract = document.getElementById('contract-filter').value;
            const searchVal = document.getElementById('campaign-search').value.toLowerCase();
            let filteredData = [];
            let isCampanaView = (currentTab === 'campana');

            if (!isCampanaView) {
                filteredData = dataResultados.filter(item => selectedContract === 'all' || item.contrato === selectedContract);
            } else {
                filteredData = dataCampanas.filter(item => {
                    const matchContract = selectedContract === 'all' || item.contrato === selectedContract;
                    const matchSearch = item.campana.toLowerCase().includes(searchVal);
                    return matchContract && matchSearch;
                });
            }
            renderDashboard(filteredData, isCampanaView);
        }

        function getCoverageStatus(ratio) {
            if (ratio >= 90 && ratio <= 105) return { class: 'badge-good', label: 'ÓPTIMO' };
            if ((ratio >= 80 && ratio < 90) || (ratio > 105 && ratio <= 115)) return { class: 'badge-warning', label: 'DESVIADO' };
            return { class: 'badge-danger', label: 'CRÍTICO' };
        }

        function renderDashboard(data, isCampanaView) {
            const headersRow = document.getElementById('table-headers');
            const tbody = document.getElementById('table-body');
            const chartList = document.getElementById('chart-list');

            if (!isCampanaView) {
                headersRow.innerHTML = `<th>Contrato</th><th style="text-align: right;">LINK con GCLID</th><th style="text-align: right;">LOGS</th><th style="text-align: right;">Cobertura %</th><th style="text-align: center;">Estado</th>`;
            } else {
                headersRow.innerHTML = `<th>Contrato</th><th>Campaña</th><th style="text-align: right;">LINK con GCLID</th><th style="text-align: right;">LOGS</th><th style="text-align: right;">Cobertura %</th><th style="text-align: center;">Estado</th>`;
            }

            tbody.innerHTML = '';
            chartList.innerHTML = '';
            let totalLink = 0; let totalLogs = 0; let alertCount = 0;
            const maxGclidVal = Math.max(...data.map(d => d.gclid), 10);

            if (data.length === 0) {
                tbody.innerHTML = `<tr><td colspan="${isCampanaView ? 6 : 5}" style="text-align: center; color: var(--text-light);">No se encontraron registros.</td></tr>`;
                document.getElementById('visual-chart-box').style.display = 'none';
            } else {
                document.getElementById('visual-chart-box').style.display = 'block';
                data.forEach(item => {
                    totalLink += item.gclid;
                    totalLogs += item.logs;
                    const status = getCoverageStatus(item.ratio);
                    if (status.label !== 'ÓPTIMO') alertCount++;

                    const tr = document.createElement('tr');
                    if (!isCampanaView) {
                        tr.innerHTML = `<td style="font-weight:600; color:var(--primary);">${item.contrato}</td><td style="text-align:right;">${item.gclid.toLocaleString()}</td><td style="text-align:right;">${item.logs.toLocaleString()}</td><td style="text-align:right; font-weight:bold;">${item.ratio.toFixed(1)}%</td><td style="text-align:center;"><span class="badge-coverage ${status.class}">${status.label}</span></td>`;
                    } else {
                        tr.innerHTML = `<td style="font-size:0.8rem; color:var(--text-light);">${item.contrato}</td><td style="font-weight:600; color:var(--primary);">${item.campana}</td><td style="text-align:right;">${item.gclid.toLocaleString()}</td><td style="text-align:right;">${item.logs.toLocaleString()}</td><td style="text-align:right; font-weight:bold;">${item.ratio.toFixed(1)}%</td><td style="text-align:center;"><span class="badge-coverage ${status.class}">${status.label}</span></td>`;
                    }
                    tbody.appendChild(tr);

                    if (chartList.children.length < 8) {
                        const displayName = isCampanaView ? item.campana : item.contrato;
                        const widthPercent = (item.gclid / maxGclidVal) * 90;
                        const logsMarkerPercent = item.gclid > 0 ? (item.logs / item.gclid) * 100 : 0;
                        const logsMarkerPosition = Math.min(logsMarkerPercent, 100);

                        const barHtml = `<div class="coverage-bar-group">
                            <div class="coverage-labels"><span class="campaign-name-label">${displayName}</span><span style="font-weight:bold; color:var(--primary);">GCLID: ${item.gclid} vs LOGS: ${item.logs} (${item.ratio.toFixed(1)}%)</span></div>
                            <div class="bar-track"><div class="bar-fill-gclid" style="width: ${Math.max(widthPercent, 2)}%;"><div class="bar-marker-logs" style="left: ${logsMarkerPosition}%;"></div></div></div>
                        </div>`;
                        chartList.insertAdjacentHTML('beforeend', barHtml);
                    }
                });
            }

            document.getElementById('kpi-link').textContent = totalLink.toLocaleString();
            document.getElementById('kpi-logs').textContent = totalLogs.toLocaleString();
            const generalRatio = totalLink > 0 ? (totalLogs / totalLink) * 100 : 0;
            document.getElementById('kpi-cobertura').textContent = `${generalRatio.toFixed(1)}%`;
            
            const alertsKpiCard = document.getElementById('kpi-alerts-card');
            const alertsKpiVal = document.getElementById('kpi-alerts');
            alertsKpiVal.textContent = alertCount;
            if (alertCount > 0) {
                alertsKpiCard.style.borderColor = 'var(--accent)';
                alertsKpiVal.style.color = 'var(--danger)';
            } else {
                alertsKpiCard.style.borderColor = 'var(--border)';
                alertsKpiVal.style.color = 'var(--success)';
            }
        }

        function exportCSV() {
            let csvContent = "data:text/csv;charset=utf-8,";
            if (currentTab === 'normal') {
                csvContent += "Contrato;LINK con GCLID;LOGS;Ratio Cobertura\\n";
                dataResultados.forEach(r => { csvContent += `"${r.contrato}";${r.gclid};${r.logs};${r.ratio}\\n`; });
            } else {
                csvContent += "Contrato;Campana;LINK con GCLID;LOGS;Ratio Cobertura\\n";
                dataCampanas.forEach(r => { csvContent += `"${r.contrato}";"${r.campana}";${r.gclid};${r.logs};${r.ratio}\\n`; });
            }
            const encodedUri = encodeURI(csvContent);
            const link = document.createElement("a");
            link.setAttribute("href", encodedUri);
            link.setAttribute("download", `Havas_Cobertura_${currentTab}.csv`);
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }

        // Se inyectarán los arrays dataResultados y dataCampanas arriba
        initFilters();
        applyFilters();
    </script>
</body>
</html>
"""

def limpiar_valores_numericos(val):
    """Limpia strings con %, comas, espacios y convierte a número."""
    if pd.isna(val):
        return 0
    if isinstance(val, (int, float)):
        return val
    s_val = str(val).replace('%', '').replace(' ', '').replace(',', '.').strip()
    try:
        return float(s_val) if '.' in s_val else int(s_val)
    except ValueError:
        return 0

def detectar_columnas_inteligente(df, es_campana=False):
    """Detecta las columnas correctas buscando palabras clave de forma flexible."""
    columnas_originales = df.columns.tolist()
    columnas_limpias = [str(c).strip().lower() for c in columnas_originales]
    
    mapa_columnas = {}
    
    # 1. Identificar columna Contrato (primera columna siempre)
    mapa_columnas["Contrato"] = columnas_originales[0]
    
    # 2. Si es la hoja de campaña, la campaña es la segunda columna
    if es_campana and len(columnas_originales) > 1:
        mapa_columnas["Campana"] = columnas_originales[1]
    
    # 3. Buscar GCLID
    for orig, limpia in zip(columnas_originales, columnas_limpias):
        if "gclid" in limpia and "%" not in limpia and "cobertura" not in limpia:
            mapa_columnas["GCLID"] = orig
            break
            
    # 4. Buscar LOGS
    for orig, limpia in zip(columnas_originales, columnas_limpias):
        if "logs" in limpia and "%" not in limpia and "cobertura" not in limpia:
            mapa_columnas["LOGS"] = orig
            break
            
    # 5. Buscar Cobertura/Ratio
    for orig, limpia in zip(columnas_originales, columnas_limpias):
        if "cobertura" in limpia or "ratio" in limpia or "%" in limpia:
            if "gclid" not in limpia:
                mapa_columnas["Ratio"] = orig
                break
                
    # Respaldos manuales por si falla la búsqueda flexible
    if "GCLID" not in mapa_columnas:
        mapa_columnas["GCLID"] = next((c for c in columnas_originales if "gclid" in str(c).lower()), None)
    if "LOGS" not in mapa_columnas:
         mapa_columnas["LOGS"] = next((c for c in columnas_originales if "logs" in str(c).lower()), None)
    if "Ratio" not in mapa_columnas:
         mapa_columnas["Ratio"] = next((c for c in columnas_originales if "%" in str(c) or "cobertura" in str(c).lower()), None)

    return mapa_columnas

def procesar_datos():
    print("🚀 Iniciando procesamiento inteligente de datos...")
    
    if not os.path.exists(EXCEL_ENTRADA):
        raise FileNotFoundError(f"No se encuentra el archivo '{EXCEL_ENTRADA}' en la carpeta Dashboard.")

    # =====================================================================
    # 1. PROCESAMIENTO: HOJA 'Resultados'
    # =====================================================================
    print("📖 Leyendo hoja 'Resultados'...")
    df_res = pd.read_excel(EXCEL_ENTRADA, sheet_name="Resultados")
    
    mapa_res = detectar_columnas_inteligente(df_res, es_campana=False)
    print(f"   🔍 Columnas detectadas para Resultados:")
    for k, v in mapa_res.items():
        print(f"      - {k} mapeada a -> '{v}'")
        
    if not mapa_res.get("GCLID") or not mapa_res.get("LOGS"):
        raise ValueError("No se pudieron identificar las columnas de GCLID o LOGS. Revisa los nombres en la hoja 'Resultados'.")

    # Filtrar y renombrar
    df_res = df_res.rename(columns={
        mapa_res["Contrato"]: "Contrato",
        mapa_res["GCLID"]: "GCLID",
        mapa_res["LOGS"]: "LOGS",
        mapa_res["Ratio"]: "Ratio"
    })
    
    df_res = df_res[df_res["Contrato"].notna()]
    df_res = df_res[~df_res["Contrato"].str.contains("Total|general|Etiquetas|Suma de", case=False, na=False)]
    
    df_res_clean = df_res[["Contrato", "GCLID", "LOGS", "Ratio"]].copy()
    df_res_clean["GCLID"] = df_res_clean["GCLID"].apply(limpiar_valores_numericos).astype(int)
    df_res_clean["LOGS"] = df_res_clean["LOGS"].apply(limpiar_valores_numericos).astype(int)
    df_res_clean["Ratio"] = df_res_clean.apply(
        lambda r: round((r["LOGS"] / r["GCLID"] * 100), 1) if r["GCLID"] > 0 else 0.0, axis=1
    )

    # =====================================================================
    # 2. PROCESAMIENTO: HOJA 'Resultados (Por campaña)'
    # =====================================================================
    print("📖 Leyendo hoja 'Resultados (Por campaña)'...")
    df_camp = pd.read_excel(EXCEL_ENTRADA, sheet_name="Resultados (Por campaña)")
    
    mapa_camp = detectar_columnas_inteligente(df_camp, es_campana=True)
    print(f"   🔍 Columnas detectadas para Campañas:")
    for k, v in mapa_camp.items():
        print(f"      - {k} mapeada a -> '{v}'")

    df_camp = df_camp.rename(columns={
        mapa_camp["Contrato"]: "Contrato",
        mapa_camp["Campana"]: "Campana",
        mapa_camp["GCLID"]: "GCLID",
        mapa_camp["LOGS"]: "LOGS",
        mapa_camp["Ratio"]: "Ratio"
    })
    
    df_camp = df_camp[df_camp["Contrato"].notna() & df_camp["Campana"].notna()]
    df_camp = df_camp[~df_camp["Campana"].str.contains("Total|general|Etiquetas|Suma de", case=False, na=False)]
    df_camp = df_camp[~df_camp["Contrato"].str.contains("Total|general", case=False, na=False)]
    
    df_camp_clean = df_camp[["Contrato", "Campana", "GCLID", "LOGS", "Ratio"]].copy()
    df_camp_clean["GCLID"] = df_camp_clean["GCLID"].apply(limpiar_valores_numericos).astype(int)
    df_camp_clean["LOGS"] = df_camp_clean["LOGS"].apply(limpiar_valores_numericos).astype(int)
    df_camp_clean["Ratio"] = df_camp_clean.apply(
        lambda r: round((r["LOGS"] / r["GCLID"] * 100), 1) if r["GCLID"] > 0 else 0.0, axis=1
    )

    return df_res_clean, df_camp_clean

def exportar_excel_visual(df_res, df_camp):
    print(f"🎨 Generando Excel visual y estilizado en '{EXCEL_SALIDA}'...")
    with pd.ExcelWriter(EXCEL_SALIDA, engine='openpyxl') as writer:
        df_res.to_excel(writer, sheet_name="Resultados", index=False)
        df_camp.to_excel(writer, sheet_name="Resultados por Campaña", index=False)
        
        workbook = writer.book
        
        # Fuentes corporativas
        font_header = Font(name="Segoe UI", size=11, bold=True, color=COLOR_WHITE)
        font_data = Font(name="Segoe UI", size=10)
        font_bold = Font(name="Segoe UI", size=10, bold=True)
        
        fill_header = PatternFill(start_color=COLOR_HAVAS_BLACK, end_color=COLOR_HAVAS_BLACK, fill_type="solid")
        fill_zebra = PatternFill(start_color=COLOR_LIGHT_BG, end_color=COLOR_LIGHT_BG, fill_type="solid")
        fill_accent = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid") # Rojo sutil
        
        thin_border = Border(
            left=Side(style='thin', color="E2E8F0"),
            right=Side(style='thin', color="E2E8F0"),
            top=Side(style='thin', color="E2E8F0"),
            bottom=Side(style='thin', color="E2E8F0")
        )
        
        for sheet_name in workbook.sheetnames:
            ws = workbook[sheet_name]
            ws.views.sheetView[0].showGridLines = True
            
            for col_idx in range(1, ws.max_column + 1):
                cell = ws.cell(row=1, column=col_idx)
                cell.font = font_header
                cell.fill = fill_header
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin_border
            
            for row_idx in range(2, ws.max_row + 1):
                is_even = row_idx % 2 == 0
                for col_idx in range(1, ws.max_column + 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    cell.font = font_data
                    cell.border = thin_border
                    
                    if is_even:
                        cell.fill = fill_zebra
                    
                    col_name = ws.cell(row=1, column=col_idx).value
                    if col_name in ["GCLID", "LOGS"]:
                        cell.number_format = '#,##0'
                        cell.alignment = Alignment(horizontal="right")
                    elif col_name == "Ratio":
                        cell.number_format = '0.0"%"'
                        cell.alignment = Alignment(horizontal="right")
                        cell.font = font_bold
                        cell.fill = fill_accent
                    else:
                        cell.alignment = Alignment(horizontal="left")
            
            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 3, 12)
                
    print("✔️ Archivo Excel visual generado.")

def actualizar_dashboard(df_res, df_camp):
    print("🌐 Sincronizando datos con el Dashboard Interactivo...")
    
    # 1. Transformación limpia a registros dict
    res_list = []
    for idx, row in df_res.iterrows():
        res_list.append({
            "id": idx + 1,
            "contrato": str(row["Contrato"]),
            "gclid": int(row["GCLID"]),
            "logs": int(row["LOGS"]),
            "ratio": float(row["Ratio"])
        })
        
    camp_list = []
    for idx, row in df_camp.iterrows():
        camp_list.append({
            "contrato": str(row["Contrato"]),
            "campana": str(row["Campana"]),
            "gclid": int(row["GCLID"]),
            "logs": int(row["LOGS"]),
            "ratio": float(row["Ratio"])
        })

    # 2. Generar el script de inyección de datos puro (sin comentarios que estorben)
    str_js_res = f"const dataResultados = {json.dumps(res_list, ensure_ascii=False, indent=12)};"
    str_js_camp = f"const dataCampanas = {json.dumps(camp_list, ensure_ascii=False, indent=12)};"

    # 3. Reemplazar usando marcas exactas en el HTML_TEMPLATE original
    html_content = HTML_TEMPLATE
    html_content = html_content.replace("// INSERT_DATA_RESULTADOS_HERE", str_js_res)
    html_content = html_content.replace("// INSERT_DATA_CAMPANAS_HERE", str_js_camp)

    # 4. Guardar archivo final
    with open(DASHBOARD_SALIDA, "w", encoding="utf-8") as f:
        f.write(html_content)

    print(f"✔️ ¡Dashboard actualizado con éxito como '{DASHBOARD_SALIDA}'!")

# =====================================================================
# EJECUCIÓN PRINCIPAL
# =====================================================================
if __name__ == "__main__":
    try:
        df_res, df_camp = procesar_datos()
        exportar_excel_visual(df_res, df_camp)
        actualizar_dashboard(df_res, df_camp)
        print("\n🎉 ¡Proceso finalizado con éxito! Revisa tu carpeta 'Dashboard'.")
    except Exception as e:
        print(f"❌ Ocurrió un error durante la ejecución: {e}")