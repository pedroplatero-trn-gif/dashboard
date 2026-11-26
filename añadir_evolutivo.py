from pathlib import Path
import json
import re

BASE = Path(__file__).parent
INDEX = BASE / "index.html"

# Datos enero-septiembre del bloque total Contratación Offline CMB/900.
# Octubre y noviembre se añadirán posteriormente desde sus ficheros propios.
DATOS = [
    {"mes": "Ene", "link": 1209, "gclid": 170, "logs": 169, "gads": 175.34,
     "coberturaGclid": 26.9, "coberturaGads": 103.8},
    {"mes": "Feb", "link": 2212, "gclid": 368, "logs": 357, "gads": 393.73,
     "coberturaGclid": 33.4, "coberturaGads": 110.3},
    {"mes": "Mar", "link": 2585, "gclid": 685, "logs": 608, "gads": 634.57,
     "coberturaGclid": 47.6, "coberturaGads": 104.4},
    {"mes": "Abr", "link": 2245, "gclid": 911, "logs": 936, "gads": 814.83,
     "coberturaGclid": 54.6, "coberturaGads": 87.1},
    {"mes": "May", "link": 1725, "gclid": 773, "logs": 736, "gads": 754.86,
     "coberturaGclid": 58.2, "coberturaGads": 102.6},
    {"mes": "Jun", "link": 1921, "gclid": 973, "logs": 911, "gads": 880.98,
     "coberturaGclid": 62.5, "coberturaGads": 96.7},
    {"mes": "Jul", "link": 1878, "gclid": 753, "logs": 723, "gads": 750.37,
     "coberturaGclid": 61.0, "coberturaGads": 103.8},
    {"mes": "Ago", "link": 2266, "gclid": 1121, "logs": 965, "gads": 992.28,
     "coberturaGclid": 57.3, "coberturaGads": 102.8},
    {"mes": "Sep", "link": 3043, "gclid": 1590, "logs": 1521, "gads": 1428.14,
     "coberturaGclid": 60.2, "coberturaGads": 93.9},
]

CSS = r"""
<!-- EVOLUTIVO_CSS_START -->
<style>
.evo-section{max-width:1400px;margin:30px auto;padding:0 20px;color:#17212b}
.evo-heading{margin-bottom:16px}
.evo-eyebrow{font-size:10px;font-weight:800;letter-spacing:.13em;text-transform:uppercase;color:#718096}
.evo-heading h2{font-size:22px;letter-spacing:-.03em;margin:4px 0}
.evo-heading p{font-size:12px;color:#64748b;margin:0}
.evo-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:16px}
.evo-card{min-width:0;background:#fff;border:1px solid #e2e8f0;border-radius:12px;padding:17px;box-shadow:0 4px 10px rgba(0,0,0,.035)}
.evo-card h3{font-size:14px;margin:0 0 3px}
.evo-subtitle{font-size:11px;color:#64748b;margin:0 0 8px}
.evo-chart{width:100%;height:auto;display:block}
.evo-gridline{stroke:#e7ebf1;stroke-width:1}
.evo-axis{fill:#7a8698;font-size:10px}
.evo-month{fill:#69768a;font-size:10px}
.evo-line{fill:none;stroke-width:2.6;stroke-linecap:round;stroke-linejoin:round}
.evo-point{stroke:#fff;stroke-width:1.5}
.evo-legend{display:flex;justify-content:center;gap:18px;flex-wrap:wrap;margin-top:4px;font-size:10px;color:#68758a}
.evo-legend-item{display:inline-flex;align-items:center;gap:6px}
.evo-dot{width:9px;height:9px;border-radius:50%;display:inline-block}
.evo-red{background:#e51c35}.evo-blue{background:#3672bd}
.evo-note{margin:12px 0 0;font-size:10px;line-height:1.5;color:#728094}
@media(max-width:760px){.evo-section{padding:0 12px}.evo-grid{grid-template-columns:1fr}.evo-card{padding:14px}}
</style>
<!-- EVOLUTIVO_CSS_END -->
"""

SECTION = r"""
<!-- EVOLUTIVO_SECTION_START -->
<section class="evo-section" aria-labelledby="evo-title">
  <header class="evo-heading">
    <div class="evo-eyebrow">Evolución · 2026</div>
    <h2 id="evo-title">Evolutivo Mensual</h2>
    <p>Enero a septiembre. Este análisis es independiente de las barras, filtros y CSV del dashboard.</p>
  </header>

  <div class="evo-grid">
    <article class="evo-card">
      <h3>Conversiones LINK vs GCLID</h3>
      <p class="evo-subtitle">Volumen mensual</p>
      <div id="evo-link"></div>
      <div class="evo-legend">
        <span class="evo-legend-item"><i class="evo-dot evo-red"></i>LINK total</span>
        <span class="evo-legend-item"><i class="evo-dot evo-blue"></i>LINK con GCLID</span>
      </div>
    </article>

    <article class="evo-card">
      <h3>Cobertura GCLID</h3>
      <p class="evo-subtitle">LINK con GCLID ÷ LINK sin CE</p>
      <div id="evo-gclid-rate"></div>
      <div class="evo-legend">
        <span class="evo-legend-item"><i class="evo-dot evo-red"></i>Cobertura GCLID</span>
      </div>
    </article>

    <article class="evo-card">
      <h3>LOGS vs conversiones de GAds</h3>
      <p class="evo-subtitle">Volumen mensual</p>
      <div id="evo-logs-gads"></div>
      <div class="evo-legend">
        <span class="evo-legend-item"><i class="evo-dot evo-red"></i>LOGS</span>
        <span class="evo-legend-item"><i class="evo-dot evo-blue"></i>GAds</span>
      </div>
    </article>

    <article class="evo-card">
      <h3>Cobertura de GAds</h3>
      <p class="evo-subtitle">GAds ÷ LOGS</p>
      <div id="evo-gads-rate"></div>
      <div class="evo-legend">
        <span class="evo-legend-item"><i class="evo-dot evo-red"></i>Cobertura GAds vs LOGS</span>
      </div>
    </article>
  </div>

  <p class="evo-note">
    Octubre y noviembre se incorporarán más adelante desde sus ficheros propios.
    Las gráficas no participan en los filtros ni en la exportación CSV del análisis principal.
  </p>
</section>

<script>
const evoData = __DATOS_MENSUALES__;
const evoNumber = n => new Intl.NumberFormat('es-ES',{maximumFractionDigits:0}).format(n);
const evoDecimal = n => new Intl.NumberFormat('es-ES',{maximumFractionDigits:2}).format(n);
const evoPercent = n => new Intl.NumberFormat('es-ES',{
  minimumFractionDigits:1,maximumFractionDigits:1
}).format(n) + '%';

function evoLineChart(targetId, series, percent=false) {
  const W=680,H=250,L=58,R=16,T=14,B=34;
  const plotW=W-L-R,plotH=H-T-B;
  const values=series.flatMap(s=>s.values).filter(Number.isFinite);
  let maxValue=percent
    ? Math.max(125,Math.ceil(Math.max(...values)/25)*25)
    : Math.ceil(Math.max(...values)/500)*500;
  if(!maxValue) maxValue=1;

  const ticks=percent
    ? Array.from({length:Math.floor(maxValue/25)+1},(_,i)=>i*25)
    : [0,.25,.5,.75,1].map(v=>Math.round(maxValue*v));

  const x=i=>L+i*(plotW/(evoData.length-1));
  const y=v=>T+plotH-(v/maxValue)*plotH;
  let svg='<svg class="evo-chart" viewBox="0 0 '+W+' '+H+'" role="img">';

  ticks.forEach(v=>{
    svg+='<line class="evo-gridline" x1="'+L+'" y1="'+y(v)+'" x2="'+(W-R)+'" y2="'+y(v)+'"></line>';
    svg+='<text class="evo-axis" x="'+(L-8)+'" y="'+(y(v)+3)+'" text-anchor="end">'+
      (percent?evoNumber(v)+'%':evoNumber(v))+'</text>';
  });

  evoData.forEach((d,i)=>{
    svg+='<text class="evo-month" x="'+x(i)+'" y="'+(H-10)+'" text-anchor="middle">'+d.mes+'</text>';
  });

  series.forEach(s=>{
    const points=s.values.map((v,i)=>x(i)+','+y(v)).join(' ');
    svg+='<polyline class="evo-line" points="'+points+'" stroke="'+s.color+'"></polyline>';
    s.values.forEach((v,i)=>{
      const shown=percent?evoPercent(v):evoDecimal(v);
      svg+='<circle class="evo-point" cx="'+x(i)+'" cy="'+y(v)+'" r="4.2" fill="'+s.color+'">'+
        '<title>'+evoData[i].mes+' · '+s.label+': '+shown+'</title></circle>';
    });
  });

  svg+='</svg>';
  document.getElementById(targetId).innerHTML=svg;
}

function evoBarChart(targetId) {
  const W=680,H=250,L=58,R=16,T=14,B=34;
  const plotW=W-L-R,plotH=H-T-B;
  const maxValue=Math.ceil(Math.max(...evoData.flatMap(d=>[d.logs,d.gads]))/500)*500;
  const ticks=[0,.25,.5,.75,1].map(v=>Math.round(maxValue*v));
  const groupW=plotW/evoData.length,barW=Math.min(16,groupW*.24);
  const y=v=>T+plotH-(v/maxValue)*plotH;
  let svg='<svg class="evo-chart" viewBox="0 0 '+W+' '+H+'" role="img">';

  ticks.forEach(v=>{
    svg+='<line class="evo-gridline" x1="'+L+'" y1="'+y(v)+'" x2="'+(W-R)+'" y2="'+y(v)+'"></line>';
    svg+='<text class="evo-axis" x="'+(L-8)+'" y="'+(y(v)+3)+'" text-anchor="end">'+evoNumber(v)+'</text>';
  });

  evoData.forEach((d,i)=>{
    const center=L+groupW*(i+.5);
    const hLogs=plotH*d.logs/maxValue,hGads=plotH*d.gads/maxValue;
    svg+='<rect x="'+(center-barW-1)+'" y="'+y(d.logs)+'" width="'+barW+'" height="'+hLogs+
      '" rx="3" fill="#e51c35"><title>'+d.mes+' · LOGS: '+evoNumber(d.logs)+'</title></rect>';
    svg+='<rect x="'+(center+1)+'" y="'+y(d.gads)+'" width="'+barW+'" height="'+hGads+
      '" rx="3" fill="#3672bd"><title>'+d.mes+' · GAds: '+evoDecimal(d.gads)+'</title></rect>';
    svg+='<text class="evo-month" x="'+center+'" y="'+(H-10)+'" text-anchor="middle">'+d.mes+'</text>';
  });

  svg+='</svg>';
  document.getElementById(targetId).innerHTML=svg;
}

evoLineChart('evo-link',[
  {label:'LINK total',values:evoData.map(d=>d.link),color:'#e51c35'},
  {label:'LINK con GCLID',values:evoData.map(d=>d.gclid),color:'#3672bd'}
]);
evoLineChart('evo-gclid-rate',[
  {label:'Cobertura GCLID',values:evoData.map(d=>d.coberturaGclid),color:'#e51c35'}
],true);
evoBarChart('evo-logs-gads');
evoLineChart('evo-gads-rate',[
  {label:'Cobertura GAds vs LOGS',values:evoData.map(d=>d.coberturaGads),color:'#e51c35'}
],true);
</script>
<!-- EVOLUTIVO_SECTION_END -->
"""

def main():
    if not INDEX.exists():
        raise FileNotFoundError(f"No encuentro el index.html: {INDEX}")

    html = INDEX.read_text(encoding="utf-8")

    section = SECTION.replace(
        "__DATOS_MENSUALES__",
        json.dumps(DATOS, ensure_ascii=False)
    )

    # Si el script se ejecuta de nuevo, reemplaza la sección anterior;
    # no duplica estilos ni gráficas.
    html = re.sub(
        r"<!-- EVOLUTIVO_CSS_START -->.*?<!-- EVOLUTIVO_CSS_END -->",
        "",
        html,
        flags=re.S
    )
    html = re.sub(
        r"<!-- EVOLUTIVO_SECTION_START -->.*?<!-- EVOLUTIVO_SECTION_END -->",
        "",
        html,
        flags=re.S
    )

    if not re.search(r"</head\s*>", html, flags=re.I):
        raise ValueError("El index.html no contiene </head>.")
    if not re.search(r"</body\s*>", html, flags=re.I):
        raise ValueError("El index.html no contiene </body>.")

    html = re.sub(
        r"</head\s*>",
        CSS + "\n</head>",
        html,
        count=1,
        flags=re.I
    )
    html = re.sub(
        r"</body\s*>",
        section + "\n</body>",
        html,
        count=1,
        flags=re.I
    )

    INDEX.write_text(html, encoding="utf-8")
    print("Listo: se añadió Evolutivo Mensual al dashboard existente.")
    print(f"Archivo actualizado: {INDEX}")

if __name__ == "__main__":
    main()
