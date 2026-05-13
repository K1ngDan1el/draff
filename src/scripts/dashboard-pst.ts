import { initializeApp } from "firebase/app";
import { getFirestore, doc, setDoc, getDoc } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyD9YseEJbs-p_XhA9-_hYlT0sQOh5NwFj4",
  authDomain: "appuestas-ee414.firebaseapp.com",
  projectId: "appuestas-ee414",
  storageBucket: "appuestas-ee414.firebasestorage.app",
  messagingSenderId: "29812166017",
  appId: "1:29812166017:web:9d77a12138245c87486c80",
  measurementId: "G-NG7FP59CPT"
};

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);

let perdidas: number[] = [], ganancias: number[] = [], charts: Record<string, any> = {};
const uploadZone = document.getElementById('uploadZone') as HTMLElement;
const fileInput  = document.getElementById('fileInput') as HTMLInputElement;
const statusEl     = document.getElementById('uploadStatus') as HTMLElement;
uploadZone.addEventListener('dragover', (e: DragEvent) => { e.preventDefault(); uploadZone.classList.add('drag-over'); });
uploadZone.addEventListener('dragleave', () => uploadZone.classList.remove('drag-over'));
uploadZone.addEventListener('drop', (e: DragEvent) => { e.preventDefault(); uploadZone.classList.remove('drag-over'); handleFile((e.dataTransfer as DataTransfer)!.files[0]); });
fileInput.addEventListener('change', (e: Event) => handleFile(((e.target as HTMLInputElement)!).files![0]));
function setStatus(msg: string, isError=false) { statusEl.textContent=msg; statusEl.className='upload-status visible'+(isError?' error':''); }
function handleFile(file: File | undefined) {
  if (!file) return;
  setStatus('Leyendo archivo…');
  const reader = new FileReader();
  reader.onload = (e: any) => {
    try {
      const wb   = (window as any).XLSX.read(e.target.result, { type: 'array' });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const rows = (window as any).XLSX.utils.sheet_to_json(ws, { defval: null });
      const colKeys = Object.keys(rows[0] || {});
      const norm = (s: string) => s.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');
      const findCol = (kws: string[]) => colKeys.find((k: string) => kws.some((kw: string) => norm(k).includes(kw)));
      const colP = findCol(['perdida','loss','perd']) as string | undefined;
      const colG = findCol(['ganancia','profit','gain','gan']) as string | undefined;
      if (!colP && !colG) { setStatus('No se encontraron columnas PÉRDIDAS o GANANCIAS.', true); return; }
      const newP: number[]=[], newG: number[]=[];
      rows.forEach((row: any) => {
        const p = parseFloat(String(row[colP!]??'').replace(/[$,]/g,''));
        const gg = parseFloat(String(row[colG!]??'').replace(/[$,]/g,''));
        if (!isNaN(p) && p>0) newP.push(p);
        if (!isNaN(gg) && gg>0) newG.push(gg);
      });
      perdidas=newP; ganancias=newG;
      setStatus(`✓ ${file.name} · ${rows.length} filas · pérd: "${colP}" · gan: "${colG}"`);
      updateDashboard();
      
      const btnSave = document.getElementById('btnSaveCloud');
      if (btnSave) {
        btnSave.style.display = 'inline-flex';
        btnSave.onclick = async () => {
          const spanSave = btnSave.querySelector('span') || btnSave;
          spanSave.textContent = 'Guardando...';
          try {
            await setDoc(doc(db, "mis_datos", "apuestas"), {
              ganancias: ganancias,
              perdidas: perdidas,
              updatedAt: new Date().toISOString()
            });
            spanSave.textContent = '✓ Guardado en la Nube';
            btnSave.className = 'glass-btn btn-green';
            setTimeout(() => { btnSave.style.display = 'none'; }, 3000);
          } catch (error) {
            console.error("Error saving:", error);
            spanSave.textContent = '❌ Error al guardar';
          }
        };
      }
    } catch(err: unknown) { setStatus('Error: '+(err instanceof Error ? err.message : 'desconocido'), true); }
  };
  reader.readAsArrayBuffer(file);
}

// Load from Firebase on button click
const btnLoad = document.getElementById('btnLoadCloud');
if (btnLoad) {
  btnLoad.onclick = async () => {
    const spanLoad = btnLoad.querySelector('span') || btnLoad;
    const originalText = spanLoad.textContent || 'Cargar Nube';
    spanLoad.textContent = '⏳ Cargando...';
    try {
      const docSnap = await getDoc(doc(db, "mis_datos", "apuestas"));
      if (docSnap.exists()) {
        const data = docSnap.data();
        if (data.ganancias && data.perdidas) {
          ganancias = data.ganancias;
          perdidas = data.perdidas;
          setStatus(`✓ Datos cargados de la nube (${ganancias.length + perdidas.length} registros)`);
          updateDashboard();
          spanLoad.textContent = '✓ Cargado exitosamente';
          setTimeout(() => { spanLoad.textContent = originalText; }, 3000);
        }
      } else {
        setStatus('No se encontraron datos en la nube.');
        spanLoad.textContent = '❌ Sin datos';
        setTimeout(() => { spanLoad.textContent = originalText; }, 3000);
      }
    } catch(e) {
      console.error("Error fetching firebase data", e);
      setStatus('Error al conectar con la nube.');
      spanLoad.textContent = '❌ Error';
      setTimeout(() => { spanLoad.textContent = originalText; }, 3000);
    }
  };
}
const fmt = (n: number) => '$'+n.toLocaleString('es-MX',{minimumFractionDigits:2,maximumFractionDigits:2});
const sum = (a: number[]) => a.reduce((x,y)=>x+y,0);
const avg = (a: number[]) => a.length ? sum(a)/a.length : 0;
const mxv = (a: number[]) => a.length ? Math.max(...a) : 0;

function updateDashboard() {
  const totalG=sum(ganancias), totalP=sum(perdidas), neto=totalG-totalP;
  const netoAbs = Math.abs(neto);
  const g = (window as any).gsap;

  // ── Helper: GSAP count-up (uses Math.round during anim for performance) ──
  function countTo(el: HTMLElement, target: number, prefix: string, dur: number) {
    if (!g) { el.textContent = prefix + target.toLocaleString('es-MX',{minimumFractionDigits:2,maximumFractionDigits:2}); return; }
    const obj = { v: 0 };
    g.to(obj, {
      v: target, duration: dur, ease: 'power2.out',
      onUpdate: () => { el.textContent = prefix + Math.round(obj.v).toLocaleString('es-MX'); },
      onComplete: () => { el.textContent = prefix + target.toLocaleString('es-MX',{minimumFractionDigits:2,maximumFractionDigits:2}); }
    });
  }

  // ── Set ALL text instantly (non-numeric) ──
  (document.getElementById('headerSub') as HTMLElement)!.textContent=`Pérdidas y Ganancias · ${perdidas.length+ganancias.length} movimientos`;
  (document.getElementById('kpiGananciasOps') as HTMLElement)!.textContent=`${ganancias.length} operaciones`;
  (document.getElementById('kpiPerdidasOps') as HTMLElement)!.textContent=`${perdidas.length} operaciones`;
  (document.getElementById('kpiNetoSub') as HTMLElement)!.textContent=neto>=0?'positivo':'negativo';
  (document.getElementById('kpiNetCard') as HTMLElement)!.className='card '+(neto>=0?'positive':'negative');
  (document.getElementById('tableTitle') as HTMLElement)!.textContent=`Datos — ${Math.max(perdidas.length,ganancias.length)} filas`;

  // ── Build data first ──
  buildTable(); buildChartLinea(); buildChartBarras(); buildChartAmbas();

  sessionStorage.setItem('tradingData', JSON.stringify({
    ganancias: { amount: fmt(totalG), badge: `${ganancias.length} operaciones` },
    perdidas:  { amount: fmt(totalP), badge: `${perdidas.length} operaciones` },
    neto:      { amount: (neto>=0?'+':'')+fmt(neto), badge: neto>=0?'positivo':'negativo' }
  }));
  sessionStorage.setItem('rawData', JSON.stringify({ ganancias, perdidas }));

  // ── Animated count-ups (NO blur — simple opacity+y only) ──
  const elG = document.getElementById('kpiGanancias')!;
  const elP = document.getElementById('kpiPerdidas')!;
  const elN = document.getElementById('kpiNeto')!;

  if (g) {
    g.fromTo(elG, {opacity:0,y:12}, {opacity:1,y:0, duration:0.4, delay:0.15, ease:'power2.out'});
    g.fromTo(elP, {opacity:0,y:12}, {opacity:1,y:0, duration:0.4, delay:0.25, ease:'power2.out'});
    g.fromTo(elN, {opacity:0,y:12}, {opacity:1,y:0, duration:0.4, delay:0.35, ease:'power2.out'});
  }
  setTimeout(() => countTo(elG, totalG, '$', 1.0), g ? 150 : 0);
  setTimeout(() => countTo(elP, totalP, '$', 1.0), g ? 250 : 0);
  setTimeout(() => countTo(elN, netoAbs, neto>=0?'+$':'-$', 1.0), g ? 350 : 0);

  // ── Stats: simple opacity + count-up ──
  const statPairs: [string, number][] = [
    ['statMaxG', mxv(ganancias)], ['statPromG', avg(ganancias)],
    ['statMaxP', mxv(perdidas)],  ['statPromP', avg(perdidas)]
  ];
  statPairs.forEach(([id, val], i) => {
    const el = document.getElementById(id)!;
    if (g) g.fromTo(el, {opacity:0,y:6}, {opacity:1,y:0, duration:0.3, delay:0.4+i*0.08, ease:'power2.out'});
    setTimeout(() => countTo(el, val, '$', 0.8), g ? 400 + i * 80 : 0);
  });

  if (!g) return;

  // ══ Lightweight visual animations ══

  // Upload zone: quick flash
  const uz = document.getElementById('uploadZone');
  if (uz) {
    g.to(uz, { boxShadow:'0 0 30px rgba(0,245,139,0.2)', duration:0.25, ease:'power2.out',
      onComplete: () => g.to(uz, { boxShadow:'none', duration:0.5, delay:0.3 })
    });
  }

  // Status badge
  const stEl = document.getElementById('uploadStatus');
  if (stEl) g.fromTo(stEl, {opacity:0,y:8}, {opacity:1,y:0, duration:0.3, delay:0.1, ease:'power2.out'});

  // KPI Cards bounce
  if ((window as any).__disableTilt) (window as any).__disableTilt();
  g.fromTo('.kpis .card', {scale:0.92, opacity:0.4, y:12}, {
    scale:1, opacity:1, y:0, stagger:0.1, duration:0.45, ease:'back.out(1.2)', delay:0.1,
    onComplete: () => {
      document.querySelectorAll<HTMLElement>('.kpis .card').forEach(c => c.style.transform = '');
      if ((window as any).__enableTilt) (window as any).__enableTilt();
    }
  });

  // Panels
  g.fromTo('.panel', {opacity:0.6,y:20}, {opacity:1,y:0, stagger:0.1, duration:0.4, ease:'power2.out', delay:0.2});

  // Table rows — only first 15, rest instant
  setTimeout(() => {
    const rows = document.querySelectorAll('#tableBody tr');
    const limit = Math.min(rows.length, 15);
    for (let i = 0; i < limit; i++) {
      g.fromTo(rows[i], {opacity:0,x:-8}, {opacity:1,x:0, duration:0.2, delay:i*0.02, ease:'power2.out'});
    }
  }, 300);
}

function buildTable() {
  const tbody=document.getElementById('tableBody') as HTMLTableSectionElement;
  const maxLen=Math.max(perdidas.length,ganancias.length);
  let html='';
  for(let i=0;i<maxLen;i++){
    const p=perdidas[i]!=null?`<td class="val-red">${fmt(perdidas[i])}</td>`:'<td class="val-empty">—</td>';
    const gg=ganancias[i]!=null?`<td class="val-green">${fmt(ganancias[i])}</td>`:'<td class="val-empty">—</td>';
    html+=`<tr><td class="val-empty">${i+1}</td>${p}${gg}</tr>`;
  }
  tbody!.innerHTML=html;
}
const CD = {
  responsive:true, animation:{duration:600, easing:'easeOutQuart' as const},
  plugins:{
    legend:{labels:{color:'rgba(255,255,255,0.45)',font:{family:'Inter, sans-serif',size:11},boxWidth:12,padding:16}},
    tooltip:{backgroundColor:'rgba(12,15,20,0.9)',borderColor:'rgba(255,255,255,0.1)',borderWidth:1,titleColor:'rgba(255,255,255,0.5)',bodyColor:'#e2e8f0',cornerRadius:12,padding:12,bodyFont:{family:'Space Grotesk, monospace',size:12},titleFont:{family:'Inter, sans-serif',size:10}}
  },
  scales:{
    x:{ticks:{color:'rgba(255,255,255,0.2)',font:{family:'Inter, sans-serif',size:9},maxTicksLimit:20},grid:{color:'rgba(255,255,255,0.04)'}},
    y:{ticks:{color:'rgba(255,255,255,0.2)',font:{family:'Inter, sans-serif',size:9},callback:(v: number)=>'$'+Math.abs(v).toLocaleString()},grid:{color:'rgba(255,255,255,0.04)'}}
  }
};
function dc(k: string){if(charts[k]){charts[k].destroy();delete charts[k];}}
function buildChartLinea(){
  dc('linea');
  const ctx=(document.getElementById('chartLinea') as HTMLCanvasElement)!.getContext('2d') as CanvasRenderingContext2D;
  let aG=0,aP=0;
  const maxLen=Math.max(perdidas.length,ganancias.length);
  const dG=[],dP=[],labels=[];
  for(let i=0;i<maxLen;i++){
    labels.push('#'+(i+1));
    if(ganancias[i]!=null)aG+=ganancias[i];
    if(perdidas[i]!=null)aP+=perdidas[i];
    dG.push(aG); dP.push(aP);
  }
  charts.linea=new (window as any).Chart(ctx,{type:'line',data:{labels,datasets:[
    {label:'Ganancias acumuladas',data:dG,borderColor:'#00f58b',backgroundColor:'rgba(0,245,139,0.06)',pointRadius:0,borderWidth:2,tension:0.4,fill:true},
    {label:'Pérdidas acumuladas',data:dP,borderColor:'#ff3b5e',backgroundColor:'rgba(255,59,94,0.06)',pointRadius:0,borderWidth:2,tension:0.4,fill:true}
  ]},options:{...CD,interaction:{mode:'index',intersect:false}}});
}
function buildChartBarras(){
  dc('barras');
  const ctx=(document.getElementById('chartBarras') as HTMLCanvasElement)!.getContext('2d') as CanvasRenderingContext2D;
  const allVals=[...perdidas,...ganancias];
  if(!allVals.length)return;
  const maxVal=Math.max(...allVals);
  let rangos;
  if(maxVal<=500)       rangos=[0,25,50,100,200,300,500];
  else if(maxVal<=2000) rangos=[0,50,100,250,500,1000,2000];
  else if(maxVal<=10000)rangos=[0,100,250,500,1000,2500,5000,10000];
  else                  rangos=[0,100,500,1000,2500,5000,10000,Infinity];
  const labels=[],cP=[],cG=[];
  for(let i=0;i<rangos.length-1;i++){
    const lo=rangos[i],hi=rangos[i+1];
    labels.push(hi===Infinity?`$${lo.toLocaleString()}+`:`$${lo.toLocaleString()}–$${hi.toLocaleString()}`);
    cP.push(perdidas.filter((v: number)=>v>lo&&v<=hi).length);
    cG.push(ganancias.filter((v: number)=>v>lo&&v<=hi).length);
  }
  charts.barras=new (window as any).Chart(ctx,{type:'bar',data:{labels,datasets:[
    {label:'Pérdidas',data:cP,backgroundColor:'rgba(255,59,94,0.6)',borderColor:'#ff3b5e',borderWidth:1,borderRadius:6},
    {label:'Ganancias',data:cG,backgroundColor:'rgba(0,245,139,0.6)',borderColor:'#00f58b',borderWidth:1,borderRadius:6}
  ]},options:{...CD,interaction:{mode:'index',intersect:false},
    plugins:{...CD.plugins,tooltip:{...CD.plugins.tooltip,callbacks:{label:(ctx: any)=>` ${ctx.dataset.label}: ${ctx.parsed.y} ops`}}},
    scales:{x:{...CD.scales.x},y:{...CD.scales.y,ticks:{color:'#334155',font:{size:9},callback:(v: number)=>v+' ops'},title:{display:true,text:'Nº de operaciones',color:'#475569',font:{size:9}}}}}});
}
function buildChartAmbas(){
  dc('ambas');
  const ctx=(document.getElementById('chartAmbas') as HTMLCanvasElement)!.getContext('2d') as CanvasRenderingContext2D;
  const maxShow=Math.min(100,Math.max(perdidas.length,ganancias.length));
  const labels=Array.from({length:maxShow},(_,i)=>'#'+(i+1));
  charts.ambas=new (window as any).Chart(ctx,{type:'bar',data:{labels,datasets:[
    {label:'Ganancias',data:ganancias.slice(0,maxShow),backgroundColor:'rgba(0,245,139,0.6)',borderColor:'#00f58b',borderWidth:1,borderRadius:6},
    {label:'Pérdidas',data:perdidas.slice(0,maxShow).map((v: number)=>-v),backgroundColor:'rgba(255,59,94,0.6)',borderColor:'#ff3b5e',borderWidth:1,borderRadius:6}
  ]},options:{...CD,interaction:{mode:'index',intersect:false},
    plugins:{...CD.plugins,tooltip:{...CD.plugins.tooltip,callbacks:{label:(ctx: any)=>` ${ctx.dataset.label}: $${Math.abs(ctx.parsed.y).toLocaleString('es-MX',{minimumFractionDigits:2})}`}}}}});
}
document.querySelectorAll('.chart-tab').forEach((tab: Element)=>{
  tab.addEventListener('click',()=>{
    document.querySelectorAll('.chart-tab').forEach((t: Element)=>t.classList.remove('active'));
    document.querySelectorAll('.chart-panel').forEach((p: Element)=>p.classList.remove('active'));
    (tab as HTMLElement).classList.add('active');
    const tabElement = tab as HTMLElement;
    const tabData = (tabElement as any).dataset?.tab;
    document.getElementById('tab-'+tabData)!.classList.add('active');
    setTimeout(()=>{
      const k=tabData==='linea'?'linea':tabData==='barras'?'barras':'ambas';
      if(charts[k])charts[k].resize();
    },50);
  });
});

// Auto-restore from session if navigating back from trading page
const storedRaw = sessionStorage.getItem('rawData');
if (storedRaw) {
  try {
    const raw = JSON.parse(storedRaw);
    if (raw.ganancias && raw.perdidas && (raw.ganancias.length > 0 || raw.perdidas.length > 0)) {
      ganancias = raw.ganancias;
      perdidas = raw.perdidas;
      updateDashboard();
      
      const btnSave = document.getElementById('btnSaveCloud');
      if (btnSave) {
        btnSave.style.display = 'inline-flex';
        btnSave.className = 'glass-btn btn-cyan';
      }
    }
  } catch(e){}
}

export {};
