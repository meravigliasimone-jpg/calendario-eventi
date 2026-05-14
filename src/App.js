import React, { useState, useMemo, useCallback, useRef, memo, useEffect } from "react";
import * as XLSX from "xlsx";

// ═══ CONSTANTS ═══
const GIORNI = ["Domenica","Lunedì","Martedì","Mercoledì","Giovedì","Venerdì","Sabato"];
const GIORNI_SHORT = ["LUN","MAR","MER","GIO","VEN","SAB","DOM"];
const MESI = ["Gennaio","Febbraio","Marzo","Aprile","Maggio","Giugno","Luglio","Agosto","Settembre","Ottobre","Novembre","Dicembre"];
const MESI_SHORT = ["OTT","NOV","DIC","GEN","FEB","MAR","APR","MAG","GIU","LUG","AGO","SET","OTT"];
const MESI_IDX = [9,10,11,0,1,2,3,4,5,6,7,8,9];
const MESI_YEAR = [2026,2026,2026,2027,2027,2027,2027,2027,2027,2027,2027,2027,2027];
const START = new Date(2026, 9, 1);
const END = new Date(2027, 9, 31);

const C = { teal:"#00C2CB", magenta:"#FF00FF", yellow:"#FFE700", black:"#000000", white:"#FFFFFF", beige:"#F5F5DC", paper:"#FAF8F0", danger:"#FF3B3B" };
const STATO_BADGE = { "Confermato":{bg:"#d4edda",text:"#155724"}, "In trattativa":{bg:"#fff3cd",text:"#856404"}, "Annullato":{bg:"#f8d7da",text:"#721c24"} };
const SET_BADGE = { "LIVE":{bg:C.magenta,text:C.white}, "DJ SET":{bg:C.teal,text:C.black}, "FORMAT":{bg:C.yellow,text:C.black} };
const PROD_BADGE = { "Nostra":{bg:C.black,text:C.white}, "Co-Prod":{bg:C.magenta,text:C.white}, "Prod Esterna":{bg:C.beige,text:C.black} };

const FIELDS = [
  { key:"artist", label:"ARTIST", type:"text", w:140 },
  { key:"agenzia", label:"AGENZIA", type:"text", w:100 },
  { key:"importo", label:"CACHET €", type:"number", w:90 },
  { key:"costoTicket", label:"COSTO TICKET €", type:"number", w:90 },
  { key:"beProvv", label:"BE PROVV.", type:"computed_be", w:90 },
  { key:"breakEven", label:"BREAK-EVEN", type:"breakeven", w:90 },
  { key:"memoDeal", label:"MEMO DEAL", type:"text", w:130 },
  { key:"chiusaDa", label:"CHIUSA DA", type:"text", w:90 },
  { key:"stato", label:"STATO", type:"select", options:["","Confermato","In trattativa","Annullato"], w:85 },
  { key:"tipoSet", label:"TIPO SET", type:"select", options:["","LIVE","DJ SET","FORMAT"], w:85 },
  { key:"produzione", label:"PRODUZIONE", type:"select", options:["","Nostra","Co-Prod","Prod Esterna"], w:100 },
  { key:"dettagliCoProd", label:"DETTAGLI CO-PROD", type:"text", w:130 },
  { key:"note", label:"NOTE", type:"text", w:130 },
];

const brutBorder = "3px solid #000";
const brutShadow = "4px 4px 0px #000";

// ═══ HELPERS ═══
function fmt(d) { try { return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`; } catch(e) { return "—"; } }
function meseAnno(d) { return `${MESI[d.getMonth()]} ${d.getFullYear()}`; }
function isWe(d) { return d.getDay()===0||d.getDay()===6; }
function isInv(d) { const m=d.getMonth(); return m>=9||m<=3; }
function isEst(d) { const m=d.getMonth(); return m>=4&&m<=8; }
const eur = n => new Intl.NumberFormat('it-IT',{style:'currency',currency:'EUR'}).format(n);

function emptySlot(dayId, idx) {
  return { slotId:`${dayId}_${idx}`, artist:"",agenzia:"",importo:"",costoTicket:"",beIva:22,beSiae:10,beComm:5,memoDeal:"",chiusaDa:"",stato:"",tipoSet:"",produzione:"",dettagliCoProd:"",note:"",consuntivo:null };
}

function genDays() {
  const days = []; let d = new Date(START);
  while (d <= END) {
    const id = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
    days.push({ id, date: new Date(d), giorno: GIORNI[d.getDay()], slots: [emptySlot(id, 0)] });
    d.setDate(d.getDate()+1);
  }
  return days;
}

// ═══ EXPORT XLSX ═══
function exportData(days, stats) {
  const headers=["Giorno","Data","Artist","Agenzia","Cachet €","Costo Ticket €","BE Provvisorio","Memo Deal","Chiusa da","Stato","Tipo Set","Produzione","Dettagli Co-Prod","Note"];
  const makeRows = (list) => list.flatMap(d=>d.slots.filter(s=>s.artist).map(s=>{
    const imp=parseFloat(s.importo)||0;const tkt=parseFloat(s.costoTicket)||0;
    const totPerc=((s.beIva||0)+(s.beSiae||0)+(s.beComm||0))/100;
    const net=tkt*(1-totPerc);
    const beP=(imp&&tkt&&net>0)?Math.ceil(imp/net):"";
    return [d.giorno,fmt(d.date),s.artist,s.agenzia,s.importo||"",s.costoTicket||"",beP?beP+" pax":"",s.memoDeal,s.chiusaDa,s.stato,s.tipoSet,s.produzione,s.dettagliCoProd,s.note];
  }));
  const allRows = makeRows(days);
  const invRows = makeRows(days.filter(d=>isInv(d.date)&&d.slots.some(s=>s.artist)));
  const estRows = makeRows(days.filter(d=>isEst(d.date)&&d.slots.some(s=>s.artist)));
  const recapRows = [
    ["RIEPILOGO STAGIONE"],[""],
    ["Totale",stats.tot],["Eventi",stats.n],["Confermati",stats.conf],["In trattativa",stats.tratt],
    ["Media",stats.avg?Math.round(stats.avg):""],[""],
    ["INVERNALE",stats.tInv],["ESTIVO",stats.tEst],[""],
    ["PER PRODUZIONE"],
    ["Nostra",stats.prodNostra.n+" ev.",stats.prodNostra.tot],
    ["Co-Prod",stats.prodCoProd.n+" ev.",stats.prodCoProd.tot],
    ["Esterna",stats.prodEsterna.n+" ev.",stats.prodEsterna.tot],
  ];
  const wb = XLSX.utils.book_new();
  const ws1 = XLSX.utils.aoa_to_sheet([headers,...allRows]);
  const ws2 = XLSX.utils.aoa_to_sheet([headers,...invRows]);
  const ws3 = XLSX.utils.aoa_to_sheet([headers,...estRows]);
  const ws4 = XLSX.utils.aoa_to_sheet(recapRows);
  XLSX.utils.book_append_sheet(wb,ws1,"Calendario");
  XLSX.utils.book_append_sheet(wb,ws2,"Invernale");
  XLSX.utils.book_append_sheet(wb,ws3,"Estivo");
  XLSX.utils.book_append_sheet(wb,ws4,"Recap");
  XLSX.writeFile(wb,"Magnolia_Calendario.xlsx");
}

// ═══ PARAMS ═══
const INITIAL_PARAMS = {
  costiProduzione: [
    { voce:"Costo Produzione - Palco Main", costo:3700 },
    { voce:"Costo Produzione - Park Stage", costo:2000 },
    { voce:"Costo Produzione - Second Stage", costo:1500 },
  ],
  costiArtistici: [
    { voce:"Compenso Artista", costo:0 },{ voce:"Compenso Supporto", costo:0 },
    { voce:"Tourbus 12.00-00.00", costo:100 },{ voce:"Hotel con colazione", costo:0 },
    { voce:"Cena Magnolia", costo:15 },{ voce:"Pranzo Magnolia", costo:13 },
    { voce:"Accoglienza Magnolia", costo:10 },{ voce:"Buoni Pizza", costo:10 },
    { voce:"Spesa Camerini (da rider)", costo:0 },{ voce:"Spesa Extra", costo:0 },
    { voce:"Free Drink Produzione", costo:5 },
  ],
  costiPersonale: [
    { voce:"Baristi", costo:90 },{ voce:"Cassieri Bar", costo:80 },
    { voce:"Sicurezza Extra (1:150)", costo:130 },{ voce:"Sicurezza Extra (on top)", costo:130 },
    { voce:"Orario Extra Sicurezza", costo:15 },{ voce:"Orario Extra Tecnici", costo:25 },
    { voce:"Tecnico Extra", costo:250 },{ voce:"Extra Tecnico Setup+Suoni", costo:100 },
    { voce:"Cassieri / Controllo Accessi", costo:80 },{ voce:"Produzione in Giornata", costo:200 },
    { voce:"Pre-Produzione", costo:200 },{ voce:"Driver", costo:0 },
    { voce:"Facchini", costo:220 },{ voce:"Addetti Pulizie", costo:20 },
  ],
  costiAllestimento: [
    { voce:"Generatore (orario)", costo:20 },{ voce:"Polizia Locale (>2500)", costo:1050 },
    { voce:"Ambulanza (120€/h)", costo:120 },{ voce:"Vigili del Fuoco (>1500)", costo:530 },
    { voce:"Noleggio Tecnico", costo:0 },{ voce:"Ledwall", costo:2100 },
    { voce:"Generatore", costo:200 },{ voce:"Assicurazione Maltempo", costo:0 },
    { voce:"Advertising", costo:0 },{ voce:"Fotografo", costo:0 },{ voce:"Truccatrice", costo:0 },
  ],
  costiRicavi: [
    { voce:"SIAE", perc:10 },{ voce:"Aggio", perc:0 },
    { voce:"Commissioni POS", perc:1 },{ voce:"Consumi Bar", perc:25 },
  ],
  ripartizione: { magnolia: 50, partner: 50 },
  commissioniTicket: 5,
};

function defaultConsuntivo(params) {
  const p = params || INITIAL_PARAMS;
  return {
    costiProduzione: p.costiProduzione.map(v => ({...v, qt:0})),
    costiArtistici: p.costiArtistici.map((v,i) => ({...v, qt:i===0?1:0})),
    costiPersonale: p.costiPersonale.map(v => ({...v, qt:0})),
    costiAllestimento: p.costiAllestimento.map(v => ({...v, qt:0})),
    costiRicavi: p.costiRicavi.map(v => ({voce:v.voce, perc:v.perc/100})),
    costiPromoter1: { voce:"Costi 1° Promoter", costo:0 },
    costiPromoter2: { voce:"Costi 2° Promoter", costo:0 },
    incassi: { ticketPrice:0, prevenditeDice:0, prevenditeTicketone:0, prevenditeTicketmaster:0, cassaDice:0, cassa:0, incassiBar:0 },
    ripartizione: { magnolia: p.ripartizione.magnolia/100, partner: p.ripartizione.partner/100 },
  };
}

function sumSection(items) { return items.reduce((s,i) => s + (i.costo||0) * (i.qt||0), 0); }
function calcTotalCosti(c) {
  return sumSection(c.costiProduzione) + sumSection(c.costiArtistici) + sumSection(c.costiPersonale) + sumSection(c.costiAllestimento) + (c.costiPromoter1.costo||0) + (c.costiPromoter2.costo||0);
}
function calcTotalIncassi(c) {
  const i = c.incassi;
  const totalTickets = (i.prevenditeDice||0)+(i.prevenditeTicketone||0)+(i.prevenditeTicketmaster||0)+(i.cassaDice||0)+(i.cassa||0);
  return { ticketRevenue: totalTickets * (i.ticketPrice||0), bar: i.incassiBar||0, totalTickets };
}
function calcBE(c) {
  if (!c) return "";
  const totalCosti = calcTotalCosti(c);
  const tp = c.incassi.ticketPrice || 0;
  if (!tp) return "";
  const totalPerc = c.costiRicavi.reduce((s,r) => s + (r.perc||0), 0);
  const netTicket = tp * (1 - totalPerc);
  if (netTicket <= 0) return "";
  return Math.ceil(totalCosti / netTicket);
}

// ═══ CONSUNTIVO MODAL ═══
function ConsuntivoModal({ day, consuntivo, params, onSave, onClose }) {
  const [c, setC] = useState(() => {
    const base = consuntivo || defaultConsuntivo(params);
    if (day.importo && base.costiArtistici[0].costo === 0) {
      return {...base, costiArtistici: base.costiArtistici.map((item,i) => i===0 ? {...item, costo: parseFloat(day.importo)||0, qt:1} : item)};
    }
    return base;
  });

  const updateItem = (section, idx, field, value) => {
    setC(prev => ({...prev, [section]: prev[section].map((item,i) => i===idx ? {...item, [field]: field==="voce"?value:(value===""?0:parseFloat(value)||0)} : item)}));
  };

  const totalCosti = calcTotalCosti(c);
  const { ticketRevenue, bar, totalTickets } = calcTotalIncassi(c);
  const costiSuRicavi = c.costiRicavi.reduce((s,r) => s + (r.perc||0) * ticketRevenue, 0);
  const totalIncassiNetto = ticketRevenue + bar - costiSuRicavi;
  const utile = totalIncassiNetto - totalCosti;
  const be = calcBE(c);

  const SH = ({title, total, bg}) => (
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"8px 12px",background:bg||"#f5f5f0",border:"2px solid #000",marginTop:14}}>
      <span style={{fontWeight:700,fontSize:12,textTransform:"uppercase",letterSpacing:.5}}>{title}</span>
      <span style={{fontFamily:"'Space Mono',monospace",fontWeight:700,fontSize:13}}>€ {(total||0).toLocaleString('it-IT')}</span>
    </div>
  );

  const IR = ({item, section, idx}) => (
    <div style={{display:"flex",gap:4,alignItems:"center",borderBottom:"1px solid #eee",padding:"3px 0"}}>
      <div style={{flex:3,fontSize:11,fontWeight:600,padding:"3px 6px"}}>{item.voce}</div>
      <input value={item.costo===0?"":item.costo||""} onChange={e=>updateItem(section,idx,"costo",e.target.value)} placeholder="€"
        style={{flex:1,border:"1px solid #ccc",padding:"3px 5px",fontFamily:"'Space Mono'",fontSize:11,fontWeight:700,textAlign:"right"}}/>
      <span style={{fontSize:9,color:"#999",width:8}}>×</span>
      <input value={item.qt===0?"":item.qt||""} onChange={e=>updateItem(section,idx,"qt",e.target.value)} placeholder="Qt"
        style={{width:40,border:"1px solid #ccc",padding:"3px 4px",fontFamily:"'Space Mono'",fontSize:11,textAlign:"center"}}/>
      <div style={{width:75,textAlign:"right",fontFamily:"'Space Mono'",fontSize:11,fontWeight:700,padding:"3px 6px",color:(item.costo||0)*(item.qt||0)>0?"#000":"#ccc"}}>
        € {((item.costo||0)*(item.qt||0)).toLocaleString('it-IT')}
      </div>
    </div>
  );

  const numInput = (val, onChange, style2) => (
    <input value={val===0?"":val||""} onChange={e=>{const v=e.target.value;onChange(v===""?0:parseFloat(v)||0);}}
      style={{border:"1px solid #ccc",padding:"4px 6px",fontFamily:"'Space Mono'",fontSize:12,fontWeight:700,textAlign:"right",...style2}}/>
  );

  return (
    <div style={{position:"fixed",inset:0,background:"#000000bb",display:"flex",alignItems:"flex-start",justifyContent:"center",zIndex:200,overflowY:"auto",padding:"20px 0"}}
      onClick={e=>e.target===e.currentTarget&&onClose()}>
      <div style={{background:C.paper,border:"3px solid #000",boxShadow:"6px 6px 0 #000",padding:0,width:880,maxWidth:"95vw",margin:"20px auto"}}>
        <div style={{background:C.black,padding:"14px 18px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div>
            <div style={{color:"#888",fontWeight:700,fontSize:9,letterSpacing:2}}>CONSUNTIVO EVENTO</div>
            <div style={{color:C.white,fontWeight:700,fontSize:18,marginTop:2}}>{day.artist || "—"}</div>
            <div style={{color:C.teal,fontFamily:"'Space Mono'",fontSize:12,fontWeight:700}}>{fmt(day.date)}</div>
          </div>
          <div style={{textAlign:"right"}}>
            <div style={{color:"#888",fontSize:9,fontWeight:700,letterSpacing:1}}>BREAK-EVEN</div>
            <div style={{color:C.yellow,fontFamily:"'Space Mono'",fontSize:28,fontWeight:700,lineHeight:1}}>{be||"—"}</div>
            <div style={{color:C.teal,fontSize:10,fontWeight:700}}>{be?"PAX":""}</div>
          </div>
        </div>

        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:0}}>
          <div style={{padding:"0 14px 14px",borderRight:"2px solid #000"}}>
            <SH title="COSTI PRODUZIONE" total={sumSection(c.costiProduzione)} bg={C.yellow+"44"}/>
            {c.costiProduzione.map((item,i) => <IR key={i} item={item} section="costiProduzione" idx={i}/>)}
            <SH title="COSTI ARTISTICI" total={sumSection(c.costiArtistici)} bg={C.magenta+"22"}/>
            {c.costiArtistici.map((item,i) => <IR key={i} item={item} section="costiArtistici" idx={i}/>)}
            <SH title="COSTI PERSONALE" total={sumSection(c.costiPersonale)} bg={C.teal+"22"}/>
            {c.costiPersonale.map((item,i) => <IR key={i} item={item} section="costiPersonale" idx={i}/>)}
            <SH title="COSTI ALLESTIMENTO" total={sumSection(c.costiAllestimento)} bg={C.yellow+"22"}/>
            {c.costiAllestimento.map((item,i) => <IR key={i} item={item} section="costiAllestimento" idx={i}/>)}
            <SH title="PROMOTER" total={(c.costiPromoter1.costo||0)+(c.costiPromoter2.costo||0)}/>
            <div style={{display:"flex",gap:6,padding:"6px 0"}}>
              <div style={{flex:1}}><div style={{fontSize:10,fontWeight:700,marginBottom:3}}>1° PROMOTER</div>{numInput(c.costiPromoter1.costo,v=>setC({...c,costiPromoter1:{...c.costiPromoter1,costo:v}}),{width:"100%"})}</div>
              <div style={{flex:1}}><div style={{fontSize:10,fontWeight:700,marginBottom:3}}>2° PROMOTER</div>{numInput(c.costiPromoter2.costo,v=>setC({...c,costiPromoter2:{...c.costiPromoter2,costo:v}}),{width:"100%"})}</div>
            </div>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 12px",background:C.black,border:"2px solid #000",marginTop:12}}>
              <span style={{color:C.white,fontWeight:700,fontSize:12}}>COSTI TOTALI</span>
              <span style={{color:C.danger,fontFamily:"'Space Mono'",fontWeight:700,fontSize:16}}>€ {totalCosti.toLocaleString('it-IT')}</span>
            </div>
          </div>

          <div style={{padding:"0 14px 14px"}}>
            <SH title="BIGLIETTERIA" total={ticketRevenue} bg={C.teal+"22"}/>
            <div style={{padding:"6px 0"}}>
              <div style={{fontSize:10,fontWeight:700,marginBottom:4}}>PREZZO BIGLIETTO</div>
              {numInput(c.incassi.ticketPrice,v=>setC({...c,incassi:{...c.incassi,ticketPrice:v}}),{width:"100%",border:brutBorder,padding:"8px",fontSize:15})}
              <div style={{fontSize:9,color:"#888",marginTop:2}}>Netto: € {((c.incassi.ticketPrice||0)*(1-c.costiRicavi.reduce((s,r)=>s+(r.perc||0),0))).toFixed(2)} (- {(c.costiRicavi.reduce((s,r)=>s+(r.perc||0),0)*100).toFixed(0)}%)</div>
            </div>
            {[{k:"prevenditeDice",l:"Prevendite DICE"},{k:"prevenditeTicketone",l:"Prevendite Ticketone"},{k:"prevenditeTicketmaster",l:"Prevendite Ticketmaster"},{k:"cassaDice",l:"Cassa DICE"},{k:"cassa",l:"Cassa"}].map(f=>(
              <div key={f.k} style={{display:"flex",justifyContent:"space-between",alignItems:"center",borderBottom:"1px solid #eee",padding:"4px 0"}}>
                <span style={{fontSize:11,fontWeight:600}}>{f.l}</span>
                {numInput(c.incassi[f.k],v=>setC({...c,incassi:{...c.incassi,[f.k]:v}}),{width:70})}
              </div>
            ))}
            <div style={{display:"flex",justifyContent:"space-between",padding:"6px 0",borderTop:"2px solid #000",marginTop:4}}>
              <span style={{fontWeight:700,fontSize:12}}>INGRESSI TOTALI</span>
              <span style={{fontFamily:"'Space Mono'",fontWeight:700,fontSize:13}}>{totalTickets}</span>
            </div>
            <SH title="BAR" total={bar} bg={C.yellow+"22"}/>
            <div style={{padding:"6px 0"}}>{numInput(c.incassi.incassiBar,v=>setC({...c,incassi:{...c.incassi,incassiBar:v}}),{width:"100%"})}</div>
            <SH title="COSTI SU RICAVI" total={costiSuRicavi} bg={C.danger+"11"}/>
            {c.costiRicavi.map((r,i) => (
              <div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",borderBottom:"1px solid #eee",padding:"4px 0"}}>
                <span style={{fontSize:11,fontWeight:600}}>{r.voce}</span>
                <span style={{fontFamily:"'Space Mono'",fontSize:11,fontWeight:700}}>€ {Math.round((r.perc||0)*ticketRevenue).toLocaleString('it-IT')}</span>
              </div>
            ))}
            <SH title="RIPARTIZIONE" total={utile} bg={utile>=0?C.teal+"22":C.danger+"22"}/>
            <div style={{display:"flex",gap:8,padding:"8px 0",textAlign:"center"}}>
              <div style={{flex:1}}>
                <div style={{fontSize:9,fontWeight:700,letterSpacing:1,marginBottom:3}}>MAGNOLIA</div>
                <input value={Math.round(c.ripartizione.magnolia*100)} onChange={e=>{const v=(parseFloat(e.target.value)||0)/100;setC({...c,ripartizione:{magnolia:v,partner:1-v}});}}
                  style={{width:50,border:"2px solid #000",padding:"5px",fontFamily:"'Space Mono'",fontSize:14,fontWeight:700,textAlign:"center"}}/>
                <span style={{fontWeight:700}}> %</span>
                <div style={{fontFamily:"'Space Mono'",fontSize:14,fontWeight:700,marginTop:4,color:utile>=0?"#007a6a":C.danger}}>{eur(Math.round(utile*(c.ripartizione.magnolia||0)))}</div>
              </div>
              <div style={{width:2,background:"#000"}}/>
              <div style={{flex:1}}>
                <div style={{fontSize:9,fontWeight:700,letterSpacing:1,marginBottom:3}}>PARTNER</div>
                <input value={Math.round(c.ripartizione.partner*100)} onChange={e=>{const v=(parseFloat(e.target.value)||0)/100;setC({...c,ripartizione:{partner:v,magnolia:1-v}});}}
                  style={{width:50,border:"2px solid #000",padding:"5px",fontFamily:"'Space Mono'",fontSize:14,fontWeight:700,textAlign:"center"}}/>
                <span style={{fontWeight:700}}> %</span>
                <div style={{fontFamily:"'Space Mono'",fontSize:14,fontWeight:700,marginTop:4,color:utile>=0?"#007a6a":C.danger}}>{eur(Math.round(utile*(c.ripartizione.partner||0)))}</div>
              </div>
            </div>
            <div style={{background:C.black,border:"2px solid #000",padding:"12px",marginTop:12}}>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
                <span style={{color:C.white,fontWeight:700,fontSize:11}}>INCASSI NETTO</span>
                <span style={{color:C.teal,fontFamily:"'Space Mono'",fontWeight:700,fontSize:13}}>€ {Math.round(totalIncassiNetto).toLocaleString('it-IT')}</span>
              </div>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
                <span style={{color:C.white,fontWeight:700,fontSize:11}}>COSTI TOTALI</span>
                <span style={{color:C.danger,fontFamily:"'Space Mono'",fontWeight:700,fontSize:13}}>€ {totalCosti.toLocaleString('it-IT')}</span>
              </div>
              <div style={{height:2,background:C.yellow,margin:"6px 0"}}/>
              <div style={{display:"flex",justifyContent:"space-between"}}>
                <span style={{color:C.yellow,fontWeight:700,fontSize:13}}>UTILE</span>
                <span style={{color:utile>=0?C.teal:C.danger,fontFamily:"'Space Mono'",fontWeight:700,fontSize:20}}>{utile>=0?"+":""}€ {Math.round(utile).toLocaleString('it-IT')}</span>
              </div>
              <div style={{display:"flex",justifyContent:"space-between",marginTop:6}}>
                <span style={{color:"#888",fontWeight:700,fontSize:11}}>BREAK-EVEN</span>
                <span style={{color:C.yellow,fontFamily:"'Space Mono'",fontWeight:700,fontSize:16}}>{be||"—"} {be?"PAX":""}</span>
              </div>
            </div>
          </div>
        </div>
        <div style={{display:"flex",gap:8,justifyContent:"flex-end",padding:"14px 18px",borderTop:"2px solid #000",background:C.white}}>
          <button onClick={onClose} style={{border:brutBorder,background:C.white,padding:"8px 20px",fontWeight:700,fontSize:13,cursor:"pointer",boxShadow:brutShadow}}>ANNULLA</button>
          <button onClick={()=>onSave(c)} style={{border:brutBorder,background:C.teal,padding:"8px 20px",fontWeight:700,fontSize:13,cursor:"pointer",boxShadow:brutShadow}}>SALVA</button>
        </div>
      </div>
    </div>
  );
}

// ═══ CALENDAR PICKER (for quick add modal) ═══
function CalPicker({ days, selected, onSelect }) {
  const [pm, setPm] = useState(0);
  const month=MESI_IDX[pm], year=MESI_YEAR[pm];
  const fd=new Date(year,month,1); let sw=fd.getDay(); sw=sw===0?6:sw-1;
  const dim=new Date(year,month+1,0).getDate();
  const cells=[]; for(let i=0;i<sw;i++)cells.push(null); for(let d=1;d<=dim;d++)cells.push(d);
  const md=days.filter(d=>d.date.getMonth()===month&&d.date.getFullYear()===year);
  const he={}; md.forEach(d=>{if(d.slots&&d.slots.some(s=>s.artist))he[d.date.getDate()]=true;});
  return (
    <div style={{border:"2px solid #000",background:C.white,padding:12}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
        <button onClick={()=>setPm(Math.max(0,pm-1))} style={{border:"2px solid #000",background:pm===0?"#eee":C.yellow,width:32,height:32,fontWeight:700,fontSize:16,cursor:"pointer",opacity:pm===0?.3:1}}>←</button>
        <span style={{fontWeight:700,fontSize:14,textTransform:"uppercase",letterSpacing:1}}>{MESI[month]} {year}</span>
        <button onClick={()=>setPm(Math.min(MESI_IDX.length-1,pm+1))} style={{border:"2px solid #000",background:pm===MESI_IDX.length-1?"#eee":C.yellow,width:32,height:32,fontWeight:700,fontSize:16,cursor:"pointer",opacity:pm===MESI_IDX.length-1?.3:1}}>→</button>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:2,marginBottom:4}}>
        {GIORNI_SHORT.map(g=><div key={g} style={{textAlign:"center",fontSize:9,fontWeight:700,color:"#888",padding:3}}>{g}</div>)}
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:2}}>
        {cells.map((day,i)=>{
          if(!day) return <div key={`e${i}`}/>;
          const ds=`${year}-${String(month+1).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
          const isSel=selected===ds;
          const we=new Date(year,month,day).getDay()===0||new Date(year,month,day).getDay()===6;
          return <button key={i} onClick={()=>onSelect(ds)} style={{border:isSel?"3px solid "+C.magenta:"2px solid "+(he[day]?C.teal:"#ddd"),background:isSel?C.magenta:he[day]?C.teal+"22":C.white,color:isSel?C.white:we?C.magenta:C.black,fontFamily:"'Space Mono'",fontWeight:700,fontSize:12,padding:"6px 0",cursor:"pointer",textAlign:"center",position:"relative"}}>{day}{he[day]&&!isSel&&<span style={{position:"absolute",bottom:1,left:"50%",transform:"translateX(-50%)",width:4,height:4,borderRadius:"50%",background:C.teal}}/>}</button>;
        })}
      </div>
      {selected&&(()=>{const [y,m,dd]=selected.split("-").map(Number);const dt=new Date(y,m-1,dd);return <div style={{marginTop:8,padding:"6px 10px",background:C.yellow,border:"2px solid #000",fontWeight:700,fontSize:12,textAlign:"center"}}>{GIORNI[dt.getDay()]} {fmt(dt)}</div>;})()}
    </div>
  );
}


// ═══ SLOT ROW (memoized) ═══
const SlotRow = memo(function SlotRow({day, slot, si, bePopup, setBePopup, update, addSlot, removeSlot, setDetailDay, setConsModal, dark}) {
  return (
    <div style={{display:"flex",alignItems:"center",borderBottom:si===day.slots.length-1?"2px solid #000":"1px dashed #bbb",background:isWe(day.date)?(dark?"#2d2040":"#f0e6ff"):(dark?"#1e1e1e":"#FFFFFF"),transition:"background .1s"}}
      onMouseEnter={e=>{if(!isWe(day.date))e.currentTarget.style.background=dark?"#2a2a2a":"#F5F5DC";}} onMouseLeave={e=>{e.currentTarget.style.background=isWe(day.date)?(dark?"#2d2040":"#f0e6ff"):(dark?"#1e1e1e":"#FFFFFF");}}>
      {si===0?(
        <><div style={{width:55,padding:"6px 5px",fontSize:10,fontWeight:isWe(day.date)?800:600,color:isWe(day.date)?"#FF00FF":"#555"}}>{day.giorno.slice(0,3).toUpperCase()}</div>
        <div style={{width:72,padding:"6px 5px",fontSize:11,fontFamily:"'Space Mono'",fontWeight:700,cursor:"pointer",textDecoration:"underline",textDecorationStyle:"dotted",textUnderlineOffset:3}} onClick={()=>setDetailDay(day.id)}>{fmt(day.date)}</div></>
      ):(
        <><div style={{width:55,padding:"6px 5px"}}/><div style={{width:72,padding:"6px 5px",fontSize:9,color:"#999",fontWeight:600}}>↳ slot {si+1}</div></>
      )}
      {FIELDS.map(f=>(
        <div key={f.key} style={{width:f.w,padding:"4px 3px"}}>
          {f.type==="computed_be"?(()=>{
            const imp=parseFloat(slot.importo)||0;const tkt=parseFloat(slot.costoTicket)||0;
            const totPerc=((slot.beIva||0)+(slot.beSiae||0)+(slot.beComm||0))/100;
            const net=tkt*(1-totPerc);const be=(imp&&tkt&&net>0)?Math.ceil(imp/net):"";
            return (<div style={{position:"relative"}}>
              <button onClick={()=>setBePopup(bePopup===slot.slotId?null:slot.slotId)} style={{width:"100%",border:"1.5px solid "+(be?"#000":"#ddd"),padding:"3px 5px",fontFamily:"'Space Mono'",fontWeight:700,fontSize:10,background:be?(dark?"#3a3000":"#fff3d0"):(dark?"#2a2a2a":"#FFFFFF"),color:dark?"#f0f0f0":"#000",cursor:"pointer",textAlign:"center"}}>{be?`${be} pax`:"—"}</button>
              {bePopup===slot.slotId&&(<div data-be-popup="true" style={{position:"absolute",top:"100%",left:0,zIndex:50,background:"#FFFFFF",border:brutBorder,boxShadow:brutShadow,padding:10,width:180,marginTop:3}}>
                <div style={{fontSize:9,fontWeight:700,letterSpacing:1,marginBottom:6,color:"#666"}}>PERCENTUALI</div>
                {[{k:"beIva",l:"IVA"},{k:"beSiae",l:"SIAE"},{k:"beComm",l:"COMM."}].map(p=>(
                  <div key={p.k} style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
                    <span style={{fontSize:10,fontWeight:700}}>{p.l}</span>
                    <div style={{display:"flex",alignItems:"center",gap:2}}>
                      <input value={slot[p.k]===0?"0":slot[p.k]||""} onChange={e=>{const v=e.target.value;update(slot.slotId,p.k,v===""?0:parseFloat(v)||0);}}
                        style={{width:40,border:"2px solid #000",padding:"2px 3px",fontFamily:"'Space Mono'",fontSize:11,fontWeight:700,textAlign:"center"}}/>
                      <span style={{fontWeight:700,fontSize:11}}>%</span></div></div>))}
                <div style={{borderTop:"2px solid #000",paddingTop:4,marginTop:4,display:"flex",justifyContent:"space-between"}}>
                  <span style={{fontSize:10,fontWeight:700}}>TOT</span>
                  <span style={{fontFamily:"'Space Mono'",fontSize:11,fontWeight:700}}>{(slot.beIva||0)+(slot.beSiae||0)+(slot.beComm||0)}%</span></div>
                <button onClick={()=>setBePopup(null)} style={{width:"100%",marginTop:6,border:"2px solid #000",background:"#FFE700",padding:"3px",fontWeight:700,fontSize:10,cursor:"pointer"}}>OK</button>
              </div>)}</div>);
          })():f.type==="breakeven"?(
            <button onClick={()=>setConsModal(slot.slotId)} style={{width:"100%",border:"1.5px solid "+(slot.consuntivo?"#000":"#ccc"),padding:"3px 5px",fontFamily:"'Space Mono'",fontWeight:700,fontSize:10,background:slot.consuntivo?"#FFE700":(dark?"#2a2a2a":"#FFFFFF"),color:slot.consuntivo?"#000":(dark?"#f0f0f0":"#000"),cursor:"pointer",textAlign:"center"}}
              onMouseEnter={e=>{e.currentTarget.style.background="#00C2CB";}} onMouseLeave={e=>{e.currentTarget.style.background=slot.consuntivo?"#FFE700":(dark?"#2a2a2a":"#FFFFFF");}}>
              {slot.consuntivo?`${calcBE(slot.consuntivo)||"—"} pax`:"CALCOLA"}</button>
          ):f.type==="select"?(
            <select value={slot[f.key]} onChange={e=>update(slot.slotId,f.key,e.target.value)} style={{width:"100%",border:"1.5px solid "+(slot[f.key]?"#000":"#ccc"),padding:"3px 4px",fontWeight:700,fontSize:10,
              background:f.key==="stato"&&STATO_BADGE[slot[f.key]]?STATO_BADGE[slot[f.key]].bg:f.key==="tipoSet"&&SET_BADGE[slot[f.key]]?SET_BADGE[slot[f.key]].bg:f.key==="produzione"&&PROD_BADGE[slot[f.key]]?PROD_BADGE[slot[f.key]].bg:"#FFFFFF",
              color:f.key==="stato"&&STATO_BADGE[slot[f.key]]?STATO_BADGE[slot[f.key]].text:f.key==="tipoSet"&&SET_BADGE[slot[f.key]]?SET_BADGE[slot[f.key]].text:f.key==="produzione"&&PROD_BADGE[slot[f.key]]?PROD_BADGE[slot[f.key]].text:"#000",
              textTransform:"uppercase"}}>
              {f.options.map(o=><option key={o} value={o}>{o||"—"}</option>)}</select>
          ):(
            <input value={slot[f.key]} onChange={e=>update(slot.slotId,f.key,e.target.value)} placeholder="—" style={{width:"100%",border:"1.5px solid "+(slot[f.key]?"#000":"#ddd"),padding:"3px 5px",
              fontFamily:(f.key==="importo"||f.key==="costoTicket")?"'Space Mono'":"inherit",fontWeight:f.key==="artist"&&slot[f.key]?700:(f.key==="importo"||f.key==="costoTicket")?700:500,fontSize:11,
              background:f.key==="importo"&&slot[f.key]?(dark?"#0a3a2a":"#d0fff0"):f.key==="costoTicket"&&slot[f.key]?(dark?"#3a3000":"#fff3d0"):(dark?"#2a2a2a":"#FFFFFF"),color:dark?"#f0f0f0":"#000"}}/>
          )}
        </div>
      ))}
      <div style={{width:22,padding:"4px 1px",display:"flex",flexDirection:"column",gap:1}}>
        {si===0&&<button onClick={()=>addSlot(day.id)} title="Aggiungi serata" style={{border:"1px solid #ccc",background:dark?"#333":"#FFFFFF",width:20,height:20,fontSize:13,fontWeight:700,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",color:dark?"#aaa":"#666"}} onMouseEnter={e=>{e.currentTarget.style.background="#00C2CB";e.currentTarget.style.color="#fff";}} onMouseLeave={e=>{e.currentTarget.style.background=dark?"#333":"#FFFFFF";e.currentTarget.style.color=dark?"#aaa":"#666";}}>+</button>}
        {si>0&&<button onClick={()=>removeSlot(day.id,slot.slotId)} title="Rimuovi" style={{border:"1px solid #ccc",background:dark?"#333":"#FFFFFF",width:20,height:20,fontSize:13,fontWeight:700,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",color:"#999"}} onMouseEnter={e=>{e.currentTarget.style.background="#FF3B3B";e.currentTarget.style.color="#fff";}} onMouseLeave={e=>{e.currentTarget.style.background=dark?"#333":"#FFFFFF";e.currentTarget.style.color="#999";}}>×</button>}
      </div>
    </div>
  );
});


const Badge = ({children,bg,color:cl}) => <span style={{display:"inline-block",padding:"2px 8px",border:"1.5px solid #000",fontWeight:700,fontSize:10,letterSpacing:.3,textTransform:"uppercase",background:bg||"#F5F5DC",color:cl||"#000",lineHeight:1.4}}>{children}</span>;

// ═══ MAIN APP ═══
export default function App() {
  const [days, setDays] = useState(genDays);
  const [history, setHistory] = useState([]); // undo stack
  const [view, setView] = useState("calendario");
  const [listMonth, setListMonth] = useState(0); // selected month index for list view
  const [search, setSearch] = useState("");
  const [fStato, setFStato] = useState("");
  const [fSet, setFSet] = useState("");
  const [fProd, setFProd] = useState("");
  const [modal, setModal] = useState(false);
  const [consModal, setConsModal] = useState(null);
  const [bePopup, setBePopup] = useState(null);
  const [detailDay, setDetailDay] = useState(null);
  const [showCal, setShowCal] = useState(true);
  const [params, setParams] = useState(JSON.parse(JSON.stringify(INITIAL_PARAMS)));
  const [dark, setDark] = useState(false);
  const [globalSearch, setGlobalSearch] = useState("");
  const [qf, setQf] = useState({date:"",artist:"",agenzia:"",importo:"",costoTicket:"",memoDeal:"",chiusaDa:"",stato:"",tipoSet:"",produzione:"",dettagliCoProd:"",note:""});

  // ═══ UNDO ═══
  const daysRef = useRef(days);
  daysRef.current = days;
  const saveHistory = useCallback(() => {
    setHistory(prev => [...prev.slice(-20), JSON.stringify(daysRef.current)]);
  },[]);

  const undo = useCallback(() => {
    if (history.length === 0) return;
    const prev = JSON.parse(history[history.length - 1]);
    // Restore date objects
    const restored = prev.map(d => ({...d, date: new Date(d.date)}));
    setDays(restored);
    setHistory(h => h.slice(0, -1));
  },[history]);

  // ═══ DATA OPERATIONS ═══
  const update = useCallback((slotId,field,value) => {
    setDays(p=>p.map(d=>{ const idx=d.slots.findIndex(s=>s.slotId===slotId); if(idx===-1) return d; const ns=[...d.slots]; ns[idx]={...ns[idx],[field]:value}; return {...d,slots:ns}; }));
  },[]);

  const addSlot = useCallback((dayId) => {
    saveHistory();
    setDays(p=>p.map(d=>d.id===dayId?{...d, slots:[...d.slots, emptySlot(dayId, Date.now())]}:d));
  },[saveHistory]);

  const removeSlot = useCallback((dayId, slotId) => {
    saveHistory();
    setDays(p=>p.map(d=>d.id===dayId&&d.slots.length>1?{...d, slots:d.slots.filter(s=>s.slotId!==slotId)}:d));
  },[saveHistory]);

  const saveConsuntivo = (slotId, consuntivo) => {
    saveHistory();
    setDays(p=>p.map(d=>{ const idx=d.slots.findIndex(s=>s.slotId===slotId); if(idx===-1) return d; const ns=[...d.slots]; ns[idx]={...ns[idx],consuntivo}; return {...d,slots:ns}; }));
    setConsModal(null);
  };

  const addEvent = () => {
    if(!qf.date||!qf.artist) return;
    saveHistory();
    const {date:_,...rest}=qf;
    setDays(p=>p.map(d=>{
      if(d.id!==qf.date) return d;
      const firstEmpty = d.slots.findIndex(s=>!s.artist);
      if(firstEmpty>=0) return {...d, slots: d.slots.map((s,i)=>i===firstEmpty?{...s,...rest}:s)};
      return {...d, slots: [...d.slots, {...emptySlot(d.id, Date.now()), ...rest}]};
    }));
    setQf({date:"",artist:"",agenzia:"",importo:"",costoTicket:"",memoDeal:"",chiusaDa:"",stato:"",tipoSet:"",produzione:"",dettagliCoProd:"",note:""});
    setModal(false);
  };

  const allEvents = useMemo(() => days.flatMap(d => d.slots.filter(s=>s.artist).map(s => ({...s, date:d.date, giorno:d.giorno, dayId:d.id}))), [days]);

  // ═══ ALERTS ═══
  const alerts = useMemo(() => {
    const warns = [];
    const today = new Date();
    allEvents.forEach(ev => {
      if (ev.artist && !ev.stato) warns.push({type:"noStato",ev,msg:`${ev.artist} (${fmt(ev.date)}) — stato non compilato`});
      if (ev.stato==="In trattativa" && ev.date < today) warns.push({type:"scaduta",ev,msg:`${ev.artist} (${fmt(ev.date)}) — trattativa su data passata`});
      if (ev.artist && !ev.importo) warns.push({type:"noCachet",ev,msg:`${ev.artist} (${fmt(ev.date)}) — cachet non compilato`});
    });
    return warns.slice(0, 20);
  }, [allEvents]);

  // ═══ GLOBAL SEARCH ═══
  const searchResults = useMemo(() => {
    if (!globalSearch || globalSearch.length < 2) return [];
    const q = globalSearch.toLowerCase();
    return allEvents.filter(s => 
      (s.artist||"").toLowerCase().includes(q) ||
      (s.agenzia||"").toLowerCase().includes(q) ||
      (s.memoDeal||"").toLowerCase().includes(q) ||
      (s.note||"").toLowerCase().includes(q) ||
      (s.chiusaDa||"").toLowerCase().includes(q) ||
      (s.dettagliCoProd||"").toLowerCase().includes(q)
    ).slice(0, 15);
  }, [globalSearch, allEvents]);

  // ═══ COMPUTED ═══
  const stats = useMemo(() => {
    const ev=allEvents;
    const tot=ev.reduce((s,d)=>s+(parseFloat(d.importo)||0),0);
    const tInv=ev.filter(d=>isInv(d.date)).reduce((s,d)=>s+(parseFloat(d.importo)||0),0);
    const tEst=ev.filter(d=>isEst(d.date)).reduce((s,d)=>s+(parseFloat(d.importo)||0),0);
    const top=[...ev].filter(d=>d.importo).sort((a,b)=>(parseFloat(b.importo)||0)-(parseFloat(a.importo)||0));
    const mk = (fn) => { const items=ev.filter(fn); const t=items.reduce((s,d)=>s+(parseFloat(d.importo)||0),0); const r=[...items].filter(d=>d.importo).sort((a,b)=>(parseFloat(b.importo)||0)-(parseFloat(a.importo)||0)); return {n:items.length,tot:t,avg:items.length?t/items.length:0,top:r,conf:items.filter(d=>d.stato==="Confermato").length,tratt:items.filter(d=>d.stato==="In trattativa").length}; };
    return { tot,tInv,tEst,n:ev.length,conf:ev.filter(d=>d.stato==="Confermato").length,tratt:ev.filter(d=>d.stato==="In trattativa").length,
      live:ev.filter(d=>d.tipoSet==="LIVE").length,dj:ev.filter(d=>d.tipoSet==="DJ SET").length,format:ev.filter(d=>d.tipoSet==="FORMAT").length,
      nostra:ev.filter(d=>d.produzione==="Nostra").length,copro:ev.filter(d=>d.produzione==="Co-Prod").length,ext:ev.filter(d=>d.produzione==="Prod Esterna").length,
      avg:ev.length?tot/ev.length:0, top,
      prodNostra:mk(d=>d.produzione==="Nostra"), prodCoProd:mk(d=>d.produzione==="Co-Prod"), prodEsterna:mk(d=>d.produzione==="Prod Esterna"),
      statoConf:mk(d=>d.stato==="Confermato"), statoTratt:mk(d=>d.stato==="In trattativa"), statoAnn:mk(d=>d.stato==="Annullato"),
      setLive:mk(d=>d.tipoSet==="LIVE"), setDj:mk(d=>d.tipoSet==="DJ SET"), setFormat:mk(d=>d.tipoSet==="FORMAT"),
    };
  },[allEvents]);

  // Current month data for calendar/list views
  const currentMonthIdx = MESI_IDX[listMonth];
  const currentMonthYear = MESI_YEAR[listMonth];
  const monthDaysRaw = useMemo(() => {
    return days.filter(d => d.date.getMonth()===currentMonthIdx && d.date.getFullYear()===currentMonthYear);
  },[days,currentMonthIdx,currentMonthYear]);

  const currentMonthDays = useMemo(() => {
    let list = monthDaysRaw;
    if (view==="invernale") list = list.filter(d=>isInv(d.date)&&d.slots.some(s=>s.artist)).map(d=>({...d,slots:d.slots.filter(s=>s.artist)}));
    else if (view==="estivo") list = list.filter(d=>isEst(d.date)&&d.slots.some(s=>s.artist)).map(d=>({...d,slots:d.slots.filter(s=>s.artist)}));
    if(search) list=list.map(d=>({...d,slots:d.slots.filter(s=>(s.artist+s.agenzia).toLowerCase().includes(search.toLowerCase()))})).filter(d=>d.slots.length>0);
    if(fStato) list=list.map(d=>({...d,slots:d.slots.filter(s=>s.stato===fStato)})).filter(d=>d.slots.length>0);
    if(fSet) list=list.map(d=>({...d,slots:d.slots.filter(s=>s.tipoSet===fSet)})).filter(d=>d.slots.length>0);
    if(fProd) list=list.map(d=>({...d,slots:d.slots.filter(s=>s.produzione===fProd)})).filter(d=>d.slots.length>0);
    return list;
  },[monthDaysRaw,view,search,fStato,fSet,fProd]);

  // Calendar grid data
  const calGrid = useMemo(() => {
    const fd = new Date(currentMonthYear, currentMonthIdx, 1);
    let sw = fd.getDay(); sw = sw===0?6:sw-1;
    const dim = new Date(currentMonthYear, currentMonthIdx+1, 0).getDate();
    const byDate = {};
    monthDaysRaw.forEach(d => { byDate[d.date.getDate()] = d; });
    const cells = [];
    for (let i=0;i<sw;i++) cells.push(null);
    for (let d=1;d<=dim;d++) cells.push(byDate[d]||null);
    return cells;
  },[monthDaysRaw,currentMonthIdx,currentMonthYear]);

  const consDay = useMemo(()=>{ if(!consModal) return null; for(const d of days) for(const s of d.slots) if(s.slotId===consModal) return {...s,date:d.date,giorno:d.giorno,dayId:d.id}; return null; },[consModal,days]);

  // ═══ UI COMPONENTS ═══


  const MonthTabs = () => {
    const visibleMonths = MESI_SHORT.map((m,i) => ({m,i,mIdx:MESI_IDX[i],mYear:MESI_YEAR[i]})).filter(({mIdx}) => {
      if (view==="invernale") return mIdx>=9||mIdx<=3;
      if (view==="estivo") return mIdx>=4&&mIdx<=8;
      return true;
    });
    return (
    <div style={{display:"flex",flexWrap:"wrap",gap:3,padding:"12px 24px",background:dark?"#2a2a2a":C.beige,borderBottom:dark?"2px solid #444":"2px solid #000"}}>
      {visibleMonths.map(({m,i,mIdx,mYear}) => {
        const evCount = days.filter(d=>d.date.getMonth()===mIdx&&d.date.getFullYear()===mYear).reduce((s,d)=>s+d.slots.filter(sl=>sl.artist).length,0);
        return (
          <button key={`${m}${i}`} onClick={()=>setListMonth(i)} style={{
            border:"2px solid #000",padding:"6px 14px",
            background:listMonth===i?C.black:C.white,
            color:listMonth===i?C.white:C.black,
            fontWeight:700,fontSize:12,cursor:"pointer",position:"relative",letterSpacing:.5
          }}>
            {m} {mYear.toString().slice(2)}
            {evCount>0 && <span style={{position:"absolute",top:-6,right:-6,background:C.teal,color:C.black,borderRadius:"50%",width:18,height:18,fontSize:9,fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",border:"1.5px solid #000"}}>{evCount}</span>}
          </button>
        );
      })}
    </div>
    );
  };


  // Auto-select valid month when switching views
  const validMonthForView = useMemo(() => {
    const mIdx = MESI_IDX[listMonth];
    if (view==="invernale" && !(mIdx>=9||mIdx<=3)) {
      return 0; // Oct 2026
    }
    if (view==="estivo" && !(mIdx>=4&&mIdx<=8)) {
      return 7; // May 2027
    }
    return listMonth;
  },[view, listMonth]);
  useEffect(() => { if (validMonthForView !== listMonth) setListMonth(validMonthForView); },[validMonthForView, listMonth]);

  const isTableView = view==="calendario"||view==="invernale"||view==="estivo";

  // ═══ RENDER ═══
  return (
    <div style={{fontFamily:"'Inter','Helvetica Neue',sans-serif",background:dark?"#1a1a1a":C.paper,color:dark?"#eee":C.black,minHeight:"100vh",transition:"background .3s, color .3s"}}>
      <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Space+Mono:wght@400;700&display=swap" rel="stylesheet"/>
      <style>{`*{box-sizing:border-box;margin:0;padding:0}::-webkit-scrollbar{width:6px;height:6px}::-webkit-scrollbar-track{background:${dark?"#222":"#F5F5DC"}}::-webkit-scrollbar-thumb{background:${dark?"#666":"#000"};border-radius:3px}input:focus,select:focus{outline:2px solid #00C2CB !important;outline-offset:-1px}select{cursor:pointer}select option{font-weight:600}${dark?"input,select{background:#2a2a2a !important;color:#f0f0f0 !important;border-color:#666 !important}button{color:#f0f0f0}table,th,td{border-color:#555 !important}":""}`}</style>

      {/* ═══ HEADER ═══ */}
      <div style={{borderBottom:dark?"3px solid #444":"3px solid #000",padding:"14px 24px",background:dark?"#222":C.white}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:10}}>
          <div style={{display:"flex",alignItems:"center",gap:14}}>
            <img src="data:image/png;base64,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" alt="Magnolia" style={{height:44}}/>
            <div style={{fontSize:11,fontWeight:600,color:"#888",letterSpacing:.5}}>OTT 2026 — OTT 2027</div>
          </div>

          <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
            <div data-global-search="true" style={{position:"relative"}}>
              <input value={globalSearch} onChange={e=>setGlobalSearch(e.target.value)} placeholder="Cerca ovunque..." style={{width:180,border:brutBorder,padding:"8px 12px",fontWeight:600,fontSize:11,background:dark?"#333":"#fff",color:dark?"#eee":"#000"}}/>
              {searchResults.length>0&&<div style={{position:"absolute",top:"100%",left:0,width:320,zIndex:200,background:dark?"#333":"#fff",border:brutBorder,boxShadow:brutShadow,maxHeight:300,overflowY:"auto"}}>
                {searchResults.map(ev=>(
                  <div key={ev.slotId} onClick={()=>{setDetailDay(ev.dayId);setGlobalSearch("");}} style={{padding:"8px 12px",borderBottom:"1px solid "+(dark?"#555":"#eee"),cursor:"pointer",display:"flex",justifyContent:"space-between",alignItems:"center"}} onMouseEnter={e=>{e.currentTarget.style.background=dark?"#444":"#f5f5f0";}} onMouseLeave={e=>{e.currentTarget.style.background="transparent";}}>
                    <div><div style={{fontWeight:700,fontSize:12,textTransform:"uppercase"}}>{ev.artist}</div><div style={{fontSize:10,color:"#888"}}>{ev.agenzia}</div></div>
                    <div style={{textAlign:"right"}}><div style={{fontFamily:"'Space Mono'",fontSize:11,fontWeight:700}}>{fmt(ev.date)}</div>{ev.stato&&<div style={{fontSize:9,fontWeight:700,color:ev.stato==="Confermato"?"#28a745":ev.stato==="In trattativa"?"#ffc107":"#dc3545"}}>{ev.stato}</div>}</div>
                  </div>
                ))}
              </div>}
            </div>
            {alerts.length>0&&<div style={{position:"relative",display:"inline-block"}}><button onClick={()=>setDetailDay("alerts")} style={{border:brutBorder,background:dark?"#333":"#fff",color:C.danger,padding:"8px 14px",fontWeight:700,fontSize:12,cursor:"pointer",boxShadow:brutShadow}}>⚠ {alerts.length}</button></div>}
            <button onClick={()=>setDark(!dark)} title={dark?"Modalità chiara":"Modalità scura"} style={{border:brutBorder,background:dark?"#333":"#fff",color:dark?"#FFE700":"#000",padding:"8px 14px",fontWeight:700,fontSize:12,cursor:"pointer",boxShadow:brutShadow}}>{dark?"DAY":"NIGHT"}</button>
            {history.length>0 && <button onClick={undo} title="Annulla ultima azione" style={{border:brutBorder,background:dark?"#333":C.white,color:dark?"#eee":C.black,padding:"8px 14px",fontWeight:700,fontSize:12,cursor:"pointer",boxShadow:brutShadow,display:"flex",alignItems:"center",gap:4}}>↩ UNDO</button>}
            <button onClick={()=>setModal(true)} style={{border:brutBorder,background:C.magenta,color:C.white,padding:"8px 18px",fontWeight:700,fontSize:12,cursor:"pointer",boxShadow:brutShadow}}>AGGIUNGI DATA</button>
            <button onClick={()=>exportData(days,stats)} style={{border:brutBorder,background:C.teal,padding:"8px 18px",fontWeight:700,fontSize:12,cursor:"pointer",boxShadow:brutShadow}}>↓ ESPORTA</button>
          </div>
        </div>
        <div style={{display:"flex",gap:0,marginTop:14,borderTop:"2px solid #000"}}>
          {[{id:"calendario",l:"CALENDARIO"},{id:"invernale",l:"INVERNALE"},{id:"estivo",l:"ESTIVO"},{id:"recap",l:"RECAP"},{id:"parametri",l:"PARAMETRI"}].map((v,i)=>
            <button key={v.id} onClick={()=>setView(v.id)} style={{padding:"8px 18px",border:"2px solid #000",borderTop:"none",marginLeft:i===0?0:-2,background:view===v.id?C.yellow:dark?"#333":"#fff",color:view===v.id?"#000":dark?"#eee":"#000",fontWeight:700,fontSize:12,cursor:"pointer",letterSpacing:.5,textTransform:"uppercase",position:"relative",zIndex:view===v.id?2:1}}>{v.l}</button>
          )}
        </div>
      </div>

      {/* ═══ MONTH TABS + FILTERS ═══ */}
      {isTableView && <>
        <MonthTabs/>
        <div style={{padding:"10px 24px",display:"flex",gap:6,flexWrap:"wrap",alignItems:"center",borderBottom:dark?"2px solid #444":"2px solid #000",background:dark?"#2a2a2a":C.white}}>
          <input placeholder="Cerca artist / agenzia..." value={search} onChange={e=>setSearch(e.target.value)} style={{border:"2px solid #000",padding:"6px 10px",fontWeight:600,fontSize:12,width:200,background:C.white}}/>
          <select value={fStato} onChange={e=>setFStato(e.target.value)} style={{border:"2px solid #000",padding:"6px 10px",fontWeight:700,fontSize:11,background:C.white}}><option value="">TUTTI STATI</option><option>Confermato</option><option>In trattativa</option><option>Annullato</option></select>
          <select value={fSet} onChange={e=>setFSet(e.target.value)} style={{border:"2px solid #000",padding:"6px 10px",fontWeight:700,fontSize:11,background:C.white}}><option value="">TUTTI SET</option><option>LIVE</option><option>DJ SET</option><option>FORMAT</option></select>
          <select value={fProd} onChange={e=>setFProd(e.target.value)} style={{border:"2px solid #000",padding:"6px 10px",fontWeight:700,fontSize:11,background:C.white}}><option value="">TUTTE PROD.</option><option>Nostra</option><option>Co-Prod</option><option>Prod Esterna</option></select>
          {(search||fStato||fSet||fProd) && <button onClick={()=>{setSearch("");setFStato("");setFSet("");setFProd("");}} style={{border:"2px solid #000",background:C.danger,color:C.white,padding:"6px 12px",fontWeight:700,fontSize:11,cursor:"pointer"}}>✕</button>}
          <div style={{marginLeft:"auto"}}><button onClick={()=>setShowCal(!showCal)} style={{border:"2px solid #000",background:showCal?C.yellow:C.white,padding:"6px 12px",fontWeight:700,fontSize:11,cursor:"pointer"}}>{showCal?"▲ NASCONDI GRIGLIA":"▼ MOSTRA GRIGLIA"}</button></div>
        </div>
      </>}

      {/* ═══ CONTENT ═══ */}
      <div style={{padding:"14px 24px 40px"}} onClick={e=>{if(bePopup && !e.target.closest('[data-be-popup]'))setBePopup(null);if(globalSearch && !e.target.closest('[data-global-search]'))setGlobalSearch("");}}>

        {/* ═══ CALENDAR GRID VIEW ═══ */}
        {view==="calendario" && showCal && !search && !fStato && !fSet && !fProd && (
          <div style={{marginBottom:20}}>
            <div style={{border:brutBorder,background:C.white,overflow:"hidden"}}>
              <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",borderBottom:"2px solid #000"}}>
                {GIORNI_SHORT.map(g=><div key={g} style={{padding:"8px",textAlign:"center",fontWeight:700,fontSize:11,color:dark?"#888":"#666",borderRight:dark?"1px solid #444":"1px solid #eee"}}>{g}</div>)}
              </div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)"}}>
                {calGrid.map((dayObj,i)=>{
                  if(!dayObj) return <div key={`e${i}`} style={{minHeight:70,background:dark?"#151515":"#f8f8f4",borderRight:dark?"1px solid #444":"1px solid #eee",borderBottom:dark?"1px solid #444":"1px solid #eee"}}/>;
                  const evts = dayObj.slots.filter(s=>s.artist);
                  const we = isWe(dayObj.date);
                  return (
                    <div key={dayObj.id} style={{minHeight:70,padding:"4px 6px",borderRight:dark?"1px solid #444":"1px solid #eee",borderBottom:dark?"1px solid #444":"1px solid #eee",background:we?(dark?"#2d2040":"#f5f0ff"):(dark?"#1e1e1e":"#fff"),cursor:"pointer",transition:"background .1s"}}
                      onClick={()=>setDetailDay(dayObj.id)}
                      onMouseEnter={e=>{e.currentTarget.style.background=we?(dark?"#3d2060":"#ebe3ff"):(dark?"#2a2a2a":"#f5f5f0");}}
                      onMouseLeave={e=>{e.currentTarget.style.background=we?(dark?"#2d2040":"#f5f0ff"):(dark?"#1e1e1e":"#fff");}}>
                      <div style={{fontSize:12,fontWeight:700,color:we?C.magenta:"#333",marginBottom:3}}>{dayObj.date.getDate()}</div>
                      {evts.slice(0,2).map(s=>(
                        <div key={s.slotId} style={{fontSize:9,fontWeight:600,padding:"1px 4px",marginBottom:2,borderLeft:`3px solid ${s.stato==="Confermato"?"#28a745":s.stato==="In trattativa"?"#ffc107":s.stato==="Annullato"?"#dc3545":"#ccc"}`,background:"#f9f9f6",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>
                          {s.artist}
                        </div>
                      ))}
                      {evts.length>2 && <div style={{fontSize:8,fontWeight:700,color:"#999"}}>+{evts.length-2} altro</div>}
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        )}

        {/* ═══ LIST TABLE ═══ */}
        {isTableView && (
          <div style={{overflowX:"auto"}}>
            <div style={{minWidth:1400,border:brutBorder,boxShadow:"4px 4px 0 #000",background:dark?"#1e1e1e":C.white}}>
              <div style={{display:"flex",borderBottom:"3px solid #000",background:C.black}}>
                <div style={{width:55,padding:"8px 5px",fontSize:9,fontWeight:700,color:C.white,letterSpacing:1}}>GIORNO</div>
                <div style={{width:72,padding:"8px 5px",fontSize:9,fontWeight:700,color:C.teal,letterSpacing:1}}>DATA</div>
                {FIELDS.map(f=><div key={f.key} style={{width:f.w,padding:"8px 6px",fontSize:9,fontWeight:700,color:C.yellow,letterSpacing:.8}}>{f.label}</div>)}
                <div style={{width:22}}/>
              </div>
              <div style={{background:C.yellow,padding:"8px 10px",borderBottom:"2px solid #000",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <span style={{fontWeight:700,fontSize:14,textTransform:"uppercase",letterSpacing:1}}>{MESI[currentMonthIdx]} {currentMonthYear}</span>
                <Badge bg={C.black} color={C.yellow}>{currentMonthDays.reduce((s,d)=>s+d.slots.filter(sl=>sl.artist).length,0)} EVENTI</Badge>
              </div>
              {currentMonthDays.map(day=>(
                <div key={day.id}>
                  {day.slots.map((slot,si)=><SlotRow key={slot.slotId} day={day} slot={slot} si={si} bePopup={bePopup} setBePopup={setBePopup} update={update} addSlot={addSlot} removeSlot={removeSlot} setDetailDay={setDetailDay} setConsModal={setConsModal} dark={dark}/>)}
                </div>
              ))}
              {currentMonthDays.length===0&&<div style={{padding:30,textAlign:"center",fontWeight:700,color:dark?"#666":"#999"}}>Nessun risultato</div>}
            </div>
          </div>
        )}

        {/* ═══ RECAP ═══ */}
        {view==="recap" && (
          <div>
            <div style={{border:brutBorder,background:dark?"#222":C.white,marginBottom:24}}>
              <div style={{background:C.black,padding:"14px 20px"}}>
                <div style={{fontSize:9,fontWeight:700,letterSpacing:2,color:"#888"}}>RIEPILOGO STAGIONE</div>
                <div style={{fontFamily:"'Space Mono'",fontSize:28,fontWeight:700,color:C.white,marginTop:2}}>{eur(stats.tot)}</div>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(6,1fr)",borderTop:brutBorder}}>
                {[{l:"EVENTI",v:stats.n},{l:"CONFERMATI",v:stats.conf,c:"#007a6a"},{l:"TRATTATIVA",v:stats.tratt,c:"#b58900"},{l:"INVERNALE",v:eur(stats.tInv)},{l:"ESTIVO",v:eur(stats.tEst)},{l:"MEDIA",v:stats.n?eur(stats.avg):"—"}].map((s,i)=>
                  <div key={i} style={{padding:"10px 12px",borderRight:i<5?"1px solid #ddd":"none",textAlign:"center"}}>
                    <div style={{fontSize:8,fontWeight:700,letterSpacing:1.5,color:"#999",marginBottom:3}}>{s.l}</div>
                    <div style={{fontFamily:"'Space Mono'",fontSize:13,fontWeight:700,color:s.c||C.black}}>{s.v}</div>
                  </div>
                )}
              </div>
            </div>

            {/* Per Produzione */}
            <div style={{fontSize:12,fontWeight:700,letterSpacing:1.5,textTransform:"uppercase",marginBottom:8,paddingBottom:5,borderBottom:"2px solid #000"}}>Per Produzione</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:24}}>
              {[{label:"NOSTRA",data:stats.prodNostra},{label:"CO-PROD",data:stats.prodCoProd},{label:"ESTERNA",data:stats.prodEsterna}].map(b=>(
                <div key={b.label} style={{border:brutBorder,background:C.white,overflow:"hidden"}}>
                  <div style={{background:C.black,padding:"8px 12px",display:"flex",justifyContent:"space-between"}}><span style={{fontWeight:700,fontSize:12,color:C.white}}>{b.label}</span><span style={{fontFamily:"'Space Mono'",fontSize:12,fontWeight:700,color:C.white}}>{b.data.n} ev.</span></div>
                  <div style={{padding:"10px 12px",borderBottom:"1px solid #eee",display:"flex",justifyContent:"space-between"}}><div><div style={{fontSize:8,fontWeight:700,color:"#999"}}>TOTALE</div><div style={{fontFamily:"'Space Mono'",fontSize:15,fontWeight:700}}>{eur(b.data.tot)}</div></div><div style={{textAlign:"right"}}><div style={{fontSize:8,fontWeight:700,color:"#999"}}>MEDIA</div><div style={{fontFamily:"'Space Mono'",fontSize:13,fontWeight:700,color:"#666"}}>{b.data.n?eur(b.data.avg):"—"}</div></div></div>
                  {b.data.top.slice(0,5).map((ev,i)=><div key={ev.slotId} style={{display:"flex",padding:"4px 12px",borderBottom:"1px solid #f0f0f0",fontSize:11}}><span style={{width:18,fontFamily:"'Space Mono'",fontWeight:700,color:"#ccc"}}>{i+1}</span><span style={{flex:2,fontWeight:600,textTransform:"uppercase"}}>{ev.artist}</span><span style={{fontFamily:"'Space Mono'",fontWeight:700,color:"#007a6a"}}>{eur(parseFloat(ev.importo))}</span></div>)}
                  {b.data.top.length===0&&<div style={{padding:14,textAlign:"center",color:"#ccc",fontSize:11}}>—</div>}
                </div>
              ))}
            </div>

            {/* Per Stato */}
            <div style={{fontSize:12,fontWeight:700,letterSpacing:1.5,textTransform:"uppercase",marginBottom:8,paddingBottom:5,borderBottom:"2px solid #000"}}>Per Stato</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:24}}>
              {[{label:"CONFERMATI",data:stats.statoConf,dot:"#007a6a"},{label:"IN TRATTATIVA",data:stats.statoTratt,dot:"#b58900"},{label:"ANNULLATI",data:stats.statoAnn,dot:C.danger}].map(b=>(
                <div key={b.label} style={{border:brutBorder,background:C.white,overflow:"hidden"}}>
                  <div style={{padding:"8px 12px",borderBottom:"1px solid #eee",display:"flex",justifyContent:"space-between",alignItems:"center"}}><div style={{display:"flex",alignItems:"center",gap:6}}><div style={{width:8,height:8,borderRadius:"50%",background:b.dot}}/><span style={{fontWeight:700,fontSize:12}}>{b.label}</span></div><span style={{fontFamily:"'Space Mono'",fontSize:11,fontWeight:700,color:"#666"}}>{b.data.n}</span></div>
                  <div style={{padding:"6px 12px",borderBottom:"1px solid #eee"}}><div style={{fontFamily:"'Space Mono'",fontSize:15,fontWeight:700}}>{eur(b.data.tot)}</div></div>
                  {b.data.top.slice(0,5).map((ev,i)=><div key={ev.slotId} style={{display:"flex",padding:"4px 12px",borderBottom:"1px solid #f0f0f0",fontSize:11}}><span style={{width:18,fontFamily:"'Space Mono'",fontWeight:700,color:"#ccc"}}>{i+1}</span><span style={{flex:2,fontWeight:600,textTransform:"uppercase"}}>{ev.artist}</span><span style={{fontFamily:"'Space Mono'",fontWeight:700,fontSize:10}}>{fmt(ev.date)}</span></div>)}
                  {b.data.top.length===0&&<div style={{padding:14,textAlign:"center",color:"#ccc",fontSize:11}}>—</div>}
                </div>
              ))}
            </div>

            {/* Per Set */}
            <div style={{fontSize:12,fontWeight:700,letterSpacing:1.5,textTransform:"uppercase",marginBottom:8,paddingBottom:5,borderBottom:"2px solid #000"}}>Per Tipo Set</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:24}}>
              {[{label:"LIVE",data:stats.setLive,dot:C.magenta},{label:"DJ SET",data:stats.setDj,dot:C.teal},{label:"FORMAT",data:stats.setFormat,dot:C.yellow}].map(b=>(
                <div key={b.label} style={{border:brutBorder,background:C.white,overflow:"hidden"}}>
                  <div style={{padding:"8px 12px",borderBottom:"1px solid #eee",display:"flex",justifyContent:"space-between",alignItems:"center"}}><div style={{display:"flex",alignItems:"center",gap:6}}><div style={{width:8,height:8,borderRadius:"50%",background:b.dot}}/><span style={{fontWeight:700,fontSize:12}}>{b.label}</span></div><span style={{fontFamily:"'Space Mono'",fontSize:11,fontWeight:700,color:"#666"}}>{b.data.n}</span></div>
                  <div style={{padding:"6px 12px",borderBottom:"1px solid #eee"}}><div style={{fontFamily:"'Space Mono'",fontSize:15,fontWeight:700}}>{eur(b.data.tot)}</div></div>
                  {b.data.top.slice(0,5).map((ev,i)=><div key={ev.slotId} style={{display:"flex",padding:"4px 12px",borderBottom:"1px solid #f0f0f0",fontSize:11}}><span style={{width:18,fontFamily:"'Space Mono'",fontWeight:700,color:"#ccc"}}>{i+1}</span><span style={{flex:2,fontWeight:600,textTransform:"uppercase"}}>{ev.artist}</span><span style={{fontFamily:"'Space Mono'",fontWeight:700,fontSize:10}}>{fmt(ev.date)}</span></div>)}
                  {b.data.top.length===0&&<div style={{padding:14,textAlign:"center",color:"#ccc",fontSize:11}}>—</div>}
                </div>
              ))}
            </div>
          </div>
        )}

        {/* ═══ PARAMETRI ═══ */}
        {view==="parametri" && (
          <div>
            <div style={{background:C.black,border:brutBorder,boxShadow:brutShadow,padding:"12px 18px",marginBottom:18}}>
              <div style={{color:C.yellow,fontSize:9,fontWeight:700,letterSpacing:2}}>IMPOSTAZIONI</div>
              <div style={{color:C.white,fontSize:16,fontWeight:700,marginTop:2}}>Parametri Default Consuntivo</div>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
              {[{key:"costiProduzione",label:"COSTI PRODUZIONE",bg:C.yellow+"33"},{key:"costiArtistici",label:"COSTI ARTISTICI",bg:C.magenta+"22"},{key:"costiPersonale",label:"COSTI PERSONALE",bg:C.teal+"22"},{key:"costiAllestimento",label:"COSTI ALLESTIMENTO",bg:C.yellow+"22"}].map(sec=>(
                <div key={sec.key} style={{border:brutBorder,background:C.white,overflow:"hidden"}}>
                  <div style={{padding:"8px 12px",background:sec.bg,borderBottom:"2px solid #000",fontWeight:700,fontSize:12,letterSpacing:.5}}>{sec.label}</div>
                  <div style={{padding:"6px 10px"}}>
                    {params[sec.key].map((item,i)=>(
                      <div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",borderBottom:"1px solid #eee",padding:"4px 0",gap:6}}>
                        <input value={item.voce} onChange={e=>{const n=[...params[sec.key]];n[i]={...n[i],voce:e.target.value};setParams({...params,[sec.key]:n});}} style={{flex:2,border:"1px solid #ddd",padding:"3px 5px",fontSize:11,fontWeight:600}}/>
                        <div style={{display:"flex",alignItems:"center",gap:3}}><span style={{fontSize:9,color:"#999"}}>€</span>
                          <input value={item.costo||""} onChange={e=>{const n=[...params[sec.key]];n[i]={...n[i],costo:parseFloat(e.target.value)||0};setParams({...params,[sec.key]:n});}} style={{width:70,border:"2px solid #000",padding:"3px 5px",fontFamily:"'Space Mono'",fontSize:11,fontWeight:700,textAlign:"right"}}/>
                        </div>
                      </div>
                    ))}
                    <button onClick={()=>{setParams({...params,[sec.key]:[...params[sec.key],{voce:"Nuova voce",costo:0}]});}} style={{border:"2px dashed #ccc",width:"100%",padding:"4px",marginTop:6,fontWeight:700,fontSize:10,cursor:"pointer",background:"transparent",color:"#888"}}>+ AGGIUNGI VOCE</button>
                  </div>
                </div>
              ))}
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginTop:16}}>
              <div style={{border:brutBorder,background:C.white,overflow:"hidden"}}>
                <div style={{padding:"8px 12px",background:C.danger+"22",borderBottom:"2px solid #000",fontWeight:700,fontSize:12}}>PERCENTUALI SU RICAVI</div>
                <div style={{padding:"10px"}}>
                  {params.costiRicavi.map((item,i)=>(
                    <div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",borderBottom:"1px solid #eee",padding:"6px 0"}}>
                      <input value={item.voce} onChange={e=>{const n=[...params.costiRicavi];n[i]={...n[i],voce:e.target.value};setParams({...params,costiRicavi:n});}} style={{flex:2,border:"1px solid #ddd",padding:"3px 5px",fontSize:11,fontWeight:600}}/>
                      <div style={{display:"flex",alignItems:"center",gap:3}}>
                        <input value={item.perc||""} onChange={e=>{const n=[...params.costiRicavi];n[i]={...n[i],perc:parseFloat(e.target.value)||0};setParams({...params,costiRicavi:n});}} style={{width:50,border:"2px solid #000",padding:"3px 5px",fontFamily:"'Space Mono'",fontSize:12,fontWeight:700,textAlign:"center"}}/>
                        <span style={{fontWeight:700,fontSize:12}}>%</span>
                      </div>
                    </div>
                  ))}
                  <div style={{borderTop:"1px solid #eee",padding:"6px 0",display:"flex",justifyContent:"space-between",alignItems:"center",marginTop:6}}>
                    <span style={{fontSize:11,fontWeight:700}}>COMMISSIONI TICKET</span>
                    <div style={{display:"flex",alignItems:"center",gap:3}}>
                      <input value={params.commissioniTicket||""} onChange={e=>setParams({...params,commissioniTicket:parseFloat(e.target.value)||0})} style={{width:50,border:"2px solid #000",padding:"3px 5px",fontFamily:"'Space Mono'",fontSize:12,fontWeight:700,textAlign:"center"}}/>
                      <span style={{fontWeight:700,fontSize:12}}>%</span>
                    </div>
                  </div>
                </div>
              </div>
              <div style={{border:brutBorder,background:C.white,overflow:"hidden"}}>
                <div style={{padding:"8px 12px",background:C.teal+"22",borderBottom:"2px solid #000",fontWeight:700,fontSize:12}}>RIPARTIZIONE DEFAULT</div>
                <div style={{padding:"16px",display:"flex",gap:14,justifyContent:"center",alignItems:"center"}}>
                  <div style={{textAlign:"center"}}><div style={{fontSize:10,fontWeight:700,marginBottom:6}}>MAGNOLIA</div>
                    <input value={params.ripartizione.magnolia||""} onChange={e=>{const v=parseFloat(e.target.value)||0;setParams({...params,ripartizione:{magnolia:v,partner:100-v}});}} style={{width:60,border:brutBorder,padding:"8px",fontFamily:"'Space Mono'",fontSize:18,fontWeight:700,textAlign:"center"}}/>
                    <span style={{fontWeight:700,fontSize:16}}> %</span></div>
                  <div style={{width:2,height:40,background:"#000"}}/>
                  <div style={{textAlign:"center"}}><div style={{fontSize:10,fontWeight:700,marginBottom:6}}>PARTNER</div>
                    <input value={params.ripartizione.partner||""} onChange={e=>{const v=parseFloat(e.target.value)||0;setParams({...params,ripartizione:{partner:v,magnolia:100-v}});}} style={{width:60,border:brutBorder,padding:"8px",fontFamily:"'Space Mono'",fontSize:18,fontWeight:700,textAlign:"center"}}/>
                    <span style={{fontWeight:700,fontSize:16}}> %</span></div>
                </div>
              </div>
            </div>
          </div>
        )}
      </div>

      {/* ═══ QUICK ADD MODAL ═══ */}
      {modal&&(
        <div style={{position:"fixed",inset:0,background:"#000000aa",display:"flex",alignItems:"center",justifyContent:"center",zIndex:100}} onClick={e=>e.target===e.currentTarget&&setModal(false)}>
          <div style={{background:C.white,border:"3px solid #000",boxShadow:"6px 6px 0 #000",padding:"24px",width:500,maxWidth:"95vw",maxHeight:"90vh",overflowY:"auto"}}>
            <h2 style={{fontSize:20,fontWeight:700,textTransform:"uppercase",borderBottom:"3px solid #000",paddingBottom:6,marginBottom:4,color:dark?"#eee":"#000"}}>Aggiungi Data</h2>
            <p style={{fontSize:11,fontWeight:600,color:"#666",marginBottom:10}}>Seleziona una data e compila i campi</p>
            <CalPicker days={days} selected={qf.date} onSelect={d=>setQf({...qf,date:d})}/>
            <label style={{display:"block",fontSize:9,fontWeight:700,letterSpacing:1.5,marginTop:14,marginBottom:3}}>ARTIST *</label>
            <input value={qf.artist} onChange={e=>setQf({...qf,artist:e.target.value})} placeholder="Nome artista..." style={{width:"100%",border:brutBorder,padding:"8px",fontWeight:600,fontSize:13,background:C.white}}/>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"0 12px"}}>
              {[{k:"agenzia",l:"AGENZIA",t:"text"},{k:"importo",l:"CACHET €",t:"number"},{k:"stato",l:"STATO",t:"select",o:["","Confermato","In trattativa","Annullato"]},{k:"tipoSet",l:"TIPO SET",t:"select",o:["","LIVE","DJ SET","FORMAT"]},{k:"produzione",l:"PRODUZIONE",t:"select",o:["","Nostra","Co-Prod","Prod Esterna"]},{k:"chiusaDa",l:"CHIUSA DA",t:"text"}].map(f=>(
                <div key={f.k}><label style={{display:"block",fontSize:9,fontWeight:700,letterSpacing:1.5,marginTop:12,marginBottom:3}}>{f.l}</label>
                  {f.t==="select"?<select value={qf[f.k]} onChange={e=>setQf({...qf,[f.k]:e.target.value})} style={{width:"100%",border:brutBorder,padding:"8px",fontWeight:700,fontSize:11,background:C.white}}>{f.o.map(o=><option key={o} value={o}>{o||"—"}</option>)}</select>
                  :<input type={f.t} value={qf[f.k]} onChange={e=>setQf({...qf,[f.k]:e.target.value})} style={{width:"100%",border:brutBorder,padding:"8px",fontWeight:600,fontSize:12,background:C.white}}/>}
                </div>
              ))}
            </div>
            {[{k:"dettagliCoProd",l:"DETTAGLI CO-PROD"},{k:"memoDeal",l:"MEMO DEAL"},{k:"note",l:"NOTE"}].map(f=>(
              <div key={f.k}><label style={{display:"block",fontSize:9,fontWeight:700,letterSpacing:1.5,marginTop:12,marginBottom:3}}>{f.l}</label>
                <input value={qf[f.k]} onChange={e=>setQf({...qf,[f.k]:e.target.value})} style={{width:"100%",border:brutBorder,padding:"8px",fontWeight:600,fontSize:12,background:C.white}}/></div>
            ))}
            <div style={{display:"flex",gap:8,marginTop:18,justifyContent:"flex-end"}}>
              <button onClick={()=>setModal(false)} style={{border:brutBorder,background:C.white,padding:"8px 20px",fontWeight:700,fontSize:13,cursor:"pointer"}}>ANNULLA</button>
              <button onClick={addEvent} disabled={!qf.date||!qf.artist} style={{border:brutBorder,background:C.magenta,color:C.white,padding:"8px 20px",fontWeight:700,fontSize:13,cursor:(!qf.date||!qf.artist)?"not-allowed":"pointer",opacity:(!qf.date||!qf.artist)?.4:1}}>AGGIUNGI</button>
            </div>
          </div>
        </div>
      )}

      {/* ═══ CONSUNTIVO MODAL ═══ */}
      {consDay && <ConsuntivoModal day={consDay} consuntivo={consDay.consuntivo} params={params} onSave={c=>saveConsuntivo(consDay.slotId,c)} onClose={()=>setConsModal(null)}/>}

      {/* ═══ ALERTS MODAL ═══ */}
      {detailDay==="alerts"&&(
        <div style={{position:"fixed",inset:0,background:"#000000aa",display:"flex",alignItems:"center",justifyContent:"center",zIndex:100}} onClick={e=>e.target===e.currentTarget&&setDetailDay(null)}>
          <div style={{background:dark?"#2a2a2a":"#fff",border:"3px solid #000",boxShadow:"6px 6px 0 #000",width:520,maxWidth:"95vw",maxHeight:"90vh",overflowY:"auto"}}>
            <div style={{background:C.danger,padding:"14px 20px"}}><div style={{fontWeight:700,fontSize:16,color:"#fff"}}>⚠ ATTENZIONE ({alerts.length})</div></div>
            {alerts.map((a,i)=>(
              <div key={i} onClick={()=>setDetailDay(a.ev.dayId)} style={{padding:"10px 20px",borderBottom:"1px solid "+(dark?"#444":"#eee"),cursor:"pointer",display:"flex",gap:10,alignItems:"center"}}
                onMouseEnter={e=>{e.currentTarget.style.background=dark?"#333":"#f9f9f6";}} onMouseLeave={e=>{e.currentTarget.style.background="transparent";}}>
                <span style={{fontSize:16}}>{a.type==="scaduta"?"🔴":a.type==="noStato"?"🟡":"🟠"}</span>
                <span style={{fontSize:12,fontWeight:600}}>{a.msg}</span>
              </div>
            ))}
            <div style={{padding:"12px 20px",textAlign:"right"}}><button onClick={()=>setDetailDay(null)} style={{border:brutBorder,background:dark?"#333":"#fff",padding:"6px 16px",fontWeight:700,fontSize:11,cursor:"pointer"}}>CHIUDI</button></div>
          </div>
        </div>
      )}

      {/* ═══ DETAIL DAY MODAL ═══ */}
      {detailDay && detailDay!=="alerts" && (()=>{
        const dayObj = days.find(d=>d.id===detailDay);
        if(!dayObj) return null;
        const filledSlots = dayObj.slots.filter(s=>s.artist);
        const DetailRow = ({label,value,mono}) => value ? <div style={{display:"flex",borderBottom:"1px solid #eee",padding:"6px 0"}}><div style={{width:130,fontSize:9,fontWeight:700,letterSpacing:1,color:"#888",textTransform:"uppercase",flexShrink:0}}>{label}</div><div style={{flex:1,fontSize:12,fontWeight:mono?700:500,fontFamily:mono?"'Space Mono',monospace":"inherit",wordBreak:"break-word"}}>{value}</div></div> : null;
        const SlotDetail = ({slot,idx}) => {
          const imp=parseFloat(slot.importo)||0;const tkt=parseFloat(slot.costoTicket)||0;
          const totPerc=((slot.beIva||0)+(slot.beSiae||0)+(slot.beComm||0))/100;const net=tkt*(1-totPerc);
          const beP=(imp&&tkt&&net>0)?Math.ceil(imp/net):"";const beC=slot.consuntivo?calcBE(slot.consuntivo):"";
          return (<div style={{padding:"14px 20px",borderBottom:"2px solid #000"}}>
            {filledSlots.length>1&&<div style={{fontSize:9,fontWeight:700,letterSpacing:2,color:"#999",marginBottom:6}}>SERATA {idx+1}</div>}
            <DetailRow label="Artist" value={slot.artist}/><DetailRow label="Agenzia" value={slot.agenzia}/>
            <DetailRow label="Cachet" value={slot.importo?eur(parseFloat(slot.importo)):null} mono/>
            <DetailRow label="Costo Ticket" value={slot.costoTicket?`€ ${slot.costoTicket}`:null} mono/>
            {beP&&<DetailRow label="BE Provv." value={`${beP} pax`} mono/>}
            {beC&&<DetailRow label="BE Consuntivo" value={`${beC} pax`} mono/>}
            <DetailRow label="Stato" value={slot.stato}/><DetailRow label="Tipo Set" value={slot.tipoSet}/>
            <DetailRow label="Produzione" value={slot.produzione}/><DetailRow label="Co-Prod" value={slot.dettagliCoProd}/>
            <DetailRow label="Chiusa da" value={slot.chiusaDa}/><DetailRow label="Memo Deal" value={slot.memoDeal}/>
            <DetailRow label="Note" value={slot.note}/>
            {slot.artist&&<button onClick={()=>{setDetailDay(null);setConsModal(slot.slotId);}} style={{border:"2px solid #000",background:C.yellow,padding:"5px 12px",fontWeight:700,fontSize:10,cursor:"pointer",marginTop:8}}>APRI CONSUNTIVO</button>}
          </div>);
        };
        return (<div style={{position:"fixed",inset:0,background:"#000000aa",display:"flex",alignItems:"center",justifyContent:"center",zIndex:100}} onClick={e=>e.target===e.currentTarget&&setDetailDay(null)}>
          <div style={{background:dark?"#2a2a2a":C.white,border:"3px solid #000",boxShadow:"6px 6px 0 #000",width:520,maxWidth:"95vw",maxHeight:"90vh",overflowY:"auto"}}>
            <div style={{background:C.black,padding:"16px 20px"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                <div><div style={{fontSize:9,fontWeight:700,letterSpacing:2,color:"#888"}}>{dayObj.giorno.toUpperCase()}</div><div style={{fontFamily:"'Space Mono'",fontSize:24,fontWeight:700,color:C.white,marginTop:2}}>{fmt(dayObj.date)}</div></div>
                <div style={{textAlign:"right"}}>
                  {filledSlots.length>1&&<div style={{display:"inline-block",padding:"3px 10px",border:"2px solid #555",fontWeight:700,fontSize:11,color:C.yellow}}>{filledSlots.length} SERATE</div>}
                  {filledSlots.length===1&&filledSlots[0].artist&&<div style={{fontSize:18,fontWeight:700,color:C.white,textTransform:"uppercase"}}>{filledSlots[0].artist}</div>}
                </div>
              </div>
            </div>
            {filledSlots.length===0&&<div style={{padding:24,textAlign:"center"}}><div style={{color:"#999",fontWeight:600,marginBottom:10}}>Nessun evento in questa data</div><button onClick={()=>{setDetailDay(null);setQf({date:dayObj.id,artist:"",agenzia:"",importo:"",costoTicket:"",memoDeal:"",chiusaDa:"",stato:"",tipoSet:"",produzione:"",dettagliCoProd:"",note:""});setModal(true);}} style={{border:brutBorder,background:C.magenta,color:C.white,padding:"6px 16px",fontWeight:700,fontSize:11,cursor:"pointer"}}>AGGIUNGI EVENTO</button></div>}
            {filledSlots.map((slot,i)=><SlotDetail key={slot.slotId} slot={slot} idx={i}/>)}
            <div style={{display:"flex",justifyContent:"flex-end",padding:"10px 20px"}}>
              <button onClick={()=>setDetailDay(null)} style={{border:brutBorder,background:C.white,padding:"6px 16px",fontWeight:700,fontSize:11,cursor:"pointer",boxShadow:brutShadow}}>CHIUDI</button>
            </div>
          </div>
        </div>);
      })()}
    </div>
  );
}
