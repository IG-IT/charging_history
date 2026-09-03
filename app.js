(() => {
  'use strict';
  const STORAGE_KEY = 'xpeng-g9-charge-log-v1';
  const DAILY_KEY = 'xpeng-g9-daily-energy-v1';
  const BACKUP_KEY = 'xpeng-g9-last-backup-at';
  const VEHICLE_KEY = 'xpeng-g9-vehicle-name';
  const DEFAULT_VEHICLE = 'XPENG G9';
  const $ = (id) => document.getElementById(id);
  const num = (v) => (v === '' || v === null || v === undefined || Number.isNaN(Number(v))) ? null : Number(v);
  const round = (v, n = 2) => v == null || !Number.isFinite(v) ? null : Number(v.toFixed(n));
  const esc = (s='') => String(s).replace(/[&<>'"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));

  let sessions = [];
  let dailyEntries = [];
  let filtered = [];
  let editingId = null;
  let toastTimer;
  let vehicleName = DEFAULT_VEHICLE;

  const headers = ['Datum','Plats / nätverk','Start-SOC %','Slut-SOC %','Laddad SOC %','Energi från laddare (kWh)','Tid (min)','Snitteffekt (kW)','Toppeffekt (kW)','Kostnad (SEK)','Snittpris (SEK/kWh)','Räckvidd start (km)','Räckvidd slut (km)','Tillagd räckvidd (km)','Kostnad / 100 km tillagd räckvidd (SEK)','Energi / 1 % SOC (kWh)','Mätarställning (km)','Anteckningar'];
  const dailyHeaders = ['Datum','Snittförbrukning (kWh/100 km)','Körsträcka (km)','Total förbrukning (kWh)','Snitthastighet (km/h)','Körning (kWh)','Regen. (kWh)','Parkering (kWh)','Elektriskt drivsystem (kWh)','A/C (kWh)','Övrigt (kWh)','Kommentar'];

  function makeId() { return (crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random().toString(16).slice(2)}`); }
  function isoDate(v) {
    if (!v) return '';
    if (v instanceof Date && !isNaN(v)) return `${v.getFullYear()}-${String(v.getMonth()+1).padStart(2,'0')}-${String(v.getDate()).padStart(2,'0')}`;
    if (typeof v === 'number' && window.XLSX) {
      const d = XLSX.SSF.parse_date_code(v); if (d) return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`;
    }
    const s = String(v).trim();
    const m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/); if (m) return `${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`;
    const d = new Date(s); return isNaN(d) ? '' : d.toISOString().slice(0,10);
  }
  function calc(s) {
    const chargedSoc = s.startSoc != null && s.endSoc != null ? s.endSoc - s.startSoc : (s.chargedSocHint ?? null);
    const avgPower = s.energy != null && s.minutes ? s.energy / s.minutes * 60 : null;
    const avgPrice = s.energy && s.cost != null ? s.cost / s.energy : null;
    const addedRange = s.rangeStart != null && s.rangeEnd != null ? s.rangeEnd - s.rangeStart : (s.addedRangeHint ?? null);
    const cost100 = s.cost != null && addedRange ? s.cost / addedRange * 100 : null;
    const energyPerSoc = s.energy != null && chargedSoc ? s.energy / chargedSoc : null;
    const {chargedSocHint, addedRangeHint, ...rest} = s;
    return {...rest, chargedSoc:round(chargedSoc,2), avgPower:round(avgPower,2), avgPrice:round(avgPrice,3), addedRange:round(addedRange,1), cost100:round(cost100,2), energyPerSoc:round(energyPerSoc,3)};
  }
  function normalize(o) {
    return calc({
      id:o.id || makeId(), date:isoDate(o.date), location:String(o.location || '').trim(),
      startSoc:num(o.startSoc), endSoc:num(o.endSoc), energy:num(o.energy), minutes:num(o.minutes),
      peakPower:num(o.peakPower), cost:num(o.cost), rangeStart:num(o.rangeStart), rangeEnd:num(o.rangeEnd),
      odometer:num(o.odometer), notes:String(o.notes || '').trim(),
      chargedSocHint:num(o.chargedSocHint??o.chargedSoc), addedRangeHint:num(o.addedRangeHint??o.addedRange)
    });
  }
  function fromSwedishRow(r) {
    return normalize({
      date:r['Datum'], location:r['Plats / nätverk'], startSoc:r['Start-SOC %'], endSoc:r['Slut-SOC %'],
      energy:r['Energi från laddare (kWh)'], minutes:r['Tid (min)'], peakPower:r['Toppeffekt (kW)'], cost:r['Kostnad (SEK)'],
      rangeStart:r['Räckvidd start (km)'], rangeEnd:r['Räckvidd slut (km)'], odometer:r['Mätarställning (km)'], notes:r['Anteckningar'],
      chargedSocHint:r['Laddad SOC %'], addedRangeHint:r['Tillagd räckvidd (km)']
    });
  }
  function fromGenericRow(r) {
    const lower = Object.fromEntries(Object.entries(r).map(([k,v]) => [String(k).toLowerCase().trim(),v]));
    const g = (...keys) => keys.map(k => lower[k]).find(v => v !== undefined);
    return normalize({
      date:g('date','datum','дата'), location:g('location','place','plats / nätverk','место / сеть','место'),
      startSoc:g('startsoc','start soc','start-soc %','start soc %'), endSoc:g('endsoc','end soc','slut-soc %','end soc %'),
      energy:g('energy','kwh','energi från laddare (kwh)','энергия, kwh'), minutes:g('minutes','time','tid (min)','время, min'),
      peakPower:g('peakpower','peak power','toppeffekt (kw)','пиковая мощность, kw'), cost:g('cost','kostnad (sek)','стоимость, sek'),
      rangeStart:g('rangestart','räckvidd start (km)','запас хода до, km'), rangeEnd:g('rangeend','räckvidd slut (km)','запас хода после, km'),
      odometer:g('odometer','mätarställning (km)','одометр, km'), notes:g('notes','anteckningar','заметки')
    });
  }

  function sessionKey(s) {
    return JSON.stringify([s.date,String(s.location||'').trim().toLowerCase(),s.startSoc,s.endSoc,s.energy,s.minutes,s.peakPower,s.cost,s.rangeStart,s.rangeEnd,s.odometer,String(s.notes||'').trim()]);
  }

  function normalizeDaily(r) {
    return {date:isoDate(r.date??r['Datum']),consumption:num(r.consumption??r['Snittförbrukning (kWh/100 km)']),distance:num(r.distance??r['Körsträcka (km)']),totalEnergy:num(r.totalEnergy??r['Total förbrukning (kWh)']),avgSpeed:num(r.avgSpeed??r['Snitthastighet (km/h)']),driving:num(r.driving??r['Körning (kWh)']),regen:num(r.regen??r['Regen. (kWh)']),parking:num(r.parking??r['Parkering (kWh)']),drivetrain:num(r.drivetrain??r['Elektriskt drivsystem (kWh)']),ac:num(r.ac??r['A/C (kWh)']),other:num(r.other??r['Övrigt (kWh)']),notes:String(r.notes??r['Kommentar']??'')};
  }

  function persist() { localStorage.setItem(STORAGE_KEY, JSON.stringify(sessions)); localStorage.setItem(DAILY_KEY,JSON.stringify(dailyEntries)); }
  function repairImportedDates(){
    const replacements=new Map();
    sessions.forEach(s=>{const m=s.notes.match(/(?:Start|Betald)\s+(\d{4}-\d{2}-\d{2})/i);if(!m||!s.date)return;const current=new Date(`${s.date}T12:00:00`),noted=new Date(`${m[1]}T12:00:00`);if(Math.abs(noted-current)===86400000)replacements.set(`${s.date}|${s.location}`,m[1]);});
    let changed=false; sessions=sessions.map(s=>{const date=replacements.get(`${s.date}|${s.location}`);if(!date)return s;changed=true;return {...s,date};}); return changed;
  }
  function loadLocal() { try { const raw = JSON.parse(localStorage.getItem(STORAGE_KEY)); if (Array.isArray(raw)) sessions = raw.map(normalize); const daily=JSON.parse(localStorage.getItem(DAILY_KEY)); if(Array.isArray(daily))dailyEntries=daily.map(normalizeDaily); if(repairImportedDates())persist(); } catch {} }

  function slug(s) { return String(s||'').trim().toLowerCase().normalize('NFKD').replace(/[̀-ͯ]/g,'').replace(/[^a-z0-9]+/g,'-').replace(/^-+|-+$/g,'') || 'charge-log'; }
  function fileSlug() { return slug(vehicleName); }
  function loadVehicleName() { vehicleName = localStorage.getItem(VEHICLE_KEY)?.trim() || DEFAULT_VEHICLE; }
  function applyVehicleName() {
    $('vehicleNameInput').value = vehicleName;
    $('vehicleEyebrow').textContent = `${vehicleName.toUpperCase()} · CHARGE LOG`;
    $('helpVehicleText').textContent = `Det här är din privata laddningslogg för ${vehicleName}. All data sparas lokalt i webbläsaren på den här enheten – inget skickas till någon server.`;
  }
  function saveVehicleName() {
    vehicleName = $('vehicleNameInput').value.trim() || DEFAULT_VEHICLE;
    localStorage.setItem(VEHICLE_KEY, vehicleName);
    applyVehicleName();
  }

  function money(v) { return v == null ? '—' : `${v.toLocaleString('sv-SE',{minimumFractionDigits:2,maximumFractionDigits:2})} SEK`; }
  function unit(v,u,d=1) { return v == null ? '—' : `${v.toLocaleString('sv-SE',{maximumFractionDigits:d})} ${u}`; }
  function dateLabel(s) { if (!s) return 'Inget datum'; const d = new Date(`${s}T12:00:00`); return isNaN(d) ? s : new Intl.DateTimeFormat('sv-SE',{day:'2-digit',month:'short',year:'numeric'}).format(d); }
  function render() {
    const q = $('searchInput').value.trim().toLowerCase();
    filtered = sessions.filter(s => !q || `${s.location} ${s.notes}`.toLowerCase().includes(q));
    const sort = $('sortSelect').value;
    filtered.sort((a,b) => sort==='date-asc' ? (a.date||'').localeCompare(b.date||'') : sort==='cost-desc' ? (b.cost||0)-(a.cost||0) : sort==='energy-desc' ? (b.energy||0)-(a.energy||0) : (b.date||'').localeCompare(a.date||''));

    const sum = k => sessions.reduce((a,s)=>a+(Number(s[k])||0),0);
    const totalEnergy=sum('energy'), totalCost=sum('cost'), totalRange=sum('addedRange');
    const powers=sessions.map(s=>s.avgPower).filter(v=>v!=null && isFinite(v));
    $('totalCost').textContent=money(totalCost);
    $('sessionCount').textContent=`${sessions.length} ${sessions.length===1?'session':'sessioner'}`;
    $('totalEnergy').textContent=unit(totalEnergy,'kWh',1);
    $('avgPrice').textContent=totalEnergy ? unit(totalCost/totalEnergy,'SEK/kWh',2) : '—';
    $('avgPower').textContent=powers.length ? unit(powers.reduce((a,b)=>a+b,0)/powers.length,'kW',1) : '—';
    $('totalRange').textContent=unit(totalRange,'km',0);
    $('listMeta').textContent=q ? `Hittade: ${filtered.length}` : 'Lokalt på denna enhet';
    renderAnalytics();
    renderDaily();
    renderBackupReminder();

    $('emptyState').hidden=filtered.length>0;
    $('sessionList').innerHTML=filtered.map(s=>{
      const socText=s.startSoc!=null || s.endSoc!=null ? `${s.startSoc ?? '?'} → ${s.endSoc ?? '?'}%` : null;
      const fill=s.endSoc!=null ? Math.max(0,Math.min(100,s.endSoc)) : 0;
      return `<button class="session-card" data-id="${esc(s.id)}">
        <div class="session-top"><div style="min-width:0"><div class="session-place">${esc(s.location||'Utan namn')}</div><div class="session-date">${esc(dateLabel(s.date))}</div></div><div class="session-cost">${esc(money(s.cost))}</div></div>
        <div class="session-bottom"><div class="metrics">
          ${s.energy!=null?`<span class="metric">⚡ ${esc(unit(s.energy,'kWh',2))}</span>`:''}
          ${s.avgPower!=null?`<span class="metric">↗ ${esc(unit(s.avgPower,'kW',1))}</span>`:''}
          ${socText?`<span class="metric">🔋 ${esc(socText)}</span>`:''}
          ${s.addedRange!=null?`<span class="metric">＋${esc(unit(s.addedRange,'km',0))}</span>`:''}
        </div>${socText?`<div class="soc-bar" title="${esc(socText)}"><div class="soc-fill" style="width:${fill}%"></div></div>`:''}</div>
      </button>`;
    }).join('');
  }
  function periodBounds(period) {
    const now=new Date(), year=now.getFullYear(), month=now.getMonth();
    if(period==='month') return [new Date(year,month,1),new Date(year,month+1,1)];
    if(period==='previous-month') return [new Date(year,month-1,1),new Date(year,month,1)];
    if(period==='year') return [new Date(year,0,1),new Date(year+1,0,1)];
    return [null,null];
  }
  function sessionsForPeriod(period) {
    const [start,end]=periodBounds(period);
    if(!start) return sessions;
    return sessions.filter(s=>{ const d=new Date(`${s.date}T12:00:00`); return !isNaN(d)&&d>=start&&d<end; });
  }
  function dailyForPeriod(period) {
    const [start,end]=periodBounds(period);
    if(!start)return dailyEntries;
    return dailyEntries.filter(x=>{const d=new Date(`${x.date}T12:00:00`);return !isNaN(d)&&d>=start&&d<end;});
  }
  function analyticsFor(items,daily=[]) {
    const total=k=>items.reduce((a,s)=>a+(Number(s[k])||0),0);
    const cost=total('cost'), energy=total('energy'), range=total('addedRange'), distance=daily.reduce((a,x)=>a+(x.distance||0),0), distanceDays=daily.filter(x=>x.distance!=null).length;
    return {cost,energy,range,distance,distanceDays,dailyRows:daily.length,avgPrice:energy?cost/energy:null,cost100:range?cost/range*100:null};
  }
  function renderAnalytics() {
    const period=$('periodSelect').value, items=sessionsForPeriod(period), a=analyticsFor(items,dailyForPeriod(period));
    const titles={month:'Denna månad','previous-month':'Förra månaden',year:'Detta år',all:'Alla tider'};
    $('analyticsTitle').textContent=titles[period];
    $('periodCost').textContent=money(a.cost);
    $('periodEnergy').textContent=unit(a.energy,'kWh',1);
    $('periodAvgPrice').textContent=unit(a.avgPrice,'SEK/kWh',2);
    $('periodCost100').textContent=unit(a.cost100,'SEK',2);
    $('periodSessions').textContent=items.length?`${items.length} ${items.length===1?'laddning':'laddningar'} · tillagd räckvidd +${unit(a.range,'km',0)}`:'Inga laddningar under perioden';
    $('periodCoverage').textContent=a.dailyRows?`Dagsloggen är ofullständig: ${unit(a.distance,'km',0)} registrerat för ${a.distanceDays} av ${a.dailyRows} dagar med data. Detta är inte den fullständiga körsträckan för perioden.`:'Faktisk körsträcka för denna period har ännu inte registrerats.';
    const comparison=$('monthComparison'); comparison.className=''; comparison.textContent='';
    if(period==='month') {
      const previous=analyticsFor(sessionsForPeriod('previous-month'),dailyForPeriod('previous-month'));
      if(previous.cost>0) {
        const change=(a.cost-previous.cost)/previous.cost*100;
        comparison.textContent=`${change>0?'▲':'▼'} ${Math.abs(change).toLocaleString('sv-SE',{maximumFractionDigits:0})}% jämfört med förra månaden`;
        comparison.className=change<=0?'positive':'negative';
      }
    }
  }
  function renderDaily(){
    const rows=dailyEntries.slice().sort((a,b)=>(b.date||'').localeCompare(a.date||''));
    const sum=k=>rows.reduce((a,x)=>a+(Number(x[k])||0),0), distance=sum('distance'), energy=sum('totalEnergy'), driving=sum('driving'), parking=sum('parking'), distanceDays=rows.filter(x=>x.distance!=null).length;
    $('dailyCount').textContent=`${distanceDays}/${rows.length} dagar`;
    $('dailySummary').innerHTML=`<div><span>Körsträcka enligt logg</span><b>${unit(distance,'km',1)}</b></div><div><span>Förbrukning under körning</span><b>${distance?unit(driving/distance*100,'kWh/100 km',1):'—'}</b></div><div><span>Total energi</span><b>${unit(energy,'kWh',1)}</b></div><div><span>Vid parkering</span><b>${unit(parking,'kWh',1)}</b></div>`;
    $('dailyList').innerHTML=rows.length?rows.map(x=>`<article class="daily-row"><div class="daily-row-head"><span>${esc(dateLabel(x.date))}</span><span>${esc(unit(x.distance,'km',1))}</span></div><div class="daily-row-metrics"><span>⚡ ${esc(unit(x.totalEnergy,'kWh',1))}</span><span>Snittförbrukning ${esc(unit(x.consumption,'kWh/100 km',1))}</span><span>Hastighet ${esc(unit(x.avgSpeed,'km/h',1))}</span><span>Parkering ${esc(unit(x.parking,'kWh',1))}</span><span>Regenerering ${esc(unit(x.regen,'kWh',1))}</span><span>A/C ${esc(unit(x.ac,'kWh',1))}</span></div></article>`).join(''):'<p class="muted">Ingen dagsdata ännu. Importera en uppdaterad Excel-fil.</p>';
  }
  function markBackup(){localStorage.setItem(BACKUP_KEY,new Date().toISOString());renderBackupReminder();}
  function renderBackupReminder(){
    const raw=localStorage.getItem(BACKUP_KEY), age=raw?(Date.now()-new Date(raw).getTime())/86400000:Infinity;
    $('backupReminder').hidden=age<30;
    $('backupReminderText').textContent=raw?'Senaste säkerhetskopian gjordes för mer än 30 dagar sedan.':'Ingen säkerhetskopia har gjorts ännu.';
  }
  function toast(msg){ const t=$('toast'); t.textContent=msg; t.classList.add('show'); clearTimeout(toastTimer); toastTimer=setTimeout(()=>t.classList.remove('show'),2600); }

  function openEditor(id=null) {
    editingId=id; const s=id?sessions.find(x=>x.id===id):null;
    $('dialogTitle').textContent=s?'Redigera laddning':'Ny laddning'; $('sessionId').value=id||'';
    const fields={date:s?.date||new Date().toISOString().slice(0,10),location:s?.location||'',startSoc:s?.startSoc??'',endSoc:s?.endSoc??'',energy:s?.energy??'',minutes:s?.minutes??'',peakPower:s?.peakPower??'',cost:s?.cost??'',rangeStart:s?.rangeStart??'',rangeEnd:s?.rangeEnd??'',odometer:s?.odometer??'',notes:s?.notes||''};
    Object.entries(fields).forEach(([k,v])=>$(k).value=v); $('deleteBtn').hidden=!s; updateComputed(); $('editDialog').showModal();
  }
  function formSession(){ const previous=editingId?sessions.find(x=>x.id===editingId):null; return normalize({id:editingId||makeId(),date:$('date').value,location:$('location').value,startSoc:$('startSoc').value,endSoc:$('endSoc').value,energy:$('energy').value,minutes:$('minutes').value,peakPower:$('peakPower').value,cost:$('cost').value,rangeStart:$('rangeStart').value,rangeEnd:$('rangeEnd').value,odometer:$('odometer').value,notes:$('notes').value,chargedSocHint:previous?.chargedSoc,addedRangeHint:previous?.addedRange}); }
  function updateComputed(){ const s=formSession(); $('computedPreview').innerHTML=`<div><span>Snitteffekt</span><b>${unit(s.avgPower,'kW',1)}</b></div><div><span>Energipris</span><b>${unit(s.avgPrice,'SEK/kWh',2)}</b></div><div><span>Tillagt</span><b>${unit(s.addedRange,'km',0)}</b></div><div><span>Kostnad / 100 km</span><b>${unit(s.cost100,'SEK',2)}</b></div>`; }
  function saveForm(e){ e.preventDefault(); const s=formSession(); if(!s.location){toast('Ange plats eller nätverk'); return;} const i=sessions.findIndex(x=>x.id===s.id); if(i>=0)sessions[i]=s; else sessions.push(s); persist(); render(); $('editDialog').close(); toast(i>=0?'Session uppdaterad':'Session tillagd'); }
  function deleteCurrent(){ if(!editingId)return; if(!confirm('Ta bort denna laddning?'))return; sessions=sessions.filter(s=>s.id!==editingId); persist(); render(); $('editDialog').close(); toast('Session borttagen'); }

  function exportRows(){ return sessions.slice().sort((a,b)=>(b.date||'').localeCompare(a.date||'')).map(s=>[s.date||'',s.location,s.startSoc,s.endSoc,s.chargedSoc,s.energy,s.minutes,s.avgPower,s.peakPower,s.cost,s.avgPrice,s.rangeStart,s.rangeEnd,s.addedRange,s.cost100,s.energyPerSoc,s.odometer,s.notes]); }
  function downloadBlob(content,name,type){ const a=document.createElement('a'); a.href=URL.createObjectURL(new Blob([content],{type})); a.download=name; a.click(); setTimeout(()=>URL.revokeObjectURL(a.href),1000); }
  function downloadJsonBackup(prefix=`${fileSlug()}-charging`) {
    const now=new Date(), stamp=`${now.toISOString().slice(0,10)}-${String(now.getHours()).padStart(2,'0')}${String(now.getMinutes()).padStart(2,'0')}${String(now.getSeconds()).padStart(2,'0')}`;
    const backup={version:2,exportedAt:now.toISOString(),sessions,dailyEntries};
    downloadBlob(JSON.stringify(backup,null,2),`${prefix}-${stamp}.json`,'application/json');
  }
  function makeWorkbook(){
    const ws=XLSX.utils.aoa_to_sheet([headers,...exportRows()]);
    ws['!cols']=[{wch:12},{wch:34},{wch:12},{wch:12},{wch:12},{wch:22},{wch:11},{wch:18},{wch:18},{wch:15},{wch:21},{wch:20},{wch:20},{wch:21},{wch:40},{wch:24},{wch:20},{wch:80}];
    const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,'Laddningslogg');
    const summary=[[`${vehicleName} — Laddningsöversikt`],[],['Mätvärde','Värde'],['Antal laddsessioner',sessions.length],['Total energi (kWh)',sumForExport('energy')],['Total kostnad (SEK)',sumForExport('cost')],['Snittpris (SEK/kWh)',sumForExport('energy')?sumForExport('cost')/sumForExport('energy'):null],['Total tillagd räckvidd (km)',sumForExport('addedRange')]];
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(summary),'Översikt');
    const dailyRows=dailyEntries.slice().sort((a,b)=>(a.date||'').localeCompare(b.date||'')).map(x=>[x.date,x.consumption,x.distance,x.totalEnergy,x.avgSpeed,x.driving,x.regen,x.parking,x.drivetrain,x.ac,x.other,x.notes]);
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([dailyHeaders,...dailyRows]),'Daglig energiförbrukning'); return wb;
  }
  function exportData(type){ const stamp=new Date().toISOString().slice(0,10), prefix=fileSlug(); if(type==='json'){ downloadBlob(JSON.stringify({version:2,exportedAt:new Date().toISOString(),sessions,dailyEntries},null,2),`${prefix}-charging-${stamp}.json`,'application/json'); }
    else if(type==='csv'){ const matrix=[headers,...exportRows()]; const csv=matrix.map(row=>row.map(v=>{const s=v==null?'':String(v).replace(/"/g,'""'); return /[;"\n]/.test(s)?`"${s}"`:s;}).join(';')).join('\r\n'); downloadBlob('\ufeff'+csv,`${prefix}-charging-${stamp}.csv`,'text/csv;charset=utf-8'); }
    else if(type==='xlsx'){ if(!window.XLSX){toast('XLSX-modulen kunde inte laddas. Kontrollera internetanslutningen.'); return;} XLSX.writeFile(makeWorkbook(),`${prefix}-charging-${stamp}.xlsx`); }
    if(type!=='csv')markBackup(); $('exportDialog').close(); toast(`Export ${type.toUpperCase()} klar`); }
  function sumForExport(k){ return sessions.reduce((a,s)=>a+(Number(s[k])||0),0); }

  async function importFile(file){ try { let incoming=[],incomingDaily=[]; const ext=(file.name.split('.').pop()||'').toLowerCase(); if(ext==='json'){ const data=JSON.parse(await file.text()); const arr=Array.isArray(data)?data:data.sessions; if(!Array.isArray(arr))throw new Error('JSON does not contain sessions'); incoming=arr.map(x => x['Plats / nätverk']!==undefined ? fromSwedishRow(x) : normalize(x)); if(Array.isArray(data.dailyEntries))incomingDaily=data.dailyEntries.map(normalizeDaily); }
      else if(ext==='csv'){ const text=await file.text(); if(window.XLSX){ const wb=XLSX.read(text,{type:'string'}); incoming=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:null}).map(r=>r['Plats / nätverk']!==undefined?fromSwedishRow(r):fromGenericRow(r)); } else { throw new Error('XLSX library unavailable for CSV parser'); } }
      else { if(!window.XLSX)throw new Error('XLSX library unavailable'); const wb=XLSX.read(await file.arrayBuffer(),{cellDates:true}); const ws=wb.Sheets['Laddningslogg']||wb.Sheets[wb.SheetNames[0]]; incoming=XLSX.utils.sheet_to_json(ws,{defval:null,raw:true}).map(r=>r['Plats / nätverk']!==undefined?fromSwedishRow(r):fromGenericRow(r)); if(wb.Sheets['Daglig energiförbrukning'])incomingDaily=XLSX.utils.sheet_to_json(wb.Sheets['Daglig energiförbrukning'],{defval:null,raw:true}).map(normalizeDaily); }
      incoming=incoming.filter(s=>s.date||s.location||s.energy!=null); incomingDaily=incomingDaily.filter(x=>x.date); if(!incoming.length&&!incomingDaily.length)throw new Error('No rows found');
      const known=new Set(sessions.map(sessionKey)), uniqueIncoming=[]; let duplicateCount=0;
      incoming.forEach(s=>{const key=sessionKey(s);if(known.has(key)){duplicateCount++;return;}known.add(key);uniqueIncoming.push(s);});
      if(sessions.length && (uniqueIncoming.length||incomingDaily.length) && !confirm(`Lägg till ${uniqueIncoming.length} nya laddningar och uppdatera ${incomingDaily.length} dagar?${duplicateCount?`\n\nDubbletter hoppades över: ${duplicateCount}.`:''}`)) return;
      sessions=[...sessions,...uniqueIncoming]; const byDate=new Map(dailyEntries.map(x=>[x.date,x])); incomingDaily.forEach(x=>byDate.set(x.date,x)); dailyEntries=[...byDate.values()]; persist(); render();
      toast(uniqueIncoming.length||incomingDaily.length?`Importerat: ${uniqueIncoming.length} laddningar, ${incomingDaily.length} dagar${duplicateCount?`; hoppade över ${duplicateCount}`:''}`:`Inga nya data — ${duplicateCount} dubbletter hoppades över`);
    } catch(e){ console.error(e); alert(`Kunde inte importera filen.\n${e.message}`); } finally { $('fileInput').value=''; } }

  async function resetData(){
    if(!window.XLSX){ alert('Återställning avbruten: Excel-modulen kunde inte laddas. Anslut till internet och försök igen — säkerhetskopior av JSON och Excel krävs innan återställning.'); return; }
    if(!confirm('Innan återställningen laddar appen ner två fullständiga säkerhetskopior: JSON och Excel. Fortsätta?'))return;
    downloadJsonBackup(`${fileSlug()}-before-reset`);
    const now=new Date(), stamp=`${now.toISOString().slice(0,10)}-${String(now.getHours()).padStart(2,'0')}${String(now.getMinutes()).padStart(2,'0')}${String(now.getSeconds()).padStart(2,'0')}`;
    XLSX.writeFile(makeWorkbook(),`${fileSlug()}-before-reset-${stamp}.xlsx`);
    markBackup();
    if(!confirm('JSON- och Excel-kopiorna har skickats till Nedladdningar. Kontrollera att Safari har sparat båda filerna. Radera lokala data nu?')){
      toast('Återställning avbruten — data sparad'); return;
    }
    localStorage.removeItem(STORAGE_KEY); localStorage.removeItem(DAILY_KEY); sessions=[]; dailyEntries=[]; render(); toast('Data raderad. Importera en fil för att ladda dem igen.');
  }
  function backupNow(){
    if(!window.XLSX){alert('Excel-modulen är inte tillgänglig. Starta om appen och försök igen.');return;}
    downloadJsonBackup(`${fileSlug()}-backup`); const stamp=new Date().toISOString().slice(0,10); XLSX.writeFile(makeWorkbook(),`${fileSlug()}-backup-${stamp}.xlsx`); markBackup(); toast('JSON och Excel har skickats till Nedladdningar');
  }

  $('addBtn').addEventListener('click',()=>openEditor()); $('emptyAddBtn').addEventListener('click',()=>openEditor());
  $('closeDialogBtn').addEventListener('click',()=>$('editDialog').close()); $('sessionForm').addEventListener('submit',saveForm); $('deleteBtn').addEventListener('click',deleteCurrent);
  $('sessionForm').addEventListener('input',updateComputed); $('searchInput').addEventListener('input',render); $('sortSelect').addEventListener('change',render);
  $('periodSelect').addEventListener('change',renderAnalytics);
  $('sessionList').addEventListener('click',e=>{ const card=e.target.closest('[data-id]'); if(card)openEditor(card.dataset.id); });
  $('importBtn').addEventListener('click',()=>$('fileInput').click()); $('fileInput').addEventListener('change',e=>e.target.files[0]&&importFile(e.target.files[0]));
  $('exportBtn').addEventListener('click',()=>$('exportDialog').showModal()); $('closeExportBtn').addEventListener('click',()=>$('exportDialog').close()); document.querySelectorAll('[data-export]').forEach(b=>b.addEventListener('click',()=>exportData(b.dataset.export)));
  $('helpBtn').addEventListener('click',()=>$('helpDialog').showModal()); $('closeHelpBtn').addEventListener('click',()=>$('helpDialog').close());
  $('vehicleNameInput').addEventListener('input',saveVehicleName);
  $('resetBtn').addEventListener('click',resetData);
  $('backupNowBtn').addEventListener('click',backupNow);
  $('settingsBtn').addEventListener('click',()=>{const menu=$('settingsMenu'),open=menu.hidden;menu.hidden=!open;$('settingsBtn').setAttribute('aria-expanded',String(open));});
  $('settingsMenu').addEventListener('click',e=>{if(e.target.closest('button')){$('settingsMenu').hidden=true;$('settingsBtn').setAttribute('aria-expanded','false');}});
  document.addEventListener('click',e=>{if(!e.target.closest('.header-buttons')){$('settingsMenu').hidden=true;$('settingsBtn').setAttribute('aria-expanded','false');}});

  loadVehicleName(); applyVehicleName(); loadLocal(); render();
  if('serviceWorker' in navigator) window.addEventListener('load',()=>navigator.serviceWorker.register('./sw.js').catch(console.warn));
})();
