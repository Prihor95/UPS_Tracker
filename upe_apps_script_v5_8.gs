// ══════════════════════════════════════════════════
//  UPE Floor Tracker — Google Apps Script v5.4
//  Hoja: Operadores con CRUD desde HTML
//  v5.4: EstadoActual separado de historial
// ══════════════════════════════════════════════════

const SS = SpreadsheetApp.getActiveSpreadsheet();
function getSheet(n){return SS.getSheetByName(n);}
function ts(){return Utilities.formatDate(new Date(),"America/Mexico_City","yyyy-MM-dd HH:mm:ss");}

function corsResponse(data){
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Init hojas ──────────────────────────────────
function initSheets(){
  const sheets={
    "EstadoActual": ["estacion","upe_id","timestamp","turno","fecha","pct_inicio","pct_fin","locomotion","etiqueta"],
    "Asignaciones": ["estacion","upe_id","timestamp","turno","fecha","pct_inicio","pct_fin","locomotion","etiqueta"],
    "Inventario":   ["upe_id","estado","timestamp"],
    "Log":          ["timestamp","mensaje"],
    "Historial":    ["upe_id","tipo","estacion","inicio","fin","turno_ini","turno_fin","dur_total_min","dur_efec_min","pct_ini","pct_fin","locomotion","etiqueta"],
    "Promedios":    ["upe_id","tipo","ciclos_normal","prom_normal","ciclos_loco","prom_loco","ultima_act"],
    "Operadores":   ["maquina","escenario","break","operador","turno"]
  };
  Object.entries(sheets).forEach(([name,headers])=>{
    let sh=SS.getSheetByName(name);
    if(!sh){
      sh=SS.insertSheet(name);
      sh.getRange(1,1,1,headers.length).setValues([headers])
        .setFontWeight("bold").setBackground("#E8F0FE");
      sh.setFrozenRows(1);
    }
  });
  // Migrar Asignaciones → EstadoActual si EstadoActual está vacía
  migrarEstadoActualSiVacia();
  SpreadsheetApp.getUi().alert("✅ Hojas creadas/verificadas");
}

// Migra el estado actual desde la hoja Asignaciones (compatibilidad v5.3→v5.4)
function migrarEstadoActualSiVacia(){
  const shEA=getSheet("EstadoActual");
  if(!shEA||shEA.getLastRow()>1) return; // ya tiene datos
  const viejo=getAsignacionesDesdeHoja("Asignaciones");
  const rows=Object.entries(viejo).map(([est,a])=>[
    est,a.upe_id,a.timestamp,a.turno||'','',
    a.pct_inicio||100,a.pct_fin||10,a.locomotion||false,a.etiqueta||''
  ]);
  if(rows.length>0)
    shEA.getRange(2,1,rows.length,9).setValues(rows);
  logEntry('Migración EstadoActual: '+rows.length+' estaciones');
}

// Upsert en EstadoActual: busca la fila de esa estación y la reemplaza (o agrega)
function upsertEstadoActual(estacion, rowData){
  const sh=getSheet("EstadoActual");
  if(!sh) return;
  const data=sh.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    if(String(data[i][0])===String(estacion)){
      sh.getRange(i+1,1,1,9).setValues([rowData]);
      return;
    }
  }
  sh.appendRow(rowData);
}

// Elimina la fila de una estación en EstadoActual (cuando se libera)
function deleteEstadoActual(estacion){
  const sh=getSheet("EstadoActual");
  if(!sh) return;
  const data=sh.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    if(String(data[i][0])===String(estacion)){
      sh.deleteRow(i+1);
      return;
    }
  }
}

// ── Settings compartidos ─────────────────────────
function getSettings(){
  try{
    const r=PropertiesService.getScriptProperties().getProperty('upe_settings');
    return r?JSON.parse(r):null;
  }catch(e){return null;}
}

function saveSettingsProps(s){
  PropertiesService.getScriptProperties()
    .setProperty('upe_settings',JSON.stringify(s));
}

function cmdGuardarSettings(body){
  if(!body.settings) return {ok:false,error:'Sin settings'};
  saveSettingsProps(body.settings);
  logEntry('Settings actualizados');
  return {ok:true};
}

// ── Layout ──────────────────────────────────────
function getLayout(){
  try{
    const r=PropertiesService.getScriptProperties().getProperty('upe_layout');
    return r?JSON.parse(r):null;
  }catch(e){return null;}
}
function saveLayout(stations,timestamp,autor){
  PropertiesService.getScriptProperties()
    .setProperty('upe_layout',JSON.stringify({stations,timestamp,autor}));
  logEntry('Layout publicado por '+(autor||'?'));
}

// ── Turnos ──────────────────────────────────────
function getTurno(d){
  const dow=d.getDay();
  if(dow===0||dow===6) return 'finde';
  const m=d.getHours()*60+d.getMinutes();
  if(m>=7*60&&m<16*60+36) return 'mañana';
  if(m>=16*60+36&&m<22*60) return 'tarde';
  return 'noche';
}

function calcBreaks(ini,fin){
  const breaks=[[13*60,14*60],[19*60+6,19*60+36],[2*60,3*60]];
  let total=0,cursor=new Date(ini);
  while(cursor<fin){
    const dow=cursor.getDay(),esLab=dow>=1&&dow<=5;
    const dayS=new Date(cursor);dayS.setHours(0,0,0,0);
    const dayE=new Date(dayS);dayE.setDate(dayE.getDate()+1);
    const sI=(Math.max(cursor,dayS)-dayS)/60000;
    const sF=(Math.min(fin,dayE)-dayS)/60000;
    if(esLab) breaks.forEach(([bs,be])=>
      total+=Math.max(0,Math.min(sF,be)-Math.max(sI,bs)));
    cursor=dayE;
  }
  return total;
}

// ── GET ─────────────────────────────────────────
function doGet(e){
  try{
    const action=(e.parameter&&e.parameter.action)||'all';
    let result;

    if(action==='all'){
      const layout=getLayout();
      const savedSettings=getSettings();
      result={
        ok:true,
        asignaciones: getAsignaciones(),
        inventario:   getInventario(),
        log:          getLog(50),
        promedios:    getPromedios(),
        operadores:   getOperadores(),
        stations:     layout?layout.stations:{},
        layout_ts:    layout?layout.timestamp:null,
        settings:     savedSettings||null
      };
    } else if(action==='operadores'){
      result={ok:true, operadores:getOperadores()};
    } else if(action==='layout'){
      const layout=getLayout();
      result=layout?{ok:true,...layout}:{ok:false,error:'Sin layout guardado'};
    } else {
      result={ok:false,error:'Acción GET no reconocida: '+action};
    }
    return corsResponse(result);
  }catch(err){return corsResponse({ok:false,error:err.message});}
}

function getAsignaciones(){
  // Lee de EstadoActual (estado presente, O(n) sobre pocas filas)
  // Si no existe la hoja, cae back a la hoja Asignaciones histórica
  const sh=getSheet("EstadoActual")||getSheet("Asignaciones");
  return getAsignacionesDesdeHoja(sh.getName());
}

// Lee asignaciones activas desde cualquier hoja con el mismo schema
function getAsignacionesDesdeHoja(nombreHoja){
  const sh=getSheet(nombreHoja);
  if(!sh) return {};
  const data=sh.getDataRange().getValues();
  if(data.length<=1) return {};
  const r={};
  data.slice(1).forEach(row=>{
    if(row[0]&&row[1]) r[String(row[0])]={
      upe_id:    String(row[1]),
      timestamp: String(row[2]),
      turno:     String(row[3]||''),
      pct_inicio:Number(row[5]||100),
      pct_fin:   Number(row[6]||10),
      locomotion:row[7]==='true'||row[7]===true,
      etiqueta:  String(row[8]||'')
    };
  });
  return r;
}

function getInventario(){
  const sh=getSheet("Inventario"),data=sh.getDataRange().getValues();
  if(data.length<=1) return {c:[],d:[]};
  const latest={};
  data.slice(1).forEach(row=>{if(row[0])latest[String(row[0])]=String(row[1]);});
  const r={c:[],d:[]};
  Object.entries(latest).forEach(([id,est])=>{
    if(est==='c')r.c.push(id);
    else if(est==='d')r.d.push(id);
  });
  return r;
}

function getLog(limit){
  const sh=getSheet("Log"),data=sh.getDataRange().getValues();
  if(data.length<=1) return [];
  return data.slice(1).slice(-limit).reverse()
    .map(r=>String(r[0])+' · '+String(r[1]));
}

function getPromedios(){
  const sh=getSheet("Promedios"),data=sh.getDataRange().getValues();
  if(data.length<=1) return {};
  const r={};
  data.slice(1).forEach(row=>{
    if(row[0]) r[String(row[0])]={
      tipo:           String(row[1]||''),
      ciclosNormal:   Number(row[2]||0),
      promedioNormal: Number(row[3]||0),
      ciclosLoco:     Number(row[4]||0),
      promedioLoco:   Number(row[5]||0)
    };
  });
  return r;
}

function getOperadores(){
  const sh=getSheet("Operadores");
  if(!sh) return [];
  const data=sh.getDataRange().getValues();
  if(data.length<=1) return [];
  const headers=data[0].map(h=>String(h).toLowerCase().trim());
  return data.slice(1)
    .map(row=>{
      const obj={};
      headers.forEach((h,i)=>obj[h]=String(row[i]||'').trim());
      // Normalizar escenario: quitar espacios internos "J 39" → "J39"
      if(obj.escenario) obj.escenario=obj.escenario.replace(/\s+/g,'').toUpperCase();
      // Break vacío = sin asignar
      if(!obj.break) obj.break='';
      return obj;
    })
    .filter(o=>o.maquina||o.escenario);
}

// ── POST ────────────────────────────────────────
function doPost(e){
  try{
    const body=JSON.parse(e.postData.contents),action=body.action;
    if(action==='asignar')            return corsResponse(cmdAsignar(body));
    if(action==='asignar_forzado')    return corsResponse(cmdAsignarForzado(body));
    if(action==='liberar')            return corsResponse(cmdLiberar(body));
    if(action==='inventario_add')     return corsResponse(cmdInvAdd(body.upe_id,body.estado));
    if(action==='inventario_move')    return corsResponse(cmdInvAdd(body.upe_id,body.estado));
    if(action==='inventario_remove')  return corsResponse(cmdInvRemove(body.upe_id));
    if(action==='guardar_layout')     return corsResponse(cmdLayout(body));
    if(action==='guardar_operadores') return corsResponse(cmdGuardarOperadores(body));
    if(action==='guardar_settings')   return corsResponse(cmdGuardarSettings(body));
    if(action==='limpiar_asignaciones') return corsResponse(cmdLimpiarAsignaciones());
    return corsResponse({ok:false,error:'Acción POST no reconocida: '+action});
  }catch(err){return corsResponse({ok:false,error:err.message});}
}

// ── Comandos ────────────────────────────────────
function getTipoBat(id){
  if(!id) return '';
  const p=String(id).charAt(0).toUpperCase();
  return p==='A'?'A':p==='O'?'O':'';
}

function cmdAsignar(body){
  const{estacion,upe_id,pct_inicio,pct_fin,locomotion,etiqueta}=body;
  if(!estacion||!upe_id) return {ok:false,error:'Faltan parámetros'};
  const now=new Date(),turno=getTurno(now),tipo=getTipoBat(upe_id);
  const tsStr=ts(),fecha=Utilities.formatDate(now,"America/Mexico_City","yyyy-MM-dd");
  const pctI=pct_inicio!==undefined?Number(pct_inicio):100;
  const pctF=pct_fin!==undefined?Number(pct_fin):10;
  const esLoco=locomotion===true||locomotion==='true';
  const etq=etiqueta||'';
  const prev=getAsignaciones()[estacion];
  if(prev&&prev.upe_id)
    cerrarCiclo(prev.upe_id,estacion,prev.timestamp,tsStr,prev.turno,turno,
                prev.pct_inicio,pctF,prev.locomotion,prev.etiqueta);
  // Historial (append log)
  getSheet("Asignaciones").appendRow([estacion,upe_id,tsStr,turno,fecha,pctI,pctF,esLoco,etq]);
  // Estado actual (upsert — lo que lee el HTML al sincronizar)
  upsertEstadoActual(estacion,[estacion,upe_id,tsStr,turno,fecha,pctI,pctF,esLoco,etq]);
  logEntry('Asignado '+upe_id+' → '+estacion+' ['+turno+'] '+(esLoco?'🚂':''));
  return {ok:true,turno,timestamp:tsStr};
}

function cmdLiberar(body){
  const{estacion,pct_fin}=body;
  if(!estacion) return {ok:false,error:'Falta estación'};
  const now=new Date(),turno=getTurno(now),tsStr=ts();
  const pctF=pct_fin!==undefined?Number(pct_fin):10;
  const prev=getAsignaciones()[estacion];
  if(prev&&prev.upe_id)
    cerrarCiclo(prev.upe_id,estacion,prev.timestamp,tsStr,prev.turno,turno,
                prev.pct_inicio,pctF,prev.locomotion,prev.etiqueta);
  // Historial (append log)
  getSheet("Asignaciones").appendRow([estacion,'',tsStr+' (liberada)',turno,'','','','','']);
  // Estado actual: eliminar la fila de esta estación
  deleteEstadoActual(estacion);
  logEntry(estacion+': liberada ['+turno+']');
  return {ok:true};
}

function cerrarCiclo(upe_id,estacion,iniStr,finStr,turnoIni,turnoFin,pctIni,pctFin,esLoco,etiqueta){
  try{
    const ini=new Date(String(iniStr).replace(' ','T'));
    const fin=new Date(String(finStr).replace(' ','T'));
    const durTotal=Math.round((fin-ini)/60000);
    const durEfec=Math.max(0,durTotal-Math.round(calcBreaks(ini,fin)));
    const tipo=getTipoBat(upe_id);
    const esFalla=String(etiqueta||'')==='falla';
    getSheet("Historial").appendRow([
      upe_id,tipo,estacion,iniStr,finStr,
      turnoIni,turnoFin,durTotal,durEfec,
      pctIni,pctFin,esLoco,etiqueta
    ]);
    if(!esFalla&&durEfec>0)
      actualizarPromedio(upe_id,tipo,durEfec,esLoco===true||esLoco==='true');
  }catch(e){logEntry('Error cerrarCiclo '+upe_id+': '+e.message);}
}

function actualizarPromedio(upe_id,tipo,durEfec,esLoco){
  const sh=getSheet("Promedios"),data=sh.getDataRange().getValues();
  let rowIdx=-1,cN=0,pN=0,cL=0,pL=0;
  data.forEach((row,i)=>{
    if(i>0&&String(row[0])===String(upe_id)){
      rowIdx=i+1;cN=Number(row[2]||0);pN=Number(row[3]||0);
      cL=Number(row[4]||0);pL=Number(row[5]||0);
    }
  });
  if(esLoco){cL++;pL=Math.round((pL*(cL-1)+durEfec)/cL);}
  else{cN++;pN=Math.round((pN*(cN-1)+durEfec)/cN);}
  const nowStr=ts();
  if(rowIdx>0) sh.getRange(rowIdx,2,1,6).setValues([[tipo,cN,pN,cL,pL,nowStr]]);
  else sh.appendRow([upe_id,tipo,cN,pN,cL,pL,nowStr]);
}

function cmdInvAdd(upe_id,estado){
  if(!upe_id||!estado) return {ok:false,error:'Faltan params'};
  getSheet("Inventario").appendRow([upe_id,estado,ts()]);
  logEntry('Inventario: '+upe_id+' → '+(estado==='c'?'cargada':estado==='d'?'descargada':'eliminada'));
  return {ok:true};
}

function cmdInvRemove(upe_id){
  getSheet("Inventario").appendRow([upe_id,'removed',ts()]);
  logEntry('Inventario: '+upe_id+' eliminada');
  return {ok:true};
}

function cmdLayout(body){
  if(!body.stations) return {ok:false,error:'Sin stations'};
  saveLayout(body.stations,body.timestamp||ts(),body.autor||'?');
  return {ok:true};
}

// ── Operadores: reescribe la hoja completa ──────
function cmdGuardarOperadores(body){
  const filas=body.operadores;
  if(!Array.isArray(filas)) return {ok:false,error:'operadores debe ser array'};

  let sh=getSheet("Operadores");
  if(!sh){
    sh=SS.insertSheet("Operadores");
    sh.setFrozenRows(1);
  }

  // Limpiar contenido (excepto header si existe)
  const lastRow=sh.getLastRow();
  const lastCol=sh.getLastColumn()||5;
  if(lastRow>1) sh.getRange(2,1,lastRow-1,lastCol).clearContent();

  // Asegurar headers
  sh.getRange(1,1,1,5).setValues([['maquina','escenario','break','operador','turno']])
    .setFontWeight('bold').setBackground('#E8F0FE');

  // Escribir filas
  if(filas.length>0){
    const rows=filas.map(f=>[
      f.maquina||'',
      (f.escenario||'').replace(/\s+/g,'').toUpperCase(),
      f.break||'',
      f.operador||'',
      f.turno||'mañana'
    ]);
    sh.getRange(2,1,rows.length,5).setValues(rows);
  }

  logEntry('Operadores actualizados: '+filas.length+' filas');
  return {ok:true,filas:filas.length};
}

// Asignar forzado — respeta el timestamp que viene del cliente
function cmdAsignarForzado(body){
  const{estacion,upe_id,timestamp,turno,pct_inicio,pct_fin,locomotion,etiqueta}=body;
  if(!estacion||!upe_id) return {ok:false,error:'Faltan parámetros'};
  const tsStr=timestamp||ts();
  const fecha=tsStr.slice(0,10);
  const pctI=pct_inicio!==undefined?Number(pct_inicio):100;
  const pctF=pct_fin!==undefined?Number(pct_fin):10;
  const esLoco=locomotion===true||locomotion==='true';
  const etq=etiqueta||'';
  // Historial
  getSheet("Asignaciones").appendRow([estacion,upe_id,tsStr,turno||'',fecha,pctI,pctF,esLoco,etq]);
  // Estado actual
  upsertEstadoActual(estacion,[estacion,upe_id,tsStr,turno||'',fecha,pctI,pctF,esLoco,etq]);
  logEntry('Asignación forzada: '+upe_id+' → '+estacion+' ('+tsStr+')');
  return {ok:true};
}

// Limpiar todas las asignaciones — borra contenido dejando headers
function cmdLimpiarAsignaciones(){
  const sh=getSheet("Asignaciones");
  const lastRow=sh.getLastRow();
  if(lastRow>1) sh.getRange(2,1,lastRow-1,sh.getLastColumn()).clearContent();
  // Limpiar también EstadoActual
  const shEA=getSheet("EstadoActual");
  if(shEA){
    const lastEA=shEA.getLastRow();
    if(lastEA>1) shEA.getRange(2,1,lastEA-1,shEA.getLastColumn()).clearContent();
  }
  logEntry('Asignaciones limpiadas manualmente');
  return {ok:true};
}

function logEntry(msg){getSheet("Log").appendRow([ts(),msg]);}
