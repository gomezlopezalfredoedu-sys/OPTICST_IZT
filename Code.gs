/*******************************************************
 * BOUTIQUE DE LENTES YAEL
 * Code.gs — V2 CORREGIDO
 * Compatible con el Sheet Maestro:
 * Boutique_de_Lentes_Yael_Sheet_Maestro
 *
 * IMPORTANTE:
 * - No usa appendRow(), porque el Sheet Maestro tiene
 *   fórmulas preparadas en las filas de datos.
 * - Respeta exactamente los nombres de hojas/columnas.
 * - Las fórmulas del Sheet siguen siendo las responsables
 *   de los campos calculados de VENTAS.
 *******************************************************/

const APP = {
  TZ: 'America/Mexico_City',
  SHEETS: [
    'CONFIGURACION','CATALOGOS','CLIENTES','HISTORIA_CLINICA',
    'LENSOMETRIA','GRADUACIONES','COTIZACIONES','VENTAS','ABONOS',
    'ENTREGAS','RECOMPENSAS','GARANTIAS','VISITAS','COSTOS_PRODUCCION'
  ],
  ID_FIELD: {
    CLIENTES:'ID_CLIENTE', HISTORIA_CLINICA:'ID_HISTORIA',
    LENSOMETRIA:'ID_LENSOMETRIA', GRADUACIONES:'ID_GRADUACION',
    COTIZACIONES:'ID_COTIZACION', VENTAS:'ID_VENTA', ABONOS:'ID_ABONO',
    ENTREGAS:'ID_ENTREGA', RECOMPENSAS:'ID_RECOMPENSA',
    GARANTIAS:'ID_GARANTIA', VISITAS:'ID_VISITA'
  }
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Boutique de Lentes Yael')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- BASE ---------- */

function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function sh_(name) {
  const s = ss_().getSheetByName(name);
  if (!s) throw new Error('No existe la hoja "' + name + '".');
  return s;
}

function headers_(sheetName) {
  const s = sh_(sheetName);
  return s.getRange(1,1,1,s.getLastColumn()).getDisplayValues()[0].filter(String);
}

function col_(sheetName, field) {
  const h = headers_(sheetName);
  const i = h.indexOf(field);
  if (i < 0) throw new Error('La columna "'+field+'" no existe en '+sheetName+'.');
  return i + 1;
}

function firstFreeRow_(sheetName) {
  const s = sh_(sheetName);
  const idField = APP.ID_FIELD[sheetName];
  if (!idField) throw new Error('No hay ID configurado para '+sheetName+'.');
  const c = col_(sheetName,idField);
  const last = Math.max(2,s.getLastRow());
  const vals = s.getRange(2,c,Math.max(1,last-1),1).getDisplayValues().flat();
  for (let i=0;i<vals.length;i++) if (!String(vals[i]).trim()) return i+2;
  return last+1;
}

function rowObject_(sheetName,rowNum) {
  const s=sh_(sheetName), h=headers_(sheetName);
  const vals=s.getRange(rowNum,1,1,h.length).getValues()[0];
  const o={_row:rowNum};
  h.forEach((k,i)=>o[k]=vals[i]);
  return o;
}

function rows_(sheetName) {
  const s=sh_(sheetName), h=headers_(sheetName);
  const last=s.getLastRow();
  if(last<2)return [];
  const vals=s.getRange(2,1,last-1,h.length).getValues();
  const idField=APP.ID_FIELD[sheetName];
  return vals.map((r,i)=>{
    const o={_row:i+2};
    h.forEach((k,j)=>o[k]=r[j]);
    return o;
  }).filter(o=>!idField || String(o[idField]??'').trim()!=='');
}

function writeRow_(sheetName,rowNum,data) {
  const s=sh_(sheetName), h=headers_(sheetName);
  const old=s.getRange(rowNum,1,1,h.length).getValues()[0];
  const row=h.map((k,i)=>Object.prototype.hasOwnProperty.call(data,k)?data[k]:old[i]);
  s.getRange(rowNum,1,1,h.length).setValues([row]);

  // Si la fila no tenía fórmulas (por ejemplo, al pasar la fila 1000),
  // copia las fórmulas de la fila anterior en las columnas no capturadas.
  if(rowNum>2){
    const prev=s.getRange(rowNum-1,1,1,h.length).getFormulas()[0];
    const cur=s.getRange(rowNum,1,1,h.length).getFormulas()[0];
    const dataRange=s.getRange(rowNum,1,1,h.length);
    const formulaRow=cur.slice();
    for(let i=0;i<h.length;i++){
      if(!Object.prototype.hasOwnProperty.call(data,h[i]) && !formulaRow[i] && prev[i]){
        formulaRow[i]=prev[i].replace(/\d+/g, String(rowNum));
      }
    }
    if(formulaRow.some(Boolean)) dataRange.setFormulas([formulaRow]);
  }
  return rowObject_(sheetName,rowNum);
}

function insertRowData_(sheetName,data) {
  const row=firstFreeRow_(sheetName);
  return writeRow_(sheetName,row,data);
}

function findOne_(sheetName,field,value) {
  const q=String(value??'').trim().toLowerCase();
  return rows_(sheetName).find(r=>String(r[field]??'').trim().toLowerCase()===q)||null;
}

function now_(){return new Date();}
function fmtDate_(d){return Utilities.formatDate(new Date(d),APP.TZ,'yyyy-MM-dd');}
function fmtTime_(d){return Utilities.formatDate(new Date(d),APP.TZ,'HH:mm:ss');}
function money_(v){const n=Number(String(v??'').replace(/[$,\s]/g,''));return isNaN(n)?0:n;}
function bool_(v){return String(v??'').toLowerCase()==='sí'||String(v??'').toLowerCase()==='si'||v===true;}

function config_(){
  const out={};
  rows_('CONFIGURACION').forEach(r=>{if(r.CLAVE)out[r.CLAVE]=r.VALOR;});
  return out;
}

function nextId_(sheetName) {
  const cfg=config_();
  const field=APP.ID_FIELD[sheetName];
  const initialKey={
    CLIENTES:'FOLIO_CLIENTE_INICIAL',
    HISTORIA_CLINICA:'FOLIO_HISTORIA_INICIAL',
    LENSOMETRIA:'FOLIO_LENSOMETRIA_INICIAL',
    GRADUACIONES:'FOLIO_GRADUACION_INICIAL',
    COTIZACIONES:'FOLIO_COTIZACION_INICIAL',
    VENTAS:'FOLIO_VENTA_INICIAL',
    ENTREGAS:'FOLIO_ENTREGA_INICIAL',
    RECOMPENSAS:'FOLIO_RECOMPENSA_INICIAL',
    GARANTIAS:'FOLIO_GARANTIA_INICIAL',
    VISITAS:'FOLIO_VISITA_INICIAL'
  }[sheetName];

  if(sheetName==='ABONOS') throw new Error('Los abonos usan folio por venta.');

  const initial=String(cfg[initialKey]||'');
  const m=initial.match(/^(.*?)(\d+)$/);
  const prefix=m?m[1]:'';
  let max=m?Number(m[2])-1:0;
  rows_(sheetName).forEach(r=>{
    const x=String(r[field]??'').match(/(\d+)$/);
    if(x)max=Math.max(max,Number(x[1]));
  });
  return prefix+String(max+1).padStart(m?m[2].length:6,'0');
}

/* ---------- CONFIGURACIÓN / CATÁLOGOS ---------- */

function getConfig(){
  return {config:config_(),catalogos:getCatalogos(),costos:rows_('COSTOS_PRODUCCION')};
}

function getCatalogos(){
  const out={};
  rows_('CATALOGOS').filter(r=>String(r.ACTIVO).toLowerCase()!=='no' && String(r.ACTIVO).toLowerCase()!=='false')
    .forEach(r=>{
      if(!r.CATALOGO)return;
      if(!out[r.CATALOGO])out[r.CATALOGO]=[];
      out[r.CATALOGO].push({valor:r.VALOR||'',subvalor:r.SUBVALOR||'',orden:Number(r.ORDEN)||0});
    });
  Object.keys(out).forEach(k=>out[k].sort((a,b)=>a.orden-b.orden));
  return out;
}

/* ---------- CLIENTES ---------- */

function buscarClientes(termino){
  const q=String(termino||'').trim().toLowerCase();
  return rows_('CLIENTES').filter(r=>{
    if(!q)return true;
    return [r.ID_CLIENTE,r.NOMBRE,r.WHATSAPP].some(v=>String(v??'').toLowerCase().includes(q));
  }).reverse().slice(0,50);
}

function obtenerCliente(id){
  const c=findOne_('CLIENTES','ID_CLIENTE',id);
  if(!c)throw new Error('Cliente no encontrado.');
  return c;
}

function crearCliente(data){
  const nombre=String(data.NOMBRE||'').trim();
  if(!nombre)throw new Error('El nombre del cliente es obligatorio.');
  const wa=String(data.WHATSAPP||'').trim();
  const dup=rows_('CLIENTES').filter(r=>
    (wa && String(r.WHATSAPP).trim()===wa) ||
    String(r.NOMBRE).trim().toLowerCase()===nombre.toLowerCase()
  );
  const n=now_(),id=nextId_('CLIENTES');
  const c=insertRowData_('CLIENTES',{
    ID_CLIENTE:id,NOMBRE:nombre,
    FECHA_NACIMIENTO:data.FECHA_NACIMIENTO?new Date(data.FECHA_NACIMIENTO):'',
    SEXO:data.SEXO||'',WHATSAPP:wa,
    FECHA_ALTA:fmtDate_(n),HORA_ALTA:fmtTime_(n),
    ULTIMA_COMPRA:'',ESTADO:'ACTIVO'
  });
  SpreadsheetApp.flush();
  return {cliente:obtenerCliente(id),duplicados:dup};
}

function editarCliente(data){
  const c=obtenerCliente(data.ID_CLIENTE);
  return writeRow_('CLIENTES',c._row,{
    NOMBRE:String(data.NOMBRE||'').trim(),
    FECHA_NACIMIENTO:data.FECHA_NACIMIENTO?new Date(data.FECHA_NACIMIENTO):'',
    SEXO:data.SEXO||'',WHATSAPP:String(data.WHATSAPP||'').trim()
  });
}


function prepararCliente(idCliente){
  const c=obtenerCliente(idCliente);
  return {
    cliente:c,
    historia:historiaCliente(idCliente),
    lensometria:ultimaLensometria(idCliente),
    graduacion:ultimaGraduacion(idCliente),
    recompensa:obtenerRecompensa(idCliente)
  };
}

/* ---------- HISTORIA CLÍNICA ---------- */

function historiaCliente(idCliente){
  return rows_('HISTORIA_CLINICA').filter(r=>String(r.ID_CLIENTE)===String(idCliente)).reverse();
}

function guardarHistoria(data){
  const n=now_();
  return insertRowData_('HISTORIA_CLINICA',{
    ID_HISTORIA:nextId_('HISTORIA_CLINICA'),
    ID_CLIENTE:data.ID_CLIENTE,FECHA:fmtDate_(n),HORA:fmtTime_(n),
    MOTIVO_CONSULTA:data.MOTIVO_CONSULTA||'',
    MOTIVO_OTRO:data.MOTIVO_OTRO||'',
    USUARIO_LENTES:data.USUARIO_LENTES||'',
    ANTECEDENTES_HEREDOFAMILIARES:Array.isArray(data.ANTECEDENTES_HEREDOFAMILIARES)?data.ANTECEDENTES_HEREDOFAMILIARES.join(', '):(data.ANTECEDENTES_HEREDOFAMILIARES||''),
    ANTECEDENTES_PERSONALES:Array.isArray(data.ANTECEDENTES_PERSONALES)?data.ANTECEDENTES_PERSONALES.join(', '):(data.ANTECEDENTES_PERSONALES||''),
    OTROS_ANTECEDENTES:data.OTROS_ANTECEDENTES||''
  });
}

/* ---------- LENSOMETRÍA ---------- */

function ultimaLensometria(idCliente){
  return rows_('LENSOMETRIA').filter(r=>String(r.ID_CLIENTE)===String(idCliente))
    .sort((a,b)=>new Date(b.FECHA)-new Date(a.FECHA))[0]||null;
}

function guardarLensometria(data){
  const n=now_();
  return insertRowData_('LENSOMETRIA',{
    ID_LENSOMETRIA:nextId_('LENSOMETRIA'),
    ID_HISTORIA:data.ID_HISTORIA||'',ID_CLIENTE:data.ID_CLIENTE,FECHA:fmtDate_(n),
    ESF_OD:data.ESF_OD||'',CIL_OD:data.CIL_OD||'',EJE_OD:data.EJE_OD||'',ADD_OD:data.ADD_OD||'',
    ESF_OI:data.ESF_OI||'',CIL_OI:data.CIL_OI||'',EJE_OI:data.EJE_OI||'',ADD_OI:data.ADD_OI||'',
    DISEÑO:data.DISEÑO||'',MATERIAL:data.MATERIAL||'',INDICE:data.INDICE||'',
    TRATAMIENTO:data.TRATAMIENTO||'',COLOR_FOTO:data.COLOR_FOTO||'',
    AV_OD:data.AV_OD||'',AV_OI:data.AV_OI||'',AV_AO:data.AV_AO||'',
    OBSERVACIONES:data.OBSERVACIONES||''
  });
}

/* ---------- GRADUACIONES ---------- */

function graduacionesCliente(idCliente){
  return rows_('GRADUACIONES').filter(r=>String(r.ID_CLIENTE)===String(idCliente)).reverse();
}

function ultimaGraduacion(idCliente){
  return graduacionesCliente(idCliente)[0]||null;
}

function guardarGraduacion(data){
  const n=now_();
  return insertRowData_('GRADUACIONES',{
    ID_GRADUACION:nextId_('GRADUACIONES'),ID_CLIENTE:data.ID_CLIENTE,
    ID_HISTORIA:data.ID_HISTORIA||'',FECHA:fmtDate_(n),TIPO:data.TIPO||'',
    AV_SC_OD:data.AV_SC_OD||'',AV_SC_OI:data.AV_SC_OI||'',AV_SC_AO:data.AV_SC_AO||'',
    ESF_OD:data.ESF_OD||'',CIL_OD:data.CIL_OD||'',EJE_OD:data.EJE_OD||'',
    ESF_OI:data.ESF_OI||'',CIL_OI:data.CIL_OI||'',EJE_OI:data.EJE_OI||'',
    ADD:data.ADD||'',AV_CC_OD:data.AV_CC_OD||'',AV_CC_OI:data.AV_CC_OI||'',AV_CC_AO:data.AV_CC_AO||'',
    DISEÑO_RECOMENDADO:data.DISEÑO_RECOMENDADO||'',
    MATERIAL_RECOMENDADO:data.MATERIAL_RECOMENDADO||'',
    INDICE_RECOMENDADO:data.INDICE_RECOMENDADO||'',
    TRATAMIENTO_RECOMENDADO:data.TRATAMIENTO_RECOMENDADO||'',
    COLOR_FOTO:data.COLOR_FOTO||'',OBSERVACIONES:data.OBSERVACIONES||''
  });
}

/* ---------- COTIZACIONES ---------- */

function crearCotizacion(data){
  const n=now_(),cfg=config_();
  const dias=Number(cfg.VIGENCIA_COTIZACION_DIAS||15);
  const f=new Date(n);f.setDate(f.getDate()+dias);
  const total=money_(data.SUBTOTAL)-money_(data.DESCUENTO);
  return insertRowData_('COTIZACIONES',{
    ID_COTIZACION:nextId_('COTIZACIONES'),ID_CLIENTE:data.ID_CLIENTE,
    FECHA:fmtDate_(n),HORA:fmtTime_(n),VIGENCIA:fmtDate_(f),
    ESTADO:'Vigente',
    LENTE_1_DISEÑO:data.LENTE_1_DISEÑO||'',LENTE_1_MATERIAL:data.LENTE_1_MATERIAL||'',
    LENTE_1_TRATAMIENTO:data.LENTE_1_TRATAMIENTO||'',LENTE_1_PRECIO:money_(data.LENTE_1_PRECIO),
    LENTE_2_DISEÑO:data.LENTE_2_DISEÑO||'',LENTE_2_MATERIAL:data.LENTE_2_MATERIAL||'',
    LENTE_2_TRATAMIENTO:data.LENTE_2_TRATAMIENTO||'',LENTE_2_PRECIO:money_(data.LENTE_2_PRECIO),
    LENTE_3_DISEÑO:data.LENTE_3_DISEÑO||'',LENTE_3_MATERIAL:data.LENTE_3_MATERIAL||'',
    LENTE_3_TRATAMIENTO:data.LENTE_3_TRATAMIENTO||'',LENTE_3_PRECIO:money_(data.LENTE_3_PRECIO),
    ARMAZON_1_MODELO:data.ARMAZON_1_MODELO||'',ARMAZON_1_DESCRIPCION:data.ARMAZON_1_DESCRIPCION||'',ARMAZON_1_PRECIO:money_(data.ARMAZON_1_PRECIO),
    ARMAZON_2_MODELO:data.ARMAZON_2_MODELO||'',ARMAZON_2_DESCRIPCION:data.ARMAZON_2_DESCRIPCION||'',ARMAZON_2_PRECIO:money_(data.ARMAZON_2_PRECIO),
    ARMAZON_3_MODELO:data.ARMAZON_3_MODELO||'',ARMAZON_3_DESCRIPCION:data.ARMAZON_3_DESCRIPCION||'',ARMAZON_3_PRECIO:money_(data.ARMAZON_3_PRECIO),
    SUBTOTAL:total+money_(data.DESCUENTO),DESCUENTO:money_(data.DESCUENTO),ID_VENTA:''
  });
}

function buscarCotizacion(id){
  const q=findOne_('COTIZACIONES','ID_COTIZACION',id);
  if(!q)throw new Error('Cotización no encontrada.');
  return q;
}

/* ---------- VENTAS ---------- */

function costoArmazon_(etiqueta,manual){
  if(String(etiqueta||'')==='Otro')return money_(manual);
  const r=rows_('COSTOS_PRODUCCION').find(x=>
    String(x.CONCEPTO)===String(etiqueta) && String(x.TIPO)==='Armazón'
  );
  return r?money_(r.COSTO):0;
}

function costoProceso_(concepto){
  const r=rows_('COSTOS_PRODUCCION').find(x=>String(x.CONCEPTO)===String(concepto));
  return r?money_(r.COSTO):0;
}

function costosVenta_(d){
  const diseño=String(d.DISEÑO_ARMAZON||'');
  const material=String(d.MATERIAL_MICA||'');
  let bisel=0,ranura=0,perforado=0,pulido=0;

  if(d.REBISSEL_MONTAJE){
    const map={
      'Completo':'Bisel - Completo',
      'Ranurado':'Bisel - Ranurado',
      '3 piezas tornillo':'Bisel - 3 piezas tornillo',
      '3 piezas grapa':'Bisel - 3 piezas grapa'
    };
    bisel=costoProceso_(map[diseño]||'');
  }

  // 3 piezas tornillo ya incluye perforaciones y montaje.
  if(diseño==='Ranurado' && (material==='CR39'||material==='Policarbonato')){
    ranura=costoProceso_('Ranura - '+material);
  }
  if(diseño==='3 piezas grapa' && (material==='CR39'||material==='Policarbonato')){
    perforado=costoProceso_('Perforado - '+material);
  }
  if(d.PULIDO){
    pulido=costoProceso_('Pulido');
  }
  return {bisel,ranura,perforado,pulido};
}

function crearOActualizarEntrega_(venta){
  const ex=rows_('ENTREGAS').find(r=>String(r.ID_VENTA)===String(venta.ID_VENTA));
  const n=now_();
  const d={
    ID_VENTA:venta.ID_VENTA,ID_CLIENTE:venta.ID_CLIENTE,
    FECHA_ENTREGA:venta.FECHA_ENTREGA||'',
    ESTADO:venta.ESTADO_ENTREGA||'Pendiente de elaboración',
    FECHA_CAMBIO_ESTADO:fmtDate_(n),RECORDATORIO_GENERADO:'NO'
  };
  if(ex)return writeRow_('ENTREGAS',ex._row,d);
  return insertRowData_('ENTREGAS',{
    ID_ENTREGA:nextId_('ENTREGAS'),...d,FECHA_RECORDATORIO_ABONO:''
  });
}

function programarVisita_(idCliente,idVenta,fechaVenta){
  const cfg=config_(),meses=Number(cfg.MESES_PROXIMA_VISITA||10);
  const f=new Date(fechaVenta);f.setMonth(f.getMonth()+meses);
  const ex=rows_('VISITAS').find(r=>String(r.ID_VENTA)===String(idVenta));
  const d={
    ID_CLIENTE:idCliente,ID_VENTA:idVenta,FECHA_PROGRAMADA:fmtDate_(f),
    FECHA_ORIGINAL:fmtDate_(f),ESTADO:'Pendiente',FECHA_RECORDATORIO:'',
    RECORDATORIO_GENERADO:'NO',OBSERVACIONES:''
  };
  if(ex)return writeRow_('VISITAS',ex._row,d);
  return insertRowData_('VISITAS',{ID_VISITA:nextId_('VISITAS'),...d});
}

function crearVenta(data){
  if(!data.ID_CLIENTE)throw new Error('Selecciona un cliente.');
  if(!data.TIPO_VENTA)throw new Error('Selecciona el tipo de venta.');

  const tipo=String(data.TIPO_VENTA);
  if(tipo==='Lentes completos'){
    if(!data.ID_GRADUACION)throw new Error('Los lentes completos requieren una graduación.');
    if(!data.MODELO_ARMAZON || !data.DISEÑO_ARMAZON || !data.MATERIAL_ARMAZON)
      throw new Error('Completa los datos del armazón.');
    if(!data.DISEÑO_MICA || !data.MATERIAL_MICA || !data.INDICE_MICA || !data.TRATAMIENTO_MICA)
      throw new Error('Completa los datos de las micas.');
  }
  if(tipo==='Solo micas'){
    if(!data.ID_GRADUACION)throw new Error('Solo micas requiere una graduación.');
    if(!data.DISEÑO_MICA || !data.MATERIAL_MICA || !data.INDICE_MICA || !data.TRATAMIENTO_MICA)
      throw new Error('Completa los datos de las micas.');
  }
  if(tipo==='Solo armazón' && !data.MODELO_ARMAZON)
    throw new Error('Solo armazón requiere el modelo.');

  const n=now_(),id=nextId_('VENTAS');
  const costos=costosVenta_(data);
  const costoAr=costoArmazon_(data.ETIQUETA_COSTO_ARMAZON,data.COSTO_ARMAZON_MANUAL);

  const precios=[
    money_(data.PRECIO_ARMAZON), money_(data.PRECIO_MICA),
    money_(data.PRECIO_ACCESORIO_1),money_(data.PRECIO_ACCESORIO_2),
    money_(data.PRECIO_ACCESORIO_3),money_(data.PRECIO_ACCESORIO_4),
    money_(data.PRECIO_ACCESORIO_5),money_(data.PRECIO_REBISSEL_MONTAJE),
    money_(data.PRECIO_OTRO)
  ];
  const subtotal=precios.reduce((a,b)=>a+b,0);
  const descuento=Math.max(0,money_(data.DESCUENTO));
  if(descuento>subtotal)throw new Error('El descuento no puede ser mayor al subtotal.');
  const total=Math.max(0,subtotal-descuento);
  const costoProduccion=costos.bisel+costos.ranura+costos.perforado+costos.pulido;
  const costoTotal=costoAr+money_(data.COSTO_INTERNO_MICA)+costoProduccion;
  const utilidad=Math.max(0,total-costoTotal);

  const venta=insertRowData_('VENTAS',{
    ID_VENTA:id,ID_CLIENTE:data.ID_CLIENTE,ID_COTIZACION:data.ID_COTIZACION||'',
    FECHA:fmtDate_(n),HORA:fmtTime_(n),TIPO_VENTA:tipo,
    ID_GRADUACION:data.ID_GRADUACION||'',ORIGEN_GRADUACION:data.ORIGEN_GRADUACION||'',
    FECHA_ENTREGA:data.FECHA_ENTREGA||'',ESTADO_PAGO:total===0?'Liquidada':'Pendiente',
    ESTADO_ENTREGA:data.ESTADO_ENTREGA||'Pendiente de elaboración',
    SUBTOTAL:subtotal,DESCUENTO:descuento,TOTAL:total,TOTAL_PAGADO:0,SALDO:total,
    FECHA_LIQUIDACION:total===0?fmtDate_(n):'',
    MODELO_ARMAZON:data.MODELO_ARMAZON||'',COLOR_ARMAZON:data.COLOR_ARMAZON||'',
    DISEÑO_ARMAZON:data.DISEÑO_ARMAZON||'',MATERIAL_ARMAZON:data.MATERIAL_ARMAZON||'',
    PRECIO_ARMAZON:money_(data.PRECIO_ARMAZON),
    ETIQUETA_COSTO_ARMAZON:data.ETIQUETA_COSTO_ARMAZON||'',
    COSTO_ARMAZON:costoAr,
    DISEÑO_MICA:data.DISEÑO_MICA||'',MATERIAL_MICA:data.MATERIAL_MICA||'',
    INDICE_MICA:data.INDICE_MICA||'',TRATAMIENTO_MICA:data.TRATAMIENTO_MICA||'',
    COLOR_FOTO:data.COLOR_FOTO||'',PRECIO_MICA:money_(data.PRECIO_MICA),
    COSTO_INTERNO_MICA:money_(data.COSTO_INTERNO_MICA),
    ACCESORIO_1:data.ACCESORIO_1||'',PRECIO_ACCESORIO_1:money_(data.PRECIO_ACCESORIO_1),
    ACCESORIO_2:data.ACCESORIO_2||'',PRECIO_ACCESORIO_2:money_(data.PRECIO_ACCESORIO_2),
    ACCESORIO_3:data.ACCESORIO_3||'',PRECIO_ACCESORIO_3:money_(data.PRECIO_ACCESORIO_3),
    ACCESORIO_4:data.ACCESORIO_4||'',PRECIO_ACCESORIO_4:money_(data.PRECIO_ACCESORIO_4),
    ACCESORIO_5:data.ACCESORIO_5||'',PRECIO_ACCESORIO_5:money_(data.PRECIO_ACCESORIO_5),
    REBISSEL_MONTAJE:data.REBISSEL_MONTAJE||'',PRECIO_REBISSEL_MONTAJE:money_(data.PRECIO_REBISSEL_MONTAJE),
    OTRO_CONCEPTO:data.OTRO_CONCEPTO||'',PRECIO_OTRO:money_(data.PRECIO_OTRO),
    COSTO_BISEL:costos.bisel,COSTO_RANURA:costos.ranura,
    COSTO_PERFORADO:costos.perforado,COSTO_PULIDO:costos.pulido,
    COSTO_PRODUCCION:costoProduccion,COSTO_TOTAL:costoTotal,UTILIDAD_ESTIMADA:utilidad
  });

  if(data.ID_COTIZACION){
    const q=findOne_('COTIZACIONES','ID_COTIZACION',data.ID_COTIZACION);
    if(q)writeRow_('COTIZACIONES',q._row,{ESTADO:'Convertida',ID_VENTA:id});
  }

  const c=obtenerCliente(data.ID_CLIENTE);
  writeRow_('CLIENTES',c._row,{ULTIMA_COMPRA:fmtDate_(n)});

  crearOActualizarEntrega_(venta);
  programarVisita_(data.ID_CLIENTE,id,n);

  if(tipo==='Lentes completos' && data.USO_RECOMPENSA!=='Sí'){
    registrarCompraRecompensa_(data.ID_CLIENTE,id,data.REFERIDO_POR||'',data.COMPRA_POR_REFERIDO||'No');
    // Si esta compra pertenece a un referido, también cuenta para quien refirió.
    if(data.COMPRA_POR_REFERIDO==='Sí' && data.REFERIDO_POR &&
       String(data.REFERIDO_POR)!==String(data.ID_CLIENTE)){
      registrarCompraRecompensa_(data.REFERIDO_POR,id,data.ID_CLIENTE,'Sí');
    }
  }
  return obtenerVenta(id);
}

function buscarVentas(termino){
  const q=String(termino||'').toLowerCase();
  const ventas=rows_('VENTAS').reverse();
  return ventas.filter(v=>{
    if(!q)return true;
    const c=findOne_('CLIENTES','ID_CLIENTE',v.ID_CLIENTE);
    return String(v.ID_VENTA).toLowerCase().includes(q) ||
      String(v.ID_CLIENTE).toLowerCase().includes(q) ||
      (c && (String(c.NOMBRE).toLowerCase().includes(q)||String(c.WHATSAPP).toLowerCase().includes(q)));
  }).slice(0,50);
}

function obtenerVenta(idVenta){
  const v=findOne_('VENTAS','ID_VENTA',idVenta);
  if(!v)throw new Error('Venta no encontrada.');
  v.cliente=obtenerCliente(v.ID_CLIENTE);
  v.abonos=rows_('ABONOS').filter(a=>String(a.ID_VENTA)===String(idVenta)).reverse();
  return v;
}

/* ---------- ABONOS ---------- */

function registrarAbono(data){
  const lock=LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    const v=findOne_('VENTAS','ID_VENTA',data.ID_VENTA);
    if(!v)throw new Error('Venta no encontrada.');

    const total=money_(v.TOTAL);
    const pagadoActual=money_(v.TOTAL_PAGADO);
    const saldo=Math.max(0,total-pagadoActual);
    const monto=money_(data.MONTO);

    if(monto<=0)throw new Error('El abono debe ser mayor a $0.');
    if(monto>saldo+0.001)throw new Error('El abono no puede superar el saldo pendiente.');
    if(!data.METODO_PAGO)throw new Error('Selecciona el método de pago.');

    const prev=rows_('ABONOS').filter(a=>String(a.ID_VENTA)===String(data.ID_VENTA));
    const numero=prev.length+1;
    const id='AB-'+data.ID_VENTA+'-'+String(numero).padStart(2,'0');
    const n=now_(),nuevo=Math.max(0,saldo-monto);
    const estado=nuevo===0?'Liquidada':'Parcial';

    insertRowData_('ABONOS',{
      ID_ABONO:id,ID_VENTA:data.ID_VENTA,ID_CLIENTE:v.ID_CLIENTE,
      NUMERO_ABONO:numero,FECHA:fmtDate_(n),HORA:fmtTime_(n),
      MONTO:monto,METODO_PAGO:data.METODO_PAGO,
      SALDO_ANTERIOR:saldo,SALDO_RESULTANTE:nuevo,
      ESTADO:nuevo===0?'Liquidado':'Abono'
    });

    writeRow_('VENTAS',v._row,{
      TOTAL_PAGADO:pagadoActual+monto,
      SALDO:nuevo,
      ESTADO_PAGO:estado,
      FECHA_LIQUIDACION:nuevo===0?fmtDate_(n):''
    });
    SpreadsheetApp.flush();
    return {abono:findOne_('ABONOS','ID_ABONO',id),venta:obtenerVenta(data.ID_VENTA)};
  }finally{lock.releaseLock();}
}

function buscarAbonos(termino){
  const q=String(termino||'').toLowerCase();
  return rows_('ABONOS').reverse().filter(a=>{
    if(!q)return true;
    const c=findOne_('CLIENTES','ID_CLIENTE',a.ID_CLIENTE);
    return String(a.ID_ABONO).toLowerCase().includes(q) ||
      String(a.ID_VENTA).toLowerCase().includes(q) ||
      (c&&(String(c.NOMBRE).toLowerCase().includes(q)||String(c.WHATSAPP).toLowerCase().includes(q)));
  }).slice(0,50);
}

/* ---------- ENTREGAS ---------- */

function entregas(filtro){
  const f=filtro||{},q=String(f.termino||'').toLowerCase();
  return rows_('ENTREGAS').reverse().filter(e=>{
    if(f.estado && String(e.ESTADO)!==String(f.estado))return false;
    if(!q)return true;
    const c=findOne_('CLIENTES','ID_CLIENTE',e.ID_CLIENTE);
    return String(e.ID_VENTA).toLowerCase().includes(q) ||
      (c&&(String(c.NOMBRE).toLowerCase().includes(q)||String(c.WHATSAPP).toLowerCase().includes(q)));
  }).slice(0,100);
}

function cambiarEntrega(data){
  const e=findOne_('ENTREGAS','ID_ENTREGA',data.ID_ENTREGA);
  if(!e)throw new Error('Entrega no encontrada.');
  const n=now_();
  const estado=data.ESTADO;
  const r=writeRow_('ENTREGAS',e._row,{
    ESTADO:estado,FECHA_ENTREGA:data.FECHA_ENTREGA||e.FECHA_ENTREGA,
    FECHA_CAMBIO_ESTADO:fmtDate_(n)
  });
  const v=findOne_('VENTAS','ID_VENTA',e.ID_VENTA);
  if(v)writeRow_('VENTAS',v._row,{
    ESTADO_ENTREGA:estado,
    FECHA_ENTREGA:data.FECHA_ENTREGA||v.FECHA_ENTREGA
  });
  return r;
}

/* ---------- RECOMPENSAS ---------- */

function recompensaActual_(idCliente){
  const rows=rows_('RECOMPENSAS').filter(r=>String(r.ID_CLIENTE)===String(idCliente));
  return rows.sort((a,b)=>Number(b.CICLO||0)-Number(a.CICLO||0))[0]||null;
}

function registrarCompraRecompensa_(idCliente,idVenta,referidoPor,compraPorReferido){
  let r=recompensaActual_(idCliente);
  let ciclo=r?Number(r.CICLO||1):1;
  let compras=r?Number(r.COMPRAS_VALIDAS||0):0;

  if(r && String(r.ESTADO)==='Redimida'){
    ciclo++; compras=0; r=null;
  }
  if(compras>=6)return r;

  compras++;
  const mica=compras>=4?'Sí':'No';
  const tratamiento=compras>=5?'Sí':'No';
  const disponible=compras>=6?'Sí':'No';
  const estado=compras>=6?'Recompensa disponible':'En curso';
  const data={
    ID_RECOMPENSA:r?r.ID_RECOMPENSA:nextId_('RECOMPENSAS'),
    ID_CLIENTE:idCliente,COMPRAS_VALIDAS:compras,
    MICA_DESBLOQUEADA:mica,TRATAMIENTO_DESBLOQUEADO:tratamiento,
    RECOMPENSA_DISPONIBLE:disponible,
    FECHA_DESBLOQUEO:compras>=6?fmtDate_(now_):(r?r.FECHA_DESBLOQUEO:''),
    FECHA_REDENCION:r?r.FECHA_REDENCION:'',
    ESTADO:estado,ID_VENTA_RECOMPENSA:r?r.ID_VENTA_RECOMPENSA:'',
    REFERIDO_POR:referidoPor||'',COMPRA_POR_REFERIDO:compraPorReferido||'No',CICLO:ciclo
  };
  if(r)return writeRow_('RECOMPENSAS',r._row,data);
  return insertRowData_('RECOMPENSAS',data);
}

function obtenerRecompensa(idCliente){return recompensaActual_(idCliente);}

function buscarRecompensas(termino){
  const q=String(termino||'').toLowerCase();
  return rows_('RECOMPENSAS').reverse().filter(r=>{
    if(!q)return true;
    const c=findOne_('CLIENTES','ID_CLIENTE',r.ID_CLIENTE);
    return String(r.ID_RECOMPENSA).toLowerCase().includes(q)||
      String(r.ID_CLIENTE).toLowerCase().includes(q)||
      (c&&(String(c.NOMBRE).toLowerCase().includes(q)||String(c.WHATSAPP).toLowerCase().includes(q)));
  }).slice(0,50);
}

function redimirRecompensa(data){
  const r=recompensaActual_(data.ID_CLIENTE);
  if(!r||Number(r.COMPRAS_VALIDAS)<6||String(r.ESTADO)!=='Recompensa disponible')
    throw new Error('El cliente no tiene una recompensa disponible.');
  return writeRow_('RECOMPENSAS',r._row,{
    FECHA_REDENCION:fmtDate_(now_),
    ESTADO:'Redimida',
    ID_VENTA_RECOMPENSA:data.ID_VENTA_RECOMPENSA||''
  });
}

/* ---------- GARANTÍAS ---------- */

function garantias(termino){
  const q=String(termino||'').toLowerCase();
  return rows_('GARANTIAS').reverse().filter(g=>{
    if(!q)return true;
    const c=findOne_('CLIENTES','ID_CLIENTE',g.ID_CLIENTE);
    return String(g.ID_GARANTIA).toLowerCase().includes(q)||
      String(g.ID_VENTA).toLowerCase().includes(q)||
      (c&&String(c.NOMBRE).toLowerCase().includes(q));
  }).slice(0,100);
}

function crearGarantia(data){
  const n=now_();
  return insertRowData_('GARANTIAS',{
    ID_GARANTIA:nextId_('GARANTIAS'),ID_VENTA:data.ID_VENTA,
    ID_CLIENTE:data.ID_CLIENTE,FECHA_REPORTE:fmtDate_(n),
    TIPO:data.TIPO||'',DESCRIPCION:data.DESCRIPCION||'',
    SOLUCION:data.SOLUCION||'',ESTADO:data.ESTADO||'Pendiente',FECHA_RESOLUCION:''
  });
}

function actualizarGarantia(data){
  const g=findOne_('GARANTIAS','ID_GARANTIA',data.ID_GARANTIA);
  if(!g)throw new Error('Garantía no encontrada.');
  return writeRow_('GARANTIAS',g._row,{
    TIPO:data.TIPO||g.TIPO,DESCRIPCION:data.DESCRIPCION||g.DESCRIPCION,
    SOLUCION:data.SOLUCION||g.SOLUCION,ESTADO:data.ESTADO||g.ESTADO,
    FECHA_RESOLUCION:data.ESTADO==='Resuelta'?fmtDate_(now_):''
  });
}

/* ---------- VISITAS ---------- */

function visitas(){
  return rows_('VISITAS').filter(v=>String(v.ESTADO)!=='Realizada')
    .sort((a,b)=>new Date(a.FECHA_PROGRAMADA)-new Date(b.FECHA_PROGRAMADA)).slice(0,100);
}

function actualizarVisita(data){
  const v=findOne_('VISITAS','ID_VISITA',data.ID_VISITA);
  if(!v)throw new Error('Visita no encontrada.');
  return writeRow_('VISITAS',v._row,{
    FECHA_PROGRAMADA:data.FECHA_PROGRAMADA||v.FECHA_PROGRAMADA,
    ESTADO:data.ESTADO||v.ESTADO,OBSERVACIONES:data.OBSERVACIONES||v.OBSERVACIONES
  });
}

/* ---------- RESUMEN ---------- */

function resumen(){
  const ventas=rows_('VENTAS'),abonos=rows_('ABONOS'),hoy=fmtDate_(now_());
  const mes=hoy.slice(0,7),anio=hoy.slice(0,4);

  const info=a=>({
    operaciones:a.length,
    total:a.reduce((s,r)=>s+money_(r.TOTAL),0),
    cobrado:a.reduce((s,r)=>s+money_(r.TOTAL_PAGADO),0),
    pendiente:a.reduce((s,r)=>s+money_(r.SALDO),0)
  });

  const metodos={};
  abonos.filter(a=>String(a.FECHA)===hoy).forEach(a=>{
    const k=a.METODO_PAGO||'Sin método';
    metodos[k]=(metodos[k]||0)+money_(a.MONTO);
  });

  return {
    dia:info(ventas.filter(v=>String(v.FECHA)===hoy)),
    semana:info(ventas.filter(v=>{const d=new Date(v.FECHA),h=new Date(hoy);return (h-d)/86400000<7;})),
    mes:info(ventas.filter(v=>String(v.FECHA).slice(0,7)===mes)),
    anio:info(ventas.filter(v=>String(v.FECHA).slice(0,4)===anio)),
    metodos,
    cuentasPorCobrar:ventas.reduce((s,v)=>s+money_(v.SALDO),0),
    recompensasDisponibles:rows_('RECOMPENSAS').filter(r=>String(r.ESTADO)==='Recompensa disponible').length,
    garantiasPendientes:rows_('GARANTIAS').filter(r=>String(r.ESTADO)!=='Resuelta').length,
    proximasVisitas:visitas().slice(0,10)
  };
}

/* ---------- WHATSAPP ---------- */

function wa_(numero,texto){
  let n=String(numero||'').replace(/\D/g,'');
  if(n.length===10)n='52'+n;
  return 'https://wa.me/'+n+'?text='+encodeURIComponent(texto||'');
}

function waEntrega(idVenta){
  const v=obtenerVenta(idVenta),c=v.cliente;
  return wa_(c.WHATSAPP,'Hola '+c.NOMBRE+', te informamos que tu pedido '+idVenta+' ya está listo para entregar. ¡Gracias por tu preferencia en Boutique de Lentes Yael!');
}

function waSaldo(idVenta){
  const v=obtenerVenta(idVenta),c=v.cliente;
  return wa_(c.WHATSAPP,'Hola '+c.NOMBRE+', te recordamos que tu venta '+idVenta+' tiene un saldo pendiente de $'+money_(v.SALDO).toFixed(2)+'. Gracias por tu preferencia en Boutique de Lentes Yael.');
}

function waComprobante(idAbono){
  const a=findOne_('ABONOS','ID_ABONO',idAbono);
  if(!a)throw new Error('Abono no encontrado.');
  const c=obtenerCliente(a.ID_CLIENTE);
  return wa_(c.WHATSAPP,'Hola '+c.NOMBRE+', te compartimos el comprobante del abono '+idAbono+' correspondiente a la venta '+a.ID_VENTA+'.');
}

function waCotizacion(idCotizacion){
  const q=buscarCotizacion(idCotizacion),c=obtenerCliente(q.ID_CLIENTE);
  return wa_(c.WHATSAPP,'Hola '+c.NOMBRE+', te compartimos la cotización '+idCotizacion+' de Boutique de Lentes Yael. Tiene vigencia de 15 días.');
}

/* ---------- DOCUMENTOS ---------- */

function documentoVenta(idVenta){return obtenerVenta(idVenta);}

function documentoAbono(idAbono){
  const a=findOne_('ABONOS','ID_ABONO',idAbono);
  if(!a)throw new Error('Abono no encontrado.');
  a.cliente=obtenerCliente(a.ID_CLIENTE);a.venta=obtenerVenta(a.ID_VENTA);return a;
}

function documentoCotizacion(id){
  const q=buscarCotizacion(id);q.cliente=obtenerCliente(q.ID_CLIENTE);return q;
}

/* ---------- API ---------- */

function api(action,payload){
  payload=payload||{};
  switch(action){
    case 'config':return getConfig();
    case 'buscarClientes':return buscarClientes(payload.termino);
    case 'cliente':return obtenerCliente(payload.id);
    case 'prepararCliente':return prepararCliente(payload.idCliente);
    case 'crearCliente':return crearCliente(payload);
    case 'editarCliente':return editarCliente(payload);
    case 'historia':return historiaCliente(payload.idCliente);
    case 'guardarHistoria':return guardarHistoria(payload);
    case 'ultimaLensometria':return ultimaLensometria(payload.idCliente);
    case 'guardarLensometria':return guardarLensometria(payload);
    case 'graduaciones':return graduacionesCliente(payload.idCliente);
    case 'ultimaGraduacion':return ultimaGraduacion(payload.idCliente);
    case 'guardarGraduacion':return guardarGraduacion(payload);
    case 'crearCotizacion':return crearCotizacion(payload);
    case 'buscarCotizacion':return buscarCotizacion(payload.id);
    case 'crearVenta':return crearVenta(payload);
    case 'buscarVentas':return buscarVentas(payload.termino);
    case 'venta':return obtenerVenta(payload.id);
    case 'registrarAbono':return registrarAbono(payload);
    case 'buscarAbonos':return buscarAbonos(payload.termino);
    case 'entregas':return entregas(payload);
    case 'cambiarEntrega':return cambiarEntrega(payload);
    case 'recompensa':return obtenerRecompensa(payload.idCliente);
    case 'buscarRecompensas':return buscarRecompensas(payload.termino);
    case 'redimirRecompensa':return redimirRecompensa(payload);
    case 'garantias':return garantias(payload.termino);
    case 'crearGarantia':return crearGarantia(payload);
    case 'actualizarGarantia':return actualizarGarantia(payload);
    case 'visitas':return visitas();
    case 'actualizarVisita':return actualizarVisita(payload);
    case 'resumen':return resumen();
    case 'waEntrega':return waEntrega(payload.idVenta);
    case 'waSaldo':return waSaldo(payload.idVenta);
    case 'waComprobante':return waComprobante(payload.idAbono);
    case 'waCotizacion':return waCotizacion(payload.idCotizacion);
    case 'documentoVenta':return documentoVenta(payload.idVenta);
    case 'documentoAbono':return documentoAbono(payload.idAbono);
    case 'documentoCotizacion':return documentoCotizacion(payload.id);
    default:throw new Error('Acción no reconocida: '+action);
  }
}

function pruebaConexion(){
  const s=ss_();
  return {
    ok:true,nombre:s.getName(),
    hojas:APP.SHEETS.map(n=>({hoja:n,existe:!!s.getSheetByName(n)}))
  };
}
