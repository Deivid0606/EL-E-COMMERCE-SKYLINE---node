/**************************************************
 * EL E-COMMERCE SKYLINE — Backend (Code.gs)
 * DB: Google Sheets (fijo por ID)
 * Roles: ADMIN | VENDEDOR | DELIVERY | DESPACHANTE
 **************************************************/


const CONFIG = {
  DB_ID: '1rSTlw8tF50JZjnehjv7O7NjuxYMsYO7qEIjMDNCsaW4', // <-- tu spreadsheet fijo
  SHEETS: {
    USERS: 'USERS',
    PRODUCTS: 'PRODUCTS',
    ORDERS: 'ORDERS',
    SESSIONS: 'SESSIONS',
    WALLETS: 'WALLETS',
    WALLET_TXS: 'WALLET_TXS',
    DELIVERY_FEES: 'DELIVERY_FEES',     // tarifas por repartidor+ciudad (ADMIN)
    CLIENT_PRICES: 'CLIENT_PRICES',      // precios al cliente por ciudad (ADMIN)
    NEWS: 'NEWS',                        // feed de novedades
    // ===== NUEVO: tokens de reseteo de contraseña =====
    RESET_TOKENS: 'RESET_TOKENS',
    // ===== NUEVO: hoja de perfiles =====
    PROFILES: 'PROFILES'
  },
  SESSION_HOURS: 24,
  ORDER_SEQ_PROP: 'ORDER_SEQ_COUNTER',
  ORDER_SEQ_PREFIX: 'A',
  ORDER_SEQ_PAD: 3,
  GUIDES_FOLDER_NAME: 'SKYLINE_GUIDES',
  // ===== (opcional) URL WebApp; si no está, se usa ScriptApp.getService().getUrl() =====
  // WEBAPP_URL: 'https://script.google.com/.../exec',
  // ===== NUEVO: duración de token de reset (minutos) =====
  RESET_TOKEN_EXP_MIN: 60
};


// ======== Precios de delivery al cliente (semilla por defecto) ========
const CLIENT_CITY_PRICES = [
  // CENTRAL
  ['Asuncion',35000],['Fernando de la Mora',35000],['Lambare',35000],['Limpio',35000],
  ['Luque',35000],['Mariano Roque Alonso',35000],['Ñemby',35000],['Villa Elisa',35000],
  ['San Lorenzo',35000],['Ypane',45000],['San Antonio',45000],['Aregua',45000],
  ['Capiata',45000],['Guarambare',45000],['Ita',45000],['Itaugua',45000],
  ['J. Augusto Saldívar',45000],['Villeta',45000],['Nueva Italia',45000],
  // CORDILLERA
  ['Caacupe',50000],['Atyrá',60000],['Altos',60000],['Emboscada',60000],['Eusebio Ayala',60000],
  ['Itacurubí de la Cordillera',60000],['Loma Grande',60000],['Piribebuy',60000],
  ['San Bernardino',60000],['Tobatí',60000],['Ypacaraí',60000],['YAGUARON',60000],
  ['Carapegua',60000],['Paraguarí',60000],
  // OTROS
  ['Benjamín Aceval',60000],['Villa Hayes',50000],['Remansito',50000],
  // ALTO PARANÁ
  ['Minga Guazu',50000],['Colonia Yguazu',50000],['Hernandarias',50000],
  ['Puerto Pdte. Franco',45000],['San Alberto',50000],['SANTA RITA',50000],
  ['Juan leon malloriquin',50000],['YAGUAZU',50000],['Ciudad del este',45000],
  // GUAIRA
  ['Villarrica',45000],['Caaguazu',45000]
];


// ======== Utils ========
function _ss(){ return SpreadsheetApp.openById(CONFIG.DB_ID); }
function _ensureSheets(){
  const ss=_ss();
  const names=Object.values(CONFIG.SHEETS);
  names.forEach(name=>{
    let sh=ss.getSheetByName(name);
    if (!sh) sh=ss.insertSheet(name);
    if (sh.getLastRow()===0){
      if (name===CONFIG.SHEETS.USERS)      sh.appendRow(['id','name','email','password_hash','created_at','role']);
      if (name===CONFIG.SHEETS.PRODUCTS)   sh.appendRow(['id','title','sku','provider_price_gs','stock','image_url','created_at','updated_at','private_to_emails']);
      if (name===CONFIG.SHEETS.ORDERS)     sh.appendRow([
        'id','created_by','customer_name','phone','city','street','district','email',
        'items_json','total_gs','delivery_gs','commission_gs','status','obs','assigned_delivery',
        'created_at','commission_credited','commission_paid','paid_at',
        'delivery_fee_gs','delivery_fee_credited','delivery_settled','delivery_paid_at','status2',
        'assigned_at'
      ]);
      if (name===CONFIG.SHEETS.SESSIONS)   sh.appendRow(['token','email','expires_at']);
      if (name===CONFIG.SHEETS.WALLETS)    sh.appendRow(['email','balance_gs','updated_at']);
      if (name===CONFIG.SHEETS.WALLET_TXS) sh.appendRow(['id','type','email','order_id','amount_gs','note','created_at']);
      if (name===CONFIG.SHEETS.DELIVERY_FEES) sh.appendRow(['id','delivery_email','city','fee_gs','updated_at']);
      if (name===CONFIG.SHEETS.CLIENT_PRICES) sh.appendRow(['city','price_gs','updated_at']);
      if (name===CONFIG.SHEETS.NEWS)       sh.appendRow(['id','ts_iso','order_id','actor_email','role_scope','target_email','message']);
      // ===== NUEVO: Hoja de tokens de reset =====
      if (name===CONFIG.SHEETS.RESET_TOKENS) sh.appendRow(['token','email','expires_at','used_at','created_at']);
      // ===== NUEVO: Hoja de perfiles =====
      if (name===CONFIG.SHEETS.PROFILES) sh.appendRow([
        'email','name','phone','doc','addr',
        'bank_name','bank_type','bank_num','bank_holder','bank_holder_ci',
        'wallet_provider','wallet_number','wallet_holder',
        'updated_at'
      ]);
    }
    // upgrades mínimos
    if (name===CONFIG.SHEETS.ORDERS){
      const shO=ss.getSheetByName(name);
      const head=shO.getRange(1,1,1,shO.getLastColumn()).getValues()[0].map(s=>String(s||'').toLowerCase());
      function addCol(l){ shO.getRange(1,shO.getLastColumn()+1).setValue(l); }
      ['street','district','delivery_fee_gs','delivery_fee_credited','delivery_settled','delivery_paid_at','status2','assigned_at'].forEach(k=>{
        if (head.indexOf(k)<0) addCol(k);
      });
    }
    if (name===CONFIG.SHEETS.PRODUCTS){
      const shP=ss.getSheetByName(name);
      const head=shP.getRange(1,1,1,shP.getLastColumn()).getValues()[0].map(s=>String(s||'').toLowerCase());
      if (head.indexOf('private_to_emails')<0){
        shP.getRange(1, shP.getLastColumn()+1).setValue('private_to_emails');
      }
    }
  });
}
function _sheet(name){ _ensureSheets(); return _ss().getSheetByName(name); }
function _nowIso(){ return new Date().toISOString(); }
function _uuid(){ return Utilities.getUuid(); }
function _zeroPad(n,p){ return String(n).padStart(p,'0'); }
function _rows(sh,cols){ const last=sh.getLastRow(); if (last<2) return []; return sh.getRange(2,1,last-1,cols).getValues(); }
function _hash(str){ const raw=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,str); return raw.map(b=>(b<0?b+256:b).toString(16).padStart(2,'0')).join(''); }
function _norm(s){ return String(s||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim(); }
function _matchQuery(q, fields){ q=_norm(q||''); if(!q) return true; const h=_norm(fields.join(' ')); return q.split(/\s+/).every(t=>h.indexOf(t)>=0); }


// ======== Sesión / Usuarios ========
// (actualizado) ahora acepta horas personalizadas para "mantener sesión"
function _newSession(email, hoursOpt){
  const hours = Number(hoursOpt||CONFIG.SESSION_HOURS);
  const token=_uuid();
  const exp=new Date(Date.now()+hours*3600*1000).toISOString();
  _sheet(CONFIG.SHEETS.SESSIONS).appendRow([token,email,exp]);
  return {token,email,expires_at:exp};
}
function _sessionEmail(token){
  if (!token) return null;
  const sh=_sheet(CONFIG.SHEETS.SESSIONS), rows=_rows(sh,3);
  for (let i=0;i<rows.length;i++){
    const [t,email,exp]=rows[i];
    if (t===token){
      if (new Date(exp).getTime()>Date.now()) return email;
      sh.deleteRow(i+2); return null;
    }
  }
  return null;
}
function _getUserByEmail(email){
  const sh=_sheet(CONFIG.SHEETS.USERS), rows=_rows(sh,6);
  const r=rows.find(x=>String(x[2]||'').toLowerCase()===String(email).toLowerCase());
  if (!r) return null; return {id:r[0],name:r[1],email:r[2],passhash:r[3],created_at:r[4],role:String(r[5]||'VENDEDOR').toUpperCase()};
}
function _userByEmail(email){ return _getUserByEmail(email) || {email,role:'VENDEDOR'}; }


function register(name,email,password,role){
  email=String(email||'').trim().toLowerCase(); if (!name||!email||!password) throw new Error('Faltan campos');
  const sh=_sheet(CONFIG.SHEETS.USERS), data=_rows(sh,6);
  if (data.some(r=>String(r[2]||'').toLowerCase()===email)) throw new Error('El email ya está registrado');
  let finalRole=String(role||'VENDEDOR').toUpperCase(); if (data.length===0) finalRole='ADMIN';
  if (!['ADMIN','VENDEDOR','DELIVERY','DESPACHANTE'].includes(finalRole)) finalRole='VENDEDOR';
  const id=_uuid(); sh.appendRow([id,name,email,_hash(password),_nowIso(),finalRole]);
  return {ok:true,user:{id,name,email,role:finalRole}};
}
// (actualizado) ahora recibe "remember" para sesiones largas
function login(email,password,remember){
  email=String(email||'').trim().toLowerCase(); const u=_getUserByEmail(email);
  if (!u) throw new Error('Usuario no encontrado'); if (_hash(password)!==u.passhash) throw new Error('Contraseña incorrecta');
  // si remember=true => 30 días; si no, la duración por defecto de CONFIG.SESSION_HOURS
  const hours = remember ? (24*30) : CONFIG.SESSION_HOURS;
  const session=_newSession(email, hours);
  return {ok:true,user:{id:u.id,name:u.name,email:u.email,role:u.role},session};
}
function me(token){ const email=_sessionEmail(token); if (!email) return {authenticated:false}; const u=_getUserByEmail(email); if (!u) return {authenticated:false}; return {authenticated:true,user:{id:u.id,name:u.name,email:u.email,role:u.role}}; }
function listUsersByRole(token, role){
  const email=_sessionEmail(token); if (!email) throw new Error('No autenticado');
  const me=_getUserByEmail(email); if (me.role!=='ADMIN') throw new Error('Solo ADMIN');
  const rows=_rows(_sheet(CONFIG.SHEETS.USERS),6);
  return rows.filter(r=>String(r[5]||'')===role).map(r=>({id:r[0],name:r[1],email:r[2],role:r[5]}));
}


// ======== Wallet ========
// (Emails normalizados para evitar duplicados y saldos inconsistentes)
function _getOrCreateWallet(email){
  email = String(email||'').trim().toLowerCase();
  const sh=_sheet(CONFIG.SHEETS.WALLETS), last=sh.getLastRow();
  if (last>=2){
    const data=sh.getRange(2,1,last-1,3).getValues();
    for (let i=0;i<data.length;i++){
      if (String(data[i][0]||'').trim().toLowerCase()===email){
        return {row:i+2,balance:Number(data[i][1]||0)};
      }
    }
  }
  sh.appendRow([email,0,_nowIso()]);
  return {row:sh.getLastRow(), balance:0};
}
function _walletTx(type,email,orderId,amount,note){
  email = String(email||'').trim().toLowerCase();
  _sheet(CONFIG.SHEETS.WALLET_TXS).appendRow([_uuid(),type,email,orderId,Number(amount||0),note||'',_nowIso()]);
}
function _walletCredit(email,amount,orderId,note){
  if (!amount) return; email=String(email||'').trim().toLowerCase();
  const w=_getOrCreateWallet(email), sh=_sheet(CONFIG.SHEETS.WALLETS);
  const newBal=Number(w.balance||0)+Number(amount||0);
  sh.getRange(w.row,2).setValue(newBal); sh.getRange(w.row,3).setValue(_nowIso());
  _walletTx('CREDIT',email,orderId,amount,note||'Crédito');
}
function _walletDebit(email,amount,orderId,note){
  if (!amount) return; email=String(email||'').trim().toLowerCase();
  const w=_getOrCreateWallet(email), sh=_sheet(CONFIG.SHEETS.WALLETS);
  const newBal=Number(w.balance||0)-Number(amount||0);
  sh.getRange(w.row,2).setValue(newBal); sh.getRange(w.row,3).setValue(_nowIso());
  _walletTx('DEBIT',email,orderId,amount,note||'Débito');
}
function getWallet(token, emailOpt){
  const email=_sessionEmail(token); if (!email) throw new Error('No autenticado');
  const me=_userByEmail(email);
  const who=(me.role==='ADMIN'&&emailOpt)? emailOpt : email;
  const w=_getOrCreateWallet(who); const txSh=_sheet(CONFIG.SHEETS.WALLET_TXS);
  let txs=[]; if (txSh.getLastRow()>=2){
    const data=txSh.getRange(2,1,txSh.getLastRow()-1,7).getValues().reverse();
    for (const r of data){
      if (String(r[2]||'').trim().toLowerCase()===String(who).trim().toLowerCase()){
        txs.push({id:r[0],type:r[1],email:r[2],order_id:r[3],amount_gs:r[4],note:r[5],created_at:r[6]});
        if (txs.length>=50) break;
      }
    }
  }
  return {email:who,balance_gs:w.balance,txs};
}


// ======== Índices ORDERS ========
function _orderIndexMap(){
  const sh=_sheet(CONFIG.SHEETS.ORDERS);
  const head=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(s=>String(s||'').toLowerCase());
  const f=n=>head.indexOf(n);
  return {
    id:f('id'), created_by:f('created_by'), customer_name:f('customer_name'),
    phone:f('phone'), city:f('city'), street:f('street'), district:f('district'),
    email:f('email'), items_json:f('items_json'), total_gs:f('total_gs'),
    delivery_gs:f('delivery_gs'), commission_gs:f('commission_gs'), status:f('status'),
    obs:f('obs'), assigned_delivery:f('assigned_delivery'), created_at:f('created_at'),
    commission_credited:f('commission_credited'), commission_paid:f('commission_paid'),
    paid_at:f('paid_at'),
    delivery_fee_gs:f('delivery_fee_gs'),
    delivery_fee_credited:f('delivery_fee_credited'),
    delivery_settled:f('delivery_settled'),
    delivery_paid_at:f('delivery_paid_at'),
    status2:f('status2'),
    assigned_at:f('assigned_at')
  };
}


// ======== Productos ========
function listProducts(token){
  const email=_sessionEmail(token); if (!email) throw new Error('No autenticado');
  const me=_userByEmail(email);
  const rows=_rows(_sheet(CONFIG.SHEETS.PRODUCTS),9);
  return rows.filter(r=>{
    const priv = String(r[8]||'').trim();
    if (!priv) return true; // público
    const allowed = priv.split(',').map(x=>x.trim().toLowerCase()).filter(Boolean);
    return allowed.includes(String(me.email||'').toLowerCase());
  }).map(r=>({
    id:r[0], title:r[1], sku:r[2], provider_price_gs:Number(r[3]||0), stock:Number(r[4]||0),
    image_url:r[5]||'', created_at:r[6], updated_at:r[7]
  }));
}
function addProduct(token,p){
  const email=_sessionEmail(token); if (!email) throw new Error('No autenticado');
  const me=_getUserByEmail(email); if (me.role!=='ADMIN') throw new Error('Solo ADMIN');
  const sh=_sheet(CONFIG.SHEETS.PRODUCTS);
  const id=_uuid(), now=_nowIso();
  const priv=Array.isArray(p.private_emails)? p.private_emails.join(',') : (p.private_emails||'');
  sh.appendRow([id, p.title||'', p.sku||'', Number(p.provider_price_gs||0), Number(p.stock||0), p.image_url||'', now, now, priv]);
  return {ok:true,id};
}


// ======== Delivery: tarifas por repartidor+ciudad (ADMIN) ========
function setDeliveryRate(token, deliveryEmail, city, feeGs){
  const email=_sessionEmail(token); if (!email) throw new Error('No autenticado');
  const me=_getUserByEmail(email); if (me.role!=='ADMIN') throw new Error('Solo ADMIN');
  deliveryEmail=String(deliveryEmail||'').trim().toLowerCase(); city=String(city||'').trim();
  const sh=_sheet(CONFIG.SHEETS.DELIVERY_FEES); const last=sh.getLastRow();
  if (last>=2){
    const data=sh.getRange(2,1,last-1,5).getValues();
    for (let i=0;i<data.length;i++){
      const r=data[i];
      if (String(r[1]||'').toLowerCase()===deliveryEmail && String(r[2]||'')===city){
        sh.getRange(i+2,4).setValue(Number(feeGs||0));
        sh.getRange(i+2,5).setValue(_nowIso());
        return {ok:true,updated:true};
      }
    }
  }
  sh.appendRow([_uuid(),deliveryEmail,city,Number(feeGs||0),_nowIso()]);
  return {ok:true,created:true};
}
function getDeliveryRates(token, emailOpt){
  const email=_sessionEmail(token); if (!email) throw new Error('No autenticado');
  const me=_getUserByEmail(email);
  const sh=_sheet(CONFIG.SHEETS.DELIVERY_FEES); const last=sh.getLastRow(); if (last<2) return [];
  const rows=sh.getRange(2,1,last-1,5).getValues();
  if (me.role==='ADMIN'){
    return rows.filter(r=>!emailOpt || String(r[1]||'').toLowerCase()===String(emailOpt||'').toLowerCase())
               .map(r=>({email:r[1],city:r[2],rate_gs:Number(r[3]||0)}));
  }
  if (me.role==='DELIVERY'){
    return rows.filter(r=>String(r[1]||'').toLowerCase()===String(me.email||'').toLowerCase())
               .map(r=>({email:r[1],city:r[2],rate_gs:Number(r[3]||0)}));
  }
  // DESPACHANTE / VENDEDOR: sin acceso (frontend ya lo oculta)
  return [];
}
function _lookupDeliveryFee(deliveryEmail, city){
  const sh=_sheet(CONFIG.SHEETS.DELIVERY_FEES), last=sh.getLastRow(); if (last<2) return 0;
  const data=sh.getRange(2,1,last-1,5).getValues();
  const de=_norm(deliveryEmail), c=_norm(city);
  const row=data.find(r=>_norm(r[1])===de && _norm(r[2])===c);
  return row? Number(row[3]||0):0;
}


// ======== City prices (cliente) ========
function getDeliveryClientPrices(token){
  const email=_sessionEmail(token); if (!email) throw new Error('No autenticado');
  const sh=_sheet(CONFIG.SHEETS.CLIENT_PRICES);
  const last=sh.getLastRow();
  if (last>=2){
    const data=sh.getRange(2,1,last-1,3).getValues();
    return data.map(r=>({city:r[0],price:Number(r[1]||0)}));
  }
  return CLIENT_CITY_PRICES.map(([city,price])=>({city,price}));
}
function setClientCityPrice(token, city, price){
  const email=_sessionEmail(token); if (!email) throw new Error('No autenticado');
  const me=_getUserByEmail(email); if (me.role!=='ADMIN') throw new Error('Solo ADMIN');
  city=String(city||'').trim(); const sh=_sheet(CONFIG.SHEETS.CLIENT_PRICES);
  const last=sh.getLastRow();
  if (last>=2){
    const data=sh.getRange(2,1,last-1,3).getValues();
    for (let i=0;i<data.length;i++){
      if (String(data[i][0]||'')===city){
        sh.getRange(i+2,2).setValue(Number(price||0));
        sh.getRange(i+2,3).setValue(_nowIso());
        return {ok:true,updated:true};
      }
    }
  }
  sh.appendRow([city, Number(price||0), _nowIso()]);
  return {ok:true,created:true};
}
function _clientCityPrice(city){
  if (!city) return 0;
  const sh=_sheet(CONFIG.SHEETS.CLIENT_PRICES); const last=sh.getLastRow();
  if (last>=2){
    const data=sh.getRange(2,1,last-1,3).getValues();
    const row=data.find(r=>_norm(r[0])===_norm(city));
    if (row) return Number(row[1]||0);
  }
  const canon=_norm(city);
  const row=CLIENT_CITY_PRICES.find(([c])=>_norm(c)===canon); return row? Number(row[1]||0):0;
}


// ======== Órdenes ========
function _nextOrderCode(){
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try{
    const props=PropertiesService.getScriptProperties();
    let n=Number(props.getProperty(CONFIG.ORDER_SEQ_PROP)||'0'); if (isNaN(n)) n=0; n+=1;
    props.setProperty(CONFIG.ORDER_SEQ_PROP,String(n));
    return CONFIG.ORDER_SEQ_PREFIX + _zeroPad(n, CONFIG.ORDER_SEQ_PAD);
  } finally { lock.releaseLock(); }
}


function addOrder(token, order){
  const email=_sessionEmail(token); if (!email) throw new Error('No autenticado');


  // Cargar catálogo para costo proveedor (por sku)
  const pr=_rows(_sheet(CONFIG.SHEETS.PRODUCTS),9);
  const mapProv={}; pr.forEach(r=>{ const sku=String(r[2]||'').trim(); if (sku) mapProv[sku]=Number(r[3]||0); });


  // Items vienen con sale_gs = TOTAL por renglón; qty para multiplicar proveedor
  let sumSale=0, sumProv=0;
  const itemsNorm=(order.items||[]).map(it=>{
    const sku=String(it.sku||'').trim(); const qty=Number(it.qty||0); const sale=Number(it.sale_gs||0);
    const provUnit=Number(mapProv[sku]||0); const provTotal=provUnit*qty;
    if (qty>0 && sale>0){ sumSale+=sale; sumProv+=provTotal; }
    return {sku,qty,sale_gs:sale,provider_gs:provUnit};
  });


  const city=String(order.city||'').trim();
  const delivery_gs=_clientCityPrice(city);
  const commission = sumSale - (sumProv + delivery_gs);


  const id=_nextOrderCode();
  _sheet(CONFIG.SHEETS.ORDERS).appendRow([
    id, email, order.customer_name||'', order.phone||'', city, order.street||'', order.district||'', (order.email||''),
    JSON.stringify(itemsNorm), sumSale, delivery_gs, commission,
    'PENDIENTE', String(order.obs||''), '', _nowIso(), false, false, '',
    0, false, false, '', '', // status2 vacío + campos delivery settle
    ''                      // assigned_at (vacío al crear)
  ]);


  return {ok:true,id};
}


// listado con filtros + búsqueda
function listOrdersFiltered(token, fromISO, toISO, q, vendorEmailOpt, deliveryEmailOpt){
  const email=_sessionEmail(token); if (!email) throw new Error('No autenticado');
  const me=_userByEmail(email), sh=_sheet(CONFIG.SHEETS.ORDERS), idx=_orderIndexMap();
  const last=sh.getLastRow(); if (last<2) return [];
  const rg=sh.getRange(2,1,last-1, sh.getLastColumn()).getValues();


  const from=fromISO? new Date(fromISO+'T00:00:00'):new Date('1970-01-01');
  const to  =toISO  ? new Date(toISO  +'T23:59:59'):new Date('2999-12-31');


  const out=[];
  for (const r of rg){
    const createdBy=r[idx.created_by]||'', assigned=r[idx.assigned_delivery]||'';
    if (me.role==='VENDEDOR' && createdBy.toLowerCase()!==email.toLowerCase()) continue;
    if (me.role==='DELIVERY' && assigned.toLowerCase()!==email.toLowerCase()) continue;
    // DESPACHANTE y ADMIN ven todo


    if (vendorEmailOpt && _norm(createdBy)!==_norm(vendorEmailOpt)) continue;
    if (deliveryEmailOpt && _norm(assigned)!==_norm(deliveryEmailOpt)) continue;


    const createdAt=new Date(r[idx.created_at]); if (createdAt<from || createdAt>to) continue;


    if(!_matchQuery(q||'', [r[idx.id], r[idx.customer_name], r[idx.phone], r[idx.email], r[idx.city], r[idx.created_by], r[idx.assigned_delivery]])) continue;


    out.push({
      id:r[idx.id], created_by:r[idx.created_by],
      customer_name:r[idx.customer_name], phone:r[idx.phone], city:r[idx.city],
      street:r[idx.street]||'', district:r[idx.district]||'',
      email:r[idx.email], items_json:r[idx.items_json],
      total_gs:Number(r[idx.total_gs]||0), delivery_gs:Number(r[idx.delivery_gs]||0),
      commission_gs:Number(r[idx.commission_gs]||0), status:r[idx.status]||'PENDIENTE',
      assigned_delivery:r[idx.assigned_delivery]||'',
      created_at:r[idx.created_at],
      delivery_fee_gs:Number(r[idx.delivery_fee_gs]||0),
      delivery_fee_credited: !!r[idx.delivery_fee_credited],
      delivery_settled: !!r[idx.delivery_settled],
      status2:r[idx.status2]||'',
      assigned_at: r[idx.assigned_at] || ''
    });
  }
  return out;
}


// Mantengo listOrders para compatibilidad
function listOrders(token){
  return listOrdersFiltered(token,'','', '', '', '');
}


function assignDelivery(token, id, deliveryEmail){
  const email=_sessionEmail(token); if (!email) throw new Error('No autenticado');
  const me=_getUserByEmail(email); if (me.role!=='ADMIN') throw new Error('Solo ADMIN');
  const sh=_sheet(CONFIG.SHEETS.ORDERS), idx=_orderIndexMap();
  const last=sh.getLastRow(); if (last<2) throw new Error('Sin pedidos');
  const rg=sh.getRange(2,1,last-1, sh.getLastColumn()), data=rg.getValues();


  if (deliveryEmail){ const di=_getUserByEmail(deliveryEmail||''); if (!di || di.role!=='DELIVERY') throw new Error('Email no es DELIVERY'); }


  for (let i=0;i<data.length;i++){
    const r=data[i]; if (String(r[idx.id])===String(id)){
      rg.getCell(i+1, idx.assigned_delivery+1).setValue(String(deliveryEmail||'').trim());
      if (String(deliveryEmail||'').trim()){
        rg.getCell(i+1, idx.assigned_at+1).setValue(_nowIso()); // marca fecha/hora de asignación
      } else {
        rg.getCell(i+1, idx.assigned_at+1).setValue(''); // limpia al desasignar
      }
      return {ok:true};
    }
  }
  throw new Error('Pedido no encontrado');
}


// ======== Estados ========
// Estado 1: PENDIENTE / EN RUTA / ENTREGADO / CANCELADO / REAGENDADO / NO CONTESTA / RECHAZADO / NO DESEA / CANCELÓ POR WHATSAPP / DEVUELTO A DEPÓSITO / RECHAZADO EN EL LUGAR
// + NUEVO: ENCOMIENDA ENTREGADA
function updateOrderStatus(token, id, status, _obsIgnored){
  const email=_sessionEmail(token); if (!email) throw new Error('No autenticado');
  const me=_getUserByEmail(email);
  // VENDEDOR y DESPACHANTE no pueden tocar Estado 1
  if (me.role==='VENDEDOR' || me.role==='DESPACHANTE') throw new Error('Sin permiso para cambiar Estado 1');


  const sh=_sheet(CONFIG.SHEETS.ORDERS), idx=_orderIndexMap();
  const last=sh.getLastRow(); if (last<2) throw new Error('Sin pedidos');
  const rg=sh.getRange(2,1,last-1, sh.getLastColumn()), data=rg.getValues();


  const notifySet=new Set(['REAGENDADO','NO CONTESTA','RECHAZADO','NO DESEA','CANCELÓ POR WHATSAPP','CANCELADO']);


  for (let i=0;i<data.length;i++){
    const r=data[i];
    if (String(r[idx.id])===String(id)){
      const prevStatus=String(r[idx.status]||'').toUpperCase();
      const newStatus=String(status||'').toUpperCase() || 'PENDIENTE';
      if (me.role==='DELIVERY' && newStatus==='DEVUELTO A DEPÓSITO') throw new Error('Delivery no puede usar este estado');


      // recomputar bases
      const total=Number(r[idx.total_gs]||0);
      const deliveryCharged=Number(r[idx.delivery_gs]||0);


      // recomputar proveedor por seguridad
      const mapProv={}; const pr=_rows(_sheet(CONFIG.SHEETS.PRODUCTS),9); pr.forEach(row=>{ const sku=String(row[2]||'').trim(); if (sku) mapProv[sku]=Number(row[3]||0); });
      let prov=0; try{ JSON.parse(r[idx.items_json]||'[]').forEach(it=>{ prov += (Number(mapProv[String(it.sku||'').trim()]||0) * Number(it.qty||0)); }); }catch(e){}
      const commission = total - (prov + deliveryCharged);


      const vendorEmail=String(r[idx.created_by]||'').trim();
      const assigned=String(r[idx.assigned_delivery]||'').trim();


      const alreadyCredVend=!!r[idx.commission_credited];
      const alreadyCredDel =!!r[idx.delivery_fee_credited];


      // ===== reversos/cargos cuando se deja "ENTREGADO"
      if (prevStatus==='ENTREGADO' && newStatus!=='ENTREGADO'){
        if (alreadyCredVend && vendorEmail && commission>0){
          _walletDebit(vendorEmail, commission, r[idx.id], 'Reverso comisión por cambio a '+newStatus);
          rg.getCell(i+1, idx.commission_credited+1).setValue(false);
        }
        if (alreadyCredDel && assigned){
          const fee=Number(r[idx.delivery_fee_gs]||0);
          if (fee>0){ _walletDebit(assigned, fee, r[idx.id], 'Reverso tarifa delivery por cambio a '+newStatus); }
          rg.getCell(i+1, idx.delivery_fee_credited+1).setValue(false);
          rg.getCell(i+1, idx.delivery_fee_gs+1).setValue(0);
        }
      }


      // ===== reverso cuando se deja "ENCOMIENDA ENTREGADA"
      if (prevStatus==='ENCOMIENDA ENTREGADA' && newStatus!=='ENCOMIENDA ENTREGADA'){
        if (alreadyCredVend && vendorEmail && commission>0){
          _walletDebit(vendorEmail, commission, r[idx.id], 'Reverso comisión por cambio desde ENCOMIENDA ENTREGADA a '+newStatus);
          rg.getCell(i+1, idx.commission_credited+1).setValue(false);
        }
        // asegurar que no quede tarifa de delivery asentada
        if (alreadyCredDel && assigned){
          const fee=Number(r[idx.delivery_fee_gs]||0);
          if (fee>0){ _walletDebit(assigned, fee, r[idx.id], 'Reverso tarifa delivery (salida de ENCOMIENDA ENTREGADA)'); }
        }
        rg.getCell(i+1, idx.delivery_fee_credited+1).setValue(false);
        rg.getCell(i+1, idx.delivery_fee_gs+1).setValue(0);
      }


      // ===== reversos cuando se deja "RECHAZADO EN EL LUGAR"
      if (prevStatus === 'RECHAZADO EN EL LUGAR' && newStatus !== 'RECHAZADO EN EL LUGAR') {
        const assignedEmail = String(r[idx.assigned_delivery]||'').trim();
        const halfCredited = Number(r[idx.delivery_fee_gs]||0); // aquí guardamos el 50% cuando lo acreditamos
        if (alreadyCredDel && assignedEmail && halfCredited>0){
          _walletDebit(assignedEmail, halfCredited, r[idx.id], 'Reverso 50% por cambio de estado');
          rg.getCell(i+1, idx.delivery_fee_credited+1).setValue(false);
          rg.getCell(i+1, idx.delivery_fee_gs+1).setValue(0);
        }
        const vendorEmailR = String(r[idx.created_by]||'').trim();
        const cityR = r[idx.city]||'';
        const clientFeeR = _clientCityPrice(cityR);
        if (vendorEmailR && clientFeeR>0){
          _walletCredit(vendorEmailR, clientFeeR, r[idx.id], 'Reverso descuento por RECHAZADO EN EL LUGAR');
        }
      }


      // si pasa a ENTREGADO acreditar si aún no se hizo (vendor + delivery)
      if (newStatus==='ENTREGADO'){
        if (!alreadyCredVend && vendorEmail && commission>0){
          _walletCredit(vendorEmail, commission, r[idx.id], 'Comisión por pedido entregado');
          rg.getCell(i+1, idx.commission_gs+1).setValue(commission);
          rg.getCell(i+1, idx.commission_credited+1).setValue(true);
        }
        if (assigned && !alreadyCredDel){
          const fee=_lookupDeliveryFee(assigned, r[idx.city]||''); // tarifa por repartidor+ciudad
          rg.getCell(i+1, idx.delivery_fee_gs+1).setValue(fee);
          if (fee>0){ _walletCredit(assigned, fee, r[idx.id], 'Tarifa de delivery (ENTREGADO)'); rg.getCell(i+1, idx.delivery_fee_credited+1).setValue(true); }
        }
      }


      // NUEVO: si pasa a ENCOMIENDA ENTREGADA -> acreditar SOLO vendedor; delivery NO
      if (newStatus==='ENCOMIENDA ENTREGADA'){
        if (!alreadyCredVend && vendorEmail && commission>0){
          _walletCredit(vendorEmail, commission, r[idx.id], 'Comisión por ENCOMIENDA ENTREGADA');
          rg.getCell(i+1, idx.commission_gs+1).setValue(commission);
          rg.getCell(i+1, idx.commission_credited+1).setValue(true);
        }
        // asegurar que el delivery no reciba nada en este estado
        rg.getCell(i+1, idx.delivery_fee_gs+1).setValue(0);
        rg.getCell(i+1, idx.delivery_fee_credited+1).setValue(false);
      }


      // aplicar nuevo estado
      r[idx.status]=newStatus;
      rg.getCell(i+1, idx.status+1).setValue(newStatus);


      // Novedades
      if (notifySet.has(newStatus)){
        const vendor=r[idx.created_by]||''; const deliv=r[idx.assigned_delivery]||'';
        _addNews(r[idx.id], email, 'ADMIN', '', `Pedido ${r[idx.id]}: ${newStatus}`);
        if (vendor) _addNews(r[idx.id], email, 'USER', vendor, `Pedido ${r[idx.id]}: ${newStatus}`);
        if (deliv)  _addNews(r[idx.id], email, 'USER', deliv,  `Pedido ${r[idx.id]}: ${newStatus}`);
      }
      return {ok:true};
    }
  }
  throw new Error('Pedido no encontrado');
}


// Estado 2 (ADMIN/DESPACHANTE): GUIA GENERADA / FUERA DE COBERTURA / CANCELADO / REPETIDO / RENDIDO
function updateOrderStatus2(token, id, status2){
  const email=_sessionEmail(token); if (!email) throw new Error('No autenticado');
  const me=_getUserByEmail(email); if (!(me.role==='ADMIN' || me.role==='DESPACHANTE')) throw new Error('Solo ADMIN o DESPACHANTE');


  const sh=_sheet(CONFIG.SHEETS.ORDERS), idx=_orderIndexMap();
  const last=sh.getLastRow(); if (last<2) throw new Error('Sin pedidos');
  const rg=sh.getRange(2,1,last-1, sh.getLastColumn()), data=rg.getValues();


  for (let i=0;i<data.length;i++){
    const r=data[i];
    if (String(r[idx.id])===String(id)){
      const s2=String(status2||'').toUpperCase();
      rg.getCell(i+1, idx.status2+1).setValue(s2);
      if (s2==='RENDIDO'){
        const settled=!!r[idx.delivery_settled];
        if (!settled){
          const assigned=String(r[idx.assigned_delivery]||'').trim();
          const fee=Number(r[idx.delivery_fee_gs]||0);
          if (assigned && fee>0){ _walletDebit(assigned, fee, r[idx.id], 'Rendición de delivery'); }
          rg.getCell(i+1, idx.delivery_settled+1).setValue(true);
          rg.getCell(i+1, idx.delivery_paid_at+1).setValue(_nowIso());
        }
      }
      return {ok:true};
    }
  }
  throw new Error('Pedido no encontrado');
}


// ======== Pago de COMISIONES (ADMIN) ========
function listVendorCommissions(token, fromISO, toISO, vendorEmailOpt, onlyStatus, q){
  const email=_sessionEmail(token); if (!email) throw new Error('No autenticado');
  const me=_getUserByEmail(email); if (me.role!=='ADMIN') throw new Error('Solo ADMIN');


  const sh=_sheet(CONFIG.SHEETS.ORDERS), idx=_orderIndexMap();
  const last=sh.getLastRow(); if (last<2) return [];
  const rg=sh.getRange(2,1,last-1, sh.getLastColumn()).getValues();


  const from=fromISO? new Date(fromISO+'T00:00:00'):new Date('1970-01-01');
  const to  =toISO  ? new Date(toISO  +'T23:59:59'):new Date('2999-12-31');


  const res=[];
  for (const r of rg){
    const createdAt=new Date(r[idx.created_at]); if (createdAt<from || createdAt>to) continue;
    const status=String(r[idx.status]||'').toUpperCase(); if (status!=='ENTREGADO') continue;


    const vendor=String(r[idx.created_by]||'').trim().toLowerCase();
    if (vendorEmailOpt && vendor!==String(vendorEmailOpt).toLowerCase()) continue;


    const paid=!!r[idx.commission_paid];
    if (onlyStatus==='PAGADO'   && !paid) continue;
    if (onlyStatus==='PENDIENTE'&&  paid) continue;


    if(!_matchQuery(q||'', [r[idx.id], r[idx.customer_name], r[idx.phone], r[idx.email], r[idx.city], r[idx.created_by]])) continue;


    res.push({
      id:r[idx.id], created_at:r[idx.created_at], city:r[idx.city]||'',
      customer_name:r[idx.customer_name]||'', vendor_email:r[idx.created_by]||'',
      total_gs:Number(r[idx.total_gs]||0), commission_gs:Number(r[idx.commission_gs]||0),
      commission_credited: !!r[idx.commission_credited],
      commission_paid: paid, paid_at:r[idx.paid_at]||''
    });
  }
  return res;
}
function payVendorCommission(token, orderId, paid){
  const email=_sessionEmail(token); if (!email) throw new Error('No autenticado');
  const me=_getUserByEmail(email); if (me.role!=='ADMIN') throw new Error('Solo ADMIN');


  const sh=_sheet(CONFIG.SHEETS.ORDERS), idx=_orderIndexMap();
  const last=sh.getLastRow(); if (last<2) throw new Error('Sin pedidos');
  const rg=sh.getRange(2,1,last-1, sh.getLastColumn()), data=rg.getValues();


  for (let i=0;i<data.length;i++){
    const r=data[i]; if (String(r[idx.id])!==String(orderId)) continue;


    const credited=!!r[idx.commission_credited];
    const alreadyPaid=!!r[idx.commission_paid];
    const vendorEmail=String(r[idx.created_by]||'').trim().toLowerCase();
    const commission=Number(r[idx.commission_gs]||0);


    if (paid && !alreadyPaid){
      if (credited && vendorEmail && commission>0){
        _walletDebit(vendorEmail, commission, r[idx.id], 'Pago de comisión (ADMIN)');
      }
      rg.getCell(i+1, idx.commission_paid+1).setValue(true);
      rg.getCell(i+1, idx.paid_at+1).setValue(_nowIso());
      return {ok:true,paid:true};
    }
    if (!paid && alreadyPaid){
      rg.getCell(i+1, idx.commission_paid+1).setValue(false);
      rg.getCell(i+1, idx.paid_at+1).setValue('');
      return {ok:true,paid:false};
    }
    return {ok:true,paid:alreadyPaid};
  }
  throw new Error('Pedido no encontrado');
}


// ======== Novedades ========
function _addNews(orderId, actorEmail, roleScope, targetEmail, message){
  _sheet(CONFIG.SHEETS.NEWS).appendRow([_uuid(), _nowIso(), orderId, actorEmail||'', roleScope||'', targetEmail||'', message||'']);
}
function listNews(token){
  const email=_sessionEmail(token); if (!email) throw new Error('No autenticado');
  const me=_userByEmail(email), sh=_sheet(CONFIG.SHEETS.NEWS);
  const last=sh.getLastRow(); if (last<2) return [];
  const rows=sh.getRange(2,1,last-1, sh.getLastColumn()).getValues().reverse();
  if (me.role==='ADMIN' || me.role==='DESPACHANTE')
    return rows.map(r=>({id:r[0],created_at:r[1],order_id:r[2],type:r[4],note:r[6]}));
  return rows.filter(r=> String(r[5]||'').toLowerCase()===String(me.email||'').toLowerCase() || String(r[4]||'')==='ALL')
            .map(r=>({id:r[0],created_at:r[1],order_id:r[2],type:r[4],note:r[6]}));
}


// ======== Guía ========
function getGuideText(token, orderId){
  const email=_sessionEmail(token); if (!email) throw new Error('No autenticado');
  const sh=_sheet(CONFIG.SHEETS.ORDERS), idx=_orderIndexMap(); const last=sh.getLastRow(); if (last<2) throw new Error('Sin pedidos');
  const rows=sh.getRange(2,1,last-1, sh.getLastColumn()).getValues();
  let order=null;
  for (const r of rows){
    if (String(r[idx.id])===String(orderId)){
      const me=_userByEmail(email);
      if (me.role==='VENDEDOR' && String(r[idx.created_by]||'').toLowerCase()!==email.toLowerCase()) throw new Error('Sin permiso');
      if (me.role==='DELIVERY' && String(r[idx.assigned_delivery]||'').toLowerCase()!==email.toLowerCase()) throw new Error('Sin permiso');
      // ADMIN y DESPACHANTE ven todos
      order={
        id:r[idx.id], vendor:r[idx.created_by]||'',
        customer_name:r[idx.customer_name]||'', phone:r[idx.phone]||'',
        city:r[idx.city]||'', street:r[idx.street]||'', district:r[idx.district]||'',
        email:r[idx.email]||'', items_json:r[idx.items_json]||'[]',
        total_gs:Number(r[idx.total_gs]||0), status:r[idx.status]||'PENDIENTE', obs:r[idx.obs]||''
      }; break;
    }
  }
  if (!order) throw new Error('Pedido no encontrado');


  let items=[]; try{ items=JSON.parse(order.items_json||'[]'); }catch(e){}
  const fmt=n=>String(Math.round(Number(n||0))).replace(/\B(?=(\d{3})+(?!\d))/g,'.');


  const lines=[];
  lines.push('EL E-COMMERCE SKYLINE');
  lines.push('----------------------');
  lines.push('Pedido: '+order.id);
  lines.push('Vendedor: '+order.vendor);
  lines.push('Cliente: '+order.customer_name);
  lines.push('Teléfono: '+order.phone);
  lines.push('Ciudad: '+order.city);
  lines.push('Calle: '+order.street);
  lines.push('Barrio: '+order.district);
  lines.push('Email: '+order.email);
  lines.push('Items:');
  items.forEach(it=>{
    const title=(it.title&&String(it.title).trim()) || (it.sku&&String(it.sku).trim()) || 'Item';
    const qty=Number(it.qty||0); const unit=fmt(Number(it.sale_gs||0));
    lines.push(`- ${title} — Cant: ${qty} — Precio: ${unit} Gs`);
  });
  lines.push('Total: '+fmt(order.total_gs)+' Gs');
  if (order.obs) { lines.push(''); lines.push('Observación: '+order.obs); }
  lines.push(''); lines.push('Estado: '+String(order.status||'PENDIENTE').toUpperCase());
  return lines.join('\r\n');
}


function generateGuidePDF(token, orderId){
  const email=_sessionEmail(token); if (!email) throw new Error('No autenticado');
  const text=getGuideText(token,orderId);
  const doc=DocumentApp.create('Guia_'+orderId), body=doc.getBody();
  body.setAttributes({FONT_FAMILY:'Courier New',FONT_SIZE:12});
  body.appendParagraph(text).setFontFamily('Courier New').setFontSize(12);
  doc.saveAndClose();
  const pdf=DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF).setName('Guia_'+orderId+'.pdf');
  const folder=(function(){ const it=DriveApp.getFoldersByName(CONFIG.GUIDES_FOLDER_NAME); return it.hasNext()?it.next():DriveApp.createFolder(CONFIG.GUIDES_FOLDER_NAME);})();
  const file=folder.createFile(pdf); DriveApp.getFileById(doc.getId()).setTrashed(true);
  return {url:file.getUrl(),fileId:file.getId(),name:file.getName()};
}


// ======== Métricas / Dashboard ========
function metrics(token, fromISO, toISO){
  const email=_sessionEmail(token); if (!email) throw new Error('No autenticado');
  const me=_userByEmail(email);
  const shO=_sheet(CONFIG.SHEETS.ORDERS); const last=shO.getLastRow();
  if (last<2) return {cards:{orders:0,sold:0,delivered:0,canceled:0,profit:0,monto_rendir:0,assigned_count:0,entregado_total:0}, series:[], pie:{}, map:[], top:[]};


  const idx=_orderIndexMap(); const rg=shO.getRange(2,1,last-1, shO.getLastColumn()).getValues();
  const from=fromISO? new Date(fromISO+'T00:00:00'):new Date('1970-01-01');
  const to  =toISO  ? new Date(toISO  +'T23:59:59'):new Date('2999-12-31');


  // catálogo costo proveedor
  const pr=_rows(_sheet(CONFIG.SHEETS.PRODUCTS),9); const mapProv={}; pr.forEach(r=>{const sku=String(r[2]||'').trim(); if (sku) mapProv[sku]=Number(r[3]||0);});


  let orders=0,sold=0,delivered=0,canceled=0,profit=0,montoRendir=0, assignedCount=0, entregadoTotal=0;
  const byDay={}, mapCity={}, topProd={}; const pie={};


  // Estados que cuentan como cancelados
  const CANCEL_STATES = new Set(['CANCELADO','RECHAZADO','RECHAZADO EN EL LUGAR','NO DESEA','CANCELÓ POR WHATSAPP']);


  for (const r of rg){
    const createdBy=r[idx.created_by]||'', assigned=r[idx.assigned_delivery]||'';
    if (me.role==='VENDEDOR' && createdBy.toLowerCase()!==email.toLowerCase()) continue;
    if (me.role==='DELIVERY' && assigned.toLowerCase()!==email.toLowerCase()) continue;


    // SOLO DELIVERY usa assigned_at (si falta, cae a created_at). Otros usan created_at.
    const assignedAtStr = r[idx.assigned_at] || '';
    const assignedAt    = assignedAtStr ? new Date(assignedAtStr) : null;
    const createdAt     = r[idx.created_at] ? new Date(r[idx.created_at]) : null;
    const baseDate      = (me.role==='DELIVERY') ? (assignedAt || createdAt) : createdAt;


    if (!baseDate || baseDate<from || baseDate>to) continue;


    // Conteo de asignados en el rango (informativo)
    if (me.role==='DELIVERY' && assignedAt && assignedAt>=from && assignedAt<=to) assignedCount++;


    const status=String(r[idx.status]||'PENDIENTE').toUpperCase();
    const total=Number(r[idx.total_gs]||0), delivery=Number(r[idx.delivery_gs]||0);


    let prov=0;
    try{
      JSON.parse(r[idx.items_json]||'[]').forEach(it=>{
        const sku=String(it.sku||'').trim(); const qty=Number(it.qty||0);
        const provUnit=Number(mapProv[sku]||0); prov+=provUnit*qty;
        const key=(it.title||sku||'Item'); if (!topProd[key]) topProd[key]={qty:0,revenue:0};
        topProd[key].qty += qty; topProd[key].revenue += Number(it.sale_gs||0);
      });
    }catch(e){}


    // KPI pedidos/entregados/cancelados (comunes)
    orders++; if (status==='ENTREGADO') delivered++;
    if (CANCEL_STATES.has(status)) canceled++;


    // >>> SOLD:
    sold += total;
    byDay[_dateYMD(baseDate)]=(byDay[_dateYMD(baseDate)]||0)+total;


    // Mapas y top productos (comunes)
    const city=(r[idx.city]||'SIN CIUDAD').toString();
    mapCity[city]=(mapCity[city]||{qty:0,revenue:0});
    mapCity[city].qty++;
    mapCity[city].revenue+=total;


    // Monto a rendir NETO (solo delivery), sin importar acreditación: ENTREGADO y no rendido
    if (me.role==='DELIVERY' && status==='ENTREGADO' && assigned.toLowerCase()===email.toLowerCase()){
      const settled=!!r[idx.delivery_settled];
      if (!settled){
        const feeStored=Number(r[idx.delivery_fee_gs]||0);
        const fee = feeStored>0 ? feeStored : _lookupDeliveryFee(assigned, r[idx.city]||'');
        montoRendir += (total - fee);
      }
      // Acumular también el TOTAL de ENTREGADO (para el KPI "Suma total ENTREGADO")
      entregadoTotal += total;
    }


    // Profit cards (usuarios no-delivery igual que antes)
    const orderProfitVendor = total - (prov + delivery);
    const orderProfitAdmin  = total - delivery;
    if (me.role==='ADMIN' || me.role==='DESPACHANTE') profit += orderProfitAdmin;
    else if (me.role==='VENDEDOR') profit += orderProfitVendor;
    else {
      // DELIVERY: dejamos profit sumando igual a vendor por compatibilidad visual previa.
      profit += orderProfitVendor;
    }


    // Pie por estado
    pie[status]=(pie[status]||0)+1;
  }


  const series=Object.keys(byDay).sort().map(d=>({date:d,value:byDay[d]}));
  const map=Object.keys(mapCity).map(c=>({city:c, qty:mapCity[c].qty, revenue:mapCity[c].revenue})).sort((a,b)=>b.revenue-a.revenue).slice(0,15);


  // <<< LÍNEA CORREGIDA: sort sin el error de sintaxis >>>
  const top = Object.keys(topProd)
    .map(k => ({ name: k, qty: topProd[k].qty, revenue: topProd[k].revenue }))
    .sort((a, b) => b.qty - a)
    .slice(0, 15);


  return {cards:{orders,sold,delivered,canceled,profit,monto_rendir:montoRendir, assigned_count:assignedCount, entregado_total:entregadoTotal}, series, pie, map, top};
}
function _dateYMD(d){ return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }


// ======== HTTP ========
function doGet(){
  _ensureSheets();
  const tpl=HtmlService.createTemplateFromFile('Index');
  return tpl.evaluate().setTitle('EL E-COMMERCE SKYLINE')
    .setFaviconUrl('https://ssl.gstatic.com/docs/doclist/images/drive_2022q3_32dp.png')
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}
function include(name){ return HtmlService.createHtmlOutputFromFile(name).getContent(); }


// =====================================================================
// ========================= NUEVO: ADMIN COUNTER =======================
// =====================================================================
function adminGetOrderCounter(token){
  const email=_sessionEmail(token); if (!email) throw new Error('No autenticado');
  const me=_getUserByEmail(email); if (me.role!=='ADMIN') throw new Error('Solo ADMIN');
  const props = PropertiesService.getScriptProperties();
  const val = Number(props.getProperty(CONFIG.ORDER_SEQ_PROP) || '0');
  return { prefix: CONFIG.ORDER_SEQ_PREFIX, pad: CONFIG.ORDER_SEQ_PAD, value: val };
}
function adminSetOrderCounter(token, num){
  const email=_sessionEmail(token); if (!email) throw new Error('No autenticado');
  const me=_getUserByEmail(email); if (me.role!=='ADMIN') throw new Error('Solo ADMIN');
  num = Number(num);
  if (!Number.isInteger(num) || num < 0) throw new Error('Número inválido. Debe ser entero >= 0.');
  const props = PropertiesService.getScriptProperties();
  props.setProperty(CONFIG.ORDER_SEQ_PROP, String(num));
  return { ok:true, value:num };
}


// =====================================================================
// ============== NUEVO: RESET DE CONTRASEÑA (PRO, CON EMAIL) ==========
// =====================================================================
function _resetTokensSheet(){ return _sheet(CONFIG.SHEETS.RESET_TOKENS); }


function requestPasswordReset(email){
  // Respuesta SIEMPRE genérica (no revelar si existe)
  email = String(email||'').trim().toLowerCase();
  const user = _getUserByEmail(email);
  // Si el usuario no existe, igual respondemos ok (para no filtrar usuarios)
  if (user){
    const token = _uuid();
    const expMs = CONFIG.RESET_TOKEN_EXP_MIN * 60 * 1000;
    const expiresAt = new Date(Date.now()+expMs).toISOString();
    // Guardar token
    _resetTokensSheet().appendRow([token, email, expiresAt, '', _nowIso()]);
    // Link
    let baseUrl = '';
    try { baseUrl = (CONFIG.WEBAPP_URL || ScriptApp.getService().getUrl() || '').trim(); } catch(e){ baseUrl=''; }
    const link = baseUrl ? (baseUrl + '?reset=' + encodeURIComponent(token)) : ('?reset=' + encodeURIComponent(token));
    // Enviar email
    try{
      MailApp.sendEmail({
        to: email,
        subject: 'Restablecer contraseña — EL E-COMMERCE SKYLINE',
        htmlBody: `
          <p>Hola ${user.name||''},</p>
          <p>Recibimos una solicitud para restablecer tu contraseña.</p>
          <p>Hacé clic en el siguiente enlace (válido por ${CONFIG.RESET_TOKEN_EXP_MIN} minutos):</p>
          <p><a href="${link}">${link}</a></p>
          <p>Si no fuiste vos, ignorá este correo.</p>
          <p>— SKYLINE</p>
        `
      });
    }catch(e){
      // Si el envío de email falla, igualmente no exponemos info sensible
      // (podés consultar logs si necesitás debuggear)
    }
  }
  return {ok:true, message:'Si el email existe, te enviamos un link de restablecimiento.'};
}


function resetPasswordWithToken(token, newPass){
  token = String(token||'').trim();
  newPass = String(newPass||'');
  if (!token || !newPass) throw new Error('Datos incompletos');


  const sh = _resetTokensSheet();
  const last = sh.getLastRow();
  if (last < 2) throw new Error('Token inválido o expirado');


  const data = sh.getRange(2,1,last-1,5).getValues(); // token,email,expires_at,used_at,created_at
  let foundRow = -1, rowVal = null;
  for (let i=0;i<data.length;i++){
    if (String(data[i][0]) === token){
      foundRow = i+2;
      rowVal = data[i];
      break;
    }
  }
  if (foundRow < 0) throw new Error('Token inválido o expirado');


  const email = String(rowVal[1]||'').trim().toLowerCase();
  const expiresAt = new Date(rowVal[2]||'1970-01-01T00:00:00Z');
  const usedAt = String(rowVal[3]||'').trim();


  if (usedAt) throw new Error('Este enlace ya fue usado');
  if (new Date() > expiresAt) throw new Error('El enlace expiró');


  // Actualizar password en USERS
  const shU = _sheet(CONFIG.SHEETS.USERS);
  const lastU = shU.getLastRow();
  if (lastU >= 2){
    const rows = shU.getRange(2,1,lastU-1,6).getValues();
    for (let i=0;i<rows.length;i++){
      const r = rows[i];
      if (String(r[2]||'').toLowerCase() === email){
        shU.getRange(i+2,4).setValue(_hash(newPass)); // password_hash
        break;
      }
    }
  }


  // Marcar token como usado
  sh.getRange(foundRow,4).setValue(_nowIso()); // used_at


  // (Opcional PRO) auto-login: crear sesión y devolver user
  const u = _getUserByEmail(email);
  const session = u ? _newSession(email) : null;


  return {
    ok:true,
    message:'Contraseña actualizada',
    session,
    user: u ? { id:u.id, name:u.name, email:u.email, role:u.role } : null
  };
}


/**
 * FLEX para Pago de Comisiones (ADMIN) — **AJUSTADA**
 * Devuelve pedidos SOLO con Estado 1 en:
 *  - ENTREGADO
 *  - ENCOMIENDA ENTREGADA
 *  - CANCELADO
 * Respeta filtros (rango, vendedor, "Pago comisión", búsqueda) y añade
 * los campos necesarios para editar Estado 1 en la UI.
 */
function listCommissionsFlex(token, fromISO, toISO, vendorEmailOpt, onlyStatus, q) {
  const email = _sessionEmail(token); if (!email) throw new Error('No autenticado');
  const me = _getUserByEmail(email); if (me.role!=='ADMIN') throw new Error('Solo ADMIN');


  const sh=_sheet(CONFIG.SHEETS.ORDERS), idx=_orderIndexMap();
  const last=sh.getLastRow(); if (last<2) return [];
  const rg=sh.getRange(2,1,last-1, sh.getLastColumn()).getValues();


  const from=fromISO? new Date(fromISO+'T00:00:00'):new Date('1970-01-01');
  const to  =toISO  ? new Date(toISO  +'T23:59:59'):new Date('2999-12-31');


  // Estados permitidos en Pago de Comisiones
  const ALLOWED = new Set(['ENTREGADO','ENCOMIENDA ENTREGADA','CANCELADO']);


  // Mapa proveedor (una sola vez)
  const pr=_rows(_sheet(CONFIG.SHEETS.PRODUCTS),9);
  const provMap={}; pr.forEach(row=>{ const sku=String(row[2]||'').trim(); if (sku) provMap[sku]=Number(row[3]||0); });


  const res=[];
  for (const r of rg){
    const createdAt=new Date(r[idx.created_at]); if (createdAt<from || createdAt>to) continue;


    const status=String(r[idx.status]||'').toUpperCase();
    if (!ALLOWED.has(status)) continue; // <<< filtro NUEVO


    const vendor=String(r[idx.created_by]||'').trim().toLowerCase();
    if (vendorEmailOpt && vendor!==String(vendorEmailOpt).toLowerCase()) continue;


    // Filtrado por "solo estado" de pago de comisión (PENDIENTE/PAGADO)
    const paid=!!r[idx.commission_paid];
    if (onlyStatus==='PAGADO'   && !paid) continue;
    if (onlyStatus==='PENDIENTE'&&  paid) continue;


    // Búsqueda flexible (cliente/teléfono/ID/ciudad/vendedor)
    const hay = [
      r[idx.id]||'',
      String(r[idx.id]||'').replace(/^[A-Za-z]+/,'') || '', // A1089 -> 1089
      r[idx.customer_name]||'',
      r[idx.phone]||'',
      r[idx.email]||'',
      r[idx.city]||'',
      r[idx.created_by]||''
    ].join(' ').toString().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');
    const qq = (q||'').toString().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');
    if (qq && hay.indexOf(qq)<0) continue;


    // Recomputo de comisión (seguro y rápido)
    const total=Number(r[idx.total_gs]||0);
    const delivery=Number(r[idx.delivery_gs]||0);
    let prov=0;
    try{
      JSON.parse(r[idx.items_json]||'[]').forEach(it=>{
        const sku=String(it.sku||'').trim(); const qty=Number(it.qty||0);
        const pu=Number(provMap[sku]||0); prov += pu*qty;
      });
    }catch(e){}
    const commission = total - (prov + delivery);


    res.push({
      id:r[idx.id], created_at:r[idx.created_at], city:r[idx.city]||'',
      customer_name:r[idx.customer_name]||'', vendor_email:r[idx.created_by]||'',
      assigned_delivery:r[idx.assigned_delivery]||'',
      total_gs: total, commission_gs: commission,
      commission_credited: !!r[idx.commission_credited],
      commission_paid: paid, paid_at:r[idx.paid_at]||'',
      status: status
    });
  }
  return res;
}
/**
 * ADMIN — Pago de comisiones en MASA.
 * Recibe un array de IDs de pedidos y marca su pago de comisión como PAGADO o PENDIENTE.
 * - Si paid=true: descuenta de la wallet del vendedor (si la comisión fue acreditada).
 * - Si paid=false: revierte el pago (solo limpia flags/fecha, NO re-acredita nada).
 *
 * No reemplaza payVendorCommission(); la deja igual. Esto es solo un atajo masivo.
 */
function payVendorCommissionBulk(token, orderIds, paid) {
  const email = _sessionEmail(token); if (!email) throw new Error('No autenticado');
  const me = _getUserByEmail(email); if (me.role !== 'ADMIN') throw new Error('Solo ADMIN');


  if (!Array.isArray(orderIds) || orderIds.length === 0) return { ok:true, updated:0 };


  const sh = _sheet(CONFIG.SHEETS.ORDERS), idx = _orderIndexMap();
  const last = sh.getLastRow(); if (last < 2) return { ok:true, updated:0 };


  // Cargamos todo para operar en memoria y evitar múltiples lecturas/escrituras
  const rg = sh.getRange(2, 1, last - 1, sh.getLastColumn());
  const data = rg.getValues();


  // Conjunto de ids a actualizar (normalizamos a string)
  const target = new Set(orderIds.map(x => String(x)));


  let updated = 0;


  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    const id = String(r[idx.id] || '');
    if (!target.has(id)) continue;


    const credited  = !!r[idx.commission_credited];   // ¿comisión acreditada al vendedor?
    const alreadyPd = !!r[idx.commission_paid];       // ¿ya marcado como pagado?
    const vendor    = String(r[idx.created_by] || '').trim().toLowerCase();
    const commission = Number(r[idx.commission_gs] || 0);


    // Nada que hacer si no hay cambio de estado
    if (paid && alreadyPd) continue;
    if (!paid && !alreadyPd) continue;


    if (paid) {
      // Marcar PAGADO: si estaba acreditado, debitamos al vendedor
      if (credited && vendor && commission > 0) {
        _walletDebit(vendor, commission, id, 'Pago de comisión (ADMIN, masivo)');
      }
      rg.getCell(i + 1, idx.commission_paid + 1).setValue(true);
      rg.getCell(i + 1, idx.paid_at + 1).setValue(_nowIso());
      updated++;
    } else {
      // Marcar PENDIENTE: solo limpiamos flags/fecha (no re-acreditamos nada aquí)
      rg.getCell(i + 1, idx.commission_paid + 1).setValue(false);
      rg.getCell(i + 1, idx.paid_at + 1).setValue('');
      updated++;
    }
  }


  return { ok:true, updated };
}


/* =====================================================================
   ===================== NUEVO: PERFIL DE USUARIO ======================
   ===================================================================== */


/** Hoja PROFILES: helpers (no tocan nada del resto) */
function _getProfilesSheet_(){
  return _sheet(CONFIG.SHEETS.PROFILES);
}
const _PROFILE_COLS_ = [
  'email','name','phone','doc','addr',
  'bank_name','bank_type','bank_num','bank_holder','bank_holder_ci',
  'wallet_provider','wallet_number','wallet_holder',
  'updated_at'
];
function _findProfileRowByEmail_(sh, email){
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const vals = sh.getRange(2,1,last-1,_PROFILE_COLS_.length).getValues();
  for (let i=0;i<vals.length;i++){
    if (String(vals[i][0]).toLowerCase() === String(email||'').toLowerCase()){
      return i+2;
    }
  }
  return -1;
}
/** Resolver sesión -> usuario mínimo */
function _requireSessionProfile_(token){
  const email = _sessionEmail(token);
  if (!email) throw new Error('No autenticado');
  const u = _getUserByEmail(email) || { email: email, name: '', role: 'VENDEDOR' };
  return u;
}


/** API: leer mi perfil */
function getMyProfile(token){
  const u = _requireSessionProfile_(token);
  const sh = _getProfilesSheet_();
  const row = _findProfileRowByEmail_(sh, u.email);
  if (row === -1){
    return {
      email: u.email,
      name: u.name || '',
      phone: '',
      doc: '',
      addr: '',
      bank_name: '',
      bank_type: '',
      bank_num: '',
      bank_holder: '',
      bank_holder_ci: '',
      wallet_provider: '',
      wallet_number: '',
      wallet_holder: ''
    };
  }
  const vals = sh.getRange(row,1,1,_PROFILE_COLS_.length).getValues()[0];
  const obj = {};
  _PROFILE_COLS_.forEach((k,i)=> obj[k] = vals[i]);
  delete obj.updated_at;
  return obj;
}


/** API: guardar mi perfil (upsert por email) */
function saveMyProfile(token, data){
  const u = _requireSessionProfile_(token);
  const sh = _getProfilesSheet_();
  const row = _findProfileRowByEmail_(sh, u.email);
  const now = new Date();


  const payload = {
    email: u.email,
    name: data && data.name || '',
    phone: data && data.phone || '',
    doc: data && data.doc || '',
    addr: data && data.addr || '',
    bank_name: data && data.bank_name || '',
    bank_type: data && data.bank_type || '',
    bank_num: data && data.bank_num || '',
    bank_holder: data && data.bank_holder || '',
    bank_holder_ci: data && data.bank_holder_ci || '',
    wallet_provider: data && data.wallet_provider || '',
    wallet_number: data && data.wallet_number || '',
    wallet_holder: data && data.wallet_holder || '',
    updated_at: now
  };
  const rowArr = _PROFILE_COLS_.map(k => payload[k]);


  if (row === -1){
    sh.appendRow(rowArr);
  } else {
    sh.getRange(row,1,1,_PROFILE_COLS_.length).setValues([rowArr]);
  }
  return { ok:true };
}
// =================================================================
// ========== APIs REST PARA VERCEL (MANTIENE TODO LO EXISTENTE) ==========
// =================================================================

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const { action, params = {} } = data;
    
    // Log para debugging
    console.log('API Call:', action, params);
    
    // Verificar autenticación para acciones protegidas
    if (requiresAuth(action)) {
      const userEmail = _sessionEmail(params.token);
      if (!userEmail) {
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          error: 'No autenticado'
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    // Ejecutar la función correspondiente
    const result = executeFunction(action, params);
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      data: result
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    console.error('API Error:', error);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function requiresAuth(action) {
  const publicActions = [
    'login', 
    'register', 
    'requestPasswordReset', 
    'resetPasswordWithToken'
  ];
  return !publicActions.includes(action);
}

function executeFunction(action, params) {
  // Mapeo completo de todas tus funciones
  const functionMap = {
    // Auth
    'login': () => login(params.email, params.password, params.remember),
    'register': () => register(params.name, params.email, params.password, params.role),
    'me': () => me(params.token),
    'listUsersByRole': () => listUsersByRole(params.token, params.role),
    
    // Products
    'listProducts': () => listProducts(params.token),
    'addProduct': () => addProduct(params.token, params.product),
    
    // Orders
    'addOrder': () => addOrder(params.token, params.order),
    'listOrders': () => listOrders(params.token),
    'listOrdersFiltered': () => listOrdersFiltered(
      params.token, params.fromISO, params.toISO, params.q, 
      params.vendorEmailOpt, params.deliveryEmailOpt
    ),
    'updateOrderStatus': () => updateOrderStatus(params.token, params.id, params.status, params.obs),
    'updateOrderStatus2': () => updateOrderStatus2(params.token, params.id, params.status2),
    'assignDelivery': () => assignDelivery(params.token, params.id, params.deliveryEmail),
    
    // Wallet
    'getWallet': () => getWallet(params.token, params.emailOpt),
    
    // Delivery Rates
    'getDeliveryRates': () => getDeliveryRates(params.token, params.emailOpt),
    'setDeliveryRate': () => setDeliveryRate(params.token, params.deliveryEmail, params.city, params.feeGs),
    'getDeliveryClientPrices': () => getDeliveryClientPrices(params.token),
    'setClientCityPrice': () => setClientCityPrice(params.token, params.city, params.price),
    
    // Metrics & News
    'metrics': () => metrics(params.token, params.fromISO, params.toISO),
    'listNews': () => listNews(params.token),
    
    // Guides
    'getGuideText': () => getGuideText(params.token, params.orderId),
    'generateGuidePDF': () => generateGuidePDF(params.token, params.orderId),
    
    // Admin
    'adminGetOrderCounter': () => adminGetOrderCounter(params.token),
    'adminSetOrderCounter': () => adminSetOrderCounter(params.token, params.num),
    
    // Password Reset
    'requestPasswordReset': () => requestPasswordReset(params.email),
    'resetPasswordWithToken': () => resetPasswordWithToken(params.token, params.newPass),
    
    // Commissions
    'listCommissionsFlex': () => listCommissionsFlex(
      params.token, params.fromISO, params.toISO, 
      params.vendorEmailOpt, params.onlyStatus, params.q
    ),
    'payVendorCommission': () => payVendorCommission(params.token, params.orderId, params.paid),
    'payVendorCommissionBulk': () => payVendorCommissionBulk(params.token, params.orderIds, params.paid),
    
    // Profiles
    'getMyProfile': () => getMyProfile(params.token),
    'saveMyProfile': () => saveMyProfile(params.token, params.data)
  };
  
  if (functionMap[action]) {
    return functionMap[action]();
  } else {
    throw new Error(`Función no implementada: ${action}`);
  }
}

// Permite CORS para Vercel
function doOptions() {
  return ContentService.createTextOutput()
    .setHeader('Access-Control-Allow-Origin', '*')
    .setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
    .setHeader('Access-Control-Allow-Headers', 'Content-Type')
    .setMimeType(ContentService.MimeType.JSON);
}
