/**
 * ============================================================
 *  KOPERASI DASHBOARD - Google Apps Script
 *  Code.gs : SEMUA backend dalam satu file (lebih mudah di-copy)
 * ============================================================
 *  CARA PAKAI:
 *  1. Copy SELURUH isi file ini ke Code.gs di Apps Script
 *  2. Jalankan fungsi setupDatabase() sekali
 *  3. Deploy sebagai Web App
 * ============================================================
 */

// ============================================================
//  BAGIAN 1: KONFIGURASI GLOBAL
// ============================================================

const SHEETS = {
  MASTER: 'MasterBahan', STOK: 'StokHarian', TRANSAKSI: 'Transaksi',
  PEMBAYARAN: 'Pembayaran', SESI_TIM: 'SesiTim', CONFIG: 'Config'
};

const FASE_LIST = ['Fase PS','Fase A','Fase B','Fase C','SFS',
                   'Makan Malam','Sarapan PJ','Guru','IPAK'];

const KATEGORI = {
  'Karbohidrat':'K', 'Protein':'P', 'Sayuran':'S', 'Buah':'B', 'Bumbu Dapur':'BD'
};

const TIMEZONE  = 'Asia/Jakarta';
const APP_TITLE = 'Koperasi Dashboard';

// ============================================================
//  BAGIAN 2: ROUTER WEB APP
// ============================================================

function doGet(e) {
  const tmpl = HtmlService.createTemplateFromFile('Index');
  tmpl.appTitle = APP_TITLE;
  return tmpl.evaluate()
    .setTitle(APP_TITLE)
    .addMetaTag('viewport','width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
//  BAGIAN 3: HELPER SPREADSHEET & UTILITAS
// ============================================================

function getSpreadsheet() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const props   = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('DB_SPREADSHEET_ID');
  if (savedId) { try { return SpreadsheetApp.openById(savedId); } catch(e){} }
  const ss = SpreadsheetApp.create('Koperasi Dashboard — Database');
  props.setProperty('DB_SPREADSHEET_ID', ss.getId());
  try {
    const f = DriveApp.getFileById(ScriptApp.getScriptId()).getParents();
    if (f.hasNext()) DriveApp.getFileById(ss.getId()).moveTo(f.next());
  } catch(e){}
  Logger.log('Spreadsheet baru: ' + ss.getUrl());
  return ss;
}

function getSheet(name) {
  const s = getSpreadsheet().getSheetByName(name);
  if (!s) throw new Error('Sheet "'+name+'" belum ada — jalankan setupDatabase() dulu.');
  return s;
}

function readSheetAsObjects(name) {
  const sheet = getSheet(name);
  const vals  = sheet.getDataRange().getValues();
  if (vals.length < 2) return [];
  const hdrs = vals[0];
  const rows = [];
  for (let i = 1; i < vals.length; i++) {
    const obj = {}; let empty = true;
    hdrs.forEach(function(h,j){ obj[h]=vals[i][j]; if(vals[i][j]!==''&&vals[i][j]!==null) empty=false; });
    if (!empty) { obj._row = i+1; rows.push(obj); }
  }
  return rows;
}

function formatDate(d) {
  if (!d) d = new Date();
  if (typeof d==='string') d = new Date(d);
  return Utilities.formatDate(d, TIMEZONE, 'yyyy-MM-dd');
}

function formatDateLabel(d) {
  if (!d) d = new Date();
  if (typeof d==='string') d = new Date(d+'T00:00:00');
  const H=['Minggu','Senin','Selasa','Rabu','Kamis','Jumat','Sabtu'];
  const B=['Januari','Februari','Maret','April','Mei','Juni',
           'Juli','Agustus','September','Oktober','November','Desember'];
  return H[d.getDay()]+', '+d.getDate()+' '+B[d.getMonth()]+' '+d.getFullYear();
}

function bulatkanRibuan(n) { return Math.round(n/1000)*1000; }
function generateId(p)     { return (p||'ID')+'-'+Date.now()+'-'+Math.floor(Math.random()*1000); }


// ============================================================
//  BAGIAN 4: SETUP DATABASE
// ============================================================

const HEADERS = {
  MasterBahan: ['KodeBarang','NamaBahan','Kategori','Satuan','HargaBeli','HargaJual','StokMin','Aktif'],
  StokHarian:  ['Tanggal','KodeBarang','NamaBahan','Kategori','Satuan','StokAwal','Masuk','Keluar','StokAkhir','Operator','UpdatedAt'],
  Transaksi:   ['ID','Tanggal','KodeBarang','NamaBahan','Kategori','Satuan','Fase','Jumlah','HargaSatuan','TotalHarga','Operator','CreatedAt'],
  Pembayaran:  ['Tanggal','Fase','TotalBelanja','Dibayar','Sisa','Status','UpdatedAt'],
  SesiTim:     ['Tanggal','Operator','AnggotaTim','CreatedAt','UpdatedAt'],
  Config:      ['Key','Value']
};

function setupDatabase() {
  const ss = getSpreadsheet();
  Object.keys(HEADERS).forEach(function(sn) {
    let s = ss.getSheetByName(sn);
    if (!s) s = ss.insertSheet(sn);
    if (!s.getRange(1,1).getValue()) {
      const h = HEADERS[sn];
      s.getRange(1,1,1,h.length).setValues([h])
       .setFontWeight('bold').setBackground('#7C3AED').setFontColor('#FFFFFF');
      s.setFrozenRows(1);
    }
  });
  const ms = ss.getSheetByName(SHEETS.MASTER);
  if (ms.getLastRow() < 2) seedMasterBahan();
  const def = ss.getSheetByName('Sheet1');
  if (def && def.getLastRow()===0 && ss.getSheets().length>1) ss.deleteSheet(def);
  setConfig('ANGGOTA_TETAP', getConfig('ANGGOTA_TETAP')||'');
  SpreadsheetApp.flush();
  return 'Setup selesai. '+getMasterBahanRaw().length+' bahan siap.';
}

function seedMasterBahan() {
  const sheet = getSheet(SHEETS.MASTER);
  const data  = getSeedBahan().map(function(b){ return [b[0],b[1],b[2],b[3],b[4],b[4],b[5]||0,true]; });
  sheet.getRange(2,1,data.length,HEADERS.MasterBahan.length).setValues(data);
}

function setConfig(key,val) {
  const s=getSheet(SHEETS.CONFIG), d=s.getDataRange().getValues();
  for(let i=1;i<d.length;i++){ if(d[i][0]===key){ s.getRange(i+1,2).setValue(val); return; } }
  s.appendRow([key,val]);
}
function getConfig(key) {
  const d=getSheet(SHEETS.CONFIG).getDataRange().getValues();
  for(let i=1;i<d.length;i++){ if(d[i][0]===key) return d[i][1]; }
  return '';
}

const SEED_BAHAN = [
  ['K01','Beras','Karbohidrat','liter',14500,5],
  ['K02','Kentang','Karbohidrat','kg',20000,0],
  ['K03','Tepung sagu','Karbohidrat','bks',7000,0],
  ['K04','Tepung tapioka','Karbohidrat','Bks',38000,0],
  ['K05','Tepung ketan','Karbohidrat','sachet',21000,0],
  ['K06','Bihun','Karbohidrat','bungkus',11000,0],
  ['K07','Ubi','Karbohidrat','Kg',15000,0],
  ['K08','Ubi madu','Karbohidrat','kg',12000,0],
  ['K09','Ketan Hitam','Karbohidrat','kg',42000,0],
  ['K10','Tepung beras','Karbohidrat','bks',10000,0],
  ['K11','Beras merah','Karbohidrat','kg',19000,0],
  ['K12','Singkong','Karbohidrat','kg',10000,0],
  ['K13','Talas','Karbohidrat','kg',17000,0],
  ['P01','Ayam Potong 10','Protein','ekor',52000,2],
  ['P02','Ayam potong 12','Protein','ekor',57000,2],
  ['P03','Telur puyuh','Protein','buah',0,0],
  ['P04','Tahu Cina','Protein','kotak',7000,0],
  ['P05','Tempe','Protein','papan',7000,0],
  ['P06','Daging sapi','Protein','kg',157000,0],
  ['P07','Daging sapi teriaki','Protein','kg',127000,0],
  ['P08','Kacang tanah kupas','Protein','kg',37000,0],
  ['P09','Ikan kembung','Protein','ekor',7000,0],
  ['P10','biji wijen','Protein','kg',63000,0],
  ['P11','Tofu','Protein','kg',10000,0],
  ['P12','keju','Protein','pcs',17000,0],
  ['P13','Telur','Protein','butir',2000,0],
  ['P14','ikan dori','Protein','kg',61000,0],
  ['P15','tahu kuning','Protein','buah',1500,0],
  ['P16','ikan nila','Protein','kg',42000,0],
  ['P17','tahu putih','Protein','pack',6000,0],
  ['P18','edamame','Protein','kg',38000,0],
  ['P19','ayam filet','Protein','kg',55000,0],
  ['P20','daging rawon','Protein','kg',112000,0],
  ['P21','tetelan','Protein','kg',87000,0],
  ['P22','daging has dalam','Protein','kg',137000,0],
  ['P23','Ikan Patin','Protein','kg',42000,0],
  ['P24','Udang Kupas','Protein','kg',97000,0],
  ['P25','Ikan Tuna','Protein','kg',47000,0],
  ['P26','kedelai','Protein','Kg',18000,0],
  ['P27','Ayam potong','Protein','Kg',6000,0],
];


const SEED_BAHAN_2 = [
  ['S01','Bayam','Sayuran','ikat',7000,0],['S02','Wortel','Sayuran','kg',25000,0],
  ['S03','Kembang kol','Sayuran','kg',32000,0],['S04','Buncis','Sayuran','kg',22000,0],
  ['S05','Kacang merah','Sayuran','kg',22000,0],['S06','Daun bawang + seledri','Sayuran','ikat',5000,0],
  ['S07','Oyong','Sayuran','buah',3000,0],['S08','Toge kecil','Sayuran','kg',22000,0],
  ['S09','Pokcoy','Sayuran','kg',20000,0],['S10','Jamur kuping','Sayuran','kg',26000,0],
  ['S11','Timun','Sayuran','kg',17000,0],['S12','Biji jagung','Sayuran','buah',0,0],
  ['S13','Kol','Sayuran','KG',16000,0],['S14','Daun pandan','Sayuran','5 lembar',3000,0],
  ['S15','Tomat','Sayuran','kg',24000,0],['S16','Melinjo','Sayuran','kg',30000,0],
  ['S17','Kecambah','Sayuran','kg',24000,0],['S18','Sayur Asem','Sayuran','pcs',6000,0],
  ['S19','Kangkung','Sayuran','iket',6000,0],['S20','Selada','Sayuran','kg',50000,0],
  ['S21','Sawi putih','Sayuran','kg',20000,0],['S22','Cesim','Sayuran','kg',13000,0],
  ['S23','baby corn','Sayuran','kg',22000,0],['S24','sawi hijau','Sayuran','kg',17000,0],
  ['S25','kemangi','Sayuran','iket',2000,0],['S26','kacang panjang','Sayuran','kg',20000,0],
  ['S27','jagung','Sayuran','buah',7000,0],['S28','labu siam','Sayuran','kg',17000,0],
  ['S29','terong','Sayuran','buah',3000,0],['S30','Sawi Hijau','Sayuran','kg',18000,0],
  ['S31','Labu kuning','Sayuran','kg',15000,0],['S32','Toge','Sayuran','kg',20000,0],
  ['S33','Daun pisang','Sayuran','gulung',4000,0],
  ['B01','Pepaya','Buah','kg',18000,0],['B02','Melon','Buah','kg',28000,0],
  ['B03','Semangka','Buah','kg',20000,0],['B04','Pisang lampung','Buah','sisir',11000,0],
  ['B05','Kismis','Buah','bks',0,0],['B06','Jambu kristal','Buah','kg',28000,0],
  ['B07','Nanas madu','Buah','buah',5000,0],['B08','Lemon','Buah','kg',33000,0],
  ['B09','Jeruk nipis','Buah','kg',1700,0],['B10','Kelengkeng','Buah','kg',52000,0],
  ['B11','Jambu Biji','Buah','kg',25000,0],['B12','nangka kupas','Buah','kg',42000,0],
  ['B13','kelapa parut','Buah','buah',18000,0],['B14','Apel','Buah','kg',42000,0],
  ['B15','Jeruk','Buah','kg',25000,0],['B16','pir','Buah','KG',28000,0],
  ['B17','Mangga','Buah','kg',38000,0],['B18','Jeruk Medan','Buah','kg',29000,0],
  ['B19','Jambu Air','Buah','kg',26000,0],['B20','Bengkoang','Buah','Kg',20000,0],
  ['B21','pisang ambon','Buah','sisir',32000,0],['B22','Alpukat','Buah','kg',28000,0],
  ['B23','stroberi','Buah','pack',12000,0],['B24','pisang uli','Buah','sisir',12000,0],
  ['B25','Kurma','Buah','Kg',24000,0],['B26','jeruk limau','Buah','buah',1200,0],
  ['B27','belimbing','Buah','gr',18000,0],['B28','Nanas','Buah','buah',12000,0],
  ['B29','buah naga','Buah','buah',30000,0],['B30','Salak','Buah','Buah',1500,0],
  ['BD01','Bawang merah','Bumbu Dapur','kg',75000,0],['BD02','Bawang putih','Bumbu Dapur','kg',55000,0],
  ['BD03','Bawang Bombay','Bumbu Dapur','buah',7000,0],['BD04','Cabe ijo besar','Bumbu Dapur','kg',55000,0],
  ['BD05','Cabe keriting','Bumbu Dapur','kg',80000,0],['BD06','Cabe Rawit','Bumbu Dapur','kg',85000,0],
  ['BD07','Cabe merah besar','Bumbu Dapur','kg',85000,0],['BD08','Cabe ijo kriting','Bumbu Dapur','buah',61000,0],
  ['BD09','Tomat ijo','Bumbu Dapur','buah',700,0],['BD10','Temu Kunci','Bumbu Dapur','bonggol',1000,0],
  ['BD11','Kunyit','Bumbu Dapur','kg',20000,0],['BD12','Biji Pala','Bumbu Dapur','buah',3000,0],
  ['BD13','Garam reina','Bumbu Dapur','kantong',6000,0],['BD14','Gula Pasir','Bumbu Dapur','kg',20000,0],
  ['BD15','Ketumbar','Bumbu Dapur','kg',51000,0],['BD16','Lengkuas','Bumbu Dapur','kg',22000,0],
  ['BD17','Kencur','Bumbu Dapur','kg',51000,0],['BD18','chiaseed','Bumbu Dapur','KG',100000,0],
  ['BD19','Kemiri','Bumbu Dapur','kg',80000,0],['BD20','Jahe','Bumbu Dapur','kg',32000,0],
  ['BD21','Terasi','Bumbu Dapur','sachet',500,0],['BD22','Daun salam/jeruk','Bumbu Dapur','per 5 lembar',1000,0],
  ['BD23','Lada bubuk','Bumbu Dapur','bks',4000,0],['BD24','Sereh','Bumbu Dapur','per batang',2000,0],
  ['BD25','Sarung tangan','Bumbu Dapur','kotak',10000,0],['BD26','Gula merah','Bumbu Dapur','kg',22000,0],
  ['BD27','Bumbu Dapur','Bumbu Dapur','bks',12000,0],['BD28','Asam kandis','Bumbu Dapur','gr',12000,0],
  ['BD29','Agar Swallow','Bumbu Dapur','bks',5500,0],['BD30','Nutrijel','Bumbu Dapur','bks',7000,0],
  ['BD31','Minyak','Bumbu Dapur','liter',24000,0],['BD32','Kecap asin','Bumbu Dapur','botol',15000,0],
  ['BD33','Santan','Bumbu Dapur','bks',9000,0],['BD34','Garam masak kecil','Bumbu Dapur','kg',3500,0],
  ['BD35','Cengkeh','Bumbu Dapur','pcs',2000,0],['BD36','Kapulaga','Bumbu Dapur','bks',1500,0],
  ['BD37','Kayu secang','Bumbu Dapur','kg',90000,0],['BD38','Kayu manis','Bumbu Dapur','kg',167000,0],
  ['BD39','bunga lawang','Bumbu Dapur','bungkus',6000,0],['BD40','Teh cap botol','Bumbu Dapur','pcs',500,0],
  ['BD41','Garam masak besar','Bumbu Dapur','pcs',6000,0],['BD42','Citric acid','Bumbu Dapur','pcs',0,0],
  ['BD43','mutiara','Bumbu Dapur','kg',0,0],['BD44','kerupuk udang','Bumbu Dapur','kg',0,0],
  ['BD45','gula halus','Bumbu Dapur','bks',0,0],['BD46','vanili','Bumbu Dapur','bks',1000,0],
  ['BD47','bumbu pecel','Bumbu Dapur','bks',0,0],['BD48','asem jawa','Bumbu Dapur','bks',8000,0],
  ['BD49','Lada hitam','Bumbu Dapur','kg',200000,0]
];

// Gabungkan dua array seed
function getSeedBahan() { return SEED_BAHAN.concat(SEED_BAHAN_2); }


// ============================================================
//  BAGIAN 5: MASTER BAHAN (CRUD)
// ============================================================

function getMasterBahanRaw() { return readSheetAsObjects(SHEETS.MASTER); }

function getMasterBahan(hanyaAktif) {
  const list = getMasterBahanRaw().map(function(r){
    return { kode:r.KodeBarang, nama:r.NamaBahan, kategori:r.Kategori, satuan:r.Satuan,
      hargaBeli:Number(r.HargaBeli)||0, hargaJual:Number(r.HargaJual)||0,
      stokMin:Number(r.StokMin)||0,
      aktif:r.Aktif===true||r.Aktif==='TRUE'||r.Aktif==='true', row:r._row };
  });
  return hanyaAktif ? list.filter(function(b){return b.aktif;}) : list;
}

function getBahanByKode(kode) {
  return getMasterBahan(false).find(function(b){return b.kode===kode;})||null;
}

function tambahBahan(data) {
  if (!data.nama||!String(data.nama).trim()) throw new Error('Nama bahan wajib diisi.');
  if (!KATEGORI[data.kategori]) throw new Error('Kategori tidak valid.');
  let kode = data.kode&&String(data.kode).trim()||generateKodeBerikutnya(data.kategori);
  if (getBahanByKode(kode)) throw new Error('Kode "'+kode+'" sudah dipakai.');
  const hb=Number(data.hargaBeli)||0, hj=Number(data.hargaJual)||hb;
  getSheet(SHEETS.MASTER).appendRow([kode,String(data.nama).trim(),data.kategori,
    data.satuan||'',hb,hj,Number(data.stokMin)||0,true]);
  SpreadsheetApp.flush();
  return getBahanByKode(kode);
}

function updateBahan(data) {
  const b=getBahanByKode(data.kode);
  if(!b) throw new Error('Bahan "'+data.kode+'" tidak ditemukan.');
  const s=getSheet(SHEETS.MASTER), r=b.row;
  s.getRange(r,2).setValue(data.nama!=null?String(data.nama).trim():b.nama);
  s.getRange(r,3).setValue(data.kategori||b.kategori);
  s.getRange(r,4).setValue(data.satuan!=null?data.satuan:b.satuan);
  s.getRange(r,5).setValue(data.hargaBeli!=null?Number(data.hargaBeli):b.hargaBeli);
  s.getRange(r,6).setValue(data.hargaJual!=null?Number(data.hargaJual):b.hargaJual);
  s.getRange(r,7).setValue(data.stokMin!=null?Number(data.stokMin):b.stokMin);
  if(data.aktif!=null) s.getRange(r,8).setValue(!!data.aktif);
  SpreadsheetApp.flush();
  return getBahanByKode(data.kode);
}

function nonaktifkanBahan(kode) { return updateBahan({kode:kode,aktif:false}); }

function generateKodeBerikutnya(kategori) {
  const prefix=KATEGORI[kategori];
  let max=0;
  getMasterBahan(false).filter(function(b){return b.kategori===kategori;}).forEach(function(b){
    const m=String(b.kode).match(new RegExp('^'+prefix+'(\\d+)$'));
    if(m){ const n=parseInt(m[1],10); if(n>max) max=n; }
  });
  const next=max+1;
  return prefix+(next<10?'0'+next:''+next);
}


// ============================================================
//  BAGIAN 6: STOK HARIAN
// ============================================================

function getStokAkhirSebelum(kode,tanggal) {
  const rows=readSheetAsObjects(SHEETS.STOK).filter(function(r){
    return r.KodeBarang===kode&&formatDate(r.Tanggal)<tanggal;
  });
  if(!rows.length) return 0;
  rows.sort(function(a,b){return formatDate(a.Tanggal)<formatDate(b.Tanggal)?1:-1;});
  return Number(rows[0].StokAkhir)||0;
}

function getStokMapByTanggal(tanggal) {
  const tgl=formatDate(tanggal), map={};
  readSheetAsObjects(SHEETS.STOK).filter(function(r){return formatDate(r.Tanggal)===tgl;})
    .forEach(function(r){map[r.KodeBarang]=r;});
  return map;
}

function getSnapshotStok(tanggal) {
  const tgl=formatDate(tanggal||new Date());
  const stokMap=getStokMapByTanggal(tgl), keluarMap=getTotalKeluarByTanggal(tgl);
  return getMasterBahan(true).map(function(b){
    const rec=stokMap[b.kode];
    const stokAwal=rec?Number(rec.StokAwal)||0:getStokAkhirSebelum(b.kode,tgl);
    const masuk=rec?Number(rec.Masuk)||0:0;
    const keluar=keluarMap[b.kode]||0, stokAkhir=stokAwal+masuk-keluar;
    return { kode:b.kode, nama:b.nama, kategori:b.kategori, satuan:b.satuan,
      hargaBeli:b.hargaBeli, hargaJual:b.hargaJual, stokMin:b.stokMin,
      stokAwal, masuk, keluar, stokAkhir,
      kritis:b.stokMin>0&&stokAkhir<=b.stokMin, habis:stokAkhir<=0 };
  });
}

function inputStokMasuk(data) {
  const tgl=formatDate(data.tanggal||new Date()), bahan=getBahanByKode(data.kode);
  if(!bahan) throw new Error('Bahan tidak ditemukan: '+data.kode);
  const sheet=getSheet(SHEETS.STOK), rows=readSheetAsObjects(SHEETS.STOK);
  const ex=rows.find(function(r){return r.KodeBarang===data.kode&&formatDate(r.Tanggal)===tgl;});
  const masuk=Number(data.masuk)||0, now=new Date(), op=data.operator||'';
  if(ex){
    const sa=data.stokAwalOverride!=null?Number(data.stokAwalOverride):Number(ex.StokAwal)||0;
    sheet.getRange(ex._row,6).setValue(sa); sheet.getRange(ex._row,7).setValue(masuk);
    sheet.getRange(ex._row,10).setValue(op); sheet.getRange(ex._row,11).setValue(now);
  } else {
    const sa=data.stokAwalOverride!=null?Number(data.stokAwalOverride):getStokAkhirSebelum(data.kode,tgl);
    sheet.appendRow([tgl,bahan.kode,bahan.nama,bahan.kategori,bahan.satuan,sa,masuk,0,sa+masuk,op,now]);
  }
  rekalkulasiStok(tgl); SpreadsheetApp.flush();
  return getSnapshotStok(tgl);
}

function rekalkulasiStok(tanggal) {
  const tgl=formatDate(tanggal), sheet=getSheet(SHEETS.STOK);
  const keluarMap=getTotalKeluarByTanggal(tgl), data=sheet.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    if(formatDate(data[i][0])===tgl){
      const keluar=keluarMap[data[i][1]]||0;
      sheet.getRange(i+1,8).setValue(keluar);
      sheet.getRange(i+1,9).setValue((Number(data[i][5])||0)+(Number(data[i][6])||0)-keluar);
    }
  }
}

function getStokKritis(tanggal) {
  return getSnapshotStok(tanggal).filter(function(b){return b.kritis||b.habis;});
}


// ============================================================
//  BAGIAN 7: TRANSAKSI
// ============================================================

function getTransaksiByTanggal(tanggal) {
  const tgl=formatDate(tanggal);
  return readSheetAsObjects(SHEETS.TRANSAKSI).filter(function(r){return formatDate(r.Tanggal)===tgl;});
}

function getTotalKeluarByTanggal(tanggal) {
  const map={};
  getTransaksiByTanggal(tanggal).forEach(function(t){
    map[t.KodeBarang]=(map[t.KodeBarang]||0)+(Number(t.Jumlah)||0);
  });
  return map;
}

function getMatriksDistribusi(tanggal) {
  const tgl=formatDate(tanggal||new Date()), snapshot=getSnapshotStok(tgl);
  const distMap={};
  getTransaksiByTanggal(tgl).forEach(function(t){
    if(!distMap[t.KodeBarang]) distMap[t.KodeBarang]={};
    distMap[t.KodeBarang][t.Fase]=Number(t.Jumlah)||0;
  });
  const bahan=snapshot.map(function(b){
    const dist={}, dm=distMap[b.kode]||{}; let totalKeluar=0;
    FASE_LIST.forEach(function(f){ dist[f]=dm[f]||0; totalKeluar+=dist[f]; });
    return { kode:b.kode, nama:b.nama, kategori:b.kategori, satuan:b.satuan,
      hargaSatuan:b.hargaBeli, stokAwal:b.stokAwal, masuk:b.masuk,
      stokAkhir:b.stokAkhir, stokMin:b.stokMin, kritis:b.kritis, distribusi:dist, totalKeluar };
  });
  return { tanggal:tgl, tanggalLabel:formatDateLabel(tgl), fase:FASE_LIST,
    bahan, ringkasan:hitungRingkasanFase(tgl) };
}

function simpanDistribusiBahan(data) {
  const tgl=formatDate(data.tanggal||new Date()), bahan=getBahanByKode(data.kode);
  if(!bahan) throw new Error('Bahan tidak ditemukan: '+data.kode);
  const sheet=getSheet(SHEETS.TRANSAKSI), vals=sheet.getDataRange().getValues();
  for(let i=vals.length-1;i>=1;i--){
    if(formatDate(vals[i][1])===tgl&&vals[i][2]===data.kode) sheet.deleteRow(i+1);
  }
  const now=new Date(), op=data.operator||'', dist=data.distribusi||{}, baris=[];
  FASE_LIST.forEach(function(f){
    const j=Number(dist[f])||0;
    if(j>0) baris.push([generateId('TRX'),tgl,bahan.kode,bahan.nama,bahan.kategori,
      bahan.satuan,f,j,bahan.hargaBeli,j*bahan.hargaBeli,op,now]);
  });
  if(baris.length) sheet.getRange(sheet.getLastRow()+1,1,baris.length,HEADERS.Transaksi.length).setValues(baris);
  rekalkulasiStok(tgl); hitungUlangPembayaran(tgl); SpreadsheetApp.flush();
  return { kode:data.kode, ringkasan:hitungRingkasanFase(tgl) };
}

function simpanDistribusiBatch(data) {
  const tgl=formatDate(data.tanggal||new Date());
  (data.items||[]).forEach(function(it){
    simpanDistribusiBahan({tanggal:tgl,kode:it.kode,distribusi:it.distribusi,operator:data.operator});
  });
  return getMatriksDistribusi(tgl);
}

function hitungRingkasanFase(tanggal) {
  const perFase={}, trx=getTransaksiByTanggal(tanggal);
  FASE_LIST.forEach(function(f){perFase[f]={total:0,totalBulat:0};});
  trx.forEach(function(t){ if(perFase[t.Fase]) perFase[t.Fase].total+=Number(t.TotalHarga)||0; });
  let grandTotal=0, grandTotalBulat=0;
  FASE_LIST.forEach(function(f){
    perFase[f].totalBulat=bulatkanRibuan(perFase[f].total);
    grandTotal+=perFase[f].total; grandTotalBulat+=perFase[f].totalBulat;
  });
  return { perFase, grandTotal, grandTotalBulat };
}


// ============================================================
//  BAGIAN 8: LAPORAN & PEMBAYARAN
// ============================================================

function hitungUlangPembayaran(tanggal) {
  const tgl=formatDate(tanggal), ring=hitungRingkasanFase(tgl);
  const sheet=getSheet(SHEETS.PEMBAYARAN), rows=readSheetAsObjects(SHEETS.PEMBAYARAN), now=new Date();
  FASE_LIST.forEach(function(f){
    const tot=ring.perFase[f].totalBulat;
    const ex=rows.find(function(r){return formatDate(r.Tanggal)===tgl&&r.Fase===f;});
    if(ex){
      const dib=Number(ex.Dibayar)||0;
      sheet.getRange(ex._row,3).setValue(tot);
      sheet.getRange(ex._row,5).setValue(tot-dib);
      sheet.getRange(ex._row,6).setValue(statusBayar(tot,dib));
      sheet.getRange(ex._row,7).setValue(now);
    } else if(tot>0) sheet.appendRow([tgl,f,tot,0,tot,'Belum Bayar',now]);
  });
}

function statusBayar(tot,dib){ if(tot<=0)return '-'; if(dib>=tot)return 'Lunas'; if(dib>0)return 'Kurang'; return 'Belum Bayar'; }

function getPembayaran(tanggal) {
  const tgl=formatDate(tanggal);
  hitungUlangPembayaran(tgl);
  return readSheetAsObjects(SHEETS.PEMBAYARAN)
    .filter(function(r){return formatDate(r.Tanggal)===tgl;})
    .map(function(r){return {fase:r.Fase,totalBelanja:Number(r.TotalBelanja)||0,
      dibayar:Number(r.Dibayar)||0,sisa:Number(r.Sisa)||0,status:r.Status};});
}

function catatPembayaran(data) {
  const tgl=formatDate(data.tanggal);
  hitungUlangPembayaran(tgl);
  const sheet=getSheet(SHEETS.PEMBAYARAN), rows=readSheetAsObjects(SHEETS.PEMBAYARAN);
  const ex=rows.find(function(r){return formatDate(r.Tanggal)===tgl&&r.Fase===data.fase;});
  if(!ex) throw new Error('Belum ada belanja untuk fase '+data.fase);
  const tot=Number(ex.TotalBelanja)||0, dib=Number(data.dibayar)||0;
  sheet.getRange(ex._row,4).setValue(dib);
  sheet.getRange(ex._row,5).setValue(tot-dib);
  sheet.getRange(ex._row,6).setValue(statusBayar(tot,dib));
  sheet.getRange(ex._row,7).setValue(new Date());
  SpreadsheetApp.flush();
  return getPembayaran(tgl);
}

let _bahanCache=null;
function getBahanByKodeCached(kode){
  if(!_bahanCache){_bahanCache={};getMasterBahan(false).forEach(function(b){_bahanCache[b.kode]=b;});}
  return _bahanCache[kode]||null;
}

function getRekapHarian(tanggal) {
  const tgl=formatDate(tanggal||new Date()), trx=getTransaksiByTanggal(tgl);
  const ring=hitungRingkasanFase(tgl), pem=getPembayaran(tgl), snap=getSnapshotStok(tgl);
  let totalModal=0; snap.forEach(function(b){totalModal+=(b.stokAwal+b.masuk)*b.hargaBeli;});
  let totalPenjualan=0;
  trx.forEach(function(t){
    const b=getBahanByKodeCached(t.KodeBarang);
    totalPenjualan+=(Number(t.Jumlah)||0)*(b?b.hargaJual:Number(t.HargaSatuan)||0);
  });
  const totalDibayar=pem.reduce(function(s,p){return s+p.dibayar;},0);
  const totalKekurangan=pem.reduce(function(s,p){return s+Math.max(0,p.sisa);},0);
  return { tanggal:tgl, tanggalLabel:formatDateLabel(tgl), totalPenjualan, totalModal,
    totalBelanja:ring.grandTotalBulat, keuntungan:totalPenjualan-totalModal,
    totalDibayar, totalKekurangan, perFase:ring.perFase, pembayaran:pem, jumlahTransaksi:trx.length };
}

function getRekapMingguan(tanggalAkhir,jumlahHari) {
  const akhir=tanggalAkhir?new Date(tanggalAkhir+'T00:00:00'):new Date(), n=jumlahHari||7;
  const harian=[]; let totP=0,totM=0,totB=0;
  for(let i=n-1;i>=0;i--){
    const d=new Date(akhir); d.setDate(d.getDate()-i);
    const tgl=formatDate(d), r=getRekapHarian(tgl);
    harian.push({tanggal:tgl,label:Utilities.formatDate(d,TIMEZONE,'EEE dd/MM'),
      penjualan:r.totalPenjualan,modal:r.totalModal,belanja:r.totalBelanja,keuntungan:r.keuntungan});
    totP+=r.totalPenjualan; totM+=r.totalModal; totB+=r.totalBelanja;
  }
  return {dari:harian[0]?harian[0].tanggal:formatDate(akhir),sampai:formatDate(akhir),
    harian,totalPenjualan:totP,totalModal:totM,totalBelanja:totB,totalKeuntungan:totP-totM};
}


function generateLaporanFormatLama(tanggal) {
  const tgl=formatDate(tanggal||new Date()), ss=getSpreadsheet(), nm='Laporan_'+tgl;
  const lama=ss.getSheetByName(nm); if(lama) ss.deleteSheet(lama);
  const sheet=ss.insertSheet(nm), snap=getSnapshotStok(tgl);
  const matriks=getMatriksDistribusi(tgl), sesi=getSesiTim(tgl);
  const bMap={}; matriks.bahan.forEach(function(b){bMap[b.kode]=b;});
  const fase=FASE_LIST, totK=8+fase.length*2;
  sheet.getRange(1,1).setValue('Anggota Koperasi');
  sheet.getRange(1,9).setValue('Laporan Belanja Harian');
  sheet.getRange(1,totK).setValue(formatDateLabel(tgl));
  if(sesi){sheet.getRange(2,1).setValue('Operator: '+sesi.operator);sheet.getRange(2,9).setValue('Tim: '+sesi.anggota.join(', '));}
  const hr=4; let h=['Kode','Nama Bahan','Stok Awal','Masuk','Stok Akhir','Satuan','Harga','HPP'];
  fase.forEach(function(f){h.push(f);h.push('');});
  sheet.getRange(hr,1,1,h.length).setValues([h]);
  const sr=hr+1; let sub=['','','','','','','',''];
  fase.forEach(function(){sub.push('Jumlah');sub.push('Harga');});
  sheet.getRange(sr,1,1,sub.length).setValues([sub]);
  let r=sr+1; const tpf={};
  fase.forEach(function(f){tpf[f]=0;});
  Object.keys(KATEGORI).forEach(function(kat){
    sheet.getRange(r,1).setValue(kat).setFontWeight('bold').setBackground('#EDE9FE'); r++;
    snap.filter(function(b){return b.kategori===kat;}).forEach(function(b){
      const bm=bMap[b.kode]||{distribusi:{}}, hpp=(b.stokAwal+b.masuk)*b.hargaBeli;
      const baris=[b.kode,b.nama,b.stokAwal,b.masuk,b.stokAkhir,b.satuan,b.hargaBeli,hpp];
      fase.forEach(function(f){
        const j=(bm.distribusi&&bm.distribusi[f])||0, hh=j*b.hargaBeli;
        baris.push(j>0?j:''); baris.push(hh>0?hh:0); tpf[f]+=hh;
      });
      sheet.getRange(r,1,1,baris.length).setValues([baris]); r++;
    });
  });
  r++; const tr=['Total Belanja Per Fase','','','','','','',''];
  fase.forEach(function(f){tr.push('');tr.push(bulatkanRibuan(tpf[f]));});
  sheet.getRange(r,1,1,tr.length).setValues([tr]); sheet.getRange(r,1).setFontWeight('bold'); r+=2;
  const rekap=getRekapHarian(tgl);
  [['Total Penjualan',rekap.totalPenjualan],['Total Modal',rekap.totalModal],
   ['Keuntungan',rekap.keuntungan],['Dibayar',rekap.totalDibayar],['Kekurangan',rekap.totalKekurangan]]
    .forEach(function(row){sheet.getRange(r,1).setValue(row[0]).setFontWeight('bold');sheet.getRange(r,8).setValue(row[1]);r++;});
  sheet.getRange(hr,1,2,h.length).setFontWeight('bold').setBackground('#7C3AED').setFontColor('#FFFFFF');
  sheet.setFrozenRows(sr); sheet.autoResizeColumns(1,Math.min(h.length,12));
  SpreadsheetApp.flush();
  return {sheetName:nm,url:ss.getUrl()+'#gid='+sheet.getSheetId()};
}

// ============================================================
//  BAGIAN 9: AUTH / SESI TIM
// ============================================================

function mulaiSesiTim(data) {
  if(!data||!data.operator||!String(data.operator).trim()) throw new Error('Nama operator wajib diisi.');
  const tanggal=formatDate(data.tanggal||new Date()), operator=String(data.operator).trim();
  let anggota=(Array.isArray(data.anggota)?data.anggota:[])
    .map(function(n){return String(n).trim();}).filter(function(n){return n!=='';});
  if(!anggota.some(function(n){return n.toLowerCase()===operator.toLowerCase();})) anggota.unshift(operator);
  const seen={}; anggota=anggota.filter(function(n){const k=n.toLowerCase();if(seen[k])return false;seen[k]=true;return true;});
  const sheet=getSheet(SHEETS.SESI_TIM), rows=readSheetAsObjects(SHEETS.SESI_TIM);
  const now=new Date(), str=anggota.join(', ');
  const ex=rows.find(function(r){return formatDate(r.Tanggal)===tanggal;});
  if(ex){sheet.getRange(ex._row,2).setValue(operator);sheet.getRange(ex._row,3).setValue(str);sheet.getRange(ex._row,5).setValue(now);}
  else sheet.appendRow([tanggal,operator,str,now,now]);
  perbaruiAnggotaTetap(anggota); SpreadsheetApp.flush();
  return getSesiTim(tanggal);
}

function getSesiTim(tanggal) {
  const tgl=formatDate(tanggal||new Date());
  const s=readSheetAsObjects(SHEETS.SESI_TIM).find(function(r){return formatDate(r.Tanggal)===tgl;});
  if(!s) return null;
  return {tanggal:tgl,tanggalLabel:formatDateLabel(tgl),operator:s.Operator,
    anggota:String(s.AnggotaTim||'').split(',').map(function(n){return n.trim();}).filter(function(n){return n!=='';})};
}

function getInitState() {
  const today=formatDate(new Date());
  return {today,todayLabel:formatDateLabel(today),sesiHariIni:getSesiTim(today),
    anggotaTetap:getAnggotaTetap(),fase:FASE_LIST,kategori:Object.keys(KATEGORI)};
}

function getAnggotaTetap() {
  const raw=getConfig('ANGGOTA_TETAP');
  return raw?String(raw).split(',').map(function(n){return n.trim();}).filter(function(n){return n!=='';}):[]; 
}

function perbaruiAnggotaTetap(baru) {
  const seen={};
  getAnggotaTetap().forEach(function(n){seen[n.toLowerCase()]=n;});
  baru.forEach(function(n){if(n&&!seen[n.toLowerCase()]) seen[n.toLowerCase()]=n;});
  setConfig('ANGGOTA_TETAP',Object.keys(seen).map(function(k){return seen[k];}).join(', '));
}
