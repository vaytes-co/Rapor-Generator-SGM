let PDFDocument, rgb, StandardFonts;

function getLibraries(){
  if(!window.XLSX) throw new Error('Library Excel belum termuat. Periksa koneksi internet atau CDN SheetJS.');
  if(!window.PDFLib) throw new Error('Library PDF belum termuat. Periksa koneksi internet atau CDN pdf-lib.');
  if(!window.JSZip) throw new Error('Library ZIP belum termuat. Periksa koneksi internet atau CDN JSZip.');
  ({PDFDocument, rgb, StandardFonts}=window.PDFLib);
  initPdfColors();
}

const state={type:'glow',level:'lower',workbook:null,students:[],photos:new Map(),manualPhotos:new Map(),selected:null,generated:new Map(),principalSignature:null};

const TEMPLATES={
  glowLower:'assets/templates/Rapor Glow Lower.pdf',
  glowMiddle:'assets/templates/Rapor Glow Middle Template.pdf',
  glowUpper:'assets/templates/Rapor Glow Upper.pdf',
  sunny:'assets/templates/Rapor Sunny Template.pdf',
  infant:'assets/templates/Rapor Infant Template.pdf'
};

const COLORS={orange:null,black:null,white:null};
function initPdfColors(){const c=window.PDFLib.rgb;COLORS.orange=c(1,.60,0);COLORS.black=c(.07,.08,.10);COLORS.white=c(1,1,1);}

const SIGNATURE_DEFAULTS={
  glow:{
    lower:{
      p2:{
        date:{x:515,y:774.3,w:70,h:18,size:10,bold:false,italic:false,underline:false,align:'left'},
        teacher:{x:290,y:885,w:95,h:18,size:10.5,bold:true,italic:false,underline:false,align:'center'},
        principal:{x:459,y:885,w:95,h:18,size:10.5,bold:true,italic:false,underline:false,align:'center',digitalEnabled:false,digital:{x:390,y:810,w:250,h:120}}
      },
      p3:{
        date:{x:513,y:735,w:70,h:18,size:10.5,bold:false,italic:false,underline:false,align:'left'},
        teacher:{x:315,y:870,w:95,h:18,size:10.5,bold:true,italic:false,underline:false,align:'center'},
        principal:{x:460,y:870,w:95,h:18,size:10.5,bold:true,italic:false,underline:false,align:'center',digitalEnabled:false,digital:{x:390,y:782,w:250,h:120}}
      }
    },
    middle:{
      p2:{
        date:{x:505,y:775.5,w:70,h:18,size:10.5,bold:false,italic:false,underline:false,align:'left'},
        teacher:{x:315,y:880,w:95,h:18,size:10.5,bold:true,italic:false,underline:false,align:'center'},
        principal:{x:460,y:880,w:95,h:18,size:10.5,bold:true,italic:false,underline:false,align:'center',digitalEnabled:false,digital:{x:390,y:800,w:250,h:120}}
      },
      p3:{
        date:{x:513,y:735,w:70,h:18,size:10.5,bold:false,italic:false,underline:false,align:'left'},
        teacher:{x:349,y:870,w:95,h:18,size:10.5,bold:true,italic:false,underline:false,align:'center'},
        principal:{x:475,y:870,w:95,h:18,size:10.5,bold:true,italic:false,underline:false,align:'center',digitalEnabled:false,digital:{x:415,y:785,w:250,h:120}}
      }
    },
    upper:{
      p2:{
        date:{x:510,y:770,w:70,h:18,size:10.5,bold:false,italic:false,underline:false,align:'left'},
        teacher:{x:304,y:875,w:95,h:18,size:10.5,bold:true,italic:false,underline:false,align:'center'},
        principal:{x:450,y:875,w:95,h:18,size:10.5,bold:true,italic:false,underline:false,align:'center',digitalEnabled:false,digital:{x:385,y:795,w:250,h:120}}
      },
      p3:{
        date:{x:500,y:753,w:70,h:18,size:10.5,bold:false,italic:false,underline:false,align:'left'},
        teacher:{x:347,y:870,w:95,h:18,size:10.5,bold:true,italic:false,underline:false,align:'center'},
        principal:{x:470,y:870,w:95,h:18,size:10.5,bold:true,italic:false,underline:false,align:'center',digitalEnabled:false,digital:{x:405,y:785,w:250,h:120}}
      }
    }
  },
  sunny:{
    p2:{
      date:{x:506,y:755,w:80,h:18,size:10.5,bold:false,italic:false,underline:false,align:'left'},
      teacher:{x:300,y:870,w:100,h:18,size:10.5,bold:true,italic:false,underline:false,align:'center'},
      principal:{x:455,y:870,w:105,h:18,size:10.5,bold:true,italic:false,underline:false,align:'center',digitalEnabled:false,digital:{x:395,y:795,w:250,h:120}}
    }
  },
  infant:{
    p2:{
      date:{x:501,y:791.3,w:64,h:18,size:10.5,bold:false,italic:false,underline:false,align:'left'},
      teacher:{x:328,y:883,w:100,h:18,size:10.5,bold:true,italic:false,underline:false,align:'center'},
      principal:{x:445,y:883,w:110,h:18,size:10.5,bold:true,italic:false,underline:false,align:'center',digitalEnabled:false,digital:{x:390,y:805,w:250,h:120}}
    }
  }
};

const BIODATA_DEFAULTS={
  glowLower:{x:0,y:5,size:12},
  glowMiddle:{x:0,y:5,size:12},
  glowUpper:{x:0,y:5,size:12},
  sunny:{x:0,y:5,size:12},
  infant:{x:0,y:5,size:12}
};

const BIODATA_STORAGE_KEY='rapor-generator-biodata-settings-v3';
let biodataStore={};

const SIGNATURE_STORAGE_KEY='rapor-generator-signature-settings-v4';
let signatureStore={};
let signaturePreviewTimer=null;
let hasUnsavedWork=false;
let allowUnload=false;
let pendingNavigation=false;

function signatureStorageKey(kind,pageIndex){
  const type=kind==='sunny'?'sunny':kind==='infant'?'infant':'glow';
  const level=state.type==='glow'?state.level:'all';
  return `${type}-${level}-p${pageIndex}`;
}

function cloneSignatureConfig(cfg){
  return JSON.parse(JSON.stringify(cfg));
}

function loadSignatureStore(){
  try{
    const raw=localStorage.getItem(SIGNATURE_STORAGE_KEY);
    signatureStore=raw?JSON.parse(raw):{};
  }catch{signatureStore={};}
}

function saveSignatureStore(){
  try{localStorage.setItem(SIGNATURE_STORAGE_KEY,JSON.stringify(signatureStore));}
  catch(err){console.warn('Pengaturan tanda tangan tidak dapat disimpan di browser ini.',err);}
}

function loadBiodataStore(){
  try{
    const raw=localStorage.getItem(BIODATA_STORAGE_KEY);
    biodataStore=raw?JSON.parse(raw):{};
  }catch{biodataStore={};}
}

function saveBiodataStore(){
  try{localStorage.setItem(BIODATA_STORAGE_KEY,JSON.stringify(biodataStore));}
  catch(err){console.warn('Pengaturan biodata tidak dapat disimpan di browser ini.',err);}
}

function biodataStorageKey(){
  return state.type==='glow'?`glow-${state.level}`:state.type;
}

function defaultBiodataConfig(){
  const key=state.type==='glow'?`glow${state.level[0].toUpperCase()+state.level.slice(1)}`:state.type;
  return cloneSignatureConfig(BIODATA_DEFAULTS[key]||BIODATA_DEFAULTS.glowLower);
}

function getBiodataConfig(){
  const key=biodataStorageKey();
  if(!biodataStore[key]){
    biodataStore[key]=defaultBiodataConfig();
    saveBiodataStore();
  }
  return biodataStore[key];
}

function updateBiodataField(field,value){
  const cfg=getBiodataConfig();
  cfg[field]=Number(value);
  saveBiodataStore();
  markDirty();
  scheduleSignaturePreview();
}

function resetBiodataConfig(){
  biodataStore[biodataStorageKey()]=defaultBiodataConfig();
  saveBiodataStore();
  markDirty();
  renderBiodataEditor();
  scheduleSignaturePreview(true);
}

function defaultSignatureConfig(kind,pageIndex){
  if(kind==='sunny')return cloneSignatureConfig(SIGNATURE_DEFAULTS.sunny.p2);
  if(kind==='infant')return cloneSignatureConfig(SIGNATURE_DEFAULTS.infant.p2);
  const level=state.level||'lower';
  const glowDefaults=SIGNATURE_DEFAULTS.glow[level]||SIGNATURE_DEFAULTS.glow.lower;
  return cloneSignatureConfig(pageIndex===3?glowDefaults.p3:glowDefaults.p2);
}

function getSignatureConfig(kind,pageIndex){
  const key=signatureStorageKey(kind,pageIndex);
  const defaults=defaultSignatureConfig(kind,pageIndex);
  if(!signatureStore[key]){
    signatureStore[key]=defaults;
    saveSignatureStore();
  }else{
    const current=signatureStore[key];
    for(const role of ['date','teacher','principal']){
      if(!current[role])current[role]=cloneSignatureConfig(defaults[role]);
      if(role==='principal'){
        if(!current[role].digital)current[role].digital=cloneSignatureConfig(defaults[role].digital);
        if(typeof current[role].digitalEnabled!=='boolean')current[role].digitalEnabled=!!defaults[role].digitalEnabled;
      }
    }
  }
  return signatureStore[key];
}

function resetSignatureConfig(kind,pageIndex){
  const key=signatureStorageKey(kind,pageIndex);
  signatureStore[key]=defaultSignatureConfig(kind,pageIndex);
  saveSignatureStore();
  markDirty();
  renderSignatureEditor();
  scheduleSignaturePreview(true);
}

function templatePageCount(){
  if(state.type==='glow')return 3;
  return 2;
}

function signatureGroupOptions(){
  return Array.from({length:Math.max(0,templatePageCount()-1)},(_,i)=>{
    const page=i+2;
    return {value:`p${page}`,label:`Halaman ${page} — Tanda tangan`};
  });
}

function signatureUiPageIndex(){
  const select=$('signaturePageSelect');
  const value=select?.value||'p2';
  return Number(String(value).replace('p',''))||2;
}

function signatureUiKind(){
  return state.type==='sunny'?'sunny':state.type==='infant'?'infant':state.level;
}

function signatureConfigForUi(){
  return getSignatureConfig(signatureUiKind(),signatureUiPageIndex());
}

function updateSignatureField(role,field,value){
  const cfg=signatureConfigForUi();
  if(field==='bold'||field==='italic'||field==='underline'||field==='digitalEnabled')cfg[role][field]=!!value;
  else if(field==='align')cfg[role][field]=value;
  else if(field.startsWith('digital')){
    if(!cfg[role].digital)cfg[role].digital={x:cfg[role].x,y:cfg[role].y-30,w:90,h:28};
    const map={digitalX:'x',digitalY:'y',digitalW:'w',digitalH:'h'};
    cfg[role].digital[map[field]]=Number(value);
  }else cfg[role][field]=Number(value);
  saveSignatureStore();
  markDirty();
}

function scheduleSignaturePreview(immediate=false){
  clearTimeout(signaturePreviewTimer);
  const run=()=>{
    state.generated.clear();
    $('downloadZipBtn').disabled=true;
    if(state.selected)refreshPreview(state.selected);
  };
  if(immediate)run();else signaturePreviewTimer=setTimeout(run,180);
}

function markDirty(){hasUnsavedWork=true;}
function showLeaveModal(){
  const modal=$('leaveModal'); if(!modal)return;
  modal.classList.add('show'); modal.setAttribute('aria-hidden','false');
  setTimeout(()=>$('leaveCancelBtn')?.focus(),30);
}
function hideLeaveModal(){
  const modal=$('leaveModal'); if(!modal)return;
  modal.classList.remove('show'); modal.setAttribute('aria-hidden','true');
}
function confirmLeave(){
  hideLeaveModal();
  allowUnload=true;
  hasUnsavedWork=false;
  if(pendingNavigation){ pendingNavigation=false; history.back(); }
  else location.reload();
}

function renderBiodataEditor(){
  const box=$('biodataEditor');
  if(!box)return;
  const cfg=getBiodataConfig();
  box.innerHTML=`
    <div class="biodata-control-grid">
      <label>Geser X <input class="sig-range" type="range" min="-80" max="80" step="0.1" data-bio-field="x" value="${cfg.x}"><input class="sig-number" type="number" min="-80" max="80" step="0.1" data-bio-field="x" value="${cfg.x}"></label>
      <label>Geser Y <input class="sig-range" type="range" min="-80" max="80" step="0.1" data-bio-field="y" value="${cfg.y}"><input class="sig-number" type="number" min="-80" max="80" step="0.1" data-bio-field="y" value="${cfg.y}"></label>
      <label>Ukuran font <input class="sig-number" type="number" min="7" max="24" step="0.1" data-bio-field="size" value="${cfg.size}"></label>
    </div>`;
  box.querySelectorAll('[data-bio-field]').forEach(el=>{
    el.addEventListener('input',()=>{
      const field=el.dataset.bioField;
      updateBiodataField(field,el.value);
      if(el.classList.contains('sig-range')){
        const number=box.querySelector(`.sig-number[data-bio-field="${field}"]`);
        if(number)number.value=el.value;
      }
      if(el.classList.contains('sig-number')){
        const range=box.querySelector(`.sig-range[data-bio-field="${field}"]`);
        if(range)range.value=el.value;
      }
    });
  });
}

function signatureControl(role,label,cfg){
  const r=cfg[role];
  return `
    <div class="signature-card">
      <div class="signature-card-title"><b>${label}</b><span>${role==='date'?'Tanggal':'Nama'}</span></div>
      <div class="signature-control-grid">
        <label>Posisi X <input class="sig-range" type="range" min="0" max="600" step="0.1" data-role="${role}" data-field="x" value="${r.x}"><input class="sig-number" type="number" min="0" max="600" step="0.1" data-role="${role}" data-field="x" value="${r.x}"></label>
        <label>Posisi Y <input class="sig-range" type="range" min="0" max="1000" step="0.1" data-role="${role}" data-field="y" value="${r.y}"><input class="sig-number" type="number" min="0" max="1000" step="0.1" data-role="${role}" data-field="y" value="${r.y}"></label>
        <label>Ukuran font <input class="sig-number sig-size" type="number" min="7" max="24" step="0.1" data-role="${role}" data-field="size" value="${r.size}"></label>
      </div>
      <div class="signature-style-row">
        <label class="sig-check"><input type="checkbox" data-role="${role}" data-field="bold" ${r.bold?'checked':''}><span>Bold</span></label>
        <label class="sig-check"><input type="checkbox" data-role="${role}" data-field="italic" ${r.italic?'checked':''}><span>Italic</span></label>
        <label class="sig-check"><input type="checkbox" data-role="${role}" data-field="underline" ${r.underline?'checked':''}><span>Garis bawah</span></label>
        <label class="sig-align">Rata <select data-role="${role}" data-field="align"><option value="left" ${r.align==='left'?'selected':''}>Kiri</option><option value="center" ${r.align==='center'?'selected':''}>Tengah</option><option value="right" ${r.align==='right'?'selected':''}>Kanan</option></select></label>
      </div>
      ${role==='principal'?`<div class="digital-signature-box">
        <div class="digital-signature-head"><b>Tanda tangan digital pimpinan</b><span id="signatureFileStatus">${state.principalSignature?'✓ File siap digunakan':'Belum ada file'}</span></div>
        <div class="digital-signature-actions"><button type="button" class="btn btn-secondary btn-xs" id="uploadSignatureBtn">Upload Tanda Tangan</button><button type="button" class="btn btn-light btn-xs" id="removeSignatureBtn" ${state.principalSignature?'':'disabled'}>Hapus</button><input id="principalSignatureInput" type="file" accept="image/png,image/jpeg,image/webp" hidden></div>
        <label class="sig-check digital-toggle"><input type="checkbox" data-role="principal" data-field="digitalEnabled" ${r.digitalEnabled?'checked':''}><span>Gunakan tanda tangan digital menggantikan tanda tangan fisik</span></label>
        <div class="signature-control-grid digital-grid">
          <label>Posisi X <input class="sig-range" type="range" min="0" max="600" step="0.1" data-role="principal" data-field="digitalX" value="${r.digital.x}"><input class="sig-number" type="number" min="0" max="600" step="0.1" data-role="principal" data-field="digitalX" value="${r.digital.x}"></label>
          <label>Posisi Y <input class="sig-range" type="range" min="0" max="1000" step="0.1" data-role="principal" data-field="digitalY" value="${r.digital.y}"><input class="sig-number" type="number" min="0" max="1000" step="0.1" data-role="principal" data-field="digitalY" value="${r.digital.y}"></label>
          <label>Lebar <input class="sig-range" type="range" min="20" max="250" step="0.1" data-role="principal" data-field="digitalW" value="${r.digital.w}"><input class="sig-number" type="number" min="20" max="250" step="0.1" data-role="principal" data-field="digitalW" value="${r.digital.w}"></label>
          <label>Tinggi <input class="sig-range" type="range" min="10" max="120" step="0.1" data-role="principal" data-field="digitalH" value="${r.digital.h}"><input class="sig-number" type="number" min="10" max="120" step="0.1" data-role="principal" data-field="digitalH" value="${r.digital.h}"></label>
        </div>
      </div>`:''}
    </div>`;
}

function renderSignatureEditor(){
  renderBiodataEditor();
  const select=$('signaturePageSelect');
  const editor=$('signatureEditor');
  if(!select||!editor)return;
  const options=signatureGroupOptions();
  const old=select.value;
  select.innerHTML=options.map(o=>`<option value="${o.value}">${o.label}</option>`).join('');
  if(options.some(o=>o.value===old))select.value=old;
  else if(options[0])select.value=options[0].value;
  const cfg=signatureConfigForUi();
  editor.innerHTML=signatureControl('date','Tanggal',cfg)+signatureControl('teacher','Guru',cfg)+signatureControl('principal','Pimpinan',cfg);

  editor.querySelectorAll('[data-field]').forEach(el=>{
    el.addEventListener('input',()=>{
      const role=el.dataset.role,field=el.dataset.field;
      updateSignatureField(role,field,el.type==='checkbox'?el.checked:el.value);
      if(el.classList.contains('sig-range')){
        const number=editor.querySelector(`.sig-number[data-role="${role}"][data-field="${field}"]`);
        if(number)number.value=el.value;
      }
      if(el.classList.contains('sig-number')){
        const range=editor.querySelector(`.sig-range[data-role="${role}"][data-field="${field}"]`);
        if(range)range.value=el.value;
      }
      scheduleSignaturePreview();
    });
  });
  const uploadBtn=$('uploadSignatureBtn');
  const input=$('principalSignatureInput');
  const removeBtn=$('removeSignatureBtn');
  if(uploadBtn&&input){
    uploadBtn.onclick=()=>input.click();
    input.onchange=async e=>{
      const file=e.target.files?.[0];
      if(!file)return;
      if(!file.type.startsWith('image/')){showMsg('Tanda tangan harus berupa file gambar PNG, JPG, atau WEBP.','error');return;}
      try{
        if(state.principalSignature?.url)URL.revokeObjectURL(state.principalSignature.url);
        state.principalSignature={file,url:URL.createObjectURL(file)};
        renderSignatureEditor();
        scheduleSignaturePreview(true);
        showMsg('Tanda tangan digital pimpinan siap digunakan.','success');
      }catch(err){console.error(err);showMsg('Tanda tangan gagal dimuat.','error');}
      input.value='';
    };
  }
  if(removeBtn){
    removeBtn.onclick=()=>{
      if(state.principalSignature?.url)URL.revokeObjectURL(state.principalSignature.url);
      state.principalSignature=null;
      markDirty();
      const cfg=signatureConfigForUi();
      cfg.principal.digitalEnabled=false;
      saveSignatureStore();
      renderSignatureEditor();
      scheduleSignaturePreview(true);
    };
  }
}

const COMMON_FIELDS={name:'Nama Anak',displayName:'Nama Anak (Nilai)',gender:'Jenis Kelamin',place:'Tempat',birth:'Tanggal Lahir',father:'Nama Ayah',mother:'Nama Ibu',job:'Pekerjaan',address:'Alamat',phone:'No Telp',note:'Catatan',teacher:'Nama Guru'};

const IDENTITY_BOXES=[
  // Kembali ke posisi biodata awal: tetap di baris orange seperti template asli.
  ['name',224.1,358.4,343.8,26.6],['gender',224.1,411.5,343.8,26.6],['place',224.1,464.7,343.8,26.6],['birth',224.1,517.8,343.8,26.6],
  ['father',224.1,570.9,343.8,26.6],['mother',224.1,624.1,343.8,26.6],['job',224.1,677.2,343.8,26.6],['address',224.1,730.3,343.8,26.6],['phone',224.1,783.5,343.8,26.6]
];

const CONFIG={
  lower:{sheet:'LOWER',key:'glowLower',fields:COMMON_FIELDS,results:[
    ['Perawatan Diri Sendiri',2,{x:201.6,y:365.1,w:90.7,h:16.9}],['Perawatan Lingkungan',2,{x:201.6,y:397.9,w:90.7,h:16.9}],['Perkembangan Psikomotorik',2,{x:201.6,y:430.7,w:90.7,h:36.3}],['Sikap',2,{x:201.6,y:482.9,w:90.1,h:16.1}],
    ['Perkembangan Visual',2,{x:201.6,y:570,w:90.1,h:16.9}],['Pengembangan Auditori',2,{x:201.6,y:602.8,w:90.1,h:15.6}],['Pengembangan Gustatori',2,{x:201.6,y:634.3,w:90.1,h:31.2}],['Pengembangan Olfactory',2,{x:201.6,y:681.4,w:90.1,h:31.2}],['Pengembangan Indra Peraba',2,{x:201.6,y:728.5,w:90.1,h:31.2}],
    ['Studi Alam',2,{x:490.2,y:363.5,w:91.6,h:15.9}],['Pengenalan dunia hewan',2,{x:490.2,y:395.3,w:91.6,h:16.9}],['Pengenalan dunia tumbuhan',2,{x:490.2,y:428.1,w:91.6,h:24.7}],['Pengenalan Sejarah',2,{x:490.2,y:471.3,w:91.6,h:24.7}],['Interaksi sosial',2,{x:487.7,y:584.3,w:91.6,h:15.9}],['Emosional',2,{x:487.7,y:616,w:91.6,h:18.8}],
    ['Komunikasi dua arah',3,{x:201.6,y:373.3,w:90.1,h:19.6}],['Gerak & Lagu',3,{x:201.6,y:408.8,w:90.1,h:21.1}],['Pengembangan kosakata',3,{x:201.6,y:445.8,w:90.1,h:21.7}],['Penguasaan fonetik',3,{x:201.6,y:483.3,w:90.1,h:19.7}],['Persiapan menulis',3,{x:201.6,y:518.9,w:90.1,h:21.9}],
    ['Pengenalan pola & bentuk',3,{x:492.8,y:373.3,w:90.1,h:19.6}],['Pengenalan kuantitas 1 - 10',3,{x:492.8,y:408.8,w:90.1,h:20.1}],['Pengenalan simbol 1 -10',3,{x:492.8,y:444.8,w:90.1,h:23.4}]
  ]},
  middle:{sheet:'MIDDLE',key:'glowMiddle',fields:COMMON_FIELDS,results:[
    ['Perawatan Diri Sendiri',2,{x:201.6,y:365.1,w:90.7,h:15.9}],['Perawatan Lingkungan',2,{x:201.6,y:396.8,w:90.7,h:15.9}],['Perkembangan Psikomotorik',2,{x:201.6,y:428.6,w:90.7,h:36.3}],['Sikap',2,{x:201.6,y:480.4,w:90.7,h:16.1}],
    ['Perkembangan Visual',2,{x:201.6,y:565.2,w:90.1,h:15.9}],['Pengembangan Auditori',2,{x:201.6,y:597,w:90.1,h:15.9}],['Pengembangan Gustatori',2,{x:201.6,y:628.8,w:90.1,h:31.2}],['Pengembangan Olfactory',2,{x:201.6,y:675.4,w:90.1,h:31.2}],['Pengembangan Indra Peraba',2,{x:201.6,y:722,w:90.1,h:31.2}],
    ['Studi Alam',2,{x:490.2,y:363.5,w:91.6,h:15.9}],['Pengenalan dunia hewan',2,{x:490.2,y:395.3,w:91.6,h:15.9}],['Pengenalan dunia tumbuhan',2,{x:490.2,y:427.1,w:91.6,h:24.7}],['Pengenalan Sejarah',2,{x:490.2,y:470.3,w:91.6,h:24.6}],['Interaksi sosial',2,{x:487.7,y:584.3,w:91.6,h:15.9}],['Emosional',2,{x:487.7,y:616,w:91.6,h:18.8}],
    ['Komunikasi dua arah',3,{x:201.6,y:377.4,w:90.1,h:23.4}],['Kemampuan bercerita',3,{x:201.6,y:416.7,w:90.1,h:20.2}],['Pengembangan kosakata',3,{x:201.6,y:452.8,w:90.1,h:26.1}],['Penguasaan fonetik',3,{x:201.6,y:494.8,w:90.1,h:25.8}],['Persiapan menulis',3,{x:201.6,y:536.5,w:90.1,h:22.3}],
    ['Pengenalan nilai tempat',3,{x:492.8,y:375.4,w:90.1,h:24.8}],['Pengenalan kuantitas 11-99',3,{x:492.8,y:416.1,w:90.1,h:20.1}],['Pengenalan simbol 11 -99',3,{x:492.8,y:452.1,w:90.1,h:24.9}]
  ]},
  upper:{sheet:'UPPER',key:'glowUpper',fields:COMMON_FIELDS,results:[
    ['Perawatan Diri Sendiri',2,{x:201.6,y:365.1,w:90.7,h:15.6}],['Perawatan Lingkungan',2,{x:201.6,y:396.5,w:90.7,h:17}],['Perkembangan Psikomotorik',2,{x:201.6,y:429.4,w:90.7,h:36.3}],['Sikap',2,{x:201.6,y:481.6,w:90.7,h:16.1}],
    ['Perkembangan Visual',2,{x:201.6,y:565.2,w:90.1,h:16.9}],['Pengembangan Auditori',2,{x:201.6,y:598,w:90.1,h:15.6}],['Pengembangan Gustatori',2,{x:201.6,y:629.5,w:90.1,h:31.2}],['Pengembangan Olfactory',2,{x:201.6,y:676.6,w:90.1,h:31.2}],['Pengembangan Indra Peraba',2,{x:201.6,y:723.7,w:90.1,h:31.2}],
    ['Studi Alam',2,{x:490.2,y:363.5,w:91.6,h:15.9}],['Pengenalan dunia hewan',2,{x:490.2,y:395.3,w:91.6,h:16.9}],['Pengenalan dunia tumbuhan',2,{x:490.2,y:428.1,w:91.6,h:24.7}],['Pengenalan Sejarah',2,{x:490.2,y:471.3,w:91.6,h:24.7}],['Interaksi sosial',2,{x:487.7,y:584.3,w:91.6,h:15.9}],['Emosional',2,{x:487.7,y:616,w:91.6,h:18.8}],
    ['Pemahaman individu & respon',3,{x:201.6,y:373.3,w:90.1,h:36.4}],['Kemampuan mendengarkan & berbicara',3,{x:201.6,y:425.8,w:90.1,h:38.9}],['Pengembangan kosakata',3,{x:201.6,y:480.6,w:90.1,h:28}],['Membaca',3,{x:201.6,y:524.5,w:90.1,h:26.5}],['Menulis',3,{x:201.6,y:566.9,w:90.1,h:25.8}],
    ['Pengoperasian sistem desimal',3,{x:498,y:375,w:90.2,h:33.4}],['Pengenalan matematika abstrak',3,{x:498,y:424.3,w:90.2,h:34.5}]
  ]}
};

const SUNNY={sheet:'RAPOT SUNNY',key:'sunny',fields:{name:'Nama Anak',displayName:'Nama Anak\n(Daftar Penilaian)',gender:'Jenis Kelamin',place:'Tempat',birth:'Tanggal Lahir',father:'Nama Ayah',mother:'Nama Ibu',job:'Pekerjaan',address:'Alamat',phone:'No Telp',note:'Catatan Keseluruhan',teacher:'NAMA GURU'},results:[
  ['Psikomotorik: Keseimbangan & Kasar',2,{x:205.1,y:364.3,w:89,h:24.8}],['Psikomotorik: Halus',2,{x:205.1,y:405,w:89,h:23.6}],['Sosial Emosional :\nKesadaran Diri',2,{x:205.1,y:487.2,w:89,h:16}],['Sosial Emosional:\nInteraksi',2,{x:205.1,y:521.7,w:89,h:16.1}],['Sosial Emosional:\nEmosional',2,{x:205.1,y:553.7,w:89,h:24}],['Latihan Kecakapan Hidup:\nPerawatan Diri',2,{x:490.9,y:365.3,w:95.5,h:15.9}],['Latihan Kecakapan Hidup: \nPerawatan Lingkungan',2,{x:490.9,y:397.1,w:95.5,h:16.9}],['Latihan Kecakapan Hidup: \nKolaborasi dengan orang dewasa',2,{x:490.9,y:429.9,w:95.5,h:24.1}],
  ['Perkembangan Kognitif: \nSensorial',2,{x:490.9,y:505.8,w:90.9,h:17}],['Perkembangan Kognitif: \nPembentukan Konsep',2,{x:490.9,y:540.2,w:90.9,h:16.9}],['Perkembangan Kognitif: \nKoordinasi Mata dan Tangan',2,{x:490.9,y:574.6,w:90.9,h:22.8}],['Bahasa: \nMendengar',2,{x:203.4,y:633.6,w:90.7,h:20.4}],['Bahasa: \nMerespon',2,{x:203.4,y:672.5,w:90.7,h:19.3}],['Bahasa: \nKomunikasi Dua Arah',2,{x:203.4,y:707.7,w:90.7,h:19.9}]
]};

const INFANT={sheet:'FEB',key:'infant',fields:{name:'Nama Anak',displayName:'Nama Anak\n(Daftar Penilaian)',gender:'Jenis Kelamin',place:'Tempat',birth:'Tanggal Lahir',father:'Nama Ayah',mother:'Nama Ibu',job:'Pekerjaan',address:'Alamat',phone:'No Telp',note:'CATATAN',teacher:'NAMA GURU'},results:[
  ['Psikomotor (keseimbangan)',2,{x:397.4,y:343.9,w:169,h:18.8}],['Perkembangan kognitif (Koordinasi Mata & Tangan)',2,{x:397.8,y:471.9,w:169,h:22.3}],['Perkembangan Sensorial (Eksploral Sensorial)',2,{x:397.4,y:539.3,w:169,h:23.6}],['Sosial Emosional (kesadaran diri dan Emosi)',2,{x:397.4,y:581.2,w:169,h:23.9}],['Sosial Emosional (interaksi)',2,{x:397.8,y:646.4,w:169,h:18.4}],['Mendengar',2,{x:397.8,y:683.1,w:169,h:18.3}],['Merespon',2,{x:397.4,y:405.7,w:169,h:21.8}]
]};

const templateCache=new Map();let fonts={};let previewUrl=null;let busy=false;
const $=id=>document.getElementById(id);const clean=v=>{if(v===null||v===undefined)return '';const s=String(v).trim();return /^#(VALUE|REF|DIV\/0|N\/A|NAME|NUM|NULL)!?$/i.test(s)?'':s;};
function norm(s){return clean(s).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\.[^.]+$/,'').replace(/[^a-z0-9]+/g,'');}
function initial(name){return clean(name).split(/\s+/).filter(Boolean).slice(0,2).map(x=>x[0]).join('').toUpperCase()||'?';}
function fmtDate(v){if(!v)return '';if(v instanceof Date&&!Number.isNaN(v.getTime()))return new Intl.DateTimeFormat('id-ID',{day:'2-digit',month:'long',year:'numeric'}).format(v);const s=clean(v);if(/^\d{4}-\d{2}-\d{2}$/.test(s)){const d=new Date(s+'T00:00:00');return new Intl.DateTimeFormat('id-ID',{day:'2-digit',month:'long',year:'numeric'}).format(d);}return s;}
function currentDate(){const d=new Date();return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;}
function showMsg(text,type='info'){const el=$('message');el.textContent=text;el.className='notice show '+type;}function clearMsg(){$('message').className='notice';$('message').textContent='';}
function safeFile(s){return clean(s).replace(/[\\/:*?"<>|]+/g,' ').replace(/\s+/g,' ').trim().slice(0,100)||'Siswa';}function escapeHtml(s){return clean(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
function settings(){return{city:clean($('city').value)||'Purwokerto',date:$('reportDate').value,period:clean($('period').value),className:clean($('className').value)||({sunny:'SUNNY CLASS',infant:'INFANT CLASS'}[state.type]||'GLOW CLASS'),teacher:clean($('teacher').value),principal:clean($('principal').value)||'Michelle Juwono'};}
async function getTemplate(key){if(templateCache.has(key))return templateCache.get(key);const res=await fetch(TEMPLATES[key]);if(!res.ok)throw new Error(`Template ${key} tidak ditemukan.`);const bytes=await res.arrayBuffer();templateCache.set(key,bytes);return bytes;}
async function setupFonts(doc){
  // Font rapor: ABeeZee dari folder assets/fonts.
  // ABeeZee hanya menyediakan Regular + Italic, jadi Regular dipakai untuk seluruh isi.
  if(!window.fontkit) throw new Error('Fontkit belum termuat. Refresh halaman dan pastikan koneksi internet aktif.');
  doc.registerFontkit(window.fontkit);

  const regBytes=await fetch('assets/fonts/ABeeZee-Regular.ttf').then(r=>{
    if(!r.ok) throw new Error('ABeeZee Regular tidak ditemukan.');
    return r.arrayBuffer();
  });
  const italicBytes=await fetch('assets/fonts/ABeeZee-Italic.ttf').then(r=>{
    if(!r.ok) throw new Error('ABeeZee Italic tidak ditemukan.');
    return r.arrayBuffer();
  });

  fonts.reg=await doc.embedFont(regBytes,{subset:true});
  // Tidak menggunakan italic sebagai bold agar bentuk tulisan tetap sesuai font template.
  fonts.bold=fonts.reg;
  fonts.italic=await doc.embedFont(italicBytes,{subset:true});
}
function pdfY(page,top,h){return page.getHeight()-top-h;}
function fitSize(text,maxWidth,size,font){let fs=size;while(fs>7&&font.widthOfTextAtSize(text,fs)>maxWidth)fs-=.25;return fs;}
function drawText(page,text,x,top,w,size=12,color=COLORS.black,align='left',bold=false){drawStyledText(page,text,x,top,w,size,color,align,{bold});}

function drawStyledText(page,text,x,top,w,size=12,color=COLORS.black,align='left',style={}){
  text=clean(text);
  if(!text)return;
  const italic=!!style.italic;
  const bold=!!style.bold;
  const underline=!!style.underline;
  const font=italic?fonts.italic:fonts.reg;
  const fs=style.fit===false?size:fitSize(text,w,size,font);
  const width=font.widthOfTextAtSize(text,fs);
  let xx=x;
  if(align==='center')xx=x+(w-width)/2;
  if(align==='right')xx=x+w-width;
  const yy=page.getHeight()-top-fs;
  const options={x:xx,y:yy,size:fs,font,color};
  page.drawText(text,options);
  // ABeeZee Regular/Italic has no bold face in the bundled files, so emulate a light bold.
  if(bold){
    page.drawText(text,{...options,x:xx+0.35});
    page.drawText(text,{...options,x:xx+0.7});
  }
  if(underline){
    const lineY=yy-1.3;
    page.drawLine({start:{x:xx,y:lineY},end:{x:xx+width,y:lineY},thickness:0.7,color});
  }
}
function wrapLines(text,w,size,font){const out=[];for(const paragraph of clean(text).split(/\n/)){let line='';for(const word of paragraph.split(/\s+/)){const test=line?`${line} ${word}`:word;if(font.widthOfTextAtSize(test,size)<=w)line=test;else{if(line)out.push(line);line=word;}}if(line)out.push(line);}return out;}
function drawMultiline(page,text,x,top,w,h,size=11,color=COLORS.black,align='left',bold=false){const font=bold?fonts.bold:fonts.reg;let fs=size;let lines=wrapLines(text,w,fs,font);while(lines.length>Math.max(1,Math.floor(h/(fs*1.18)))&&fs>7){fs-=.3;lines=wrapLines(text,w,fs,font);}const lh=fs*1.18;lines.slice(0,Math.floor(h/lh)).forEach((line,i)=>drawText(page,line,x,top+i*lh,w,fs,color,align,bold));}
function writeTransparent(page,box,text,opt={}){if(!clean(text))return;if(opt.center)drawText(page,text,box.x,box.y+(box.h-(opt.size||12))/2,box.w,opt.size||12,opt.color||COLORS.black,'center',!!opt.bold);else drawMultiline(page,text,box.x+2,box.y+2,box.w-4,box.h-4,opt.size||12,opt.color||COLORS.black,opt.align||'left',!!opt.bold);}
function fillRect(page,box,color=COLORS.orange){page.drawRectangle({x:box.x,y:page.getHeight()-box.y-box.h,width:box.w,height:box.h,color,borderWidth:0});}

function photoFor(student){if(!student)return null;if(state.manualPhotos.has(student.id))return state.manualPhotos.get(student.id);const candidates=[student.name,student.displayName].map(norm).filter(Boolean);for(const c of candidates)if(state.photos.has(c))return state.photos.get(c);const matches=[];for(const [key,val] of state.photos){for(const c of candidates)if(key.includes(c)||c.includes(key))matches.push({score:Math.max(key.length,c.length),val});}matches.sort((a,b)=>b.score-a.score);return matches[0]?.val||null;}
function parseRows(ws){return XLSX.utils.sheet_to_json(ws,{defval:'',raw:true});}
function getFieldValue(row,header){if(Object.prototype.hasOwnProperty.call(row,header))return row[header];const target=norm(header);const key=Object.keys(row).find(k=>norm(k)===target);return key?row[key]:'';}
function findWorksheet(cfg){const direct=state.workbook.Sheets[cfg.sheet];if(direct)return direct;const wanted=Object.values(cfg.fields).slice(0,10).map(norm);let best=null,bestScore=-1;for(const name of state.workbook.SheetNames){const ws=state.workbook.Sheets[name];const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:true});const header=(rows[0]||[]).map(norm);const score=wanted.filter(x=>x&&header.includes(x)).length;if(score>bestScore){bestScore=score;best=ws;}}if(bestScore>=2)return best;return null;}
function parseWorkbook(){if(!state.workbook)return;if(!window.XLSX){showMsg('Library Excel belum termuat.','error');return;}const cfg=state.type==='sunny'?SUNNY:state.type==='infant'?INFANT:CONFIG[state.level];const ws=findWorksheet(cfg);if(!ws){showMsg(`Sheet data untuk ${state.type==='sunny'?'Sunny':state.type==='infant'?'Infant':state.level} tidak ditemukan.`,'error');return;}const rows=parseRows(ws);state.students=rows.filter(r=>clean(getFieldValue(r,cfg.fields.name))).map((r,index)=>{const obj={raw:r,index,id:`${state.type}-${state.level}-${index}`};for(const [k,h] of Object.entries(cfg.fields))obj[k]=getFieldValue(r,h);obj.name=clean(obj.name);obj.displayName=clean(obj.displayName);obj.note=clean(obj.note);obj.teacher=clean(obj.teacher);return obj;});state.selected=state.students[0]||null;state.generated.clear();state.manualPhotos.clear();$('downloadZipBtn').disabled=true;updateUI();}
function expectedTemplateLabel(){if(state.type==='sunny')return 'Sunny';if(state.type==='infant')return 'Infant';return `Glow ${state.level[0].toUpperCase()+state.level.slice(1)}`;}
function detectWorkbookTemplate(wb){
  const names=wb?.SheetNames||[];
  const normNames=names.map(n=>({raw:n,n:norm(n).replace(/\s+/g,'')}));
  const known=[
    {label:'Glow Lower',keys:['lower']},
    {label:'Glow Middle',keys:['middle']},
    {label:'Glow Upper',keys:['upper']},
    {label:'Sunny',keys:['rapotsunny','sunny']},
    {label:'Infant',keys:['feb','infant']}
  ];
  for(const item of known){if(normNames.some(x=>item.keys.includes(x.n)))return item.label;}
  // Fallback: inspect headers when sheet names are custom/renamed.
  const signatures=[
    {label:'Sunny',terms:['latihan kecakapan hidup','komunikasi dua arah','pembentukan konsep']},
    {label:'Infant',terms:['eksplorasi sensorial','keseimbangan','koordinasi mata dan tangan']},
    {label:'Glow Lower',terms:['gerak dan lagu','pengenalan kuantitas 1 - 10']},
    {label:'Glow Middle',terms:['kemampuan bercerita','pengenalan kuantitas 11 - 99']},
    {label:'Glow Upper',terms:['pemahaman individu dan respon','matematika abstrak']}
  ];
  let best=null,bestScore=0;
  for(const sh of names){
    const ws=wb.Sheets[sh];
    const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:true});
    const text=rows.slice(0,8).flat().map(v=>norm(v)).join(' | ');
    for(const sig of signatures){const score=sig.terms.filter(t=>text.includes(norm(t))).length;if(score>bestScore){bestScore=score;best=sig.label;}}
  }
  return best;
}
function showTemplateMismatch(expected,detected,fileName){
  const modal=$('templateModal');if(!modal)return;
  $('templateExpected').textContent=expected;
  $('templateDetected').textContent=detected||'Format tidak dikenali';
  $('templateModalText').textContent=detected?`File “${fileName}” terdeteksi sebagai format ${detected}, bukan ${expected}.`:`File “${fileName}” tidak terlihat menggunakan struktur Excel rapor yang didukung.`;
  modal.classList.add('show');modal.setAttribute('aria-hidden','false');
}
function hideTemplateMismatch(){const modal=$('templateModal');if(!modal)return;modal.classList.remove('show');modal.setAttribute('aria-hidden','true');}
async function loadExcel(file){
  clearMsg();
  if(!window.XLSX){showMsg('Library Excel belum termuat. Refresh halaman dan pastikan internet aktif saat pertama membuka website.','error');return;}
  $('excelInfo').textContent=`Memeriksa ${file.name}…`;
  try{
    const buf=await file.arrayBuffer();
    const wb=XLSX.read(buf,{type:'array',cellDates:true});
    const expected=expectedTemplateLabel();
    const detected=detectWorkbookTemplate(wb);
    if(detected && detected!==expected){
      showTemplateMismatch(expected,detected,file.name);
      $('excelInfo').textContent='Belum ada file Excel yang sesuai.';
      $('excelInfo').classList.remove('ok');
      return;
    }
    if(!detected){
      showTemplateMismatch(expected,null,file.name);
      $('excelInfo').textContent='Format Excel belum dikenali.';
      $('excelInfo').classList.remove('ok');
      return;
    }
    markDirty();
    state.workbook=wb;
    parseWorkbook();
    $('excelInfo').textContent=`✓ ${file.name}`;
    $('excelInfo').classList.add('ok');
  }catch(err){console.error(err);showMsg('Excel gagal dibaca. Pastikan file .xlsx/.xls yang valid.','error');$('excelInfo').textContent='Excel gagal dibaca.';$('excelInfo').classList.remove('ok');}
}
function loadPhotos(files){markDirty();state.photos.clear();for(const f of files){if(!f.type.startsWith('image/'))continue;const key=norm(f.name);if(key)state.photos.set(key,{file:f,url:URL.createObjectURL(f)});}$('photoInfo').textContent=files.length?`✓ ${files.length} foto siap dicocokkan`:'Belum ada foto.';$('photoInfo').classList.toggle('ok',!!files.length);if(state.students.length)updateUI();}
function updateUI(){if(!state.students.length){$('emptyState').style.display='grid';$('result').classList.remove('show');$('generateBtn').disabled=true;return;}$('emptyState').style.display='none';$('result').classList.add('show');$('generateBtn').disabled=false;$('statStudents').textContent=state.students.length;$('statPhotos').textContent=state.students.filter(s=>photoFor(s)).length;$('statTemplate').textContent=state.type==='sunny'?'Sunny':state.type==='infant'?'Infant':`Glow ${state.level[0].toUpperCase()+state.level.slice(1)}`;$('studentCount').textContent=`${state.students.length} siswa`;renderStudents();if(state.selected)selectStudent(state.selected,false);}
function renderStudents(){const list=$('studentList');list.innerHTML='';for(const s of state.students){const p=photoFor(s);const b=document.createElement('button');b.className='student'+(state.selected===s?' active':'');b.type='button';b.innerHTML=`<span class="avatar">${p?`<img src="${p.url}" alt="">`:initial(s.name)}</span><span class="student-info"><b>${escapeHtml(s.name)}</b><small>${state.type==='sunny'?'SUNNY':state.type==='infant'?'INFANT':state.level.toUpperCase()}</small></span><span class="student-status ${p?'ok':''}">${p?'✓':'○'}</span><span class="student-photo" title="Tambah/ganti foto">📷</span>`;b.onclick=()=>selectStudent(s,true);b.querySelector('.student-photo').onclick=e=>{e.stopPropagation();chooseManualPhoto(s);};list.appendChild(b);}}
function chooseManualPhoto(student){if(!student)return;state.selected=student;renderStudents();$('manualPhotoInput').value='';$('manualPhotoInput').click();}
function setPreview(blob,student){
  if(previewUrl){URL.revokeObjectURL(previewUrl);previewUrl=null;}
  previewUrl=URL.createObjectURL(blob);
  const frame=$('previewFrame');
  frame.innerHTML=`<object data="${previewUrl}#page=1&zoom=page-width" type="application/pdf" aria-label="Preview rapor"><div class="preview-fallback">Preview PDF tidak tampil di browser ini. <a href="${previewUrl}" target="_blank" rel="noopener">Buka PDF</a></div></object>`;
  $('previewMeta').textContent=photoFor(student)?'Foto ditemukan.':'Foto belum ditemukan — klik 📷 untuk menambahkan.';
}
async function refreshPreview(student){$('previewTitle').textContent=student.name;$('previewMeta').textContent='Membuat preview…';$('manualPhotoBtn').disabled=false;$('downloadOneBtn').disabled=true;try{const blob=await generatePdf(student);await setPreview(blob,student);state.generated.set(student.id,blob);$('downloadOneBtn').disabled=false;}catch(err){console.error(err);$('previewMeta').textContent='PDF belum bisa dibuat';showMsg(err.message||'Gagal membuat PDF.','error');}}
function selectStudent(student,rerender=true){state.selected=student;if(rerender)renderStudents();refreshPreview(student);}

function drawIdentity(page,student){const vals={name:student.name,gender:student.gender,place:student.place,birth:fmtDate(student.birth),father:student.father,mother:student.mother,job:student.job,address:student.address,phone:student.phone};const cfg=getBiodataConfig();for(const [key,x,y,w,h] of IDENTITY_BOXES)writeTransparent(page,{x:x+cfg.x,y:y+cfg.y,w,h},vals[key],{size:cfg.size});}
function drawHeader(page,student,set){const name=clean(student.displayName)||student.name;const infant=state.type==='infant';const bars=[{x:312.3,y:235,w:infant?255:269.5,h:infant?23.3:20.6,label:'NAMA',value:name},{x:312.3,y:262.3,w:infant?255:269.5,h:infant?22.2:20.6,label:'PERIODE / TA',value:set.period},{x:312.3,y:289.6,w:infant?255:269.5,h:infant?21.5:20.6,label:'KELAS',value:set.className},{x:397.4,y:205.5,w:infant?169.8:185.6,h:infant?19.2:18.7,label:'CABANG',value:set.city}];for(const b of bars){fillRect(page,b);drawText(page,b.label,b.x+3,b.y+3,110,12,COLORS.white,'left',true);drawText(page,':',b.x+(b.label==='CABANG'?58:112),b.y+3,12,12,COLORS.white,'left',true);const vx=b.x+(b.label==='CABANG'?66:118);drawText(page,b.value,vx,b.y+3,b.w-(vx-b.x)-5,12,COLORS.white,'left',true);}}
async function drawPhoto(page,photo,doc){if(!photo)return;const bytes=new Uint8Array(await photo.file.arrayBuffer());let img;try{img=await doc.embedJpg(bytes);}catch{try{img=await doc.embedPng(bytes);}catch{const png=await imageFileToPng(photo.file);img=await doc.embedPng(png);}}const box={x:146.6,y:218.9,w:61.1,h:84};const scale=Math.min(box.w/img.width,box.h/img.height);const w=img.width*scale,h=img.height*scale;page.drawImage(img,{x:box.x+(box.w-w)/2,y:page.getHeight()-box.y-box.h+(box.h-h)/2,width:w,height:h});}
async function imageFileToPng(file){const bitmap=await createImageBitmap(file);const canvas=document.createElement('canvas');canvas.width=bitmap.width;canvas.height=bitmap.height;canvas.getContext('2d').drawImage(bitmap,0,0);bitmap.close();const blob=await new Promise(resolve=>canvas.toBlob(resolve,'image/png'));if(!blob)throw new Error('Foto tidak dapat diproses. Gunakan JPG atau PNG.');return new Uint8Array(await blob.arrayBuffer());}
function noteBox(kind,pageIndex){
  if(kind==='infant'){
    return {x:45.8,y:746.3,w:521.4,h:69.2};
  }

  if(kind==='sunny'){
    return {x:299.8,y:647.5,w:294.8,h:106};
  }

  // Glow Middle
  if(state.level === 'middle'){
    return {x:20,y:635,w:562,h:133};
  }

  // Glow Upper
  if(state.level === 'upper'){
    return {x:20,y:640,w:562,h:133};
  }

  // Glow Lower
  return {x:20,y:615,w:562,h:133};
}

function signaturePositions(kind,pageIndex){
  const cfg=getSignatureConfig(kind,pageIndex);
  return {date:cfg.date,teacher:cfg.teacher,principal:cfg.principal};
}


async function embedSignatureImage(doc,file){
  const bytes=new Uint8Array(await file.arrayBuffer());
  try{return await doc.embedPng(bytes);}catch{
    try{return await doc.embedJpg(bytes);}catch{
      const png=await imageFileToPng(file);
      return await doc.embedPng(png);
    }
  }
}

async function drawDigitalSignature(page,img,box){
  if(!img||!box)return;
  const scale=Math.min(box.w/img.width,box.h/img.height);
  const w=img.width*scale,h=img.height*scale;
  page.drawImage(img,{x:box.x+(box.w-w)/2,y:page.getHeight()-box.y-box.h+(box.h-h)/2,width:w,height:h});
}

async function generatePdf(student){getLibraries();const cfg=state.type==='sunny'?SUNNY:state.type==='infant'?INFANT:CONFIG[state.level];const bytes=await getTemplate(cfg.key);const doc=await PDFDocument.load(bytes);await setupFonts(doc);const set=settings();const pages=doc.getPages();const kind=state.type==='sunny'?'sunny':state.type==='infant'?'infant':state.level;drawIdentity(pages[0],student);const photo=photoFor(student);for(let i=1;i<pages.length;i++){drawHeader(pages[i],student,set);if(i===1)await drawPhoto(pages[i],photo,doc);for(const [header,pageNo,box] of cfg.results){if(pageNo===i+1){const val=clean(getFieldValue(student.raw,header));writeTransparent(pages[i],box,val,{size:12,center:true});}}const sig=signaturePositions(kind,i+1);const dateText=set.date?fmtDate(set.date):'';if(dateText)drawStyledText(pages[i],dateText,sig.date.x,sig.date.y,sig.date.w,sig.date.size,COLORS.orange,sig.date.align,{...sig.date,fit:false});const teacher=clean(set.teacher)||student.teacher;if(teacher)drawStyledText(pages[i],teacher,sig.teacher.x,sig.teacher.y,sig.teacher.w,sig.teacher.size,COLORS.orange,sig.teacher.align,sig.teacher);
    if(set.principal){
      if(sig.principal.digitalEnabled&&state.principalSignature?.file){
        const sigImg=await embedSignatureImage(doc,state.principalSignature.file);
        await drawDigitalSignature(pages[i],sigImg,sig.principal.digital);
      }
      drawStyledText(pages[i],set.principal,sig.principal.x,sig.principal.y,sig.principal.w,sig.principal.size,COLORS.orange,sig.principal.align,sig.principal);
    }}
const note=clean(student.note);if(note){const p=pages[pages.length-1];const b=noteBox(kind,pages.length-1);writeTransparent(p,b,note,{size:10.5});}return new Blob([await doc.save()],{type:'application/pdf'});}

async function generateAll(showSuccess=true){if(!state.students.length||busy)return false;busy=true;state.generated.clear();$('progress').classList.add('show');$('generateBtn').disabled=true;$('downloadZipBtn').disabled=true;$('statStatus').textContent='Membuat…';$('progressBar').style.width='0%';$('progressPct').textContent='0%';const errors=[];try{for(let i=0;i<state.students.length;i++){const s=state.students[i];$('progressText').textContent=`Membuat ${i+1}/${state.students.length}: ${s.name}`;try{state.generated.set(s.id,await generatePdf(s));}catch(err){console.error('Gagal membuat',s.name,err);errors.push(`${s.name}: ${err.message||'gagal membuat PDF'}`);}const pct=Math.round((i+1)/state.students.length*100);$('progressBar').style.width=pct+'%';$('progressPct').textContent=pct+'%';}const ok=state.generated.size;$('progressText').textContent=errors.length?`${ok} selesai, ${errors.length} gagal.`:'Semua rapor selesai.';$('statStatus').textContent=errors.length?`${ok} berhasil`:'Selesai';$('downloadZipBtn').disabled=!ok;if(showSuccess){if(errors.length)showMsg(`${ok} rapor berhasil dibuat. ${errors.length} siswa gagal: ${errors.slice(0,3).join(' | ')}`,ok?'success':'error');else showMsg(`${ok} rapor berhasil dibuat.`,'success');}return ok>0;}finally{busy=false;$('generateBtn').disabled=false;}}
async function downloadZip(){try{getLibraries();}catch(err){showMsg(err.message,'error');return;}if(!state.students.length){showMsg('Upload Excel terlebih dahulu.','info');return;}if(state.generated.size!==state.students.length){const ok=await generateAll(false);if(!ok){showMsg('Tidak ada raport yang berhasil dibuat.','error');return;}}const zip=new JSZip();for(const s of state.students){const blob=state.generated.get(s.id);if(blob)zip.file(`${safeFile(s.name)}.pdf`,blob);}const out=await zip.generateAsync({type:'blob',compression:'DEFLATE',compressionOptions:{level:6}},meta=>{const pct=Math.round(meta.percent);$('progress').classList.add('show');$('progressText').textContent='Mengemas ZIP…';$('progressPct').textContent=pct+'%';$('progressBar').style.width=pct+'%';});saveBlob(out,`Rapor_${state.type==='sunny'?'Sunny':state.type==='infant'?'Infant':`Glow_${state.level}`}_Semua.zip`);showMsg(`${state.generated.size} raport berhasil dikemas ke ZIP.`,'success');}
function saveBlob(blob,name){const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(a.href),1500);}

loadSignatureStore();
loadBiodataStore();
renderSignatureEditor();
$('signaturePageSelect')?.addEventListener('change',()=>renderSignatureEditor());
$('resetBiodataBtn')?.addEventListener('click',resetBiodataConfig);
$('resetSignatureBtn')?.addEventListener('click',()=>{const kind=signatureUiKind();resetSignatureConfig(kind,signatureUiPageIndex());});

$('reportDate').value=currentDate();$('excelBtn').onclick=()=>$('excelInput').click();$('photoBtn').onclick=()=>$('photoInput').click();$('excelInput').onchange=e=>{if(e.target.files[0])loadExcel(e.target.files[0]);e.target.value='';};$('photoInput').onchange=e=>{loadPhotos([...e.target.files]);e.target.value='';};$('manualPhotoBtn').onclick=()=>{if(state.selected)chooseManualPhoto(state.selected);};$('manualPhotoInput').onchange=e=>{const f=e.target.files[0];if(!f||!state.selected)return;const old=state.manualPhotos.get(state.selected.id);if(old?.url)URL.revokeObjectURL(old.url);state.manualPhotos.set(state.selected.id,{file:f,url:URL.createObjectURL(f)});markDirty();state.generated.delete(state.selected.id);$('photoInfo').textContent=`✓ Foto manual untuk ${state.selected.name}`;$('photoInfo').classList.add('ok');renderStudents();refreshPreview(state.selected);e.target.value='';};$('generateBtn').onclick=()=>generateAll(true);$('downloadZipBtn').onclick=downloadZip;$('downloadOneBtn').onclick=async()=>{if(!state.selected)return;try{const blob=state.generated.get(state.selected.id)||await generatePdf(state.selected);state.generated.set(state.selected.id,blob);saveBlob(blob,`${safeFile(state.selected.name)}.pdf`);}catch(e){showMsg(e.message||'Gagal membuat PDF.','error');}};$('resetBtn').onclick=()=>{if(hasUnsavedWork)showLeaveModal();else location.reload();};
['city','reportDate','period','className','teacher','principal'].forEach(id=>$(id).addEventListener('input',()=>{markDirty();state.generated.clear();$('downloadZipBtn').disabled=true;if(state.selected)refreshPreview(state.selected);}));
// Lindungi data saat refresh / tutup tab. Browser memang hanya mengizinkan dialog native untuk beforeunload.
window.addEventListener('beforeunload',e=>{
  if(!hasUnsavedWork||allowUnload)return;
  e.preventDefault(); e.returnValue='';
});

// Tombol Back memakai modal custom yang lebih rapi.
try{
  history.pushState({raporGeneratorGuard:true},'',location.href);
  window.addEventListener('popstate',()=>{
    if(!hasUnsavedWork){ return; }
    pendingNavigation=true;
    history.pushState({raporGeneratorGuard:true},'',location.href);
    showLeaveModal();
  });
}catch(err){console.warn('Proteksi tombol Back tidak tersedia.',err);}

$('leaveCancelBtn')?.addEventListener('click',()=>{pendingNavigation=false;hideLeaveModal();});
$('leaveConfirmBtn')?.addEventListener('click',confirmLeave);
$('templateOkBtn')?.addEventListener('click',hideTemplateMismatch);
$('templateModal')?.querySelector('[data-template-cancel]')?.addEventListener('click',hideTemplateMismatch);

$('leaveModal')?.querySelector('[data-leave-cancel]')?.addEventListener('click',()=>{pendingNavigation=false;hideLeaveModal();});

document.addEventListener('keydown',e=>{
  if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='r'&&hasUnsavedWork){e.preventDefault();pendingNavigation=false;showLeaveModal();return;}
  if(e.key==='Escape'&&$('leaveModal')?.classList.contains('show')){pendingNavigation=false;hideLeaveModal();}
});

function checkLibraries(){const missing=[];if(!window.XLSX)missing.push('Excel (SheetJS)');if(!window.PDFLib)missing.push('PDF (pdf-lib)');if(!window.JSZip)missing.push('ZIP (JSZip)');if(!window.fontkit)missing.push('fontkit (font custom)');if(!window.XLSX)showMsg('Library Excel belum termuat. Refresh halaman dan pastikan internet aktif saat pertama membuka website.','error');else if(missing.length)showMsg(`Excel siap. ${missing.includes('fontkit (font custom)')?'Font custom belum termuat; PDF tetap bisa dibuat dengan font fallback. ':''}${missing.filter(x=>x!=='fontkit (font custom)').length?'Library: '+missing.filter(x=>x!=='fontkit (font custom)').join(', '):''}`,'info');}
window.addEventListener('load',()=>setTimeout(checkLibraries,50));
function workbookMatchesSelection(){if(!state.workbook)return true;const detected=detectWorkbookTemplate(state.workbook);return !!detected&&detected===expectedTemplateLabel();}
for(const btn of document.querySelectorAll('.type-card'))btn.onclick=()=>{
  const nextType=btn.dataset.type;
  const previousType=state.type;
  state.type=nextType;
  const matches=workbookMatchesSelection();
  state.type=previousType;
  if(!matches){showTemplateMismatch(nextType==='sunny'?'Sunny':nextType==='infant'?'Infant':`Glow ${state.level[0].toUpperCase()+state.level.slice(1)}`,detectWorkbookTemplate(state.workbook),'file Excel yang sedang dipakai');return;}
  document.querySelectorAll('.type-card').forEach(x=>x.classList.remove('active'));btn.classList.add('active');state.type=nextType;$('glowLevels').classList.toggle('show',state.type==='glow');$('className').value=state.type==='sunny'?'SUNNY CLASS':state.type==='infant'?'INFANT CLASS':'GLOW CLASS';clearMsg();renderSignatureEditor();if(state.workbook)parseWorkbook();
};
for(const btn of document.querySelectorAll('.seg button'))btn.onclick=()=>{
  const nextLevel=btn.dataset.level;
  const previousLevel=state.level;
  state.level=nextLevel;
  const matches=workbookMatchesSelection();
  state.level=previousLevel;
  if(!matches){showTemplateMismatch(`Glow ${nextLevel[0].toUpperCase()+nextLevel.slice(1)}`,detectWorkbookTemplate(state.workbook),'file Excel yang sedang dipakai');return;}
  document.querySelectorAll('.seg button').forEach(x=>x.classList.remove('active'));btn.classList.add('active');state.level=nextLevel;$('className').value='GLOW CLASS';clearMsg();renderSignatureEditor();if(state.workbook)parseWorkbook();
};
function wireDrop(id,handler){const el=$(id);for(const ev of ['dragenter','dragover'])el.addEventListener(ev,e=>{e.preventDefault();el.classList.add('drag');});for(const ev of ['dragleave','drop'])el.addEventListener(ev,e=>{e.preventDefault();el.classList.remove('drag');});el.addEventListener('drop',e=>handler(e.dataTransfer.files));}
wireDrop('excelDrop',files=>{const f=[...files].find(x=>/\.xlsx?$|\.xlsm$/i.test(x.name));if(f)loadExcel(f);});wireDrop('photoDrop',files=>loadPhotos([...files].filter(f=>f.type.startsWith('image/'))));
