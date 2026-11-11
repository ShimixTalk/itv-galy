(function(){
'use strict';
let pending=[];
const clean=s=>String(s??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();
const aliases={site:['gare','station','arret','nom gare','nom station','site'],line:['ligne','line','numero ligne','n ligne'],mode:['mode','reseau','transport','type transport'],local:['local','locaux','nom local','type local','localisation','emplacement','lieu'],name:['equipement','equipment','designation','description','nom equipement','nom','materiel'],reference:['reference','ref','identifiant','code','numero'],category:['categorie','category','famille'],sub:['sous categorie','sub'],notes:['observations','commentaires','remarques','notes'],access:['acces','guidage'],phone:['telephone','tel'],supply:['alimentation','disjoncteur']};
const headerMap=row=>row.map(h=>{let n=clean(h);return Object.keys(aliases).find(k=>aliases[k].includes(n))||''});
function parseTechnical(rows){
 let line='',headers=[],out=[],site='',stationCode='';
 for(let i=0;i<rows.length;i++){
  let row=rows[i].map(v=>String(v??'').trim()),first=row[0]||'';
  let m=first.match(/^LIGNE\s*0?(\d{1,2})\s*$/i);
  if(m){line=m[1].padStart(2,'0');headers=[];site='';continue}
  if(/^station$/i.test(first)){headers=row.slice(1).map(s=>s.trim());site='';continue}
  if(!line||!headers.length)continue;
  if(!first){if(site&&row.slice(1).some(Boolean)){let note=row.slice(1).filter(Boolean).join(' | ');out.filter(x=>x.site===site&&x.line===line).slice(-headers.length).forEach(x=>x.notes=[x.notes,note].filter(Boolean).join(' | '))}continue}
  if(/^\s*(?:total|remarque|observation)/i.test(first))continue;
  let sm=first.match(/^(.*?)\s*\[([^\]]+)\]\s*$/);site=sm?sm[1].trim():first;stationCode=sm?sm[2]:'';
  if(!site||!row.slice(1).some(Boolean))continue;
  headers.forEach((type,j)=>{let location=row[j+1];if(!location||!type)return;
   out.push({site,stationCode,line,mode:'Métro',name:type,category:'Locaux techniques',sub:type,local:location,access:location,reference:stationCode,notes:'Source : listing technique, ligne '+line, _row:i+1})
  });
 }
 return out
}
function parseRows(rows){
 if(!rows.length)return [];
 if(rows.some(r=>/^LIGNE\s*\d{1,2}\s*$/i.test(String(r[0]||'').trim()))&&rows.some(r=>/^station$/i.test(String(r[0]||'').trim())))return parseTechnical(rows);
 let best=-1,score=-1;rows.slice(0,25).forEach((r,i)=>{let count=headerMap(r).filter(Boolean).length;if(count>score){score=count;best=i}});
 if(score<2)throw Error('Colonnes non reconnues : utilisez des en-têtes Gare/Station et Local/Équipement.');
 let headers=headerMap(rows[best]);return rows.slice(best+1).map((r,i)=>{let o={};headers.forEach((k,j)=>{if(k&&r[j]!=null&&!o[k])o[k]=String(r[j]).trim()});return {...o,site:o.site||'',name:o.name||o.local||'',_row:i+best+2}}).filter(o=>o.site&&o.name)
}
function key(r){return [clean(r.site),clean(r.name),clean(r.local),clean(r.reference),clean(r.line)].join('|')}
function preview(){let root=document.getElementById('smartPreview');if(!root)return;let seen=new Set(db.records.map(key)),unique=[];let skipped=0;for(let r of pending){let k=key(r);if(seen.has(k)){skipped++;continue}seen.add(k);unique.push(r)}window.smartReady=unique;root.innerHTML='<div class="item"><b>'+unique.length+' fiche(s) à créer</b> · '+skipped+' doublon(s) ignoré(s)<p class="small">'+unique.slice(0,8).map(r=>esc(r.site+' → '+(r.local||r.name))).join('<br>')+'</p><p class="small">Contrôlez les données avant de confirmer. Les informations ambiguës restent à vérifier.</p><button class="primary" id="smartConfirm" '+(!unique.length?'disabled':'')+'>Confirmer l’importation</button></div>';document.getElementById('smartConfirm').onclick=commit}
function commit(){let rows=window.smartReady||[];if(!rows.length)return;let before=JSON.stringify(db),now=stamp();for(let r of rows){if(!db.sites.includes(r.site))db.sites.push(r.site);let category=r.category||'Locaux techniques';let cat=db.categories.find(c=>c.name===category);if(!cat){cat={id:id(),name:category,subs:[]};db.categories.push(cat)}let sub=r.sub||'';if(sub&&!cat.subs.includes(sub))cat.subs.push(sub);db.records.push({id:id(),site:r.site,name:r.name,category,sub,local:r.local||'',reference:r.reference||'',phone:r.phone||'',supply:r.supply||'',access:r.access||'',notes:[r.line?'Ligne '+r.line:'',r.mode||'',r.notes||''].filter(Boolean).join(' | '),status:'À vérifier',links:'',created:now,updated:now,author:actor(),image:'',annotated:''})}log('Import intelligent',rows.length+' fiches');if(!persist()){db=JSON.parse(before);return}pending=[];window.smartReady=[];alert(rows.length+' fiches importées. Pensez à exporter une sauvegarde JSON.');render()}
function library(){return new Promise((resolve,reject)=>{if(window.XLSX)return resolve();let s=document.createElement('script');s.src='https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';s.onload=resolve;s.onerror=()=>reject(Error('Impossible de charger le lecteur Excel. Vérifiez Internet.'));document.head.appendChild(s)})}
async function read(f){let ext=f.name.split('.').pop().toLowerCase();if(['xlsx','xls','ods'].includes(ext)){await library();let book=XLSX.read(await f.arrayBuffer(),{type:'array'});return book.SheetNames.flatMap(n=>parseRows(XLSX.utils.sheet_to_json(book.Sheets[n],{header:1,defval:''})))}if(ext==='csv'||ext==='tsv'){let content=(await f.text()).replace(/^\ufeff/,'');return parseRows(parseCSV(content))}throw Error('Ce format sera ajouté ultérieurement. Pour le moment : Excel, ODS et CSV.')}
window.smartImportFile=async function(input){let f=input.files?.[0];if(!f)return;let out=document.getElementById('smartPreview');out.textContent='Lecture du fichier…';try{pending=await read(f);if(!pending.length)throw Error('Aucune fiche reconnue. Vérifiez les intitulés des colonnes.');preview()}catch(e){out.textContent='Import impossible : '+e.message}finally{input.value=''}};
const originalSettings=settings;settings=function(){return originalSettings()+'<div class="panel"><h2>Importation intelligente</h2><p class="small">Choisissez un fichier Excel (.xls, .xlsx), ODS ou CSV sur votre iPhone. Prévisualisation et détection des doublons avant ajout. Traitement local : aucune donnée envoyée à un serveur par cette fonction. PDF et photos à venir.</p><input type="file" accept=".xls,.xlsx,.ods,.csv" onchange="smartImportFile(this)"><div id="smartPreview"></div></div>'};
})();