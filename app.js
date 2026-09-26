const SUMATRA=[[105.817655,-5.852356],[104.710384,-5.873285],[103.868213,-5.037315],[102.584261,-4.220259],[102.156173,-3.614146],[101.399113,-2.799777],[100.902503,-2.050262],[100.141981,-0.650348],[99.26374,0.183142],[98.970011,1.042882],[98.601351,1.823507],[97.699598,2.453184],[97.176942,3.308791],[96.424017,3.86886],[95.380876,4.970782],[95.293026,5.479821],[95.936863,5.439513],[97.484882,5.246321],[98.369169,4.26837],[99.142559,3.59035],[99.693998,3.174329],[100.641434,2.099381],[101.658012,2.083697],[102.498271,1.3987],[103.07684,0.561361],[103.838396,0.104542],[103.437645,-0.711946],[104.010789,-1.059212],[104.369991,-1.084843],[104.53949,-1.782372],[104.887893,-2.340425],[105.622111,-2.428844],[106.108593,-3.061777],[105.857446,-4.305525],[105.817655,-5.852356]];
const BORNEO=[[117.875627,1.827641],[118.996747,0.902219],[117.811858,0.784242],[117.478339,0.102475],[117.521644,-0.803723],[116.560048,-1.487661],[116.533797,-2.483517],[116.148084,-4.012726],[116.000858,-3.657037],[114.864803,-4.106984],[114.468652,-3.495704],[113.755672,-3.43917],[113.256994,-3.118776],[112.068126,-3.478392],[111.703291,-2.994442],[111.04824,-3.049426],[110.223846,-2.934032],[110.070936,-1.592874],[109.571948,-1.314907],[109.091874,-0.459507],[108.952658,0.415375],[109.069136,1.341934],[109.66326,2.006467],[110.396135,1.663775],[111.168853,1.850637],[111.370081,2.697303],[111.796928,2.885897],[112.995615,3.102395],[113.712935,3.893509],[114.204017,4.525874],[114.599961,4.900011],[115.45071,5.44773],[116.220741,6.143191],[116.725103,6.924771],[117.129626,6.928053],[117.643393,6.422166],[117.689075,5.98749],[118.347691,5.708696],[119.181904,5.407836],[119.110694,5.016128],[118.439727,4.966519],[118.618321,4.478202],[117.882035,4.137551],[117.313232,3.234428],[118.04833,2.28769],[117.875627,1.827641]];

const T={
th:{subtitle:"จำแนกจุดความร้อนระหว่าง Sumatra, Borneo และเกาะอื่นของอินโดนีเซีย",pickTitle:"เลือกไฟล์ข้อมูล",pickHint:"รองรับ .xlsx/.xls ที่มี LATITUDE, LONGITUDE และ CT_EN",pickButton:"📂 เลือกไฟล์ Excel",none:"ยังไม่ได้เลือกไฟล์",sheet:"Worksheet",sum:"SUM HOTSPOT",totalNote:"จำนวนจุดที่มีพิกัดใช้งานได้",sb:"Sumatra + Borneo",id:"Indonesia, all islands",otherId:"Other Indonesian Islands",unit:"จุด",borneoTitle:"Borneo แยกตามประเทศ",borneoHint:"เป็นรายละเอียดภายในยอด Borneo และไม่ถูกบวกซ้ำใน SUM",otherTitle:"Other Countries",otherHint:"ประเทศอื่นที่อยู่นอก Sumatra/Borneo",otherSum:"รวม",copy:"คัดลอกสรุป",preview:"ตัวอย่างข้อมูลที่จำแนกแล้ว",privacy:"ไฟล์ประมวลผลภายใน browser เท่านั้น ไม่มีการอัปโหลดข้อมูลขึ้น server",ready:"อ่านไฟล์สำเร็จ",missing:"ไม่พบคอลัมน์ LATITUDE/LONGITUDE",copied:"คัดลอกสรุปแล้ว",unknown:"ไม่ระบุประเทศ"},
en:{subtitle:"Classify hotspots across Sumatra, Borneo, and other Indonesian islands",pickTitle:"Select data file",pickHint:"Supports .xlsx/.xls with LATITUDE, LONGITUDE, and CT_EN columns",pickButton:"📂 Choose Excel File",none:"No file selected",sheet:"Worksheet",sum:"SUM HOTSPOT",totalNote:"Hotspots with usable coordinates",sb:"Sumatra + Borneo",id:"Indonesia, all islands",otherId:"Other Indonesian Islands",unit:"hotspots",borneoTitle:"Borneo by Country",borneoHint:"Breakdown within the Borneo total; not added again to SUM",otherTitle:"Other Countries",otherHint:"Countries outside Sumatra/Borneo",otherSum:"Total",copy:"Copy Summary",preview:"Classified Data Preview",privacy:"Files are processed locally in your browser and are not uploaded to a server",ready:"File processed successfully",missing:"LATITUDE/LONGITUDE columns not found",copied:"Summary copied",unknown:"Unknown"}};

const COUNTRY_TH={Indonesia:"อินโดนีเซีย",Malaysia:"มาเลเซีย",Brunei:"บรูไน",China:"จีน",Vietnam:"เวียดนาม",Thailand:"ไทย",Laos:"ลาว",Myanmar:"เมียนมา",Cambodia:"กัมพูชา",Philippines:"ฟิลิปปินส์",Singapore:"สิงคโปร์",Timor_Leste:"ติมอร์-เลสเต"};
let lang="th", workbook=null, currentFile="", analyzed=null;

const $=id=>document.getElementById(id);
function nfmt(n){return Number(n||0).toLocaleString("en-US")}
function pointInPoly(x,y,p){let inside=false;for(let i=0,j=p.length-1;i<p.length;j=i++){const xi=p[i][0],yi=p[i][1],xj=p[j][0],yj=p[j][1];const hit=((yi>y)!=(yj>y))&&(x<(xj-xi)*(y-yi)/(yj-yi)+xi);if(hit)inside=!inside}return inside}
function keyCountry(v){let s=String(v??"").trim();if(!s)return "Unknown";const u=s.toUpperCase().replace(/[._-]+/g," ").replace(/\s+/g," ");const map={"INDONESIA":"Indonesia","MALAYSIA":"Malaysia","BRUNEI":"Brunei","BRUNEI DARUSSALAM":"Brunei","CHINA":"China","VIETNAM":"Vietnam","VIET NAM":"Vietnam","THAILAND":"Thailand","LAOS":"Laos","LAO PDR":"Laos","LAO PEOPLE'S DEMOCRATIC REPUBLIC":"Laos","MYANMAR":"Myanmar","BURMA":"Myanmar","CAMBODIA":"Cambodia","PHILIPPINES":"Philippines","SINGAPORE":"Singapore","TIMOR LESTE":"Timor_Leste","TIMOR-LESTE":"Timor_Leste"};return map[u]||s.replace(/\b\w/g,c=>c.toUpperCase())}
function countryLabel(k){if(k==="Unknown")return T[lang].unknown;if(lang==="th"&&COUNTRY_TH[k])return COUNTRY_TH[k];return k.replaceAll("_"," ")}
function colIndex(headers,names){const norm=headers.map(v=>String(v??"").trim().toUpperCase().replace(/\s+/g,"_"));for(const n of names){const i=norm.indexOf(n);if(i>=0)return i}return-1}
function analyzeSheet(){
  const name=$("sheetSelect").value, ws=workbook.Sheets[name];
  const aoa=XLSX.utils.sheet_to_json(ws,{header:1,defval:null,raw:true});
  if(!aoa.length)return;
  const h=aoa[0], ilat=colIndex(h,["LATITUDE","LAT"]), ilon=colIndex(h,["LONGITUDE","LONG","LON"]), ict=colIndex(h,["CT_EN","COUNTRY","COUNTRY_EN"]);
  if(ilat<0||ilon<0){$("status").className="status err";$("status").textContent=T[lang].missing;return}
  let total=0,sumatra=0,borneo=0,otherId=0,idTotal=0; const bc={},oc={},rows=[];
  for(let r=1;r<aoa.length;r++){
    const a=aoa[r];
    const rawLat=a[ilat], rawLon=a[ilon];
    if(rawLat===null||rawLat===undefined||rawLon===null||rawLon===undefined||String(rawLat).trim()===""||String(rawLon).trim()==="")continue;
    const lat=Number(rawLat), lon=Number(rawLon);
    if(!Number.isFinite(lat)||!Number.isFinite(lon)||lat<-90||lat>90||lon<-180||lon>180)continue;
    total++;
    const country=keyCountry(ict>=0?a[ict]:"");
    if(country==="Indonesia")idTotal++;
    let cls="",bCountry="";
    if(pointInPoly(lon,lat,SUMATRA)){cls="Sumatra";sumatra++}
    else if(pointInPoly(lon,lat,BORNEO)){cls="Borneo";borneo++;bCountry=country;bc[country]=(bc[country]||0)+1}
    else if(country==="Indonesia"){cls="Other Indonesian Islands";otherId++}
    else{cls="Other Countries";oc[country]=(oc[country]||0)+1}
    const obj={};h.forEach((x,i)=>{const k=String(x??"").trim()||("Column_"+(i+1));obj[k]=a[i]??""});
    obj.Hotspot_Class=cls;obj.Borneo_Country=bCountry==="Unknown"?"":bCountry;rows.push(obj);
  }
  analyzed={sheet:name,total,sumatra,borneo,otherId,idTotal,bc,oc,rows,otherTotal:Object.values(oc).reduce((a,b)=>a+b,0)};
  render();
  $("status").className="status ok";$("status").textContent=T[lang].ready+" • "+name;
}
function renderCountries(el,data){
  el.innerHTML="";
  const e=Object.entries(data).sort((a,b)=>b[1]-a[1]);
  if(!e.length){el.innerHTML='<div class="muted">-</div>';return}
  for(const [k,v] of e){const d=document.createElement("div");d.className=el.id==="borneoCountries"?"country-card":"country-row";d.innerHTML=el.id==="borneoCountries"?'<span>'+countryLabel(k)+'</span><b>'+nfmt(v)+'</b>':'<span>'+countryLabel(k)+'</span><b>'+nfmt(v)+'</b>';el.appendChild(d)}
}
function render(){
  if(!analyzed)return;
  $("results").classList.remove("hidden");
  $("totalCount").textContent=nfmt(analyzed.total);$("sumatra").textContent=nfmt(analyzed.sumatra);$("borneo").textContent=nfmt(analyzed.borneo);$("otherId").textContent=nfmt(analyzed.otherId);$("sbCount").textContent=nfmt(analyzed.sumatra+analyzed.borneo);$("idCount").textContent=nfmt(analyzed.idTotal);$("otherSum").textContent=nfmt(analyzed.otherTotal);
  renderCountries($("borneoCountries"),analyzed.bc);renderCountries($("otherCountries"),analyzed.oc);
  const cols=["LATITUDE","LONGITUDE","CT_EN","Hotspot_Class","Borneo_Country"], head=$("previewHead"), body=$("previewBody");head.innerHTML="";body.innerHTML="";
  cols.forEach(c=>{const th=document.createElement("th");th.textContent=c;head.appendChild(th)});
  analyzed.rows.slice(0,12).forEach(o=>{const tr=document.createElement("tr");cols.forEach(c=>{const td=document.createElement("td");td.textContent=o[c]??"";tr.appendChild(td)});body.appendChild(tr)});
}
function applyLang(l){
  lang=l;document.documentElement.lang=l;$("thBtn").classList.toggle("active",l==="th");$("enBtn").classList.toggle("active",l==="en");
  const t=T[l];$("subtitle").textContent=t.subtitle;$("pickTitle").textContent=t.pickTitle;$("pickHint").textContent=t.pickHint;$("pickButton").textContent=t.pickButton;if(!currentFile)$("fileName").textContent=t.none;$("sheetLabel").textContent=t.sheet;$("sumLabel").textContent=t.sum;$("totalNote").textContent=t.totalNote;$("sbLabel").textContent=t.sb;$("idLabel").textContent=t.id;$("otherIdTitle").textContent=t.otherId;["unit1","unit2","unit3"].forEach(x=>$(x).textContent=t.unit);$("borneoTitle").textContent=t.borneoTitle;$("borneoHint").textContent=t.borneoHint;$("otherTitle").textContent=t.otherTitle;$("otherHint").textContent=t.otherHint;$("otherSumLabel").textContent=t.otherSum;$("copyBtn").textContent=t.copy;$("previewTitle").textContent=t.preview;$("privacy").textContent=t.privacy;if(analyzed)render();
}
$("thBtn").onclick=()=>applyLang("th");$("enBtn").onclick=()=>applyLang("en");
$("fileInput").addEventListener("change",async e=>{
  const f=e.target.files[0];if(!f)return;currentFile=f.name;$("fileName").textContent=f.name;$("status").className="status";$("status").textContent="...";
  try{const ab=await f.arrayBuffer();workbook=XLSX.read(ab,{type:"array"});const sel=$("sheetSelect");sel.innerHTML="";workbook.SheetNames.forEach(n=>{const o=document.createElement("option");o.value=o.textContent=n;sel.appendChild(o)});sel.value=workbook.SheetNames.includes("ALL")?"ALL":workbook.SheetNames[0];$("sheetBox").classList.remove("hidden");analyzeSheet()}catch(err){$("status").className="status err";$("status").textContent=String(err.message||err)}
});
$("sheetSelect").addEventListener("change",analyzeSheet);
function summaryText(){if(!analyzed)return"";let s="Hotspot Island Counter\nFile: "+currentFile+"\nWorksheet: "+analyzed.sheet+"\nTotal hotspots: "+analyzed.total+"\nSumatra: "+analyzed.sumatra+"\nBorneo: "+analyzed.borneo+"\nOther Indonesian Islands: "+analyzed.otherId+"\nOther Countries: "+analyzed.otherTotal+"\nSumatra + Borneo: "+(analyzed.sumatra+analyzed.borneo)+"\nIndonesia (all islands): "+analyzed.idTotal+"\n\nBorneo by Country:\n";Object.entries(analyzed.bc).sort((a,b)=>b[1]-a[1]).forEach(([k,v])=>s+="  "+k+": "+v+"\n");s+="\nOther Countries:\n";Object.entries(analyzed.oc).sort((a,b)=>b[1]-a[1]).forEach(([k,v])=>s+="  "+k+": "+v+"\n");return s}
$("copyBtn").onclick=async()=>{if(!analyzed)return;await navigator.clipboard.writeText(summaryText());$("status").className="status ok";$("status").textContent=T[lang].copied};
function baseName(){return(currentFile||"hotspots").replace(/\.[^.]+$/,"")}
$("csvBtn").onclick=()=>{if(!analyzed)return;const ws=XLSX.utils.json_to_sheet(analyzed.rows), csv=XLSX.utils.sheet_to_csv(ws), blob=new Blob(["\uFEFF"+csv],{type:"text/csv;charset=utf-8"}),a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=baseName()+"_island_classified.csv";a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000)};
$("xlsxBtn").onclick=()=>{if(!analyzed)return;const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(analyzed.rows),"Classified");const summary=[["Metric","Count"],["Total Hotspots",analyzed.total],["Sumatra",analyzed.sumatra],["Borneo",analyzed.borneo],["Other Indonesian Islands",analyzed.otherId],["Other Countries",analyzed.otherTotal],["Sumatra + Borneo",analyzed.sumatra+analyzed.borneo],["Indonesia (all islands)",analyzed.idTotal],[],["Borneo by Country",""],...Object.entries(analyzed.bc).sort((a,b)=>b[1]-a[1]),[],["Other Countries",""],...Object.entries(analyzed.oc).sort((a,b)=>b[1]-a[1])];XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(summary),"Summary");XLSX.writeFile(wb,baseName()+"_island_classified.xlsx")};
applyLang("th");