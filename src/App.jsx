import { useState, useRef, useCallback } from "react";

const G = {
  bg:"#0e1117",card:"#161b27",border:"#1e2a3d",accent:"#3b82f6",
  green:"#22c55e",yellow:"#f59e0b",red:"#ef4444",muted:"#4a5a7a",
  text:"#e2e8f0",sub:"#8899aa",faint:"#1a2235"
};

const PASSWORD = "sqli2026";

// ─── API — via Vercel proxy (clé côté serveur) ────────────────────────────────
async function callClaude(system, userText, images=[]) {
  const content=[];
  for(const img of images) content.push({type:"image",source:{type:"base64",media_type:img.type,data:img.data}});
  content.push({type:"text",text:userText});
  const res=await fetch("/api/claude",{
    method:"POST",
    headers:{"Content-Type":"application/json"},
    body:JSON.stringify({
      model:"claude-sonnet-4-20250514",
      max_tokens:8000,
      system,
      messages:[{role:"user",content}]
    }),
  });
  if(!res.ok){const e=await res.json();throw new Error(e.error?.message||`API ${res.status}`);}
  return (await res.json()).content[0].text;
}

function parseJSON(raw){
  try{return JSON.parse(raw.trim());}catch{}
  const stripped=raw.replace(/```json|```/g,"").trim();
  try{return JSON.parse(stripped);}catch{}
  const objMatch=stripped.match(/\{[\s\S]*\}/);
  if(objMatch)try{return JSON.parse(objMatch[0]);}catch{}
  throw new Error("Réponse JSON invalide — réessaie.");
}

// ─── FILE UTILS ───────────────────────────────────────────────────────────────
function readFileAsText(file,enc="UTF-8"){
  return new Promise((res,rej)=>{const r=new FileReader();r.onload=e=>res(e.target.result);r.onerror=rej;r.readAsText(file,enc);});
}
function readFileAsBase64(file){
  return new Promise((res,rej)=>{const r=new FileReader();r.onload=e=>res(e.target.result.split(",")[1]);r.onerror=rej;r.readAsDataURL(file);});
}
function extractKeywordsFromText(text){
  text=text.replace(/^\uFEFF/,"");
  const lines=text.split(/\r?\n/).filter(Boolean);
  let headerIdx=-1,kwCol=-1,volCol=-1;
  for(let i=0;i<Math.min(lines.length,5);i++){
    const cols=lines[i].split("\t");
    const ki=cols.findIndex(c=>c.trim().toLowerCase()==="keyword");
    if(ki>=0){headerIdx=i;kwCol=ki;volCol=cols.findIndex(c=>c.toLowerCase().includes("avg")&&c.toLowerCase().includes("monthly"));break;}
  }
  if(headerIdx>=0){
    return lines.slice(headerIdx+1).map(l=>{const c=l.split("\t");return{keyword:c[kwCol]?.trim(),volume:parseInt(c[volCol]?.replace(/\D/g,"")||"0")};}).filter(r=>r.keyword&&r.keyword.length>1);
  }
  return lines.map(l=>({keyword:l.split(",")[0].replace(/^"|"$/g,"").trim(),volume:0})).filter(r=>r.keyword&&r.keyword.length>1&&r.keyword.length<150);
}
async function parseXlsxNative(file){
  try{
    const buf=await file.arrayBuffer();
    if(!window.JSZip){
      await new Promise((res,rej)=>{const s=document.createElement("script");s.src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js";s.onload=res;s.onerror=rej;document.head.appendChild(s);});
    }
    const zip=await window.JSZip.loadAsync(buf);
    const ssFile=zip.file("xl/sharedStrings.xml");
    const ss=[];
    if(ssFile){const x=await ssFile.async("string");for(const m of x.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g))ss.push(m[1].replace(/&amp;/g,"&").replace(/&lt;/g,"<").replace(/&gt;/g,">").replace(/&quot;/g,'"').replace(/&#(\d+);/g,(_,n)=>String.fromCharCode(n)));}
    const sf=zip.file("xl/worksheets/sheet1.xml")||zip.file("xl/worksheets/Sheet1.xml");
    if(!sf)return[];
    const xml=await sf.async("string");
    const grid=[];
    for(const rm of xml.matchAll(/<row[^>]*>([\s\S]*?)<\/row>/g)){
      const row={};
      for(const cell of rm[1].matchAll(/<c r="([A-Z]+)\d+"([^>]*)>([\s\S]*?)<\/c>/g)){
        const col=cell[1],attrs=cell[2],inner=cell[3];
        const vm=inner.match(/<v>([\s\S]*?)<\/v>/);
        if(vm)row[col]=attrs.includes('t="s"')?(ss[parseInt(vm[1])]||""):vm[1];
      }
      grid.push(row);
    }
    if(!grid.length)return[];
    let hrow=grid[0],hidx=0;
    for(let i=0;i<Math.min(grid.length,5);i++){
      if(Object.values(grid[i]).some(v=>String(v).toLowerCase()==="keyword")){hrow=grid[i];hidx=i;break;}
    }
    let kwCol=null,volCol=null;
    for(const [col,val] of Object.entries(hrow)){
      const v=String(val).toLowerCase().trim();
      if(["keyword","keywords","mot-clé","mot clé","query"].includes(v))kwCol=col;
      if(v==="volume"||v==="avg. monthly searches"||v==="search volume"||(v.includes("volume")&&!v.includes("url")&&!v.includes("page")&&!v.includes("density")))volCol=col;
    }
    if(!kwCol)kwCol=Object.keys(hrow)[0];
    const rows=[];
    for(let i=hidx+1;i<grid.length;i++){
      const kw=String(grid[i][kwCol]||"").trim();
      const vol=volCol?parseInt(String(grid[i][volCol]||"0").replace(/[^0-9]/g,""))||0:0;
      if(kw&&kw.length>1&&kw.length<200&&!kw.toLowerCase().includes("http"))rows.push({keyword:kw,volume:vol});
    }
    return rows;
  }catch(e){console.error("XLSX:",e);return[];}
}
async function parseFile(file){
  const isXlsx=file.name.toLowerCase().endsWith(".xlsx")||file.name.toLowerCase().endsWith(".xls");
  if(isXlsx)return parseXlsxNative(file);
  let text="";
  try{text=await readFileAsText(file,"UTF-16LE");if(!text.includes("eyword")&&!text.includes(","))text=await readFileAsText(file,"UTF-8");}
  catch{text=await readFileAsText(file,"UTF-8");}
  return extractKeywordsFromText(text.replace(/^\uFEFF/,""));
}
function dedup(rows){
  const map=new Map();
  for(const r of rows){const k=r.keyword.toLowerCase().trim();if(!map.has(k)||map.get(k).volume<r.volume)map.set(k,r);}
  return Array.from(map.values()).sort((a,b)=>b.volume-a.volume);
}
const BRAND_BLACKLIST=["tefal","rowenta","krups","moulinex","calor","seb ","imusa","wmf","all-clad","lagostina","t-fal","supor","cuisinart","kitchenaid","breville","ninja","instant pot","hamilton beach","black+decker","black decker","blackdecker","philips","braun","bosch","siemens","kenwood","delonghi","de'longhi","nespresso","keurig","amazon","walmart","target","lidl","aldi","costco","ikea","presto","lodge","le creuset","staub","mauviel","demeyere","fissler","tramontina"];
const NOISE_RX=[/\brecette[s]?\b/i,/\brecipe[s]?\b/i,/\breceta[s]?\b/i,/\b20[0-9]{2}\b/,/black friday/i,/cyber monday/i,/\bsoldes?\b/i,/wikipedia/i,/youtube/i,/instagram/i,/facebook/i,/\bebay\b/i,/\bcdiscount\b/i];
function ruleBasedClean(kws){return kws.filter(({keyword})=>{const kw=keyword.toLowerCase();if(BRAND_BLACKLIST.some(b=>kw.includes(b)))return false;if(NOISE_RX.some(p=>p.test(kw)))return false;return true;});}
function buildBalancedList(kws,target){
  const groups={};
  for(const k of kws){const cat=k.category||"Autre";if(!groups[cat])groups[cat]=[];groups[cat].push(k);}
  const cats=Object.keys(groups);if(!cats.length)return kws.slice(0,target);
  const total=kws.length;const selected=[];
  cats.forEach(cat=>{const quota=Math.max(1,Math.round(target*(groups[cat].length/total)));selected.push(...[...groups[cat]].sort((a,b)=>b.volume-a.volume).slice(0,quota));});
  return selected.sort((a,b)=>b.volume-a.volume).slice(0,target);
}
function dl(content,filename,type="text/plain;charset=utf-8;"){
  const a=document.createElement("a");a.href=URL.createObjectURL(new Blob([content],{type}));a.download=filename;
  document.body.appendChild(a);a.click();setTimeout(()=>{document.body.removeChild(a);URL.revokeObjectURL(a.href);},200);
}
function dlCSV(rows,filename){dl("\uFEFF"+rows.map(r=>r.map(c=>`"${String(c).replace(/"/g,'""')}"`).join(",")).join("\n"),filename,"text/csv;charset=utf-8;");}

const SEB_CATS=["LINEN CARE","HOME COMFORT","PERSONAL CARE","ELECTRICAL COOKING","FOOD PREPARATION","BEVERAGE","FLOOR CARE","COOKWARE & BAKEWARE","KITCHENWARE & DINNER"];
const STEPS=["Projet","Catégories","Mots-clés","Volumes","Nettoyage","Catégorisation","Export"];

// ─── UI COMPONENTS ────────────────────────────────────────────────────────────
function Card({children,style={}}){return <div style={{background:G.card,border:`1px solid ${G.border}`,borderRadius:12,padding:28,...style}}>{children}</div>;}
function Btn({onClick,disabled,children,secondary,color,style={}}){
  const bg=color||(secondary?"transparent":G.accent);
  return <button onClick={onClick} disabled={disabled} style={{background:bg,color:secondary&&!color?G.sub:"#fff",border:secondary?`1px solid ${G.border}`:"none",padding:"10px 22px",borderRadius:7,fontFamily:"inherit",fontSize:13,fontWeight:600,cursor:disabled?"not-allowed":"pointer",opacity:disabled?.4:1,...style}}>{children}</button>;
}
function SBtn({onClick,children,color=G.accent}){return <button onClick={onClick} style={{background:color+"18",border:`1px solid ${color}44`,color,padding:"5px 12px",borderRadius:5,fontFamily:"inherit",fontSize:11,fontWeight:600,cursor:"pointer"}}>{children}</button>;}
function Inp({label,placeholder,value,onChange,type="text"}){return <div><label style={{fontSize:11,color:G.muted,display:"block",marginBottom:5,textTransform:"uppercase",letterSpacing:".07em"}}>{label}</label><input type={type} value={value} onChange={e=>onChange(e.target.value)} placeholder={placeholder} style={{background:G.faint,border:`1px solid ${G.border}`,color:G.text,padding:"10px 14px",borderRadius:7,fontFamily:"inherit",fontSize:13,width:"100%",outline:"none",boxSizing:"border-box"}}/></div>;}
function Stat({value,label,color=G.accent}){return <div style={{background:G.faint,border:`1px solid ${G.border}`,borderRadius:8,padding:"14px 18px",textAlign:"center"}}><div style={{fontSize:26,fontWeight:700,color,fontFamily:"'Space Grotesk',sans-serif"}}>{typeof value==="number"?value.toLocaleString():value}</div><div style={{fontSize:11,color:G.muted,marginTop:3}}>{label}</div></div>;}
function Spin({msg}){return <div style={{textAlign:"center",padding:"24px 0"}}><div style={{display:"inline-block",width:30,height:30,border:`3px solid ${G.border}`,borderTopColor:G.accent,borderRadius:"50%",animation:"spin 1s linear infinite",marginBottom:10}}/><div style={{fontSize:13,color:G.accent}}>{msg}</div></div>;}
function ErrBox({msg}){if(!msg)return null;return <div style={{background:G.red+"18",border:`1px solid ${G.red}44`,borderRadius:8,padding:"12px 16px",color:"#f87171",fontSize:13,marginBottom:16}}>⚠ {msg}</div>;}
function Info({children,color=G.accent}){return <div style={{background:color+"12",border:`1px solid ${color}33`,borderRadius:8,padding:"12px 16px",fontSize:12,color:G.sub,lineHeight:1.7,marginBottom:16}}>{children}</div>;}
function Drop({label,note,accept,multiple,onFiles,children}){
  const ref=useRef();
  return <div onClick={()=>ref.current.click()} onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();onFiles(Array.from(e.dataTransfer.files));}} style={{background:G.faint,border:`2px dashed ${G.border}`,borderRadius:10,padding:"22px 20px",textAlign:"center",cursor:"pointer"}}>
    <div style={{fontSize:22,marginBottom:6}}>📂</div>
    <div style={{fontSize:13,color:G.sub,marginBottom:3}}>{label}</div>
    <div style={{fontSize:11,color:G.muted}}>{note}</div>
    {children}
    <input ref={ref} type="file" accept={accept} multiple={multiple} style={{display:"none"}} onChange={e=>onFiles(Array.from(e.target.files))}/>
  </div>;
}
function StepBar({step}){return <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>{STEPS.map((l,i)=><div key={i} style={{padding:"4px 11px",borderRadius:4,border:`1px solid ${i===step?G.accent:i<step?G.green:G.border}`,color:i===step?G.accent:i<step?G.green:G.muted,fontSize:11}}>{i<step?"✓":i+1}. {l}</div>)}</div>;}
function Tag({n,children}){return <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}><span style={{background:G.accent,color:"#fff",fontSize:11,fontWeight:700,padding:"2px 9px",borderRadius:20}}>{n}</span><span style={{fontSize:13,color:G.sub}}>{children}</span></div>;}
function AddCompForm({onAdd}){
  const [open,setOpen]=useState(false);const [name,setName]=useState("");const [url,setUrl]=useState("");
  if(!open)return <button onClick={()=>setOpen(true)} style={{marginTop:10,background:"none",border:`1px dashed ${G.border}`,color:G.muted,padding:"7px 16px",borderRadius:7,fontFamily:"inherit",fontSize:12,cursor:"pointer",width:"100%",textAlign:"center"}}>+ Ajouter un concurrent manuellement</button>;
  return <div style={{marginTop:10,background:G.faint,border:`1px solid ${G.border}`,borderRadius:8,padding:"12px 14px"}}>
    <div style={{fontSize:12,color:G.text,fontWeight:600,marginBottom:10}}>Ajouter un concurrent</div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:10}}>
      <div><div style={{fontSize:10,color:G.muted,marginBottom:3,textTransform:"uppercase",letterSpacing:".06em"}}>Nom</div><input value={name} onChange={e=>setName(e.target.value)} placeholder="ex: Cuisinart" style={{background:G.card,border:`1px solid ${G.border}`,color:G.text,padding:"7px 10px",borderRadius:5,fontFamily:"inherit",fontSize:12,width:"100%",outline:"none"}}/></div>
      <div><div style={{fontSize:10,color:G.muted,marginBottom:3,textTransform:"uppercase",letterSpacing:".06em"}}>Domaine</div><input value={url} onChange={e=>setUrl(e.target.value)} placeholder="ex: cuisinart.com" style={{background:G.card,border:`1px solid ${G.border}`,color:G.accent,padding:"7px 10px",borderRadius:5,fontFamily:"inherit",fontSize:12,width:"100%",outline:"none"}}/></div>
    </div>
    <div style={{display:"flex",gap:8}}>
      <button onClick={()=>{if(name&&url){onAdd({name,url:url.replace(/^https?:\/\//,""),reason:"Ajouté manuellement"});setName("");setUrl("");setOpen(false);}}} disabled={!name||!url} style={{background:G.accent,color:"#fff",border:"none",padding:"7px 16px",borderRadius:5,fontFamily:"inherit",fontSize:12,cursor:"pointer",opacity:!name||!url?.4:1}}>Ajouter</button>
      <button onClick={()=>setOpen(false)} style={{background:"none",border:`1px solid ${G.border}`,color:G.muted,padding:"7px 12px",borderRadius:5,fontFamily:"inherit",fontSize:12,cursor:"pointer"}}>Annuler</button>
    </div>
  </div>;
}

// ─── LOGIN SCREEN ─────────────────────────────────────────────────────────────
function LoginScreen({onSuccess}){
  const [pwd,setPwd]=useState("");
  const [shake,setShake]=useState(false);
  const submit=()=>{
    if(pwd===PASSWORD){onSuccess();}
    else{setShake(true);setTimeout(()=>setShake(false),500);}
  };
  return(
    <div style={{background:G.bg,minHeight:"100vh",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",fontFamily:"'IBM Plex Mono','Courier New',monospace",padding:20}}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600&family=Space+Grotesk:wght@400;500;700&display=swap');
        *{box-sizing:border-box}
        @keyframes spin{to{transform:rotate(360deg)}}
        @keyframes shake{0%,100%{transform:translateX(0)}20%,60%{transform:translateX(-8px)}40%,80%{transform:translateX(8px)}}
        @keyframes fadein{from{opacity:0;transform:translateY(16px)}to{opacity:1;transform:translateY(0)}}
      `}</style>

      {/* Logo */}
      <div style={{marginBottom:40,textAlign:"center",animation:"fadein .5s ease"}}>
        <div style={{width:56,height:56,background:"linear-gradient(135deg,#3b82f6,#22c55e)",borderRadius:16,display:"flex",alignItems:"center",justifyContent:"center",fontSize:26,margin:"0 auto 16px"}}>⚡</div>
        <div style={{fontFamily:"'Space Grotesk',sans-serif",fontWeight:700,fontSize:22,color:"#fff",letterSpacing:"-.02em"}}>SEO Keyword Pipeline</div>
        <div style={{fontSize:12,color:G.muted,marginTop:4}}>Groupe SEB — Accès restreint</div>
      </div>

      {/* Card */}
      <div style={{
        background:G.card,border:`1px solid ${G.border}`,borderRadius:16,
        padding:"32px 36px",width:"100%",maxWidth:380,
        animation:`fadein .5s ease .1s both,${shake?"shake .4s ease":"none"}`
      }}>
        <div style={{fontSize:13,color:G.sub,marginBottom:20,lineHeight:1.6}}>Entre le mot de passe pour accéder à l'outil.</div>

        <div style={{marginBottom:16}}>
          <label style={{fontSize:11,color:G.muted,display:"block",marginBottom:6,textTransform:"uppercase",letterSpacing:".08em"}}>Mot de passe</label>
          <input
            type="password"
            value={pwd}
            onChange={e=>setPwd(e.target.value)}
            onKeyDown={e=>e.key==="Enter"&&submit()}
            placeholder="••••••••"
            autoFocus
            style={{
              width:"100%",background:G.faint,border:`1px solid ${G.border}`,
              color:G.text,padding:"12px 14px",borderRadius:8,
              fontFamily:"inherit",fontSize:14,outline:"none",
              transition:"border-color .2s",
            }}
            onFocus={e=>e.target.style.borderColor=G.accent}
            onBlur={e=>e.target.style.borderColor=G.border}
          />
        </div>

        <button
          onClick={submit}
          style={{
            width:"100%",background:`linear-gradient(135deg,${G.accent},#2563eb)`,
            color:"#fff",border:"none",padding:"12px",borderRadius:8,
            fontFamily:"inherit",fontSize:14,fontWeight:600,cursor:"pointer",
            transition:"opacity .2s",letterSpacing:".02em"
          }}
          onMouseEnter={e=>e.target.style.opacity=".9"}
          onMouseLeave={e=>e.target.style.opacity="1"}
        >
          Accéder →
        </button>
      </div>

      <div style={{marginTop:24,fontSize:11,color:G.muted}}>Usage interne SQLI · Groupe SEB</div>
    </div>
  );
}

// ─── MAIN APP ─────────────────────────────────────────────────────────────────
export default function App(){
  const [authed,setAuthed]=useState(false);

  // ALL hooks at top — even before auth check
  const [step,setStep]=useState(0);
  const [err,setErr]=useState("");
  const [loading,setLoading]=useState(false);
  const [loadMsg,setLoadMsg]=useState("");
  const [log,setLog]=useState([]);
  const [config,setConfig]=useState({brand:"",country:"",language:"",siteUrl:""});
  const [catImages,setCatImages]=useState([]);
  const [categories,setCategories]=useState([]);
  const [catManual,setCatManual]=useState("");
  const [enrichedKws,setEnrichedKws]=useState([]);
  const [enrichPreview,setEnrichPreview]=useState("");
  const [enrichValidated,setEnrichValidated]=useState(false);
  const [kwSub,setKwSub]=useState(0);
  const [competitors,setCompetitors]=useState([]);
  const [sourceKws,setSourceKws]=useState([]);
  const [rawWithVol,setRawWithVol]=useState([]);
  const [cleanedKws,setCleanedKws]=useState([]);
  const [cleanStats,setCleanStats]=useState(null);
  const [categorizedKws,setCategorizedKws]=useState([]);
  const [targetCount,setTargetCount]=useState(500);

  const addLog=useCallback(m=>setLog(l=>[...l,m]),[]);
  const nav=s=>{setStep(s);setErr("");};
  const allSourceKws=()=>dedup([...sourceKws,...enrichedKws]);

  // Show login if not authed
  if(!authed) return <LoginScreen onSuccess={()=>setAuthed(true)}/>;

  function getDb(){
    const c=config.country.toLowerCase();
    if(c.includes("col"))return"co";if(c.includes("mex"))return"mx";
    if(c.includes("esp"))return"es";if(c.includes("fra"))return"fr";
    if(c.includes("usa")||c.includes("états")||c.includes("etats"))return"us";
    if(c.includes("arg"))return"ar";if(c.includes("br"))return"br";
    if(c.includes("ital"))return"it";if(c.includes("allem")||c.includes("deutsch"))return"de";
    return"us";
  }

  async function extractCategories(){
    setLoading(true);setErr("");setLoadMsg("Analyse du menu…");
    try{
      const resp=await callClaude(`Expert SEO. Extrais les sous-catégories produits génériques du menu. Règles : types de produits génériques uniquement · pas de marques SEB · pas de gammes/collections · pas de catégories générales · "A & B" → deux entrées · pratique → type produit. JSON : {"categories":["cat1",...]}`,`Site:${config.siteUrl}|Pays:${config.country}|Langue:${config.language}`,catImages);
      const{categories:cats}=parseJSON(resp);setCategories(cats);addLog(`✅ ${cats.length} catégories extraites`);
    }catch(e){setErr("Erreur: "+e.message);}
    setLoading(false);
  }
  async function runEnrichment(){
    setLoading(true);setErr("");setLoadMsg("Génération mots-clés sémantiques…");
    try{
      const resp=await callClaude(`Assistant SEO multilingue. Génère 10 requêtes SEO génériques, evergreen, hors-marque par catégorie. Langue locale uniquement. Formulations : meilleur, avis, acheter, pas cher, entretien, quel choisir… Jamais de marques, années, promos. JSON : {"enriched":[{"category":"...","keywords":["kw1",...10]},...]}`,`Pays:${config.country}|Langue:${config.language}\nCatégories:\n${categories.join("\n")}`);
      const{enriched}=parseJSON(resp);
      const preview=enriched.slice(0,2).map(e=>`${e.category}\n${e.keywords.slice(0,4).map(k=>"  • "+k).join("\n")}\n  …`).join("\n\n");
      setEnrichPreview(preview);setEnrichedKws(enriched.flatMap(e=>e.keywords.map(k=>({keyword:k,volume:0}))));
      addLog(`✅ ${enriched.flatMap(e=>e.keywords).length} mots-clés sémantiques générés`);
    }catch(e){setErr("Erreur: "+e.message);}
    setLoading(false);
  }
  async function findCompetitors(){
    setLoading(true);setErr("");setLoadMsg("Identification concurrents…");
    try{
      const resp=await callClaude(`Expert marché équipement domestique. Identifie 5 concurrents en ligne d'une marque Groupe SEB. Pour chacun : nom, domaine exact (sans https://), explication 1 phrase. JSON : {"competitors":[{"name":"...","url":"...","reason":"..."},...]}`,`Marque:${config.brand}|Pays:${config.country}|Langue:${config.language}|Site:${config.siteUrl}`);
      const{competitors:comps}=parseJSON(resp);setCompetitors(comps);addLog(`✅ ${comps.length} concurrents identifiés`);
    }catch(e){setErr("Erreur: "+e.message);}
    setLoading(false);
  }
  async function replaceComp(idx){
    setLoading(true);setErr("");setLoadMsg(`Remplacement…`);
    const snap=[...competitors];
    try{
      const resp=await callClaude(`Propose UN seul concurrent alternatif pertinent, différent des autres. JSON : {"name":"...","url":"...","reason":"..."}`,`Marque:${config.brand}|Pays:${config.country}\nÀ remplacer:${snap[idx].name}\nDéjà présents:${snap.filter((_,j)=>j!==idx).map(x=>x.name).join(", ")}`);
      const alt=parseJSON(resp);setCompetitors(prev=>prev.map((x,j)=>j===idx?alt:x));addLog(`🔄 ${snap[idx].name} → ${alt.name}`);
    }catch(e){setErr("Erreur: "+e.message);}
    setLoading(false);
  }
  async function importSourceFiles(files){
    let all=[];
    for(const f of files){let rows=[];try{rows=await parseFile(f);}catch(e){addLog(`⚠ ${f.name}: ${e.message}`);}all.push(...rows.map(r=>({keyword:r.keyword,volume:0})));addLog(`📄 ${f.name} → ${rows.length} mots-clés`);}
    setSourceKws(prev=>dedup([...prev,...all]));
  }
  async function importVolumeFile(files){
    const f=files[0];if(!f)return;
    let rows=[];try{rows=await parseFile(f);}catch(e){setErr(`Erreur ${f.name}: ${e.message}`);return;}
    const withVol=rows.filter(r=>r.volume>0);
    if(!withVol.length){setErr(`Aucun volume trouvé dans "${f.name}".`);return;}
    const volMap=new Map(withVol.map(r=>[r.keyword.toLowerCase().trim(),r.volume]));
    const sources=allSourceKws();
    const merged=sources.map(k=>({...k,volume:volMap.get(k.keyword.toLowerCase().trim())||0})).filter(k=>k.volume>0).sort((a,b)=>b.volume-a.volume);
    setRawWithVol(merged);setTargetCount(Math.min(500,merged.length));
    addLog(`📊 ${f.name} → ${merged.length} mots-clés avec volumes (${sources.length-merged.length} sans volume écartés)`);setErr("");
  }
  function runRuleBasedCleaning(){
    const before=rawWithVol.length;const cleaned=ruleBasedClean(rawWithVol);
    setCleanedKws(cleaned);setCleanStats({before,afterRules:cleaned.length});
    addLog(`✅ Nettoyage règles : ${cleaned.length}/${before} conservés`);
  }
  async function runClaudeDedup(){
    if(!cleanedKws.length)return;setLoading(true);setErr("");
    const BATCH=80;const total=Math.ceil(cleanedKws.length/BATCH);const toRemove=new Set();
    try{
      for(let b=0;b<total;b++){
        setLoadMsg(`Dédoublonnage — batch ${b+1}/${total}…`);
        const batch=cleanedKws.slice(b*BATCH,(b+1)*BATCH);
        const resp=await callClaude(`Liste de mots-clés SEO triés par volume décroissant. Identifie les quasi-doublons (pluriel, accent, faute, ordre des mots). Garde toujours celui avec le plus grand volume. Liste uniquement ceux à supprimer. JSON : {"remove":["kw1",...]}`,`Mots-clés (keyword|volume):\n${batch.map(k=>`${k.keyword}|${k.volume}`).join("\n")}`);
        const parsed=parseJSON(resp);for(const kw of(parsed.remove||[]))toRemove.add(kw.toLowerCase().trim());
      }
      const deduped=cleanedKws.filter(k=>!toRemove.has(k.keyword.toLowerCase().trim()));
      setCleanedKws(deduped);setCleanStats(s=>({...s,afterDedup:deduped.length,dedupRemoved:toRemove.size}));
      addLog(`✅ Dédoublonnage : ${deduped.length} conservés (${toRemove.size} supprimés)`);
    }catch(e){setErr("Erreur: "+e.message);}
    setLoading(false);
  }
  async function runClaudeSemanticPass(){
    if(!cleanedKws.length)return;setLoading(true);setErr("");
    const BATCH=80;const total=Math.ceil(cleanedKws.length/BATCH);const kept=[];
    try{
      for(let b=0;b<total;b++){
        setLoadMsg(`Affinage sémantique — batch ${b+1}/${total}…`);
        const batch=cleanedKws.slice(b*BATCH,(b+1)*BATCH);
        const resp=await callClaude(`Expert SEO. Supprime uniquement les mots-clés clairement hors-périmètre. Conserve tout ce qui est pertinent. JSON : {"kept":["kw1",...]}`,`Catégories:${categories.join(", ")}\nMots-clés:\n${batch.map(k=>k.keyword).join("\n")}`);
        const parsed=parseJSON(resp);const keptSet=new Set((parsed.kept||[]).map(k=>k.toLowerCase().trim()));
        kept.push(...batch.filter(k=>keptSet.has(k.keyword.toLowerCase().trim())));
      }
      setCleanedKws(kept);setCleanStats(s=>({...s,afterSemantic:kept.length}));addLog(`✅ Affinage sémantique : ${kept.length} conservés`);
    }catch(e){setErr("Erreur: "+e.message);}
    setLoading(false);
  }
  async function runCategorization(){
    setLoading(true);setErr("");
    const kws=cleanedKws.filter(k=>k.volume>0);const BATCH=80;const allResults=[];const total=Math.ceil(kws.length/BATCH);
    try{
      for(let b=0;b<total;b++){
        setLoadMsg(`Catégorisation batch ${b+1}/${total}…`);
        const batch=kws.slice(b*BATCH,(b+1)*BATCH);
        const resp=await callClaude(`Expert SEO et produit Groupe SEB. Pour chaque mot-clé assigne : 1. "category" : catégorie produit exacte parmi la liste fournie 2. "sebCategory" : parmi ${SEB_CATS.join(", ")} 100% assignés, jamais Uncategorized, ne pas modifier le texte. JSON : {"results":[{"keyword":"...","category":"...","sebCategory":"..."},...]}`,`Marque:${config.brand}|Pays:${config.country}|Langue:${config.language}\nCatégories:\n${categories.join("\n")}\nMots-clés:\n${batch.map(k=>`${k.keyword}|${k.volume}`).join("\n")}`);
        const parsed=parseJSON(resp);allResults.push(...(parsed.results||[]));
      }
      const map=new Map(allResults.map(r=>[r.keyword.toLowerCase().trim(),r]));
      const final=kws.map(k=>{const r=map.get(k.keyword.toLowerCase().trim());return{...k,category:r?.category||categories[0],sebCategory:r?.sebCategory||SEB_CATS[0]};});
      setCategorizedKws(final);setTargetCount(Math.min(targetCount,final.length));
      addLog(`✅ Catégorisation : ${final.length} mots-clés (${total} batches)`);nav(6);
    }catch(e){setErr("Erreur catégorisation: "+e.message);}
    setLoading(false);
  }

  const H2=({c})=><h2 style={{fontFamily:"'Space Grotesk',sans-serif",fontWeight:700,fontSize:20,margin:"0 0 8px",color:"#fff"}}>{c}</h2>;
  const Sub=({c})=><p style={{fontSize:13,color:G.sub,marginBottom:20,lineHeight:1.6}}>{c}</p>;

  return(
    <div style={{fontFamily:"'IBM Plex Mono','Courier New',monospace",background:G.bg,minHeight:"100vh",color:G.text}}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600&family=Space+Grotesk:wght@400;500;700&display=swap');
        html,body,#root{background:${G.bg};margin:0;padding:0;}
        *{box-sizing:border-box}input,textarea{font-family:'IBM Plex Mono',monospace}
        ::-webkit-scrollbar{width:5px}::-webkit-scrollbar-track{background:#0e1117}::-webkit-scrollbar-thumb{background:#1e2a3d;border-radius:3px}
        button:hover:not(:disabled){filter:brightness(1.12)}@keyframes spin{to{transform:rotate(360deg)}}
      `}</style>

      {/* Header */}
      <div style={{background:"#0a0d14",borderBottom:`1px solid ${G.border}`,padding:"14px 24px",display:"flex",alignItems:"center",gap:14,flexWrap:"wrap"}}>
        <div style={{width:34,height:34,background:"linear-gradient(135deg,#3b82f6,#22c55e)",borderRadius:9,display:"flex",alignItems:"center",justifyContent:"center",fontSize:16}}>⚡</div>
        <div><div style={{fontFamily:"'Space Grotesk',sans-serif",fontWeight:700,fontSize:15,color:"#fff"}}>SEO Keyword Pipeline</div><div style={{fontSize:11,color:G.muted}}>Groupe SEB</div></div>
        <div style={{marginLeft:"auto",display:"flex",alignItems:"center",gap:12}}>
          <StepBar step={step}/>
          {step>0&&<button onClick={()=>{if(!window.confirm("Recommencer ? Toutes les données seront effacées."))return;setStep(0);setErr("");setLog([]);setConfig({brand:"",country:"",language:"",siteUrl:""});setCatImages([]);setCategories([]);setCatManual("");setEnrichedKws([]);setEnrichPreview("");setEnrichValidated(false);setKwSub(0);setCompetitors([]);setSourceKws([]);setRawWithVol([]);setCleanedKws([]);setCleanStats(null);setCategorizedKws([]);setTargetCount(500);}} style={{background:"none",border:`1px solid ${G.red}44`,color:G.red+"cc",padding:"5px 12px",borderRadius:6,fontFamily:"inherit",fontSize:11,cursor:"pointer",whiteSpace:"nowrap",flexShrink:0}}>🔄 Nouveau projet</button>}
        </div>
      </div>
      <div style={{height:3,background:G.border}}><div style={{height:"100%",width:`${(step/6)*100}%`,background:"linear-gradient(90deg,#3b82f6,#22c55e)",transition:"width .4s"}}/></div>

      <div style={{maxWidth:860,margin:"0 auto",padding:"28px 20px"}}>
        <ErrBox msg={err}/>

        {step===0&&<Card>
          <H2 c="📋 Nouveau projet"/>
          <Sub c="Ces informations servent de contexte à Claude pour toutes les étapes."/>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
            <Inp label="Marque" placeholder="ex: Imusa" value={config.brand} onChange={v=>setConfig(c=>({...c,brand:v}))}/>
            <Inp label="Pays cible" placeholder="ex: USA" value={config.country} onChange={v=>setConfig(c=>({...c,country:v}))}/>
            <Inp label="Langue" placeholder="ex: Anglais" value={config.language} onChange={v=>setConfig(c=>({...c,language:v}))}/>
            <Inp label="URL du site" placeholder="ex: imusausa.com" value={config.siteUrl} onChange={v=>setConfig(c=>({...c,siteUrl:v}))}/>
          </div>
          <div style={{marginTop:24,display:"flex",justifyContent:"flex-end"}}>
            <Btn onClick={()=>{nav(1);addLog(`📋 ${config.brand} / ${config.country} / ${config.language}`);}} disabled={!config.brand||!config.country||!config.language}>Continuer →</Btn>
          </div>
        </Card>}

        {step===1&&<Card>
          <H2 c="🗂 Catégories + enrichissement sémantique"/>
          <Sub c="Extraire les catégories du menu, puis générer 10 mots-clés génériques par catégorie."/>
          <Tag n="2.1">Captures d'écran du menu → extraction automatique</Tag>
          <Info>Screenshots de toutes les entrées et sous-entrées du menu → Claude extrait les catégories produits.</Info>
          <Drop label="Captures d'écran du menu (PNG, JPG)" note="Toutes les entrées et sous-entrées" accept="image/*" multiple onFiles={async files=>{const imgs=await Promise.all(files.map(async f=>({name:f.name,data:await readFileAsBase64(f),type:f.type})));setCatImages(prev=>[...prev,...imgs]);}}>
            {catImages.length>0&&<div style={{marginTop:8,display:"flex",flexWrap:"wrap",gap:4,justifyContent:"center"}}>{catImages.map((img,i)=><span key={i} style={{fontSize:10,background:G.accent+"22",border:`1px solid ${G.accent}44`,color:G.accent,padding:"2px 7px",borderRadius:4}}>🖼 {img.name}</span>)}</div>}
          </Drop>
          {catImages.length>0&&!loading&&categories.length===0&&<div style={{marginTop:10,display:"flex",justifyContent:"flex-end"}}><Btn onClick={extractCategories}>🤖 Extraire les catégories</Btn></div>}
          <div style={{margin:"16px 0",display:"flex",alignItems:"center",gap:10}}><div style={{flex:1,height:1,background:G.border}}/><span style={{fontSize:11,color:G.muted}}>OU</span><div style={{flex:1,height:1,background:G.border}}/></div>
          <textarea value={catManual} onChange={e=>setCatManual(e.target.value)} rows={5} placeholder={"Sartenes\nOllas\nCacerolas\n..."} style={{width:"100%",background:G.faint,border:`1px solid ${G.border}`,color:G.text,padding:"10px 14px",borderRadius:7,fontSize:13,resize:"vertical",outline:"none",marginBottom:8}}/>
          <div style={{display:"flex",justifyContent:"flex-end",marginBottom:16}}>
            <Btn secondary onClick={()=>{const c=catManual.split("\n").map(x=>x.trim()).filter(Boolean);setCategories(c);addLog(`✅ ${c.length} catégories`);}} disabled={!catManual.trim()}>Utiliser ces catégories</Btn>
          </div>
          {loading&&<Spin msg={loadMsg}/>}
          {categories.length>0&&<>
            <div style={{fontSize:12,color:G.green,marginBottom:8}}>✅ {categories.length} catégories :</div>
            <div style={{display:"flex",flexWrap:"wrap",gap:5,marginBottom:20}}>{categories.map((c,i)=><span key={i} style={{display:"inline-flex",alignItems:"center",gap:4,background:G.accent+"18",border:`1px solid ${G.accent}33`,color:G.accent,padding:"3px 10px",borderRadius:20,fontSize:12}}>{c}<button onClick={()=>setCategories(p=>p.filter((_,j)=>j!==i))} style={{background:"none",border:"none",color:G.muted,cursor:"pointer",padding:0,fontSize:11}}>×</button></span>)}</div>
            <div style={{borderTop:`1px solid ${G.border}`,paddingTop:18}}>
              <Tag n="2.3">Enrichissement — 10 mots-clés SEO génériques par catégorie</Tag>
              <Info color={G.green}>Claude génère des requêtes evergreen hors-marque qui seront ajoutées à ta liste avant les volumes.</Info>
              {!loading&&enrichedKws.length===0&&<Btn onClick={runEnrichment} color={G.green}>🤖 Générer</Btn>}
              {enrichPreview&&!enrichValidated&&<><div style={{background:G.faint,borderRadius:8,padding:12,fontSize:12,color:G.sub,whiteSpace:"pre-wrap",lineHeight:1.7,marginTop:12,marginBottom:10}}>{enrichPreview}</div><div style={{display:"flex",gap:10}}><SBtn onClick={runEnrichment} color={G.yellow}>🔄 Regénérer</SBtn><Btn onClick={()=>{setEnrichValidated(true);addLog(`✅ ${enrichedKws.length} mots-clés sémantiques validés`);}} color={G.green}>✅ Valider</Btn></div></>}
              {enrichValidated&&<div style={{background:G.green+"14",border:`1px solid ${G.green}44`,borderRadius:8,padding:"10px 14px",fontSize:12,color:G.green}}>✅ {enrichedKws.length} mots-clés sémantiques validés</div>}
            </div>
          </>}
          <div style={{marginTop:24,display:"flex",justifyContent:"space-between"}}>
            <Btn secondary onClick={()=>nav(0)}>← Retour</Btn>
            <Btn onClick={()=>nav(2)} disabled={categories.length===0}>Continuer →</Btn>
          </div>
        </Card>}

        {step===2&&<>
          {kwSub===0&&<Card>
            <H2 c="📥 Collecte des mots-clés"/>
            <Sub c="Récupère toutes les sources de mots-clés disponibles."/>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
              <div onClick={()=>setKwSub(1)} style={{background:G.faint,border:`1px solid ${G.border}`,borderRadius:10,padding:20,cursor:"pointer"}} onMouseEnter={e=>e.currentTarget.style.borderColor=G.accent} onMouseLeave={e=>e.currentTarget.style.borderColor=G.border}>
                <div style={{fontSize:26,marginBottom:8}}>🔍</div><div style={{fontSize:14,color:G.text,fontWeight:700,marginBottom:5}}>Je pars de zéro</div>
                <div style={{fontSize:12,color:G.sub,lineHeight:1.6}}>Claude identifie les concurrents et génère les liens SEMrush Keyword Gap.</div>
              </div>
              <div onClick={()=>setKwSub(3)} style={{background:G.faint,border:`1px solid ${G.border}`,borderRadius:10,padding:20,cursor:"pointer"}} onMouseEnter={e=>e.currentTarget.style.borderColor=G.accent} onMouseLeave={e=>e.currentTarget.style.borderColor=G.border}>
                <div style={{fontSize:26,marginBottom:8}}>📂</div><div style={{fontSize:14,color:G.text,fontWeight:700,marginBottom:5}}>J'ai déjà des fichiers</div>
                <div style={{fontSize:12,color:G.sub,lineHeight:1.6}}>SEMrush, anciennes listes, SEA, retailers… importe directement.</div>
              </div>
            </div>
            <div style={{marginTop:20}}><Btn secondary onClick={()=>nav(1)}>← Retour</Btn></div>
          </Card>}

          {kwSub===1&&<Card>
            <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:20}}>
              <Btn secondary onClick={()=>setKwSub(0)} style={{padding:"5px 12px",fontSize:12}}>← Retour</Btn>
              <div style={{flex:1,height:1,background:G.border}}/><span style={{fontSize:11,color:G.muted}}>1/3 — Concurrents</span>
            </div>
            <H2 c="🔍 Identification des concurrents"/>
            {loading?<Spin msg={loadMsg}/>:<>
              {competitors.length===0&&<Btn onClick={findCompetitors}>🤖 Générer automatiquement</Btn>}
              {competitors.length>0&&<>
                <div style={{fontSize:12,color:G.green,marginBottom:10}}>✅ {competitors.length} concurrent{competitors.length>1?"s":""} :</div>
                {competitors.map((c,i)=>{const idx=i;return <div key={i} style={{background:G.faint,border:`1px solid ${G.border}`,borderRadius:8,padding:"12px 14px",marginBottom:8}}>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
                    <div><div style={{fontSize:10,color:G.muted,marginBottom:3,textTransform:"uppercase",letterSpacing:".06em"}}>Nom</div><input value={c.name} onChange={e=>setCompetitors(prev=>prev.map((x,j)=>j===i?{...x,name:e.target.value}:x))} style={{background:G.card,border:`1px solid ${G.border}`,color:G.text,padding:"7px 10px",borderRadius:5,fontFamily:"inherit",fontSize:12,width:"100%",outline:"none"}}/></div>
                    <div><div style={{fontSize:10,color:G.muted,marginBottom:3,textTransform:"uppercase",letterSpacing:".06em"}}>Domaine</div><input value={c.url} onChange={e=>setCompetitors(prev=>prev.map((x,j)=>j===i?{...x,url:e.target.value}:x))} style={{background:G.card,border:`1px solid ${G.border}`,color:G.accent,padding:"7px 10px",borderRadius:5,fontFamily:"inherit",fontSize:12,width:"100%",outline:"none"}}/></div>
                  </div>
                  <div style={{fontSize:11,color:G.sub,marginBottom:8}}>{c.reason}</div>
                  <div style={{display:"flex",gap:8}}><SBtn onClick={()=>replaceComp(idx)} color={G.yellow}>🔄 Non pertinent — remplacer</SBtn><SBtn onClick={()=>setCompetitors(prev=>prev.filter((_,j)=>j!==idx))} color={G.red}>🗑 Supprimer</SBtn></div>
                </div>;})}
                <AddCompForm onAdd={c=>setCompetitors(prev=>[...prev,c])}/>
                <div style={{marginTop:16,display:"flex",gap:10,justifyContent:"space-between"}}><Btn secondary onClick={findCompetitors} style={{fontSize:12}}>🔄 Regénérer</Btn><Btn onClick={()=>setKwSub(2)}>Valider → Liens SEMrush</Btn></div>
              </>}
            </>}
          </Card>}

          {kwSub===2&&<Card>
            <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:20}}>
              <Btn secondary onClick={()=>setKwSub(1)} style={{padding:"5px 12px",fontSize:12}}>← Retour</Btn>
              <div style={{flex:1,height:1,background:G.border}}/><span style={{fontSize:11,color:G.muted}}>2/3 — SEMrush</span>
            </div>
            <H2 c="🔗 Liens Keyword Gap SEMrush"/>
            <Info>Ouvre chaque vue, exporte en CSV/XLSX, reviens importer. <strong style={{color:G.text}}>⚠ Un concurrent à la fois.</strong></Info>
            {competitors.map((c,i)=>{
              const site="https://"+config.siteUrl.replace(/^https?:\/\//,"").replace(/\/$/,"");
              const comp=c.url.replace(/^https?:\/\//,"").replace(/\/$/,"");
              const db=getDb();
              const f50=encodeURIComponent(JSON.stringify({search:null,volume:null,kd:null,intent:null,position:{value:"0-50",type:"competitors"},advanced:{}}));
              const cw=encodeURIComponent(comp+":domain:organic");
              const base=`https://www.semrush.com/analytics/keywordgap/?q=${encodeURIComponent(site)}&searchType=domain&keywordType=organic&compareWith=${cw}&db=${db}`;
              return <div key={i} style={{background:G.faint,border:`1px solid ${G.border}`,borderRadius:10,padding:16,marginBottom:12}}>
                <div style={{marginBottom:10}}><span style={{fontSize:13,color:G.text,fontWeight:700}}>{config.brand}</span><span style={{color:G.muted}}> vs </span><span style={{fontSize:13,color:G.accent,fontWeight:700}}>{c.name}</span><div style={{fontSize:11,color:G.muted,marginTop:2}}>{config.siteUrl} ↔ {comp}</div></div>
                <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:8}}>
                  {[{label:"Shared",rt:"common",color:G.accent,desc:"Mots-clés en commun",f:""},{label:"Missing",rt:"missing",color:G.yellow,desc:"Top 50 concurrent, absent chez vous",f:`&filter=${f50}`},{label:"Untapped",rt:"untapped",color:G.green,desc:"Top 50 concurrent, opportunités",f:`&filter=${f50}`}].map(({label,rt,color,desc,f})=>(
                    <a key={rt} href={`${base}&rankType=${rt}${f}`} target="_blank" rel="noreferrer" style={{display:"block",background:color+"14",border:`1px solid ${color}44`,borderRadius:7,padding:"10px 12px",textDecoration:"none"}}>
                      <div style={{fontSize:12,color,fontWeight:700,marginBottom:3}}>{label} →</div>
                      <div style={{fontSize:10,color:G.muted,lineHeight:1.4}}>{desc}</div>
                    </a>
                  ))}
                </div>
              </div>;
            })}
            <div style={{display:"flex",justifyContent:"flex-end",marginTop:8}}><Btn onClick={()=>setKwSub(3)}>Import des fichiers →</Btn></div>
          </Card>}

          {kwSub===3&&<Card>
            {competitors.length>0&&<div style={{display:"flex",alignItems:"center",gap:10,marginBottom:20}}><Btn secondary onClick={()=>setKwSub(2)} style={{padding:"5px 12px",fontSize:12}}>← Retour SEMrush</Btn><div style={{flex:1,height:1,background:G.border}}/><span style={{fontSize:11,color:G.muted}}>3/3 — Import sources</span></div>}
            <H2 c="📂 Import des sources de mots-clés"/>
            <Info>Tous tes fichiers sources ici : SEMrush, anciennes listes, SEA, retailers…<br/><strong style={{color:G.text}}>Uniquement les keywords — les volumes viennent à l'étape suivante.</strong></Info>
            <Drop label="Tous tes fichiers sources" note="XLSX, CSV, TSV — tout en même temps" accept=".csv,.tsv,.txt,.xlsx,.xls" multiple onFiles={importSourceFiles}>
              {sourceKws.length>0&&<div style={{marginTop:10,display:"flex",alignItems:"center",gap:10,justifyContent:"center",flexWrap:"wrap"}}>
                <span style={{fontSize:13,color:G.green,fontWeight:600}}>✅ {sourceKws.length} mots-clés uniques</span>
                <SBtn onClick={e=>{e.stopPropagation();dl(dedup([...sourceKws,...enrichedKws]).map(k=>k.keyword).join("\n"),`${config.brand}_preview_sources.txt`);}} color={G.accent}>📥 Aperçu TXT</SBtn>
                <button onClick={e=>{e.stopPropagation();setSourceKws([]);}} style={{background:"none",border:`1px solid ${G.border}`,color:G.muted,padding:"4px 10px",borderRadius:5,fontFamily:"inherit",fontSize:11,cursor:"pointer"}}>Vider</button>
              </div>}
            </Drop>
            {(sourceKws.length>0||enrichedKws.length>0)&&<div style={{marginTop:10,background:G.faint,borderRadius:8,padding:"10px 14px",fontSize:12,color:G.sub}}><strong style={{color:G.text}}>Total collecté :</strong> {allSourceKws().length} mots-clés{enrichedKws.length>0&&<span style={{color:G.green}}> (dont {enrichedKws.length} sémantiques)</span>}</div>}
            <div style={{display:"flex",justifyContent:"space-between",marginTop:20}}>
              <Btn secondary onClick={()=>setKwSub(competitors.length>0?2:0)}>← Retour</Btn>
              <Btn onClick={()=>nav(3)} disabled={allSourceKws().length===0}>Continuer → Import volumes</Btn>
            </div>
          </Card>}
        </>}

        {step===3&&<Card>
          <H2 c="📊 Import des volumes"/>
          <Sub c="Les volumes arrivent AVANT le nettoyage — c'est eux qui guident le dédoublonnage."/>
          <Info color={G.green}>
            <strong style={{color:G.text}}>Flux :</strong><br/>
            1. Télécharge la liste pour KWP → <SBtn onClick={()=>dl(allSourceKws().map(k=>k.keyword).join("\n"),`${config.brand}_${config.country}_pour_KWP.txt`)} color={G.accent}>📥 Télécharger ({allSourceKws().length} mots-clés)</SBtn><br/>
            2. <a href="https://ads.google.com/aw/keywordplanner/ideas/existing" target="_blank" rel="noreferrer" style={{color:G.accent}}>Google Keyword Planner</a> → Get search volume → pays → exporte CSV<br/>
            3. Glisse ce fichier ici.
          </Info>
          <Drop label="Fichier de volumes" note="Keyword Planner CSV · SEMrush XLSX · tout fichier keyword + volume" accept=".csv,.tsv,.txt,.xlsx,.xls" multiple={false} onFiles={importVolumeFile}>
            {rawWithVol.length>0&&<div style={{marginTop:8}}><span style={{fontSize:13,color:G.green,fontWeight:600}}>✅ {rawWithVol.length} mots-clés avec volumes</span><span style={{fontSize:11,color:G.muted,marginLeft:10}}>Max : {rawWithVol.reduce((m,r)=>r.volume>m?r.volume:m,0).toLocaleString()}</span></div>}
          </Drop>
          {rawWithVol.length>0&&<>
            <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10,marginTop:16}}>
              <Stat value={allSourceKws().length} label="Sources collectées" color={G.muted}/>
              <Stat value={rawWithVol.length} label="Avec volumes" color={G.accent}/>
              <Stat value={rawWithVol.filter(k=>k.volume>1000).length} label="Volume > 1 000" color={G.green}/>
            </div>
            <div style={{marginTop:12,maxHeight:160,overflowY:"auto",fontSize:12}}>
              {rawWithVol.slice(0,10).map((k,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",padding:"3px 0",borderBottom:`1px solid ${G.faint}`}}><span style={{color:G.sub}}>{k.keyword}</span><span style={{color:G.accent}}>{k.volume.toLocaleString()}</span></div>)}
              {rawWithVol.length>10&&<div style={{color:G.muted,fontSize:11,padding:"3px 0"}}>… {rawWithVol.length-10} autres</div>}
            </div>
          </>}
          <div style={{display:"flex",justifyContent:"space-between",marginTop:24}}>
            <Btn secondary onClick={()=>nav(2)}>← Retour</Btn>
            <Btn onClick={()=>nav(4)} disabled={rawWithVol.length===0}>Nettoyage →</Btn>
          </div>
        </Card>}

        {step===4&&<Card>
          <H2 c="🧹 Nettoyage + dédoublonnage intelligent"/>
          <Sub c="Volumes disponibles — les quasi-doublons sont gérés par volume : on garde toujours le meilleur."/>
          <Tag n="Passe 1">Nettoyage par règles — instantané</Tag>
          <Info>Supprime automatiquement : marques connues · recettes · années · promos.</Info>
          <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10,marginBottom:14}}>
            <Stat value={rawWithVol.length} label="Avec volumes" color={G.yellow}/>
            <Stat value={cleanStats?.afterRules??"?"} label="Après règles" color={G.accent}/>
            <Stat value={cleanStats?.afterDedup??cleanStats?.afterRules??"?"} label="Après dédup." color={G.green}/>
          </div>
          {!cleanStats?<Btn onClick={()=>{runRuleBasedCleaning();addLog("🧹 Nettoyage règles lancé");}}>⚡ Lancer le nettoyage (instantané)</Btn>
            :<div style={{background:G.green+"14",border:`1px solid ${G.green}44`,borderRadius:8,padding:"12px 14px",fontSize:12,color:G.green,marginBottom:12}}>
              ✅ Règles : {cleanStats.afterRules} conservés{cleanStats.afterDedup!==undefined&&` · Dédup : ${cleanStats.afterDedup} (${cleanStats.dedupRemoved} supprimés)`}{cleanStats.afterSemantic!==undefined&&` · Sémantique : ${cleanStats.afterSemantic}`}
            </div>}
          {cleanStats&&<>
            <div style={{marginBottom:16}}>
              <div style={{fontSize:12,color:G.text,fontWeight:600,marginBottom:8}}>Liste actuelle ({cleanedKws.length} mots-clés) :</div>
              <div style={{maxHeight:220,overflowY:"auto",background:G.faint,borderRadius:8,padding:"8px 12px"}}>
                <div style={{display:"grid",gridTemplateColumns:"1fr 80px 24px",paddingBottom:6,position:"sticky",top:0,background:G.faint}}>
                  <div style={{fontSize:10,color:G.muted,textTransform:"uppercase",letterSpacing:".06em"}}>Mot-clé</div>
                  <div style={{fontSize:10,color:G.muted,textTransform:"uppercase",letterSpacing:".06em",textAlign:"right"}}>Volume</div>
                </div>
                {cleanedKws.map((k,i)=><div key={i} style={{display:"grid",gridTemplateColumns:"1fr 80px 24px",padding:"3px 0",borderBottom:`1px solid ${G.border}`}}>
                  <span style={{color:G.sub,fontSize:12}}>{k.keyword}</span>
                  <span style={{color:G.accent,fontSize:12,textAlign:"right"}}>{k.volume.toLocaleString()}</span>
                  <button onClick={()=>setCleanedKws(prev=>prev.filter((_,j)=>j!==i))} style={{background:"none",border:"none",color:G.red+"88",cursor:"pointer",fontSize:11,padding:0}}>✕</button>
                </div>)}
              </div>
              <div style={{fontSize:11,color:G.muted,marginTop:4}}>Clique ✕ pour supprimer manuellement</div>
            </div>
            <div style={{borderTop:`1px solid ${G.border}`,paddingTop:18,display:"flex",flexDirection:"column",gap:12}}>
              <Tag n="Passe 2">Dédoublonnage intelligent — similarité + volume</Tag>
              <Info color={G.yellow}>Claude repère les quasi-doublons et garde toujours celui avec le plus grand volume. Batches de 80.</Info>
              {loading&&loadMsg.includes("Dédoublon")?<Spin msg={loadMsg}/>:<Btn onClick={runClaudeDedup} color={G.yellow}>🤖 Dédoublonner intelligemment</Btn>}
            </div>
            <div style={{borderTop:`1px solid ${G.border}`,paddingTop:18,marginTop:4}}>
              <Tag n="Passe 3">Affinage sémantique — optionnel</Tag>
              <Info color={G.muted}>Claude supprime les mots-clés hors-périmètre restants. Batches de 80.</Info>
              {loading&&loadMsg.includes("sémantique")?<Spin msg={loadMsg}/>:<Btn secondary onClick={runClaudeSemanticPass}>🤖 Affinage sémantique (optionnel)</Btn>}
            </div>
          </>}
          <div style={{display:"flex",justifyContent:"space-between",marginTop:24}}>
            <Btn secondary onClick={()=>nav(3)}>← Retour</Btn>
            <Btn onClick={()=>nav(5)} disabled={cleanedKws.length===0}>Catégorisation →</Btn>
          </div>
        </Card>}

        {step===5&&<Card>
          <H2 c="🏷 Catégorisation"/>
          <Sub c="Catégorie produit + grande catégorie SEB. Traitement par lots de 80."/>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:16}}>
            <div><div style={{fontSize:11,color:G.muted,marginBottom:8,textTransform:"uppercase",letterSpacing:".06em"}}>Catégories site ({categories.length})</div><div style={{display:"flex",flexWrap:"wrap",gap:5}}>{categories.map((c,i)=><span key={i} style={{background:G.accent+"18",border:`1px solid ${G.accent}33`,color:G.accent,padding:"3px 10px",borderRadius:20,fontSize:11}}>{c}</span>)}</div></div>
            <div><div style={{fontSize:11,color:G.muted,marginBottom:8,textTransform:"uppercase",letterSpacing:".06em"}}>Catégories SEB (9)</div><div style={{display:"flex",flexWrap:"wrap",gap:5}}>{SEB_CATS.map((c,i)=><span key={i} style={{background:G.green+"18",border:`1px solid ${G.green}33`,color:G.green,padding:"3px 10px",borderRadius:20,fontSize:11}}>{c}</span>)}</div></div>
          </div>
          <Stat value={cleanedKws.filter(k=>k.volume>0).length} label="Mots-clés à catégoriser"/>
          {loading?<Spin msg={loadMsg}/>:<div style={{display:"flex",justifyContent:"space-between",marginTop:20}}><Btn secondary onClick={()=>nav(4)}>← Retour</Btn><Btn onClick={runCategorization}>🤖 Catégoriser</Btn></div>}
        </Card>}

        {step===6&&<Card>
          <H2 c="✅ Sélection finale + Export"/>
          <Sub c="Choisis ta cible — l'app répartit proportionnellement par catégorie en conservant le ratio naturel."/>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:20}}>
            <Stat value={allSourceKws().length} label="Sources" color={G.muted}/>
            <Stat value={rawWithVol.length} label="Avec volumes" color={G.yellow}/>
            <Stat value={cleanedKws.length} label="Après nettoyage" color={G.accent}/>
            <Stat value={categorizedKws.length} label="Catégorisés" color={G.green}/>
          </div>
          {categorizedKws.length>0&&(()=>{
            const balanced=buildBalancedList(categorizedKws,targetCount);
            const groups={};for(const k of balanced)groups[k.category]=(groups[k.category]||0)+1;
            return <>
              <div style={{background:G.faint,border:`1px solid ${G.border}`,borderRadius:10,padding:20,marginBottom:20}}>
                <div style={{fontSize:13,color:G.text,fontWeight:600,marginBottom:14}}>🎯 Taille de la liste finale</div>
                <div style={{display:"flex",alignItems:"center",gap:16,marginBottom:12}}>
                  <input type="range" min={50} max={Math.min(2000,categorizedKws.length)} step={50} value={targetCount} onChange={e=>setTargetCount(Number(e.target.value))} style={{flex:1,accentColor:G.accent}}/>
                  <div style={{fontSize:24,fontWeight:700,color:G.accent,fontFamily:"'Space Grotesk',sans-serif",minWidth:60,textAlign:"right"}}>{targetCount}</div>
                </div>
                <div style={{fontSize:12,color:G.sub,marginBottom:10}}><strong style={{color:G.text}}>{Object.keys(groups).length} catégories</strong> · ratio naturel · meilleurs volumes</div>
                <div style={{display:"flex",flexWrap:"wrap",gap:5,marginBottom:16}}>{Object.entries(groups).sort((a,b)=>b[1]-a[1]).map(([cat,n])=><span key={cat} style={{fontSize:11,background:G.accent+"18",border:`1px solid ${G.accent}33`,color:G.accent,padding:"2px 8px",borderRadius:12}}>{cat} <strong>{n}</strong></span>)}</div>
                <Btn onClick={()=>{try{dlCSV([["Main SEB Category","Keyword","Avg. Monthly Volume","Product Category"],...balanced.map(k=>[k.sebCategory||"",k.keyword,k.volume,k.category||""])],`${config.brand}_${config.country}_keywords_${targetCount}.csv`);setTimeout(()=>addLog(`📥 ${balanced.length} mots-clés exportés`),0);}catch(e){setErr("Erreur export: "+e.message);}}}>📥 Télécharger — {targetCount} mots-clés</Btn>
              </div>
              <div style={{maxHeight:200,overflowY:"auto",fontSize:12}}>
                <div style={{display:"grid",gridTemplateColumns:"140px 1fr 70px 130px",position:"sticky",top:0,background:G.card,paddingBottom:5}}>{["SEB Cat.","Mot-clé","Volume","Catégorie"].map(h=><div key={h} style={{color:G.muted,fontSize:10,textTransform:"uppercase",letterSpacing:".06em",padding:"4px 6px"}}>{h}</div>)}</div>
                {balanced.slice(0,20).map((k,i)=><div key={i} style={{display:"grid",gridTemplateColumns:"140px 1fr 70px 130px",background:i%2===0?G.faint:"transparent",borderRadius:4}}>
                  <div style={{padding:"5px 6px",fontSize:11,color:"#7ab3ff"}}>{k.sebCategory}</div>
                  <div style={{padding:"5px 6px",color:G.text}}>{k.keyword}</div>
                  <div style={{padding:"5px 6px",color:G.accent}}>{k.volume>0?k.volume.toLocaleString():"—"}</div>
                  <div style={{padding:"5px 6px",color:G.sub,fontSize:11}}>{k.category}</div>
                </div>)}
                {balanced.length>20&&<div style={{color:G.muted,padding:"6px",fontSize:11}}>… {balanced.length-20} dans l'export</div>}
              </div>
            </>;
          })()}
          <div style={{display:"flex",justifyContent:"flex-start",marginTop:20}}><Btn secondary onClick={()=>nav(5)}>← Retour</Btn></div>
        </Card>}

        {log.length>0&&<div style={{marginTop:24}}>
          <div style={{fontSize:11,color:G.muted,marginBottom:6,textTransform:"uppercase",letterSpacing:".07em"}}>Journal</div>
          <div style={{background:G.card,border:`1px solid ${G.border}`,borderRadius:8,padding:"12px 16px"}}>
            {log.map((l,i)=><div key={i} style={{fontSize:12,color:G.sub,padding:"3px 0",borderBottom:`1px solid ${G.faint}`}}>{l}</div>)}
          </div>
        </div>}
      </div>
    </div>
  );
}
