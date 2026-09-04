const $=id=>document.getElementById(id);
const editor=$("editor"), form=$("pileForm"), rows=$("pileRows"), message=$("message");
let piles=[];

const boringFields=["date","time_from","time_to","depth","description","remarks"];
const reinFields=["description","diameter","nos","length"];

const topLevelFields=[
  "company_name","client_name","serial_no","report_date","project_location",
  "pile_id","rig_no","location","drawing_no","dia_of_pile","depth_of_pile_ctl","founding_level",
  "type","coordinates","cutoff_level","top_of_casing","existing_gl","cage_length","boring_completed_date",
  "cage_started_date","cage_started_time","cage_completed_date","cage_completed_time",
  "trimmer_started_date","trimmer_started_time","trimmer_completed_date","trimmer_completed_time",
  "flushing_started_date","flushing_started_time","flushing_completed_date","flushing_completed_time",
  "specific_gravity","bentonite_before","bentonite_after",
  "cement_type","concrete_grade","design_mix","slump","cubes_taken",
  "concrete_commence_date","concrete_commence_time","concrete_completed_date","concrete_completed_time",
  "actual_concrete_volume"
];

function makeInput(cls,type="text"){
  const i=document.createElement("input"); i.className=cls; i.type=type; return i;
}

function addBoringRow(values){
  const tr=document.createElement("tr");
  ["date","time","time","number","text","text"].forEach((t,i)=>{
    const td=document.createElement("td"); const inp=makeInput("b-"+boringFields[i],t);
    if(t==="number") inp.step="0.001";
    if(values) inp.value = values[boringFields[i]] ?? "";
    td.appendChild(inp); tr.appendChild(td);
  });
  const td=document.createElement("td");
  const del=document.createElement("button"); del.type="button"; del.textContent="Remove"; del.className="danger small-btn";
  del.onclick=()=>tr.remove();
  td.appendChild(del); tr.appendChild(td);
  $("boringBody").appendChild(tr);
}

function buildRows(boringData){
  $("boringBody").innerHTML="";
  const initial = Array.isArray(boringData) && boringData.length ? boringData : [null,null,null,null,null,null];
  initial.forEach(v=>addBoringRow(v));

  $("reinforcementBody").innerHTML="";
  const defaults=[
    ["Vertical Top",25,5,10.05],["Vertical Top",25,5,11.675],["Vertical Top",25,10,9.95],
    ["Vertical Bottom",16,5,""],["Vertical Bottom",16,5,""],["Master Ring",16,"","1.513"],
    ["Helical Ring Top 75mm Spacing",8,1,""],["Helical Ring Bottom 200mm Spacing",8,1,""]
  ];
  defaults.forEach(d=>{
    const tr=document.createElement("tr");
    d.forEach((v,i)=>{
      const td=document.createElement("td");
      const isNumber = i===1 || i===2 || i===3;
      const inp = makeInput("r-"+reinFields[i], isNumber ? "number" : "text");
      if(i===3){ inp.step = "0.001"; inp.min = "0"; inp.inputMode = "decimal"; }
      else if(isNumber) { inp.step = "1"; inp.min = "0"; }
      inp.value=v;
      td.appendChild(inp);
      tr.appendChild(td);
    });
    $("reinforcementBody").appendChild(tr);
  });
}
function setVal(id,v){const el=$(id); if(el) el.value=v??""}
function resetForm(){
  form.reset(); $("editingId").value=""; $("formTitle").textContent="New Pile"; message.textContent="";
  buildRows();
}
function openNew(){resetForm();editor.classList.remove("hidden");editor.scrollIntoView({behavior:"smooth"})}
async function load(){
  const r=await fetch("/api/piles"); piles=await r.json();
  $("count").textContent=piles.length; $("latest").textContent=piles[0]?.pile_id||"—"; render();
}
function render(){
  const q=$("search").value.toLowerCase();
  const filtered=piles.filter(p=>(p.pile_id+" "+(p.coordinates||"")).toLowerCase().includes(q));
  rows.innerHTML=filtered.length?filtered.map(p=>`
    <tr><td><b>${esc(p.pile_id)}</b></td><td>${esc(p.coordinates||"")}</td><td>${esc(p.boring_logs?.[0]?.date||"")}</td>
    <td>${p.actual_concrete_volume??""}</td><td>${esc(p.updated_at||"")}</td>
    <td><button data-edit="${esc(p.pile_id)}">Edit</button> <button class="danger" data-del="${esc(p.pile_id)}">Delete</button></td></tr>`).join("")
    : `<tr><td colspan="6" class="empty">No piles found.</td></tr>`;
}
function esc(s){return String(s).replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[m]))}
async function edit(id){
  const p=await (await fetch("/api/piles/"+encodeURIComponent(id))).json();
  resetForm(); $("editingId").value=p.pile_id; $("formTitle").textContent="Edit "+p.pile_id;
  topLevelFields.forEach(k=>setVal(k,p[k]));
  buildRows(p.boring_logs);
  [...document.querySelectorAll("#reinforcementBody tr")].forEach((tr,i)=>{const r=p.reinforcement?.[i]||{};tr.querySelectorAll("input").forEach((x,j)=>x.value=r[reinFields[j]]??x.value)});
  editor.classList.remove("hidden");editor.scrollIntoView({behavior:"smooth"});
}
function collect(){
  const d={};
  topLevelFields.forEach(id=>{const el=$(id); if(el) d[id]=el.value});
  d.pile_id=$("pile_id").value;
  d.boring_logs=[...document.querySelectorAll("#boringBody tr")].map(tr=>{
    const o={};tr.querySelectorAll("input").forEach((x,i)=>o[boringFields[i]]=x.value);return o;
  }).filter(x=>Object.values(x).some(Boolean));
  d.reinforcement=[...document.querySelectorAll("#reinforcementBody tr")].map(tr=>{
    const o={};tr.querySelectorAll("input").forEach((x,i)=>o[reinFields[i]]=x.value);return o;
  }).filter(x=>Object.values(x).some(Boolean));
  return d;
}
form.addEventListener("submit",async e=>{
  e.preventDefault();message.textContent="Saving…";const d=collect();const old=$("editingId").value;
  const r=await fetch(old?"/api/piles/"+encodeURIComponent(old):"/api/piles",{method:old?"PUT":"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(d)});
  const j=await r.json(); if(!r.ok){message.textContent=j.error||"Save failed";message.className="message danger";return}
  message.className="message";message.textContent="Saved successfully.";await load();
});
rows.addEventListener("click",async e=>{
  const editBtn=e.target.closest("[data-edit]"), delBtn=e.target.closest("[data-del]");
  if(editBtn) edit(editBtn.dataset.edit);
  if(delBtn && confirm("Delete pile "+delBtn.dataset.del+"?")){await fetch("/api/piles/"+encodeURIComponent(delBtn.dataset.del),{method:"DELETE"});await load()}
});
$("newBtn").onclick=openNew;$("closeBtn").onclick=()=>editor.classList.add("hidden");$("cancelBtn").onclick=()=>editor.classList.add("hidden");
$("refreshBtn").onclick=load;$("search").oninput=render;
$("addBoringRow").onclick=()=>addBoringRow(null);
$("exportBtn").onclick=()=>{const ids=piles.map(p=>p.pile_id).join(",");if(ids)location.href="/api/export?ids="+encodeURIComponent(ids);else alert("No piles to export.")};
$("logoutBtn").onclick=async()=>{await fetch("/api/logout",{method:"POST"});location.href="/login.html"};
buildRows();load();
