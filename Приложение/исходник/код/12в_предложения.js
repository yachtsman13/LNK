// ── Предложения по улучшению программы (3.85) ───────────────
// Просьба заказчика: «раздел для внесения предложений по улучшению
// программы». Пишет любой; предложения ходят в обмене (видны на всех
// местах) и выгружаются в Excel — этот файл и пересылается разработчику.
// Ответ и отметка «сделано в версии …» вписываются в ту же карточку.
const REG_ПРЕДЛ="lnk_ideas";
let ПРЕДЛ=[];
const ПРЕДЛ_СОСТОЯНИЯ=["Новое","Принято","Сделано","Отклонено"];
// Разделы программы — как в меню.
function разделыПрограммы(){
  const из=[...document.querySelectorAll("#side .nav")].map(b=>[...b.childNodes].filter(n=>n.nodeType===3)
    .map(n=>n.textContent).join("").trim()).filter(Boolean);
  return ["Вся программа"].concat(из);
}
REGS.idea={ arr:()=>ПРЕДЛ, key:REG_ПРЕДЛ, view:"ideas", pfx:"ПР", title:"предложение", name:"текст",
  fields:[
    {k:"текст",l:"Что предлагаете",wide:1,type:"area",
     подсказка:"Что неудобно или чего не хватает, где именно, как хотелось бы. Можно приложить снимок экрана файлом."},
    {k:"раздел",l:"Раздел программы",type:"select",get options(){ return разделыПрограммы(); }},
    {k:"важность",l:"Насколько важно",type:"select",options:["Мешает работать","Важно","Было бы удобно"]},
    {k:"автор",l:"Кто предложил"},
    {k:"дата",l:"Дата",type:"date"},
    {k:"файл",l:"Файл (снимок экрана, пример документа)",wide:1,файлДок:1},
    {k:"состояние",l:"Состояние",type:"select",options:ПРЕДЛ_СОСТОЯНИЯ},
    {k:"ответ",l:"Ответ / в какой версии сделано",wide:1,type:"area"}],
  поУмолчанию:()=>({дата:сегодняМестн(), состояние:"Новое", автор:getWS().код}),
  после:()=>switchView("ideas") };
ДОП_РЕЕСТРЫ.push({поле:"предложения", ключ:REG_ПРЕДЛ, н:"Предложения по программе", вид:"ideas", kind:"idea",
  arr:()=>ПРЕДЛ, set:v=>{ ПРЕДЛ=v; }, п:x=>String(x.текст||"").slice(0,80),
  поиск:{поля:x=>[x.текст,x.раздел,x.автор,x.ответ], п:x=>[x.дата?ru(x.дата):"", String(x.текст||"").slice(0,80)].filter(Boolean).join(" · ")}});

let предлОтбор="открытые";
function рядыПредложений(){
  const q=(query||"").toLowerCase();
  return live(ПРЕДЛ).filter(x=>{
    const с=String(x.состояние||"Новое");
    if(предлОтбор==="открытые" && (с==="Сделано"||с==="Отклонено")) return false;
    return !q || [x.текст,x.раздел,x.автор,x.ответ].join(" ").toLowerCase().includes(q);
  }).sort((a,b)=>String(b.дата||"").localeCompare(String(a.дата||"")));
}
function предлНовых(){ return live(ПРЕДЛ).filter(x=>String(x.состояние||"Новое")==="Новое").length; }
function setПредлОтбор(v){ предлОтбор=v; renderПредложения(); }
function renderПредложения(){
  const rows=рядыПредложений();
  const кл={"Новое":"new","Принято":"work","Сделано":"ok","Отклонено":""};
  const body=rows.length?rows.map(x=>{ const с=String(x.состояние||"Новое");
    return `<tr data-id="${esc(x.id)}" onclick="editReg('idea','${esc(x.id)}')">
      <td class="num">${ru(x.дата)||"—"}</td>
      <td class="obj" style="white-space:pre-line">${esc(x.текст||"")}${x.ответ?`<div class="muted" style="font-size:12px;margin-top:4px">Ответ: ${esc(x.ответ)}</div>`:""}</td>
      <td>${esc(x.раздел||"")}</td><td>${esc(x.важность||"")}</td><td>${esc(x.автор||"")}</td>
      <td><span class="st ${кл[с]||""}">${esc(с)}</span></td></tr>`; }).join("")
    :`<tr><td colspan="6" class="empty">${live(ПРЕДЛ).length?"Ничего не найдено.":"Предложений пока нет — «＋ Предложение»."}</td></tr>`;
  document.getElementById("main").innerHTML=`
    ${printHead("Предложения по улучшению программы", "предложений "+rows.length)}
    <div class="head"><h2>Предложения по программе</h2><span class="sub">${rows.length} из ${live(ПРЕДЛ).length}</span>
      <div class="seg">${[["открытые","Открытые"],["все","Все"]].map(([к,т])=>
        `<button class="btn small${предлОтбор===к?" on":""}" onclick="setПредлОтбор('${к}')">${т}</button>`).join("")}</div>
      <div class="search"><span class="ic">⌕</span><input id="q" placeholder="Поиск…" value="${esc(query)}"></div>
      <div class="spring"></div>
      <button class="btn primary" onclick="editReg('idea',null)">＋ Предложение</button>
      <button class="btn" onclick="exportXlsx()" title="Этот файл и пересылается разработчику">▤ Excel для разработчика</button>
      <button class="btn" onclick="window.print()">⎙ Печать / PDF</button></div>
    <div class="card"><div class="tblwrap"><table data-список="idea">
      <thead><tr><th>Дата</th><th>Предложение</th><th>Раздел</th><th>Важность</th><th>Кто</th><th>Состояние</th></tr></thead>
      <tbody>${body}</tbody></table></div></div>
    <div class="note справ">Предложения видны на всех местах (идут в обмене). Чтобы передать разработчику — «▤ Excel для разработчика»
      и отправить файл; ответ и версию, в которой сделано, вписывайте в само предложение.</div>`;
  wireSearch(renderПредложения); setChrome();
}
