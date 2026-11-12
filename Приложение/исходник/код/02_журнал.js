// ── helpers ───────────────────────────────────────────────
function esc(s){return (s==null?"":String(s)).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));}
function ru(d){ if(!d) return ""; const m=/^(\d{4})-(\d{2})-(\d{2})/.exec(d); return m?`${m[3]}.${m[2]}.${m[1]}`:d; }
function statusClass(s){ s=(s||"").toLowerCase();
  if(s.includes("зарегистр")) return "ok";
  if(s.includes("аннул")) return "bad";
  if(s.includes("работ")||s.includes("чернов")) return "work";
  return "new"; }
function methodsOf(z){ return Object.keys(z.методы||{}); }

// Шапка печатного документа: на экране не видна, в PDF/на бумаге —
// заголовок, что именно отобрано и когда сформировано. Без этого
// распечатка не документ, а безымянная таблица.
function printHead(title, подзаголовок){
  // Организация берётся из карточки лаборатории — там же, откуда её берёт
  // акт. Отдельное поле «наименование организации» жило в удалённом
  // черновике документа и разошлось бы с настройками.
  const орг=getЛаб().организация;
  return `<div class="print-only print-head">
    ${орг?`<div class="ph-sub" style="margin:0 0 6px">${esc(орг)}</div>`:""}
    <div class="ph-title">${esc(title)}</div>
    <div class="ph-sub">${подзаголовок?esc(подзаголовок)+" · ":""}сформировано ${ru(new Date().toISOString().slice(0,10))}</div>
  </div>`;
}
function printSign(){
  return `<div class="print-only print-sign">Составил: <span class="sl"></span> /
    <span class="sl"></span>/<div style="margin-top:10px;font-size:11px;color:#555">
    (подпись) &nbsp;&nbsp;&nbsp;&nbsp; (расшифровка)</div></div>`;
}

function setChrome(){
  // вход включили (обменом, копией), а никто не вошёл — спросить
  if(window.ЛНК_ГОТОВ && входВкл() && _я==null && !_окноВхода) показатьВход();
  document.getElementById("cJournal").textContent=liveJ().length;
  const пд=document.getElementById("cPodr"); if(пд) пд.textContent=live(ПОДР).length;
  const цНд=document.getElementById("cНд"); if(цНд) цНд.textContent=live(НД).length;
  const цПом=document.getElementById("cПом"); if(цПом) цПом.textContent=live(ПОМ).length;
  // Документы СМ — число действующих; предложения — новые (не разобранные).
  const цСмк=document.getElementById("cСмк"); if(цСмк) цСмк.textContent=live(СМК||[]).filter(д=>!смкУтратил(д)).length;
  const цПр=document.getElementById("cПредл"); if(цПр){ const н=предлНовых(); цПр.textContent=н; цПр.style.display=н?"":"none"; }
  const цПг=document.getElementById("cPg");
  if(цПг){ const м=пгМесяца(); цПг.textContent=м.осталось; цПг.title="работ по планам-графикам в этом месяце осталось"; }
  const цВиб=document.getElementById("cVibro");
  if(цВиб){ const к=ВИБК||[]; цВиб.textContent=live(к).length;
    цВиб.classList.toggle("warn", (ВИБИ||[]).length>0 && live(к).some(x=>состояниеКарты(x).ур==="прев")); }
  const цз=document.getElementById("cTasks");
  // (ЗАД||[]) — счётчики могут обновляться ещё до того, как реестр задач
  // загружен: сохранение из кода начальной загрузки (например, перевод
  // справочника на новые id) идёт раньше. Без защиты такой ранний вызов
  // ронял всю загрузку, и до пульта дело не доходило вовсе.
  if(цз){ const зд=ЗАД||[];
          const ну=наУтверждение().length;
          const n=зд.filter(з=>живаяЗадача(з)&&з.статус!=="закрыта").length+ну;
          цз.textContent=n; цз.classList.toggle("warn", ну>0 || зд.some(просроченаЗадача)); }
  const цзв=document.getElementById("cЗаяв");
  // try: счётчики могут обновляться до того, как объявлен реестр заявок
  if(цзв) try{ const откр=ЗАЯВ.filter(з=>заявкаОткрыта(з)).length; цзв.textContent=откр;
           цзв.classList.toggle("warn", заявокПросрочено()>0 || заявокНаУтверждение()>0); }catch(e){}
  const пр=document.getElementById("cCheck");
  if(пр){ const n=числоПроблем(); пр.textContent=n; пр.classList.toggle("warn", n>0); }
  document.getElementById("cObjects").textContent=live(O).length;
  document.getElementById("cSi").textContent=live(S).length;
  document.getElementById("cPeople").textContent=live(P).length;
  const cd=document.getElementById("cDocs");
  if(cd) cd.textContent=списокАктов().length;
  document.getElementById("fRows").textContent=liveJ().length;
  document.getElementById("fBuilt").textContent=ru(DB.собрано)||DB.собрано||"—";
  document.getElementById("verTag").textContent="v"+(DB.версия||"?")+" • реальные данные";
  обновитьКнопкуОтмены();
}

// ── Журнал: фильтры и сортировка ──────────────────────────
// Фильтры намеренно НЕ запоминаются между открытиями: иначе можно
// открыть журнал, увидеть 12 записей вместо 315 и решить, что данные
// пропали. Сортировка запоминается — она ничего не прячет.
let jf={год:"",метод:"",статус:"",исп:"",с:"",по:"",докум:""};
// Отбор «Документ» (с 3.74 — здесь, а не вкладками раздела «Акты и
// протоколы»): ждут номера — у метода записи нет номера протокола; без
// заключения — номера есть, заключение не вписано; готовы — и то и другое.
// Те же признаки, что раньше (ждётНомеров, оформлен).
const ЖУРНАЛ_ДОКУМ=[["без","ждут номера"],["нужно","без заключения"],["готовые","готовы"]];
function состояниеДокумента(z){ return ждётНомеров(z)?"без":оформлен(z)?"готовые":"нужно"; }
function журналОтбор(k,v){ switchView("journal"); resetJF(); setJF(k,v); }
let jsort=(function(){ try{ const r=ХР.getItem("lnk_jsort");
  if(r){ const o=JSON.parse(r); if(o&&o.c) return {c:o.c,d:o.d===1?1:-1}; } }catch(e){}
  return {c:"дата", d:-1}; })();   // по умолчанию — свежие сверху

function jsortSave(){ try{ ХР.setItem("lnk_jsort",JSON.stringify(jsort)); }catch(e){} }
function setJF(k,v){ if(jf[k]===v) return; jf[k]=v; renderJournal(); }
// Поле даты в фильтре, после которого раздел перерисовывается. Chromium
// шлёт change уже после первых цифр года (значение «0020-10-05»), а
// перерисовка заменяет само поле — фокус слетал, и год не удавалось
// допечатать (баг, найденный заказчиком). Поэтому применяю, только когда
// год полный (с 1900) или поле очищено; неполный — по уходу из поля.
function датаВвод(el, fn, ушёл){
  const v=el.value;
  if(!(!v || ушёл || +v.slice(0,4)>=1900)) return;
  const id=el.id; fn(v);
  // после перерисовки — фокус обратно в то же поле (по id)
  if(!ушёл && id){ const н=document.getElementById(id); if(н && н!==document.activeElement) н.focus(); }
}
function resetJF(){ jf={год:"",метод:"",статус:"",исп:"",с:"",по:"",докум:""}; query=""; renderJournal(); }
function sortJournal(c){
  if(jsort.c===c) jsort.d=-jsort.d; else { jsort.c=c; jsort.d=(c==="дата")?-1:1; }
  jsortSave(); renderJournal();
}
// Исполнители одной записи — без повторов, в порядке методов записи
// (обычно один человек делает несколько методов, повторять его незачем).
function исполнителиЗаписи(z){
  const н=[]; methodsOf(z).forEach(м=>{ const и=String(((z.методы||{})[м]||{}).исп||"").trim(); if(и && н.indexOf(и)<0) н.push(и); });
  return н;
}
// Исполнители, как они вписаны в журнал (по всем методам всех записей).
// Список берётся из самого журнала, а не из справочника: отбор должен
// находить и тех, кого в справочнике специалистов нет.
function исполнителиЖурнала(){
  const н=new Set();
  liveJ().forEach(z=>methodsOf(z).forEach(м=>{ const и=String(((z.методы||{})[м]||{}).исп||"").trim(); if(и) н.add(и); }));
  if(jf.исп) н.add(jf.исп);
  return [...н].sort((a,b)=>a.localeCompare(b,"ru"));
}
// Отобранные и отсортированные записи. Одна функция на список, экспорт
// и печать — чтобы в Excel уходило ровно то, что видно на экране.
function journalRows(){
  const q=(query||"").toLowerCase();
  const rows=liveJ().filter(z=>{
    if(jf.год && (z.дата||"").slice(0,4)!==jf.год) return false;
    if(jf.метод && !(z.методы&&z.методы[jf.метод])) return false;
    if(jf.статус && (z.статус||"")!==jf.статус) return false;
    if(jf.докум && (methodsOf(z).length?состояниеДокумента(z):"")!==jf.докум) return false;
    // Исполнитель — хотя бы по одному методу записи, так, как вписан в журнал.
    if(jf.исп && !methodsOf(z).some(м=>String(((z.методы||{})[м]||{}).исп||"").trim()===jf.исп)) return false;
    // Период — по дате контроля; запись без даты при заданном периоде не
    // показывается (её место — «Проверка данных»).
    if((jf.с||jf.по) && !z.дата) return false;
    if(jf.с && z.дата<jf.с) return false;
    if(jf.по && z.дата>jf.по) return false;
    if(q && !((z.объект||"")+(z.акт||"")+(z.наименование||"")+(z.зав||"")+(z.рег||"")+
              methodsOf(z).join(" ")+(z.дата||"")).toLowerCase().includes(q)) return false;
    return true;
  });
  const key=z=>{ switch(jsort.c){
    case "акт": return z.акт||"";
    case "объект": return z.объект||"";
    case "методы": return methodsOf(z).join(" ");
    case "исп": return исполнителиЗаписи(z).join(", ");
    case "статус": return z.статус||"";
    default: return z.дата||""; } };
  // Пустые значения всегда внизу — и по возрастанию, и по убыванию:
  // иначе сортировка по № акта поднимает наверх записи без номера.
  return rows.sort((a,b)=>{
    const A=String(key(a)||""), B=String(key(b)||"");
    if(!A && !B) return 0;
    if(!A) return 1;
    if(!B) return -1;
    return A.localeCompare(B,"ru",{numeric:true})*jsort.d;
  });
}

// ── Журнал ────────────────────────────────────────────────
function renderJournal(){
  const rows=journalRows();
  const годы=[...new Set(liveJ().map(z=>(z.дата||"").slice(0,4)).filter(Boolean))].sort().reverse();
  const статусы=[...new Set(liveJ().map(z=>z.статус).filter(Boolean))].sort();
  const методыЕсть=METHODS.filter(m=>liveJ().some(z=>z.методы&&z.методы[m]));
  const исполнители=исполнителиЖурнала();
  const естьФильтр=jf.год||jf.метод||jf.статус||jf.исп||jf.с||jf.по||jf.докум||query;
  const сМет=liveJ().filter(z=>methodsOf(z).length);
  const докСчёт=k=>сМет.filter(z=>состояниеДокумента(z)===k).length;
  const opt=(cur,arr,пусто)=>`<option value="">${пусто}</option>`+
    arr.map(v=>`<option value="${esc(v)}"${v===cur?" selected":""}>${esc(v)}</option>`).join("");
  const стрелка=c=>jsort.c===c?(jsort.d===1?" ▲":" ▼"):"";
  const th=(c,l)=>`<th class="sortable" onclick="sortJournal('${c}')">${l}${стрелка(c)}</th>`;
  // Колонка «№ п/п» существует всегда, но видна только в печати:
  // на экране она не нужна, а в распечатке без неё строки не сослаться.
  const body=rows.length?rows.map((z,i)=>`
    <tr onclick="showJournal('${esc(z.id)}')"${жОтм.has(String(z.id))?' class="отм"':""}>
      <td class="noprint жчб" onclick="event.stopPropagation()"><input type="checkbox" ${жОтм.has(String(z.id))?"checked":""}
        onchange="жОтметить('${esc(z.id)}',this.checked)" title="Отметить для массовой правки"></td>
      <td class="pnum num">${i+1}</td>
      <td class="num">${esc(z.акт)||"—"}</td>
      <td class="num">${ru(z.дата)}${записьЗакрыта(z)?` <span title="год закрыт — правка с подтверждением">🔒</span>`:""}</td>
      <td class="obj">${esc(z.объект)}${z.зав?`<div class="muted" style="font-size:11.5px">зав. № ${esc(z.зав)}${z.рег?` · рег. ${esc(z.рег)}`:""}</div>`:""}</td>
      <td>${methodsOf(z).map(m=>{const н=номерМетода(z,m);
        const бланк=естьФормаПротокола(m);
        return `<span class="pill${н?"":" nonum"}${бланк?" пбланк":""}" title="${бланк?"Открыть бланк документа · ":""}${н?"протокол "+esc(обозначениеПротокола(z,m)):"номер не выдан"}"${
          бланк?` onclick="event.stopPropagation();открытьПротокол('${esc(z.id)}','${m}')"`:""}>${m}${н?" "+esc(н):""}</span>`;
        }).join("")||"<span class='muted'>—</span>"}</td>
      <td class="жисп">${исполнителиЗаписи(z).map(и=>`<div>${esc(и)}</div>`).join("")||"<span class='muted'>—</span>"}</td>
      <td><span class="st ${statusClass(z.статус)}">${esc(z.статус)||"—"}</span>${ждётНомеров(z)?`<div class="muted" style="font-size:11px">ждёт номера</div>`
        :jf.докум&&methodsOf(z).length?`<div class="muted" style="font-size:11px">${состояниеДокумента(z)==="готовые"?"готов":"нет заключения"}</div>`:""}</td>
    </tr>`).join(""):`<tr><td colspan="8" class="empty">Ничего не найдено.</td></tr>`;
  const всеПоказаны=rows.length && rows.every(z=>жОтм.has(String(z.id)));
  // Что именно отобрано — попадёт в подзаголовок распечатки.
  const отбор=[jf.год?("год "+jf.год):"", jf.метод?("метод "+jf.метод):"",
    jf.статус?("статус: "+jf.статус):"", jf.докум?("документ: "+(ЖУРНАЛ_ДОКУМ.find(x=>x[0]===jf.докум)||["",""])[1]):"", jf.исп?("исполнитель: "+jf.исп):"",
    (jf.с||jf.по)?("дата контроля "+(jf.с?"с "+ru(jf.с):"")+(jf.с&&jf.по?" ":"")+(jf.по?"по "+ru(jf.по):"")):"",
    query?("поиск: «"+query+"»"):""].filter(Boolean).join(", ");
  document.getElementById("main").innerHTML=`
    ${printHead("Журнал учёта заключений ЛНК",
      (отбор?("отобрано: "+отбор+" · "):"")+"записей "+rows.length+
      (jf.год&&годЗакрыт(jf.год)?" · год закрыт "+ru(((закрытиеГода(jf.год)||{}).когда||"").slice(0,10)):""))}
    <div class="head">
      <h2>Журнал</h2><span class="sub" id="cnt">${rows.length} из ${liveJ().length}</span>
      <div class="search"><span class="ic">⌕</span><input id="q" placeholder="Поиск: объект, № акта, зав. №, метод…" value="${esc(query)}"></div>
      <div class="spring"></div>
      <button class="btn primary" onclick="editJournal(null)">＋ Добавить запись</button>
      <button class="btn" onclick="exportXlsx()">▤ Excel</button>
      <button class="btn" onclick="выгрузкаЖурналаДляПравки()" title="Выгрузить показанные записи в Excel, поправить и загрузить обратно тем же импортом">⤓ Выгрузить для правки</button>
      <button class="btn" onclick="document.getElementById('imp_journal').click()">⬆ Импорт из Excel</button>
      <input type="file" id="imp_journal" accept=".xlsx" style="display:none" onchange="импортЖурналВыбран(this)">
      <button class="btn" onclick="window.print()">⎙ PDF</button>
      <button class="btn" onclick="backupSave()" title="Резервная копия всей базы в файл">▼ Копия</button>
    </div>
    <div class="filters noprint">
      <label>Год <select onchange="setJF('год',this.value)">${opt(jf.год,годы,"все")}</select></label>
      ${блокЗакрытияГода(jf.год)}
      <label>Метод <select onchange="setJF('метод',this.value)">${opt(jf.метод,методыЕсть,"любой")}</select></label>
      <label>Статус <select onchange="setJF('статус',this.value)">${opt(jf.статус,статусы,"любой")}</select></label>
      <label title="По номерам протоколов и заключению записи">Документ <select onchange="setJF('докум',this.value)"><option value="">любой</option>
        ${ЖУРНАЛ_ДОКУМ.map(([k,t])=>`<option value="${k}"${jf.докум===k?" selected":""}>${t} (${докСчёт(k)})</option>`).join("")}</select></label>
      <label>Исполнитель <select onchange="setJF('исп',this.value)">${opt(jf.исп,исполнители,"любой")}</select></label>
      <label>Дата с <input id="jf_с" type="date" value="${esc(jf.с)}" onchange="датаВвод(this,v=>setJF('с',v))" onblur="датаВвод(this,v=>setJF('с',v),1)"></label>
      <label>по <input id="jf_по" type="date" value="${esc(jf.по)}" onchange="датаВвод(this,v=>setJF('по',v))" onblur="датаВвод(this,v=>setJF('по',v),1)"></label>
      ${jf.с&&jf.по&&jf.с>jf.по?`<span class="fnote" style="color:var(--danger)">«с» позже «по» — ничего не попадёт</span>`:""}
      ${естьФильтр?`<button class="btn" onclick="resetJF()">✕ Сбросить отбор</button>
        <span class="fnote">показано ${rows.length} из ${liveJ().length}</span>`:
        `<span class="fnote">заголовок таблицы — сортировка по клику</span>`}
    </div>
    ${блокМассовойПравки()}
    <div class="card"><div class="tblwrap"><table data-отбор="нет">
      <thead><tr><th class="noprint жчб"><input type="checkbox" ${всеПоказаны?"checked":""}
        onchange="жОтметитьПоказанные(this.checked)" title="Отметить все показанные записи"></th>
        <th class="pnum">№</th>${th("акт","№ акта")}${th("дата","Дата")}${th("объект","Объект контроля")}${th("методы","Методы")}${th("исп","Исполнители")}${th("статус","Статус")}</tr></thead>
      <tbody>${body}</tbody></table></div></div>
    ${printSign()}`;
  const qb=document.getElementById("q");
  qb.addEventListener("input",()=>{query=qb.value;const p=qb.selectionStart;renderJournal();
    const n=document.getElementById("q");n.focus();n.setSelectionRange(p,p);});
}

// ── Массовая правка журнала ───────────────────────────────
// В журнале одно и то же значение часто стоит сразу у многих записей:
// «Исходящий» — 18 разных значений на 225 записей (одно у 74), «№ акта» —
// одно у 36. Вписывать его в каждую запись по отдельности незачем:
// отмечаются записи флажками и поле ставится разом. Отметка «изменено»
// ставится только реально изменившимся.
// Очистка (версия 3.15, выбор заказчика «пусто = стереть»): стирает поле,
// которое ОЧИСТИЛИ — вписали и стёрли, или выбрали «— очистить —» в
// списке; оно подсвечивается и подписано «будет очищено». Поле, которого
// не касались, не меняется: иначе, проставляя один статус, стёрли бы у
// всех отмеченных записей № акта и исходящий. Номера протоколов здесь не ставятся: у каждой записи свой
// номер, его выдаёт отдельный механизм с правом выдачи.
const жОтм=new Set();
// Введённое в полосу правки живёт отдельно от разметки: полоса
// перерисовывается при каждом щелчке по флажку записи.
const жмЗнач={};   // k → {v, т} (т — поле трогали)
const ЖМ_ОЧИСТИТЬ="__очистить__";   // значение пункта «— очистить —» (статусов с таким именем нет)
function жмВвод(k,el){ жмЗнач[k]={v:el.value, т:1}; жмВид(el); }
function жмВид(el){
  const очистка=el.value===ЖМ_ОЧИСТИТЬ || (el.tagName==="INPUT" && !el.value.trim() && el.dataset.т==="1");
  el.classList.toggle("жмочист", очистка);
  if(el.tagName==="INPUT") el.placeholder=очистка?"будет очищено":"не менять";
}
function жмСброс(){ Object.keys(жмЗнач).forEach(k=>delete жмЗнач[k]); renderJournal(); }
// что делать с полем: undefined — не менять, "" — очистить, иначе значение
function жмДействие(k){
  const с=жмЗнач[k]; if(!с || !с.т) return undefined;
  if(с.v===ЖМ_ОЧИСТИТЬ) return "";
  const v=String(с.v||"").trim();
  return v;   // тронутое и пустое — очистить
}
const МАССОВЫЕ_ПОЛЯ=[
  {k:"статус",l:"Статус",список:()=>STATUSES},   // STATUSES объявлен ниже — только при вызове
  {k:"акт",l:"№ акта"},
  {k:"исходящий",l:"Исходящий"},
  {k:"архив",l:"В архив ЛНК"},
  {k:"отправитель",l:"Отправитель",спец:1}
];
function жОтметить(ид,да){ if(да) жОтм.add(String(ид)); else жОтм.delete(String(ид)); renderJournal(); }
function жОтметитьПоказанные(да){
  journalRows().forEach(z=>{ if(да) жОтм.add(String(z.id)); else жОтм.delete(String(z.id)); });
  renderJournal();
}
function жСнятьОтметки(){ жОтм.clear(); renderJournal(); }
function блокМассовойПравки(){
  const живые=new Set(liveJ().map(z=>String(z.id)));
  [...жОтм].forEach(ид=>{ if(!живые.has(ид)) жОтм.delete(ид); });   // удалённые тем временем
  if(!жОтм.size) return "";
  const видно=new Set(journalRows().map(z=>String(z.id)));
  const скрыто=[...жОтм].filter(ид=>!видно.has(ид)).length;
  const спецы=live(P).map(p=>p.вдок||p.фио).filter(Boolean);
  const поле=f=>{
    const с=жмЗнач[f.k]||{}, v=с.v||"";
    const оч=с.т && (v===ЖМ_ОЧИСТИТЬ || (!f.список && !v.trim()));
    return f.список
    ? `<label>${esc(f.l)}<select id="жм_${f.k}" class="${оч?"жмочист":""}" onchange="жмВвод('${f.k}',this)">
        <option value=""${!v?" selected":""}>— не менять —</option>
        <option value="${esc(ЖМ_ОЧИСТИТЬ)}"${v===ЖМ_ОЧИСТИТЬ?" selected":""}>— очистить —</option>
        ${f.список().map(x=>`<option${x===v?" selected":""}>${esc(x)}</option>`).join("")}</select></label>`
    : `<label>${esc(f.l)}<input id="жм_${f.k}" value="${esc(v)}" data-т="${с.т?1:0}" class="${оч?"жмочист":""}"
        placeholder="${оч?"будет очищено":"не менять"}" oninput="this.dataset.т='1';жмВвод('${f.k}',this)"${f.спец?' list="жмСпец"':""}></label>`;
  };
  return `<div class="card noprint" style="padding:12px 14px;border-color:var(--accent)">
    <datalist id="жмСпец">${спецы.map(x=>`<option value="${esc(x)}">`).join("")}</datalist>
    <div style="display:flex;gap:10px;align-items:flex-end;flex-wrap:wrap" class="form жмасс">
      <b style="align-self:center">Отмечено: ${жОтм.size}${скрыто?` <span class="muted" style="font-weight:normal">(из них скрыто отбором: ${скрыто})</span>`:""}</b>
      ${МАССОВЫЕ_ПОЛЯ.map(поле).join("")}
      <button class="btn primary" onclick="жПрименить()">Проставить отмеченным</button>
      <button class="btn" onclick="жмСброс()" title="Все поля снова «не менять»">Сбросить поля</button>
      <button class="btn" onclick="жСнятьОтметки()">Снять отметки</button>
    </div>
    <div class="note" style="margin:6px 0 0">Поле, которого не касались, у записей не меняется. Чтобы СТЕРЕТЬ
      значение у отмеченных записей — сотрите текст в поле (или выберите «— очистить —» в списке): поле подсветится
      и будет подписано «будет очищено». Номера протоколов здесь не ставятся — у каждой записи свой.</div>
  </div>`;
}
function жПрименить(){
  const знач={};
  МАССОВЫЕ_ПОЛЯ.forEach(f=>{ const д=жмДействие(f.k); if(д!==undefined) знач[f.k]=д; });
  const поля=Object.keys(знач);
  if(!поля.length){ toast("Заполните или очистите хотя бы одно поле"); return; }
  const записи=liveJ().filter(z=>жОтм.has(String(z.id)));
  const меняются=записи.filter(z=>поля.some(k=>String(z[k]==null?"":z[k]).trim()!==знач[k]));
  if(!меняются.length){ toast("У всех отмеченных записей эти значения уже стоят"); return; }
  if(!разрешитьЗакрытые(меняются, "Массовая правка")) return;
  const подп=k=>МАССОВЫЕ_ПОЛЯ.find(f=>f.k===k).l;
  const перезапишут=поля.filter(k=>знач[k]).map(k=>{
    const n=записи.filter(z=>String(z[k]==null?"":z[k]).trim() && String(z[k]).trim()!==знач[k]).length;
    return n?("«"+подп(k)+"» уже заполнено иначе у "+n):"";
  }).filter(Boolean);
  const стёрто=new Map();
  поля.filter(k=>!знач[k]).forEach(k=>{
    const n=записи.filter(z=>String(z[k]==null?"":z[k]).trim()).length;
    if(n) стёрто.set(подп(k),n);
  });
  if(!confirm("Отмечено записей: "+записи.length+".\n\n"+
    поля.map(k=>"· "+подп(k)+": "+(знач[k]?"«"+знач[k]+"»":"ОЧИСТИТЬ")).join("\n")+
    "\n\nИзменится записей: "+меняются.length+"."+
    (перезапишут.length?("\nСтарое значение будет заменено: "+перезапишут.join("; ")+"."):"")+
    (стёрто.size?("\n\n⚠ БУДЕТ СТЁРТО значений: "+[...стёрто.values()].reduce((a,b)=>a+b,0)+" — "+
      [...стёрто].map(([л,n])=>"«"+л+"»: "+n).join(", ")+"."):""))) return;
  отменаНачать("массовая правка журнала", [STORE_KEY]);
  const ts=new Date().toISOString(), место=getWS().код;
  меняются.forEach(z=>{ поля.forEach(k=>{ z[k]=знач[k]; }); z.изменено=ts; z.место=место; });
  saveJournal();
  отменаЗавершить();
  жОтм.clear(); Object.keys(жмЗнач).forEach(k=>delete жмЗнач[k]);
  renderJournal();
  toast("Изменено записей: "+меняются.length);
}
function showJournal(id){
  const z=J.find(x=>String(x.id)===String(id)); if(!z) return;
  const mrows=methodsOf(z).map(m=>{const об=обозначениеПротокола(z,m);
    const бланк=естьФормаПротокола(m);
    return `<tr${бланк?` style="cursor:pointer" title="Открыть бланк документа" onclick="closeDrawer();открытьПротокол('${esc(z.id)}','${m}')"`:""}><td>${m}</td><td>${об?(бланк?`<u>${esc(об)}</u>`:esc(об)):"<span class='muted'>номер не выдан</span>"}</td><td>${esc(z.методы[m].исп)||"—"}</td></tr>`;
    }).join("")||`<tr><td colspan="3" class="muted">методы не указаны</td></tr>`;
  const F=(k,v)=>v?`<div class="fld"><div class="k">${k}</div><div class="v">${esc(v)}</div></div>`:"";
  openDrawer("Акт № "+(z.акт||"—")+" от "+ru(z.дата), `
    <div class="dbtns">
      <button class="btn primary" onclick="closeDrawer();editJournal('${esc(z.id)}')">✎ Изменить</button>
      <button class="btn" onclick="closeDrawer();открытьДокумент('${esc(z.id)}')">▦ Документ</button>
      ${ждётНомеров(z)?`<button class="btn" onclick="выдатьНомера('${esc(z.id)}')">№ Выдать номера</button>`:""}
      ${String(z.акт||"").trim()?"":`<button class="btn" onclick="выдатьАкт('${esc(z.id)}')">№ Номер акта</button>`}
      <button class="btn danger" onclick="delJournal('${esc(z.id)}')">✕ Удалить</button>
    </div>
    ${F("Статус",z.статус)}${F("Объект",z.объект)}${F("Наименование в документе",z.наименование)}
    ${F("Заводской №",z.зав)}${F("Рег. / уч. №",z.рег)}${F("Подразделение",z.подразделение)}
    <div class="sect">Методы контроля</div>
    <table class="mtab"><thead><tr><th>Метод</th><th>№</th><th>Исполнитель</th></tr></thead><tbody>${mrows}</tbody></table>
    <div class="sect">Движение экземпляров</div>
    ${F("В архив ЛНК",z.архив)}${F("Исходящий",z.исходящий)}${F("Отправитель",z.отправитель)}
    ${F("Примечание",z.примечание)}`);
}



// ── Закрытые годы журнала (3.58) ──────────────────────────
// Указание заказчика: «закрыть год — запрет правки записей с
// подтверждением». Закрытый год — запись этого года форма открывает только
// для просмотра; править можно, но лишь после явного «да» (кнопка
// «Разрешить правку»). Те же вопросы — у всех путей, которыми запись
// меняется пачкой или из другого места программы: массовая правка, импорт
// из Excel, «Повторить по объектам», выдача номеров, удаление, возврат
// удалённой или отброшенной версии. Отмена последней загрузки и приём
// файла обмена не спрашивают: первая возвращает то, что сделали здесь же,
// а у второго свой разбор (правки, сделанные на другом месте после
// закрытия, видны в «Проверке данных»).
// Реестр — по записи на год, id «ЗГ-ГГГГ» (одинаковый на всех местах),
// закрытие и открытие — поле «закрыт» со своей отметкой времени: в обмене
// побеждает свежее, без истории правок (это не карточка).
function сохранитьЗакрытые(){
  try{ ХР.setItem(REG_ЗАКР, JSON.stringify({версия:DB.версия, сохранено:new Date().toISOString(), rows:ЗАКР})); }
  catch(e){ переполнение(e); }
}
function слитьЗакрытые(вход){
  let изм=0; const поId={}; ЗАКР.forEach((r,i)=>поId[r.id]=i);
  (вход||[]).forEach(r=>{ if(!r || r.id==null || r.id==="") return;
    const i=поId[r.id];
    if(i==null){ ЗАКР.push(r); поId[r.id]=ЗАКР.length-1; изм++; }
    else if(lastTs(r)>lastTs(ЗАКР[i])){ ЗАКР[i]=r; изм++; } });
  if(изм) сохранитьЗакрытые();
  return изм;
}
function закрытиеГода(год){ return год ? ЗАКР.find(r=>String(r.год)===String(год)) : null; }
function годЗакрыт(год){ const r=закрытиеГода(год); return !!(r && r.закрыт); }
function записьЗакрыта(z){ return !!z && годЗакрыт(годЗаписи(z)); }
function закрытыеГоды(){ return ЗАКР.filter(r=>r.закрыт).map(r=>String(r.год)).sort(); }
// id записей, правку которых в этом сеансе уже подтвердили — чтобы форма
// не спрашивала второй раз при сохранении. Не хранится: после перезагрузки
// запись снова только для просмотра.
const _закрРазрешено=new Set();
// Затрагивает ли действие записи закрытых лет; если да — спросить.
// записи — сами записи или {id?, дата} (новая запись, перенос даты).
function разрешитьЗакрытые(записи, что){
  const видел=new Set();
  const закр=(записи||[]).filter(z=>{
    if(!z || !годЗакрыт(годЗаписи(z))) return false;
    const ид=z.id!=null && z.id!==""?String(z.id):"";
    if(ид && _закрРазрешено.has(ид)) return false;
    if(ид){ if(видел.has(ид)) return false; видел.add(ид); }
    return true;
  });
  if(!закр.length) return true;
  const годы=[...new Set(закр.map(годЗаписи))].sort();
  // Править записи закрытого года — тоже только место начальника.
  if(!яНачальник()){
    alert("🔒 "+(годы.length>1?"Годы "+годы.join(", ")+" закрыты":"Год "+годы[0]+" закрыт")+".\n\n"+что+
      " не выполнено: записи закрытого года правит только место начальника.");
    return false;
  }
  const ок=confirm("🔒 "+(годы.length>1?"Годы "+годы.join(", ")+" закрыты":"Год "+годы[0]+" закрыт")+
    " — записи закрытого года не правятся.\n\n"+что+": затрагивает "+
    закр.length+" "+числослово(закр.length,["запись","записи","записей"])+" закрытого года.\n\n"+
    "Всё равно изменить?");
  if(ок) закр.forEach(z=>{ if(z.id!=null && z.id!=="") _закрРазрешено.add(String(z.id)); });
  return ок;
}
// Закрывает и открывает год и разрешает правку записей закрытого года
// только место начальника (указания заказчика) — тот же признак, что у
// постановки задач.
function закрытьГод(год){
  год=String(год||""); if(!год) return;
  if(!яНачальник()){ toast("Закрыть год может только место начальника"); return; }
  const n=liveJ().filter(z=>годЗаписи(z)===год).length;
  const идёт=+год>=new Date().getFullYear();
  if(!confirm("Закрыть "+год+" год?\n\nЗаписей журнала за год: "+n+". Они станут только для "+
    "просмотра: правка, удаление, массовая правка, импорт из Excel и выдача номеров по ним — "+
    "только после подтверждения.\n\nЗакрытие уйдёт на другие места при обмене. Открыть год "+
    "можно здесь же."+(идёт?"\n\n⚠ "+год+" год ещё не закончился.":""))) return;
  поставитьЗакрытие(год, true);
  renderJournal(); toast("Год "+год+" закрыт");
}
function открытьГод(год){
  год=String(год||""); if(!год) return;
  if(!яНачальник()){ toast("Открыть год может только место начальника"); return; }
  if(!confirm("Открыть "+год+" год?\n\nЗаписи года снова можно будет править без вопросов. "+
    "Открытие уйдёт на другие места при обмене.")) return;
  поставитьЗакрытие(год, false);
  renderJournal(); toast("Год "+год+" открыт");
}
function поставитьЗакрытие(год, да){
  const ts=new Date().toISOString(), место=getWS().код;
  const нов={id:"ЗГ-"+год, год:год, закрыт:!!да, когда:ts, кто:место, изменено:ts, место:место};
  const i=ЗАКР.findIndex(r=>String(r.год)===год);
  if(i>=0) ЗАКР[i]=нов; else ЗАКР.push(нов);
  сохранитьЗакрытые();
}
// Кнопка и пометка у фильтра «Год» журнала.
function блокЗакрытияГода(год){
  if(!год) return закрытыеГоды().length?`<span class="fnote">🔒 закрыты: ${esc(закрытыеГоды().join(", "))}</span>`:"";
  const r=закрытиеГода(год);
  if(r && r.закрыт) return `<span class="st закрыт" title="закрыл ${esc(r.кто||"")} ${esc(ru((r.когда||"").slice(0,10)))}">🔒 ${esc(год)} год закрыт ${esc(ru((r.когда||"").slice(0,10)))}${r.кто?" · "+esc(r.кто):""}</span>
    ${яНачальник()?`<button class="btn" onclick="открытьГод('${esc(год)}')">🔓 Открыть год</button>`:""}`;
  if(!яНачальник()) return "";
  return `<button class="btn" onclick="закрытьГод('${esc(год)}')" title="Записи года станут только для просмотра; правка — с подтверждением">🔒 Закрыть ${esc(год)} год</button>`;
}
// Шапка формы записи закрытого года.
function плашкаЗакрытойЗаписи(z){
  const г=годЗаписи(z), r=закрытиеГода(г)||{};
  return `<div class="закрплашка">🔒 ${esc(г)} год закрыт${r.когда?" "+esc(ru(r.когда.slice(0,10))):""}${r.кто?" (место "+esc(r.кто)+")":""} —
    запись только для просмотра.
    ${яНачальник()?`<button class="btn small" type="button" onclick="разрешитьПравкуЗаписи('${esc(z.id)}')">Разрешить правку</button>`:
      `<span style="font-weight:400">Править может только место начальника.</span>`}</div>`;
}
function разрешитьПравкуЗаписи(ид){
  const z=J.find(x=>String(x.id)===String(ид)); if(!z) return;
  if(!разрешитьЗакрытые([z], "Правка записи")) return;
  editJournal(ид);
}
