// ── Защита от потери незаполненной формы ──────────────────
// Журнальная запись, карточка справочника, задача, карточка лаборатории
// и методика сохраняются только по кнопке. Уйти из них можно тремя
// способами — закрыть вкладку, нажать F5 и просто выбрать другой раздел
// в меню, — и во всех трёх введённое пропадало молча.
// Признак «в форме есть несохранённое» ставится по первому же вводу в
// поле. Клики по кнопкам не считаются: они либо сохраняют сами, либо
// перерисовывают свой блок, не теряя полей. Выбор файла в «Обмене» —
// тоже change, но импорт там происходит сразу, поэтому раздел в список
// не входит: иначе после каждого импорта выскакивал бы ложный вопрос.
const ВИДЫ_С_ФОРМОЙ=["edit","regedit","taskedit","lab","metod","planedit","partsedit"];
let _несохранено=false;
function формаСохранена(){ _несохранено=false; }
function уходСФормы(){
  if(!_несохранено) return true;
  if(!confirm("В форме есть несохранённые изменения.\n\n"+
              "Уйти и потерять их? Чтобы сохранить — останьтесь и нажмите «Сохранить».")) return false;
  _несохранено=false; return true;
}
["input","change"].forEach(тип=>
  document.getElementById("main").addEventListener(тип, ()=>{
    if(ВИДЫ_С_ФОРМОЙ.indexOf(curView)>=0) _несохранено=true;
  }));
// Закрытие вкладки и перезагрузка: текст предупреждения задаёт сам
// браузер, своё показать нельзя — достаточно того, что он спросит.
window.addEventListener("beforeunload", e=>{
  if(!_несохранено) return;
  e.preventDefault(); e.returnValue="";
});
// ── Единый поиск по всей базе ─────────────────────────────────────────
// Поиск в программе был в каждом разделе свой (одна переменная query на
// всех, но искала она только внутри открытого списка). Чтобы найти, где
// встречается заводской номер или фамилия, приходилось обходить восемь
// разделов подряд. Здесь одно поле на всё: ищет по журналу, объектам,
// приборам, специалистам, подразделениям, норм. документам, помещениям и
// задачам; строка результата открывает саму карточку.
// Удалённые записи не показываются (live) — их место в «Что изменилось».
const ПОИСК_РАЗДЕЛЫ=[
  {н:"Журнал", arr:()=>live(J), вид:"journal",
   поля:z=>[z.объект,z.наименование,z.зав,z.рег,z.акт,z.заключение,z.примечание,z.подразделение,
            ...Object.keys(z.методы||{}).map(м=>м+" "+((z.методы[м]||{}).н||"")+" "+((z.методы[м]||{}).исп||""))],
   п:z=>[z.дата?ru(z.дата):"", z.объект||z.наименование||"", z.акт?("акт "+z.акт):""].filter(Boolean).join(" · "),
   открыть:z=>"editJournal('"+esc(String(z.id))+"')"},
  {н:"Объекты", arr:()=>live(O), вид:"objects", kind:"obj",
   поля:x=>[x.наименование,x.зав,x.рег,x.инв,x.учет,расположениеОб(x),x.тип,x.изготовитель,x.примечание],
   п:x=>строкаОбъектаСпр(x)},
  {н:"Приборы (СИ)", arr:()=>live(S), вид:"si", kind:"si",
   поля:x=>[x.наименование,x.тип,x.зав,x.инв,x.карточка,x.свидетельство,x.изготовитель,
            x.характеристики,x.назначение,x.комплектность],
   п:x=>[x.наименование, x.тип?("«"+x.тип+"»"):"", x.зав?("зав. № "+x.зав):""].filter(Boolean).join(" ")},
  {н:"Специалисты", arr:()=>live(P), вид:"people", kind:"ppl",
   поля:x=>[x.фио,x.вдок,x.должность,x.функции,
            ...(x.аттестации||[]).map(а=>[а.метод,а.уровень,а.номер].filter(Boolean).join(" "))],
   п:x=>[x.вдок||x.фио, x.должность].filter(Boolean).join(" · ")},
  {н:"Подразделения", arr:()=>live(ПОДР), вид:"podr", kind:"podr",
   поля:x=>[x.наименование,x.предложная,x.эксп_организация,x.эксп_фио,
            ...контрактыКарточки(x).map(к=>[к.год,к.номер].filter(Boolean).join(" "))],
   п:x=>x.наименование||""},
  {н:"Норм. документы", arr:()=>live(НД), вид:"nd", kind:"nd",
   поля:x=>[x.наименование,x.обозначение,x.издатель,x.место_год,x.метод,x.категория],
   п:x=>[x.обозначение, x.наименование].filter(Boolean).join(" — ")},
  {н:"Помещения", arr:()=>live(ПОМ), вид:"premises", kind:"pom",
   поля:x=>[x.идентификация,x.назначение,x.площадь,x.параметры,x.сети,x.основание],
   п:x=>[x.идентификация, x.назначение].filter(Boolean).join(" · ")},
  {н:"Задачи", arr:()=>live(ЗАД), вид:"tasks",
   поля:x=>[x.что,x.кому,x.объект,x.заметки&&x.заметки.map?x.заметки.map(з=>з.текст).join(" "):"",x.источник],
   п:x=>[x.срок?ru(x.срок):"", x.что, x.кому].filter(Boolean).join(" · ")}
];
let _поиск="";
function найтиПоВсем(q){
  const н=String(q||"").trim().toLowerCase();
  if(н.length<2) return [];
  const из=[];
  ПОИСК_РАЗДЕЛЫ.forEach(р=>{
    const найдено=(р.arr()||[]).filter(x=>
      (р.поля(x)||[]).some(v=>String(v==null?"":v).toLowerCase().includes(н)));
    if(найдено.length) из.push({раздел:р, записи:найдено});
  });
  return из;
}
function поискВвод(v){ _поиск=v; renderНайти(true); }
// Недавние запросы. Поиск идёт на каждую букву, поэтому запоминать каждое
// состояние поля нельзя — в списке оказались бы «12», «129», «1297»…
// Запрос запоминается, когда он себя оправдал: по нему открыли найденное
// или нажали Enter. Хранится на этом месте (lnk_find_recent), в обмен и
// копию не идёт — это привычки места, не данные. Последние 10, без
// повторов (без учёта регистра), свежий — первым.
const ПОИСК_НЕДАВНИХ=10, REG_ПОИСК_НЕДАВНИЕ="lnk_find_recent";
function недавниеПоиски(){
  try{ const a=JSON.parse(ХР.getItem(REG_ПОИСК_НЕДАВНИЕ)||"[]"); return Array.isArray(a)?a.filter(x=>typeof x==="string"):[]; }
  catch(e){ return []; }
}
function запомнитьПоиск(q){
  q=String(q||"").trim(); if(q.length<2) return;
  const н=q.toLowerCase();
  const сп=[q].concat(недавниеПоиски().filter(x=>x.toLowerCase()!==н)).slice(0,ПОИСК_НЕДАВНИХ);
  try{ ХР.setItem(REG_ПОИСК_НЕДАВНИЕ, JSON.stringify(сп)); }catch(e){}
}
function повторитьПоиск(i){
  const q=недавниеПоиски()[i]; if(q==null) return;
  _поиск=q; renderНайти();
}
function забытьПоиски(){
  try{ ХР.removeItem(REG_ПОИСК_НЕДАВНИЕ); }catch(e){}
  renderНайти();
}
function блокНедавнихПоисков(){
  const сп=недавниеПоиски(); if(!сп.length) return "";
  return `<div class="недавние"><span class="muted" style="font-size:12px">Недавние:</span>
    ${сп.map((q,i)=>`<button class="btn small" onclick="повторитьПоиск(${i})" title="Повторить поиск">${esc(q)}</button>`).join("")}
    <button class="btn small" style="opacity:.7" onclick="забытьПоиски()" title="Очистить список недавних запросов">✕ очистить</button></div>`;
}
function открытьИзПоиска(вид, kind, ид){
  if(kind) editReg(kind, ид); else if(вид==="journal") editJournal(ид); else switchView(вид);
}

// ── Календарь сроков ──────────────────────────────────────
// Все сроки на одном экране по месяцам: свои задачи, работы ППР,
// окончание поверки, аттестации по методам и сроки специалистов,
// сроки лаборатории. Своего списка нет — это тот же расчёт, что у
// производных задач (производныеЗадачи), плюс свои задачи и окончание
// аттестаций по методам (их в задачах нет: там это «Прогноз аттестации»).
// Щелчок по дню — список дел этого дня, по делу — его карточка.
let кал={база:"", день:"", источник:""};
function календарьДела(база, вид){
  вид=вид||"месяц";
  const [с,по]=границыПериода(вид, база);
  const из=производныеЗадачи(вид, база).map(x=>Object.assign({}, x));
  ЗАД.filter(з=>живаяЗадача(з) && з.срок && з.срок>=с && з.срок<=по).forEach(з=>из.push({
    ид:"ЗД|"+з.id, источник:"Задачи", что:з.что+(з.исполнитель?(" — "+з.исполнитель):""),
    срок:з.срок, сделано:з.статус==="выполнена"||з.статус==="закрыта", куда:"tasks"}));
  live(P).forEach(c=>аттестации(c).forEach(а=>{
    if(!а.до || а.до<с || а.до>по) return;
    из.push({ид:"АТТ|"+c.id+"|"+а.метод+"|"+а.до, источник:"Аттестация",
      что:"Истекает аттестация "+[а.метод,а.уровень].filter(Boolean).join(" ")+": "+(c.вдок||c.фио),
      срок:а.до, сделано:false, куда:"people"});
  }));
  return из.sort((a,b)=>String(a.срок).localeCompare(String(b.срок))||String(a.что).localeCompare(String(b.что),"ru"));
}
function открытьДелоКалендаря(ид){
  const ч=String(ид).split("|");
  if(!уходСФормы()) return;
  if(ч[0]==="ЗД"){ editTask(ч.slice(1).join("|")); return; }
  if(ч[0]==="СИ" && найтиПоИд(live(S),ч[1])){ editReg("si",ч[1]); return; }
  if(["ИНСТР","ОГРПР","ЭЛБЕЗ","ПОЖБЕЗ","ПБ","КЛАСС","АТТ"].includes(ч[0]) && найтиПоИд(live(P),ч[1])){ editReg("ppl",ч[1]); return; }
  if(ч[0]==="ППР"){ const п=найтиПоИд(live(ППР_ПОЗ),ч[1]); if(п) план.год=+ч[2]; switchView("plan"); return; }
  if(ч[0]==="ЛАБ"){ switchView("lab"); return; }
  toast("Карточка не найдена");
}
function калСдвиг(шаг){
  const d=new Date((кал.база||new Date().toISOString().slice(0,8)+"01")+"T00:00:00");
  d.setDate(1); d.setMonth(d.getMonth()+шаг);
  кал.база=new Date(Date.UTC(d.getFullYear(),d.getMonth(),1)).toISOString().slice(0,10); кал.день="";
  renderКалендарь();
}
function калСегодня(){ кал.база=""; кал.день=""; renderКалендарь(); }
function калДень(д){ кал.день=кал.день===д?"":д; renderКалендарь(); }
function калИсточник(v){ кал.источник=v; renderКалендарь(); }
// Дела календаря списком в Excel — за месяц, квартал, полугодие или год
// того месяца, что открыт в календаре. Источник берётся тот же, что
// выбран на экране. Состояние — на день выгрузки: сделано / просрочено /
// предстоит. Своего расчёта нет: те же календарьДела(), что и на экране.
function exportКалендарь(вид){
  const сегодня=new Date().toISOString().slice(0,10);
  const база=кал.база||сегодня.slice(0,8)+"01";
  let дела=календарьДела(база, вид);
  if(кал.источник) дела=дела.filter(x=>x.источник===кал.источник);
  const имя=имяПериода(вид, база);
  if(!дела.length){ toast("Дел нет ("+имя+") — выгружать нечего"); return; }
  const сост=x=>x.сделано?"сделано":(x.срок<сегодня?"просрочено":"предстоит");
  xlsxФайл("Календарь сроков "+имя+(кал.источник?" — "+кал.источник:"")+".xlsx",[{имя:"Календарь",
    колонки:[{з:"№",ш:6,тип:"число"},{з:"Срок",ш:12,тип:"дата"},{з:"Что",ш:70},{з:"Источник",ш:16},
      {з:"Состояние на "+ru(сегодня),ш:18},{з:"Выполнено",ш:12,тип:"дата"}],
    строки:дела.map((x,i)=>[i+1, x.срок, x.что, x.источник, сост(x), x.сделано&&x.когда?x.когда:""])}]);
  toast("Выгружено дел: "+дела.length+" ("+имя+")");
}
function renderКалендарь(){
  const сегодня=new Date().toISOString().slice(0,10);
  const база=кал.база||сегодня.slice(0,8)+"01";
  const [с,по]=границыПериода("месяц", база);
  const всеДела=календарьДела(база);
  const источники=[...new Set(всеДела.map(x=>x.источник))].sort((a,b)=>a.localeCompare(b,"ru"));
  const дела=кал.источник?всеДела.filter(x=>x.источник===кал.источник):всеДела;
  const поДню={}; дела.forEach(x=>(поДню[x.срок]=поДню[x.срок]||[]).push(x));
  const кл=x=>x.сделано?"сдел":(x.срок<сегодня?"прос":"");
  const d0=new Date(с+"T00:00:00"); const сдвиг=(d0.getDay()+6)%7;
  const клетки=[];
  for(let i=0;i<42;i++){
    const d=new Date(d0); d.setDate(1-сдвиг+i);
    const iso=new Date(Date.UTC(d.getFullYear(),d.getMonth(),d.getDate())).toISOString().slice(0,10);
    if(i>=35 && iso>по) break;
    const сп=поДню[iso]||[];
    клетки.push(`<div class="дн${iso<с||iso>по?" чуж":""}${iso===сегодня?" сег":""}${iso===кал.день?" выб":""}"
      onclick="калДень('${iso}')" title="${esc(сп.map(x=>x.что).join("\n"))}">
      <div class="чис"><span>${d.getDate()}</span>${сп.length?`<span class="muted">${сп.length}</span>`:""}</div>
      ${сп.slice(0,3).map(x=>`<div class="сб ${кл(x)}">${esc(x.что)}</div>`).join("")}
      ${сп.length>3?`<div class="muted" style="font-size:11px">ещё ${сп.length-3}</div>`:""}
    </div>`);
  }
  const список=кал.день?(поДню[кал.день]||[]):дела;
  const просрочено=дела.filter(x=>!x.сделано && x.срок<сегодня).length;
  const месяц=МЕСЯЦЫ[+с.slice(5,7)-1]+" "+с.slice(0,4);
  document.getElementById("main").innerHTML=`
    <div class="head"><h2>Календарь сроков</h2>
      <span class="sub">${esc(месяц)} · дел: <b>${дела.length}</b>${просрочено?` · просрочено: <b style="color:var(--danger)">${просрочено}</b>`:""}</span>
      <div class="seg" style="margin-left:auto">
        <button class="btn" onclick="калСдвиг(-1)">‹</button>
        <button class="btn" onclick="калСегодня()">Сегодня</button>
        <button class="btn" onclick="калСдвиг(1)">›</button>
      </div>
      <select onchange="калИсточник(this.value)"><option value="">Все источники</option>
        ${источники.map(и=>`<option${и===кал.источник?" selected":""}>${esc(и)}</option>`).join("")}</select>
      <span style="display:flex;gap:4px;align-items:center" title="Дела периода списком в Excel — с учётом выбранного источника">
        <select id="калВыгрПериод">${[["месяц","месяц"],["квартал","квартал"],["полугодие","полугодие"],["год","год"]]
          .map(([v,l])=>`<option value="${v}">${l}</option>`).join("")}</select>
        <button class="btn" onclick="exportКалендарь(document.getElementById('калВыгрПериод').value)">▤ Excel</button></span>
    </div>
    <div class="card" style="padding:14px">
      <div class="кал">${["пн","вт","ср","чт","пт","сб","вс"].map(x=>`<div class="шап">${x}</div>`).join("")}${клетки.join("")}</div>
      <div class="note" style="margin:8px 0 0">Красное — срок прошёл, а дело не сделано; бледное — сделано.
        У работ ППР срок — последний день месяца (точной даты в плане нет). Отмечаются дела там, где им место:
        работа ППР — в плане, поверка — в карточке прибора, задача — в самой задаче.</div>
    </div>
    <div class="card" style="padding:14px 18px">
      <div class="grp" style="padding-left:0">${кал.день?("Дела на "+ru(кал.день)):("Все дела за "+esc(месяц.toLowerCase()))} — ${список.length}</div>
      <div class="plist">${список.length?список.map(x=>`<div class="pr" onclick="открытьДелоКалендаря('${esc(x.ид)}')">
          <span style="min-width:84px">${ru(x.срок)}</span>
          <span class="st ${x.сделано?"ok":(x.срок<сегодня?"bad":"new")}">${x.сделано?"сделано":(x.срок<сегодня?"просрочено":esc(x.источник))}</span>
          <span style="flex:1">${esc(x.что)}</span><span class="muted">${esc(x.источник)}</span></div>`).join("")
        :`<div class="note">Дел нет.</div>`}</div>
    </div>`;
}
function renderНайти(толькоРезультат){
  curView="find";
  document.querySelectorAll(".nav").forEach(n=>n.classList.toggle("on",n.dataset.view==="find"));
  const гр=найтиПоВсем(_поиск);
  const всего=гр.reduce((n,г)=>n+г.записи.length,0);
  const тело=(_поиск.trim().length<2)
    ? `${блокНедавнихПоисков()}<div class="note">Введите хотя бы два знака. Ищу по журналу, объектам, приборам,
        специалистам, подразделениям, норм. документам, помещениям и задачам — сразу по всем.
        Запрос попадает в недавние, когда по нему открыли найденное или нажали Enter.</div>`
    : (всего
      ? гр.map(г=>`<div style="margin-top:10px">
          <div class="grp" style="padding-left:0">${esc(г.раздел.н)} — ${г.записи.length}</div>
          ${г.записи.slice(0,20).map(x=>`<div class="pr"
            onclick="${г.раздел.открыть?г.раздел.открыть(x):("открытьИзПоиска('"+г.раздел.вид+"','"+(г.раздел.kind||"")+"','"+esc(String(x.id))+"')")}">
            <span class="g" style="white-space:normal">${esc(г.раздел.п(x)||"(без названия)")}</span>
            <span class="st">открыть</span></div>`).join("")}
          ${г.записи.length>20?`<div class="note" style="margin:2px 0 0">… и ещё ${г.записи.length-20};
            уточните запрос</div>`:""}
        </div>`).join("")
      : `<div class="note">Ничего не нашлось.</div>`);
  if(толькоРезультат && document.getElementById("поискРез")){
    document.getElementById("поискРез").innerHTML=тело;
    const с=document.getElementById("поискСчёт"); if(с) с.textContent=всего;
    return;
  }
  document.getElementById("main").innerHTML=`
    <div class="head"><h2>Поиск по базе</h2><span class="sub">найдено: <b id="поискСчёт">${всего}</b></span></div>
    <div class="card" style="padding:20px">
      <div class="search"><span class="ic">⌕</span>
        <input id="поискQ" placeholder="Заводской номер, фамилия, ГОСТ, № акта, объект…"
          value="${esc(_поиск)}" oninput="поискВвод(this.value)"
          onkeydown="if(event.key==='Enter'){ запомнитьПоиск(this.value); }"></div>
      <div id="поискРез">${тело}</div>
    </div>`;
  // Открыли найденное — запрос себя оправдал. Перехват в фазе захвата:
  // обработчик строки сам перерисовывает #main, но _поиск к этому моменту
  // ещё тот, по которому нашли.
  document.getElementById("поискРез").addEventListener("click",e=>{
    if(e.target.closest(".pr")) запомнитьПоиск(_поиск);
  },true);
  const q=document.getElementById("поискQ");
  if(q){ q.focus(); q.setSelectionRange(q.value.length,q.value.length); }
}
// Ctrl+K — открыть поиск с клавиатуры (в Chrome страница это сочетание
// получает, поэтому preventDefault работает; F3 и Ctrl+F не трогаю —
// они ищут по видимой странице, это другая задача).
// Ctrl+S в форме — то же, что кнопка «✓ Сохранить» этой формы (ищется
// по самой кнопке, а не по виду: так не надо помнить десяток функций
// сохранения). Вне форм сочетание остаётся браузеру. «ы» — та же клавиша в
// русской раскладке: e.key там кириллический, поэтому смотрю и e.code.
document.addEventListener("keydown",e=>{
  if(!(e.ctrlKey||e.metaKey) || e.altKey || !(e.code==="KeyS" || /^[sSыЫ]$/.test(e.key))) return;
  if(ВИДЫ_С_ФОРМОЙ.indexOf(curView)<0) return;
  const кн=[...document.querySelectorAll("#main button.btn.primary")].find(b=>/Сохранить|Добавить части/.test(b.textContent) && b.offsetParent!==null);
  if(!кн) return;
  e.preventDefault();
  if(document.activeElement && document.activeElement.blur) document.activeElement.blur();   // дописанное поле — в его значение
  кн.click();
});
// Alt+← / Alt+→ в форме записи журнала и карточки справочника — соседняя
// запись списка (те же кнопки «◀ ▶» в шапке формы). В других видах сочетание
// остаётся браузеру.
document.addEventListener("keydown",e=>{
  if(!e.altKey || e.ctrlKey || e.metaKey || e.shiftKey || (curView!=="edit" && curView!=="regedit")) return;
  if(e.key!=="ArrowLeft" && e.key!=="ArrowRight") return;
  e.preventDefault();
  if(document.activeElement && document.activeElement.blur) document.activeElement.blur();
  листатьНаШаг(e.key==="ArrowLeft"?-1:1);
});
document.addEventListener("keydown",e=>{
  if((e.ctrlKey||e.metaKey) && (e.key==="k"||e.key==="K")){
    e.preventDefault(); if(уходСФормы()){ _возвратВПроверку=false; renderНайти(); }
  }
});

function switchView(v){
  if(v!==curView && !уходСФормы()) return;
  if(_возвратВПроверку && ВИДЫ_С_ФОРМОЙ.indexOf(curView)>=0) v="check";
  curView=v;query=""; _актИзЗаписи=null;
  document.querySelectorAll(".nav").forEach(n=>n.classList.toggle("on",n.dataset.view===v));
  ({dash:renderDash,find:renderНайти,journal:renderJournal,objects:renderObjects,si:renderSi,people:renderPeople,docs:renderDocs,exchange:renderExchange,plan:renderPlan,check:renderCheck,podr:renderPodr,nd:renderNd,premises:renderPremises,lab:renderLab,tasks:renderTasks,passport:renderПаспорт,metod:renderMetod,calendar:renderКалендарь,setup:renderМастер}[v]
    ||(()=>renderPlaceholder(v==="docs"?"Акты и протоколы":"План работ",
       v==="docs"?"Формирование документов по записям журнала — следующий шаг.":"Плановые и внеплановые работы — следующий шаг.")))();}
document.getElementById("side").addEventListener("click",e=>{const b=e.target.closest(".nav");if(b){ _возвратВПроверку=false; switchView(b.dataset.view); }});

