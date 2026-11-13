// ── Назад: возврат на предыдущую страницу в то же место (3.97) ─
// Просьба заказчика: кнопка «Назад» на всех нужных страницах, и по
// возвращении — та же страница в том же месте, откуда перешли.
//
// Как устроено. Страницу открывают десятки функций (switchView, render…,
// edit…, открыть…) — каждая своя. Вместо правки каждой они обёрнуты здесь
// (тот же приём, что `подПравом` в 23б): обёртка запоминает, где мы были
// (вид, прокрутка, поиск раздела, вкладка), зовёт функцию и, если страница
// действительно сменилась (в #main новое содержимое — отказ «есть
// несохранённое» ничего не меняет), кладёт прежнюю страницу в историю.
//
// Формы (запись журнала, карточка, задача…) в историю НЕ кладутся: форма —
// не место, куда возвращаются, а действие. Поэтому:
//  • «Назад» из формы — на страницу, с которой её открыли (поиск, проверка
//    данных, календарь, список), в то же место прокрутки;
//  • «Сохранить» ведёт, как и раньше, в свой список; если с него форму и
//    открывали — список встаёт на прежнее место, а не в начало.
// Перерисовка той же страницы (поиск по буквам, отбор, вкладка) историей
// не считается.

let _навСтек=[], _навТек=null, _навГлуб=0, _навЗапрос=null;
const НАВ_МАКС=40;

// Подпись страницы для кнопки: пункт меню, иначе заголовок страницы.
function навПодпись(){
  const м=document.querySelector('.side .nav[data-view="'+curView+'"]');
  if(м) return Array.from(м.childNodes).filter(n=>n.nodeType===3).map(n=>n.textContent).join(" ").replace(/\s+/g," ").trim();
  const h=document.querySelector("#main .head h2");
  const т=h ? Array.from(h.childNodes).filter(n=>n.nodeType===3).map(n=>n.textContent).join(" ").replace(/\s+/g," ").trim() : "";
  return т || (НАЗАД_КУДА[curView]||"").replace(/^(к|в|на)\s+/,"") || "предыдущая страница";
}
function навСнимок(){
  const m=document.getElementById("main");
  return { вид:curView, ключ:_навТек?_навТек.ключ:"v:"+curView, открыть:_навТек?_навТек.открыть:(в=>()=>switchView(в))(curView),
    прокрутка:m?m.scrollTop:0, запрос:query, вкл:(typeof ВКЛ_РАЗДЕЛА!=="undefined"?ВКЛ_РАЗДЕЛА[curView]:undefined),
    лабвкл:(typeof лабВкладка!=="undefined"?лабВкладка:undefined), подпись:навПодпись() };
}
function этоФорма(вид){ return ВИДЫ_С_ФОРМОЙ.indexOf(вид)>=0; }
function навКлюч(имя, арг, вид){
  if(имя==="switchView" || /^render/.test(имя) && !арг.length) return "v:"+вид;
  let а; try{ а=JSON.stringify(арг); }catch(e){ а=String(арг.length); }
  return вид+"|"+имя+"|"+а;
}
function навПрокрутить(п){
  const m=document.getElementById("main"); if(!m) return;
  const ставь=()=>{ m.scrollTop=п; };
  ставь(); requestAnimationFrame(ставь); setTimeout(ставь,60); setTimeout(ставь,200);
}

function навОбернуть(имя){
  const f=window[имя]; if(typeof f!=="function" || f.__нав) return;
  const обёртка=function(...арг){
    if(_навГлуб>0) return f.apply(this,арг);
    const m=document.getElementById("main"), узел=m&&m.firstElementChild;
    const до=навСнимок(), изФормы=этоФорма(до.вид);
    // Выход из формы на страницу, с которой её открыли (вершина истории), —
    // это возврат: поиск раздела восстанавливается ДО отрисовки.
    const верх=_навСтек[_навСтек.length-1];
    if(изФормы && верх && имя==="switchView"){
      let ц=String(арг[0]||"").split(":")[0];
      if(_возвратВПроверку) ц=typeof _возвратВПроверку==="string"?_возвратВПроверку:"check";
      if(верх.ключ==="v:"+ц && верх.запрос) _навЗапрос=верх.запрос;
    }
    const готово=()=>{
      _навЗапрос=null;
      const m2=document.getElementById("main");
      if(!m2 || m2.firstElementChild===узел) return;           // не перешли
      const ключ=навКлюч(имя, арг, curView);
      // страница раздела открывается переходом (render… сам вид не ставит)
      const открыть=ключ.slice(0,2)==="v:" ? (в=>()=>switchView(в))(curView) : ()=>обёртка.apply(null,арг);
      if(_навТек && ключ===_навТек.ключ){ _навТек={ключ, открыть}; навКнопки(); return; } // перерисовка
      if(изФормы){
        // форма в историю не кладётся; вернулись туда, откуда её открыли, —
        // прежнее место прокрутки
        if(верх && верх.ключ===ключ){ _навСтек.pop(); навПрокрутить(верх.прокрутка); }
      } else if(_навТек){
        _навСтек.push(до); if(_навСтек.length>НАВ_МАКС) _навСтек.shift();
      }
      _навТек={ключ, открыть}; навКнопки();
    };
    _навГлуб++;
    let r; try{ r=f.apply(this,арг); } finally{ _навГлуб--; }
    if(r && typeof r.then==="function") r.then(готово, готово); else готово();
    return r;
  };
  обёртка.__нав=true; window[имя]=обёртка;
}

// Назад: открыть прежнюю страницу тем же способом, каким её открывали, и
// вернуть прокрутку, поиск и вкладку. Не открылась (несохранённое, «Отмена»)
// — история не трогается.
// Форма, открытая из «Проверки данных», паспорта, карточки лаборатории,
// календаря (_возвратВПроверку), возвращается туда, как и раньше, — даже
// если в истории перед ней другая страница.
function навЦельФормы(){ return _возвратВПроверку && этоФорма(curView) ? (typeof _возвратВПроверку==="string"?_возвратВПроверку:"check") : null; }
function навНазад(запасной){
  const e=_навСтек[_навСтек.length-1], цель=навЦельФормы();
  if(цель && (!e || e.вид!==цель)){ switchView(цель); return; }
  if(!e){ if(запасной) switchView(запасной); return; }
  const m=document.getElementById("main"), узел=m&&m.firstElementChild;
  if(e.вкл!==undefined && typeof ВКЛ_РАЗДЕЛА!=="undefined") ВКЛ_РАЗДЕЛА[e.вид]=e.вкл;
  if(e.лабвкл!==undefined && e.вид==="lab") лабВкладка=e.лабвкл;
  const стек=_навСтек.slice();
  _навЗапрос=e.запрос||null; _возвратВПроверку=false;
  _навГлуб++;                       // сам переход историей не считается
  try{ e.открыть(); } finally{ _навГлуб--; _навЗапрос=null; }
  if(!m || m.firstElementChild===узел) return;
  _навСтек=стек; _навСтек.pop();
  _навТек={ключ:e.ключ, открыть:e.открыть};
  if(e.запрос && query!==e.запрос){ query=e.запрос; _навГлуб++; try{ e.открыть(); } finally{ _навГлуб--; } query=e.запрос; }
  навПрокрутить(e.прокрутка);
  if(typeof подсветитьМеню==="function" && МЕНЮ_РОДИТЕЛЬ[curView]) подсветитьМеню(curView);
  навКнопки();
}
function навЕсть(){ return _навСтек.length>0; }
// Кнопка в шапке и кнопки «← …» форм (кнопкаНазад): подпись — куда ведёт.
function навКнопки(){
  const e=_навСтек[_навСтек.length-1];
  const к=document.getElementById("назадКн");
  if(к){ к.disabled=!e; к.title=e?"Назад: "+e.подпись+" (Alt+←)":"Назад — некуда"; }
  if(!e || навЦельФормы()) return;
  document.querySelectorAll("#main button.назад[data-нав]").forEach(b=>{ b.textContent="← "+e.подпись; });
}

// Что оборачивается: переход по меню и все функции, открывающие страницу.
[
  "switchView",
  "renderJournal","renderObjects","renderSi","renderPeople","renderDocs","renderExchange","renderPlan","renderCheck",
  "renderPodr","renderNd","renderPremises","renderLab","renderTasks","renderПаспорт","renderMetod","renderКалендарь",
  "renderВыдачаСИ","renderВибрация","renderНормыВибр","renderСМК","renderПредложения","renderЗаявки","renderРаскладка",
  "renderПГ","renderЛистУдостоверений","renderРуководство","renderБирки","renderПодготовка","renderDash","renderНайти",
  "renderОтчёт","renderAkt","renderРазмножить",
  "editReg","editJournal","editMetodika","карточкаОбъектаЛист","editЗамерВибр","открытьКартуВибрации",
  "печатьКартВибрации","частиСпискомОткрыть","editППР","editПГ","editПГП","открытьПротокол","editTask",
  "editЗаметка","открытьЗаявку"
].forEach(навОбернуть);

// Мышь: боковая кнопка «назад»; клавиатура: Alt+← вне форм (в форме записи
// и карточки Alt+← листает записи — там своя кнопка «←»).
window.addEventListener("mouseup",e=>{ if(e.button===3){ e.preventDefault(); навНазад(); } });
window.addEventListener("keydown",e=>{
  if(e.defaultPrevented || !e.altKey || e.ctrlKey || e.metaKey || e.shiftKey || e.key!=="ArrowLeft") return;
  if(curView==="edit" || curView==="regedit") return;
  const эл=document.activeElement;
  if(эл && (эл.tagName==="INPUT" || эл.tagName==="TEXTAREA" || эл.isContentEditable)) return;
  if(!навЕсть()) return;
  e.preventDefault(); навНазад();
});
