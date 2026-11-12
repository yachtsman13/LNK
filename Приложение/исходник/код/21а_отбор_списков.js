
// ── Сортировка по заголовкам и отбор во всех списках ──────────────
// Один механизм на все таблицы-списки программы, а не своя сортировка в
// каждом разделе: после каждой отрисовки он находит в #main таблицы и
//  - делает заголовки щёлкаемыми: ▲ по возрастанию, ▼ по убыванию, третий
//    щелчок — прежний порядок раздела. Пустые значения всегда внизу, как в
//    журнале. Даты ДД.ММ.ГГГГ и числа сравниваются как даты и числа;
//  - под заголовком (в списках от 8 строк) ставит строку отбора: у колонки
//    с немногими разными значениями — выпадающий список, у остальных —
//    поле «содержит». В поле понимаются «<01.01.2027», «>=5», «!текст»
//    (не содержит) и «—» (пусто).
// Строки только переставляются и прячутся на месте — раздел не
// перерисовывается, поэтому ввод в поле отбора не теряет фокус.
// Порядок сортировки помнится (настройка места lnk_list_sort), отбор — до
// перезагрузки страницы: иначе после перезапуска список выглядел бы
// «потерявшим» записи.
//
// Таблицы со своей сортировкой (журнал, задачи, специалисты, прогноз
// аттестации — у них заголовки уже th.sortable) он не пересортировывает,
// но отбор добавляет. Журнал и список задач — со своими отборами, туда
// он не встаёт вовсе (data-отбор="нет"). Не трогает листы документов
// (паспорт, акт, протоколы), карту работ, календарь, формы и таблицы с
// полями ввода или строками-разделами.
//
// Excel и листание карточек идут «как на экране»: у таблиц справочников
// есть data-список и у строк data-id; какНаЭкране() упорядочивает и
// отбирает массив записей так же, как показаны строки.
const КЛЮЧ_СОРТ_СПИСКОВ="lnk_list_sort";
const ПОРОГ_СТРОКИ_ОТБОРА=8;      // меньше строк — строка отбора не нужна
const ПОРОГ_СПИСКА_ОТБОРА=20;     // столько разных значений — ещё выпадающий список
let _сортСписков=null;            // {ключ:{i, н}} — н: 1 по возрастанию, -1 по убыванию
const _отборСписков={};           // {ключ:{i: значение}}
const _видимыеСписков={};         // {ключ:[id …]} — как показано после отбора и сортировки
const _текстЯчеек=new WeakMap();  // строка → [тексты ячеек]
function сортСписков(){
  if(!_сортСписков){ try{ _сортСписков=JSON.parse(ХР.getItem(КЛЮЧ_СОРТ_СПИСКОВ)||"{}")||{}; }catch(e){ _сортСписков={}; } }
  return _сортСписков;
}
function сохранитьСортСписков(){ try{ ХР.setItem(КЛЮЧ_СОРТ_СПИСКОВ, JSON.stringify(сортСписков())); }catch(e){} }
function нормОтб(s){ return String(s==null?"":s).replace(/\s+/g," ").trim(); }
function нижеОтб(s){ return нормОтб(s).toLowerCase().replace(/ё/g,"е"); }
// Ключ сравнения: {ч:число} для дат и чисел, {с:строка} для текста, null — пусто.
function ключСортСписка(т){
  const s=нормОтб(т);
  if(!s || /^[—–-]+$/.test(s)) return null;
  let m=/^(\d{2})\.(\d{2})\.(\d{4})/.exec(s); if(m) return {ч:+(m[3]+m[2]+m[1])};
  m=/^(\d{4})-(\d{2})-(\d{2})/.exec(s);       if(m) return {ч:+(m[1]+m[2]+m[3])};
  const n=s.replace(/[\s ]/g,"").replace(",",".");
  if(/^[-+]?\d+(\.\d+)?$/.test(n)) return {ч:+n};
  return {с:s.toLowerCase()};
}
function сравнитьКлючиСписка(a,b){
  if(a==null && b==null) return 0;
  if(a==null) return 1; if(b==null) return -1;          // пустые — вниз при любом направлении
  if(a.ч!=null && b.ч!=null) return a.ч-b.ч;
  if(a.ч!=null) return -1; if(b.ч!=null) return 1;
  return a.с.localeCompare(b.с,"ru",{numeric:true});
}
// Условие отбора из поля «содержит».
function условиеОтбора(значение){
  const v=нормОтб(значение); if(!v) return null;
  const m=/^(<=|>=|≤|≥|<|>)\s*(.+)$/.exec(v);
  if(m){ const к=ключСортСписка(m[2]);
    if(к && к.ч!=null){ const оп=m[1].replace("≤","<=").replace("≥",">=");
      return т=>{ const x=ключСортСписка(т); if(!x||x.ч==null) return false;
        return оп==="<"?x.ч<к.ч : оп===">"?x.ч>к.ч : оп==="<="?x.ч<=к.ч : x.ч>=к.ч; }; } }
  if(/^[—–-]+$/.test(v) || v.toLowerCase()==="пусто") return т=>ключСортСписка(т)==null;
  if(v[0]==="!" && v.length>1){ const ч=нижеОтб(v.slice(1)); return т=>!нижеОтб(т).includes(ч); }
  const ч=нижеОтб(v); return т=>нижеОтб(т).includes(ч);
}
function текстЯчейкиОтб(td){
  if(!td) return "";
  if(td.dataset && td.dataset.s!=null) return td.dataset.s;
  return td.innerText;
}
function годнаДляОтбора(t){
  if(t.dataset.отбор==="нет") return false;
  if(t.closest(".sheet,.doctbl,.тепло,.form,.drawer,.протлист,.календарь,.удлист,.noотбор")) return false;
  if(/doctbl|тепло|mtab|вбтабл|вбввод|ппрформа|карт\b/.test(t.className)) return false;
  if(!t.tHead || t.tHead.rows.length!==1 || !t.tBodies[0]) return false;
  if([...t.tHead.rows[0].cells].some(c=>c.colSpan>1)) return false;
  const тело=[...t.tBodies[0].rows];
  if(тело.length<2) return false;
  if(тело.some(r=>[...r.cells].some(c=>c.colSpan>1))) return false;
  if(t.tBodies[0].querySelector("select,textarea,input:not([type=checkbox])")) return false;
  return true;
}
function ключСписка(t){
  if(t.dataset.список) return t.dataset.список;
  return (typeof curView!=="undefined"?curView:"")+"|"+[...t.tHead.rows[0].cells].map(c=>нормОтб(c.innerText).replace(/[▲▼]/g,"").trim()).join("/");
}
function обработатьТаблицуОтбора(t){
  if(t.dataset.отб) return;
  if(!годнаДляОтбора(t)){ t.dataset.отб="0"; return; }
  t.dataset.отб="1";
  const ключ=ключСписка(t); t.dataset.отбКлюч=ключ;
  const шапка=[...t.tHead.rows[0].cells];
  const ряды=[...t.tBodies[0].rows];
  ряды.forEach((r,k)=>{ r.dataset.поряд=k; _текстЯчеек.set(r,[...r.cells].map(текстЯчейкиОтб)); });
  const колонки=шапка.map((th,i)=>({i, имя:нормОтб(th.innerText),
    свой:th.classList.contains("sortable"), пусто:!нормОтб(th.innerText)||!!th.querySelector("input")}));
  // Щелчок по заголовку
  колонки.forEach(к=>{ if(к.свой||к.пусто) return;
    const th=шапка[к.i]; th.classList.add("отбсорт"); th.title="Сортировать по этой колонке (щёлкните ещё раз — обратный порядок, третий раз — как было)";
    th.addEventListener("click",()=>{ const с=сортСписков(), б=с[ключ];
      if(!б||б.i!==к.i) с[ключ]={i:к.i,н:1}; else if(б.н===1) с[ключ]={i:к.i,н:-1}; else delete с[ключ];
      сохранитьСортСписков(); применитьОтбор(t); }); });
  // Строка отбора
  const отб=_отборСписков[ключ]||{};
  const есть=Object.keys(отб).some(i=>отб[i]);
  if(ряды.length>=ПОРОГ_СТРОКИ_ОТБОРА || есть){
    const стр=document.createElement("tr"); стр.className="отбстр noprint";
    колонки.forEach(к=>{ const td=document.createElement("td");
      if(!к.пусто){
        const знач=[...new Set(ряды.map(r=>нормОтб(_текстЯчеек.get(r)[к.i])))];
        // Даты и числа — всегда поле: в нём понимаются «<01.01.2027», «>=5».
        const непуст=знач.map(ключСортСписка).filter(Boolean);
        const числовая=непуст.length>0 && непуст.filter(x=>x.ч!=null).length>=непуст.length*0.6;
        const список=!числовая && знач.length<=ПОРОГ_СПИСКА_ОТБОРА && знач.length>1 && знач.length<ряды.length;
        if(список){
          const s=document.createElement("select"); s.dataset.i=к.i; s.title="Отбор по колонке «"+к.имя+"»";
          const пустые=знач.some(x=>!x||/^[—–-]+$/.test(x));
          s.innerHTML=`<option value="">все</option>`+знач.filter(x=>x&&!/^[—–-]+$/.test(x))
            .sort((a,b)=>сравнитьКлючиСписка(ключСортСписка(a),ключСортСписка(b)))
            .map(x=>`<option value="=${esc(x)}">${esc(x.length>40?x.slice(0,39)+"…":x)}</option>`).join("")+
            (пустые?`<option value="—">(пусто)</option>`:"");
          s.value=отб[к.i]||"";
          s.addEventListener("change",()=>{ задатьОтбор(ключ,к.i,s.value); применитьОтбор(t); });
          td.appendChild(s);
        } else if(знач.length>1){
          const inp=document.createElement("input"); inp.type="search"; inp.dataset.i=к.i;
          inp.placeholder="отбор…"; inp.value=отб[к.i]||"";
          inp.title="Отбор по колонке «"+к.имя+"»: текст — содержит; !текст — не содержит; <01.01.2027, >=5 — даты и числа; — пусто";
          inp.addEventListener("input",()=>{ задатьОтбор(ключ,к.i,inp.value); применитьОтбор(t); });
          td.appendChild(inp);
        }
      }
      стр.appendChild(td); });
    t.tHead.appendChild(стр);
  }
  // Итог отбора — над таблицей
  const итог=document.createElement("div"); итог.className="отбитог"; итог.style.display="none";
  const куда=t.closest(".tblwrap")||t;
  куда.parentNode.insertBefore(итог, куда);
  t._отбИтог=итог;
  применитьОтбор(t);
}
function задатьОтбор(ключ,i,v){
  const о=_отборСписков[ключ]||(_отборСписков[ключ]={});
  if(v) о[i]=v; else delete о[i];
}
function сброситьОтборТаблицы(ключ){
  delete _отборСписков[ключ];
  document.querySelectorAll("#main table[data-отб='1']").forEach(t=>{ if(t.dataset.отбКлюч!==ключ) return;
    t.querySelectorAll(".отбстр select,.отбстр input").forEach(x=>{ x.value=""; }); применитьОтбор(t); });
}
function применитьОтбор(t){
  const ключ=t.dataset.отбКлюч, тело=t.tBodies[0]; if(!тело) return;
  const ряды=[...тело.rows].filter(r=>r.dataset.поряд!=null);
  const отб=_отборСписков[ключ]||{};
  const условия=Object.keys(отб).filter(i=>отб[i]).map(i=>{ const v=отб[i];
    if(v[0]==="="){ const н=нормОтб(v.slice(1)); return {i:+i, ф:т=>нормОтб(т)===н}; }
    if(v==="—") return {i:+i, ф:т=>ключСортСписка(т)==null};
    const ф=условиеОтбора(v); return ф?{i:+i,ф}:null; }).filter(Boolean);
  let видно=0;
  ряды.forEach(r=>{ const тт=_текстЯчеек.get(r)||[];
    const ок=условия.every(у=>у.ф(тт[у.i]||"")); r.style.display=ок?"":"none"; r.classList.toggle("отбскрыт",!ок); if(ок) видно++; });
  const с=сортСписков()[ключ];
  const порядок=ряды.slice().sort((a,b)=>{
    if(с){ const тa=(_текстЯчеек.get(a)||[])[с.i], тb=(_текстЯчеек.get(b)||[])[с.i];
      const ка=ключСортСписка(тa), кб=ключСортСписка(тb);
      if(ка==null||кб==null){ const d=сравнитьКлючиСписка(ка,кб); if(d) return d; }
      else { const d=сравнитьКлючиСписка(ка,кб)*с.н; if(d) return d; } }
    return (+a.dataset.поряд)-(+b.dataset.поряд); });
  порядок.forEach(r=>тело.appendChild(r));
  // Пустой результат — своя строка
  let пусто=тело.querySelector("tr.отбпусто");
  if(!видно && ряды.length){
    if(!пусто){ пусто=document.createElement("tr"); пусто.className="отбпусто";
      пусто.innerHTML=`<td colspan="${t.tHead.rows[0].cells.length}" class="empty">По этому отбору ничего нет.</td>`; }
    тело.appendChild(пусто);
  } else if(пусто) пусто.remove();
  // Стрелки
  [...t.tHead.rows[0].cells].forEach((th,i)=>{ if(!th.classList.contains("отбсорт")) return;
    let ст=th.querySelector(".отбстрелка");
    if(!ст){ ст=document.createElement("span"); ст.className="отбстрелка"; th.appendChild(ст); }
    ст.textContent=с&&с.i===i?(с.н===1?" ▲":" ▼"):""; });
  // Итог
  const итог=t._отбИтог;
  if(итог){
    if(условия.length){
      const имена=[...t.tHead.rows[0].cells].map(c=>нормОтб(c.innerText).replace(/[▲▼]/g,"").trim());
      итог.innerHTML=`Отбор: ${условия.map(у=>`<b>${esc(имена[у.i]||"")}</b> — ${esc(String(отб[у.i]).replace(/^=/,""))}`).join("; ")}
        · показано <b>${видно}</b> из ${ряды.length}
        <button class="btn small noprint" onclick="сброситьОтборТаблицы(${esc(JSON.stringify(ключ))})">✕ сбросить отбор</button>`;
      итог.style.display="";
    } else { итог.style.display="none"; итог.innerHTML=""; }
  }
  if(t.dataset.список) _видимыеСписков[ключ]=порядок.filter(r=>r.style.display!=="none" && r.dataset.id!=null).map(r=>r.dataset.id);
}
// Массив записей — как показан список: тем же отбором и порядком. Пока
// список не открывали или отбора и сортировки нет — как есть.
function отборАктивен(ключ){ const о=_отборСписков[ключ]; return !!(о && Object.keys(о).some(i=>о[i])); }
// ид(x, k) — чем строка помечена в таблице (data-id); по умолчанию x.id.
function какНаЭкране(ключ, ряды, ид){
  const вид=_видимыеСписков[ключ], с=сортСписков()[ключ], ф=отборАктивен(ключ);
  if(!вид || (!с && !ф)) return ряды;
  const поз=new Map(вид.map((id,k)=>[String(id),k]));
  const пары=ряды.map((x,k)=>({x, к:String(ид?ид(x,k):x.id)}));
  const есть=пары.filter(п=>поз.has(п.к)).sort((a,b)=>поз.get(a.к)-поз.get(b.к)).map(п=>п.x);
  return ф ? есть : есть.concat(пары.filter(п=>!поз.has(п.к)).map(п=>п.x));
}
function обработатьТаблицыОтбора(){
  document.querySelectorAll("#main table:not([data-отб])").forEach(t=>{ try{ обработатьТаблицуОтбора(t); }catch(e){ t.dataset.отб="0"; } });
}
(function(){
  const m=document.getElementById("main"); if(!m || typeof MutationObserver==="undefined") return;
  new MutationObserver(()=>обработатьТаблицыОтбора()).observe(m,{childList:true, subtree:true});
})();
