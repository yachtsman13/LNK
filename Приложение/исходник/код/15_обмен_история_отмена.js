// ── Заглушки и обмен ──────────────────────────────────────
function renderPlaceholder(t,x){document.getElementById("main").innerHTML=
  `<div class="head"><h2>${t}</h2></div><div class="placeholder"><div><div class="big">▦</div>${x}</div></div>`;}
// ── Обмен между рабочими местами ──────────────────────────
// Файлами, без сети. Слияние по id записи: новых нет — добавляем;
// есть — берём ту, у кого «изменено» свежее. Код места входит в id
// новых записей, поэтому места не затирают друг друга.
function mergeJournal(incoming){
  let added=0, updated=0, skipped=0;
  const idx={}; J.forEach((z,i)=>{ idx[z.id]=i; });
  incoming.forEach(r=>{
    if(!r||!r.id) return;
    if(idx[r.id]==null){ J.push(r); idx[r.id]=J.length-1; if(!r.удалено) added++; }
    else { const cur=J[idx[r.id]];
      if(конфликтПравок(STORE_KEY,cur,r)) записатьКонфликт(STORE_KEY,cur,r);
      if(lastTs(r) > lastTs(cur)){ J[idx[r.id]]=r; updated++; }
      else skipped++; }
  });
  if(added||updated) saveJournal();
  return {added,updated,skipped};
}
// ── Правки одной записи с двух мест ───────────────────────
// Слияние при обмене — «свежее побеждает». Если одну и ту же запись между
// обменами поправили на ДВУХ местах, более ранняя правка молча терялась:
// её не видно ни в отчёте приёма, ни в «Что изменилось». Теперь такая
// пара распознаётся, отброшенная версия сохраняется (lnk_conflicts) и её
// можно сравнить и вернуть в разделе «Обмен файлами». Само правило
// слияния НЕ менялось — меняется только то, что потеря больше не молчит.
//
// Как отличить «две независимые правки» от «правка поверх моей»: у записи
// появилась отметка `база` — версия (отметка времени), от которой пошла
// цепочка правок ЭТОГО места. Ставится сама в saveReg/saveJournal по
// снимку, снятому при прошлой записи: запись изменилась, последним её
// правило это место, а до того — другое место (или никто) → база = то,
// что было. Поэтому её не надо помнить в каждой из десятков функций
// сохранения. Пришедшая версия — конфликт, если у меня есть своя правка
// (место = я, база есть), пришла не та версия, от которой я правил, и не
// правка поверх моей (её база ≠ моей версии).
const REG_КОНФЛ="lnk_conflicts";
let КОНФЛ=[];
const _снимки={};
// У задачи «место» — поле «Рабочее место исполнителя», отметка правки —
// место_правки (та же ловушка, что была у «Расположения» объекта).
function местоПравки(key,r){ return key===REG_ЗАД ? String(r.место_правки||"") : String(r.место||""); }
function снятьСнимок(key,arr){
  const с={};
  // j — содержимое записи: по нему «История правок» видит, ЧТО поменялось
  (arr||[]).forEach(r=>{ if(r && r.id!=null) с[r.id]={изм:lastTs(r), место:местоПравки(key,r), j:JSON.stringify(r)}; });
  _снимки[key]=с;
}
function проставитьБазу(key,arr){
  const с=_снимки[key]; if(!с || key===REG_КОНФЛ) return;
  const я=getWS().код;
  (arr||[]).forEach(r=>{
    if(!r || r.id==null) return;
    const s=с[r.id]; if(!s) return;                 // новая запись — у неё нет чужой версии
    // У записи из книги/сида отметки времени нет вовсе — база тогда
    // «исходная версия», а не пустота (пустая база значила бы «не знаю»).
    // Своя цепочка правок, у которой базы ещё нет (запись заведена здесь же
    // или правилась старой версией программы), получает базой свою прошлую
    // версию: иначе правка другого места поверх неё прошла бы незамеченной.
    if(lastTs(r)!==s.изм && местоПравки(key,r)===я && (s.место!==я || !r.база))
      r.база=s.изм||БАЗА_ИСХОДНАЯ;
  });
}
function реестрПоКлючу(key){
  return ({[STORE_KEY]:()=>J,[REG_OBJ]:()=>O,[REG_SI]:()=>S,[REG_PPL]:()=>P,[REG_AKT]:()=>A,
    [REG_ППР]:()=>ОТМ,[REG_ППР_ПЛАН]:()=>ППР_ПОЗ,[REG_ПОДР]:()=>ПОДР,[REG_ПАДЕЖ]:()=>ПАДЕЖИ,
    [REG_ЗАД]:()=>ЗАД,[REG_НД]:()=>НД,[REG_ПОМ]:()=>ПОМ}[key]||(()=>null))();
}
function снятьВсеСнимки(){
  [STORE_KEY,REG_OBJ,REG_SI,REG_PPL,REG_AKT,REG_ППР,REG_ППР_ПЛАН,REG_ПОДР,REG_ПАДЕЖ,REG_ЗАД,REG_НД,REG_ПОМ]
    .forEach(k=>{ const a=реестрПоКлючу(k); if(a) снятьСнимок(k,a); });
  _историяВключена=true;
}
const БАЗА_ИСХОДНАЯ="исходная";

// ── История правок записи ─────────────────────────────────
// Кто, когда и с какого значения на какое менял поле. Пишется в ОДНОМ
// месте — в saveReg/saveJournal, по снимку прошлой записи (тот же
// _снимки, что у конфликтов), поэтому не надо помнить о ней в каждой из
// десятков функций сохранения. Запись попадает в историю, только если у
// неё сменилась отметка времени (правка, удаление, приём из обмена —
// тогда место правки другое); служебные пересохранения без отметки
// (починки при загрузке) историей не считаются. Хранится у каждого места
// своя (в обмен не идёт — у другого места свой глаз), в резервную копию
// идёт. На запись — последние ИСТ_НА_ЗАПИСЬ правок, значения обрезаны:
// память браузера не резиновая.
const REG_ИСТ="lnk_history";
const ИСТ_НА_ЗАПИСЬ=15, ИСТ_ДЛИНА=300;
let ИСТ=null;
function ист(){
  if(!ИСТ){ try{ ИСТ=JSON.parse(ХР.getItem(REG_ИСТ)||"{}")||{}; }catch(e){ ИСТ={}; } }
  return ИСТ;
}
function сохранитьИсторию(){
  try{ ХР.setItem(REG_ИСТ, JSON.stringify(ист())); }catch(e){ переполнение(e); }
}
const ИСТ_СЛУЖЕБНЫЕ=new Set(["изменено","база","место_правки","удалено"]);
function плоскоИст(key,o,пр,выход,гл){
  Object.keys(o||{}).forEach(k=>{
    if(!пр && (ИСТ_СЛУЖЕБНЫЕ.has(k) || (k==="место" && key!==REG_ЗАД))) return;
    const v=o[k];
    if(v && typeof v==="object" && !Array.isArray(v) && гл<3) плоскоИст(key,v,пр+k+".",выход,гл+1);
    else выход[пр+k]=v==null?"":(Array.isArray(v)?(v.length?JSON.stringify(v):""):String(v));
  });
  return выход;
}
// Пока идёт начальная загрузка, правки делает сама программа (починки,
// добавление документов из актуализации) — это не правки людей, в
// историю не пишу. Включается в конце загрузки (снятьВсеСнимки).
let _историяВключена=false;
function записатьИсторию(key,arr){
  if(!_историяВключена) return;
  const с=_снимки[key];
  if(!с || key===REG_КОНФЛ) return;          // снимка нет — идёт начальная загрузка
  const Н=ист(); let было=false;
  const обрез=v=>{ v=String(v); return v.length>ИСТ_ДЛИНА?v.slice(0,ИСТ_ДЛИНА)+"…":v; };
  (arr||[]).forEach(r=>{
    if(!r || r.id==null) return;
    const s=с[r.id], ts=lastTs(r);
    if(s && ts===s.изм) return;
    if(!s && !ts) return;
    const зап={когда:ts||new Date().toISOString(), место:местоПравки(key,r)};
    if(!s) зап.новая=1;
    else{
      const ст=JSON.parse(s.j);
      if(r.удалено && !ст.удалено) зап.удалена=1;
      if(!r.удалено && ст.удалено) зап.возвращена=1;
      const а=плоскоИст(key,ст,"",{},0), б=плоскоИст(key,r,"",{},0), поля=[];
      new Set(Object.keys(а).concat(Object.keys(б))).forEach(k=>{
        const x=а[k]||"", y=б[k]||"";
        if(x!==y) поля.push([k,обрез(x),обрез(y)]);
      });
      if(!поля.length && !зап.удалена && !зап.возвращена) return;
      зап.поля=поля;
    }
    const ключ=key+"|"+r.id;
    Н[ключ]=[зап].concat(Н[ключ]||[]).slice(0,ИСТ_НА_ЗАПИСЬ);
    было=true;
  });
  if(было) сохранитьИсторию();
}
const ИСТ_МЕТОД={н:"№ протокола",исп:"исполнитель",файл:"файл протокола",приборы:"приборы",нд:"норм. документы"};
function полеИст(key,k){
  if(key===STORE_KEY) return ЖУРНАЛ_ПОЛЯ.find(f=>f.k===k);
  if(key===REG_ЗАД) return ЗАДАЧИ_ПОЛЯ.find(f=>f.k===k);
  if(key===REG_ППР_ПЛАН) return ПЛАН_ПОЛЯ.find(f=>f.k===k);
  const C=Object.values(REGS).find(C=>C.key===key);
  return C ? C.fields.find(f=>f.k===k) : null;
}
function подписьИст(key,путь){
  const ч=путь.split(".");
  if(key===STORE_KEY && ч[0]==="методы" && ч.length>1)
    return ч[1]+(ч[2]?": "+(ИСТ_МЕТОД[ч[2]]||ч.slice(2).join(".")):"");
  if(key===REG_ППР_ПЛАН && ч[0]==="план" && ч.length===3)
    return "план "+ч[1]+", "+String(МЕСЯЦЫ[+ч[2]-1]||ч[2]).toLowerCase()+" (ч)";
  const особые={объект_ид:"связь с карточкой объекта",статус:"Состояние",выполнено:"Дата выполнения",
    нп:"№ п/п",год:"Год",месяцы_поверки:"Месяц представления на поверку"};
  const f=полеИст(key,ч[0]);
  const л=f?(f.l||f.л):(особые[ч[0]]||ч[0]);
  return л+(ч.length>1?" ("+ч.slice(1).join(".")+")":"");
}
function значениеИст(key,путь,v){
  if(v==="") return "<span class='muted'>пусто</span>";
  const f=полеИст(key,путь.split(".")[0]);
  if(f && f.type==="podr"){ const x=найтиПоИд(ПОДР,v); if(x) return esc(x.наименование); }
  if(f && f.type==="ref" && f.ref){ const x=найтиПоИд(f.ref()||[],v); if(x) return esc(x[f.refName]||v); }
  if(/^\d{4}-\d{2}-\d{2}$/.test(v)) return ru(v);
  return esc(v);
}
function блокИстории(key,id){
  const сп=ист()[key+"|"+id]||[];
  const строки=сп.map(е=>{
    const когда=ru(String(е.когда).slice(0,10))+" "+String(е.когда).slice(11,16);
    const что=е.новая?"<i>запись заведена</i>":
      [е.удалена?"<i>запись удалена</i>":"", е.возвращена?"<i>запись возвращена</i>":""]
        .concat((е.поля||[]).map(([k,x,y])=>`<b>${esc(подписьИст(key,k))}</b>: ${значениеИст(key,k,x)} → ${значениеИст(key,k,y)}`))
        .filter(Boolean).join("<br>");
    return `<tr><td style="white-space:nowrap">${esc(когда)}</td><td>${esc(е.место||"—")}</td><td>${что}</td></tr>`;
  }).join("");
  return `<details class="история noprint" style="margin-top:12px"><summary>История правок (${сп.length})</summary>
    ${сп.length?`<table class="mtab" style="margin-top:6px"><thead><tr><th>Когда</th><th>Место</th><th>Что изменилось</th></tr></thead>
      <tbody>${строки}</tbody></table>
      <div class="note">Показаны последние ${ИСТ_НА_ЗАПИСЬ} правок, которые видело это рабочее место (свои и пришедшие
        обменом). Правки до версии 3.16 не записывались.</div>`
    :`<div class="note">Правок пока не было — или они сделаны до версии 3.16, когда история ещё не записывалась.</div>`}
  </details>`;
}

// ── Отмена последней загрузки ─────────────────────────────
// Импорт из Excel и массовая правка меняют сразу много записей, а пустая
// ячейка теперь стирает значение — ошибиться легко, а вручную собрать
// обратно нечего. Перед такой правкой программа запоминает, как было
// (отменаНачать), после — ЧТО именно изменилось (отменаЗавершить): старые
// версии изменённых записей и id добавленных. Хранится одна, последняя.
// Отмена возвращает содержимое записей и ставит им свежую отметку
// «изменено» (иначе на местах, куда правка уже ушла обменом, отмена бы не
// дошла), добавленные записи — удаляет пометкой (тоже уходит в обмен).
// Запись, которую после загрузки успели поправить ещё раз, не трогаю —
// называю: откатывать поверх новой правки значило бы потерять её.
const REG_ОТМЕНА="lnk_undo";
let _отменаПодг=null;
function отменаЗагрузить(){ try{ return JSON.parse(ХР.getItem(REG_ОТМЕНА)||"null"); }catch(e){ return null; } }
function отменаСохранить(о){
  try{ if(о) ХР.setItem(REG_ОТМЕНА, JSON.stringify(о)); else ХР.removeItem(REG_ОТМЕНА); }
  catch(e){ toast("Не хватило памяти запомнить состояние для отмены — эту загрузку отменить будет нельзя"); }
  обновитьКнопкуОтмены();
}
function отменаНачать(описание, ключи){
  const до={};
  ключи.forEach(k=>{ const м={}; (реестрПоКлючу(k)||[]).forEach(r=>{ if(r && r.id!=null) м[String(r.id)]=JSON.stringify(r); }); до[k]=м; });
  _отменаПодг={описание, ключи, до, вид:curView};
}
function отменаЗавершить(){
  const п=_отменаПодг; _отменаПодг=null; if(!п) return;
  const части=[];
  п.ключи.forEach(k=>{
    const до=п.до[k], старые=[], новые=[], после={};
    (реестрПоКлючу(k)||[]).forEach(r=>{
      if(!r || r.id==null) return;
      const ид=String(r.id);
      if(!(ид in до)){ новые.push(ид); после[ид]=lastTs(r); }
      else if(до[ид]!==JSON.stringify(r)){ старые.push(JSON.parse(до[ид])); после[ид]=lastTs(r); }
    });
    if(старые.length||новые.length) части.push({k, старые, новые, после});
  });
  if(!части.length) return;
  отменаСохранить({описание:п.описание, вид:curView, когда:new Date().toISOString(), место:getWS().код, части});
}
function отменаСчёт(о){ return о.части.reduce((n,ч)=>n+ч.старые.length+ч.новые.length,0); }
function обновитьКнопкуОтмены(){
  const el=document.getElementById("fОтмена"); if(!el) return;
  const о=отменаЗагрузить();
  el.innerHTML=о?`<button class="btn" style="padding:2px 10px" onclick="отменитьПоследнее()"
    title="Вернуть записи к виду до этой правки">↶ Отменить: ${esc(о.описание)} (${String(о.когда).slice(11,16)})</button>`:"";
}
function отменитьПоследнее(){
  const о=отменаЗагрузить(); if(!о){ toast("Отменять нечего"); return; }
  const поздние=[];
  о.части.forEach(ч=>{
    const arr=реестрПоКлючу(ч.k)||[];
    ч.старые.map(x=>String(x.id)).concat(ч.новые).forEach(ид=>{
      const r=найтиПоИд(arr,ид);
      if(r && lastTs(r)!==ч.после[ид]) поздние.push(r.наименование||r.объект||r.что||ид);
    });
  });
  const всего=отменаСчёт(о);
  if(!confirm("Отменить «"+о.описание+"» ("+ru(о.когда.slice(0,10))+" "+о.когда.slice(11,16)+")?\n\n"+
    "Записей затронуто: "+всего+". Изменённые вернутся к прежнему виду, добавленные этой правкой — будут удалены."+
    (поздние.length?("\n\nПосле загрузки поправлены ещё раз — их НЕ трогаю, чтобы не потерять новую правку ("+
      поздние.length+"): "+поздние.slice(0,8).join("; ")+(поздние.length>8?" …":"")):"")+
    "\n\nОтмена уйдёт и на другие места при следующем обмене."+
    (о.место!==getWS().код?("\n\nВнимание: правка сделана на месте "+о.место+"."):""))) return;
  const ts=new Date().toISOString(), я=getWS().код;
  const штамп=(k,r)=>{ r.изменено=ts; if(k===REG_ЗАД) r.место_правки=я; else r.место=я; };
  let сделано=0;
  о.части.forEach(ч=>{
    const arr=реестрПоКлючу(ч.k); if(!arr) return;
    ч.старые.forEach(ст=>{
      const r=найтиПоИд(arr,ст.id); if(!r || lastTs(r)!==ч.после[String(ст.id)]) return;
      Object.keys(r).forEach(k=>delete r[k]); Object.assign(r, ст);
      if(r.удалено) r.удалено=ts;           // было удалено — остаётся удалённым, но свежее
      штамп(ч.k,r); сделано++;
    });
    ч.новые.forEach(ид=>{
      const r=найтиПоИд(arr,ид); if(!r || lastTs(r)!==ч.после[ид]) return;
      r.удалено=ts; штамп(ч.k,r); сделано++;
    });
    if(ч.k===STORE_KEY) saveJournal(); else saveReg(ч.k,arr);
    if(ч.k===REG_SI && синхронизироватьЧасти()) saveReg(REG_SI,S);
    if(ч.k===REG_OBJ) rebuildObjIndex();
  });
  отменаСохранить(null);
  setChrome();
  switchView(ВИДЫ_С_ФОРМОЙ.indexOf(curView)>=0?(о.вид||"dash"):curView);
  toast("Отменено: "+сделано+(поздние.length?(", не тронуто (правлены позже): "+поздние.length):""));
  if(о.части.some(ч=>ч.k===STORE_KEY)){
    const столк=проверитьСтолкновения();
    if(столк.length) alert("ВНИМАНИЕ: одинаковые номера у разных записей.\n\n"+столк.slice(0,15).join("\n"));
  }
}
const базаКакВремя=б=>б===БАЗА_ИСХОДНАЯ?"":б;
function конфликтПравок(key,L,I){
  if(!key || key===REG_КОНФЛ || !L || !I) return false;
  if(lastTs(L)===lastTs(I)) return false;
  if(местоПравки(key,L)!==getWS().код) return false;   // своей правки нет — просто приходит свежее
  if(!L.база) return false;       // правлено старой версией программы — судить не по чему
  if(lastTs(I)===базаКакВремя(L.база)) return false;   // пришла та версия, от которой я правил
  if(I.база && базаКакВремя(I.база)===lastTs(L)) return false;   // пришла правка поверх моей
  if(местоПравки(key,I)===местоПравки(key,L)) return false;
  return true;
}
let _новыхКонфликтов=0;
function записатьКонфликт(key,L,I){
  const поб=lastTs(I)>lastTs(L)?I:L, проиг=поб===I?L:I;
  const уже=КОНФЛ.some(к=>к.ключ===key && String(к.запись)===String(L.id) &&
    к.проигравшая && lastTs(к.проигравшая)===lastTs(проиг));
  if(уже) return;
  КОНФЛ.push({id:newId("КФ"), ключ:key, запись:L.id, когда:new Date().toISOString(),
    победила:JSON.parse(JSON.stringify(поб)), проигравшая:JSON.parse(JSON.stringify(проиг)),
    решено:"", место:getWS().код});
  _новыхКонфликтов++;
  saveReg(REG_КОНФЛ,КОНФЛ);
}
function конфликтыОткрытые(){ return (КОНФЛ||[]).filter(к=>!к.решено); }
const СЛУЖЕБНЫЕ_ДЛЯ_СРАВНЕНИЯ=["id","изменено","удалено","место","место_правки","база"];
function подписьПоля(key,k){
  const рег=Object.values(REGS).find(r=>r.key===key);
  const f=рег && (рег.fields||[]).find(x=>x.k===k);
  return f ? f.l : k;
}
function значениеДляСравнения(v){
  if(v==null || v==="") return "—";
  if(typeof v==="object") return JSON.stringify(v).slice(0,160);
  return String(v);
}
function разницаВерсий(к){
  const a=к.победила||{}, b=к.проигравшая||{}, из=[];
  new Set(Object.keys(a).concat(Object.keys(b))).forEach(k=>{
    if(СЛУЖЕБНЫЕ_ДЛЯ_СРАВНЕНИЯ.includes(k)) return;
    const x=значениеДляСравнения(a[k]), y=значениеДляСравнения(b[k]);
    if(x!==y) из.push({поле:подписьПоля(к.ключ,k), осталось:x, отброшено:y});
  });
  if(!!a.удалено!==!!b.удалено) из.push({поле:"(запись)", осталось:a.удалено?"удалена":"есть", отброшено:b.удалено?"удалена":"есть"});
  return из;
}
function разделКонфликта(к){
  const р=РАЗДЕЛЫ_ИЗМЕНЕНИЙ.find(р=>{ try{ return р.arr()===реестрПоКлючу(к.ключ); }catch(e){ return false; } });
  const x=к.победила||к.проигравшая||{};
  return {н:р?р.н:(КЛЮЧИ_ДАННЫХ[к.ключ]||к.ключ), вид:р?р.вид:"", подпись:(р?р.п(x):"")||String(к.запись)};
}
function конфликтОставить(ид){
  const к=КОНФЛ.find(x=>x.id===ид); if(!к) return;
  к.решено="оставлена свежая версия"; к.решено_когда=new Date().toISOString();
  saveReg(REG_КОНФЛ,КОНФЛ); renderExchange();
}
function конфликтВернуть(ид){
  const к=КОНФЛ.find(x=>x.id===ид); if(!к) return;
  const arr=реестрПоКлючу(к.ключ);
  if(!arr){ alert("Раздел этой записи не найден — вернуть не во что."); return; }
  const i=arr.findIndex(z=>String(z.id)===String(к.запись));
  if(!confirm("Вернуть отброшенную версию записи «"+разделКонфликта(к).подпись+"»?\n\n"+
    "Она станет самой свежей и при следующем обмене уйдёт на остальные места. "+
    "Та версия, что сейчас в базе, сохранится здесь же как отброшенная — передумаете, вернёте её.")) return;
  const ts=new Date().toISOString(), я=getWS().код;
  const тек=i>=0?JSON.parse(JSON.stringify(arr[i])):null;
  const нов=JSON.parse(JSON.stringify(к.проигравшая));
  delete нов.база;
  нов.изменено=ts;
  if(к.ключ===REG_ЗАД) нов.место_правки=я; else нов.место=я;
  if(i>=0) arr[i]=нов; else arr.push(нов);
  к.решено="возвращена отброшенная версия"; к.решено_когда=ts;
  if(тек) КОНФЛ.push({id:newId("КФ"), ключ:к.ключ, запись:к.запись, когда:ts,
    победила:JSON.parse(JSON.stringify(нов)), проигравшая:тек, решено:"", место:я, возврат:true});
  if(к.ключ===STORE_KEY) saveJournal(); else saveReg(к.ключ,arr);
  if(к.ключ===REG_OBJ) rebuildObjIndex();
  if(к.ключ===REG_SI && синхронизироватьЧасти()) saveReg(REG_SI,S);
  saveReg(REG_КОНФЛ,КОНФЛ);
  renderExchange();
  toast("Отброшенная версия возвращена");
}
function блокКонфликтов(){
  const откр=конфликтыОткрытые();
  const решено=(КОНФЛ||[]).length-откр.length;
  if(!откр.length) return `<div class="panel"><h3>⇄ Правки одной записи с двух мест</h3>
    <div class="note" style="margin-top:0">Неразобранных нет${решено?(" (разобрано раньше: "+решено+")"):""}.
      Если одну запись между обменами поправят на двух местах, при приёме файла здесь появятся
      обе версии — останется более свежая, отброшенную можно будет вернуть.</div></div>`;
  return `<div class="panel" style="border-color:var(--warn)">
    <h3>⇄ Правки одной записи с двух мест — ${откр.length}</h3>
    <div class="note" style="margin-top:0">Одну и ту же запись между обменами правили на разных местах.
      По правилу обмена в базе осталась более свежая версия, а более ранняя — отброшена. Сравните и
      решите: оставить как есть или вернуть отброшенную${решено?(" · разобрано раньше: "+решено):""}.</div>
    ${откр.map(к=>{
      const р=разделКонфликта(к), рз=разницаВерсий(к), п=к.победила||{}, о=к.проигравшая||{};
      return `<div class="card" style="padding:10px;margin-top:8px">
        <div><b>${esc(р.н)}</b> · ${esc(р.подпись)}</div>
        <div class="note" style="margin:2px 0 6px">Осталась: ${esc(местоПравки(к.ключ,п)||"—")}, ${ru(lastTs(п).slice(0,10))} ${esc(lastTs(п).slice(11,16))}
          · отброшена: ${esc(местоПравки(к.ключ,о)||"—")}, ${ru(lastTs(о).slice(0,10))} ${esc(lastTs(о).slice(11,16))}</div>
        ${рз.length?`<div class="tblwrap"><table><thead><tr><th>Поле</th><th>Осталось в базе</th><th>Отброшено</th></tr></thead>
          <tbody>${рз.slice(0,12).map(d=>`<tr><td>${esc(d.поле)}</td><td>${esc(d.осталось)}</td><td>${esc(d.отброшено)}</td></tr>`).join("")}</tbody></table></div>`
          :`<div class="note">Содержимое версий совпадает — различаются только отметки времени.</div>`}
        <div class="formbtns" style="margin-top:8px">
          <button class="btn" onclick="конфликтОставить('${esc(к.id)}')">Оставить как есть</button>
          <button class="btn" onclick="конфликтВернуть('${esc(к.id)}')">Вернуть отброшенную</button>
        </div></div>`;
    }).join("")}</div>`;
}

// Два места могли выдать один номер, если оба брали право порознь.
// Молчать об этом нельзя: номер уже мог уйти на подписанный документ.
function проверитьСтолкновения(){
  const карта={}, из=[];
  liveJ().forEach(z=>{ const г=годЗаписи(z); if(!г) return;
    methodsOf(z).forEach(м=>{ const н=номерМетода(z,м); if(!/^\d+$/.test(н)) return;
      const ключ=м+"-"+г+"-"+(+н); (карта[ключ]=карта[ключ]||[]).push(z); });
    const m=/^АКТ\s+(\d+)-(\d{4})$/.exec(String(z.акт||"").trim());
    if(m){ const ключ="АКТ-"+m[2]+"-"+(+m[1]); (карта[ключ]=карта[ключ]||[]).push(z); }
  });
  Object.keys(карта).forEach(k=>{
    const сп=карта[k];
    // Акт законно повторяется: один акт охватывает несколько объектов.
    if(k.indexOf("АКТ-")===0) return;
    if(сп.length>1) из.push(k.replace(/^(\S+?)-(\d{4})-(\d+)$/,"$3-$2-$1")+
      ": " + сп.map(z=>(z.объект||"без объекта").slice(0,28)).join(" / "));
  });
  return из;
}
function exchangeExport(){
  const ws=getWS();
  const payload=Object.assign({ тип:"обмен ЛНК", формат:2, место:ws.код, мастер:ws.мастер,
    версия:DB.версия, выгружено:new Date().toISOString(), право:getGrant() }, всеДанные());
  const name="Обмен ЛНК "+new Date().toISOString().slice(0,10)+" ("+ws.код+").json";
  dl(name, JSON.stringify(payload));
  try{ ХР.setItem("lnk_last_export", new Date().toISOString()); }catch(e){}
  toast("Выгружено: журнал "+liveJ().length+", объекты "+live(O).length+", СИ "+live(S).length+", спец. "+live(P).length);
  if(curView==="exchange") renderExchange();
}
let _exIn;
function exchangeImport(){
  if(!_exIn){
    _exIn=document.createElement("input"); _exIn.type="file"; _exIn.accept=".json"; _exIn.multiple=true; _exIn.style.display="none"; document.body.appendChild(_exIn);
    _exIn.addEventListener("change",e=>{
      const files=[...e.target.files]; if(!files.length) return;
      _новыхКонфликтов=0;
      let done=0, tot={j:0,o:0,s:0,p:0}, отчёт=[];
      files.forEach(f=>{
        const rd=new FileReader();
        rd.onload=()=>{
          try{
            const d=JSON.parse(rd.result);
            const jarr=Array.isArray(d)?d:d.записи;
            if(!Array.isArray(jarr) && !d.объекты && !d.приборы && !d.специалисты)
              throw new Error("не файл обмена");
            отметитьПриёмОтМеста(d.место, d.выгружено, f.name);
            if(d.место) _местаИзФайлов.add(String(d.место));
            const rj=Array.isArray(jarr)?mergeJournal(jarr):{added:0,updated:0,skipped:0};
            // Право выдачи — свежее побеждает, как и всё остальное.
            if(d.право && d.право.изменено && d.право.изменено > getGrant().изменено){
              saveGrant({место:d.право.место||"", изменено:d.право.изменено});
            }
            const слито=применитьДанные(d, false);
            const ro=слито.o, rs=слито.s, rp=слито.p;
            // Карточка лаборатории — свежее побеждает, как право и названия.
            if(d.лаборатория && d.лаборатория.изменено &&
               d.лаборатория.изменено > (getЛаб().изменено||""))
              saveЛаб(d.лаборатория);
            // Методики контроля — свежее побеждает, тем же приёмом, что и
            // карточка лаборатории.
            if(d.методики && d.методики.изменено &&
               d.методики.изменено > (getМетодики().изменено||""))
              saveМетодики(d.методики);
            // Названия документов по методам — свежее побеждает.
            if(d.названия && d.названия.изменено && d.названия.изменено > имена().изменено)
              saveИмена({имена:Object.assign({}, имена().имена, d.названия.имена||{}),
                         ед:Object.assign({}, имена().ед, d.названия.ед||{}),
                         изменено:d.названия.изменено});
            if(ro.added||ro.updated) rebuildObjIndex();
            tot.j+=rj.added+rj.updated; tot.o+=ro.added+ro.updated; tot.s+=rs.added+rs.updated; tot.p+=rp.added+rp.updated;
            отчёт.push((d.место?("["+d.место+"] "):"")+f.name+
              ": журнал +"+rj.added+"/"+rj.updated+", объекты +"+ro.added+"/"+ro.updated+
              ", СИ +"+rs.added+"/"+rs.updated+", спец. +"+rp.added+"/"+rp.updated);
          }catch(err){ отчёт.push(f.name+": ошибка — "+err.message); }
          if(++done===files.length){
            try{ ХР.setItem("lnk_last_import", new Date().toISOString()); }catch(e){}
            обновитьОбменИлиМастер();
            const столк=проверитьСтолкновения();
            if(столк.length) alert("ВНИМАНИЕ: одинаковые номера у разных записей.\n\n"+
              столк.slice(0,12).join("\n")+(столк.length>12?("\n… всего "+столк.length):"")+
              "\n\nТакое бывает, если два места выдавали номера порознь. Исправьте вручную: "+
              "откройте записи и задайте верные номера.");
            alert("Принято файлов: "+files.length+"\n\n"+отчёт.join("\n")+
              "\n\nВсего изменений: журнал "+tot.j+", объекты "+tot.o+", приборы "+tot.s+", специалисты "+tot.p+
              ".\n(в записи «+новых/обновлено»)"+
              (_новыхКонфликтов?("\n\nОдни и те же записи правили на двух местах: "+_новыхКонфликтов+
                ". Осталась более свежая версия, отброшенная сохранена — сравнить и при "+
                "необходимости вернуть можно здесь же, в блоке «Правки одной записи с двух мест»."):""));
            toast("Обмен принят");
          }
        };
        rd.readAsText(f);
      });
      e.target.value="";
    });
  }
  _exIn.click();
}
function saveWSfromForm(){
  const код=(document.getElementById("wsCode").value||"").trim()||"М1";
  const мастер=document.getElementById("wsMaster").checked;
  const нач=document.getElementById("wsChief").checked;
  saveWS({код:код, мастер:мастер, начальник:нач});
  toast("Рабочее место: "+код+(мастер?" (мастер)":"")+(нач?" (начальник)":""));
  renderExchange();
}
