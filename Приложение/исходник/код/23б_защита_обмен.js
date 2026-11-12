// ── Защита настроек ПИН-кодом (3.77, п. 10 списка задач, вариант «б») ──
// ПИН один на все места: хранится как ДАННЫЕ (`lnk_pin`) и приходит с
// обменом и копией — свежее побеждает. В памяти лежит не сам ПИН, а его
// отпечаток SHA-256 с солью. Честно: программа — один файл без сервера,
// ПИН защищает от случайной правки не на своём месте, а не от человека,
// который захочет его обойти (это сказано и на экране).
//
// Что закрыто ПИНом, выбирают сами (флажки в «Обмене»). Проверка стоит не в
// каждой функции, а обёрткой вокруг функций сохранения (`подПИН`) — список
// ниже; с ПИНом спрашивается окно, без ПИНа функции работают как раньше.
// Открытые настройки остаются открытыми 30 минут с последнего действия
// под ПИНом или до «Закрыть сейчас» / перезагрузки.
const REG_ПИН="lnk_pin";
const ПИН_МИНУТ=30;
const ПИН_ГРУППЫ=[
  ["лаб","Карточка лаборатории (сохранение, новое свидетельство)",1,["saveЛабИзФормы","новоеСвидетельство"]],
  ["право","Право выдачи номеров (принять, передать)",1,["принятьПраво","передатьПраво"]],
  ["место","Код рабочего места и роли (начальник, мастер)",1,["saveWSfromForm"]],
  ["методики","Методики контроля (стандартные наборы)",1,["saveMetodika","сохранитьСтандартМетода"]],
  ["год","Закрытие и открытие года журнала",1,["закрытьГод","открытьГод"]],
  ["копия","Восстановление из копии, переезд, правки старой версии",1,["восстановитьИзДанных","переездПринять","принятьПравкиСтаройВерсии"]],
  ["импорт","Загрузка из Excel (все разделы)",1,["импортПрименить","импортЖурналПрименить","импортЗадачПрименить",
    "импортПланПрименить","импортПланаПозицийПрименить","импортПравкиПГ","импортОтметокПГ","импортПГПрименить"]],
  ["план","План ППР и планы-графики: позиции (не отметки о выполнении)",0,["saveППР","delППР","ппрКакВКниге","saveПГП","delПГП","saveПГ","delПГ"]],
  ["спр:obj","Справочник «Объекты»",0,["quickAddObject"]],
  ["спр:si","Справочник «Приборы»",0,["частиСоздать"]],
  ["спр:ppl","Справочник «Специалисты»",1],
  ["спр:podr","Справочник «Подразделения»",1],
  ["спр:nd","Справочник «Норм. документы»",1],
  ["спр:pom","Справочник «Помещения»",1],
  ["спр:vk","Карточки учёта вибрации",0],
  ["спр:vn","Нормы вибрации",1]];
const ПИН_ПО_УМОЛЧАНИЮ=()=>ПИН_ГРУППЫ.filter(г=>г[2]).map(г=>г[0]);
function пин(){
  let v=null; try{ v=JSON.parse(ХР.getItem(REG_ПИН)||"null"); }catch(e){}
  return Object.assign({хеш:"", соль:"", что:ПИН_ПО_УМОЛЧАНИЮ(), изменено:"", место:"", сброшен:null}, v||{});
}
function сохранитьПин(п, какЕсть){
  if(!какЕсть){ п.изменено=new Date().toISOString(); п.место=getWS().код; }
  try{ ХР.setItem(REG_ПИН, JSON.stringify(п)); }catch(e){ переполнение(e); }
}
function пинЕсть(){ return !!пин().хеш; }
let _пинДо=0;
function пинОткрыт(){ return Date.now()<_пинДо; }
function пинЗакрыть(){ _пинДо=0; if(curView==="exchange") renderExchange(); toast("Настройки закрыты ПИН-кодом"); }
function пинНужен(ключ){
  const п=пин(); if(!п.хеш) return false;
  if(ключ!=="пин" && (п.что||[]).indexOf(ключ)<0) return false;
  return !пинОткрыт();
}
async function хешПИН(значение, соль){
  const текст=соль+":"+String(значение);
  try{
    const h=await crypto.subtle.digest("SHA-256", new TextEncoder().encode(текст));
    return [...new Uint8Array(h)].map(x=>x.toString(16).padStart(2,"0")).join("");
  }catch(e){ return "открыто:"+текст; }
}
async function пинВерен(значение){ const п=пин(); return !!п.хеш && (await хешПИН(значение,п.соль))===п.хеш; }
function пинПродлить(){ _пинДо=Date.now()+ПИН_МИНУТ*60000; }
// Пропустить действие или спросить ПИН. true — можно продолжать сразу;
// false — спрошено окно, действие выполнится само после верного ПИНа.
function черезПИН(ключ, дальше){
  if(!пинНужен(ключ)){ if(пинЕсть() && пинОткрыт()) пинПродлить(); return true; }
  const г=ПИН_ГРУППЫ.find(x=>x[0]===ключ);
  спроситьПИН(г?г[1]:"управление ПИН-кодом").then(ок=>{ if(ок) дальше(); });
  return false;
}
function подПИН(ключ, имя){
  const f=window[имя]; if(typeof f!=="function" || f.__подПИН) return false;
  const обёртка=function(){ const а=arguments, я=this;
    const к=typeof ключ==="function"?ключ.apply(null,а):ключ;
    if(!черезПИН(к, ()=>f.apply(я,а))) return; return f.apply(я,а); };
  обёртка.__подПИН=true; window[имя]=обёртка; return true;
}
// Окно ввода ПИНа (свои поля, а не prompt — prompt показывает ПИН открытым).
let _пинОкно=null;
function спроситьПИН(что){
  if(_пинОкно) _пинОкно.отмена();
  return new Promise(готово=>{
    const фон=document.createElement("div"); фон.className="пинфон";
    фон.innerHTML=`<div class="пинокно" role="dialog" aria-label="ПИН-код">
      <h3>🔒 Нужен ПИН-код</h3>
      <div class="note" style="margin-top:0">Закрыто ПИНом: ${esc(что)}.</div>
      <input type="password" id="пинВвод" autocomplete="off" placeholder="ПИН-код">
      <div id="пинОш" class="пинош"></div>
      <div class="formbtns"><button class="btn primary" id="пинОк">Открыть</button>
        <button class="btn" id="пинОтм">Отмена</button></div>
      <div class="note"><a href="#" id="пинЗабыл">Забыли ПИН?</a> · открытые настройки сами закроются через ${ПИН_МИНУТ} мин.</div></div>`;
    document.body.appendChild(фон);
    const ввод=фон.querySelector("#пинВвод");
    const закрыть=(ок)=>{ фон.remove(); _пинОкно=null; готово(ок); };
    const проверить=async()=>{
      if(await пинВерен(ввод.value)){ пинПродлить(); закрыть(true); if(curView==="exchange") renderExchange(); }
      else{ фон.querySelector("#пинОш").textContent="Неверный ПИН-код"; ввод.select(); }
    };
    фон.querySelector("#пинОк").onclick=проверить;
    фон.querySelector("#пинОтм").onclick=()=>закрыть(false);
    фон.querySelector("#пинЗабыл").onclick=(e)=>{ e.preventDefault(); закрыть(false); пинСбросить(); };
    ввод.addEventListener("keydown",e=>{ if(e.key==="Enter"){ e.preventDefault(); проверить(); }
      else if(e.key==="Escape"){ e.preventDefault(); закрыть(false); } });
    _пинОкно={отмена:()=>закрыть(false), ввести:async(v)=>{ ввод.value=v; await проверить(); }};
    setTimeout(()=>ввод.focus(),0);
  });
}
async function пинУстановить(значение, повтор, что){
  if(пинЕсть()){ toast("ПИН уже установлен — его можно сменить"); return false; }
  return пинЗаписать(значение, повтор, что, "установлен");
}
async function пинЗаписать(значение, повтор, что, слово){
  значение=String(значение||""); повтор=String(повтор==null?значение:повтор);
  if(значение.length<4){ alert("ПИН-код — не меньше 4 знаков."); return false; }
  if(значение!==повтор){ alert("ПИН-код и повтор не совпадают."); return false; }
  const соль=[...crypto.getRandomValues(new Uint8Array(12))].map(x=>x.toString(16).padStart(2,"0")).join("");
  const п=пин();
  п.соль=соль; п.хеш=await хешПИН(значение, соль); п.сброшен=null;
  if(Array.isArray(что)) п.что=что;
  сохранитьПин(п); пинПродлить();
  if(curView==="exchange") renderExchange();
  toast("ПИН-код "+слово+" — уйдёт на остальные места с обменом");
  return true;
}
function пинФорма(){
  const з=id=>(document.getElementById(id)||{}).value||"";
  const что=[...document.querySelectorAll(".пингр:checked")].map(x=>x.value);
  if(!пинЕсть()) return пинУстановить(з("пинНов"), з("пинПовт"), что);
  if(!черезПИН("пин", ()=>пинФорма())) return;
  return пинЗаписать(з("пинНов"), з("пинПовт"), null, "сменён");
}
function пинСписок(){
  if(!черезПИН("пин", ()=>пинСписок())) return;
  const п=пин();
  п.что=[...document.querySelectorAll(".пингр:checked")].map(x=>x.value);
  сохранитьПин(п); renderExchange(); toast("Список закрытого ПИНом сохранён");
}
function пинСнять(){
  if(!черезПИН("пин", ()=>пинСнять())) return;
  if(!confirm("Снять ПИН-код? Настройки станут доступны на всех местах после обмена.")) return;
  const п=пин(); п.хеш=""; п.соль=""; п.сброшен=null; сохранитьПин(п); _пинДо=0;
  renderExchange(); toast("ПИН-код снят");
}
// Забыли ПИН: сброс возможен, но не тихий — запись «сброшен на месте … »
// уходит с обменом на все места и висит в «Проверке данных», пока не
// установят новый.
function пинСбросить(){
  const ответ=prompt("Сбросить ПИН-код?\n\nЗащита снимется на ВСЕХ местах после обмена, а в «Проверке данных» "+
    "будет видно, на каком месте и когда ПИН сброшен.\n\nДля сброса впишите слово СБРОСИТЬ:","");
  if(String(ответ||"").trim().toUpperCase()!=="СБРОСИТЬ") return;
  const п=пин(); п.хеш=""; п.соль=""; п.сброшен={место:getWS().код, когда:new Date().toISOString()};
  сохранитьПин(п); _пинДо=0;
  if(curView==="exchange") renderExchange();
  alert("ПИН-код сброшен. Установите новый в «Обмене файлами» → «Защита настроек».");
}
function принятьПин(вход){
  if(!вход || typeof вход!=="object") return false;
  if(String(вход.изменено||"")>String(пин().изменено||"")){ сохранитьПин(JSON.parse(JSON.stringify(вход)), true); return true; }
  return false;
}
function блокПИН(){
  const п=пин(), есть=!!п.хеш;
  const флажки=ПИН_ГРУППЫ.map(г=>`<label class="пинстр"><input type="checkbox" class="пингр" value="${esc(г[0])}"
      ${(п.что||[]).indexOf(г[0])>=0?"checked":""}> ${esc(г[1])}</label>`).join("");
  return `<div class="panel">
    <h3>🔒 Защита настроек ПИН-кодом</h3>
    ${есть?`<div class="bar"><span class="g" style="flex:1">ПИН-код</span><b>установлен</b></div>
      <div class="note" style="margin-top:0">Установлен или изменён на месте ${esc(п.место||"—")} ${п.изменено?ru(п.изменено.slice(0,10)):""}.
        ${пинОткрыт()?`Сейчас настройки <b>открыты</b> до ${new Date(_пинДо).toTimeString().slice(0,5)}.
          <button class="btn small" onclick="пинЗакрыть()">🔒 Закрыть сейчас</button>`:"Сейчас настройки закрыты."}</div>`
     :`<div class="bar"><span class="g" style="flex:1">ПИН-код</span><b>не установлен</b></div>
      ${п.сброшен?`<div class="note" style="color:var(--warn)">ПИН был сброшен на месте ${esc(п.сброшен.место||"—")}
        ${ru(String(п.сброшен.когда||"").slice(0,10))} — установите новый.</div>`:""}`}
    <div class="sect">Что закрыто ПИНом</div>
    <div class="пинсписок">${флажки}</div>
    ${есть?`<div class="formbtns"><button class="btn" onclick="пинСписок()">Сохранить список</button></div>`:""}
    <div class="sect">${есть?"Сменить ПИН-код":"Установить ПИН-код"}</div>
    <div class="form"><div class="grid">
      <label>${есть?"Новый ПИН":"ПИН-код"}<input type="password" id="пинНов" autocomplete="new-password"></label>
      <label>Ещё раз<input type="password" id="пинПовт" autocomplete="new-password"></label></div>
      <div class="formbtns"><button class="btn primary" onclick="пинФорма()">${есть?"Сменить":"Установить"}</button>
        ${есть?`<button class="btn" onclick="пинСнять()">Снять ПИН</button>`:""}</div></div>
    <div class="note">ПИН один на все места: он уходит с обменом и в резервную копию (в базе лежит не сам ПИН, а
      его отпечаток). Защищает от случайной правки не на своём месте — человек, который захочет обойти
      запрет, сможет это сделать: программа — один файл без сервера.</div>
  </div>`;
}

// ── Обмен одной кнопкой и копии в несколько мест (3.77, п. 10) ──────
// «▼ Сохранить на флешку» кладёт в выбранную папку файл обмена, резервную
// копию и недостающие у других мест файлы документов. «▲ Принять» берёт из
// выбранной папки файлы обмена других мест, недостающие документы и
// замечает там новую версию программы. Папка запоминается — при запуске
// программа сама заглядывает в неё (если браузер уже дал разрешение) и
// предлагает принять новое.
let _папкаОбмена=null;
async function папкаОбмена(){ if(!_папкаОбмена) _папкаОбмена=await идбЧтение("папкаОбмена"); return _папкаОбмена; }
async function запомнитьПапкуОбмена(h){ _папкаОбмена=h; try{ await идбЗапись("папкаОбмена",h); }catch(e){} }
function имяКопииДня(){ return "Резервная копия ЛНК "+new Date().toISOString().slice(0,10)+" ("+getWS().код+").json"; }
function данныеКопииДня(){
  return JSON.stringify(Object.assign({ тип:"резервная копия ЛНК", формат:2, место:getWS().код,
    версия:DB.версия, сохранено:new Date().toISOString() }, всеДанные()), null, 1);
}
async function записатьТекст(папка, имя, текст){
  const ф=await папка.getFileHandle(имя,{create:true}); const w=await ф.createWritable();
  await w.write(new Blob([текст],{type:"application/json"})); await w.close();
}
async function сохранитьНаФлешку(){
  if(!ЕСТЬ_ПАПКИ){
    exchangeExport(); backupSave();
    alert("Этот браузер не умеет писать в выбранную папку, поэтому файл обмена и резервная копия скачаны "+
      "в папку загрузок — перенесите их на флешку. Файлы документов в этом браузере копируются вручную.");
    return;
  }
  const цель=await выбратьФлешку(true); if(!цель) return;
  await запомнитьПапкуОбмена(цель);
  const флВсе=document.getElementById("флВсе");
  const текст=await выгрузитьВПапку(цель, {копия:true, все:!!(флВсе&&флВсе.checked)});
  const старая=await найтиВерсиюПрограммы(цель);
  alert(текст+(старая && верс(старая.версия)<верс(DB.версия)?("\n\nВ папке лежит программа версии "+старая.версия+
    " («"+старая.имя+"»), а у вас "+DB.версия+". Чтобы обновить другие места, положите туда файл программы этой версии."):""));
}
async function принятьСФлешки(){
  if(!ЕСТЬ_ПАПКИ){ принятьФайлы(); return; }
  return обменСФайламиПринять();
}
// ── Версия программы в папке и обновление файла программы ─────────
// В сборке у страницы есть метка <meta name="lnk-version"> (build_app.py) —
// по ней в папке на флешке находится файл программы новее этой. Обновление
// копирует его поверх файла программы на этом компьютере, а прежний
// оставляет рядом «… (прежняя).html». Данные лежат в браузере, а не в файле
// программы, поэтому замена файла их не трогает (проверено: хранилище
// у страниц file:// общее).
function верс(v){ return String(v||"0").split(".").map(x=>parseInt(x,10)||0).reduce((a,x,i)=>a+x/Math.pow(1000,i),0); }
function версияИзТекста(текст){
  const m=/<meta name="lnk-version" content="([^"]+)"(?: data-built="([^"]*)")?/.exec(текст||"");
  return m?{версия:m[1], собрано:m[2]||""}:null;
}
async function найтиВерсиюПрограммы(папка){
  let лучшая=null;
  try{
    for await (const [имя,х] of папка.entries()){
      if(х.kind!=="file" || !/\.html?$/i.test(имя)) continue;
      try{ const ф=await х.getFile(); const в=версияИзТекста(await ф.slice(0,4096).text());
        if(в && (!лучшая || верс(в.версия)>верс(лучшая.версия))) лучшая=Object.assign({имя, h:х}, в); }catch(e){}
    }
  }catch(e){}
  return лучшая;
}
function перезагрузитьПрограмму(){ location.reload(); }
function имяФайлаПрограммы(){ try{ return decodeURIComponent(location.pathname.split("/").pop()||""); }catch(e){ return ""; } }
function путьПрограммы(){
  try{ const п=decodeURIComponent(location.pathname);
    return /^\/[A-Za-z]:/.test(п)?п.slice(1).replace(/\//g,"\\"):п; }catch(e){ return location.href; } }
async function естьФайлВ(папка, имя){ try{ await папка.getFileHandle(имя); return true; }catch(e){ return false; } }
async function обновитьПрограмму(новый){
  const имя=имяФайлаПрограммы();
  if(!/\.html?$/i.test(имя)){ alert("Не понять, как называется файл программы. Замените его вручную."); return false; }
  if(!новый){ toast("Нового файла программы нет"); return false; }
  if(верс(новый.версия)<=верс(DB.версия)){ toast("Это не новее вашей версии "+DB.версия); return false; }
  let п=await идбЧтение("папкаПрограммы");
  let право="нет";
  if(п){ try{ право=await п.queryPermission({mode:"readwrite"}); if(право!=="granted") право=await п.requestPermission({mode:"readwrite"}); }catch(e){ право="нет"; } }
  if(!п || право!=="granted" || !(await естьФайлВ(п,имя))){
    alert("Выберите папку на этом компьютере, где лежит файл программы «"+имя+"».\n\nСейчас программа открыта отсюда:\n"+путьПрограммы());
    try{ п=await window.showDirectoryPicker({mode:"readwrite", id:"lnk-programma"}); }
    catch(e){ if(e.name!=="AbortError") toast("Не удалось выбрать папку: "+e.name); return false; }
    if(!(await естьФайлВ(п,имя))){ alert("В папке «"+п.name+"» нет файла «"+имя+"». Выберите папку, где он лежит."); return false; }
    await идбЗапись("папкаПрограммы",п);
  }
  if(!уходСФормы()) return false;
  try{ await ХР.дождаться(); }catch(e){}
  try{
    const данные=await (await новый.h.getFile()).arrayBuffer();
    const в=версияИзТекста(new TextDecoder().decode(данные.slice(0,4096)));
    if(!в || в.версия!==новый.версия){ alert("Файл на флешке не похож на программу версии "+новый.версия+" — не заменяю."); return false; }
    const старое=await (await (await п.getFileHandle(имя)).getFile()).arrayBuffer();
    const имяПрежней=имя.replace(/\.html?$/i,"")+" "+DB.версия+" (прежняя).html";
    const ф1=await п.getFileHandle(имяПрежней,{create:true}); const w1=await ф1.createWritable(); await w1.write(старое); await w1.close();
    const ф2=await п.getFileHandle(имя); const w2=await ф2.createWritable(); await w2.write(данные); await w2.close();
    const провер=версияИзТекста(new TextDecoder().decode((await (await ф2.getFile()).arrayBuffer()).slice(0,4096)));
    if(!провер || провер.версия!==новый.версия){ alert("Файл записан, но проверка версии не сошлась. Откройте «"+имяПрежней+"» и сообщите мне."); return false; }
    if(confirm("Программа на этом компьютере обновлена до версии "+новый.версия+".\nПрежняя сохранена рядом: «"+имяПрежней+"».\n\n"+
      "Данные не тронуты — они хранятся в браузере.\n\nОткрыть новую версию сейчас (перезагрузить страницу)?")) перезагрузитьПрограмму();
    return true;
  }catch(e){
    alert("Не удалось заменить файл программы: "+(e&&e.name||e)+"\n\nЗамените вручную: закройте вкладку, скопируйте «"+новый.имя+
      "» с флешки поверх «"+имя+"» и откройте снова.");
    return false;
  }
}

// ── Заглянуть в папку обмена сама ──────────────────────────
// При запуске и при возврате в окно (не чаще раза в минуту): если браузер
// уже разрешил читать запомненную папку обмена — найти там файлы обмена
// других мест новее принятых и программу новее этой. Разрешения нет —
// только кнопка «Проверить» в «Обмене» (без нажатия браузер не спросит).
let _вПапкеОбмена=null, _папкаОбменаКогда=0;
async function проверитьПапкуОбмена(спросить){
  const h=await папкаОбмена(); if(!h){ _вПапкеОбмена=null; return null; }
  let p="denied";
  try{ p=await h.queryPermission({mode:"read"}); if(p!=="granted" && спросить) p=await h.requestPermission({mode:"read"}); }catch(e){}
  if(p!=="granted"){ _вПапкеОбмена={папка:h.name, нужноРазрешение:true, новые:[]}; return _вПапкеОбмена; }
  const я=getWS().код, места=обменМеста().места||{}, лучш={};
  try{
    for await (const [имя,х] of h.entries()){
      if(х.kind!=="file" || !/\.json$/i.test(имя)) continue;
      let d; try{ d=JSON.parse(await (await х.getFile()).text()); }catch(e){ continue; }
      if(!d || d.тип!=="обмен ЛНК" || !d.место || d.место===я) continue;
      const был=(места[d.место]||{}).выгружен||"";
      if(String(d.выгружено||"")<=был) continue;
      if(!лучш[d.место] || String(d.выгружено)>String(лучш[d.место].выгружено)) лучш[d.место]={место:d.место, выгружено:d.выгружено||"", имя};
    }
  }catch(e){ _вПапкеОбмена={папка:h.name, нетДоступа:true, новые:[]}; return _вПапкеОбмена; }
  const пр=await найтиВерсиюПрограммы(h);
  _вПапкеОбмена={папка:h.name, новые:Object.values(лучш), программа:(пр && верс(пр.версия)>верс(DB.версия))?пр:null};
  return _вПапкеОбмена;
}
async function автоПроверкаПапкиОбмена(){
  if(!ЕСТЬ_ПАПКИ) return;
  _папкаОбменаКогда=Date.now();
  try{ await проверитьПапкуОбмена(false); }catch(e){}
  if(curView==="dash") renderDash(); else if(curView==="exchange") renderExchange();
}
window.addEventListener("focus",()=>{ if(window.ЛНК_ГОТОВ && Date.now()-_папкаОбменаКогда>60000) автоПроверкаПапкиОбмена(); });
async function проверитьПапкуОбменаКнопкой(){
  const р=await проверитьПапкуОбмена(true);
  if(!р){ toast("Папка обмена ещё не выбрана — нажмите «▲ Принять»"); return; }
  if(curView==="exchange") renderExchange(); else if(curView==="dash") renderDash();
  toast(р.нужноРазрешение?"Нет разрешения на папку":р.нетДоступа?"Папка недоступна (флешка не вставлена?)":
    (р.новые.length?("Новых файлов обмена: "+р.новые.length):"Нового в папке нет"));
}
async function принятьИзПапкиОбмена(){
  const h=await папкаОбмена(); if(!h){ принятьСФлешки(); return; }
  try{ let p=await h.queryPermission({mode:"read"}); if(p!=="granted") p=await h.requestPermission({mode:"read"});
    if(p!=="granted"){ toast("Нет разрешения на папку обмена"); return; } }catch(e){ toast("Папка недоступна"); return; }
  await обменСФайламиПринять(h);
  _вПапкеОбмена=null; автоПроверкаПапкиОбмена();
}
async function обновитьИзПапкиОбмена(){
  const р=_вПапкеОбмена||await проверитьПапкуОбмена(true);
  if(!р || !р.программа){ toast("Новой версии программы в папке обмена нет"); return; }
  await обновитьПрограмму(р.программа);
}
function плашкаПапкиОбмена(наПульте){
  const р=_вПапкеОбмена; if(!р) return "";
  const стр=[];
  if(р.новые.length) стр.push(`<div class="draftnote" style="margin-bottom:10px">⇄ В папке обмена «${esc(р.папка)}» новые файлы:
    ${р.новые.map(x=>esc(x.место)+" (выгружен "+ru(String(x.выгружено).slice(0,10))+" "+esc(String(x.выгружено).slice(11,16))+")").join(", ")}.
    <button class="btn small primary" onclick="принятьИзПапкиОбмена()">▲ Принять</button></div>`);
  if(р.программа) стр.push(`<div class="draftnote" style="margin-bottom:10px">⬆ В папке обмена «${esc(р.папка)}» новая версия программы
    <b>${esc(р.программа.версия)}</b> («${esc(р.программа.имя)}»), у вас ${esc(DB.версия)}.
    <button class="btn small primary" onclick="обновитьИзПапкиОбмена()">Обновить программу</button></div>`);
  if(!наПульте && р.нужноРазрешение) стр.push(`<div class="note">Папка обмена «${esc(р.папка)}»: браузер не дал разрешения
    читать её без нажатия. <button class="btn small" onclick="проверитьПапкуОбменаКнопкой()">Проверить папку</button></div>`);
  return стр.join("");
}

// ── Копии сразу в несколько папок (до трёх: ПК, флешка, сеть) ───────
const СЛОТЫ_КОПИЙ=["папкаКопий","папкаКопий2","папкаКопий3"];
const ПОДПИСИ_СЛОТОВ=["Папка 1 (например, на этом ПК)","Папка 2 (например, флешка)","Папка 3 (например, сетевая)"];
const _папкиСлотов={};
async function папкаКопийСлот(i){
  if(i===0) return папкаКопий();
  if(!_папкиСлотов[i]) _папкиСлотов[i]=await идбЧтение(СЛОТЫ_КОПИЙ[i]);
  return _папкиСлотов[i];
}
function копииСлотов(){ try{ return JSON.parse(ХР.getItem("lnk_copy_slots")||"{}")||{}; }catch(e){ return {}; } }
function отметитьСлот(i, ок, ош){
  const в=копииСлотов(); в[i]=Object.assign({}, в[i]||{}, ок?{когда:new Date().toISOString(), ошибка:""}:{ошибка:ош||"", попытка:new Date().toISOString()});
  try{ ХР.setItem("lnk_copy_slots", JSON.stringify(в)); }catch(e){}
}
async function выбратьПапкуКопийСлот(i){
  if(i===0) return выбратьПапкуКопий();
  if(!ЕСТЬ_ПАПКИ){ alert("Этот браузер не умеет писать копии в выбранную папку. Так умеют Chrome и Edge."); return; }
  try{
    const h=await window.showDirectoryPicker({mode:"readwrite", id:"lnk-kopii-"+i});
    _папкиСлотов[i]=h; await идбЗапись(СЛОТЫ_КОПИЙ[i],h);
    toast("Папка для копий выбрана"); await копияВПапку(false); обновитьОбменИлиМастер();
  }catch(e){ if(e.name!=="AbortError") toast("Не удалось выбрать папку: "+e.name); }
}
async function забытьПапкуСлот(i){
  if(i===0) return забытьПапку();
  _папкиСлотов[i]=null; await идбЗапись(СЛОТЫ_КОПИЙ[i],null);
  toast("Папка забыта"); renderExchange();
}
async function правоНаРучку(h, спросить){
  try{ let p=await h.queryPermission({mode:"readwrite"}); if(p==="granted") return "есть";
    if(спросить){ p=await h.requestPermission({mode:"readwrite"}); if(p==="granted") return "есть"; }
  }catch(e){} return "нужно разрешение";
}
// Запись копии во все выбранные папки: что недоступно — названо, но не
// мешает записи в остальные.
async function копииВоВсеПапки(тихо){
  const данные=данныеКопии(), имя=имяКопии(); let записано=0; const не=[];
  for(let i=0;i<СЛОТЫ_КОПИЙ.length;i++){
    const h=await папкаКопийСлот(i); if(!h) continue;
    if(await правоНаРучку(h,!тихо)!=="есть"){ отметитьСлот(i,false,"нет разрешения"); не.push(h.name+" — нет разрешения"); continue; }
    try{
      const ф=await h.getFileHandle(имя,{create:true}); const w=await ф.createWritable(); await w.write(данные); await w.close();
      await чиститьСтарые(h); отметитьСлот(i,true); записано++;
    }catch(e){ отметитьСлот(i,false,e.name||String(e)); не.push(h.name+" — "+(e.name||e)); }
  }
  if(записано) try{ ХР.setItem("lnk_last_backup", new Date().toISOString()); }catch(e){}
  if(!тихо) toast(записано?("Копия записана в папок: "+записано+(не.length?("; не записана: "+не.join("; ")):"")):
    (не.length?("Копия не записана: "+не.join("; ")):"Папки для копий не выбраны"));
  return записано>0;
}
function блокПапокКопий(){
  if(!ЕСТЬ_ПАПКИ) return `<div class="panel"><h3>⟳ Копии сами</h3>
    <div class="note">Этот браузер не умеет писать в выбранную папку — так умеют <b>Chrome и Edge</b>. Здесь остаётся
      «▼ Сохранить на флешку»: файлы уходят в папку загрузок, оттуда переносите на флешку.</div></div>`;
  const сл=копииСлотов();
  const ручки=[_папка, _папкиСлотов[1], _папкиСлотов[2]];
  return `<div class="panel"><h3>⟳ Копии сами — в несколько папок сразу</h3>
    <div class="note" style="margin-top:0">Программа сама пишет копию во все выбранные папки (не чаще раза в 10 минут
      работы) и держит в каждой последние ${ДЕРЖАТЬ_КОПИЙ} копий. Папка недоступна (флешку вынули, сети нет) — копия
      ложится в остальные, а здесь видно, где её нет.</div>
    ${СЛОТЫ_КОПИЙ.map((к,i)=>{ const h=ручки[i], с=сл[i]||{};
      return `<div class="bar"><span class="g" style="flex:1">${esc(ПОДПИСИ_СЛОТОВ[i])}</span>
        <b>${h?esc(h.name):"не выбрана"}</b></div>
        ${h?`<div class="note" style="margin-top:0">${с.когда?("последняя копия "+ru(с.когда.slice(0,10))+" "+с.когда.slice(11,16)):"копий ещё не было"}${
          с.ошибка?` · <span style="color:var(--danger)">не записана: ${esc(с.ошибка)}</span>`:""}</div>`:""}
        <div class="formbtns" style="margin-top:2px"><button class="btn small" onclick="выбратьПапкуКопийСлот(${i})">${h?"Другая папка…":"Указать папку…"}</button>
          ${h?`<button class="btn small" onclick="забытьПапкуСлот(${i})">Забыть</button>`:""}</div>`; }).join("")}
    <div class="formbtns"><button class="btn primary" onclick="копииВоВсеПапки(false)">Записать копию сейчас</button></div></div>`;
}
// Подгрузить ручки слотов 2–3 для экрана (первая — как раньше).
Promise.resolve().then(async()=>{ for(let i=1;i<СЛОТЫ_КОПИЙ.length;i++) await папкаКопийСлот(i); });

// ── «Принять файл…»: тип файла программа узнаёт сама ───────────
// Файл обмена — слияние; резервная копия — восстановление (замена базы, с
// вопросом); файл переезда — переезд места. Несколько файлов — только обмен.
let _прФ;
function принятьФайлы(){
  if(!_прФ){
    _прФ=document.createElement("input"); _прФ.type="file"; _прФ.accept=".json"; _прФ.multiple=true; _прФ.style.display="none";
    document.body.appendChild(_прФ);
    _прФ.addEventListener("change",async e=>{
      const files=[...e.target.files]; e.target.value=""; if(!files.length) return;
      const прочитано=[];
      for(const f of files){ try{ прочитано.push({f, d:JSON.parse(await f.text())}); }catch(err){ прочитано.push({f, ошибка:"не читается"}); } }
      const особые=прочитано.filter(x=>x.d && (x.d.тип===ТИП_ПЕРЕЕЗДА || x.d.тип==="резервная копия ЛНК"));
      if(особые.length){
        if(прочитано.length>1){ alert("Резервную копию и файл переезда принимают по одному, отдельно от файлов обмена."); return; }
        const d=особые[0].d;
        if(d.тип===ТИП_ПЕРЕЕЗДА) переездПринять(d); else восстановитьИзДанных(d);
        return;
      }
      const годные=[], плохие=[];
      прочитано.forEach(x=>{
        if(x.ошибка){ плохие.push(x); return; }
        const d=x.d, jarr=Array.isArray(d)?d:d.записи;
        if(!Array.isArray(jarr) && !d.объекты && !d.приборы && !d.специалисты) плохие.push(Object.assign(x,{ошибка:"не файл обмена"}));
        else годные.push(x);
      });
      if(!годные.length){ alert("Принять нечего:\n\n"+плохие.map(x=>x.f.name+": "+x.ошибка).join("\n")); return; }
      if(принятьПрочитанныеОбмены(годные, плохие)) автоСверкаФайлов();
    });
  }
  _прФ.click();
}
function блокГлавногоОбмена(){
  const last=(k)=>{ let t=""; try{t=ХР.getItem(k)||"";}catch(e){} return t?(ru(t.slice(0,10))+" "+t.slice(11,16)):"—"; };
  return `<div class="panel главобмен">
    <h3>⇄ Обмен с другими местами</h3>
    <div class="formbtns">
      <button class="btn primary big" onclick="сохранитьНаФлешку()" title="Файл обмена, резервная копия и недостающие файлы документов — в выбранную папку">▼ Сохранить на флешку…</button>
      <button class="btn primary big" onclick="принятьСФлешки()" title="Файлы обмена других мест, недостающие документы и новая версия программы — из выбранной папки">▲ Принять…</button>
      ${ЕСТЬ_ПАПКИ?`<button class="btn" onclick="принятьФайлы()" title="Один файл не с флешки: резервная копия, файл переезда места или файл обмена — программа узнает сама">Принять отдельный файл…</button>`:""}
    </div>
    ${ЕСТЬ_ПАПКИ?`<label class="флвсе" title="Обычно кладутся только документы, которых нет у других мест (это видно по их файлам обмена). Отметьте, если появилось новое место или есть сомнения — положатся все прикреплённые документы.">
      <input type="checkbox" id="флВсе"> класть все файлы документов, а не только недостающие у других мест</label>`:""}
    <div class="note">${ЕСТЬ_ПАПКИ
      ?`<b>«Сохранить»</b> кладёт в выбранную папку (флешка или общая папка) файл обмена, резервную копию и файлы документов,
        которых нет у других мест. <b>«Принять»</b> берёт из выбранной папки файлы обмена других мест (по самому свежему от
        каждого), недостающие документы и замечает там новую версию программы. Папка запоминается: при запуске программа
        сама смотрит, нет ли в ней нового.`
      :`Этот браузер не умеет работать с папкой: «Сохранить» скачает файл обмена и копию в папку загрузок, «Принять»
        откроет выбор файлов — программа сама узнает, обмен это, резервная копия или переезд места.`}</div>
    <div class="note">Последняя выгрузка: ${last("lnk_last_export")} · приём: ${last("lnk_last_import")} · копия: ${last("lnk_last_backup")}
      ${_папкаОбмена?` · папка обмена «${esc(_папкаОбмена.name)}»`:""}</div>
    ${плашкаПапкиОбмена(false)}
    <div class="note"><b>Как обновить программу:</b> положите новый файл программы на флешку рядом с обменом. На каждом
      месте «▲ Принять» заметит его и предложит «Обновить программу»: новый файл ляжет поверх старого на этом компьютере,
      прежний останется рядом с пометкой «(прежняя)». Данные не теряются — они в браузере, а не в файле.</div>
  </div>`;
}

// Поставить ПИН на функции из списка ПИН_ГРУППЫ — один раз, при загрузке.
// Справочники — по виду карточки (первый аргумент saveRegRec/delReg).
ПИН_ГРУППЫ.forEach(г=>(г[3]||[]).forEach(имя=>подПИН(г[0], имя)));
подПИН(kind=>"спр:"+kind, "saveRegRec");
подПИН(kind=>"спр:"+kind, "delReg");
