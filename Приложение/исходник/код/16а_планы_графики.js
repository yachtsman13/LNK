// ── Планы-графики проведения работ: ВД, ТК, НК ──────────────
// Три присланных книги («План-график ВД 2027», «ТК 2026-2027», «НК 2027»)
// перенесены как есть скриптом scripts/export_pg.py и вшиты в программу
// (тег lnk-pg). Дальше это рабочие данные, как план ППР: правятся,
// ходят в обмене и в копии.
//
// Устройство — три реестра:
//  · графики (lnk_pg) — сам документ: вид контроля и период «с — по»;
//  · позиции (lnk_pg_pos) — строки графика; месяцы — объект
//    {"2027-06":"", "2027-12":"2027-12-14"}: ключ — месяц, где стоит «+»,
//    значение — уточнённый день (пусто — уточнения ещё нет);
//  · отметки о выполнении (lnk_pg_otm) — отдельным реестром, как у ППР:
//    отмечает исполнитель, а позицию правит планирующий, и при обмене
//    одно не должно перетирать другое.
// У ВД выполнение видно и без отметки: замер в карточке учёта вибрации
// с датой в этом месяце засчитывает работу сам.
const ПГ_СИД=(function(){
  try{ const t=document.getElementById("lnk-pg").textContent.trim();
       const d=t?JSON.parse(t):null;
       return (d && Array.isArray(d.графики)) ? d : {графики:[],позиции:[],карты:[]};
  }catch(e){ return {графики:[],позиции:[],карты:[]}; }
})();
const REG_ПГ="lnk_pg", REG_ПГП="lnk_pg_pos", REG_ПГО="lnk_pg_otm";
let ПГ=[], ПГП=[], ПГО=[];
// Подписи полей — для истории правок.
const ПГ_ПОЛЯ=[{k:"нп",l:"№ п/п"},{k:"наименование",l:"Наименование"},{k:"подразделение",l:"Заказывающее подразделение",type:"podr"},
  {k:"количество",l:"Количество"},{k:"методы",l:"Методы контроля"},{k:"доля",l:"Периодичность, доля года"},
  {k:"мес_первый",l:"Месяц первого измерения"},{k:"трудозатраты",l:"Трудозатраты"},{k:"карта_ид",l:"Карточка учёта вибрации"},
  {k:"срок_службы",l:"Дата окончания срока службы",type:"date"},{k:"примечание",l:"Примечание"},{k:"месяцы",l:"Месяц (день)"},
  {k:"с",l:"Период с"},{k:"по",l:"Период по"},{k:"выполнено",l:"Выполнено",type:"date"}];
// Названия — дословно из заголовков присланных графиков.
const ПГ_ВИДЫ={
  ВД:{полно:"по вибродиагностическому контролю", наим:"Наименование технических средств", методы:["ВД"]},
  ТК:{полно:"по тепловому контролю", наим:"Наименование объекта контроля", методы:["ТК"]},
  НК:{полно:"по неразрушающему контролю", наим:"Наименование и тип объекта контроля", методы:null}};
const ПГ_ПОРЯДОК=["ВД","ТК","НК"];
// Периодичность ВД — доля года, как в столбце «Периодичность2» книги;
// словами — ровно как формула столбца «Периодичность» той же книги.
const ПГ_ДОЛИ=[[1,"раз в год"],[0.5,"раз в полгода"],[0.25,"раз в квартал"],[0.083,"раз в месяц"]];
function периодичностьПГ(доля){
  const x=ПГ_ДОЛИ.find(([д])=>Math.abs(д-(+доля))<1e-9); return x?x[1]:"";
}
function ymСдвиг(ym,n){
  const [г,м]=String(ym).split("-").map(Number); const i=г*12+(м-1)+n;
  return Math.floor(i/12)+"-"+String(i%12+1).padStart(2,"0");
}
function ymТекущий(){ const н=new Date(); return н.getFullYear()+"-"+String(н.getMonth()+1).padStart(2,"0"); }
function ymИмя(ym){ const [г,м]=String(ym).split("-"); return МЕСЯЦЫ[+м-1]+" "+г; }
function ymПосл(ym){ const [г,м]=String(ym).split("-").map(Number); return new Date(Date.UTC(г,м,0)).toISOString().slice(0,10); }
const РЕ_YM=/^\d{4}-(0[1-9]|1[0-2])$/;
function месяцыГрафика(г){
  const из=[]; if(!г || !РЕ_YM.test(г.с||"") || !РЕ_YM.test(г.по||"")) return из;
  let ym=г.с; for(let k=0;k<60 && ym<=г.по;k++){ из.push(ym); ym=ymСдвиг(ym,1); }
  return из;
}
function периодТекстПГ(г){
  const гс=String(г.с||"").slice(0,4), гп=String(г.по||"").slice(0,4);
  return гс===гп ? гс+" г." : гс+" - "+гп+" гг.";
}
function названиеПГ(г){ return ((ПГ_ВИДЫ[г.вид]||{}).полно||"")+" на "+периодТекстПГ(г); }
function краткоПГ(г){
  const гс=String(г.с||"").slice(0,4), гп=String(г.по||"").slice(0,4);
  return г.вид+" "+(гс===гп?гс:гс+"–"+гп);
}
function графикиПГ(){
  return live(ПГ).slice().sort((a,b)=>ПГ_ПОРЯДОК.indexOf(a.вид)-ПГ_ПОРЯДОК.indexOf(b.вид)||String(a.с).localeCompare(String(b.с)));
}
function позицииГрафика(ид){
  return live(ПГП).filter(п=>тотЖеИд(п.пг,ид)).sort((a,b)=>(+a.нп||0)-(+b.нп||0));
}
function графикПозиции(п){ return найтиПоИд(live(ПГ),п.пг); }
function имяПозицииПГ(п){ return String(п.наименование||"").trim()||("позиция № "+(п.нп||"")+" (наименование не внесено)"); }
// Правило формул книги ВД: в каждом году первый месяц, затем шаг
// round(доля × 12) до конца года.
function месяцыПоПравилу(доля, первый, г){
  const шаг=Math.round((+доля)*12), м0=+первый;
  if(!шаг || !(м0>=1 && м0<=12)) return [];
  const годы=[...new Set(месяцыГрафика(г).map(ym=>ym.slice(0,4)))];
  const из=[], есть=new Set(месяцыГрафика(г));
  годы.forEach(год=>{ for(let м=м0;м<=12;м+=шаг){ const ym=год+"-"+String(м).padStart(2,"0"); if(есть.has(ym)) из.push(ym); } });
  return из;
}
// Методы позиции — для подсказок о поверке приборов. У НК они в самой
// строке («ВИК, УК, АЭ, ТВ»); УК — ультразвук, им аттестуют сразу и на
// УЗД, и на УЗТ (см. аттестацияОхватывает), здесь так же.
function методыПозицииПГ(п, г){
  const в=(ПГ_ВИДЫ[(г||графикПозиции(п)||{}).вид]||{}).методы;
  if(в) return в.slice();
  const из=[];
  String(п.методы||"").split(/[,;\s]+/).map(x=>x.trim().toUpperCase()).filter(Boolean).forEach(м=>{
    if(м==="УК") из.push("УЗД","УЗТ"); else if(METHODS.indexOf(м)>=0) из.push(м); });
  return [...new Set(из)];
}

// ── Отметки о выполнении ──────────────────────────────────
function ключПГО(п,ym){ return п.id+"|"+ym; }
function отметкаПГ(п,ym){
  const о=live(ПГО).find(x=>x.id===ключПГО(п,ym));
  if(о) return {выполнено:о.выполнено, кто:о.кто||"", ист:"отметка"};
  if(п.карта_ид){
    const з=замерыКарты(п.карта_ид).find(з=>String(з.дата||"").slice(0,7)===ym);
    if(з) return {выполнено:з.дата, кто:з.исп||"", ист:"замер"};
  }
  return null;
}
function отметитьПГ(позИд, ym, дата){
  const п=найтиПоИд(live(ПГП),позИд); if(!п) return;
  if(!/^\d{4}-\d{2}-\d{2}$/.test(дата||"")){ toast("Укажите дату выполнения"); return; }
  if(дата>new Date().toISOString().slice(0,10)){ toast("Дата выполнения позже сегодняшней"); return; }
  const к=ключПГО(п,ym), ts=new Date().toISOString(), м=getWS().код;
  const был=ПГО.find(о=>о.id===к);
  if(был) Object.assign(был,{удалено:"", выполнено:дата, изменено:ts, место:м});
  else ПГО.push({id:к, поз:п.id, мес:ym, выполнено:дата, кто:"", изменено:ts, место:м});
  saveReg(REG_ПГО,ПГО);
  closeDrawer(); renderПГ();
  toast("Отмечено: "+имяПозицииПГ(п)+", "+ymИмя(ym));
}
function снятьОтметкуПГ(позИд, ym){
  const о=ПГО.find(x=>x.id===String(позИд)+"|"+ym && !x.удалено); if(!о) return;
  о.удалено=new Date().toISOString(); о.изменено=о.удалено; о.место=getWS().код;
  saveReg(REG_ПГО,ПГО);
  closeDrawer(); renderПГ();
}
// Уточнённый день — в самой позиции (это планирование, а не отметка).
function деньПГ(позИд, ym, дата){
  const п=найтиПоИд(live(ПГП),позИд); if(!п || !(ym in (п.месяцы||{}))) return;
  if(дата && String(дата).slice(0,7)!==ym){ toast("День должен быть в том же месяце — "+ymИмя(ym)); return; }
  п.месяцы[ym]=дата||""; п.изменено=new Date().toISOString(); п.место=getWS().код;
  saveReg(REG_ПГП,ПГП);
  closeDrawer(); renderПГ();
  toast(дата?("Запланировано на "+ru(дата)):"Уточнение дня снято");
}

// ── Что ещё приходится на месяц: ППР и поверка ─────────────
// Заказчик просил распределять работы с учётом обслуживания оборудования
// и поверки. Своего правила «можно/нельзя» здесь нет — программа
// показывает, что ещё стоит в этом месяце, решает планирующий.
function приборыМетодов(методы){
  const мт=getМетодики(), из=new Map();
  приборыСПоверкой().forEach(п=>{
    if(методы.some(м=>п.метод===м || ((мт[м]||{}).приборы||[]).some(x=>тотЖеИд(x,п.id)))) из.set(String(п.id),п);
  });
  return [...из.values()];
}
function поверкаВМесяце(методы, ym){
  const [г,м]=ym.split("-"), имя=МЕСЯЦЫ[+м-1];
  return приборыМетодов(методы).map(п=>{
    const кон=String(п.поверка_до||"").slice(0,7)===ym;
    const пред=месяцПрибора(п,+г)===имя;
    return (кон||пред)?{п, кон, пред}:null;
  }).filter(Boolean);
}
function загрузкаМесяцаПГ(ym){
  const [г,м]=ym.split("-").map(Number);
  let пг=0; live(ПГП).forEach(п=>{ if(ym in (п.месяцы||{}) && графикПозиции(п)) пг++; });
  return {пг, ппр:позицииППР().length?работыМесяца(г,м).length:0};
}

// ── Раздел ────────────────────────────────────────────────
const пгСост={граф:"", подр:"", поиск:""};
function текущийГрафикПГ(){
  const все=графикиПГ();
  let г=найтиПоИд(все,пгСост.граф);
  if(!г){
    const тек=ymТекущий();
    г=все.find(x=>x.с<=тек && тек<=x.по) || все.find(x=>x.по>=тек) || все[все.length-1] || null;
    пгСост.граф=г?г.id:"";
  }
  return г;
}
function setПГ(k,v){ пгСост[k]=v; renderПГ(); }
function подрПГ(п){ const д=п.подразделение?найтиПоИд(live(ПОДР),п.подразделение):null; return д?д.наименование:""; }
function рядыПГ(г){
  const q=(пгСост.поиск||"").toLowerCase();
  return позицииГрафика(г.id).filter(п=>{
    if(пгСост.подр==="—" ? п.подразделение : (пгСост.подр && !тотЖеИд(п.подразделение,пгСост.подр))) return false;
    if(q && ![п.наименование,п.методы,п.примечание,подрПГ(п),п.нп].join(" ").toLowerCase().includes(q)) return false;
    return true;
  });
}
function итогиПГ(г, ряды){
  const тек=ymТекущий(); let работ=0, сделано=0, просрочено=0, уточнено=0;
  ряды.forEach(п=>Object.keys(п.месяцы||{}).forEach(ym=>{
    работ++; const о=отметкаПГ(п,ym);
    if(о) сделано++; else if(ym<тек) просрочено++;
    if(п.месяцы[ym]) уточнено++;
  }));
  return {работ, сделано, просрочено, уточнено};
}
function renderПГ(){
  curView="pg"; setChrome();
  document.querySelectorAll(".nav").forEach(n=>n.classList.toggle("on",n.dataset.view==="pg"));
  const все=графикиПГ(), г=текущийГрафикПГ();
  if(!г){
    document.getElementById("main").innerHTML=`<div class="head"><h2>Планы-графики</h2></div>
      <div class="placeholder"><div><div class="big">▦</div>Планов-графиков нет.<br><br>
      <button class="btn primary" onclick="editПГ('')">＋ План-график</button>
      <button class="btn" onclick="document.getElementById('imp_pg').click()">⬆ Из Excel</button>
      <input type="file" id="imp_pg" accept=".xlsx" style="display:none" onchange="импортПГВыбран(this)"></div></div>`;
    return;
  }
  const мес=месяцыГрафика(г), тек=ymТекущий(), ряды=рядыПГ(г), и=итогиПГ(г,ряды);
  const вд=г.вид==="ВД", нк=г.вид==="НК";
  const подрВГр=[...new Set(позицииГрафика(г.id).map(п=>String(п.подразделение||"")).filter(Boolean))]
    .map(id=>найтиПоИд(live(ПОДР),id)).filter(Boolean).sort((a,b)=>String(a.наименование).localeCompare(String(b.наименование),"ru"));
  const безПодр=позицииГрафика(г.id).filter(п=>!п.подразделение).length;
  const ячейка=(п,ym)=>{
    if(!(ym in (п.месяцы||{}))) return `<td class="мес пусто"></td>`;
    const о=отметкаПГ(п,ym), день=п.месяцы[ym];
    const кл=о?"сделано":(ym<тек?"просрочено":"план");
    const т=о?"✓":(день?String(+день.slice(8)):"+");
    const подск=о?("выполнено "+ru(о.выполнено)+(о.ист==="замер"?" (замер в карточке вибрации)":"")):
      (ymИмя(ym)+(день?", на "+ru(день):"")+" — нажмите: отметить, уточнить день");
    return `<td class="мес ${кл}${ym===тек?" тек":""}" title="${esc(подск)}" onclick="пгЯчейка('${esc(п.id)}','${ym}')">${т}</td>`;
  };
  const строки=ряды.map(п=>{
    const к=п.карта_ид?найтиПоИд(live(ВИБК),п.карта_ид):null;
    return `<tr>
      <td class="num">${esc(п.нп)}</td>
      <td class="obj" style="cursor:pointer" title="Править позицию" onclick="editПГП('${esc(п.id)}')">${esc(имяПозицииПГ(п))}
        ${вд&&к?`<div class="muted" style="font-size:11.5px">карточка вибрации № ${esc(к.номер)}</div>`:""}
        ${нк&&п.примечание?`<div class="muted" style="font-size:11.5px">${esc(п.примечание)}</div>`:""}</td>
      ${вд?`<td>${esc(периодичностьПГ(п.доля))}</td>`:""}
      ${нк?`<td>${esc(п.методы)}</td>`:""}
      ${!вд?`<td class="num">${esc(п.количество)}</td>`:""}
      <td>${esc(подрПГ(п))||"<span class='muted'>—</span>"}</td>
      ${мес.map(ym=>ячейка(п,ym)).join("")}
      ${нк?`<td class="num">${п.срок_службы?ru(п.срок_службы):""}</td>`:""}
    </tr>`;
  }).join("");
  const колДо=2+(вд?1:0)+(нк?1:0)+(!вд?1:0)+1;
  const годШапки=мес.length && мес[0].slice(0,4)!==мес[мес.length-1].slice(0,4);
  const шапкаМес=мес.map(ym=>`<th class="мес${ym===тек?" тек":""}">${МЕС3[+ym.slice(5)-1]}${годШапки?`<br><span class="muted">${ym.slice(0,4)}</span>`:""}</th>`).join("");
  const методы=new Set(); позицииГрафика(г.id).forEach(п=>методыПозицииПГ(п,г).forEach(м=>методы.add(м)));
  const итогРабот=мес.map(ym=>`<td class="мес"><b>${ряды.filter(п=>ym in (п.месяцы||{})).length||""}</b></td>`).join("");
  const итогППР=мес.map(ym=>{ const n=загрузкаМесяцаПГ(ym).ппр; return `<td class="мес">${n||""}</td>`; }).join("");
  const итогПов=мес.map(ym=>{ const n=поверкаВМесяце([...методы],ym).length; return `<td class="мес${n?" пгпов":""}"${n?` title="Нажмите — какие приборы" onclick="пгПоверкаМесяца('${ym}')" style="cursor:pointer"`:""}>${n||""}</td>`; }).join("");
  const безКарт=вд?позицииГрафика(г.id).filter(п=>!п.карта_ид || !найтиПоИд(live(ВИБК),п.карта_ид)).length:0;
  document.getElementById("main").innerHTML=`
    ${printHead("ПЛАН-ГРАФИК проведения работ", названиеПГ(г)+(пгСост.подр&&пгСост.подр!=="—"?" · "+(найтиПоИд(live(ПОДР),пгСост.подр)||{}).наименование:""))}
    <div class="head"><h2>Планы-графики</h2><span class="sub">${esc(названиеПГ(г))}</span>
      <div class="search"><span class="ic">⌕</span><input id="q" placeholder="Поиск по наименованию, методам…" value="${esc(пгСост.поиск)}"></div>
      <div class="spring"></div>
      <button class="btn primary" onclick="editПГП('','${esc(г.id)}')">＋ Позиция</button>
      <button class="btn" onclick="editПГ('${esc(г.id)}')" title="Период, вид; удалить план-график">✎ План-график</button>
      <button class="btn" onclick="editПГ('')" title="Новый план-график — пустой или по образцу имеющегося">＋ План-график</button>
      <button class="btn" onclick="document.getElementById('imp_pg').click()" title="План-график в той же форме, что присланные книги ВД, ТК, НК">⬆ Из Excel</button>
      <input type="file" id="imp_pg" accept=".xlsx" style="display:none" onchange="импортПГВыбран(this)">
      <button class="btn" onclick="exportПГWord()" title="Показанные позиции — выписка для подразделения, если оно выбрано">⤓ Word</button>
      <button class="btn" onclick="exportПГXlsx()">▤ Excel</button>
      <button class="btn" onclick="window.print()">⎙ PDF</button></div>
    <div class="filters noprint seg">
      ${все.map(x=>`<button class="btn ${x.id===г.id?"on":""}" onclick="пгСост.граф='${esc(x.id)}';пгСост.подр='';renderПГ()">${esc(краткоПГ(x))}</button>`).join("")}
      <select onchange="setПГ('подр',this.value)" style="margin-left:10px" title="Выписка для заказывающего подразделения: выберите его и выгрузите в Word или Excel">
        <option value="">Все подразделения</option>
        ${подрВГр.map(д=>`<option value="${esc(д.id)}" ${String(д.id)===String(пгСост.подр)?"selected":""}>${esc(д.наименование)}</option>`).join("")}
        ${безПодр?`<option value="—" ${пгСост.подр==="—"?"selected":""}>Подразделение не указано (${безПодр})</option>`:""}</select>
    </div>
    <div class="tiles">
      <div class="tile"><div class="n">${ряды.length}</div><div class="l">позиций${ряды.length!==позицииГрафика(г.id).length?" (по отбору)":""}</div></div>
      <div class="tile"><div class="n">${и.работ}</div><div class="l">работ по плану-графику</div></div>
      <div class="tile ${и.сделано?"ok":""}"><div class="n">${и.сделано}</div><div class="l">выполнено</div></div>
      <div class="tile ${и.просрочено?"bad":"ok"}"><div class="n">${и.просрочено}</div><div class="l">прошедший месяц без отметки</div></div>
      <div class="tile"><div class="n">${и.уточнено}</div><div class="l">уточнён день</div></div>
    </div>
    ${безПодр?`<div class="draftnote noprint">Заказывающее подразделение не указано у ${безПодр} из ${позицииГрафика(г.id).length} позиций —
      выписка по подразделению соберёт только тех, у кого оно есть. Указывается в позиции (щелчок по наименованию).</div>`:""}
    ${вд&&безКарт?`<div class="draftnote noprint">Без карточки учёта вибрации: ${безКарт} ${склЧисло(безКарт,"позиция","позиции","позиций")}.
      <button class="btn small" onclick="картыПоПлануВД('${esc(г.id)}')">Завести карточки</button></div>`:""}
    <div class="card"><div class="tblwrap"><table class="плантбл пгтбл">
      <thead><tr>
        <th style="width:44px">№ п/п</th><th>${esc((ПГ_ВИДЫ[г.вид]||{}).наим||"Наименование")}</th>
        ${вд?`<th style="width:110px">Периодичность</th>`:""}
        ${нк?`<th style="width:110px">Методы контроля</th>`:""}
        ${!вд?`<th style="width:60px">Кол.</th>`:""}
        <th style="width:150px">Заказывающее подразделение</th>
        ${шапкаМес}${нк?`<th style="width:96px">Окончание срока службы</th>`:""}
      </tr></thead>
      <tbody>${строки||`<tr><td colspan="${колДо+мес.length+(нк?1:0)}" class="empty">${позицииГрафика(г.id).length?"По отбору позиций нет.":"Позиций нет — «＋ Позиция»."}</td></tr>`}</tbody>
      <tfoot>
        <tr><td colspan="${колДо}" class="num"><b>Работ по плану-графику</b></td>${итогРабот}${нк?"<td></td>":""}</tr>
        <tr class="noprint"><td colspan="${колДо}" class="num muted">Работ ППР (план работ)</td>${итогППР}${нк?"<td></td>":""}</tr>
        <tr class="noprint"><td colspan="${колДо}" class="num muted">Приборов ${esc([...методы].join(", "))} на поверку</td>${итогПов}${нк?"<td></td>":""}</tr>
      </tfoot>
    </table></div></div>
    <div class="note noprint">«+» — работа по плану-графику в этом месяце; число — уточнённый день; ✓ — выполнено
      (зелёным), красным — прошедший месяц без отметки. Нажатие на месяц: отметить выполнение, уточнить день,
      посмотреть, что ещё в этом месяце (ППР, поверка приборов).
      ${вд?"Работа ВД засчитывается и сама — замером в карточке учёта вибрации с датой в этом месяце.":""}
      Две нижние строки — для распределения работ: сколько работ ППР в месяце и сколько приборов этих методов
      уходит на поверку (по месяцу представления или окончанию поверки).</div>`;
  const q=document.getElementById("q");
  if(q) q.oninput=()=>{ пгСост.поиск=q.value; const п=q.selectionStart; renderПГ(); const q2=document.getElementById("q"); q2.focus(); try{ q2.setSelectionRange(п,п); }catch(e){} };
}
function пгЯчейка(позИд, ym){
  const п=найтиПоИд(live(ПГП),позИд); if(!п) return;
  const г=графикПозиции(п), о=отметкаПГ(п,ym), день=(п.месяцы||{})[ym]||"";
  const сег=new Date().toISOString().slice(0,10);
  const методы=методыПозицииПГ(п,г), пов=поверкаВМесяце(методы,ym);
  const [гг,мм]=ym.split("-").map(Number);
  const ппр=позицииППР().length?работыМесяца(гг,мм):[];
  const тотЖеДень=день?live(ПГП).filter(x=>x!==п && (x.месяцы||{})[ym]===день):[];
  const к=п.карта_ид?найтиПоИд(live(ВИБК),п.карта_ид):null;
  const до=ymПосл(ym), отДаты=день||(до<сег?до:(ym+"-01">сег?"":сег));
  openDrawer(ymИмя(ym)+" — "+имяПозицииПГ(п), `
    <div class="note" style="margin-top:0">${esc(краткоПГ(г||{}))}${п.методы?" · "+esc(п.методы):""}${подрПГ(п)?" · "+esc(подрПГ(п)):""}</div>
    <div class="sect">Выполнение</div>
    ${о?`<p><span class="st ok">выполнено ${ru(о.выполнено)}</span> ${о.ист==="замер"?"— по замеру в карточке учёта вибрации":""}</p>
        ${о.ист==="отметка"?`<button class="btn" onclick="снятьОтметкуПГ('${esc(п.id)}','${ym}')">Снять отметку</button>`:""}`
       :`<p>Дата выполнения <input type="date" id="пгДата" value="${отДаты}" max="${сег}">
         <button class="btn primary" onclick="отметитьПГ('${esc(п.id)}','${ym}',document.getElementById('пгДата').value)">✓ Отметить</button></p>`}
    ${к?`<p><button class="btn" onclick="closeDrawer();открытьКартуВибрации('${esc(к.id)}')">Карточка вибрации № ${esc(к.номер)}</button>
        <button class="btn" onclick="closeDrawer();editЗамерВибр('${esc(к.id)}','')">＋ Замер</button></p>`:""}
    <div class="sect">Уточнить день</div>
    <p><input type="date" id="пгДень" value="${день}" min="${ym}-01" max="${до}">
      <button class="btn" onclick="деньПГ('${esc(п.id)}','${ym}',document.getElementById('пгДень').value)">Сохранить день</button>
      ${день?`<button class="btn" onclick="деньПГ('${esc(п.id)}','${ym}','')">Убрать</button>`:""}</p>
    ${тотЖеДень.length?`<div class="note">На ${ru(день)} по планам-графикам также: ${тотЖеДень.slice(0,8).map(x=>esc(имяПозицииПГ(x))).join("; ")}${тотЖеДень.length>8?" …":""}</div>`:""}
    <div class="sect">Что ещё в этом месяце</div>
    <p><b>ППР:</b> ${ппр.length?склЧисло(ппр.length,"работа","работы","работ")+": "+ппр.slice(0,8).map(р=>esc(р.поз.наименование)).join("; ")+(ппр.length>8?" …":""):"работ нет"}</p>
    <p><b>Поверка приборов (${esc(методы.join(", ")||"методы не указаны")}):</b>
      ${пов.length?пов.map(x=>esc(x.п.наименование)+(x.п.зав?", зав. № "+esc(x.п.зав):"")+" — "+
        [x.пред?"представляется на поверку":"", x.кон?"поверка до "+ru(x.п.поверка_до):""].filter(Boolean).join(", ")+
        (день && x.кон && x.п.поверка_до<день?" <b style='color:var(--danger)'>— на "+ru(день)+" уже не поверен</b>":"")).join("<br>"):"в этом месяце не уходят"}</p>
    <div class="note">Решение, когда проводить, — за вами: программа только показывает, что ещё стоит в этом месяце.</div>`);
}
function пгПоверкаМесяца(ym){
  const г=текущийГрафикПГ(); if(!г) return;
  const методы=new Set(); позицииГрафика(г.id).forEach(п=>методыПозицииПГ(п,г).forEach(м=>методы.add(м)));
  const пов=поверкаВМесяце([...методы],ym);
  openDrawer("Поверка приборов — "+ymИмя(ym), пов.map(x=>`<p>${esc(x.п.наименование)}${x.п.зав?", зав. № "+esc(x.п.зав):""}
    <span class="muted">${esc(x.п.метод||"")}</span><br>${[x.пред?"представляется на поверку в этом месяце":"", x.кон?"поверка действует до "+ru(x.п.поверка_до):""].filter(Boolean).join("; ")}</p>`).join("")||"Нет.");
}

// ── Позиция ───────────────────────────────────────────────
let _пгПозГраф="";
function editПГП(id, пгИд){
  const п=id?найтиПоИд(live(ПГП),id):null;
  if(id && !п){ toast("Позиция не найдена"); return; }
  const г=п?графикПозиции(п):найтиПоИд(live(ПГ),пгИд);
  if(!г){ toast("План-график не найден"); return; }
  if(curView!=="pgedit" && !уходСФормы()) return;
  формаСохранена();
  _пгПозГраф=г.id; curView="pgedit";
  document.querySelectorAll(".nav").forEach(n=>n.classList.toggle("on",n.dataset.view==="pg"));
  const v=x=>esc(x==null?"":String(x)), вд=г.вид==="ВД", нк=г.вид==="НК";
  const мес=месяцыГрафика(г), м=(п&&п.месяцы)||{};
  const нп=п?п.нп:(позицииГрафика(г.id).reduce((a,x)=>Math.max(a,+x.нп||0),0)+1);
  const отм=п?live(ПГО).filter(о=>тотЖеИд(о.поз,п.id)).length:0;
  const клетки=мес.map(ym=>`<label class="пгмес"><span><input type="checkbox" id="пгм_${ym}" ${ym in м?"checked":""}> ${МЕС3[+ym.slice(5)-1]} ${ym.slice(0,4)}</span>
    <input type="date" id="пгд_${ym}" value="${v(м[ym]||"")}" min="${ym}-01" max="${ymПосл(ym)}" title="Уточнённый день — необязательно"></label>`).join("");
  const карты=live(ВИБК).slice().sort((a,b)=>String(a.номер||"").localeCompare(String(b.номер||""),"ru",{numeric:true}));
  document.getElementById("main").innerHTML=`
    <div class="head"><h2>${п?"Позиция плана-графика":"Новая позиция плана-графика"}</h2>
      <span class="sub">${esc(краткоПГ(г))}</span><div class="spring"></div>
      <button class="btn" onclick="switchView('pg')">← К плану-графику</button></div>
    <div class="card" style="padding:20px"><div class="form">
      <div class="grid">
        <label>№ п/п<input id="пг_нп" inputmode="numeric" value="${v(нп)}"></label>
        <label>Заказывающее подразделение<select id="пг_подразделение"><option value="">— не указано —</option>
          ${live(ПОДР).map(д=>`<option value="${esc(д.id)}" ${п&&тотЖеИд(п.подразделение,д.id)?"selected":""}>${esc(д.наименование)}</option>`).join("")}</select></label>
        <label class="wide">${esc((ПГ_ВИДЫ[г.вид]||{}).наим)}<input id="пг_наименование" value="${v(п&&п.наименование)}"></label>
        ${вд?`<label>Периодичность<select id="пг_доля">${ПГ_ДОЛИ.map(([д,т])=>`<option value="${д}" ${п&&Math.abs(+п.доля-д)<1e-9?"selected":""}>${т}</option>`).join("")}</select></label>
          <label>Месяц первого измерения<select id="пг_мес_первый">${МЕСЯЦЫ.map((x,i)=>`<option value="${i+1}" ${п&&+п.мес_первый===i+1?"selected":""}>${x}</option>`).join("")}</select></label>
          <label>Трудозатраты<input id="пг_трудозатраты" value="${v(п&&п.трудозатраты)}"></label>
          <label>Карточка учёта вибрации<select id="пг_карта_ид"><option value="">— нет —</option>
            ${карты.map(к=>`<option value="${esc(к.id)}" ${п&&тотЖеИд(п.карта_ид,к.id)?"selected":""}>№ ${esc(к.номер)} ${esc(заголовокКарты(к))}</option>`).join("")}</select></label>`:""}
        ${нк?`<label>Методы контроля<input id="пг_методы" value="${v(п&&п.методы)}" placeholder="ВИК, УК, АЭ, ТВ"></label>
          <label>Дата окончания срока службы${п&&п.срок_службы&&!/^\d{4}-\d{2}-\d{2}$/.test(п.срок_службы)
            ?`<input id="пг_срок_службы" value="${v(п.срок_службы)}" title="В книге здесь текст, а не дата — оставлен как есть">`
            :`<input type="date" id="пг_срок_службы" value="${v(п&&п.срок_службы)}">`}</label>`:""}
        ${!вд?`<label>Количество<input id="пг_количество" value="${v(п&&п.количество)}"></label>`:""}
        <label class="wide">Примечание<input id="пг_примечание" value="${v(п&&п.примечание)}"></label>
      </div>
      <div class="sect" style="margin-top:14px">Месяцы проведения работ
        ${вд?`<button class="btn small" type="button" onclick="пгПоПравилу()" title="Отметить месяцы по периодичности и месяцу первого измерения — как формулы присланной книги">по периодичности</button>`:""}</div>
      <div class="пгмесяцы">${клетки}</div>
      <div class="note">Галочка — работа в этом месяце («+» в плане-графике). День — необязательное уточнение,
        когда именно. ${вд?"«По периодичности» ставит месяцы так же, как формулы книги ВД: первый месяц, дальше через "+
        "12 × периодичность месяцев до конца года.":""}</div>
      ${отм?`<div class="note">Отметок о выполнении у позиции: ${отм}. Убранный месяц отметку не удаляет — она вернётся вместе с месяцем.</div>`:""}
      <div class="formbtns">
        <button class="btn primary" onclick="saveПГП(${п?`'${esc(п.id)}'`:"null"})">✓ Сохранить</button>
        <button class="btn" onclick="switchView('pg')">Отмена</button>
        ${п?`<button class="btn danger" style="margin-left:auto" onclick="delПГП('${esc(п.id)}')">✕ Удалить</button>`:""}
      </div>
      ${п?блокИстории(REG_ПГП,п.id):""}
    </div></div>`;
}
function пгПоПравилу(){
  const г=найтиПоИд(live(ПГ),_пгПозГраф); if(!г) return;
  const нужно=new Set(месяцыПоПравилу(+document.getElementById("пг_доля").value, +document.getElementById("пг_мес_первый").value, г));
  месяцыГрафика(г).forEach(ym=>{ const ч=document.getElementById("пгм_"+ym); if(ч) ч.checked=нужно.has(ym); });
  _несохранено=true;
}
function saveПГП(id){
  const г=найтиПоИд(live(ПГ),_пгПозГраф); if(!г) return;
  const был=id?ПГП.find(x=>тотЖеИд(x.id,id)):null;
  const зн=k=>{ const e=document.getElementById("пг_"+k); return e?String(e.value||"").trim():undefined; };
  const месяцы={}, ошибки=[];
  месяцыГрафика(г).forEach(ym=>{
    const ч=document.getElementById("пгм_"+ym), д=document.getElementById("пгд_"+ym).value;
    if(ч.checked) месяцы[ym]=д||"";
    else if(д) ошибки.push(ymИмя(ym)+": день указан, а месяц не отмечен");
    if(д && д.slice(0,7)!==ym) ошибки.push(ymИмя(ym)+": день не из этого месяца");
  });
  if(ошибки.length){ alert(ошибки.join("\n")); return; }
  if(!зн("наименование") && !confirm("Наименование не вписано. Сохранить позицию без наименования?")) return;
  if(был){
    const убраны=live(ПГО).filter(о=>тотЖеИд(о.поз,был.id) && !(о.мес in месяцы));
    if(убраны.length && !confirm("Убираются месяцы с отметкой о выполнении: "+убраны.map(о=>ymИмя(о.мес)).join(", ")+
      ".\nОтметки не удаляю — вернёте месяц, вернутся и они. Сохранить?")) return;
  }
  const поля={нп:+зн("нп")||зн("нп")||"", наименование:зн("наименование"), подразделение:зн("подразделение"),
    примечание:зн("примечание"), месяцы, изменено:new Date().toISOString(), место:getWS().код};
  ["доля","мес_первый"].forEach(k=>{ if(зн(k)!==undefined) поля[k]=+зн(k); });
  ["трудозатраты","карта_ид","методы","срок_службы","количество"].forEach(k=>{ if(зн(k)!==undefined) поля[k]=зн(k); });
  if(был) Object.assign(был,поля);
  else ПГП.push(Object.assign({id:newId("ПГП"), пг:г.id},поля));
  saveReg(REG_ПГП,ПГП);
  формаСохранена(); switchView("pg");
  toast(был?"Позиция сохранена":"Позиция добавлена");
}
function delПГП(id){
  const п=ПГП.find(x=>тотЖеИд(x.id,id)); if(!п) return;
  if(!confirm("Удалить из плана-графика «"+имяПозицииПГ(п)+"»? Удаление уйдёт в обмен на другие места.")) return;
  п.удалено=new Date().toISOString(); п.изменено=п.удалено; п.место=getWS().код;
  saveReg(REG_ПГП,ПГП);
  формаСохранена(); switchView("pg"); toast("Позиция удалена");
}

// ── План-график: период, новый по образцу, удаление ────────
function editПГ(id){
  const г=id?найтиПоИд(live(ПГ),id):null;
  if(curView!=="pgedit" && !уходСФормы()) return;
  формаСохранена();
  curView="pgedit"; _пгПозГраф="";
  document.querySelectorAll(".nav").forEach(n=>n.classList.toggle("on",n.dataset.view==="pg"));
  const след=new Date().getFullYear()+1;
  const образцы=графикиПГ();
  document.getElementById("main").innerHTML=`
    <div class="head"><h2>${г?"План-график "+esc(краткоПГ(г)):"Новый план-график"}</h2><div class="spring"></div>
      <button class="btn" onclick="switchView('pg')">← К планам-графикам</button></div>
    <div class="card" style="padding:20px"><div class="form">
      <div class="grid">
        <label>Вид контроля<select id="пгг_вид" ${г?"disabled":""} onchange="пгОбразцыОбновить()">${ПГ_ПОРЯДОК.map(в=>`<option value="${в}" ${г&&г.вид===в?"selected":""}>${в} — ${ПГ_ВИДЫ[в].полно}</option>`).join("")}</select></label>
        <span></span>
        <label>Период — с месяца<input type="month" id="пгг_с" value="${г?esc(г.с):след+"-01"}"></label>
        <label>по месяц<input type="month" id="пгг_по" value="${г?esc(г.по):след+"-12"}"></label>
        ${г?"":`<label class="wide">Позиции по образцу<select id="пгг_образец"></select></label>
          <label class="wide" style="flex-direction:row;align-items:center;gap:8px"><input type="checkbox" id="пгг_сдвиг" checked style="width:auto">
            перенести месяцы образца на новый период (ВД — по периодичности позиции, ТК и НК — со сдвигом на разницу начала периодов)</label>`}
      </div>
      <div class="note">Заголовок документа строится сам: «ПЛАН-ГРАФИК проведения работ ${г?esc(названиеПГ(г)):"по … контролю на ГГГГ г."}».
        ${г?"":"По образцу переносятся наименования, подразделения, количество, методы, периодичность и карточки вибрации; отметки о выполнении — нет."}</div>
      <div class="formbtns">
        <button class="btn primary" onclick="saveПГ(${г?`'${esc(г.id)}'`:"null"})">✓ Сохранить</button>
        <button class="btn" onclick="switchView('pg')">Отмена</button>
        ${г?`<button class="btn danger" style="margin-left:auto" onclick="delПГ('${esc(г.id)}')">✕ Удалить план-график</button>`:""}
      </div>
      ${г?блокИстории(REG_ПГ,г.id):""}
    </div></div>`;
  if(!г) пгОбразцыОбновить();
}
function пгОбразцыОбновить(){
  const s=document.getElementById("пгг_образец"); if(!s) return;
  const в=document.getElementById("пгг_вид").value;
  s.innerHTML=`<option value="">— пустой, без позиций —</option>`+графикиПГ().filter(x=>x.вид===в)
    .map(x=>`<option value="${esc(x.id)}">${esc(краткоПГ(x))} (${позицииГрафика(x.id).length} поз.)</option>`).join("");
  if(s.options.length>1) s.selectedIndex=s.options.length-1;
}
function saveПГ(id){
  const с=document.getElementById("пгг_с").value, по=document.getElementById("пгг_по").value;
  if(!РЕ_YM.test(с)||!РЕ_YM.test(по)||по<с){ alert("Период: укажите месяцы «с» и «по», «по» не раньше «с»."); return; }
  if(месяцыГрафика({с,по}).length>24){ alert("Период длиннее двух лет — так план-график не составляется. Проверьте месяцы."); return; }
  const ts=new Date().toISOString(), место=getWS().код;
  const был=id?ПГ.find(x=>тотЖеИд(x.id,id)):null;
  if(был){
    const вне=позицииГрафика(был.id).filter(п=>Object.keys(п.месяцы||{}).some(ym=>ym<с||ym>по)).length;
    if(вне && !confirm("У "+вне+" позиций есть месяцы вне нового периода — в плане-графике их видно не будет (данные остаются). Сохранить?")) return;
    Object.assign(был,{с,по,изменено:ts,место});
    saveReg(REG_ПГ,ПГ);
    пгСост.граф=был.id; формаСохранена(); switchView("pg"); toast("План-график сохранён"); return;
  }
  const вид=document.getElementById("пгг_вид").value;
  if(live(ПГ).some(x=>x.вид===вид && x.с===с && x.по===по)){ alert("План-график "+вид+" на этот период уже есть."); return; }
  const нов={id:newId("ПГ"), вид, с, по, изменено:ts, место};
  ПГ.push(нов);
  const обр=найтиПоИд(live(ПГ),document.getElementById("пгг_образец").value);
  let n=0;
  if(обр){
    const сдв=document.getElementById("пгг_сдвиг").checked;
    const [г1,м1]=обр.с.split("-").map(Number), [г2,м2]=с.split("-").map(Number);
    const разн=(г2*12+м2)-(г1*12+м1);
    const мес=new Set(месяцыГрафика(нов));
    позицииГрафика(обр.id).forEach(п=>{
      const месяцы={};
      if(сдв){
        const список=вид==="ВД" && п.доля && п.мес_первый ? месяцыПоПравилу(п.доля,п.мес_первый,нов)
          : Object.keys(п.месяцы||{}).map(ym=>ymСдвиг(ym,разн)).filter(ym=>мес.has(ym));
        список.forEach(ym=>месяцы[ym]="");
      }
      const к=JSON.parse(JSON.stringify(п));
      ["id","удалено","база","изменено","место","месяцы"].forEach(x=>delete к[x]);
      if(вид==="НК") delete к.срок_службы;
      ПГП.push(Object.assign(к,{id:newId("ПГП"), пг:нов.id, месяцы, изменено:ts, место}));
      n++;
    });
  }
  saveReg(REG_ПГ,ПГ); if(n) saveReg(REG_ПГП,ПГП);
  пгСост.граф=нов.id; пгСост.подр=""; формаСохранена(); switchView("pg");
  toast("План-график "+краткоПГ(нов)+" создан"+(n?", позиций по образцу: "+n:""));
}
function delПГ(id){
  const г=ПГ.find(x=>тотЖеИд(x.id,id)); if(!г) return;
  const поз=позицииГрафика(г.id);
  if(!confirm("Удалить план-график "+краткоПГ(г)+" вместе с его позициями ("+поз.length+")? Удаление уйдёт в обмен на другие места; "+
    "вернуть можно в «Обмене» → «Что изменилось» → «все удалённые».")) return;
  const ts=new Date().toISOString(), место=getWS().код;
  г.удалено=ts; г.изменено=ts; г.место=место;
  поз.forEach(п=>{ п.удалено=ts; п.изменено=ts; п.место=место; });
  saveReg(REG_ПГ,ПГ); saveReg(REG_ПГП,ПГП);
  пгСост.граф=""; формаСохранена(); switchView("pg"); toast("План-график удалён");
}

// ── Карточки учёта вибрации по плану ВД ────────────────────
// Присланный план ВД 2027 несёт технические данные оборудования
// (производитель, год, мощность, подшипники, частоты вращения, моточасы)
// — из них заводятся карточки (сид, scripts/export_pg.py). У позиции,
// добавленной позже, карточку заводит эта кнопка: номер — № п/п позиции,
// наименование — из позиции; остальное заполняется в карточке.
function картыПоПлануВД(пгИд){
  const г=найтиПоИд(live(ПГ),пгИд); if(!г) return;
  const без=позицииГрафика(г.id).filter(п=>!п.карта_ид || !найтиПоИд(live(ВИБК),п.карта_ид));
  if(!без.length){ toast("У всех позиций есть карточки"); return; }
  if(!confirm("Завести карточки учёта вибрации для "+без.length+" позиций? Номер карточки — № п/п позиции, вид и тип — из наименования позиции.")) return;
  const ts=new Date().toISOString(), место=getWS().код;
  без.forEach(п=>{
    const id=newId("ВК");
    ВИБК.push(Object.assign({id, номер:String(п.нп||""), подразделение:п.подразделение||"", изменено:ts, место},картаИзПозиции(п.наименование)));
    п.карта_ид=id; п.изменено=ts; п.место=место;
  });
  saveReg(REG_ВИБК,ВИБК); saveReg(REG_ПГП,ПГП);
  renderПГ(); toast("Заведено карточек: "+без.length);
}
// Вид оборудования в родительном падеже — в заголовок карточки. Падежи
// программа сама не склоняет: закрытый список ровно тех начал, что есть
// в присланном плане ВД 2027 (тот же, что в scripts/export_pg.py).
const ПГ_РОДИТЕЛЬНЫЙ=[["Электровентилятор","электровентилятора"],["Электронасосный агрегат","электронасосного агрегата"],
  ["Вентиляционная установка","вентиляционной установки"],["Вентилятор","вентилятора"],["Насос","насоса"]];
function картаИзПозиции(наим){
  наим=String(наим||"").replace(/\s+/g," ").trim();
  for(const [нач,род] of ПГ_РОДИТЕЛЬНЫЙ){ const ост=наим.slice(нач.length).trim();
    if(наим.startsWith(нач) && ост) return {вид:род, наименование:ост}; }
  return {вид:"", наименование:наим};
}
// Совпадает ли наименование позиции с карточкой: целиком или «вид + тип».
function позицияКарты(наим, к){
  const н=x=>String(x||"").replace(/\s+/g," ").trim().toLowerCase(), р=картаИзПозиции(наим);
  return н(к.наименование)===н(наим) || (н(к.наименование)===н(р.наименование) && (!к.вид || н(к.вид)===н(р.вид)));
}
// Сид: карточки из присланного плана ВД с постоянными id — добавляются
// один раз (удалённую не возвращает: у неё остаётся id с пометкой).
function картыИзСидаПГ(){
  const есть=new Set((ВИБК||[]).map(к=>String(к.id))); let n=0;
  (ПГ_СИД.карты||[]).forEach(к=>{ if(!есть.has(String(к.id))){ ВИБК.push(Object.assign({},к)); n++; } });
  if(n) saveReg(REG_ВИБК,ВИБК);
  return n;
}
// Что по планам-графикам приходится на карточку вибрации — для самой карточки.
function планВДКарты(картаИд){
  const из=[];
  live(ПГП).forEach(п=>{ if(!тотЖеИд(п.карта_ид,картаИд)) return; const г=графикПозиции(п); if(!г) return;
    Object.keys(п.месяцы||{}).sort().forEach(ym=>из.push({г, п, ym, о:отметкаПГ(п,ym)})); });
  return из;
}

// ── Выгрузки: Word и Excel в форме присланных графиков ─────
function колонкиПГ(г){
  const мес=месяцыГрафика(г);
  const к=[{з:"№ п/п",ш:6,ф:п=>п.нп,тип:"число"},{з:(ПГ_ВИДЫ[г.вид]||{}).наим,ш:46,ф:п=>п.наименование||""}];
  if(г.вид==="ВД") к.push({з:"Периодичность",ш:16,ф:п=>периодичностьПГ(п.доля)});
  if(г.вид==="ТК") к.push({з:"Количество",ш:11,ф:п=>п.количество||""});
  к.push({з:"Заказывающее подразделение",ш:26,ф:п=>подрПГ(п)});
  if(г.вид==="НК"){ к.push({з:"Методы контроля (ВИК, УК, АЭ, ТВ)",ш:16,ф:п=>п.методы||""}); к.push({з:"Количество",ш:11,ф:п=>п.количество||""}); }
  мес.forEach(ym=>к.push({з:МЕСЯЦЫ[+ym.slice(5)-1]+(мес[0].slice(0,4)!==мес[мес.length-1].slice(0,4)?" "+ym.slice(0,4):""), ш:10, мес:ym,
    ф:п=>!(ym in (п.месяцы||{}))?"":(п.месяцы[ym]?ru(п.месяцы[ym]).slice(0,5):"+")}));
  if(г.вид==="НК"){ к.push({з:"Дата окончания срока службы",ш:14,ф:п=>п.срок_службы||"",тип:"дата"}); к.push({з:"Примечание",ш:24,ф:п=>п.примечание||""}); }
  return к;
}
function имяФайлаПГ(г, расш){
  const д=пгСост.подр==="—"?"без подразделения":(пгСост.подр?(найтиПоИд(live(ПОДР),пгСост.подр)||{}).наименование:"");
  return "План-график "+краткоПГ(г).replace("–","-")+(д?" — "+String(д).replace(/[\\/:*?"<>|]/g,"_"):"")+расш;
}
function exportПГXlsx(){
  const г=текущийГрафикПГ(); if(!г) return;
  const ряды=рядыПГ(г); if(!ряды.length){ toast("Позиций по отбору нет"); return; }
  const к=колонкиПГ(г);
  xlsxФайл(имяФайлаПГ(г,".xlsx"),[{имя:"План-график",
    колонки:к.map(x=>({з:x.з,ш:x.ш,тип:x.тип})),
    строки:ряды.map(п=>к.map(x=>x.ф(п)))}]);
  toast("План-график выгружен в Excel: "+ряды.length+" позиций");
}
let _кегльПГ=0;    // 10 пт таблицы плана-графика в Word (кегльТабл/строкаТабл)
function exportПГWord(){
  const г=текущийГрафикПГ(); if(!г) return;
  const ряды=рядыПГ(г); if(!ряды.length){ toast("Позиций по отбору нет"); return; }
  const т=типографика, к=колонкиПГ(г);
  // Ширины — доли от ширины альбомного листа (14003 твипа), месяцы
  // поровну; сумма точно равна ширине (остаток — в наименование).
  const вес=x=>x.мес?1:({"№ п/п":1.1,"Периодичность":2.3,"Количество":2,"Заказывающее подразделение":3.6,
    "Методы контроля (ВИК, УК, АЭ, ТВ)":2.6,"Дата окончания срока службы":2.3,"Примечание":3}[x.з]||0);
  const мес=к.filter(x=>x.мес).length, ед=Math.floor(ШИРИНА_АЛЬБОМ/(мес+14+к.filter(x=>!x.мес).reduce((a,x)=>a+вес(x),0)));
  const ширины=к.map(x=>Math.round(вес(x)*ед));
  const наим=1; ширины[наим]=ШИРИНА_АЛЬБОМ-ширины.reduce((a,b,i)=>a+(i===наим?0:b),0);
  const шапка=к.map(x=>x.з);
  const строки=[к.map((x,i)=>String(i+1))].concat(ряды.map(п=>к.map(x=>{ const з=x.ф(п); return x.тип==="дата"?(з?ru(з):""):т(String(з==null?"":з)); })));
  const д=пгСост.подр&&пгСост.подр!=="—"?(найтиПоИд(live(ПОДР),пгСост.подр)||{}).наименование:"";
  let тело=wP("ПЛАН-ГРАФИК проведения работ",{jc:"center",жирный:true,отступ:false});
  тело+=wP(т(названиеПГ(г)),{jc:"center",жирный:true,отступ:false});
  if(д) тело+=wP(т("Заказывающее подразделение: "+д),{jc:"center",отступ:false});
  тело+=wP("",{отступ:false});
  _кегльПГ=20;
  try{ тело+=wТаблШапкаШир(шапка,ширины,строки,true,false); } finally { _кегльПГ=0; }
  const sectPr='<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape" w:code="9"/>'+
    '<w:pgMar w:top="1134" w:right="1276" w:bottom="1134" w:left="1559" '+
    'w:header="567" w:footer="0" w:gutter="0"/></w:sectPr>';
  wФайлDocx(имяФайлаПГ(г,".docx"), тело, sectPr);
  toast("План-график выгружен в Word: "+ряды.length+" позиций");
}

// ── Новый план-график из Excel — в форме присланных книг ───
// Узнаётся по заголовку: «…по вибродиагностическому / тепловому /
// неразрушающему контролю на ГГГГ г.». Столбцы — по ТЕКСТУ шапки
// (строка с «№ п/п»), месяцы — по названиям; год месяца — из строки над
// шапкой («2026 г.» над ноябрём, как в книге ТК), иначе из заголовка, а
// при переходе с декабря на январь — следующий. Отметка работы — «+».
// Импорт ДОБАВЛЯЕТ новый план-график; такой же (вид и период) уже есть —
// отказ: правка имеющегося — в программе.
async function импортПГВыбран(input){
  const f=input.files&&input.files[0]; input.value=""; if(!f) return;
  try{ импортПГПрименить(await xlsxРазобрать(f), f.name); }
  catch(e){ alert("Не удалось прочитать файл: "+(e&&e.message||e)); }
}
function норм_заг(s){ return String(s==null?"":s).replace(/\s+/g," ").trim(); }
function разборПГ(строки){
  const весьТекст=строки.slice(0,4).map(р=>Object.values(р).join(" ")).join(" ").toLowerCase();
  const вид=/вибродиагност/.test(весьТекст)?"ВД":/теплов/.test(весьТекст)?"ТК":/неразрушающ/.test(весьТекст)?"НК":"";
  if(!вид) throw new Error("в заголовке не нашлось «по вибродиагностическому / тепловому / неразрушающему контролю»");
  // Годы из заголовка: «на 2027 г.» или «на 2026 - 2027 гг.»
  const годыТит=[...((весьТекст.match(/на\s+\d{4}[\s\d\-–гг.]*/)||[""])[0]).matchAll(/(\d{4})/g)].map(m=>+m[1]);
  const iш=строки.findIndex(р=>Object.values(р).some(v=>норм_заг(v)==="№ п/п"));
  if(iш<0) throw new Error("не нашлась строка шапки с «№ п/п»");
  const ш=строки[iш], букв=Object.keys(ш).sort((a,b)=>a.length-b.length||a.localeCompare(b));
  const кол=(тест)=>букв.find(б=>тест(норм_заг(ш[б])));
  const над=iш>0?строки[iш-1]:{};
  // Месяцы слева направо. Год — из строки над шапкой, иначе из заголовка;
  // назад по месяцам (декабрь → январь) без года над шапкой — следующий
  // год, если заголовок на два года, а иначе дальше уже другие столбцы
  // (в книге ВД за отметками идут трудозатраты по тем же месяцам).
  const мес=[], видели=new Set(); let год=годыТит[0]||null, годНад=null, прев=0, стоп=false;
  букв.forEach(б=>{
    if(стоп) return;
    const гн=(String(над[б]||"").match(/(\d{4})/)||[])[1]; if(гн) годНад=+гн;
    const м=МЕСЯЦЫ.findIndex(x=>x.toLowerCase()===норм_заг(ш[б]).toLowerCase());
    if(м<0) return;
    let г=годНад;
    if(!г){
      if(прев && м+1<=прев){ if(годыТит.length>1 && год<годыТит[годыТит.length-1]) год++; else { стоп=true; return; } }
      г=год;
    }
    if(!г) return;
    const ym=г+"-"+String(м+1).padStart(2,"0");
    if(видели.has(ym)){ стоп=true; return; }
    видели.add(ym); прев=м+1; мес.push({б, ym});
  });
  if(!мес.length) throw new Error("в шапке не нашлись столбцы месяцев (или не понятен год — его нет ни в заголовке, ни над месяцами)");
  const к={нп:кол(т=>т==="№ п/п"), наим:кол(т=>/^наименование/i.test(т)), подр:кол(т=>/^заказывающее/i.test(т)),
    кол:кол(т=>т==="Количество"), мет:кол(т=>/^методы контроля/i.test(т)), срок:кол(т=>/^дата окончания срока службы/i.test(т)),
    прим:кол(т=>т==="Примечание"), доля:кол(т=>т==="Периодичность2"), пер:кол(т=>т==="Периодичность"),
    первый:кол(т=>т==="Месяц первого измерения"), труд:кол(т=>т==="Трудозатраты"), прич:кол(т=>т==="Причина"),
    изг:кол(т=>т==="Данные о производителе"), год:кол(т=>т==="Год выпуска"), мощн:кол(т=>/^мощность/i.test(т)),
    подш:кол(т=>/^количество и тип подшипников/i.test(т)), лоп:кол(т=>/^количество лопастей/i.test(т)),
    колёс:кол(т=>т==="Количество рабочих колес"||т==="Количество рабочих колёс"), фэд:кол(т=>/^f вращения вала эд/i.test(т)),
    фрк:кол(т=>/^f вращения вала рабочего/i.test(т)), моточ:кол(т=>/^количество моточасов/i.test(т))};
  if(!к.наим) throw new Error("не нашёлся столбец наименования");
  const позиции=[];
  строки.slice(iш+1).forEach(р=>{
    const нп=норм_заг(р[к.нп]); if(!/^\d+$/.test(нп)) return;
    if(норм_заг(р[к.наим])==="2" && нп==="1") return;    // строка номеров столбцов под шапкой
    const зн=б=>б?норм_заг(р[б]):"";
    const п={нп:+нп, наименование:зн(к.наим), подр:зн(к.подр), количество:зн(к.кол), методы:зн(к.мет),
      примечание:зн(к.прим), месяцы:{}};
    мес.forEach(x=>{ if(зн(x.б)==="+") п.месяцы[x.ym]=""; });
    if(к.срок){ const с=зн(к.срок); п.срок_службы=значениеДаты(с)||с; }
    if(вид==="ВД"){
      let доля=+String(зн(к.доля)).replace(",",".");
      if(!доля){ const x=ПГ_ДОЛИ.find(([,т])=>т===зн(к.пер)); if(x) доля=x[0]; }
      if(доля) п.доля=доля;
      if(+зн(к.первый)) п.мес_первый=+зн(к.первый);
      п.трудозатраты=зн(к.труд); п.причина=зн(к.прич);
      п.тех={изготовитель:зн(к.изг), год_выпуска:зн(к.год), мощность:зн(к.мощн), подшипники:зн(к.подш), лопасти:зн(к.лоп),
        колёса:зн(к.колёс), обороты:зн(к.фэд), обороты_рк:зн(к.фрк), моточасы:зн(к.моточ)};
    }
    позиции.push(п);
  });
  // № п/п повторяются — так в присланной книге ВД 2027 (в трёх строках
  // формула нумерации сбита). Тогда нумерую по порядку строк — это и
  // считает формула «№ п/п» — и называю расхождения в подтверждении.
  const повт=позиции.filter((п,i)=>позиции.findIndex(x=>x.нп===п.нп)!==i).map(п=>п.нп);
  const перенум=[];
  if(повт.length) позиции.forEach((п,i)=>{ if(п.нп!==i+1) перенум.push(п.нп+" → "+(i+1)); п.нп=i+1; });
  return {вид, с:мес[0].ym, по:мес[мес.length-1].ym, позиции, перенум, повт:[...new Set(повт)]};
}
function импортПГПрименить(строки, имяФайла){
  const р=разборПГ(строки);
  const нов={id:newId("ПГ"), вид:р.вид, с:р.с, по:р.по, источник:имяФайла||""};
  if(live(ПГ).some(x=>x.вид===р.вид && x.с===р.с && x.по===р.по)){
    alert("План-график "+краткоПГ(нов)+" уже есть в программе. Загружается только новый; имеющийся правится в программе "+
      "(или удалите его и загрузите заново)."); return; }
  // Подразделение — по точному наименованию карточки (без учёта регистра)
  const поИмени=new Map(live(ПОДР).map(д=>[String(д.наименование||"").trim().toLowerCase(),д.id]));
  const ненайд=new Set();
  р.позиции.forEach(п=>{ if(п.подр){ const id=поИмени.get(п.подр.toLowerCase()); if(id) п.подразделение=id; else ненайд.add(п.подр); } delete п.подр; });
  // ВД: карточка вибрации — имеющаяся с тем же номером И наименованием, иначе новая
  let связано=0, новых=0;
  if(р.вид==="ВД") р.позиции.forEach(п=>{
    const к=live(ВИБК).find(к=>String(к.номер||"").trim()===String(п.нп) && позицияКарты(п.наименование,к));
    if(к){ п.карта_ид=к.id; связано++; } else новых++;
  });
  const работ=р.позиции.reduce((a,п)=>a+Object.keys(п.месяцы).length,0);
  if(!confirm("План-график "+краткоПГ(нов)+" ("+названиеПГ(нов)+"): позиций "+р.позиции.length+", работ (отметок «+») "+работ+"."+
    (р.вид==="ВД"?"\nКарточки учёта вибрации: связано с имеющимися (тот же № и наименование) — "+связано+", будет заведено новых — "+новых+".":"")+
    (р.перенум.length?"\n№ п/п в файле повторяются ("+р.повт.join(", ")+") — пронумеровано по порядку строк: "+р.перенум.slice(0,10).join("; ")+(р.перенум.length>10?" …":""):"")+
    (ненайд.size?"\nПодразделения не нашлись в справочнике и не подставлены: "+[...ненайд].slice(0,10).join("; ")+(ненайд.size>10?" …":""):"")+
    "\n\nЗагрузить? (отменить можно кнопкой «↶ Отменить» внизу)")) return;
  отменаНачать("загрузка плана-графика "+краткоПГ(нов), [REG_ПГ,REG_ПГП,REG_ВИБК]);
  const ts=new Date().toISOString(), место=getWS().код;
  нов.изменено=ts; нов.место=место; ПГ.push(нов);
  р.позиции.forEach(п=>{
    if(р.вид==="ВД" && !п.карта_ид){
      const id=newId("ВК");
      ВИБК.push(Object.assign({id, номер:String(п.нп), подразделение:п.подразделение||"", изменено:ts, место},картаИзПозиции(п.наименование),
        Object.fromEntries(Object.entries(п.тех||{}).filter(([,v])=>v!==""))));
      п.карта_ид=id;
    }
    delete п.тех;
    ПГП.push(Object.assign({id:newId("ПГП"), пг:нов.id, изменено:ts, место},п));
  });
  saveReg(REG_ПГ,ПГ); saveReg(REG_ПГП,ПГП); if(р.вид==="ВД") saveReg(REG_ВИБК,ВИБК);
  отменаЗавершить();
  пгСост.граф=нов.id; пгСост.подр=""; renderПГ();
  toast("Загружен план-график "+краткоПГ(нов)+": "+р.позиции.length+" позиций");
}

// ── Для пульта, задач и календаря ──────────────────────────
function пгМесяца(ym){
  ym=ym||ymТекущий(); let всего=0, сделано=0;
  live(ПГП).forEach(п=>{ if(!(ym in (п.месяцы||{})) || !графикПозиции(п)) return; всего++; if(отметкаПГ(п,ym)) сделано++; });
  return {всего, сделано, осталось:всего-сделано};
}
function пгПросрочено(){
  const тек=ymТекущий(); let n=0;
  live(ПГП).forEach(п=>{ const г=графикПозиции(п); if(!г) return;
    Object.keys(п.месяцы||{}).forEach(ym=>{ if(ym<тек && ym>=г.с && ym<=г.по && !отметкаПГ(п,ym)) n++; }); });
  return n;
}
// Производные задачи «План-график»: срок — уточнённый день, а без него
// последний день месяца (как у ППР).
function задачиПГ(вокне){
  const из=[];
  live(ПГП).forEach(п=>{
    const г=графикПозиции(п); if(!г) return;
    Object.keys(п.месяцы||{}).forEach(ym=>{
      if(ym<г.с || ym>г.по) return;
      const срок=п.месяцы[ym]||ymПосл(ym);
      if(!вокне(срок)) return;
      const о=отметкаПГ(п,ym);
      из.push({ид:"ПГ|"+п.id+"|"+ym, источник:"План-график",
        что:г.вид+": "+имяПозицииПГ(п)+(п.методы?" ("+п.методы+")":""),
        срок, горизонт:"месяц", сделано:!!о, когда:(о&&о.выполнено)||"", куда:"pg"});
    });
  });
  return из;
}
function проверкиПГ(доб, сс){
  графикиПГ().forEach(г=>{
    const поз=позицииГрафика(г.id), к=краткоПГ(г);
    const безПодр=поз.filter(п=>!п.подразделение);
    if(безПодр.length) доб("наблюдение","Планы-графики","План-график "+к+": заказывающее подразделение не указано у "+безПодр.length+" из "+поз.length+" позиций",
      ["Выписка для подразделения соберёт только позиции, где оно указано."]);
    const безИмени=поз.filter(п=>!String(п.наименование||"").trim());
    if(безИмени.length) доб("внимание","Планы-графики","План-график "+к+": наименование не внесено у "+безИмени.length+" позиций",
      безИмени.slice(0,10).map(п=>сс("№ п/п "+п.нп,"pgpos",п.id)));
    if(г.вид==="ВД"){
      const битые=поз.filter(п=>п.карта_ид && !найтиПоИд(live(ВИБК),п.карта_ид));
      if(битые.length) доб("внимание","Планы-графики","План-график "+к+": позиция ссылается на удалённую карточку учёта вибрации — "+битые.length,
        битые.slice(0,10).map(п=>сс("№ "+п.нп+" "+имяПозицииПГ(п),"pgpos",п.id)));
      const несх=поз.filter(п=>п.доля && п.мес_первый && месяцыПоПравилу(п.доля,п.мес_первый,г).join()!==Object.keys(п.месяцы||{}).sort().join());
      if(несх.length) доб("наблюдение","Планы-графики","План-график "+к+": месяцы не сходятся с периодичностью и месяцем первого измерения — "+несх.length,
        несх.slice(0,10).map(п=>сс("№ "+п.нп+" "+имяПозицииПГ(п)+": "+периодичностьПГ(п.доля)+", с "+МЕСЯЦЫ[п.мес_первый-1].toLowerCase(),"pgpos",п.id)));
    }
    const вне=поз.filter(п=>Object.keys(п.месяцы||{}).some(ym=>ym<г.с||ym>г.по));
    if(вне.length) доб("наблюдение","Планы-графики","План-график "+к+": у "+вне.length+" позиций есть месяцы вне периода — их не видно",
      вне.slice(0,10).map(п=>сс("№ "+п.нп+" "+имяПозицииПГ(п),"pgpos",п.id)));
  });
}
