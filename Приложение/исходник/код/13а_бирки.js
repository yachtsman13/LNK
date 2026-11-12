// ── Бирки по результатам поверки (3.84) ───────────────────
// По образцу заказчика («Бирки … 2026», Word): бирка — таблица в рамке
// (снаружи тройная линия, внутри тонкая), сверху наименование с типом
// жирным по центру, ниже две колонки: слева учётный (№ карточки учёта),
// заводской и инвентарный номера, справа даты последней и следующей
// поверки и ответственный за эксплуатацию. Самая маленькая (2 см) —
// одной колонкой, всё строками друг под другом.
//
// Ширина и кегль в образце у каждой бирки свои — подобраны под прибор.
// Поэтому размер — поле карточки прибора (бирка_размер), а список
// размеров и кегль у каждого взяты из образца: самый частый кегль для
// этой ширины. Поле «кегль» карточки его перекрывает.
//
// «Дата следующей поверки» — поле «Действительно до»: в образце у
// дальномера поверка 03.09.2026, следующая 02.09.2027 — ровно дата
// окончания поверки (тот же пример заказчик давал для расчёта срока).
// Если у прибора отмечена калибровка, пишется «калибровки» — этого в
// образце нет, моё решение (иначе бирка называла бы калибровку поверкой).
const БИРКА_РАЗМЕРЫ=[
  {имя:"2 см",  ш:2,    наим:4,  текст:3,  одна:true},
  {имя:"4 см",  ш:4,    наим:4,  текст:4},
  {имя:"4,2 см",ш:4.2,  наим:10, текст:6},
  {имя:"5 см",  ш:5,    наим:9,  текст:8},
  {имя:"6 см",  ш:6,    наим:10, текст:10},
  {имя:"7 см",  ш:7,    наим:10, текст:8},
  {имя:"8,4 см",ш:8.4,  наим:12, текст:9},
  {имя:"9 см",  ш:9,    наим:12, текст:12},
  {имя:"10,5 см",ш:10.5,наим:12, текст:8},
  {имя:"12 см", ш:12,   наим:10, текст:10},
  {имя:"16 см", ш:16,   наим:18, текст:13}];
const БИРКА_ПО_УМОЛЧАНИЮ="4 см";
// Промежуток между бирками на листе — под ножницы. Моё число, в образце
// бирки разбросаны по листу как придётся.
const БИРКА_ЗАЗОР_ММ=3;
function размерБирки(п){
  const v=String((п&&п.бирка_размер)||"").trim();
  return БИРКА_РАЗМЕРЫ.find(р=>р.имя===v)||БИРКА_РАЗМЕРЫ.find(р=>р.имя===БИРКА_ПО_УМОЛЧАНИЮ);
}
function кегльБирки(п){
  const р=размерБирки(п);
  const к=parseFloat(String((п&&п.бирка_кегль)||"").replace(",","."));
  if(!(к>0 && к<=40)) return {наим:р.наим, текст:р.текст};
  return {наим:Math.max(к, к+(р.наим-р.текст)), текст:к};
}
function ответственныйСИ(п){
  const сп=найтиПоИд(live(P), п&&п.ответственный);
  return сп?фамилияИО(сп):String((п&&п.ответственный)||"");
}
// Учётный номер у неотъемлемой части, если свой не вписан, — основного
// прибора (в образце у мер «СО-2» и «СО-3» из одного комплекта номер один).
function учётныйБирки(п){
  if(String(п.карточка||"").trim()) return String(п.карточка).trim();
  const поId={}; live(S).forEach(x=>поId[String(x.id)]=x);
  const к=п.основной_ид?корневойПрибор(п,поId):null;
  return к&&к!==п?String(к.карточка||"").trim():"";
}
function строкиБирки(п){
  const сл=п.калибровка?"калибровки":"поверки";
  return {
    наим: имяПрибораДок(п),
    слева: ["Учетный № "+учётныйБирки(п), "Заводской № "+String(п.зав||"").trim(), "Инвентарный № "+String(п.инв||"").trim()],
    справа: ["Дата последней "+сл+" "+ru(п.поверка_дата||""), "Дата следующей "+сл+" "+ru(п.поверка_до||""),
             "Ответственный за эксплуатацию: "+ответственныйСИ(п)]};
}
// Чего не хватает на бирке — строки остаются, но пустыми.
function пробелыБирки(п){
  const н=[];
  if(!учётныйБирки(п)) н.push("учётный №");
  if(!String(п.зав||"").trim()) н.push("заводской №");
  if(!String(п.инв||"").trim()) н.push("инвентарный №");
  if(!п.поверка_дата) н.push("дата поверки");
  if(!п.поверка_до) н.push("действительно до");
  if(!ответственныйСИ(п)) н.push("ответственный");
  return н;
}
// Сводка пробелов по полям: «инвентарный № — 28 (…)», без перечня каждой бирки.
function сводкаПробеловБирок(список){
  const по={}; список.forEach(п=>пробелыБирки(п).forEach(к=>(по[к]=по[к]||[]).push(п)));
  return Object.keys(по).map(к=>к+" — "+по[к].length+(по[к].length<=3?" ("+по[к].map(п=>имяПрибораДок(п)).join(", ")+")":""));
}
// Кого можно печатать: несписанные. По умолчанию отмечены те, у кого
// есть дата поверки (бирка «по результатам поверки»).
function приборыДляБирок(){
  return приборыВДокументы().slice().sort((a,b)=>{
    const ка=parseFloat(учётныйБирки(a)), кб=parseFloat(учётныйБирки(b));
    if(isFinite(ка)&&isFinite(кб)&&ка!==кб) return ка-кб;
    if(isFinite(ка)!==isFinite(кб)) return isFinite(ка)?-1:1;
    if(!!a.основной_ид!==!!b.основной_ид) return a.основной_ид?1:-1;   // часть — после своего прибора
    return String(a.наименование||"").localeCompare(String(b.наименование||""),"ru"); });
}
let _бирки={выбор:null};
function выбранныеБирки(){
  const все=приборыДляБирок();
  if(!_бирки.выбор) return все.filter(п=>!!п.поверка_дата);
  const н=new Set(_бирки.выбор.map(String));
  return все.filter(п=>н.has(String(п.id)));
}
function бирки(ид){
  if(curView!=="labels" && !уходСФормы()) return;
  _бирки.выбор=ид?ид.map(String):null; renderБирки();
}
function биркаВыбор(id, вкл){
  const н=new Set(выбранныеБирки().map(п=>String(п.id)));
  if(вкл) н.add(String(id)); else н.delete(String(id));
  _бирки.выбор=[...н]; renderБирки();
}
function биркиВсе(как){
  _бирки.выбор=как==="все"?приборыДляБирок().map(п=>String(п.id)):как==="нет"?[]:null; renderБирки();
}
// Правка размера, кегля и ответственного прямо в списке — в саму
// карточку, с теми же отметками, что обычное сохранение (иначе правка не
// ушла бы в обмен).
function setБирка(id, поле, v){
  const п=S.find(x=>String(x.id)===String(id)); if(!п) return;
  if(["бирка_размер","бирка_кегль","ответственный"].indexOf(поле)<0) return;
  if(String(п[поле]||"")===String(v||"")) return;
  п[поле]=v; п.изменено=new Date().toISOString(); п.место=getWS().код;
  saveReg(REG_SI,S); renderБирки();
}
function биркаHTML(п){
  const р=размерБирки(п), к=кегльБирки(п), с=строкиБирки(п);
  const ст=`width:${р.ш}cm;font-size:${к.текст}pt`;
  const стр=a=>a.map(x=>`<div>${esc(типографика(x))}</div>`).join("");
  if(р.одна) return `<table class="бирка" data-id="${esc(п.id)}" style="${ст}"><tr><td>
      <div class="биркаимя" style="font-size:${к.наим}pt">${esc(типографика(с.наим))}</div>${стр(с.слева.concat(с.справа))}</td></tr></table>`;
  return `<table class="бирка" data-id="${esc(п.id)}" style="${ст}"><colgroup><col style="width:40%"><col style="width:60%"></colgroup>
    <tr><td colspan="2"><div class="биркаимя" style="font-size:${к.наим}pt">${esc(типографика(с.наим))}</div></td></tr>
    <tr><td>${стр(с.слева)}</td><td>${стр(с.справа)}</td></tr></table>`;
}
function renderБирки(){
  if(curView!=="labels" && !уходСФормы()) return;
  curView="labels"; query="";
  document.querySelectorAll(".nav").forEach(n=>n.classList.toggle("on",n.dataset.view==="si"));
  const все=приборыДляБирок(), выбр=выбранныеБирки(), н=new Set(выбр.map(п=>String(п.id)));
  const пробелы=выбр.filter(п=>пробелыБирки(п).length);
  const спец=live(P).slice().sort((a,b)=>String(a.фио||"").localeCompare(String(b.фио||""),"ru"));
  document.getElementById("main").innerHTML=`
    <div class="head noprint"><button class="btn" onclick="switchView('si')">← К приборам</button>
      <h2>Бирки по результатам поверки</h2><span class="sub">выбрано ${выбр.length} из ${все.length}</span>
      <div class="spring"></div>
      <button class="btn" onclick="биркиWord()">⤓ Word</button>
      <button class="btn primary" onclick="window.print()">⎙ Печать / PDF</button></div>
    <div class="note noprint">Бирки — по вашему образцу: наименование с типом, учётный (№ карточки учёта),
      заводской и инвентарный номера, даты последней и следующей («действительно до») поверки и ответственный за
      эксплуатацию. Размер и кегль у каждой бирки свои, как в образце, — выбираются здесь же (пишется в карточку
      прибора); пустой размер — «${БИРКА_ПО_УМОЛЧАНИЮ}». Бирки ставятся на лист А4 рядами с зазором ${БИРКА_ЗАЗОР_ММ} мм
      под ножницы. В печати — масштаб 100 %. По умолчанию отмечены приборы с датой поверки; списанные не показываются.
      Бирки на стеллаж — отдельно: их образца ещё нет.</div>
    ${пробелы.length?`<div class="note noprint" style="color:var(--warn)">На выбранных бирках не заполнено (строка
      напечатается пустой, её можно дописать от руки): ${сводкаПробеловБирок(выбр).map(esc).join("; ")}.</div>`:""}
    <details class="noprint биркивыбор" ${_бирки.выбор?"open":""}><summary>Какие бирки печатать и какого размера</summary>
      <div style="margin:6px 0"><button class="btn small" onclick="биркиВсе('все')">все</button>
        <button class="btn small" onclick="биркиВсе('поверка')">с датой поверки</button>
        <button class="btn small" onclick="биркиВсе('нет')">ни одной</button></div>
      <div class="tblwrap"><table data-отбор="нет"><thead><tr><th></th><th>Учётный №</th><th>Прибор</th><th>Зав. №</th>
        <th>Поверка</th><th>Ответственный</th><th>Размер</th><th>Кегль, пт</th></tr></thead><tbody>${все.map(п=>`<tr>
        <td><input type="checkbox" ${н.has(String(п.id))?"checked":""} onchange="биркаВыбор('${esc(п.id)}',this.checked)"></td>
        <td class="num">${esc(учётныйБирки(п))||"—"}</td><td class="obj">${esc(имяПрибораДок(п))}${п.основной_ид?' <span class="muted">(часть)</span>':""}</td>
        <td class="num">${esc(п.зав)||"—"}</td><td class="num">${п.поверка_дата?ru(п.поверка_дата)+" — "+ru(п.поверка_до):"—"}</td>
        <td><select onchange="setБирка('${esc(п.id)}','ответственный',this.value)"><option value="">—</option>${спец.map(с=>
          `<option value="${esc(с.id)}" ${тотЖеИд(с.id,п.ответственный)?"selected":""}>${esc(фамилияИО(с))}</option>`).join("")}</select></td>
        <td><select onchange="setБирка('${esc(п.id)}','бирка_размер',this.value)"><option value="">${esc(БИРКА_ПО_УМОЛЧАНИЮ)} (по умолч.)</option>${
          БИРКА_РАЗМЕРЫ.map(р=>`<option ${String(п.бирка_размер||"")===р.имя?"selected":""}>${esc(р.имя)}</option>`).join("")}</select></td>
        <td><input class="бирккегль" value="${esc(п.бирка_кегль||"")}" placeholder="${кегльБирки(Object.assign({},п,{бирка_кегль:""})).текст}"
          onchange="setБирка('${esc(п.id)}','бирка_кегль',this.value.trim())"></td></tr>`).join("")}</tbody></table></div></details>
    ${выбр.length?`<div class="бирклист">${выбр.map(биркаHTML).join("")}</div>`
      :`<div class="note">Бирок не выбрано.</div>`}`;
}
// ── Word ──────────────────────────────────────────────────
// Бирки рядами: каждый ряд — таблица без рамок, в её ячейках — таблицы
// бирок (вложенные таблицы), между ними — пустые ячейки-зазоры. Так Word
// сам переносит ряды на следующую страницу и ничего не налезает друг на
// друга (в образце бирки — «плавающие» таблицы, у разных программ они
// разъезжаются). Ряд не разрывается между страницами (cantSplit).
const БИРКА_ТВИП_СМ=567, БИРКИ_ПОЛЕ=567, БИРКИ_ШИРИНА=11906-2*БИРКИ_ПОЛЕ;
function рядыБирок(список){
  const зазор=Math.round(БИРКА_ЗАЗОР_ММ/10*БИРКА_ТВИП_СМ), ряды=[]; let ряд=[], ш=0;
  список.forEach(п=>{
    const w=Math.round(размерБирки(п).ш*БИРКА_ТВИП_СМ);
    if(ряд.length && ш+зазор+w>БИРКИ_ШИРИНА){ ряды.push(ряд); ряд=[]; ш=0; }
    ш+=(ряд.length?зазор:0)+w; ряд.push(п);
  });
  if(ряд.length) ряды.push(ряд);
  return {ряды, зазор};
}
function wБиркаАбзац(t, кегль, вид){
  вид=вид||{};
  const rp='<w:rPr>'+ФОНТ+(вид.жирный?'<w:b/>':'')+'<w:sz w:val="'+Math.round(кегль*2)+'"/><w:szCs w:val="'+Math.round(кегль*2)+'"/></w:rPr>';
  return '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'+(вид.jc?'<w:jc w:val="'+вид.jc+'"/>':'')+rp+'</w:pPr>'+
    (t?'<w:r>'+rp+'<w:t xml:space="preserve">'+XT(типографика(t))+'</w:t></w:r>':'')+'</w:p>';
}
function wБирка(п){
  const р=размерБирки(п), к=кегльБирки(п), с=строкиБирки(п);
  const w=Math.round(р.ш*БИРКА_ТВИП_СМ), мал=р.ш<=4.2, поле=мал?28:57;
  const лин=(вид,sz)=>'w:val="'+вид+'" w:sz="'+sz+'" w:space="0" w:color="auto"';
  const рамка='<w:tblBorders><w:top '+лин("thinThickThinSmallGap",24)+'/><w:left '+лин("thinThickThinSmallGap",24)+'/>'+
    '<w:bottom '+лин("thinThickThinSmallGap",24)+'/><w:right '+лин("thinThickThinSmallGap",24)+'/>'+
    '<w:insideH '+лин("single",4)+'/><w:insideV '+лин("single",4)+'/></w:tblBorders>';
  const шл=р.одна?w:Math.round(w*0.4), шп=w-шл;
  const tblPr='<w:tblPr><w:tblW w:w="'+w+'" w:type="dxa"/>'+рамка+'<w:tblLayout w:type="fixed"/>'+
    '<w:tblCellMar><w:left w:w="'+поле+'" w:type="dxa"/><w:right w:w="'+поле+'" w:type="dxa"/></w:tblCellMar></w:tblPr>';
  const яч=(ш,содерж,span)=>'<w:tc><w:tcPr><w:tcW w:w="'+ш+'" w:type="dxa"/>'+(span?'<w:gridSpan w:val="'+span+'"/>':'')+
    '<w:vAlign w:val="center"/></w:tcPr>'+содерж+'</w:tc>';
  const стр=a=>a.map(x=>wБиркаАбзац(x,к.текст)).join("");
  const имя=wБиркаАбзац(с.наим,к.наим,{жирный:true,jc:"center"});
  if(р.одна) return '<w:tbl>'+tblPr+'<w:tblGrid><w:gridCol w:w="'+w+'"/></w:tblGrid>'+
    '<w:tr><w:trPr><w:cantSplit/></w:trPr>'+яч(w, имя+стр(с.слева.concat(с.справа)))+'</w:tr></w:tbl>';
  return '<w:tbl>'+tblPr+'<w:tblGrid><w:gridCol w:w="'+шл+'"/><w:gridCol w:w="'+шп+'"/></w:tblGrid>'+
    '<w:tr><w:trPr><w:cantSplit/></w:trPr>'+яч(w, имя, 2)+'</w:tr>'+
    '<w:tr><w:trPr><w:cantSplit/></w:trPr>'+яч(шл, стр(с.слева))+яч(шп, стр(с.справа))+'</w:tr></w:tbl>';
}
function wРядБирок(ряд, зазор){
  const ширины=[]; ряд.forEach((п,i)=>{ if(i) ширины.push(зазор); ширины.push(Math.round(размерБирки(п).ш*БИРКА_ТВИП_СМ)); });
  const всего=ширины.reduce((a,b)=>a+b,0);
  const пусто='<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>';
  const нет='<w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/>'+
    '<w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>';
  let яч="", к=0;
  ряд.forEach((п,i)=>{
    if(i) яч+='<w:tc><w:tcPr><w:tcW w:w="'+ширины[к++]+'" w:type="dxa"/></w:tcPr>'+пусто+'</w:tc>';
    яч+='<w:tc><w:tcPr><w:tcW w:w="'+ширины[к++]+'" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>'+wБирка(п)+пусто+'</w:tc>';
  });
  return '<w:tbl><w:tblPr><w:tblW w:w="'+всего+'" w:type="dxa"/>'+нет+'<w:tblLayout w:type="fixed"/>'+
    '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'+
    '<w:tblGrid>'+ширины.map(x=>'<w:gridCol w:w="'+x+'"/>').join("")+'</w:tblGrid>'+
    '<w:tr><w:trPr><w:cantSplit/></w:trPr>'+яч+'</w:tr></w:tbl>';
}
function биркиWord(){
  const выбр=выбранныеБирки();
  if(!выбр.length){ toast("Бирок не выбрано — файл не создан"); return; }
  const {ряды, зазор}=рядыБирок(выбр);
  // Между рядами — абзац высотой в зазор: и промежуток под ножницы, и
  // без него Word склеил бы соседние таблицы рядов в одну.
  const между='<w:p><w:pPr><w:spacing w:after="0" w:line="'+зазор+'" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>';
  const тело=ряды.map(р=>wРядБирок(р,зазор)).join(между)+между;
  const sectPr='<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="'+БИРКИ_ПОЛЕ+'" w:right="'+БИРКИ_ПОЛЕ+
    '" w:bottom="'+БИРКИ_ПОЛЕ+'" w:left="'+БИРКИ_ПОЛЕ+'" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>';
  wФайлDocx("Бирки по результатам поверки "+сегодня()+".docx", тело, sectPr);
  toast("Бирки выгружены: "+выбр.length);
}
