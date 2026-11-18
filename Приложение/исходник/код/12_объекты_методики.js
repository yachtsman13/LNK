// ── Объекты ───────────────────────────────────────────────
function рядыПодр(){
  const q=(query||"").toLowerCase();
  return live(ПОДР).filter(x=>!q||((x.наименование||"")+(x.эксп_организация||"")).toLowerCase().includes(q))
    .sort((a,b)=>String(a.наименование||"").localeCompare(String(b.наименование||""),"ru"));
}
function renderPodr(){
  const rows=рядыПодр();
  const body=rows.length?rows.map(x=>{
    const к=контрактыКарточки(x).slice().sort((a,b)=>String(b.год).localeCompare(String(a.год)));
    return `<tr data-id="${esc(x.id)}" onclick="editReg('podr','${esc(x.id)}')">
      <td class="obj"><b>${esc(x.наименование)}</b></td>
      <td>${esc(x.тип_контракта)||"<span class='muted'>тип не задан</span>"}</td>
      <td class="obj">${esc(x.эксп_организация)||"<span class='muted'>—</span>"}</td>
      <td class="obj">${esc(x.эксп_фио)||"<span class='muted'>—</span>"}</td>
      <td>${к.length?к.slice(0,3).map(y=>`<span class="pill">${esc(y.год)}</span>`).join("")+
            (к.length>3?`<span class="muted"> +${к.length-3}</span>`:"")
          :"<span class='st new'>контрактов нет</span>"}</td>
    </tr>`;}).join(""):`<tr><td colspan="5" class="empty">Подразделений пока нет.</td></tr>`;
  document.getElementById("main").innerHTML=`
    ${вкладкиРаздела("podr", [["список","Подразделения"],["сводка","Сводка работ"]])}
    <div data-вклр="список">
    ${printHead("Заказывающие подразделения","карточек "+rows.length)}
    <div class="head"><h2>Подразделения</h2><span class="sub">${rows.length}</span>
      <div class="search"><span class="ic">⌕</span><input id="q" placeholder="Поиск по названию…" value="${esc(query)}"></div>
      <div class="spring"></div>
      <button class="btn primary" onclick="editReg('podr','')">+ Добавить</button>
      <button class="btn" onclick="exportXlsx()">▤ Excel</button>
      <button class="btn" onclick="выгрузкаДляПравки('podr')" title="Выгрузить текущие карточки в Excel, поправить и загрузить обратно тем же импортом">⤓ Выгрузить для правки</button>
      <button class="btn" onclick="шаблонИмпортаXlsx('podr')">⤓ Шаблон для импорта</button>
      <button class="btn" onclick="document.getElementById('imp_podr').click()">⬆ Импорт из Excel</button>
      <input type="file" id="imp_podr" accept=".xlsx" style="display:none" onchange="импортФайлВыбран('podr',this)">
      <button class="btn" onclick="window.print()">⎙ PDF</button></div>
    <div class="note справ noprint">Карточка нужна акту: из неё берутся тип контракта, реквизиты
      контракта на нужный год и подписант со стороны экспертной организации.
      Контракт свой на каждый год — так у вас и заведено.</div>
    <div class="card"><div class="tblwrap"><table data-список="podr">
      <thead><tr><th>Подразделение</th><th style="width:180px">Тип контракта</th>
        <th style="width:220px">Экспертная организация</th><th style="width:150px">Руководитель</th>
        <th style="width:150px">Контракты</th></tr></thead>
      <tbody>${body}</tbody></table></div></div></div>
    <div data-вклр="сводка">${блокСводкиПодр()}</div>`;
  показатьВкладкиРаздела("podr");
  wireSearch(renderPodr); инициализироватьКрестикиПоиска();
}
// Сводка работ по заказывающим подразделениям за период — для сверки с
// контрактами. Подразделение записи — поле записи журнала, как оно
// вписано; если оно пустое — подразделение из карточки объекта. Ни того,
// ни другого — строка «(подразделение не указано)». Период и единицы —
// те же, что у «Нагрузки»: дата контроля, работа = метод в записи.
// Отчёт внутренний: образца формы нет.
const сводкаПодр={год:new Date().getFullYear(), период:"год"};
function подразделениеЗаписи(z){
  const т=String(z.подразделение||"").trim(); if(т) return т;
  const o=найтиПоИд(O,z.объект_ид); const п=o?найтиПоИд(ПОДР,o.подразделение):null;
  return п?String(п.наименование||"").trim():"";
}
function сводкаПодрРасчёт(год,период){
  const [с,по]=периодНагрузки(год,период);
  const ряды={}, методы=new Set(); let безДаты=0;
  liveJ().forEach(z=>{
    if(!z.дата){ безДаты++; return; }
    if(z.дата<с || z.дата>по) return;
    const п=подразделениеЗаписи(z)||"(подразделение не указано)";
    const р=ряды[п]=ряды[п]||{п, записей:0, объекты:new Set(), акты:new Set(), поМетодам:{}, работ:0};
    р.записей++;
    р.объекты.add(String(z.объект_ид!=null&&z.объект_ид!==""?z.объект_ид:(z.объект||"")));
    if(String(z.акт||"").trim()) р.акты.add(String(z.акт).trim());
    methodsOf(z).forEach(м=>{ р.поМетодам[м]=(р.поМетодам[м]||0)+1; р.работ++; методы.add(м); });
  });
  const список=Object.values(ряды).sort((a,b)=>
    (a.п.startsWith("(")?1:0)-(b.п.startsWith("(")?1:0) || a.п.localeCompare(b.п,"ru"));
  const порядок=METHODS.filter(м=>методы.has(м));
  return {ряды:список, методы:порядок, безДаты, с, по};
}
function setСводкаПодр(поле,знач){ сводкаПодр[поле]=поле==="год"?+знач:знач; renderPodr(); }
function блокСводкиПодр(){
  const св=сводкаПодрРасчёт(сводкаПодр.год, сводкаПодр.период);
  const опц=[["год","весь год"],["к1","I квартал"],["к2","II квартал"],["к3","III квартал"],["к4","IV квартал"]]
    .concat(МЕСЯЦЫ.map((м,i)=>["м"+(i+1),м]));
  const подп=подписьПериодаНагрузки(сводкаПодр.год,сводкаПодр.период);
  const итог=к=>св.ряды.reduce((n,р)=>n+(typeof р[к]==="number"?р[к]:р[к].size),0);
  const все=св.ряды.every(р=>р.п==="(подразделение не указано)");
  const тело=св.ряды.length?св.ряды.map(р=>`<tr><td class="obj">${р.п.startsWith("(")?`<span class="muted">${esc(р.п)}</span>`:esc(р.п)}</td>
      <td class="num">${р.записей}</td><td class="num">${р.объекты.size}</td><td class="num">${р.акты.size}</td>
      ${св.методы.map(м=>`<td class="num">${р.поМетодам[м]||"<span class='muted'>—</span>"}</td>`).join("")}
      <td class="num"><b>${р.работ}</b></td></tr>`).join("")
    :`<tr><td colspan="${св.методы.length+5}" class="empty">За ${esc(подп)} записей в журнале нет.</td></tr>`;
  return `<div class="card noprint" style="margin-top:14px">
    <div class="head"><h2>Сводка работ по подразделениям</h2><span class="sub">${esc(подп)}</span>
      <div class="spring"></div>
      ${годыНагрузки().map(г=>`<button class="btn ${г===сводкаПодр.год?"on":""}" onclick="setСводкаПодр('год',${г})">${г}</button>`).join("")}
      <select onchange="setСводкаПодр('период',this.value)" style="margin-left:6px">
        ${опц.map(([k,t])=>`<option value="${k}"${k===сводкаПодр.период?" selected":""}>${esc(t)}</option>`).join("")}</select>
      <button class="btn" onclick="exportСводкаПодрXlsx()">⤓ Excel</button></div>
    <div class="tblwrap"><table>
      <thead><tr><th>Подразделение</th><th class="num">Записей</th><th class="num">Объектов</th><th class="num">Актов</th>
        ${св.методы.map(м=>`<th class="num">${esc(м)}</th>`).join("")}<th class="num">Всего работ</th></tr></thead>
      <tbody>${тело}</tbody>
      ${св.ряды.length>1?`<tfoot><tr><td><b>Итого</b></td><td class="num"><b>${итог("записей")}</b></td><td></td><td></td>
        ${св.методы.map(м=>`<td class="num"><b>${св.ряды.reduce((n,р)=>n+(р.поМетодам[м]||0),0)}</b></td>`).join("")}
        <td class="num"><b>${итог("работ")}</b></td></tr></tfoot>`:""}</table></div>
    <div class="note справ" style="margin:8px 12px 12px">Подразделение — как вписано в записи журнала, а если там пусто — из
      карточки объекта. Работа — метод в записи; объекты и акты считаются без повторов. Период — по дате контроля
      (${ru(св.с)} — ${ru(св.по)}).${св.безДаты?` Записей без даты — ${св.безДаты}, в период не попадают.`:""}
      ${все&&св.ряды.length?`<br><b>Сейчас подразделение не указано ни в одной записи и ни в одной карточке объекта</b> —
        поэтому всё в одной строке. Проставьте подразделение в карточках объектов (оно само подставится и в акт),
        и сводка разложится.`:""} Отчёт внутренний — образца формы нет.</div>
  </div>`;
}
function листСводкиПодр(год, период){
  const св=сводкаПодрРасчёт(год, период);
  const колонки=[{з:"Подразделение",ш:36},{з:"Записей",ш:10,тип:"число"},{з:"Объектов",ш:10,тип:"число"},{з:"Актов",ш:10,тип:"число"}]
    .concat(св.методы.map(м=>({з:м,ш:8,тип:"число"}))).concat([{з:"Всего работ",ш:12,тип:"число"}]);
  const строки=св.ряды.map(р=>[р.п,р.записей,р.объекты.size,р.акты.size].concat(св.методы.map(м=>р.поМетодам[м]||0)).concat([р.работ]));
  строки.push(["Итого",св.ряды.reduce((n,р)=>n+р.записей,0),"",""].concat(св.методы.map(м=>св.ряды.reduce((n,р)=>n+(р.поМетодам[м]||0),0)))
    .concat([св.ряды.reduce((n,р)=>n+р.работ,0)]));
  if(!св.ряды.length) строки.length=0;
  return {имя:"Сводка", колонки, строки, пусто:!св.ряды.length};
}
function exportСводкаПодрXlsx(){
  const л=листСводкиПодр(сводкаПодр.год, сводкаПодр.период);
  if(л.пусто){ toast("За этот период записей в журнале нет — выгружать нечего"); return; }
  const подп=подписьПериодаНагрузки(сводкаПодр.год,сводкаПодр.период);
  xlsxФайл("Сводка по подразделениям "+подп.replace(/ г\.$/,"")+".xlsx",[л]);
  toast("Сводка за "+подп+" выгружена");
}
// Раздел 9 паспорта ЛНК: перечень нормативных технических и методических
// документов. Форма — по вашему бланку паспорта (Форма 3 СДАНК-01-2020).
//
// Актуализация: поле «проверено» — дата, когда в последний раз сверяли,
// что документ не отменён и не заменён новым изданием. Своего срока
// действия нормативные документы не имеют — это не поверка, а ваша
// внутренняя процедура. Один год — принятая практика (так же считает
// «Проверка данных»): без даты — «внимание» (могли не проверять),
// проверка старше года — «ошибка» (пора пересмотреть).
function ндАктуальноДо(дата){
  if(!дата) return "";
  const d=new Date(дата+"T00:00:00");
  d.setFullYear(d.getFullYear()+1);
  return d.toISOString().slice(0,10);
}
function ндСтатус(x){
  if(!x.проверено) return ["new","не проверялся"];
  const до=ндАктуальноДо(x.проверено);
  if(до<сегодня()) return ["bad","проверен "+ru(x.проверено)];
  return ["ok","проверен "+ru(x.проверено)];
}
function рядыНд(){
  const q=(query||"").toLowerCase();
  return live(НД).filter(x=>!q||((x.наименование||"")+(x.обозначение||"")).toLowerCase().includes(q))
    .filter(x=>!ндФильтрМетод || x.метод===ндФильтрМетод)
    .sort((a,b)=>String(a.наименование||"").localeCompare(String(b.наименование||""),"ru"));
}
function renderNd(){
  const rows=рядыНд();
  const body=rows.length?rows.map(x=>{const[c,l]=ндСтатус(x);return`
    <tr data-id="${esc(x.id)}" onclick="editReg('nd','${esc(x.id)}')">
      <td class="obj">${ссылкаНД(x,esc(x.наименование))}${x.файл?` <span class="muted" title="файл: ${esc(x.файл)}">📄</span>`:""}
        ${x.метод?`<div class="muted" style="font-size:11.5px">${esc(x.метод)}</div>`:""}</td>
      <td class="num">${esc(x.обозначение)||"—"}</td>
      <td>${esc(x.издатель)||"<span class='muted'>—</span>"}</td>
      <td class="num">${esc(x.место_год)||"—"}</td>
      <td><span class="st ${c}">${esc(l)}</span></td>
    </tr>`;}).join(""):`<tr><td colspan="5" class="empty">Документов пока нет.</td></tr>`;
  document.getElementById("main").innerHTML=`
    ${printHead("Нормативные технические и методические документы","документов "+rows.length)}
    <div class="head"><h2>Нормативные документы</h2><span class="sub">${rows.length}</span>
      <div class="search"><span class="ic">⌕</span><input id="q" placeholder="Поиск по названию, обозначению…" value="${esc(query)}"></div>
      <select id="ндМетод" onchange="ндФильтрМетод=this.value;renderNd()" style="margin-left:8px">
        <option value="">Все методы</option>
        ${METHODS.map(м=>`<option value="${esc(м)}" ${ндФильтрМетод===м?"selected":""}>${esc(м)}</option>`).join("")}
      </select>
      <div class="spring"></div>
      <button class="btn primary" onclick="editReg('nd','')">+ Добавить</button>
      <button class="btn" onclick="exportXlsx()">▤ Excel</button>
      <button class="btn" onclick="выгрузкаДляПравки('nd')" title="Выгрузить текущие карточки в Excel, поправить и загрузить обратно тем же импортом">⤓ Выгрузить для правки</button>
      <button class="btn" onclick="шаблонИмпортаXlsx('nd')">⤓ Шаблон для импорта</button>
      <button class="btn" onclick="document.getElementById('imp_nd').click()">⬆ Импорт из Excel</button>
      <input type="file" id="imp_nd" accept=".xlsx" style="display:none" onchange="импортФайлВыбран('nd',this)">
      <button class="btn" onclick="window.print()">⎙ PDF</button></div>
    <div class="note справ noprint">Идёт в раздел 9 паспорта ЛНК как есть — список ваш, я его не
      выдумываю и не переношу из образцов. Колонка «Актуальность» в паспорт не идёт (в Форме 3
      её нет), она только для внутреннего контроля — раз в год стоит сверять список с действующими
      изданиями. Поле «Метод» — необязательное, для фильтра здесь же: у документов общего
      назначения (законы, ФНП, ГОСТы на термины) метод обычно не проставлен, это не ошибка.</div>
    <div class="card"><div class="tblwrap"><table data-список="nd">
      <thead><tr><th>Наименование</th><th style="width:120px">Обозначение</th>
        <th style="width:220px">Издательство / разработчик</th><th style="width:110px">Место, год</th>
        <th style="width:140px">Актуальность</th></tr></thead>
      <tbody>${body}</tbody></table></div></div>`;
  wireSearch(renderNd);
}
// Раздел 10 паспорта ЛНК: помещения, используемые для проведения НК.
function рядыПом(){
  const q=(query||"").toLowerCase();
  return live(ПОМ).filter(x=>!q||(x.идентификация||"").toLowerCase().includes(q))
    .sort((a,b)=>String(a.идентификация||"").localeCompare(String(b.идентификация||""),"ru"));
}
function renderPremises(){
  const rows=рядыПом();
  const body=rows.length?rows.map(x=>`
    <tr data-id="${esc(x.id)}" onclick="editReg('pom','${esc(x.id)}')">
      <td class="obj">${esc(x.идентификация)}</td>
      <td>${esc(x.назначение)||"<span class='muted'>—</span>"}</td>
      <td class="num">${esc(x.площадь)||"—"}</td>
      <td>${esc(x.сети)||"<span class='muted'>—</span>"}</td>
    </tr>`).join(""):`<tr><td colspan="4" class="empty">Помещений пока нет.</td></tr>`;
  document.getElementById("main").innerHTML=`
    ${printHead("Помещения, используемые для проведения НК","помещений "+rows.length)}
    ${вкладкиЛаб("premises")}
    <div class="head"><h2>Помещения</h2><span class="sub">${rows.length}</span>
      <div class="search"><span class="ic">⌕</span><input id="q" placeholder="Поиск по номеру, наименованию…" value="${esc(query)}"></div>
      <div class="spring"></div>
      <button class="btn primary" onclick="editReg('pom','')">+ Добавить</button>
      <button class="btn" onclick="exportXlsx()">▤ Excel</button>
      <button class="btn" onclick="выгрузкаДляПравки('pom')" title="Выгрузить текущие карточки в Excel, поправить и загрузить обратно тем же импортом">⤓ Выгрузить для правки</button>
      <button class="btn" onclick="шаблонИмпортаXlsx('pom')">⤓ Шаблон для импорта</button>
      <button class="btn" onclick="document.getElementById('imp_pom').click()">⬆ Импорт из Excel</button>
      <input type="file" id="imp_pom" accept=".xlsx" style="display:none" onchange="импортФайлВыбран('pom',this)">
      <button class="btn" onclick="window.print()">⎙ PDF</button></div>
    <div class="note справ noprint">Идёт в раздел 10 паспорта ЛНК.</div>
    <div class="card"><div class="tblwrap"><table data-список="pom">
      <thead><tr><th>Идентификация</th><th style="width:220px">Назначение</th>
        <th style="width:90px">Площадь</th><th style="width:220px">Инженерные сети</th></tr></thead>
      <tbody>${body}</tbody></table></div></div>`;
  wireSearch(renderPremises); инициализироватьКрестикиПоиска();
}
// ── Методики контроля: стандартный набор приборов и норм. документов
// по методу ──────────────────────────────────────────────────────
// Не справочник записей (CRUD не нужен — методов ровно 8, весь список
// в METHODS), а фиксированный набор ссылок на уже существующие карточки
// приборов и норм. документов. Хранится тем же приёмом, что и карточка
// лаборатории (getЛаб/saveЛаб): один плоский объект плюс «изменено» —
// не отдельный REGS-справочник.
function getМетодики(){
  let сырое=null;
  try{ const r=ХР.getItem("lnk_metodiki"); if(r) сырое=JSON.parse(r); }catch(e){}
  return нормМетодики(сырое);
}
function нормМетодики(сырое){
  const пусто={}; METHODS.forEach(м=>{ пусто[м]={приборы:[],нд:[]}; });
  let рез=Object.assign({},пусто);
  if(сырое) рез=JSON.parse(JSON.stringify(Object.assign({},пусто,сырое)));
  METHODS.forEach(м=>{ if(!рез[м]||typeof рез[м]!=="object") рез[м]={приборы:[],нд:[]};
    рез[м].приборы=рез[м].приборы||[]; рез[м].нд=рез[м].нд||[]; });
  return рез;
}
// какЕсть — как у saveЛаб: набор пришёл с другого места или это починка
// при загрузке, своей отметкой времени и местом не перебиваю.
function saveМетодики(мт, какЕсть){
  мт=Object.assign({},мт);
  if(!какЕсть){ мт.изменено=new Date().toISOString(); мт.место=getWS().код; if(ктоЯ()) мт.кто=ктоЯ(); else delete мт.кто; }
  одиночноеПередЗаписью(REG_МЕТОДИКИ, мт);
  try{ ХР.setItem("lnk_metodiki", JSON.stringify(мт)); }catch(e){ переполнение(e); }
  одиночноеПослеЗаписи(REG_МЕТОДИКИ);
  копияПоСлучаю();
}
// Сортировка списка для чекбоксов: карточки с совпадающим полем «метод»
// (оно уже есть у приборов и у норм. документов) — сверху, остальные —
// ниже по алфавиту. Не фильтрует, а только облегчает поиск в длинном
// списке — прибор без размеченного метода вполне может входить в набор.
function методСортировка(список, м){
  return список.slice().sort((a,b)=>{
    const ам=a.метод===м?0:1, бм=b.метод===м?0:1;
    if(ам!==бм) return ам-бм;
    return String(a.наименование||"").localeCompare(String(b.наименование||""),"ru");
  });
}
function списокМетодЧекбоксы(список, выбранные, класс, показать){
  if(!список.length) return `<div class="note">Список пуст.</div>`;
  // Сравнение через тотЖеИд, а не indexOf: отмеченные приборы лежат в
  // наборе строками (значение чекбокса), а id прибора из книги — число.
  // Строгий indexOf их не равнял, поэтому галочка при повторном открытии
  // пропадала, а следующее сохранение выбрасывало прибор из набора.
  const отм=выбранные||[];
  return список.map(x=>`<label class="мет-стр"><input type="checkbox" class="${класс}" value="${esc(x.id)}"
    ${отм.some(v=>тотЖеИд(v,x.id))?"checked":""}> ${показать(x)}</label>`).join("");
}
// Текстовый фильтр над списком чекбоксов — общий и для «Методик
// контроля», и для панели прибор/НД в записи журнала.
function фильтрСписка(el, contId){
  const q=(el.value||"").toLowerCase();
  document.querySelectorAll("#"+contId+" .мет-стр").forEach(l=>{
    l.style.display = !q || l.textContent.toLowerCase().includes(q) ? "" : "none";
  });
}
function renderMetod(){
  формаСохранена();
  curView="metod";
  document.querySelectorAll(".nav").forEach(n=>n.classList.toggle("on",n.dataset.view==="metod"));
  const мт=getМетодики();
  const body=METHODS.map(м=>{
    const н=мт[м]||{приборы:[],нд:[]};
    const сд=сданкМетода(м), вОбл=методВОбласти(м,getЛаб().область_методы);
    return `<tr onclick="editMetodika('${esc(м)}')">
      <td class="obj"><b>${esc(м)}</b></td>
      <td>${сд?esc(сд.н+". "+сд.т):(м===МЕТОД_БЕЗ_АТТЕСТАЦИИ?"<span class='muted'>пункта в СДАНК нет — проводит любой с действующей аттестацией по методу НК и ПБ</span>":"<span class='muted'>пункта в СДАНК нет</span>")}${
        вОбл===false?` <span class="st bad">нет в области аттестации</span>`:""}</td>
      <td class="num">${н.приборы.length}</td>
      <td class="num">${н.нд.length}</td>
    </tr>`;
  }).join("");
  document.getElementById("main").innerHTML=`
    ${printHead("Методики контроля по методам","методов "+METHODS.length)}
    <div class="head"><h2>Методики контроля</h2>
      <span class="sub">стандартный набор приборов и норм. документов на каждый метод</span></div>
    <div class="note справ noprint">На каждый из ${METHODS.length} методов — свой набор приборов и
      норм. документов. Он подставляется при заполнении записи журнала (кнопка «⚙ Приборы и
      НД» у метода) — там его можно скорректировать только для этой записи, а можно сохранить
      как новый стандарт для метода.</div>
    <div class="card"><div class="tblwrap"><table>
      <thead><tr><th>Метод</th><th>По классификатору СДАНК</th><th style="width:140px">Приборов</th><th style="width:170px">Норм. документов</th></tr></thead>
      <tbody>${body}</tbody></table></div>
      <div style="padding:0 16px 12px">${блокИстории(REG_МЕТОДИКИ,ОДИНОЧНЫЕ[REG_МЕТОДИКИ].ид)}</div></div>`;
}
function editMetodika(м){
  формаСохранена();
  curView="metod";
  const мт=getМетодики(); const н=мт[м]||{приборы:[],нд:[]};
  const si=методСортировка(live(S), м), nd=методСортировка(live(НД), м);
  document.getElementById("main").innerHTML=`
    <div class="head"><h2>Методика: ${esc(м)}</h2>
      <span class="sub">стандартный набор приборов и норм. документов</span></div>
    <div class="card" style="padding:20px"><div class="form">
      <div class="grp" style="padding-left:0">Приборы (СИ)</div>
      <div id="метСпис_си" class="сд-список">
        ${списокМетодЧекбоксы(si, н.приборы, "мчб_си", x=>esc(x.наименование)+(x.зав?(" · зав. № "+esc(x.зав)):"")+
          (x.метод?` <span class="muted">— ${esc(x.метод)}</span>`:""))}
      </div>

      <div class="grp" style="padding-left:0;margin-top:14px">Нормативные документы</div>
      <div id="метСпис_нд" class="сд-список">
        ${списокМетодЧекбоксы(nd, н.нд, "мчб_нд", x=>ссылкаНД(x,esc(x.наименование))+(x.обозначение?(" · "+esc(x.обозначение)):"")+
          (x.метод?` <span class="muted">— ${esc(x.метод)}</span>`:""))}
      </div>

      <div class="formbtns закреп">
        <button class="btn primary" onclick="saveMetodika('${esc(м)}')">✓ Сохранить</button>
        ${кнопкаНазад('metod')}
      </div>
      ${блокИстории(REG_МЕТОДИКИ,ОДИНОЧНЫЕ[REG_МЕТОДИКИ].ид,путь=>путь.split(".")[0]===м)}
    </div></div>`;
}
function saveMetodika(м){
  const приборы=[...document.querySelectorAll(".мчб_си:checked")].map(c=>c.value);
  const нд=[...document.querySelectorAll(".мчб_нд:checked")].map(c=>c.value);
  const мт=getМетодики(); мт[м]={приборы:приборы, нд:нд};
  saveМетодики(мт);
  формаСохранена(); toast("Методика «"+м+"» сохранена"); renderMetod();
}
function рядыОбъектов(){
  const q=(query||"").toLowerCase();
  return live(O).filter(o=>!q||(o.наименование+o.зав+o.рег+o.тип+расположениеОб(o)).toLowerCase().includes(q));
}
function renderObjects(){
  const rows=рядыОбъектов();
  const body=rows.length?rows.map(o=>`
    <tr data-id="${esc(o.id)}" onclick="showObject('${esc(o.id)}')">
      <td class="obj">${esc(o.наименование)}</td>
      <td class="num">${esc(o.зав)||"—"}</td>
      <td class="num">${esc(o.рег)||"—"}</td>
      <td>${esc(o.тип)||"<span class='muted'>—</span>"}</td>
      <td class="num">${esc(расположениеОб(o))||"—"}</td>
    </tr>`).join(""):`<tr><td colspan="5" class="empty">Ничего не найдено.</td></tr>`;
  document.getElementById("main").innerHTML=`
    ${printHead("Справочник объектов контроля", (query?("поиск: «"+query+"» · "):"")+"объектов "+rows.length)}
    <div class="head"><h2>Объекты</h2><span class="sub">${rows.length} из ${live(O).length}</span>
      <div class="search"><span class="ic">⌕</span><input id="q" placeholder="Поиск: наименование, зав. №, тип…" value="${esc(query)}"></div>
      <div class="spring"></div>
      <button class="btn primary" onclick="editReg('obj',null)">＋ Добавить</button>
      <button class="btn" onclick="exportXlsx()">▤ Excel</button>
      <button class="btn" onclick="выгрузкаДляПравки('obj')" title="Выгрузить текущие карточки в Excel, поправить и загрузить обратно тем же импортом">⤓ Выгрузить для правки</button>
      <button class="btn" onclick="шаблонИмпортаXlsx('obj')">⤓ Шаблон для импорта</button>
      <button class="btn" onclick="document.getElementById('imp_obj').click()">⬆ Импорт из Excel</button>
      <input type="file" id="imp_obj" accept=".xlsx" style="display:none" onchange="импортФайлВыбран('obj',this)">
      <button class="btn" onclick="window.print()">⎙ Печать / PDF</button></div>
    <div class="card"><div class="tblwrap"><table data-список="obj">
      <thead><tr><th>Наименование</th><th>Заводской №</th><th>Рег. / уч. №</th><th>Тип / группа</th><th>Расположение</th></tr></thead>
      <tbody>${body}</tbody></table></div></div>`;
  wireSearch(renderObjects); инициализироватьКрестикиПоиска(); показатьТопПоиск("objects","Поиск: наименование, зав. №, тип…");
}
function showObject(id){
  const o=objById[id]||O.find(x=>String(x.id)===String(id)); if(!o) return;
  const F=(k,v)=>v?`<div class="fld"><div class="k">${k}</div><div class="v">${esc(v)}</div></div>`:"";
  openDrawer(o.наименование,`
    <div class="dbtns">
      <button class="btn primary" onclick="closeDrawer();editReg('obj','${esc(o.id)}')">✎ Изменить</button>
      <button class="btn danger" onclick="delReg('obj','${esc(o.id)}')">✕ Удалить</button>
      <button class="btn" onclick="closeDrawer();карточкаОбъектаЛист('${esc(o.id)}')" title="Карточка и история контроля на отдельном листе — для печати или PDF">⎙ Печать</button>
      <button class="btn" onclick="exportКарточкаОбъекта('${esc(o.id)}')">▤ Excel</button>
      <button class="btn" onclick="копияКарточки('obj','${esc(o.id)}')" title="Новая карточка по образцу этой — номера не переносятся">⧉ Копия</button>
      ${яНачальник()?`<button class="btn" onclick="задачаПоКарточке('obj','${esc(o.id)}')">✎ Задача</button>`:""}
    </div>
    ${F("Заводской №",o.зав)}${F("Регистрационный №",o.рег)}${F("Учётный №",o.учет)}${F("Инвентарный №",o.инв)}
    ${F("Тип / группа",o.тип)}${F("Расположение",расположениеОб(o))}${F("Изготовитель",o.изготовитель)}${F("Паспорт",o.паспорт)}
    <div class="sect">Параметры для контроля</div>
    ${F("Марка материала",o.материал)}${F("Объём",o.объем)}${F("Толщины",o.толщины)}
    ${F("Мин. толщина, мм",o.мин_толщина)}${F("Доп. твёрдость, НВ",o.твердость)}
    ${F("Рабочее давление",o.давление)}${F("Рабочая среда",o.среда)}
    ${F("Примечание",o.примечание)}
    ${блокИсторииОбъекта(o)}`);
}
// История контроля объекта — записи журнала, привязанные к ЭТОЙ карточке
// (объект_ид). Раньше карточка показывала только число записей, и считалось
// оно по совпадению наименования: у двенадцати одноимённых теплообменников
// счёт был общий на всех. Запись без привязки (заведена до справочника или
// из книги с неразобранной строкой) берётся, если строка объекта ведёт к
// этой карточке тем же поиском, что и форма записи.
function записиОбъекта(o){
  return liveJ().filter(z=>z.объект_ид!=null && z.объект_ид!==""
      ? тотЖеИд(z.объект_ид,o.id)
      : (!!z.объект && карточкаПоСтроке(z.объект,z.зав,z.рег)===o))
    .sort((a,b)=>String(b.дата||"").localeCompare(String(a.дата||"")));
}
function последнийКонтрольПоМетодам(записи){
  const п={};
  записи.forEach(z=>methodsOf(z).forEach(м=>{ if(z.дата && (!п[м] || z.дата>п[м].дата)) п[м]={дата:z.дата, z}; }));
  return METHODS.filter(м=>п[м]).map(м=>({м, дата:п[м].дата, z:п[м].z}));
}
function блокИсторииОбъекта(o){
  const зз=записиОбъекта(o);
  if(!зз.length) return `<div class="sect">История контроля</div><div class="note">В журнале записей по этому объекту нет.</div>`;
  const посл=последнийКонтрольПоМетодам(зз);
  const строки=зз.map(z=>{
    const мм=methodsOf(z).map(м=>{ const об=обозначениеПротокола(z,м);
      return `<div><b>${esc(м)}</b>${об?" · "+esc(об):" · <span class='muted'>без номера</span>"}${
        (z.методы[м]||{}).исп?` <span class="muted">— ${esc(z.методы[м].исп)}</span>`:""}</div>`; }).join("")
      || "<span class='muted'>методы не указаны</span>";
    return `<tr style="cursor:pointer" onclick="showJournal('${esc(z.id)}')">
      <td style="white-space:nowrap">${ru(z.дата)||"<span class='muted'>без даты</span>"}</td>
      <td>${мм}</td><td style="white-space:nowrap">${esc(z.акт)||"<span class='muted'>—</span>"}</td></tr>`;
  }).join("");
  return `<div class="sect">История контроля — ${зз.length} ${числослово(зз.length,["запись","записи","записей"])}</div>
    ${посл.length?`<div class="fld"><div class="k">Последний контроль по методам</div><div class="v">${
      посл.map(x=>esc(x.м)+" — "+ru(x.дата)).join("; ")}</div></div>`:""}
    <table class="mtab"><thead><tr><th>Дата</th><th>Методы, протоколы, исполнители</th><th>Акт</th></tr></thead>
      <tbody>${строки}</tbody></table>
    <div class="note справ">Щелчок по строке открывает запись журнала.</div>`;
}

// ── Карточка объекта на печать и в Excel ─────────────────────
// Все поля карточки (в порядке формы, REGS.obj.fields, пустые не
// печатаются) и вся история контроля. Отдельный лист в основной области,
// а не боковая панель: панель при печати скрыта, а на листе своя шапка
// печати, как у остальных разделов. Форма внутренняя — образца
// «карточки объекта» у заказчика нет, поэтому она не выдаётся за
// официальный документ.
function поляОбъектаДляВывода(o){
  return REGS.obj.fields.map(f=>{
    let v=o[f.k];
    if(f.type==="podr"){ const п=найтиПоИд(live(ПОДР),v); v=п?п.наименование:(v||""); }
    if(f.k==="расположение") v=расположениеОб(o);
    return {l:f.l, v:String(v==null?"":v).trim()};
  }).filter(x=>x.v);
}
function строкиИсторииОбъекта(o){
  const из=[];
  записиОбъекта(o).forEach(z=>{
    const мм=methodsOf(z);
    if(!мм.length) из.push({z, м:"", протокол:"", исп:""});
    мм.forEach(м=>из.push({z, м, протокол:обозначениеПротокола(z,м)||"", исп:String(((z.методы||{})[м]||{}).исп||"")}));
  });
  return из;
}
function карточкаОбъектаЛист(id){
  const o=найтиПоИд(O,id); if(!o) return;
  if(!уходСФормы()) return;
  curView="objcard"; query="";
  document.querySelectorAll(".nav").forEach(n=>n.classList.toggle("on",n.dataset.view==="objects"));
  const поля=поляОбъектаДляВывода(o), ист=строкиИсторииОбъекта(o), посл=последнийКонтрольПоМетодам(записиОбъекта(o));
  document.getElementById("main").innerHTML=`
    ${printHead("Карточка объекта контроля", строкаОбъектаСпр(o))}
    <div class="head"><h2>Карточка объекта</h2><span class="sub">${esc(строкаОбъектаСпр(o))}</span>
      <div class="spring"></div>
      <button class="btn primary" onclick="window.print()">⎙ Печать / PDF</button>
      <button class="btn" onclick="exportКарточкаОбъекта('${esc(o.id)}')">▤ Excel</button>
      <button class="btn" onclick="switchView('objects')">← К объектам</button></div>
    <div class="card" style="padding:16px 20px">
      <h3 class="прз" style="margin-top:0">Сведения об объекте</h3>
      <table class="mtab карт"><tbody>${поля.map(x=>`<tr><th style="width:34%;text-align:left">${esc(x.l)}</th><td>${esc(x.v)}</td></tr>`).join("")}</tbody></table>
      <h3 class="прз">История контроля — ${ист.length?записиОбъекта(o).length+" "+числослово(записиОбъекта(o).length,["запись","записи","записей"]):"записей нет"}</h3>
      ${посл.length?`<p style="margin:4px 0 8px">Последний контроль по методам: ${посл.map(x=>esc(x.м)+" — "+ru(x.дата)).join("; ")}</p>`:""}
      ${ист.length?`<table class="mtab"><thead><tr><th>Дата</th><th>Метод</th><th>№ протокола</th><th>Исполнитель</th><th>№ акта</th><th>Заключение</th></tr></thead>
        <tbody>${ист.map(r=>`<tr><td style="white-space:nowrap">${ru(r.z.дата)||"—"}</td><td>${esc(r.м)||"—"}</td>
          <td style="white-space:nowrap">${esc(r.протокол)||"—"}</td><td>${esc(r.исп)||"—"}</td>
          <td style="white-space:nowrap">${esc(r.z.акт)||"—"}</td><td>${esc(r.z.заключение||"")||"—"}</td></tr>`).join("")}</tbody></table>`
        :`<div class="note">В журнале записей по этому объекту нет.</div>`}
    </div>
    ${printSign()}`;
}
function exportКарточкаОбъекта(id){
  const o=найтиПоИд(O,id); if(!o) return;
  const ист=строкиИсторииОбъекта(o);
  const листы=[{имя:"Карточка", колонки:[{з:"Поле",ш:32},{з:"Значение",ш:60}],
      строки:поляОбъектаДляВывода(o).map(x=>[x.l,x.v])},
    {имя:"История контроля", колонки:[{з:"Дата",ш:12,тип:"дата"},{з:"Метод",ш:8},{з:"№ протокола",ш:18},
      {з:"Исполнитель",ш:20},{з:"№ акта",ш:14},{з:"Заключение",ш:50}],
      строки:ист.map(r=>[r.z.дата||"", r.м, r.протокол, r.исп, r.z.акт||"", r.z.заключение||""])}];
  const имя=String(o.наименование||"объект").replace(/[\\/:*?"<>|]/g," ").slice(0,60).trim();
  xlsxФайл("Карточка объекта "+имя+(o.зав?" зав "+String(o.зав).replace(/[\\/:*?"<>|]/g," "):"")+".xlsx", листы);
  toast("Карточка выгружена: "+ист.length+" строк истории");
}

