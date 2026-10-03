// ── Акты и протоколы ──────────────────────────────────────
// Раздел отвечает на один рабочий вопрос: по каким записям документ уже
// оформлен, а по каким ещё нет. «Оформлен» = в документе вписано
// заключение (оно и хранится в записи).
let df="акты";                          // акты | без | нужно | готовые | все
function setDF(v){ df=v; renderDocs(); }
function оформлен(z){ return !!String(z.заключение||"").trim(); }
function docsRows(){
  const q=(query||"").toLowerCase();
  return liveJ().filter(z=>{
    if(df==="без" && !ждётНомеров(z)) return false;
    if(df==="готовые" && !(оформлен(z) && !ждётНомеров(z))) return false;
    if(df==="нужно" && (оформлен(z) || ждётНомеров(z))) return false;
    if(q && !((z.объект||"")+(z.акт||"")+(z.заключение||"")).toLowerCase().includes(q)) return false;
    return true;
  }).sort((a,b)=>String(b.дата||"").localeCompare(String(a.дата||"")));
}
// ── Реестр актов ───────────────────────────────────────────
// Список актов с тем, что по ним известно из журнала: даты, подразделение,
// объекты, протоколы по методам, записи без номера и итог «Готовности к
// выдаче». Отбор — тот же, что на экране (поиск в разделе). Форма
// внутренняя, образца реестра у заказчика нет.
function актыПоказанные(){
  const q=(query||"").toLowerCase(), акты=списокАктов();
  return q?акты.filter(g=>(g.ключ+" "+g.записи.map(z=>z.объект).join(" ")).toLowerCase().includes(q)):акты;
}
function реестрАктов(акты){
  const методы=new Set();
  const ряды=акты.map(g=>{
    const zs=g.записи, д=zs.map(z=>z.дата).filter(Boolean).sort();
    const пм={}; спискиАкта(zs).forEach(x=>{ пм[x.м]=x.номера.length; методы.add(x.м); });
    const п=подрАкта(g.ключ), гот=готовностьАкта(g.ключ);
    return {ключ:g.ключ, номер:номерАкта(g.ключ), год:годАкта(g.ключ), с:д[0]||"", по:д[д.length-1]||"",
      подр:п?(п.наименование||""):String(актЗапись(g.ключ,false).подразделение||"").trim(),
      записей:zs.length, объекты:[...new Set(zs.map(именоватьОбъект))].join("; "), пм,
      протоколов:Object.values(пм).reduce((a,b)=>a+b,0), безНомера:zs.filter(ждётНомеров).length,
      ошибок:гот.filter(x=>x.ур==="ошибка").length, внимания:гот.filter(x=>x.ур==="внимание").length};
  });
  const порядок=ПОРЯДОК_В_АКТЕ.concat(METHODS).filter((м,i,a)=>a.indexOf(м)===i && методы.has(м));
  return {ряды, методы:порядок};
}
function листРеестраАктов(акты){
  const р=реестрАктов(акты);
  const колонки=[{з:"№ акта",ш:12},{з:"Год",ш:7,тип:"число"},{з:"Дата работ с",ш:13,тип:"дата"},{з:"по",ш:13,тип:"дата"},
    {з:"Подразделение",ш:24},{з:"Записей журнала",ш:10,тип:"число"},{з:"Объекты контроля",ш:60}]
    .concat(р.методы.map(м=>({з:"Протоколов "+м,ш:10,тип:"число"})))
    .concat([{з:"Протоколов всего",ш:10,тип:"число"},{з:"Записей без номера",ш:10,тип:"число"},
      {з:"Готовность: ошибок",ш:10,тип:"число"},{з:"Готовность: внимание",ш:10,тип:"число"}]);
  const строки=р.ряды.map(x=>[x.номер,+x.год||x.год,x.с,x.по,x.подр,x.записей,x.объекты]
    .concat(р.методы.map(м=>x.пм[м]||0)).concat([x.протоколов,x.безНомера,x.ошибок,x.внимания]));
  return {имя:"Реестр актов", колонки, строки, пусто:!р.ряды.length};
}
function exportРеестрАктов(){
  const л=листРеестраАктов(актыПоказанные());
  if(л.пусто){ toast("Актов не найдено — выгружать нечего"); return; }
  xlsxФайл("Реестр актов "+сегодня()+".xlsx",[л]);
  toast("Реестр выгружен: актов "+л.строки.length);
}
// ── Акты за период — в папку с документами ────────────────
// Каждый акт периода — своим файлом Word, как при выгрузке одного акта.
// Период — по дате работ акта (датаРаботАкта: самая поздняя дата
// контроля его записей). Акт с ошибками «Готовности к выдаче» молча не
// выгружается и не пропадает: он пропускается и называется в итоге вместе
// с ошибками — поправить и выгрузить отдельно.
let периодАктов="месяц";
const ПЕРИОДЫ_АКТОВ=[["месяц","этот месяц"],["прошлый_месяц","прошлый месяц"],["квартал","этот квартал"],
  ["прошлый_квартал","прошлый квартал"],["год","этот год"]];
function границыПериодаАктов(вид, сег){
  const d=сег?new Date(сег+"T00:00:00"):new Date(), г=d.getFullYear(), м=d.getMonth();
  const iso=(y,mm,dd)=>new Date(Date.UTC(y,mm,dd)).toISOString().slice(0,10);
  const кв=Math.floor(м/3);
  switch(вид){
    case "прошлый_месяц": return {с:iso(г,м-1,1), по:iso(г,м,0), имя:МЕСЯЦЫ[(м+11)%12].toLowerCase()+" "+(м?г:г-1)};
    case "квартал": return {с:iso(г,кв*3,1), по:iso(г,кв*3+3,0), имя:(кв+1)+" квартал "+г};
    case "прошлый_квартал": { const к=кв?кв-1:3, гг=кв?г:г-1; return {с:iso(гг,к*3,1), по:iso(гг,к*3+3,0), имя:(к+1)+" квартал "+гг}; }
    case "год": return {с:iso(г,0,1), по:iso(г,11,31), имя:String(г)+" год"};
    default: return {с:iso(г,м,1), по:iso(г,м+1,0), имя:МЕСЯЦЫ[м].toLowerCase()+" "+г};
  }
}
function актыПериода(вид, сег){
  const гр=границыПериодаАктов(вид, сег);
  return {гр, акты:списокАктов().filter(а=>{ const д=датаРаботАкта(а.ключ); return д && д>=гр.с && д<=гр.по; })};
}
async function актыВПапку(){
  const {гр, акты}=актыПериода(периодАктов);
  if(!акты.length){ toast("За "+гр.имя+" актов нет (по дате работ)"); return; }
  const h=await папкаДок();
  if(!h){ toast("Сначала выберите папку с документами — «Обмен файлами»"); switchView("exchange"); return; }
  if(await правоНаПапкуДок(true)!=="есть"){ toast("Нет разрешения на папку с документами"); return; }
  const сОшибками=[], готовы=[];
  акты.forEach(а=>{ const ош=готовностьАкта(а.ключ).filter(x=>x.ур==="ошибка");
    (ош.length?сОшибками:готовы).push({ключ:а.ключ, ош}); });
  if(!confirm("Акты за "+гр.имя+": "+акты.length+".\nВыгрузить в Word: "+готовы.length+
    (сОшибками.length?"\nПропустить с ошибками готовности: "+сОшибками.length+" ("+сОшибками.map(x=>номерАкта(x.ключ)).join(", ")+")":"")+
    "\n\nФайлы лягут в папку с документами, в подпапку «Акты — "+гр.имя+"»; файлы с тем же именем заменятся.")) return;
  const имяПапки="Акты — "+гр.имя; let записано=0; const сбой=[];
  try{
    const цель=await h.getDirectoryHandle(имяПапки,{create:true});
    for(const а of готовы){
      const д=актВWord(а.ключ,true); if(!д){ сбой.push(номерАкта(а.ключ)); continue; }
      try{ const вых=await цель.getFileHandle(д.имя,{create:true}); const w=await вых.createWritable();
        await w.write(д.данные); await w.close(); записано++; }
      catch(e){ сбой.push(номерАкта(а.ключ)+" ("+e.name+")"); }
    }
  }catch(e){ toast("Не удалось записать в папку: "+e.name); return; }
  alert("Папка «"+имяПапки+"»: записано актов — "+записано+" из "+акты.length+"."+
    (сОшибками.length?"\n\nПропущены из-за ошибок готовности к выдаче:\n"+сОшибками.map(x=>"· акт № "+номерАкта(x.ключ)+": "+
      x.ош.slice(0,3).map(y=>y.т).join("; ")+(x.ош.length>3?" …":"")).join("\n")+
      "\nПоправьте и выгрузите их по одному (откроется вопрос «всё равно выгрузить»).":"")+
    (сбой.length?"\n\nНе записались: "+сбой.join(", "):""));
}
function renderDocs(){
  const rows=docsRows();
  const всего=liveJ().length;
  const гот=liveJ().filter(z=>оформлен(z)&&!ждётНомеров(z)).length;
  const безном=liveJ().filter(ждётНомеров).length;
  const seg=(v,l,n)=>`<button class="btn ${df===v?"on":""}" onclick="setDF('${v}')">${l} <b>${n}</b></button>`;
  const акты=списокАктов();
  const актыП=актыПоказанные();
  const bodyА=актыП.length?актыП.map(g=>{
    const zs=g.записи, дата=zs.map(z=>z.дата).filter(Boolean).sort();
    const мет=спискиАкта(zs);
    const ждут=zs.filter(ждётНомеров).length;
    return `<tr onclick="renderAkt('${esc(g.ключ)}')">
      <td class="num">${esc(номерАкта(g.ключ))}</td>
      <td class="num">${ru(дата[0])}${дата.length>1&&дата[дата.length-1]!==дата[0]?(" – "+ru(дата[дата.length-1])):""}</td>
      <td class="num">${zs.length}</td>
      <td class="obj">${esc([...new Set(zs.map(именоватьОбъект))].join("; ").slice(0,64))}</td>
      <td class="mets">${мет.map(g2=>`${g2.м}&nbsp;${g2.номера.length}`).join(" · ")||"<span class='muted'>—</span>"}
        ${ждут?`<span class="st work" style="margin-left:6px">${ждут} без номера</span>`:""}
        ${(()=>{ const ош=готовностьАкта(g.ключ).filter(x=>x.ур==="ошибка").length;
          return ош?`<span class="st bad" style="margin-left:6px" title="откройте акт — список в блоке «Готовность к выдаче»">ошибок ${ош}</span>`:""; })()}</td>
    </tr>`;}).join(""):`<tr><td colspan="5" class="empty">Актов не найдено.</td></tr>`;
  const body=rows.length?rows.map(z=>`
    <tr onclick="открытьДокумент('${esc(z.id)}')">
      <td class="num">${esc(z.акт)||"—"}</td>
      <td class="num">${ru(z.дата)}</td>
      <td class="obj">${esc(z.объект)}</td>
      <td>${methodsOf(z).map(m=>`<span class="pill">${m}</span>`).join("")||"<span class='muted'>—</span>"}</td>
      <td>${ждётНомеров(z) ? `<span class="st work">ждёт номера</span>`
        : оформлен(z) ? `<span class="st ok">готов</span>`
        : `<span class="st new">нет заключения</span>`}</td>
      <td class="obj muted" style="font-size:12px">${esc(String(z.заключение||"").slice(0,90))}${String(z.заключение||"").length>90?"…":""}</td>
    </tr>`).join(""):`<tr><td colspan="6" class="empty">Ничего не найдено.</td></tr>`;
  document.getElementById("main").innerHTML=`
    ${printHead("Акты и протоколы", df==="акты"?("актов "+актыП.length)
      : (df==="без"?"ждут номера":df==="готовые"?"готовые":df==="нужно"?"без заключения":"все")+" · записей "+rows.length)}
    <div class="head">
      <h2>Акты и протоколы</h2><span class="sub">${df==="акты"?(актыП.length+" актов"):(rows.length+" из "+всего)}</span>
      <div class="search"><span class="ic">⌕</span><input id="q" placeholder="Поиск: объект, № акта, текст заключения…" value="${esc(query)}"></div>
      <div class="spring"></div>
      ${df==="акты"?`<button class="btn" onclick="exportРеестрАктов()" title="Показанные акты в Excel: даты, подразделение, объекты, протоколы по методам, готовность">▤ Реестр актов в Excel</button>
        <select onchange="периодАктов=this.value" title="Период по дате работ акта">${ПЕРИОДЫ_АКТОВ.map(([k,t])=>`<option value="${k}"${k===периодАктов?" selected":""}>${t}</option>`).join("")}</select>
        <button class="btn" onclick="актыВПапку()" title="Каждый акт периода — своим файлом Word в папку с документами; акты с ошибками готовности пропускаются и называются">⤓ Акты за период в папку</button>`:""}
      <button class="btn" onclick="window.print()">⎙ PDF</button>
    </div>
    <div class="filters noprint seg">
      ${seg("акты","Акты",акты.length)}
      ${seg("без","Ждут номера",безном)}
      ${seg("нужно","Без заключения",всего-гот-безном)}
      ${seg("готовые","Готовы",гот)}
      ${seg("все","Все",всего)}
      <span class="fnote">клик по строке — открыть документ</span>
    </div>
    <div class="card"><div class="tblwrap"><table>
      ${df==="акты"
        ? `<thead><tr><th style="width:90px">№ акта</th><th style="width:110px">Дата</th><th style="width:70px">Объектов</th><th>Объект контроля</th><th>Документы</th></tr></thead>
           <tbody>${bodyА}</tbody>`
        : `<thead><tr><th>№ акта</th><th>Дата</th><th>Объект контроля</th><th>Методы</th><th>Документ</th><th>Заключение</th></tr></thead>
           <tbody>${body}</tbody>`}
    </table></div></div>`;
  wireSearch(renderDocs);
}

