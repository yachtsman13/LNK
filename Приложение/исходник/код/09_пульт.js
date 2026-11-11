// ── Пульт (сводка) ────────────────────────────────────────
// Что по задачам горит сегодня — для пульта. Считаются только свои
// задачи: производные и так видны на своих местах (план ППР, приборы).
function задачиНаСегодня(){
  const сегодня=new Date().toISOString().slice(0,10);
  const мои=ЗАД.filter(з=>живаяЗадача(з) && з.статус!=="закрыта")
    .filter(з=>з.срок && з.срок<=сегодня)
    .sort((a,b)=>String(a.срок).localeCompare(String(b.срок)));
  return { всего:мои.length, просрочено:мои.filter(просроченаЗадача).length, список:мои };
}
// Что горит на этой неделе (срок завтра..+7 дней) — чтобы не узнавать
// о просрочке постфактум, как «задачи на сегодня» выше. Тоже только свои —
// производные видны на своих местах.
function задачиСкоро(){
  const d=new Date(); const сегодня=d.toISOString().slice(0,10);
  d.setDate(d.getDate()+7); const через7=d.toISOString().slice(0,10);
  const мои=ЗАД.filter(з=>живаяЗадача(з) && з.статус!=="закрыта")
    .filter(з=>з.срок && з.срок>сегодня && з.срок<=через7)
    .sort((a,b)=>String(a.срок).localeCompare(String(b.срок)));
  return { всего:мои.length, список:мои };
}

// ── Работы по месяцам (пульт) ─────────────────────────────
// Сетка «метод × месяц» за год: в клетке — число работ (метод в записи,
// та же единица, что у «Нагрузки»), цвет — одна синяя шкала от светлого к
// тёмному. Столбики с восемью методами потребовали бы восьми разных красок и
// всё равно не дали бы прочитать число, а в сетке число стоит прямо в
// клетке. Щелчок по клетке — журнал с отбором по году и методу.
let годРабот=new Date().getFullYear();
function setГодРабот(г){ годРабот=+г; renderDash(); }
function работыПоМесяцам(год){
  const поМ={}, итогМ=Array(12).fill(0);
  liveJ().forEach(z=>{
    if(!z.дата || String(z.дата).slice(0,4)!==String(год)) return;
    const мес=+String(z.дата).slice(5,7)-1; if(!(мес>=0&&мес<12)) return;
    methodsOf(z).forEach(м=>{ (поМ[м]=поМ[м]||Array(12).fill(0))[мес]++; итогМ[мес]++; });
  });
  const методы=METHODS.filter(м=>поМ[м]);
  const макс=Math.max(0,...методы.flatMap(м=>поМ[м]));
  return {поМ, итогМ, методы, макс, всего:итогМ.reduce((a,b)=>a+b,0)};
}
// Сравнение с прошлым годом — итоги по месяцам того же журнала. Разница
// считается только за прошедшие месяцы: у текущего года будущие месяцы
// пусты, и «минус» за них был бы неправдой. Итог — тоже за одинаковые
// месяцы обоих лет («с начала года»), а не весь прошлый год против части
// текущего. Нет записей за прошлый год — так и сказано, без нулевой строки.
function строкиСравнения(год, р){
  const п=работыПоМесяцам(год-1);
  const сегодня=new Date(), текГ=сегодня.getFullYear(), текМ=сегодня.getMonth();
  const прошло=i=>год<текГ || (год===текГ && i<=текМ);
  if(!п.всего) return `<tr class="срав"><td class="м" colspan="14" style="text-align:left;font-weight:normal">
      <span class="muted">За ${год-1} год в журнале записей нет — сравнивать не с чем.</span></td></tr>`;
  const разн=i=>{ if(!прошло(i)) return `<td class="итог"><span class="muted">·</span></td>`;
    const d=р.итогМ[i]-п.итогМ[i]; return `<td class="итог">${d>0?"+"+d:d<0?"−"+(-d):"0"}</td>`; };
  const мес=[...Array(12).keys()].filter(прошло);
  const тек=мес.reduce((a,i)=>a+р.итогМ[i],0), прош=мес.reduce((a,i)=>a+п.итогМ[i],0), д=тек-прош;
  const проц=прош?Math.round(д/прош*100):null;
  return `<tr class="срав"><td class="м">${год-1}</td>${п.итогМ.map(v=>`<td class="итог">${v||"<span class='muted'>·</span>"}</td>`).join("")}
      <td class="итог">${п.всего}</td></tr>
    <tr class="срав"><td class="м" title="${esc(год+" минус "+(год-1)+", только прошедшие месяцы")}">Разница</td>${[...Array(12).keys()].map(разн).join("")}
      <td class="итог" title="${esc("за "+(мес.length===12?"год":"январь — "+МЕСЯЦЫ[мес.length-1].toLowerCase())+": "+тек+" против "+прош)}">${д>0?"+"+д:д<0?"−"+(-д):"0"}${проц!==null?`<div class="muted" style="font-size:11px">${проц>0?"+":""}${проц}%</div>`:""}</td></tr>`;
}
// ── Итоги года одной книгой Excel ─────────────────────────────
// Собирает на отдельные листы то, что уже считается в разных разделах —
// своего расчёта здесь нет: работы по месяцам (карта на пульте), нагрузка
// специалистов, сводка по подразделениям, реестр актов года и выполнение
// плана ППР. Первый лист — короткая сводка цифр. Год журнала — по дате
// контроля, год акта — по его номеру (годАкта). Форма внутренняя, образца
// годового отчёта у заказчика нет.
function итогиГодаСводка(год){
  const г=String(год);
  const зз=liveJ().filter(z=>String(z.дата||"").slice(0,4)===г);
  const объекты=new Set(), протоколы={с:0,без:0};
  зз.forEach(z=>{ объекты.add(String(z.объект_ид!=null&&z.объект_ид!==""?z.объект_ид:(z.объект||"")));
    methodsOf(z).forEach(м=>{ if(номерМетода(z,м)) протоколы.с++; else протоколы.без++; }); });
  const акты=списокАктов().filter(x=>String(годАкта(x.ключ))===г);
  const ппр=годыППР().includes(+год)?итогиГода(+год):null;
  // «Прошедший месяц» — тот же прошедший(), что красит план и считает
  // лист «План ППР»: текущий месяц ещё идёт и прошедшим не считается.
  let ппрПрошло=0, ппрПрошлоСделано=0;
  if(ппр) for(let м=1;м<=12;м++) if(прошедший(+год,м)) работыМесяца(+год,м).forEach(р=>{ ппрПрошло++; if(р.отм) ппрПрошлоСделано++; });
  return {зз, записей:зз.length, работ:зз.reduce((n,z)=>n+methodsOf(z).length,0), объектов:объекты.size,
    актов:акты.length, акты, протоколы, ппр, ппрПрошло, ппрПрошлоСделано, безДаты:liveJ().filter(z=>!z.дата).length};
}
function листПланаГода(год){
  const строки=[];
  let итог=[0,0,0,0];
  for(let м=1;м<=12;м++){
    const рр=работыМесяца(+год,м), сд=рр.filter(р=>р.отм);
    const часы=+рр.reduce((a,р)=>a+р.часы,0).toFixed(2), часыС=+сд.reduce((a,р)=>a+р.часы,0).toFixed(2);
    строки.push([МЕСЯЦЫ[м-1], рр.length, сд.length, прошедший(+год,м)?рр.length-сд.length:"", часы, часыС]);
    итог=[итог[0]+рр.length, итог[1]+сд.length, итог[2]+(прошедший(+год,м)?рр.length-сд.length:0), итог[3]+часы];
  }
  строки.push(["Итого", итог[0], итог[1], итог[2], +итог[3].toFixed(2),
    +строки.reduce((a,r)=>a+(typeof r[5]==="number"?r[5]:0),0).toFixed(2)]);
  return {имя:"План ППР", колонки:[{з:"Месяц",ш:12},{з:"Работ по плану",ш:12,тип:"число"},{з:"Отмечено выполненных",ш:14,тип:"число"},
    {з:"Прошедший месяц: не отмечено",ш:16,тип:"число"},{з:"Трудозатраты по плану, ч",ш:14,тип:"число"},{з:"Трудозатраты выполненных, ч",ш:14,тип:"число"}], строки};
}
function exportИтогиГода(год){
  const св=итогиГодаСводка(год);
  if(!св.записей && !св.актов && !св.ппр){ toast("За "+год+" год нет ни записей журнала, ни актов, ни плана ППР — выгружать нечего"); return; }
  const р=работыПоМесяцам(год);
  const м3=МЕСЯЦЫ.map(м=>м.slice(0,3).toLowerCase());
  const листы=[
    {имя:"Сводка", колонки:[{з:"Показатель",ш:52},{з:String(год),ш:14,тип:"число"}], строки:[
      ["Записей журнала (по дате контроля)", св.записей],
      ["Работ — методов в записях", св.работ],
      ["Объектов контроля (без повторов)", св.объектов],
      ["Актов (по году в номере акта)", св.актов],
      ["Методов с номером протокола", св.протоколы.с],
      ["Методов без номера протокола", св.протоколы.без],
    ].concat(св.ппр?[
      ["План ППР: работ за год", св.ппр.работ],
      ["План ППР: отмечено выполненных", св.ппр.сделано],
      ["План ППР: прошедших месяцев — работ по плану", св.ппрПрошло],
      ["План ППР: прошедших месяцев — не отмечено", св.ппрПрошло-св.ппрПрошлоСделано],
    ]:[["План ППР", "на "+год+" год позиций нет"]]).concat([
      ["Записей журнала без даты (не попали ни в один год)", св.безДаты],
      ["Сформировано", ru(new Date().toISOString().slice(0,10))],
      ["Примечание", "Внутренняя сводка из данных программы; образца годового отчёта нет."]])},
    {имя:"Работы по месяцам", колонки:[{з:"Метод",ш:10}].concat(м3.map(м=>({з:м,ш:7,тип:"число"}))).concat([{з:"Итого",ш:9,тип:"число"}]),
     строки:р.методы.map(м=>[м].concat(р.поМ[м]).concat([р.поМ[м].reduce((a,b)=>a+b,0)]))
       .concat(р.методы.length?[["Все"].concat(р.итогМ).concat([р.всего])]:[])},
    Object.assign(листНагрузки(год,"год"),{имя:"Специалисты"}),
    Object.assign(листСводкиПодр(год,"год"),{имя:"Подразделения"}),
    Object.assign(листРеестраАктов(св.акты),{имя:"Акты"})
  ];
  if(св.ппр) листы.push(листПланаГода(год));
  листы.forEach(л=>{ if(!л.строки.length) л.строки=[["За "+год+" год данных нет"]]; delete л.пусто; });
  xlsxФайл("Итоги "+год+" года ЛНК.xlsx", листы);
  toast("Итоги "+год+" года выгружены: листов "+листы.length);
}
function блокРаботПоМесяцам(){
  const годы=годыНагрузки(); if(!годы.includes(годРабот)) годРабот=годы[годы.length-1];
  const р=работыПоМесяцам(годРабот);
  const корзина=v=>v?Math.max(1,Math.min(7,Math.ceil(v/р.макс*7))):0;
  const м3=["янв","фев","мар","апр","май","июн","июл","авг","сен","окт","ноя","дек"];
  const строки=р.методы.map(м=>`<tr><td class="м">${esc(м)}</td>${р.поМ[м].map((v,i)=>v
      ? `<td class="т${корзина(v)}" data-п="${esc(м)} · ${МЕСЯЦЫ[i].toLowerCase()} ${годРабот}: ${v} ${числослово(v,["работа","работы","работ"])}"
          onclick="работыВЖурнал('${esc(м)}')">${v}</td>`
      : `<td class="пусто">·</td>`).join("")}
      <td class="итог">${р.поМ[м].reduce((a,b)=>a+b,0)}</td></tr>`).join("");
  return `<div class="panel" style="margin-bottom:14px">
    <h3 style="display:flex;align-items:center;gap:8px;flex-wrap:wrap">▦ Работы по месяцам
      <span class="muted" style="font-weight:normal;font-size:12px">${годРабот} · всего ${р.всего}</span>
      <span style="margin-left:auto;display:flex;gap:4px">${годы.map(г=>
        `<button class="btn small ${г===годРабот?"primary":""}" onclick="setГодРабот(${г})">${г}</button>`).join("")}
        <button class="btn small" onclick="exportИтогиГода(годРабот)" title="Книга Excel: сводка, работы по месяцам, специалисты, подразделения, акты, план ППР">⤓ Итоги ${годРабот} года</button></span></h3>
    ${р.методы.length?`<div class="tblwrap" style="border:0"><table class="тепло" onmousemove="теплоПодсказка(event)" onmouseleave="теплоПодсказка(null)">
      <thead><tr><th></th>${м3.map(x=>`<th>${x}</th>`).join("")}<th>Итого</th></tr></thead>
      <tbody>${строки}</tbody>
      <tfoot><tr><td class="м">Все</td>${р.итогМ.map(v=>`<td class="итог">${v||"<span class='muted'>·</span>"}</td>`).join("")}
        <td class="итог">${р.всего}</td></tr>${строкиСравнения(годРабот,р)}</tfoot></table></div>
      <div class="note" style="margin:6px 0 0">Работа — метод в записи журнала, месяц — по дате контроля.
        Чем сильнее цвет клетки отличается от фона, тем больше работ; число — в самой клетке. Щелчок по клетке открывает журнал с отбором по году и методу.
        Строка «Разница» — ${годРабот} минус ${годРабот-1} только за прошедшие месяцы; итог справа — за те же месяцы обоих лет.</div>`
    :`<div class="note">За ${годРабот} год в журнале работ нет.</div>`}
    <div class="тепло-подск" id="теплоПодск"></div>
  </div>`;
}
function теплоПодсказка(e){
  const п=document.getElementById("теплоПодск"); if(!п) return;
  const td=e&&e.target.closest&&e.target.closest("td[data-п]");
  if(!td){ п.style.display="none"; return; }
  п.textContent=td.dataset.п; п.style.display="block";
  п.style.left=Math.min(e.clientX+12, window.innerWidth-п.offsetWidth-8)+"px";
  п.style.top=(e.clientY+14)+"px";
}
function работыВЖурнал(м){
  switchView("journal");
  if(typeof resetJF==="function") resetJF();
  setJF("год",String(годРабот)); setJF("метод",м);
}
function renderDash(){
  догнатьПовторы();
  const L=liveJ();
  const year=new Date().getFullYear();
  const заГод=L.filter(z=>(z.дата||"").slice(0,4)===String(year)).length;
  const st=(k)=>L.filter(z=>statusClass(z.статус)===k).length;

  // поверка приборов — списанные не считаются: их не представляют на
  // поверку, поэтому их просрочка не дело, а просто история
  const groups={bad:[],work:[],ok:[],none:[]};
  приборыСПоверкой().forEach(s=>{ const[c]=siStatus(s);
    if(c==="bad")groups.bad.push(s); else if(c==="work")groups.work.push(s);
    else if(c==="ok")groups.ok.push(s); else groups.none.push(s); });
  const attn=groups.bad.concat(groups.work)
    .sort((a,b)=>(a.поверка_до||"").localeCompare(b.поверка_до||""));

  // записи по методам
  const mc=METHODS.map(m=>({m,n:L.filter(z=>z.методы&&z.методы[m]).length}));
  const mmax=Math.max(1,...mc.map(x=>x.n));

  // последние записи по дате
  const last=L.slice().filter(z=>z.дата).sort((a,b)=>b.дата.localeCompare(a.дата)).slice(0,6);

  const tile=(n,l,cls,view)=>`<div class="tile ${cls||""} ${view?"click":""}"${view?` onclick="switchView('${view}')"`:""}>
    <div class="n">${n}</div><div class="l">${l}</div></div>`;

  let поверкаHTML;
  if(приборыСПоверкой().length===0){ поверкаHTML=`<div class="note">${
    !live(S).length?"Приборы не заведены.":приборыВДокументы().length?"Поверяемых приборов нет (только вспомогательное оборудование).":"Все приборы списаны — следить за поверкой нечего."}</div>`; }
  else {
    поверкаHTML=`
      <div class="tiles" style="margin-bottom:12px">
        ${tile(groups.bad.length,"просрочено","bad")}
        ${tile(groups.work.length,"истекает (30 дн.)","warn")}
        ${tile(groups.ok.length,"в норме","ok")}
        ${tile(groups.none.length,"дата не указана","")}
      </div>`;
    if(attn.length){
      поверкаHTML+=`<div class="plist">`+attn.map(s=>{const[c,l]=siStatus(s);
        return `<div class="pr" onclick="switchView('si')">
          <span class="st ${c}">${l}</span>
          <span class="g">${esc(s.наименование)}${s.зав?` · зав. ${esc(s.зав)}`:""}</span>
          <span class="num">${ru(s.поверка_до)||"—"}</span></div>`;}).join("")+`</div>`;
    } else if(groups.none.length){
      поверкаHTML+=`<div class="note">Просроченных нет. Даты поверки не заполнены — статус не отслеживается, пока даты не внесены (приборов: ${groups.none.length}).</div>`;
    } else {
      поверкаHTML+=`<div class="note">Все поверки в норме.</div>`;
    }
  }

  const дн=днейБезКопии(), об=объёмХранилища();
  const тревога = дн===null
    ? `<b>Резервной копии ещё не было.</b> Память браузера могут очистить — и работа пропадёт вся сразу.`
    : (дн>=3 ? `<b>Резервной копии не было ${дн} дн.</b> Это единственная страховка от очистки памяти браузера.` : "");
  document.getElementById("main").innerHTML=`
    <div class="head"><h2>Пульт</h2><span class="sub">на ${ru(new Date().toISOString().slice(0,10))}</span></div>
    ${плашкаПереезда()}
    ${тревога?`<div class="draftnote" style="cursor:pointer" onclick="switchView('exchange')">
      ◆ ${тревога} Нажмите — открою «Обмен», там одна кнопка.</div>`:""}
    ${(()=>{ const п=предупрежденияОбмена(); return п.length?`<div class="draftnote" style="cursor:pointer"
      onclick="switchView('exchange')">⇄ Обмен: ${esc(п.join("; ").replace(/\.$/,""))}. Нажмите — открою «Обмен».</div>`:""; })()}
    ${об.доля>0.8?`<div class="draftnote" style="cursor:pointer;border-color:var(--danger);color:var(--danger)"
      onclick="switchView('exchange')">▲ Память браузера заполнена на ${(об.доля*100).toFixed(0)}%.
      Скоро правки перестанут сохраняться — откройте «Обмен».</div>`:""}

    <div class="tiles">
      ${tile(L.length,"записей журнала","accent","journal")}
      ${tile(заГод,"за "+year+" год","")}
      ${tile(live(O).length,"объектов","","objects")}
      ${tile(live(S).length,"приборов (СИ)","","si")}
      ${tile(live(P).length,"специалистов","","people")}
      ${tile(L.filter(ждётНомеров).length,"ждут номера",
        L.filter(ждётНомеров).length?"warn":"ok","docs")}
      ${плиткаВибрации()}
      ${позицииППР().length?tile(ппрМесяца().осталось,"ППР в этом месяце",
        ппрМесяца().осталось?"warn":"ok","plan"):""}
      ${(()=>{ const з=задачиНаСегодня();
        return з.всего||з.просрочено ? tile(з.просрочено||з.всего,
          з.просрочено?"задач просрочено":"задач на сегодня",
          з.просрочено?"bad":"warn","tasks") : ""; })()}
      ${(()=>{ const с=задачиСкоро();
        return с.всего ? tile(с.всего,"горит на этой неделе","warn","tasks") : ""; })()}
      ${(()=>{ const м=задачиМоегоМеста(); if(!м.всего) return "";
        return `<div class="tile ${м.просрочено||м.вернули?"bad":"warn"} click" onclick="document.getElementById('мои_задачи').scrollIntoView({behavior:'smooth'})">
          <div class="n">${м.всего}</div><div class="l">мои задачи</div></div>`; })()}
      ${(()=>{ const л=getЛаб(); if(!л.свид_до) return "";
        const [к,т]=срокЛаб(л.свид_до);
        return `<div class="tile ${к} click" onclick="switchView('lab')">
          <div class="n" style="font-size:19px">${esc(т)}</div>
          <div class="l">аттестация ЛНК</div></div>`; })()}
      ${(()=>{ const г=готовностьПродления(); if(г.областьПуста) return "";
        const гот=г.методы.filter(x=>!x.нет.length).length;
        return `<div class="tile ${гот===г.методы.length?"ok":"warn"} click" onclick="открытьГотовностьПродления()"
          title="Сводка в карточке лаборатории: допущенные, приборы с поверкой, документы по каждому методу области аттестации">
          <div class="n">${гот} из ${г.методы.length}</div><div class="l">методов готово к продлению аттестации</div></div>`; })()}
      ${(()=>{ const л=getЛаб(); if(!л.проверка && !л.проверка_пройдена) return "";
        const [к,т]=состояниеПроверкиЛНК(л);
        return `<div class="tile ${к} click" onclick="switchView('lab')">
          <div class="n" style="font-size:${т.length>24?14:19}px">${esc(т)}</div>
          <div class="l">плановая проверка ЛНК</div></div>`; })()}
    </div>
    <div id="мои_задачи">${блокМоихЗадач()}</div>
    ${блокРаботПоМесяцам()}
    ${блокПоследнихПравок()}
    ${(()=>{ const з=задачиНаСегодня();
      if(!з.всего && !з.просрочено) return "";
      return `<div class="panel" style="margin-bottom:14px">
        <h3>✓ Задачи</h3>
        <div class="bar"><span class="g" style="flex:1">Срок сегодня или раньше</span>
          <b>${з.всего}</b></div>
        ${з.просрочено?`<div class="bar"><span class="g" style="flex:1">Из них просрочено</span>
          <b style="color:var(--danger)">${з.просрочено}</b></div>`:""}
        <div class="plist">${з.список.slice(0,6).map(x=>
          `<div class="pr" onclick="editTask('${esc(x.id)}')">
             <span class="st ${состКласс(x)}">${esc(состТекст(x))}</span>
             <span class="g">${esc(x.что)}</span>
             <span class="num">${ru(x.срок)}</span></div>`).join("")}
          ${з.всего>6?`<div class="pr" style="cursor:pointer" onclick="switchView('tasks')">
             <span class="g">…и ещё ${з.всего-6}</span></div>`:""}</div>
      </div>`; })()}
    ${(()=>{ const м=ппрМесяца();
      if(!позицииППР().length || !м.всего) return "";
      return `<div class="panel" style="margin-bottom:14px">
        <h3>◷ Обслуживание приборов в этом месяце</h3>
        <div class="bar"><span class="g" style="flex:1">По плану ППР на ${esc(м.имя)}</span>
          <b>${м.всего}</b></div>
        <div class="bar"><span class="g" style="flex:1">Отмечено выполненными</span><b>${м.сделано}</b></div>
        <div class="bf"><span style="width:${м.всего?Math.round(м.сделано/м.всего*100):0}%"></span></div>
        ${м.осталось?`<div class="plist">`+м.список.slice(0,6).map(р=>
          `<div class="pr" onclick="switchView('plan')">
             <span class="st work">${р.часы} ч</span>
             <span class="g">${esc(р.поз.наименование)}</span>
             <span class="num">${esc(р.поз.вид)}</span></div>`).join("")+
          (м.осталось>6?`<div class="pr" style="cursor:pointer" onclick="switchView('plan')">
             <span class="g">…и ещё ${м.осталось-6}</span></div>`:"")+`</div>`
          :`<div class="note">Всё за месяц отмечено выполненным.</div>`}
      </div>`; })()}

    <div class="cols">
      <div class="panel">
        <h3>⌗ Поверка приборов</h3>
        ${поверкаHTML}
      </div>
      <div class="panel">
        <h3>▤ Записи по методам <span class="pc">всего ${L.length}</span></h3>
        ${mc.map(x=>`<div class="bar"><span class="bn">${x.m}</span>
          <span class="bt"><span class="bf" style="width:${Math.round(x.n/mmax*100)}%"></span></span>
          <span class="bc">${x.n}</span></div>`).join("")}
      </div>
    </div>

    <div class="cols">
      <div class="panel">
        <h3>◷ По статусу</h3>
        ${[["Зарегистрировано","ok"],["В работе / черновик","work"],["Новые","new"],["Аннулировано","bad"]]
          .map(([l,c])=>`<div class="bar"><span class="bt" style="max-width:none"><span class="bf ${c}" style="background:var(--${c==='ok'?'ok':c==='work'?'warn':c==='bad'?'danger':'brand'});width:${Math.round(st(c)/Math.max(1,L.length)*100)}%"></span></span><span class="bc">${st(c)}</span></div>`).join("")}
        <div class="note">Зарегистрировано ${st("ok")}, в работе/черновик ${st("work")}, новых ${st("new")}, аннулировано ${st("bad")}.</div>
      </div>
      <div class="panel">
        <h3>◈ Последние записи</h3>
        <div class="plist">${last.map(z=>`<div class="pr" onclick="showJournal('${esc(z.id)}')">
          <span class="num">${ru(z.дата)}</span>
          <span class="g">${esc(z.объект)}</span>
          <span>${methodsOf(z).map(m=>`<span class="pill">${m}</span>`).join("")}</span></div>`).join("")
          ||`<div class="note">Записей нет.</div>`}</div>
      </div>
    </div>`;
}

