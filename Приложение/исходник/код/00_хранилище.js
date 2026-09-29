// ── Хранилище базы: IndexedDB с зеркалом в памяти ─────────
// До 3.30 база жила в localStorage: предел ~5 МБ, и он молчаливый —
// браузер просто отказывает в записи. С историей правок (3.16) и
// отброшенными версиями (3.13) база растёт с каждой правкой, поэтому она
// перенесена в IndexedDB (в file:// работает, квота браузера — сотни МБ
// и больше, замерено: navigator.storage.estimate отдаёт квоту и из file://).
//
// Программа написана на синхронном чтении (localStorage.getItem везде), а
// IndexedDB асинхронная. Поэтому ХР — та же пятёрка методов, что у
// localStorage, над картой в памяти: при запуске вся база один раз
// читается из IndexedDB в память (основной код стартует только после
// этого, см. запуск внизу страницы), чтение идёт из памяти, запись — в
// память сразу и в IndexedDB следом.
//
// Первый запуск новой версии переносит ключи lnk_* из localStorage в
// IndexedDB и ОСТАВЛЯЕТ localStorage как было — нетронутой копией: вдруг
// на этом же компьютере откроют старый файл программы, он увидит свои
// данные, а не пустоту. Правки, сделанные в старой версии после переноса,
// новая версия замечает по отпечатку localStorage и предлагает принять
// (см. правкиСтаройВерсии).
//
// Если IndexedDB недоступна (запрет браузера, частный режим) — работа
// по-старому прямо в localStorage, с прежним пределом и счётчиком.
const ХР=(function(){
  const БД="lnk_data", ХРАН="kv", МЕТКА="__перенос";
  const м=new Map();
  let db=null, режим="ls", причина="", метка=null, ожидает=0, сообщено=false;
  const висят=new Set();
  let квота=null, занято=null;
  function отпечатокLS(){
    const кк=[];
    try{ for(let i=0;i<localStorage.length;i++){ const k=localStorage.key(i);
      if(k && k.indexOf("lnk_")===0) кк.push(k); } }catch(e){ return ""; }
    кк.sort();
    let h=0x811c9dc5, n=0;
    кк.forEach(k=>{ const s=k+"\u0000"+(localStorage.getItem(k)||"")+"\u0001"; n+=s.length;
      for(let i=0;i<s.length;i++){ h^=s.charCodeAt(i); h=Math.imul(h,16777619); } });
    return кк.length+":"+n+":"+(h>>>0).toString(16);
  }
  function открыть(){
    return new Promise((ок,нет)=>{
      let r; try{ r=indexedDB.open(БД,1); }catch(e){ return нет(e); }
      r.onupgradeneeded=()=>{ try{ r.result.createObjectStore(ХРАН); }catch(e){} };
      r.onsuccess=()=>ок(r.result); r.onerror=()=>нет(r.error);
      r.onblocked=()=>нет(new Error("база занята другой вкладкой"));
    });
  }
  function прочестьВсё(){
    return new Promise((ок,нет)=>{
      const t=db.transaction(ХРАН,"readonly"), st=t.objectStore(ХРАН);
      const кл=st.getAllKeys(), зн=st.getAll();
      t.oncomplete=()=>ок([кл.result||[], зн.result||[]]);
      t.onerror=()=>нет(t.error); t.onabort=()=>нет(t.error);
    });
  }
  function записатьПачку(пары){
    return new Promise((ок,нет)=>{
      const t=db.transaction(ХРАН,"readwrite"), st=t.objectStore(ХРАН);
      пары.forEach(([k,v])=>st.put(v,k));
      t.oncomplete=()=>ок(true); t.onerror=()=>нет(t.error); t.onabort=()=>нет(t.error);
    });
  }
  // Отказ записи — правка НЕ сохранена. Как и в прежнем переполнении,
  // молчать нельзя: окно, а не подсказка внизу. Одно на сеанс, чтобы не
  // засыпать окнами при серии сохранений.
  function сбой(e){
    if(сообщено) return; сообщено=true;
    alert("ЗАПИСЬ В ХРАНИЛИЩЕ БРАУЗЕРА НЕ УДАЛАСЬ — последняя правка может быть не сохранена.\n\n"+
      "Причина: "+((e&&(e.name||e.message))||"неизвестна")+"\n\n"+
      "Сделайте сейчас «Обмен» → «Сохранить копию» и перезапустите программу.");
  }
  function вБазу(k,v){
    ожидает++;
    const обещ=new Promise(готово=>{
      let конец=false;
      const итог=(ош)=>{ if(конец) return; конец=true; ожидает--; висят.delete(обещ); if(ош) сбой(ош); готово(); };
      try{
        const t=db.transaction(ХРАН,"readwrite"), st=t.objectStore(ХРАН);
        if(v===null) st.delete(k); else st.put(v,k);
        t.oncomplete=()=>итог(null); t.onerror=()=>итог(t.error||new Error("ошибка записи"));
        t.onabort=()=>итог(t.error||new Error("запись прервана"));
      }catch(e){ итог(e); }
    });
    висят.add(обещ);
  }
  async function загрузить(){
    try{
      if(typeof indexedDB==="undefined" || !indexedDB) throw new Error("IndexedDB нет в этом браузере");
      db=await Promise.race([открыть(),
        new Promise((_,нет)=>setTimeout(()=>нет(new Error("IndexedDB не ответила за 5 с")),5000))]);
      const [кл,зн]=await прочестьВсё();
      кл.forEach((k,i)=>{ k=String(k); if(k===МЕТКА){ try{ метка=JSON.parse(зн[i]); }catch(e){ метка={}; } }
                          else м.set(k, String(зн[i])); });
      if(!метка){
        // Первый запуск версии с IndexedDB: переносим, что было.
        const пары=[];
        for(let i=0;i<localStorage.length;i++){ const k=localStorage.key(i);
          if(k && k.indexOf("lnk_")===0 && !м.has(k)){ const v=localStorage.getItem(k); if(v!=null){ м.set(k,v); пары.push([k,v]); } } }
        метка={ когда:new Date().toISOString(), ключей:пары.length, отпечаток:отпечатокLS() };
        пары.push([МЕТКА, JSON.stringify(метка)]);
        await записатьПачку(пары);
      }
      режим="idb";
      try{ if(navigator.storage && navigator.storage.estimate){ const е=await navigator.storage.estimate();
             квота=е.quota||null; занято=(е.usage!=null)?е.usage:null; } }catch(e){}
    }catch(e){
      режим="ls"; причина=String((e&&(e.message||e.name))||e); db=null; м.clear(); метка=null;
    }
  }
  window.addEventListener("beforeunload", e=>{ if(ожидает>0){ e.preventDefault(); e.returnValue=""; } });
  const ls=()=>localStorage;
  return {
    загрузить,
    getItem(k){ if(режим!=="idb") return ls().getItem(k); k=String(k); return м.has(k)?м.get(k):null; },
    setItem(k,v){ if(режим!=="idb") return ls().setItem(k,v); k=String(k); v=String(v); м.set(k,v); вБазу(k,v); },
    removeItem(k){ if(режим!=="idb") return ls().removeItem(k); k=String(k); if(!м.has(k)) return; м.delete(k); вБазу(k,null); },
    key(i){ if(режим!=="idb") return ls().key(i); return [...м.keys()][i]||null; },
    get length(){ return режим!=="idb" ? ls().length : м.size; },
    режим(){ return режим; }, причина(){ return причина; }, метка(){ return метка; },
    ожидает(){ return ожидает; }, квота(){ return квота; }, занято(){ return занято; },
    // Все начатые записи дошли до IndexedDB (для проверок и перед
    // перезагрузкой; в работе программы ждать не нужно).
    дождаться(){ return Promise.all([...висят]); },
    // Правки в localStorage после переноса = открывали старую версию.
    стараяВерсияПравила(){ return режим==="idb" && !!метка && !!метка.отпечаток && отпечатокLS()!==метка.отпечаток; },
    // Правки старой версии приняты (или отклонены) — запомнить нынешнее
    // состояние localStorage, чтобы не спрашивать снова.
    запомнитьСтаруюВерсию(){ if(режим!=="idb"||!метка) return Promise.resolve(false);
      метка=Object.assign({},метка,{отпечаток:отпечатокLS(), принято:new Date().toISOString()});
      return записатьПачку([[МЕТКА, JSON.stringify(метка)]]); },
    прочестьLS(k){ try{ return localStorage.getItem(k); }catch(e){ return null; } }
  };
})();
