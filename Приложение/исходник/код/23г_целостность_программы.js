// ── Целостность файла программы при переносе и обновлении ──────────
// SHA-256 нужен только против случайной порчи или неполного копирования.
// Это не подпись издателя: рядом с HTML автоматически лежит <имя>.sha256.
// Функции ниже заменяют одноимённые из 23б_защита_обмен.js, сохраняя их
// порядок выбора папок, резервную копию и перезагрузку без новых шагов.
function имяХешаПрограммы(имя){ return имя+".sha256"; }
async function хешПрограммы(данные){
  if(!window.crypto || !crypto.subtle) throw new Error("Web Crypto API недоступен");
  const х=await crypto.subtle.digest("SHA-256",данные);
  return Array.from(new Uint8Array(х),б=>б.toString(16).padStart(2,"0")).join("");
}
async function ожидаемыйХешПрограммы(папка,имя){
  try{
    const ф=await папка.getFileHandle(имяХешаПрограммы(имя));
    const m=/^\s*([0-9a-f]{64})\s*$/i.exec(await (await ф.getFile()).text());
    return m?m[1].toLowerCase():"";
  }catch(e){ return ""; }
}
async function программуНаФлешку(п, спросить){
  const есть=await найтиВерсиюПрограммы(п);
  if(есть && верс(есть.версия)>=верс(DB.версия)) return "";
  const имя=имяФайлаПрограммы(); if(!/\.html?$/i.test(имя)) return "";
  const пп=await идбЧтение("папкаПрограммы");
  let право="нет";
  if(пп){ try{ право=await пп.queryPermission({mode:"read"}); if(право!=="granted" && спросить) право=await пп.requestPermission({mode:"read"}); }catch(e){ право="нет"; } }
  if(!пп || право!=="granted") return "нет папки";
  try{
    const данные=await (await (await пп.getFileHandle(имя)).getFile()).arrayBuffer();
    const в=версияИзТекста(new TextDecoder().decode(данные.slice(0,4096)));
    if(!в || в.версия!==DB.версия) return "не та версия";
    const хеш=await хешПрограммы(данные);
    const пр=await п.getDirectoryHandle(ПАПКА_ПРОГРАММЫ_ФЛЕШКИ,{create:true});
    const ф=await пр.getFileHandle(имя,{create:true}); const w=await ф.createWritable(); await w.write(данные); await w.close();
    const записанное=await (await ф.getFile()).arrayBuffer();
    if(await хешПрограммы(записанное)!==хеш) return "ошибка копии";
    await записатьТекст(пр,имяХешаПрограммы(имя),хеш+"\n");
    return "положена";
  }catch(e){ return "нет папки"; }
}
async function найтиВерсиюПрограммы(папка, корень){
  let лучшая=null;
  const где=[папка];
  try{ где.push(await папка.getDirectoryHandle(ПАПКА_ПРОГРАММЫ_ФЛЕШКИ)); }catch(e){}
  if(корень && корень!==папка) где.push(корень);
  for(const п of где){
    try{
      for await (const [имя,х] of п.entries()){
        if(х.kind!=="file" || !/\.html?$/i.test(имя)) continue;
        try{
          const данные=await (await х.getFile()).arrayBuffer();
          const в=версияИзТекста(new TextDecoder().decode(данные.slice(0,4096)));
          const ожидаемый=await ожидаемыйХешПрограммы(п,имя);
          if(в && ожидаемый && await хешПрограммы(данные)===ожидаемый &&
            (!лучшая || верс(в.версия)>верс(лучшая.версия))) лучшая=Object.assign({имя,h:х,папка:п},в);
        }catch(e){}
      }
    }catch(e){}
  }
  return лучшая;
}
async function обновитьПрограмму(новый){
  const имя=имяФайлаПрограммы();
  if(!/\.html?$/i.test(имя)){ alert("Не понять, как называется файл программы. Замените его вручную."); return false; }
  if(!новый){ toast("Нового файла программы нет"); return false; }
  if(верс(новый.версия)<=верс(DB.версия)){ toast("Это не новее вашей версии "+DB.версия); return false; }
  let п=await идбЧтение("папкаПрограммы");
  let право="нет";
  if(п){ try{ право=await п.queryPermission({mode:"readwrite"}); if(право!=="granted") право=await п.requestPermission({mode:"readwrite"}); }catch(e){ право="нет"; } }
  // Папка программы не запомнена — пробуем уже выбранную папку с документами
  // (обычно это корень флешки, где лежит и программа).
  if(!п || право!=="granted" || !(await естьФайлВ(п,имя))){
    if(typeof _папкаДок!=="undefined" && _папкаДок){
      try{
        if(await естьФайлВ(_папкаДок,имя)){
          let правоДок="denied";
          try{ правоДок=await _папкаДок.queryPermission({mode:"readwrite"}); }catch(e){}
          if(правоДок==="granted"){ п=_папкаДок; }
          else {
            try{ правоДок=await _папкаДок.requestPermission({mode:"readwrite"}); }catch(e){}
            if(правоДок==="granted"){ п=_папкаДок; }
          }
        }
      }catch(e){}
    }
    if(!п || !(await естьФайлВ(п,имя))){
      alert("Выберите папку на этом компьютере, где лежит файл программы «"+имя+"».\n\nСейчас программа открыта отсюда:\n"+путьПрограммы());
      try{ п=await window.showDirectoryPicker({mode:"readwrite", id:"lnk-programma"}); }
      catch(e){ if(e.name!=="AbortError") toast("Не удалось выбрать папку: "+e.name); return false; }
      if(!(await естьФайлВ(п,имя))){ alert("В папке «"+п.name+"» нет файла «"+имя+"». Выберите папку, где он лежит."); return false; }
      await идбЗапись("папкаПрограммы",п);
    }
  }
  if(!уходСФормы()) return false;
  try{ await ХР.дождаться(); }catch(e){}
  try{
    const данные=await (await новый.h.getFile()).arrayBuffer();
    const в=версияИзТекста(new TextDecoder().decode(данные.slice(0,4096)));
    if(!в || в.версия!==новый.версия){ alert("Файл на флешке не похож на программу версии "+новый.версия+" — не заменяю."); return false; }
    const ожидаемый=await ожидаемыйХешПрограммы(новый.папка,новый.имя);
    if(!ожидаемый){ alert("Рядом с новой программой нет читаемого файла «"+имяХешаПрограммы(новый.имя)+"» — не заменяю. Положите программу на флешку ещё раз новой версией ЛНК."); return false; }
    if(await хешПрограммы(данные)!==ожидаемый){ alert("Контрольная сумма полного HTML не совпала — файл повреждён или скопирован не полностью. Программу не заменяю."); return false; }
    const старое=await (await (await п.getFileHandle(имя)).getFile()).arrayBuffer();
    const имяПрежней=имя.replace(/\.html?$/i,"")+" "+DB.версия+" (прежняя).html";
    const ф1=await п.getFileHandle(имяПрежней,{create:true}); const w1=await ф1.createWritable(); await w1.write(старое); await w1.close();
    const ф2=await п.getFileHandle(имя); const w2=await ф2.createWritable(); await w2.write(данные); await w2.close();
    const записанное=await (await ф2.getFile()).arrayBuffer();
    const провер=версияИзТекста(new TextDecoder().decode(записанное.slice(0,4096)));
    if(!провер || провер.версия!==новый.версия || await хешПрограммы(записанное)!==ожидаемый){
      alert("Файл записан, но проверка целостности не сошлась. Откройте «"+имяПрежней+"» и сообщите мне."); return false;
    }
    if(confirm("Программа на этом компьютере обновлена до версии "+новый.версия+".\nПрежняя сохранена рядом: «"+имяПрежней+"».\n\n"+
      "Данные не тронуты — они хранятся в браузере.\n\nОткрыть новую версию сейчас (перезагрузить страницу)?")) перезагрузитьПрограмму();
    return true;
  }catch(e){
    alert("Не удалось заменить файл программы: "+(e&&e.name||e)+"\n\nЗамените вручную: закройте вкладку, скопируйте «"+новый.имя+
      "» с флешки поверх «"+имя+"» и откройте снова.");
    return false;
  }
}
