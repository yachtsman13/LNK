// Проверка программы ЛНК целиком — одной командой перед каждой выдачей файла:
//
//     python3 scripts/build_app.py && node scripts/проверка_программы.js
//
// Открывает собранный Приложение/ЛНК.html в безголовом Chromium (как у
// заказчика — с диска, file://) и прогоняет то, что до сих пор проверялось
// разовыми скриптами во временной папке сессии и пропадало вместе с ней:
//
//   1. запуск: база из IndexedDB, все разделы открываются без ошибок;
//   2. «Проверка данных»: нет ошибок, кроме известного состояния данных
//      заказчика (просроченная поверка приборов), нет реестров мимо копии и
//      полей на служебных ключах;
//   3. все выгрузки Word/Excel собираются и проходят
//      scripts/проверка_docx.py и scripts/проверка_xlsx.py;
//   4. хранилище: правка переживает перезагрузку;
//   5. перенос базы из localStorage (как у заказчика при первом запуске
//      версии 3.30) и правки старой версии, сделанные после переноса;
//   6. обмен между двумя рабочими местами (у каждого своё хранилище):
//      правка, новая запись, повторный приём, перенос удаления;
//   7. резервная копия: сохранить → стереть всё → восстановить.
//
// Итог — список ✓/✗ и код выхода: 0 — всё прошло, 1 — есть провалы.
// Файлы выгрузок остаются во временной папке, путь печатается в конце.
//
// Где что лежит (переопределяется переменными окружения):
//   PW_PATH — модуль playwright (по умолчанию $(npm root -g)/playwright);
//   CHROME  — Chromium (по умолчанию /opt/pw-browsers/chromium).
"use strict";
const path = require("path"), fs = require("fs"), os = require("os");
const { execSync, spawnSync } = require("child_process");

const КОРЕНЬ = path.resolve(__dirname, "..");
const ПРОГРАММА = "file://" + path.join(КОРЕНЬ, "Приложение", "ЛНК.html");
const CHROME = process.env.CHROME || "/opt/pw-browsers/chromium";
// Латиницей нарочно: Playwright молча не подаёт в поле выбора файл, если
// в его пути есть кириллица (замерено — событие change просто не приходит).
const ВЫХОД = fs.mkdtempSync(path.join(os.tmpdir(), "lnk-check-"));
// Ошибки «Проверки данных», которые есть в самих присланных данных, а не в
// программе. Появится другая — это провал.
const ИЗВЕСТНЫЕ_ОШИБКИ = [/приборов с просроченной поверкой/];
// Разделы программы (switchView).
const РАЗДЕЛЫ = ["dash","find","journal","objects","si","people","docs","exchange","plan",
  "check","podr","nd","premises","lab","tasks","passport","metod","calendar"];

function playwright(){
  if(process.env.PW_PATH) return require(process.env.PW_PATH);
  const корень = execSync("npm root -g").toString().trim();
  return require(path.join(корень, "playwright"));
}

let провалов = 0;
const итоги = [];
function раздел(имя){ итоги.push("\n" + имя); console.log("…", имя); }
function да(усл, что, подробно){
  итоги.push((усл ? "  ✓ " : "  ✗ ") + что + (!усл && подробно != null ? " — " + подробно : ""));
  if(!усл) провалов++;
}

// Открыть программу и дождаться конца загрузки. Выгрузки через dlBlob
// перехватываются в окно (__ф), чтобы забрать их файлами.
async function открыть(p){
  await p.goto(ПРОГРАММА);
  await p.waitForFunction(() => window.ЛНК_ГОТОВ, null, { timeout: 30000 });
  await p.evaluate(() => {
    window.__ф = [];
    window.dlBlob = (имя, blob) => { const fr = new FileReader();
      fr.onload = () => __ф.push({ имя, b64: fr.result.split(",")[1] }); fr.readAsDataURL(blob); };
  });
}
// Рабочее место — отдельный контекст браузера, то есть своё хранилище, как
// отдельный компьютер заказчика.
async function место(b, код){
  const c = await b.newContext({ viewport: { width: 1400, height: 900 }, acceptDownloads: true });
  const p = await c.newPage();
  p.ошибки = [];
  p.on("pageerror", e => p.ошибки.push(e.message));
  p.on("dialog", d => d.accept());
  await открыть(p);
  if(код){
    await p.evaluate(к => { saveWS({ код: к, мастер: false, начальник: false }); return ХР.дождаться(); }, код);
  }
  return { c, p };
}
async function перезагрузить(p){ await p.evaluate(() => ХР.дождаться()); await открыть(p); }
// Файлы, выгруженные через dlBlob (Word, Excel).
async function собратьФайлы(p){
  await p.waitForTimeout(1000);
  const ф = await p.evaluate(() => { const x = window.__ф; window.__ф = []; return x; });
  return ф.map(f => { const п = path.join(ВЫХОД, f.имя.replace(/[\/\\]/g, "_"));
    fs.writeFileSync(п, Buffer.from(f.b64, "base64")); return п; });
}
// Файл, который программа отдаёт обычной загрузкой (копия, обмен).
async function скачать(p, вызов, имя){
  const [з] = await Promise.all([p.waitForEvent("download"), p.evaluate(вызов)]);
  const п = path.join(ВЫХОД, имя); await з.saveAs(п); return п;
}
// Подать файл туда, где программа просит выбрать его. Окно выбора
// безголовый Chromium без нажатия человека не открывает, поэтому
// запоминаем поле, по которому программа «щёлкнула», и кладём файл прямо
// в него — дальше работает её собственный обработчик, как у заказчика.
async function подать(p, вызов, путь){
  await p.evaluate(() => { window.__поле = null;
    if(!HTMLInputElement.prototype.__щёлк){
      HTMLInputElement.prototype.__щёлк = HTMLInputElement.prototype.click;
      HTMLInputElement.prototype.click = function(){ if(this.type === "file"){ window.__поле = this; return; } return this.__щёлк(); };
    } });
  await p.evaluate(вызов);
  const поле = await p.evaluateHandle(() => window.__поле);
  const лат = path.join(ВЫХОД, "podat-" + Date.now() + path.extname(путь));   // см. ВЫХОД
  fs.copyFileSync(путь, лат);
  await поле.setInputFiles(лат);
  await p.waitForTimeout(1000);
}

async function главное(){
  const b = await playwright().chromium.launch({ executablePath: CHROME });
  const все = [];
  try{
    // ── 1. Запуск и разделы ──────────────────────────────────
    раздел("1. Запуск и разделы");
    const m1 = await место(b, "М1"); все.push(m1.p);
    const p = m1.p;
    const св = await p.evaluate(() => ({ режим: ХР.режим(), журнал: liveJ().length,
      объекты: live(O).length, приборы: live(S).length, версия: DB.версия }));
    да(св.режим === "idb", "база в IndexedDB", "режим " + св.режим);
    да(св.журнал > 0 && св.объекты > 0 && св.приборы > 0, "данные загружены: журнал " + св.журнал +
      ", объекты " + св.объекты + ", приборы " + св.приборы);
    for(const v of РАЗДЕЛЫ){ await p.evaluate(x => switchView(x), v); await p.waitForTimeout(60); }
    да(p.ошибки.length === 0, "все " + РАЗДЕЛЫ.length + " разделов открываются без ошибок", p.ошибки.join("; "));

    // ── 2. Проверка данных ───────────────────────────────────
    раздел("2. «Проверка данных»");
    const пр = await p.evaluate(() => ({
      ошибки: проверки().filter(x => x.ур === "ошибка").map(x => x.текст),
      мимо: реестрыМимоКопии(), служебные: полеНаСлужебномКлюче() }));
    const лишние = пр.ошибки.filter(t => !ИЗВЕСТНЫЕ_ОШИБКИ.some(r => r.test(t)));
    да(лишние.length === 0, "ошибок сверх известных нет (всего ошибок: " + пр.ошибки.length + ")", лишние.join("; "));
    да(пр.мимо.length === 0, "реестров мимо резервной копии нет", пр.мимо.join(", "));
    да(пр.служебные.length === 0, "полей на служебных ключах нет", пр.служебные.join(", "));

    // ── 3. Выгрузки ──────────────────────────────────────────
    раздел("3. Выгрузки Word и Excel");
    const вызовы = await p.evaluate(() => {
      const out = []; const t = (n, f) => { try{ f(); out.push([n, ""]); }catch(e){ out.push([n, e.message]); } };
      t("вся база", () => exportXlsxВсё());
      ["journal","objects","si","people","podr","nd","premises"].forEach(v => t("Excel " + v, () => { curView = v; exportXlsx(); }));
      t("паспорт", () => паспортВWord()); t("план ППР Word", () => exportПланWord()); t("план ППР Excel", () => exportПланXlsx());
      t("проверка данных", () => exportПроверкуXlsx()); t("задачи", () => exportЗадачиXlsx());
      t("нагрузка", () => exportНагрузкаXlsx()); t("сводка подразделений", () => exportСводкаПодрXlsx());
      t("заявка на поверку", () => заявкаПоверкиWord()); t("план-график поверки", () => планПоверкиWord());
      t("шаблон импорта", () => шаблонИмпортаXlsx("si")); t("реестр актов", () => exportРеестрАктов());
      t("итоги года", () => exportИтогиГода(new Date().getFullYear()));
      const а = списокАктов().find(x => !готовностьАкта(x.ключ).some(y => y.ур === "ошибка"));
      if(а) t("акт", () => актВWord(а.ключ));
      return out; });
    вызовы.forEach(([n, ош]) => да(!ош, "выгрузка «" + n + "» без исключений", ош));
    const ф = await собратьФайлы(p);
    да(ф.length >= вызовы.length - 1, "получено файлов: " + ф.length);
    let плохих = 0;
    for(const п of ф){
      const пров = п.endsWith(".docx") ? "проверка_docx.py" : п.endsWith(".xlsx") ? "проверка_xlsx.py" : null;
      if(!пров) continue;
      const r = spawnSync("python3", [path.join(КОРЕНЬ, "scripts", пров), п], { encoding: "utf8" });
      if(r.status !== 0){ плохих++; да(false, "файл «" + path.basename(п) + "»", (r.stdout || r.stderr).trim().split("\n").slice(0, 3).join(" / ")); }
    }
    да(плохих === 0, "все файлы проходят проверка_docx.py / проверка_xlsx.py");

    // ── 4. Хранилище ─────────────────────────────────────────
    раздел("4. Хранилище");
    await p.evaluate(() => { const s = live(S)[0]; s.примечание = "проверка хранилища"; saveReg(REG_SI, S); });
    await перезагрузить(p);
    const пер = await p.evaluate(() => ({ прим: live(S)[0].примечание, место: getWS().код, ждёт: ХР.ожидает() }));
    да(пер.прим === "проверка хранилища", "правка прибора пережила перезагрузку");
    да(пер.место === "М1", "настройка места пережила перезагрузку");

    // ── 5. Перенос из localStorage ───────────────────────────
    раздел("5. Перенос базы из localStorage (первый запуск версии с IndexedDB)");
    {
      const c = await b.newContext(); const q = await c.newPage(); q.ошибки = [];
      q.on("pageerror", e => q.ошибки.push(e.message)); q.on("dialog", d => d.accept()); все.push(q);
      // Старая версия держала всё в localStorage. Кладём туда базу, как у
      // заказчика, с чистой страницы того же file:// — без запуска программы.
      const пусто = path.join(ВЫХОД, "пусто.html"); fs.writeFileSync(пусто, "<!doctype html><title>.</title>");
      await q.goto("file://" + пусто);
      await q.evaluate(() => {
        localStorage.setItem("lnk_ws", JSON.stringify({ код: "М7", мастер: false, начальник: true }));
        localStorage.setItem("lnk_pom", JSON.stringify({ версия: "3.28", rows: [
          { id: "ПМ-ПЕРЕНОС", идентификация: "99", назначение: "Перенесённое помещение", изменено: "2026-01-01T00:00:00.000Z", место: "М7" }] }));
      });
      await открыть(q);
      const а = await q.evaluate(() => ({ режим: ХР.режим(), место: getWS().код,
        пом: live(ПОМ).map(x => x.идентификация), вLS: !!ХР.прочестьLS("lnk_pom"), старая: ХР.стараяВерсияПравила() }));
      да(а.режим === "idb", "после переноса база в IndexedDB");
      да(а.место === "М7" && а.пом.includes("99"), "настройки и данные из localStorage перенесены", JSON.stringify(а));
      да(а.вLS, "localStorage оставлен нетронутой копией для старой версии");
      да(!а.старая, "сразу после переноса предупреждения о старой версии нет");
      // Старая версия программы, открытая после переноса, пишет в localStorage.
      await q.evaluate(() => ХР.дождаться());
      await q.goto("file://" + пусто);
      await q.evaluate(() => {
        const r = JSON.parse(localStorage.getItem("lnk_pom"));
        r.rows.push({ id: "ПМ-СТАРАЯ", идентификация: "98", назначение: "Заведено в старой версии",
          изменено: new Date().toISOString(), место: "М7" });
        localStorage.setItem("lnk_pom", JSON.stringify(r));
      });
      await открыть(q);
      const б = await q.evaluate(() => ({ старая: ХР.стараяВерсияПравила(),
        вПроверке: проверки().some(x => /старую версию/.test(x.текст)), пом: live(ПОМ).map(x => x.идентификация) }));
      да(б.старая && б.вПроверке, "правки старой версии замечены (и в «Проверке данных»)");
      да(!б.пом.includes("98"), "до принятия они в работе не участвуют");
      await q.evaluate(() => принятьПравкиСтаройВерсии());
      await перезагрузить(q);
      const в = await q.evaluate(() => ({ старая: ХР.стараяВерсияПравила(), пом: live(ПОМ).map(x => x.идентификация) }));
      да(в.пом.includes("98") && в.пом.includes("99"), "правки старой версии приняты слиянием", JSON.stringify(в.пом));
      да(!в.старая, "после принятия предупреждение ушло и не вернулось после перезагрузки");
      да(q.ошибки.length === 0, "без ошибок на странице", q.ошибки.join("; "));
      await c.close();
    }

    // ── 6. Обмен между двумя местами ─────────────────────────
    раздел("6. Обмен между двумя местами");
    const m2 = await место(b, "М2"); все.push(m2.p);
    const p2 = m2.p;
    await p.evaluate(() => {
      const s = live(S)[1]; s.примечание = "правка с М1"; s.изменено = new Date().toISOString(); s.место = "М1"; saveReg(REG_SI, S);
      ПОМ.push({ id: newId("ПМ"), идентификация: "77", назначение: "Новое с М1", изменено: new Date().toISOString(), место: "М1" });
      saveReg(REG_ПОМ, ПОМ);
    });
    const обм1 = await скачать(p, () => exchangeExport(), "обмен-М1.json");
    await подать(p2, () => exchangeImport(), обм1);
    const г = await p2.evaluate(() => ({ прим: live(S)[1].примечание, пом: live(ПОМ).filter(x => x.идентификация === "77").length,
      приборов: S.length, помещений: ПОМ.length }));
    да(г.прим === "правка с М1", "правка прибора дошла на М2");
    да(г.пом === 1, "новое помещение дошло на М2");
    await подать(p2, () => exchangeImport(), обм1);
    const г2 = await p2.evaluate(() => ({ приборов: S.length, помещений: ПОМ.length }));
    да(г2.приборов === г.приборов && г2.помещений === г.помещений, "повторный приём того же файла ничего не задвоил");
    await p2.evaluate(() => { const x = ПОМ.find(x => x.идентификация === "77"); x.удалено = new Date().toISOString(); x.место = "М2"; saveReg(REG_ПОМ, ПОМ); });
    const обм2 = await скачать(p2, () => exchangeExport(), "обмен-М2.json");
    await подать(p, () => exchangeImport(), обм2);
    да(await p.evaluate(() => !live(ПОМ).some(x => x.идентификация === "77")), "удаление с М2 перенесено на М1");
    await перезагрузить(p2);
    да(await p2.evaluate(() => live(S)[1].примечание === "правка с М1"), "принятое на М2 пережило перезагрузку");

    // ── 7. Резервная копия ───────────────────────────────────
    раздел("7. Резервная копия");
    const до = await p.evaluate(() => ({ ж: liveJ().length, о: live(O).length, п: live(S).length, нд: live(НД).length,
      прим: live(S)[0].примечание }));
    const копия = await скачать(p, () => backupSave(), "копия-М1.json");
    await p.evaluate(() => { const кк = []; for(let i = 0; i < ХР.length; i++) кк.push(ХР.key(i));
      кк.forEach(k => { if(k !== "lnk_ws") ХР.removeItem(k); }); });
    await перезагрузить(p);
    да(await p.evaluate(() => live(S)[0].примечание !== "проверка хранилища"), "после стирания памяти правок нет (программа с чистого листа)");
    await подать(p, () => backupOpen(), копия);
    await перезагрузить(p);
    const после = await p.evaluate(() => ({ ж: liveJ().length, о: live(O).length, п: live(S).length, нд: live(НД).length,
      прим: live(S)[0].примечание }));
    да(JSON.stringify(после) === JSON.stringify(до), "копия восстановила всё и это пережило перезагрузку",
      "было " + JSON.stringify(до) + ", стало " + JSON.stringify(после));

    раздел("Ошибки на страницах");
    все.forEach((x, i) => да(x.ошибки.length === 0, "страница " + (i + 1) + " без ошибок JavaScript", x.ошибки.join("; ")));
  }finally{ await b.close(); }
}

главное().then(() => {
  console.log(итоги.join("\n"));
  console.log("\nФайлы выгрузок: " + ВЫХОД);
  console.log(провалов ? "\nПРОВАЛОВ: " + провалов : "\nВСЁ ПРОШЛО");
  process.exit(провалов ? 1 : 0);
}).catch(e => { console.log(итоги.join("\n")); console.error("\nПРОВЕРКА ОБОРВАЛАСЬ:", e); process.exit(1); });
