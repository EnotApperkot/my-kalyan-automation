// ==========================================
// 1. МЕНЮ (появляется при открытии таблицы)
// ==========================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Мои инструменты')
    .addItem('Добавить накладную (Форма)', 'pokazatFormu')
    .addItem('Ввод кальянов (Форма)', 'pokazatFormuKalyans')
    .addItem('Импорт из фото (Gemini)', 'avtoImport')
    .addItem('Посчитать сводку', 'svodkaPoPostavshikam')
    .addItem('Посчитать ЗП', 'raschetZarplaty')
    .addItem('Рассчитать прибыль', 'raschetPribyli')
    .addSeparator() 
    .addItem('Создать лист Склад (Разово)', 'sozdatListSklad')
    .addItem('Пересчитать склад (Приход)', 'pereschetSklada')
    .addItem('Списать расход (Продажи)', 'spisanieRaskhoda')
    .addItem('Провести Инвентаризацию', 'provestiInventarizatsiyu')
    .addItem('ЗАПУСТИТЬ ВСЁ (main)', 'main') 
    .addToUi();
}

// ==========================================
// 2. ФОРМЫ (открытие окон)
// ==========================================
function pokazatFormu() {
  var html = HtmlService.createHtmlOutputFromFile('Index')
      .setWidth(450)
      .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Новая накладная');
}

function pokazatFormuKalyans() {
  var html = HtmlService.createHtmlOutputFromFile('IndexKalyans')
      .setWidth(400)
      .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Ввод кальянов');
}

// ==========================================
// 3. СОХРАНЕНИЕ ИЗ ФОРМ
// ==========================================
function dobavitNakladnuyu(dannye) {
  var list = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Приход Продукции");
  if (!list) return SpreadsheetApp.getUi().alert("Ошибка: Лист 'Приход Продукции' не найден!");
  var novayaStroka = list.getLastRow() + 1;
  list.getRange(novayaStroka, 1).setValue(dannye.data);       
  list.getRange(novayaStroka, 2).setValue(dannye.postavshik); 
  list.getRange(novayaStroka, 3).setValue(dannye.tovar);      
  list.getRange(novayaStroka, 4).setValue(Number(dannye.kolichestvo)); 
  list.getRange(novayaStroka, 5).setValue(Number(dannye.cena));        
  list.getRange(novayaStroka, 6).setValue(Number(dannye.kolichestvo) * Number(dannye.cena)); 
}

function dobavitKalyany(dannye) {
  var list = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Кальянщики");
  if (!list) return SpreadsheetApp.getUi().alert("Ошибка: Лист 'Кальянщики' не найден!");
  var novayaStroka = list.getLastRow() + 1;
  list.getRange(novayaStroka, 1).setValue(dannye.data);       
  list.getRange(novayaStroka, 2).setValue(dannye.sotrudnik); 
  list.getRange(novayaStroka, 3).setValue(Number(dannye.voda));      
  list.getRange(novayaStroka, 4).setValue(Number(dannye.greip)); 
}

// ==========================================
// 4. ТРИГГЕР: Автоматическая дата
// ==========================================
function onEdit(e) {
  var izmenennayaKolonka = e.range.getColumn();
  var izmenennayaStroka = e.range.getRow();
  if (izmenennayaKolonka === 2 && izmenennayaStroka > 1) {
    var yacheykaDaty = e.source.getActiveSheet().getRange(izmenennayaStroka, 1);
    if (yacheykaDaty.getValue() === "") yacheykaDaty.setValue(new Date()); 
  }
}

// ==========================================
// 5. СВОДКА ПО ПОСТАВЩИКАМ
// ==========================================
function svodkaPoPostavshikam() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var list = ss.getSheetByName("Приход Продукции"); 
  if (!list) return SpreadsheetApp.getUi().alert("Ошибка: Лист 'Приход Продукции' не найден!");
  
  var dannye = list.getDataRange().getValues(); 
  var itogi = {}; 
  
  for (var i = 1; i < dannye.length; i++) {
    var postavshik = dannye[i][1]; 
    var ochistka = String(dannye[i][5]).replace(/[^\d,.-]/g, "").replace(",", "."); 
    var summa = Number(ochistka);
    
    if (postavshik != "") {
      if (summa === 0 || isNaN(summa)) {
        list.getRange(i + 1, 7).setValue("НЕТ ЦЕНЫ!");
        continue; 
      }
      if (itogi[postavshik] == undefined) itogi[postavshik] = summa; 
      else itogi[postavshik] += summa;
    }
  }
  
  var listSvodki = ss.getSheetByName("Сводка") || ss.insertSheet("Сводка");
  listSvodki.clear(); 
  listSvodki.getRange(1, 1, 1, 2).setValues([["Поставщик", "Общая сумма"]]);
  
  var stroka = 2; 
  for (var key in itogi) {
    listSvodki.getRange(stroka, 1).setValue(key);       
    listSvodki.getRange(stroka, 2).setValue(itogi[key]); 
    stroka++; 
  }
}

// ==========================================
// 6. ИМПОРТ ИЗ JSON (Ручной)
// ==========================================
function importJSONvTabel() {
  var list = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Приход Продукции");
  if (!list) return;
  var jsonText = `[{"tovar": "Тест", "kolichestvo": 1, "cena": 100}]`;
  var dannye = JSON.parse(jsonText);
  var novayaStroka = list.getLastRow() + 1;
  
  for (var i = 0; i < dannye.length; i++) {
    var summa = dannye[i].kolichestvo * dannye[i].cena;
    list.getRange(novayaStroka + i, 1).setValue("15.10.2026"); 
    list.getRange(novayaStroka + i, 2).setValue('ООО "Продым"');    
    list.getRange(novayaStroka + i, 3).setValue(dannye[i].tovar);         
    list.getRange(novayaStroka + i, 4).setValue(dannye[i].kolichestvo);   
    list.getRange(novayaStroka + i, 5).setValue(dannye[i].cena);          
    list.getRange(novayaStroka + i, 6).setValue(summa);         
  }
}

// ==========================================
// 7. ИМПОРТ ИЗ ФОТО ЧЕРЕЗ GEMINI
// ==========================================
function avtoImport() {
  var ui = SpreadsheetApp.getUi();
  var otvetPolzovatelya = ui.prompt('Импорт из фото', 'Вставьте ссылку на файл:', ui.ButtonSet.OK_CANCEL);
  if (otvetPolzovatelya.getSelectedButton() != ui.Button.OK) return;
  
  var link = otvetPolzovatelya.getResponseText().trim();
  var fileId = link;
  if (link.indexOf('/d/') !== -1) {
    var posleD = link.split('/d/'); 
    fileId = posleD[1].split('/')[0];
  }
  
  var list = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Приход Продукции");
  if (!list) return ui.alert("Ошибка: Лист 'Приход Продукции' не найден!");
  
  try {
    var file = DriveApp.getFileById(fileId); 
    var blob = file.getBlob();               
    var tipFayla = blob.getContentType(); 
    var base64Data = Utilities.base64Encode(blob.getBytes()); 
    
    var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    var url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=" + apiKey;
    
    var prompt = "Прочитай этот документ. Выведи данные строго в формате JSON-массива. Формат: [{\"tovar\": \"название\", \"kolichestvo\": 1, \"cena\": 1550}]. ВАЖНО: цены и количества должны быть числами без пробелов и запятых. ОБЯЗАТЕЛЬНО переноси в поле 'tovar' полное название, включая вес (например, '200г', '200гр') или количество штук в пачке (например, '72шт', '80шт'). Если в документе НЕ УКАЗАН вес или штуки, напиши в конце названия ' [НЕТ ВЕСА]'. Не добавляй никакого лишнего текста, только JSON.";
    
    var payload = { "contents": [{ "parts": [ {"text": prompt}, {"inlineData": {"mimeType": tipFayla, "data": base64Data}} ] }] };
    var options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true };
    
    Logger.log("Отправили файл в Gemini, ждем ответ...");
    var response = UrlFetchApp.fetch(url, options);
    var json = JSON.parse(response.getContentText());
    
    if (json.error) {
      Logger.log("ОШИБКА СЕРВЕРА: " + JSON.stringify(json.error));
      return;
    }
    
    var otvet = json.candidates[0].content.parts[0].text;
    otvet = otvet.replace(/```json/g, "").replace(/```/g, "").trim();
    var dannye = JSON.parse(otvet);
    
    var novayaStroka = list.getLastRow() + 1;
    for (var i = 0; i < dannye.length; i++) {
      var summa = Number(dannye[i].kolichestvo) * Number(dannye[i].cena);
      list.getRange(novayaStroka + i, 1).setValue("12.11.2026"); 
      list.getRange(novayaStroka + i, 2).setValue('ООО "Продым"');    
      list.getRange(novayaStroka + i, 3).setValue(dannye[i].tovar);         
      list.getRange(novayaStroka + i, 4).setValue(Number(dannye[i].kolichestvo));   
      list.getRange(novayaStroka + i, 5).setValue(Number(dannye[i].cena));          
      list.getRange(novayaStroka + i, 6).setValue(summa);         
    }
    ui.alert("Готово!", "Успешно добавлено строк: " + dannye.length, ui.ButtonSet.OK); 
  } catch (e) {
    ui.alert("Ошибка", "Не удалось обработать файл. Проверьте Журнал.", ui.ButtonSet.OK);
  }
}

// ==========================================
// 8. СВЕРКА С БУХГАЛТЕРИЕЙ
// ==========================================
function sverkaObsshihKalyans() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var listKal = ss.getSheetByName("Кальянщики"); 
  var listBuh = ss.getSheetByName("Бухгалтерия"); 
  if (!listKal || !listBuh) return SpreadsheetApp.getUi().alert("Ошибка: Проверьте названия листов.");
  
  var dannyeKal = listKal.getDataRange().getValues();
  var totalVodaKal = 0, totalGreipKal = 0;
  for (var i = 1; i < dannyeKal.length; i++) {
    totalVodaKal += Number(dannyeKal[i][2]) || 0; 
    totalGreipKal += Number(dannyeKal[i][3]) || 0; 
  }
  
  var dannyeBuh = listBuh.getDataRange().getValues();
  var totalVodaBuh = Number(dannyeBuh[1][1]) || 0; 
  var totalGreipBuh = Number(dannyeBuh[1][2]) || 0; 
  
  if (totalVodaKal === totalVodaBuh && totalGreipKal === totalGreipBuh) {
    SpreadsheetApp.getUi().alert("ИДЕАЛЬНО! Общие цифры полностью сошлись.");
  } else {
    var textOshibki = "ВНИМАНИЕ! РАСХОЖДЕНИЕ.\n\n";
    if (totalVodaKal !== totalVodaBuh) textOshibki += "Вода: у нас " + totalVodaKal + ", а у бухгалтера " + totalVodaBuh + "\n";
    if (totalGreipKal !== totalGreipBuh) textOshibki += "Грейп: у нас " + totalGreipKal + ", а у бухгалтера " + totalGreipBuh + "\n";
    SpreadsheetApp.getUi().alert(textOshibki);
  }
}

// ==========================================
// 9. РАСЧЕТ ЗАРПЛАТЫ
// ==========================================
function parseDate(str) {
  var parts = str.split('.'); 
  if (parts.length === 3) return new Date(parts[2], parts[1] - 1, parts[0]);
  return new Date(str); 
}

function raschetZarplaty() {
  var ui = SpreadsheetApp.getUi();
  var otvet1 = ui.prompt('Расчет ЗП', 'Начало периода (ДД.ММ.ГГГГ):', ui.ButtonSet.OK_CANCEL);
  if (otvet1.getSelectedButton() != ui.Button.OK) return;
  var startStr = otvet1.getResponseText().trim();
  
  var otvet2 = ui.prompt('Расчет ЗП', 'Конец периода (ДД.ММ.ГГГГ):', ui.ButtonSet.OK_CANCEL);
  if (otvet2.getSelectedButton() != ui.Button.OK) return;
  var endStr = otvet2.getResponseText().trim();
  
  var startDate = parseDate(startStr);
  var endDate = parseDate(endStr);
  endDate.setHours(23, 59, 59); 
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var listKal = ss.getSheetByName("Кальянщики"); 
  if (!listKal) return ui.alert("Ошибка: Лист 'Кальянщики' не найден!");
  
  var dannye = listKal.getDataRange().getValues();
  var sotrudniki = {}; 
  
  for (var i = 1; i < dannye.length; i++) {
    var dataIzTablitsy = (typeof dannye[i][0] === 'string' && dannye[i][0].includes('.')) ? parseDate(dannye[i][0]) : new Date(dannye[i][0]);
    var imya = dannye[i][1];
    var voda = Number(dannye[i][2]) || 0;
    var greip = Number(dannye[i][3]) || 0;
    
    if (dataIzTablitsy >= startDate && dataIzTablitsy <= endDate && imya != "") {
      if (sotrudniki[imya] == undefined) sotrudniki[imya] = { dni: [], voda: 0, greip: 0 };
      sotrudniki[imya].dni.push(dataIzTablitsy.toDateString()); 
      sotrudniki[imya].voda += voda;
      sotrudniki[imya].greip += greip;
    }
  }
  
  var listZarplata = ss.getSheetByName("Зарплата") || ss.insertSheet("Зарплата");
  listZarplata.clear();
  listZarplata.getRange(1, 1, 1, 6).setValues([["Период", "Сотрудник", "Дней", "Вода", "Грейп", "К выплате"]]);
  
  var stroka = 2, periodText = startStr + " - " + endStr;
  for (var imya in sotrudniki) {
    var d = sotrudniki[imya];
    var unikalnyeDni = [...new Set(d.dni)].length;
    var kVyplate = (unikalnyeDni * 500) + (d.voda * 200) + (d.greip * 300);
    
    listZarplata.getRange(stroka, 1).setValue(periodText);
    listZarplata.getRange(stroka, 2).setValue(imya);
    listZarplata.getRange(stroka, 3).setValue(unikalnyeDni);
    listZarplata.getRange(stroka, 4).setValue(d.voda);
    listZarplata.getRange(stroka, 5).setValue(d.greip);
    listZarplata.getRange(stroka, 6).setValue(kVyplate);
    stroka++;
  }
  ui.alert("Зарплата за период " + periodText + " рассчитана!");
}

// ==========================================
// 10. РАСЧЕТ ПРИБЫЛИ
// ==========================================
function raschetPribyli() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var listBuh = ss.getSheetByName("Бухгалтерия");
  var listSklad = ss.getSheetByName("Склад");
  var listZarplata = ss.getSheetByName("Зарплата");
  
  if (!listBuh || !listSklad) {
    ui.alert("Ошибка: Нет листа 'Бухгалтерия' или 'Склад'!");
    return;
  }
  
  // 1. Берем средние цены со склада
  var avgTabak = Number(listSklad.getRange("B9").getValue()) || 0;
  var avgUgol = Number(listSklad.getRange("B10").getValue()) || 0;
  var avgMund = Number(listSklad.getRange("B11").getValue()) || 0;
  
  // 2. Читаем данные бухгалтерии и считаем общую выручку и кол-во кальянов
  var dannyeBuh = listBuh.getDataRange().getValues();
  var totalVyruchka = 0;
  var totalKalyans = 0;
  var periodText = "Весь период";
  
  for (var i = 1; i < dannyeBuh.length; i++) {
    var status = dannyeBuh[i][4];
    var vyruchkaZaPeriod = Number(dannyeBuh[i][3]) || 0;
    var kalyanyZaPeriod = (Number(dannyeBuh[i][1]) || 0) + (Number(dannyeBuh[i][2]) || 0);
    
    // Если период еще не списан, мы его не учитываем в прибыли
    if (status === "Списано") {
      totalVyruchka += vyruchkaZaPeriod;
      totalKalyans += kalyanyZaPeriod;
      if (dannyeBuh[i][0] !== "") periodText = dannyeBuh[i][0]; // Берем текст последнего периода
    }
  }
  
  if (totalVyruchka === 0) {
    return ui.alert("Внимание! В листе 'Бухгалтерия' нет списанных периодов с выручкой.");
  }
  
  // 3. Считаем РАСХОД СЫРЬЯ АВТОМАТИЧЕСКИ
  // Формула: (Кол-во * Норма) * Средняя цена
  var costTabak = (totalKalyans * 25) * avgTabak;
  var costUgol = (totalKalyans * 6) * avgUgol;
  var costMund = (totalKalyans * 4) * avgMund;
  var totalRaskhodSyrja = costTabak + costUgol + costMund;
  
  // 4. Считаем общую ЗП из листа "Зарплата"
  var totalZarplata = 0;
  if (listZarplata) {
    var dannyeZarplaty = listZarplata.getDataRange().getValues();
    for (var j = 1; j < dannyeZarplaty.length; j++) {
      totalZarplata += Number(dannyeZarplaty[j][5]) || 0; 
    }
  }
  
  // 5. Считаем Мост (30%) и Чистую Прибыль
  var most = totalVyruchka * 0.3;
  var pribyl = totalVyruchka - most - totalRaskhodSyrja - totalZarplata;
  
  // 6. Записываем отчет
  var listPribyl = ss.getSheetByName("Прибыль") || ss.insertSheet("Прибыль");
  listPribyl.clear();
  
  listPribyl.getRange(1, 1, 1, 2).setValues([["Период", "Сводный отчет"]]);
  listPribyl.getRange(2, 1).setValue("Выручка (всего)");
  listPribyl.getRange(2, 2).setValue(totalVyruchka);
  
  listPribyl.getRange(3, 1).setValue("Мост (30%)");
  listPribyl.getRange(3, 2).setValue(most);
  
  listPribyl.getRange(4, 1).setValue("Себестоимость сырья (Авто)");
  listPribyl.getRange(4, 2).setValue(totalRaskhodSyrja);
  
  // Расшифровка сырья
  listPribyl.getRange(5, 1).setValue("  в т.ч. Табак");
  listPribyl.getRange(5, 2).setValue(costTabak);
  listPribyl.getRange(6, 1).setValue("  в т.ч. Уголь");
  listPribyl.getRange(6, 2).setValue(costUgol);
  listPribyl.getRange(7, 1).setValue("  в т.ч. Мундштуки");
  listPribyl.getRange(7, 2).setValue(costMund);
  
  listPribyl.getRange(8, 1).setValue("Зарплата (итого)");
  listPribyl.getRange(8, 2).setValue(totalZarplata);
  
  listPribyl.getRange(10, 1).setValue("ЧИСТАЯ ПРИБЫЛЬ");
  listPribyl.getRange(10, 2).setValue(pribyl);
  
  // Красим итоговую строку
  listPribyl.getRange(10, 1, 1, 2).setBackground("#d9ead3").setFontWeight("bold");
  
  // Форматируем числа
  listPribyl.getRange("B2:B10").setNumberFormat("#,##0.00");
  
  ui.alert("Прибыль рассчитана!\n\nВыручка: " + totalVyruchka + "\nСебестоимость: " + totalRaskhodSyrja.toFixed(2) + "\nЧИСТАЯ ПРИБЫЛЬ: " + pribyl.toFixed(2));
}

// ==========================================
// 11. СОЗДАНИЕ ЛИСТА СКЛАД
// ==========================================
function sozdatListSklad() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var listSklad = ss.getSheetByName("Склад") || ss.insertSheet("Склад");
  
  listSklad.clear();
  listSklad.clearFormats();
  
  listSklad.getRange("A1").setValue("ОБЩИЕ ИТОГИ").setFontWeight("bold").setFontSize(12);
  listSklad.getRange("A2:C2").setValues([["Показатель", "Кол-во", "Сумма (руб)"]]).setFontWeight("bold").setBackground("#d9ead3");
  listSklad.getRange("A3:C5").setValues([
    ["Всего табака", 0, 0],
    ["Всего угля", 0, 0],
    ["Всего мундштуков", 0, 0]
  ]);
  listSklad.getRange("A6:C6").setValues([["ИТОГО СКЛАД", "", 0]]).setFontWeight("bold").setBackground("#fce5cd");
  
  listSklad.getRange("A8").setValue("Средние цены:").setFontWeight("bold");
  listSklad.getRange("A9:B11").setValues([
    ["Ср. цена табака (за 1 гр)", 0],
    ["Цена угля (за 1 шт)", 0],
    ["Цена мундштука (за 1 шт)", 0]
  ]);
  
  listSklad.getRange("D1").setValue("ДЕТАЛИЗАЦИЯ ТАБАКА").setFontWeight("bold").setFontSize(12);
  listSklad.getRange("D2:F2").setValues([["Бренд табака", "Остаток (гр)", "Сумма в деньгах"]]).setFontWeight("bold").setBackground("#d9ead3");
  
  listSklad.getRange("H1").setValue("КОРРЕКТИРОВКА (ИНВЕНТАРИЗАЦИЯ)").setFontWeight("bold").setFontSize(12);
  listSklad.getRange("H2:I2").setValues([["Показатель", "Значение"]]).setFontWeight("bold").setBackground("#f4cccc");
  listSklad.getRange("H3:I5").setValues([
    ["Табак (гр)", 0],
    ["Уголь (шт)", 0],
    ["Мундштуки (шт)", 0]
  ]);
  
  listSklad.setColumnWidth(1, 220);
  listSklad.setColumnWidth(2, 100);
  listSklad.setColumnWidth(3, 120);
  listSklad.setColumnWidth(4, 150);
  listSklad.setColumnWidth(5, 120);
  listSklad.setColumnWidth(6, 120);
  listSklad.setColumnWidth(8, 180);
  listSklad.setColumnWidth(9, 100);
}

// ==========================================
// 12. ПЕРЕСЧЕТ СКЛАДА (ПРИХОД - РАСХОД + КОРРЕКТИРОВКА)
// ==========================================
function pereschetSklada() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var listPrihod = ss.getSheetByName("Приход Продукции");
  var listBuh = ss.getSheetByName("Бухгалтерия");
  var listSklad = ss.getSheetByName("Склад");
  
  if (!listPrihod || !listSklad) return SpreadsheetApp.getUi().alert("Ошибка: Нет листа 'Приход Продукции' или 'Склад'!");
  
  var korrTabak = Number(listSklad.getRange("I3").getValue()) || 0;
  var korrUgol = Number(listSklad.getRange("I4").getValue()) || 0;
  var korrMund = Number(listSklad.getRange("I5").getValue()) || 0;

  var dannyePrihoda = listPrihod.getDataRange().getValues();
  var buyTabakGr = 0, buyTabakMoney = 0, buyUgolSht = 0, buyUgolMoney = 0, buyMundSht = 0, buyMundMoney = 0;
  var brendyBuy = {}; var oshibki = []; 
  
  for (var i = 1; i < dannyePrihoda.length; i++) {
    var nazvanie = String(dannyePrihoda[i][2]); 
    var kolichestvoPachek = Number(dannyePrihoda[i][3]) || 0; 
    var summaDeneg = Number(dannyePrihoda[i][5]) || 0; 
    
    if (nazvanie === "" || kolichestvoPachek === 0) continue; 
    var nizhniyRegistr = nazvanie.toLowerCase();
    
    if (nizhniyRegistr.includes("уголь") || nizhniyRegistr.includes("crown")) {
      var match = nazvanie.match(/(\d+)\s*шт/i);
      if (!match) { oshibki.push(nazvanie); continue; }
      buyUgolSht += (kolichestvoPachek * Number(match[1]));
      buyUgolMoney += summaDeneg;
    } 
    else if (nizhniyRegistr.includes("мундштук")) {
      var match = nazvanie.match(/(\d+)\s*шт/i);
      if (!match) { oshibki.push(nazvanie); continue; }
      buyMundSht += (kolichestvoPachek * Number(match[1]));
      buyMundMoney += summaDeneg;
    } 
    else {
      var match = nazvanie.match(/(\d+)\s*г/i);
      if (!match) { oshibki.push(nazvanie); continue; }
      var vsegoGramm = kolichestvoPachek * Number(match[1]);
      buyTabakGr += vsegoGramm;
      buyTabakMoney += summaDeneg;
      
      var brend = nazvanie.split("-")[0].trim();
      if (brendyBuy[brend] == undefined) brendyBuy[brend] = 0;
      brendyBuy[brend] += vsegoGramm;
    }
  }
  
  var soldKalyans = 0;
  if (listBuh) {
    var dannyeBuh = listBuh.getDataRange().getValues();
    for (var j = 1; j < dannyeBuh.length; j++) {
      if (dannyeBuh[j][4] === "Списано") {
        soldKalyans += (Number(dannyeBuh[j][1]) || 0) + (Number(dannyeBuh[j][2]) || 0);
      }
    }
  }
  
  var remTabakGr = buyTabakGr - (soldKalyans * 25) + korrTabak;
  var remUgolSht = buyUgolSht - (soldKalyans * 6) + korrUgol;
  var remMundSht = buyMundSht - (soldKalyans * 4) + korrMund;
  
  var realBuyTabakGr = buyTabakGr > 0 ? buyTabakGr : 1;
  var avgTabak = buyTabakMoney / realBuyTabakGr;
  var avgUgol = buyUgolSht > 0 ? buyUgolMoney / buyUgolSht : 0;
  var avgMund = buyMundSht > 0 ? buyMundMoney / buyMundSht : 0;
  
  var remTabakMoney = remTabakGr * avgTabak;
  var remUgolMoney = remUgolSht * avgUgol;
  var remMundMoney = remMundSht * avgMund;
  var totalRemMoney = remTabakMoney + remUgolMoney + remMundMoney;
  
  listSklad.getRange("B3:C5").setValues([
    [remTabakGr, remTabakMoney],
    [remUgolSht, remUgolMoney],
    [remMundSht, remMundMoney]
  ]);
  listSklad.getRange("C6").setValue(totalRemMoney);
  listSklad.getRange("B9:B11").setValues([[avgTabak.toFixed(2)], [avgUgol.toFixed(2)], [avgMund.toFixed(2)]]);
  
  var lastRowSklad = listSklad.getLastRow();
  if (lastRowSklad >= 3) listSklad.getRange(3, 4, lastRowSklad - 2, 3).clearContent();
  
  var stroka = 3;
  for (var brendImya in brendyBuy) {
    var dolya = buyTabakGr > 0 ? (brendyBuy[brendImya] / buyTabakGr) : 0;
    var remBrendGr = remTabakGr * dolya;
    var remBrendMoney = remBrendGr * avgTabak;
    
    listSklad.getRange(stroka, 4).setValue(brendImya);               
    listSklad.getRange(stroka, 5).setValue(remBrendGr);    
    listSklad.getRange(stroka, 6).setValue(remBrendMoney); 
    stroka++;
  }
  
  if (oshibki.length > 0) {
    SpreadsheetApp.getUi().alert("Внимание!", "Склад пересчитан, но есть проблемы!\n\nНе указан вес:\n- " + oshibki.join("\n- "), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ==========================================
// 13. СПИСАНИЕ РАСХОДА (СТАВИМ СТАТУС И ПЕРЕСЧИТЫВАЕМ)
// ==========================================
function spisanieRaskhoda() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var listBuh = ss.getSheetByName("Бухгалтерия");
  if (!listBuh) return SpreadsheetApp.getUi().alert("Ошибка: Нет листа 'Бухгалтерия'!");
  
  var dannyeBuh = listBuh.getDataRange().getValues();
  var byloSpisano = false; 
  
  for (var i = 1; i < dannyeBuh.length; i++) {
    var status = dannyeBuh[i][4]; 
    if (status === "Списано") continue;
    
    var totalKalyans = (Number(dannyeBuh[i][1]) || 0) + (Number(dannyeBuh[i][2]) || 0);
    if (totalKalyans === 0) continue;
    
    listBuh.getRange(i + 1, 5).setValue("Списано"); 
    byloSpisano = true;
  }
  
  if (byloSpisano) {
    pereschetSklada(); 
    SpreadsheetApp.getUi().alert("Списание выполнено! Склад пересчитан с учетом новых продаж.");
  } else {
    SpreadsheetApp.getUi().alert("Нет новых данных для списания.");
  }
}

// ==========================================
// 14. ИНВЕНТАРИЗАЦИЯ (ВВОД РЕАЛЬНЫХ ОСТАТКОВ)
// ==========================================
function provestiInventarizatsiyu() {
  var ui = SpreadsheetApp.getUi();
  var listSklad = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Склад");
  
  var otvet1 = ui.prompt('Инвентаризация', 'Введите РЕАЛЬНЫЙ остаток ТАБАКА (в граммах):', ui.ButtonSet.OK_CANCEL);
  if (otvet1.getSelectedButton() != ui.Button.OK) return;
  var realTabak = Number(otvet1.getResponseText().trim()) || 0;
  
  var otvet2 = ui.prompt('Инвентаризация', 'Введите РЕАЛЬНЫЙ остаток УГЛЯ (в штуках):', ui.ButtonSet.OK_CANCEL);
  if (otvet2.getSelectedButton() != ui.Button.OK) return;
  var realUgol = Number(otvet2.getResponseText().trim()) || 0;
  
  var otvet3 = ui.prompt('Инвентаризация', 'Введите РЕАЛЬНЫЙ остаток МУНДШТУКОВ (в штуках):', ui.ButtonSet.OK_CANCEL);
  if (otvet3.getSelectedButton() != ui.Button.OK) return;
  var realMund = Number(otvet3.getResponseText().trim()) || 0;
  
  var scriptTabak = Number(listSklad.getRange("B3").getValue()) || 0;
  var scriptUgol = Number(listSklad.getRange("B4").getValue()) || 0;
  var scriptMund = Number(listSklad.getRange("B5").getValue()) || 0;
  
  var korrTabak = realTabak - scriptTabak;
  var korrUgol = realUgol - scriptUgol;
  var korrMund = realMund - scriptMund;
  
  listSklad.getRange("I3").setValue(korrTabak);
  listSklad.getRange("I4").setValue(korrUgol);
  listSklad.getRange("I5").setValue(korrMund);
  
  pereschetSklada();
  
  ui.alert("Инвентаризация проведена!\n\nРазница записана в блок корректировок.\nСклад пересчитан.");
}

// ==========================================
// 15. ГЛАВНЫЙ ПУЛЬТ (main)
// ==========================================
function main() {
  var ui = SpreadsheetApp.getUi();
  var otvet = ui.alert('Запуск всех отчетов', 'Обновить Сводку, рассчитать ЗП и Прибыль?', ui.ButtonSet.YES_NO);
  if (otvet != ui.Button.YES) return;
  
  svodkaPoPostavshikam(); 
  raschetZarplaty(); 
  raschetPribyli(); 
  
  ui.alert("Готово!", "Все отчеты (Сводка, ЗП, Прибыль) обновлены.", ui.ButtonSet.OK);
}