# R7ConsultLibs

**R7ConsultLibs** —  это мощный инструмент для работы с Excel-файлами в JavaScript. Он предоставляет разработчикам простой и интуитивно понятный API, основанный на библиотеке ExcelJS, для обработки .xlsx-документов в табличном редакторе Р7-Офис через макросы. 

## Возможности и примеры работы с файлами Excel

### Установка и выполнение тестового примера 
- Установите плагин `R7ConsultLibs(0.1.0).plugin` в Р7-Офис. Обратите внимание, что данный плагин является системным и не отображается в списке установленных плагинов Р7-Офис.
- Загрузите файл `Книга1.xlsx` в Р7-Офис
- Откройте панель макросов и выполните макрос `Тест загрузки из внешнего файла`
-Выберите файл `Книга2.xlsx` и нажмите кнопку `Открыть`
-Файл `Книга2.xlsx` считается и его данные отобразятся на активном листе
- Дополнительно отобразится окно с сообщением о тестировании считывания ячейки A1


### Как начать использовать разработчику макросов
После установки плагина `R7ConsultLibs(0.1.0).plugin` в Р7-Офис,перед использованием любых функций `ExternalXLSXApi`, вам нужно его инициализировать и выбрать библиотеку.

#### `init()`

Эта функция запускает API. Вы должны вызвать ее первой.
(!)В текущей версии 0.1.0 временно убрана возможность выбора библиотеки и используется пока только ExcelJS!
-   `library`: Укажите `'sheetjs'` или `'exceljs'`. Это говорит API, какую внутреннюю библиотеку использовать для работы с файлами.

Пример:
```javascript
ExternalXLSXApi.init();
//ExternalXLSXApi.init('sheetjs'); // Используем SheetJS
// или
// ExternalXLSXApi.init('exceljs'); // Используем ExcelJS
```

### Загрузка файлов Excel

Прежде чем читать или записывать данные, вам нужно загрузить файл Excel. API сохраняет загруженные файлы во внутренней памяти, чтобы вы могли быстро работать с ними.

#### `loadWorkbook(filePath, fileData)`

Загружает файл Excel из данных, которые у вас уже есть (например, прочитаны из `<input type="file">` или получены другим способом).

-   `filePath`: Уникальное имя или путь для этого файла, чтобы потом к нему обращаться.
-   `fileData`: Сами данные файла. Обычно это `ArrayBuffer` или строка.

Пример:
```javascript
// Предположим, fileArrayBuffer содержит данные вашего .xlsx файла
// ExternalXLSXApi.loadWorkbook('мой_файл.xlsx', fileArrayBuffer);
```

#### `loadWorkbookFromUrl(filePath, url)`

Асинхронно загружает файл Excel с указанного URL с помощью `$ajax`.

-   `filePath`: Уникальное имя или путь для этого файла, чтобы потом к нему обращаться.
-   `url`: Веб-адрес файла .xlsx.

Пример:
```javascript
// Асинхронная функция для загрузки и использования
async function loadFile() {
  try {
    await ExternalXLSXApi.loadWorkbookFromUrl('удаленный_файл.xlsx', 'http://example.com/data/spreadsheet.xlsx');
    console.log('Файл загружен!');
    // Теперь можно читать/записывать данные
    const sheetData = ExternalXLSXApi.readSheet('удаленный_файл.xlsx', 'Лист1');
    console.log(sheetData);
  } catch (error) {
    console.error('Ошибка загрузки:', error);
  }
}
// loadFile(); // Вызываем функцию
```

### Чтение данных из файла

После загрузки файла вы можете читать данные из листов и ячеек.

#### `readSheet(filePath, sheetIdentifier)`

Читает все данные с указанного листа.

-   `filePath`: Имя файла, которое вы использовали при загрузке.
-   `sheetIdentifier`: Имя листа (например, 'Sheet1') или его номер (начиная с 0 для SheetJS, с 1 для ExcelJS).

Возвращает: Массив массивов, представляющий все строки и ячейки листа.

Пример:
```javascript
// const всеДанныеЛиста = ExternalXLSXApi.readSheet('мой_файл.xlsx', 'Лист1');
// console.log(всеДанныеЛиста);
```

#### `readRange(filePath, sheetIdentifier, range)`

Читает данные из определенного диапазона ячеек на листе.

-   `filePath`: Имя файла.
-   `sheetIdentifier`: Имя или номер листа.
-   `range`: Строка, указывающая диапазон ячеек (например, `'A1:C5'`).

Возвращает: Массив массивов, содержащий данные только из указанного диапазона.

Пример:
```javascript
// const данныеДиапазона = ExternalXLSXApi.readRange('мой_файл.xlsx', 'Лист1', 'B2:D4');
// console.log(данныеДиапазона);
```

#### `readCell(filePath, sheetIdentifier, cellAddress)`

Читает значение отдельной ячейки.

-   `filePath`: Имя файла.
-   `sheetIdentifier`: Имя или номер листа.
-   `cellAddress`: Адрес ячейки (например, `'A1'`).

Возвращает: Значение ячейки.

Пример:
```javascript
// const значениеЯчейки = ExternalXLSXApi.readCell('мой_файл.xlsx', 'Лист1', 'A1');
// console.log(значениеЯчейки);
```

### Запись данных в файл

Вы можете изменить содержимое файла, записывая новые данные.

#### `writeData(filePath, sheetIdentifier, startCellAddress, dataArray)`

Записывает данные из массива в указанный лист, начиная с определенной ячейки. **Важно:** Этот метод только изменяет данные во внутренней памяти. Чтобы сохранить изменения в файл, нужно вызвать `saveWorkbook`.

-   `filePath`: Имя файла.
-   `sheetIdentifier`: Имя или номер листа.
-   `startCellAddress`: Адрес ячейки, с которой начнется запись (например, `'A1'`).
-   `dataArray`: Массив массивов с данными для записи.

Пример:
```javascript
// const новыеДанные = [['Hello', 'World'], [123, 456]];
// ExternalXLSXApi.writeData('мой_файл.xlsx', 'Лист1', 'E1', новыеДанные);
// console.log('Данные записаны во внутреннюю структуру.');
```

### Сохранение изменений

После внесения изменений с помощью `writeData`, вам нужно сохранить их.

#### `saveWorkbook(filePath)`

Преобразует измененную рабочую книгу из внутренней памяти обратно в формат файла (ArrayBuffer). **Важно:** Эта функция только возвращает данные файла. Вам нужно будет использовать другие средства (например, API браузера для скачивания или функции Node.js для записи на диск) для фактического сохранения файла на компьютере пользователя или сервере.

-   `filePath`: Имя файла.

Возвращает: `ArrayBuffer` или `Promise<ArrayBuffer>` (для ExcelJS), содержащий данные измененного файла.

Пример:
```javascript
// async function saveFile() {
//   try {
//     const обновленныеДанныеФайла = await ExternalXLSXApi.saveWorkbook('мой_файл.xlsx');
//     // Здесь код для сохранения updatedFileData в файл
//     console.log('Данные файла получены для сохранения.');
//   } catch (error) {
//     console.error('Ошибка сохранения:', error);
//   }
// }
// saveFile();
```

### Управление памятью

Если вы закончили работать с файлом и он больше не нужен во внутренней памяти, вы можете его выгрузить.

#### `unloadWorkbook(filePath)`

Удаляет рабочую книгу из внутренней памяти API.

-   `filePath`: Имя файла, которое нужно выгрузить.

Пример:
```javascript
// ExternalXLSXApi.unloadWorkbook('мой_файл.xlsx');
// console.log('Файл выгружен из памяти.');
```

## Примеры использования

### Тест загрузки данных из Excel файла
```javascript
//Пример загрузки данных из Excel файла
(function(){
    if(Common.R7ConsultLibs)
        console.info("We have ExternalXLSXApi library!");
        if(Common.R7ConsultLibs.ExternalXLSXApi){
            let extApi=Common.R7ConsultLibs.ExternalXLSXApi;
            //extApi.init('sheetjs');
            extApi.init();            
            var myfile = AscDesktopEditor.OpenFilenameDialog("Excel(*xlsx)",false, function(_file) {
                // Если файл выбрали
                var file = _file;
                if (Array.isArray(file))
                    file = file[0]; //Если выбрали несколько берем первый
                if (!file) 
                    return; // Если не выбран файл закрыть макрос            
                file = file.replace(/\\/g,"/"); //Замена бэкслешей на слэши
                let idFile='test_file';
                loadAndProcessFile(extApi,idFile,file);
            });
        }    
})();

// Загрузка рабочей книги из URL (асинхронно)
async function loadAndProcessFile(extApi,idFile,pathToFile) {
  try {
    await extApi.loadWorkbookFromUrl(idFile, pathToFile);

    // Теперь вы можете работать с загруженной книгой
    const sheetData = extApi.readSheet(idFile, 1);
     if(sheetData){
        let worksheet=Api.GetActiveSheet();                                        
        let rowIdx=1;
        
        for(let row of sheetData){ 
            if(row){
                for(let colIdx=1;colIdx<=row.length;colIdx++){
                    worksheet.GetCells(rowIdx,colIdx).SetValue(row[colIdx-1]);
                }
            }
            rowIdx++;
        }

        let cellA1=extApi.readCell(idFile, 1,'A1');
        messageWindow("Тестирования считывания ячейки A1",cellA1);
        
        Api.asc_calculate(Asc.c_oAscCalculateType.ActiveSheet);
        }
    
    extApi.unloadWorkbook(idFile); // Выгрузить из кеша после использования

  } catch (error) {
    console.error("Ошибка при загрузке или обработке файла:", error);
  }
}
//вспомогательная функция для вывода сообщений
function messageWindow(title,textMessage){
    Common.UI.alert(
            {
                title: title,
                msg: textMessage,
                width: 600,            
                closable: true            
            }); 
}
```

## Структура проекта

```text
R7ConsultLibs/
├── example/
│   ├── R7ConsultLibs(0.1.0).plugin  # Плагин для Р7-Офис
│   ├── Книга1.xlsx                  # Пример файла Excel
│   └── Книга2.xlsx                  # Файл для тестирования загрузки
├── docs/
│   └── Презентация решения R7ConsultLibs.pdf  # Презентация решения
└── README.md                        # Документация проекта
```

## Презентация решения
 <a href="./docs/Презентация решения R7ConsultLibs.pdf">Презентация</a>

## Контакты
- Сайт: [Р7-Консалт](https://r7-consult.ru/)
- Email: er@exceldb.pro
- Телефон: +7 915 258-0371
- Telegram: https://t.me/r7_js
