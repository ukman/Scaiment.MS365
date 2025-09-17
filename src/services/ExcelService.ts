
export class ExcelService {
    constructor (private context : Excel.RequestContext) {

    }

    public static async create(ctx : Excel.RequestContext) : Promise<ExcelService> {
        const service = new ExcelService(ctx);
        return service;
    }

    /**
     * 
     * @param sheetName 
     * @returns array with filled values from sheet <sheetName>
     */
    public async getNamedRangesWithValues(sheetName : string): Promise<{ [key: string]: any }> {
        try {
                const worksheet = this.context.workbook.worksheets.getItem(sheetName);
               
                worksheet.load("name");
                
                const workbookNames = this.context.workbook.names;
                const worksheetNames = worksheet.names;
                
                // Загружаем базовую информацию
                workbookNames.load("items");
                worksheetNames.load("items");
                
                await this.context.sync();
                
                const result: { [key: string]: any } = {};
                
                // Функция для безопасного получения диапазона
                const safeGetRange = async (namedItem: Excel.NamedItem, itemName: string) => {
                    try {
                        const range = namedItem.getRange();
                        range.load(["address", "values"]);
                        await this.context.sync();
                        
                        // Проверяем принадлежность к активному листу
                        const belongsToSheet = 
                            range.address.includes(sheetName) || 
                            range.address.includes(`'${sheetName}'`) ||
                            range.address.includes(`${sheetName}!`);
                        
                        if (belongsToSheet) {
                            let values = range.values;

                            for(let k = 0; k < values.length; k++) {
                                const row = values[k];
                                for(let m = 0; m < row.length; m++) {
                                    row[m] = this.convertExcelValue(row[m], namedItem.name);
                                }
                            }

                            while(Array.isArray(values) && values.length == 1) {
                                values = values[0];
                            }
                            result[itemName] = values;
                            console.log(`✓ Processed named range: ${itemName}`);
                        }
                    } catch (error) {
                        console.warn(`✗ Failed to process named range '${itemName}':`, error.message);
                    }
                };
                
                // Обрабатываем workbook names
                for (const namedItem of workbookNames.items) {
                    await safeGetRange(namedItem, namedItem.name);
                }
                
                // Обрабатываем worksheet names
                for (const namedItem of worksheetNames.items) {
                    await safeGetRange(namedItem, namedItem.name);
                    // await safeGetRange(namedItem, `${activeSheetName}.${namedItem.name}`);
                }
                
                console.log(`Total named ranges processed: ${Object.keys(result).length}`);
                return result;
        } catch (error) {
            console.error("Ошибка при получении именованных диапазонов:", error);
            throw error;
        }
    }      


/**
 * Извлекает данные из всех таблиц на указанном листе Excel
 * @param sheetName - имя листа Excel
 * @returns Promise с объектом, где ключи - имена таблиц, значения - массивы объектов с данными
 */
    public async getTablesDataFromSheet(sheetName: string): Promise<Record<string, any[]>> {
        try {
            // Получаем лист по имени
            const worksheet = this.context.workbook.worksheets.getItem(sheetName);
            
            // Получаем все таблицы на листе
            const tables = worksheet.tables;
            tables.load("items/name");
            
            await this.context.sync();
            
            const result: Record<string, any[]> = {};
            
            // Проходим по каждой таблице
            for (const table of tables.items) {
              const tableName = table.name;
              
              // Загружаем данные таблицы
              table.load("rows/values, columns/name");
              await this.context.sync();
              
              // Получаем заголовки колонок
              const columnNames = table.columns.items.map(col => col.name);
              
              // Конвертируем строки в массив объектов
              const tableData = table.rows.items.map(row => {
                const rowObject: Record<string, any> = {};
                row.values[0].forEach((value, index) => {
                  const columnName = columnNames[index];
                  // Конвертируем значения в подходящие типы
                  rowObject[columnName] = this.convertExcelValue(value, columnName);
                });
                return rowObject;
              });
              
              result[tableName] = tableData;
            }
            
            return result;
            
          } catch (error) {
            throw error;
          }  
    }
  
/**
 * Конвертирует значения Excel в подходящие JavaScript типы
 * @param value - значение из Excel ячейки
 * @param columnName - имя колонки для определения специальных типов
 * @returns конвертированное значение
 */
public convertExcelValue(value: any, columnName?: string): any {
    // Если значение null или undefined
    if (value === null || value === undefined) {
      return null;
    }
    
    // Проверяем, нужно ли конвертировать в дату
    const isDateField = columnName && columnName.toLowerCase().endsWith('date');
    
    // Если это строка
    if (typeof value === 'string') {
      // Проверяем, является ли строка числом
      const numValue = Number(value);
      if (!isNaN(numValue) && value.trim() !== '') {
        // Если это поле даты и число, конвертируем в дату
        if (isDateField) {
          return this.convertExcelSerialToDate(numValue);
        }
        return numValue;
      }
      
      // Проверяем булевы значения
      if (value.toLowerCase() === 'true') return true;
      if (value.toLowerCase() === 'false') return false;
      
      return value;
    }
    
    // Если это число
    if (typeof value === 'number') {
      // Если это поле даты, конвертируем в дату
      if (isDateField) {
        return this.convertExcelSerialToDate(value);
      }
      return value;
    }
    
    // Если это уже булево значение
    if (typeof value === 'boolean') {
      return value;
    }
    
    // По умолчанию возвращаем как есть
    return value;
  }
    
/**
 * Конвертирует серийный номер Excel в дату JavaScript
 * @param serialNumber - серийный номер даты Excel
 * @returns объект Date или null если число некорректно
 */
public convertExcelSerialToDate(serialNumber: number): Date | null {
    // Проверяем, что это разумное значение для даты Excel
    // Excel считает дни с 1 января 1900 года (с некоторыми особенностями)
    if (serialNumber < 1 || serialNumber > 2958465) { // примерно до 31.12.9999
      return null;
    }
    
    try {
      // Excel считает 1 января 1900 года как день 1, но есть баг с 1900 годом (не високосный)
      // JavaScript Date начинается с 1 января 1970 года
      const excelEpoch = new Date(1899, 11, 30); // 30 декабря 1899 (Excel day 0)
      const millisecondsPerDay = 24 * 60 * 60 * 1000;
      
      const jsDate = new Date(excelEpoch.getTime() + serialNumber * millisecondsPerDay);
      
      // Проверяем, что получилась валидная дата
      if (isNaN(jsDate.getTime())) {
        return null;
      }
      
      return jsDate;
    } catch (error) {
      return null;
    }
  }
  
  /**
   * Альтернативная версия функции с более подробной обработкой ошибок
   * @param sheetName - имя листа Excel
   * @returns Promise с объектом данных таблиц или с информацией об ошибке
   */
  public async getTablesDataFromSheetSafe(sheetName: string): Promise<{
    success: boolean;
    data?: Record<string, any[]>;
    error?: string;
    tablesCount?: number;
  }> {
    try {
      const data = await this.getTablesDataFromSheet(sheetName);
      return {
        success: true,
        data,
        tablesCount: Object.keys(data).length
      };
    } catch (error) {
      let errorMessage = 'Неизвестная ошибка';
      
      if (error instanceof Error) {
        errorMessage = error.message;
      } else if (typeof error === 'string') {
        errorMessage = error;
      }
      
      // Проверяем специфичные ошибки Excel
      if (errorMessage.includes('ItemNotFound')) {
        errorMessage = `Лист с именем "${sheetName}" не найден`;
      }
      
      return {
        success: false,
        error: errorMessage
      };
    }
  }

    // Заполняет именованные поля в sheetName данными из data.
    public async fillSheet(context: Excel.RequestContext, data: any, sheetName: string) {
        // Получаем указанный лист
        const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
        await context.sync();
  
        if (sheet.isNullObject) {
          throw new Error(`Лист с именем "${sheetName}" не найден.`);
        }
  
        // Получаем все именованные диапазоны в области действия листа
        const namedItems = sheet.names;
        namedItems.load("name, value");
        await context.sync();
  
        // Перебираем все именованные диапазоны
        for (const namedItem of namedItems.items) {
          const rangeName = namedItem.name;
  
          // Проверяем, есть ли поле с таким именем в JSON-объекте
          if (data.hasOwnProperty(rangeName)) {
            const range = namedItem.getRange();
            const value = data[rangeName];
  
            // Записываем значение в диапазон
            range.values = [[value]];
  
            // Если значение — это дата, устанавливаем соответствующий формат
            if (value instanceof Date) {
              range.numberFormat = [["dd.mm.yyyy hh:mm:ss"]];
            }
          }
          // Если поля нет в JSON, ничего не делаем (пропускаем)
        }
  
        // Синхронизируем изменения
        await context.sync();
  
        // console.log(`Именованные диапазоны на листе "${sheetName}" заполнены данными из JSON.`);
    }


/**
 * Ищет все листы, содержащие именованный диапазон "__requisitionDraftMarker".
 * Учитывает:
 *  - Имя на уровне книги (Workbook Names)
 *  - Имя на уровне листа (Worksheet Names)
 * Возвращает уникальный список имен листов.
 */
    public async findSheetsWithMarker(ctx : Excel.RequestContext, markerName : string): Promise<string[]> {
        // const MARKER_NAME = "__requisitionDraftMarker";
  
        const wb = ctx.workbook;
        const result = new Set<string>();
    
        // Загрузим коллекцию листов (имена нужны для возврата и для обхода)
        const worksheets = wb.worksheets;
        worksheets.load("items/name");
        // Проверим наличие имени на уровне книги
        const wbNamed = wb.names.getItemOrNullObject(markerName);
    
        await ctx.sync();
    
        // Если имя существует на уровне книги — получим его диапазон и лист
        /*
        if (!wbNamed.isNullObject) {
            // На всякий случай убедимся, что это именно Range
            wbNamed.load("type");
            await ctx.sync();
    
            if (wbNamed.type === Excel.NamedItemType.range) {
            const r = wbNamed.getRange();
            r.load("worksheet/name");
            await ctx.sync();
            result.add(r.worksheet.name);
            }
        }
            */
    
        // Теперь проверим имена на уровне каждого листа
        // Сначала создаём "прокси" объекты для всех листов
        const perSheetNamedItems = worksheets.items.map((ws) => {
            const named = ws.names.getItemOrNullObject(markerName);
            return { ws, named };
        });
    
        // Выполним sync, чтобы узнать какие из них существуют
        await ctx.sync();
    
        // Для существующих имен получим соответствующие диапазоны и листы
        const rangesToLoad: Excel.Range[] = [];
        for (const { named } of perSheetNamedItems) {
            if (named && !named.isNullObject) {
            // Убедимся, что это Range
            named.load("type");
            }
        }
        await ctx.sync();
    
        for (const { named } of perSheetNamedItems) {
            if (named && !named.isNullObject && named.type === Excel.NamedItemType.range) {
            rangesToLoad.push(named.getRange());
            }
        }
    
        if (rangesToLoad.length > 0) {
            rangesToLoad.forEach((r) => r.load("worksheet/name"));
            await ctx.sync();
            rangesToLoad.forEach((r) => result.add(r.worksheet.name));
        }
    
        return Array.from(result);
    }
  
  
  /*
  // Пример использования:
  async function example() {
    try {
      // Получаем данные из всех таблиц на листе "Data"
      const tablesData = await getTablesDataFromSheet("Data");
      
      console.log('Извлеченные данные:', tablesData);
      
      // Можно обращаться к конкретным таблицам
      if (tablesData.persons) {
        console.log('Данные таблицы persons:', tablesData.persons);
        // Пример: если есть поле creationDate, оно будет объектом Date
        tablesData.persons.forEach(person => {
          if (person.creationDate instanceof Date) {
            console.log(`Person ${person.id} created on: ${person.creationDate.toLocaleDateString()}`);
          }
        });
      }
      
      if (tablesData.companies) {
        console.log('Данные таблицы companies:', tablesData.companies);
        // Пример: если есть поле foundationDate, оно будет объектом Date
        tablesData.companies.forEach(company => {
          if (company.foundationDate instanceof Date) {
            console.log(`Company ${company.name} founded on: ${company.foundationDate.toLocaleDateString()}`);
          }
        });
      }
      
    } catch (error) {
      console.error('Ошибка при извлечении данных:', error);
    }
  }
  
  // Пример использования безопасной версии:
  /*
  async function exampleSafe() {
    const result = await getTablesDataFromSheetSafe("Data");
    
    if (result.success) {
      console.log(`Успешно извлечено ${result.tablesCount} таблиц:`, result.data);
    } else {
      console.error('Ошибка:', result.error);
    }
  } 
    */   
}