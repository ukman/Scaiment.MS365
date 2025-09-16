export async function excelLog2(message: string): Promise<void> {
    try {
        await Excel.run(async (context: Excel.RequestContext) => {
            // Получаем коллекцию листов
            const sheets = context.workbook.worksheets;
            
            // Загружаем имена листов
            sheets.load("items/name");
            await context.sync();
            
            // Проверяем, существует ли лист "__logs"
            const logSheet = sheets.items.find(sheet => sheet.name === "__logs");
            
            if (logSheet) {
                // Активируем лист логов
                logSheet.load("name");
                await context.sync();
                
                // Находим последнюю заполненную строку в столбце A
                const range = logSheet.getUsedRange();
                let nextRow = 1;
                
                if (range) {
                    range.load("rowCount");
                    await context.sync();
                    nextRow = range.rowCount + 1;
                }
                
                // Записываем сообщение и текущую дату/время
                const logRange = logSheet.getCell(nextRow - 0, 0);
                logRange.values = [[`${new Date().toISOString()}: ${message}`]];
                
                await context.sync();
            }
            // Если лист "__logs" не найден, ничего не делаем
        });
    } catch (error) {
        console.error("Error in log function:", error);
    }
}



let logBuffer: string[] = [];
let timeoutId: number | null = null;
const BUFFER_DELAY = 1000; // Задержка в 1 секунду

export async function  excelLog(message: string): Promise<void> {
    // Добавляем сообщение в буфер
    logBuffer.push(message);

    // Если таймер уже запущен, ждем его завершения
    if (timeoutId !== null) {
        return;
    }

    // Устанавливаем таймер для записи логов
    timeoutId = window.setTimeout(async () => {
        try {
            await Excel.run(async (context: Excel.RequestContext) => {
                // Получаем коллекцию листов
                const sheets = context.workbook.worksheets;
                sheets.load("items/name");
                await context.sync();

                // Проверяем, существует ли лист "__logs"
                const logSheet = sheets.items.find(sheet => sheet.name === "__logs");

                if (logSheet) {
                    // Получаем диапазон столбца A, чтобы найти последнюю заполненную строку
                    const columnA = logSheet.getRange("A:A");
                    const usedRange = logSheet.getUsedRange(true); // true для учета форматированных ячеек
                    let nextRow = 1;

                    if (usedRange) {
                        // Загружаем свойства диапазона
                        usedRange.load(["rowIndex", "rowCount"]);
                        await context.sync();

                        // Вычисляем следующую свободную строку
                        nextRow = usedRange.rowIndex + usedRange.rowCount + 1;
                    } else {
                        // Если лист пустой, начинаем с первой строки
                        nextRow = 1;
                    }

                    // Формируем значения для записи (временная метка + сообщение)
                    const values = logBuffer.map(msg => [`${new Date().toISOString()}: ${msg}`]);

                    // Записываем все сообщения из буфера в лист
                    const logRange = logSheet.getRange(`A${nextRow}:A${nextRow + values.length - 1}`);
                    logRange.values = values;

                    await context.sync();

                    // Диагностика: выводим информацию о записи
                    console.log(`Записано ${values.length} логов начиная со строки ${nextRow}`);
                }

                // Очищаем буфер и таймер
                logBuffer = [];
                timeoutId = null;
            });
        } catch (error) {
            console.error("Error in log function:", error);
            // Очищаем буфер и таймер в случае ошибки
            logBuffer = [];
            timeoutId = null;
        }
    }, BUFFER_DELAY);
}