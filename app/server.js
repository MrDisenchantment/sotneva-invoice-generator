import express from 'express';
import cors from 'cors';
import ExcelJS from 'exceljs';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs/promises';
import os from 'os';
import libre from 'libreoffice-convert';
import { config } from 'dotenv';
import { spawn } from 'child_process';

config()

/**
 * Функция для автоматического обновления документа LibreOffice
 * Принудительно пересчитывает формулы, высоту строк, ширину столбцов и другие параметры форматирования
 * @param {string} filePath - Путь к файлу Excel
 * @returns {Promise<Buffer>} - Обновленный файл в виде буфера
 */
async function updateLibreOfficeDocument(filePath) {
    return new Promise((resolve, reject) => {
        // Создаем временный файл для обновленного документа
        const updatedFilePath = filePath.replace('.xlsx', '_updated.xlsx');
        
        // Команда LibreOffice для открытия, обновления и сохранения документа
        // --headless - запуск без GUI
        // --calc - использовать Calc для Excel файлов
        // --convert-to xlsx - конвертировать в xlsx (это принудительно пересчитает все)
        // --outdir - директория для сохранения
        const libreOfficeArgs = [
            '--headless',
            '--calc',
            '--convert-to', 'xlsx',
            '--outdir', path.dirname(updatedFilePath),
            filePath
        ];

        console.log('Обновляем документ LibreOffice...');
        
        const libreOfficeProcess = spawn('libreoffice', libreOfficeArgs);
        
        let stderr = '';
        
        libreOfficeProcess.stderr.on('data', (data) => {
            stderr += data.toString();
        });
        
        libreOfficeProcess.on('close', async (code) => {
            if (code !== 0) {
                console.warn(`LibreOffice завершился с кодом ${code}, stderr: ${stderr}`);
                // Если LibreOffice недоступен, возвращаем исходный файл
                try {
                    const originalBuffer = await fs.readFile(filePath);
                    resolve(originalBuffer);
                } catch (error) {
                    reject(new Error(`Не удалось прочитать исходный файл: ${error.message}`));
                }
                return;
            }
            
            try {
                // Читаем обновленный файл
                const updatedBuffer = await fs.readFile(updatedFilePath);
                
                // Удаляем временный файл
                await fs.unlink(updatedFilePath).catch(e => 
                    console.warn('Не удалось удалить временный файл:', e.message)
                );
                
                console.log('Документ LibreOffice успешно обновлен');
                resolve(updatedBuffer);
            } catch (error) {
                // Если не удалось прочитать обновленный файл, возвращаем исходный
                console.warn('Не удалось прочитать обновленный файл, используем исходный:', error.message);
                try {
                    const originalBuffer = await fs.readFile(filePath);
                    resolve(originalBuffer);
                } catch (readError) {
                    reject(new Error(`Не удалось прочитать исходный файл: ${readError.message}`));
                }
            }
        });
        
        libreOfficeProcess.on('error', async (error) => {
            console.warn('LibreOffice недоступен:', error.message);
            // Если LibreOffice недоступен, возвращаем исходный файл
            try {
                const originalBuffer = await fs.readFile(filePath);
                resolve(originalBuffer);
            } catch (readError) {
                reject(new Error(`LibreOffice недоступен и не удалось прочитать исходный файл: ${readError.message}`));
            }
        });
    });
}

const app = express();
app.use(cors());
app.use(express.json());

// frontend
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
app.use(express.static(path.join(__dirname, './www')));

app.get('/', function (req, res) {
    res.sendFile(__dirname + '/index.html');
});

/**
 * Проверка доступности LibreOffice в системе
 * @returns {Promise<boolean>} - true если LibreOffice доступен
 */
async function checkLibreOfficeAvailability() {
    return new Promise((resolve) => {
        const testProcess = spawn('libreoffice', ['--version']);
        
        testProcess.on('close', (code) => {
            resolve(code === 0);
        });
        
        testProcess.on('error', () => {
            resolve(false);
        });
        
        // Таймаут на случай зависания
        setTimeout(() => {
            testProcess.kill();
            resolve(false);
        }, 5000);
    });
}

// Эндпоинт для получения токена DaData
app.get('/api/dadata-token', (req, res) => {
    res.json({ token: process.env.DADATA_TOKEN });
});

// Эндпоинт для проверки статуса системы обновления документов
app.get('/api/system-status', async (req, res) => {
    try {
        const libreOfficeAvailable = await checkLibreOfficeAvailability();
        
        res.json({
            status: 'ok',
            features: {
                exceljs_formatting: true,
                libreoffice_available: libreOfficeAvailable,
                document_update: true,
                auto_formatting: libreOfficeAvailable ? 'full' : 'basic'
            },
            message: libreOfficeAvailable 
                ? 'Полная поддержка обновления документов через LibreOffice'
                : 'Базовая поддержка обновления через ExcelJS (LibreOffice недоступен)'
        });
    } catch (error) {
        res.status(500).json({
            status: 'error',
            message: 'Ошибка при проверке статуса системы',
            error: error.message
        });
    }
});

app.post('/api/generate-invoice', async (req, res) => {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(path.join(__dirname, 'templates/schet-shablon.xlsx'));
        const worksheet = workbook.getWorksheet(1);

        // Заменяем маркеры в шаблоне
        worksheet.eachRow((row) => {
            row.eachCell((cell) => {
                if (cell.value && typeof cell.value === 'string') {
                    Object.entries(req.body.replacements).forEach(([marker, value]) => {
                        cell.value = cell.value.replace(marker, value);
                    });
                }
            });
        });

        // Добавляем строки товаров
        const items = req.body.items;
        const startRow = 16; // Начинаем с 16-й строки

        // Сохраняем шаблонную строку
        const templateRow = worksheet.getRow(startRow);
        const templateStyle = {};

        // Сохраняем стили шаблонной строки
        templateRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            templateStyle[colNumber] = {
                style: cell.style,
                alignment: cell.alignment,
                border: cell.border,
                fill: cell.fill,
                font: cell.font,
                numFmt: cell.numFmt
            };
        });

        // Соответствие столбцов согласно требованиям
        const columnMapping = {
            number: 1,        // Номер (A)
            name: [2, 10],   // Наименование товара (B-J)
            quantity: [11, 12], // Количество (K-L)
            unit: [13, 14],  // Ед. измерения (M-N)
            price: 15,       // Цена (O)
            sum: [16, 17]    // Сумма (P-Q)
        };

        // Добавляем дополнительные строки, если нужно
        if (items.length > 1) {
            for (let i = 1; i < items.length; i++) {
                const newRowIndex = startRow + i;
                console.log(`Создаем новую строку ${newRowIndex}`);

                // Вставляем пустую строку
                worksheet.spliceRows(newRowIndex, 0, []);

                // Получаем новую строку и шаблонную строку
                const newRow = worksheet.getRow(newRowIndex);
                const templateRow = worksheet.getRow(startRow);

                // Копируем только высоту строки
                newRow.height = templateRow.height;
            }
        }

        // Добавляем строки для каждого товара
        console.log(`Начинаем обработку ${items.length} товаров`);
        items.forEach((item, index) => {
            const currentRowNumber = startRow + index;
            console.log(`\n--- Обработка товара ${index + 1} в строке ${currentRowNumber} ---`);
            console.log(`Товар: ${item.name}, Количество: ${item.quantity}, Цена: ${item.price}`);

            const currentRow = worksheet.getRow(currentRowNumber);

            // Копируем стили из шаблона для всех ячеек
            console.log(`Копируем стили для строки ${currentRowNumber}`);
            Object.entries(templateStyle).forEach(([col, style]) => {
                const cell = currentRow.getCell(parseInt(col));
                Object.assign(cell, style);
            });
            console.log(`✓ Стили скопированы`);

            // Теперь объединяем ячейки (разъединение уже выполнено выше)
            console.log(`Начинаем объединение/разъединение ячеек в строке ${currentRowNumber}`);
            try {

                // KOSTIL

                for (let key in worksheet._merges) {
                    if (key.includes(`${currentRowNumber}`)) {
                        let left = getColumnLetter(worksheet._merges[key].left)
                        let right = getColumnLetter(worksheet._merges[key].right)
                        worksheet.unMergeCells(`${left}${currentRowNumber}:${right}${currentRowNumber}`)
                    }
                }


                // Наименование товара (B-J)
                mergeCellsIfNeeded(worksheet, `B${currentRowNumber}:J${currentRowNumber}`)
                // worksheet.unMergeCells(`B${currentRowNumber}:J${currentRowNumber}`);
                // worksheet.mergeCells(`B${currentRowNumber}:J${currentRowNumber}`);

                // Количество (K-L)
                // worksheet.unMergeCells(`K${currentRowNumber}:L${currentRowNumber}`);
                // worksheet.mergeCells(`K${currentRowNumber}:L${currentRowNumber}`);
                mergeCellsIfNeeded(worksheet, `K${currentRowNumber}:L${currentRowNumber}`)

                // Единица измерения (M-N)
                // worksheet.unMergeCells(`M${currentRowNumber}:N${currentRowNumber}`);
                // worksheet.mergeCells(`M${currentRowNumber}:N${currentRowNumber}`);
                mergeCellsIfNeeded(worksheet, `M${currentRowNumber}:N${currentRowNumber}`)


                // Сумма (P-Q)
                // worksheet.unMergeCells(`P${currentRowNumber}:Q${currentRowNumber}`);
                // worksheet.mergeCells(`P${currentRowNumber}:Q${currentRowNumber}`);
                mergeCellsIfNeeded(worksheet, `P${currentRowNumber}:Q${currentRowNumber}`)

                console.log(`✓ Все ячейки в строке ${currentRowNumber} успешно объединены`);
            } catch (error) {
                console.log(`❌ ОШИБКА объединения ячеек в строке ${currentRowNumber}: ${error.message}`);
                console.log(`Детали ошибки:`, error);
            }

            // Заполняем данные в соответствующие столбцы
            console.log(`Заполняем данные в строке ${currentRowNumber}`);
            // Номер по порядку
            currentRow.getCell(columnMapping.number).value = index + 1;
            console.log(`✓ Номер: ${index + 1}`);

            // Наименование товара
            currentRow.getCell(columnMapping.name[0]).value = item.name;
            console.log(`✓ Наименование: ${item.name}`);

            // Количество
            currentRow.getCell(columnMapping.quantity[0]).value = item.quantity;
            console.log(`✓ Количество: ${item.quantity}`);

            // Единица измерения
            currentRow.getCell(columnMapping.unit[0]).value = item.unit;
            console.log(`✓ Единица: ${item.unit}`);

            // Цена
            currentRow.getCell(columnMapping.price).value = item.price;
            console.log(`✓ Цена: ${item.price}`);

            // Сумма - используем готовое значение
            currentRow.getCell(columnMapping.sum[0]).value = item.sum;
            console.log(`✓ Сумма: ${item.sum}`);

            // Устанавливаем высоту строки как в шаблоне
            currentRow.height = templateRow.height;
            console.log(`✓ Высота строки установлена: ${templateRow.height}`);

            console.log(`✅ Товар ${index + 1} полностью обработан`);
        });

        console.log(`\n=== Обработка товаров завершена ===`);
        console.log('\n=== НАЧИНАЕМ ПРОЦЕСС ОБНОВЛЕНИЯ ДОКУМЕНТА ===');
        console.log('Этап 1: Обновление итоговой суммы...');
        
        // Обновляем итоговую сумму
        const lastItemRow = startRow + items.length - 1;
        const totalRow = worksheet.getRow(lastItemRow + 1);

        // Используем готовое значение для итоговой суммы
        if (req.body.replacements["{{total_sum}}"]) {
            totalRow.getCell(columnMapping.sum[0]).value = req.body.replacements["{{total_sum}}"];
            console.log(`✓ Итоговая сумма установлена из replacements: ${req.body.replacements["{{total_sum}}"]}`);            
        } else {
            // Если итоговая сумма не предоставлена, считаем сумму всех товаров
            let total = 0;
            items.forEach(item => {
                total += parseFloat(item.sum);
            });
            totalRow.getCell(columnMapping.sum[0]).value = total;
            console.log(`✓ Итоговая сумма рассчитана автоматически: ${total}`);
        }
        
        console.log('Этап 1 завершен: Итоговая сумма обновлена');
        console.log('Этап 2: Генерация Excel файла...');

        // Генерируем файл
        const buffer = await workbook.xlsx.writeBuffer();
        console.log('✓ Excel файл сгенерирован в буфер');

        const tempFilePath = path.join(os.tmpdir(), `invoice-${Date.now()}.xlsx`);
        console.log(`Сохраняем временный файл: ${tempFilePath}`);

        await fs.writeFile(tempFilePath, buffer);
        console.log('Этап 2 завершен: Файл Excel сохранен');

        const fileContent = await fs.readFile(tempFilePath);
        console.log('Этап 3 завершен: Файл прочитан для конвертации');
        
        console.log('=== ПРОЦЕСС ПОДГОТОВКИ ФАЙЛА ЗАВЕРШЕН ===\n');}

        // Затем обновляем документ через LibreOffice для финального пересчета
        console.log('Запускаем обновление документа через LibreOffice...');
        const updatedBuffer = await updateLibreOfficeDocument(tempFilePath);
        
        // Перезаписываем файл обновленной версией
        await fs.writeFile(tempFilePath, updatedBuffer);
        
        const fileContent = await fs.readFile(tempFilePath);


        // Затем обновляем документ через LibreOffice для финального пересчета
        console.log('Запускаем расширенное обновление документа через LibreOffice...');
        const updatedBuffer = await updateLibreOfficeWithMacro(tempFilePath);

        const pdfBuf = await new Promise((resolve, reject) => {
            libre.convert(fileContent, '.pdf', undefined, (err, done) => {
                if (err) {
                    fs.unlink(tempFilePath).catch(e => console.error("Couldn't remove temp file", e));
                    return reject(err);
                }
                fs.unlink(tempFilePath).catch(e => console.error("Couldn't remove temp file", e));
                resolve(done);
            });
        });

        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition', 'attachment; filename=invoice.pdf');
        res.send(pdfBuf);
    } catch (error) {
        console.error(error);
        res.status(500).send('Ошибка при генерации счета');
    }
});

const PORT = 3000;
app.listen(PORT, () => {
    console.log(`server started on post: ${PORT}`);
});

const mergeCellsIfNeeded = (
    worksheet,
    range
) => {
    const [startCell, endCell] = range.split(":");
    const startRow = worksheet.getCell(startCell).row;
    const endRow = worksheet.getCell(endCell).row;
    const startCol = worksheet.getCell(startCell).col;
    const endCol = worksheet.getCell(endCell).col;

    let alreadyMerged = false;

    for (let row = startRow; row <= endRow; row++) {
        for (let col = startCol; col <= endCol; col++) {
            const cell = worksheet.getCell(row, col);
            if (cell.isMerged) {
                alreadyMerged = true;
                break;
            }
        }
        if (alreadyMerged) break;
    }

    if (!alreadyMerged) {
        worksheet.mergeCells(range);
    }
};


/**
 * Функция для обновления форматирования через ExcelJS
 * Принудительно пересчитывает размеры ячеек и обновляет форматирование
 * @param {ExcelJS.Workbook} workbook - Рабочая книга ExcelJS
 */
function updateWorkbookFormatting(workbook) {
    console.log('Обновляем форматирование через ExcelJS...');
    
    workbook.eachSheet((worksheet) => {
        // Принудительно пересчитываем высоту строк на основе содержимого
        worksheet.eachRow((row, rowNumber) => {
            let maxHeight = 15; // Минимальная высота строки
            
            row.eachCell({ includeEmpty: false }, (cell) => {
                if (cell.value && typeof cell.value === 'string') {
                    // Приблизительный расчет высоты на основе длины текста и ширины ячейки
                    const textLength = cell.value.length;
                    const estimatedLines = Math.ceil(textLength / 50); // Примерно 50 символов на строку
                    const estimatedHeight = estimatedLines * 15; // 15 пунктов на строку
                    maxHeight = Math.max(maxHeight, estimatedHeight);
                }
            });
            
            // Устанавливаем высоту строки, но не больше разумного максимума
            row.height = Math.min(maxHeight, 100);
        });
        
        // Автоматически подгоняем ширину столбцов
        worksheet.columns.forEach((column, index) => {
            if (column.values) {
                let maxWidth = 10; // Минимальная ширина
                
                column.values.forEach(value => {
                    if (value && typeof value === 'string') {
                        maxWidth = Math.max(maxWidth, value.length * 1.2);
                    } else if (value && typeof value === 'number') {
                        maxWidth = Math.max(maxWidth, value.toString().length * 1.2);
                    }
                });
                
                // Устанавливаем ширину столбца, но не больше разумного максимума
                column.width = Math.min(maxWidth, 50);
            }
        });
        
        // Принудительно обновляем все формулы
        worksheet.eachRow((row) => {
            row.eachCell((cell) => {
                if (cell.formula) {
                    // Принудительно пересчитываем формулу
                    const formula = cell.formula;
                    cell.formula = formula;
                }
            });
        });
    });
    
    console.log('Этап 1 завершен: Форматирование ExcelJS обновлено');
    
    console.log('Этап 2: Генерация файла Excel...');
    // Генерируем файл
    const buffer = await workbook.xlsx.writeBuffer();

    const tempFilePath = path.join(os.tmpdir(), `invoice-${Date.now()}.xlsx`);
    console.log(`Сохраняем временный файл: ${tempFilePath}`);

    await fs.writeFile(tempFilePath, buffer);
    console.log('Этап 2 завершен: Файл Excel сохранен');

    const fileContent = await fs.readFile(tempFilePath);
    console.log('Этап 3 завершен: Файл прочитан для конвертации');
    
    console.log('=== ПРОЦЕСС ПОДГОТОВКИ ФАЙЛА ЗАВЕРШЕН ===\n');}

    // Затем обновляем документ через LibreOffice для финального пересчета
    console.log('Запускаем расширенное обновление документа через LibreOffice...');
    const updatedBuffer = await updateLibreOfficeWithMacro(tempFilePath);

    const pdfBuf = await new Promise((resolve, reject) => {
        libre.convert(fileContent, '.pdf', undefined, (err, done) => {
            if (err) {
                fs.unlink(tempFilePath).catch(e => console.error("Couldn't remove temp file", e));
                return reject(err);
            }
            fs.unlink(tempFilePath).catch(e => console.error("Couldn't remove temp file", e));
            resolve(done);
        });
    });

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename=invoice.pdf');
    res.send(pdfBuf);
}

/**
 * Расширенная функция обновления LibreOffice с макросами
 * Использует встроенные команды LibreOffice для точного пересчета
 * @param {string} filePath - Путь к файлу Excel
 * @returns {Promise<Buffer>} - Обновленный файл в виде буфера
 */
async function updateLibreOfficeWithMacro(filePath) {
    return new Promise((resolve, reject) => {
        const updatedFilePath = filePath.replace('.xlsx', '_macro_updated.xlsx');
        
        // Создаем временный макрос для обновления документа
        const macroCommands = [
            '--headless',
            '--calc',
            '--eval',
            // Макрос для обновления всех вычислений и форматирования
            'Sub UpdateDocument\n' +
            '    Dim oDoc As Object\n' +
            '    oDoc = ThisComponent\n' +
            '    \n' +
            '    \' Пересчитываем все формулы\n' +
            '    oDoc.calculateAll()\n' +
            '    \n' +
            '    \' Обновляем автоподбор высоты строк\n' +
            '    Dim oSheets As Object\n' +
            '    Dim oSheet As Object\n' +
            '    Dim i As Integer\n' +
            '    \n' +
            '    oSheets = oDoc.getSheets()\n' +
            '    For i = 0 To oSheets.getCount() - 1\n' +
            '        oSheet = oSheets.getByIndex(i)\n' +
            '        \' Автоподбор высоты всех строк\n' +
            '        oSheet.getRows().setPropertyValue("OptimalHeight", True)\n' +
            '        \' Автоподбор ширины всех столбцов\n' +
            '        oSheet.getColumns().setPropertyValue("OptimalWidth", True)\n' +
            '    Next i\n' +
            '    \n' +
            '    \' Сохраняем документ\n' +
            '    oDoc.store()\n' +
            'End Sub\n' +
            'UpdateDocument',
            filePath
        ];

        console.log('Обновляем документ LibreOffice с макросами...');
        
        const libreOfficeProcess = spawn('libreoffice', macroCommands);
        
        let stderr = '';
        
        libreOfficeProcess.stderr.on('data', (data) => {
            stderr += data.toString();
        });
        
        libreOfficeProcess.on('close', async (code) => {
            // Независимо от результата макроса, пробуем стандартное обновление
            try {
                const standardUpdated = await updateLibreOfficeDocument(filePath);
                resolve(standardUpdated);
            } catch (error) {
                reject(error);
            }
        });
        
        libreOfficeProcess.on('error', async (error) => {
            console.warn('Макрос LibreOffice недоступен, используем стандартное обновление:', error.message);
            try {
                const standardUpdated = await updateLibreOfficeDocument(filePath);
                resolve(standardUpdated);
            } catch (standardError) {
                reject(standardError);
            }
        });
    });
}

async function updateLibreOfficeDocument(filePath) {
    return new Promise((resolve, reject) => {
        // Создаем временный файл для обновленного документа
        const updatedFilePath = filePath.replace('.xlsx', '_updated.xlsx');
        
        // Команда LibreOffice для открытия, обновления и сохранения документа
        // --headless - запуск без GUI
        // --calc - использовать Calc для Excel файлов
        // --convert-to xlsx - конвертировать в xlsx (это принудительно пересчитает все)
        // --outdir - директория для сохранения
        const libreOfficeArgs = [
            '--headless',
            '--calc',
            '--convert-to', 'xlsx',
            '--outdir', path.dirname(updatedFilePath),
            filePath
        ];

        console.log('Обновляем документ LibreOffice...');
        
        const libreOfficeProcess = spawn('libreoffice', libreOfficeArgs);
        
        let stderr = '';
        
        libreOfficeProcess.stderr.on('data', (data) => {
            stderr += data.toString();
        });
        
        libreOfficeProcess.on('close', async (code) => {
            if (code !== 0) {
                console.warn(`LibreOffice завершился с кодом ${code}, stderr: ${stderr}`);
                // Если LibreOffice недоступен, возвращаем исходный файл
                try {
                    const originalBuffer = await fs.readFile(filePath);
                    resolve(originalBuffer);
                } catch (error) {
                    reject(new Error(`Не удалось прочитать исходный файл: ${error.message}`));
                }
                return;
            }
            
            try {
                // Читаем обновленный файл
                const updatedBuffer = await fs.readFile(updatedFilePath);
                
                // Удаляем временный файл
                await fs.unlink(updatedFilePath).catch(e => 
                    console.warn('Не удалось удалить временный файл:', e.message)
                );
                
                console.log('Документ LibreOffice успешно обновлен');
                resolve(updatedBuffer);
            } catch (error) {
                // Если не удалось прочитать обновленный файл, возвращаем исходный
                console.warn('Не удалось прочитать обновленный файл, используем исходный:', error.message);
                try {
                    const originalBuffer = await fs.readFile(filePath);
                    resolve(originalBuffer);
                } catch (readError) {
                    reject(new Error(`Не удалось прочитать исходный файл: ${readError.message}`));
                }
            }
        });
        
        libreOfficeProcess.on('error', async (error) => {
            console.warn('LibreOffice недоступен:', error.message);
            // Если LibreOffice недоступен, возвращаем исходный файл
            try {
                const originalBuffer = await fs.readFile(filePath);
                resolve(originalBuffer);
            } catch (readError) {
                reject(new Error(`LibreOffice недоступен и не удалось прочитать исходный файл: ${readError.message}`));
            }
        });
    });
}

function getColumnLetter(col) {
    let letter = '';
    while (col > 0) {
        let rem = (col - 1) % 26;
        letter = String.fromCharCode(65 + rem) + letter;
        col = Math.floor((col - 1) / 26);
    }
    return letter;
}