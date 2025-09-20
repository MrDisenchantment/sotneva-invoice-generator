const express = require('express');
const cors = require('cors');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs').promises;
const os = require('os');
const libre = require('libreoffice-convert');
const { start } = require('repl');
require('dotenv').config();

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, '../frontend')));

// Эндпоинт для получения токена DaData
app.get('/api/dadata-token', (req, res) => {
    res.json({ token: process.env.DADATA_TOKEN });
});

app.post('/generate-invoice', async (req, res) => {
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
            currentRow.getCell(columnMapping.quantity[0]).value = parseFloat(item.quantity);
            console.log(`✓ Количество: ${item.quantity}`);

            // Единица измерения
            currentRow.getCell(columnMapping.unit[0]).value = item.unit;
            console.log(`✓ Единица: ${item.unit}`);

            // Цена
            currentRow.getCell(columnMapping.price).value = parseFloat(item.price);
            console.log(`✓ Цена: ${item.price}`);

            // Сумма - используем готовое значение
            currentRow.getCell(columnMapping.sum[0]).value = parseFloat(item.price * item.quantity);
            console.log(`✓ Сумма: ${item.total}`);

            // Устанавливаем высоту строки как в шаблоне
            currentRow.height = templateRow.height;
            console.log(`✓ Высота строки установлена: ${templateRow.height}`);

            console.log(`✅ Товар ${index + 1} полностью обработан`);
        });

        console.log(`\n=== Обработка товаров завершена ===`);
        // Обновляем итоговую сумму
        const lastItemRow = startRow + items.length - 1;
        const totalRow = worksheet.getRow(lastItemRow + 1);

        // Используем готовое значение для итоговой суммы
        if (req.body.totalSum) {
            totalRow.getCell(columnMapping.sum[0]).value = parseFloat(req.body.totalSum);
        } else {
            // Если итоговая сумма не предоставлена, считаем сумму всех товаров
            let total = 0;
            items.forEach(item => {
                total += parseFloat(item.price * item.quantity);
            });
            totalRow.getCell(columnMapping.sum[0]).value = total;
        }

        // Генерируем файл
        const buffer = await workbook.xlsx.writeBuffer();

        const tempFilePath = path.join(os.tmpdir(), `invoice-${Date.now()}.xlsx`);

        await fs.writeFile(tempFilePath, buffer);

        const fileContent = await fs.readFile(tempFilePath);

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
    console.log(`Сервер запущен на http://localhost:${PORT}`);
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


function getColumnLetter(col) {
    let letter = '';
    while (col > 0) {
        let rem = (col - 1) % 26;
        letter = String.fromCharCode(65 + rem) + letter;
        col = Math.floor((col - 1) / 26);
    }
    return letter;
}