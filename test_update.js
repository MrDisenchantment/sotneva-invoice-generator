const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

// Импортируем функции из server.js
const { 
    updateWorkbookFormatting, 
    updateLibreOfficeWithMacro, 
    updateLibreOfficeDocument,
    checkLibreOfficeAvailability 
} = require('./app/server.js');

async function testDocumentUpdate() {
    console.log('🧪 Тестирование функций обновления документов...\n');
    
    try {
        // 1. Проверяем доступность LibreOffice
        console.log('1️⃣ Проверка доступности LibreOffice...');
        const isLibreOfficeAvailable = await checkLibreOfficeAvailability();
        console.log(`   LibreOffice доступен: ${isLibreOfficeAvailable ? '✅ Да' : '❌ Нет'}\n`);
        
        // 2. Создаем тестовый Excel файл
        console.log('2️⃣ Создание тестового Excel файла...');
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Тест');
        
        // Добавляем тестовые данные
        worksheet.getCell('A1').value = 'Очень длинный текст для проверки автоматического изменения ширины столбца';
        worksheet.getCell('A2').value = 'Короткий';
        worksheet.getCell('B1').value = 'Формула:';
        worksheet.getCell('B2').value = { formula: 'SUM(1,2,3)' };
        
        const testFilePath = path.join(__dirname, 'test_document.xlsx');
        await workbook.xlsx.writeFile(testFilePath);
        console.log(`   Тестовый файл создан: ${testFilePath}\n`);
        
        // 3. Тестируем обновление через ExcelJS
        console.log('3️⃣ Тестирование обновления через ExcelJS...');
        const testWorkbook = new ExcelJS.Workbook();
        await testWorkbook.xlsx.readFile(testFilePath);
        await updateWorkbookFormatting(testWorkbook);
        
        const updatedFilePath = path.join(__dirname, 'test_document_updated.xlsx');
        await testWorkbook.xlsx.writeFile(updatedFilePath);
        console.log(`   Файл обновлен через ExcelJS: ${updatedFilePath}\n`);
        
        // 4. Тестируем обновление через LibreOffice (если доступен)
        if (isLibreOfficeAvailable) {
            console.log('4️⃣ Тестирование обновления через LibreOffice...');
            
            try {
                await updateLibreOfficeWithMacro(testFilePath);
                console.log('   ✅ Обновление через макрос LibreOffice выполнено успешно\n');
            } catch (error) {
                console.log('   ⚠️ Обновление через макрос не удалось, пробуем базовое обновление...');
                await updateLibreOfficeDocument(testFilePath);
                console.log('   ✅ Базовое обновление LibreOffice выполнено успешно\n');
            }
        } else {
            console.log('4️⃣ Пропускаем тестирование LibreOffice (не доступен)\n');
        }
        
        console.log('🎉 Тестирование завершено успешно!');
        console.log('\nРезультаты:');
        console.log(`- LibreOffice доступен: ${isLibreOfficeAvailable ? 'Да' : 'Нет'}`);
        console.log(`- Обновление через ExcelJS: Работает`);
        console.log(`- Обновление через LibreOffice: ${isLibreOfficeAvailable ? 'Работает' : 'Недоступно'}`);
        
    } catch (error) {
        console.error('❌ Ошибка при тестировании:', error.message);
    } finally {
        // Очищаем тестовые файлы
        const filesToClean = ['test_document.xlsx', 'test_document_updated.xlsx'];
        filesToClean.forEach(file => {
            const filePath = path.join(__dirname, file);
            if (fs.existsSync(filePath)) {
                fs.unlinkSync(filePath);
                console.log(`Удален тестовый файл: ${file}`);
            }
        });
    }
}

// Запускаем тест
if (require.main === module) {
    testDocumentUpdate();
}

module.exports = { testDocumentUpdate };