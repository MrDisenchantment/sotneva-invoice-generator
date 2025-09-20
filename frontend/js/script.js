// Глобальные переменные
let itemCounter = 1;

// Функция для добавления новой строки в таблицу товаров
function addTableRow() {
    const tbody = document.getElementById('itemsTableBody');
    const tr = document.createElement('tr');
    tr.innerHTML = `
        <td><input type="text" class="item-name"></td>
        <td><input type="number" class="item-quantity" min="1" value="1"></td>
        <td><input type="text" class="item-unit" value="шт."></td>
        <td><input type="number" class="item-price" value="0" step="0,001"></td>
        <td><input type="number" class="item-total" value="0" step="0,001"></td>
        <td><button type="button" class="delete-button">Удалить</button></td>
    `;
    tbody.appendChild(tr);

    // Добавляем обработчики событий для автоматического расчета
    const row = tbody.lastElementChild;
    const quantityInput = row.querySelector('.item-quantity');
    const priceInput = row.querySelector('.item-price');
    const totalInput = row.querySelector('.item-total');
    const deleteButton = row.querySelector('.delete-button');

    // Обработчики для количества и цены
    quantityInput.addEventListener('input', () => calculateRowTotal(row));
    priceInput.addEventListener('input', () => calculateRowTotal(row));
    totalInput.addEventListener('input', () => updateTotals());
    
    // Обработчик для кнопки удаления
    deleteButton.addEventListener('click', () => {
        const container = document.getElementById('itemsContainer');
        const totalCards = container.children.length;
        
        // Удаляем карточку
        itemCard.remove();
        updateItemNumbers();
        
        // Если это была последняя карточка - создаем новую пустую
        if (totalCards === 1) {
            addTableRow();
        }
        
        updateTotals();
    });
}

// Функция для расчета суммы по строке
function calculateRowTotal(row) {
    const quantity = parseFloat(row.querySelector('.item-quantity').value) || 0;
    const price = parseFloat(row.querySelector('.item-price').value) || 0;
    const total = quantity * price;
    row.querySelector('.item-total').value = total.toFixed(2);
    updateTotals();
}

// Функция для обновления итоговых сумм
function updateTotals() {
    const totals = Array.from(document.getElementsByClassName('item-total'))
        .map(input => parseFloat(input.value) || 0)
        .reduce((sum, current) => sum + current, 0);

    const vat = totals * 0.20; // НДС 20%
    const grandTotal = totals + vat;

    document.getElementById('totalAmount').textContent = totals.toFixed(2);
    document.getElementById('vatAmount').textContent = vat.toFixed(2);
    document.getElementById('grandTotal').textContent = grandTotal.toFixed(2);
}

// Функция для преобразования числа в слова (сумма прописью)
function convertNumberToWords(num) {
    const units = ['', 'один', 'два', 'три', 'четыре', 'пять', 'шесть', 'семь', 'восемь', 'девять'];
    const unitsF = ['', 'одна', 'две', 'три', 'четыре', 'пять', 'шесть', 'семь', 'восемь', 'девять']; // для тысяч (женский род)
    const teens = ['десять', 'одиннадцать', 'двенадцать', 'тринадцать', 'четырнадцать', 'пятнадцать', 'шестнадцать', 'семнадцать', 'восемнадцать', 'девятнадцать'];
    const tens = ['', 'десять', 'двадцать', 'тридцать', 'сорок', 'пятьдесят', 'шестьдесят', 'семьдесят', 'восемьдесят', 'девяносто'];
    const hundreds = ['', 'сто', 'двести', 'триста', 'четыреста', 'пятьсот', 'шестьсот', 'семьсот', 'восемьсот', 'девятьсот'];
    
    // Функция для склонения слов
    function declension(number, forms) {
        const cases = [2, 0, 1, 1, 1, 2];
        return forms[(number % 100 > 4 && number % 100 < 20) ? 2 : cases[Math.min(number % 10, 5)]];
    }
    
    // Функция для преобразования трёхзначного числа
    function convertHundreds(num, isFeminine = false) {
        let result = '';
        
        if (num === 0) return '';
        
        // Сотни
        if (num >= 100) {
            result += hundreds[Math.floor(num / 100)] + ' ';
        }
        
        // Десятки и единицы
        const remainder = num % 100;
        
        if (remainder >= 10 && remainder < 20) {
            result += teens[remainder - 10];
        } else {
            if (remainder >= 20) {
                result += tens[Math.floor(remainder / 10)] + ' ';
            }
            const unit = remainder % 10;
            if (unit > 0) {
                result += isFeminine ? unitsF[unit] : units[unit];
            }
        }
        
        return result.trim();
    }
    
    // Обрабатываем целую часть
    const intPart = Math.floor(num);
    
    if (intPart === 0) {
        return 'ноль';
    }
    
    let result = '';
    
    // Миллиарды
    if (intPart >= 1000000000) {
        const billions = Math.floor(intPart / 1000000000);
        result += convertHundreds(billions) + ' ' + declension(billions, ['миллиард', 'миллиарда', 'миллиардов']) + ' ';
    }
    
    // Миллионы
    const millions = Math.floor((intPart % 1000000000) / 1000000);
    if (millions > 0) {
        result += convertHundreds(millions) + ' ' + declension(millions, ['миллион', 'миллиона', 'миллионов']) + ' ';
    }
    
    // Тысячи
    const thousands = Math.floor((intPart % 1000000) / 1000);
    if (thousands > 0) {
        result += convertHundreds(thousands, true) + ' ' + declension(thousands, ['тысяча', 'тысячи', 'тысяч']) + ' ';
    }
    
    // Единицы, десятки, сотни
    const remainder = intPart % 1000;
    if (remainder > 0) {
        result += convertHundreds(remainder);
    }
    
    return result.trim();
}

// Функция для форматирования номера счета
function formatAccountNumber(accountNumber) {
    if (!accountNumber) return '';
    // Удаляем все нецифровые символы
    let value = accountNumber.replace(/\D/g, '');
    // Ограничиваем длину до 20 цифр
    value = value.substring(0, 20);
    // Добавляем пробелы после каждых 4 цифр
    return value.replace(/(\d{4})/g, '$1 ').trim();
}

// Единый обработчик DOMContentLoaded
document.addEventListener('DOMContentLoaded', function() {
    // Установка текущей даты
    const today = new Date().toISOString().split('T')[0];
    document.getElementById('invoiceDate').value = today;

    // Добавление первой строки товаров
    addTableRow();

    // Обработчик кнопки добавления товара
    document.getElementById('addItemButton').addEventListener('click', addTableRow);

    // Обработчик отправки формы
    // В обработчике отправки формы (около строки 180)
    document.getElementById('invoiceForm').addEventListener('submit', async function(e) {
        e.preventDefault();
        
        const submitButton = document.querySelector('.submit-button');
        const originalText = submitButton.textContent;
        
        // Добавляем состояние загрузки
        submitButton.classList.add('loading');
        submitButton.textContent = 'Генерируется...';
        
        // Собираем данные формы
        const formData = {
            replacements: {
                // Банковские реквизиты
                '{{bank_name}}': document.getElementById('receiverBank').value,
                '{{bik}}': document.getElementById('receiverBIK').value,
                '{{cor_account}}': document.getElementById('receiverCorAccount').value,
                '{{inn}}': document.getElementById('receiverINN').value,
                '{{kpp}}': document.getElementById('receiverKPP').value,
                '{{account}}': document.getElementById('receiverAccount').value,
                '{{recipient}}': document.getElementById('receiverName').value,
                
                // Информация о счете
                '{{invoice_number}}': document.getElementById('invoiceNumber').value,
                '{{invoice_date}}': document.getElementById('invoiceDate').value,
                '{{head_warning}}': document.getElementById('headWarning').value,
                
                // Информация о контрагентах
                '{{supplier_info}}': `${document.getElementById('receiverName').value}, ИНН ${document.getElementById('receiverINN').value}, КПП ${document.getElementById('receiverKPP').value}, ${document.getElementById('receiverAddress').value}`,
                '{{customer_info}}': `${document.getElementById('payerName').value}, ИНН ${document.getElementById('payerINN').value}, КПП ${document.getElementById('payerKPP').value}, ${document.getElementById('payerAddress').value}`,
                
                // Итоговая информация
                '{{total_sum}}': document.getElementById('totalAmount').textContent,
                '{{total_vat}}': document.getElementById('vatAmount').textContent,
                '{{total_to_pay}}': document.getElementById('grandTotal').textContent,
                '{{total_items_count}}': document.querySelectorAll('#itemsTableBody tr').length.toString(),
                '{{total_sum_text}}': convertNumberToWords(parseFloat(document.getElementById('grandTotal').textContent)),
                '{{nds_title}}': (() => {
                const vatRate = parseFloat(document.getElementById('vatRate').value) || 0;
                return vatRate === 0 ? 'Без НДС' : `НДС (${vatRate}%)`;
                })(),  
                // Подписи
                '{{director_name}}': document.getElementById('receiverDirector').value,
                '{{accountant_name}}': document.getElementById('receiverAccountant').value
            },
            // В обработчике отправки формы замените сбор данных товаров на:
            items: Array.from(document.querySelectorAll('.item-card')).map((card, index) => ({
                number: index + 1,
                name: card.querySelector('.item-name').value,
                quantity: card.querySelector('.item-quantity').value,
                unit: card.querySelector('.item-unit').value,
                price: card.querySelector('.item-price').value,
                sum: (parseFloat(card.querySelector('.item-quantity').value) || 0) * (parseFloat(card.querySelector('.item-price').value) || 0)
            }))
        };
    
        try {
            const response = await fetch('http://localhost:3000/generate-invoice', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(formData)
            });
    
            if (!response.ok) throw new Error('Ошибка при генерации счета');
    
            // Скачиваем файл
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `Счет №${formData.replacements['{{invoice_number}}']} от ${formData.replacements['{{invoice_date}}']}.pdf`;
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            a.remove();
        } catch (error) {
            console.error(error);
            alert('Ошибка при генерации счета');
        }
        finally {
            // Убираем состояние загрузки
            submitButton.classList.remove('loading');
            submitButton.textContent = originalText;
        }
    });

    // Инициализация подсказок Dadata
    let dadataToken = null;
    
    // Получаем токен с сервера
    fetch('/api/dadata-token')
        .then(response => response.json())
        .then(data => {
            dadataToken = data.token;
            initializeDadataSuggestions();
        })
        .catch(error => {
            console.error('Ошибка при получении токена DaData:', error);
        });
    
    function initializeDadataSuggestions() {
        // Подсказки для организации-продавца
        $("#receiverName").suggestions({
            token: dadataToken,
            type: "PARTY",
            minChars: 4,
            onSelect: function(suggestion) {
                const data = suggestion.data;
                document.getElementById("receiverINN").value = data.inn || '';
                document.getElementById("receiverKPP").value = data.kpp || '';
                document.getElementById("receiverAddress").value = data.address.value || '';
                
                // Автозаполнение данных руководителя
                if (data.management && data.management.name) {
                    document.getElementById("receiverDirector").value = data.management.name;
                }
            }
        });

        // Подсказки для ИНН продавца с автозаполнением руководителя
        $("#receiverINN").on("input", function() {
            const inn = this.value;
            if (inn.length >= 10) {
                fetch("https://suggestions.dadata.ru/suggestions/api/4_1/rs/findById/party", {
                    method: "POST",
                    headers: {
                        "Content-Type": "application/json",
                        "Accept": "application/json",
                        "Authorization": "Token " + dadataToken
                    },
                    body: JSON.stringify({query: inn})
                })
                .then(response => response.json())
                .then(result => {
                    if (result.suggestions && result.suggestions[0]) {
                        const company = result.suggestions[0].data;
                        document.getElementById("receiverName").value = company.name.short_with_opf || company.name.full;
                        document.getElementById("receiverKPP").value = company.kpp || '';
                        document.getElementById("receiverAddress").value = company.address.value || '';
                        
                        // Автозаполнение данных руководителя
                        if (company.management && company.management.name) {
                            document.getElementById("receiverDirector").value = company.management.name;
                        }
                    }
                })
                .catch(error => console.log("error", error));
            }
        });

        // Подсказки для организации-покупателя
        $("#payerName").suggestions({
            token: dadataToken,
            type: "PARTY",
            minChars: 4,
            onSelect: function(suggestion) {
                const data = suggestion.data;
                document.getElementById("payerINN").value = data.inn || '';
                document.getElementById("payerKPP").value = data.kpp || '';
                document.getElementById("payerAddress").value = data.address.value || '';
            }
        });

        // Подсказки для банка
        $("#receiverBank").suggestions({
            token: dadataToken,
            type: "BANK",
            minChars: 4,
            onSelect: function(suggestion) {
                $("#receiverBIK").val(suggestion.data.bic);
                // Форматируем корреспондентский счет при автозаполнении
                const formattedCorAccount = formatAccountNumber(suggestion.data.correspondent_account);
                $("#receiverCorAccount").val(formattedCorAccount);
            }
        });

        // Подсказки для БИК
        $("#receiverBIK").suggestions({
            token: dadataToken,
            type: "BANK",
            minChars: 4,
            onSelect: function(suggestion) {
                $("#receiverBank").val(suggestion.value);
                // Форматируем корреспондентский счет при автозаполнении
                const formattedCorAccount = formatAccountNumber(suggestion.data.correspondent_account);
                $("#receiverCorAccount").val(formattedCorAccount);
            }
        });
    }

    // Форматирование расчетного счета при вводе
    document.getElementById('receiverAccount').addEventListener('input', function(e) {
        this.value = formatAccountNumber(this.value);
    });
    
    // Форматирование корреспондентского счета при вводе
    document.getElementById('receiverCorAccount').addEventListener('input', function(e) {
        this.value = formatAccountNumber(this.value);
    });
});

$(function() {
    // Получаем токен с сервера и инициализируем подсказки для банков
    fetch('/api/dadata-token')
        .then(response => response.json())
        .then(data => {
            $("#bank").suggestions({
                token: data.token,
                type: "BANK",
                onSelect: function(suggestion) {
                    console.log(suggestion);
                }
            });
        })
        .catch(error => {
            console.error('Ошибка при получении токена DaData:', error);
        });
});

// Функция для генерации счета
function generateInvoice() {
    // Получаем данные из формы
    const invoiceData = {
        // ... существующий код сбора данных ...
    };

    // Загружаем шаблон и генерируем файл
    fetch('schet-shablon.xlsx')
        .then(response => {
            if (!response.ok) {
                throw new Error('Не удалось загрузить шаблон');
            }
            return response.arrayBuffer();
        })
        .then(buffer => {
            // Проверяем, что библиотека XLSX загружена
            if (typeof XLSX === 'undefined') {
                throw new Error('Библиотека XLSX не загружена');
            }

            // Читаем файл с сохранением стилей
            const wb = XLSX.read(new Uint8Array(buffer), { 
                type: 'array',
                cellStyles: true,
                cellNF: true,
                cellFormula: true
            });
            const ws = wb.Sheets[wb.SheetNames[0]];
            
            // Сохраняем оригинальные стили и форматирование
            const originalStyles = {};
            Object.keys(ws).forEach(cell => {
                if (cell[0] !== '!') {
                    originalStyles[cell] = {
                        s: ws[cell].s, // стили
                        z: ws[cell].z, // формат
                        f: ws[cell].f, // формулы
                        t: ws[cell].t, // тип
                        w: ws[cell].w  // форматированный текст
                    };
                }
            });
            
            // Заполняем данные в шаблоне
            fillTemplate(ws, invoiceData);
            
            // Восстанавливаем стили и форматирование
            Object.keys(originalStyles).forEach(cell => {
                if (ws[cell]) {
                    ws[cell].s = originalStyles[cell].s;
                    if (originalStyles[cell].z) ws[cell].z = originalStyles[cell].z;
                    if (originalStyles[cell].f) ws[cell].f = originalStyles[cell].f;
                    if (originalStyles[cell].t) ws[cell].t = originalStyles[cell].t;
                }
            });
            
            // Копируем свойства листа
            const wsProps = ['!merges', '!margins', '!cols', '!rows', '!autofilter'];
            wsProps.forEach(prop => {
                if (wb.Sheets[wb.SheetNames[0]][prop]) {
                    ws[prop] = wb.Sheets[wb.SheetNames[0]][prop];
                }
            });
            
            // Сохраняем файл с сохранением стилей
            const newWb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(newWb, ws, "Счет");
            XLSX.writeFile(newWb, `Счет №${invoiceData.invoiceNumber} от ${invoiceData.invoiceDate}.xlsx`, {
                bookType: 'xlsx',
                bookSST: true,
                type: 'array',
                cellStyles: true
            });
        })
        .catch(error => {
            console.error('Ошибка при генерации файла:', error);
            alert('Произошла ошибка при генерации файла: ' + error.message);
        });
}

// Функция для добавления новой карточки товара
function addTableRow() {
    const container = document.getElementById('itemsContainer');
    const itemNumber = container.children.length + 1;
    
    const itemCard = document.createElement('div');
    itemCard.className = 'item-card';
    itemCard.innerHTML = `
        <div class="item-card-header">
            <span class="item-number">${itemNumber}.</span>
            <button type="button" class="delete-button">Удалить</button>
        </div>
        <div class="item-name-input">
            <div class="label-for-input">Товар:</div>
            <input type="text" class="item-name" placeholder="Товар">
        </div>
        <div class="item-details">
            <div class="item-field quantity">
                <div class="label-for-input">Кол-во:</div>
                <input type="number" class="item-quantity" min="1" value="">
                <div class="unit-container">
                <input type="text" class="item-unit" value="шт.">
                </div>
            </div>
            <span class="multiply-sign">×</span>
            <div class="item-field price">
                <div class="label-for-input">Цена:</div>
                <input type="number" class="item-price" value="" step="0,001">
            </div>
            <span class="equals-sign">=</span>
            <div class="item-total">
                <span class="total-value">0</span> ₽
            </div>
        </div>
        <div style="display: flex; padding: 12px 24px; flex-direction: column; align-items: flex-start; gap: 8px; align-self: stretch;"> 
            <div style="height: 1px; align-self: stretch; border-radius: 999px; background: #E3E3E3;"></div> 
        </div>
    `;
    
    container.appendChild(itemCard);

    // Добавляем обработчики событий
    const quantityInput = itemCard.querySelector('.item-quantity');
    const priceInput = itemCard.querySelector('.item-price');
    const deleteButton = itemCard.querySelector('.delete-button');

    // Обработчики для количества и цены
    quantityInput.addEventListener('input', () => calculateRowTotal(itemCard));
    priceInput.addEventListener('input', () => calculateRowTotal(itemCard));
    
    // Обработчик для кнопки удаления
    deleteButton.addEventListener('click', () => {
        const container = document.getElementById('itemsContainer');
        const totalCards = container.children.length;
        
        // Удаляем карточку
        itemCard.remove();
        updateItemNumbers();
        
        // Если это была последняя карточка - создаем новую пустую
        if (totalCards === 1) {
            addTableRow();
        }
        
        updateTotals();
    });
    
    // Обновляем итоги
    updateTotals();
}

// Функция для расчета суммы по карточке
function calculateRowTotal(itemCard) {
    const quantity = parseFloat(itemCard.querySelector('.item-quantity').value) || 0;
    const price = parseFloat(itemCard.querySelector('.item-price').value) || 0;
    const total = quantity * price;
    
    const totalElement = itemCard.querySelector('.total-value');
    totalElement.textContent = total.toLocaleString('ru-RU', {
        minimumFractionDigits: 0,
        maximumFractionDigits: 2
    });
    
    updateTotals();
}

// Функция для обновления номеров карточек
function updateItemNumbers() {
    const cards = document.querySelectorAll('.item-card');
    cards.forEach((card, index) => {
        const numberElement = card.querySelector('.item-number');
        numberElement.textContent = `${index + 1}.`;
    });
}

// Функция для обновления итоговых сумм
// Функция для обновления заголовка НДС
function updateVatTitle() {
    const vatRate = parseFloat(document.getElementById('vatRate').value) || 0;
    const vatTitle = document.getElementById('nds_title');
    
    if (vatRate === 0) {
        vatTitle.textContent = 'Без НДС:';
    } else {
        vatTitle.textContent = 'НДС:';
    }
}

// Обновленная функция для расчета итогов
function updateTotals() {
    const cards = document.querySelectorAll('.item-card');
    let totalSum = 0;
    
    cards.forEach(card => {
        const quantity = parseFloat(card.querySelector('.item-quantity').value) || 0;
        const price = parseFloat(card.querySelector('.item-price').value) || 0;
        totalSum += quantity * price;
    });

    const vatRate = parseFloat(document.getElementById('vatRate').value) || 0;
    const vat = totalSum * (vatRate / 100); // Динамический НДС
    const grandTotal = totalSum + vat;

    document.getElementById('totalAmount').textContent = totalSum.toFixed(2);
    document.getElementById('vatAmount').textContent = vat.toFixed(2);
    document.getElementById('grandTotal').textContent = grandTotal.toFixed(2);
    
    // Обновляем заголовок НДС
    updateVatTitle();
}

// Добавить в DOMContentLoaded обработчик для поля НДС
document.addEventListener('DOMContentLoaded', function() {
    // Обработчик для поля НДС
    document.getElementById('vatRate').addEventListener('input', function() {
        updateTotals();
    });
    
    // Инициализация заголовка НДС
    updateVatTitle();
});

// Функция для очистки карточек товаров
function clearAllItems() {
    // Подтверждение действия
    if (confirm('Вы уверены, что хотите удалить все карточки товаров?')) {
        const itemsContainer = document.getElementById('itemsContainer');
        
        // Удаляем все существующие карточки
        itemsContainer.innerHTML = '';
        
        // Добавляем одну пустую карточку
        addTableRow();
        
        // Обновляем итоговые суммы
        updateTotals();
        
        console.log('Все карточки товаров очищены, добавлена новая пустая карточка');
    }
}