let selectedM29Sheet = null; // Храним выбранный лист М-29

async function uploadM29() {
    const fileInput = document.getElementById('file-upload-m29');
    const file = fileInput.files[0];

    if (!file) {
        alert("Пожалуйста, выберите файл.");
        return;
    }

    const formData = new FormData();
    formData.append('file', file);

    try {
        // Загружаем файл
        const uploadResponse = await fetch('/upload/m29', {
            method: 'POST',
            body: formData
        });

        const uploadResult = await uploadResponse.json();
        document.getElementById('message').innerText = uploadResult.message;

        // Получаем список листов
        const sheetsResponse = await fetch('/get_sheets');
        const sheetsData = await sheetsResponse.json();

        // Отображаем список листов
        const sheetSelection = document.getElementById('sheet-selection-m29');
        const sheetSelect = document.getElementById('sheet-name-m29');

        sheetSelect.innerHTML = ''; // Очищаем список
        sheetsData.sheets.forEach(sheet => {
            const option = document.createElement('option');
            option.value = sheet;
            option.textContent = sheet;
            sheetSelect.appendChild(option);
        });

        sheetSelection.style.display = 'block'; // Показываем выбор листа

        // Сохраняем выбранный лист
        sheetSelect.addEventListener('change', () => {
            selectedM29Sheet = sheetSelect.value;
        });
    } catch (error) {
        console.error("Ошибка при загрузке файла:", error);
        document.getElementById('message').innerText = "Ошибка при загрузке файла.";
    }
}

async function uploadKs2() {
    const fileInput = document.getElementById('file-upload-ks2');
    const file = fileInput.files[0];

    if (!file) {
        alert("Пожалуйста, выберите файл.");
        return;
    }

    const formData = new FormData();
    formData.append('file', file);

    try {
        // Загружаем файл
        const uploadResponse = await fetch('/upload/ks2', {
            method: 'POST',
            body: formData
        });

        const uploadResult = await uploadResponse.json();
        console.log("Результат загрузки КС-2:", uploadResult); // Логируем результат
        document.getElementById('message').innerText = uploadResult.message;

        // Получаем список листов
        const sheetsResponse = await fetch('/get_sheets_ks2');
        const sheetsData = await sheetsResponse.json();
        console.log("Листы КС-2:", sheetsData); // Логируем список листов

        // Отображаем список листов
        const sheetSelection = document.getElementById('sheet-selection-ks2');
        const sheetSelect = document.getElementById('sheet-name-ks2');

        sheetSelect.innerHTML = ''; // Очищаем список
        sheetsData.sheets.forEach(sheet => {
            const option = document.createElement('option');
            option.value = sheet;
            option.textContent = sheet;
            sheetSelect.appendChild(option);
        });

        sheetSelection.style.display = 'block'; // Показываем выбор листа
    } catch (error) {
        console.error("Ошибка при загрузке файла КС-2:", error);
        document.getElementById('message').innerText = "Ошибка при загрузке файла КС-2.";
    }
}


async function uploadSap() {
    const fileInput = document.getElementById('file-upload-sap');
    const file = fileInput.files[0];

    if (!file) {
        alert("Пожалуйста, выберите файл.");
        return;
    }

    const formData = new FormData();
    formData.append('file', file);

    try {
        // Загружаем файл
        const uploadResponse = await fetch('/upload/sap', {
            method: 'POST',
            body: formData
        });

        const uploadResult = await uploadResponse.json();
        console.log("Результат загрузки SAP:", uploadResult); // Логируем результат
        document.getElementById('message').innerText = uploadResult.message;
    } catch (error) {
        console.error("Ошибка при загрузке файла SAP:", error);
        document.getElementById('message').innerText = "Ошибка при загрузке файла SAP.";
    }
}

// Получаем маску для поиска из поля ввода
function getMtrMask() {
    const mtrMaskInput = document.getElementById('mtr-mask').value;
    return mtrMaskInput.split(',').map(item => item.trim());
}

async function unwrapM29() {
    if (!selectedM29Sheet) {
        document.getElementById('message').innerText = "Пожалуйста, выберите лист для М-29.";
        return;
    }

    const mtrMask = getMtrMask(); // Получаем маску из поля ввода

    try {
        // Проверяем, загружен ли файл
        const statusResponse = await fetch('/status');
        const status = await statusResponse.json();

        if (!status.m29_loaded) {
            document.getElementById('message').innerText = "Файл М-29 не загружен.";
            return;
        }

        // Отправляем запрос на /m_unwrap
        const unwrapResponse = await fetch('/m_unwrap', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                m29_name: selectedM29Sheet,
                mtr_mask: mtrMask
            })
        });

        if (unwrapResponse.ok) {
            const blob = await unwrapResponse.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'расскрытая м29.xlsx';
            document.body.appendChild(a);
            a.click();
            a.remove();
            document.getElementById('message').innerText = "Файл успешно обработан и скачан.";
        } else {
            const error = await unwrapResponse.json();
            document.getElementById('message').innerText = error.detail || "Ошибка при обработке файла.";
        }
    } catch (error) {
        console.error("Ошибка при вызове /m_unwrap:", error);
        document.getElementById('message').innerText = "Ошибка при обработке файла.";
    }
}

async function compareM29Ks2() {
    const ks2Sheet = document.getElementById('sheet-name-ks2').value;
    const addedInt = parseInt(document.getElementById('added-int').value, 10); // Получаем значение added_int
    const mtrMask = getMtrMask(); // Получаем маску из поля ввода

    if (!selectedM29Sheet) {
        document.getElementById('message').innerText = "Пожалуйста, выберите лист для М-29.";
        return;
    }

    try {
        // Проверяем, загружены ли файлы
        const statusResponse = await fetch('/status');
        const status = await statusResponse.json();
        console.log("Статус загрузки файлов:", status); // Логируем статус

        if (!status.m29_loaded) {
            document.getElementById('message').innerText = "Файл М-29 не загружен.";
            return;
        }
        if (!status.ks2_loaded) {
            document.getElementById('message').innerText = "Файл КС-2 не загружен.";
            return;
        }

        // Формируем тело запроса
        const requestBody = {
            m29_name: selectedM29Sheet,
            ks2_name: ks2Sheet,
            mtr_mask: mtrMask,
            added_int: addedInt
        };
        console.log("Тело запроса:", requestBody); // Логируем тело запроса

        // Отправляем запрос на /compare_m29_ks2
        const compareResponse = await fetch('/compare_m29_ks2', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(requestBody)
        });

        if (compareResponse.ok) {
            const blob = await compareResponse.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'расхождения м29 и кс2.xlsx';
            document.body.appendChild(a);
            a.click();
            a.remove();
            document.getElementById('message').innerText = "Файл успешно обработан и скачан.";
        } else {
            const error = await compareResponse.json();
            console.error("Ошибка при обработке файла:", error); // Логируем ошибку
            document.getElementById('message').innerText = error.detail || "Ошибка при обработке файла.";
        }
    } catch (error) {
        console.error("Ошибка при вызове /compare_m29_ks2:", error);
        document.getElementById('message').innerText = "Ошибка при обработке файла.";
    }
}

async function compareM29Sap() {
    const mtrMask = getMtrMask(); // Получаем маску из поля ввода

    if (!selectedM29Sheet) {
        document.getElementById('message').innerText = "Пожалуйста, выберите лист для М-29.";
        return;
    }

    try {
        // Проверяем, загружены ли файлы
        const statusResponse = await fetch('/status');
        const status = await statusResponse.json();
        console.log("Статус загрузки файлов:", status); // Логируем статус

        if (!status.m29_loaded) {
            document.getElementById('message').innerText = "Файл М-29 не загружен.";
            return;
        }
        if (!status.sap_loaded) {
            document.getElementById('message').innerText = "Файл SAP не загружен.";
            return;
        }

        // Формируем тело запроса
        const requestBody = {
            m29_name: selectedM29Sheet,
            mtr_mask: mtrMask
        };
        console.log("Тело запроса:", requestBody); // Логируем тело запроса

        // Отправляем запрос на /compare_m29_sap
        const compareResponse = await fetch('/compare_m29_sap', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(requestBody)
        });

        if (compareResponse.ok) {
            const blob = await compareResponse.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'расхождения м29 и sap.xlsx';
            document.body.appendChild(a);
            a.click();
            a.remove();
            document.getElementById('message').innerText = "Файл успешно обработан и скачан.";
        } else {
            const error = await compareResponse.json();
            console.error("Ошибка при обработке файла:", error); // Логируем ошибку
            document.getElementById('message').innerText = error.detail || "Ошибка при обработке файла.";
        }
    } catch (error) {
        console.error("Ошибка при вызове /compare_m29_sap:", error);
        document.getElementById('message').innerText = "Ошибка при обработке файла.";
    }
}
