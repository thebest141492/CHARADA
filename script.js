const users = {}; // Almacén de usuarios registrados

function showRegister() {
    document.getElementById('login-container').style.display = 'none';
    document.getElementById('register-container').style.display = 'block';
}

function showLogin() {
    document.getElementById('register-container').style.display = 'none';
    document.getElementById('login-container').style.display = 'block';
}

function showMenu() {
    document.getElementById('login-container').style.display = 'none';
    document.getElementById('menu-bar').style.display = 'block';
    document.getElementById('qr-generator').style.display = 'none';
    document.getElementById('sales-container').style.display = 'none';
}

function showQRGenerator() {
    document.getElementById('qr-generator').style.display = 'block';
    document.getElementById('sales-container').style.display = 'none';
    document.getElementById('table-container').style.display = 'none';
    document.getElementById('menu-bar').style.display = 'none'; // Ocultar el menú desplegableee
    document.getElementById('menu-toggle-btn').style.display = 'none'; // Ocultar el botón del menú
}

function showSales() {
    document.getElementById('qr-generator').style.display = 'none';
    document.getElementById('sales-container').style.display = 'block';
    document.getElementById('table-container').style.display = 'none';
    document.getElementById('menu-bar').style.display = 'none'; // Ocultar el menú desplegable
    document.getElementById('menu-toggle-btn').style.display = 'none'; // Ocultar el botón del menú
}

function logout() {
    // Ocultar todos los módulos y reiniciar al estado inicial
    document.getElementById('menu-bar').style.display = 'none'; // Ocultar el menú
    document.getElementById('menu-toggle-btn').style.display = 'none'; // Ocultar el botón del menú
    document.getElementById('qr-generator').style.display = 'none'; // Ocultar el generador de QR
    document.getElementById('sales-container').style.display = 'none'; // Ocultar la sección de ventas
    document.getElementById('table-container').style.display = 'none'; // Ocultar la tabla
    document.getElementById('login-container').style.display = 'block'; // Mostrar el inicio de sesión

    // Ocultar el contenido principal si está visible
    document.getElementById('main-content').style.display = 'none';

    // Reiniciar campos de entrada
    document.getElementById('username').value = '';
    document.getElementById('password').value = '';
    document.getElementById('login-error').style.display = 'none'; // Ocultar mensaje de error
}

function register() {
    const username = document.getElementById('register-username').value.trim();
    const password = document.getElementById('register-password').value.trim();
    const errorMessage = document.getElementById('register-error');
    const successMessage = document.getElementById('register-success');

    if (username === '' || password === '') {
        errorMessage.style.display = 'block';
        successMessage.style.display = 'none';
        return;
    }

    if (users[username]) {
        errorMessage.textContent = 'El usuario ya existe';
        errorMessage.style.display = 'block';
        successMessage.style.display = 'none';
        return;
    }

    users[username] = password;
    errorMessage.style.display = 'none';
    successMessage.style.display = 'block';

    setTimeout(() => {
        showLogin(); // Redirigir al inicio de sesión
    }, 2000);
}

function login() {
    const username = document.getElementById('username').value.trim();
    const password = document.getElementById('password').value.trim();
    const errorMessage = document.getElementById('login-error');

    if (users[username] && users[username] === password) {
        document.getElementById('login-container').style.display = 'none';
        document.getElementById('main-content').style.display = 'block';
    } else {
        errorMessage.style.display = 'block';
    }
}

function generarQR() {
    let text = document.getElementById("text").value;
    let qrContainer = document.getElementById("qr-container");

    // Limpiar el contenedor antes de generar un nuevo QR
    qrContainer.innerHTML = "";

    if (text.trim() === "") {
        alert("Ingresa un texto o URL válido");
        return;
    }

    new QRCode(qrContainer, {
        text: text,
        width: 320,
        height: 320
    });
}

function handleExcelUpload() {
    const fileInput = document.getElementById('excel-file');
    const file = fileInput.files[0];

    if (file) {
        const reader = new FileReader();
        reader.onload = function (e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

            generateTable(jsonData);
        };
        reader.readAsArrayBuffer(file);
    }
}

function generateTable(data) {
    const tableHead = document.querySelector('#data-table thead');
    const tableBody = document.querySelector('#data-table tbody');

    tableHead.innerHTML = '';
    tableBody.innerHTML = '';

    if (data.length > 0) {
        const headers = data[0];
        const headerRow = document.createElement('tr');
        headers.forEach(header => {
            const th = document.createElement('th');
            th.textContent = header;
            headerRow.appendChild(th);
        });
        tableHead.appendChild(headerRow);

        data.slice(1).forEach(row => {
            const tableRow = document.createElement('tr');
            row.forEach(cell => {
                const td = document.createElement('td');
                td.textContent = cell;
                tableRow.appendChild(td);
            });
            tableBody.appendChild(tableRow);
        });
    }
}

function populateDataTable(data) {
    const tableContainer = document.getElementById('table-container');
    const table = document.getElementById('data-table');

    // Limpiar tabla existente
    table.innerHTML = '';

    // Generar encabezados
    const headers = data[0];
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    headers.forEach((header) => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);

    // Generar filas de datos
    const tbody = document.createElement('tbody');
    data.slice(1).forEach((row) => {
        const tr = document.createElement('tr');
        row.forEach((cell) => {
            const td = document.createElement('td');
            td.textContent = cell || '';
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
    table.appendChild(tbody);

    // Inicializar DataTables.js
    $(table).DataTable();

    tableContainer.style.display = 'block'; // Mostrar la tabla
}

function formatDate(dateString) {
    // Convertir fecha en formato DD/MM/YYYY a YYYY-MM-DD
    const [day, month, year] = dateString.split('/');
    return `${year}-${month}-${day}`;
}

function toggleEditRow(row, modifyIcon) {
    const inputs = row.querySelectorAll('input');
    const isEditing = modifyIcon.classList.contains('fa-save');

    inputs.forEach((input, index) => {
        input.disabled = isEditing; // Alternar entre habilitar/deshabilitar

        if (!isEditing && index === 9) { // Columna 10 (índice 9)
            input.focus(); // Enfocar el campo
            input.setSelectionRange(input.value.length, input.value.length); // Mover cursor al final
        }
    });

    modifyIcon.className = isEditing ? 'fas fa-edit' : 'fas fa-save';
    modifyIcon.title = isEditing ? 'Modificar' : 'Guardar';

    if (isEditing) {
        alert('Cambios guardados en la fila.');
    }
}

function duplicateRow(row) {
    const tableBody = document.querySelector('#data-table tbody');
    const newRow = document.createElement('tr');

    row.forEach((cell) => {
        const td = document.createElement('td');
        const input = document.createElement('input');
        input.type = 'text';
        input.value = cell || '';
        input.disabled = true; // Deshabilitado por defecto
        td.appendChild(input);
        newRow.appendChild(td);
    });

    // Columna de acciones
    const actionTd = document.createElement('td');
    const actionDiv = document.createElement('div');
    actionDiv.className = 'action-buttons';

    const modifyIcon = document.createElement('i');
    modifyIcon.className = 'fas fa-edit';
    modifyIcon.title = 'Modificar';
    modifyIcon.onclick = () => toggleEditRow(newRow, modifyIcon);

    const deleteIcon = document.createElement('i');
    deleteIcon.className = 'fas fa-trash';
    deleteIcon.title = 'Eliminar';
    deleteIcon.onclick = () => newRow.remove();

    const duplicateIcon = document.createElement('i');
    duplicateIcon.className = 'fas fa-copy';
    duplicateIcon.title = 'Duplicar';
    duplicateIcon.onclick = () => duplicateRow(row);

    actionDiv.appendChild(modifyIcon);
    actionDiv.appendChild(deleteIcon);
    actionDiv.appendChild(duplicateIcon);
    actionTd.appendChild(actionDiv);
    newRow.appendChild(actionTd);

    tableBody.appendChild(newRow);
}

function filterTable() {
    const filterInput = document.getElementById('filter-input').value.toLowerCase();
    const tableBody = document.querySelector('#data-table tbody');
    const rows = tableBody.querySelectorAll('tr');

    rows.forEach((row) => {
        const cells = row.querySelectorAll('td');
        const rowText = Array.from(cells).map(cell => cell.textContent.toLowerCase()).join(' ');
        row.style.display = rowText.includes(filterInput) ? '' : 'none';
    });
}

function toggleMenu() {
    const menuBar = document.getElementById('menu-bar');
    const isMenuVisible = menuBar.style.display === 'block';
    menuBar.style.display = isMenuVisible ? 'none' : 'block';
}

function filterByModel() {
    const selectedModel = document.getElementById('model-select').value;
    const tables = document.querySelectorAll('#data-table-container .group-container');

    tables.forEach(group => {
        const table = group.querySelector('table');
        let hasVisibleRows = false;

        const rows = table.querySelectorAll('tbody tr');
        rows.forEach(row => {
            const modelCell = row.children[5]; // Columna 6
            if (modelCell) {
                const isVisible = selectedModel ? modelCell.textContent.trim() === selectedModel : true;
                row.style.display = isVisible ? '' : 'none';
                if (isVisible) hasVisibleRows = true;
            }
        });

        group.style.display = hasVisibleRows ? '' : 'none'; // Mostrar u ocultar la tabla completa
    });
}

function clearModelFilter() {
    document.getElementById('model-select').value = ''; // Resetear el selector
    document.getElementById('model-select').style.display = 'none'; // Ocultar el selector
    document.getElementById('clear-model-filter').style.display = 'none'; // Ocultar el botón
    filterByModel(); // Aplicar el filtro vacío para mostrar todas las filas
}

function filterByStore() {
    const selectedStore = document.getElementById('store-select').value;
    const tables = document.querySelectorAll('#data-table-container .group-container');
    const clearButton = document.getElementById('clear-store-filter');

    tables.forEach(group => {
        const table = group.querySelector('table');
        let hasVisibleRows = false;

        const rows = table.querySelectorAll('tbody tr');
        rows.forEach(row => {
            const storeCell = row.children[3]; // Columna 4
            if (storeCell) {
                const isVisible = selectedStore ? storeCell.textContent.trim() === selectedStore : true;
                row.style.display = isVisible ? '' : 'none';
                if (isVisible) hasVisibleRows = true;
            }
        });

        group.style.display = hasVisibleRows ? '' : 'none'; // Mostrar u ocultar la tabla completa
    });

    clearButton.style.display = selectedStore ? 'block' : 'none'; // Mostrar botón si hay filtro
}

function clearStoreFilter() {
    document.getElementById('store-select').value = ''; // Resetear el selector
    document.getElementById('store-select').style.display = 'none'; // Ocultar el selector
    document.getElementById('clear-store-filter').style.display = 'none'; // Ocultar el botón
    filterByStore(); // Aplicar el filtro vacío para mostrar todas las filas
}

function toggleModelFilter() {
    const modelSelect = document.getElementById('model-select');
    const clearButton = document.getElementById('clear-model-filter');

    // Alternar visibilidad del selector y el botón
    const isVisible = modelSelect.style.display === 'block';
    modelSelect.style.display = isVisible ? 'none' : 'block';
    clearButton.style.display = isVisible ? 'none' : 'block';

    if (isVisible) {
        // Si se oculta, limpiar el filtro
        modelSelect.value = '';
        filterByModelLogic(); // Aplicar el filtro vacío
    }
}


