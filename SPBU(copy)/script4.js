document.addEventListener('DOMContentLoaded', () => {
    const fuelTypeSelect = document.getElementById('fuelType');
    const fuelTables = {
        solar: document.getElementById('solarTable'),
        pertalite: document.getElementById('pertaliteTable'),
        pertamax: document.getElementById('pertamaxTable'),
        pertaminadex: document.getElementById('pertaminadexTable'),
    };
    const tableBodies = {
        solar: document.getElementById('solarTableBody'),
        pertalite: document.getElementById('pertaliteTableBody'),
        pertamax: document.getElementById('pertamaxTableBody'),
        pertaminadex: document.getElementById('pertaminadexTableBody'),
    };

    const singleUnitPriceInputs = {
        solar: document.getElementById('hargaSatuanInputSolar'),
        pertalite: document.getElementById('hargaSatuanInputPertalite'),
        pertamax: document.getElementById('hargaSatuanInputPertamax'),
        pertaminadex: document.getElementById('hargaSatuanInputPertaminadex'), // Perbaikan ID di sini
    };

    const initialDateInput = document.getElementById('initialDate');
    const receiveDateInput = document.getElementById('receiveDate');
    const setDateInput = document.getElementById('setDate');

    let fuelData = {
        solar: { history: [], lastBalance: { liter: 0, rupiah: 0, hasat: 0 }, lastHargaSatuan: 0 },
        pertalite: { history: [], lastBalance: { liter: 0, rupiah: 0, hasat: 0 }, lastHargaSatuan: 0 },
        pertamax: { history: [], lastBalance: { liter: 0, rupiah: 0, hasat: 0 }, lastHargaSatuan: 0 },
        pertaminadex: { history: [], lastBalance: { liter: 0, rupiah: 0, hasat: 0 }, lastHargaSatuan: 0 },
    };

    function saveToLocalStorage() {
        localStorage.setItem('fuelData', JSON.stringify(fuelData));
    }

    function loadFromLocalStorage() {
        const storedData = localStorage.getItem('fuelData');
        if (storedData) {
            const parsedData = JSON.parse(storedData);
            if (parsedData.premium && !parsedData.solar) {
                parsedData.solar = parsedData.premium;
                delete parsedData.premium;
            }
            for (const type in fuelData) {
                if (fuelData.hasOwnProperty(type)) {
                    if (!parsedData.hasOwnProperty(type)) {
                        parsedData[type] = { history: [], lastBalance: { liter: 0, rupiah: 0, hasat: 0 }, lastHargaSatuan: 0 };
                    } else if (!parsedData[type].hasOwnProperty('lastHargaSatuan')) {
                        parsedData[type].lastHargaSatuan = 0;
                    }
                }
            }
            fuelData = parsedData;
        }
    }

    loadFromLocalStorage();

    function setAllDatesToToday() {
        const today = new Date();
        const year = today.getFullYear();
        const month = String(today.getMonth() + 1).padStart(2, '0');
        const day = String(today.getDate()).padStart(2, '0');
        const formattedToday = `${year}-${month}-${day}`;

        if (initialDateInput) initialDateInput.value = formattedToday;
        if (receiveDateInput) receiveDateInput.value = formattedToday;
        if (setDateInput) setDateInput.value = formattedToday;
    }

    setAllDatesToToday();

    function updateSingleUnitPriceInput(fuelType) {
        if (singleUnitPriceInputs[fuelType]) {
            singleUnitPriceInputs[fuelType].value = fuelData[fuelType].lastHargaSatuan || '';
        }
    }

    function showHideUnitPriceInputs() {
        const selectedFuelType = fuelTypeSelect.value;
        for (const type in singleUnitPriceInputs) {
            if (singleUnitPriceInputs.hasOwnProperty(type) && singleUnitPriceInputs[type]) {
                if (type === selectedFuelType) {
                    singleUnitPriceInputs[type].style.display = 'block';
                    updateSingleUnitPriceInput(selectedFuelType);
                } else {
                    singleUnitPriceInputs[type].style.display = 'none';
                }
            }
        }
    }

    function renderTable(fuelType) {
        const tbody = tableBodies[fuelType];
        tbody.innerHTML = '';

        if (!fuelData[fuelType] || !Array.isArray(fuelData[fuelType].history)) {
            fuelData[fuelType] = { history: [], lastBalance: { liter: 0, rupiah: 0, hasat: 0 }, lastHargaSatuan: 0 };
        }

        let totalSetoran = { liter: 0, rupiah: 0, hasat: 0 };
        let totalTerima = { liter: 0, rupiah: 0, hasat: 0 };
        let currentSaldo = { liter: 0, rupiah: 0, hasat: 0 };

        let lastSetoranHasat = 0;
        let lastTerimaHasat = 0;

        if (fuelData[fuelType].history.length > 0) {
            currentSaldo = { ...fuelData[fuelType].history[0].saldo };
        } else {
            currentSaldo = { ...fuelData[fuelType].lastBalance };
        }

        fuelData[fuelType].history.forEach((entry) => {
            const row = tbody.insertRow();
            row.insertCell().textContent = entry.tanggal;
            row.insertCell().textContent = entry.keterangan;

            row.insertCell().textContent = entry.setoran.liter ? entry.setoran.liter.toLocaleString('id-ID') : '';
            row.insertCell().textContent = entry.setoran.rupiah ? entry.setoran.rupiah.toLocaleString('id-ID') : '';
            row.insertCell().textContent = entry.setoran.hasat ? entry.setoran.hasat.toLocaleString('id-ID') : '';
            totalSetoran.liter += parseFloat(entry.setoran.liter) || 0;
            totalSetoran.rupiah += parseFloat(entry.setoran.rupiah) || 0;
            if (entry.setoran.hasat) {
                lastSetoranHasat = parseFloat(entry.setoran.hasat);
            }

            row.insertCell().textContent = entry.terima.liter ? entry.terima.liter.toLocaleString('id-ID') : '';
            row.insertCell().textContent = entry.terima.rupiah ? entry.terima.rupiah.toLocaleString('id-ID') : '';
            row.insertCell().textContent = entry.terima.hasat ? entry.terima.hasat.toLocaleString('id-ID') : '';
            totalTerima.liter += parseFloat(entry.terima.liter) || 0;
            totalTerima.rupiah += parseFloat(entry.terima.rupiah) || 0;
            if (entry.terima.hasat) {
                lastTerimaHasat = parseFloat(entry.terima.hasat);
            }

            currentSaldo = { ...entry.saldo };

            row.insertCell().textContent = currentSaldo.liter.toLocaleString('id-ID');
            row.insertCell().textContent = currentSaldo.rupiah.toLocaleString('id-ID');
            row.insertCell().textContent = currentSaldo.hasat.toLocaleString('id-ID');
        });

        const totalRow = tbody.insertRow();
        totalRow.style.fontWeight = 'bold';
        totalRow.style.backgroundColor = '#e3e3e3';

        totalRow.insertCell().textContent = 'Jumlah';
        totalRow.insertCell().textContent = '';

        totalRow.insertCell().textContent = totalSetoran.liter.toLocaleString('id-ID');
        totalRow.insertCell().textContent = totalSetoran.rupiah.toLocaleString('id-ID');
        totalRow.insertCell().textContent = lastSetoranHasat.toLocaleString('id-ID');

        totalRow.insertCell().textContent = totalTerima.liter.toLocaleString('id-ID');
        totalRow.insertCell().textContent = totalTerima.rupiah.toLocaleString('id-ID');
        totalRow.insertCell().textContent = lastTerimaHasat.toLocaleString('id-ID');

        totalRow.insertCell().textContent = currentSaldo.liter.toLocaleString('id-ID');
        totalRow.insertCell().textContent = currentSaldo.rupiah.toLocaleString('id-ID');
        totalRow.insertCell().textContent = currentSaldo.hasat.toLocaleString('id-ID');

        const akhirRow = tbody.insertRow();
        akhirRow.style.fontWeight = 'bold';
        akhirRow.style.backgroundColor = '#d0ffd0';

        akhirRow.insertCell().textContent = 'Saldo Akhir';
        akhirRow.insertCell().textContent = '';

        akhirRow.insertCell().textContent = '';
        akhirRow.insertCell().textContent = '';
        akhirRow.insertCell().textContent = '';

        akhirRow.insertCell().textContent = '';
        akhirRow.insertCell().textContent = '';
        akhirRow.insertCell().textContent = '';

        akhirRow.insertCell().textContent = currentSaldo.liter.toLocaleString('id-ID');
        akhirRow.insertCell().textContent = currentSaldo.rupiah.toLocaleString('id-ID');
        akhirRow.insertCell().textContent = currentSaldo.hasat.toLocaleString('id-ID');
    }

    function renderAllTables() {
        for (const type in fuelData) {
            if (fuelData.hasOwnProperty(type)) {
                renderTable(type);
            }
        }
    }

    showHideUnitPriceInputs();
    fuelTypeSelect.addEventListener('change', () => {
        console.log(`Jenis BBM terpilih untuk input: ${fuelTypeSelect.value}`);
        showHideUnitPriceInputs();
    });
    renderAllTables();


    function calculateRupiah(literInputId, rupiahOutputId) {
        const liter = parseFloat(document.getElementById(literInputId).value) || 0;
        const selectedFuelType = fuelTypeSelect.value;
        const hargaSatuan = parseFloat(singleUnitPriceInputs[selectedFuelType].value) || 0;

        const rupiahResult = liter * hargaSatuan;
        document.getElementById(rupiahOutputId).value = rupiahResult.toFixed(2);
    }

    const initialLiter = document.getElementById('initialLiter');
    const initialRupiah = document.getElementById('initialRupiah');
    if (initialLiter && initialRupiah) {
        initialLiter.addEventListener('input', () => calculateRupiah('initialLiter', 'initialRupiah'));
    }

    const receiveLiter = document.getElementById('receiveLiter');
    const receiveRupiah = document.getElementById('receiveRupiah');
    if (receiveLiter && receiveRupiah) {
        receiveLiter.addEventListener('input', () => calculateRupiah('receiveLiter', 'receiveRupiah'));
    }

    const setorLiter = document.getElementById('setorLiter');
    const setorRupiah = document.getElementById('setorRupiah');
    if (setorLiter && setorRupiah) {
        setorLiter.addEventListener('input', () => calculateRupiah('setorLiter', 'setorRupiah'));
    }

    for (const type in singleUnitPriceInputs) {
        if (singleUnitPriceInputs.hasOwnProperty(type) && singleUnitPriceInputs[type]) {
            singleUnitPriceInputs[type].addEventListener('input', (event) => {
                const selectedFuelType = fuelTypeSelect.value;
                if (type === selectedFuelType) {
                    fuelData[selectedFuelType].lastHargaSatuan = parseFloat(event.target.value) || 0;
                    saveToLocalStorage();
                    calculateRupiah('initialLiter', 'initialRupiah');
                    calculateRupiah('receiveLiter', 'receiveRupiah');
                    calculateRupiah('setorLiter', 'setorRupiah');
                }
            });
        }
    }

    window.addInitialBalance = function () {
        const selectedFuelType = fuelTypeSelect.value;
        const liter = parseFloat(document.getElementById('initialLiter').value) || 0;
        const hasatInput = parseFloat(singleUnitPriceInputs[selectedFuelType].value) || 0;
        const rupiah = liter * hasatInput;

        let transactionDate = '';
        if (initialDateInput && initialDateInput.value) {
            const [year, month, day] = initialDateInput.value.split('-');
            transactionDate = `${day}/${month}/${year}`;
        } else {
            const now = new Date();
            const day = String(now.getDate()).padStart(2, '0');
            const month = String(now.getMonth() + 1).padStart(2, '0');
            const year = now.getFullYear();
            transactionDate = `${day}/${month}/${year}`;
        }

        if (liter === 0 && rupiah === 0 && hasatInput === 0) {
            alert('Harap masukkan nilai untuk saldo awal.');
            return;
        }

        const newBalance = { liter: liter, rupiah: rupiah, hasat: hasatInput };
        fuelData[selectedFuelType].lastBalance = newBalance;
        fuelData[selectedFuelType].lastHargaSatuan = hasatInput;

        fuelData[selectedFuelType].history.push({
            tanggal: transactionDate,
            keterangan: 'Saldo Awal',
            setoran: { liter: '', rupiah: '', hasat: '' },
            terima: { liter: '', rupiah: '', hasat: '' },
            saldo: { ...newBalance },
        });

        saveToLocalStorage();
        renderTable(selectedFuelType);
        alert('Saldo awal berhasil ditambahkan! 🥳');
        clearInitialInputs();
        setAllDatesToToday();
    };

    window.addReceiveEntry = function () {
        const selectedFuelType = fuelTypeSelect.value;
        const literTerima = parseFloat(document.getElementById('receiveLiter').value) || 0;
        const hasatTerima = parseFloat(singleUnitPriceInputs[selectedFuelType].value) || 0;
        const rupiahTerima = literTerima * hasatTerima;

        let transactionDate = '';
        if (receiveDateInput && receiveDateInput.value) {
            const [year, month, day] = receiveDateInput.value.split('-');
            transactionDate = `${day}/${month}/${year}`;
        } else {
            const now = new Date();
            const day = String(now.getDate()).padStart(2, '0');
            const month = String(now.getMonth() + 1).padStart(2, '0');
            const year = now.getFullYear();
            transactionDate = `${day}/${month}/${year}`;
        }

        if (literTerima === 0 && rupiahTerima === 0 && hasatTerima === 0) {
            alert('Harap masukkan nilai untuk data terima.');
            return;
        }

        const previousBalanceLiter = fuelData[selectedFuelType].history.length > 0
            ? fuelData[selectedFuelType].history[fuelData[selectedFuelType].history.length - 1].saldo.liter
            : fuelData[selectedFuelType].lastBalance.liter;

        const previousBalanceRupiah = fuelData[selectedFuelType].history.length > 0
            ? fuelData[selectedFuelType].history[fuelData[selectedFuelType].history.length - 1].saldo.rupiah
            : fuelData[selectedFuelType].lastBalance.rupiah;

        const newLiterBalance = previousBalanceLiter - literTerima;
        const newRupiahBalance = previousBalanceRupiah - rupiahTerima;
        const newHasatBalance = hasatTerima;

        fuelData[selectedFuelType].lastBalance.liter = newLiterBalance;
        fuelData[selectedFuelType].lastBalance.rupiah = newRupiahBalance;
        fuelData[selectedFuelType].lastBalance.hasat = newHasatBalance;
        fuelData[selectedFuelType].lastHargaSatuan = hasatTerima;

        fuelData[selectedFuelType].history.push({
            tanggal: transactionDate,
            keterangan: 'Terima',
            setoran: { liter: '', rupiah: '', hasat: '' },
            terima: { liter: literTerima, rupiah: rupiahTerima, hasat: hasatTerima },
            saldo: { ...fuelData[selectedFuelType].lastBalance },
        });

        saveToLocalStorage();
        renderTable(selectedFuelType);
        alert('Data terima berhasil ditambahkan! 📥');
        clearReceiveInputs();
        setAllDatesToToday();
    };

    window.addSetorEntry = function () {
        const selectedFuelType = fuelTypeSelect.value;
        const literSetor = parseFloat(document.getElementById('setorLiter').value) || 0;
        const hasatSetor = parseFloat(singleUnitPriceInputs[selectedFuelType].value) || 0;
        const rupiahSetor = literSetor * hasatSetor;

        let transactionDate = '';
        if (setDateInput && setDateInput.value) {
            const [year, month, day] = setDateInput.value.split('-');
            transactionDate = `${day}/${month}/${year}`;
        } else {
            const now = new Date();
            const day = String(now.getDate()).padStart(2, '0');
            const month = String(now.getMonth() + 1).padStart(2, '0');
            const year = now.getFullYear();
            transactionDate = `${day}/${month}/${year}`;
        }

        if (literSetor === 0 && rupiahSetor === 0 && hasatSetor === 0) {
            alert('Harap masukkan nilai untuk data setoran.');
            return;
        }

        const previousBalanceLiter = fuelData[selectedFuelType].history.length > 0
            ? fuelData[selectedFuelType].history[fuelData[selectedFuelType].history.length - 1].saldo.liter
            : fuelData[selectedFuelType].lastBalance.liter;

        const previousBalanceRupiah = fuelData[selectedFuelType].history.length > 0
            ? fuelData[selectedFuelType].history[fuelData[selectedFuelType].history.length - 1].saldo.rupiah
            : fuelData[selectedFuelType].lastBalance.rupiah;

        const newLiterBalance = previousBalanceLiter + literSetor;
        const newRupiahBalance = previousBalanceRupiah + rupiahSetor;
        const newHasatBalance = hasatSetor;

        fuelData[selectedFuelType].lastBalance.liter = newLiterBalance;
        fuelData[selectedFuelType].lastBalance.rupiah = newRupiahBalance;
        fuelData[selectedFuelType].lastBalance.hasat = newHasatBalance;
        fuelData[selectedFuelType].lastHargaSatuan = hasatSetor;

        fuelData[selectedFuelType].history.push({
            tanggal: transactionDate,
            keterangan: 'Setoran',
            setoran: { liter: literSetor, rupiah: rupiahSetor, hasat: hasatSetor },
            terima: { liter: '', rupiah: '', hasat: '' },
            saldo: { ...fuelData[selectedFuelType].lastBalance },
        });

        saveToLocalStorage();
        renderTable(selectedFuelType);
        alert('Data setoran berhasil ditambahkan! ⬆️');
        clearSetorInputs();
        setAllDatesToToday();
    };

    function clearInitialInputs() {
        document.getElementById('initialLiter').value = '';
        document.getElementById('initialRupiah').value = '';
    }

    function clearReceiveInputs() {
        document.getElementById('receiveLiter').value = '';
        document.getElementById('receiveRupiah').value = '';
    }

    function clearSetorInputs() {
        document.getElementById('setorLiter').value = '';
        document.getElementById('setorRupiah').value = '';
    }

    window.clearAllData = function () {
        if (confirm('Apakah Anda yakin ingin menghapus semua data yang tersimpan? Tindakan ini tidak dapat dibatalkan.')) {
            fuelData = {
                solar: { history: [], lastBalance: { liter: 0, rupiah: 0, hasat: 0 }, lastHargaSatuan: 0 },
                pertalite: { history: [], lastBalance: { liter: 0, rupiah: 0, hasat: 0 }, lastHargaSatuan: 0 },
                pertamax: { history: [], lastBalance: { liter: 0, rupiah: 0, hasat: 0 }, lastHargaSatuan: 0 },
                pertaminadex: { history: [], lastBalance: { liter: 0, rupiah: 0, hasat: 0 }, lastHargaSatuan: 0 },
            };
            saveToLocalStorage();
            renderAllTables();
            for (const type in singleUnitPriceInputs) {
                if (singleUnitPriceInputs.hasOwnProperty(type) && singleUnitPriceInputs[type]) {
                    singleUnitPriceInputs[type].value = '';
                }
            }
            alert('Semua data berhasil dihapus! 🗑️');
            setAllDatesToToday();
        }
    };

    // Helper function for current date for filename
    function getCurrentDate() {
        const now = new Date();
        const day = String(now.getDate()).padStart(2, '0');
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const year = now.getFullYear();
        return `${day}/${month}/${year}`;
    }

    window.exportAllFuelsSideBySideToExcel = function () {
        if (typeof XLSX === 'undefined') {
            alert('Pustaka ekspor Excel belum dimuat. Mohon refresh halaman atau periksa koneksi internet Anda.');
            console.error('XLSX library is not loaded.');
            return;
        }

        const fuelTypesInOrder = ['solar', 'pertalite', 'pertamax', 'pertaminadex'];
        const fuelTypeDisplayNames = {
            solar: 'Solar',
            pertalite: 'Pertalite',
            pertamax: 'Pertamax',
            pertaminadex: 'Pertamina Dex',
        };

        const allHistory = {};
        let maxRows = 0;

        const fuelTotals = {};

        for (const fuelType of fuelTypesInOrder) {
            allHistory[fuelType] = fuelData[fuelType].history;
            if (allHistory[fuelType].length > maxRows) {
                maxRows = allHistory[fuelType].length;
            }

            let totalSetoranLiter = 0;
            let totalSetoranRupiah = 0;
            let lastSetoranHasatForExport = 0;

            let totalTerimaLiter = 0;
            let totalTerimaRupiah = 0;
            let lastTerimaHasatForExport = 0;

            let finalSaldo = { liter: 0, rupiah: 0, hasat: 0 };

            // Calculate initial saldo from lastBalance if history is empty
            if (allHistory[fuelType].length > 0) {
                // If there's history, the final saldo will be the saldo of the last entry
                finalSaldo = { ...allHistory[fuelType][allHistory[fuelType].length - 1].saldo };
            } else {
                // If no history, use the lastBalance (which might be the initial set balance)
                finalSaldo = { ...fuelData[fuelType].lastBalance };
            }


            allHistory[fuelType].forEach((entry) => {
                totalSetoranLiter += parseFloat(entry.setoran.liter) || 0;
                totalSetoranRupiah += parseFloat(entry.setoran.rupiah) || 0;
                // Only update lastSetoranHasat if the current entry has a value
                if (entry.setoran.hasat) {
                    lastSetoranHasatForExport = parseFloat(entry.setoran.hasat);
                }

                totalTerimaLiter += parseFloat(entry.terima.liter) || 0;
                totalTerimaRupiah += parseFloat(entry.terima.rupiah) || 0;
                // Only update lastTerimaHasat if the current entry has a value
                if (entry.terima.hasat) {
                    lastTerimaHasatForExport = parseFloat(entry.terima.hasat);
                }

                // The final saldo is the saldo of the last entry in the loop
                // This line updates finalSaldo with each iteration, so it correctly holds the last saldo
                finalSaldo = { ...entry.saldo };
            });

            fuelTotals[fuelType] = {
                totalSetoranLiter: totalSetoranLiter,
                totalSetoranRupiah: totalSetoranRupiah,
                lastSetoranHasat: lastSetoranHasatForExport,
                totalTerimaLiter: totalTerimaLiter,
                totalTerimaRupiah: totalTerimaRupiah,
                lastTerimaHasat: lastTerimaHasatForExport,
                finalSaldo: finalSaldo, // This will correctly be the very last calculated saldo
            };
        }

        if (maxRows === 0 && Object.values(fuelData).every((f) => f.lastBalance.liter === 0 && f.lastBalance.rupiah === 0 && f.lastBalance.hasat === 0 && f.history.length === 0)) {
             // Added check for history length to ensure truly no data
            alert('Tidak ada data untuk diekspor!');
            return;
        }

        const totalRowsForExport = maxRows + 3 + 2; // Header (3) + Data (maxRows) + Jumlah (1) + Saldo Akhir (1)

        const dataForExport = [];
        const headerRow0 = [];
        const headerRow1 = [];
        const headerRow2 = [];
        const merges = [];
        const cols = [];

        const columnsPerFuelType = 11;
        const separatorColumns = 2;
        const totalColumnsPerBlock = columnsPerFuelType + separatorColumns;

        const headerFillColor = 'FFADD8E6'; // Light blue
        const sumRowFillColor = 'FFE3E3E3'; // Light grey
        const finalBalanceRowFillColor = 'FFD0FFD0'; // Light green

        const borderStyle = {
            top: { style: 'thin', color: { auto: 1 } },
            bottom: { style: 'thin', color: { auto: 1 } },
            left: { style: 'thin', color: { auto: 1 } },
            right: { style: 'thin', color: { auto: 1 } },
        };
        const headerStyle = {
            font: { bold: true },
            alignment: { horizontal: 'center', vertical: 'center' },
            border: borderStyle,
            fill: { fgColor: { rgb: headerFillColor } },
        };
        const subHeaderStyle = {
            font: { bold: true },
            alignment: { horizontal: 'center' },
            border: borderStyle,
            fill: { fgColor: { rgb: headerFillColor } },
        };

        let currentColumnOffset = 0;

        fuelTypesInOrder.forEach((fuelType, index) => {
            headerRow0[currentColumnOffset] = fuelTypeDisplayNames[fuelType];
            // Merge for the fuel type header will be added later when applying styles
            
            headerRow1[currentColumnOffset] = 'Tanggal';
            headerRow1[currentColumnOffset + 1] = 'Keterangan';

            headerRow1[currentColumnOffset + 2] = 'Setoran';
            headerRow1[currentColumnOffset + 5] = 'Terima';
            headerRow1[currentColumnOffset + 8] = 'Saldo';

            headerRow2[currentColumnOffset] = 'Tanggal';
            headerRow2[currentColumnOffset + 1] = 'Keterangan';
            headerRow2[currentColumnOffset + 2] = 'Qty (Ltr)';
            headerRow2[currentColumnOffset + 3] = 'Rupiah';
            headerRow2[currentColumnOffset + 4] = 'Hasat';
            headerRow2[currentColumnOffset + 5] = 'Qty (Ltr)';
            headerRow2[currentColumnOffset + 6] = 'Rupiah';
            headerRow2[currentColumnOffset + 7] = 'Hasat';
            headerRow2[currentColumnOffset + 8] = 'Qty (Ltr)';
            headerRow2[currentColumnOffset + 9] = 'Rupiah';
            headerRow2[currentColumnOffset + 10] = 'Hasat';

            cols.push({ wch: 12 }, { wch: 15 }, { wch: 12 }, { wch: 15 }, { wch: 12 }, { wch: 12 }, { wch: 15 }, { wch: 12 }, { wch: 12 }, { wch: 15 }, { wch: 12 });

            if (index < fuelTypesInOrder.length - 1) {
                for (let s = 0; s < separatorColumns; s++) {
                    headerRow0[currentColumnOffset + columnsPerFuelType + s] = '';
                    headerRow1[currentColumnOffset + columnsPerFuelType + s] = '';
                    headerRow2[currentColumnOffset + columnsPerFuelType + s] = '';
                    cols.push({ wch: 3 }); // Spacer column width
                }
            }

            currentColumnOffset += totalColumnsPerBlock;
        });

        dataForExport.push(headerRow0);
        dataForExport.push(headerRow1);
        dataForExport.push(headerRow2);

        for (let i = 0; i < maxRows; i++) {
            const rowData = [];
            let currentCellOffset = 0;
            for (const fuelType of fuelTypesInOrder) {
                const entry = allHistory[fuelType][i];
                if (entry) {
                    rowData[currentCellOffset] = entry.tanggal;
                    rowData[currentCellOffset + 1] = entry.keterangan;
                    rowData[currentCellOffset + 2] = entry.setoran.liter ? entry.setoran.liter : '';
                    rowData[currentCellOffset + 3] = entry.setoran.rupiah ? entry.setoran.rupiah : '';
                    rowData[currentCellOffset + 4] = entry.setoran.hasat ? entry.setoran.hasat : '';
                    rowData[currentCellOffset + 5] = entry.terima.liter ? entry.terima.liter : '';
                    rowData[currentCellOffset + 6] = entry.terima.rupiah ? entry.terima.rupiah : '';
                    rowData[currentCellOffset + 7] = entry.terima.hasat ? entry.terima.hasat : '';
                    rowData[currentCellOffset + 8] = entry.saldo.liter;
                    rowData[currentCellOffset + 9] = entry.saldo.rupiah;
                    rowData[currentCellOffset + 10] = entry.saldo.hasat;
                } else {
                    for (let k = 0; k < columnsPerFuelType; k++) {
                        rowData[currentCellOffset + k] = '';
                    }
                }
                if (fuelType !== fuelTypesInOrder[fuelTypesInOrder.length - 1]) {
                    for (let s = 0; s < separatorColumns; s++) {
                        rowData[currentCellOffset + columnsPerFuelType + s] = '';
                    }
                }
                currentCellOffset += totalColumnsPerBlock;
            }
            dataForExport.push(rowData);
        }

        // ROW: "Jumlah" for all fuel types
        const totalSumRow = [];
        currentColumnOffset = 0;
        fuelTypesInOrder.forEach((fuelType, index) => {
            totalSumRow[currentColumnOffset] = 'Jumlah';
            totalSumRow[currentColumnOffset + 1] = '';
            totalSumRow[currentColumnOffset + 2] = fuelTotals[fuelType].totalSetoranLiter;
            totalSumRow[currentColumnOffset + 3] = fuelTotals[fuelType].totalSetoranRupiah;
            totalSumRow[currentColumnOffset + 4] = fuelTotals[fuelType].lastSetoranHasat;
            totalSumRow[currentColumnOffset + 5] = fuelTotals[fuelType].totalTerimaLiter;
            totalSumRow[currentColumnOffset + 6] = fuelTotals[fuelType].totalTerimaRupiah;
            totalSumRow[currentColumnOffset + 7] = fuelTotals[fuelType].lastTerimaHasat;

            totalSumRow[currentColumnOffset + 8] = fuelTotals[fuelType].finalSaldo.liter;
            totalSumRow[currentColumnOffset + 9] = fuelTotals[fuelType].finalSaldo.rupiah;
            totalSumRow[currentColumnOffset + 10] = fuelTotals[fuelType].finalSaldo.hasat;
            if (index < fuelTypesInOrder.length - 1) {
                for (let s = 0; s < separatorColumns; s++) {
                    totalSumRow[currentColumnOffset + columnsPerFuelType + s] = '';
                }
            }
            currentColumnOffset += totalColumnsPerBlock;
        });
        dataForExport.push(totalSumRow);

        // ROW: "Saldo Akhir" for all fuel types
        const finalBalanceRow = [];
        currentColumnOffset = 0;
        fuelTypesInOrder.forEach((fuelType, index) => {
            finalBalanceRow[currentColumnOffset] = 'Saldo Akhir';
            finalBalanceRow[currentColumnOffset + 1] = '';
            for (let k = 0; k < 6; k++) { // Empty cells for Setoran (3) and Terima (3) sections
                finalBalanceRow[currentColumnOffset + 2 + k] = '';
            }
            finalBalanceRow[currentColumnOffset + 8] = fuelTotals[fuelType].finalSaldo.liter;
            finalBalanceRow[currentColumnOffset + 9] = fuelTotals[fuelType].finalSaldo.rupiah;
            finalBalanceRow[currentColumnOffset + 10] = fuelTotals[fuelType].finalSaldo.hasat;
            if (index < fuelTypesInOrder.length - 1) {
                for (let s = 0; s < separatorColumns; s++) {
                    finalBalanceRow[currentColumnOffset + columnsPerFuelType + s] = '';
                }
            }
            currentColumnOffset += totalColumnsPerBlock;
        });
        dataForExport.push(finalBalanceRow);

        const ws = XLSX.utils.aoa_to_sheet(dataForExport);

        ws['!cols'] = cols;

        // Add merges for headers and summary rows
        let colIndex = 0;
        fuelTypesInOrder.forEach((fuelType, index) => {
            // Main fuel type header merge (Row 0)
            merges.push({
                s: { r: 0, c: colIndex },
                e: { r: 0, c: colIndex + columnsPerFuelType - 1 },
            });

            // "Tanggal" and "Keterangan" headers merge over two rows (Row 1 to Row 2)
            merges.push({ s: { r: 1, c: colIndex }, e: { r: 2, c: colIndex } });
            merges.push({ s: { r: 1, c: colIndex + 1 }, e: { r: 2, c: colIndex + 1 } });

            // "Setoran", "Terima", "Saldo" sub-headers merge (Row 1)
            merges.push({ s: { r: 1, c: colIndex + 2 }, e: { r: 1, c: colIndex + 4 } });
            merges.push({ s: { r: 1, c: colIndex + 5 }, e: { r: 1, c: colIndex + 7 } });
            merges.push({ s: { r: 1, c: colIndex + 8 }, e: { r: 1, c: colIndex + 10 } });

            colIndex += totalColumnsPerBlock;
        });

        colIndex = 0;
        for (let i = 0; i < fuelTypesInOrder.length; i++) {
            const sumRowIndex = maxRows + 3; // +3 because of 0-indexed rows and 3 header rows
            merges.push({
                s: { r: sumRowIndex, c: colIndex },
                e: { r: sumRowIndex, c: colIndex + 1 }, // Merge "Jumlah" and empty cell
            });
            const finalBalanceRowIndex = maxRows + 4; // +4 for the "Saldo Akhir" row
            merges.push({
                s: { r: finalBalanceRowIndex, c: colIndex },
                e: { r: finalBalanceRowIndex, c: colIndex + 7 }, // Merge "Saldo Akhir" and the 7 empty cells
            });
            colIndex += totalColumnsPerBlock;
        }

        ws['!merges'] = merges;

        // Adjust sheet range to include all data and summary rows
        const newRange = XLSX.utils.decode_range(ws['!ref'] || 'A1');
        newRange.e.r = maxRows + 4; // Ensure the range covers all rows, including new totals
        ws['!ref'] = XLSX.utils.encode_range(newRange);

        // Apply borders to all cells within the defined range
        for (let R = 0; R <= newRange.e.r; ++R) {
            for (let C = 0; C <= newRange.e.c; ++C) {
                const cell_address = XLSX.utils.encode_cell({ r: R, c: C });
                if (!ws[cell_address]) {
                    ws[cell_address] = { v: undefined }; // Ensure cell exists to apply style
                }
                // Only apply border if no other style is explicitly set (e.g., header, sum, final balance styles will override this)
                // Or, ensure border is part of all styles
                if (!ws[cell_address].s) {
                    ws[cell_address].s = { border: borderStyle };
                } else if (!ws[cell_address].s.border) {
                    ws[cell_address].s.border = borderStyle;
                }
            }
        }

        // Apply specific styles for headers and summary rows
        colIndex = 0;
        fuelTypesInOrder.forEach((fuelType, index) => {
            // Apply style to the merged fuel type header cell (Row 0)
            const cellFuelTypeHeader = ws[XLSX.utils.encode_cell({ r: 0, c: colIndex })];
            if (cellFuelTypeHeader) cellFuelTypeHeader.s = { ...headerStyle }; // Already includes border

            // Apply style to "Tanggal" and "Keterangan" headers (merged, Row 1 & 2)
            const tanggalHeaderCell1 = ws[XLSX.utils.encode_cell({ r: 1, c: colIndex })];
            if (tanggalHeaderCell1) tanggalHeaderCell1.s = { ...subHeaderStyle, alignment: { horizontal: 'center', vertical: 'center' } };
            const keteranganHeaderCell1 = ws[XLSX.utils.encode_cell({ r: 1, c: colIndex + 1 })];
            if (keteranganHeaderCell1) keteranganHeaderCell1.s = { ...subHeaderStyle, alignment: { horizontal: 'center', vertical: 'center' } };

            // Apply style to "Setoran", "Terima", "Saldo" headers (merged, Row 1)
            const setoranHeaderCell = ws[XLSX.utils.encode_cell({ r: 1, c: colIndex + 2 })];
            if (setoranHeaderCell) setoranHeaderCell.s = { ...subHeaderStyle };
            const terimaHeaderCell = ws[XLSX.utils.encode_cell({ r: 1, c: colIndex + 5 })];
            if (terimaHeaderCell) terimaHeaderCell.s = { ...subHeaderStyle };
            const saldoHeaderCell = ws[XLSX.utils.encode_cell({ r: 1, c: colIndex + 8 })];
            if (saldoHeaderCell) saldoHeaderCell.s = { ...subHeaderStyle };

            // Apply style to the second row of sub-headers (Qty, Rupiah, Hasat) (Row 2)
            // Ensure borders are applied correctly for these specific cells
            for (let C = 0; C < columnsPerFuelType; ++C) {
                if (C === 0 || C === 1) { // Skip Tanggal and Keterangan as they are merged from row 1
                    continue;
                }
                const cell_address = XLSX.utils.encode_cell({ r: 2, c: colIndex + C });
                if (!ws[cell_address]) {
                    ws[cell_address] = { v: undefined };
                }
                ws[cell_address].s = { ...subHeaderStyle };
            }

            colIndex += totalColumnsPerBlock;
        });

        colIndex = 0;
        for (let i = 0; i < fuelTypesInOrder.length; i++) {
            const sumRowIndex = maxRows + 3;
            const finalBalanceRowIndex = maxRows + 4;

            // Style for "Jumlah" row (including the merged 'Jumlah' cell)
            for (let C = colIndex; C < colIndex + columnsPerFuelType; ++C) {
                const cell_address = XLSX.utils.encode_cell({ r: sumRowIndex, c: C });
                if (!ws[cell_address]) ws[cell_address] = { v: undefined }; // Ensure cell exists
                if (C === colIndex) { // "Jumlah" cell
                     ws[cell_address].s = {
                        font: { bold: true },
                        fill: { fgColor: { rgb: sumRowFillColor } },
                        border: borderStyle,
                        alignment: { horizontal: 'center', vertical: 'center' },
                    };
                } else if (C === colIndex + 1) { // Merged empty cell next to "Jumlah"
                     ws[cell_address].s = {
                        font: { bold: true }, // Inherit bold from merge region
                        fill: { fgColor: { rgb: sumRowFillColor } },
                        border: borderStyle,
                    };
                } else { // Other data cells in "Jumlah" row
                    ws[cell_address].s = {
                        font: { bold: true },
                        fill: { fgColor: { rgb: sumRowFillColor } },
                        border: borderStyle,
                    };
                }
            }

            // Style for "Saldo Akhir" row (including the merged 'Saldo Akhir' cell)
            for (let C = colIndex; C < colIndex + columnsPerFuelType; ++C) {
                const cell_address = XLSX.utils.encode_cell({ r: finalBalanceRowIndex, c: C });
                if (!ws[cell_address]) ws[cell_address] = { v: undefined }; // Ensure cell exists
                if (C === colIndex) { // "Saldo Akhir" cell
                     ws[cell_address].s = {
                        font: { bold: true },
                        fill: { fgColor: { rgb: finalBalanceRowFillColor } },
                        border: borderStyle,
                        alignment: { horizontal: 'center', vertical: 'center' },
                    };
                } else if (C >= colIndex + 2 && C <= colIndex + 7) { // Merged empty cells for Setoran and Terima
                    ws[cell_address].s = {
                        font: { bold: true }, // Inherit bold from merge region
                        fill: { fgColor: { rgb: finalBalanceRowFillColor } },
                        border: borderStyle,
                    };
                } else if (C >= colIndex + 8 && C <= colIndex + 10) { // Saldo cells
                    ws[cell_address].s = {
                        font: { bold: true },
                        fill: { fgColor: { rgb: finalBalanceRowFillColor } },
                        border: borderStyle,
                    };
                }
            }
            colIndex += totalColumnsPerBlock;
        }

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Rekap BBM');

        const fileName = `Rekap_BBM_Side_By_Side_${getCurrentDate().replace(/\//g, '-')}.xlsx`;
        XLSX.writeFile(wb, fileName);
        alert('Data berhasil diekspor ke Excel! 📊');
    };
});
