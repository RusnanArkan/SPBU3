document.addEventListener('DOMContentLoaded', () => {
    // ... (existing variable declarations) ...
    const initialBalanceInput = document.getElementById('initialBalance');
    const setBalanceBtn = document.getElementById('setBalanceBtn');
    const descriptionInput = document.getElementById('description');
    const accountSearchInput = document.getElementById('accountSearch');
    const accountSelect = document.getElementById('accountSelect');
    const transactionTypeSelect = document.getElementById('transactionType');
    const nominalInput = document.getElementById('nominal');
    const addTransactionBtn = document.getElementById('addTransactionBtn');
    const transactionTableBody = document.getElementById('transactionTableBody');
    const clearAllDataBtn = document.getElementById('clearAllDataBtn');
    const downloadXLSXBtn = document.getElementById('downloadXLSXBtn');
    const viewSelector = document.getElementById('viewSelector');
    const mainTableSection = document.getElementById('mainTableSection');
    const recapTableSection = document.getElementById('recapTableSection');
    const recapTableBody = document.getElementById('recapTableBody');

    const accountDetailSection = document.getElementById('accountDetailSection');
    const detailAccountSelector = document.getElementById('detailAccountSelector');
    const detailTransactionTableBody = document.getElementById('detailTransactionTableBody');
    const detailAccountSearchInput = document.getElementById('detailAccountSearch');

    // Tambahan elemen untuk tombol Export Excel
    const exportMainTableBtn = document.getElementById('exportMainTableBtn');
    const exportRecapTableBtn = document.getElementById('exportRecapTableBtn');
    const exportDetailTableBtn = document.getElementById('exportDetailTableBtn');

    // NEW: Date input element
    const transactionDateInput = document.getElementById('transactionDate'); // Make sure this ID matches your HTML

    let initialStartingBalance = 0;
    let transactions = [];
    let accounts = [];

    const defaultAccounts = [
        { code: '111000', name: 'Kas' },
        { code: '112001', name: 'Bank Mandiri' },
        { code: '112002', name: 'Bank BTPN' },
        { code: '112009', name: 'Bank BNI' },
        { code: '114100', name: 'Piutang Dagang' },
        { code: '114200', name: 'Piutang Karyawan' },
        { code: '114300', name: 'Piutang Pajak' },
        { code: '114900', name: 'Piutang Lain-Lain' },
        { code: '115100', name: 'Uang Muka Solar' },
        { code: '115110', name: 'Uang Muka Pertamina Dex' },
        { code: '115120', name: 'Uang Muka Pertamax' },
        { code: '115200', name: 'Uang Muka Premium' },
        { code: '116001', name: 'Uang Muka Pertalite' },
        { code: '116100', name: 'Uang Muka Bonus Manajemen' },
        { code: '116200', name: 'Uang Muka Bagi Laba' },
        { code: '117100', name: 'Biaya Dibayar Dimuka' },
        { code: '118100', name: 'Persediaan Solar' },
        { code: '118200', name: 'Persediaan Premium' },
        { code: '118300', name: 'Persediaan Pelumas' },
        { code: '131000', name: 'TANAH' },
        { code: '132000', name: 'BANGUNAN' },
        { code: '133000', name: 'MESIN DAN INSTALASI' },
        { code: '134000', name: 'FURNITURE DAN FIXTURE' },
        { code: '134100', name: 'Akum Penyus. Furniture Dan Fixture' },
        { code: '134200', name: 'Akum Penyus. Inst/Mesin/ alat' },
        { code: '211000', name: 'Utang Dagang' },
        { code: '221000', name: 'Utang Kredit Js Raharaja' },
        { code: '229100', name: 'Utang Pajak' },
        { code: '229200', name: 'Utang Pada Pihak Ketiga' },
        { code: '229210', name: 'Utang Pada M Thamrin Yh' },
        { code: '229220', name: 'Utang Pada Ibu Zauriah' },
        { code: '229230', name: 'Utang Pada PT ADNANTA' },
        { code: '229290', name: 'Utang Lain-Lain' },
        { code: '240000', name: 'Penjualan Diterima Dimuka' },
        { code: '250000', name: 'Biaya Yang Masih Harus Dibayar' },
        { code: '310000', name: 'MODAL SAHAM DISETOR' },
        { code: '320000', name: 'LABA DITAHAN' },
        { code: '330000', name: 'CADANGAN UMUM' },
        { code: '340000', name: 'LABA PERIODE LALU' },
        { code: '350000', name: 'LABA PERIODE BERJALAN' },
        { code: '410100', name: 'Penjualan Solar' },
        { code: '410110', name: 'Penjualan Pertamina Dex' },
        { code: '410200', name: 'Penjualan Premium' },
        { code: '410300', name: 'Penjualan Pertamax' },
        { code: '410301', name: 'Penjualan Pertalite' },
        { code: '421000', name: 'Laba Kotor' },
        { code: '430000', name: 'PENDAPATAN LAIN-LAIN' },
        { code: '510000', name: 'HARGA POKOK PENJUALAN' },
        { code: '521100', name: 'Ongkos Setor' },
        { code: '521200', name: 'Biaya Material' },
        { code: '521300', name: 'Biaya Hewanana Migas' },
        { code: '521400', name: 'Biaya Tip Sopir Tangki' },
        { code: '522100', name: 'Biaya Promosi dan Iklan' },
        { code: '523100', name: 'Biaya Perbaikan/Perawatan' },
        { code: '524100', name: 'Biaya Pegawai' },
        { code: '524110', name: 'Biaya Gaji Karyawan' },
        { code: '524120', name: 'Biaya Lembur Karyawan' },
        { code: '524130', name: 'Biaya Astek/Insentif/Thr' },
        { code: '524200', name: 'Biaya Kantor / Administrasi' },
        { code: '529100', name: 'Biaya Bunga Pinjaman' },
        { code: '529200', name: 'Biaya Pajak' },
        { code: '529300', name: 'Biaya Pph 21/25/29' },
    ];

    const accountStorageKey = 'accountsData_kaskecil';

    function saveToLocalStorage() {
        localStorage.setItem('initialStartingBalance_kaskecil', JSON.stringify(initialStartingBalance));
        localStorage.setItem('transactions_kaskecil', JSON.stringify(transactions));
        localStorage.setItem(accountStorageKey, JSON.stringify(accounts));
        console.log('Data saved to Local Storage.');
    }

    function loadFromLocalStorage() {
        const savedInitialBalance = localStorage.getItem('initialStartingBalance_kaskecil');
        const savedTransactions = localStorage.getItem('transactions_kaskecil');
        const savedAccountsData = localStorage.getItem(accountStorageKey);

        if (savedInitialBalance !== null) {
            initialStartingBalance = JSON.parse(savedInitialBalance);
        }
        if (savedTransactions !== null) {
            transactions = JSON.parse(savedTransactions);
            // NEW: Ensure transactions have a 'date' property. If not, set it to today's date.
            transactions.forEach(t => {
                if (!t.date) {
                    t.date = new Date().toLocaleDateString('id-ID');
                }
            });
        }

        if (savedAccountsData !== null) {
            try {
                const parsedAccounts = JSON.parse(savedAccountsData);
                if (Array.isArray(parsedAccounts) && parsedAccounts.length > 0) {
                    accounts = parsedAccounts;
                } else {
                    console.warn(`Data 'accountsData' di Local Storage (Kas Kecil) tidak valid atau kosong. Menggunakan default accounts.`);
                    accounts = [...defaultAccounts];
                    saveToLocalStorage();
                }
            } catch (e) {
                console.error(`Error parsing 'accountsData' (Kas Kecil) from Local Storage:`, e);
                accounts = [...defaultAccounts];
                saveToLocalStorage();
            }
        } else {
            accounts = [...defaultAccounts];
            saveToLocalStorage();
        }
        console.log(`Data loaded from Local Storage for Kas Kecil.`);
    }

    function loadAccountsToSelects(mainFilter = '', detailFilter = '') {
        accountSelect.innerHTML = '';
        const filteredMainAccounts = accounts.filter((account) => account.name.toLowerCase().includes(mainFilter.toLowerCase()) || account.code.toLowerCase().includes(mainFilter.toLowerCase()));

        if (filteredMainAccounts.length === 0) {
            const option = document.createElement('option');
            option.value = '';
            option.textContent = 'Tidak ada akun ditemukan';
            option.disabled = true;
            accountSelect.appendChild(option);
        } else {
            filteredMainAccounts.forEach((account) => {
                const optionMain = document.createElement('option');
                optionMain.value = `${account.code} - ${account.name}`;
                optionMain.textContent = `${account.code} - ${account.name}`;
                optionMain.dataset.defaultDescription = account.defaultDescription || '';
                accountSelect.appendChild(optionMain);
            });
        }

        detailAccountSelector.innerHTML = '<option value="">-- Pilih Akun --</option>';
        const filteredDetailAccounts = accounts.filter((account) => account.name.toLowerCase().includes(detailFilter.toLowerCase()) || account.code.toLowerCase().includes(detailFilter.toLowerCase()));

        if (filteredDetailAccounts.length === 0) {
            const detailOption = document.createElement('option');
            detailOption.value = '';
            detailOption.textContent = 'Tidak ada akun ditemukan';
            detailOption.disabled = true;
            detailAccountSelector.appendChild(detailOption);
        } else {
            filteredDetailAccounts.forEach((account) => {
                const optionDetail = document.createElement('option');
                optionDetail.value = account.code;
                optionDetail.textContent = `${account.code} - ${account.name}`;
                detailAccountSelector.appendChild(optionDetail);
            });
        }
    }

    function exportTableToExcel(tableID, filename = '') {
        const table = document.getElementById(tableID);
        if (!table) {
            console.error(`Table with ID "${tableID}" not found.`);
            return;
        }

        const wb = XLSX.utils.table_to_book(table, { sheet: "SheetJS" });
        filename = filename ? filename + '.xlsx' : 'excel_data.xlsx';
        XLSX.writeFile(wb, filename);
        alert(`Tabel berhasil diekspor ke ${filename}`);
    }

    setBalanceBtn.addEventListener('click', () => {
        const balance = parseFloat(initialBalanceInput.value);
        // NEW: Get the date for the initial balance
        const initialBalanceDate = transactionDateInput.value;

        if (!isNaN(balance) && balance >= 0) {
            initialStartingBalance = balance;
            transactions = []; // Clear existing transactions
            // NEW: Add a "Saldo Awal" transaction with the selected date
            if (initialBalanceDate) {
                transactions.push({
                    date: new Date(initialBalanceDate).toLocaleDateString('id-ID'), // Format date
                    description: 'Saldo Awal',
                    accountCode: '-',
                    accountName: '-',
                    type: '-',
                    nominal: initialStartingBalance,
                    penerimaan: initialStartingBalance,
                    pengeluaran: 0,
                });
            } else {
                 transactions.push({
                    date: new Date().toLocaleDateString('id-ID'), // Fallback to current date
                    description: 'Saldo Awal',
                    accountCode: '-',
                    accountName: '-',
                    type: '-',
                    nominal: initialStartingBalance,
                    penerimaan: initialStartingBalance,
                    pengeluaran: 0,
                });
            }


            renderTables();
            saveToLocalStorage();
            alert(`Saldo awal berhasil diatur: Rp ${initialStartingBalance.toLocaleString('id-ID')}`);
        } else {
            alert('Mohon masukkan saldo awal yang valid (angka positif).');
        }
    });

    accountSearchInput.addEventListener('input', (e) => {
        loadAccountsToSelects(e.target.value, detailAccountSearchInput.value);
    });

    accountSelect.addEventListener('change', () => {
        const selectedOption = accountSelect.options[accountSelect.selectedIndex];
        if (selectedOption) {
            accountSearchInput.value = selectedOption.value;
        }
    });

    detailAccountSearchInput.addEventListener('input', (e) => {
        loadAccountsToSelects(accountSearchInput.value, e.target.value);
        if (detailAccountSelector.options.length === 2 && detailAccountSelector.options[1].value !== '') {
            detailAccountSelector.value = detailAccountSelector.options[1].value;
            renderAccountDetailTable(detailAccountSelector.value);
        } else {
            renderAccountDetailTable('');
        }
    });

    addTransactionBtn.addEventListener('click', () => {
        const description = descriptionInput.value.trim();
        const selectedAccount = accountSelect.value;
        const transactionType = transactionTypeSelect.value;
        const nominal = parseFloat(nominalInput.value);
        // NEW: Get the selected transaction date
        const transactionDate = transactionDateInput.value;

        if (!description || !selectedAccount || !transactionType || isNaN(nominal) || nominal <= 0 || !transactionDate) {
            alert('Mohon lengkapi semua kolom dengan data yang valid, termasuk tanggal transaksi.');
            return;
        }

        // Use the selected date, format it for display
        const dateForDisplay = new Date(transactionDate).toLocaleDateString('id-ID');

        let penerimaan = 0;
        let pengeluaran = 0;

        if (transactionType === 'D') {
            penerimaan = nominal;
        } else if (transactionType === 'K') {
            pengeluaran = nominal;
        }

        const [accountCode, accountName] = selectedAccount.split(' - ');

        const newTransaction = {
            date: dateForDisplay, // Use the selected date
            description: description,
            accountCode: accountCode,
            accountName: accountName,
            type: transactionType,
            nominal: nominal,
            penerimaan: penerimaan,
            pengeluaran: pengeluaran,
        };

        transactions.push(newTransaction);
        renderTables();
        saveToLocalStorage();
        clearForm();
    });

    clearAllDataBtn.addEventListener('click', () => {
        if (confirm('Apakah Anda yakin ingin menghapus semua data transaksi dan saldo awal? Tindakan ini tidak dapat dibatalkan.')) {
            initialStartingBalance = 0;
            transactions = [];
            localStorage.removeItem('initialStartingBalance_kaskecil');
            localStorage.removeItem('transactions_kaskecil');
            localStorage.removeItem(accountStorageKey);

            accounts = [...defaultAccounts];
            saveToLocalStorage();

            renderTables();
            clearForm();
            alert('Semua data berhasil dihapus.');
        }
    });

    if (exportMainTableBtn) {
        exportMainTableBtn.addEventListener('click', () => {
            exportTableToExcel('mainTable', 'Buku_Kas_Harian');
        });
    }
    if (exportRecapTableBtn) {
        exportRecapTableBtn.addEventListener('click', () => {
            exportTableToExcel('recapTable', 'Rekap_Akun_Transaksi');
        });
    }
    if (exportDetailTableBtn) {
        exportDetailTableBtn.addEventListener('click', () => {
            const selectedAccount = detailAccountSelector.options[detailAccountSelector.selectedIndex];
            let filename = 'Detail_Transaksi_Akun';
            if (selectedAccount && selectedAccount.value) {
                filename = `Detail_Transaksi_${selectedAccount.textContent.replace(/ - /g, '_')}`;
            }
            exportTableToExcel('detailTable', filename);
        });
    }

    viewSelector.addEventListener('change', () => {
        const selectedView = viewSelector.value;
        mainTableSection.classList.add('hidden');
        recapTableSection.classList.add('hidden');
        accountDetailSection.classList.add('hidden');

        if (exportMainTableBtn) exportMainTableBtn.classList.add('hidden');
        if (exportRecapTableBtn) exportRecapTableBtn.classList.add('hidden');
        if (exportDetailTableBtn) exportDetailTableBtn.classList.add('hidden');

        if (selectedView === 'recap') {
            recapTableSection.classList.remove('hidden');
            if (exportRecapTableBtn) exportRecapTableBtn.classList.remove('hidden');
        } else if (selectedView === 'account_detail') {
            accountDetailSection.classList.remove('hidden');
            if (exportDetailTableBtn) exportDetailTableBtn.classList.remove('hidden');
            detailAccountSearchInput.value = '';
            loadAccountsToSelects(accountSearchInput.value, '');
            renderAccountDetailTable('');
        } else {
            mainTableSection.classList.remove('hidden');
            if (exportMainTableBtn) exportMainTableBtn.classList.remove('hidden');
        }
        renderTables();
    });

    function renderMainTable() {
        transactionTableBody.innerHTML = '';

        let runningBalance = initialStartingBalance;

        // NEW: Display initial balance row with its stored date
        const initialBalanceEntry = transactions.find(t => t.description === 'Saldo Awal');
        const initialBalanceDate = initialBalanceEntry ? initialBalanceEntry.date : new Date().toLocaleDateString('id-ID'); // Fallback

        const initialBalanceRow = transactionTableBody.insertRow();
        initialBalanceRow.innerHTML = `
            <td>${initialBalanceDate}</td>
            <td>Saldo Awal</td>
            <td>-</td>
            <td>-</td>
            <td>-</td>
            <td>-</td>
            <td>-</td>
            <td>-</td>
            <td>${initialStartingBalance.toLocaleString('id-ID')}</td>
        `;

        const otherTransactions = transactions.filter(t => t.description !== 'Saldo Awal'); // Filter out the initial balance entry for subsequent rendering

        if (otherTransactions.length === 0 && initialStartingBalance === 0) { // Adjust condition
            const noDataRow = transactionTableBody.insertRow();
            noDataRow.innerHTML = `<td colspan="9" style="text-align: center;">Belum ada transaksi lain.</td>`;
        } else {
            // Sort transactions by date (optional, but good for history)
            otherTransactions.sort((a, b) => new Date(a.date.split('/').reverse().join('-')) - new Date(b.date.split('/').reverse().join('-'))); // Assumes DD/MM/YYYY

            // Recalculate running balance based on sorted transactions
            runningBalance = initialStartingBalance; // Reset for recalculation from the initial balance

            otherTransactions.forEach((transaction) => {
                if (transaction.type === 'D') {
                    runningBalance += transaction.nominal;
                } else if (transaction.type === 'K') {
                    runningBalance -= transaction.nominal;
                }

                const row = transactionTableBody.insertRow();
                row.innerHTML = `
                    <td>${transaction.date}</td>
                    <td>${transaction.description}</td>
                    <td>${transaction.accountCode}</td>
                    <td>${transaction.accountName}</td>
                    <td>${transaction.type}</td>
                    <td>${transaction.nominal.toLocaleString('id-ID')}</td>
                    <td>${transaction.penerimaan.toLocaleString('id-ID')}</td>
                    <td>${transaction.pengeluaran.toLocaleString('id-ID')}</td>
                    <td>${runningBalance.toLocaleString('id-ID')}</td>
                `;
            });
        }
    }

    function generateRecapTable() {
        recapTableBody.innerHTML = '';

        const initialBalanceRecapRow = recapTableBody.insertRow();
        initialBalanceRecapRow.innerHTML = `
            <td>-</td>
            <td>-</td>
            <td>Saldo Awal</td>
            <td>${initialStartingBalance.toLocaleString('id-ID')}</td>
            <td>-</td>
        `;

        const accountRecap = {};
        let totalOverallPenerimaan = 0;
        let totalOverallPengeluaran = 0;

        // Exclude the 'Saldo Awal' entry from recap calculations if it's explicitly added as a transaction
        const transactionsForRecap = transactions.filter(t => t.description !== 'Saldo Awal');

        transactionsForRecap.forEach((transaction) => {
            const key = `${transaction.accountCode}-${transaction.accountName}`;
            if (!accountRecap[key]) {
                accountRecap[key] = {
                    code: transaction.accountCode,
                    name: transaction.accountName,
                    totalPenerimaan: 0,
                    totalPengeluaran: 0,
                };
            }

            if (transaction.type === 'D') {
                accountRecap[key].totalPenerimaan += transaction.nominal;
                totalOverallPenerimaan += transaction.nominal;
            } else if (transaction.type === 'K') {
                accountRecap[key].totalPengeluaran += transaction.nominal;
                totalOverallPengeluaran += transaction.nominal;
            }
        });

        const sortedRecap = Object.values(accountRecap).sort((a, b) => {
            return a.code.localeCompare(b.code);
        });

        if (sortedRecap.length === 0 && initialStartingBalance === 0) {
            const noDataRow = recapTableBody.insertRow();
            noDataRow.innerHTML = `<td colspan="5" style="text-align: center;">Belum ada data rekap transaksi.</td>`;
        } else {
            sortedRecap.forEach((recap) => {
                const rub = recap.code.substring(0, 2);
                const row = recapTableBody.insertRow();
                row.innerHTML = `
                    <td>${rub}</td>
                    <td>${recap.code}</td>
                    <td>${recap.name}</td>
                    <td>${recap.totalPenerimaan.toLocaleString('id-ID')}</td>
                    <td>${recap.totalPengeluaran.toLocaleString('id-ID')}</td>
                `;
            });
        }

        const totalRow = recapTableBody.insertRow();
        totalRow.style.fontWeight = 'bold';
        totalRow.innerHTML = `
            <td colspan="3" style="text-align: center;">Jumlah Keseluruhan</td>
            <td>${totalOverallPenerimaan.toLocaleString('id-ID')}</td>
            <td>${totalOverallPengeluaran.toLocaleString('id-ID')}</td>
        `;

        let currentOverallBalance = initialStartingBalance;
        transactionsForRecap.forEach((transaction) => { // Use transactionsForRecap here
            if (transaction.type === 'D') {
                currentOverallBalance += transaction.nominal;
            } else if (transaction.type === 'K') {
                currentOverallBalance -= transaction.nominal;
            }
        });

        const finalBalanceRow = recapTableBody.insertRow();
        finalBalanceRow.style.fontWeight = 'bold';
        finalBalanceRow.style.backgroundColor = '#e0f7fa';
        finalBalanceRow.innerHTML = `
            <td colspan="3" style="text-align: center;">Saldo Akhir Buku Kas Harian</td>
            <td colspan="2" style="text-align: center;">${currentOverallBalance.toLocaleString('id-ID')}</td>
        `;
    }

    function renderAccountDetailTable(accountCodeToFilter) {
        detailTransactionTableBody.innerHTML = '';

        if (!accountCodeToFilter) {
            const noSelectionRow = detailTransactionTableBody.insertRow();
            noSelectionRow.innerHTML = `<td colspan="4" style="text-align: center;">Pilih Kode Akun untuk melihat detail.</td>`;
            return;
        }

        // Filter out 'Saldo Awal' for detail view as it's not tied to a specific account code
        const filteredTransactions = transactions.filter((t) => t.accountCode === accountCodeToFilter && t.description !== 'Saldo Awal');

        if (filteredTransactions.length === 0) {
            const noDataRow = detailTransactionTableBody.insertRow();
            noDataRow.innerHTML = `<td colspan="4" style="text-align: center;">Tidak ada transaksi untuk akun ini.</td>`;
        } else {
            filteredTransactions.forEach((transaction) => {
                const row = detailTransactionTableBody.insertRow();
                row.innerHTML = `
                    <td>${transaction.date}</td>
                    <td>${transaction.description}</td>
                    <td>${transaction.penerimaan.toLocaleString('id-ID')}</td>
                    <td>${transaction.pengeluaran.toLocaleString('id-ID')}</td>
                `;
            });
        }
    }

    detailAccountSelector.addEventListener('change', (event) => {
        const selectedCode = event.target.value;
        renderAccountDetailTable(selectedCode);
    });

    function renderTables() {
        renderMainTable();
        generateRecapTable();
        loadAccountsToSelects(accountSearchInput.value, detailAccountSearchInput.value);

        const selectedView = viewSelector.value;
        mainTableSection.classList.add('hidden');
        recapTableSection.classList.add('hidden');
        accountDetailSection.classList.add('hidden');

        if (exportMainTableBtn) exportMainTableBtn.classList.add('hidden');
        if (exportRecapTableBtn) exportRecapTableBtn.classList.add('hidden');
        if (exportDetailTableBtn) exportDetailTableBtn.classList.add('hidden');

        if (selectedView === 'recap') {
            recapTableSection.classList.remove('hidden');
            if (exportRecapTableBtn) exportRecapTableBtn.classList.remove('hidden');
        } else if (selectedView === 'account_detail') {
            accountDetailSection.classList.remove('hidden');
            if (exportDetailTableBtn) exportDetailTableBtn.classList.remove('hidden');
            renderAccountDetailTable(detailAccountSelector.value);
        } else {
            mainTableSection.classList.remove('hidden');
            if (exportMainTableBtn) exportMainTableBtn.classList.remove('hidden');
        }
    }

    function clearForm() {
        descriptionInput.value = '';
        accountSearchInput.value = '';
        accountSelect.value = '';
        transactionTypeSelect.value = '';
        nominalInput.value = '';
        transactionDateInput.valueAsDate = new Date(); // Set default date to today
        accountSearchInput.value = '';
        detailAccountSearchInput.value = '';
        loadAccountsToSelects();
    }

    // MODIFIED downloadXLSXBtn EVENT LISTENER
    downloadXLSXBtn.addEventListener('click', () => {
        const wb = XLSX.utils.book_new();

        // --- Prepare Main Transaction Table Data (Riwayat Transaksi) ---
        let mainWsData = [];

        // Header baris 1: Tanggal, Uraian Transaksi, Kode Akun, Nama Akun, D/K, Nominal, Mutasi Bank Harian (merge 3 sel)
        mainWsData.push(["Tanggal", "Uraian Transaksi", "Kode Akun", "Nama Akun", "D/K", "Nominal", "Mutasi Bank Harian", "", ""]);
        // Header baris 2: Penerimaan, Pengeluaran, Saldo (di bawah Mutasi Bank Harian)
        mainWsData.push(["", "", "", "", "", "", "Penerimaan", "Pengeluaran", "Saldo"]);

        // Data Saldo Awal (ensure it's the first entry for accurate running balance in Excel)
        const initialBalanceEntry = transactions.find(t => t.description === 'Saldo Awal');
        if (initialBalanceEntry) {
            mainWsData.push([
                initialBalanceEntry.date,
                "Saldo Awal",
                "",
                "",
                "",
                "",
                initialBalanceEntry.penerimaan, // Use penerimaan for initial balance
                initialBalanceEntry.pengeluaran,
                initialStartingBalance
            ]);
        }


        let runningBalance = initialStartingBalance;

        // Data Transaksi (filter out the "Saldo Awal" entry as it's handled above)
        const transactionsForExcel = transactions.filter(t => t.description !== 'Saldo Awal');

        // Sort transactions for Excel export to ensure correct running balance calculation
        transactionsForExcel.sort((a, b) => new Date(a.date.split('/').reverse().join('-')) - new Date(b.date.split('/').reverse().join('-')));

        transactionsForExcel.forEach((transaction) => {
            if (transaction.type === 'D') {
                runningBalance += transaction.nominal;
            } else if (transaction.type === 'K') {
                runningBalance -= transaction.nominal;
            }
            mainWsData.push([
                transaction.date,
                transaction.description,
                transaction.accountCode,
                transaction.accountName,
                transaction.type,
                transaction.nominal,
                transaction.penerimaan,
                transaction.pengeluaran,
                runningBalance
            ]);
        });

        const wsMain = XLSX.utils.aoa_to_sheet(mainWsData);

        // Merge cells for "Mutasi Bank Harian" (cells G1, H1, I1 -> 0-indexed: c=6, r=0 to c=8, r=0)
        if (!wsMain['!merges']) wsMain['!merges'] = [];
        wsMain['!merges'].push({ s: { r: 0, c: 6 }, e: { r: 0, c: 8 } });

        // Add autofilter to all columns starting from the second header row (row index 1)
        wsMain['!autofilter'] = { ref: "A2:I" + (mainWsData.length) };

        // Optional: Set column widths for main sheet for better readability
        wsMain['!cols'] = [
            { wch: 15 }, // Tanggal
            { wch: 30 }, // Uraian Transaksi
            { wch: 15 }, // Kode Akun
            { wch: 20 }, // Nama Akun
            { wch: 5 },  // D/K
            { wch: 15 }, // Nominal
            { wch: 18 }, // Penerimaan
            { wch: 18 }, // Pengeluaran
            { wch: 18 }  // Saldo
        ];

        // Format angka sebagai angka di Excel, bukan teks (start from row 2 because of 2 header rows)
        for (let R = 2; R < mainWsData.length; ++R) {
            const nominalCell = XLSX.utils.encode_cell({ r: R, c: 5 });
            const penerimaanCell = XLSX.utils.encode_cell({ r: R, c: 6 });
            const pengeluaranCell = XLSX.utils.encode_cell({ r: R, c: 7 });
            const saldoCell = XLSX.utils.encode_cell({ r: R, c: 8 });

            if (wsMain[nominalCell]) wsMain[nominalCell].t = 'n';
            if (wsMain[penerimaanCell]) wsMain[penerimaanCell].t = 'n';
            if (wsMain[pengeluaranCell]) wsMain[pengeluaranCell].t = 'n';
            if (wsMain[saldoCell]) wsMain[saldoCell].t = 'n';
        }

        XLSX.utils.book_append_sheet(wb, wsMain, "Riwayat Transaksi");


        // --- Prepare Recap Table Data (Rekapitulasi Akun) ---
        let recapWsData = [];

        recapWsData.push(["Rub", "Kode Akun", "Nama Akun", "Penerimaan", "Pengeluaran"]);

        recapWsData.push([
            "",
            "",
            "Saldo Awal",
            initialStartingBalance,
            ""
        ]);

        const accountRecap = {};
        let totalOverallPenerimaan = 0;
        let totalOverallPengeluaran = 0;

        // Exclude 'Saldo Awal' from recap calculation
        const transactionsForRecapExcel = transactions.filter(t => t.description !== 'Saldo Awal');

        transactionsForRecapExcel.forEach((transaction) => {
            const key = `${transaction.accountCode}-${transaction.accountName}`;
            if (!accountRecap[key]) {
                accountRecap[key] = {
                    code: transaction.accountCode,
                    name: transaction.accountName,
                    totalPenerimaan: 0,
                    totalPengeluaran: 0,
                };
            }

            if (transaction.type === 'D') {
                accountRecap[key].totalPenerimaan += transaction.nominal;
                totalOverallPenerimaan += transaction.nominal;
            } else if (transaction.type === 'K') {
                accountRecap[key].totalPengeluaran += transaction.nominal;
                totalOverallPengeluaran += transaction.nominal;
            }
        });

        const sortedRecap = Object.values(accountRecap).sort((a, b) => {
            return a.code.localeCompare(b.code);
        });

        sortedRecap.forEach((recap) => {
            const rub = recap.code.substring(0, 2);
            recapWsData.push([
                rub,
                recap.code,
                recap.name,
                recap.totalPenerimaan,
                recap.totalPengeluaran
            ]);
        });

        recapWsData.push([
            "Jumlah Keseluruhan", "", "",
            totalOverallPenerimaan,
            totalOverallPengeluaran
        ]);

        let currentOverallBalance = initialStartingBalance;
        transactionsForRecapExcel.forEach((transaction) => { // Use the filtered list for final balance
            if (transaction.type === 'D') {
                currentOverallBalance += transaction.nominal;
            } else if (transaction.type === 'K') {
                currentOverallBalance -= transaction.nominal;
            }
        });
        recapWsData.push([
            "Saldo Akhir Buku Kas Harian", "", "",
            currentOverallBalance,
            ""
        ]);

        const wsRecap = XLSX.utils.aoa_to_sheet(recapWsData);

        let totalRowRecapIndex = -1;
        for (let i = 0; i < recapWsData.length; i++) {
            if (recapWsData[i][0] === "Jumlah Keseluruhan") {
                totalRowRecapIndex = i;
                break;
            }
        }

        if (!wsRecap['!merges']) wsRecap['!merges'] = [];
        if (totalRowRecapIndex !== -1) {
            wsRecap['!merges'].push({ s: { r: totalRowRecapIndex, c: 0 }, e: { r: totalRowRecapIndex, c: 2 } });
        }

        let finalBalanceRowRecapIndex = -1;
        for (let i = 0; i < recapWsData.length; i++) {
            if (recapWsData[i][0] === "Saldo Akhir Buku Kas Harian") {
                finalBalanceRowRecapIndex = i;
                break;
            }
        }

        if (finalBalanceRowRecapIndex !== -1) {
            wsRecap['!merges'].push({ s: { r: finalBalanceRowRecapIndex, c: 0 }, e: { r: finalBalanceRowRecapIndex, c: 2 } });
            wsRecap['!merges'].push({ s: { r: finalBalanceRowRecapIndex, c: 3 }, e: { r: finalBalanceRowRecapIndex, c: 4 } });
        }

        for (let R = 1; R < recapWsData.length; ++R) {
            const penerimaanCell = XLSX.utils.encode_cell({ r: R, c: 3 });
            const pengeluaranCell = XLSX.utils.encode_cell({ r: R, c: 4 });

            if (wsRecap[penerimaanCell] && typeof recapWsData[R][3] === 'number') {
                wsRecap[penerimaanCell].t = 'n';
            }
            if (wsRecap[pengeluaranCell] && typeof recapWsData[R][4] === 'number') {
                wsRecap[pengeluaranCell].t = 'n';
            }
        }

        wsRecap['!cols'] = [
            { wch: 8 },
            { wch: 15 },
            { wch: 30 },
            { wch: 20 },
            { wch: 20 }
        ];

        XLSX.utils.book_append_sheet(wb, wsRecap, "Rekapitulasi Akun");

        if (wb.SheetNames.length === 0) {
            alert('Tidak ada data transaksi atau rekap untuk diunduh.');
            return;
        }

        const fileName = `Laporan_Mutasi_Kas_Kecil_${new Date().toLocaleDateString('id-ID').replace(/\//g, '-')}.xlsx`;
        XLSX.writeFile(wb, fileName);
        alert('Laporan berhasil diunduh dalam format XLSX.');
    });

    // --- Panggil Fungsi Saat Halaman Dimuat ---
    loadFromLocalStorage();
    loadAccountsToSelects();
    clearForm(); // Set default date on load
    renderTables();
});
