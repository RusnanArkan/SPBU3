document.addEventListener('DOMContentLoaded', () => {
    const jenisBBMSelect = document.getElementById('jenisBBM');
    const hargaJualSekarangInput = document.getElementById('hargaJualSekarang');
    const literTerjualInput = document.getElementById('literTerjual');
    const tanggalInput = document.getElementById('tanggalInput');
    const hitungBtn = document.getElementById('hitungBtn');
    const exportExcelBtn = document.getElementById('exportExcelBtn');
    const clearAllDataBtn = document.getElementById('clearAllDataBtn');
    const separatedBBMReportsDiv = document.getElementById('separatedBBMReports');

    let riwayatPenjualan = JSON.parse(localStorage.getItem('riwayatPenjualan')) || [];
    let lastHargaPerJenis = JSON.parse(localStorage.getItem('lastHargaPerJenis')) || {};

    const today = new Date();
    const year = today.getFullYear();
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const day = String(today.getDate()).padStart(2, '0');
    tanggalInput.value = `${year}-${month}-${day}`;

    renderSeparatedBBMReports();
    updateHargaJualSekarangBasedOnSelection();

    jenisBBMSelect.addEventListener('change', updateHargaJualSekarangBasedOnSelection);

    function updateHargaJualSekarangBasedOnSelection() {
        const selectedJenisBBM = jenisBBMSelect.value;
        if (lastHargaPerJenis[selectedJenisBBM]) {
            hargaJualSekarangInput.value = lastHargaPerJenis[selectedJenisBBM];
        } else {
            hargaJualSekarangInput.value = '';
        }
    }

    hitungBtn.addEventListener('click', () => {
        const jenisBBM = jenisBBMSelect.value;
        const hargaJualSekarang = parseFloat(hargaJualSekarangInput.value);
        const literTerjual = parseFloat(literTerjualInput.value);
        const tanggalManual = tanggalInput.value;

        if (isNaN(hargaJualSekarang) || isNaN(literTerjual) || hargaJualSekarang <= 0 || literTerjual <= 0) {
            alert('Mohon masukkan angka yang valid dan lebih besar dari nol untuk harga jual dan liter terjual.');
            return;
        }

        const nilaiJual = hargaJualSekarang * literTerjual;

        let displayDate;
        if (tanggalManual) {
            const [yearManual, monthManual, dayManual] = tanggalManual.split('-');
            displayDate = `${dayManual}/${monthManual}/${yearManual}`;
        } else {
            const now = new Date();
            displayDate = now.toLocaleString('id-ID', {
                year: 'numeric',
                month: '2-digit',
                day: '2-digit'
            });
        }

        const newEntry = {
            waktu: displayDate,
            jenisBBM: jenisBBM,
            hargaJual: hargaJualSekarang,
            literTerjual: literTerjual,
            nilaiJual: nilaiJual
        };

        riwayatPenjualan.push(newEntry);
        localStorage.setItem('riwayatPenjualan', JSON.stringify(riwayatPenjualan));

        lastHargaPerJenis[jenisBBM] = hargaJualSekarang;
        localStorage.setItem('lastHargaPerJenis', JSON.stringify(lastHargaPerJenis));

        renderSeparatedBBMReports();

        literTerjualInput.value = '';
    });

    exportExcelBtn.addEventListener('click', () => {
        if (typeof XLSX === 'undefined') {
            alert('Pustaka ekspor Excel belum dimuat. Mohon refresh halaman atau periksa koneksi internet Anda.');
            console.error('XLSX library is not loaded.');
            return;
        }

        if (riwayatPenjualan.length === 0) {
            alert('Tidak ada data untuk diekspor.');
            return;
        }

        const uniqueJenisBBM = [...new Set(riwayatPenjualan.map(item => item.jenisBBM))].sort();

        const wb = XLSX.utils.book_new();

        uniqueJenisBBM.forEach(jenisBBM => {
            // **Perubahan di sini:** Tambahkan kembali "Jenis BBM" ke header Excel
            const dataForSheet = [
                ["No.", "Waktu", "Jenis BBM", "Harga Jual (per Liter)", "Liter Terjual", "Nilai Jual"]
            ];

            const filteredData = riwayatPenjualan
                .filter(item => item.jenisBBM === jenisBBM)
                .sort((a, b) => {
                    const [dayA, monthA, yearA] = a.waktu.split('/');
                    const dateA = new Date(`${monthA}/${dayA}/${yearA}`);
                    const [dayB, monthB, yearB] = b.waktu.split('/');
                    const dateB = new Date(`${monthB}/${dayB}/${yearB}`);
                    return dateA - dateB;
                });

            let currentGroupNumber = 0;
            let lastDateForNumbering = null;
            let lastDateForDisplay = null;
            let lastNumberForDisplay = null;

            let subtotalLiter = 0;
            let subtotalNilai = 0;

            filteredData.forEach(item => {
                if (item.waktu !== lastDateForNumbering) {
                    currentGroupNumber++;
                }

                let displayWaktu = "";
                let displayGroupNumber = "";

                if (item.waktu !== lastDateForDisplay) {
                    displayWaktu = item.waktu;
                }

                if (currentGroupNumber !== lastNumberForDisplay) {
                    displayGroupNumber = currentGroupNumber;
                }

                // **Perubahan di sini:** Tambahkan item.jenisBBM ke dataForSheet
                dataForSheet.push([
                    displayGroupNumber,
                    displayWaktu,
                    item.jenisBBM, // Menampilkan Jenis BBM di kolom
                    item.hargaJual,
                    item.literTerjual,
                    item.nilaiJual
                ]);

                subtotalLiter += item.literTerjual;
                subtotalNilai += item.nilaiJual;

                lastDateForNumbering = item.waktu;
                lastDateForDisplay = item.waktu;
                lastNumberForDisplay = currentGroupNumber;
            });

            // Tambahkan baris subtotal untuk jenis BBM ini
            // Sesuaikan colSpan untuk subtotal karena kolom "Jenis BBM" kembali
            dataForSheet.push(["", "", `Subtotal ${jenisBBM}`, "", subtotalLiter, subtotalNilai]);

            const ws = XLSX.utils.aoa_to_sheet(dataForSheet);
            XLSX.utils.book_append_sheet(wb, ws, jenisBBM);
        });

        const summaryData = [
            ["Ringkasan Total Penjualan Keseluruhan", "", "", "", ""],
            ["Jenis BBM", "Total Liter Terjual", "Total Nilai Jual", "", ""]
        ];
        const overallSummary = calculateSummary(riwayatPenjualan);
        const grandTotal = calculateGrandTotal(riwayatPenjualan);

        for (const jenis in overallSummary) {
            summaryData.push([
                jenis,
                overallSummary[jenis].totalLiter,
                overallSummary[jenis].totalNilai,
                "", ""
            ]);
        }
        summaryData.push(["", "", "", "", ""]);
        summaryData.push(["TOTAL KESELURUHAN NILAI JUAL", "", grandTotal, "", ""]);

        const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
        XLSX.utils.book_append_sheet(wb, wsSummary, "Ringkasan Total");

        XLSX.writeFile(wb, `Buku_Jual_Kontan_Per_BBM_${new Date().toISOString().slice(0, 10)}.xlsx`);
    });

    clearAllDataBtn.addEventListener('click', () => {
        if (confirm('Apakah Anda yakin ingin menghapus semua data penjualan? Tindakan ini tidak dapat dibatalkan.')) {
            localStorage.removeItem('riwayatPenjualan');
            localStorage.removeItem('lastHargaPerJenis');
            riwayatPenjualan = [];
            lastHargaPerJenis = {};
            renderSeparatedBBMReports();
            updateHargaJualSekarangBasedOnSelection();
            tanggalInput.value = `${year}-${month}-${day}`;
            alert('Semua data penjualan telah dihapus.');
        }
    });

    function calculateSummary(data) {
        const summary = {};
        data.forEach(entry => {
            if (!summary[entry.jenisBBM]) {
                summary[entry.jenisBBM] = { totalLiter: 0, totalNilai: 0 };
            }
            summary[entry.jenisBBM].totalLiter += entry.literTerjual;
            summary[entry.jenisBBM].totalNilai += entry.nilaiJual;
        });
        return summary;
    }

    function calculateGrandTotal(data) {
        return data.reduce((total, entry) => total + entry.nilaiJual, 0);
    }

    function renderSeparatedBBMReports() {
        separatedBBMReportsDiv.innerHTML = '';

        if (riwayatPenjualan.length === 0) {
            separatedBBMReportsDiv.innerHTML = '<p>Tidak ada data penjualan untuk ditampilkan.</p>';
            return;
        }

        const uniqueJenisBBM = [...new Set(riwayatPenjualan.map(item => item.jenisBBM))].sort();

        uniqueJenisBBM.forEach(jenisBBM => {
            const tableContainer = document.createElement('div');
            tableContainer.className = 'bbm-report-section mb-4';

            const sectionTitle = document.createElement('h4');
            sectionTitle.textContent = `Riwayat Penjualan ${jenisBBM}`;
            tableContainer.appendChild(sectionTitle);

            const table = document.createElement('table');
            table.className = 'table table-bordered table-striped';
            tableContainer.appendChild(table);

            const thead = table.createTHead();
            const headerRow = thead.insertRow();
            // Header untuk UI, tanpa "Jenis BBM" karena sudah dipisahkan per tabel
            ['No.', 'Waktu', 'Harga Jual (per Liter)', 'Liter Terjual', 'Nilai Jual'].forEach(text => {
                const th = document.createElement('th');
                th.textContent = text;
                headerRow.appendChild(th);
            });

            const tbody = table.createTBody();

            const filteredData = riwayatPenjualan
                .filter(item => item.jenisBBM === jenisBBM)
                .sort((a, b) => {
                    const [dayA, monthA, yearA] = a.waktu.split('/');
                    const dateA = new Date(`${monthA}/${dayA}/${yearA}`);
                    const [dayB, monthB, yearB] = b.waktu.split('/');
                    const dateB = new Date(`${monthB}/${dayB}/${yearB}`);
                    return dateA - dateB;
                });

            let currentGroupNumberUI = 0;
            let lastDateForNumberingUI = null;
            let lastDateForDisplayUI = null;
            let lastNumberForDisplayUI = null;

            let subtotalLiterUI = 0;
            let subtotalNilaiUI = 0;

            filteredData.forEach(entry => {
                if (entry.waktu !== lastDateForNumberingUI) {
                    currentGroupNumberUI++;
                }

                const row = tbody.insertRow();
                let displayWaktu = "";
                let displayGroupNumber = "";

                if (entry.waktu !== lastDateForDisplayUI) {
                    displayWaktu = entry.waktu;
                }

                if (currentGroupNumberUI !== lastNumberForDisplayUI) {
                    displayGroupNumber = currentGroupNumberUI;
                }

                row.insertCell().textContent = displayGroupNumber;
                row.insertCell().textContent = displayWaktu;
                row.insertCell().textContent = `Rp ${entry.hargaJual.toLocaleString('id-ID')}`;
                row.insertCell().textContent = `${entry.literTerjual.toLocaleString('id-ID')} L`;
                row.insertCell().textContent = `Rp ${entry.nilaiJual.toLocaleString('id-ID')}`;

                subtotalLiterUI += entry.literTerjual;
                subtotalNilaiUI += entry.nilaiJual;

                lastDateForNumberingUI = entry.waktu;
                lastDateForDisplayUI = entry.waktu;
                lastNumberForDisplayUI = currentGroupNumberUI;
            });

            const subtotalRow = tbody.insertRow();
            subtotalRow.style.backgroundColor = '#e0f7fa';
            const subtotalLabelCell = subtotalRow.insertCell();
            subtotalLabelCell.colSpan = 3; // Colspan 3 karena tanpa "Jenis BBM" di UI
            subtotalLabelCell.textContent = `Subtotal ${jenisBBM}`;
            subtotalLabelCell.style.textAlign = 'right';
            subtotalLabelCell.style.fontWeight = 'bold';
            subtotalRow.insertCell().textContent = `${subtotalLiterUI.toLocaleString('id-ID')} L`;
            subtotalRow.insertCell().textContent = `Rp ${subtotalNilaiUI.toLocaleString('id-ID')}`;

            separatedBBMReportsDiv.appendChild(tableContainer);
        });

        const summary = calculateSummary(riwayatPenjualan);
        const grandTotal = calculateGrandTotal(riwayatPenjualan);

        if (riwayatPenjualan.length > 0) {
            const summaryContainer = document.createElement('div');
            summaryContainer.className = 'overall-summary mt-5 p-3 border rounded bg-light';

            const summaryHeader = document.createElement('h5');
            summaryHeader.textContent = 'Ringkasan Total Penjualan Keseluruhan';
            summaryContainer.appendChild(summaryHeader);

            const summaryList = document.createElement('ul');
            summaryList.className = 'list-unstyled';
            for (const jenis in summary) {
                const li = document.createElement('li');
                li.innerHTML = `<strong>${jenis}:</strong> ${summary[jenis].totalLiter.toLocaleString('id-ID')} L / Rp ${summary[jenis].totalNilai.toLocaleString('id-ID')}`;
                summaryList.appendChild(li);
            }
            summaryContainer.appendChild(summaryList);

            const grandTotalParagraph = document.createElement('p');
            grandTotalParagraph.innerHTML = `<strong>TOTAL KESELURUHAN NILAI JUAL: Rp ${grandTotal.toLocaleString('id-ID')}</strong>`;
            grandTotalParagraph.style.fontSize = '1.2em';
            grandTotalParagraph.style.marginTop = '15px';
            summaryContainer.appendChild(grandTotalParagraph);

            separatedBBMReportsDiv.appendChild(summaryContainer);
        }
    }

    const script = document.createElement('script');
    script.src = 'https://unpkg.com/xlsx/dist/xlsx.full.min.js';
    script.onload = () => console.log('XLSX library loaded successfully!');
    script.onerror = () => {
        console.error('Failed to load XLSX library. Check your internet connection or the CDN link.');
        alert('Gagal memuat fungsi ekspor Excel. Mohon periksa koneksi internet Anda.');
    };
    document.head.appendChild(script);
});
