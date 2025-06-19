function analyzeSurveyData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheetName = 'Sheet1'; // Pastikan nama sheet ini sudah benar
  const sourceSheet = ss.getSheetByName(sourceSheetName);

  if (!sourceSheet) {
    Browser.msgBox('Error', `Sheet "${sourceSheetName}" tidak ditemukan. Pastikan nama sheet sudah benar.`, Browser.Buttons.OK);
    return;
  }

  const dataRange = sourceSheet.getDataRange();
  const data = dataRange.getValues(); // Mengambil semua data dari sheet sumber
  const headerRow = data[0]; // Baris pertama adalah header (F1, G1, H1, dst.)

  // Membuat sheet baru atau menghapus yang lama jika sudah ada
  let newSheet = ss.getSheetByName('Hasil Analisis Survey');
  if (newSheet) {
    ss.deleteSheet(newSheet);
  }
  newSheet = ss.insertSheet('Hasil Analisis Survey');

  // --- Tabel Pertama: Rekapitulasi Jumlah Responden ---
  const headers1 = ['Aspek', 'Indikator', 'Tidak Puas', 'Kurang Puas', 'Cukup Puas', 'Puas', 'Sangat Puas'];
  newSheet.getRange(1, 1, 1, headers1.length).setValues([headers1]).setFontWeight('bold');

  // MODIFIKASI BERIKUTNYA ADA DI BAWAH INI
  const aspects = {
    'Tangibles': { startCol: 'D', endCol: 'E' }, // Berubah dari F-T menjadi D-E
    'Reliability': { startCol: 'F', endCol: 'F' }, // Berubah dari U-AE menjadi F saja
    'Responsiveness': { startCol: 'G', endCol: 'G' }, // Berubah dari AF-AO menjadi G saja
    'Assurance': { startCol: 'H', endCol: 'J' }, // Berubah dari AP-AS menjadi H-J
    'Empathy': { startCol: 'K', endCol: 'L' } // Berubah dari AT-AX menjadi K-L, aspek Transparency dihapus
  };
  // AKHIR MODIFIKASI

  const sentimentOptions = ['Tidak Puas', 'Kurang Puas', 'Cukup Puas', 'Puas', 'Sangat Puas'];
  let currentRow = 2; // Baris awal untuk data

  for (const aspectName in aspects) {
    const { startCol, endCol } = aspects[aspectName];
    const startColIndex = letterToColumn(startCol);
    const endColIndex = letterToColumn(endCol);

    // Menentukan baris awal untuk aspek saat ini
    const startRowForAspect = currentRow; 
    let indicatorCountForAspect = 0; // Menghitung berapa banyak indikator dalam aspek ini

    for (let colIndex = startColIndex; colIndex <= endColIndex; colIndex++) {
      const indicator = headerRow[colIndex - 1]; 
      const rowData = [aspectName, indicator]; // Aspek dan Indikator

      sentimentOptions.forEach(sentiment => {
        let count = 0;
        for (let i = 1; i < data.length; i++) {
          const cellValue = String(data[i][colIndex - 1]).trim();
          if (cellValue.toLowerCase() === sentiment.toLowerCase()) {
            count++;
          }
        }
        rowData.push(count);
      });
      newSheet.getRange(currentRow, 1, 1, rowData.length).setValues([rowData]);
      currentRow++;
      indicatorCountForAspect++;
    }

    // Setelah semua indikator untuk aspek ini ditulis, merge sel di kolom "Aspek"
    if (indicatorCountForAspect > 1) {
      newSheet.getRange(startRowForAspect, 1, indicatorCountForAspect, 1).merge();
      // Opsional: Atur perataan vertikal ke tengah
      newSheet.getRange(startRowForAspect, 1, indicatorCountForAspect, 1).setVerticalAlignment('middle');
    }
  }

  // --- Tabel Kedua: Persentase Responden per Aspek ---
  currentRow++; // Jeda satu baris
  const headers2 = ['Aspek', 'Tidak Puas (%)', 'Kurang Puas (%)', 'Cukup Puas (%)', 'Puas (%)', 'Sangat Puas (%)'];
  newSheet.getRange(currentRow, 1, 1, headers2.length).setValues([headers2]).setFontWeight('bold');
  currentRow++;

  for (const aspectName in aspects) {
    const { startCol, endCol } = aspects[aspectName];
    const startColIndex = letterToColumn(startCol);
    const endColIndex = letterToColumn(endCol);

    const aspectRowData = [aspectName];
    const aspectCounts = Array(sentimentOptions.length).fill(0);
    let totalResponsesInAspect = 0;

    for (let colIndex = startColIndex; colIndex <= endColIndex; colIndex++) {
      for (let i = 1; i < data.length; i++) {
        const cellValue = String(data[i][colIndex - 1]).trim();
        const sentimentIndex = sentimentOptions.findIndex(s => s.toLowerCase() === cellValue.toLowerCase());

        if (sentimentIndex !== -1) {
          aspectCounts[sentimentIndex]++;
          totalResponsesInAspect++;
        }
      }
    }

    sentimentOptions.forEach((_, index) => {
      const percentage = totalResponsesInAspect > 0 ? (aspectCounts[index] / totalResponsesInAspect) * 100 : 0;
      aspectRowData.push(percentage.toFixed(2) + '%');
    });
    newSheet.getRange(currentRow, 1, 1, aspectRowData.length).setValues([aspectRowData]);
    currentRow++;
  }

  // Auto-fit kolom agar lebih rapi
  newSheet.autoResizeColumns(1, newSheet.getLastColumn());
  Browser.msgBox('Sukses', 'Analisis data survey telah selesai dan disimpan di sheet "Hasil Analisis Survey".', Browser.Buttons.OK);
}

// Fungsi pembantu untuk mengkonversi huruf kolom ke indeks angka (A=1, B=2, dst)
function letterToColumn(letter) {
  let column = 0;
  const length = letter.length;
  for (let i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}

// Fungsi pembantu untuk mengkonversi indeks angka kolom ke huruf (1=A, 2=B, dst)
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}