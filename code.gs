// 1D Cutting Stock Optimizer - Google Apps Script

// Fungsi untuk menampilkan HTML di sidebar
function showOptimizer() {
  const html = HtmlService.getUrl() + '/usercodeappsscript?url=' + encodeURIComponent(ScriptApp.getService().getUrl());
  const ui = HtmlService.createHtmlOutputFromFile('index');
  ui.setWidth(800).setHeight(900);
  SpreadsheetApp.getUi().showModelessDialog(ui, '1D Cutting Optimizer');
}

// Fungsi untuk mendapatkan HTML
function getHtml() {
  return HtmlService.createHtmlOutputFromFile('index');
}

// Fungsi optimasi cutting stock menggunakan First Fit Decreasing (FFD)
function optimizeCutting(materialLength, kerfThickness, cutLists) {
  try {
    // Validasi input
    if (!materialLength || materialLength <= 0) {
      return { error: 'Panjang material harus lebih dari 0' };
    }
    if (!kerfThickness || kerfThickness < 0) {
      return { error: 'Tebal gergaji tidak valid' };
    }
    if (!cutLists || cutLists.length === 0) {
      return { error: 'Daftar cut list tidak boleh kosong' };
    }

    // Parse cut lists dan buat array item dengan quantity
    const items = [];
    for (let i = 0; i < cutLists.length; i++) {
      const cut = cutLists[i];
      if (cut.length && cut.length > 0 && cut.quantity && cut.quantity > 0) {
        for (let q = 0; q < cut.quantity; q++) {
          items.push({
            id: cut.id,
            length: cut.length,
            quantity: 1,
            originalQuantity: cut.quantity
          });
        }
      }
    }

    if (items.length === 0) {
      return { error: 'Tidak ada item yang valid dalam cut list' };
    }

    // Sort items by length (descending) - First Fit Decreasing
    items.sort((a, b) => b.length - a.length);

    // Optimasi pemotongan
    const batches = [];
    const usedItems = {};

    for (const item of items) {
      if (usedItems[item.id + '_' + item.quantity]) {
        continue;
      }

      let placed = false;

      // Coba tempatkan di batch yang sudah ada
      for (const batch of batches) {
        if (canFitInBatch(batch, item, materialLength, kerfThickness)) {
          batch.items.push(item);
          batch.usedLength += item.length + kerfThickness;
          placed = true;
          break;
        }
      }

      // Jika tidak bisa, buat batch baru
      if (!placed) {
        batches.push({
          items: [item],
          usedLength: item.length
        });
      }

      usedItems[item.id + '_' + item.quantity] = true;
    }

    // Hitung waste dan hasil
    const totalBatches = batches.length;
    const totalMaterialUsed = totalBatches * materialLength;
    const totalCutLength = items.reduce((sum, item) => sum + item.length, 0);
    const totalKerfLength = (totalCutLength > 0) ? (items.length - 1) * kerfThickness : 0;
    const totalWaste = totalMaterialUsed - totalCutLength - totalKerfLength;
    const wastePercentage = (totalWaste / totalMaterialUsed * 100).toFixed(2);

    // Format diagram
    const diagrams = batches.map((batch, batchIndex) => {
      const diagram = formatDiagram(batch, materialLength, kerfThickness, batchIndex + 1);
      return diagram;
    });

    return {
      success: true,
      totalBatches: totalBatches,
      wastePercentage: parseFloat(wastePercentage),
      totalMaterialUsed: totalMaterialUsed,
      totalCutLength: totalCutLength,
      totalKerfLength: totalKerfLength,
      totalWaste: totalWaste,
      diagrams: diagrams,
      batches: batches
    };

  } catch (error) {
    return { error: 'Error: ' + error.message };
  }
}

// Fungsi untuk cek apakah item bisa masuk ke dalam batch
function canFitInBatch(batch, item, materialLength, kerfThickness) {
  const currentUsedLength = batch.usedLength;
  const requiredLength = item.length + kerfThickness; // Tambah kerf untuk potongan berikutnya

  return (currentUsedLength + requiredLength) <= materialLength;
}

// Fungsi untuk format diagram potongan
function formatDiagram(batch, materialLength, kerfThickness, batchNumber) {
  let diagram = `BATANG ${batchNumber} (Panjang Total: ${materialLength}mm)\n`;
  diagram += '═'.repeat(50) + '\n';

  let currentPosition = 0;
  const items = batch.items;

  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    const itemLength = item.length;

    // Tampilkan potongan
    diagram += `${item.id.toUpperCase()}: ${itemLength}mm (posisi: ${currentPosition}mm - ${currentPosition + itemLength}mm)\n`;
    currentPosition += itemLength;

    // Tampilkan kerf jika bukan potongan terakhir
    if (i < items.length - 1) {
      diagram += `[KERF: ${kerfThickness}mm]\n`;
      currentPosition += kerfThickness;
    }
  }

  // Sisa material (waste)
  const waste = materialLength - currentPosition;
  diagram += `\nSISA: ${waste}mm\n`;
  diagram += '═'.repeat(50) + '\n\n';

  return diagram;
}

// Fungsi untuk generate laporan lengkap
function generateReport(materialLength, kerfThickness, cutLists) {
  const result = optimizeCutting(materialLength, kerfThickness, cutLists);
  
  if (result.error) {
    return { error: result.error };
  }

  let report = '';
  report += 'LAPORAN OPTIMASI CUTTING STOCK 1D\n';
  report += '═'.repeat(60) + '\n\n';

  report += 'PARAMETER INPUT:\n';
  report += `- Panjang Material: ${materialLength}mm\n`;
  report += `- Tebal Gergaji (Kerf): ${kerfThickness}mm\n`;
  report += `- Jumlah Tipe Potongan: ${cutLists.length}\n\n`;

  report += 'HASIL OPTIMASI:\n';
  report += `- Total Batang Dibutuhkan: ${result.totalBatches}\n`;
  report += `- Total Material Digunakan: ${result.totalMaterialUsed}mm\n`;
  report += `- Total Panjang Potongan: ${result.totalCutLength}mm\n`;
  report += `- Total Panjang Kerf: ${result.totalKerfLength}mm\n`;
  report += `- Total Waste: ${result.totalWaste}mm\n`;
  report += `- Persentase Waste: ${result.wastePercentage}%\n`;
  report += `- Efisiensi: ${(100 - result.wastePercentage).toFixed(2)}%\n\n`;

  report += 'DIAGRAM POTONGAN:\n';
  report += '═'.repeat(60) + '\n\n';
  report += result.diagrams.join('\n');

  return { success: true, report: report, data: result };
}
