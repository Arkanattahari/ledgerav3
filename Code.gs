// Konstanta Global untuk Nama Sheet
const PRODUCT_SHEET_NAME = 'Products';
const CATEGORY_SHEET_NAME = 'Categories';
const BASE_UNIT_SHEET_NAME = 'BaseUnits';
const UNIT_SHEET_NAME = 'Units';
const WAREHOUSE_SHEET_NAME = 'Warehouses';
const SUPPLIER_SHEET_NAME = 'Suppliers';
const STOCK_INVENTORY_SHEET_NAME = 'StockInventory'
const STOCK_LEDGER_SHEET_NAME = 'StockLedger';
const PURCHASES_SHEET_NAME = 'Purchases';
const PURCHASE_ITEMS_SHEET_NAME = 'PurchaseItems';

// ==========================================================================
// FUNGSI UTAMA WEB APP (doGet & include)
// ==========================================================================
function doGet(e) {
  try {
    const template = HtmlService.createTemplateFromFile('index');
    const htmlOutput = template.evaluate();
    htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).setTitle('Ledgera App');
    return htmlOutput;
  } catch (error) {
    Logger.log("LOG_ERROR: Critical error in doGet: " + error.toString());
    return HtmlService.createHtmlOutput("<h3>Terjadi kesalahan kritis saat memuat aplikasi.</h3>");
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==========================================================================
// FUNGSI HELPER UTAMA
// ==========================================================================
/**
 * Helper generik baru yang terbukti bekerja.
 * @param {Array<Array<any>>} data - Data mentah dari sheet (termasuk header).
 * @param {Object} fieldMap - Peta untuk mengubah nama header ke properti JS.
 * @returns {Array<Object>} Array objek yang sudah diproses.
 */
function _processSheetData_(data, fieldMap) {
  if (!data || data.length < 2) return [];
  const headers = data[0].map(h => String(h).toLowerCase().replace(/ /g, ''));
  const colIndexes = {};
  for (const key in fieldMap) {
    const colName = fieldMap[key];
    const colIndex = headers.indexOf(colName.toLowerCase().replace(/ /g, ''));
    if (colIndex === -1) {
      Logger.log(`WARNING: _processSheetData_ - Kolom '${colName}' tidak ditemukan.`);
    }
    colIndexes[key] = colIndex;
  }
  return data.slice(1).map(row => {
    const item = {};
    for (const key in colIndexes) {
      const index = colIndexes[key];
      let value = (index !== -1 && row[index] !== undefined) ? row[index] : null;
      if (typeof value === 'string') {
        value = manualEscapeHtml(value);
      }
      // Khusus untuk kolom tanggal, konversi ke ISO string agar tidak error
      if (value instanceof Date) {
        value = value.toISOString();
      }
      item[key] = value;
    }
    return item;
  }).filter(item => item.id && String(item.id).trim() !== '');
}

/**
 * Mengambil data dari beberapa sheet sekaligus secara efisien.
 * @param {string[]} sheetNames Array berisi nama-nama sheet yang akan diambil datanya.
 * @returns {Object} Objek dengan key nama sheet dan value data array 2D dari sheet tersebut.
 */
function getDataFromSheets_(sheetNames) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  const sheetData = {};
  const requestedSheets = allSheets.filter(sheet => sheetNames.includes(sheet.getName()));
  requestedSheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const range = sheet.getDataRange();
    sheetData[sheetName] = (range.getNumRows() > 0) ? range.getValues() : [];
  });
  return sheetData;
}

// ==========================================================================
// FUNGSI UTILITAS
// ==========================================================================
function manualEscapeHtml(text) {
  if (typeof text !== 'string') {
    return text; // Kembalikan apa adanya jika bukan string
  }
  return text
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;"); // atau &apos;
}

/**
 * Membuat objek map dari data sheet untuk lookup nama berdasarkan ID.
 * @param {Array<Array<string>>} dataValues Data dari getValues(). Baris pertama harus header.
 * @param {string} idHeaderName Nama header kolom ID (setelah dinormalisasi).
 * @param {string} nameHeaderName Nama header kolom Nama (setelah dinormalisasi).
 * @returns {Object} Objek map { id: nama }.
 */
function createNameMap_(dataValues, idHeaderName, nameHeaderName) {
  const nameMap = {};
  if (!dataValues || dataValues.length <= 1) {
    return nameMap;
  }
  
  const headers = dataValues[0].map(h => String(h).toLowerCase().replace(/ /g, ''));
  const idIndex = headers.indexOf(idHeaderName);
  const nameIndex = headers.indexOf(nameHeaderName);

  if (idIndex === -1 || nameIndex === -1) {
    Logger.log(`LOG_WARNING: createNameMap_ - Could not find headers '${idHeaderName}' or '${nameHeaderName}'.`);
    return nameMap;
  }

  for (let i = 1; i < dataValues.length; i++) {
    const row = dataValues[i];
    const id = row[idIndex];
    const name = row[nameIndex];
    if (id && String(id).trim() !== '') {
      nameMap[String(id).trim()] = String(name).trim();
    }
  }
  return nameMap;
}

/**
 * Mendapatkan objek Sheet dan menangani kasus jika tidak ditemukan.
 * @param {string} sheetName Nama sheet yang dicari.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null} Objek sheet atau null jika tidak ada.
 */
function getSheet_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`LOG_ERROR: getSheet_ - Sheet "${sheetName}" tidak ditemukan.`);
  }
  return sheet;
}

/**
 * FUNGSI BARU UNTUK NOMOR REFERENSI TRANSAKSI
 * Menghasilkan nomor referensi dengan format PREFIX+YYMM+NNN yang direset setiap bulan.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Objek sheet transaksi.
 * @param {string} prefix Awalan untuk referensi (misal: "PU", "SA").
 * @param {string} refColumnName Nama header kolom referensi (setelah normalisasi).
 * @param {number} [padLength=3] Jumlah digit untuk angka urut.
 * @returns {string} Nomor referensi baru yang sudah diformat.
 */
function generateReferenceNumber_(sheet, prefix, refColumnName, padLength = 3) {
  const now = new Date();
  const year = now.getFullYear().toString().slice(-2); // '25'
  const month = (now.getMonth() + 1).toString().padStart(2, '0'); // '06'
  const currentYYMM = year + month; // '2506'
  const fullPrefix = prefix + currentYYMM; // 'PU2506'

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return fullPrefix + "1".padStart(padLength, '0'); // Jika sheet kosong, mulai dari 1
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(h => String(h).toLowerCase().replace(/ /g, ''));
  const refColIndex = headers.indexOf(refColumnName);

  if (refColIndex === -1) {
    throw new Error(`Kolom Referensi '${refColumnName}' tidak ditemukan di sheet.`);
  }

  const allReferences = sheet.getRange(2, refColIndex + 1, lastRow - 1, 1).getValues().flat();
  
  const numbersThisMonth = allReferences
    .map(ref => {
      const refString = String(ref);
      if (refString.startsWith(fullPrefix)) {
        const seqStr = refString.substring(fullPrefix.length);
        return parseInt(seqStr, 10);
      }
      return null;
    })
    .filter(num => num !== null && !isNaN(num));

  const lastNumber = numbersThisMonth.length > 0 ? Math.max(0, ...numbersThisMonth) : 0;
  const nextNumber = lastNumber + 1;

  return fullPrefix + String(nextNumber).padStart(padLength, '0');
}

/**
 * Menghasilkan ID berikutnya untuk baris baru.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Objek sheet.
 * @param {string} idColumnName Nama header kolom ID (setelah normalisasi).
 * @param {string} prefix Awalan untuk ID (misal: "CAT", "PD").
 * @param {number} [padLength=3] Jumlah digit untuk angka ID. Default adalah 3.
 * @returns {string} ID baru yang sudah diformat.
 */
function generateNextId_(sheet, idColumnName, prefix, padLength = 3) {
  const lastRow = sheet.getLastRow();
  const defaultId = prefix + "1".padStart(padLength, '0');
  
  if (lastRow < 2) {
    return defaultId;
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(h => String(h).toLowerCase().replace(/ /g, ''));
  const idColIndex = headers.indexOf(idColumnName);

  if (idColIndex === -1) throw new Error(`Kolom ID '${idColumnName}' tidak ditemukan.`);

  const idRange = sheet.getRange(2, idColIndex + 1, lastRow - 1, 1).getValues();
  const allIds = idRange.map(row => {
    const idStr = String(row[0]).replace(new RegExp(prefix, 'i'), '');
    return parseInt(idStr, 10);
  }).filter(id => !isNaN(id));

  const nextIdNumber = allIds.length > 0 ? Math.max(0, ...allIds) + 1 : 1;
  return prefix + String(nextIdNumber).padStart(padLength, '0');
}

/**
 * Mencari nomor baris berdasarkan nilai unik di sebuah kolom.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Objek sheet.
 * @param {string|number} valueToFind Nilai yang dicari.
 * @param {string} columnName Nama header kolom (setelah normalisasi).
 * @returns {number} Nomor baris (dimulai dari 1) atau -1 jika tidak ditemukan.
 */
function findRowIndexByValue_(sheet, valueToFind, columnName) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).toLowerCase().replace(/ /g, ''));
  const colIndex = headers.indexOf(columnName);

  if (colIndex === -1) return -1;

  for (let i = 1; i < data.length; i++) {
    if (data[i][colIndex] !== undefined && String(data[i][colIndex]).trim() === String(valueToFind).trim()) {
      return i + 1; // Mengembalikan nomor baris aktual (bukan indeks array)
    }
  }
  return -1;
}

/**
 * Memeriksa duplikasi nilai di kolom tertentu, dengan opsi untuk mengecualikan baris tertentu.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Objek sheet.
 * @param {string} valueToCheck Nilai yang akan dicek.
 * @param {string} columnName Nama header kolom (setelah normalisasi).
 * @param {number} [excludeRowIndex=-1] Nomor baris yang akan diabaikan dalam pengecekan (untuk mode update).
 * @returns {boolean} True jika duplikat ditemukan, false jika tidak.
 */
function checkForDuplicate_(sheet, valueToCheck, columnName, excludeRowIndex = -1) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).toLowerCase().replace(/ /g, ''));
  const colIndex = headers.indexOf(columnName);

  if (colIndex === -1) throw new Error(`Kolom '${columnName}' tidak ditemukan untuk cek duplikasi.`);

  const normalizedValue = String(valueToCheck).trim().toLowerCase();

  for (let i = 1; i < data.length; i++) {
    const currentRowIndex = i + 1;
    if (excludeRowIndex !== -1 && currentRowIndex === excludeRowIndex) {
      continue; // Lewati baris yang sedang diedit
    }
    if (data[i][colIndex] && String(data[i][colIndex]).trim().toLowerCase() === normalizedValue) {
      return true; // Duplikat ditemukan
    }
  }
  return false; // Tidak ada duplikat
}

// ==========================================================================
// FUNGSI-FUNGSI GETTER
// ==========================================================================
function getProductFormDependencies() {
  try {
    const sheetNames = [
      PRODUCT_SHEET_NAME, CATEGORY_SHEET_NAME, BASE_UNIT_SHEET_NAME,
      UNIT_SHEET_NAME, WAREHOUSE_SHEET_NAME, SUPPLIER_SHEET_NAME
    ];
    const allData = getDataFromSheets_(sheetNames);

    const dependencies = {
      categories: _processSheetData_(allData[CATEGORY_SHEET_NAME], { id: 'CategoryID', name: 'NamaKategori' }),
      baseUnits: _processSheetData_(allData[BASE_UNIT_SHEET_NAME], { id: 'BaseUnitID', name: 'NamaUnitDasar' }),
      units: _processSheetData_(allData[UNIT_SHEET_NAME], { id: 'UnitID', name: 'NamaUnit', ref: 'BaseUnitID_Ref' }),
      warehouses: _processSheetData_(allData[WAREHOUSE_SHEET_NAME], { id: 'WarehouseID', name: 'NamaGudang' }),
      suppliers: _processSheetData_(allData[SUPPLIER_SHEET_NAME], { id: 'SupplierID', name: 'NamaPemasok' })
    };
    return { success: true, data: dependencies };
  } catch (e) {
    return { success: false, message: 'Gagal memuat data pendukung form: ' + e.message };
  }
}

function getCategories() {
  try {
    const sheetData = getDataFromSheets_([CATEGORY_SHEET_NAME, PRODUCT_SHEET_NAME]);
    const categoriesData = sheetData[CATEGORY_SHEET_NAME];
    const productsData = sheetData[PRODUCT_SHEET_NAME];

    const headerMap = { 'id': 'CategoryID', 'name': 'NamaKategori' };
    const categories = _processSheetData_(categoriesData, headerMap);
    
    // Logika spesifik untuk menghitung produk
    categories.forEach(cat => cat.jumlahProduk = 0);

    if (productsData && productsData.length > 1) {
      const productHeaders = productsData[0].map(h => String(h).toLowerCase().replace(/ /g, ''));
      const categoryIdRefIndex = productHeaders.indexOf('categoryid_ref');
      if (categoryIdRefIndex !== -1) {
        const productCounts = {};
        for (let i = 1; i < productsData.length; i++) {
          const catId = String(productsData[i][categoryIdRefIndex]).trim();
          if (catId) {
            productCounts[catId] = (productCounts[catId] || 0) + 1;
          }
        }
        categories.forEach(category => {
          category.jumlahProduk = productCounts[category.id] || 0;
        });
      }
    }
    return { success: true, data: categories, totalRecords: categories.length };
  } catch (e) {
    return { success: false, message: 'Gagal mengambil data kategori: ' + e.message, data: [] };
  }
}

function getBaseUnits() {
  try {
    const sheetData = getDataFromSheets_([BASE_UNIT_SHEET_NAME])[BASE_UNIT_SHEET_NAME];
    const headerMap = { 'id': 'BaseUnitID', 'name': 'NamaUnitDasar' };
    const processedData = _processSheetData_(sheetData, headerMap);
    return { success: true, data: processedData, totalRecords: processedData.length };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function getUnits(options = {}) {
  try {
    const allSheetData = getDataFromSheets_([UNIT_SHEET_NAME, BASE_UNIT_SHEET_NAME]);
    const unitsData = allSheetData[UNIT_SHEET_NAME];
    const baseUnitsData = allSheetData[BASE_UNIT_SHEET_NAME];

    // Buat Peta Nama untuk BaseUnit
    const baseUnitNameMap = {};
    if(baseUnitsData) {
      const processedBaseUnits = _processSheetData_(baseUnitsData, { 'id': 'BaseUnitID', 'name': 'NamaUnitDasar' });
      processedBaseUnits.forEach(bu => { 
          baseUnitNameMap[bu.id] = bu.name; 
      });
    }

    // Proses data Unit
    const unitHeaderMap = {
        'id': 'UnitID', 'name': 'NamaUnit', 'singkatanUnit': 'SingkatanUnit',
        'baseUnitIdRef': 'BaseUnitID_Ref', 'dibuatPada': 'DibuatPada'
    };
    let units = _processSheetData_(unitsData, unitHeaderMap);

    // Tambahkan nama base unit ke setiap objek unit
    units.forEach(unit => {
        unit.namaBaseUnit = baseUnitNameMap[unit.baseUnitIdRef] || `[ID: ${unit.baseUnitIdRef}]`;
        if (unit.dibuatPada instanceof Date) {
            unit.dibuatPada = unit.dibuatPada.toISOString();
        }
    });

    return { success: true, data: units, totalRecords: units.length };
  } catch (e) {
    return { success: false, message: 'Gagal mengambil Unit Pengukuran: ' + e.message };
  }
}

function getWarehouses(options = {}) {
  try {
    const sheetData = getDataFromSheets_([WAREHOUSE_SHEET_NAME])[WAREHOUSE_SHEET_NAME];
    const headerMap = {
      'id': 'WarehouseID', 'name': 'NamaGudang', 'emailGudang': 'EmailGudang',
      'nomorTelepon': 'NomorTeleponGudang', 'kotaKabupaten': 'KotaKabupatenGudang',
      'kodePos': 'KodePosGudang', 'dibuatPada': 'DibuatPada'
    };
    let processedData = _processSheetData_(sheetData, headerMap);
    if (options.searchTerm) {
      const searchTerm = options.searchTerm.toLowerCase();
      processedData = processedData.filter(item =>
        (item.name && item.name.toLowerCase().includes(searchTerm)) ||
        (item.kotaKabupaten && item.kotaKabupaten.toLowerCase().includes(searchTerm))
      );
    }
    return { success: true, data: processedData, totalRecords: processedData.length };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getSuppliers(options = {}) {
  try {
    const sheetData = getDataFromSheets_([SUPPLIER_SHEET_NAME])[SUPPLIER_SHEET_NAME];
    const headerMap = {
      'id': 'SupplierID', 'name': 'NamaPemasok', 'nomorTelepon': 'NomorTelepon',
      'emailPemasok': 'EmailPemasok', 'alamatPemasok': 'AlamatPemasok', 'dibuatPada': 'DibuatPada'
    };
    let processedData = _processSheetData_(sheetData, headerMap);
    if (options.searchTerm) {
      const searchTerm = options.searchTerm.toLowerCase();
      processedData = processedData.filter(item =>
        (item.name && item.name.toLowerCase().includes(searchTerm)) ||
        (item.emailPemasok && item.emailPemasok.toLowerCase().includes(searchTerm))
      );
    }
    return { success: true, data: processedData, totalRecords: processedData.length };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ==========================================================================
// FUNGSI-FUNGSI UNTUK MODUL KATEGORI
// ==========================================================================
function addCategory(categoryData) {
  Logger.log("LOG_INFO: addCategory - Received: " + JSON.stringify(categoryData));
  const categoryName = String(categoryData.namaKategori || '').trim();
  if (!categoryName) {
    return { success: false, message: 'Nama Kategori wajib diisi.' };
  }

  try {
    const sheet = getSheet_(CATEGORY_SHEET_NAME);
    if (!sheet) return { success: false, message: `Sheet "${CATEGORY_SHEET_NAME}" tidak ditemukan.` };

    // Cek Duplikasi menggunakan helper
    if (checkForDuplicate_(sheet, categoryName, 'namakategori')) {
      return { success: false, message: 'Nama kategori sudah ada.' };
    }

    // Generate ID menggunakan helper
    const newCategoryId = generateNextId_(sheet, 'categoryid', 'CAT');
    
    // Siapkan baris baru (asumsi kolomnya: CategoryID, NamaKategori)
    const newRow = [newCategoryId, categoryName];
    sheet.appendRow(newRow);
    
    SpreadsheetApp.flush();
    Logger.log("LOG_INFO: addCategory - Category successfully added. ID: " + newCategoryId);
    return { 
        success: true, 
        message: 'Kategori berhasil ditambahkan!', 
        category: { 
            categoryId: newCategoryId,
            namaKategori: manualEscapeHtml(categoryName),
            jumlahProduk: 0 
        }
    };
  } catch (e) {
    Logger.log("LOG_ERROR: Error in addCategory: " + e.toString() + " Stack: " + e.stack);
    return { success: false, message: 'Gagal menambahkan kategori: ' + e.message };
  }
}

function updateCategory(categoryData) {
  Logger.log("LOG_INFO: updateCategory - Received: " + JSON.stringify(categoryData));
  const categoryId = String(categoryData.categoryId || '').trim();
  const categoryName = String(categoryData.namaKategori || '').trim();

  if (!categoryId || !categoryName) {
    return { success: false, message: 'CategoryID dan Nama Kategori wajib diisi.' };
  }

  try {
    const sheet = getSheet_(CATEGORY_SHEET_NAME);
    if (!sheet) return { success: false, message: `Sheet "${CATEGORY_SHEET_NAME}" tidak ditemukan.` };

    // Cari baris menggunakan helper
    const rowIndexToUpdate = findRowIndexByValue_(sheet, categoryId, 'categoryid');
    if (rowIndexToUpdate === -1) return { success: false, message: 'Kategori dengan ID tersebut tidak ditemukan.' };
    
    // Cek duplikasi di baris lain menggunakan helper
    if (checkForDuplicate_(sheet, categoryName, 'namakategori', rowIndexToUpdate)) {
      return { success: false, message: 'Nama kategori sudah ada untuk entri lain.' };
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).toLowerCase().replace(/ /g, ''));
    const nameColIndex = headers.indexOf('namakategori');
    if (nameColIndex === -1) throw new Error('Kolom "NamaKategori" tidak ditemukan.');
    
    sheet.getRange(rowIndexToUpdate, nameColIndex + 1).setValue(categoryName);
    SpreadsheetApp.flush();
    
    Logger.log("LOG_INFO: updateCategory - Category successfully updated. ID: " + categoryId);
    return { 
        success: true, 
        message: 'Kategori berhasil diperbarui!',
        updatedCategory: {
            categoryId: categoryId,
            namaKategori: manualEscapeHtml(categoryName)
        }
    };
  } catch (e) {
    Logger.log("LOG_ERROR: Error in updateCategory: " + e.toString() + " Stack: " + e.stack);
    return { success: false, message: 'Gagal memperbarui kategori: ' + e.message };
  }
}

function deleteCategory(categoryId) {
  Logger.log("LOG_INFO: deleteCategory - Attempting to delete: " + categoryId);
  const idToDelete = String(categoryId || '').trim();
  if (!idToDelete) {
    return { success: false, message: 'CategoryID tidak valid untuk dihapus.' };
  }

  try {
    const productSheet = getSheet_(PRODUCT_SHEET_NAME);
    // Cek dependensi di sheet Produk
    if (productSheet) {
      const isUsed = findRowIndexByValue_(productSheet, idToDelete, 'categoryid_ref') !== -1;
      if (isUsed) {
        return { success: false, message: `Gagal menghapus: Kategori ini masih digunakan oleh produk.` };
      }
    }

    const categorySheet = getSheet_(CATEGORY_SHEET_NAME);
    if (!categorySheet) return { success: false, message: `Sheet "${CATEGORY_SHEET_NAME}" tidak ditemukan.` };

    // Cari baris yang akan dihapus menggunakan helper
    const rowNumberToDelete = findRowIndexByValue_(categorySheet, idToDelete, 'categoryid');

    if (rowNumberToDelete !== -1) {
      categorySheet.deleteRow(rowNumberToDelete);
      SpreadsheetApp.flush();
      Logger.log("LOG_INFO: deleteCategory - Category successfully deleted. ID: " + idToDelete);
      return { success: true, message: 'Kategori berhasil dihapus.' };
    } else {
      Logger.log("LOG_WARNING: deleteCategory - Category not found for deletion. ID: " + idToDelete);
      return { success: false, message: 'Kategori dengan ID tersebut tidak ditemukan.' };
    }
  } catch (e) {
    Logger.log("LOG_ERROR: Error in deleteCategory: " + e.toString() + " Stack: " + e.stack);
    return { success: false, message: 'Gagal menghapus kategori: ' + e.message };
  }
}

// ==========================================================================
// FUNGSI-FUNGSI UNTUK MODUL PRODUK
// ==========================================================================

function getProducts(options) {
  try {
    Logger.log("LOG_INFO: getProducts (SORTABLE) - Options: " + JSON.stringify(options));
    
    // 1. Ambil semua data yang diperlukan dalam SATU KALI PANGGILAN
    const requiredSheets = [
      PRODUCT_SHEET_NAME,
      CATEGORY_SHEET_NAME,
      BASE_UNIT_SHEET_NAME,
      UNIT_SHEET_NAME,
      WAREHOUSE_SHEET_NAME,
      SUPPLIER_SHEET_NAME
    ];
    const allData = getDataFromSheets_(requiredSheets);

    // 2. Buat semua "peta nama" (name maps) menggunakan data yang sudah ada
    const categoryNameMap = createNameMap_(allData[CATEGORY_SHEET_NAME], 'categoryid', 'namakategori');
    const baseUnitNameMap = createNameMap_(allData[BASE_UNIT_SHEET_NAME], 'baseunitid', 'namaunitdasar');
    const unitNameMap = createNameMap_(allData[UNIT_SHEET_NAME], 'unitid', 'namaunit');
    const warehouseNameMap = createNameMap_(allData[WAREHOUSE_SHEET_NAME], 'warehouseid', 'namagudang');
    const supplierNameMap = createNameMap_(allData[SUPPLIER_SHEET_NAME], 'supplierid', 'namapemasok');

    Logger.log("LOG_INFO: getProducts - All necessary data maps created efficiently.");

    // 3. Proses data produk
    const productValues = allData[PRODUCT_SHEET_NAME];
    if (!productValues || productValues.length <= 1) {
      return { success: true, data: [], totalRecords: 0, page: 1, totalPages: 0 };
    }

    const page = options && options.page ? parseInt(options.page, 10) : 1;
    const limit = options && options.limit ? parseInt(options.limit, 10) : 10;

    // Ambil parameter sorting dari options
    const sortBy = 'dibuatpada';
    const sortOrder = 'desc';

    const actualHeaders = productValues[0];
    const normalizedSheetHeaders = actualHeaders.map(h => String(h).toLowerCase().replace(/ /g, ''));
    
    const statusAktifColIndex = normalizedSheetHeaders.indexOf("statusaktif");
    if (statusAktifColIndex === -1) {
        Logger.log("LOG_WARNING: getProducts - Kolom 'StatusAktif' tidak ditemukan. Menampilkan semua produk.");
    }

    const dataRows = productValues.slice(1);

    // Filter baris yang aktif
    let activeRows = dataRows;
    if (statusAktifColIndex !== -1) {
        activeRows = dataRows.filter(row => row[statusAktifColIndex] === true);
    }

    // Filter berdasarkan pencarian
    let filteredRows = activeRows;
    const searchTerm = options && options.searchTerm ? String(options.searchTerm).toLowerCase().trim() : '';
    if (searchTerm) {
      const searchColumnIndices = [];
      const nameIndex = normalizedSheetHeaders.indexOf('namaproduk');
      const codeIndex = normalizedSheetHeaders.indexOf('kodeproduk');
      if (nameIndex !== -1) searchColumnIndices.push(nameIndex);
      if (codeIndex !== -1) searchColumnIndices.push(codeIndex);

      if (searchColumnIndices.length > 0) {
        filteredRows = activeRows.filter(row => {
          return searchColumnIndices.some(colIndex => 
            row[colIndex] && String(row[colIndex]).toLowerCase().includes(searchTerm)
          );
        });
      }
    }

    // --- BLOK SORTING DATA (TAMBAHAN BARU) ---
    const sortColIndex = normalizedSheetHeaders.indexOf(sortBy);
    if (sortColIndex !== -1) {
      filteredRows.sort((a, b) => {
        let valA = a[sortColIndex];
        let valB = b[sortColIndex];

        // Penanganan untuk tipe data yang berbeda
        if (typeof valA === 'string' && valA.includes('T') && !isNaN(new Date(valA))) { // Cek apakah ini string tanggal ISO
            valA = new Date(valA);
            valB = new Date(valB);
        } else if (typeof valA === 'number' || !isNaN(Number(valA))) { // Cek angka
            valA = Number(valA);
            valB = Number(valB);
        }

        let comparison = 0;
        if (valA > valB) {
          comparison = 1;
        } else if (valA < valB) {
          comparison = -1;
        }
        
        return sortOrder === 'desc' ? comparison * -1 : comparison;
      });
    }
    // --- AKHIR BLOK SORTING ---

    const totalFilteredRecords = filteredRows.length;
    const startIndex = (page - 1) * limit;
    const rowsForPage = filteredRows.slice(startIndex, startIndex + limit);

    const productHeaderMap = {
      'productid': 'productId',
      'namaproduk': 'productName',
      'kodeproduk': 'kodeProduk',
      'categoryid_ref': 'productCategoryId',      // Kolom ID Kategori di sheet Products
      'unitprodukid_ref': 'unitProdukId',         // BARU: Menyimpan BaseUnitID
      'unitpenjualanid_ref': 'unitPenjualanId',   // BARU: Menyimpan UnitID
      'unitpembelianid_ref': 'unitPembelianId',   // BARU: Menyimpan UnitID
      'hargapokok': 'hargaPokok',
      'hargajual': 'hargaJual',
      'stok': 'stok',
      'stokminimum': 'stokMinimum',
      'catatan': 'catatan',
      'bataskuantitaspembelian': 'batasKuantitasPembelian',
      'warehouseid_ref': 'warehouseId',
      'supplierid_ref': 'pemasokId',
      'statusproduk': 'statusProduk',
      'pajakpenjualan': 'pajakPenjualan',
      'tipepajak': 'tipePajak',
      'dibuatpada': 'dibuatPada'
    };

const productsForPage = rowsForPage.map(row => {
        const product = {};
        let hasIdentifier = false;

        normalizedSheetHeaders.forEach((sheetHeader, index) => {
          const propName = productHeaderMap[sheetHeader];
          if (propName) {
            let value = row[index];
            product[propName] = (value instanceof Date) ? value.toISOString() : value;
            if ((propName === 'productId' || propName === 'productName') && value && String(value).trim() !== '') {
              hasIdentifier = true;
            }
          }
        });
        
        if (hasIdentifier) {
          // Lookup dan escape data
          product.productCategoryName = manualEscapeHtml(categoryNameMap[product.productCategoryId] || `[ID: ${product.productCategoryId}]`);
          product.unitProdukName = manualEscapeHtml(baseUnitNameMap[product.unitProdukId] || `[ID: ${product.unitProdukId}]`);
          product.unitPenjualanName = manualEscapeHtml(unitNameMap[product.unitPenjualanId] || `[ID: ${product.unitPenjualanId}]`);
          product.unitPembelianName = manualEscapeHtml(unitNameMap[product.unitPembelianId] || `[ID: ${product.unitPembelianId}]`);
          product.namaGudang = manualEscapeHtml(warehouseNameMap[product.warehouseId] || `[ID: ${product.warehouseId}]`);
          product.namaPemasok = manualEscapeHtml(supplierNameMap[product.pemasokId] || `[ID: ${product.pemasokId}]`);
          
          // Escape field teks lainnya
          if (product.productName) product.productName = manualEscapeHtml(product.productName);
          if (product.catatan) product.catatan = manualEscapeHtml(product.catatan);

          return product;
        }
        return null;
      }).filter(p => p !== null); // Hapus baris yang tidak valid

    const totalPages = Math.ceil(totalFilteredRecords / limit);
    Logger.log(`LOG_INFO: getProducts (OPTIMIZED) - Returning ${productsForPage.length} products for page ${page}.`);
    
    return { 
        success: true, 
        data: productsForPage, 
        totalRecords: totalFilteredRecords,
        page: page,
        totalPages: totalPages 
    };
  } catch (e) {
    Logger.log("LOG_ERROR: Error in getProducts (OPTIMIZED): " + e.toString() + " Stack: " + e.stack);
    return { success: false, message: e.toString(), data: [], totalRecords: 0, page: 1, totalPages: 0 };
  }
}

function addProduct(productData) {
  Logger.log("LOG_INFO: addProduct - Received: " + JSON.stringify(productData));
  
  // --- Blok Validasi (TETAP SAMA, TIDAK DIUBAH) ---
  if (!productData) { return { success: false, message: 'Data produk tidak diterima oleh server.' }; }

  // Validasi Field Wajib Server-Side
  if (!productData.productName || String(productData.productName).trim() === '') {
    return { success: false, message: 'Nama Produk wajib diisi.' };
  }
  if (!productData.kodeProduk || String(productData.kodeProduk).trim() === '') {
    return { success: false, message: 'Kode Produk wajib diisi.' };
  }
  if (!productData.productCategory || String(productData.productCategory).trim() === '') { // Ini adalah CategoryID
    return { success: false, message: 'Kategori Produk wajib dipilih.' };
  }
  if (!productData.unitProdukId || String(productData.unitProdukId).trim() === '') { // Ini adalah BaseUnitID
    return { success: false, message: 'Unit Produk (Dasar) wajib dipilih.' };
  }
  if (!productData.unitPenjualanId || String(productData.unitPenjualanId).trim() === '') { // Ini adalah UnitID
    return { success: false, message: 'Unit Penjualan wajib dipilih.' };
  }
  if (!productData.unitPembelianId || String(productData.unitPembelianId).trim() === '') { // Ini adalah UnitID
    return { success: false, message: 'Unit Pembelian wajib dipilih.' };
  }
  if (!productData.warehouseId || String(productData.warehouseId).trim() === '') {
    return { success: false, message: 'Gudang wajib dipilih.' };
  }
  if (!productData.pemasokId || String(productData.pemasokId).trim() === '') {
    return { success: false, message: 'Pemasok wajib dipilih.' };
  }
  if (!productData.statusProduk || String(productData.statusProduk).trim() === '') {
    return { success: false, message: 'Status Produk wajib dipilih.' };
  }
  if (productData.hargaPokok === undefined || productData.hargaPokok === null || String(productData.hargaPokok).trim() === '') {
    return { success: false, message: 'Harga Pokok wajib diisi.' };
  }
  if (productData.hargaJual === undefined || productData.hargaJual === null || String(productData.hargaJual).trim() === '') {
    return { success: false, message: 'Harga Jual wajib diisi.' };
  }
  if (!productData.tipePajak || String(productData.tipePajak).trim() === '') {
    return { success: false, message: 'Tipe Pajak wajib dipilih.' };
  }
  if (productData.stok === undefined || productData.stok === null || String(productData.stok).trim() === '') { // Untuk stok awal
    return { success: false, message: 'Kuantitas Stok Awal wajib diisi.' };
  }

  // Validasi Tipe Data Angka
  const hargaPokok = parseFloat(productData.hargaPokok);
  if (isNaN(hargaPokok) || hargaPokok < 0) {
    return { success: false, message: 'Harga Pokok harus berupa angka positif.' };
  }
  const hargaJual = parseFloat(productData.hargaJual);
  if (isNaN(hargaJual) || hargaJual < 0) {
    return { success: false, message: 'Harga Jual harus berupa angka positif.' };
  }
  const stok = parseInt(productData.stok);
  if (isNaN(stok) || stok < 0) {
    return { success: false, message: 'Stok harus berupa angka non-negatif.' };
  }
  if (productData.stokMinimum !== null && productData.stokMinimum !== undefined && (isNaN(parseInt(productData.stokMinimum)) || parseInt(productData.stokMinimum) < 0) ) {
    return { success: false, message: 'Stok Minimum harus berupa angka non-negatif jika diisi.' };
  }
  if (productData.pajakPenjualan !== null && productData.pajakPenjualan !== undefined && (isNaN(parseFloat(productData.pajakPenjualan)) || parseFloat(productData.pajakPenjualan) < 0 || parseFloat(productData.pajakPenjualan) > 100) ) {
    return { success: false, message: 'Pajak Penjualan harus antara 0-100 jika diisi.' };
  }
  if (productData.batasKuantitasPembelian !== null && productData.batasKuantitasPembelian !== undefined && (isNaN(parseFloat(productData.batasKuantitasPembelian)) || parseFloat(productData.batasKuantitasPembelian) < 0) ) {
    return { success: false, message: 'Batas Kuantitas Pembelian harus berupa angka positif jika diisi.' };
  }

  try {
    const sheet = getSheet_(PRODUCT_SHEET_NAME);
    if (!sheet) return { success: false, message: `Sheet "${PRODUCT_SHEET_NAME}" tidak ditemukan.` };

    const inventorySheet = getSheet_(STOCK_INVENTORY_SHEET_NAME);
    if (!inventorySheet) return { success: false, message: `Sheet "${STOCK_INVENTORY_SHEET_NAME}" tidak ditemukan.` };
    
    // Cek duplikasi untuk Kode Produk (contoh, bisa juga nama produk)
    if (checkForDuplicate_(sheet, productData.kodeProduk, 'kodeproduk')) {
      return { success: false, message: 'Kode Produk sudah ada.' };
    }

    // Gunakan helper untuk generate ID
    const newProductId = generateNextId_(sheet, 'productid', 'PD');
    
    const dibuatPada = new Date();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).toLowerCase().replace(/ /g, ''));
    
    // Siapkan data baris baru dalam bentuk objek agar lebih terbaca
    const newRowData = {
      productid: newProductId,
      namaproduk: productData.productName?.trim() || '',
      kodeproduk: productData.kodeProduk?.trim() || '',
      categoryid_ref: productData.productCategory,
      unitprodukid_ref: productData.unitProdukId || '',
      unitpenjualanid_ref: productData.unitPenjualanId || '',
      unitpembelianid_ref: productData.unitPembelianId || '',
      warehouseid_ref: productData.warehouseId || '',
      supplierid_ref: productData.pemasokId || '',
      hargapokok: parseFloat(productData.hargaPokok),
      hargajual: parseFloat(productData.hargaJual),
      stok: parseInt(productData.stok, 10),
      stokminimum: productData.stokMinimum ? parseInt(productData.stokMinimum, 10) : '',
      catatan: productData.catatan?.trim() || '',
      bataskuantitaspembelian: productData.batasKuantitasPembelian ? parseFloat(productData.batasKuantitasPembelian) : '',
      statusproduk: productData.statusProduk?.trim() || '',
      pajakpenjualan: productData.pajakPenjualan ? parseFloat(productData.pajakPenjualan) : '',
      tipepajak: productData.tipePajak?.trim() || '',
      statusaktif: true,
      dibuatpada: dibuatPada
    };

    const newRow = headers.map(header => newRowData[header] !== undefined ? newRowData[header] : '');
    sheet.appendRow(newRow);

    // --- Logika Stok Inventaris (TETAP SAMA, TIDAK DIUBAH) ---
    const warehouseId = productData.warehouseId;
    const initialStock = parseInt(productData.stok, 10);

    if (warehouseId && !isNaN(initialStock) && initialStock > 0) { // Hanya catat jika ada stok awal
      // Langkah 1: Catat di Ledger sebagai 'Sumber Kebenaran'
      _addStockLedgerEntry_({
        productId: newProductId,
        warehouseId: warehouseId,
        transactionType: 'STOK_AWAL',
        quantityChange: initialStock,
        referenceId: newProductId, // Untuk stok awal, referensinya adalah produk itu sendiri
        notes: 'Stok awal dari penambahan produk baru.'
      });
      
      // Langkah 2: Update 'Cache Cepat' di StockInventory
      const stockUpdateSuccess = _updateStockInInventory(inventorySheet, newProductId, warehouseId, initialStock);
      if (stockUpdateSuccess) {
        _updateTotalStockInProducts(SpreadsheetApp.getActiveSpreadsheet(), newProductId);
      } else {
        Logger.log(`LOG_WARNING: addProduct - Produk ${newProductId} berhasil dibuat, tetapi pencatatan stok awal di gudang ${warehouseId} GAGAL.`);
      }
    }
    // --- Akhir Logika Stok Inventaris ---

    SpreadsheetApp.flush();
    Logger.log("LOG_INFO: addProduct - Product successfully added. ID: " + newProductId);
    return { success: true, message: 'Produk berhasil ditambahkan!', productId: newProductId };
  } catch (e) {
    Logger.log("LOG_ERROR: Error in addProduct: " + e.toString() + " Stack: " + e.stack);
    return { success: false, message: 'Gagal menambahkan produk: ' + e.message };
  }
}

function updateProduct(productData) {
  Logger.log("LOG_INFO: updateProduct - Received: " + JSON.stringify(productData));
  
  // Validasi Awal Data Diterima
  if (!productData || !productData.productId) {
       return { success: false, message: 'Product ID tidak lengkap untuk melakukan pembaruan.' };
  }
  // Validasi untuk field yang BOLEH diubah saat edit
  if (!productData.productName || String(productData.productName).trim() === '') {
    return { success: false, message: 'Nama Produk wajib diisi.' };
  }
  if (!productData.kodeProduk || String(productData.kodeProduk).trim() === '') {
    return { success: false, message: 'Kode Produk wajib diisi.' };
  }
  if (!productData.productCategory || String(productData.productCategory).trim() === '') {
    return { success: false, message: 'Kategori Produk wajib dipilih (ID).' };
  }
  if (!productData.unitProdukId || String(productData.unitProdukId).trim() === '') {
    return { success: false, message: 'Unit Produk (Dasar) wajib dipilih.' };
  }
  if (!productData.unitPenjualanId || String(productData.unitPenjualanId).trim() === '') {
    return { success: false, message: 'Unit Penjualan wajib dipilih.' };
  }
  if (!productData.unitPembelianId || String(productData.unitPembelianId).trim() === '') {
    return { success: false, message: 'Unit Pembelian wajib dipilih.' };
  }
  // Validasi format untuk field opsional (jika diisi)
  if (productData.batasKuantitasPembelian !== null && productData.batasKuantitasPembelian !== undefined && (isNaN(parseFloat(productData.batasKuantitasPembelian)) || parseFloat(productData.batasKuantitasPembelian) < 0) ) {
    return { success: false, message: 'Batas Kuantitas Pembelian harus berupa angka positif jika diisi.' };
  }

  try {
    const sheet = getSheet_(PRODUCT_SHEET_NAME);
    if (!sheet) return { success: false, message: `Sheet "${PRODUCT_SHEET_NAME}" tidak ditemukan.` };

    const productIdToFind = String(productData.productId).trim();
    
    // Gunakan helper untuk menemukan baris
    const rowIndex = findRowIndexByValue_(sheet, productIdToFind, 'productid');
    if (rowIndex === -1) {
      return { success: false, message: 'Produk dengan ID tersebut tidak ditemukan.' };
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colMap = {};
    headers.forEach((header, index) => { colMap[String(header).toLowerCase().replace(/ /g, '')] = index; });

    const currentRowValues = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];

    // === HANYA UPDATE FIELD YANG BOLEH DIEDIT DARI FORM INI ===
    if(productData.productName !== undefined) currentRowValues[colMap['namaproduk']] = productData.productName.trim();
    if(productData.kodeProduk !== undefined) currentRowValues[colMap['kodeproduk']] = productData.kodeProduk.trim();
    if(productData.productCategory !== undefined) currentRowValues[colMap['categoryid_ref']] = productData.productCategory;
    if(productData.unitProdukId !== undefined) currentRowValues[colMap['unitprodukid_ref']] = productData.unitProdukId;
    if(productData.unitPenjualanId !== undefined) currentRowValues[colMap['unitpenjualanid_ref']] = productData.unitPenjualanId;
    if(productData.unitPembelianId !== undefined) currentRowValues[colMap['unitpembelianid_ref']] = productData.unitPembelianId;
    if(productData.catatan !== undefined) currentRowValues[colMap['catatan']] = productData.catatan.trim() || '';
    if(productData.batasKuantitasPembelian !== undefined) {
      const bqp = parseFloat(productData.batasKuantitasPembelian);
      currentRowValues[colMap['bataskuantitaspembelian']] = !isNaN(bqp) && bqp >= 0 ? bqp : '';
    }
    // Tambahkan kolom "DiperbaruiPada" jika ada
    //if (colMap['diperbaruipada'] !== undefined) {
    //  currentRowValues[colMap['diperbaruipada']] = new Date();
    // }
    
    sheet.getRange(rowIndex, 1, 1, currentRowValues.length).setValues([currentRowValues]);
    SpreadsheetApp.flush();
    
    Logger.log("LOG_INFO: updateProduct - Product successfully updated. ID: " + productIdToFind);
    return { success: true, message: 'Produk berhasil diperbarui!' };
  } catch (e) {
    Logger.log("LOG_ERROR: Error in updateProduct: " + e.toString() + " Stack: " + e.stack);
    return { success: false, message: 'Gagal memperbarui produk: ' + e.message };
  }
}

/**
 * Mengarsipkan produk (Soft Delete) dengan mengubah StatusAktif menjadi FALSE.
 * @param {string} productId ID Produk yang akan diarsipkan.
 */
function archiveProduct(productId) {
  Logger.log("LOG_INFO: archiveProduct - Attempting to archive product ID: " + productId);
  const idToArchive = String(productId || '').trim();
  if (!idToArchive) {
      return { success: false, message: 'Product ID tidak valid untuk pengarsipan.' };
  }

  try {
    const sheet = getSheet_(PRODUCT_SHEET_NAME);
    if (!sheet) return { success: false, message: `Sheet "${PRODUCT_SHEET_NAME}" tidak ditemukan.` };

    const rowIndex = findRowIndexByValue_(sheet, idToArchive, 'productid');
    if (rowIndex === -1) {
      return { success: false, message: 'Produk dengan ID tersebut tidak ditemukan.' };
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).toLowerCase().replace(/ /g, ''));
    const statusAktifColIndex = headers.indexOf("statusaktif");

    if (statusAktifColIndex === -1) {
      return { success: false, message: 'Kolom "StatusAktif" tidak ditemukan di sheet Products.' };
    }
    
    // Update sel di kolom StatusAktif menjadi FALSE
    sheet.getRange(rowIndex, statusAktifColIndex + 1).setValue(false); // Gunakan boolean false
    SpreadsheetApp.flush();
    
    Logger.log("LOG_INFO: archiveProduct - Product with ID " + idToArchive + " successfully archived.");
    return { success: true, message: 'Produk berhasil diarsipkan.' };

  } catch (e) {
    Logger.log("LOG_ERROR: Error in archiveProduct: " + e.toString() + " Stack: " + e.stack);
    return { success: false, message: 'Gagal mengarsipkan produk: ' + e.message };
  }
}

// ==========================================================================
// FUNGSI-FUNGSI UNTUK MODUL UNIT DASAR (BaseUnits) - VERSI REVISI
// ==========================================================================
function addBaseUnit(data) {
  Logger.log("LOG_INFO: addBaseUnit - Received: " + JSON.stringify(data));
  const name = String(data.namaBaseUnit || '').trim();
  if (!name) {
    return { success: false, message: 'Nama Unit Dasar wajib diisi.' };
  }

  try {
    const sheet = getSheet_(BASE_UNIT_SHEET_NAME);
    if (!sheet) return { success: false, message: `Sheet "${BASE_UNIT_SHEET_NAME}" tidak ditemukan.` };

    if (checkForDuplicate_(sheet, name, 'namaunitdasar')) {
      return { success: false, message: 'Nama Unit Dasar sudah ada.' };
    }

    const newId = generateNextId_(sheet, 'baseunitid', 'BU');
    const newRow = [newId, name, new Date()]; // Asumsi kolom: ID, Nama, DibuatPada
    sheet.appendRow(newRow);
    
    SpreadsheetApp.flush();
    Logger.log("LOG_INFO: addBaseUnit - BaseUnit added: " + newId);
    return { 
        success: true, 
        message: 'Unit Dasar berhasil ditambahkan!',
        baseUnit: { baseUnitId: newId, namaBaseUnit: manualEscapeHtml(name) }
    };
  } catch (e) {
    Logger.log("LOG_ERROR: Error in addBaseUnit: " + e.toString() + "\nStack: " + e.stack);
    return { success: false, message: 'Gagal menambahkan Unit Dasar: ' + e.message };
  }
}

function updateBaseUnit(data) {
  Logger.log("LOG_INFO: updateBaseUnit - Received: " + JSON.stringify(data));
  const id = String(data.baseUnitId || '').trim();
  const name = String(data.namaBaseUnit || '').trim();
  if (!id || !name) {
    return { success: false, message: 'ID dan Nama Unit Dasar wajib diisi.' };
  }

  try {
    const sheet = getSheet_(BASE_UNIT_SHEET_NAME);
    if (!sheet) return { success: false, message: `Sheet "${BASE_UNIT_SHEET_NAME}" tidak ditemukan.` };

    const rowIndex = findRowIndexByValue_(sheet, id, 'baseunitid');
    if (rowIndex === -1) {
      return { success: false, message: 'Unit Dasar dengan ID tersebut tidak ditemukan.' };
    }

    if (checkForDuplicate_(sheet, name, 'namaunitdasar', rowIndex)) {
      return { success: false, message: 'Nama Unit Dasar sudah ada untuk entri lain.' };
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).toLowerCase().replace(/ /g, ''));
    const nameColIndex = headers.indexOf('namaunitdasar');
    if (nameColIndex === -1) throw new Error('Kolom "NamaUnitDasar" tidak ditemukan.');
    
    sheet.getRange(rowIndex, nameColIndex + 1).setValue(name);
    SpreadsheetApp.flush();
    
    Logger.log("LOG_INFO: updateBaseUnit - BaseUnit updated: " + id);
    return { 
        success: true, 
        message: 'Unit Dasar berhasil diperbarui!',
        updatedBaseUnit: { baseUnitId: id, namaBaseUnit: manualEscapeHtml(name) }
    };
  } catch (e) {
    Logger.log("LOG_ERROR: Error in updateBaseUnit: " + e.toString() + "\nStack: " + e.stack);
    return { success: false, message: 'Gagal memperbarui Unit Dasar: ' + e.message };
  }
}

function deleteBaseUnit(baseUnitId) {
  Logger.log("LOG_INFO: deleteBaseUnit - Attempting to delete: " + baseUnitId);
  const idToDelete = String(baseUnitId || '').trim();
  if (!idToDelete) {
    return { success: false, message: 'ID Unit Dasar tidak valid.' };
  }

  try {
    // Cek dependensi di sheet "Units" menggunakan helper
    const unitSheet = getSheet_(UNIT_SHEET_NAME);
    if (unitSheet && findRowIndexByValue_(unitSheet, idToDelete, 'baseunitid_ref') !== -1) {
      return { success: false, message: 'Gagal menghapus: Unit Dasar ini masih digunakan oleh Unit Pengukuran lain.' };
    }
    
    // Cek dependensi di sheet "Products"
    const productSheet = getSheet_(PRODUCT_SHEET_NAME);
    if (productSheet && findRowIndexByValue_(productSheet, idToDelete, 'unitprodukid_ref') !== -1) {
      return { success: false, message: 'Gagal menghapus: Unit Dasar ini masih digunakan oleh satu atau lebih Produk.' };
    }

    const baseUnitSheet = getSheet_(BASE_UNIT_SHEET_NAME);
    if (!baseUnitSheet) return { success: false, message: `Sheet "${BASE_UNIT_SHEET_NAME}" tidak ditemukan.` };

    const rowToDelete = findRowIndexByValue_(baseUnitSheet, idToDelete, 'baseunitid');
    if (rowToDelete !== -1) {
      baseUnitSheet.deleteRow(rowToDelete);
      SpreadsheetApp.flush();
      Logger.log("LOG_INFO: deleteBaseUnit - BaseUnit deleted: " + idToDelete);
      return { success: true, message: 'Unit Dasar berhasil dihapus.' };
    } else {
      return { success: false, message: 'Unit Dasar tidak ditemukan.' };
    }
  } catch (e) {
    Logger.log("LOG_ERROR: Error in deleteBaseUnit: " + e.toString() + "\nStack: " + e.stack);
    return { success: false, message: 'Gagal menghapus Unit Dasar: ' + e.message };
  }
}

// ==========================================================================
// FUNGSI-FUNGSI UNTUK MODUL UNIT PENGUKURAN (Units)
// ==========================================================================
function addUnit(data) {
  Logger.log("LOG_INFO: addUnit - Received: " + JSON.stringify(data));
  const name = String(data.namaUnit || '').trim();
  const baseUnitRef = String(data.baseUnitIdRef || '').trim();
  if (!name || !baseUnitRef) {
    return { success: false, message: 'Nama Unit dan Referensi Unit Dasar wajib diisi.' };
  }

  try {
    const sheet = getSheet_(UNIT_SHEET_NAME);
    if (!sheet) return { success: false, message: `Sheet "${UNIT_SHEET_NAME}" tidak ditemukan.` };

    if (checkForDuplicate_(sheet, name, 'namaunit')) {
      return { success: false, message: 'Nama Unit sudah ada.' };
    }

    const newId = generateNextId_(sheet, 'unitid', 'U');
    const newRowData = {
      unitid: newId,
      namaunit: name,
      singkatanunit: String(data.singkatanUnit || '').trim(),
      baseunitid_ref: baseUnitRef,
      dibuatpada: new Date()
    };
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).toLowerCase().replace(/ /g, ''));
    const newRow = headers.map(header => newRowData[header] || ''); // Build row based on header order
    
    sheet.appendRow(newRow);
    SpreadsheetApp.flush();

    // ... (sisa logika untuk mendapatkan nama base unit dan mengembalikan respons tetap sama)
    // Untuk mempersingkat, bagian mengambil `baseUnitName` bisa kita pertahankan
    let baseUnitName = 'N/A';
    const baseUnitsResponse = getBaseUnits();
    if (baseUnitsResponse && baseUnitsResponse.success && Array.isArray(baseUnitsResponse.data)) {
        const foundBaseUnit = baseUnitsResponse.data.find(bu => bu.baseUnitId === baseUnitRef);
        if (foundBaseUnit) baseUnitName = foundBaseUnit.namaBaseUnit;
    }
    
    return { 
        success: true, 
        message: 'Unit berhasil ditambahkan!',
        unit: { 
            unitId: newId,
            namaUnit: manualEscapeHtml(name),
            singkatanUnit: manualEscapeHtml(String(data.singkatanUnit || '').trim()),
            baseUnitIdRef: baseUnitRef,
            namaBaseUnit: baseUnitName,
            dibuatPada: newRowData.dibuatpada.toISOString()
        }
    };
  } catch (e) {
    Logger.log("LOG_ERROR: Error in addUnit: " + e.toString() + "\nStack: " + e.stack);
    return { success: false, message: 'Gagal menambahkan Unit: ' + e.message };
  }
}

function updateUnit(data) {
  Logger.log("LOG_INFO: updateUnit - Received: " + JSON.stringify(data));
  const id = String(data.unitId || '').trim();
  const name = String(data.namaUnit || '').trim();
  const baseUnitRef = String(data.baseUnitIdRef || '').trim();
  if (!id || !name || !baseUnitRef) {
    return { success: false, message: 'ID Unit, Nama Unit, dan Referensi Unit Dasar wajib diisi.' };
  }

  try {
    const sheet = getSheet_(UNIT_SHEET_NAME);
    if (!sheet) return { success: false, message: `Sheet "${UNIT_SHEET_NAME}" tidak ditemukan.` };

    const rowIndex = findRowIndexByValue_(sheet, id, 'unitid');
    if (rowIndex === -1) {
      return { success: false, message: 'Unit dengan ID tersebut tidak ditemukan.' };
    }

    if (checkForDuplicate_(sheet, name, 'namaunit', rowIndex)) {
      return { success: false, message: 'Nama Unit sudah ada untuk entri lain.' };
    }

    // Untuk update, kita bisa update beberapa kolom sekaligus dengan setValues
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).toLowerCase().replace(/ /g, ''));
    const nameCol = headers.indexOf('namaunit') + 1;
    const abbrCol = headers.indexOf('singkatanunit') + 1;
    const baseRefCol = headers.indexOf('baseunitid_ref') + 1;

    if (nameCol > 0) sheet.getRange(rowIndex, nameCol).setValue(name);
    if (baseRefCol > 0) sheet.getRange(rowIndex, baseRefCol).setValue(baseUnitRef);
    if (abbrCol > 0) sheet.getRange(rowIndex, abbrCol).setValue(String(data.singkatanUnit || '').trim());
    
    SpreadsheetApp.flush();
    // ... (sisa logika untuk mendapatkan nama base unit dan mengembalikan respons tetap sama)
    // ...
    return { 
        success: true, 
        message: 'Unit berhasil diperbarui!',
        // ... (objek updatedUnit bisa disamakan seperti sebelumnya)
    };
  } catch (e) {
    Logger.log("LOG_ERROR: Error in updateUnit: " + e.toString() + "\nStack: " + e.stack);
    return { success: false, message: 'Gagal memperbarui Unit: ' + e.message };
  }
}

function deleteUnit(unitId) {
  Logger.log("LOG_INFO: deleteUnit - Attempting to delete: " + unitId);
  const idToDelete = String(unitId || '').trim();
  if (!idToDelete) {
    return { success: false, message: 'ID Unit tidak valid.' };
  }

  try {
    // Cek dependensi di sheet "Products" menggunakan helper
    const productSheet = getSheet_(PRODUCT_SHEET_NAME);
    if (productSheet) {
      const isUsedInSales = findRowIndexByValue_(productSheet, idToDelete, 'unitpenjualanid_ref') !== -1;
      if (isUsedInSales) {
        return { success: false, message: 'Gagal menghapus: Unit ini masih digunakan sebagai Unit Penjualan pada produk.' };
      }
      const isUsedInPurchase = findRowIndexByValue_(productSheet, idToDelete, 'unitpembelianid_ref') !== -1;
      if (isUsedInPurchase) {
        return { success: false, message: 'Gagal menghapus: Unit ini masih digunakan sebagai Unit Pembelian pada produk.' };
      }
    }
    
    const unitSheet = getSheet_(UNIT_SHEET_NAME);
    if (!unitSheet) return { success: false, message: `Sheet "${UNIT_SHEET_NAME}" tidak ditemukan.` };

    const rowToDelete = findRowIndexByValue_(unitSheet, idToDelete, 'unitid');

    if (rowToDelete !== -1) {
      unitSheet.deleteRow(rowToDelete);
      SpreadsheetApp.flush();
      Logger.log("LOG_INFO: deleteUnit - Unit deleted: " + idToDelete);
      return { success: true, message: 'Unit berhasil dihapus.' };
    } else {
      return { success: false, message: 'Unit tidak ditemukan.' };
    }
  } catch (e) {
    Logger.log("LOG_ERROR: Error in deleteUnit: " + e.toString() + "\nStack: " + e.stack);
    return { success: false, message: 'Gagal menghapus Unit: ' + e.message };
  }
}

// ==========================================================================
// FUNGSI-FUNGSI UNTUK MODUL GUDANG (Warehouses)
// ==========================================================================
function addWarehouse(data) {
  Logger.log("LOG_INFO: addWarehouse - Received: " + JSON.stringify(data));
  const namaGudang = String(data.namaGudang || '').trim();
  if (!namaGudang) {
    return { success: false, message: 'Nama Gudang wajib diisi.' };
  }

  try {
    const sheet = getSheet_(WAREHOUSE_SHEET_NAME);
    if (!sheet) return { success: false, message: `Sheet "${WAREHOUSE_SHEET_NAME}" tidak ditemukan.` };

    if (checkForDuplicate_(sheet, namaGudang, 'namagudang')) {
      return { success: false, message: 'Nama Gudang sudah ada.' };
    }

    const newId = generateNextId_(sheet, 'warehouseid', 'WH');
    
    // Membangun baris baru sesuai urutan header di sheet
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).toLowerCase().replace(/ /g, ''));
    const newRowData = {
      warehouseid: newId,
      namagudang: namaGudang,
      emailgudang: String(data.emailGudang || '').trim(),
      nomortelepongudang: String(data.nomorTelepon || '').trim() ? "'" + String(data.nomorTelepon).trim() : '',
      kotakabupatengudang: String(data.kotaKabupaten || '').trim(),
      kodeposgudang: String(data.kodePos || '').trim(),
      dibuatpada: new Date()
    };
    const newRow = headers.map(header => newRowData[header] || '');
    
    sheet.appendRow(newRow);
    SpreadsheetApp.flush();
    
    Logger.log("LOG_INFO: addWarehouse - Warehouse added: " + newId);
    return { 
        success: true, 
        message: 'Gudang berhasil ditambahkan!',
        warehouse: { warehouseId: newId, namaGudang: manualEscapeHtml(namaGudang) }
    };
  } catch (e) {
    Logger.log("LOG_ERROR: Error in addWarehouse: " + e.toString() + "\nStack: " + e.stack);
    return { success: false, message: 'Gagal menambahkan Gudang: ' + e.message };
  }
}

function updateWarehouse(data) {
  Logger.log("LOG_INFO: updateWarehouse - Received: " + JSON.stringify(data));
  const id = String(data.warehouseId || '').trim();
  const namaGudang = String(data.namaGudang || '').trim();
  if (!id || !namaGudang) {
    return { success: false, message: 'ID Gudang dan Nama Gudang wajib diisi.' };
  }

  try {
    const sheet = getSheet_(WAREHOUSE_SHEET_NAME);
    if (!sheet) return { success: false, message: `Sheet "${WAREHOUSE_SHEET_NAME}" tidak ditemukan.` };

    const rowIndex = findRowIndexByValue_(sheet, id, 'warehouseid');
    if (rowIndex === -1) {
      return { success: false, message: 'Gudang dengan ID tersebut tidak ditemukan.' };
    }

    if (checkForDuplicate_(sheet, namaGudang, 'namagudang', rowIndex)) {
      return { success: false, message: 'Nama Gudang sudah ada untuk entri lain.' };
    }

    // Update data pada baris yang ditemukan
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const oldValues = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
    
    const colMap = {};
    headers.forEach((header, index) => { colMap[String(header).toLowerCase().replace(/ /g, '')] = index; });

    const updatedValues = [...oldValues]; // salin nilai lama
    updatedValues[colMap['namagudang']] = namaGudang;
    updatedValues[colMap['emailgudang']] = String(data.emailGudang || '').trim();
    updatedValues[colMap['nomortelepongudang']] = String(data.nomorTelepon || '').trim() ? "'" + String(data.nomorTelepon).trim() : '';
    updatedValues[colMap['kotakabupatengudang']] = String(data.kotaKabupaten || '').trim();
    updatedValues[colMap['kodeposgudang']] = String(data.kodePos || '').trim();
    
    sheet.getRange(rowIndex, 1, 1, updatedValues.length).setValues([updatedValues]);
    SpreadsheetApp.flush();

    Logger.log("LOG_INFO: updateWarehouse - Warehouse updated: " + id);
    return { success: true, message: 'Gudang berhasil diperbarui!' };
  } catch (e) {
    Logger.log("LOG_ERROR: Error in updateWarehouse: " + e.toString() + "\nStack: " + e.stack);
    return { success: false, message: 'Gagal memperbarui Gudang: ' + e.message };
  }
}

function deleteWarehouse(warehouseId) {
  Logger.log("LOG_INFO: deleteWarehouse - Attempting to delete: " + warehouseId);
  const idToDelete = String(warehouseId || '').trim();
  if (!idToDelete) {
    return { success: false, message: 'ID Gudang tidak valid.' };
  }

  try {
    const productSheet = getSheet_(PRODUCT_SHEET_NAME);
    if (productSheet && findRowIndexByValue_(productSheet, idToDelete, 'warehouseid_ref') !== -1) {
      return { success: false, message: 'Gagal menghapus: Gudang ini masih digunakan oleh satu atau lebih produk.' };
    }

    const warehouseSheet = getSheet_(WAREHOUSE_SHEET_NAME);
    if (!warehouseSheet) return { success: false, message: `Sheet "${WAREHOUSE_SHEET_NAME}" tidak ditemukan.` };

    const rowToDelete = findRowIndexByValue_(warehouseSheet, idToDelete, 'warehouseid');
    if (rowToDelete !== -1) {
      warehouseSheet.deleteRow(rowToDelete);
      SpreadsheetApp.flush();
      return { success: true, message: 'Gudang berhasil dihapus.' };
    } else {
      return { success: false, message: 'Gudang tidak ditemukan.' };
    }
  } catch (e) {
    Logger.log("LOG_ERROR: Error in deleteWarehouse: " + e.toString() + "\nStack: " + e.stack);
    return { success: false, message: 'Gagal menghapus Gudang: ' + e.message };
  }
}

// ==========================================================================
// FUNGSI-FUNGSI UNTUK MODUL PEMASOK (Suppliers)
// ==========================================================================
function updateSupplier(data) {
  Logger.log("LOG_INFO: updateSupplier - Received: " + JSON.stringify(data));
  const id = String(data.supplierId || '').trim();
  const namaPemasok = String(data.namaPemasok || '').trim();
  if (!id || !namaPemasok) {
    return { success: false, message: 'ID Pemasok dan Nama Pemasok wajib diisi.' };
  }

  try {
    const sheet = getSheet_(SUPPLIER_SHEET_NAME);
    if (!sheet) return { success: false, message: `Sheet "${SUPPLIER_SHEET_NAME}" tidak ditemukan.` };

    const rowIndex = findRowIndexByValue_(sheet, id, 'supplierid');
    if (rowIndex === -1) {
      return { success: false, message: 'Pemasok dengan ID tersebut tidak ditemukan.' };
    }

    if (checkForDuplicate_(sheet, namaPemasok, 'namapemasok', rowIndex)) {
      return { success: false, message: 'Nama Pemasok sudah ada untuk entri lain.' };
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colMap = {};
    headers.forEach((header, index) => { colMap[String(header).toLowerCase().replace(/ /g, '')] = index; });
    
    const updatedValues = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
    updatedValues[colMap['namapemasok']] = namaPemasok;
    updatedValues[colMap['emailpemasok']] = String(data.emailPemasok || '').trim();
    updatedValues[colMap['nomortelepon']] = String(data.nomorTelepon || '').trim() ? "'" + String(data.nomorTelepon).trim() : '';
    updatedValues[colMap['alamatpemasok']] = String(data.alamatPemasok || '').trim();
    
    sheet.getRange(rowIndex, 1, 1, updatedValues.length).setValues([updatedValues]);
    SpreadsheetApp.flush();
    
    Logger.log("LOG_INFO: updateSupplier - Supplier updated: " + id);
    return { success: true, message: 'Pemasok berhasil diperbarui!' };
  } catch (e) {
    Logger.log("LOG_ERROR: Error in updateSupplier: " + e.toString() + "\nStack: " + e.stack);
    return { success: false, message: 'Gagal memperbarui Pemasok: ' + e.message };
  }
}

function deleteSupplier(supplierId) {
  Logger.log("LOG_INFO: deleteSupplier - Attempting to delete: " + supplierId);
  const idToDelete = String(supplierId || '').trim();
  if (!idToDelete) {
    return { success: false, message: 'ID Pemasok tidak valid.' };
  }

  try {
    const productSheet = getSheet_(PRODUCT_SHEET_NAME);
    if (productSheet && findRowIndexByValue_(productSheet, idToDelete, 'supplierid_ref') !== -1) {
      return { success: false, message: 'Gagal menghapus: Pemasok ini masih digunakan oleh satu atau lebih produk.' };
    }

    const supplierSheet = getSheet_(SUPPLIER_SHEET_NAME);
    if (!supplierSheet) return { success: false, message: `Sheet "${SUPPLIER_SHEET_NAME}" tidak ditemukan.` };

    const rowToDelete = findRowIndexByValue_(supplierSheet, idToDelete, 'supplierid');
    if (rowToDelete !== -1) {
      supplierSheet.deleteRow(rowToDelete);
      SpreadsheetApp.flush();
      return { success: true, message: 'Pemasok berhasil dihapus.' };
    } else {
      return { success: false, message: 'Pemasok tidak ditemukan.' };
    }
  } catch (e) {
    Logger.log("LOG_ERROR: Error in deleteSupplier: " + e.toString() + "\nStack: " + e.stack);
    return { success: false, message: 'Gagal menghapus Pemasok: ' + e.message };
  }
}

// ==========================================================================
// FUNGSI-FUNGSI UNTUK MODUL INVENTARIS STOK (StockInventory)
// ==========================================================================

/**
 * Mendapatkan data stok produk berdasarkan ProductID atau WarehouseID.
 * @param {object} options Objek berisi { productId: string } ATAU { warehouseId: string }.
 * @return {object} Objek respons { success, data }. 
 * Data berisi array objek [{ productId, warehouseId, quantity, productName, warehouseName, ... }]
 */
function getStockInventory(options) {
  try {
    Logger.log("LOG_INFO: getStockInventory execution started. Options: " + JSON.stringify(options));
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(STOCK_INVENTORY_SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${STOCK_INVENTORY_SHEET_NAME}" tidak ditemukan.`);

    const range = sheet.getDataRange();
    const values = range.getValues();
    let inventoryData = [];

    if (values.length > 1) {
      const headers = values[0].map(h => String(h).toLowerCase().replace(/ /g, ''));
      const productIdRefIndex = headers.indexOf('productid_ref');
      const warehouseIdRefIndex = headers.indexOf('warehouseid_ref');
      const quantityIndex = headers.indexOf('quantity');
      
      if (productIdRefIndex === -1 || warehouseIdRefIndex === -1 || quantityIndex === -1) {
        throw new Error("Header penting (ProductID_Ref, WarehouseID_Ref, Quantity) tidak ditemukan di sheet StockInventory.");
      }

      for (let i = 1; i < values.length; i++) {
        const row = values[i];
        // Tambahkan semua data ke array awal
        inventoryData.push({
          productId: String(row[productIdRefIndex]).trim(),
          warehouseId: String(row[warehouseIdRefIndex]).trim(),
          quantity: parseInt(row[quantityIndex]) || 0
        });
      }

      // Filter berdasarkan productId jika ada
      if (options && options.productId) {
        inventoryData = inventoryData.filter(item => item.productId === String(options.productId).trim());
      }
      // Filter berdasarkan warehouseId jika ada
      if (options && options.warehouseId) {
        inventoryData = inventoryData.filter(item => item.warehouseId === String(options.warehouseId).trim());
      }

      // Jika perlu, tambahkan lookup untuk Nama Produk dan Nama Gudang
      if (inventoryData.length > 0) {
        // (Logika lookup nama produk dan nama gudang bisa ditambahkan di sini agar lebih informatif)
      }
    }
    Logger.log(`LOG_INFO: getStockInventory - Retrieved ${inventoryData.length} inventory records.`);
    return { success: true, data: inventoryData };
  } catch (e) {
    Logger.log("LOG_ERROR: Error in getStockInventory: " + e.toString() + "\nStack: " + e.stack);
    return { success: false, message: 'Gagal mengambil data inventaris: ' + e.toString(), data: [] };
  }
}


/**
 * Fungsi internal untuk mencari dan/atau mengupdate stok di sheet StockInventory.
 * @param {Sheet} sheet Objek sheet StockInventory.
 * @param {string} productId ID Produk.
 * @param {string} warehouseId ID Gudang.
 * @param {number} quantityChange Jumlah yang akan ditambahkan (positif) atau dikurangi (negatif).
 * @return {boolean} true jika berhasil, false jika gagal.
 */
function _updateStockInInventory(sheet, productId, warehouseId, quantityChange) {
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).toLowerCase().replace(/ /g, ''));
  const productIdRefIndex = headers.indexOf('productid_ref');
  const warehouseIdRefIndex = headers.indexOf('warehouseid_ref');
  const quantityIndex = headers.indexOf('quantity');
  const inventoryIdIndex = headers.indexOf('inventoryid');
  const lastUpdateIndex = headers.indexOf('lastupdate');

  let foundRowIndex = -1;
  // Cari baris yang cocok dengan productId dan warehouseId
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][productIdRefIndex]).trim() === productId && String(values[i][warehouseIdRefIndex]).trim() === warehouseId) {
      foundRowIndex = i;
      break;
    }
  }

  if (foundRowIndex !== -1) { // Jika entri inventaris sudah ada, update
    const currentRow = values[foundRowIndex];
    const currentQuantity = parseInt(currentRow[quantityIndex]) || 0;
    const newQuantity = currentQuantity + quantityChange;
    if (newQuantity < 0) {
      Logger.log(`LOG_WARNING: _updateStockInInventory - Gagal update. Stok akan menjadi negatif untuk produk ${productId} di gudang ${warehouseId}.`);
      return false; // Gagal
    }
    sheet.getRange(foundRowIndex + 1, quantityIndex + 1).setValue(newQuantity);
    if (lastUpdateIndex !== -1) {
      sheet.getRange(foundRowIndex + 1, lastUpdateIndex + 1).setValue(new Date());
    }
  } else { // Jika entri belum ada, buat baris baru
    if (quantityChange < 0) {
      Logger.log(`LOG_WARNING: _updateStockInInventory - Gagal update. Mencoba mengurangi stok untuk produk ${productId} yang belum ada di gudang ${warehouseId}.`);
      return false; // Gagal
    }
    
    // ---- Gunakan helper untuk membuat ID baru ----
    const newInventoryId = generateNextId_(sheet, 'inventoryid', 'INV', 4);

    const newRow = Array(headers.length).fill('');
    newRow[inventoryIdIndex] = newInventoryId;
    newRow[productIdRefIndex] = productId;
    newRow[warehouseIdRefIndex] = warehouseId;
    newRow[quantityIndex] = quantityChange;
    if (lastUpdateIndex !== -1) {
        newRow[lastUpdateIndex] = new Date();
    }
    sheet.appendRow(newRow);
  }
  return true; // Berhasil
}

/**
 * Fungsi untuk mengupdate stok total di sheet "Products".
 * @param {Spreadsheet} ss Objek Spreadsheet.
 * @param {string} productId ID Produk yang stok totalnya akan diupdate.
 */
function _updateTotalStockInProducts(ss, productId) {
  const inventorySheet = ss.getSheetByName(STOCK_INVENTORY_SHEET_NAME);
  const productSheet = ss.getSheetByName(PRODUCT_SHEET_NAME);
  if (!inventorySheet || !productSheet) return;

  const inventoryValues = inventorySheet.getDataRange().getValues();
  const inventoryHeaders = inventoryValues[0].map(h => String(h).toLowerCase().replace(/ /g, ''));
  const invProdIdIndex = inventoryHeaders.indexOf('productid_ref');
  const invQtyIndex = inventoryHeaders.indexOf('quantity');
  
  // Hitung total stok dari semua gudang untuk produk ini
  let totalStock = 0;
  for (let i = 1; i < inventoryValues.length; i++) {
    if (String(inventoryValues[i][invProdIdIndex]).trim() === productId) {
      totalStock += parseInt(inventoryValues[i][invQtyIndex]) || 0;
    }
  }

  // Update stok total di sheet "Products"
  const productValues = productSheet.getDataRange().getValues();
  const productHeaders = productValues[0].map(h => String(h).toLowerCase().replace(/ /g, ''));
  const prodIdIndex = productHeaders.indexOf('productid');
  const prodStokIndex = productHeaders.indexOf('stok'); // Asumsi nama kolomnya 'Stok'
  
  if (prodIdIndex === -1 || prodStokIndex === -1) return;

  for (let i = 1; i < productValues.length; i++) {
    if (String(productValues[i][prodIdIndex]).trim() === productId) {
      productSheet.getRange(i + 1, prodStokIndex + 1).setValue(totalStock);
      break;
    }
  }
}

/**
 * Helper internal untuk menambahkan entri baru ke dalam StockLedger.
 * Ini adalah 'source of truth' untuk semua pergerakan stok.
 * @param {object} entry - Objek yang berisi detail transaksi.
 * @param {string} entry.productId - ID produk yang bergerak.
 * @param {string} entry.warehouseId - ID gudang tempat stok bergerak.
 * @param {string} entry.transactionType - Tipe transaksi (misal: 'STOK_AWAL', 'PENJUALAN').
 * @param {number} entry.quantityChange - Jumlah perubahan (+ untuk masuk, - untuk keluar).
 * @param {string} entry.referenceId - ID referensi (misal: ID faktur penjualan, ID produk baru).
 * @param {string} entry.notes - Catatan tambahan.
 */
function _addStockLedgerEntry_(entry) {
  try {
    const ledgerSheet = getSheet_(STOCK_LEDGER_SHEET_NAME); // Kita perlu definisikan konstantanya
    if (!ledgerSheet) {
      Logger.log("LOG_CRITICAL: Sheet StockLedger tidak ditemukan. Pencatatan gagal.");
      return;
    }
    
    // Gunakan helper kita untuk membuat ID Transaksi baru, dengan 6 digit angka
    const transactionId = generateNextId_(ledgerSheet, 'transactionid', 'TRN', 6);
    const timestamp = new Date();

    const newRow = [
      transactionId,
      timestamp,
      entry.productId,
      entry.warehouseId,
      entry.transactionType,
      entry.quantityChange,
      entry.referenceId || '', // Beri nilai default jika tidak ada
      entry.notes || ''      // Beri nilai default jika tidak ada
    ];
    
    ledgerSheet.appendRow(newRow);
    Logger.log(`LOG_INFO: StockLedger entry created. TxID: ${transactionId}, Product: ${entry.productId}, Change: ${entry.quantityChange}`);

  } catch (e) {
    Logger.log("LOG_ERROR: Failed to add stock ledger entry: " + e.toString());
  }
}

// ==========================================================================
// FUNGSI-FUNGSI UNTUK MODUL PEMBELIAN
// ==========================================================================
/**
 * Menambahkan transaksi pembelian baru beserta item-itemnya.
 * @param {object} purchaseData - Data pembelian dari frontend.
 * @param {string} purchaseData.supplierId - ID Pemasok.
 * @param {string} purchaseData.referenceNo - No. referensi dari supplier.
 * @param {string} purchaseData.status - Status pembelian.
 * @param {Array<object>} purchaseData.items - Array item produk.
 * @returns {object} Objek hasil operasi.
 */

function addPurchase(purchaseData) {
  try {
    // Validasi data utama
    if (!purchaseData || !purchaseData.supplierId || !purchaseData.warehouseId || !purchaseData.items || purchaseData.items.length === 0) {
      return { success: false, message: 'Pemasok, Gudang, dan minimal satu item produk wajib diisi.' };
    }

    const purchasesSheet = getSheet_(PURCHASES_SHEET_NAME);
    const purchaseItemsSheet = getSheet_(PURCHASE_ITEMS_SHEET_NAME);
    const inventorySheet = getSheet_(STOCK_INVENTORY_SHEET_NAME);
    if (!purchasesSheet || !purchaseItemsSheet || !inventorySheet) {
      return { success: false, message: 'Satu atau lebih sheet penting (Purchases, PurchaseItems, StockInventory) tidak ditemukan.' };
    }

    // 1. Buat Primary Key Internal (tidak berubah)
    const newPurchaseId = generateNextId_(purchasesSheet, 'purchaseid', 'PUR', 5);
    // 2. Buat Nomor Referensi yang Dilihat Pengguna (BARU)
    const newReferenceNo = generateReferenceNumber_(purchasesSheet, 'PU', 'referenceno'); 
    
    const timestamp = new Date();
    let totalAmount = 0;

    // Proses setiap item yang dibeli
    for (const item of purchaseData.items) {
      if (!item.productId || !item.quantity || !item.costPrice || Number(item.quantity) <= 0) {
        continue; // Lewati item yang tidak valid
      }
    }
    
    // Kalkulasi Grand Total
    const orderTax = Number(purchaseData.orderTax) || 0;
    const orderDiscount = Number(purchaseData.orderDiscount) || 0;
    const shipping = Number(purchaseData.shipping) || 0;
    const grandTotal = totalAmount + orderTax - orderDiscount + shipping;
    
    // Terakhir, catat di Purchases (Master) menggunakan nomor referensi yang baru dibuat
    const purchaseRow = [
      newPurchaseId, 
      timestamp, 
      purchaseData.warehouseId, 
      purchaseData.supplierId, 
      newReferenceNo, // Menggunakan nomor referensi dari sistem
      orderTax, 
      orderDiscount, 
      shipping, 
      grandTotal, 
      purchaseData.paymentType, 
      purchaseData.status, 
      purchaseData.notes
    ];
    purchasesSheet.appendRow(purchaseRow);
    
    SpreadsheetApp.flush();
    return { success: true, message: `Pembelian ${newReferenceNo} berhasil dibuat!`, purchaseId: newPurchaseId };
  } catch (e) {
    Logger.log("LOG_ERROR: Error in addPurchase: " + e.toString() + " Stack: " + e.stack);
    return { success: false, message: 'Terjadi kesalahan server saat membuat pembelian: ' + e.message };
  }
}

/**
 * FUNGSI DEBUGGING
 * Fungsi ini hanya untuk memeriksa nama header asli dari sheet.
 * Tidak akan mempengaruhi fungsi aplikasi lainnya.
 */
function DEBUG_cekNamaHeader() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = ['Warehouses', 'Suppliers']; // Nama sheet yang mau kita cek

  sheetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      const lastColumn = sheet.getLastColumn();
      if (lastColumn > 0) {
        const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
        Logger.log(`Header untuk sheet "${name}": ${JSON.stringify(headers)}`);
      } else {
        Logger.log(`Sheet "${name}" kosong atau tidak memiliki kolom.`);
      }
    } else {
      Logger.log(`Sheet dengan nama "${name}" TIDAK DITEMUKAN.`);
    }
  });
}
