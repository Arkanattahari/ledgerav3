const PRODUCT_SHEET_NAME = 'Products';
const CATEGORY_SHEET_NAME = 'Categories';

function doGet(e) {
  try {
    Logger.log("CAMELCASE_LOG: doGet started");
    const template = HtmlService.createTemplateFromFile('index');
    const htmlOutput = template.evaluate();
    
    htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .setTitle('Ledgera App')

    Logger.log("CAMELCASE_LOG: doGet evaluation complete, returning HtmlOutput.");
    return htmlOutput;

  } catch (error) {
    Logger.log("CAMELCASE_ERROR: Error in doGet: " + error.toString() + " Stack: " + error.stack);
    return HtmlService.createHtmlOutput("<h3>Terjadi kesalahan kritis saat memuat aplikasi.</h3><p>Detail: " + error.toString() + "</p><p>Silakan cek log server untuk informasi lebih lanjut.</p>");
  }
}

function include(filename) {
  try {
    if (!filename || typeof filename !== 'string') {
      Logger.log(`CAMELCASE_ERROR: Invalid filename for include: ${filename}`);
      throw new Error(`Invalid filename provided to include function: ${filename}`);
    }
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (e) {
    Logger.log(`CAMELCASE_ERROR: Error including file: ${filename}. Error: ${e.toString()}`);
    throw new Error(`Gagal include file: ${filename}. Detail: ${e.toString()}`);
  }
}

/**
 * Mendapatkan semua data produk dari Spreadsheet.
 * Mengembalikan objek dengan kunci camelCase.
 */
function manualEscapeHtml(text) {
  if (typeof text !== 'string') {
    return text;
  }
  return text
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
}

function getProducts(options) {
  try {
    Logger.log("getProducts (paginated): Execution started. Options: " + JSON.stringify(options));
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(PRODUCT_SHEET_NAME);
    if (!sheet) {
      Logger.log(`getProducts: Sheet "${PRODUCT_SHEET_NAME}" tidak ditemukan.`);
      throw new Error(`Sheet "${PRODUCT_SHEET_NAME}" tidak ditemukan.`);
    }

    const page = options && options.page ? parseInt(options.page) : 1;
    const limit = options && options.limit ? parseInt(options.limit) : 10; // Default limit 10
    const searchTerm = options && options.searchTerm ? String(options.searchTerm).toLowerCase().trim() : '';
    // const unitFilter = options && options.unitFilter ? String(options.unitFilter).toLowerCase().trim() : ''; // Untuk nanti

    const allValues = sheet.getDataRange().getValues(); // Ambil semua data termasuk header
    if (allValues.length <= 1) {
      Logger.log("getProducts: No data found beyond header row.");
      return { success: true, data: [], totalRecords: 0, page: 1, totalPages: 0 };
    }

    const actualHeaders = allValues[0];
    const normalizedSheetHeaders = actualHeaders.map(h => String(h).toLowerCase().replace(/ /g, ''));
    
    // Tentukan indeks kolom yang akan dicari berdasarkan searchTerm
    const searchColumnIndices = [];
    const nameIndex = normalizedSheetHeaders.indexOf('namaproduk');
    const codeIndex = normalizedSheetHeaders.indexOf('kodeproduk');
    if (nameIndex !== -1) searchColumnIndices.push(nameIndex);
    if (codeIndex !== -1) searchColumnIndices.push(codeIndex);
    // Tambahkan indeks kolom lain jika ingin disertakan dalam pencarian

    // 1. Filter data berdasarkan searchTerm (server-side)
    const filteredRows = [];
    if (searchTerm && searchColumnIndices.length > 0) {
      for (let i = 1; i < allValues.length; i++) { // Mulai dari baris setelah header
        const row = allValues[i];
        let match = false;
        for (const colIndex of searchColumnIndices) {
          if (row[colIndex] && String(row[colIndex]).toLowerCase().includes(searchTerm)) {
            match = true;
            break;
          }
        }
        if (match) {
          filteredRows.push(row);
        }
      }
    } else {
      // Jika tidak ada searchTerm, semua baris data (setelah header) adalah hasil filter awal
      for (let i = 1; i < allValues.length; i++) {
        filteredRows.push(allValues[i]);
      }
    }
    // TODO: Implementasi filter unit di sini jika diperlukan nanti

    const totalFilteredRecords = filteredRows.length;
    Logger.log(`getProducts: Total records after filtering: ${totalFilteredRecords}`);

    // 2. Ambil semua kategori untuk mapping ID ke Nama (hanya sekali)
    const allCategories = getCategories(); 
    const categoryNameMap = {}; 
    if (allCategories && Array.isArray(allCategories)) {
        allCategories.forEach(cat => {
            if (cat && cat.categoryId && cat.namaKategori) {
                categoryNameMap[cat.categoryId] = cat.namaKategori; 
            }
        });
    }
    Logger.log("getProducts: Category name map created.");

    // 3. Terapkan Paginasi (slice) pada filteredRows
    const startIndex = (page - 1) * limit;
    const endIndex = startIndex + limit;
    const rowsForPage = filteredRows.slice(startIndex, endIndex);
    Logger.log(`getProducts: Slicing for page ${page}, limit ${limit}. Start: ${startIndex}, End: ${endIndex}. Rows for page: ${rowsForPage.length}`);

    // 4. Konversi baris untuk halaman saat ini menjadi objek produk (camelCase)
    const productsForPage = [];
    const productHeaderMap = { // Ini harus SAMA dengan yang Anda definisikan sebelumnya
      'productid': 'productId',
      'namaproduk': 'productName',
      'kodeproduk': 'kodeProduk',
      'categoryid_ref': 'productCategoryId',
      'unitproduk': 'unitProduk',
      'hargajual': 'hargaJual',
      'stok': 'stok',
      'dibuatpada': 'dibuatPada',
      'hargapokok': 'hargaPokok',
      'unitpenjualan': 'unitPenjualan',
      'unitpembelian': 'unitPembelian',
      'stokminimum': 'stokMinimum',
      'catatan': 'catatan',
      'bataskuantitaspembelian': 'batasKuantitasPembelian',
      'gudang': 'gudang',
      'pemasok': 'pemasok',
      'statusproduk': 'statusProduk',
      'pajakpenjualan': 'pajakPenjualan',
      'tipepajak': 'tipePajak'
    };

    for (const row of rowsForPage) {
      const product = {};
      let hasIdentifier = false;
      normalizedSheetHeaders.forEach((sheetHeader, index) => {
        const propName = productHeaderMap[sheetHeader];
        if (propName) {
          let value = row[index];
          if (propName !== 'urlGambar' && propName !== 'productCategoryId' && typeof value === 'string' && value.trim() !== '') {
            value = manualEscapeHtml(value);
          }
          product[propName] = (value instanceof Date) ? value.toISOString() : value;
          if ((propName === 'productId' || propName === 'productName') && value && String(value).trim() !== '') {
            hasIdentifier = true;
          }
        }
      });
      
      if (hasIdentifier) { // Seharusnya selalu true jika baris ada di filteredRows
        if (product.productCategoryId && categoryNameMap[product.productCategoryId]) {
          product.productCategory = categoryNameMap[product.productCategoryId];
        } else if (product.productCategoryId) {
          product.productCategory = manualEscapeHtml(`[ID: ${product.productCategoryId}]`);
        } else {
          product.productCategory = 'Tanpa Kategori';
        }
        for (const key in productHeaderMap) {
            if (product[productHeaderMap[key]] === undefined) {
                 product[productHeaderMap[key]] = null; 
            }
        }
        if (product.urlGambar === null) product.urlGambar = '';
        productsForPage.push(product);
      }
    }
    
    const totalPages = Math.ceil(totalFilteredRecords / limit);
    Logger.log(`getProducts: Returning ${productsForPage.length} products for page ${page}. Total filtered: ${totalFilteredRecords}. Total pages: ${totalPages}.`);
    
    return { 
        success: true, 
        data: productsForPage, 
        totalRecords: totalFilteredRecords,
        page: page,
        totalPages: totalPages 
    };

  } catch (e) {
    Logger.log("Error in getProducts (paginated): " + e.toString() + " Stack: " + e.stack);
    // Kembalikan objek error agar withFailureHandler di klien bisa menanganinya dengan baik
    return { success: false, message: e.toString(), data: [], totalRecords: 0, page: 1, totalPages: 0 };
  }
}

// ADD PRODUCT
function addProduct(productData) {
  Logger.log("addProduct: Received productData: " + JSON.stringify(productData));
  if (!productData) {
    return { success: false, message: 'Data produk tidak diterima oleh server.' };
  }
  if (typeof productData.productName === 'undefined' || String(productData.productName).trim() === '') {
    return { success: false, message: 'Nama Produk (productName) wajib diisi.' };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(PRODUCT_SHEET_NAME);
    if (!sheet) return { success: false, message: `Sheet "${PRODUCT_SHEET_NAME}" tidak ditemukan.` };

    // ... (Validasi Server-Side lengkap untuk semua field wajib lainnya) ...
    if (!productData.productCategory || String(productData.productCategory).trim() === '') {
        return { success: false, message: 'ID Kategori Produk wajib diisi.' };
    }

    const hargaPokok = parseFloat(productData.hargaPokok);
    if (isNaN(hargaPokok) || hargaPokok < 0) return { success: false, message: 'Harga Pokok harus angka positif.'};

    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const sheetColMap = {};
    headerRow.forEach((header, index) => {
        sheetColMap[String(header).toLowerCase().replace(/ /g, '')] = index;
    });

    if (sheetColMap['categoryid_ref'] === undefined) { // Pastikan kolom untuk ID Kategori ada di sheet Products
        return { success: false, message: 'Kolom untuk CategoryID_Ref tidak ditemukan di sheet Products.' };
    }

    let nextId = 1;
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const productIDColumnIndex = sheetColMap['productid'];
      if (productIDColumnIndex === undefined) return { success: false, message: 'Kolom "ProductID" tidak ditemukan di sheet Products.' };
      const idRange = sheet.getRange(2, productIDColumnIndex + 1, lastRow - 1, 1).getValues();
      const allIds = idRange.map(row => Number(row[0])).filter(id => !isNaN(id) && id !== null && id !== '');
      nextId = allIds.length ? Math.max(0, ...allIds) + 1 : 1;
    }
    
    const dibuatPada = new Date();
    const newRowData = Array(headerRow.length).fill('');

    // Mengisi newRowData berdasarkan sheetColMap dan productData (camelCase)
    newRowData[sheetColMap['productid']] = nextId;
    newRowData[sheetColMap['namaproduk']] = productData.productName?.trim() || '';
    newRowData[sheetColMap['kodeproduk']] = productData.kodeProduk?.trim() || '';
    
    // SIMPAN CATEGORY ID KE KOLOM YANG TEPAT
    newRowData[sheetColMap['categoryid_ref']] = productData.productCategory; // productData.productCategory adalah ID

    newRowData[sheetColMap['unitproduk']] = productData.unitProduk?.trim() || '';
    newRowData[sheetColMap['unitpenjualan']] = productData.unitPenjualan?.trim() || '';
    newRowData[sheetColMap['unitpembelian']] = productData.unitPembelian?.trim() || '';
    newRowData[sheetColMap['hargapokok']] = parseFloat(productData.hargaPokok);
    newRowData[sheetColMap['hargajual']] = parseFloat(productData.hargaJual);
    newRowData[sheetColMap['stok']] = parseInt(productData.stok);
    newRowData[sheetColMap['stokminimum']] = productData.stokMinimum !== null && productData.stokMinimum !== undefined ? productData.stokMinimum : '';
    newRowData[sheetColMap['catatan']] = productData.catatan?.trim() || '';
    newRowData[sheetColMap['bataskuantitaspembelian']] = productData.batasKuantitasPembelian !== null && productData.batasKuantitasPembelian !== undefined ? productData.batasKuantitasPembelian : '';
    newRowData[sheetColMap['gudang']] = productData.gudang?.trim() || '';
    newRowData[sheetColMap['pemasok']] = productData.pemasok?.trim() || '';
    newRowData[sheetColMap['statusproduk']] = productData.statusProduk?.trim() || '';
    newRowData[sheetColMap['pajakpenjualan']] = productData.pajakPenjualan !== null && productData.pajakPenjualan !== undefined ? productData.pajakPenjualan : '';
    newRowData[sheetColMap['tipepajak']] = productData.tipePajak?.trim() || '';
    if (sheetColMap['dibuatpada'] !== undefined) {
        newRowData[sheetColMap['dibuatpada']] = dibuatPada;
    }
    
    sheet.appendRow(newRowData);
    SpreadsheetApp.flush();
    Logger.log("addProduct: Produk berhasil ditambahkan dengan ID: " + nextId);
    return { 
        success: true, 
        message: 'Produk berhasil ditambahkan!', 
        productId: nextId };
  } catch (e) {
    Logger.log("Error in addProduct: " + e.toString() + " Stack: " + e.stack);
    return { success: false, message: 'Gagal menambahkan produk: ' + e.message };
  }
}

// UPDATE PRODUK
function updateProduct(productData) {
  Logger.log("updateProduct: Received productData: " + JSON.stringify(productData));
  if (!productData || typeof productData.productId === 'undefined') {
    return { success: false, message: 'Product ID tidak lengkap untuk pembaruan.' };
  }
  if (typeof productData.productId === 'undefined') {
    Logger.log("updateProduct: productData.productId is undefined.");
    return { success: false, message: 'Product ID (productId) tidak ada dalam data yang dikirim untuk pembaruan.' };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(PRODUCT_SHEET_NAME);
    if (!sheet) return { success: false, message: `Sheet "${PRODUCT_SHEET_NAME}" tidak ditemukan.` };

    const productIdToFind = Number(productData.productId);
    if (isNaN(productIdToFind)) return { success: false, message: 'Product ID tidak valid.'};

    // Validasi Server-Side (serupa dengan addProduct)
    const hargaPokok = parseFloat(productData.hargaPokok);
    if (isNaN(hargaPokok) || hargaPokok < 0) return { success: false, message: 'Harga Pokok harus angka positif.'};
    // ... (validasi lainnya untuk field wajib) ...

    const dataRange = sheet.getDataRange();
    const allData = dataRange.getValues();
    const headers = allData[0].map(header => String(header).toLowerCase().replace(/ /g, ''));
    let rowIndexToUpdate = -1; // Index baris di sheet (1-based)

    const productIdHeaderIndex = headers.indexOf("productid");
    if (productIdHeaderIndex === -1) return { success: false, message: 'Kolom "ProductID" tidak ditemukan.' };

    for (let i = 1; i < allData.length; i++) {
      if (allData[i][productIdHeaderIndex] !== undefined && Number(allData[i][productIdHeaderIndex]) == productIdToFind) {
        rowIndexToUpdate = i + 1; 
        break;
      }
    }

    if (rowIndexToUpdate === -1) return { success: false, message: 'Produk dengan ID tersebut tidak ditemukan.' };

    const colMap = {}; 
    headers.forEach((header, index) => { colMap[header] = index; });
    const currentRowValues = allData[rowIndexToUpdate - 1]; 

    // Update nilai berdasarkan productData (camelCase)
    if(productData.productName !== undefined && colMap['namaproduk'] !== undefined) currentRowValues[colMap['namaproduk']] = productData.productName.trim();
    if(productData.kodeProduk !== undefined && colMap['kodeproduk'] !== undefined) currentRowValues[colMap['kodeproduk']] = productData.kodeProduk.trim();
    // ... (Lengkapi untuk semua field yang bisa diupdate, serupa dengan addProduct) ...
    if(productData.productCategory !== undefined && colMap['kategoriproduk'] !== undefined) currentRowValues[colMap['kategoriproduk']] = productData.productCategory.trim();
    if(productData.unitProduk !== undefined && colMap['unitproduk'] !== undefined) currentRowValues[colMap['unitproduk']] = productData.unitProduk.trim();
    if(productData.unitPenjualan !== undefined && colMap['unitpenjualan'] !== undefined) currentRowValues[colMap['unitpenjualan']] = productData.unitPenjualan.trim();
    if(productData.unitPembelian !== undefined && colMap['unitpembelian'] !== undefined) currentRowValues[colMap['unitpembelian']] = productData.unitPembelian.trim();
    if(productData.hargaPokok !== undefined && colMap['hargapokok'] !== undefined) currentRowValues[colMap['hargapokok']] = hargaPokok;
    if(productData.hargaJual !== undefined && colMap['hargajual'] !== undefined) currentRowValues[colMap['hargajual']] = parseFloat(productData.hargaJual); // Pastikan konversi
    if(productData.stok !== undefined && colMap['stok'] !== undefined) currentRowValues[colMap['stok']] = parseInt(productData.stok); // Pastikan konversi
    if(productData.stokMinimum !== undefined && colMap['stokminimum'] !== undefined) currentRowValues[colMap['stokminimum']] = productData.stokMinimum !== null ? productData.stokMinimum : '';
    if(productData.catatan !== undefined && colMap['catatan'] !== undefined) currentRowValues[colMap['catatan']] = productData.catatan.trim() || '';
    if(productData.batasKuantitasPembelian !== undefined && colMap['bataskuantitaspembelian'] !== undefined) currentRowValues[colMap['bataskuantitaspembelian']] = productData.batasKuantitasPembelian !== null ? productData.batasKuantitasPembelian : '';
    if(productData.gudang !== undefined && colMap['gudang'] !== undefined) currentRowValues[colMap['gudang']] = productData.gudang.trim();
    if(productData.pemasok !== undefined && colMap['pemasok'] !== undefined) currentRowValues[colMap['pemasok']] = productData.pemasok.trim();
    if(productData.statusProduk !== undefined && colMap['statusproduk'] !== undefined) currentRowValues[colMap['statusproduk']] = productData.statusProduk.trim();
    if(productData.pajakPenjualan !== undefined && colMap['pajakpenjualan'] !== undefined) currentRowValues[colMap['pajakpenjualan']] = productData.pajakPenjualan !== null ? productData.pajakPenjualan : '';
    if(productData.tipePajak !== undefined && colMap['tipepajak'] !== undefined) currentRowValues[colMap['tipepajak']] = productData.tipePajak.trim();
    // DibuatPada biasanya tidak diupdate. Anda bisa menambahkan kolom DiperbaruiPada jika mau.

    sheet.getRange(rowIndexToUpdate, 1, 1, currentRowValues.length).setValues([currentRowValues]);
    SpreadsheetApp.flush();
    Logger.log("updateProduct: Produk berhasil diperbarui untuk ID: " + productData.productId);
        return { 
            success: true, 
            message: 'Produk berhasil diperbarui!'};
  } catch (e) {
    Logger.log("Error in updateProduct (dg gambar): " + e.toString() + " Stack: " + e.stack);
    return { success: false, message: 'Gagal memperbarui produk: ' + e.message };
  }
}

function deleteProduct(productId) {
  Logger.log("deleteProduct: Attempting to delete product with ID: " + productId);
  if (productId === undefined || productId === null) {
      Logger.log("deleteProduct: Invalid Product ID received.");
      return { success: false, message: 'Product ID tidak valid untuk dihapus.' };
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PRODUCT_SHEET_NAME);
  if (!sheet) {
    Logger.log(`deleteProduct: Sheet "${PRODUCT_SHEET_NAME}" tidak ditemukan.`);
    return { success: false, message: `Sheet "${PRODUCT_SHEET_NAME}" tidak ditemukan.` };
  }

  try {
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(header => String(header).toLowerCase().replace(/ /g, ''));
    const productIdHeaderIndex = headers.indexOf("productid");

    if (productIdHeaderIndex === -1) {
         Logger.log("deleteProduct: Kolom 'ProductID' tidak ditemukan sebagai header.");
         return { success: false, message: 'Kolom "ProductID" tidak ditemukan sebagai header untuk penghapusan.' };
    }
    
    let rowNumberToDelete = -1; 
    const idToDelete = Number(productId);

    for (let i = 1; i < data.length; i++) { 
        if (data[i][productIdHeaderIndex] !== undefined && Number(data[i][productIdHeaderIndex]) === idToDelete) {
            rowNumberToDelete = i + 1; 
            break;
        }
    }

    if (rowNumberToDelete !== -1) {
      sheet.deleteRow(rowNumberToDelete);
      SpreadsheetApp.flush();
      Logger.log("deleteProduct: Product with ID " + idToDelete + " successfully deleted from row " + rowNumberToDelete);
      return { success: true, message: 'Produk berhasil dihapus.' };
    } else {
      Logger.log("deleteProduct: Product with ID " + idToDelete + " not found.");
      return { success: false, message: 'Produk dengan ID tersebut tidak ditemukan.' };
    }
  } catch (e) {
    Logger.log("Error in deleteProduct: " + e.toString() + " Stack: " + e.stack);
    return { success: false, message: 'Gagal menghapus produk: ' + e.message };
  }
}

// --- FUNGSI-FUNGSI UNTUK MODUL KATEGORI (SESUAI DESAIN BARU) ---

/**
 * Mendapatkan semua data kategori dari Spreadsheet beserta jumlah produknya.
 * Mengembalikan objek dengan kunci camelCase.
 */
function getCategories() {
  try {
    Logger.log("getCategories: Execution started.");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const categorySheet = ss.getSheetByName(CATEGORY_SHEET_NAME);
    if (!categorySheet) {
      Logger.log(`getCategories: Sheet "${CATEGORY_SHEET_NAME}" tidak ditemukan.`);
      throw new Error(`Sheet "${CATEGORY_SHEET_NAME}" tidak ditemukan.`);
    }
    const categoryRange = categorySheet.getDataRange();
    const categoryValues = categoryRange.getValues();
    const categories = [];

    if (categoryValues.length > 1) {
        const categoryActualHeaders = categoryValues[0];
        const categoryHeaderMap = {
          'categoryid': 'categoryId',
          'namakategori': 'namaKategori'
        };
        const normalizedCategoryHeaders = categoryActualHeaders.map(h => String(h).toLowerCase().replace(/ /g, ''));

        for (let i = 1; i < categoryValues.length; i++) {
          const row = categoryValues[i];
          const category = { jumlahProduk: 0 }; 
          let hasIdentifier = false;
          normalizedCategoryHeaders.forEach((sheetHeader, index) => {
            const propName = categoryHeaderMap[sheetHeader];
            if (propName) {
              let value = row[index];
              // Untuk namaKategori yang akan ditampilkan, kita escape. 
              // ID Kategori biasanya aman dan tidak perlu di-escape.
              if (propName === 'namaKategori' && typeof value === 'string' && value.trim() !== '') {
                value = manualEscapeHtml(value);
              }
              category[propName] = value;
              if ((propName === 'categoryId' || propName === 'namaKategori') && value && String(value).trim() !== '') {
                hasIdentifier = true;
              }
            }
          });
          if (hasIdentifier) {
            for (const key in categoryHeaderMap) {
                if (category[categoryHeaderMap[key]] === undefined) {
                     category[categoryHeaderMap[key]] = null; 
                }
            }
            categories.push(category);
          }
        }
    }
    Logger.log(`getCategories: Retrieved ${categories.length} base categories.`);

    // Menghitung jumlah produk per kategori berdasarkan CategoryID_Ref di sheet Products
    const productSheet = ss.getSheetByName(PRODUCT_SHEET_NAME);
    if (!productSheet || categories.length === 0) {
      Logger.log(`getCategories: Sheet "${PRODUCT_SHEET_NAME}" tidak ditemukan atau tidak ada kategori untuk dihitung produknya.`);
      return categories; // Kembalikan kategori dengan jumlahProduk = 0
    }
    const productRange = productSheet.getRange(1, 1, productSheet.getLastRow(), productSheet.getLastColumn()); // Lebih efisien
    const productValues = productRange.getValues();

    if (productValues.length > 1) {
        const productActualHeaders = productValues[0];
        const normalizedProductHeaders = productActualHeaders.map(h => String(h).toLowerCase().replace(/ /g, ''));
        
        // Cari index kolom 'CategoryID_Ref' (atau nama kolom yang Anda gunakan) di sheet Products
        const productCategoryIDRefColName = 'categoryid_ref'; // SESUAIKAN JIKA NAMA KOLOM ID KATEGORI DI SHEET PRODUK BERBEDA
        const productCategoryIDRefColIndex = normalizedProductHeaders.indexOf(productCategoryIDRefColName);

        if (productCategoryIDRefColIndex !== -1) {
            const productCounts = {}; // { 'CAT001': count, 'CAT002': count }

            for (let i = 1; i < productValues.length; i++) {
                const productCategoryID = productValues[i][productCategoryIDRefColIndex];
                if (productCategoryID && String(productCategoryID).trim() !== '') {
                    const cleanedProductCategoryID = String(productCategoryID).trim();
                    productCounts[cleanedProductCategoryID] = (productCounts[cleanedProductCategoryID] || 0) + 1;
                }
            }
            Logger.log("getCategories: Product counts by category ID: " + JSON.stringify(productCounts));

            categories.forEach(category => {
                category.jumlahProduk = productCounts[category.categoryId] || 0;
            });
        } else {
            Logger.log(`getCategories: Kolom referensi ID Kategori ('${productCategoryIDRefColName}') tidak ditemukan di sheet Products.`);
        }
    }
    
    Logger.log("getCategories: Successfully processed categories with product counts.");
    return categories;
  } catch (e) {
    Logger.log("Error in getCategories: " + e.toString() + " Stack: " + e.stack);
    throw e; 
  }
}


function addCategory(categoryData) {
  Logger.log("addCategory: Received categoryData: " + JSON.stringify(categoryData));
  if (!categoryData || typeof categoryData.namaKategori === 'undefined' || String(categoryData.namaKategori).trim() === '') {
    Logger.log("addCategory: namaKategori is invalid or missing.");
    return { success: false, message: 'Nama Kategori wajib diisi.' };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CATEGORY_SHEET_NAME);
    if (!sheet) return { success: false, message: `Sheet "${CATEGORY_SHEET_NAME}" tidak ditemukan.` };

    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const categoryIDColumnIndex = headerRow.findIndex(header => String(header).toLowerCase().replace(/ /g, '') === "categoryid");
    const namaKategoriColumnIndex = headerRow.findIndex(header => String(header).toLowerCase().replace(/ /g, '') === "namakategori");

    if (categoryIDColumnIndex === -1 || namaKategoriColumnIndex === -1) {
        return { success: false, message: 'Kolom "CategoryID" atau "NamaKategori" tidak ditemukan di sheet Categories.' };
    }

    const existingCategoriesValues = sheet.getRange(2, namaKategoriColumnIndex + 1, Math.max(1, sheet.getLastRow() -1), 1).getValues();
    const normalizedNewCategoryName = String(categoryData.namaKategori).trim().toLowerCase();
    for (let i = 0; i < existingCategoriesValues.length; i++) {
        if (existingCategoriesValues[i][0] && String(existingCategoriesValues[i][0]).trim().toLowerCase() === normalizedNewCategoryName) {
            return { success: false, message: 'Nama kategori sudah ada.' };
        }
    }

    let nextIdNumber = 1;
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
        const idRange = sheet.getRange(2, categoryIDColumnIndex + 1, lastRow - 1, 1).getValues();
        const allIds = idRange.map(row => {
            const idStr = String(row[0]).replace(/CAT/i, '');
            return parseInt(idStr, 10);
        }).filter(id => !isNaN(id));
        if (allIds.length > 0) {
            nextIdNumber = Math.max(0, ...allIds) + 1;
        }
    }
    const newCategoryId = "CAT" + String(nextIdNumber).padStart(3, '0');

    const newRowData = Array(headerRow.length).fill('');
    newRowData[categoryIDColumnIndex] = newCategoryId;
    newRowData[namaKategoriColumnIndex] = String(categoryData.namaKategori).trim();
    
    sheet.appendRow(newRowData);
    SpreadsheetApp.flush();
    Logger.log("addCategory: Category successfully added with ID: " + newCategoryId);
    return { 
        success: true, 
        message: 'Kategori berhasil ditambahkan!', 
        category: { // Kirim kembali objek kategori yang baru, sudah di-escape untuk nama
            categoryId: newCategoryId,
            namaKategori: manualEscapeHtml(String(categoryData.namaKategori).trim()),
            jumlahProduk: 0 
        }
    };
  } catch (e) {
    Logger.log("Error in addCategory: " + e.toString() + " Stack: " + e.stack);
    return { success: false, message: 'Gagal menambahkan kategori: ' + e.message };
  }
}

function updateCategory(categoryData) {
  Logger.log("updateCategory: Received categoryData: " + JSON.stringify(categoryData));
  if (!categoryData || !categoryData.categoryId) {
    return { success: false, message: 'CategoryID wajib ada untuk pembaruan.' };
  }
  if (!categoryData.namaKategori || String(categoryData.namaKategori).trim() === '') {
    return { success: false, message: 'Nama Kategori wajib diisi.' };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CATEGORY_SHEET_NAME);
    if (!sheet) return { success: false, message: `Sheet "${CATEGORY_SHEET_NAME}" tidak ditemukan.` };

    const dataRange = sheet.getDataRange();
    const allData = dataRange.getValues();
    const headers = allData[0].map(header => String(header).toLowerCase().replace(/ /g, ''));
    let rowIndexToUpdate = -1; 
    const categoryIdToFind = String(categoryData.categoryId).trim();

    const categoryIdHeaderIndex = headers.indexOf("categoryid");
    const namaKategoriHeaderIndex = headers.indexOf("namakategori");

    if (categoryIdHeaderIndex === -1 || namaKategoriHeaderIndex === -1) {
        return { success: false, message: 'Kolom "CategoryID" atau "NamaKategori" tidak ditemukan.' };
    }

    for (let i = 1; i < allData.length; i++) {
      if (allData[i][categoryIdHeaderIndex] !== undefined && String(allData[i][categoryIdHeaderIndex]).trim() == categoryIdToFind) {
        rowIndexToUpdate = i + 1; 
        break;
      }
    }

    if (rowIndexToUpdate === -1) return { success: false, message: 'Kategori dengan ID tersebut tidak ditemukan.' };
    
    const normalizedNewCategoryName = String(categoryData.namaKategori).trim().toLowerCase();
    for (let i = 1; i < allData.length; i++) {
        if (i === (rowIndexToUpdate -1) ) continue; 
        if (allData[i][namaKategoriHeaderIndex] && String(allData[i][namaKategoriHeaderIndex]).trim().toLowerCase() === normalizedNewCategoryName) {
            return { success: false, message: 'Nama kategori sudah ada untuk entri lain.' };
        }
    }

    sheet.getRange(rowIndexToUpdate, namaKategoriHeaderIndex + 1).setValue(String(categoryData.namaKategori).trim());
    
    SpreadsheetApp.flush();
    Logger.log("updateCategory: Category with ID " + categoryIdToFind + " successfully updated.");
    return { 
        success: true, 
        message: 'Kategori berhasil diperbarui!',
        // Kirim juga data kategori yang diupdate untuk pembaruan UI jika perlu
        updatedCategory: {
            categoryId: categoryIdToFind,
            namaKategori: manualEscapeHtml(String(categoryData.namaKategori).trim())
            // Jumlah produk tidak diupdate di sini, perlu load ulang jika mau update jumlah
        }
    };
  } catch (e) {
    Logger.log("Error in updateCategory: " + e.toString() + " Stack: " + e.stack);
    return { success: false, message: 'Gagal memperbarui kategori: ' + e.message };
  }
}

/**
 * Menghapus kategori dari Spreadsheet berdasarkan CategoryID.
 * @param {string} categoryId ID kategori yang akan dihapus.
 */
function deleteCategory(categoryId) {
  Logger.log("deleteCategory: Attempting to delete category with ID: " + categoryId);
  if (categoryId === undefined || categoryId === null || String(categoryId).trim() === '') {
      Logger.log("deleteCategory: Invalid CategoryID received.");
      return { success: false, message: 'CategoryID tidak valid untuk dihapus.' };
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CATEGORY_SHEET_NAME);
  if (!sheet) {
    Logger.log(`deleteCategory: Sheet "${CATEGORY_SHEET_NAME}" tidak ditemukan.`);
    return { success: false, message: `Sheet "${CATEGORY_SHEET_NAME}" tidak ditemukan.` };
  }

  try {
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(header => String(header).toLowerCase().replace(/ /g, ''));
    const categoryIdHeaderIndex = headers.indexOf("categoryid");

    if (categoryIdHeaderIndex === -1) {
         Logger.log("deleteCategory: Kolom 'CategoryID' tidak ditemukan sebagai header.");
         return { success: false, message: 'Kolom "CategoryID" tidak ditemukan.' };
    }
    
    let rowNumberToDelete = -1; 
    const idToDelete = String(categoryId).trim();

    for (let i = 1; i < data.length; i++) { 
        if (data[i][categoryIdHeaderIndex] !== undefined && String(data[i][categoryIdHeaderIndex]).trim() === idToDelete) {
            rowNumberToDelete = i + 1; 
            break;
        }
    }

    if (rowNumberToDelete !== -1) {
      sheet.deleteRow(rowNumberToDelete);
      SpreadsheetApp.flush();
      Logger.log("deleteCategory: Category with ID " + idToDelete + " successfully deleted from row " + rowNumberToDelete);
      return { success: true, message: 'Kategori berhasil dihapus.' };
    } else {
      Logger.log("deleteCategory: Category with ID " + idToDelete + " not found.");
      return { success: false, message: 'Kategori dengan ID tersebut tidak ditemukan.' };
    }
  } catch (e) {
    Logger.log("Error in deleteCategory: " + e.toString() + " Stack: " + e.stack);
    return { success: false, message: 'Gagal menghapus kategori: ' + e.message };
  }
}
