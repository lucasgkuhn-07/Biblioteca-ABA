var SPREADSHEET_ID = '1_UGG4_idUSciGESuTax2qyJXZpvE0NzQlHEyVmDg9Ek';

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Biblioteca ABA')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

// --- Database Connection ---
function getSheet(sheetName) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheets = ss.getSheets();
  
  // Try to find sheet case-insensitive and trimmed
  var target = sheetName.toLowerCase().trim();
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName().toLowerCase().trim();
    // Also check for accented versions manually if needed
    if (name === target || 
        (target === 'solicitacoes' && name === 'solicitações') ||
        (target === 'emprestimosativos' && name === 'empréstimos ativos') ||
        (target === 'emprestimosativos' && name === 'empréstimos') || // Fallback to existing 'Empréstimos'
        (target === 'historico' && name === 'histórico')) {
      return sheets[i];
    }
  }
  
  // If not found, create it
  var sheet = ss.insertSheet(sheetName);
  if (sheetName === 'Usuários') {
    sheet.appendRow(['Nome', 'Telefone', 'Email', 'Rede', 'GC', 'Data Cadastro']);
  } else if (sheetName === 'Livros') {
    sheet.appendRow(['Código', 'Título', 'Autor', 'Categoria', 'Capa', 'Status', 'Descrição']);
  } else if (sheetName === 'Solicitacoes') {
    sheet.appendRow(['ID', 'Código Livro', 'Título Livro', 'Telefone Usuário', 'Nome Usuário', 'Data Solicitação', 'Status']);
  } else if (sheetName === 'EmprestimosAtivos') {
    sheet.appendRow(['ID', 'Código Livro', 'Título Livro', 'Telefone Usuário', 'Nome Usuário', 'Data Empréstimo', 'Data Devolução Prevista', 'Status']);
  } else if (sheetName === 'Historico') {
    sheet.appendRow(['ID', 'Código Livro', 'Título Livro', 'Telefone Usuário', 'Nome Usuário', 'Data Empréstimo', 'Data Devolução Efetiva', 'Status']);
  }
  return sheet;
}

// --- Public API ---

/**
 * Fetch all books from 'Livros' sheet.
 * Assumed Columns: Código [0], Título [1], Autor [2], Categoria [3], Capa [4], Status [5], Descrição [6]
 */
/**
 * Fetch all books from all Category sheets.
 * Excludes system sheets: 'Usuários', 'Empréstimos'.
 * Assumed Columns in Category Sheets: Código [0], Título [1], Autor [2], Capa [3] (Optional), Status [4] (Optional), Descrição [5] (Optional)
 * Uses Sheet Name as Category.
 */
function getBooks() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheets = ss.getSheets();
    var allBooks = [];
    var systemSheets = ['Usuários', 'Solicitacoes', 'EmprestimosAtivos', 'Historico', 'Configurações', 'Página1', 'Empréstimos'];

    sheets.forEach(function(sheet) {
      var sheetName = sheet.getName();
      if (systemSheets.indexOf(sheetName) === -1) {
        // It's a category sheet
        try {
          var data = sheet.getDataRange().getValues();
          if (data.length > 1) { // Has data
            var headers = data.shift(); // Remove header
            
            var sheetBooks = data.map(function(row) {
              // User specified columns: 
              // A: Nome (Title), B: Código (Code), C: Descrição, D: Autor, E: Editora, F: Categoria
              return {
                title: row[0],
                code: row[1],
                description: row[2],
                author: row[3],
                publisher: row[4],
                category: row[5] || sheetName, // Use column F, fallback to sheet name
                cover: '', 
                status: row[headers.indexOf('Status')] || 'Disponível'
              };
            });
            allBooks = allBooks.concat(sheetBooks);
          }
        } catch (e) {
          Logger.log('Error reading sheet ' + sheetName + ': ' + e);
        }
      }
    });
    
    return allBooks;
  } catch (e) {
    Logger.log(e);
    return [];
  }
}

/**
 * Verify userLogin or Register
 * user: { name, phone, email, network, gc, termsAccepted }
 */
function loginUser(userData) {
  try {
    var sheet = getSheet('Usuários');
    var data = sheet.getDataRange().getValues();
    
    var uName = String(userData.name || "").toLowerCase().trim();
    var uPhone = String(userData.phone || "").toLowerCase().trim();
    var uGC = String(userData.gc || "").toLowerCase().trim();
    var uRede = String(userData.rede || "").toLowerCase().trim();

    var foundUser = null;
    for (var i = 1; i < data.length; i++) {
      var dbName = String(data[i][0] || "").toLowerCase().trim();
      var dbPhone = String(data[i][1] || "").toLowerCase().trim();
      var dbGC = String(data[i][4] || "").toLowerCase().trim();
      var dbRede = String(data[i][3] || "").toLowerCase().trim();

      // Duplicate Detection Logic:
      // Match by (Phone) OR (Name AND Phone) OR (Name AND GC AND Rede)
      if (dbPhone === uPhone || 
          (dbName === uName && dbPhone === uPhone) ||
          (dbName === uName && dbGC === uGC && dbRede === uRede)) {
        foundUser = data[i];
        break;
      }
    }
    
    if (foundUser) {
      // Return existing user data from DB to ensure consistency
      var existingUser = {
        name: foundUser[0],
        phone: foundUser[1],
        email: foundUser[2],
        rede: foundUser[3],
        gc: foundUser[4]
      };
      return { success: true, user: existingUser, isNew: false };
    } else {
      // Register new user
      sheet.appendRow([
        userData.name,
        "'" + userData.phone,
        userData.email,
        userData.rede, // Updated to match implementation plan fields
        userData.gc,
        new Date()
      ]);
      return { success: true, user: userData, isNew: true };
    }
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Handle Loan Request
 */
function requestLoan(loanData) {
  try {
    var sheet = getSheet('Solicitacoes');
    var bookTitle = getBookTitleByCode(loanData.bookCode);
    var id = Utilities.getUuid();
    
    sheet.appendRow([
      id,
      loanData.bookCode,
      bookTitle,
      "'" + loanData.userPhone,
      loanData.userName,
      new Date(),
      'Pendente'
    ]);

    // Check for auto-approval permission
    if (userHasAutoRetirada(loanData.userPhone)) {
      approveLoan(id, 'Auto aprovado');
      return { success: true, autoApproved: true };
    }

    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Get User Loans
 */
function getUserLoans(phone) {
  var sheet = getSheet('EmprestimosAtivos');
  var data = sheet.getDataRange().getValues();
  var loans = [];
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][3]) === String(phone)) { // Phone is in column D [3]
      loans.push({
        id: data[i][0],
        bookCode: data[i][1],
        bookTitle: data[i][2],
        dueDate: data[i][6],
        status: data[i][7]
      });
    }
  }
  return loans;
}

// --- Admin API ---

function getPendingLoans() {
  return getSheetData('Solicitacoes');
}

function getActiveLoans() {
  return getSheetData('EmprestimosAtivos');
}

function getLoanHistory() {
  return getSheetData('Historico');
}

function getAdminLoans() {
  try {
    return {
      requests: getPendingLoans(),
      active: getActiveLoans(),
      history: getLoanHistory()
    };
  } catch (e) {
    Logger.log(e);
    return { requests: [], active: [], history: [] };
  }
}

function getSheetData(sheetName) {
  try {
    var sheet = getSheet(sheetName);
    var fullData = sheet.getDataRange().getValues();
    
    // Skip empty leading rows
    var data = fullData.filter(function(row) {
      return row.some(function(cell) { return cell !== "" && cell !== null; });
    });
    
    if (data.length <= 1) return [];
    
    var headers = data.shift().map(function(h) { return String(h).trim(); });
    
    // Header to key mapping
    var keyMap = {
      'ID': 'id',
      'Código Livro': 'bookCode',
      'Título Livro': 'bookTitle',
      'Telefone Usuário': 'userPhone',
      'Nome Usuário': 'userName',
      'Data Solicitação': 'dateSolicited',
      'Data Empréstimo': 'dateSolicited',
      'Data Devolução Prevista': 'dueDate',
      'Data Devolução Efetiva': 'actualReturnDate',
      'Status': 'status',
      'Data Cadastro': 'dateRegistered',
      'Email': 'email',
      'Rede': 'rede',
      'GC': 'gc',
      'Tags': 'tags',
      'Nome': 'userName',
      'Telefone': 'userPhone'
    };

    return data.map(function(row) {
      var obj = {};
      headers.forEach(function(header, i) {
        var key = keyMap[header] || header;
        // Convert dates to ISO strings for safer transfer
        var val = row[i];
        if (val instanceof Date) val = val.toISOString();
        obj[key] = val;
      });
      return obj;
    });
  } catch (e) {
    Logger.log('Error in getSheetData for ' + sheetName + ': ' + e);
    return [];
  }
}

function approveLoan(loanId, customStatus) {
  try {
    var solSheet = getSheet('Solicitacoes');
    var activeSheet = getSheet('EmprestimosAtivos');
    var data = solSheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(loanId)) {
            var rowData = data[i];

            // DUPLICITY PROTECTION: Check if book is already taken
            var bookStatus = getBookStatusByCode(rowData[1]);
            if (bookStatus === 'Indisponível') {
                return { success: false, error: 'Este livro já está em uso por outra pessoa.' };
            }
            

        
        // Calculate dates
        var loanDate = new Date();
        var returnDate = new Date();
        returnDate.setDate(loanDate.getDate() + 7);
        
        // Append to Active (ID, Code, Title, Phone, Name, LoanDate, ReturnDate, Status)
        activeSheet.appendRow([
          rowData[0], // ID
          rowData[1], // Code
          rowData[2], // Title
          "'" + rowData[3], // Phone (ensure prefix)
          rowData[4], // Name
          loanDate,
          returnDate,
          customStatus || 'Ativo'
        ]);
        
        // Update Book Status to 'Indisponível'
        updateBookStatus(rowData[1], 'Indisponível');
        
        // Remove from Solicitacoes
        solSheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'Solicitação não encontrada' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function rejectLoan(loanId) {
  try {
    var solSheet = getSheet('Solicitacoes');
    var histSheet = getSheet('Historico');
    var data = solSheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(loanId)) {
        var rowData = data[i];
        
        // Log to Historico (ID, Code, Title, Phone, Name, LoanDate(Solicited), ReturnDate(Rejected), Status)
        histSheet.appendRow([
          rowData[0], // ID
          rowData[1], // Code
          rowData[2], // Title
          "'" + rowData[3], // Phone (ensure prefix)
          rowData[4], // Name
          rowData[5], // Date Solicited
          new Date(), // Date Rejected
          'Negado'
        ]);
        
        solSheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'Solicitação não encontrada' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function renewLoan(loanId) {
  try {
    var activeSheet = getSheet('EmprestimosAtivos');
    var data = activeSheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(loanId)) {
        var currentDueDate = new Date(data[i][6]);
        var newDueDate = new Date();
        newDueDate.setDate(new Date().getDate() + 7); // Renew for 7 days from today
        
        activeSheet.getRange(i + 1, 7).setValue(newDueDate);
        return { success: true };
      }
    }
    return { success: false, error: 'Empréstimo ativo não encontrado' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function returnLoan(loanId) {
  try {
    var activeSheet = getSheet('EmprestimosAtivos');
    var historySheet = getSheet('Historico');
    var data = activeSheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(loanId)) {
        var rowData = data[i];
        
        // Append to History (ID, Code, Title, Phone, Name, LoanDate, ReturnDateEffective, Status)
        historySheet.appendRow([
          rowData[0], // ID
          rowData[1], // Code
          rowData[2], // Title
          "'" + rowData[3], // Phone
          rowData[4], // Name
          rowData[5], // Loan Date
          new Date(), // Actual Return Date
          'Devolvido'
        ]);
        
        // Update Book Status to 'Disponível'
        updateBookStatus(rowData[1], 'Disponível');
        
        // Remove from Active
        activeSheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'Empréstimo ativo não encontrado' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function updateBookStatus(bookCode, status) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheets = ss.getSheets();
  var systemSheets = ['Usuários', 'Solicitacoes', 'EmprestimosAtivos', 'Historico', 'Configurações', 'Página1', 'Empréstimos'];

  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    if (systemSheets.indexOf(sheet.getName()) === -1) {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        // User columns: A: Title [0], B: Code [1], C: Desc [2], D: Author [3], E: Publisher [4], F: Category [5]
        // Wait, where is the Status? The user didn't have a status column in the scrap.
        // I should probably ADD a status column if it's not there, or assume the last column?
        // Let's check headers first. 
        if (String(data[i][1]) === String(bookCode)) {
           // For now, let's assume we can add it at the end if not present or if we know which column it is.
           // However, my getBooks currently returns "Disponível" by default.
           // I'll update the sheet by searching the header "Status".
           var headers = data[0];
           var statusIdx = headers.indexOf('Status');
           if (statusIdx === -1) {
             // Add Status column
             statusIdx = headers.length;
             sheet.getRange(1, statusIdx + 1).setValue('Status');
           }
           sheet.getRange(i + 1, statusIdx + 1).setValue(status);
           return;
        }
      }
    }
  }
}

function getBookTitleByCode(bookCode) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheets = ss.getSheets();
  var systemSheets = ['Usuários', 'Solicitacoes', 'EmprestimosAtivos', 'Historico', 'Configurações', 'Página1', 'Empréstimos'];

  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    if (systemSheets.indexOf(sheet.getName()) === -1) {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][1]) === String(bookCode)) {
          return data[i][0]; // Title is in column A
        }
      }
    }
  }
  return 'Título Desconhecido';
}
function handleLoanAction(id, action) {
  if (action === 'approve') return approveLoan(id);
  if (action === 'reject') return rejectLoan(id);
  if (action === 'return') return returnLoan(id);
  if (action === 'renew') return renewLoan(id);
  return { success: false, error: 'Ação inválida' };
}

function getAdminUsers() {
  try {
    var sheet = getSheet('Usuários');
    var fullData = sheet.getDataRange().getValues();
    
    // Ensure Tags column exists
    var headers = fullData[0];
    var tagsIdx = headers.indexOf('Tags');
    if (tagsIdx === -1) {
      tagsIdx = headers.length;
      sheet.getRange(1, tagsIdx + 1).setValue('Tags');
      headers = sheet.getDataRange().getValues()[0]; // Refresh headers
    }
    
    var data = getSheetData('Usuários');
    
    // Enrich with loan counts
    var activeLoans = getSheetData('EmprestimosAtivos');
    var historyLoans = getSheetData('Historico');

    return data.map(function(user) {
      var phone = String(user.userPhone || "").trim();
      user.activeLoansCount = activeLoans.filter(function(l) { return String(l.userPhone || "").trim() === phone; }).length;
      user.totalLoansCount = user.activeLoansCount + historyLoans.filter(function(l) { return String(l.userPhone || "").trim() === phone; }).length;
      return user;
    });
  } catch (e) {
    Logger.log(e);
    return [];
  }
}

function addUserTag(phone, newTag) {
  try {
    var sheet = getSheet('Usuários');
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var phoneIdx = headers.indexOf('Telefone');
    var tagsIdx = headers.indexOf('Tags');
    
    if (phoneIdx === -1 || tagsIdx === -1) return { success: false, error: 'Colunas não encontradas' };

    for (var i = 1; i < data.length; i++) {
      var dbPhone = String(data[i][phoneIdx] || "").trim();
      var targetPhone = String(phone || "").trim();
      
      if (dbPhone === targetPhone) {
        var currentTags = String(data[i][tagsIdx] || "");
        var tagsArray = currentTags.split(',').map(function(t) { return t.trim(); }).filter(Boolean);
        
        if (tagsArray.indexOf(newTag) === -1) {
          tagsArray.push(newTag);
          sheet.getRange(i + 1, tagsIdx + 1).setValue(tagsArray.join(', '));
        }
        return { success: true };
      }
    }
    return { success: false, error: 'Usuário não encontrado' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function userHasAutoRetirada(phone) {
  try {
    var users = getAdminUsers();
    var user = users.find(function(u) { return String(u.userPhone).trim() === String(phone).trim(); });
    if (!user || !user.tags) return false;
    
    var tagsArray = user.tags.split(',').map(function(t) { return t.trim().toLowerCase(); });
    return tagsArray.indexOf('auto-retirada') !== -1 || tagsArray.indexOf('vip') !== -1;
  } catch (e) {
    return false;
  }
}

function getBookStatusByCode(bookCode) {
  var books = getBooks();
  var book = books.find(function(b) { return String(b.code) === String(bookCode); });
  return book ? book.status : 'Desconhecido';
}

function manualLoanLaunch(loanData) {
  try {
    var activeSheet = getSheet('EmprestimosAtivos');
    var id = Utilities.getUuid();
    var loanDate = new Date();
    var returnDate = new Date();
    returnDate.setDate(loanDate.getDate() + 7);

    activeSheet.appendRow([
      id,
      loanData.bookCode,
      loanData.bookTitle,
      "'" + loanData.userPhone,
      loanData.userName,
      loanDate,
      returnDate,
      'Manual'
    ]);

    updateBookStatus(loanData.bookCode, 'Indisponível');
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}
