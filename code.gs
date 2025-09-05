// Sheet URL
const SHEET_URL = "https://docs.google.com/spreadsheets/d/1-iql_f7zY54V9KBkBSRtzhXXXXXXXXXXXXXzSL8HM3r2P6zkQ/edit?gid=0#gid=0";

// Sheet names
const SHEET_NAMES = {
  CUSTOMERS: 'Customers',
  ITEMS: 'Items',
  ORDERS: 'Master',
  PAYMENTS: 'Payments'
};

// ---------- ONE TIME ONLY: INITIALIZATION ----------
function initializeSheets() {
  try {
    const ss = SpreadsheetApp.openById(getSheetIdFromUrl(SHEET_URL));

    // Helper function to reset sheet
    function resetSheet(sheet, headers, rows = 100, cols = headers.length) {
      // Clear existing data
      sheet.clear();
      
      // Ensure correct number of rows & columns
      if (sheet.getMaxRows() > rows) {
        sheet.deleteRows(rows + 1, sheet.getMaxRows() - rows);
      } else if (sheet.getMaxRows() < rows) {
        sheet.insertRowsAfter(sheet.getMaxRows(), rows - sheet.getMaxRows());
      }
      
      if (sheet.getMaxColumns() > cols) {
        sheet.deleteColumns(cols + 1, sheet.getMaxColumns() - cols);
      } else if (sheet.getMaxColumns() < cols) {
        sheet.insertColumnsAfter(sheet.getMaxColumns(), cols - sheet.getMaxColumns());
      }

      // Add headers in first row
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    // CUSTOMERS sheet
    let customersSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMERS);
    if (!customersSheet) {
      customersSheet = ss.insertSheet(SHEET_NAMES.CUSTOMERS);
    }
    resetSheet(customersSheet, ["Id", "Name", "Mobile", "Address", "Remarks"]);

    // ITEMS sheet
    let itemsSheet = ss.getSheetByName(SHEET_NAMES.ITEMS);
    if (!itemsSheet) {
      itemsSheet = ss.insertSheet(SHEET_NAMES.ITEMS);
    }
    resetSheet(itemsSheet, ["Id", "Name", "Price", "Remarks"]);

    // ORDERS sheet
    let ordersSheet = ss.getSheetByName(SHEET_NAMES.ORDERS);
    if (!ordersSheet) {
      ordersSheet = ss.insertSheet(SHEET_NAMES.ORDERS);
    }
    resetSheet(ordersSheet, ["Id", "Date", "Customer", "Items", "BillAmount", "Remarks"]);

    // PAYMENTS sheet
    let paymentsSheet = ss.getSheetByName(SHEET_NAMES.PAYMENTS);
    if (!paymentsSheet) {
      paymentsSheet = ss.insertSheet(SHEET_NAMES.PAYMENTS);
    }
    resetSheet(paymentsSheet, ["Id", "Date", "Customer", "PaidAmount", "Remarks"]);

    return "Sheets initialized and reset successfully!";
  } catch (error) {
    console.error("Error initializing sheets: " + error.toString());
    throw new Error("Failed to initialize sheets: " + error.message);
  }
}

// ---------- INITIALIZATION ----------
function initializeSheets1() {
  try {
    const ss = SpreadsheetApp.openById(getSheetIdFromUrl(SHEET_URL));
    
    // Create CUSTOMERS sheet if it doesn't exist
    let customersSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMERS);
    if (!customersSheet) {
      customersSheet = ss.insertSheet(SHEET_NAMES.CUSTOMERS);
      customersSheet.appendRow(["Id", "Name", "Mobile", "Address", "Remarks"]);
    }
    
    // Create ITEMS sheet if it doesn't exist
    let itemsSheet = ss.getSheetByName(SHEET_NAMES.ITEMS);
    if (!itemsSheet) {
      itemsSheet = ss.insertSheet(SHEET_NAMES.ITEMS);
      itemsSheet.appendRow(['Id', 'Name', 'Price','Remarks']);
    }
    
    // Create ORDERS sheet if it doesn't exist
    let ordersSheet = ss.getSheetByName(SHEET_NAMES.ORDERS);
    if (!ordersSheet) {
      ordersSheet = ss.insertSheet(SHEET_NAMES.ORDERS);
      ordersSheet.appendRow(["Id", "Date", "Customer", "Items", "BillAmount", "Remarks"]);
    }

    // Create PAYMENTS sheet if it doesn't exist
    let paymentsSheet = ss.getSheetByName(SHEET_NAMES.PAYMENTS);
    if (!paymentsSheet) {
      paymentsSheet = ss.insertSheet(SHEET_NAMES.PAYMENTS);
      paymentsSheet.appendRow(["Id", "Date", "Customer", "PaidAmount", "Remarks"]);
    }
    
    return "Sheets initialized successfully!";
  } catch (error) {
    console.error("Error initializing sheets: " + error.toString());
    throw new Error("Failed to initialize sheets: " + error.message);
  }
}

function getSheetIdFromUrl(url) {
  const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : null;
}

// Get active spreadsheet
function getActiveSpreadsheet() {
  try {
    return SpreadsheetApp.openByUrl(SHEET_URL);
  } catch (error) {
    throw new Error('Unable to access spreadsheet. Please check the URL and permissions.');
  }
}

// Get sheet by name
function getSheet(sheetName) {
  try {
    const ss = getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found.`);
    }
    return sheet;
  } catch (error) {
    throw new Error(`Error accessing sheet: ${error.message}`);
  }
}

// Get all data for the web app
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Tea Shop Management')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Get all data for the web app
function getAllData() {
  try {
    return {
      customers: getCustomers(),
      items: getItems(),
      orders: getOrders(),
      payments: getPayments()
    };
  } catch (error) {
    console.error('Error in getAllData:', error);
    throw new Error('Failed to load data. Please try again.');
  }
}

// Get all customers
function getCustomers() {
  try {
    const sheet = getSheet(SHEET_NAMES.CUSTOMERS);
    const data = sheet.getDataRange().getValues();
    
    // Remove header row
    if (data.length > 0) data.shift();
    
    return data.map(row => ({
      id: row[0],
      name: row[1],
      mobile: row[2],
      address: row[3],
      remarks: row[4]
    })).sort((a, b) => b.id - a.id);
  } catch (error) {
    console.error('Error in getCustomers:', error);
    throw new Error('Failed to load customers data.');
  }
}

// Get all items
function getItems() {
  try {
    const sheet = getSheet(SHEET_NAMES.ITEMS);
    const data = sheet.getDataRange().getValues();
    
    // Remove header row
    if (data.length > 0) data.shift();
    
    return data.map(row => ({
      id: row[0],
      name: row[1],
      price: row[2],
      remarks: row[3]
    })).sort((a, b) => b.id - a.id);
  } catch (error) {
    console.error('Error in getItems:', error);
    throw new Error('Failed to load items data.');
  }
}

// Get all orders
function getOrders() {
  try {
    const sheet = getSheet(SHEET_NAMES.ORDERS);
    const data = sheet.getDataRange().getValues();
    
    // Remove header row
    if (data.length > 0) data.shift();
    
    return data.map(row => ({
      id: row[0],
      date: formatDate(row[1]),
      customerId: row[2],
      items: row[3],
      billAmount: row[4],
      remarks: row[5]
    })).sort((a, b) => b.id - a.id);
  } catch (error) {
    console.error('Error in getOrders:', error);
    throw new Error('Failed to load orders data.');
  }
}

// Get all payments
function getPayments() {
  try {
    const sheet = getSheet(SHEET_NAMES.PAYMENTS);
    const data = sheet.getDataRange().getValues();
    
    // Remove header row
    if (data.length > 0) data.shift();
    
    return data.map(row => ({
      id: row[0],
      date: formatDate(row[1]),
      customerId: row[2],
      amount: row[3],
      remarks: row[4]
    })).sort((a, b) => b.id - a.id);
  } catch (error) {
    console.error('Error in getPayments:', error);
    throw new Error('Failed to load payments data.');
  }
}

// Format date to YYYY-MM-DD
function formatDate(date) {
  if (!date) return '';
  if (typeof date === 'string') return date;
  
  try {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (error) {
    console.error('Error formatting date:', date, error);
    return '';
  }
}

// Validate customer data
function validateCustomerData(customerData) {
  if (!customerData.name || customerData.name.trim() === '') {
    throw new Error('Customer name is required.');
  }
  if (!customerData.mobile || customerData.mobile.trim() === '') {
    throw new Error('Customer mobile number is required.');
  }
  return true;
}

// Validate item data
function validateItemData(itemData) {
  if (!itemData.name || itemData.name.trim() === '') {
    throw new Error('Item name is required.');
  }
  if (!itemData.price || isNaN(parseFloat(itemData.price)) || parseFloat(itemData.price) <= 0) {
    throw new Error('Valid item price is required.');
  }
  return true;
}

// Validate order data
function validateOrderData(orderData) {
  if (!orderData.date || orderData.date.trim() === '') {
    throw new Error('Order date is required.');
  }
  if (!orderData.customerId) {
    throw new Error('Customer selection is required.');
  }
  if (!orderData.items || orderData.items.trim() === '') {
    throw new Error('At least one item must be selected.');
  }
  if (!orderData.billAmount || isNaN(parseFloat(orderData.billAmount)) || parseFloat(orderData.billAmount) <= 0) {
    throw new Error('Valid bill amount is required.');
  }
  return true;
}

// Validate payment data
function validatePaymentData(paymentData) {
  if (!paymentData.date || paymentData.date.trim() === '') {
    throw new Error('Payment date is required.');
  }
  if (!paymentData.customerId) {
    throw new Error('Customer selection is required.');
  }
  if (!paymentData.amount || isNaN(parseFloat(paymentData.amount)) || parseFloat(paymentData.amount) <= 0) {
    throw new Error('Valid payment amount is required.');
  }
  return true;
}

// Save customer (create or update)
function saveCustomer(customerData) {
  try {
    validateCustomerData(customerData);
    
    const sheet = getSheet(SHEET_NAMES.CUSTOMERS);
    const data = sheet.getDataRange().getValues();
    
    if (customerData.id) {
      // Update existing customer
      let found = false;
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] == customerData.id) {
          sheet.getRange(i + 1, 2).setValue(customerData.name.trim());
          sheet.getRange(i + 1, 3).setValue(customerData.mobile.trim());
          sheet.getRange(i + 1, 4).setValue((customerData.address || '').trim());
          sheet.getRange(i + 1, 5).setValue((customerData.remarks || '').trim());
          found = true;
          break;
        }
      }
      if (!found) {
        throw new Error('Customer not found for update.');
      }
    } else {
      // Create new customer
      const id = generateNextId(data);
      sheet.appendRow([
        id,
        customerData.name.trim(),
        customerData.mobile.trim(),
        (customerData.address || '').trim(),
        (customerData.remarks || '').trim()
      ]);
    }
    
    return { success: true, message: 'Customer saved successfully.' };
  } catch (error) {
    console.error('Error in saveCustomer:', error);
    throw error;
  }
}

// Save item (create or update)
function saveItem(itemData) {
  try {
    validateItemData(itemData);
    
    const sheet = getSheet(SHEET_NAMES.ITEMS);
    const data = sheet.getDataRange().getValues();
    
    if (itemData.id) {
      // Update existing item
      let found = false;
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] == itemData.id) {
          sheet.getRange(i + 1, 2).setValue(itemData.name.trim());
          sheet.getRange(i + 1, 3).setValue(parseFloat(itemData.price));
          sheet.getRange(i + 1, 4).setValue((itemData.remarks || '').trim());
          found = true;
          break;
        }
      }
      if (!found) {
        throw new Error('Item not found for update.');
      }
    } else {
      // Create new item
      const id = generateNextId(data);
      sheet.appendRow([
        id,
        itemData.name.trim(),
        parseFloat(itemData.price),
        (itemData.remarks || '').trim()
      ]);
    }
    
    return { success: true, message: 'Item saved successfully.' };
  } catch (error) {
    console.error('Error in saveItem:', error);
    throw error;
  }
}

// Save order (create or update)
function saveOrder(orderData) {
  try {
    validateOrderData(orderData);
    
    const sheet = getSheet(SHEET_NAMES.ORDERS);
    const data = sheet.getDataRange().getValues();
    
    if (orderData.id) {
      // Update existing order
      let found = false;
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] == orderData.id) {
          sheet.getRange(i + 1, 2).setValue(orderData.date);
          sheet.getRange(i + 1, 3).setValue(orderData.customerId);
          sheet.getRange(i + 1, 4).setValue(orderData.items);
          sheet.getRange(i + 1, 5).setValue(parseFloat(orderData.billAmount));
          sheet.getRange(i + 1, 6).setValue((orderData.remarks || '').trim());
          found = true;
          break;
        }
      }
      if (!found) {
        throw new Error('Order not found for update.');
      }
    } else {
      // Create new order
      const id = generateNextId(data);
      sheet.appendRow([
        id,
        orderData.date,
        orderData.customerId,
        orderData.items,
        parseFloat(orderData.billAmount),
        (orderData.remarks || '').trim()
      ]);
    }
    
    return { success: true, message: 'Order saved successfully.' };
  } catch (error) {
    console.error('Error in saveOrder:', error);
    throw error;
  }
}

// Save payment (create or update)
function savePayment(paymentData) {
  try {
    validatePaymentData(paymentData);
    
    const sheet = getSheet(SHEET_NAMES.PAYMENTS);
    const data = sheet.getDataRange().getValues();
    
    if (paymentData.id) {
      // Update existing payment
      let found = false;
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] == paymentData.id) {
          sheet.getRange(i + 1, 2).setValue(paymentData.date);
          sheet.getRange(i + 1, 3).setValue(paymentData.customerId);
          sheet.getRange(i + 1, 4).setValue(parseFloat(paymentData.amount));
          sheet.getRange(i + 1, 5).setValue((paymentData.remarks || '').trim());
          found = true;
          break;
        }
      }
      if (!found) {
        throw new Error('Payment not found for update.');
      }
    } else {
      // Create new payment
      const id = generateNextId(data);
      sheet.appendRow([
        id,
        paymentData.date,
        paymentData.customerId,
        parseFloat(paymentData.amount),
        (paymentData.remarks || '').trim()
      ]);
    }
    
    return { success: true, message: 'Payment saved successfully.' };
  } catch (error) {
    console.error('Error in savePayment:', error);
    throw error;
  }
}

// Generate next ID based on existing data
function generateNextId(data) {
  if (data.length <= 1) return 1;
  
  let maxId = 0;
  for (let i = 1; i < data.length; i++) {
    const id = Number(data[i][0]);
    if (!isNaN(id) && id > maxId) maxId = id;
  }
  
  return maxId + 1;
}

// Delete customer
function deleteCustomer(id) {
  try {
    const sheet = getSheet(SHEET_NAMES.CUSTOMERS);
    const data = sheet.getDataRange().getValues();
    
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        sheet.deleteRow(i + 1);
        found = true;
        break;
      }
    }
    
    if (!found) {
      throw new Error('Customer not found for deletion.');
    }
    
    return { success: true, message: 'Customer deleted successfully.' };
  } catch (error) {
    console.error('Error in deleteCustomer:', error);
    throw error;
  }
}

// Delete item
function deleteItem(id) {
  try {
    const sheet = getSheet(SHEET_NAMES.ITEMS);
    const data = sheet.getDataRange().getValues();
    
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        sheet.deleteRow(i + 1);
        found = true;
        break;
      }
    }
    
    if (!found) {
      throw new Error('Item not found for deletion.');
    }
    
    return { success: true, message: 'Item deleted successfully.' };
  } catch (error) {
    console.error('Error in deleteItem:', error);
    throw error;
  }
}

// Delete order
function deleteOrder(id) {
  try {
    const sheet = getSheet(SHEET_NAMES.ORDERS);
    const data = sheet.getDataRange().getValues();
    
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        sheet.deleteRow(i + 1);
        found = true;
        break;
      }
    }
    
    if (!found) {
      throw new Error('Order not found for deletion.');
    }
    
    return { success: true, message: 'Order deleted successfully.' };
  } catch (error) {
    console.error('Error in deleteOrder:', error);
    throw error;
  }
}

// Delete payment
function deletePayment(id) {
  try {
    const sheet = getSheet(SHEET_NAMES.PAYMENTS);
    const data = sheet.getDataRange().getValues();
    
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        sheet.deleteRow(i + 1);
        found = true;
        break;
      }
    }
    
    if (!found) {
      throw new Error('Payment not found for deletion.');
    }
    
    return { success: true, message: 'Payment deleted successfully.' };
  } catch (error) {
    console.error('Error in deletePayment:', error);
    throw error;
  }
}

// Get orders by date range
function getOrdersByDateRange(dateFrom, dateTo) {
  try {
    const allOrders = getOrders();
    
     if (!dateFrom && !dateTo) {
      return allOrders.sort((a, b) => b.id - a.id);
    }
    
    return allOrders.filter(order => {
      const orderDate = order.date;
      return (!dateFrom || orderDate >= dateFrom) && 
             (!dateTo || orderDate <= dateTo);
    }).sort((a, b) => b.id - a.id); // sort by Id DESC;
  } catch (error) {
    console.error('Error in getOrdersByDateRange:', error);
    throw new Error('Failed to filter orders by date range.');
  }
}

// Get payments by date range
function getPaymentsByDateRange(dateFrom, dateTo) {
  try {
    const allPayments = getPayments();
    
    if (!dateFrom && !dateTo) {
      return allPayments.sort((a, b) => b.id - a.id);
    }
    
    return allPayments.filter(payment => {
      const paymentDate = payment.date;
      return (!dateFrom || paymentDate >= dateFrom) && 
             (!dateTo || paymentDate <= dateTo);
    }).sort((a, b) => b.id - a.id);
  } catch (error) {
    console.error('Error in getPaymentsByDateRange:', error);
    throw new Error('Failed to filter payments by date range.');
  }
}

// Get sales report data
function getSalesReport(dateFrom, dateTo) {
  try {
    const orders = getOrdersByDateRange(dateFrom, dateTo);
    const payments = getPaymentsByDateRange(dateFrom, dateTo);
    const customers = getCustomers();
    const items = getItems();
    
    // Calculate total revenue
    const totalRevenue = orders.reduce((sum, order) => sum + parseFloat(order.billAmount || 0), 0);
    
    // Calculate total payments
    const totalPayments = payments.reduce((sum, payment) => sum + parseFloat(payment.amount || 0), 0);
    
    // Group orders by customer
    const customerOrders = {};
    orders.forEach(order => {
      if (!customerOrders[order.customerId]) {
        customerOrders[order.customerId] = {
          total: 0,
          count: 0
        };
      }
      customerOrders[order.customerId].total += parseFloat(order.billAmount || 0);
      customerOrders[order.customerId].count += 1;
    });
    
    // Find top customers
    const topCustomers = Object.entries(customerOrders)
      .map(([customerId, data]) => {
        const customer = customers.find(c => c.id == customerId) || { name: 'Unknown' };
        return {
          name: customer.name,
          amount: data.total,
          orders: data.count
        };
      })
      .sort((a, b) => b.amount - a.amount)
      .slice(0, 5);
    
    // Get popular items
    const itemCounts = {};
    orders.forEach(order => {
      if (order.items) {
        const itemIds = order.items.split(',');
        itemIds.forEach(itemId => {
          itemId = itemId.trim();
          if (!itemCounts[itemId]) {
            itemCounts[itemId] = 0;
          }
          itemCounts[itemId] += 1;
        });
      }
    });
    
    const popularItems = Object.entries(itemCounts)
      .map(([itemId, count]) => {
        const item = items.find(i => i.id == itemId) || { name: `Item#${itemId}` };
        return { name: item.name, count };
      })
      .sort((a, b) => b.count - a.count)
      .slice(0, 5);
    
    return {
      totalRevenue,
      totalPayments,
      totalOrders: orders.length,
      topCustomers,
      popularItems,
      orders,
      payments
    };
  } catch (error) {
    console.error('Error in getSalesReport:', error);
    throw new Error('Failed to generate sales report.');
  }
}
