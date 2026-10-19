// ==========================================
// CONFIGURATION
// ==========================================
const ADMIN_PIN = "1234"; // Change this to your secret Admin Password/PIN

/**
 * Run this function ONCE inside the Apps Script editor 
 * to automatically create the sheets, headers, and sample data.
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Setup Products Sheet
  let prodSheet = ss.getSheetByName("Products");
  if (!prodSheet) {
    prodSheet = ss.insertSheet("Products");
    prodSheet.appendRow(["ID", "Name", "Category", "Price", "Stock", "ImageURL", "Description"]);
    prodSheet.appendRow([
      "P101", 
      "Custom Wooden Engraved Frame", 
      "Frames", 
      599, 
      15, 
      "https://images.unsplash.com/photo-1513519245088-0e12902e5a38?w=600", 
      "Personalized wooden photo frame with custom text engraving."
    ]);
    prodSheet.appendRow([
      "P102", 
      "Retro Polaroid Photo Prints (Set of 20)", 
      "Polaroids", 
      249, 
      40, 
      "https://images.unsplash.com/photo-1526170375885-4d8ecf77b99f?w=600", 
      "High-gloss vintage style mini polaroid prints with custom captions."
    ]);
    prodSheet.appendRow([
      "P103", 
      "Acrylic LED Night Lamp", 
      "LED Gifts", 
      799, 
      8, 
      "https://images.unsplash.com/photo-1507473885765-e6ed057f782c?w=600", 
      "Warm LED glowing acrylic plaque customized with couple names or Spotify code."
    ]);
  }

  // 2. Setup Coupons Sheet
  let couponSheet = ss.getSheetByName("Coupons");
  if (!couponSheet) {
    couponSheet = ss.insertSheet("Coupons");
    couponSheet.appendRow(["Code", "Type", "Value", "MinOrder", "Active"]);
    couponSheet.appendRow(["WELCOME10", "PERCENT", 10, 300, true]);
    couponSheet.appendRow(["GIFT100", "FLAT", 100, 799, true]);
  }

  // 3. Setup Orders Sheet
  let orderSheet = ss.getSheetByName("Orders");
  if (!orderSheet) {
    orderSheet = ss.insertSheet("Orders");
    orderSheet.appendRow([
      "OrderID", "Timestamp", "CustomerName", "Phone", 
      "Address", "CustomNote", "ItemsJSON", "Subtotal", 
      "CouponCode", "Discount", "FinalTotal", "PaymentStatus"
    ]);
  }
}

/**
 * GET Request: Fetches live products and active coupons for the storefront
 */
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const prodSheet = ss.getSheetByName("Products");
  const couponSheet = ss.getSheetByName("Coupons");
  
  const products = getSheetData(prodSheet);
  const coupons = getSheetData(couponSheet);

  return jsonResponse({
    status: "success",
    products: products,
    coupons: coupons
  });
}

/**
 * POST Request: Handles Orders, Stock Updates, Product Creation, and Coupons
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000); // Prevents stock race conditions if two people order at once

  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // --- PUBLIC ACTION: PLACE ORDER & AUTO-DEDUCT STOCK ---
    if (action === "placeOrder") {
      const prodSheet = ss.getSheetByName("Products");
      const orderSheet = ss.getSheetByName("Orders");
      const prodData = prodSheet.getDataRange().getValues();

      // 1. Verify and deduct stock
      const cartItems = body.items; // [{id, name, qty, price}]
      for (let item of cartItems) {
        let found = false;
        for (let i = 1; i < prodData.length; i++) {
          if (String(prodData[i][0]) === String(item.id)) {
            found = true;
            let currentStock = Number(prodData[i][4]);
            if (currentStock < item.qty) {
              return jsonResponse({ 
                status: "error", 
                message: `Insufficient stock for ${item.name}. Available: ${currentStock}` 
              });
            }
            // Deduct stock in sheet
            prodSheet.getRange(i + 1, 5).setValue(currentStock - item.qty);
            break;
          }
        }
        if (!found) {
          return jsonResponse({ status: "error", message: `Product ${item.name} not found.` });
        }
      }

      // 2. Record order in Orders sheet
      const orderId = "ORD-" + Math.floor(100000 + Math.random() * 900000);
      const itemSummary = cartItems.map(i => `${i.name} (x${i.qty})`).join(", ");

      orderSheet.appendRow([
        orderId,
        new Date().toLocaleString("en-IN"),
        body.customerName,
        body.phone,
        body.address,
        body.customNote || "-",
        itemSummary,
        body.subtotal,
        body.couponCode || "NONE",
        body.discount || 0,
        body.finalTotal,
        "Pending on WhatsApp"
      ]);

      return jsonResponse({ status: "success", orderId: orderId });
    }

    // --- ADMIN ACTIONS (Protected by ADMIN_PIN) ---
    if (body.pin !== ADMIN_PIN) {
      return jsonResponse({ status: "error", message: "Invalid Admin PIN!" });
    }

    // 1. Add New Product
    if (action === "addProduct") {
      const prodSheet = ss.getSheetByName("Products");
      const newId = "P" + Math.floor(1000 + Math.random() * 9000);
      prodSheet.appendRow([
        newId,
        body.name,
        body.category,
        Number(body.price),
        Number(body.stock),
        body.imageUrl,
        body.description
      ]);
      return jsonResponse({ status: "success", message: "Product added!", id: newId });
    }

    // 2. Update Product Stock & Price
    if (action === "updateStock") {
      const prodSheet = ss.getSheetByName("Products");
      const data = prodSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(body.id)) {
          prodSheet.getRange(i + 1, 4).setValue(Number(body.price)); // Column D: Price
          prodSheet.getRange(i + 1, 5).setValue(Number(body.stock)); // Column E: Stock
          return jsonResponse({ status: "success", message: "Stock & price updated!" });
        }
      }
      return jsonResponse({ status: "error", message: "Product ID not found." });
    }

    // 3. Delete Product
    if (action === "deleteProduct") {
      const prodSheet = ss.getSheetByName("Products");
      const data = prodSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(body.id)) {
          prodSheet.deleteRow(i + 1);
          return jsonResponse({ status: "success", message: "Product deleted." });
        }
      }
      return jsonResponse({ status: "error", message: "Product not found." });
    }

    // 4. Add Coupon
    if (action === "addCoupon") {
      const couponSheet = ss.getSheetByName("Coupons");
      const cleanCode = String(body.code).trim().toUpperCase();
      couponSheet.appendRow([
        cleanCode,
        body.type, // "PERCENT" or "FLAT"
        Number(body.value),
        Number(body.minOrder || 0),
        true
      ]);
      return jsonResponse({ status: "success", message: "Coupon created!" });
    }

    // 5. Delete Coupon
    if (action === "deleteCoupon") {
      const couponSheet = ss.getSheetByName("Coupons");
      const data = couponSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).toUpperCase() === String(body.code).toUpperCase()) {
          couponSheet.deleteRow(i + 1);
          return jsonResponse({ status: "success", message: "Coupon removed." });
        }
      }
      return jsonResponse({ status: "error", message: "Coupon not found." });
    }

    return jsonResponse({ status: "error", message: "Unknown action." });

  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

// Helper: Convert Sheet Rows to Array of Objects
function getSheetData(sheet) {
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  const headers = rows[0];
  return rows.slice(1).map(row => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

// Helper: Format JSON Output
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
