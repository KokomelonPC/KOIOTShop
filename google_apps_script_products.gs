const PRODUCT_SPREADSHEET_ID = "1VV11XQZyzsRVpUXa7nipGWvIwHClDz_PDGhuD6ZeHCc";
const PRODUCT_SHEET_NAME = "Item ForSell";
const PRODUCT_ORDER_SHEET_NAME = "ProductOrders";
const PRODUCT_IMAGE_FOLDER_ID = "1Nj1e3sbJqGg-kNqYWeIhYMNbsUTev6X1";

function doGet(e) {
  e = e || { parameter: {} };
  const action = (e.parameter.action || "").trim();

  try {
    if (action === "getProducts") {
      return jsonResponse(getProducts());
    }

    if (action === "getProductOrdersByEmail") {
      return jsonResponse(getProductOrdersByEmail(e.parameter.email, e.parameter.uid));
    }

    if (action === "getProductOrders") {
      return jsonResponse(getProductOrders());
    }

    return jsonResponse({ success: false, message: "Unknown action" });
  } catch (error) {
    return jsonResponse({ success: false, message: error.message });
  }
}

function doPost(e) {
  e = e || { postData: { contents: "{}" } };
  try {
    const body = JSON.parse(e.postData.contents || "{}");
    const action = (body.action || "").trim();

    if (action === "createProduct") {
      return jsonResponse(createProduct(body));
    }

    if (action === "updateProduct") {
      return jsonResponse(updateProduct(body));
    }

    if (action === "createProductOrder") {
      return jsonResponse(createProductOrder(body));
    }

    if (action === "updateProductOrder") {
      return jsonResponse(updateProductOrder(body));
    }

    return jsonResponse({ success: false, message: "Unknown action" });
  } catch (error) {
    return jsonResponse({ success: false, message: error.message });
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getProductSheet() {
  const ss = SpreadsheetApp.openById(PRODUCT_SPREADSHEET_ID);
  let sheet = ss.getSheetByName(PRODUCT_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(PRODUCT_SHEET_NAME);
  }

  ensureProductHeaders(sheet);
  return sheet;
}

function ensureProductHeaders(sheet) {
  const headers = [
    "productId",
    "name",
    "description",
    "price",
    "stock",
    "status",
    "imageUrl",
    "imageFileId",
    "imageName",
    "galleryImageUrls",
    "galleryImageFileIds",
    "specs",
    "condition",
    "warranty",
    "adminEmail",
    "createdAt",
    "updatedAt"
  ];

  const currentLastColumn = Math.max(sheet.getLastColumn(), headers.length);
  const firstRow = sheet.getRange(1, 1, 1, currentLastColumn).getValues()[0];
  const hasHeader = firstRow.some(value => String(value || "").trim() !== "");

  if (!hasHeader) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
    return;
  }

  const currentHeaders = firstRow.map(value => String(value || "").trim()).filter(Boolean);
  const missingHeaders = headers.filter(header => currentHeaders.indexOf(header) === -1);
  if (missingHeaders.length) {
    sheet.getRange(1, currentHeaders.length + 1, 1, missingHeaders.length).setValues([missingHeaders]);
    sheet.autoResizeColumns(1, currentHeaders.length + missingHeaders.length);
  }
}

function getProducts() {
  const sheet = getProductSheet();
  const values = sheet.getDataRange().getValues();

  if (values.length <= 1) {
    return { success: true, products: [] };
  }

  const headers = values[0].map(header => String(header || "").trim());
  const products = values.slice(1)
    .filter(row => row.some(cell => String(cell || "").trim() !== ""))
    .map(rowToObject(headers))
    .filter(product => String(product.status || "active") !== "hidden")
    .sort((a, b) => new Date(b.updatedAt || b.createdAt || 0) - new Date(a.updatedAt || a.createdAt || 0));

  return { success: true, products };
}

function createProduct(body) {
  const sheet = getProductSheet();
  const now = new Date();
  const productId = "PRD-" + Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyyMMdd-HHmmss");
  let image = { imageUrl: "", imageFileId: "" };
  try {
    image = saveProductImage(body, productId);
  } catch (error) {
    Logger.log("Product image upload skipped: " + error.message);
  }
  const gallery = saveProductGalleryImages(body, productId);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(header => String(header || "").trim());
  const next = {
    productId,
    name: body.name || "",
    description: body.description || "",
    price: Number(body.price || 0),
    stock: Number(body.stock || 0),
    status: body.status || "active",
    imageUrl: image.imageUrl,
    imageFileId: image.imageFileId,
    imageName: body.imageName || "",
    galleryImageUrls: gallery.imageUrls.join("\n"),
    galleryImageFileIds: gallery.imageFileIds.join("\n"),
    specs: body.specs || "",
    condition: body.condition || "",
    warranty: body.warranty || "",
    adminEmail: body.adminEmail || "",
    createdAt: now,
    updatedAt: now
  };

  sheet.appendRow(headers.map(header => next[header] ?? ""));

  return {
    success: true,
    productId,
    imageUrl: image.imageUrl
  };
}

function updateProduct(body) {
  const sheet = getProductSheet();
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(header => String(header || "").trim());
  const productIdIndex = headers.indexOf("productId");

  if (!body.productId || productIdIndex === -1) {
    return { success: false, message: "Missing productId" };
  }

  const rowIndex = values.findIndex((row, index) => index > 0 && row[productIdIndex] === body.productId);
  if (rowIndex === -1) {
    return { success: false, message: "Product not found" };
  }

  const current = rowToObject(headers)(values[rowIndex]);
  const image = body.imageBase64 ? saveProductImage(body, body.productId) : {
    imageUrl: current.imageUrl || "",
    imageFileId: current.imageFileId || ""
  };
  const gallery = saveProductGalleryImages(body, body.productId);
  const currentGalleryUrls = current.galleryImageUrls || "";
  const currentGalleryFileIds = current.galleryImageFileIds || "";
  const now = new Date();

  const next = {
    productId: body.productId,
    name: body.name ?? current.name ?? "",
    description: body.description ?? current.description ?? "",
    price: Number(body.price ?? current.price ?? 0),
    stock: Number(body.stock ?? current.stock ?? 0),
    status: body.status ?? current.status ?? "active",
    imageUrl: image.imageUrl,
    imageFileId: image.imageFileId,
    imageName: body.imageName || current.imageName || "",
    galleryImageUrls: [
      currentGalleryUrls,
      body.galleryImageUrls || "",
      gallery.imageUrls.join("\n")
    ].filter(Boolean).join("\n"),
    galleryImageFileIds: [
      currentGalleryFileIds,
      gallery.imageFileIds.join("\n")
    ].filter(Boolean).join("\n"),
    specs: body.specs ?? current.specs ?? "",
    condition: body.condition ?? current.condition ?? "",
    warranty: body.warranty ?? current.warranty ?? "",
    adminEmail: body.adminEmail || current.adminEmail || "",
    createdAt: current.createdAt || now,
    updatedAt: now
  };

  const nextRow = headers.map(header => next[header] ?? "");
  sheet.getRange(rowIndex + 1, 1, 1, nextRow.length).setValues([nextRow]);

  return { success: true, productId: body.productId, imageUrl: image.imageUrl };
}

function saveProductImage(body, productId) {
  if (!body.imageBase64) {
    return { imageUrl: "", imageFileId: "" };
  }

  const folder = DriveApp.getFolderById(PRODUCT_IMAGE_FOLDER_ID);
  const bytes = Utilities.base64Decode(body.imageBase64);
  const extension = getExtensionFromMimeType(body.imageType || "image/jpeg");
  const safeName = sanitizeFileName(body.imageName || productId + extension);
  const fileName = productId + "-" + safeName;
  const blob = Utilities.newBlob(bytes, body.imageType || "image/jpeg", fileName);
  const file = folder.createFile(blob);
return {
    imageUrl: "https://drive.google.com/uc?export=view&id=" + file.getId(),
    imageFileId: file.getId()
  };
}

function saveProductGalleryImages(body, productId) {
  const images = Array.isArray(body.galleryImages) ? body.galleryImages : [];
  const saved = { imageUrls: [], imageFileIds: [] };

  images.forEach((item, index) => {
    if (!item || !item.imageBase64) return;
    try {
      const folder = DriveApp.getFolderById(PRODUCT_IMAGE_FOLDER_ID);
      const bytes = Utilities.base64Decode(item.imageBase64);
      const extension = getExtensionFromMimeType(item.imageType || "image/jpeg");
      const safeName = sanitizeFileName(item.imageName || "gallery-" + (index + 1) + extension);
      const fileName = productId + "-gallery-" + (index + 1) + "-" + safeName;
      const blob = Utilities.newBlob(bytes, item.imageType || "image/jpeg", fileName);
      const file = folder.createFile(blob);
      saved.imageUrls.push("https://drive.google.com/uc?export=view&id=" + file.getId());
      saved.imageFileIds.push(file.getId());
    } catch (error) {
      Logger.log("Product gallery image upload skipped: " + error.message);
    }
  });

  if (body.galleryImageUrls) {
    String(body.galleryImageUrls).split(/\n|,/).map(url => url.trim()).filter(Boolean).forEach(url => saved.imageUrls.push(url));
  }

  return saved;
}

function getExtensionFromMimeType(mimeType) {
  const map = {
    "image/jpeg": ".jpg",
    "image/png": ".png",
    "image/webp": ".webp",
    "image/gif": ".gif"
  };
  return map[mimeType] || ".jpg";
}

function sanitizeFileName(name) {
  return String(name)
    .replace(/[\\/:*?"<>|]/g, "-")
    .replace(/\s+/g, "-")
    .slice(0, 80);
}

function rowToObject(headers) {
  return row => {
    const item = {};
    headers.forEach((header, index) => {
      item[header] = row[index];
    });
    return item;
  };
}

function getProductOrdersByEmail(email, uid) {
  const ss = SpreadsheetApp.openById(PRODUCT_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(PRODUCT_ORDER_SHEET_NAME);

  if (!sheet) {
    return { success: true, orders: [] };
  }

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return { success: true, orders: [] };
  }

  const headers = values[0].map(header => String(header || "").trim());
  const orders = values.slice(1)
    .filter(row => row.some(cell => String(cell || "").trim() !== ""))
    .map(rowToObject(headers))
    .filter(order => {
      const orderEmail = String(order.email || "").toLowerCase();
      const orderUid = String(order.uid || "");
      return (email && orderEmail === String(email).toLowerCase()) || (uid && orderUid === String(uid));
    });

  return { success: true, orders };
}

function getProductOrders() {
  const sheet = getProductOrderSheet();
  const values = sheet.getDataRange().getValues();

  if (values.length <= 1) {
    return { success: true, orders: [] };
  }

  const headers = values[0].map(header => String(header || "").trim());
  const orders = values.slice(1)
    .filter(row => row.some(cell => String(cell || "").trim() !== ""))
    .map(rowToObject(headers))
    .sort((a, b) => new Date(b.updatedAt || b.createdAt || 0) - new Date(a.updatedAt || a.createdAt || 0));

  return { success: true, orders };
}

function getProductOrderSheet() {
  const ss = SpreadsheetApp.openById(PRODUCT_SPREADSHEET_ID);
  let sheet = ss.getSheetByName(PRODUCT_ORDER_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(PRODUCT_ORDER_SHEET_NAME);
  }

  ensureProductOrderHeaders(sheet);
  return sheet;
}

function ensureProductOrderHeaders(sheet) {
  const headers = [
    "orderId",
    "productId",
    "item",
    "quantity",
    "amount",
    "uid",
    "email",
    "customerName",
    "status",
    "adminReply",
    "paymentMethod",
    "contactPhone",
    "contactLine",
    "contactFacebook",
    "note",
    "slipUrl",
    "slipFileId",
    "createdAt",
    "updatedAt"
  ];

  const currentLastColumn = Math.max(sheet.getLastColumn(), headers.length);
  const firstRow = sheet.getRange(1, 1, 1, currentLastColumn).getValues()[0];
  const hasHeader = firstRow.some(value => String(value || "").trim() !== "");

  if (!hasHeader) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
    return;
  }

  const currentHeaders = firstRow.map(value => String(value || "").trim()).filter(Boolean);
  const missingHeaders = headers.filter(header => currentHeaders.indexOf(header) === -1);
  if (missingHeaders.length) {
    sheet.getRange(1, currentHeaders.length + 1, 1, missingHeaders.length).setValues([missingHeaders]);
    sheet.autoResizeColumns(1, currentHeaders.length + missingHeaders.length);
  }
}

function createProductOrder(body) {
  if (!body.productId || !body.item) {
    return { success: false, message: "Missing product data" };
  }

  if (!body.uid && !body.email) {
    return { success: false, message: "Missing customer data" };
  }

  const sheet = getProductOrderSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(header => String(header || "").trim());
  const now = new Date();
  const orderId = "PORDER-" + Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyyMMdd-HHmmss");
  const slip = saveProductPaymentSlip(body, orderId);
  const next = {
    orderId,
    productId: body.productId || "",
    item: body.item || "",
    quantity: Number(body.quantity || 1),
    amount: Number(body.amount || 0),
    uid: body.uid || "",
    email: body.email || "",
    customerName: body.customerName || "",
    status: body.status || "รอตรวจสอบ",
    adminReply: body.adminReply || "",
    paymentMethod: body.paymentMethod || "",
    contactPhone: body.contactPhone || "",
    contactLine: body.contactLine || "",
    contactFacebook: body.contactFacebook || "",
    note: body.note || "",
    slipUrl: slip.slipUrl || "",
    slipFileId: slip.slipFileId || "",
    createdAt: now,
    updatedAt: now
  };

  sheet.appendRow(headers.map(header => next[header] ?? ""));

  return { success: true, orderId };
}

function updateProductOrder(body) {
  const sheet = getProductOrderSheet();
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(header => String(header || "").trim());
  const orderIdIndex = headers.indexOf("orderId");

  if (!body.orderId || orderIdIndex === -1) {
    return { success: false, message: "Missing orderId" };
  }

  const rowIndex = values.findIndex((row, index) => index > 0 && row[orderIdIndex] === body.orderId);
  if (rowIndex === -1) {
    return { success: false, message: "Product order not found" };
  }

  const current = rowToObject(headers)(values[rowIndex]);
  const next = {
    ...current,
    status: body.status ?? current.status ?? "รอตรวจสอบ",
    adminReply: body.adminReply ?? current.adminReply ?? "",
    updatedAt: new Date()
  };

  const nextRow = headers.map(header => next[header] ?? "");
  sheet.getRange(rowIndex + 1, 1, 1, nextRow.length).setValues([nextRow]);

  return { success: true, orderId: body.orderId };
}

function saveProductPaymentSlip(body, orderId) {
  if (!body.image) {
    return { slipUrl: "", slipFileId: "" };
  }

  const folder = DriveApp.getFolderById(PRODUCT_IMAGE_FOLDER_ID);
  const imageData = String(body.image).split(",").pop();
  const bytes = Utilities.base64Decode(imageData);
  const blob = Utilities.newBlob(bytes, "image/jpeg", orderId + "-payment-slip.jpg");
  const file = folder.createFile(blob);

  return {
    slipUrl: "https://drive.google.com/uc?export=view&id=" + file.getId(),
    slipFileId: file.getId()
  };
}
function testGetProducts() {
  Logger.log(JSON.stringify(getProducts()));
}
function testDriveAccess() {
  const folder = DriveApp.getFolderById(PRODUCT_IMAGE_FOLDER_ID);
  Logger.log(folder.getName());
}


