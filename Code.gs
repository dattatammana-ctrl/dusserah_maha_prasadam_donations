// Google Apps Script – bound to the donation Google Sheet (Extensions > Apps Script)
const SHEET_NAME = "Donations";
const ITEMS_SHEET = "Items";
const ADMIN_PASSWORD = "Dusserah2026";   // checked server-side, never sent to the public page

// Starting list, written to the "Items" tab only the first time (when the tab doesn't exist)
// [Category, Item, Qty, Unit, Total amount ₹]
const SEED = [
 ["Milk products","Ghee",3,"kg",1800],["Milk products","Milk",10,"ltr",800],
 ["Dry Fruits","Cashew nuts",5,"kg",4000],["Dry Fruits","Almond",1,"kg",800],["Dry Fruits","Tarbuz seeds",2,"kg",800],["Dry Fruits","Cashew pieces",1,"kg",700],["Dry Fruits","Kismis",1,"kg",400],
 ["Big Groceries","Rice",100,"kg",8000],["Big Groceries","Oil",45,"kg",7400],["Big Groceries","Curd",60,"kg",4800],["Big Groceries","Cylinders",4,"",4000],["Big Groceries","Appadalu",5,"kg",1000],["Big Groceries","Chilli powder",3,"kg",900],["Big Groceries","Besan",10,"kg",900],["Big Groceries","Sesame powder",3,"kg",750],["Big Groceries","Groundnuts",5,"kg",700],["Big Groceries","Tamarind",3,"kg",450],["Big Groceries","Coriander powder",2,"kg",400],["Big Groceries","Red gram (toor dal)",2.5,"kg",350],["Big Groceries","Red chilli",1,"kg",300],["Big Groceries","Coconut powder",1,"kg",300],
 ["Small Groceries","Atta",10,"kg",450],["Small Groceries","Sugar",10,"kg",450],["Small Groceries","Green peas (dry)",5,"kg",450],["Small Groceries","Suji rava",8,"kg",400],["Small Groceries","Paper napkins",10,"packets",300],["Small Groceries","Jaggery",2,"kg",140],["Small Groceries","Urad gota",1,"kg",130],["Small Groceries","Salt (big)",8,"kg",160],["Small Groceries","Salt (small)",5,"kg",125],["Small Groceries","Green gram",1,"kg",120],
 ["Vegetables","Tomato",30,"kg",1200],["Vegetables","Potato",25,"kg",875],["Vegetables","Brinjal (small)",20,"kg",1000],["Vegetables","Carrot",6,"kg",360],["Vegetables","Bajji mirchi",10,"kg",600],["Vegetables","Green chilli",5,"kg",300],["Vegetables","Cucumber",2,"kg",80],["Vegetables","Keera",4,"kg",160],["Vegetables","Beans",2,"kg",140],["Vegetables","Green peas",2,"kg",240],["Vegetables","Curry leaves",1,"kg",100],["Vegetables","Coriander",50,"bunches",500],["Vegetables","Mint",30,"bunches",300],["Vegetables","Lemon",50,"",250],["Vegetables","Drumsticks",100,"",800],["Vegetables","Bottle gourd",5,"",150],["Vegetables","Radish",10,"",150],["Vegetables","Cauliflower",6,"",240],["Vegetables","Garlic paste",1,"kg",150]
];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(["Timestamp","Name","Tower - Flat","Mobile","Category","Item","Qty","Amount (₹)","Payment mode","Payment proof"]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function items_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ITEMS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(ITEMS_SHEET);
    sh.appendRow(["Category","Item","Qty","Unit","Total amount (₹)"]);
    sh.getRange(2, 1, SEED.length, 5).setValues(SEED);
    sh.setFrozenRows(1);
    sh.getRange("A1:E1").setFontWeight("bold");
  }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const rows = body.rows, now = new Date();
  let url = "";
  if (body.proof) {   // save payment screenshot to Drive (private, owner-only)
    const it = DriveApp.getFoldersByName("OnCloud33 Payment Proofs");
    const folder = it.hasNext() ? it.next() : DriveApp.createFolder("OnCloud33 Payment Proofs");
    const name = rows[0].flat + " - " + rows[0].name + " - " + Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyyMMdd-HHmmss") + ".jpg";
    url = folder.createFile(Utilities.newBlob(Utilities.base64Decode(body.proof), "image/jpeg", name)).getUrl();
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  const sh = sheet_();
  rows.forEach(r => sh.appendRow([now, r.name, r.flat, "'" + r.mobile, r.category, r.item, r.qty, r.amount, r.mode, r.mode === "Cash" ? url : ""]));
  lock.releaseLock();
  return ContentService.createTextOutput("ok");
}

function doGet(e) {
  const v = sheet_().getDataRange().getValues();

  // Admin: full table, only with the correct password
  if (e && e.parameter && e.parameter.action === "admin") {
    if (e.parameter.pw !== ADMIN_PASSWORD) return json_({error: "unauthorized"});
    const tz = Session.getScriptTimeZone();
    return json_({rows: v.map((r, i) => i === 0 ? r : [Utilities.formatDate(new Date(r[0]), tz, "dd-MMM-yyyy HH:mm")].concat(r.slice(1)))});
  }

  // Public: item list (from the Items tab) + per-item pledge totals
  const taken = {};
  v.slice(1).forEach(r => {
    const o = taken[r[5]] = taken[r[5]] || {qty: 0, phys: 0, amt: 0, by: ""};
    const q = parseFloat(r[6]) || 0;
    o.qty += q;
    o.amt += parseFloat(r[7]) || 0;
    if (r[8] === "Physical item") o.phys += q;
    o.by += (o.by ? ", " : "") + r[2];
  });
  const items = items_().getDataRange().getValues().slice(1)
    .filter(r => r[1] !== "" && parseFloat(r[2]) > 0 && parseFloat(r[4]) >= 0);
  return json_({items, taken});
}
