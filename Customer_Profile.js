function loadCustomerProfile(){

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const profile = ss.getSheetByName("Customer_Profile");
  const customers = ss.getSheetByName("Customers");
  const quotations = ss.getSheetByName("Quotations");

  const pos = ss.getSheetByName("Purchase_Orders");

  const selected = profile.getRange("B3").getValue();

  if(!selected) return SpreadsheetApp.getUi().alert("Select customer");

  const [customerID] = selected.split("|").map(s => s.trim());

  // 🧹 مسح القديم
  profile.getRange("B6:B10").clearContent();
  profile.getRange("A14:F100").clearContent();
  profile.getRange("H14:M100").clearContent();

  // =========================
  // 👤 بيانات العميل
  // =========================
  const data = customers.getRange(3,1,customers.getLastRow()-2,14).getValues();
  const customer = data.find(r => r[0] === customerID);

  if(!customer) return;

  profile.getRange("B6").setValue(customer[1]); // Name
  profile.getRange("B7").setValue(customer[2]); // Type
  profile.getRange("B8").setValue(customer[3]); // City
  profile.getRange("B9").setValue(customer[7]); // Phone
  profile.getRange("B10").setValue(customer[8]); // Email

  // =========================
  // 📄 Quotations
  // =========================



  if (quotations && quotations.getLastRow() >= 4) {

    const qData = quotations
      .getRange(
        4,
        1,
        quotations.getLastRow() - 3,
        14
      )
      .getValues();

    let totalSales = 0;
    let totalQuotes = 0;
    let wonQuotes = 0;

    const qFiltered = qData
      .filter(r => r[2] === customerID)
      .map(r => {

        const amount =
          Number(r[13]) || 0;

        const status = r[5];

        totalQuotes++;

        if(status === "Won"){

          totalSales += amount;

          wonQuotes++;
        }

        return [

          r[0],
          r[3],
          r[6],
          r[9],
          amount,
          status

        ];
      });

    if(qFiltered.length){

      profile
        .getRange(
          14,
          1,
          qFiltered.length,
          6
        )
        .setValues(qFiltered);
    }

    const conversionRate =
      totalQuotes > 0
        ? (
            (wonQuotes / totalQuotes) * 100
          ).toFixed(1) + "%"
        : "0%";

    profile.getRange("I8")
      .setValue(totalSales);

    profile.getRange("I6")
      .setValue(totalQuotes);

    profile.getRange("I10")
      .setValue(conversionRate);

  }


  // =========================
  // 📦 Purchase Orders
  // =========================
  if(pos){

    const poData = pos.getDataRange().getValues();

    const poFiltered = poData
      .filter(r => r[1] === customerID)
      .map(r => [
        r[0],
        r[2],
        r[3],
        r[4],
        r[5],
        r[6]
      ]);

    if(poFiltered.length){
      profile.getRange(14,8,poFiltered.length,6).setValues(poFiltered);
    }
  }
}




function editCustomerFromProfile(){

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const profile = ss.getSheetByName("Customer_Profile");
  const form = ss.getSheetByName("Register_Customer");
  const db = ss.getSheetByName("Customers");

  const selected = profile.getRange("B3").getValue();

  if(!selected) return SpreadsheetApp.getUi().alert("Select customer");

  const [customerID] = selected.split("|").map(s => s.trim());

  // 📥 هات كل بيانات العميل من الداتا بيز
  const data = db.getRange(3,1,db.getLastRow()-2,14).getValues();

  const customer = data.find(r => r[0] === customerID);

  if(!customer) {
    SpreadsheetApp.getUi().alert("Customer not found");
    return;
  }

    clearCustomerForm();

  // =========================
  // 🧾 Fill Form بالكامل
  // =========================

  form.getRange("B4").setValue(customer[1]); // Name
  form.getRange("B5").setValue(customer[2]); // Type
  form.getRange("B6").setValue(customer[3]); // City
  form.getRange("B7").setValue(customer[4]); // Address
  form.getRange("B11").setValue(customer[5]); // Contact
  form.getRange("B8").setValue(customer[6]); // Referred
  form.getRange("B12").setValue(customer[7]); // Phone
  form.getRange("B13").setValue(customer[8]); // Email
  form.getRange("B15").setValue(customer[11]); // Notes

  // 🆔 مهم جدًا (Edit Mode)
  form.getRange("D4").setValue(customerID);

  // 🔀 تحويلك للفورم
  ss.setActiveSheet(form);

}



function clearCustomerProfile(){

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Customer_Profile");

  sheet.getRange("B3").clearContent(); // Customer Selector

  // 🧹 مسح بيانات العميل
  sheet.getRange("B6:B10").clearContent();

  // 🧹 مسح الجداول
  sheet.getRange("A14:F100").clearContent();
  sheet.getRange("H14:M100").clearContent();

  // 🧹 مسح KPIs
  sheet.getRange("I6:I10").clearContent();

}




function buildCustomerProfileSelector() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const db = ss.getSheetByName(CONFIG.SHEETS.CUSTOMERS);
  const profile = ss.getSheetByName("Customer_Profile");

  if (!db || !profile || db.getLastRow() < 3) return;

  const data = db.getRange(3, 1, db.getLastRow() - 2, 11).getValues();

  const list = data
    .filter(r =
      r[0] &&
      r[1] &&
      r[10] === "Active"
    ) 
    .map(r => `${r[0]} | ${r[1]}`);

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true)
    .setAllowInvalid(false)
    .build();

  profile.getRange("B3").setDataValidation(rule);
}
