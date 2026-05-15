function createInitialQuotationTerms(
  qID,
  revisionNo,
  currency
) {

  const sheet =
    getRequiredSheet_(CONFIG.SHEETS.QUOTATION_TERMS);

  const termsID =
    qID + "-" + revisionNo + "-TERMS";

  const now = new Date();
  const user = getCurrentUserName();

  const rowData = [[
    termsID,
    qID,
    revisionNo,
    currency || CONFIG.CURRENCY.EGP,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    "",
    "",
    30,
    now,
    user,
    now,
    user,
    ""
  ]];

  const nextRow =
    sheet.getLastRow() < 2
      ? 2
      : sheet.getLastRow() + 1;

  sheet.getRange(nextRow, 1, 1, 19)
    .setValues(rowData);

  return termsID;
}









function getQuotationTerms_(
  qID,
  revisionNo
) {

  const sheet =
    getRequiredSheet_(CONFIG.SHEETS.QUOTATION_TERMS);

  if (sheet.getLastRow() < 2) return null;

  const data =
    sheet.getRange(
      2,
      1,
      sheet.getLastRow() - 1,
      19
    ).getValues();

  for (let i = 0; i < data.length; i++) {

    if (
      data[i][1] === qID &&
      data[i][2] === revisionNo
    ) {

      return {
        row: i + 2,
        termsID: data[i][0],
        qID: data[i][1],
        revisionNo: data[i][2],
        currency: data[i][3],
        subTotal: Number(data[i][4]) || 0,
        discountPercent: normalizePercent_(data[i][5]),
        vatPercent: normalizePercent_(data[i][6]),
        advancePercent: normalizePercent_(data[i][7]),
        advanceAmount: Number(data[i][8]) || 0,
        remainingAmount: Number(data[i][9]) || 0,
        grandTotal: Number(data[i][10]) || 0,
        deliveryTerms: data[i][11],
        paymentTerms: data[i][12],
        validityDays: data[i][13],
        notes: data[i][18]
      };
    }
  }

  return null;
}





function normalizePercent_(value) {

  let n = Number(value) || 0;

  if (n > 0 && n < 1) {
    n = n * 100;
  }

  return n;
}



function updateQuotationCommercialTerms(data) {

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {

    if (!data) {
      throw new Error("Missing commercial terms data.");
    }

    if (!data.qID) {
      throw new Error("Quotation ID is required.");
    }

    const quotation =
      getQuotationById_(data.qID);

    if (!quotation) {
      throw new Error("Quotation not found.");
    }

    if (!canEditQuotation_(quotation.status)) {
      throw new Error(
        "Commercial terms cannot be edited while status is: " +
        quotation.status
      );
    }

    const revisionNo =
      data.revisionNo || quotation.currentRevision;

    let terms =
      getQuotationTerms_(data.qID, revisionNo);

    if (!terms) {
      createInitialQuotationTerms(
        data.qID,
        revisionNo,
        data.currency || quotation.currency || CONFIG.CURRENCY.EGP
      );

      terms =
        getQuotationTerms_(data.qID, revisionNo);
    }

    validateCommercialTerms_(data);

    const sheet =
      getRequiredSheet_(CONFIG.SHEETS.QUOTATION_TERMS);

    const oldValue =
      JSON.stringify({
        currency: terms.currency,
        discountPercent: terms.discountPercent,
        vatPercent: terms.vatPercent,
        advancePercent: terms.advancePercent,
        deliveryTerms: terms.deliveryTerms,
        paymentTerms: terms.paymentTerms,
        validityDays: terms.validityDays
      });

    const currency =
      data.currency || terms.currency || CONFIG.CURRENCY.EGP;

    const discountPercent =
      normalizePercent_(
        data.discountPercent !== undefined
          ? data.discountPercent
          : terms.discountPercent
      );

    const vatPercent =
      normalizePercent_(
        data.vatPercent !== undefined
          ? data.vatPercent
          : terms.vatPercent
      );

    const advancePercent =
      normalizePercent_(
        data.advancePercent !== undefined
          ? data.advancePercent
          : terms.advancePercent
      );

    const deliveryTerms =
      data.deliveryTerms !== undefined
        ? data.deliveryTerms
        : terms.deliveryTerms;

    const paymentTerms =
      data.paymentTerms !== undefined
        ? data.paymentTerms
        : terms.paymentTerms;

    const validityDays =
      data.validityDays !== undefined
        ? data.validityDays
        : terms.validityDays;

    const notes =
      data.notes !== undefined
        ? data.notes
        : terms.notes;

    sheet.getRange(terms.row, 4).setValue(currency);
    sheet.getRange(terms.row, 6).setValue(discountPercent / 100);
    sheet.getRange(terms.row, 7).setValue(vatPercent / 100);
    sheet.getRange(terms.row, 8).setValue(advancePercent / 100);
    sheet.getRange(terms.row, 12).setValue(deliveryTerms);
    sheet.getRange(terms.row, 13).setValue(paymentTerms);
    sheet.getRange(terms.row, 14).setValue(validityDays);
    sheet.getRange(terms.row, 17).setValue(new Date());
    sheet.getRange(terms.row, 18).setValue(getCurrentUserName());
    sheet.getRange(terms.row, 19).setValue(notes);

    updateQuotationTotals(
      data.qID,
      revisionNo
    );

    if (typeof logAction === "function") {

      logAction({
        module: "Quotations",
        action: "UPDATE_TERMS",
        recordID: data.qID + "-" + revisionNo,
        field: "Commercial Terms",
        oldValue: oldValue,
        newValue: JSON.stringify({
          currency: currency,
          discountPercent: discountPercent,
          vatPercent: vatPercent,
          advancePercent: advancePercent,
          deliveryTerms: deliveryTerms,
          paymentTerms: paymentTerms,
          validityDays: validityDays
        }),
        notes: "Quotation commercial terms updated"
      });

    }

    return {
      success: true,
      qID: data.qID,
      revisionNo: revisionNo
    };

  } catch (err) {

    SpreadsheetApp.getUi().alert(
      "Update Commercial Terms Error: " + err.message
    );

    Logger.log(err);

    return {
      success: false,
      error: err.message
    };

  } finally {

    lock.releaseLock();

  }
}









function validateCommercialTerms_(data) {

  if (
    data.discountPercent !== undefined &&
    normalizePercent_(data.discountPercent) < 0
  ) {
    throw new Error("Discount percent cannot be negative.");
  }

  if (
    data.vatPercent !== undefined &&
    normalizePercent_(data.vatPercent) < 0
  ) {
    throw new Error("VAT percent cannot be negative.");
  }

  if (
    data.advancePercent !== undefined &&
    (
      normalizePercent_(data.advancePercent) < 0 ||
      normalizePercent_(data.advancePercent) > 100
    )
  ) {
    throw new Error("Advance percent must be between 0 and 100.");
  }

  if (
    data.validityDays !== undefined &&
    Number(data.validityDays) < 0
  ) {
    throw new Error("Validity days cannot be negative.");
  }
}




function testUpdateQuotationCommercialTerms() {

  const result =
    updateQuotationCommercialTerms({

      qID: "Q-0001",

      revisionNo: "R00",

      currency: CONFIG.CURRENCY.EGP,

      discountPercent: 5,

      vatPercent: 14,

      advancePercent: 30,

      deliveryTerms: "Ex-Works EEI Factory",

      paymentTerms: "30% advance payment, 70% before delivery",

      validityDays: 30,

      notes: "Commercial terms test"

    });

  Logger.log(result);
}






