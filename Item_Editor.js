// @ts-nocheck
/*****************************************************
 EEI ITEM EDITOR
*****************************************************/


/*****************************************************
LOAD ITEM TO EDITOR
K17
*****************************************************/

function loadSelectedItemToEditor(){

const sh=QFORM.SHEET();

const qID=
sh.getRange("B7").getValue();

const revision=
sh.getRange("E2").getValue();

const lineNo=
Number(
sh.getRange("B20").getValue()
);

if(!qID){
SpreadsheetApp.getUi()
.alert("Load quotation first");
return;
}

if(!lineNo){
SpreadsheetApp.getUi()
.alert("Enter line number");
return;
}

const item=
getQuotationItemByLine_(
qID,
revision,
lineNo
);

if(!item){

SpreadsheetApp.getUi()
.alert("Item not found");

return;
}

sh.getRange("B21")
.setValue(item.description);

sh.getRange("F21")
.setValue(item.transformerType);

sh.getRange("B22")
.setValue(item.powerKVA);

sh.getRange("F22")
.setValue(item.voltage);

sh.getRange("B23")
.setValue(item.quantity);

sh.getRange("D23")
.setValue(item.unitPrice);

sh.getRange("F23")
.setValue(item.warranty);

sh.getRange("H23")
.setValue(item.delivery);

sh.getRange("B24")
.setValue(item.notes);

SpreadsheetApp.getUi()
.alert("Item loaded");
}



/*****************************************************
SAVE ITEM
K18
*****************************************************/

function saveSelectedItemChanges() {

const sh = QFORM.SHEET();

const qID =
String(
sh.getRange("B7")
.getDisplayValue()
).trim();

const revision =
String(
sh.getRange("E2")
.getDisplayValue()
).trim();

const lineNo =
Number(
sh.getRange("B20")
.getValue()
);

if (!qID) {

SpreadsheetApp
.getUi()
.alert(
"Load quotation first"
);

return;

}

if (!revision) {

SpreadsheetApp
.getUi()
.alert(
"Select revision first"
);

return;

}

if (!lineNo) {

SpreadsheetApp
.getUi()
.alert(
"Line number missing"
);

return;

}


/*****************************************************
READ ITEM EDITOR
*****************************************************/

const description =
String(
sh.getRange("B21")
.getDisplayValue()
).trim();

const transformerType =
String(
sh.getRange("F21")
.getDisplayValue()
).trim();

const powerKVA =
String(
sh.getRange("B22")
.getDisplayValue()
).trim();

const voltage =
String(
sh.getRange("F22")
.getDisplayValue()
).trim();

const quantity =
Number(
sh.getRange("B23")
.getValue()
);

const unitPrice =
Number(
sh.getRange("D23")
.getValue()
);

const warranty =
String(
sh.getRange("F23")
.getDisplayValue()
).trim();

const deliveryTime =
String(
sh.getRange("H23")
.getDisplayValue()
).trim();

const notes =
String(
sh.getRange("B24")
.getDisplayValue()
).trim();


/*****************************************************
VALIDATION
*****************************************************/

if (!description) {

SpreadsheetApp
.getUi()
.alert(
"Description is required"
);

return;

}

if (
isNaN(quantity) ||
quantity <= 0
) {

SpreadsheetApp
.getUi()
.alert(
"Quantity must be greater than zero"
);

return;

}

if (
isNaN(unitPrice) ||
unitPrice <= 0
) {

SpreadsheetApp
.getUi()
.alert(
"Unit price must be greater than zero"
);

return;

}


/*****************************************************
ITEM OBJECT
*****************************************************/

const itemData = {

qID: qID,

revisionNo: revision,

description: description,

transformerType: transformerType,

powerKVA: powerKVA,

voltage: voltage,

quantity: quantity,

unitPrice: unitPrice,

currency:
String(
sh.getRange("B11")
.getDisplayValue()
).trim(),

deliveryTime: deliveryTime,

warranty: warranty,

notes: notes

};


/*****************************************************
CHECK EXISTING
*****************************************************/

const existingItem =
getQuotationItemByLine_(
qID,
revision,
lineNo
);


/*****************************************************
NEW ITEM
*****************************************************/

if (!existingItem) {

const result =
addQuotationItem(
itemData
);

if (
!result ||
!result.success
) {

SpreadsheetApp
.getUi()
.alert(
result.error ||
"Add item failed"
);

return;

}

SpreadsheetApp
.getUi()
.alert(
"New item added"
);

}


/*****************************************************
UPDATE ITEM
*****************************************************/

else {

const items =
getRequiredSheet_(
CONFIG.SHEETS
.QUOTATION_ITEMS
);

const row =
existingItem.row;

const totalPrice =
quantity *
unitPrice;

items
.getRange(
row,
5,
1,
11
)
.setValues([[

description,
transformerType,
powerKVA,
voltage,
quantity,
unitPrice,
totalPrice,

itemData.currency,

deliveryTime,
warranty,
notes

]]);

items
.getRange(
row,
18
)
.setValue(
new Date()
);

items
.getRange(
row,
19
)
.setValue(
getCurrentUserName()
);

items
.getRange(
row,
21
)
.setValue(
getCurrentUserName()
);

items
.getRange(
row,
22
)
.setValue(
new Date()
);

SpreadsheetApp
.getUi()
.alert(
"Item updated"
);

}


/*****************************************************
REFRESH
*****************************************************/

updateQuotationTotals(
qID,
revision
);

loadQuotationItemsToForm_(
qID,
revision
);

refreshQuotationKPIsFromForm();

}


/*****************************************************
DELETE SOFT
K19
*****************************************************/

function deleteSelectedItemSoft(){

const sh=QFORM.SHEET();

const qID=
sh.getRange("B7").getValue();

const revision=
sh.getRange("E2").getValue();

const lineNo=
Number(
sh.getRange("B20")
.getValue()
);

const item=
getQuotationItemByLine_(
qID,
revision,
lineNo
);

if(!item){

SpreadsheetApp
.getUi()
.alert(
"Item not found"
);

return;
}

const result=
SpreadsheetApp
.getUi()
.alert(
"Delete item ?",
SpreadsheetApp
.getUi()
.ButtonSet.YES_NO
);

if(
result!==
SpreadsheetApp
.getUi()
.Button.YES
){
return;
}

const items=
getRequiredSheet_(
CONFIG.SHEETS
.QUOTATION_ITEMS
);

items
.getRange(
item.row,
20
)
.setValue(
"Deleted"
);

items
.getRange(
item.row,
23
)
.setValue(
getCurrentUserName()
);

items
.getRange(
item.row,
24
)
.setValue(
new Date()
);

updateQuotationTotals(
qID,
revision
);

loadQuotationItemsToForm_(
qID,
revision
);

refreshQuotationKPIsFromForm();

SpreadsheetApp
.getUi()
.alert(
"Item deleted"
);

}



/*****************************************************
ARCHIVE QUOTATION
K20
*****************************************************/

function archiveQuotationFromForm(){

const sh=QFORM.SHEET();

const qID=
sh.getRange("B7")
.getValue();

if(!qID){

SpreadsheetApp
.getUi()
.alert(
"Load quotation"
);

return;
}

archiveQuotation_(
qID
);

SpreadsheetApp
.getUi()
.alert(
"Quotation archived"
);

}



/*****************************************************
GET ITEM
*****************************************************/

function getQuotationItemByLine_(
qID,
revision,
lineNo
){

const items=
getRequiredSheet_(
CONFIG.SHEETS
.QUOTATION_ITEMS
);

if(
items.getLastRow()<2
){
return null;
}

const data=
items
.getRange(
2,
1,
items.getLastRow()-1,
items.getLastColumn()
)
.getValues();

for(
let i=0;
i<data.length;
i++
){

const row=
data[i];

if(

row[1]===qID &&

row[2]===revision &&

Number(
row[3]
)===lineNo &&

row[19]!=="Deleted"

){

return{

row:i+2,

description:
row[4],

transformerType:
row[5],

powerKVA:
row[6],

voltage:
row[7],

quantity:
row[8],

unitPrice:
row[9],

delivery:
row[12],

warranty:
row[13],

notes:
row[14]

};

}

}

return null;

}