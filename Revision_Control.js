// @ts-nocheck
/*****************************************************
 EEI REVISION CONTROL
*****************************************************/

function createQuotationRevision(qID){

const lock=LockService.getScriptLock();
lock.waitLock(30000);

try{

const quotation=getQuotationById_(qID);

if(!quotation){
throw new Error("Quotation not found");
}

if(
quotation.recordStatus==="Archived"
){
throw new Error(
"Archived quotation cannot create revisions"
);
}

const oldRevision=
quotation.currentRevision;

const newRevision=
generateNextRevisionNo_(oldRevision);

const revisions=
getRequiredSheet_(
CONFIG.SHEETS.QUOTATION_REVISIONS
);

const items=
getRequiredSheet_(
CONFIG.SHEETS.QUOTATION_ITEMS
);

const terms=
getRequiredSheet_(
CONFIG.SHEETS.QUOTATION_TERMS
);

const now=new Date();
const user=getCurrentUserName();

copyRevisionItems_(
qID,
oldRevision,
newRevision
);

copyRevisionTerms_(
qID,
oldRevision,
newRevision
);

const revisionID=
qID+"-"+newRevision;

revisions
.getRange(
getNextDataRow_(revisions),
1,
1,
20
)
.setValues([[
revisionID,
qID,
newRevision,
CONFIG.QUOTATION_STATUS.DRAFT,
"Revision created",
now,
user,
oldRevision,
"",
0,
0,
0,
0,
"",
"",
"",
"",
"",
true,
""
]]);

const quotations=
getRequiredSheet_(
CONFIG.SHEETS.QUOTATIONS
);

quotations
.getRange(
quotation.row,
5
)
.setValue(
newRevision
);

quotations
.getRange(
quotation.row,
19
)
.setValue(now);

quotations
.getRange(
quotation.row,
20
)
.setValue(user);

logQuotationStatus_(
qID,
newRevision,
oldRevision,
newRevision,
"Revision",
"Revision created"
);

return{
success:true,
revision:newRevision
};

}
catch(err){

SpreadsheetApp
.getUi()
.alert(err.message);

return{
success:false,
error:err.message
};

}
finally{
lock.releaseLock();
}

}


/*****************************************************
COPY ITEMS
*****************************************************/

function copyRevisionItems_(
qID,
oldRevision,
newRevision
){

const sheet=
getRequiredSheet_(
CONFIG.SHEETS.QUOTATION_ITEMS
);

if(sheet.getLastRow()<2){
return;
}

const data=
sheet
.getRange(
2,
1,
sheet.getLastRow()-1,
24
)
.getValues();

const rows=[];

data.forEach(function(r){

const itemQID=r[1];
const itemRevision=r[2];
const itemStatus=r[19]||"Active";

if(
itemQID===qID &&
itemRevision===oldRevision &&
itemStatus!=="Deleted"
){

const lineNo=r[3];

const newItemID=
qID+
"-"+
newRevision+
"-L"+
("00"+lineNo)
.slice(-2);

const row=[...r];

row[0]=newItemID;
row[2]=newRevision;

row[15]=new Date();
row[16]=getCurrentUserName();

row[17]=new Date();
row[18]=getCurrentUserName();

row[19]="Active";

row[20]="";
row[21]="";
row[22]="";
row[23]="";

rows.push(row);

}

});

if(rows.length){

sheet
.getRange(
sheet.getLastRow()+1,
1,
rows.length,
24
)
.setValues(rows);

}

}


/*****************************************************
COPY TERMS
*****************************************************/

function copyRevisionTerms_(
qID,
oldRevision,
newRevision
){

const terms=
getRequiredSheet_(
CONFIG.SHEETS.QUOTATION_TERMS
);

if(terms.getLastRow()<2){
return;
}

const data=
terms
.getRange(
2,
1,
terms.getLastRow()-1,
terms.getLastColumn()
)
.getValues();

for(let i=0;i<data.length;i++){

if(
data[i][1]===qID &&
data[i][2]===oldRevision
){

const row=[...data[i]];

row[2]=newRevision;

row[16]=new Date();
row[17]=getCurrentUserName();

terms
.getRange(
terms.getLastRow()+1,
1,
1,
row.length
)
.setValues([row]);

return;

}

}

}


/*****************************************************
REVISION NUMBER
*****************************************************/

function generateNextRevisionNo_(
currentRevision
){

const n=
parseInt(
currentRevision
.replace("R","")
)||0;

return "R"+
("00"+(n+1))
.slice(-2);

}