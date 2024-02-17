var fs = require('fs');
var fetch = require('node-fetch');
var path = require("path");
var xlsx = require("xlsx");
const _undrscr = require('underscore');
const JSONData=require('./MicronInvoiceMigrationData.json');
const params = require("./paramsList.json")

async function validateData(){
    let errorInvoice='';
    let invCount=0;
try{
    var currentdate = new Date();
    console.log("Validation has started. Please wait...\n Current Time:"+currentdate.getHours()+":"+currentdate.getMinutes()+":"+currentdate.getSeconds()+":"+currentdate.getMilliseconds());
    let headers = { "Authorization": "Bearer " + params.token }
    const method = 'GET'

    const filePath = path.resolve(__dirname, "InvoiceDataMigrationOutput.xlsx");
    const workbook = xlsx.readFile(filePath, {cellDates: true});
    
    //T2 Validation data to be inserted in sheet
    const worksheet = xlsx.utils.sheet_to_json(workbook.Sheets['Invoice Validation Result'], {raw: false, cellDates: true, dateNF:'mm/dd/yyyy'});  //validationresults

        let rowCount=0;                          //it will hold individual row in excel sheet and help to push in individual row
        for(let i=0;i<JSONData.length;i++){      //looping through each PO payload in JSON format in listPOMigration file      
            let counter=0;                       //counter=0 for new PO payload and will increment on line count
            invCount++;
            let partyList=JSONData[i].privacyGroup.replaceAll('-',',');                    //patyList will hold privacy group (like a006-a002-a001-)

            let invoiceAPIURL = params.baseURL.PPE + "/api/invoices/" + JSONData[i].invNumber + "?participants="+partyList+"&fromBlockchain=false";
            let responseT2Invoice = await fetch(invoiceAPIURL, { method, headers }).then(res => res.json());
            errorInvoice=JSONData[i].invNumber;
            let T2headerCID = JSONData[i].cid;
            if(responseT2Invoice.success){
            while(counter<responseT2Invoice.data.result[0].lineItems.length){  //iterating loop for line items in PO
                let cosmosRslt=responseT2Invoice.data.result[0];
                let lineArray=_undrscr.filter(JSONData[i].lineItems,function(arr){ if(arr.line==cosmosRslt.lineItems[counter].line){ return arr; } })
                let productArray=(lineArray[0] && lineArray[0].product)?lineArray[0].product:"";                
                let cosmosProduct=(cosmosRslt.lineItems[counter] && cosmosRslt.lineItems[counter].product)?cosmosRslt.lineItems[counter].product:"";
                
                worksheet[rowCount]={
                        "T2 Invoice": JSONData[i].invNumber,
                        "T2 CreatedBy": (JSONData[i] && JSONData[i].createdBy)? JSONData[i].createdBy : "T2 CreatedBy unavailable",
                        "T2 InvType": (JSONData[i] && JSONData[i].invType)? JSONData[i].invType : "T2 Header invType unavailable",
                        "T2 Ref" : (JSONData[i] && JSONData[i].ref)? JSONData[i].ref.length : "T2 Ref unavailable",
                        "T2 PaymentTerms" : (JSONData[i] && JSONData[i].paymentTerms)?'1' : "T2 PaymentTerms unavailable", 
                        "T2 Prices" : (JSONData[i] && JSONData[i].prices)? JSONData[i].prices.length: "T2 Prices unavailable", 
                        "T2 Currency": (JSONData[i] && JSONData[i].currency)? '1' : "T2 Currency unavailable",
                        "T2 DTM length" :(JSONData[i] && JSONData[i].dtm)? JSONData[i].dtm.length : "T2 dtm unavailable",
                        "T2 Parties" : (JSONData[i] && JSONData[i].parties)? JSONData[i].parties.length : " T2 Parties unavailable",
                        "T2 Status" : (JSONData[i] && JSONData[i].status)? JSONData[i].status.length : "T2 Ref unavailable",
                        "T2 Invoice Type": (JSONData[i] && JSONData[i].ref)?_undrscr.isEmpty(JSONData[i].ref.filter(e=>e.idQualf=="ZF"))?"Type unavailable":JSONData[i].ref.filter(e=>e.idQualf=="ZF")[0].id:"Type unavailable",
                        "T2 Invoice Line#": _undrscr.isEmpty(lineArray)?"Line# unavailable":lineArray[0].line,
                        "T2 BPN": productArray!=""? _undrscr.isEmpty(productArray.filter(e=>e.productQualf=='BP'))?"No BPN":productArray.filter(e=>e.productQualf=='BP')[0].value : "No BPN",   
                        "T2 MPN": productArray!=""?  _undrscr.isEmpty(productArray.filter(e=>e.productQualf=='VP'))?"No MPN":productArray.filter(e=>e.productQualf=='VP')[0].value : "No MPN", 
                        "T2 Qty" : (lineArray[0] && lineArray[0].qty)?_undrscr.isEmpty(lineArray[0].qty.filter(e=>e.type=="BILLED_QTY"))?"No BILLED_QTY":lineArray[0].qty.filter(e=>e.type=="BILLED_QTY")[0].value:"No Quantity",
                        "T2 Price" : (lineArray[0] && lineArray[0].prices)? _undrscr.isEmpty(lineArray[0].prices.filter(e=>e.type=="UNIT_PRICE"))?"No UNIT_PRICE":lineArray[0].prices.filter(e=>e.type=="UNIT_PRICE")[0].value:"No PRICE",
                        "T2 CID check" : (lineArray[0] && lineArray[0].cid && T2headerCID)?((T2headerCID===lineArray[0].cid)?"T2 CID MATCH" : '1'):" T2 CID MISS",
                        "":"",
                        "COSMOS Invoice" : cosmosRslt.invNumber,
                        "COSMOS CreatedBy" : (cosmosRslt && cosmosRslt.createdBy)?  cosmosRslt.createdBy : "CreatedBy unavailable",
                        "COSMOS InvType" : (cosmosRslt && cosmosRslt.invType)?  cosmosRslt.invType : "InvType unavailable",
                        "COSMOS REF" : (cosmosRslt && cosmosRslt.ref)?  cosmosRslt.ref.length : "Ref unavailable ",
                        "COSMOS PaymentTerms" : (cosmosRslt && cosmosRslt.paymentTerms)?  '1' : "PaymentTerms unavailable ",
                        "COSMOS Prices" : (cosmosRslt && cosmosRslt.prices)?  cosmosRslt.prices.length : "Prices unavailable ",
                        "COSMOS Currency" : (cosmosRslt && cosmosRslt.currency)?  '1' : "Currency unavailable ",
                        "COSMOS DTM length" : (cosmosRslt && cosmosRslt.dtm)?  cosmosRslt.dtm.length : "DTM unavailable ",
                        "COSMOS Parties" : (cosmosRslt && cosmosRslt.parties)?  cosmosRslt.parties.length : "Parties unavailable ",
                        "COSMOS Status" : (cosmosRslt && cosmosRslt.status)?  cosmosRslt.status.length : "Status unavailable ",
                        "COSMOS Invoice Type" : (cosmosRslt && cosmosRslt.ref)?_undrscr.isEmpty(cosmosRslt.ref.filter(e=>e.idQualf=="ZF"))?"Type unavailable":cosmosRslt.ref.filter(e=>e.idQualf=="ZF")[0].id:"Type unavailable",
                        "COSMOS Invoice Line#" : (cosmosRslt.lineItems[counter] && cosmosRslt.lineItems[counter].line)?cosmosRslt.lineItems[counter].line:"Line# unavailable",
                        "COSMOS BPN": cosmosProduct!=""? _undrscr.isEmpty(cosmosProduct.filter(e=>e.productQualf=="BP"))?"No BPN":cosmosProduct.filter(e=>e.productQualf=="BP")[0].value:"No BPN", 
                        "COSMOS MPN": cosmosProduct!=""? _undrscr.isEmpty(cosmosProduct.filter(e=>e.productQualf=="VP"))?"No MPN":cosmosProduct.filter(e=>e.productQualf=="VP")[0].value:"No MPN",  
                        "COSMOS Qty" : (cosmosRslt.lineItems[counter] && cosmosRslt.lineItems[counter].qty)?_undrscr.isEmpty(cosmosRslt.lineItems[counter].qty.filter(e=>e.type=="BILLED_QTY"))?"No BILLED_QTY":cosmosRslt.lineItems[counter].qty.filter(e=>e.type=="BILLED_QTY")[0].value:"No Quantity",
                        "COSMOS Price" : (cosmosRslt.lineItems[counter] && cosmosRslt.lineItems[counter].prices)? _undrscr.isEmpty(cosmosRslt.lineItems[counter].prices.filter(e=>e.type=="UNIT_PRICE"))?"No UNIT_PRICE":cosmosRslt.lineItems[counter].prices.filter(e=>e.type=="UNIT_PRICE")[0].value:"No Price",                        
                        " ":" ",
                        "Match Invoice": (JSONData[i].invNumber==cosmosRslt.invNumber) ? 1 : 0,
                        "Match CreatedBy" : ((JSONData[i] && JSONData[i].createdBy)? JSONData[i].createdBy : "T2 CreatedBy unavailable")==((cosmosRslt && cosmosRslt.createdBy)?  cosmosRslt.createdBy : "CreatedBy unavailable")?1:0,
                        "Match InvType" : ((JSONData[i] && JSONData[i].invType)? JSONData[i].invType : "T2 Header invType unavailable")==((cosmosRslt && cosmosRslt.invType)?  cosmosRslt.invType : "InvType unavailable")?1:0,
                        "Match REF" : ((JSONData[i] && JSONData[i].ref)? JSONData[i].ref.length : "T2 Ref unavailable")==((cosmosRslt && cosmosRslt.ref)?  cosmosRslt.ref.length : "Ref unavailable ")?1:0,
                        "Match PaymentTerms" : ((JSONData[i] && JSONData[i].paymentTerms)?'1' : "T2 PaymentTerms unavailable")==((cosmosRslt && cosmosRslt.paymentTerms)?  '1' : "PaymentTerms unavailable ")?1:0,
                        "Match Prices" : ((JSONData[i] && JSONData[i].prices)? JSONData[i].prices.length: "T2 Prices unavailable")==((cosmosRslt && cosmosRslt.prices)?  cosmosRslt.prices.length : "Prices unavailable ")?1:0,
                        "Match Currency" : ((JSONData[i] && JSONData[i].currency)? '1' : "T2 Currency unavailable")==((cosmosRslt && cosmosRslt.currency)?  '1' : "Currency unavailable ")?1:0,
                        "Match DTM length" : ((JSONData[i] && JSONData[i].dtm)? JSONData[i].dtm.length : " T2 dtm unavailable")==((cosmosRslt && cosmosRslt.dtm)?  cosmosRslt.dtm.length : "DTM unavailable ")?1:0,
                        "Match Parties" : ((JSONData[i] && JSONData[i].parties)? JSONData[i].parties.length : " T2 Parties unavailable")==((cosmosRslt && cosmosRslt.parties)?  cosmosRslt.parties.length : "Parties unavailable ")?1:0,
                        "Match Status" : ((JSONData[i] && JSONData[i].status)? JSONData[i].status.length : "T2 Ref unavailable")==((cosmosRslt && cosmosRslt.status)?  cosmosRslt.status.length : "Status unavailable ")?1:0,
                        "Match Invoice Type": ((JSONData[i] && JSONData[i].ref)?_undrscr.isEmpty(JSONData[i].ref.filter(e=>e.idQualf=="ZF"))?"Type unavailable":JSONData[i].ref.filter(e=>e.idQualf=="ZF")[0].id:"Type unavailable")==((cosmosRslt && cosmosRslt.ref)?_undrscr.isEmpty(cosmosRslt.ref.filter(e=>e.idQualf=="ZF"))?"Type unavailable":cosmosRslt.ref.filter(e=>e.idQualf=="ZF")[0].id:"Type unavailable")? 1:0,
                        "Match Line#": (_undrscr.isEmpty(lineArray)?"Line# unavailable":lineArray[0].line)==((cosmosRslt.lineItems[counter] && cosmosRslt.lineItems[counter].line)?cosmosRslt.lineItems[counter].line:"Line# unavailable")? 1:0,
                        "Match BPN": (productArray!=""? _undrscr.isEmpty(productArray.filter(e=>e.productQualf=='BP'))?"No BPN":productArray.filter(e=>e.productQualf=='BP')[0].value : "No BPN")==(cosmosProduct!=""? _undrscr.isEmpty(cosmosProduct.filter(e=>e.productQualf=="BP"))?"No BPN":cosmosProduct.filter(e=>e.productQualf=="BP")[0].value:"No BPN")? 1:0,
                        "Match MPN": (productArray!=""?  _undrscr.isEmpty(productArray.filter(e=>e.productQualf=='VP'))?"No MPN":productArray.filter(e=>e.productQualf=='VP')[0].value : "No MPN")==(cosmosProduct!=""? _undrscr.isEmpty(cosmosProduct.filter(e=>e.productQualf=="VP"))?"No MPN":cosmosProduct.filter(e=>e.productQualf=="VP")[0].value:"No MPN")? 1:0,
                        "Match Qty": ((lineArray[0] && lineArray[0].qty)?_undrscr.isEmpty(lineArray[0].qty.filter(e=>e.type=="BILLED_QTY"))?"No BILLED_QTY":lineArray[0].qty.filter(e=>e.type=="BILLED_QTY")[0].value:"No Quantity")===((cosmosRslt.lineItems[counter] && cosmosRslt.lineItems[counter].qty)?_undrscr.isEmpty(cosmosRslt.lineItems[counter].qty.filter(e=>e.type=="BILLED_QTY"))?"No BILLED_QTY":cosmosRslt.lineItems[counter].qty.filter(e=>e.type=="BILLED_QTY")[0].value:"No Quantity")? 1:0,
                        "Match Price": ((lineArray[0] && lineArray[0].prices)? _undrscr.isEmpty(lineArray[0].prices.filter(e=>e.type=="UNIT_PRICE"))?"No UNIT_PRICE":lineArray[0].prices.filter(e=>e.type=="UNIT_PRICE")[0].value:"No Price")===((cosmosRslt.lineItems[counter] && cosmosRslt.lineItems[counter].prices)? _undrscr.isEmpty(cosmosRslt.lineItems[counter].prices.filter(e=>e.type=="UNIT_PRICE"))?"No UNIT_PRICE":cosmosRslt.lineItems[counter].prices.filter(e=>e.type=="UNIT_PRICE")[0].value:"No Price")? 1:0                       
                        };
                    counter++;
                    rowCount++;
                }
            }
            else{
                worksheet[rowCount]={
                        "T2 Invoice": JSONData[i].invNumber,
                        "T2 CreatedBy": (JSONData[i] && JSONData[i].createdBy)? JSONData[i].createdBy : "T2 CreatedBy unavailable",
                        "T2 InvType": (JSONData[i] && JSONData[i].invType)? JSONData[i].invType : "T2 Header invType unavailable",
                        "T2 Ref" : (JSONData[i] && JSONData[i].ref)? JSONData[i].ref.length : "T2 Ref unavailable",
                        "T2 PaymentTerms" : (JSONData[i] && JSONData[i].paymentTerms)?'1' : "T2 PaymentTerms unavailable", 
                        "T2 Prices" : (JSONData[i] && JSONData[i].prices)? JSONData[i].prices.length: "T2 Prices unavailable", 
                        "T2 Currency": (JSONData[i] && JSONData[i].currency)? '1' : "T2 Currency unavailable",
                        "T2 DTM length" :(JSONData[i] && JSONData[i].dtm)? JSONData[i].dtm.length : " T2 dtm unavailable",
                        "T2 Parties" : (JSONData[i] && JSONData[i].parties)? JSONData[i].parties.length : " T2 Parties unavailable",
                        "T2 Status" : (JSONData[i] && JSONData[i].status)? JSONData[i].status.length : "T2 Ref unavailable",
                        "T2 Invoice Type": (JSONData[i] && JSONData[i].ref)?_undrscr.isEmpty(JSONData[i].ref.filter(e=>e.idQualf=="ZF"))?"Type unavailable":JSONData[i].ref.filter(e=>e.idQualf=="ZF")[0].id:"Type unavailable",
                        "T2 Invoice Line#": (JSONData[i].lineItems[counter] && JSONData[i].lineItems[counter].line)?JSONData[i].lineItems[counter].line:"Line# unavailable",
                        "T2 BPN": (JSONData[i].lineItems[counter] && JSONData[i].lineItems[counter].product)?_undrscr.isEmpty(JSONData[i].lineItems[counter].product.filter(e=>e.productQualf=="BP"))?"No BPN":JSONData[i].lineItems[counter].product.filter(e=>e.productQualf=="BP")[0].value:"No BPN",
                        "T2 MPN": (JSONData[i].lineItems[counter] && JSONData[i].lineItems[counter].product)?_undrscr.isEmpty(JSONData[i].lineItems[counter].product.filter(e=>e.productQualf=="VP"))?"No MPN":JSONData[i].lineItems[counter].product.filter(e=>e.productQualf=="VP")[0].value:"No MPN",
                        "T2 Qty" : (JSONData[i].lineItems[counter] && JSONData[i].lineItems[counter].qty)?_undrscr.isEmpty(JSONData[i].lineItems[counter].qty.filter(e=>e.type=='BILLED_QTY'))?"No BILLED_QTY":JSONData[i].lineItems[counter].qty.filter(e=>e.type=='BILLED_QTY')[0].value:"No Quantity",
                        "T2 Price" : (JSONData[i].lineItems[counter] && JSONData[i].lineItems[counter].prices)?_undrscr.isEmpty(JSONData[i].lineItems[counter].prices.filter(e=>e.type=="UNIT_PRICE"))?"No UNIT_PRICE":JSONData[i].lineItems[counter].prices.filter(e=>e.type=="UNIT_PRICE")[0].value:"No Price",
                        "T2 CID check" : (JSONData[i] && JSONData[i].lineItems[counter] && JSONData[i].lineItems[counter].cid && T2headerCID)?((T2headerCID===JSONData[i].lineItems[counter].cid)?"T2 CID MATCH" : '1'):" T2 CID MISS",
                        "":"",
                        "COSMOS Invoice" : "NOT AVAILABLE",
                        "COSMOS CreatedBy" : "NOT AVAILABLE",
                        "COSMOS InvType" :"NOT AVAILABLE",
                        "COSMOS REF" : "NOT AVAILABLE",
                        "COSMOS PaymentTerms" :"NOT AVAILABLE",
                        "COSMOS Prices" : "NOT AVAILABLE",
                        "COSMOS Currency" :"NOT AVAILABLE",
                        "COSMOS DTM length" :"NOT AVAILABLE",
                        "COSMOS Parties" : "NOT AVAILABLE",
                        "COSMOS Status" : "NOT AVAILABLE",
                        "COSMOS Invoice Type": "NOT AVAILABLE",
                        "COSMOS Invoice Line#" : "NOT AVAILABLE",
                        "COSMOS BPN": "NOT AVAILABLE",
                        "COSMOS MPN": "NOT AVAILABLE",
                        "COSMOS Qty" : "NOT AVAILABLE",
                        "COSMOS Price" : "NOT AVAILABLE",                       
                        " ":" ",                        
                        "Match Invoice": 0,
                        "Match CreatedBy" :0,
                        "Match InvType" : 0,
                        "Match REF" : 0,
                        "Match PaymentTerms" : 0,
                        "Match Prices" : 0,
                        "Match Currency" : 0,
                        "Match DTM length" : 0,
                        "Match Parties" : 0,
                        "Match Status" : 0,
                        "Match Invoice Type":0,
                        "Match Line#": 0,
                        "Match BPN": 0,
                        "Match MPN": 0,
                        "Match Qty": 0,
                        "Match Price": 0
                    };
                rowCount++;
                counter++;
            }
        }
    //Save the workbook
     xlsx.utils.sheet_add_json(workbook.Sheets["Invoice Validation Result"], worksheet)
     xlsx.writeFile(workbook, 'InvoiceDataMigrationOutput.xlsx');
     console.log("Total Invoice:"+invCount+"\nTotal Lines:"+rowCount)
     let currentdateEnd = new Date();
     console.log("Validation has completed @ "+currentdateEnd.getHours()+":"+currentdateEnd.getMinutes()+":"+currentdateEnd.getSeconds()+":"+currentdateEnd.getMilliseconds());
}
catch(e){
    console.log("\n Error in Invoice#"+errorInvoice);
    console.log('Error:',e)
}
}

validateData();