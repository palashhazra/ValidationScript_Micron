var fs = require('fs');
var fetch = require('node-fetch');
var path = require("path");
var xlsx = require("xlsx");
const _undrscr = require('underscore');
const JSONData=require('./MicronPaymentMigrationData.json');
const params = require("./paramsList.json");

async function validateData(){
    let errorPayment='';
    let paymentCount=0;
try{
    var currentdate = new Date();
    console.log("Validation has started. Please wait...\n Current Time:"+currentdate.getHours()+":"+currentdate.getMinutes()+":"+currentdate.getSeconds()+":"+currentdate.getMilliseconds());
    let headers = { "Authorization": "Bearer " + params.token }
    const method = 'GET'

    const filePath = path.resolve(__dirname, "PaymentDataMigrationOutput.xlsx");
    const workbook = xlsx.readFile(filePath, {cellDates: true});
    const sheetNames = workbook.SheetNames;
    //T2 Validation data to be inserted in sheet
    const worksheet = xlsx.utils.sheet_to_json(workbook.Sheets['Payment Validation Result'], {raw: false, cellDates: true, dateNF:'mm/dd/yyyy'});  //validationresults

        let rowCount=0;                          //it will hold individual row in excel sheet and help to push in individual row
        for(let i=0;i<JSONData.length;i++){      //looping through each PO payload in JSON format in listPOMigration file      
            let counter=0;                       //counter=0 for new PO payload and will increment on line count
            paymentCount++;
            let partyList=JSONData[i].privacyGroup.replaceAll('-',','); 

            let paymentAPIURL = params.baseURL.PPE + "/api/payments/" + JSONData[i].paymentNumber + "?participants="+partyList+"&fromBlockchain=false";
            let responseT2Payment = await fetch(paymentAPIURL, { method, headers }).then(res => res.json());
            errorPayment=JSONData[i].paymentNumber;
            let cosmosRslt=responseT2Payment.data.result[0];
            let T2headerCID = JSONData[i].cid;

            if(responseT2Payment.success){
            while(counter<cosmosRslt.lineItems.length){  //iterating loop for line items in PO

                let lineArray=_undrscr.filter(JSONData[i].lineItems,function(arr){ if(arr.line==cosmosRslt.lineItems[counter].line){ return arr; } })
                //let productArray=lineArray[0].product.filter(e=>e.productQualf=="ZM");
                worksheet[rowCount]={
                        "T2 Payment": JSONData[i].paymentNumber,
                        "T2 CreatedBy": JSONData[i].createdBy,
                        "T2 Payment type": (JSONData[i].paymentType && JSONData[i])?JSONData[i].paymentType:"Payment type unavailable",
                        "T2 Payment Status": (JSONData[i].status && JSONData[i])?_undrscr.isEmpty(JSONData[i].status.filter(e=>e.type == "PAYMENT_STATUS"))?"Status unavailable":JSONData[i].status.filter(e=>e.type == "PAYMENT_STATUS")[0].value:"Status unavailable",
                        "T2 Currency":(JSONData[i].currency && JSONData[i])? JSONData[i].currency:"Currency unavailable",
                        "T2 Amount":(JSONData[i].amount && JSONData[i])? JSONData[i].amount.length:"Amount unavailable",
                        "T2 DTM": (JSONData[i].dtm && JSONData[i])?JSONData[i].dtm.length:"dtm unavailable",
                        "T2 Ref": (JSONData[i].ref && JSONData[i])?JSONData[i].ref.length:"ref unavailable",
                        "T2 Parties": (JSONData[i].parties && JSONData[i])?JSONData[i].parties.length:"ref unavailable",
                        "T2 Payment Line#": _undrscr.isEmpty(lineArray)?"Line# unavailable":lineArray[0].line,                        
                        "T2 Price" : _undrscr.isEmpty(lineArray)?"Price unavailable":_undrscr.isEmpty(lineArray[0].amount.filter(e=>e.type=="PAID_AMT"))?"Price unavailable":lineArray[0].amount.filter(e=>e.type=="PAID_AMT")[0].value,
                        "T2 CID check" : (lineArray[0] && lineArray[0].cid && T2headerCID)?((T2headerCID===lineArray[0].cid)?"T2 CID MATCH" : '1'):" T2 CID MISS",
                        "":"",
                        "COSMOS Payment" : cosmosRslt.paymentNumber,
                        "COSMOS CreatedBy": cosmosRslt.createdBy,
                        "COSMOS Payment Type" : (cosmosRslt.paymentType && cosmosRslt)?cosmosRslt.paymentType:"Payment type unavailable",
                        "COSMOS Payment Status" : (cosmosRslt.status && cosmosRslt)?_undrscr.isEmpty(cosmosRslt.status.filter(e=>e.type == "PAYMENT_STATUS"))?"Status unavailable":cosmosRslt.status.filter(e=>e.type == "PAYMENT_STATUS")[0].value:"Status unavailable",
                        "COSMOS Currency":(cosmosRslt.currency && cosmosRslt)? cosmosRslt.currency:"Currency unavailable",
                        "COSMOS Amount":(cosmosRslt.amount && cosmosRslt)?cosmosRslt.amount.length:"Amount unavailable",
                        "COSMOS DTM":(cosmosRslt.dtm && cosmosRslt)?cosmosRslt.dtm.length:"dtm unavailable",
                        "COSMOS Ref":(cosmosRslt.ref && cosmosRslt)?cosmosRslt.ref.length:"ref unavailable",
                        "COSMOS Parties": (cosmosRslt.parties && cosmosRslt)?cosmosRslt.parties.length:"Parties unavailable",
                        "COSMOS Payment Line#" : (cosmosRslt.lineItems[counter] && cosmosRslt.lineItems[counter].line)?cosmosRslt.lineItems[counter].line:"Line# unavailable",
                        "COSMOS Price" : (cosmosRslt.lineItems[counter] && cosmosRslt.lineItems[counter].amount)?_undrscr.isEmpty(cosmosRslt.lineItems[counter].amount.filter(e=>e.type=="PAID_AMT"))?"Price unavailable":cosmosRslt.lineItems[counter].amount.filter(e=>e.type=="PAID_AMT")[0].value:"Price unavailable",                        
                        " ":" ",
                        "Match Payment": (JSONData[i].paymentNumber==cosmosRslt.paymentNumber) ? 1 : 0,
                        "Match CreatedBy": JSONData[i].createdBy==cosmosRslt.createdBy?1:0,
                        "Match Payment Type": ((JSONData[i].paymentType && JSONData[i])?JSONData[i].paymentType:"Payment type unavailable")==((cosmosRslt.paymentType && cosmosRslt)?cosmosRslt.paymentType:"Payment type unavailable")?1:0,
                        "Match Payment Status": ((JSONData[i].status && JSONData[i])?_undrscr.isEmpty(JSONData[i].status.filter(e=>e.type == "PAYMENT_STATUS"))?"Status unavailable":JSONData[i].status.filter(e=>e.type == "PAYMENT_STATUS")[0].value:"Status unavailable")==((cosmosRslt.status && cosmosRslt)?_undrscr.isEmpty(cosmosRslt.status.filter(e=>e.type == "PAYMENT_STATUS"))?"Status unavailable":cosmosRslt.status.filter(e=>e.type == "PAYMENT_STATUS")[0].value:"Status unavailable")? 1:0,
                        "Match Currency": ((JSONData[i].currency && JSONData[i])? JSONData[i].currency:"Currency unavailable")==((cosmosRslt.currency && cosmosRslt)? cosmosRslt.currency:"Currency unavailable")?1:0,
                        "Match Amount": ((JSONData[i].amount && JSONData[i])? JSONData[i].amount.length:"Amount unavailable")==((cosmosRslt.amount && cosmosRslt)?cosmosRslt.amount.length:"Amount unavailable")?1:0,
                        "Match DTM": ((JSONData[i].dtm && JSONData[i])?JSONData[i].dtm.length:"dtm unavailable")==((cosmosRslt.dtm && cosmosRslt)?cosmosRslt.dtm.length:"dtm unavailable")?1:0,
                        "Match Ref": ((JSONData[i].ref && JSONData[i])?JSONData[i].ref.length:"ref unavailable")==((cosmosRslt.ref && cosmosRslt)?cosmosRslt.ref.length:"ref unavailable")?1:0,
                        "Match Parties": ((JSONData[i].parties && JSONData[i])?JSONData[i].parties.length:"ref unavailable")==((cosmosRslt.parties && cosmosRslt)?cosmosRslt.parties.length:"Parties unavailable")?1:0,
                        "Match Payment Line#": (_undrscr.isEmpty(lineArray)?"Line# unavailable":lineArray[0].line)==((cosmosRslt.lineItems[counter] && cosmosRslt.lineItems[counter].line)?cosmosRslt.lineItems[counter].line:"Line# unavailable")? 1:0,
                        "Match Price": (_undrscr.isEmpty(lineArray)?"Price unavailable":_undrscr.isEmpty(lineArray[0].amount.filter(e=>e.type=="PAID_AMT"))?"Price unavailable":lineArray[0].amount.filter(e=>e.type=="PAID_AMT")[0].value)===((cosmosRslt.lineItems[counter] && cosmosRslt.lineItems[counter].amount)?_undrscr.isEmpty(cosmosRslt.lineItems[counter].amount.filter(e=>e.type=="PAID_AMT"))?"Price unavailable":cosmosRslt.lineItems[counter].amount.filter(e=>e.type=="PAID_AMT")[0].value:"Price unavailable")? 1:0                       
                        };
                    counter++;
                    rowCount++;
                }
            }
            else{
                worksheet[rowCount]={
                        "T2 Payment": JSONData[i].paymentNumber,
                        "T2 CreatedBy": JSONData[i].createdBy,
                        "T2 Payment type": (JSONData[i].paymentType && JSONData[i])?JSONData[i].paymentType:"Payment type unavailable",
                        "T2 Payment Status": (JSONData[i].status && JSONData[i])? _undrscr.isEmpty(JSONData[i].status.filter(e=>e.type == "PAYMENT_STATUS"))?"Status unavailable":JSONData[i].status.filter(e=>e.type == "PAYMENT_STATUS")[0].value:"Status unavailable",
                        "T2 Currency":(JSONData[i].currency && JSONData[i])? JSONData[i].currency:"Currency unavailable",
                        "T2 Amount":(JSONData[i].amount && JSONData[i])? JSONData[i].amount.length:"Amount unavailable",
                        "T2 DTM": (JSONData[i].dtm && JSONData[i])?JSONData[i].dtm.length:"dtm unavailable",
                        "T2 Ref": (JSONData[i].ref && JSONData[i])?JSONData[i].ref.length:"ref unavailable",
                        "T2 Parties": (JSONData[i].parties && JSONData[i])?JSONData[i].parties.length:"ref unavailable",
                        "T2 Payment Line#": (JSONData[i].lineItems[counter].line && JSONData[i].lineItems[counter])?JSONData[i].lineItems[counter].line:"Line# unavailable",
                        "T2 Price" : (JSONData[i].lineItems[counter].amount && JSONData[i].lineItems[counter])? _undrscr.isEmpty(JSONData[i].lineItems[counter].amount.filter(e=>e.type=="PAID_AMT"))?"Price unavailable":JSONData[i].lineItems[counter].amount.filter(e=>e.type=="PAID_AMT")[0].value:"Price unavailable",
                        "T2 CID check" : (lineArray[0] && lineArray[0].cid && T2headerCID)?((T2headerCID===lineArray[0].cid)?"T2 CID MATCH" : '1'):" T2 CID MISS",
                        "":"",
                        "COSMOS Payment" : "NOT AVAILABLE",
                        "COSMOS CreatedBy": "NOT AVAILABLE",
                        "COSMOS Payment type": "NOT AVAILABLE",
                        "COSMOS Payment Status": "NOT AVAILABLE",
                        "COSMOS Currency":"NOT AVAILABLE",
                        "COSMOS Amount":"NOT AVAILABLE",
                        "COSMOS DTM": "NOT AVAILABLE",
                        "COSMOS Ref": "NOT AVAILABLE",
                        "COSMOS Parties":"NOT AVAILABLE",
                        "COSMOS Payment Line#" : "NOT AVAILABLE",
                        "COSMOS Price" : "NOT AVAILABLE",                      
                        " ":" ",
                        "Match Payment": 0,
                        "Match CreatedBy":0,
                        "Match Payment Type":0,
                        "Match Payment Status": 0,
                        "Match Currency": 0,
                        "Match Amount": 0,
                        "Match DTM": 0,
                        "Match Ref": 0,
                        "Match Parties": 0,
                        "Match Payment Line#": 0,
                        "Match Price": 0
                    };
                rowCount++;
                counter++;                
            }
        }
    //Save the workbook
     xlsx.utils.sheet_add_json(workbook.Sheets["Payment Validation Result"], worksheet)
     xlsx.writeFile(workbook, 'PaymentDataMigrationOutput.xlsx');
     console.log("Total Payment:"+paymentCount+"\nTotal Lines:"+rowCount)
     let currentdateEnd = new Date();
     console.log("Validation has completed @ "+currentdateEnd.getHours()+":"+currentdateEnd.getMinutes()+":"+currentdateEnd.getSeconds()+":"+currentdateEnd.getMilliseconds());
}
catch(e){
    console.log("\nError in Payment#"+errorPayment);
    console.log('Error:',e)
}
}

validateData();