var fs = require('fs');
var fetch = require('node-fetch');
var path = require("path");
var xlsx = require("xlsx");
const _undrscr = require('underscore');
const JSONData=require('./MicronGRMigrationData.json');
const params = require("./paramsList.json")
const grtype = require("./GRType.json")

async function validateData(){
    let errorGR='';
    let grCount=0;
try{
    var currentdate = new Date();
    console.log("Validation has started. Please wait...\n Current Time:"+currentdate.getHours()+":"+currentdate.getMinutes()+":"+currentdate.getSeconds()+":"+currentdate.getMilliseconds());
    let headers = { "Authorization": "Bearer " + params.token }
    const method = 'GET'

    const filePath = path.resolve(__dirname, "GRDataMigrationOutput.xlsx");
    const workbook = xlsx.readFile(filePath, {cellDates: true});
    const sheetNames = workbook.SheetNames;
    //T2 Validation data to be inserted in sheet
    const worksheet = xlsx.utils.sheet_to_json(workbook.Sheets['GR Validation Result'], {raw: false, cellDates: true, dateNF:'mm/dd/yyyy'});  //validationresults

        let rowCount=0;                          //it will hold individual row in excel sheet and help to push in individual row
        for(let i=0;i<JSONData.length;i++){      //looping through each PO payload in JSON format in listPOMigration file      
            let counter=0;                       //counter=0 for new PO payload and will increment on line count
            grCount++;
            let partyList=JSONData[i].privacyGroup.replaceAll('-',',');      //to fetch & store all the party codes for dynamic URL

            let grAPIURL = params.baseURL.PPE + "/api/goodsreceipts/" + JSONData[i].grNumber + "?participants="+partyList+"&fromBlockchain=false";
            let responseT2GR = await fetch(grAPIURL, { method, headers }).then(res => res.json());
            errorGR=JSONData[i].grNumber;
            let T2headerCID = JSONData[i].cid;
            if(responseT2GR.success){
            while(counter<responseT2GR.data.result[0].lineItems.length){  //iterating loop for line items in PO
                let cosmosRslt=responseT2GR.data.result[0];
                let lineArray=_undrscr.filter(JSONData[i].lineItems,function(arr){ if(arr.line==cosmosRslt.lineItems[counter].line){ return arr; } })
                
                //let cartonQty=_undrscr.reduce(JSONData[i].handlingUnits,function (x, y) { return x + parseInt(y.quantity); }, 0);
                //let cosmosCartonQty=_undrscr.reduce(responseT2GR.data.result[0].handlingUnits,function (x, y) { return x + parseInt(y.quantity); }, 0);
                worksheet[rowCount]={
                        "T2 GR": JSONData[i].grNumber,
                        "T2 Created By" : (JSONData[i] && JSONData[i].createdBy)? JSONData[i].createdBy : "T2 CreatedBy unavailable",
                        "T2 Ref length" :(JSONData[i] && JSONData[i].ref)? JSONData[i].ref.length : " T2 Ref unavailable",
                        "T2 DTM length" :(JSONData[i] && JSONData[i].dtm)? JSONData[i].dtm.length : " T2 dtm unavailable",
                        "T2 GR type": (JSONData[i] && JSONData[i].ref)?_undrscr.isEmpty(JSONData[i].ref.filter(e=>e.idQualf=="ZF"))?"No GR type": JSONData[i].ref.filter(e=>e.idQualf=="ZF")[0].id : "Ref unavailable",
                        "T2 GR Line#": _undrscr.isEmpty(lineArray)?"Line unavailable" : lineArray[0].line,
                        "T2 VPN": (lineArray[0] && lineArray[0].product)?_undrscr.isEmpty(lineArray[0].product.filter(e=>e.productQualf=="VP"))?"No VPN":lineArray[0].product.filter(e=>e.productQualf=="VP")[0].value:"Product unavailable",   //[2].productQualf=="ZM"?lineArray[0].product[2].value:"Invalid MSPN",
                        "T2 Qty" : (lineArray[0] && lineArray[0].qty)?_undrscr.isEmpty(lineArray[0].qty.filter(e=>e.type=="GR_QTY"))?"No GR_QTY":lineArray[0].qty.filter(e=>e.type=="GR_QTY")[0].value:"No Quantity",
                        "T2 CID check" : (lineArray[0] && lineArray[0].cid && T2headerCID)?((T2headerCID===lineArray[0].cid)?"T2 CID MATCH" : '1'):" T2 CID MISS",
                        "":"",
                        "COSMOS GR" : cosmosRslt.grNumber,
                        "COSMOS Created By" : (cosmosRslt && cosmosRslt.createdBy)? cosmosRslt.createdBy : "Cosmos CreatedBy unavailable",
                        "COSMOS GR type" : (cosmosRslt && cosmosRslt.ref)?_undrscr.isEmpty(cosmosRslt.ref.filter(e=>e.idQualf =="ZF"))?"NO GR Type" : cosmosRslt.ref.filter(e=>e.idQualf=="ZF")[0].id : "REF unavailable",
                        "COSMOS Ref length" : (cosmosRslt && cosmosRslt.ref)? cosmosRslt.ref.length : "Ref unavailable",
                        "COSMOS DTM length" : (cosmosRslt && cosmosRslt.dtm)? cosmosRslt.dtm.length : "DTM unavailable",
                        "COSMOS GR Line#" : _undrscr.isEmpty(cosmosRslt.lineItems[counter])?"Line unavailable" :cosmosRslt.lineItems[counter].line, 
                        "COSMOS VPN": (cosmosRslt.lineItems[counter] && cosmosRslt.lineItems[counter].product)?_undrscr.isEmpty(cosmosRslt.lineItems[counter].product.filter(e=>e.productQualf=="VP"))?"No VPN":cosmosRslt.lineItems[counter].product.filter(e=>e.productQualf=="VP")[0].value:"Product unavailable", 
                        "COSMOS Qty": (cosmosRslt.lineItems[counter] && cosmosRslt.lineItems[counter].qty)?_undrscr.isEmpty(cosmosRslt.lineItems[counter].qty.filter(e=>e.type=="GR_QTY"))?"No GR_QTY":cosmosRslt.lineItems[counter].qty.filter(e=>e.type=='GR_QTY')[0].value:"No Quantity",                        
                        " ":" ",
                        "Match GR": (JSONData[i].grNumber==cosmosRslt.grNumber) ? 1 : 0,
                        "Match Created By": ((JSONData[i] && JSONData[i].createdBy)? JSONData[i].createdBy : "T2 CreatedBy unavailable")==((cosmosRslt && cosmosRslt.createdBy)? cosmosRslt.createdBy : "Cosmos CreatedBy unavailable") ? 1 : 0,
                        "Match GR Type": ((JSONData[i] && JSONData[i].ref)?_undrscr.isEmpty(JSONData[i].ref.filter(e=>e.idQualf=="ZF"))?"No GR type": JSONData[i].ref.filter(e=>e.idQualf=="ZF")[0].id : "Ref unavailable")==((cosmosRslt && cosmosRslt.ref)?_undrscr.isEmpty(cosmosRslt.ref.filter(e=>e.idQualf =="ZF"))?"NO GR Type" : cosmosRslt.ref.filter(e=>e.idQualf=="ZF")[0].id : "REF unavailable")? 1:0,
                        "Match Ref length" : ((JSONData[i] && JSONData[i].ref)? JSONData[i].ref.length : " T2 Ref unavailable") == ((cosmosRslt && cosmosRslt.ref)? cosmosRslt.ref.length : "Ref unavailable")? 1:0,
                        "Match DTM length" : ((JSONData[i] && JSONData[i].dtm)? JSONData[i].dtm.length : " T2 dtm unavailable") == ((cosmosRslt && cosmosRslt.dtm)? cosmosRslt.dtm.length : "DTM unavailable")? 1:0,
                        // "Match GR Line#": (_undrscr.isEmpty(lineArray)?"Line unavailable" : _undrscr.isEmpty(cosmosRslt.lineItems[counter])?"Line unavailable" :cosmosRslt.lineItems[counter].line)? 1:0,                        
                        "Match GR Line#": (_undrscr.isEmpty(lineArray)?"Line unavailable" : lineArray[0].line)==(_undrscr.isEmpty(cosmosRslt.lineItems[counter])?"Line unavailable" :cosmosRslt.lineItems[counter].line)? 1:0,                        
                        "Match VPN": ((lineArray[0] && lineArray[0].product)?_undrscr.isEmpty(lineArray[0].product.filter(e=>e.productQualf=="VP"))?"No VPN":lineArray[0].product.filter(e=>e.productQualf=="VP")[0].value:"Product unavailable")==((cosmosRslt.lineItems[counter] && cosmosRslt.lineItems[counter].product)?_undrscr.isEmpty(cosmosRslt.lineItems[counter].product.filter(e=>e.productQualf=="VP"))?"No VPN":cosmosRslt.lineItems[counter].product.filter(e=>e.productQualf=="VP")[0].value:"Product unavailable")? 1:0,
                        "Match Qty": ((lineArray[0] && lineArray[0].qty)?_undrscr.isEmpty(lineArray[0].qty.filter(e=>e.type=="GR_QTY"))?"No Quantity":lineArray[0].qty.filter(e=>e.type=="GR_QTY")[0].value:"Product unavailable")===((cosmosRslt.lineItems[counter] && cosmosRslt.lineItems[counter].qty)?_undrscr.isEmpty(cosmosRslt.lineItems[counter].qty.filter(e=>e.type=="GR_QTY"))?"No Quantity":cosmosRslt.lineItems[counter].qty.filter(e=>e.type=='GR_QTY')[0].value:"Quantity unavailable")? 1:0,
                        };
                    counter++;
                    rowCount++;
                }
            }
            else{
                worksheet[rowCount]={
                        "T2 GR": JSONData[i].grNumber,
                        "T2 Created By": (JSONData[i] && JSONData[i].createdBy)? JSONData[i].createdBy : "T2 CreatedBy unavailable",
                        "T2 GR type": (JSONData[i] && JSONData[i].ref)?_undrscr.isEmpty(JSONData[i].ref.filter(e=>e.idQualf=="ZF"))?"No GR type": JSONData[i].ref.filter(e=>e.idQualf=="ZF")[0].id : "Ref unavailable",
                        "T2 Ref length":(JSONData[i] && JSONData[i].ref)? JSONData[i].ref.length : " T2 Ref unavailable",
                        "T2 DTM length":(JSONData[i] && JSONData[i].dtm)? JSONData[i].dtm.length : " T2 dtm unavailable",
                        "T2 GR Line#": JSONData[i].lineItems[counter].line, 
                        "T2 VPN": (JSONData[i].lineItems[counter] && JSONData[i].lineItems[counter].product)?_undrscr.isEmpty(JSONData[i].lineItems[counter].product.filter(e=>e.productQualf=="VP"))?"No VPN":JSONData[i].lineItems[counter].product.filter(e=>e.productQualf=="VP")[0].value:"Product unavailable",
                        "T2 Qty": (JSONData[i].lineItems[counter] && JSONData[i].lineItems[counter].qty)?_undrscr.isEmpty(JSONData[i].lineItems[counter].qty.filter(e=>e.type=="GR_QTY"))?"No Quantity":JSONData[i].lineItems[counter].qty.filter(e=>e.type=="GR_QTY")[0].value:"Quantity unavailable",
                        "T2 CID check" : (JSONData[i] && T2headerCID && JSONData[i].lineItems[counter] && JSONData[i].lineItems[counter].cid)?((T2headerCID===JSONData[i].lineItems[counter].cid)?"T2 CID MATCH" : '1'):" T2 CID MISS",
                        "":"",
                        "COSMOS GR": "NOT AVAILABLE",
                        "COSMOS Created By" : "NOT AVAILABLE",
                        "COSMOS GR type": "NOT AVAILABLE",
                        "COSMOS Ref length" : "NOT AVAILABLE",
                        "COSMOS DTM length" : "NOT AVAILABLE",
                        "COSMOS GR Line#": "NOT AVAILABLE",
                        "COSMOS VPN": "NOT AVAILABLE",
                        "COSMOS Qty": "NOT AVAILABLE",
                        " ":" ",
                        "Match GR": 0,
                        "Match Created By":0,
                        "Match GR Type": 0,
                        "Match Ref length" : 0,
                        "Match DTM length" : 0,
                        "Match GR Line#": 0,                       
                        "Match VPN": 0,
                        "Match Qty": 0,
                    };
                rowCount++;
            }
        }
    //Save the workbook
     xlsx.utils.sheet_add_json(workbook.Sheets["GR Validation Result"], worksheet)     
     xlsx.writeFile(workbook, 'GRDataMigrationOutput.xlsx'); 
     console.log("Total GR:"+grCount+"\n Total Lines:"+rowCount);
     let currentdateEnd = new Date(); 
     console.log("Validation has completed @ "+currentdateEnd.getHours()+":"+currentdateEnd.getMinutes()+":"+currentdateEnd.getSeconds()+":"+currentdateEnd.getMilliseconds());
}
catch(e){
    console.log("\nError in GR#"+errorGR);
    console.log('Error:',e)
}
}

validateData();