var fs = require('fs');
var fetch = require('node-fetch');
var path = require("path");
var xlsx = require("xlsx");
const _undrscr = require('underscore');
const JSONData=require('./MicronPOMigrationData.json');
const params = require("./paramsList.json")

async function validateData(){
    let errorOrder='';
    let POCount=0;
try{   
    var currentdate = new Date();
    console.log("Validation has started. Please wait...\n Current Time:"+currentdate.getHours()+":"+currentdate.getMinutes()+":"+currentdate.getSeconds()+":"+currentdate.getMilliseconds());
    let headers = { "Authorization": "Bearer " + params.token }   
    const method = 'GET'
      
    const filePath = path.resolve(__dirname, "PODataMigrationOutput.xlsx");  
    const workbook = xlsx.readFile(filePath, {cellDates: true});
    const sheetNames = workbook.SheetNames;
    //T2 Validation data to be inserted in sheet
    const worksheet = xlsx.utils.sheet_to_json(workbook.Sheets['PO Validation Result'], {raw: false, cellDates: true, dateNF:'mm/dd/yyyy'});  //validationresults
    
        let rowCount=0;                          //it will hold individual row in excel sheet and help to push in individual row
        for(let i=0;i<JSONData.length;i++){      //looping through each PO payload in JSON format in listPOMigration file      
            let counter=0;                       //counter=0 for new PO payload and will increment on line count
            POCount++;
            
            let prvgroupval = JSONData[i].privacyGroup.replaceAll("-", ',');            

            let poAPIURL = params.baseURL.PPE + "/api/purchaseorders/" + JSONData[i].orderNumber + "?participants="+prvgroupval+"&fromBlockchain=false";
            let responseT2PO = await fetch(poAPIURL, { method, headers }).then(res => res.json());
            let cosmosRslt=responseT2PO.data.result[0];
            errorOrder=JSONData[i].orderNumber;
            let T2headerCID = JSONData[i].cid;
            if(responseT2PO.success){
            while(counter<cosmosRslt.lineItems.length){  //iterating loop for line items in PO               
                let lineArray=_undrscr.filter(JSONData[i].lineItems,function(arr){ if(arr.line==cosmosRslt.lineItems[counter].line){ return arr; } })
                if (!_undrscr.isEmpty(lineArray)){
                worksheet[rowCount]={
                        "T2 Order Type":(JSONData[i] && JSONData[i].orderType)?JSONData[i].orderType:"Order Type unavailable",
                        "T2 PO": JSONData[i].orderNumber,   
                        "T2 Currency":(JSONData[i] && JSONData[i].currency)?JSONData[i].currency:"Currency unavailable",
                        "T2 Status Length":(JSONData[i].status && JSONData[i])?JSONData[i].status.length:"Status unavailable",
                        "T2 DTM Length":(JSONData[i] && JSONData[i].dtm)?JSONData[i].dtm.length:"dtm unavailable",
                        "T2 Payment Term": (JSONData[i] && JSONData[i].paymentTerms)?JSONData[i].paymentTerms.terms:"PaymentTerm unavailable",
                        "T2 Incoterm":(JSONData[i] && JSONData[i].incoTerms)?JSONData[i].incoTerms.incoTerms1:"Incoterm unavailable",
                        "T2 Ref":(JSONData[i] && JSONData[i].ref)?JSONData[i].ref.length:"Ref unavailable",
                        "T2 Parties": (JSONData[i] && JSONData[i].parties)?JSONData[i].parties.length:"Parties unavailable",
                        "T2 PO Line#": _undrscr.isEmpty(lineArray)?"Line unavailable" : lineArray[0].line,
                        "T2 MSPN": (lineArray[0].product && lineArray[0])?_undrscr.isEmpty(lineArray[0].product.filter(e=>e.productQualf=="ZM"))?"No MSPN":lineArray[0].product.filter(e=>e.productQualf=="ZM")[0].value : "Product unavailable",
                        "T2 MPN": (lineArray[0].product && lineArray[0])?_undrscr.isEmpty(lineArray[0].product.filter(e=>e.productQualf=="VP"))?"No MPN":lineArray[0].product.filter(e=>e.productQualf=="VP")[0].value : "Product unavailable",
                        "T2 BPN": (lineArray[0].product && lineArray[0])?_undrscr.isEmpty(lineArray[0].product.filter(e=>e.productQualf=="BP"))?"No BPN":lineArray[0].product.filter(e=>e.productQualf=="BP")[0].value : "Product unavailable",
                        "T2 Qty" :(lineArray[0].qty && lineArray[0])?_undrscr.isEmpty(lineArray[0].qty.filter(e=>e.type=='PO_QTY'))?"No Quantity":lineArray[0].qty.filter(e=>e.type=='PO_QTY')[0].value : "Quantity unavailable",
                        "T2 Price": (lineArray[0].prices && lineArray[0])?_undrscr.isEmpty(lineArray[0].prices.filter(e=>e.type=="UNIT_PRICE"))?"No Unit Price":lineArray[0].prices.filter(e=>e.type=="UNIT_PRICE")[0].value : "Price unavailable",
                        "T2 CID check" : (lineArray[0] && lineArray[0].cid && T2headerCID)?((T2headerCID===lineArray[0].cid)?"T2 CID MATCH" : '1'):" T2 CID MISS",
                        "":"",
                        "COSMOS Order Type":(cosmosRslt.orderType && cosmosRslt)?cosmosRslt.orderType:"Order Type unavailable",
                        "COSMOS PO" : cosmosRslt.orderNumber,
                        "COSMOS Currency":(cosmosRslt && cosmosRslt.currency)?cosmosRslt.currency:"Currency unavailable",
                        "COSMOS Status Length":(cosmosRslt && cosmosRslt.status)?cosmosRslt.status.length:"Status unavailable",
                        "COSMOS DTM Length":(cosmosRslt && cosmosRslt.dtm)?cosmosRslt.dtm.length:"dtm unavailable",
                        "COSMOS Payment Term":(cosmosRslt && cosmosRslt.paymentTerms)?cosmosRslt.paymentTerms.terms:"PaymentTerm unavailable",
                        "COSMOS Incoterm":(cosmosRslt && cosmosRslt.incoTerms)?cosmosRslt.incoTerms.incoTerms1:"Incoterm unavailable",
                        "COSMOS Ref":(cosmosRslt && cosmosRslt.ref)?cosmosRslt.ref.length:"Ref unavailable",
                        "COSMOS Parties": (cosmosRslt && cosmosRslt.parties)?cosmosRslt.parties.length:"Parties unavailable",
                        "COSMOS PO Line#" : _undrscr.isEmpty(cosmosRslt.lineItems[counter])?"Line unavailable" : cosmosRslt.lineItems[counter].line,                        
                        "COSMOS MSPN": (cosmosRslt.lineItems[counter].product && cosmosRslt.lineItems[counter])?_undrscr.isEmpty(cosmosRslt.lineItems[counter].product.filter(e=>e.productQualf=="ZM"))?"No MSPN": cosmosRslt.lineItems[counter].product.filter(e=>e.productQualf=='ZM')[0].value :"Product unavailable",
                        "COSMOS MPN": (cosmosRslt.lineItems[counter].product && cosmosRslt.lineItems[counter])?_undrscr.isEmpty(cosmosRslt.lineItems[counter].product.filter(e=>e.productQualf=="VP"))? "No MPN":cosmosRslt.lineItems[counter].product.filter(e=>e.productQualf=='VP')[0].value : "Product unavailable",
                        "COSMOS BPN": (cosmosRslt.lineItems[counter].product && cosmosRslt.lineItems[counter])?_undrscr.isEmpty(cosmosRslt.lineItems[counter].product.filter(e=>e.productQualf=="BP"))? "No BPN":cosmosRslt.lineItems[counter].product.filter(e=>e.productQualf=='BP')[0].value : "Product unavailable",
                        "COSMOS Qty" : (cosmosRslt.lineItems[counter].qty && cosmosRslt.lineItems[counter])?_undrscr.isEmpty(cosmosRslt.lineItems[counter].qty.filter(e=>e.type=="PO_QTY"))?"No Quantity":cosmosRslt.lineItems[counter].qty.filter(e=>e.type=='PO_QTY')[0].value : "Quantity unavailable",
                        "COSMOS Price": (cosmosRslt.lineItems[counter].prices && cosmosRslt.lineItems[counter])?_undrscr.isEmpty(cosmosRslt.lineItems[counter].prices.filter(e=>e.type=='UNIT_PRICE'))?"No Price":cosmosRslt.lineItems[counter].prices.filter(e=>e.type=='UNIT_PRICE')[0].value : "Price unavailable",
                        " ":" ",
                        "Match Order Type":((JSONData[i] && JSONData[i].orderType)?JSONData[i].orderType:"Order Type unavailable")==((cosmosRslt.orderType && cosmosRslt)?cosmosRslt.orderType:"Order Type unavailable")?1:0,
                        "Match PO": (JSONData[i].orderNumber==cosmosRslt.orderNumber)? 1 : 0,
                        "Match Currency": ((JSONData[i] && JSONData[i].currency)?JSONData[i].currency:"Currency unavailable")==((cosmosRslt && cosmosRslt.currency)?cosmosRslt.currency:"Currency unavailable")?1:0,
                        "Match Status": ((JSONData[i].status && JSONData[i])?JSONData[i].status.length:"Status unavailable")==((cosmosRslt && cosmosRslt.status)?cosmosRslt.status.length:"Status unavailable")?1:0,
                        "Match DTM Length": ((JSONData[i] && JSONData[i].dtm)?JSONData[i].dtm.length:"dtm unavailable")==((cosmosRslt && cosmosRslt.dtm)?cosmosRslt.dtm.length:"dtm unavailable")?1:0,
                        "Match Payment Term": ((JSONData[i] && JSONData[i].paymentTerms)?JSONData[i].paymentTerms.terms:"PaymentTerm unavailable")==((cosmosRslt && cosmosRslt.paymentTerms)?cosmosRslt.paymentTerms.terms:"PaymentTerm unavailable")?1:0,
                        "Match Incoterm": ((JSONData[i] && JSONData[i].incoTerms)?JSONData[i].incoTerms.incoTerms1:"Incoterm unavailable")==((cosmosRslt && cosmosRslt.incoTerms)?cosmosRslt.incoTerms.incoTerms1:"Incoterm unavailable")?1:0,
                        "Match Ref":((JSONData[i] && JSONData[i].ref)?JSONData[i].ref.length:"Ref unavailable")==((cosmosRslt && cosmosRslt.ref)?cosmosRslt.ref.length:"Ref unavailable")?1:0,
                        "Match Parties":((JSONData[i] && JSONData[i].parties)?JSONData[i].parties.length:"Parties unavailable")==((cosmosRslt && cosmosRslt.parties)?cosmosRslt.parties.length:"Parties unavailable")?1:0,
                        "Match Line#": (_undrscr.isEmpty(lineArray)?"Line unavailable" : lineArray[0].line==_undrscr.isEmpty(cosmosRslt.lineItems[counter])?"Line unavailable" : cosmosRslt.lineItems[counter].line)?1:0,
                        "Match MSPN": ((lineArray[0].product && lineArray[0])?_undrscr.isEmpty(lineArray[0].product.filter(e=>e.productQualf=="ZM"))?"No MSPN":lineArray[0].product.filter(e=>e.productQualf=="ZM")[0].value : "Product unavailable")==((cosmosRslt.lineItems[counter].product && cosmosRslt.lineItems[counter])?_undrscr.isEmpty(cosmosRslt.lineItems[counter].product.filter(e=>e.productQualf=="ZM"))?"No MSPN": cosmosRslt.lineItems[counter].product.filter(e=>e.productQualf=='ZM')[0].value :"Product unavailable")? 1:0,
                        "Match MPN": ((lineArray[0].product && lineArray[0])?_undrscr.isEmpty(lineArray[0].product.filter(e=>e.productQualf=="VP"))?"No MPN":lineArray[0].product.filter(e=>e.productQualf=="VP")[0].value : "Product unavailable")==((cosmosRslt.lineItems[counter].product && cosmosRslt.lineItems[counter])?_undrscr.isEmpty(cosmosRslt.lineItems[counter].product.filter(e=>e.productQualf=="VP"))? "No MPN":cosmosRslt.lineItems[counter].product.filter(e=>e.productQualf=='VP')[0].value : "Product unavailable")?1:0,
                        "Match BPN": ((lineArray[0].product && lineArray[0])?_undrscr.isEmpty(lineArray[0].product.filter(e=>e.productQualf=="BP"))?"No BPN":lineArray[0].product.filter(e=>e.productQualf=="BP")[0].value : "Product unavailable")==((cosmosRslt.lineItems[counter].product && cosmosRslt.lineItems[counter])?_undrscr.isEmpty(cosmosRslt.lineItems[counter].product.filter(e=>e.productQualf=="BP"))? "No BPN":cosmosRslt.lineItems[counter].product.filter(e=>e.productQualf=='BP')[0].value : "Product unavailable")?1:0,
                        "Match Qty": ((lineArray[0].qty && lineArray[0])?_undrscr.isEmpty(lineArray[0].qty.filter(e=>e.type=='PO_QTY'))?"No Quantity":lineArray[0].qty.filter(e=>e.type=='PO_QTY')[0].value : "Quantity unavailable")===((cosmosRslt.lineItems[counter].qty && cosmosRslt.lineItems[counter])?_undrscr.isEmpty(cosmosRslt.lineItems[counter].qty.filter(e=>e.type=="PO_QTY"))?"No Quantity":cosmosRslt.lineItems[counter].qty.filter(e=>e.type=='PO_QTY')[0].value : "Quantity unavailable")?1:0,
                        "Match Price": ((lineArray[0].prices && lineArray[0])?_undrscr.isEmpty(lineArray[0].prices.filter(e=>e.type=="UNIT_PRICE"))?"No Price":lineArray[0].prices.filter(e=>e.type=="UNIT_PRICE")[0].value : "Price unavailable")===((cosmosRslt.lineItems[counter].prices && cosmosRslt.lineItems[counter])?_undrscr.isEmpty(cosmosRslt.lineItems[counter].prices.filter(e=>e.type=='UNIT_PRICE'))?"No Price":cosmosRslt.lineItems[counter].prices.filter(e=>e.type=='UNIT_PRICE')[0].value : "Price unavailable")?1:0                       
                        };
                    }
                    else{
                        worksheet[rowCount]={
                            "T2 Order Type":JSONData[i].orderType,
                            "T2 PO": JSONData[i].orderNumber,
                            "T2 Currency":JSONData[i].currency,
                            "T2 Status Length":JSONData[i].status.length,
                            "T2 DTM Length":JSONData[i].dtm.length,
                            "T2 Payment Term": JSONData[i].paymentTerms.terms,
                            "T2 Incoterm":JSONData[i].incoTerms.incoTerms1,
                            "T2 Ref":JSONData[i].ref.length,
                            "T2 Parties": JSONData[i].parties.length,
                            "T2 PO Line#":"No Partner Line",
                            "T2 MSPN": "No Partner MSPN",
                            "T2 MPN": "No Partner MPN",
                            "T2 BPN": "No Partner BPN",
                            "T2 Qty" : "No Partner Qty",
                            "T2 Price": "No Partner Price",
                            "T2 CID check" : (lineArray[0] && lineArray[0].cid && T2headerCID)?((T2headerCID===lineArray[0].cid)?"T2 CID MATCH" : '1'):" T2 CID MISS",
                            "":"",
                            "COSMOS Order Type":cosmosRslt.orderType,
                            "COSMOS PO" : cosmosRslt.orderNumber,
                            "COSMOS Currency":cosmosRslt.currency,
                            "COSMOS Status Length":cosmosRslt.status.length,
                            "COSMOS DTM Length":cosmosRslt.dtm.length,
                            "COSMOS Payment Term":cosmosRslt.paymentTerms.terms,
                            "COSMOS Incoterm":cosmosRslt.incoTerms.incoTerms1,
                            "COSMOS Ref": cosmosRslt.ref.length,
                            "COSMOS Parties": cosmosRslt.parties.length,
                            "COSMOS PO Line#" : _undrscr.isEmpty(cosmosRslt.lineItems[counter])?"Line unavailable" : cosmosRslt.lineItems[counter].line,                        
                            "COSMOS MSPN": (cosmosRslt.lineItems[counter].product && cosmosRslt.lineItems[counter])?_undrscr.isEmpty(cosmosRslt.lineItems[counter].product.filter(e=>e.productQualf=="ZM"))?"No MSPN": cosmosRslt.lineItems[counter].product.filter(e=>e.productQualf=='ZM')[0].value :"Product unavailable",
                            "COSMOS MPN": (cosmosRslt.lineItems[counter].product && cosmosRslt.lineItems[counter])?_undrscr.isEmpty(cosmosRslt.lineItems[counter].product.filter(e=>e.productQualf=="VP"))? "No MPN":cosmosRslt.lineItems[counter].product.filter(e=>e.productQualf=='VP')[0].value : "Product unavailable",
                            "COSMOS BPN": (cosmosRslt.lineItems[counter].product && cosmosRslt.lineItems[counter])?_undrscr.isEmpty(cosmosRslt.lineItems[counter].product.filter(e=>e.productQualf=="BP"))? "No BPN":cosmosRslt.lineItems[counter].product.filter(e=>e.productQualf=='BP')[0].value : "Product unavailable",
                            "COSMOS Qty" : (cosmosRslt.lineItems[counter].qty && cosmosRslt.lineItems[counter])?_undrscr.isEmpty(cosmosRslt.lineItems[counter].qty.filter(e=>e.type=="PO_QTY"))?"No Quantity":cosmosRslt.lineItems[counter].qty.filter(e=>e.type=='PO_QTY')[0].value : "Quantity unavailable",
                            "COSMOS Price": (cosmosRslt.lineItems[counter].prices && cosmosRslt.lineItems[counter])?_undrscr.isEmpty(cosmosRslt.lineItems[counter].prices.filter(e=>e.type=='UNIT_PRICE'))?"No Price":cosmosRslt.lineItems[counter].prices.filter(e=>e.type=='UNIT_PRICE')[0].value : "Price unavailable",
                            " ":" ",
                            "Match Order Type":JSONData[i].orderType==cosmosRslt.orderType?1:0,
                            "Match PO": (JSONData[i].orderNumber==cosmosRslt.orderNumber)? 1 : 0,
                            "Match Currency":JSONData[i].currency==cosmosRslt.currency?1:0,
                            "Match Status Length":JSONData[i].status.length==cosmosRslt.status.length?1:0,
                            "Match DTM Length": JSONData[i].dtm.length==cosmosRslt.dtm.length?1:0,
                            "Match Payment Term": JSONData[i].paymentTerms.terms==cosmosRslt.paymentTerms.terms?1:0,
                            "Match Incoterm": JSONData[i].incoTerms.incoTerms1==cosmosRslt.incoTerms.incoTerms1?1:0,
                            "Match Ref": JSONData[i].ref.length==cosmosRslt.ref.length?1:0,
                            "Match Parties":JSONData[i].parties.length==cosmosRslt.parties.length?1:0,
                            "Match Line#": 0,
                            "Match MSPN": 0,
                            "Match MPN": 0,
                            "Match BPN": 0,
                            "Match Qty": 0,
                            "Match Price": 0
                            };
                    }
                    counter++;
                    rowCount++;
                }            
            }
            else{
                worksheet[rowCount]={
                        "T2 Order Type":JSONData[i].orderType,
                        "T2 PO": JSONData[i].orderNumber,
                        "T2 Currency":JSONData[i].currency,
                        "T2 Status Length":JSONData[i].status.length,
                        "T2 DTM Length":JSONData[i].dtm.length,
                        "T2 Payment Term": JSONData[i].paymentTerms.terms,
                        "T2 Incoterm":JSONData[i].incoTerms.incoTerms1,
                        "T2 Ref":JSONData[i].ref.length,
                        "T2 Parties": JSONData[i].parties.length,
                        "T2 PO Line#": _undrscr.isEmpty(lineArray)?"Line unavailable" : JSONData[i].lineItems[counter].line,
                        "T2 MSPN": (JSONData[i].lineItems[counter].product && JSONData[i].lineItems[counter])?_undrscr.isEmpty(JSONData[i].lineItems[counter].product.filter(e=>e.productQualf=="ZM"))?"No MSPN":JSONData[i].lineItems[counter].product.filter(e=>e.productQualf=="ZM")[0].value : "Product unavailable",
                        "T2 MPN": (JSONData[i].lineItems[counter].product && JSONData[i].lineItems[counter])?_undrscr.isEmpty(JSONData[i].lineItems[counter].product.filter(e=>e.productQualf=="VP"))?"No MPN":JSONData[i].lineItems[counter].product.filter(e=>e.productQualf=="VP")[0].value : "Product unavailable",
                        "T2 BPN": (JSONData[i].lineItems[counter].product && JSONData[i].lineItems[counter])?_undrscr.isEmpty(JSONData[i].lineItems[counter].product.filter(e=>e.productQualf=="BP"))?"No BPN":JSONData[i].lineItems[counter].product.filter(e=>e.productQualf=="BP")[0].value : "Product unavailable",
                        "T2 Qty" :(JSONData[i].lineItems[counter].qty && JSONData[i].lineItems[counter])?_undrscr.isEmpty(JSONData[i].lineItems[counter].qty.filter(e=>e.type=='PO_QTY'))?"No Quantity":JSONData[i].lineItems[counter].qty.filter(e=>e.type=='PO_QTY')[0].value : "Quantity unavailable",
                        "T2 Price": (JSONData[i].lineItems[counter].prices && JSONData[i].lineItems[counter])?_undrscr.isEmpty(JSONData[i].lineItems[counter].prices.filter(e=>e.type=="UNIT_PRICE"))?"No Price":JSONData[i].lineItems[counter].prices.filter(e=>e.type=="UNIT_PRICE")[0].value : "Price unavailable",
                        "T2 CID check" : (JSONData[i] && JSONData[i].lineItems[counter] && JSONData[i].lineItems[counter].cid && T2headerCID)?((T2headerCID===lineArray[0].cid)?"T2 CID MATCH" : '1'):" T2 CID MISS",
                        "":"",
                        "COSMOS Order Type":"NOT AVAILABLE",
                        "COSMOS PO" : "NOT AVAILABLE",
                        "COSMOS Currency":"NOT AVAILABLE",
                        "COSMOS Status Length":"NOT AVAILABLE",
                        "COSMOS DTM Length":"NOT AVAILABLE",
                        "COSMOS Payment Term":"NOT AVAILABLE",
                        "COSMOS Incoterm":"NOT AVAILABLE",
                        "COSMOS Ref":"NOT AVAILABLE",
                        "COSMOS Parties": "NOT AVAILABLE",
                        "COSMOS PO Line#" : "NOT AVAILABLE",
                        "COSMOS MSPN": "NOT AVAILABLE",
                        "COSMOS MPN": "NOT AVAILABLE",
                        "COSMOS BPN": "NOT AVAILABLE",
                        "COSMOS Qty" : "NOT AVAILABLE",
                        "COSMOS Price" : "NOT AVAILABLE",                     
                        " ":" ",       
                        "Match Order Type":0,                 
                        "Match PO": 0,
                        "Match Currency":0,
                        "Match Status Length":0,
                        "Match DTM Length": 0,
                        "Match Payment Term":0,
                        "Match Incoterm":0,
                        "Match Ref":0,
                        "Match Parties":0,
                        "Match Line#": 0,
                        "Match MSPN": 0,
                        "Match MPN": 0,
                        "Match BPN": 0,
                        "Match Qty": 0,
                        "Match Price": 0
                    };
                rowCount++;
            }
        }
    //Save the workbook
     xlsx.utils.sheet_add_json(workbook.Sheets["PO Validation Result"], worksheet)     
     xlsx.writeFile(workbook, 'PODataMigrationOutput.xlsx'); 
     console.log("Total PO:"+POCount+"\n Total Lines:"+rowCount)
     let currentdateEnd = new Date(); 
     console.log("Validation has completed @ "+currentdateEnd.getHours()+":"+currentdateEnd.getMinutes()+":"+currentdateEnd.getSeconds()+":"+currentdateEnd.getMilliseconds());
}
catch(e){
    console.log("\n Error in Order#"+errorOrder);
    console.log('Error:',e)
}
}

validateData();