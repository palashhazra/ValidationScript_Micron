var fs = require('fs');
var fetch = require('node-fetch');
var path = require("path");
var xlsx = require("xlsx");
const _undrscr = require('underscore');
const JSONData=require('./MicronShipmentMigrationData.json');
const params = require("./paramsList.json")
//const partner = require("./partnerList.json")

async function ValidateShipment(){
    let errorShipment='';
    let ShipmentCount=0;
try {
    var currentdate = new Date();
    console.log("Validation has started. Please wait...\n Current Time:"+currentdate.getHours()+":"+currentdate.getMinutes()+":"+currentdate.getSeconds()+":"+currentdate.getMilliseconds());
    let headers = { "Authorization": "Bearer " + params.token }
    const method = 'GET'

    const filePath = path.resolve(__dirname, "ShipmentDataMigrationOutput.xlsx");
    const workbook = xlsx.readFile(filePath, {cellDates: true});

    //Selecting the worksheet where to insert data
    const worksheet = xlsx.utils.sheet_to_json(workbook.Sheets['Shipment Validation Result'], {raw: false, cellDates: true, dateNF:'mm/dd/yyyy'});

        let rowCount=0;  
        for(let j=0;j<JSONData.length;j++){      //loop to iterate on each lines of given Shipment list
            let counter=0;
            ShipmentCount++;
            let participants=JSONData[j].privacyGroup.replaceAll("-",",");
            let shipmentAPIURL = params.baseURL.PROD + "/api/shipments/" + JSONData[j].shipmentNumber + "?participants="+participants+"&fromBlockchain=false";
            const responseShipment = await fetch(shipmentAPIURL, { method, headers }).then(res => res.json());
            errorShipment=JSONData[j].shipmentNumber;
            let T2headerCID = (JSONData[j].cid && JSONData[j])?JSONData[j].cid:'';
            let cartonLength=(responseShipment.data.result[0] && responseShipment.data.result[0].handlingUnits)?responseShipment.data.result[0].handlingUnits.huCarton.carton.length:0;

            let lineItemsCountJSON=(JSONData[j] && JSONData[j].lineItems)?JSONData[j].lineItems.length:0;
            let lineItemQtyJSON=0;
            for(let i=0;i<lineItemsCountJSON;i++){
                lineItemQtyJSON += JSONData[j].lineItems[i].qty.reduce((acc,currentValue)=>acc+currentValue.value, 0);
            }

            if(responseShipment.data!=null && (responseShipment.data.result[0] && responseShipment.data.result[0].handlingUnits)){
            //looping through total number of cartons            
            while(counter<cartonLength){

                //cartonDetails containing cartonId and associated carton quantity and carton related details
                let cartonDetails=JSONData[j].handlingUnits.huCarton.carton.filter(element=>element.cartonId==responseShipment.data.result[0].handlingUnits.huCarton.carton[counter].cartonId);

                let countQty=_undrscr.reduce(_undrscr.filter(JSONData[j].handlingUnits.huCarton.carton,function(arr){
                            if(arr.contents[0].shipmentLineId==cartonDetails[0].contents[0].shipmentLineId){
                                return arr;
                            }
                        }),function(memo,num){
                            return memo + parseInt(num.contents[0].quantity);
                    },0);

                let lineQtyResponse=0;
                if((responseShipment.data.result[0] && responseShipment.data.result[0].lineItems)){
                    for(let r=0;r<responseShipment.data.result[0].lineItems.length;r++){
                        lineQtyResponse += responseShipment.data.result[0].lineItems[r].qty.reduce((acc,currentValue)=>acc+currentValue.value, 0);
                    }
                }
                let lineCID='';
                for(let l=0;l<JSONData[j].lineItems.length;l++){
                    if(!(JSONData[j].lineItems[l].cid && JSONData[j].lineItems[l])){
                        lineCID='T2 CID MISS';
                    }
                }
                
                let response = [
                    responseShipment.data.result[0].shipmentNumber,
                    (responseShipment.data.result[0] && responseShipment.data.result[0].createdBy)?responseShipment.data.result[0].createdBy:"",
                    (responseShipment.data.result[0] && responseShipment.data.result[0].shipmentType)?responseShipment.data.result[0].shipmentType:"Shipment Type unavailable",
                    (responseShipment.data.result[0] && responseShipment.data.result[0].shipmentDetails)?responseShipment.data.result[0].shipmentDetails.length:"Shipment Details unavailable",
                    (responseShipment.data.result[0] && responseShipment.data.result[0].ref)?responseShipment.data.result[0].ref.length:"ref unavailable",
                    (responseShipment.data.result[0] && responseShipment.data.result[0].coo)?responseShipment.data.result[0].coo:"coo unavailable",
                    (responseShipment.data.result[0] && responseShipment.data.result[0].dtm)?responseShipment.data.result[0].dtm.length:"dtm unavailable",
                    (responseShipment.data.result[0] && responseShipment.data.result[0].parties)?responseShipment.data.result[0].parties.length:"Parties unavailable",
                    (responseShipment.data.result[0] && responseShipment.data.result[0].status)? _undrscr.isEmpty(responseShipment.data.result[0].status.filter(e=>e.type=="SHIP_STATUS"))?"Status unavailable":responseShipment.data.result[0].status.filter(e=>e.type=="SHIP_STATUS")[0].value:"Status unavailable",
                    responseShipment.data.result[0].parties.filter(x=> x.partnQualf=='SF' || x.partnQualf=='VN')[0].partnerId,   
                    _undrscr.isEmpty(cartonDetails)?"No cartons":cartonDetails[0].cartonId,
                    _undrscr.isEmpty(cartonDetails)?"No cartons":cartonDetails[0].contents[0].quantity,
                    (responseShipment.data.result[0] && responseShipment.data.result[0].lineItems)?responseShipment.data.result[0].lineItems.length:0,
                    lineQtyResponse
                    // _undrscr.isEmpty(lineDetails)?"Wrong carton data":(lineDetails[0] && lineDetails[0].qty[0])? lineDetails[0].qty[0].value:"Line Quantity unavailable"
                    ];

                worksheet[rowCount]={
                    "T2 Shipment": JSONData[j].shipmentNumber,
                    "T2 Shipment CreatedBy": (JSONData[j] && JSONData[j].createdBy)?JSONData[j].createdBy:"",
                    "T2 Shipment Type": (JSONData[j] && JSONData[j].shipmentType)?JSONData[j].shipmentType:"Shipment Type unavailable",
                    "T2 Shipment Details": (JSONData[j] && JSONData[j].shipmentDetails)?JSONData[j].shipmentDetails.length:"Shipment Details unavailable",
                    "T2 Ref": (JSONData[j] && JSONData[j].ref)?JSONData[j].ref.length:"ref unavailable",
                    "T2 COO": (JSONData[j] && JSONData[j].coo)? JSONData[j].coo : "coo unavailable",
                    "T2 DTM": (JSONData[j] && JSONData[j].dtm)? JSONData[j].dtm.length:"dtm unavailable",
                    "T2 Parties": (JSONData[j] && JSONData[j].parties)? JSONData[j].parties.length:"Parties unavailable",
                    "T2 Shipment Status": (JSONData[j] && JSONData[j].status)? _undrscr.isEmpty(JSONData[j].status.filter(e=>e.type=="SHIP_STATUS"))?"Status unavailable":JSONData[j].status.filter(e=>e.type=="SHIP_STATUS")[0].value:"Status unavailable",
                    "T2 Supplying Vendor": (JSONData[j].parties.filter(element=>element.partnQualf=='SF' || element.partnQualf=='VN')[0].partnerId)?JSONData[j].parties.filter(element=>element.partnQualf=='SF' || element.partnQualf=='VN')[0].partnerId:0,
                    "T2 Carton ID": (JSONData[j].handlingUnits.huCarton && JSONData[j].handlingUnits.huCarton.carton[counter])?JSONData[j].handlingUnits.huCarton.carton[counter].cartonId:"No cartons",
                    "T2 Quantity": (JSONData[j].handlingUnits.huCarton && JSONData[j].handlingUnits.huCarton.carton[counter])?JSONData[j].handlingUnits.huCarton.carton[counter].contents[0].quantity:"No cartons",
                    "T2 Line Count":lineItemsCountJSON,
                    "T2 Line Quantity": lineItemQtyJSON,
                    "T2 CID check": (lineCID!='')?"T2 CID MISS":_undrscr.isEmpty(JSONData[j].lineItems.filter(x=>x.cid==T2headerCID))?"1":"T2 CID MATCH",
                    "":"",
                    "COSMOS Shipment" : response[0],
                    "COSMOS Shipment CreatedBy":response[1],
                    "COSMOS Shipment Type":response[2],
                    "COSMOS Shipment Details":response[3],
                    "COSMOS Ref":response[4],
                    "COSMOS COO":response[5],
                    "COSMOS DTM":response[6],
                    "COSMOS Parties":response[7],
                    "COSMOS Shipment Status":response[8],
                    "COSMOS Supplying Vendor" : response[9],
                    "COSMOS Carton ID": response[10],
                    "COSMOS Quantity" : response[11],
                    "COSMOS Line Count" : response[12],
                    "COSMOS Line Quantity" : response[13],
                    " ":" ",

                    "Match Shipment": (JSONData[j].shipmentNumber==response[0])?1:0,
                    "Match CreatedBy": ((JSONData[j] && JSONData[j].createdBy)?JSONData[j].createdBy:"")==response[1]?1:0,
                    "Match Shipment Type": ((JSONData[j] && JSONData[j].shipmentType)?JSONData[j].shipmentType:"Shipment Type unavailable")==response[2]?1:0,
                    "Match Shipment Details": ((JSONData[j] && JSONData[j].shipmentDetails)?JSONData[j].shipmentDetails.length:"Shipment Details unavailable")==response[3]?1:0,
                    "Match Ref": ((JSONData[j] && JSONData[j].ref)?JSONData[j].ref.length:"ref unavailable")==response[4]?1:0,
                    "Match COO": ((JSONData[j] && JSONData[j].coo)? JSONData[j].coo : "coo unavailable")==response[5]?1:0,
                    "Match DTM": ((JSONData[j] && JSONData[j].dtm)? JSONData[j].dtm.length:"dtm unavailable")==response[6]?1:0,
                    "Match Parties": ((JSONData[j] && JSONData[j].parties)? JSONData[j].parties.length:"Parties unavailable")==response[7]?1:0,
                    "Match Shipment Status": ((JSONData[j] && JSONData[j].status)? _undrscr.isEmpty(JSONData[j].status.filter(e=>e.type=="SHIP_STATUS"))?"Status unavailable":JSONData[j].status.filter(e=>e.type=="SHIP_STATUS")[0].value:"Status unavailable")==response[8]?1:0,
                    "Match Supplying Vendor": (JSONData[j].parties.filter(element=>element.partnQualf=='SF' || element.partnQualf=='VN')[0].partnerId==response[9])?1:0,
                    "Match Carton ID": ((JSONData[j].handlingUnits.huCarton && JSONData[j].handlingUnits.huCarton.carton[counter])?JSONData[j].handlingUnits.huCarton.carton[counter].cartonId:"No cartons")==response[10]?1:0,
                    "Match Carton Quantity": ((JSONData[j].handlingUnits.huCarton && JSONData[j].handlingUnits.huCarton.carton[counter])?JSONData[j].handlingUnits.huCarton.carton[counter].contents[0].quantity:"No cartons")==response[11]?1:0,
                    "Match Line Count": lineItemsCountJSON==response[12]?1:0,
                    "Match Line Quantity": (lineItemQtyJSON)==response[13]?1:0
                };
                rowCount++;
                counter++;
            }
        }
        else{
            let lineCID='';
                for(let l=0;l<JSONData[j].lineItems.length;l++){
                    if(!(JSONData[j].lineItems[l].cid && JSONData[j].lineItems[l])){
                        lineCID='T2 CID MISS';
                    }
                }

            let lineQtyResponse=0;             
            if((responseShipment.data.result[0] && responseShipment.data.result[0].lineItems)){
                for(let r=0;r<responseShipment.data.result[0].lineItems.length;r++){
                    lineQtyResponse += responseShipment.data.result[0].lineItems[r].qty.reduce((acc,currentValue)=>acc+currentValue.value, 0);
                }
            }

            worksheet[rowCount]={
                "T2 Shipment": JSONData[j].shipmentNumber,
                "T2 Shipment CreatedBy": (JSONData[j] && JSONData[j].createdBy)?JSONData[j].createdBy:"",
                "T2 Shipment Type": (JSONData[j] && JSONData[j].shipmentType)?JSONData[j].shipmentType:"Shipment Type unavailable",
                "T2 Shipment Details": (JSONData[j] && JSONData[j].shipmentDetails)?JSONData[j].shipmentDetails.length:"Shipment Details unavailable",
                "T2 Ref": (JSONData[j] && JSONData[j].ref)?JSONData[j].ref.length:"ref unavailable",
                "T2 COO": (JSONData[j] && JSONData[j].coo)? JSONData[j].coo : "coo unavailable",
                "T2 DTM": (JSONData[j] && JSONData[j].dtm)? JSONData[j].dtm.length:"dtm unavailable",
                "T2 Parties": (JSONData[j] && JSONData[j].parties)? JSONData[j].parties.length:"Parties unavailable",
                "T2 Shipment Status": (JSONData[j] && JSONData[j].status)? _undrscr.isEmpty(JSONData[j].status.filter(e=>e.type=="SHIP_STATUS"))?"Status unavailable":JSONData[j].status.filter(e=>e.type=="SHIP_STATUS")[0].value:"Status unavailable",
                "T2 Supplying Vendor": (JSONData[j] && JSONData[j].parties)?JSONData[j].parties.filter(element=>element.partnQualf=='SF' || element.partnQualf=='VN')[0].partnerId:"Party unavailable",
                "T2 Carton ID": (JSONData[j].handlingUnits && JSONData[j])?JSONData[j].handlingUnits.huCarton.carton[counter].cartonId:"No Cartons",
                "T2 Quantity": (JSONData[j].handlingUnits && JSONData[j])?JSONData[j].handlingUnits.huCarton.carton[counter].contents[0].quantity:"No Cartons",
                "T2 Line Count": lineItemsCountJSON,
                "T2 Line Quantity": lineItemQtyJSON,
                "T2 CID check": (lineCID!='')?"T2 CID MISS":_undrscr.isEmpty(JSONData[j].lineItems.filter(x=>x.cid==T2headerCID))?"1":"T2 CID MATCH",
                "":"",

                "COSMOS Shipment" : responseShipment.data.result[0].shipmentNumber,
                "COSMOS Shipment CreatedBy":(responseShipment.data.result[0] && responseShipment.data.result[0].createdBy)?responseShipment.data.result[0].createdBy:'',
                "COSMOS Shipment Type":(responseShipment.data.result[0] && responseShipment.data.result[0].shipmentType)?responseShipment.data.result[0].shipmentType:'Shipment Type unavailable',
                "COSMOS Shipment Details":(responseShipment.data.result[0] && responseShipment.data.result[0].shipmentDetails)?responseShipment.data.result[0].shipmentDetails.length:'Shipment Details unavailable',
                "COSMOS Ref":(responseShipment.data.result[0] && responseShipment.data.result[0].ref)?responseShipment.data.result[0].ref.length:'ref unavailable',
                "COSMOS COO":(responseShipment.data.result[0] && responseShipment.data.result[0].coo)?responseShipment.data.result[0].coo:'coo unavailable',
                "COSMOS DTM":(responseShipment.data.result[0] && responseShipment.data.result[0].dtm)?responseShipment.data.result[0].dtm.length:'dtm unavailable',
                "COSMOS Parties":(responseShipment.data.result[0] && responseShipment.data.result[0].parties)?responseShipment.data.result[0].parties.length:'Parties unavailable',
                "COSMOS Shipment Status":(responseShipment.data.result[0] && responseShipment.data.result[0].status)? _undrscr.isEmpty(responseShipment.data.result[0].status.filter(e=>e.type=="SHIP_STATUS"))?"Status unavailable":responseShipment.data.result[0].status.filter(e=>e.type=="SHIP_STATUS")[0].value:'Status unavailable',
                "COSMOS Supplying Vendor": _undrscr.isEmpty(responseShipment.data.result[0].parties.filter(x=> x.partnQualf=='SF' || x.partnQualf=='VN')) ? 'Party unavailable': responseShipment.data.result[0].parties.filter(x=> x.partnQualf=='SF' || x.partnQualf=='VN')[0].partnerId,
                "COSMOS Carton ID": "No Cartons",
                "COSMOS Quantity" :  "No Cartons",
                "COSMOS Line Count" : (responseShipment.data.result[0] && responseShipment.data.result[0].lineItems)? responseShipment.data.result[0].lineItems.length:0,
                "COSMOS Line Quantity" : lineQtyResponse,
                " ":" ",

                "Match Shipment": JSONData[j].shipmentNumber==responseShipment.data.result[0].shipmentNumber?1:0,
                "Match CreatedBy": ((JSONData[j] && JSONData[j].createdBy)?JSONData[j].createdBy:"Not Available")==((responseShipment.data.result[0] && responseShipment.data.result[0].createdBy)?responseShipment.data.result[0].createdBy:'Not Available')?1:0,
                "Match Shipment Type": ((JSONData[j] && JSONData[j].shipmentType)?JSONData[j].shipmentType:"Not Available")==((responseShipment.data.result[0] && responseShipment.data.result[0].shipmentType)?responseShipment.data.result[0].shipmentType:'Not Available')?1: 0,
                "Match Shipment Details": ((JSONData[j] && JSONData[j].shipmentDetails)?JSONData[j].shipmentDetails.length:"Not Available")==((responseShipment.data.result[0] && responseShipment.data.result[0].shipmentDetails)?responseShipment.data.result[0].shipmentDetails.length:'Not Available')?1:0,
                "Match Ref": ((JSONData[j] && JSONData[j].ref)?JSONData[j].ref.length:"Not Available")==((responseShipment.data.result[0] && responseShipment.data.result[0].ref)?responseShipment.data.result[0].ref.length:'Not Available')?1: 0,
                "Match COO": ((JSONData[j] && JSONData[j].coo)? JSONData[j].coo : "Not Available")==((responseShipment.data.result[0] && responseShipment.data.result[0].coo)?responseShipment.data.result[0].coo:'Not Available')? 1:0,
                "Match DTM": ((JSONData[j] && JSONData[j].dtm)? JSONData[j].dtm.length:"Not Available")==((responseShipment.data.result[0] && responseShipment.data.result[0].dtm)?responseShipment.data.result[0].dtm.length:'Not Available')?1:0,
                "Match Parties": ((JSONData[j] && JSONData[j].parties)? JSONData[j].parties.length:"Not Available")==((responseShipment.data.result[0] && responseShipment.data.result[0].parties)?responseShipment.data.result[0].parties.length:'Not Available')?1: 0,
                "Match Shipment Status":((JSONData[j] && JSONData[j].status)? _undrscr.isEmpty(JSONData[j].status.filter(e=>e.type=="SHIP_STATUS"))?"Not Available":JSONData[j].status.filter(e=>e.type=="SHIP_STATUS")[0].value:"Not Available")==((responseShipment.data.result[0] && responseShipment.data.result[0].status)? _undrscr.isEmpty(responseShipment.data.result[0].status.filter(e=>e.type=="SHIP_STATUS"))?"Not Available":responseShipment.data.result[0].status.filter(e=>e.type=="SHIP_STATUS")[0].value:'Not Available')?1:0,
                "Match Supplying Vendor": ((JSONData[j] && JSONData[j].parties)?JSONData[j].parties.filter(element=>element.partnQualf=='SF' || element.partnQualf=='VN')[0].partnerId:"Not Available")==(_undrscr.isEmpty(responseShipment.data.result[0].parties.filter(x=> x.partnQualf=='SF' || x.partnQualf=='VN')) ? 'Not Available': responseShipment.data.result[0].parties.filter(x=> x.partnQualf=='SF' || x.partnQualf=='VN')[0].partnerId)? 1:0,
                "Match Carton ID": ((JSONData[j].handlingUnits && JSONData[j])?JSONData[j].handlingUnits.huCarton.carton[counter].cartonId:"No Cartons")=="No Cartons"?1:0,
                "Match Carton Quantity": ((JSONData[j].handlingUnits && JSONData[j])?JSONData[j].handlingUnits.huCarton.carton[counter].contents[0].quantity:"No Cartons")=="No Cartons"?1:0,
                "Match Line Count": lineItemsCountJSON==((responseShipment.data.result[0] && responseShipment.data.result[0].lineItems)? responseShipment.data.result[0].lineItems.length:0)?1:0,
                "Match Line Quantity": lineItemQtyJSON==lineQtyResponse?1:0
            };
            rowCount++;
            counter++;
        }
    }
 // Save the workbook
 xlsx.utils.sheet_add_json(workbook.Sheets["Shipment Validation Result"], worksheet)
 xlsx.writeFile(workbook, 'ShipmentDataMigrationOutput.xlsx');
 console.log("Total Shipment:"+ShipmentCount+"\nTotal Lines:"+rowCount)
 let currentdateEnd = new Date(); 
 console.log("Validation has completed @ "+currentdateEnd.getHours()+":"+currentdateEnd.getMinutes()+":"+currentdateEnd.getSeconds()+":"+currentdateEnd.getMilliseconds());
}
catch (e) {
    console.log("\n Error in Shipment#"+errorShipment);
    console.log('Error:', e.stack);
}
}

ValidateShipment(); 