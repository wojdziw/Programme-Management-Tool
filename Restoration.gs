/*
The purpose of this function is to restore the sheet if it has become unusable after a crash of the script or any other reason
*/

function restore() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var gatewaySpecificSheet = ss.getSheets()[1];
  var numSheets = ss.getNumSheets();
  var dataUpdateSheet = ss.getSheetByName('Data update');
  dataUpdateSheet.showSheet();
  
  //DELETES ALL GATEWAY SHEETS AND THE OVERVIEW
  for (var s = numSheets-8; s>=0; s--){
      gatewaySpecificSheet = ss.getSheets()[s];
      ss.setActiveSheet(gatewaySpecificSheet);
      ss.deleteActiveSheet();
      }
  
  //COPIES THE DESIRED SHEETS FROM ANOTHER FILE (https://docs.google.com/a/jaguarlandrover.com/spreadsheet/ccc?key=0Aj0AzxO8qOEodGZiNzlFYkhaeGtaWFphOFB5RVA2dHc#gid=2)
  var sourceSpreadsheet = SpreadsheetApp.openById('0Aj0AzxO8qOEodGZiNzlFYkhaeGtaWFphOFB5RVA2dHc');
  var sourceOverview = sourceSpreadsheet.getSheetByName('Overview')
  var sourceGatewaySheet = sourceSpreadsheet.getSheetByName('PS');
  
  sourceGatewaySheet.copyTo(ss).setName('PS');
  sourceOverview.copyTo(ss).setName('Overview');
  
  //HIDES THE SPREADSHEET RESTORATION SHEET
  ss.getSheetByName("Spreadsheet restoration").hideSheet()
  
  //RUNS THE GATEWAY UPDATER, FEATURES UPDATER AND DELIVERABLES UPDATER TO RESTORE THE PREVIOUS LAYOUT
  restoregateways();
  restorefeaturesndeliverables();
}  
  
  
function restoregateways(){  
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var numSheets = ss.getNumSheets();                                                                                                                 //Registers the number of sheets in the spreadsheet
  
  var gatewaysReferenceSheet = ss.getSheetByName('Gateways reference');                                                                              //Saves the Gateways reference sheet
  var gatewaysReferenceData = gatewaysReferenceSheet.getDataRange();                                                                                 //Saves the Gateways reference data range
  var numGatewaysReferenceRows = gatewaysReferenceSheet.getLastRow();                                                                                //Saves the number of Gateways reference rows
  var numGateways = numGatewaysReferenceRows-4;                                                                                                      //Calculates the number of new gateways
  
  var featuresReferenceSheet = ss.getSheetByName('Features reference');                                                                              //Saves the Features reference sheet
  var featuresReferenceData = featuresReferenceSheet.getDataRange();                                                                                 //Saves the Features reference data range
  var numFeaturesReferenceRows = featuresReferenceSheet.getLastRow();                                                                                //Saves the number of Features reference rows
  var numFeatures = numFeaturesReferenceRows-5;                                                                                                      //Calculates the number of features used in the sheet
  
  var gatewaySpecificSheet = ss.getSheets()[1];                                                                                                      //Saves the Gateway-specific sheet (e.g. PS, PSC, etc.)
  var gatewaySpecificData = gatewaySpecificSheet.getDataRange();                                                                                     //Saves the Gateway-specific data range
  var numGatewaySpecificCols = gatewaySpecificSheet.getLastColumn();                                                                                 //Saves the number of Gateway-specific columns
  var numGatewaySpecificRows = gatewaySpecificSheet.getLastRow();                                                                                    //Saves the number of Gateway-specific rows
  
  var overviewSheet = ss.getSheetByName("Overview");                                                                                                 //Saves the Overview sheet
  var overviewData = ss.getSheetByName("Overview").getDataRange();                                                                                   //Saves the Overview data range
  var numOverviewRows = ss.getSheetByName('Overview').getLastRow();                                                                                  //Saves the number of Overview rows
  var numOverviewCols = ss.getSheetByName('Overview').getLastColumn();                                                                               //Saves the number of Overview columns
  
  var numDeliverablesBefore = 0;                                                                                                                     //Introduces the variable storing the number of deliverables before (calculated later on)
  
  //REMOVES ALL CHARTS FROM THE OVERVIEW
  var overviewCharts = overviewSheet.getCharts();                                                                                                    //Saves the charts within the overview sheet
  for (var i in overviewCharts){overviewSheet.removeChart(overviewCharts[i])};                                                                       //Removes the charts in the overview sheet, one by one
  
  //DELETES ALL GATEWAY SHEETS (EXCEPT THE FIRST ONE)
  for (var s=numSheets-8; s>=2; s--){                                                                                                                //Counts down the number of gateway-specific sheets' indexes
      gatewaySpecificSheet = ss.getSheets()[s];                                                                                                      //Saves the Gateway-specific sheet
      ss.setActiveSheet(gatewaySpecificSheet);                                                                                                       //Sets the current sheet the active one
      ss.deleteActiveSheet();                                                                                                                        //Deletes the active sheete)
      }
    
  //CALCULATES THE NUMBER OF DELIVERABLES IN THE FIRST SHEET
  gatewaySpecificSheet = ss.getSheets()[1];                                                                                                          //Saves the first gateway-specific sheet
  gatewaySpecificData = ss.getSheets()[1].getDataRange();                                                                                            //Saves the first gateway-specific data range
  numGatewaySpecificCols = gatewaySpecificSheet.getLastColumn();                                                                                     //Saves the number of the first gateway spacific sheet's columns
  numGatewaySpecificRows = gatewaySpecificSheet.getLastRow();                                                                                        //Saves the number of the first gateway spacific sheet's rows
  for (var i=19; i<=numGatewaySpecificCols; i++){if (gatewaySpecificData.getCell(2,i).getValue()!==""){numDeliverablesBefore++}};                    //Calculates the number of deliverables in the first gateway specific sheet

  //DELETES ALL DELIVERABLE COLUMNS (EXCEPT THE FIRST ONE)
  if (numDeliverablesBefore>1){gatewaySpecificSheet.deleteColumns(20,numDeliverablesBefore*2-2)};                                                    //Removes all deliverable columns except the first one
  
  //RESETS THE FEATURE STATUSES IN THE FIRST GATEWAY SHEET
  for (var i=0; i<2; i++){                                                                                                                 //Iterates through all the features
      gatewaySpecificData.getCell(i*3+7,19).setValue("Not updated");                                                                                 //Resets all features statuses to Not Updated
      gatewaySpecificData.getCell(i*3+8,19).setValue("");                                                                                            //Resets all features coments
      };
  gatewaySpecificData.getCell(2,19).setValue("Deliverable name");                                                                                    //Sets the first deliverable's name to "Deliverable name"
  gatewaySpecificData.getCell(4,19).setValue("");                                                                                                    //Resets the first deliverable's measure of completeness
  
  //INSERTS NEW GATEWAY SHEETS
  for (var i=1; i<numGateways; i++){ss.setActiveSheet(ss.getSheets()[1]); ss.duplicateActiveSheet()};                                                //Inserts appropriate number of sheets by duplicating the first one
  
  //UPDATES THE NAMES OF THE NEW GATEWAY SHEETS
  for (var i=1; i<=numGateways; i++){                                                                                                                //Iterates through sheets
      ss.setActiveSheet(ss.getSheets()[i]);                                                                                                          //Sets a new active sheet
      ss.renameActiveSheet(gatewaysReferenceData.getCell(4+i,2).getValue());                                                                         //Renames the active sheet
      ss.getSheets()[i].getDataRange().getCell(2,2).setValue(gatewaysReferenceData.getCell(4+i,2).getValue());                                       //Inserts the name of the gateway in the sheet's top left corner
      }
  
  //DELETES ALL GATEWAY ROWS IN THE OVERVIEW (EXCEPT THE FIRST ONE)
  if (numOverviewRows>9){overviewSheet.deleteRows(11,numOverviewRows-9)};                                                                            //Deletes all gateway rows except the first one in the overview
  
  //ADDS APPROPRIATE NUMBER OF ROWS IN THE OVERVIEW
  if (numGateways>1){overviewSheet.insertRowsAfter(10,numGateways*2-2)};                                                                             //Adds an appropriate number of rows

  //UPDATES THE OVERVIEW DATA RANGE
  overviewData = overviewSheet.getRange(1,1,8+numGateways*2,numOverviewCols+1);                                                                      //Saves the new Overview data range
  
  //SHOWS A MESSAGE BOX TO PREVENT CRASHES
  Browser.msgBox("Data update","Data update 90% complete.", Browser.Buttons.OK);                                                                     //Shows a message box, otherwise the script CRASHes
  
  //REMOVES REDUNDANT BORDERS AND SETS NEW ONES
  overviewSheet.getRange(8+numGateways*2,1,1,numOverviewCols+1).setBorder(false,false,false,false,false,false);                                      //Removes redundant borders from the sheet
  overviewData.getCell(9,6).setBorder(true,true,true,true,false,false);                                                                              //Sets new borders
  
  //PREPARES THE NEW GATEWAY ROWS IN THE OVERVIEW
  for (var i=0; i<numGateways; i++){                                                                                                                 //Iterates through the rows in the overview
      
      //COPIES THE FIRST GATEWAY'S FORMATTING AND PASTES IT ELSEWHERE
      var sourceOfFormatting = overviewSheet.getRange(9,6,1,7);                                                                                      //Takes the formatting of the first gateway row
      var destinationOfFormatting = overviewSheet.getRange(9+i*2,6,1,7);                                                                             //Takes the range of each consecutive row
      sourceOfFormatting.copyTo(destinationOfFormatting);                                                                                            //Pastes the formatting to new gateways
      
      //SETS BACKGROUND OF THE GATEWAY ROWS
      if (i%2==0){overviewSheet.getRange(9+i*2, 2, 1, 13).setBackground("#e4effc")};                                                                 //Sets the background colour in the even gateways
      if (i%2==1){overviewSheet.getRange(9+i*2, 2, 1, 13).setBackground("#d9f0ad")};                                                                 //Sets the background colour in the odd gateways
      
      //SETS APPROPRIATE HEIGHT OF THE GATEWAY ROWS
      overviewSheet.setRowHeight(9+2*i, 113);                                                                                                        //Sets appropriate row height
      overviewSheet.setRowHeight(10+2*i, 20);                                                                                                        //Sets appropriate row height
      
      //INSERTS REFERENCES TO THE GATEWAY LOGOS, STATUSES AND DEADLINES
      var gatewayName = gatewaysReferenceData.getCell(5+i,2).getValue();                                                                             //Saves the name of the gateway-specific sheet
      overviewData.getCell(9+i*2,6).setFormula('=\''+gatewayName+'\'!$I$2');                                                                         //Sets a reference to the status cell in the gateway-specific sheet
      overviewData.getCell(9+i*2,12).setFormula('=\''+gatewayName+'\'!$F$2').setBackground("FFFFFF");                                                //Sets a reference to the deadline cell in the gateway specific sheet
          if (gatewayName == "PS"){overviewData.getCell(9+i*2,2)                                                                                     //Inserts the gateway logos to the Overview
              .setFormula('=image("https://lh3.googleusercontent.com/DcNLJopiHr05BsBqtyD6TEz6LJx0Ir77YRFZ4RtPtGKolff9lt3rtSZ4rwbjs4OL4yc7XL8YfxLk78ZYBt3iWeJUdavsdn-c5JOUa_KcUVoRo9puzgGvXf8qSA",1)')};
          if (gatewayName == "PSC"){overviewData.getCell(9+i*2,2)
              .setFormula('=image("https://lh6.googleusercontent.com/p1ZaItR7prRin_VJP6c3WmhKQKqdIfRt48GbTcc9uNxp8ybFmG-SXClX0SSMx5W-LIWl41NYdSZkx7gv2y1rptcs9eP9XfJo88NVMFRkIu0qqeorE8R1FEgApw",1)')};
          if (gatewayName == "PTCC"){overviewData.getCell(9+i*2,2)
              .setFormula('=image("https://lh3.googleusercontent.com/BbsC4-aulyAF-iUPXXuWaSINTfRkVnvRIpjc1o7OyLixum3HEsU43kzm79tZyBlfjrFcNWlMIgOmIpVXhXbwRtj4JBeTKCdTMbB4UQiZMVmFoxDzMBu5Kjvn8w",1)')};
          if (gatewayName == "PTC"){overviewData.getCell(9+i*2,2)
              .setFormula('=image("https://lh5.googleusercontent.com/WTbl6C3pHgn36MuOGAGpoZmME6JqV4zPj3Fro0xhe_W2uqyZxTxUV0HfLeewCyJskrzwPtPg2_cjiOPbXb73URswVE2nR6SjASOCxJryF0XZjLHAOXQopQrp_A",1)')};
          if (gatewayName == "M-1DJ"){overviewData.getCell(9+i*2,2)
              .setFormula('=image("https://lh6.googleusercontent.com/w0akvtb2RxEzQn5D-lpUDrtmyEl6xITE5u0sTHkKvRPwdMlpE4cqF8HLTS1Cz1GtLn9SJNv5bH5OpY-iecQWnHyJSFHSsJodkH5tFoCtOGiiJ8nvayt_cZ9Lhw",1)')};
          if (gatewayName == "PTC/M-1DJ"){overviewData.getCell(9+i*2,2)
              .setFormula('=image("https://lh3.googleusercontent.com/dPRYJxA0VM1BKQE53LHrlsoMdpwS1L7GZg0KImpRR3ZKFl3GUUvVhMv2zi7mdhxqxJHpqEmIal0-LRvRhByG8rz8qucTRhskRmmfpK98qb8Cok9bDCbDp8DPUQ",1)')};
          if (gatewayName == "AA1"){overviewData.getCell(9+i*2,2)
              .setFormula('=image("https://lh3.googleusercontent.com/fqEqrxVqtANX5Nvc-LBFzmMb5Q2hQkvDFKDtYBVERKgjriZRHdTsSyXj3WH_3aoX8F_iarrG_e3Yz5y_gJF0lvX4HkRWHiS1tjsFX6uOPSRgluFERlj01PJ09Q",1)')};
          if (gatewayName == "PA"){overviewData.getCell(9+i*2,2)
              .setFormula('=image("https://lh6.googleusercontent.com/2FdqZrm636NSLYwu9KuctIlZjFnDPB3uNkvv2sZYiHTkPJiCZ3BIBr7PIOXfpNGEMZ3_awYHzqE9G68WPnLMZRtJD6XZ2ENiLamHjVXJx2MOdJC5UP4u0iCFww",1)')};
          if (gatewayName == "PA/AA2"){overviewData.getCell(9+i*2,2)
              .setFormula('=image("https://lh3.googleusercontent.com/qzrg4aSRabf6GafnFez1uEqjsdZMUOhzIf3m2gdKJ1woWXa1y5zt_AB7UOTAxWZjfYDONQeZC3JoD0hsxmpNZlNWaXfYEBrXu8anoCcL267veMxyEaYVjEKvAQ",1)')};
          if (gatewayName == "AA2"){overviewData.getCell(9+i*2,2)
              .setFormula('=image("https://lh5.googleusercontent.com/RSXyfK-NIFWkY2MNtpCZOUFKABuEjnNnHyWNqJhEOuUa20vxNHPe-djUyQHVYYZvtaDtL44iPlwltMXde-cxMncX4t7pZatwCqvzZk_w5_xEfO0L1iTjmUL1Zw",1)')};
          if (gatewayName == "M-1DC"){overviewData.getCell(9+i*2,2)
              .setFormula('=image("https://lh6.googleusercontent.com/THg7p8wEJI7beR4rzyHuZatXqYiNPOto8HmdalPHKhjQiI-BTDAGaEWDNDeiFlONhRfzgAT0vnk1EXDBLeuvXyNHa4oteacDqi17M9IdrhfUY_zcGvZRBzAASg",1)')};
          if (gatewayName == "FAA"){overviewData.getCell(9+i*2,2)
              .setFormula('=image("https://lh6.googleusercontent.com/kqDmN8zBsVqcgFRma3UNG0kLuqqiy4g6nYTly3xcVH5ogxZztD9AwT6fro4Sms_0bnbhQe5b8FBiwA83wuPV9cowwO04T4KIOiS7VePnWzVJcKV7L5F4BkoA_A",1)')};
          if (gatewayName == "FDJ"){overviewData.getCell(9+i*2,2)
              .setFormula('=image("https://lh5.googleusercontent.com/ewe0luCic38pIG2VpmpF6ACWOym65HaPkwz-nSqlNZW48yPq0x2l9u_WLuOUyLT2g52HvGpzet6magzcqGnAHt8dYX5q7LG8cBjNiqcPgcKcotxhX4TNNq4C-g",1)')};
          if (gatewayName == "VP"){overviewData.getCell(9+i*2,2)
              .setFormula('=image("https://lh3.googleusercontent.com/_K2DJtHPm6PJur_oVroDVe8fpL59TyEK_NiurepHrxlP_XBTNdgYZ8eXB-MPY0HewjPmbv1dJb2lbjSCmKVhqvR3GSzo4XKcEN07uDMiP4zEAMqN296-rSZ2EQ",1)')};
          if (gatewayName == "PEC"){overviewData.getCell(9+i*2,2)
              .setFormula('=image("https://lh5.googleusercontent.com/dOZY9Xes6GwWxKh8CFi_QBK7NZ6568OdgbX_304XA6dNQeViteRGSAbFiS07ZJp8u1QxfQOfgSg1xY2wg9XeqWpskgoScTkSQKHVOmGc0szL4Ic-FWv4T-UmaA",1)')};
          if (gatewayName == "FEC"){overviewData.getCell(9+i*2,2)
              .setFormula('=image("https://lh6.googleusercontent.com/thimUe6SHlMPSohqC1dflb-4T1DDe8cX9JB8yqgS7nadd0b3hp5edYbqu5FEdNA5ZAxA00v9JSvTuZ4fhsUY782FDPcOK2HElz4FPFt1cX0KsWrveKf1Gl8y_A",1)')};
          if (gatewayName == "LR"){overviewData.getCell(9+i*2,2)
              .setFormula('=image("https://lh6.googleusercontent.com/AIydW9ujqpyeUzt3yERNxjNtYzZBQYlh2RxfF4CBDNd3SpXerJY4vr5c7wJTKK1ESbPbxDV1tkt8NFib6aXnjaMPAWfU5rXHqlDityVtiQknfNw2hVhIscJucg",1)')};
          if (gatewayName == "LS"){overviewData.getCell(9+i*2,2)
              .setFormula('=image("https://lh4.googleusercontent.com/bkww6X8ZMe2bQG5Kcd-pqKKFfwtvFhVOJ5j3kHZxpPEX6vuu-kOl7qfcOYwvtuNGE7WcXr0FsBCKvOLtPStGJ_OYYJjJx2KmIJ49GGaKgAUIWsTunJfWko4Tog",1)')};
          if (gatewayName == "J1"){overviewData.getCell(9+i*2,2)
              .setFormula('=image("https://lh4.googleusercontent.com/bG-VnNZxBW90RddMVLsaNqYkPjRrnCC3midkcqmcp3ixeSZuJPHCS_IfFlK81_ZwCvLiH3edJR579TlolSawFekjeHphIVDPHSNyGNzmzur5_fKW5eCuHlHOKA",1)')};
          if (gatewayName == "FSR"){overviewData.getCell(9+i*2,2)
              .setFormula('=image("https://lh3.googleusercontent.com/t0K-zeV_TNugqirEcxmcv9Ppe8M1hGH3YB65GVQdin0gFqW46N0xx-gUFoEgONn9KcVB7cW8ruN1r4fwx1qhz6G3ftcDhJuxBjVLjpaUwZgDWgvtN8rp-TS_sg",1)')};
          if (gatewayName == "Unit PS"){overviewData.getCell(9+i*2,2)
              .setFormula('=image("https://lh5.googleusercontent.com/JXY4tTDSUBQPmEU6r5XhGPD6hk1gAZrtzWzfn3N9w0AXg-F1cNOY5HmXKr29aqMCKArU2EpnJmQMMwa-5wWsbKCN1fleOXoyR9LH_6F2fV4ihBT5F17wf-h2Mg",1)')};
          if (gatewayName == "Unit PSC"){overviewData.getCell(9+i*2,2)
              .setFormula('=image("https://lh6.googleusercontent.com/aGyd1k8Wntjb1zlQxkYkxqON6JeTZOK8wDAjbApW9bMXU5c7heYVwddUxcY0A1rTn45CV85A1E5E5ZqoNHmyHenqlV73Nbv0I6dGzqQ_iFMIJW7SCAi6PwBTUg",1)')};
          if (gatewayName == "Unit PTC"){overviewData.getCell(9+i*2,2)
              .setFormula('=image("https://lh5.googleusercontent.com/3auO2Yx5zk1sg24QvHgXneZnmeeShZkZsW-p6bmwsOrAmoeWIxnUir8DzRiA2jXKrMV5s-fMJfqB56JxnCDNpMVnfn2vstuOZlPzq6lusMR7hcFfUZww7K_HRQ",1)')};
          if (gatewayName == "Unit PA"){overviewData.getCell(9+i*2,2)
              .setFormula('=image("https://lh4.googleusercontent.com/fo3hOiA6VhKYaZEicKXez585tOEEosiS9wHxIehISureIrJ_ejN8VGrw3dG4wp1OtRgDhFrVSAkmty0wL8i3bGZfQAEOJoSFscCfr3RXJ8FqTz8Rh3VDzcPrAw",1)')};
          if (gatewayName == "Unit DC"){overviewData.getCell(9+i*2,2)
              .setFormula('=image("https://lh5.googleusercontent.com/iEl2rGWTJ77_YXhopSw8xnRFFL3TJYf8nPIV_RTAX-jQ_bswUC4aWzrsagPjLXYurZKaRN53WnS-bcsghVvgDQZFRbIxb1tdnKLf5uXDw61gsHdy0GcmI5LyCQ",1)')};
          if (gatewayName == "X-1"){overviewData.getCell(9+i*2,2)
              .setFormula('=image("https://lh3.googleusercontent.com/QCp7XXJq6-Lku2lwBPQf-rLf1bK-A3KcEJebv4iAlqhtgSmeaNWmbcPNmmTLT6wIakRs4BSRpiTS7gJxvWU794vhva5LGtqQAc2a6UkVgGXuXXlAfovtXMSdJA",1)')};
          if (gatewayName == "M-1"){overviewData.getCell(9+i*2,2)
              .setFormula('=image("https://lh3.googleusercontent.com/dzqFtVu--ZCzTJRZxe1-70SWHxR9c80HIQl6UpW4MIH2PXjC4HETgoh11HmpE-zsvTynlfduE8d2zOE4zABJs8wPqf04L5SVfRuMbTxWmQcsGpYzef4AtRDYag",1)')};
          if (gatewayName == "DCV"){overviewData.getCell(9+i*2,2)
              .setFormula('=image("https://lh4.googleusercontent.com/J1UyIKVkb-Cgoh7qEX4S8ndrQI0i69UnESUywNUmkc5_ZotlHxg_6HqpkXTLvmT7mzZmkWDbdtqwFvdy9TZNm5qULBtEOb48LeG56XoGqnsXVQ9K5M3bY4mZbQ",1)')};
          if (gatewayName == "VP/DCV"){overviewData.getCell(9+i*2,2)
              .setFormula('=image("https://lh4.googleusercontent.com/6OyObNyysc0iU_pz1MINC-pb0yFzcAJLRsjLfm3mxwwIFxOFBgA0ktILb-PE5AY5j4jJtSPpzVs0bBjcOkbZl1_tCK9aFBRxnwCULEJtK7tfcZ18IkX5AqmeiQ",1)')};
          if (gatewayName == "TT"){overviewData.getCell(9+i*2,2)
              .setFormula('=image("https://lh3.googleusercontent.com/F4oNvN6Tn82pliTnTM-hlfnAKGLhNls5XyC1ahvztCAbRYuazZmOr-lHIJQQUgo61Jv3OmoEiAtX4qcXsnI0dzkI6Vgn3Ew8c0agryaC76gpLWwRTpzNWazrIQ",1)')};
          if (gatewayName == "PP"){overviewData.getCell(9+i*2,2)
              .setFormula('=image("https://lh4.googleusercontent.com/WrVvftiJxseQDWORrY80hVXbXNWYG1fmPLai4t96gNPAj-EBzGOGmgFxwWLHA8Rv6HJxkiqLgdOai_XY6LcnN-2TaiN8dOz6lz164O3qboM2qUYUeMPwEf_VYg",1)')};
          if (gatewayName == "MP1"){overviewData.getCell(9+i*2,2)
              .setFormula('=image("https://lh6.googleusercontent.com/ERIFFY7iofpnJBH8G0vZxlaUT-MtME0Prmnq2waN1sUuynyQk49FLQJsCs0hOzXJATXOeFImPj3UsZ9OBM-SFwHdd4PXZ3MzmFtXeDb2Pk2g77hA0gJEKPkSLA",1)')};
          if (gatewayName == "MP2"){overviewData.getCell(9+i*2,2)
              .setFormula('=image("https://lh4.googleusercontent.com/OImzZqhXZT0Fvn0ODYRfb8981jk-99zZ8jLTj-MYMUrU238ptbpEI_BKR-EPS3QMiq1-KX8Ea7oWZJDxvrvxvDiCM1EOU1_PfURrzlxqEGM_uEK8OaeJTrDjNA",1)')};
      }
}

function restorefeaturesndeliverables(){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var overviewSheet = ss.getSheetByName("Overview");                                                                                                //Saves the Overview sheet
  var overviewData = overviewSheet.getDataRange();                                                                                                  //Saves the Overview data
  var numOverviewCols = overviewSheet.getLastColumn();                                                                                              //Saves the number of columns in the Overview
  var numOverviewRows = overviewSheet.getLastRow();                                                                                                 //Saves the number of rows in the Overview
  
  var featuresReferenceSheet = ss.getSheetByName("Features reference");                                                                             //Saves the Features reference sheet
  var featuresReferenceData = featuresReferenceSheet.getDataRange();                                                                                //Saves the Features reference data
  var numFeaturesReferenceRows = featuresReferenceSheet.getLastRow();                                                                               //Saves the number of rows in Features reference
  var numFeatures = numFeaturesReferenceRows-5;                                                                                                     //Calculates the number of new features
  var numFeaturesBefore = 0;                                                                                                                        //Introduces the number of features before variable

  var gatewaysReferenceSheet = ss.getSheetByName("Gateways reference");                                                                             //Saves the Gateways reference sheet
  var gatewaysReferenceData = gatewaysReferenceSheet.getDataRange();                                                                                //Saves the Gateways reference data
  var numGatewaysReferenceRows = gatewaysReferenceSheet.getLastRow();                                                                               //Saves the number of rows in Gateways reference
  var numGateways = ss.getNumSheets()-8;                                                                                                            //Introduces the number of gateways variable

  var numSheets = ss.getNumSheets();                                                                                                                //Saves the number of sheets in the spreadsheet

  //TAKES THE USER TO THE DATA UPDATE SHEET
  ss.setActiveSheet(ss.getSheetByName("Data update"));                                                                                              //Turns on "Data update" (usually hidden) as active
  var completion = 0;                                                                                                                               //Introduces the completion variable
  ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);                                                                //Uploads the completion level (one chceckpoint) for the chart's (loading bar's) reference
  var totalToBeCompleted = 36*numGateways+9;                                                                                                        //Sets total number of checkpoints for the chart's reference
  ss.getSheetByName("Data update").getDataRange().getCell(4,4).setFormula('='+totalToBeCompleted+'-D3');                                            //Uploads the the total number of checkpoints for the chart's reference  
  ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3).setValue("Starting script...");                                                   //Displays the starting message in the status box below the loading bar
  
  completion++;                                                                                                                                     //Registers that one more checkpoint is completed
  ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);                                                                //Uploads the new completion level to the completion sheet
  ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3).setValue("Unmerging the Features bar in the overview");                           //Updates the status box
  
  //UNMERGES THE "FEATURES" BAR AT THE TOP OF OVERVIEW
  overviewSheet.getRange(4,14,1,numOverviewCols-13).breakApart().setBorder(false,false,false,false,false,false);                                    //Unmerges the "Features" bar and removes its border

  completion++;
  ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
  ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
      .setValue("Removing charts from the Overview...");
  
  //REMOVES ALL CHARTS FROM THE OVERVIEW
  var overviewCharts = overviewSheet.getCharts();                                                                                                   //Saves the charts in the sheet
  for (var i in overviewCharts){                                                                                                                    //Iterates through the charts
      overviewSheet.removeChart(overviewCharts[i])                                                                                                  //Removes each of them
      }

  completion++;
  ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
  ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
      .setValue("Calculating the number of new features...");
  
  //CALCULATES THE NUMBER OF OLD FEATURES
  for (var i=15; i<=numOverviewCols;i++){                                                                                                           //Iterates through the columns in the overview
      if (overviewData.getCell(6,i).getValue() !== ""){numFeaturesBefore++}                                                                         //Calculates how many features were there before
      }
   
  completion++;
  ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
  ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
      .setValue("Removing columns from the Overview...");
  
  //IF THERE ARE MORE THAN ONE FEATURES IT DELETES THE OLD ONES (EXCEPT THE FIRST)
  if (numFeaturesBefore>1){                                                                                                                         //Checks if there was more than one feature before
      overviewSheet.deleteColumns(17,numFeaturesBefore*2-3)}                                                                                        //If yes, it deletes all the other ones

  completion++;
  ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
  ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
      .setValue("Removing columns from the Overview...");
  
  //DELETES ONE MORE COLUMN IF THERE IS JUST ONE FEATURE (DESIGN REQUIREMENT)
  if (numFeatures==1 && numFeaturesBefore>1){                                                                                                       //Checks if there is just one new feature and if there were more thatn one features before
      overviewSheet.deleteColumns(16,1);                                                                                                            //If yes, it deletes one more column
      overviewData.getCell(4,15).setBorder(true,true,true,true,false,false);                                                                        //Sets the "Features" bar borders
      overviewData.getCell(6,15).setBorder(true,true,true,true,false,false);                                                                        //Sets the first feature's borders
      }

  completion++;
  ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
  ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
      .setValue("Adding new feature columns to the Overview...");
  
  //IF THERE ARE MORE THAN ONE FEATURES IT ADDS THEM
  if (numFeatures>1){                                                                                                                               //Checks if there are more than on new featues
  
      //ADDS APPROPRIATE NUMBER OF FEATURE COLUMNS
      overviewSheet.insertColumnsAfter(16,numFeatures*2-3);                                                                                         //If yes, it adds new columns

      //ADDS ONE MORE COLUMN IF THERE WAS JUST ONE FEATURE
      if (numFeaturesBefore==1){overviewSheet.insertColumnsAfter(16,1)};                                                                            //Adds one more column if there was just one feature before

      //MERGES THE FEATURES BAR
      overviewSheet.getRange(4,15,1,numFeatures*2-1).merge()                                                                                        //Merges the "Features" bar back
      }

  completion++;
  ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
  ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
      .setValue("Changing the features' formatting in the overview..."); 
  
  //CHANGES THE FEATURES' COLUMNS' FORMATTING
  overviewSheet.getRange(6,15,1,numFeatures*2-1).setFontFamily('Trebuchet MS').setFontWeight('Bold');                                               //Changes the formatting of the features' columns
  
  completion++;
  ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
  ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
      .setValue("Populating new feature columns in the overview...");
  
  //POPULATES THE NEW FEATURES AND INSERTS THEIR PROPER FORMATTING
  overviewData = overviewSheet.getRange(1,1,numOverviewRows,numFeatures*2+14);                                                                      //Updates the overview data variable
  for (var i=0; i<numFeatures; i++){                                                                                                                //Iterates through the features
      overviewData.getCell(6,i*2+15)                                                                                                                //Takes a feature cell
          .setValue(featuresReferenceData.getCell(i+6,3).getValue())                                                                                //And populates it
          .setBackground('#E6B8AF').setHorizontalAlignment("center")                                                                                //Changes its background
          .setVerticalAlignment("center");                                                                                                          //and vertical alignment
      if (i<numFeatures-1){
          overviewSheet.setColumnWidth(i*2+15,50);                                                                                                  //Sets the feature's column width
          overviewSheet.setColumnWidth(i*2+16,5);                                                                                                   //Sets the gap between the feature's column width
          }
      }
  overviewSheet.setColumnWidth(numFeatures*2+13,50);                                                                                                //Sets column width of the last feature
  overviewSheet.setColumnWidth(numFeatures*2+14,35);                                                                                                //Sets column width of the last blank column
  
  completion++;
  ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
  ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
      .setValue("Starting to iterate through the sheets...");
  
  //ITERATES THROUGH THE SHEETS MAKING CHANGES
  for (var s=1; s<=numGateways; s++){                                                                                                               //Iterates through the sheets

      var gatewaySpecificSheet = ss.getSheets()[s];                                                                                                 //Saves the gateway specific sheet
      var gatewaySpecificData = ss.getSheets()[s].getDataRange();                                                                                   //Saves the gateway specific data range
      var gatewayName = gatewaySpecificSheet.getSheetName();                                                                                        //Saves the sheet name
      var numGatewaySpecificRows = gatewaySpecificSheet.getLastRow();                                                                               //Saves the number of rows in the gateway specific sheet
      var numGatewaySpecificCols = gatewaySpecificSheet.getLastColumn();                                                                            //Saves the number of columns in the gateway specific sheet

      var deliverablesReferenceSheet = ss.getSheetByName("Deliverables reference");                                                                 //Saves the Deliverables reference sheet
      var deliverableReferenceData = deliverablesReferenceSheet.getDataRange();                                                                     //Saves the Deliverables reference data
      var numDeliverablesReferenceRows = deliverablesReferenceSheet.getLastRow();                                                                   //Saves the number of rows in the Deliverables reference sheet
      var numDeliverables = 0;                                                                                                                      //Introduces the number of deliverables variable

      //CALCULATES THE NUMBER OF DELIVERABLES
      for (var i=5; i<=numDeliverablesReferenceRows; i++){                                                                                          //Iterates through the rows in Deliverables reference
          if (deliverableReferenceData.getCell(i,2).getValue()==gatewayName){numDeliverables++}                                                     //Calculates the number of deliverables
          }
  
      //IF THE FIRST DELIVERABLE'S NAME IS "DELIVERABLE NAME", IT SETS THE NUMBER OF DELIVERABLES TO 1 (TO PREVENT CRASHING IF THIS SCRIPTS IS TRIGGERED AFTER THE GATEWAY UPDATER)
      if (gatewaySpecificData.getCell(2,19).getValue()=="Deliverable name"){numDeliverables=1};
  
      //IF THERE ARE ANY DELIVERABLES, IT STARTS TO MAKE CHANGES
      if (numDeliverables>0){                                                                                                                       //If there are any deliverables, it starts to make changes

          completion++;
          ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
          ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Removing charts from "+ss.getSheets()[s].getSheetName()+"...");
          
          //REMOVES ALL CHARTS FROM THE SHEET
          var gatewaycharts = ss.getSheets()[s].getCharts();                                                                                        //Saves all charts in the sheet
          for (var i in gatewaycharts){                                                                                                             //Iterates through the charts
              gatewaySpecificSheet.removeChart(gatewaycharts[i])                                                                                    //Removes each of them
              }

          completion++;
          ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
          ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Unmerging lead engineer cells in "+ss.getSheets()[s].getSheetName()+"...");
          
          //UNMERGES THE LEADER AND SYSTEM CELLS ON THE LEFT
          gatewaySpecificSheet.getRange(6, 2, numGatewaySpecificRows-5, 1).breakApart().setBorder(false, false, false, false, false, false);

          completion++;
          ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
          ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Merging the first lead engineer cell in "+ss.getSheets()[s].getSheetName()+"...");
          
          //MERGES THE FIRST LEADER'S CELL
          gatewaySpecificSheet.getRange(7, 2, 2, 1).merge().setBorder(true, true, true, true, false, false);

          completion++;
          ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
          ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Removing redundant borders in "+ss.getSheets()[s].getSheetName()+"...");
          
          //REMOVES ALL THE BORDERS
          gatewaySpecificSheet.getRange(10,1,numGatewaySpecificRows-9,numGatewaySpecificCols-9).setBorder(false,false,false,false,false,false);

          completion++;
          ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
          ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Removing the reference columns in "+ss.getSheets()[s].getSheetName()+"...");
          
          //REMOVES THE REFERENCE COLUMNS FOR THE CHARTS AT THE END
          gatewaySpecificSheet.deleteColumns(19+numDeliverables*2,numGatewaySpecificCols-numDeliverables*2-18);

          completion++;
          ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
          ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Removing redundant rows in "+ss.getSheets()[s].getSheetName()+"...");
          
          //REMOVES REDUNDANT ROWS
          gatewaySpecificSheet.deleteRows(10,numGatewaySpecificRows-9);

          completion++;
          ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
          ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Adding new rows in "+ss.getSheets()[s].getSheetName()+"...");
          
          //ADDS AN APPROPRIATE NUMBER OF NEW ROWS
          gatewaySpecificSheet.insertRowsAfter(9, 3*numFeatures-3+numFeatures+6);

          completion++;
          ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
          ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Resizing a row in "+ss.getSheets()[s].getSheetName()+"...");
          
          //RESIZES THE ROW BELOW THE FIRST FEATURE
          gatewaySpecificSheet.setRowHeight(9, 10);

          completion++;
          ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
          ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Adding new columns in "+ss.getSheets()[s].getSheetName()+"...");
          
          //ADDS AN APPROPRIATE NUMBER OF COLUMNS FOR THE CHARTS' REFERENCE                                                                         //If there is more than three features
          gatewaySpecificSheet.insertColumnsAfter(18+numDeliverables*2, numFeatures*2+2)                                                            //It inserts the number of feature columns in the end


          completion++;
          ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
          ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Updating the data range in "+ss.getSheets()[s].getSheetName()+"...");
          
          //UPDATES THE GATEWAY SPECIFIC DATA RANGE
          gatewaySpecificData = gatewaySpecificSheet.getRange(1,1,9+3*numFeatures-3+numFeatures+6,18+numDeliverables*2+numFeatures*2+2);            //It updates the Gateway specific data range
          numGatewaySpecificRows = 9+3*numFeatures-3+numFeatures+6;                                                                                 //It updates the number of Gateway specific rows
          numGatewaySpecificCols = 18+numDeliverables*2+numFeatures*2+2;                                                                            //It updates the number of Gateway specific columns

          completion++;
          ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
          ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Clearing the formatting in "+ss.getSheets()[s].getSheetName()+"...");

          //CLEARS THE FORMATTING OF THE FIRST FEATURE
          gatewaySpecificData.getCell(7,2).setValue("").setBackgroundColor("#FFFFFF");                                                              //It clears the Lead Engineer cell
          gatewaySpecificData.getCell(7,4).setValue("");                                                                                            //It clears the Feature name cell
          gatewaySpecificData.getCell(7,6).setValue("");                                                                                            //It clears the Engineer in charge cell
          for (var i=0; i<numDeliverables; i++){                                                                                                    //Iterates through the feature columns
              gatewaySpecificData.getCell(7,19+i*2).setValue("Not updated");                                                                        //Sets the status to "Not updated"
              gatewaySpecificData.getCell(8,19+i*2).setValue("");                                                                                   //Resets the comment box under status cell
              }

          completion++;
          ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
          ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Copying the first feature's formatting in "+ss.getSheets()[s].getSheetName()+"...");
          
          //COPIES THE FORMATTING OF THE FIRST FEATURE ELSEWHERE
          var sourceOfFormatting = gatewaySpecificSheet.getRange(7,1,2,numGatewaySpecificCols);                                                     //Takes the first row's formatting
          for(var i=1; i<numFeatures; i++){                                                                                                         //Iterates through the new rows
              var destinationOfFormatting = gatewaySpecificSheet.getRange(i*3+7,1,2,numGatewaySpecificCols);                                        //Takes a new row
              sourceOfFormatting.copyTo(destinationOfFormatting);                                                                                   //And copies the first row's formatting into the new row
              }
 
          completion++;
          ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
          ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Resizing the new rows in "+ss.getSheets()[s].getSheetName()+"...");
          
          //RESIZES THE NEW FEATURE ROWS
          for(var i=0; i<numFeatures-1; i++){                                                                                                       //Iterates through the rows
              gatewaySpecificSheet.setRowHeight(i*3+10, 45)
              gatewaySpecificSheet.setRowHeight(i*3+11, 60)
              gatewaySpecificSheet.setRowHeight(i*3+12, 10)
              }

          completion++;
          ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
          ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Populating the new feature rows in "+ss.getSheets()[s].getSheetName()+"...");
          
          //POPULATES THE NEW FEATURE ROWS
          for(var i=0; i<numFeatures; i++){                                                                                                         //Iterates through the new features
              gatewaySpecificData.getCell(i*3+7,2).setValue(featuresReferenceData.getCell(i+6,2).getValue())                                        //Populates the lead engineer cell end sets its formatting
                  .setBackground(featuresReferenceData.getCell(i+6,5).getBackgrounds())
                  .setFontColor(featuresReferenceData.getCell(i+6,6).getBackgrounds());
              gatewaySpecificData.getCell(i*3+7,4).setValue(featuresReferenceData.getCell(i+6,3).getValue());                                       //Populates the Feature name cell
              gatewaySpecificData.getCell(i*3+7,6).setValue(featuresReferenceData.getCell(i+6,4).getValue());                                       //Populates the Engineer in charge cell
              }

          completion++;
          ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
          ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Merging the lead engineers in "+ss.getSheets()[s].getSheetName()+"...");
          
          //MERGES THE LEADER CELLS FOR THE FEATURES WITH MATCHING LEADERS
          var sameLeader=0;                                                                                                                         //Introduces the sameLeader variable
          for (var i=0; i<numFeatures; i++){                                                                                                        //Iterates through the features
              sameLeader=0;                                                                                                                         //Resets the sameLeader variable
              for (var j=i; j<numFeatures; j++){                                                                                                    //Iterates through the features below the analyzed one
                  if (gatewaySpecificData.getCell(j*3+7,2).getValue()==gatewaySpecificData.getCell(i*3+7,2).getValue()){sameLeader++};              //Calculates how many features are under the same lead engineer
                  if (gatewaySpecificData.getCell(j*3+7,2).getValue()!==gatewaySpecificData.getCell(i*3+7,2).getValue()){break}                     //If there is a feature not under the same lead engineer, it stops the loop
                  }
              gatewaySpecificSheet.getRange(i*3+7, 2, sameLeader*3-1).merge().setBorder(true,true,true,true,true,true);                             //Merges the Lead engineer cells for different features
              gatewaySpecificSheet.setRowHeight(i*3+7+sameLeader*3-1, 20)                                                                           //Increases the gap between features with different lead engineers
              i=i+sameLeader-1;                                                                                                                     //Omits the features whose lead engineer's have already been merged
              }

          completion++;
          ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
          ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Populating the column chart's reference in "+ss.getSheets()[s].getSheetName()+"...");
          
          //POPULATES THE OVERVIEW'S COLUMN CHARTS REFERENCE
          for (var i=0; i<numFeatures; i++){                                                                                                        //Iterates through the features
              var index=i*3+7;
              gatewaySpecificData.getCell(numGatewaySpecificRows-numFeatures+1+i,numGatewaySpecificCols-numFeatures*2-1)                            //Populates the reference
                  .setValue(featuresReferenceData.getCell(i+6,3).getValue());
              gatewaySpecificData.getCell(numGatewaySpecificRows-numFeatures+1+i,numGatewaySpecificCols-numFeatures*2)                            
                  .setFormula('=IF(COUNTIF($S'+index+':'+index+', "*o*")=0, 1,(COUNTIF($S'+index+':'+index+', "OK")/COUNTIF($S'+index+':'+index+', "*o*")))');
              gatewaySpecificData.getCell(numGatewaySpecificRows-numFeatures+1+i,numGatewaySpecificCols-numFeatures*2+1)
                  .setFormula('=IF(COUNTIF($S'+index+':'+index+', "*o*")=0, 0,(COUNTIF($S'+index+':'+index+', "In progress")/COUNTIF($S'+index+':'+index+', "*o*")))');
              gatewaySpecificData.getCell(numGatewaySpecificRows-numFeatures+1+i,numGatewaySpecificCols-numFeatures*2+2)
                  .setFormula('=IF(COUNTIF($S'+index+':'+index+', "*o*")=0, 0,(COUNTIF($S'+index+':'+index+', "Not OK")+COUNTIF($S'+index+':'+index+', "Not started"))/COUNTIF($S'+index+':'+index+', "*o*"))');
              gatewaySpecificData.getCell(numGatewaySpecificRows-numFeatures+1+i,numGatewaySpecificCols-numFeatures*2+3)
                  .setFormula('=IF(COUNTIF($S'+index+':'+index+', "*o*")=0, 0,(COUNTIF($S'+index+':'+index+', "Not updated")/COUNTIF($S'+index+':'+index+', "*o*")))');
              }

          completion++;
          ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
          ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Populating the pie chart's reference in "+ss.getSheets()[s].getSheetName()+"...");
          
          //POPULATES THE SHEET'S PIE CHARTS REFERENCE
          for (var i=0; i<numFeatures; i++){                                                                                                        //Iterates through the features
              var index=i*3+7;
              
              gatewaySpecificData.getCell(numGatewaySpecificRows-numFeatures-5,numGatewaySpecificCols-numFeatures*2+i*2+1)                          //Populates the reference 
                  .setValue("Feature");
              gatewaySpecificData.getCell(numGatewaySpecificRows-numFeatures-5,numGatewaySpecificCols-numFeatures*2+i*2+2)                          //Populates the reference 
                  .setValue(featuresReferenceData.getCell(i+6,3).getValue());
              
              gatewaySpecificData.getCell(numGatewaySpecificRows-numFeatures-4,numGatewaySpecificCols-numFeatures*2+i*2+1)                          //Populates the reference 
                  .setValue("Green");
              gatewaySpecificData.getCell(numGatewaySpecificRows-numFeatures-4,numGatewaySpecificCols-numFeatures*2+i*2+2)
                  .setFormula('=IF(COUNTIF($S'+index+':'+index+', "*o*")=0, 1,(COUNTIF($S'+index+':'+index+', "OK")/COUNTIF($S'+index+':'+index+', "*o*")))');
              
              gatewaySpecificData.getCell(numGatewaySpecificRows-numFeatures-3,numGatewaySpecificCols-numFeatures*2+i*2+1)                          //Populates the reference 
                  .setValue("Amber");
              gatewaySpecificData.getCell(numGatewaySpecificRows-numFeatures-3,numGatewaySpecificCols-numFeatures*2+i*2+2)
                  .setFormula('=IF(COUNTIF($S'+index+':'+index+', "*o*")=0, 0,(COUNTIF($S'+index+':'+index+', "In progress")/COUNTIF($S'+index+':'+index+', "*o*")))');
              
              gatewaySpecificData.getCell(numGatewaySpecificRows-numFeatures-2,numGatewaySpecificCols-numFeatures*2+i*2+1)                          //Populates the reference 
                  .setValue("Red");
              gatewaySpecificData.getCell(numGatewaySpecificRows-numFeatures-2,numGatewaySpecificCols-numFeatures*2+i*2+2)
                  .setFormula('=IF(COUNTIF($S'+index+':'+index+', "*o*")=0, 0,(COUNTIF($S'+index+':'+index+', "Not OK")+COUNTIF($S'+index+':'+index+', "Not started"))/COUNTIF($S'+index+':'+index+', "*o*"))');
              
              gatewaySpecificData.getCell(numGatewaySpecificRows-numFeatures-1,numGatewaySpecificCols-numFeatures*2+i*2+1)                          //Populates the reference 
                  .setValue("Blue");
              gatewaySpecificData.getCell(numGatewaySpecificRows-numFeatures-1,numGatewaySpecificCols-numFeatures*2+i*2+2)
                  .setFormula('=IF(COUNTIF($S'+index+':'+index+', "*o*")=0, 0,(COUNTIF($S'+index+':'+index+', "Not updated")/COUNTIF($S'+index+':'+index+', "*o*")))');
              }
  
          completion++;
          ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
          ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Populating the pie chart's reference in "+ss.getSheets()[s].getSheetName()+"...");
          
          //POPULATES THE OVERVIEW'S PIE CHARTS REFERENCE
          gatewaySpecificData.getCell(numGatewaySpecificRows-numFeatures-5,numGatewaySpecificCols-numFeatures*2-1).setValue("Overall");
          gatewaySpecificData.getCell(numGatewaySpecificRows-numFeatures-5,numGatewaySpecificCols-numFeatures*2).setValue("Overall")
          
          gatewaySpecificData.getCell(numGatewaySpecificRows-numFeatures-4,numGatewaySpecificCols-numFeatures*2-1).setValue("Green");
          gatewaySpecificData.getCell(numGatewaySpecificRows-numFeatures-4,numGatewaySpecificCols-numFeatures*2).setFormula("=$I$4")
          
          gatewaySpecificData.getCell(numGatewaySpecificRows-numFeatures-3,numGatewaySpecificCols-numFeatures*2-1).setValue("Amber");
          gatewaySpecificData.getCell(numGatewaySpecificRows-numFeatures-3,numGatewaySpecificCols-numFeatures*2).setFormula("=$K$4")
          
          gatewaySpecificData.getCell(numGatewaySpecificRows-numFeatures-2,numGatewaySpecificCols-numFeatures*2-1).setValue("Red");
          gatewaySpecificData.getCell(numGatewaySpecificRows-numFeatures-2,numGatewaySpecificCols-numFeatures*2).setFormula("=$M$4")
          
          gatewaySpecificData.getCell(numGatewaySpecificRows-numFeatures-1,numGatewaySpecificCols-numFeatures*2-1).setValue("Blue");
          gatewaySpecificData.getCell(numGatewaySpecificRows-numFeatures-1,numGatewaySpecificCols-numFeatures*2).setFormula("=$O$4")
  
          completion++;
          ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);          
          ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Removing redundant borders in "+ss.getSheets()[s].getSheetName()+"...");
          
          //REMOVES REDUNDANT BORDERES
          gatewaySpecificSheet.getRange(numGatewaySpecificRows-numFeatures-6, 1, numFeatures+7, numGatewaySpecificCols)
              .setBorder(false,false,false,false,false,false);
  
          completion++;
          ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
          ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Improving the formatting in "+ss.getSheets()[s].getSheetName()+"...");
          
          //MAKES SURE THERE IS GOOD FORMATTING INSERTED IN THE LAST FEATURE (VERY BOTTOM)
          var sourceOfFormatting = gatewaySpecificSheet.getRange(7,9,2,numGatewaySpecificCols-8);                                                   //Takes the first feature's formatting
          var destinationOfFormatting = gatewaySpecificSheet.getRange(numFeatures*3+4,9,2,numGatewaySpecificCols-8);                                //Takes the last feature's range
          sourceOfFormatting.copyTo(destinationOfFormatting);                                                                                       //Copies the formatting
   
          completion++;
          ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
          ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Improving the formatting in "+ss.getSheets()[s].getSheetName()+"...");
          
          //SETS THE CHARTS' REFERENCE'S FORMATTING
          gatewaySpecificSheet.getRange(numGatewaySpecificRows-numFeatures-5,numGatewaySpecificCols-numFeatures*2-1,numFeatures+6,numFeatures*2+2)
              .setFontFamily('Trebuchet MS')
              .setNumberFormat("0%")
              .setFontFamily('Trebuchet MS')
              .setFontSize(8)
              .setHorizontalAlignment("center")
              .setVerticalAlignment("center");
    
          completion++;
          ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
          ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Hiding the charts' reference in "+ss.getSheets()[s].getSheetName()+"...");
          
          //HIDES THE CHARTS' REFERENCE
          gatewaySpecificSheet.hideRows(numGatewaySpecificRows-numFeatures-5,numFeatures+6);
          gatewaySpecificSheet.hideColumns(numGatewaySpecificCols-numFeatures*2-1,numFeatures*2+2);

          completion++;
          ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
          ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Inserting pie charts to "+ss.getSheets()[s].getSheetName()+"...");
          
          //INSERT PIE CHARTS TO THE SHEET
          for (var i=0; i<numFeatures; i++){                                                                                                        //Iterates through the features
              var index=i*3+7;
              var gatewayPieChart = gatewaySpecificSheet.newChart()                                                                                 //Inserts pie charts
                  .setOption('width', 197)
                  .setOption('height', 101)
                  .setOption('theme','maximized')
                  .setOption('legend', 'none')
                  .setOption('pieHole', 0.5)
                  .setOption('pieSliceTextStyle',{color:'black'})
                  .setOption('fontSize', 9)
                  .setOption('pieSliceText','value')
                  .setOption('colors', ['#00B050','#FFC000','#FF0000','#6FA8DC'])
                  .setChartType(Charts.ChartType.PIE)
                  .addRange(gatewaySpecificSheet.getRange(numGatewaySpecificRows-numFeatures-4,numGatewaySpecificCols-numFeatures*2+i*2+1,4,2))
                  .setPosition(index,9, 243, 0)
                  .build();
              gatewaySpecificSheet.insertChart(gatewayPieChart);
              }

          completion++;
          ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
          ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Inserting "+ss.getSheets()[s].getSheetName()+" charts to the overview...");
          
          //INSERTS CHARTS TO THE OVERVIEW
          var ss = SpreadsheetApp.getActiveSpreadsheet();
          var pieChartRange = gatewaySpecificSheet                                                                                                  //Saves the pie chart's range
              .getRange(numGatewaySpecificRows-numFeatures-4,numGatewaySpecificCols-numFeatures*2-1,4,2);
          var columnChartRange = gatewaySpecificSheet                                                                                               //Saves the column chart's range
              .getRange(numGatewaySpecificRows-numFeatures+1,numGatewaySpecificCols-numFeatures*2-1,numFeatures,5);
          
          completion++;
          ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
          ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Inserting "+ss.getSheets()[s].getSheetName()+" pie chart to the overview...");
          
          //INSERTS THE PIE CHART
          var overviewPieChart = overviewSheet.newChart()                                                                                           //Inserts pie chart to the overview
              .setOption('width', 110)
              .setOption('height', 110)
              .setOption('theme','maximized')
              .setOption('legend', 'none')
              .setOption('fontSize', 9)
              .setOption('pieSliceText','value')
              .setOption('colors', ['#00B050','#FFC000','#FF0000','#6FA8DC'])
              .setChartType(Charts.ChartType.PIE)
              .addRange(pieChartRange)
              .setPosition(7+s*2,9, 37, 0)
              .build();
          overviewSheet.insertChart(overviewPieChart); 
          
          completion++;
          ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
          ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Inserting "+ss.getSheets()[s].getSheetName()+" column chart to the overview...");
          
          //INSERTS THE COLUMN CHART
          var overviewColumnChart = overviewSheet.newChart()                                                                                        //Inserts column chart to the overview
              .setOption('width', numFeatures*57+2)
              .setOption('height', 110)
              .setOption('theme','maximized')
              .setOption('legend','none')
              .setOption('hAxis', {'textPosition': 'none'})
              .setOption('vAxis.textPosition', 'none')
              .setOption('colors', ['#00B050','#FFC000','#FF0000','#6FA8DC'])
              .setChartType(Charts.ChartType.COLUMN)
              .addRange(columnChartRange)
              .setOption("isStacked", true)
              .setPosition(7+s*2,1,683,0)
              .build();
          overviewSheet.insertChart(overviewColumnChart);
          }
      }

  //SETS BORDERS IN OVERVIEW FEATURES' CELLS AND THE 'FEATURES' BAR
  for (var i=0; i<numFeatures; i++){
      overviewData.getCell(6,i*2+15).setBorder(true,true,true,true,true,true);
      overviewData.getCell(5,i*2+15).setBorder(true,false,true,false,false,false);
      }
      
  overviewSheet.getRange(4,15,1,numFeatures*2-1).setBorder(true,true,true,true,false,false);
  overviewSheet.getRange(3,15,1,numFeatures*2-1).setBorder(false,false,true,false,false,false);

  
  ss.getSheetByName("Features reference").hideSheet()

  var deliverablesReferenceSheet = ss.getSheetByName('Deliverables reference');                                                                      //Saves the Deliverables reference sheet
  var deliverablesReferenceData = deliverablesReferenceSheet.getDataRange();                                                                         //Saves the Deliverables reference data
  var numDeliverablesReferenceRows = deliverablesReferenceSheet.getLastRow();                                                                        //Saves the number of rows in Deliverables reference

  var deliverablesReferenceArray = new Array (numDeliverablesReferenceRows);                                                                         //Sets up a new array storing the Deliverables Reference Sheet
  for (var i=0; i<numDeliverablesReferenceRows; i++){                                                                                                //For each row of the array it sets up new columns
      deliverablesReferenceArray [i] = new Array (7);
      }
  
  deliverablesReferenceArray = deliverablesReferenceData.getValues();                                                                                //Saves the data from Deliverables reference in the array
  
  var featureReferenceSheet = ss.getSheetByName('Features reference');                                                                               //Saves the Feature reference sheet
  var numFeaturesReferenceRows = featureReferenceSheet.getLastRow();                                                                                 //Saves the number of rows in Features reference
  var numFeatures = numFeaturesReferenceRows-5;                                                                                                      //Saves the number of features used in the spreadsheet

  var gatewaysReferenceSheet = ss.getSheetByName('Gateways reference');                                                                              //Saves the Gateways reference sheet
  var numGatewaysReferenceRows = gatewaysReferenceSheet.getLastRow();                                                                                //Saves the number of rows in Gateways reference
  var numGateways = ss.getNumSheets()-8;                                                                                                             //Saves the number of gateways

  //TAKES THE USER TO THE DATA UPDATE SHEET
  ss.setActiveSheet(ss.getSheetByName("Data update"));                                                                                               //Turns on "Data update" (usually hidden) as active
  var completion = 0;                                                                                                                                //Introduces the completion variable
  ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);                                                                 //Uploads the completion level (one chceckpoint) for the chart's (loading bar's) reference
  var totalToBeCompleted = 10*numGateways;                                                                                                           //Sets total number of checkpoints for the chart's reference
  ss.getSheetByName("Data update").getDataRange().getCell(4,4).setFormula('='+totalToBeCompleted+'-D3');                                             //Uploads the the total number of checkpoints for the chart's reference
  ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3).setValue("Starting script...");                                                    //Displays the starting message in the status box

  for (var s=1; s<=numGateways; s++){                                                                                                                //Iterates through the gateway specific sheets
      
      var gatewaySpecificSheet = ss.getSheets()[s];                                                                                                  //Saves the gateway specific sheet
      var gatewaySpecificData = gatewaySpecificSheet.getDataRange();                                                                                 //Saves the gateway specific data range
      var numGatewaySpecificRows = gatewaySpecificSheet.getLastRow();                                                                                //Saves the number of rows in the gateway specific sheet
      var numGatewaySpecificCols = gatewaySpecificSheet.getLastColumn();                                                                             //Saves the number of columns in the gateway specific sheet
      var gatewayName = gatewaySpecificSheet.getSheetName();                                                                                         //Saves the gateway's name

      var numDeliverablesBefore = 0;                                                                                                                 //Introduces the number of deliverables before variable
      var numDeliverables = 0;                                                                                                                       //Introduces the number of deliverables variable
      var numOtherDeliverables = 0;                                                                                                                  //Introduces the number of other deliverables variable
      var deliverableDifference = 0;                                                                                                                 //Introduces the deliverable difference variable
 
      completion++;
      ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
      ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
          .setValue("Calculating the number of new deliverables in "+ss.getSheets()[s].getSheetName()+"..."); 
 
      //CALCULATES THE NUMBER OF DELIVERABLES TO INSERT
      for (var i=5; i<=numDeliverablesReferenceRows; i++){                                                                                           //Iterates through the rows in the Deliverables reference
          if (deliverablesReferenceArray[i-1][1]==gatewayName){numDeliverables++}                                                                    //Calculates the number of deliverables to insert
          }

      completion++;                                                                                                                                  //Registers that one more checkpoint is completed
      ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);                                                             //Uploads the new completion level to the completion sheet
      ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
          .setValue("Calculating the number of old deliverables in "+ss.getSheets()[s].getSheetName()+"...");
      
      if (numDeliverables>0){
      //CALCULATES THE NUMBER OF DELIVERABLES BEFORE                                                           
      for (var i=19; i<=numGatewaySpecificCols; i++){                                                                                                //Iterates through the columns in the gateway sheet
          if (gatewaySpecificData.getCell(2,i).getValue()!==""){numDeliverablesBefore++}                                                             //Calculates the number of deliverables present in the sheet before
          }
      
      //IF THERE IS THE SAME NUMBER OF NEW AND OLD DELIVERABLES, IT JUST GOES THROUGH THE CHECKPOINT
      if (numDeliverablesBefore==numDeliverables){                                                                                                   //Checks if there is as many new deliverables as the old ones        
          completion=completion+5;
          ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
          ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Proceeding...");
          }
      
      //IF THERE IS MORE OLD DELIVERABLES THAN NEW ONES, IT REMOVES REDUNDANT ONES
      if (numDeliverablesBefore>numDeliverables){                                                                                                    //Checks if there were more deliverables before than there are new ones          
          completion=completion+5;
          ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
          ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Removing redundant columns in "+ss.getSheets()[s].getSheetName()+"...");
          
          //REMOVES COLUMNS
          deliverableDifference = numDeliverablesBefore-numDeliverables;                                                                             //If yes, calculates how many to remove
          gatewaySpecificSheet.deleteColumns(18+numDeliverables*2,deliverableDifference*2);                                                          //Removes the redundant columns
          }
  
      //IF THERE IS MORE NEW DELIVERABLES THAN OLD ONES, IT ADDS APPROPRIATE NUMBER OF COLUMNS
      if (numDeliverablesBefore<numDeliverables){                                                                                                    //Checks if there are more new deliverables than the old ones
          completion++;
          ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
          ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Adding columns in "+ss.getSheets()[s].getSheetName()+"...");
              
          //ADDS COLUMNS
          deliverableDifference = numDeliverables-numDeliverablesBefore;                                                                             //If yes, calculates how many to add
          gatewaySpecificSheet.insertColumnsAfter(18+numDeliverablesBefore*2, deliverableDifference*2)                                               //Adds desired number of columns
  
          completion++;
          ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
          ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Resizing columns in "+ss.getSheets()[s].getSheetName()+"...");
          
          //RESIZES THE NEW COLUMNS
          for (var i=numDeliverablesBefore; i<numDeliverables; i++){                                                                                 //Iterates through the new columns
              gatewaySpecificSheet.setColumnWidth(i*2+19, 195)                                                                                       //Chenges the new columns width
              gatewaySpecificSheet.setColumnWidth(i*2+20, 40)                                                                                        //Changes the new columns width (this is the gap between deliverables)
              }
    
          completion++;
          ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
          ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Copying the deliverable formatting in "+ss.getSheets()[s].getSheetName()+"...");
         
          //COPIES THE FIRST DELIVERABLE'S FORMATTING AND PASTES IT TO THE FIRST NEW DELIVERABLE
          var sourceOfFormatting = gatewaySpecificSheet.getRange(1,19,numGatewaySpecificRows,1);                                                     //Takes formatting from the first deliverables
          var destinationOfFormatting = gatewaySpecificSheet.getRange(1, 19+2*numDeliverablesBefore,numGatewaySpecificRows,1);                       //Takes the first new deliverable
          sourceOfFormatting.copyTo(destinationOfFormatting);                                                                                        //And copies the formatting to it
          
          completion++;
          ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
          ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Reseting the status of the first new deliverable "+ss.getSheets()[s].getSheetName()+"...");
          
          //RESETS THE STATUSES IN THE NEW DELIVERABLE
          for (var i=0; i<numFeatures; i++){                                                                                                         //Iterates through all the features
              gatewaySpecificData.getCell(i*3+7,19+2*numDeliverablesBefore).setValue("Not updated");                                                 //Resets all features statuses to Not Updated
              gatewaySpecificData.getCell(i*3+8,19+2*numDeliverablesBefore).setValue("");                                                            //Resets all features coments
              };
              
          completion++;
          ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
          ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Copying the deliverable formatting in "+ss.getSheets()[s].getSheetName()+"...");
          
          //COPIES THE NEW FORMATTING ELSEWHERE
          var sourceOfFormatting = gatewaySpecificSheet.getRange(1, 19+2*numDeliverablesBefore,numGatewaySpecificRows,1);                            //Takes formatting from the first new deliverable (reset statuses)
          for(var i=numDeliverablesBefore+1; i<numDeliverables; i++){                                                                                //Iterates through the new deliverables
              var destinationOfFormatting = gatewaySpecificSheet.getRange(1, 19+2*i,numGatewaySpecificRows,1);                                       //Takes the new deliverable
              sourceOfFormatting.copyTo(destinationOfFormatting);                                                                                    //And copies the formatting to it
              }          
          }
      
      completion++;
      ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
      ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Calculating the number of preceding deliverables in "+ss.getSheets()[s].getSheetName()+"...");
              
      //CALCULATES THE NUMBER OF DELIVERABLES PRECEDING THE DESIRED ONES
      for (var i=5; i<=numDeliverablesReferenceRows; i++){                                                                                           //Iterates through the Deliverables reference rows
          if (deliverablesReferenceArray[i-1][1]==gatewayName){numOtherDeliverables=i-5; break}                                                      //Calculates how many deliverables there are before the gateway that we're working on and then breaks
          }

      completion++;
      ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
      ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Updating the data range in "+ss.getSheets()[s].getSheetName()+"...");
      
      //UPDATES THE GATEWAY SPECIFIC DATA RANGE
      var gatewaySpecificData = ss.getSheets()[s].getDataRange();                                                                                    //Updates the gateway specific data range

      completion++;
      ss.getSheetByName("Data update").getDataRange().getCell(3,4).setValue(completion);
      ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3)
              .setValue("Populating the new deliverables in "+ss.getSheets()[s].getSheetName()+"...");
         
      var gatewaySpecificArray = new Array (3);                                                                                                      //Adds a new array which stores the values to be pasted into the specific sheet
      for (var i=0; i<3; i++){
          gatewaySpecificArray[i] = new Array (numDeliverables*2)
          }
      
      //POPULATES THE DELIVERABLES
      for (var j=0; j<numDeliverables; j++){                                                                                                         //Iterates through the deliverables
          gatewaySpecificArray[0][j*2] = deliverablesReferenceArray[numOtherDeliverables+j+4][2];                                                    //Updates the deliverable name
          gatewaySpecificArray[2][j*2] = deliverablesReferenceArray[numOtherDeliverables+j+4][3];                                                    //Updates the measure of completeness
          gatewaySpecificArray[0][j*2+1] = "";                                                                                                       //Resets the value of the cells around
          gatewaySpecificArray[1][j*2] = "";
          gatewaySpecificArray[1][j*2+1] = "";
          gatewaySpecificArray[2][j*2+1] = "";
          }     
      gatewaySpecificSheet.getRange(2,19, 3, numDeliverables*2).setValues(gatewaySpecificArray);
      }
  }      
      
  ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3).setValue("");
  ss.setActiveSheet(ss.getSheetByName("Overview"));                                                                                                 //Takes the user to the overview
  Browser.msgBox("Data update","Sheet restoration completed.", Browser.Buttons.OK);                                                                 //Shows a finishing message box
  ss.getSheetByName("Data update").hideSheet()                                                                                                      //Hides the completion sheeet
  ss.getSheetByName("Deliverables reference").hideSheet()

}