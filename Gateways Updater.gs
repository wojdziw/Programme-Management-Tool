/*
The purpose of this function is to alter the number and kind of the gateways appearing in the spreadsheet.
It makes changes both in the overview sheet and each of the gateway sheets.

Due to the nature of the scripts two types of comments were added. The one in capital letters above the 
bits of the code describe the general actions of the processes while the ones on the side of the script 
refer to particular action that the script performs to achieve its goal.
*/

function gatewayupdater(){

  Browser.msgBox("Data update","The data update will now start. The process is going to take from 20 seconds to 4 minutes.", Browser.Buttons.OK);    //Shows a starting message box

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
  for (var i=0; i<numFeatures; i++){                                                                                                                 //Iterates through all the features
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
  
  //INSERTS CHARTS TO THE OVERVIEW
  for (var s=1; s<=numGateways; s++){                                                                                                                //Iterates through the gateways to insert the graphs to the overview
      gatewaySpecificSheet = ss.getSheets()[s];                                                                                                      //Saves consecutive sheets
      numGatewaySpecificRows = gatewaySpecificSheet.getLastRow();                                                                                    //Saves the number of rows in the gateway specific sheet
      numGatewaySpecificCols = gatewaySpecificSheet.getLastColumn();                                                                                 //Saves the number of columnss in the gateway specific sheet

      //SAVES THE RANGES FOR THE CHARTS
      var pieChartRange = gatewaySpecificSheet.getRange(numGatewaySpecificRows-numFeatures-4,numGatewaySpecificCols-numFeatures*2-1,4,2);            //Saves the range for the pie chart
      var columnChartRange = gatewaySpecificSheet
          .getRange(numGatewaySpecificRows-numFeatures+1,numGatewaySpecificCols-numFeatures*2-1,numFeatures,5);                                      //Saves the range for the column chart
  
      //INSERTS A PIE CHART
      var pieChart = overviewSheet.newChart()                                                                                                        //Sets the pie chart's parameters:
          .setOption('width', 110)                                                                                                                   //width
          .setOption('height', 110)                                                                                                                  //height
          .setOption('theme','maximized')                                                                                                            //whether it should be maximized or not
          .setOption('legend', 'none')                                                                                                               //whether or not there should be a legend
          .setOption('fontSize', 9)                                                                                                                  //the font size of any text inside
          .setOption('pieSliceText','value')                                                                                                         //the type of number shown by the slices
          .setOption('colors', ['#00B050','#FFC000','#FF0000','#6FA8DC'])                                                                            //colours of the slices
          .setChartType(Charts.ChartType.PIE)                                                                                                        //whether it should be a pie chart
          .addRange(pieChartRange)                                                                                                                   //Adds the chart's range
          .setPosition(7+s*2,9, 37, 0)                                                                                                               //sets its position
          .build();                                                                                                                                  //creates the chart
      overviewSheet.insertChart(pieChart);                                                                                                           //inserts the chart into the sheet
  
      //INSERTS A COLUMN CHART
      var columnChart = overviewSheet.newChart()                                                                                                     //Sets the column chart's paremeters:
           .setOption('width', numFeatures*57+2)                                                                                                     //width
           .setOption('height', 110)                                                                                                                 //height
           .setOption('theme','maximized')                                                                                                           //whether it should be maximized or not
           .setOption('legend','none')                                                                                                               //whether or not there should be a legend
           .setOption('vAxis.textPosition', 'none')                                                                                                  //whether or not there should be any axis title
           .setOption('colors', ['#00B050','#FFC000','#FF0000','#6FA8DC'])                                                                           //colours
           .setChartType(Charts.ChartType.COLUMN)                                                                                                    //whether it should be a pie chart
           .addRange(columnChartRange)                                                                                                               //Adds the chart's range
           .setOption('hAxis', {'textPosition': 'none'})
           .setOption("isStacked", true)                                                                                                             //whether the chart should be stacked
           .setPosition(7+s*2,1,683,0)                                                                                                               //sets it position
           .build();                                                                                                                                 //creates the chart
      overviewSheet.insertChart(columnChart);                                                                                                        //inserts the chart into the sheet
      }

  //TAKES THE USER TO THE DELIVERABLE SHEET
  ss.setActiveSheet(ss.getSheetByName("Deliverables reference"));                                                                                    //Takes user to the Deliverables reference sheet
  Browser.msgBox("Data update","Now please review the deliverables in this sheet.", Browser.Buttons.OK);                                             //Shows a message box
  ss.getSheetByName("Gateways reference").hideSheet()
}