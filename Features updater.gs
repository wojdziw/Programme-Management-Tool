/*
The purpose of this function is to alter the number and kind of the features appearing in the spreadsheet.
It makes changes both in the overview sheet and each of the gateway sheets.

Due to the nature of the scripts two types of comments were added. The one in capital letters above the 
bits of the code describe the general actions of the processes while the ones on the side of the script 
refer to particular action that the script performs to achieve its goal.
*/

function featureupdater(){

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
  var totalToBeCompleted = 26*numGateways+9;                                                                                                        //Sets total number of checkpoints for the chart's reference
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

  ss.getSheetByName("Data update").getRange(1,1,8,6).getCell(7,3).setValue("");
  ss.setActiveSheet(ss.getSheetByName("Overview"));                                                                                                 //Takes the user to the overview
  Browser.msgBox("Data update","Data update finished.", Browser.Buttons.OK);                                                                        //Shows a finishing message box
  ss.getSheetByName("Data update").hideSheet()                                                                                                      //Hides the completion sheeet
  ss.getSheetByName("Features reference").hideSheet()
  
}