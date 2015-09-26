/*
The purpose of this function is to alter the number and kind of the deliverables appearing in the spreadsheet.
It makes changes both in the gateway sheets only.

Due to the nature of the scripts two types of comments were added. The one in capital letters above the 
bits of the code describe the general actions of the processes while the ones on the side of the script 
refer to particular action that the script performs to achieve its goal.
*/

function deliverableupdater(){

  var ss = SpreadsheetApp.getActiveSpreadsheet();

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
  ss.setActiveSheet(ss.getSheetByName("Overview"));                                                                                                  //Takes the user to the overview
  Browser.msgBox("Data update","Data update finished.", Browser.Buttons.OK);                                                                         //Shows a finishing message box
  ss.getSheetByName("Data update").hideSheet()                                                                                                       //Hides the completion sheeet
  ss.getSheetByName("Deliverables reference").hideSheet()
}