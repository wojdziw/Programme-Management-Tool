/*
The purpose of this function is to control the "Adding/deleting" buttons in the 'Deliverables reference', 
'Features reference' and 'Gateways reference' sheets. It is an onEdit function which means it is triggered
whenever user modifies the documented. In this particular script actions are undertaken if and only if
the "Adding/deleting" buttons are selected.
*/

function onEdit(event){

  var whichSheetWasChanged = event.source.getActiveSheet();                                                                                          //Registers which sheet was subject to a change

  if (whichSheetWasChanged.getName() == "Features reference"){                                                                                       //Checks if the changed sheet is the "Features reference" one
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var featuresReferenceSheet = ss.getSheetByName('Features reference');                                                                          //Saves the Features reference sheet
      var featureReferenceData = ss.getSheetByName('Features reference').getRange(1,1,100,9);                                                        //Saves the Features reference data range
      var changedCell= event.source.getActiveRange();                                                                                                //Registers the changed cell
      var changedCellColumn = changedCell.getColumn();                                                                                               //Registers the changed cell's column index
      var changedCellRow = changedCell.getRow();                                                                                                     //Registers the changed cell's row index

      if (featureReferenceData.getCell(changedCellRow,changedCellColumn).getValue()=="Add Below"){                                                   //Checks if the changed cell indicates adding a row
          
          featuresReferenceSheet.insertRowAfter(changedCellRow);                                                                                     //Adds a row below
          
          featureReferenceData.getCell(changedCellRow,changedCellColumn).setValue("-");                                                              //Resets the content of the changed cell (from "Add Below" to "-")
        
          var sourceOfDataValidation = featuresReferenceSheet.getRange(changedCellRow,changedCellColumn,1,1);                                        //Saves the "Adding/deleting" cell (there is no way to create data validation programatically therefore appropriate formatting must be copied)
          var destinationOfDataValidation = featuresReferenceSheet.getRange(changedCellRow+1,changedCellColumn,1,1);                                 //Saves the "Adding/deleting" cell in the new row
          sourceOfDataValidation.copyTo(destinationOfDataValidation);                                                                                //Copies data validation to the new row
          
          featureReferenceData.getCell(changedCellRow+1,changedCellColumn-6)                                                                         //Resets the values and formatting and sets the borders in the new cells (some borders may not seem needed in theory but in fact they don't appear otherwise)
              .setValue("")
              .setBorder(true,true,true,true,true,true);
          featureReferenceData.getCell(changedCellRow+1,changedCellColumn-5)
              .setValue("")
              .setBorder(true,true,true,true,true,true);
          featureReferenceData.getCell(changedCellRow+1,changedCellColumn-4)
              .setValue("")
              .setBorder(true,true,true,true,true,true);
          featureReferenceData.getCell(changedCellRow+1,changedCellColumn-3)
              .setValue("")
              .setBackground("#FFFFFF")
              .setBorder(true,true,true,true,true,true);
          featureReferenceData.getCell(changedCellRow+1,changedCellColumn-2)
              .setValue("")
              .setBackground("#000000")
              .setBorder(true,true,true,true,true,true);
          featureReferenceData.getCell(changedCellRow+1,changedCellColumn)
              .setValue("-")
              .setBorder(true,true,true,true,true,true);
          }
  
      if (featureReferenceData.getCell(changedCellRow,changedCellColumn).getValue() == "Delete"){                                                    //Checks if the changed cell indicates deleting a row
          featuresReferenceSheet.deleteRows(changedCellRow,1);                                                                                       //Deletes the desired row
          }
  }

  if (whichSheetWasChanged.getName() == "Deliverables reference"){                                                                                   //Checks if the changed sheet is the "Deliverables reference" one
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var deliverablesReferenceSheet = ss.getSheetByName('Deliverables reference');                                                                  //Saves the Deliverables reference sheet
      var deliverableReferenceData = ss.getSheetByName('Deliverables reference').getRange(1,1,700,7);                                                //Saves the Deliverables reference data range
      var changedCell = event.source.getActiveRange();                                                                                               //Registers the changed cell
      var changedCellColumn = changedCell.getColumn();                                                                                               //Registers the changed cell's column index
      var changedCellRow = changedCell.getRow();                                                                                                     //Registers the changed cell's row index

      if (deliverableReferenceData.getCell(changedCellRow,changedCellColumn).getValue() == "Add Below"){                                             //Checks if the changed cell indicates adding a row
      
          deliverablesReferenceSheet.insertRowAfter(changedCellRow);                                                                                 //Adds a row below
          
          deliverableReferenceData.getCell(changedCellRow,changedCellColumn).setValue("-");                                                          //Resets the content of the changed cell (from "Add Below" to "-")
          
          var sourceOfFormatting = deliverablesReferenceSheet.getRange(changedCellRow,changedCellColumn-4,1,5);                                      //Saves the changed row formatting (there is no way to create data validation programatically therefore appropriate formatting must be copied)
          var destinationOfFormatting = deliverablesReferenceSheet.getRange(changedCellRow+1,changedCellColumn-4,1,5);                               //Saves the newly created row
          sourceOfFormatting.copyTo(destinationOfFormatting);                                                                                        //Copies formatting to the new row
          
          deliverableReferenceData.getCell(changedCellRow+1,changedCellColumn-4)                                                                     //Resets the values and formatting and sets the borders in the new cells (some borders may not seem needed in theory but in fact they don't appear otherwise)
              .setValue("")
              .setBorder(true,true,true,true,true,true);
          deliverableReferenceData.getCell(changedCellRow+1,changedCellColumn-3)
              .setValue("")
              .setBorder(true,true,true,true,true,true);
          deliverableReferenceData.getCell(changedCellRow+1,changedCellColumn-2)
              .setValue("")
              .setBorder(true,true,true,true,true,true);
          deliverableReferenceData.getCell(changedCellRow+1,changedCellColumn)
              .setValue("-")
              .setBorder(true,true,true,true,true,true);
          deliverableReferenceData.getCell(changedCellRow+1,changedCellColumn+1)
              .setBorder(false, true, false, false, false, false);
          deliverableReferenceData.getCell(changedCellRow,changedCellColumn+1)
              .setBorder(false, null, false, false, false, false);
          deliverableReferenceData.getCell(changedCellRow+2,changedCellColumn+1)
              .setBorder(false, null, false, false, false, false);
          }

      if (deliverableReferenceData.getCell(changedCellRow,changedCellColumn).getValue() == "Delete"){                                                //Checks if the changed cell indicates deleting a row
          deliverablesReferenceSheet.deleteRows(changedCellRow,1);                                                                                   //Deletes the desired row
          }
  }


  if (whichSheetWasChanged.getName() == "Gateways reference"){                                                                                       //Checks if the changed sheet is the "Gateways reference" one
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var gatewaysReferenceSheet = ss.getSheetByName('Gateways reference');                                                                              //Saves the Gateways reference sheet
  var gatewaysReferenceData = ss.getSheetByName('Gateways reference').getRange(1,1,700,7);                                                           //Saves the Gateways reference data range
  var changedCell= event.source.getActiveRange();                                                                                                    //Registers the changed cell
  var changedCellColumn = changedCell.getColumn();                                                                                                   //Registers the changed cell's column index
  var changedCellRow = changedCell.getRow();                                                                                                         //Registers the changed cell's row index

      if (gatewaysReferenceData.getCell(changedCellRow,changedCellColumn).getValue() == "Add Below"){                                                //Checks if the changed cell indicates adding a row
      
          gatewaysReferenceSheet.insertRowAfter(changedCellRow);                                                                                     //Adds a row below
          
          gatewaysReferenceData.getCell(changedCellRow,changedCellColumn).setValue("-");                                                             //Resets the content of the changed cell (from "Add Below" to "-")
          
          var sourceOfFormatting = gatewaysReferenceSheet.getRange(changedCellRow,changedCellColumn-4,1,5);                                          //Saves the changed row formatting (there is no way to create data validation programatically therefore appropriate formatting must be copied)
          var destinationOfFormatting = gatewaysReferenceSheet.getRange(changedCellRow+1,changedCellColumn-4,1,5);                                   //Saves the newly created row
          sourceOfFormatting.copyTo(destinationOfFormatting);                                                                                        //Copies formatting to the new row
          
          gatewaysReferenceData.getCell(changedCellRow+1,changedCellColumn-4)                                                                        //Resets the values and formatting and sets the borders in the new cells (some borders may not seem needed in theory but in fact they don't appear otherwise)
              .setValue("")
              .setBorder(true,true,true,true,true,true);
          gatewaysReferenceData.getCell(changedCellRow+1,changedCellColumn)
              .setValue("-")
              .setBorder(true,true,true,true,true,true);  
          gatewaysReferenceData.getCell(changedCellRow+1,changedCellColumn+1)
              .setBorder(false, true, false, false, false, false);
          gatewaysReferenceData.getCell(changedCellRow,changedCellColumn+1)
              .setBorder(false, null, false, false, false, false);
          gatewaysReferenceData.getCell(changedCellRow+2,changedCellColumn+1)
              .setBorder(false, null, false, false, false, false);
          }
  
      if (gatewaysReferenceData.getCell(changedCellRow,changedCellColumn).getValue()=="Delete"){                                                     //Checks if the changed cell indicates deleting a row
          ss.getSheetByName('Gateways reference').deleteRows(changedCellRow,1);                                                                      //Deletes the desired row
          }

  }
}