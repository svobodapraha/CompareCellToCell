   //start compare real
    //find position of "Object Type" Column and add headers to list

    int iNewColObjectType = -1;
    QStringList lstNewHeaders;
    lstNewHeaders.clear();
    QMap<QString, int> mapNewHeaders;
    mapNewHeaders.clear();

    for(int col= kn_ReqIDCol; col <= newLastCol; col++)
    {
      QString asTemp = getCellValue(kn_HeaderRow, col, newReqDoc, true);
      asTemp.replace("."," ");
      if (ui->cb_DetectOriginalID->isChecked())
      {
         if(ignoreWhiteCaseAndSpaces(asTemp).contains("idoriginal")) asTemp = "ID Original (DXL)";
      }
      lstNewHeaders << asTemp;
      mapNewHeaders[asTemp] = col;

      if(ignoreWhiteCaseAndSpaces(asTemp).contains("objecttype"))
      {
        iNewColObjectType = col;
      }
    }


    int iOldColObjectType = -1;
    QStringList lstOldHeaders;
    lstOldHeaders.clear();
    QMap<QString, int> mapOldHeaders;
    mapOldHeaders.clear();

    for(int col= kn_ReqIDCol; col <= oldLastCol; col++)
    {
        QString asTemp = getCellValue(kn_HeaderRow, col, oldReqDoc, true);
        asTemp.replace("."," ");
        if (ui->cb_DetectOriginalID->isChecked())
        {
            if(ignoreWhiteCaseAndSpaces(asTemp).contains("idoriginal")) asTemp = "ID Original (DXL)";

        }
        lstOldHeaders << asTemp;
        mapOldHeaders[asTemp] = col;

        if(ignoreWhiteCaseAndSpaces(asTemp).contains("objecttype"))
        {
          iOldColObjectType = col;
        }
    }


    //qDebug() << "Object Type Col" << iNewColObjectType << iOldColObjectType;
    //qDebug() << lstNewHeaders << mapNewHeaders;
    //qDebug() << lstOldHeaders << mapOldHeaders;

    //intersection - headers in both files..
    QStringList lstNewOldHeader =  (QSet<QString>(lstNewHeaders.begin(), lstNewHeaders.end())
                                   .intersect(QSet<QString>(lstOldHeaders.begin(), lstOldHeaders.end()))).values();

    lstNewOldHeader.sort();
    //qDebug() << lstNewOldHeader;

    //Test Columns
    foreach (QString asTemp, lstNewOldHeader)
    {
      //qDebug() << "new: " << mapNewHeaders[asTemp] << "old: " << mapOldHeaders[asTemp];
    }

    //TODO exit when Object Type not found


    QStringList lstNewReqIDs;
    lstNewReqIDs.clear();
    QStringList lstOldReqIDs;
    lstOldReqIDs.clear();

    //START COMPARING - FIRST LOOK TO THE OLD FILES...
    //read all rows in old files
    for (int oldRow = kn_FistDataRow; oldRow <= oldLastRow; ++oldRow)
    {
       QString asOldReqID = getCellValue(oldRow, kn_ReqIDCol, oldReqDoc,  true);
       //ignore project prefix
       if(ui->cb_IgnoreProject->isChecked())
       {
           asOldReqID = asOldReqID.remove(0, asOldReqID.indexOf("TSAnS"));
       }


       //USE GUID
       if(ui->cb_UseGUID->isChecked())
       {
           int iMappedColumn = mapOldHeaders["GUID"];
           if(iMappedColumn > 0)
           {
             asOldReqID = getCellValue(oldRow, iMappedColumn, oldReqDoc, true);
           }
           else
           {
               if (!boBatchProcessing)
               {
                   QMessageBox::information(this, "Error", "No GUID in old", QMessageBox::Ok);
               }
               else
               {
                  qWarning() << "Error " <<  "No GUID in old";
                  //TODO exit code
               }
               return;
            }
       }





       lstOldReqIDs << asOldReqID;
       int iOldReqID      = -1; //TODO
       Q_UNUSED(iOldReqID)

       //FOR CURRENT ITEM IN THE OLD FILE CHECK EVERY ITEM IN NEW FILE
       //and compare with every row in new file..
       bool bo_idOldFound = false;
       for(int newRow = kn_FistDataRow; newRow <= newLastRow; ++newRow)
       {
         QString asNewReqIDForCompare = getCellValue(newRow, kn_ReqIDCol, newReqDoc, true);
         //ignore project prefix
         if(ui->cb_IgnoreProject->isChecked())
         {
              asNewReqIDForCompare = asNewReqIDForCompare.remove(0, asNewReqIDForCompare.indexOf("TSAnS"));
         }
         //real new ID
         QString asNewReqID = asNewReqIDForCompare;

         //USE GUID
         if(ui->cb_UseGUID->isChecked())
         {
             int iMappedColumn = mapNewHeaders["GUID"];
             if(iMappedColumn > 0)
             {
               asNewReqIDForCompare = getCellValue(newRow, iMappedColumn, newReqDoc, true);
             }
             else
             {
                 if (!boBatchProcessing)
                 {
                     QMessageBox::information(this, "Error", "No GUID in new", QMessageBox::Ok);
                 }
                 else
                 {
                    qWarning() << "Error " <<  "No GUID in new";
                    //TODO exit code
                 }
                 return;

             }

         }


         //USE ID Original (DXL)
         if(ui->cb_UseOriginalID->isChecked())
         {
             int iMappedColumn = mapNewHeaders["ID Original (DXL)"];
             if(iMappedColumn > 0)
             {
               asNewReqIDForCompare = getCellValue(newRow, iMappedColumn, newReqDoc, true);
             }
             else
             {
                 if (!boBatchProcessing)
                 {
                     QMessageBox::information(this, "Error", "No ID Original (DXL) in new", QMessageBox::Ok);
                 }
                 else
                 {
                    qWarning() << "Error " <<  "No ID Original (DXL) in new";
                    //TODO exit code
                 }
                 return;
              }

         }







         int iNewReqID    = -1; //TODO
         Q_UNUSED(iNewReqID)
         //add ids only once
         if(oldRow == kn_FistDataRow)
         {
           lstNewReqIDs << asNewReqIDForCompare;
         }

         //if ID requirments are same, check every column which are in both files
         //qDebug() << "1" << asOldReqID << asNewReqID << asNewReqIDForCompare;
         if(asOldReqID == asNewReqIDForCompare)
         {
             //qDebug() << "2" << asOldReqID << asNewReqID << asNewReqIDForCompare;
             bo_idOldFound = true;
             bool boSameReq = false;
             //and now check every column which is in both files.. //TODO maybe better in at least one!!
             foreach (QString asHeader, lstNewOldHeader)
             {
                //ignore column with links
                if(ui->cb_IgnoreLinks->isChecked())
                {
                   if(asHeader.contains("Out-links", Qt::CaseInsensitive)) continue;
                }

                //ignore other columns
                if(ui->cb_IgnoreOthers->isChecked())
                {
                   if(asHeader.contains("ID Original", Qt::CaseInsensitive)) continue;
                   if(asHeader.contains("Source",      Qt::CaseInsensitive)) continue;
                   if(asHeader.contains("GUID",        Qt::CaseInsensitive)) continue;
                   if(asHeader.contains("Preview",     Qt::CaseInsensitive)) continue;

                }

                if(ui->cb_IgnoreID->isChecked())
                {
                   if(asHeader == "ID") continue;
                }

                if(ui->cb_TextOnly->isChecked())
                {
                   if(asHeader.toLower() != "text") continue;
                }


                QString asNewColValue =  getCellValue(newRow, mapNewHeaders[asHeader], newReqDoc, true);
                QString asOldColValue =  getCellValue(oldRow, mapOldHeaders[asHeader], oldReqDoc, true);
                if(ui->cb_IngnoreWaC->isChecked())
                {
                   asNewColValue = ignoreWhiteAndCase(asNewColValue);
                   asOldColValue = ignoreWhiteAndCase(asOldColValue);
                }
                //qDebug() << getCellValue(newRow, iNewColObjectType, newReqDoc, true) << getCellValue(oldRow, iOldColObjectType, oldReqDoc, true);

                //compare value for parameters..
                if(
                     (asOldColValue != asNewColValue) &&
                     (
                       (
                          (getCellValueWCIgnored(newRow, iNewColObjectType, newReqDoc) == "functional requirement") ||
                          (getCellValueWCIgnored(oldRow, iOldColObjectType, oldReqDoc) == "functional requirement") ||
                          (!ui->cbOnlyFuncReq->isChecked())
                       )
                     )

                  )
                //different
                {
                   //qDebug() << "different:";
                   //qDebug() << asOldReqID << asHeader <<  asOldColValue << asNewColValue;

                   ui->tableWidget_Changes->insertRow(ui->tableWidget_Changes->rowCount());
                   ui->tableWidget_Changes->setItem(ui->tableWidget_Changes->rowCount()-1, col_infotable_requirement, new QTableWidgetItem(asNewReqID));
                   QString asShort = "";
                   if (!ui->cb_UseGUID->isChecked())
                   {
                      asShort = asNewReqID.section("_", -1);
                   }
                   else
                   {
                      asShort = getCellValue(oldRow, kn_ReqIDCol, oldReqDoc,  true).section("_", -1) + "/" +
                                getCellValue(newRow, kn_ReqIDCol, newReqDoc, true).section("_", -1);
                   }
                   ui->tableWidget_Changes->setItem(ui->tableWidget_Changes->rowCount()-1, col_infotable_req_short, new QTableWidgetItem(asShort));

                   //another column from the same request.. - set Requirement column to light grey backroud
                   if(boSameReq)
                   {
                       ui->tableWidget_Changes->item(ui->tableWidget_Changes->rowCount()-1, col_infotable_requirement)->setBackground(QBrush(QColor(Qt::lightGray)));
                       ui->tableWidget_Changes->item(ui->tableWidget_Changes->rowCount()-1, col_infotable_req_short)->setBackground(QBrush(QColor(Qt::lightGray)));

                       //denote previous line also - should belog to the same req id
                       if (ui->tableWidget_Changes->rowCount() > 1)
                       {
                           QTableWidgetItem * tmpTableWidgetItem = nullptr;
                           tmpTableWidgetItem = ui->tableWidget_Changes->item((ui->tableWidget_Changes->rowCount()-1)-1, col_infotable_requirement);
                           if(tmpTableWidgetItem)tmpTableWidgetItem->setBackground(QBrush(QColor(Qt::lightGray)));

                           tmpTableWidgetItem = ui->tableWidget_Changes->item((ui->tableWidget_Changes->rowCount()-1)-1, col_infotable_req_short);
                           if(tmpTableWidgetItem)tmpTableWidgetItem->setBackground(QBrush(QColor(Qt::lightGray)));
                       }
                   }
                   boSameReq = true;


                   ui->tableWidget_Changes->setItem(ui->tableWidget_Changes->rowCount()-1, col_infotable_status,    new QTableWidgetItem("Changed"));
                   ui->tableWidget_Changes->setItem(ui->tableWidget_Changes->rowCount()-1, col_infotable_par_name,  new QTableWidgetItem(asHeader));
                   ui->tableWidget_Changes->setItem(ui->tableWidget_Changes->rowCount()-1, col_infotable_old_value, new QTableWidgetItem(asOldColValue));
                   ui->tableWidget_Changes->setItem(ui->tableWidget_Changes->rowCount()-1, col_infotable_new_value, new QTableWidgetItem(asNewColValue));
                   QString asOrigID = "";
                   if(ui->cb_FindSimilarFromNew->isChecked())
                     asOrigID = getCellValue(oldRow, mapOldHeaders[knOrigIDText], oldReqDoc, true);
                   if(ui->cb_UseOriginalID->isChecked())
                     asOrigID = asOldReqID;
                   ui->tableWidget_Changes->setItem(ui->tableWidget_Changes->rowCount()-1, col_infotable_origID, new QTableWidgetItem(asOrigID));

                   //if compare without demand for ignorig white space and case, compare once more ignored and if same (e.g. difference is only in spaces and case
                   //set the backroud to light grey
                   if (!ui->cb_IngnoreWaC->isChecked())
                   {

                       asNewColValue = ignoreWhiteAndCase(asNewColValue);
                       asOldColValue = ignoreWhiteAndCase(asOldColValue);
                       if(asOldColValue == asNewColValue)
                       {
                           ui->tableWidget_Changes->item(ui->tableWidget_Changes->rowCount()-1, col_infotable_old_value)->setBackground(QBrush(QColor(Qt::lightGray)));
                           ui->tableWidget_Changes->item(ui->tableWidget_Changes->rowCount()-1, col_infotable_new_value)->setBackground(QBrush(QColor(Qt::lightGray)));
                       }

                   }


                   //set status yellow backround if not function requiments
                   if(
                         (getCellValueWCIgnored(newRow, iNewColObjectType, newReqDoc) != "functional requirement") &&
                         (getCellValueWCIgnored(oldRow, iOldColObjectType, oldReqDoc) != "functional requirement")
                     )
                     {
                       ui->tableWidget_Changes->item(ui->tableWidget_Changes->rowCount()-1, col_infotable_status)->setText("Changed(not FR)");
                       ui->tableWidget_Changes->item(ui->tableWidget_Changes->rowCount()-1, col_infotable_status)->setBackground(QBrush(QColor(Qt::yellow)));
                     }

                }

                //if there will be another diff in the same request (another column)
             }//foreach (QString asHeader, stNewOldHeader)
         }//if(asOldReqID == asNewReqID)

       }//for(int newRow = kn_FistDataRow; newRow <= newLastRow; ++newRow)

       //not found in new - only in old req
       if(!ui->cb_HideMissing->isChecked())
       {
           if(!bo_idOldFound)
           {
               //qDebug() << "not in new file:";
               //qDebug() << asOldReqID;
               ui->tableWidget_Changes->insertRow(ui->tableWidget_Changes->rowCount());
               ui->tableWidget_Changes->setItem(ui->tableWidget_Changes->rowCount()-1, col_infotable_requirement, new QTableWidgetItem(asOldReqID));
               ui->tableWidget_Changes->setItem(ui->tableWidget_Changes->rowCount()-1, col_infotable_req_short, new QTableWidgetItem(getCellValue(oldRow, kn_ReqIDCol, oldReqDoc,  true).section("_", -1)));
               ui->tableWidget_Changes->setItem(ui->tableWidget_Changes->rowCount()-1, col_infotable_status,      new QTableWidgetItem("Missing"));
               ui->tableWidget_Changes->item(ui->tableWidget_Changes->rowCount()-1, col_infotable_status)->setBackground(QBrush(QColor(Qt::red)));
           }
       }

    }//for (int oldRow = kn_FistDataRow; oldRow <= oldLastRow; ++oldRow)

    //TO DO CHECK

    //Find new without reference to OLD or mising in OLD
    QStringList lstNewReqIDsOnly;
    lstNewReqIDsOnly.clear();

    if (ui->cb_UseOriginalID->isChecked())
    {
        for(int newRow = kn_FistDataRow; newRow <= newLastRow; ++newRow)
        {

          QString asIDOrigin;
          asIDOrigin.clear();
          QString asNewReqID = getCellValue(newRow, kn_ReqIDCol, newReqDoc, true);
          //ignore project prefix
          if(ui->cb_IgnoreProject->isChecked())
          {
               asNewReqID = asNewReqID.remove(0, asNewReqID.indexOf("TSAnS"));
          }
          //real new ID


          int iMappedColumn = mapNewHeaders["ID Original (DXL)"];
          if(iMappedColumn > 0)
          {
            asIDOrigin = getCellValue(newRow, iMappedColumn, newReqDoc, true);
          }
          else
          {
            asIDOrigin.clear();
          }


          if(ui->cb_IgnoreProject->isChecked())
          {
             asIDOrigin = asIDOrigin.remove(0, asIDOrigin.indexOf("TSAnS"));
          }

          //if asIDOrigin is missing or set to N/A add to new list
           if(!lstOldReqIDs.contains(asIDOrigin))
          {
             lstNewReqIDsOnly << asNewReqID;

          }




        }//for(int newRow = kn_FistDataRow; newRow <= newLastRow; ++newRow)

    }
    else
    {
        lstNewReqIDsOnly =  (QSet<QString>(lstNewReqIDs.begin(), lstNewReqIDs.end())
                                     .subtract(QSet<QString>(lstOldReqIDs.begin(), lstOldReqIDs.end()))).values();
    }

    qDebug() << "new: " << lstNewReqIDs;
    qDebug() << "old: " << lstOldReqIDs;
    qDebug() << "newonly: " << lstNewReqIDsOnly;


    lstNewReqIDsOnly.sort();

    if(!ui->cb_HideNew->isChecked())
    {

        //Only in new file (new req)
        foreach (QString asNewIDOnly, lstNewReqIDsOnly)
        {
           //qDebug() << "New ID only "  << asNewIDOnly;
           ui->tableWidget_Changes->insertRow(ui->tableWidget_Changes->rowCount());
           ui->tableWidget_Changes->setItem(ui->tableWidget_Changes->rowCount()-1, col_infotable_requirement, new QTableWidgetItem(asNewIDOnly));
           if (!ui->cb_UseGUID->isChecked())
             ui->tableWidget_Changes->setItem(ui->tableWidget_Changes->rowCount()-1, col_infotable_req_short, new QTableWidgetItem(asNewIDOnly.section("_", -1)));
           else
             ui->tableWidget_Changes->setItem(ui->tableWidget_Changes->rowCount()-1, col_infotable_req_short, new QTableWidgetItem(""));  //TODO
           ui->tableWidget_Changes->setItem(ui->tableWidget_Changes->rowCount()-1, col_infotable_status,      new QTableWidgetItem("New"));
           ui->tableWidget_Changes->item(ui->tableWidget_Changes->rowCount()-1, col_infotable_status)->setBackground(QBrush(QColor(Qt::green)));

        }
    }


    //adjust table width
    ui->tableWidget_Changes->resizeColumnsToContents();
    ui->tableWidget_Changes->setColumnWidth(col_infotable_old_value,250);
    ui->tableWidget_Changes->setColumnWidth(col_infotable_new_value,250);

    //Add checkboxes to the table
    for (int irow = 0; irow < ui->tableWidget_Changes->rowCount(); ++irow)
    {
        QCheckBox *cb = new QCheckBox(this);
        ui->tableWidget_Changes->setCellWidget(irow, col_infotable_check_box, cb);
    }

    //set checkboxes to checked state
    checkAll(true);

    //all cell in table except user comments should not editable and not selectable
    setChangesTableNotEditable(); //except user comments

    //enable sorting
    //ui->tableWidget_Changes->setSortingEnabled(true);


    if (!boBatchProcessing)
    {
        QMessageBox::information(this, "Compared", "Lines: " + QString::number(ui->tableWidget_Changes->rowCount()), QMessageBox::Ok);
    }
    else
    {
       qWarning() << "Compared " <<  QString::number(ui->tableWidget_Changes->rowCount()) << " Lines";
    }

    asCompareCondition = "";
    if(ui->cbOnlyFuncReq->isChecked())          asCompareCondition += ":Only Function Requiments:";
    if(ui->cb_IngnoreWaC->isChecked())          asCompareCondition += ":Ignore White Spaces and Case:";
    if(ui->cb_IgnoreProject->isChecked())       asCompareCondition += ":Ignore Project:";
    if(ui->cb_IgnoreLinks->isChecked())         asCompareCondition += ":Ignore Links:";
    if(ui->cb_IgnoreOthers->isChecked())        asCompareCondition += ":Ignore Others:";
    if(ui->cb_FindSimilarFromNew->isChecked())  asCompareCondition += ":Find Similar:";
    if(ui->cb_UseGUID->isChecked())             asCompareCondition += ":Use GUID:";
    if(ui->cb_IgnoreID->isChecked())            asCompareCondition += ":Ignore ID:";
    if(ui->cb_HideMissing->isChecked())         asCompareCondition += ":Hide Missing:";
    if(ui->cb_HideNew->isChecked())             asCompareCondition += ":Hide New:";
    if(ui->cb_UseOriginalID->isChecked())       asCompareCondition += ":Use Original ID:";
    if(ui->cb_DetectOriginalID->isChecked())    asCompareCondition += ":Detect Original ID:";
    if(ui->cb_TextOnly->isChecked())             asCompareCondition += ":Text Only:";
    if(boBatchProcessing)                       asCompareCondition += ":Batch Processing:";

//Try to find if some req with differnt ID did not match in some parts
    if(ui->cb_FindSimilarFromNew->isChecked())
    {
      QXlsx::Document oldReqDocCor(fileName_OldReq);
      //insert column for Original ID
      for (int oldRow = kn_HeaderRow; oldRow <= oldLastRow; ++oldRow)
      {
         for(int col= oldLastCol; col > kn_ReqIDCol; col--)
         {
            oldReqDocCor.write(oldRow, col+1, getCellValue(oldRow, col, oldReqDocCor));
            if(col == kn_ReqIDCol+1)
                oldReqDocCor.write(oldRow, col, "");

         }
      }
      //Add Header for Original ID
      oldReqDocCor.write(kn_HeaderRow, kn_ReqIDCol + 1, knOrigIDText);

      //similarView->clear();
      foreach (QString asNewIDOnly, lstNewReqIDsOnly)
      {
        //find again - we need data from another columns
        for(int newRow = kn_FistDataRow; newRow <= newLastRow; ++newRow)
        {  	   
          QString asNewReqID = getCellValue(newRow, kn_ReqIDCol, newReqDoc, true);
          //ignore project prefix
          if(ui->cb_IgnoreProject->isChecked())
          {
            asNewReqID = asNewReqID.remove(0, asNewReqID.indexOf("TSAnS"));
          }
            
          //save data from "Requirement" and "Architecture Remark" column
          if(asNewReqID == asNewIDOnly)
          {
            QString asNewRequirement = getCellValueWCIgnored(newRow, mapNewHeaders["Requirement"], newReqDoc);
            QString asNewArchRemark     = getCellValueWCIgnored(newRow, mapNewHeaders["Architecture Remark"], newReqDoc);
            //qDebug() << asNewReqID << asArchRemark;
            //try to find in old doc
            for (int oldRow = kn_FistDataRow; oldRow <= oldLastRow; ++oldRow)
            {
              QString asOldReqID = getCellValue(oldRow, kn_ReqIDCol, oldReqDoc,  true);
              //ignore project prefix
              if(ui->cb_IgnoreProject->isChecked())
              {
                asOldReqID = asOldReqID.remove(0, asOldReqID.indexOf("TSAnS"));
              }
              QString asOldRequirement = getCellValueWCIgnored(oldRow, mapNewHeaders["Requirement"], oldReqDoc);
              QString asOldArchRemark  = getCellValueWCIgnored(oldRow, mapNewHeaders["Architecture Remark"], oldReqDoc);
              if(
                   ((asNewRequirement == asOldRequirement) && (!asNewRequirement.isEmpty())) ||
                   ((asNewArchRemark  == asOldArchRemark)  && (!asNewArchRemark.isEmpty()))
                )
              {
                oldReqDocCor.write(oldRow, kn_ReqIDCol, asNewReqID);
                oldReqDocCor.write(oldRow, kn_ReqIDCol + 1, asOldReqID);

//                similarView->addRow(asNewReqID,
//                                    asOldReqID,
//                                    asNewRequirement,
//                                    asOldRequirement,
//                                    asNewArchRemark,
//                                    asOldArchRemark);
//                //qDebug() << oldRow << newRow
//                         << asNewReqID
//                         << asOldReqID
//                         << asNewRequirement
//                         << asOldRequirement
//                         << asNewArchRemark;
              }//if
            }//for (int oasOldArchRemark);ldRow = kn_FistDataRow; oldRow <= oldLastRow; ++oldRow)
          }//if(asNewReqID == asNewIDOnly)
        }//for(int newRow = kn_FistDataRow; newRow <= newLastRow; ++newRow)
      }//foreach (QString asNewIDOnly, lstNewReqIDsOnly)

      //write corrected file
      if(!oldReqDocCor.saveAs(fileName_OldReqCor))
      {

          if (!boBatchProcessing)
          {
              QMessageBox::information(this, "Problem", "Problem write corrected file (opened?)", QMessageBox::Ok);
          }
          else
          {
             qCritical() << "Error: "<< "Problem write corrected file (opened?)";
             iExitCode = knExitStatusCorrectedFileOpened;
          }

      }
      else
      {

          if (!boBatchProcessing)
          {
              QMessageBox::information(this, "Corrected file written", fileName_OldReqCor, QMessageBox::Ok);
          }
          else
          {
             qWarning() << "Success: " << "Corrected file writen to " + fileName_OldReqCor;
          }

      }
    }//if(ui->cb_FindSimilarFromNew.isChecked())

    //allow export, disable similar
    ui->cb_FindSimilarFromNew->setChecked(false);
    ui->btnWrite->setEnabled(true);



QStringList tableHeader = QStringList() <<"Sel."
                                              <<"Sheet"
                                              <<"Row"
                                              <<"Column"
                                              <<"Old Value"
                                              <<"New Value"
                                              <<"User Note"



#define col_infotable_check_box         0
#define col_Sheet                       1
#define col_Row                         2
#define col_Column                      3
#define col_OldValue                    4
#define col_NewValue                    5
#define col_UserNote                    6





