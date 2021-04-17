#include "mainwindow.h"
#include "ui_mainwindow.h"



#define nullptr NULL

QString getCellValue(int row, int col, const QXlsx::Document &doc, bool boStrip = false)
{
    QString asValue;
    Cell* cell = doc.cellAt(row, col);
    if(cell != nullptr)
    {
        QVariant var = cell->readValue();
        asValue = var.toString();
    }
    else
    {
        asValue = "";
    }

    if (boStrip)
    {
      asValue = asValue.trimmed();
      asValue.replace("\r"," ").replace("\n", " ");
    }

    return asValue;

}

QString getCellValueValue(int row, int col, const QXlsx::Document &doc, bool boStrip = false)
{
    QString asValue;
    Cell* cell = doc.cellAt(row, col);
    if(cell != nullptr)
    {
        asValue = cell->value().toString();
    }
    else
    {
        asValue = "";
    }

    if (boStrip)
    {
      asValue = asValue.trimmed();
      asValue.replace("\r"," ").replace("\n", " ");
    }

    return asValue;

}



//ignore white spaces except one space
QString getCellValueWCIgnored(int row, int col, const QXlsx::Document &doc)
{
    QString asValue;
    Cell* cell = doc.cellAt(row, col);
    if(cell != nullptr)
    {
        QVariant var = cell->readValue();
        asValue = var.toString();
    }
    else
    {
        asValue = "";
    }

    asValue = asValue.simplified();
    asValue = asValue.toUpper();
    asValue = asValue.trimmed();

    return asValue;

}

QString ignoreWhiteAndCase(QString asValue)
{
    asValue = asValue.simplified();
    asValue = asValue.toUpper();

    return asValue;
}

QString ignoreWhiteCaseAndSpaces(QString asValue)
{
    asValue = asValue.simplified();
    asValue = asValue.toUpper();
    asValue.replace(" ","");

    return asValue;
}

int CsvToExcel(QString fileNameCsv, QXlsx::Document &xlsxDocument)
{
    //7.6.2020 TODO .. include /r to text inside the quotes...until now not working...

    //Open file and convert it to QTextStrem

    QFile inputFile(fileNameCsv);
    if(!inputFile.open(QIODevice::ReadOnly))
    {
        return(knExitStatusCannotLoadInputCsvFiles);
    }
    QTextStream inputStream(&inputFile);


    //List of column lists
    QList <QStringList> lsCSVDoc;
    lsCSVDoc.clear();

    //Column list
    QStringList lstOneRowFields;
    lstOneRowFields.clear();

    QString asDelimiter = ";";
    QString asCurrentField = "";
    QString oneChar;

    //iterate over the characters
    bool boInQuotes = false;
    bool boQuotesInQuotes = false;
    bool boNewLine = false;
    lstOneRowFields.clear();
    while (!inputStream.atEnd())
    {
        boNewLine = false;
        oneChar = inputStream.read(1);
        if(oneChar == "\r") continue;

        //last character is " after quote section. Could be end of quote or escape for "
        if (boQuotesInQuotes)
        {
            if (oneChar == "\"")
            {
               boQuotesInQuotes = false;
               asCurrentField+=oneChar;
               continue;
            }
            else
            {
                 boQuotesInQuotes = false;
                 boInQuotes = false;
            }
        }
        else
        {

        }

        if (!boInQuotes && oneChar == "\"" )
        {
             boInQuotes = true;
        }
        else if (boInQuotes && oneChar== "\"" )
        {
            boQuotesInQuotes = true;
        }
        else if (!boInQuotes && oneChar == asDelimiter)
        {
            lstOneRowFields << asCurrentField;
            asCurrentField.clear();
        }
        else if (!boInQuotes && (oneChar == "\n"))
        {
            lstOneRowFields << asCurrentField;
            asCurrentField.clear();
            boNewLine = true;
            //new row
            lsCSVDoc << lstOneRowFields;
            lstOneRowFields.clear();
        }
        else
        {
            //do not add CR if it is outside of quoted text: if(!(!boInQuotes &&  oneChar == "\r"))  asCurrentField+=oneChar;
            asCurrentField+=oneChar;
        }

    }

    //if last LF in file is missing
    if(!boNewLine)
    {
      lstOneRowFields << asCurrentField;
      asCurrentField.clear();
      //new row
      lsCSVDoc << lstOneRowFields;
      lstOneRowFields.clear();
    }


    //write to Excel document
    int iCitacFields = 0;
    int iCitacRow = 0;
    Format text_num_format;
    text_num_format.setNumberFormat("@");
    foreach (QStringList lstRow, lsCSVDoc)
    {
        iCitacRow++;
        iCitacFields = 0;
        foreach (QString asItem, lstRow)
        {
            iCitacFields++;
            xlsxDocument.write(iCitacRow, iCitacFields, asItem, text_num_format);
            //qDebug() <<  iCitacRow  << QChar(64+iCitacFields) << ":" << asItem;
        }
    }

    return 0;
}




MainWindow::MainWindow(QStringList arguments, QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);

    setCentralWidget(ui->MainFrame);

    //display version
    setWindowTitle(QFileInfo( QCoreApplication::applicationFilePath() ).completeBaseName() + " V: " + QApplication::applicationVersion());
    ui->plainTextEdit_Statistic->setReadOnly(true);




    //other forms
    detailView = new DetailView(this);



    //init.
    newReqFileLastPath = ".//";
    oldReqFileLastPath = ".//";
    asCompareCondition = "";

    //ini files
    QString asIniFileName = QFileInfo( QCoreApplication::applicationFilePath() ).filePath().section(".",0,0)+".ini";
    iniSettings = new QSettings(asIniFileName, QSettings::IniFormat);
    //in other forms:



    QVariant VarTemp;



    VarTemp = iniSettings->value("paths/newReqFileLastPath");
    if (VarTemp.isValid())
    {
      newReqFileLastPath = VarTemp.toString();
    }
    else
    {
    }

    VarTemp = iniSettings->value("paths/oldReqFileLastPath");
    if (VarTemp.isValid())
    {
      oldReqFileLastPath = VarTemp.toString();
    }
    else
    {
    }

    VarTemp = iniSettings->value("extTools/WinMerge");
    if (VarTemp.isValid())
    {
      detailView->asWinMergePath = VarTemp.toString();
    }
    else
    {
      detailView->asWinMergePath = WIN_MERGE_LOCATION;
    }


    //set query table header
      QStringList tableHeader = QStringList() <<"Sel."
                                              <<"Sheet"
                                              <<"Row"
                                              <<"Column"
                                              <<"Old Value"
                                              <<"New Value"
                                              <<"User Note"
                                              ;

      ui->tableWidget_Changes->setColumnCount(tableHeader.count());
      ui->tableWidget_Changes->setHorizontalHeaderLabels(tableHeader);



      //get host and user name
      asUserAndHostName = "";
#if defined(Q_OS_WIN)
      char acUserName[100];
      DWORD sizeofUserName = sizeof(acUserName);
      if (GetUserNameA(acUserName, &sizeofUserName))asUserAndHostName = acUserName;
#endif
#if defined(Q_OS_LINUX)
      {
         QProcess process(this);
         process.setProgram("whoami");
         process.start();
         while (process.state() != QProcess::NotRunning) qApp->processEvents();
         asUserAndHostName = process.readAll();
         asUserAndHostName = asUserAndHostName.trimmed();
      }
#endif
      asUserAndHostName += "@" + QHostInfo::localHostName();
      //qDebug() << asUserAndHostName;

      boBatchProcessing = false;

      //CLI
      //queued connection - start after leaving constructor in case of batch processing
      connect(this,SIGNAL(startBatchProcessing(int)),
              SLOT(batchProcessing(int)),
              Qt::QueuedConnection);


      //process command line arguments
      comLineArgList = arguments;
      ////remove exec name
      if(comLineArgList.size() > 0) comLineArgList.removeFirst();

      if(comLineArgList.size() >= 2)
      {
        ui->lineEdit_NewReq->setText(comLineArgList.at(0).trimmed());
        comLineArgList.removeFirst();
        ui->lineEdit_OldReq->setText(comLineArgList.at(0).trimmed());
        comLineArgList.removeFirst();



        if(comLineArgList.contains("-w")) ui->cb_IngnoreWaC->setChecked(true);
        else                              ui->cb_IngnoreWaC->setChecked(false);



        //check for batch processing (without window)
        boBatchProcessing = false;
        iExitCode = 0;
        if(comLineArgList.contains("-b"))
        {
           //do not show main window, process without it
           boBatchProcessing = true;
           startBatchProcessing(knBatchProcessingID);
           //qDebug() << "BATCH PROCESSING STARTED, SHOULD BE FIRST";

        }

     }



     EmptyChangesTable();

     //set visible "compare" page in result tab
     ui->tabWidget_Results->setCurrentWidget(ui->pageCompare);




}


int MainWindow::batchProcessing(int iID)
{
  iExitCode = 0;
  if (iID != knBatchProcessingID) iExitCode = knExitStatusBadSignal;
  if(!iExitCode)  this->on_btnCompare_clicked();
  if(!iExitCode)  this->on_btnWrite_clicked();

  QCoreApplication::exit(iExitCode);

  return(iExitCode);


}

void MainWindow::EmptyChangesTable()
{
    //delete old if exist
     for (int irow = 0; irow < ui->tableWidget_Changes->rowCount(); ++irow)
     {
       for (int icol = 0; icol < ui->tableWidget_Changes->columnCount(); ++icol)
       {
         if(ui->tableWidget_Changes->item(irow, icol) != NULL)
         {
           delete ui->tableWidget_Changes->item(irow, icol);
           ui->tableWidget_Changes->setItem(irow, icol, NULL);
         }

       }
     }
     ui->tableWidget_Changes->setRowCount(0);
     ui->btnWrite->setEnabled(false);


    //clear statistic list
    ui->plainTextEdit_Statistic->clear();


}


void MainWindow::setChangesTableNotEditable()
{
    //set table not editable, not selectable - exeptions is user comment column
     for (int irow = 0; irow < ui->tableWidget_Changes->rowCount(); ++irow)
     {
       for (int icol = 0; icol < ui->tableWidget_Changes->columnCount(); ++icol)
       {

           if (icol != col_UserNote) //user comment of course editable...
           {


             QTableWidgetItem *item = ui->tableWidget_Changes->item(irow, icol);
             if(item== nullptr)
             {
                item = new QTableWidgetItem();
                ui->tableWidget_Changes->setItem(irow,icol, item);
             }
             item->setFlags(item->flags() & ~(Qt::ItemIsEditable) & ~(Qt::ItemIsSelectable));
          }
       }
     }
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::on_btnWrite_clicked()
{

    //write result
    QXlsx::Document reportDoc;

    //add document properties
    reportDoc.setDocumentProperty("creator", asUserAndHostName);
    reportDoc.setDocumentProperty("description", "Requirements Comparsion");


    reportDoc.addSheet("TABLE");

    int iReportCurrentRow = kn_ReportFirstInfoRow;
    //info to report file
    reportDoc.write(iReportCurrentRow++, 1, "generated by: " +
                                          QFileInfo( QCoreApplication::applicationName()).fileName() +
                                          " V:" + QApplication::applicationVersion() +
                                          ", "  + asUserAndHostName);
    reportDoc.write(iReportCurrentRow++, 1, "generated on: " + QDateTime::currentDateTime().toString("dd.MM.yyyy hh:mm:ss"));
    reportDoc.write(iReportCurrentRow++, 1, "new file: " + fileName_NewReq  + ", sheet: " + asNewSheetName);
    reportDoc.write(iReportCurrentRow++, 1, "old file: " + fileName_OldReq  + ", sheet: " + asOldSheetName);
    reportDoc.write(iReportCurrentRow++, 1, asCompareCondition);
    int iRowWhereToWriteNumberOfReqWritten = iReportCurrentRow++;

// one blank line
     iReportCurrentRow++;

// headers
     Format font_bold;
     font_bold.setFontBold(true);
     reportDoc.write(iReportCurrentRow,     1, "Sheet", font_bold);       reportDoc.setColumnWidth(1, 23);
     reportDoc.write(iReportCurrentRow,     2, "Row", font_bold);             reportDoc.setColumnWidth(2, 8);
     reportDoc.write(iReportCurrentRow,     3, "Column", font_bold);            reportDoc.setColumnWidth(3, 16);
     reportDoc.write(iReportCurrentRow,     4, "Old Value", font_bold);    reportDoc.setColumnWidth(4, 25);
     reportDoc.write(iReportCurrentRow,     5, "New Value", font_bold);         reportDoc.setColumnWidth(5, 75);
     reportDoc.write(iReportCurrentRow,     6, "User Notes", font_bold);         reportDoc.setColumnWidth(6, 75);


     //new row
     iReportCurrentRow++;
     int iReqWritten = 0;

     QSet<QString> setStatisticReq;
     setStatisticReq.clear();
     //copy from table to excel sheet "TABLE"
     for (int irow = 0; irow < ui->tableWidget_Changes->rowCount(); ++irow)
     {
       //it is selected by user (check box)
       QCheckBox *checkBox = dynamic_cast<QCheckBox *> (ui->tableWidget_Changes->cellWidget(irow, col_infotable_check_box));
       if(checkBox != nullptr)
       {
           if(!(checkBox->isChecked())) continue; //not selected, next row
       }
       iReqWritten++;


       //write to excel
       for (int icol = 0; icol < ui->tableWidget_Changes->columnCount(); ++icol)
       {
         int iExcelCol = icol;
         if(iExcelCol < 1) continue;
         if(ui->tableWidget_Changes->item(irow, icol) != NULL)
         {
           Format cellFormat;
           QColor cellColor = ui->tableWidget_Changes->item(irow, icol)->background().color();
           if((cellColor.isValid()) && (cellColor != Qt::black))
           {
               //qDebug() << cellColor;
               cellFormat.setPatternBackgroundColor(cellColor);
           }


           QString asTableContent = ui->tableWidget_Changes->item(irow, icol)->text();
           reportDoc.write(iReportCurrentRow, iExcelCol, asTableContent, cellFormat);                   
         }

       }
       iReportCurrentRow++;
     }


     qDebug() << ui->plainTextEdit_Statistic->toPlainText().replace("\r","").replace("\n","::");

     QString asReqWritten = "";
     if (iReqWritten < ui->tableWidget_Changes->rowCount())
     {
         asReqWritten = QString("Only %1 table row of %2 written").arg(iReqWritten).arg(ui->tableWidget_Changes->rowCount());
     }
     else
     {
         asReqWritten = QString("All %1 table row written").arg(ui->tableWidget_Changes->rowCount());
     }

     reportDoc.write(iRowWhereToWriteNumberOfReqWritten, 1,
                     asReqWritten + "::"+ui->plainTextEdit_Statistic->toPlainText().replace("\r","").replace("\n","::"));






     //write excel sheet "EASY TO READ"
     reportDoc.addSheet("EASY TO READ");

     reportDoc.setColumnWidth(2, 23);
     reportDoc.setColumnWidth(3, 8);
     reportDoc.setColumnWidth(4, 16);
     reportDoc.setColumnWidth(5, 25);

     int iWriteToExcelRow = -1;
     int iWriteToExcelCol = -1;

     iWriteToExcelRow = 1;
     for (int irow = 0; irow < ui->tableWidget_Changes->rowCount(); ++irow)
     {

       //it is selected by user (check box)
       QCheckBox *checkBox = dynamic_cast<QCheckBox *> (ui->tableWidget_Changes->cellWidget(irow, col_infotable_check_box));
       if(checkBox != nullptr)
       {
          if(!(checkBox->isChecked())) continue; //not selected, next row
       }


       bool boExtraRow = false;
       for (int icol = 0; icol < ui->tableWidget_Changes->columnCount(); ++icol)
       {
         if(ui->tableWidget_Changes->item(irow, icol) != NULL)
         {
           Format cellFormat;
           QColor cellColor = ui->tableWidget_Changes->item(irow, icol)->background().color();
           if((cellColor.isValid()) && (cellColor != Qt::black))
           {
               //qDebug() << cellColor;
               cellFormat.setPatternBackgroundColor(cellColor);
           }


           QString asTableContent = ui->tableWidget_Changes->item(irow, icol)->text();
           asTableContent = asTableContent.trimmed();
           if(asTableContent.isEmpty() || asTableContent.isNull())
           {
               //boExtraRow = true;
               continue;
           }

           iWriteToExcelCol = icol+1;
           if(icol == col_OldValue)
           {
             iWriteToExcelRow++;
             iWriteToExcelCol = 1;
             asTableContent = "OLD: "+asTableContent;
             boExtraRow = true;
           }

           if(icol == col_NewValue)
           {
             iWriteToExcelRow++;
             iWriteToExcelCol = 1;
             asTableContent = "NEW: "+asTableContent;
             boExtraRow = true;
           }

           if(icol == col_UserNote)
           {
             iWriteToExcelRow++;
             iWriteToExcelCol = 1;
             asTableContent = "USER: "+asTableContent;
             boExtraRow = true;
           }
           reportDoc.write(iWriteToExcelRow, iWriteToExcelCol, asTableContent, cellFormat);




         }

       }
       iWriteToExcelRow++;
       if(boExtraRow) iWriteToExcelRow++;

     }

    reportDoc.selectSheet("TABLE");
    if(true)
    {
       QString fileName_Report_complet = fileName_Report + ui->lineEdit_ReportSuffix->text() + ".xlsx";
       if(!reportDoc.saveAs(fileName_Report_complet))
       {

           if (!boBatchProcessing)
           {
               QMessageBox::information(this, "Problem", "Error, not opened?", QMessageBox::Ok);
           }
           else
           {
              qCritical() << "Error: "<< "Problem write to report file (opened?)";
              iExitCode = knExitStatusReporFileOpened;
           }

       }
       else
       {

           if (!boBatchProcessing)
           {
               QMessageBox::information(this, "Written", fileName_Report_complet+"\r\n\r\n"+asReqWritten, QMessageBox::Ok);
           }
           else
           {
              qWarning() << "Success: " << asReqWritten;
           }

       }

    }
    else
    {
        if (!boBatchProcessing)
        {
            QMessageBox::information(this, "Problem", "No result doc", QMessageBox::Ok);
        }
        else
        {
           qCritical() << "Error: " << "No result doc";
           iExitCode = knExitStatusNoResultDoc;
        }

    }
}

void MainWindow::on_btnNewReq_clicked()
{
    EmptyChangesTable();
    ui->comboBox_SheetsNew->clear();
    //delete sheet names
    asNewSheetName.clear();




    // local only, it is use later from edit window
    QString fileName_NewReq = QFileDialog::getOpenFileName
    (
      this,
      "Open New Requirements",
      newReqFileLastPath,
      "Excel Files (*.xlsx;*.xlsm; *.csv)"
    );

    if(!fileName_NewReq.isEmpty() && !fileName_NewReq.isNull())
    {
        ui->lineEdit_NewReq->setText(fileName_NewReq);
        newReqFileLastPath = QFileInfo(fileName_NewReq).path();
        if (QFileInfo(fileName_NewReq).suffix().toLower().contains("xls"))
        {
            QXlsx::Document  docToListSheets(fileName_NewReq);
            foreach(QString name, docToListSheets.sheetNames())
            {
                //qDebug() << "new sheets" <<  name;
                ui->comboBox_SheetsNew->addItem(name);
            }

        }

    }
    else
    {
        ui->lineEdit_NewReq->clear();
    }


}


void MainWindow::on_lineEdit_NewReq_textEdited(const QString &arg1)
{
    Q_UNUSED(arg1)
    //new files - table is not valid for them
    EmptyChangesTable();
    ui->comboBox_SheetsNew->clear();
    asNewSheetName.clear();


}

void MainWindow::on_btnOldReq_clicked()
{
    EmptyChangesTable();
    ui->comboBox_SheetsOld->clear();
    //delete sheet names
    asOldSheetName.clear();

    // local only, it is use later from edit window
    QString fileName_OldReq = QFileDialog::getOpenFileName
    (
      this,
      "Open Old Requirements",
      oldReqFileLastPath,
      "Excel Files (*.xlsx;*.xlsm; *.csv)"
    );

    if(!fileName_OldReq.isEmpty() && !fileName_OldReq.isNull())
    {
        ui->lineEdit_OldReq->setText(fileName_OldReq);
        oldReqFileLastPath = QFileInfo(fileName_OldReq).path();
        //load sheets
        if (QFileInfo(fileName_OldReq).suffix().toLower().contains("xls"))
        {
           QXlsx::Document  docToListSheets(fileName_OldReq);
           foreach(QString name, docToListSheets.sheetNames())
            {
               //qDebug() << "old sheets" <<  name;
               ui->comboBox_SheetsOld->addItem(name);
            }

        }
    }
    else
    {
        ui->lineEdit_OldReq->clear();
    }
}

void MainWindow::on_lineEdit_OldReq_textEdited(const QString &arg1)
{
      Q_UNUSED(arg1)
    //new files - table is not valid for them
      EmptyChangesTable();
      ui->comboBox_SheetsOld->clear();
      asOldSheetName.clear();

}

void MainWindow::on_btnCompare_clicked()
{
    //set visible "compare" page in result tab
    ui->tabWidget_Results->setCurrentWidget(ui->pageCompare);

    fileName_NewReq = ui->lineEdit_NewReq->text();
    fileName_OldReq = ui->lineEdit_OldReq->text();

    //check validity of filenames
    if(
            fileName_NewReq.isEmpty() ||
            fileName_NewReq.isNull()  ||
            fileName_OldReq.isEmpty() ||
            fileName_OldReq.isNull()
        )
    {
        if (!boBatchProcessing)
        {
            QMessageBox::information(this, "No filenames", "Select Files", QMessageBox::Ok);
        }
        else
        {
           qCritical() << "Error: " << "Invalid input filenames";
           iExitCode = knExitStatusInvalidInputFilenames;
        }
        return;
    }






    //Load the new document or convert it from csv file
    qDebug() << __LINE__;
    QXlsx::Document  newReqDoc(fileName_NewReq);
    qDebug() << __LINE__;

    if(QFileInfo(fileName_NewReq).suffix().toLower() == "csv")
    {
       int iResult = CsvToExcel(fileName_NewReq, newReqDoc);
       if(iResult)
       {
           if (!boBatchProcessing)
           {
               QMessageBox::information(this, "Problem", "Problem to load new csv file", QMessageBox::Ok);
           }
           else
           {
               qCritical() << "Error: " << "Problem to load new csv files";
               iExitCode = iResult;
           }
           return;
       }
    }
    else
    {
       if (!newReqDoc.load())
       {
           if (!boBatchProcessing)
           {
               QMessageBox::information(this, "Problem", "Problem to load new file", QMessageBox::Ok);
           }
           else
           {
               qCritical() << "Error: " << "Can not load new input file";
               iExitCode = knExitStatusCannotLoadInputFiles;
           }

           return;
       }
    }


    if (ui->comboBox_SheetsNew->count() > 1)   //if only one, then it is default...
    {
        newReqDoc.selectSheet(ui->comboBox_SheetsNew->currentText());
    }
    asNewSheetName = newReqDoc.currentSheet()->sheetName();


    //newReqDoc.saveAs("pn.xlsx");  //DEBUG CODE

    //Load the old document or convert it from csv file
    QXlsx::Document  oldReqDoc(fileName_OldReq);

    if(QFileInfo(fileName_OldReq).suffix().toLower() == "csv")
    {
       int iResult = CsvToExcel(fileName_OldReq, oldReqDoc);
       if(iResult)
       {
           if (!boBatchProcessing)
           {
               QMessageBox::information(this, "Problem", "Problem to load old csv file", QMessageBox::Ok);
           }
           else
           {
               qCritical() << "Error: " << "Problem to load old csv files";
               iExitCode = iResult;
           }
           return;
       }
    }
    else
    {
       if (!oldReqDoc.load())
       {
           if (!boBatchProcessing)
           {
               QMessageBox::information(this, "Problem", "Problem to load old file", QMessageBox::Ok);
           }
           else
           {
               qCritical() << "Error: " << "Can not load old input files";
               iExitCode = knExitStatusCannotLoadInputFiles;
           }

           return;
       }
    }

    //oldReqDoc.saveAs("po.xlsx");  //DEBUG CODE



    if (ui->comboBox_SheetsOld->count() > 1)   //if only one, then it is default...
    {
        oldReqDoc.selectSheet(ui->comboBox_SheetsOld->currentText());
    }
    asOldSheetName = oldReqDoc.currentSheet()->sheetName();


    //save paths to ini file
    iniSettings->setValue("paths/newReqFileLastPath", newReqFileLastPath);
    iniSettings->setValue("paths/oldReqFileLastPath", oldReqFileLastPath);


    //prepare report file
    fileName_Report =   QFileInfo(fileName_NewReq).path() + "/" +
                        QFileInfo(fileName_NewReq).completeBaseName() +
                        "_to_" +
                        QFileInfo(fileName_OldReq).completeBaseName();



    fileName_OldReqCor = QFileInfo(fileName_OldReq).path() + "/" +
                         QFileInfo(fileName_OldReq).completeBaseName()+ 
                         "_cor"
                         ".xlsx";                         

    //qDebug() << fileName_NewReq << fileName_OldReq << fileName_Report;









    //Delete changes table
    EmptyChangesTable();

    //disable sorting
    ui->tableWidget_Changes->setSortingEnabled(false);
    //close detail
    detailView->close();


    //start compare demo
    boCompareJustSelectedSheet = !ui->cb_CompAllSheets->isChecked();

    foreach(QString comparedSheetName, newReqDoc.sheetNames())
    {

        bool boJustOnePass =      false;
        bool boNewSheetSelected = false;
        bool boOldSheetSelected = false;

        if(!boCompareJustSelectedSheet)
        {
            boNewSheetSelected = newReqDoc.selectSheet(comparedSheetName);
            boOldSheetSelected = oldReqDoc.selectSheet(comparedSheetName);

            if(!boNewSheetSelected)
            {
                qDebug() << "No Sheet" << comparedSheetName << " in NEW";
                continue;
            }

            if(!boOldSheetSelected)
            {
                qDebug() << "No Sheet" << comparedSheetName << " in OLD";
                continue;

            }
        }
        else
        {
          boJustOnePass = true;
          boNewSheetSelected = true; //already selected
          boOldSheetSelected = true; //already selected
          comparedSheetName = newReqDoc.currentSheet()->sheetName() + "/" + oldReqDoc.currentSheet()->sheetName();
        }

        int newLastRow = newReqDoc.dimension().lastRow();
        int newLastCol = newReqDoc.dimension().lastColumn();
        int oldLastRow = oldReqDoc.dimension().lastRow();
        int oldLastCol = oldReqDoc.dimension().lastColumn();

        int maxLastRow = (newLastRow > oldLastRow) ? newLastRow : oldLastRow;
        int maxLastCol = (newLastCol > oldLastCol) ? newLastCol : oldLastCol;
        qDebug() << boNewSheetSelected << boOldSheetSelected << comparedSheetName <<newLastRow << newLastCol << oldLastRow << oldLastCol << maxLastRow << maxLastCol;


        for(int row =1; row <= maxLastRow; row++)
        {
            for(int col=1; col <= maxLastCol; col++)
            {
                QString asNewValue = getCellValueValue(row, col, newReqDoc);
                QString asOldValue = getCellValueValue(row, col, oldReqDoc);

                if(ui->cb_IngnoreWaC->isChecked())
                {
                  asNewValue = ignoreWhiteAndCase(asNewValue);
                  asOldValue = ignoreWhiteAndCase(asOldValue);
                }

                if(  asNewValue != asOldValue)
                {
                    qDebug() << "SHEET:" << comparedSheetName << "NEW:" << asNewValue << "OLD:" << asOldValue;
                    //insert to table
                    ui->tableWidget_Changes->insertRow(ui->tableWidget_Changes->rowCount()); //new row
                    ui->tableWidget_Changes->setItem(ui->tableWidget_Changes->rowCount()-1, col_Sheet,    new QTableWidgetItem(comparedSheetName));
                    ui->tableWidget_Changes->setItem(ui->tableWidget_Changes->rowCount()-1, col_Row,      new QTableWidgetItem(QString::number(row)));
                    ui->tableWidget_Changes->setItem(ui->tableWidget_Changes->rowCount()-1, col_Column,   new QTableWidgetItem(QString::number(col)));
                    ui->tableWidget_Changes->setItem(ui->tableWidget_Changes->rowCount()-1, col_OldValue, new QTableWidgetItem(asOldValue));
                    ui->tableWidget_Changes->setItem(ui->tableWidget_Changes->rowCount()-1, col_NewValue, new QTableWidgetItem(asNewValue));
                    ui->tableWidget_Changes->setItem(ui->tableWidget_Changes->rowCount()-1, col_UserNote, new QTableWidgetItem(""));

                    //multiline for new and old text?
                    QTextEdit *edit;
                    edit = new QTextEdit();
                    edit->setText(ui->tableWidget_Changes->item(ui->tableWidget_Changes->rowCount()-1, col_OldValue)->text());
                    ui->tableWidget_Changes->setCellWidget     (ui->tableWidget_Changes->rowCount()-1, col_OldValue, edit);

                    edit = new QTextEdit();
                    edit->setText(ui->tableWidget_Changes->item(ui->tableWidget_Changes->rowCount()-1, col_NewValue)->text());
                    ui->tableWidget_Changes->setCellWidget     (ui->tableWidget_Changes->rowCount()-1, col_NewValue, edit);

                }

            }
        }

       if(boJustOnePass) break;
    }


    //Add checkboxes to the table
     for (int irow = 0; irow < ui->tableWidget_Changes->rowCount(); ++irow)
     {
         QCheckBox *cb = new QCheckBox(this);
         ui->tableWidget_Changes->setCellWidget(irow, col_infotable_check_box, cb);
     }

     //set checkboxes to checked state
     checkAll(true);







     //allow export
     ui->btnWrite->setEnabled(true);






}

void MainWindow::on_tableWidget_Changes_clicked(const QModelIndex &index)
{
    //qDebug() << index.row() << index.column();


    QString asOldText = "";
    QString asNewText = "";


    //texts
    if(ui->tableWidget_Changes->item(index.row(), col_OldValue) != nullptr)
      asOldText = ui->tableWidget_Changes->item(index.row(), col_OldValue)->text();
    if( ui->tableWidget_Changes->item(index.row(), col_NewValue) != nullptr)
      asNewText = ui->tableWidget_Changes->item(index.row(), col_NewValue)->text();
    detailView->setTexts(asOldText, asNewText);



    detailView->setTitle("Differences");

    if(index.column() == col_infotable_check_box) return; //select field
    if(index.column() == col_UserNote) return; //user notes field

    detailView->show();


}

void MainWindow::on_btn_Debug_clicked()
{
    //if(detailView != nullptr) detailView->exec();
    detailView->show();
    detailView->setTexts("My Old", "My New");
}





void MainWindow::on_btn_CheckAll_clicked()
{
  bool static boNowCheckAll = true;
  if(boNowCheckAll) checkAll(true);
  else              checkAll(false);
  boNowCheckAll = !boNowCheckAll;
}


void MainWindow::checkAll(bool boCheck)
{
    for (int irow = 0; irow < ui->tableWidget_Changes->rowCount(); ++irow)
    {
        QCheckBox *checkBox = dynamic_cast<QCheckBox *> (ui->tableWidget_Changes->cellWidget(irow, col_infotable_check_box));
        if(checkBox != nullptr)  checkBox->setChecked(boCheck);
    }

}

void MainWindow::on_btn_SortReq_clicked()
{
   bool static boSortAsc = true;
   if(boSortAsc) {}
   else          {}


   boSortAsc = !boSortAsc;
}


