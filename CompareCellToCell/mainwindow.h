#ifndef MAINWINDOW_H
#define MAINWINDOW_H


#include <QMainWindow>
#include <QDebug>
#include <QFileDialog>
#include <QMessageBox>
#include <QSettings>
#include <QHostInfo>
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QSqlTableModel>
#include <QSqlQueryModel>
#include <QSqlError>
#include <QListWidgetItem>


#include "xlsxdocument.h"
#include "xlsxchartsheet.h"
#include "xlsxcellrange.h"
#include "xlsxchart.h"
#include "xlsxrichstring.h"
#include "xlsxworkbook.h"

#include "detailview.h"
#include <windows.h>


using namespace QXlsx;

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QStringList arguments,QWidget *parent = 0);
    ~MainWindow();

    bool boBatchProcessing;
    int iExitCode;

signals:
    void startBatchProcessing(int iID);


private slots:
    void on_btnWrite_clicked();

    void on_btnNewReq_clicked();

    void on_btnOldReq_clicked();

    void on_btnCompare_clicked();

    void on_tableWidget_Changes_clicked(const QModelIndex &index);

    void on_btn_Debug_clicked();

    void on_lineEdit_NewReq_textEdited(const QString &arg1);

    void on_lineEdit_OldReq_textEdited(const QString &arg1);

    void on_btn_CheckAll_clicked();

    void on_btn_SortReq_clicked();

    int batchProcessing(int iID);





private:
    Ui::MainWindow *ui;

private:
    QString fileName_NewReq, asNewSheetName;
    QString fileName_OldReq, asOldSheetName;
    QString fileName_OldReqCor;
    QString newReqFileLastPath;
    QString oldReqFileLastPath;
    QString fileName_Report;
    QString asCompareCondition;
    QString asQueryTextToReport;


    QSettings *iniSettings;
    void EmptyChangesTable();
    DetailView *detailView;

    void setChangesTableNotEditable();
    void checkAll(bool boCheck);

    QString asUserAndHostName;
    QStringList comLineArgList;

    bool boCompareJustSelectedSheet;



};


#define kn_HeaderRow      1
#define kn_FistDataRow    2
#define kn_ReqIDCol       1
#define kn_FirstHeaderCol 2

#define col_infotable_check_box         0
#define col_Sheet                       1
#define col_Row                         2
#define col_Column                      3
#define col_OldValue                    4
#define col_NewValue                    5
#define col_UserNote                    6


#define kn_ReportFirstInfoRow           1

#define knBatchProcessingID  1

#define knExitStatusBadSignal               1
#define knExitStatusReporFileOpened         2
#define knExitStatusInvalidInputFilenames   3
#define knExitStatusCannotLoadInputFiles    4
#define knExitStatusNoResultDoc             5
#define knExitStatusCorrectedFileOpened     6
#define knExitStatusCannotLoadInputCsvFiles 7

#define knOrigIDText     ("Original Old ID")
#define knMAX_VIEW_COLUMN_WIDTH            400
#define WIN_MERGE_LOCATION   "c:\\Program Files (x86)\\WinMerge\\WinMergeU.exe"



#endif // MAINWINDOW_H
