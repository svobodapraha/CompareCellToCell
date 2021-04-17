#include "detailview.h"
#include "ui_detailview.h"
#include <QClipboard>

DetailView::DetailView(QWidget *parent) :
    QDialog(parent),
    ui(new Ui::DetailView),
    newFileForCompare("New_XXXXXXX.tmp"),
    oldFileForCompare("Old_XXXXXXX.tmp")
{
    ui->setupUi(this);

    ui->MainFrame->setSizePolicy(QSizePolicy::Expanding, QSizePolicy::Expanding);
    QVBoxLayout* vbox = new QVBoxLayout( this );
    vbox->addWidget( ui->MainFrame );

    //qDebug() << "detailedview form";


}

DetailView::~DetailView()
{
    delete ui;
}







void DetailView::setTexts(QString oldText, QString newText)
{
    ui->textBrowser_oldReq->setText(oldText);
    ui->textBrowser_newReq->setText(newText);

    //prepare temporary files for external compare
    newFileForCompare.close();
    oldFileForCompare.close();
    newFileForCompare.remove();
    oldFileForCompare.remove();
    newFileForCompare.open();
    oldFileForCompare.open();
    QTextStream newStreamForCompare(&newFileForCompare);
    QTextStream oldStreamForCompare(&oldFileForCompare);
    newStreamForCompare << newText;
    newStreamForCompare.flush();
    oldStreamForCompare << oldText;
    oldStreamForCompare.flush();


}

void DetailView::setTitle(QString asTitle)
{
    this->setWindowTitle(asTitle);
}

void DetailView::on_btnDiff_clicked()
{
#if defined(Q_OS_WIN)
    QStringList lstArg;
    lstArg << "/dl" << this->windowTitle() << "/dr" << this->windowTitle()
           << "/s" << "/e" << "/u" << "/wl" << "/wm" << "/wr"
           << oldFileForCompare.fileName()
           << newFileForCompare.fileName();
    qDebug() << lstArg;
    QProcess::startDetached(asWinMergePath, lstArg);

#endif
}

void DetailView::on_btnCopyToClipboard_clicked()
{
    QString asTextForClipboard = "";
    asTextForClipboard = "OLD:\r****\r" +
                         ui->textBrowser_oldReq->toPlainText() +
                         "\rNEW:\r****\r" +
                         ui->textBrowser_newReq->toPlainText();

    QClipboard *p_Clipboard = QApplication::clipboard();
    p_Clipboard->setText(asTextForClipboard);
}

void DetailView::on_btnClose_clicked()
{
    this->close();
}


