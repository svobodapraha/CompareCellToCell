#ifndef DETAILVIEW_H
#define DETAILVIEW_H

#include <QDialog>
#include <QDebug>
#include <QSaveFile>
#include <QProcess>
#include <QTemporaryFile>

namespace Ui {
class DetailView;
}

class DetailView : public QDialog
{
    Q_OBJECT

public:
    explicit DetailView(QWidget *parent = 0);
    ~DetailView();

    void setTexts(QString oldText, QString newText);
    void setTitle(QString asTitle);
    QString asWinMergePath;

private slots:


    void on_btnCopyToClipboard_clicked();

    void on_btnClose_clicked();

    void on_btnDiff_clicked();

private:
    Ui::DetailView *ui;
    //ext compare stuff
     QTemporaryFile newFileForCompare;
     QTemporaryFile oldFileForCompare;

};

#endif // DETAILVIEW_H
