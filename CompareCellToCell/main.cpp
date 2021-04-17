#include "mainwindow.h"
#include <QApplication>
#include <QDebug>
#include <QStyleFactory>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    //set version
    a.setApplicationVersion(APP_VERSION);
    //set style
    a.setStyle(QStyleFactory::create("Fusion"));

    MainWindow w(a.arguments());
    if(!w.boBatchProcessing)w.show();

    return a.exec();
}
