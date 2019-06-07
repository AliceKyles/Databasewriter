#include "mainwindow.h"
#include <QApplication>
#include <QTextCodec>
#include "mainwindow.h"
int main(int argc, char *argv[])
{
    QTextCodec::setCodecForLocale(QTextCodec::codecForName("UTF-8"));
    QApplication a(argc, argv);
    MainWindow w;
    w.setWindowTitle("DatabaseWriter");
    a.setWindowIcon(QIcon("icon.png"));
    w.show();

    return a.exec();
}
