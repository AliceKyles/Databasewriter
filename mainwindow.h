#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QtSql/QSqlDatabase>
#include <QSqlRecord>
#include <QSqlQuery>
#include <QAxObject>

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

private slots:

    void addchan();
    void del();
    void logo();
    void search();
    void excel();
    void image();
    void lang(int state);
    void clear();

private:
    Ui::MainWindow *ui;
    QString loga, imaga;
    QSqlDatabase db;
    bool o, p;
    int ao, ap, uo, up;
    QString select_image();
    void openSQL();
    bool checko();
    bool checkp();
    void updatelogo(int id);
    void updateimage(int id);
    int ido();
    void updateproj(int id);
    void updatepo(QString s, QString s1, QString s2, int id);
    void updateint(QString s, QString s1, QString s2, int id);
    void fillo(QString s, int id);
    void fillp(QString s, int id);
    QString select(QString s1, QString s2, int id);
    QString getcell(int c, int r, QAxObject* sheet);
};

#endif // MAINWINDOW_H
