#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QFileDialog>
#include <QFile>




MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    connect(ui->excel, SIGNAL(clicked()), this, SLOT(excel()));
    connect(ui->del, SIGNAL(clicked()), this, SLOT(del()));
    connect(ui->addchan, SIGNAL(clicked()), this, SLOT(addchan()));
    connect(ui->search, SIGNAL(clicked()), this, SLOT(search()));
    connect(ui->logo, SIGNAL(clicked()), this, SLOT(logo()));
    connect(ui->imag, SIGNAL(clicked()), this, SLOT(image()));
    connect(ui->lang_2, SIGNAL(stateChanged ( int )), this, SLOT(lang(int)));
    connect(ui->clear, SIGNAL(clicked()), this, SLOT(clear()));
    openSQL();
}

void MainWindow::lang(int state) //позволяет вводить данные на другом языке после отметки соответствующего поля
{
    ui->olang->setEnabled(state);
    ui->plang->setEnabled(state);
}

void MainWindow::addchan() //добавляет или меняет данные в базе данных
{
    QSqlQuery query(db);
    if (ui->oname->text().size()>0){
        if (checko()){
            int id=ido();
            updatelogo(id);
            updatepo(ui->cont->text(), "contact","Organization", id);
            updatepo(ui->odesc->toPlainText(),"desc","Organization", id);
            if (o==1){
                ui->info->append("Информация об организации обновилась");
                o=0;
            }
        }
        else{
            int id=0;
            if (ui->lang_2->isChecked()){
                if (ui->lang->currentText()=="Английский")
                   query.prepare("SELECT ID FROM Organizationru WHERE name = '"+ui->olang->text()+"'");
                else query.prepare("SELECT ID FROM Organizationen WHERE name = '"+ui->olang->text()+"'");
                query.exec();
                if (query.next()){
                    id=query.value(0).toInt();
                }
                if (ui->lang->currentText()=="Английский")
                   query.prepare("SELECT ID FROM Organizationru WHERE ID = "+ui->olang->text());
                else query.prepare("SELECT ID FROM Organizationen WHERE ID = "+ui->olang->text());
                query.exec();
                query.exec();
                if (query.next()){
                    id=query.value(0).toInt();
                }
            }
            if (id!=0)
                updatelogo(id);
            else {
                id =1;
                while (true){
                    query.prepare("SELECT * FROM Organizationes WHERE ID =:id");
                    query.bindValue(":id", id);
                    query.exec();
                    if (!query.next())
                        break;
                    id++;
                }
                query.prepare("INSERT INTO Organizationes VALUES (:id, :image)");
                query.bindValue(":id", id);
                QByteArray data;
                if (ui->checklogo->isChecked()){
                    QFile file(loga);
                    file.open(QIODevice::ReadOnly);
                    data = file.readAll();
                    file.close();
                    loga.clear();
                    ui->checklogo->setChecked(0);
                }
                query.bindValue(":image", data);
                query.exec();
            }
            if (ui->lang->currentText()=="Английский")
               query.prepare("INSERT INTO Organizationen VALUES (:id, :name, :contact, :desc)");
            else query.prepare("INSERT INTO Organizationru VALUES (:id, :name, :contact, :desc)");
            query.bindValue(":id", id);
            query.bindValue(":name", ui->oname->text());
            query.bindValue(":contact", ui->cont->text());
            query.bindValue(":desc", ui->odesc->toPlainText());
            query.exec();
            ao++;
            ui->info->append("Добавлена организация");
        }
    }
    if (ui->pname->text().size()>0){
        if (checkp()){
            int id;
            if (ui->lang->currentText()=="Английский")
               query.prepare("SELECT ID FROM Projecten WHERE name = '"+ui->pname->text()+"'");
            else query.prepare("SELECT ID FROM Projectru WHERE name = '"+ui->pname->text()+"'");
            query.exec();
            if (query.next()){
                id=query.value(0).toInt();
            }
            else id=ui->pname->text().toInt();
            updateproj(id);
            updatepo(ui->problem->toPlainText(),"problem", "Project",id);
            updatepo(ui->solution->toPlainText(),"solution","Project",id);
            updatepo(ui->actual->toPlainText(),"actuality","Project",id);
            updatepo(ui->examp->text(),"examples","Project",id);
            updatepo(ui->prot->text(),"protection","Project",id);
            updatepo(ui->stage->text(),"stage","Project",id);
            updatepo(ui->author->text(),"author","Project",id);
            updatepo(ui->categ->text(),"category","Project",id);
            if (p==1){
                ui->info->append("Информация о проекте обновилась");
                p=0;
            }
        }
        else {
            int id=0;
            if (ui->lang_2->isChecked()){
                if (ui->lang->currentText()=="Английский")
                   query.prepare("SELECT ID FROM Projectru WHERE name = '"+ui->plang->text()+"'");
                else query.prepare("SELECT ID FROM Projecten WHERE name = '"+ui->plang->text()+"'");
                query.exec();
                if (query.next()){
                    id=query.value(0).toInt();
                }
                if (ui->lang->currentText()=="Английский")
                   query.prepare("SELECT ID FROM Projectru WHERE ID = "+ui->plang->text());
                else query.prepare("SELECT ID FROM Projecten WHERE ID = "+ui->plang->text());
                query.exec();
                query.exec();
                if (query.next()){
                    id=query.value(0).toInt();
                }
                if (ui->lang->currentText()=="Английский")
                   query.prepare("SELECT OID FROM Projectru WHERE ID =:1");
                else query.prepare("SELECT OID FROM Projecten WHERE ID =:1");
                query.bindValue(":1", id);
                query.exec();
                if (query.next()){
                    ui->oname->setText(query.value(0).toString());
                }
            }
            if (id!=0){
                updateproj(id);
            }
            else {
                id =1;
                while (true){
                    query.prepare("SELECT * FROM Projectes WHERE ID =:id");
                    query.bindValue(":id", id);
                    query.exec();
                    if (!query.next())
                        break;
                    id++;
                }
                query.prepare("INSERT INTO Projectes VALUES (:id, :oid, :image, :invest, :datedev, :dateinv)");
                query.bindValue(":id", id);
                query.bindValue(":oid", ido());
                QByteArray data;
                if (ui->checkimage->isChecked()){
                    QFile file(imaga);
                    file.open(QIODevice::ReadOnly);
                    data = file.readAll();
                    file.close();
                    imaga.clear();
                    ui->checkimage->setChecked(0);
                }
                query.bindValue(":image", data);
                if (ui->invest->text().toInt()>0)
                    query.bindValue(":invest", ui->invest->text().toInt());
                QDate d(2000,1,1);
                if (ui->datedev->date()!=d)
                    query.bindValue(":datedev", ui->datedev->date());
                if (ui->dateinv->date()!=d)
                    query.bindValue(":dateinv", ui->dateinv->date());
                query.exec();
            }
            if (ui->lang->currentText()=="Английский")
               query.prepare("INSERT INTO Projecten VALUES (:id, :name, :problem, :solution, :actuality, :examples, :protection, :stage, :author, :category)");
            else query.prepare("INSERT INTO Projectru VALUES (:id, :name, :problem, :solution, :actuality, :examples, :protection, :stage, :author, :category)");
            query.bindValue(":id", id);
            query.bindValue(":name", ui->pname->text());
            query.bindValue(":problem", ui->problem->toPlainText());
            query.bindValue(":solution", ui->solution->toPlainText());
            query.bindValue(":actuality", ui->actual->toPlainText());
            query.bindValue(":examples", ui->examp->text());
            query.bindValue(":protection", ui->prot->text());
            query.bindValue(":stage", ui->stage->text());
            query.bindValue(":author", ui->author->text());
            query.bindValue(":category", ui->categ->text());
            query.exec();
            ap++;
            ui->info->append("Добавлен проект");
        }
    }
}

void MainWindow::updateproj(int id) //обновляет не зависимые от языка данные о проекте
{
    QSqlQuery query(db);
    updateimage(id);
    if (ui->oname->text().size()>0){
        query.prepare("UPDATE Projectes SET OID=:1 WHERE ID=:2");
        query.bindValue(":1", ido());
        query.bindValue(":2", id);
        query.exec();
        if (!p){
            p=1;
            up++;
        }
    }
    if (ui->invest->text().toInt()>0)
        updateint(ui->invest->text(),"invest","Projectes", id);
    QDate d(2000,1,1);
    if (ui->datedev->date()!=d){
        query.prepare("UPDATE Projectes SET datedev=:1 WHERE ID=:2");
        query.bindValue(":1", ui->datedev->date());
        query.bindValue(":2", id);
        query.exec();
        if (!p){
            p=1;
            up++;
        }
    }
    if (ui->dateinv->date()!=d){
        query.prepare("UPDATE Projectes SET dateinv=:1 WHERE ID=:2");
        query.bindValue(":1", ui->dateinv->date());
        query.bindValue(":2", id);
        query.exec();
        if (!p){
            p=1;
            up++;
        }
    }
    }

void MainWindow::updateint(QString s, QString s1,QString s2, int id) //обновляет данные типа Int
{
    QSqlQuery query(db);
    if (s.toInt()>0){
        query.prepare("UPDATE "+s2+" SET "+s1+"=:1 WHERE ID=:2");
        query.bindValue(":1", s.toInt());
        query.bindValue(":2", id);
        query.exec();
        if (!p){
            p=1;
            up++;
        }
    }
    }


void MainWindow::updatepo(QString s, QString s1, QString s2, int id) //обновляет зависимые от языка данные
{
    QSqlQuery query(db);
    if (s.size()>0){
        if (ui->lang->currentText()=="Английский")
           query.prepare("UPDATE "+s2+"en SET "+s1+"='"+s+"' WHERE ID=:1");
        else query.prepare("UPDATE "+s2+"ru SET "+s1+"='"+s+"' WHERE ID=:1");
        query.bindValue(":1", id);
        query.exec();
        if (!o&&s2=="Organization"){
            o=1;
            uo++;
        }
        if (!p&&s2=="Project"){
            p=1;
            up++;
        }
    }
    }

int MainWindow::ido() //возвращает Id организации
{
    QSqlQuery query(db);
    int id;
    if (ui->lang->currentText()=="Английский")
       query.prepare("SELECT ID FROM Organizationen WHERE name = '"+ui->oname->text()+"'");
    else query.prepare("SELECT ID FROM Organizationru WHERE name = '"+ui->oname->text()+"'");
    query.exec();
    if (query.next())
        id = query.value(0).toInt();
    else id=ui->oname->text().toInt();
    return id;
    }


void MainWindow::updateimage(int id) //обновляет изображение проекта
{
    QSqlQuery query(db);
    if (ui->checkimage->isChecked()){
        QFile file(imaga);
        file.open(QIODevice::ReadOnly);
        QByteArray data;
        data = file.readAll();
        file.close();
        query.prepare("UPDATE Projectes SET picture=:1 WHERE ID=:2");
        query.bindValue(":1", data);
        query.bindValue(":2", id);
        query.exec();
        imaga.clear();
        ui->checkimage->setChecked(0);
        if (!p){
            p=1;
            up++;
        }
    }
}

void MainWindow::updatelogo(int id) //обновляет логотип организации
{
    QSqlQuery query(db);
    if (ui->checklogo->isChecked()){
        QFile file(loga);
        file.open(QIODevice::ReadOnly);
        QByteArray data;
        data = file.readAll();
        file.close();
        query.prepare("UPDATE Organizationes SET picture=:1 WHERE ID=:2");
        query.bindValue(":1", data);
        query.bindValue(":2", id);
        query.exec();
        loga.clear();
        ui->checklogo->setChecked(0);
        if (!o){
            o=1;
            uo++;
        }
    }
}

bool MainWindow::checko() //проверяет наличие организации в базе данных
{
    QSqlQuery query(db);
    if (ui->lang->currentText()=="Английский")
       query.prepare("SELECT ID FROM Organizationen WHERE name = '"+ui->oname->text()+"'");
    else query.prepare("SELECT ID FROM Organizationru WHERE name = '"+ui->oname->text()+"'");
    query.exec();
    if (query.next()){
        return 1;
    }
    if (ui->lang->currentText()=="Английский")
       query.prepare("SELECT ID FROM Organizationen WHERE ID = "+ui->oname->text());
    else query.prepare("SELECT ID FROM Organizationru WHERE ID = "+ui->oname->text());
    query.exec();
    if (query.next()){
        return 1;
    }
    return 0;
}

bool MainWindow::checkp() //проверяет наличие проекта в базе данных
{
    QSqlQuery query(db);
    if (ui->lang->currentText()=="Английский")
       query.prepare("SELECT ID FROM Projecten WHERE name = '"+ui->pname->text()+"'");
    else query.prepare("SELECT ID FROM Projectru WHERE name = '"+ui->pname->text()+"'");
    query.exec();
    if (query.next()){
        return 1;
    }
    if (ui->lang->currentText()=="Английский")
       query.prepare("SELECT ID FROM Projecten WHERE ID = "+ui->pname->text());
    else query.prepare("SELECT ID FROM Projectru WHERE ID = "+ui->pname->text());
    query.exec();
    if (query.next()){
        return 1;
    }
    return 0;
}

void MainWindow::del() //удаляет данные из базы данных
{
    QSqlQuery query(db);
    int id;
    if (checkp()){
        if (ui->lang->currentText()=="Английский")
           query.prepare("SELECT ID FROM Projecten WHERE name = '"+ui->pname->text()+"'");
        else query.prepare("SELECT ID FROM Projectru WHERE name = '"+ui->pname->text()+"'");
        query.exec();
        if (query.next())
            id = query.value(0).toInt();
        else id=ui->pname->text().toInt();
        if (ui->lang->currentText()=="Английский")
           query.prepare("DELETE FROM Projecten WHERE ID =:1");
        else query.prepare("DELETE FROM Projectru WHERE ID =:1");
        query.bindValue(":1", id);
        query.exec();
        if (ui->lang->currentText()=="Английский")
           query.prepare("SELECT id FROM Projectru WHERE ID =:1");
        else query.prepare("SELECT id FROM Projecten WHERE ID =:1");
        query.bindValue(":1", id);
        query.exec();
        if (query.next()){
            if (ui->lang_2->isChecked()){
                if (ui->lang->currentText()=="Английский")
                   query.prepare("DELETE FROM Projectru WHERE ID =:1");
                else query.prepare("DELETE FROM Projecten WHERE ID =:1");
                query.bindValue(":1", id);
                query.exec();
                query.prepare("DELETE FROM Projectes WHERE ID =:1");
                query.bindValue(":1", id);
                query.exec();
                ui->info->setText("Удалён проект");
            }
        }
        else {
            query.prepare("DELETE FROM Projectes WHERE ID =:1");
            query.bindValue(":1", id);
            query.exec();
            ui->info->setText("Удалён проект");
        }
    }
    if (ui->pname->text().size()==0){
        if (checko()){
            id=ido();
            if (ui->lang->currentText()=="Английский")
               query.prepare("DELETE FROM Organizationen WHERE ID =:1");
            else query.prepare("DELETE FROM Organizationru WHERE ID =:1");
            query.bindValue(":1", id);
            query.exec();
            if (ui->lang->currentText()=="Английский")
               query.prepare("SELECT id FROM Organizationru WHERE ID =:1");
            else query.prepare("SELECT id FROM Organizationen WHERE ID =:1");
            query.bindValue(":1", id);
            query.exec();
            if (query.next()){
                if (ui->lang_2->isChecked()){
                    if (ui->lang->currentText()=="Английский")
                       query.prepare("DELETE FROM Organizationru WHERE ID =:1");
                    else query.prepare("DELETE FROM Organizationen WHERE ID =:1");
                    query.bindValue(":1", id);
                    query.exec();
                    query.prepare("DELETE FROM Organizationes WHERE ID =:1");
                    query.bindValue(":1", id);
                    query.exec();
                    ui->info->setText("Удалена организация");
                }
            }
            else {
                query.prepare("DELETE FROM Organizationes WHERE ID =:1");
                query.bindValue(":1", id);
                query.exec();
                ui->info->setText("Удалена организация");
            }
        }
    }
}

void MainWindow::search() //находит данные в базе данных и вводит их в соответствующие поля
{
    QSqlQuery query(db);
    int id, oid;
    id=oid=0;
    if (checko())
        oid=ido();
    if (checkp()){
        if (ui->lang->currentText()=="Английский")
           query.prepare("SELECT ID FROM Projecten WHERE name = '"+ui->pname->text()+"'");
        else query.prepare("SELECT ID FROM Projectru WHERE name = '"+ui->pname->text()+"'");
        query.exec();
        if (query.next())
            id = query.value(0).toInt();
        else id=ui->pname->text().toInt();
    }
    if (oid!=0){
        if (ui->lang->currentText()=="Английский")
           fillo("en",oid);
        else fillo("ru",oid);
    }
    if (id!=0){
        if (ui->lang->currentText()=="Английский")
           fillp("en",id);
        else fillp("ru",id);
    }
    if ((id==0)&&(oid==0)){
        ui->info->setText("Поиск не удался");
    }
}

void MainWindow::fillp(QString s, int id) //вводит в соответствующие поля данные о проекте
{
    QSqlQuery query(db);
    query.prepare("SELECT OID FROM Projectes WHERE ID =:1");
    query.bindValue(":1", id);
    query.exec();
    query.next();
    fillo(s, query.value(0).toInt());
    query.prepare("SELECT invest FROM Projectes WHERE ID =:1");
    query.bindValue(":1", id);
    query.exec();
    query.next();
    ui->invest->setText(query.value(0).toString());
    query.prepare("SELECT datedev FROM Projectes WHERE ID =:1");
    query.bindValue(":1", id);
    query.exec();
    query.next();
    ui->datedev->setDate(query.value(0).toDate());
    query.prepare("SELECT dateinv FROM Projectes WHERE ID =:1");
    query.bindValue(":1", id);
    query.exec();
    query.next();
    ui->dateinv->setDate(query.value(0).toDate());
    ui->pname->setText(select(s,"name",id));
    ui->problem->setText(select(s,"problem",id));
    ui->solution->setText(select(s,"solution",id));
    ui->actual->setText(select(s,"actuality",id));
    ui->examp->setText(select(s,"examples",id));
    ui->prot->setText(select(s,"protection",id));
    ui->stage->setText(select(s,"stage",id));
    ui->author->setText(select(s,"author",id));
    ui->categ->setText(select(s,"category",id));
}

QString MainWindow::select(QString s1, QString s2, int id) //возвращает параметр проекта типа Text
{
    QSqlQuery query(db);
    query.prepare("SELECT "+s2+" FROM Project"+s1+" WHERE ID =:1");
    query.bindValue(":1", id);
    query.exec();
    query.next();
    return query.value(0).toString();
}

void MainWindow::fillo(QString s, int id) //вводит в соответствующие поля данные об организации
{
    QSqlQuery query(db);
    query.prepare("SELECT name FROM Organization"+s+" WHERE ID =:1");
    query.bindValue(":1", id);
    query.exec();
    query.next();
    ui->oname->setText(query.value(0).toString());
    query.prepare("SELECT contact FROM Organization"+s+" WHERE ID =:1");
    query.bindValue(":1", id);
    query.exec();
    query.next();
    ui->cont->setText(query.value(0).toString());
    query.prepare("SELECT desc FROM Organization"+s+" WHERE ID =:1");
    query.bindValue(":1", id);
    query.exec();
    query.next();
    ui->odesc->setText(query.value(0).toString());
}

void MainWindow::openSQL() // находит базу данных и открывает её
{
    db = QSqlDatabase::addDatabase("QSQLITE");
    db.setDatabaseName("proj.db");
    db.open();
    uo=up=ao=ap=0;
    o=p=0;
}



void MainWindow::logo() // выбирает логотип организации
{
    loga = select_image();
    if (loga.isNull()==0)
        ui->checklogo->setChecked(1);
}

void MainWindow::image() // выбирает изображение проекта
{
    imaga = select_image();
    if (imaga.isNull()==0)
        ui->checkimage->setChecked(1);
}

void MainWindow::excel() // берёт данные из файла Excel
{
    QFileDialog qf;
    QString eadres=(qf.getOpenFileName(0, "Open Dialog", "", "*.xlsx"));
    if (eadres.isNull())
        return void();
    QAxObject* excel = new QAxObject("Excel.Application", 0);
    QAxObject* workbooks = excel->querySubObject("Workbooks");
    QAxObject* workbook = workbooks->querySubObject("Open(const QString&)", eadres);
    QAxObject* sheets = workbook->querySubObject("Worksheets");
    QAxObject* sheet = sheets->querySubObject("Item(int)", 1);
    QAxObject* usedRange = sheet->querySubObject("UsedRange");
    QAxObject* rows = usedRange->querySubObject("Rows");
    eadres.chop(eadres.size()-eadres.lastIndexOf("/")-1);
    int countRows = rows->dynamicCall("Count").toInt();
    for ( int row = 2; row <= countRows; row++ ){
        ui->oname->setText(getcell(1,row,sheet));
        loga=(getcell(2,row,sheet));
        if (loga.isNull()==0)
            ui->checklogo->setChecked(1);
        else loga.clear();
        loga=eadres+loga;
        ui->cont->setText(getcell(3,row,sheet));
        ui->odesc->setText(getcell(4,row,sheet));
        ui->pname->setText(getcell(5,row,sheet));
        imaga=(getcell(6,row,sheet));
        if (imaga.isNull()==0)
            ui->checkimage->setChecked(1);
        else imaga.clear();
        imaga=eadres+imaga;
        ui->problem->setText(getcell(7,row,sheet));
        ui->solution->setText(getcell(8,row,sheet));
        ui->actual->setText(getcell(9,row,sheet));
        ui->examp->setText(getcell(10,row,sheet));
        ui->author->setText(getcell(11,row,sheet));
        ui->categ->setText(getcell(12,row,sheet));
        ui->prot->setText(getcell(13,row,sheet));
        ui->invest->setText(getcell(14,row,sheet));
        ui->stage->setText(getcell(15,row,sheet));
        QAxObject* cell = sheet->querySubObject("Cells(int,int)", row, 16);
        QVariant value = cell->property("Value");
        ui->datedev->setDate(value.toDate());
        cell = sheet->querySubObject("Cells(int,int)", row, 17);
        value = cell->property("Value");
        ui->dateinv->setDate(value.toDate());
        ui->lang->setCurrentText(getcell(18,row,sheet));
        ui->olang->setText(getcell(19,row,sheet));
        ui->plang->setText(getcell(20,row,sheet));
        if (ui->olang->text().isEmpty()==0&&ui->plang->text().isEmpty()==0)
            ui->lang_2->setChecked(1);
        if (ui->oname->text()==ui->pname->text())
            break;
        addchan();
        clear();
    }
    clear();
    ui->info->setText(QString::number(ao)+" организаций добавилось\n"+QString::number(uo)+" организаций обновилось\n"+QString::number(ap)+" проектов добавилось\n"+QString::number(up)+" проектов обновилось");
    ao=ap=up=uo=0;
    workbook->dynamicCall("Close()");
    excel->dynamicCall("Quit()");
}

QString MainWindow::getcell(int c, int r, QAxObject* sheet) // возвращает данные из ячейки файла Excel
{
    QAxObject* cell = sheet->querySubObject("Cells(int,int)", r, c);
    QVariant value = cell->property("Value");
    return value.toString();
}


void MainWindow::clear() // очищает поля с данными
{
    ui->oname->clear();
    ui->checklogo->setChecked(0);
    loga.clear();
    ui->cont->clear();
    ui->odesc->clear();
    ui->pname->clear();
    ui->checkimage->setChecked(0);
    imaga.clear();
    ui->author->clear();
    ui->categ->clear();
    ui->problem->clear();
    ui->solution->clear();
    ui->actual->clear();
    ui->examp->clear();
    ui->prot->clear();
    ui->invest->clear();
    ui->stage->clear();
    ui->lang_2->setChecked(0);
    ui->olang->clear();
    ui->plang->clear();
}

QString MainWindow::select_image() // выбирает файл изображения
{
    return QFileDialog::getOpenFileName(0, "Open Dialog", "", "*.jpg *.bmp *.png");
}

MainWindow::~MainWindow()
{
    db.close();
    delete ui;
}

