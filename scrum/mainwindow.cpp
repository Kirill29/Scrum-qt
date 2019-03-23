#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <iostream>
#include <fstream>
#include <string>
#include <QFile>
#include <QStandardPaths>
#include <QtDebug>
#include <QFileDialog>
#include <QtWidgets>
#include <QtGui>
#include <iostream>
#include <fstream>
#include <string>
#include <list>
#include <QAxObject>
#include <QtCore>
using namespace std;



MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);

}

MainWindow::~MainWindow()
{
    delete ui;
}



//открываем файл ворд
void MainWindow::on_pushButton_clicked()
{
    QList<string> sentences;//список под предложения

    string line;//предложение в тексте
    //выбираем файл
    QString str = QFileDialog::getOpenFileName( this,"Open Dialog", "", "*");
    //если формат файла docx/doc
    if ((str.indexOf(".docx")>0) || (str.indexOf((".docx")))){
        sentences.clear();//чистим список

        try {
            QAxObject wordApplication("Word.Application");//привязка к приложению WORD(необходимо наличие на компьютере ПОЛЬЗОВАТЕЛЯ)

            QAxObject *documents = wordApplication.querySubObject("Documents");
            //открываем документ по ранее выбранному адресу
            QAxObject *document = documents->querySubObject("Open(const QString&, bool)", str, true);
            //указываем,что будем считывать предложения
            QAxObject *words = document->querySubObject("Sentences");
           //подсчитываем число предложений в файле
            int countWord = words->dynamicCall("Count()").toInt();
            if (countWord==0){
                 QMessageBox::warning(0,"Warning", "Файл пустой!");
            }
            for (int a = 1; a <= countWord; a++){
                //записываем предложение в список
                sentences.push_back(line);
            }
            //закрываем файл
            document->dynamicCall("Close (boolean)", false);
            //выводим текст по предложениям на форму
            foreach (string sentence,sentences){
                QString qstr = QString::fromStdString(sentence);
                ui->out_file->append(qstr);
            }
                delete words;


        } catch (exception ex) {

             QMessageBox::warning(0,"Warning", ex.what());
        }

    }
    //если формат файла txt
    else if (str.indexOf(".txt")>0){
        sentences.clear();//чистим список
        //открываем поток на чтение
        ifstream in(str.toStdString());
        //проверка на правильный путь к файлу
         if (!(in.is_open()))
         {
             QMessageBox::warning(0,"Warning", "Невозможно открыть файл!");
             return;

         }
         //если первый символт переход на новую строку-файл пустой
          else if (in.peek() == EOF) {
              QMessageBox::warning(0,"Warning", "Файл пустой!");
              return;
          }else{


             //считываем файл по строчке.Конец строчки-точка.
             while (getline(in, line, '.'))
             {
                 //записываем предложение в список
                 sentences.push_back(line);
             }
         }
         //выводим текст по предложениям на форму
         foreach (string sentence,sentences){
             QString qstr = QString::fromStdString(sentence);
             ui->out_file->append(qstr);
         }

         //закрываем файл
         in.close();



    }



 //другие форматы?
    else{
        QMessageBox::warning(0,"Warning", "Неправильный формат файла");
        return;
    }


}
