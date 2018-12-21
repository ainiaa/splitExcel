#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QFileDialog>
#include <QMessageBox>
#include<QHashIterator>
#include <QDateTime>

#include "xlsxdocument.h"

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

    QString openFile();

    void doSplitXls(int groupby, QString savePath);

    QHash<QString, QList<QStringList>> readXls(int groupby);
    void writeXls(QHash<QString, QList<QStringList>> qhash, QString savePath);

    void writeXlsHeader(QXlsx::Document *currXls);

    void convertToColName(int data, QString &res);
    QString to26AlphabetString(int data);


private slots:
    void on_selectFilePushButton_clicked();

    void on_savePathPushButton_clicked();

    void on_submitPushButton_clicked();

private:
    Ui::MainWindow *ui;
    QXlsx::Document *xlsx;
    QStringList *header = new QStringList();
};

#endif // MAINWINDOW_H
