#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QFileDialog>
#include <QMessageBox>
#include<QHashIterator>
#include <QDateTime>
#include<QErrorMessage>
#include<QThread>
#include <QDir>

#include "xlsxdocument.h"
#include "configsetting.h"
#include "config.h"
#include "emailsender.h"
#include "common.h"
#include "processwindow.h"
#include "xlsxparser.h"

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

    void doSplitXls(QString dataSheetName, QString emailSheetName, QString savePath);

    QHash<QString, QList<QStringList>> readXls(QString groupByText, QString selectedSheetName, bool isEmail);
    void writeXls(QHash<QString, QList<QStringList>> qHash, QString savePath);
    void sendemail(QHash<QString, QList<QStringList>> qHash, QString savePath);
    void writeXlsHeader(QXlsx::Document *currXls);

    void convertToColName(int data, QString &res);
    QString to26AlphabetString(int data);

    void loadConfig();

signals:
    void doSend();
    void doSplit();

private slots:
    void on_selectFilePushButton_clicked();

    void on_savePathPushButton_clicked();

    void on_submitPushButton_clicked();

    void changeGroupby(QString selectedSheetName);

    void on_dataComboBox_currentTextChanged(const QString &arg1);

    void showConfigSetting();

    void receiveMessage(const int msgType, const QString &result);

private:
    Ui::MainWindow *ui;
    QXlsx::Document *xlsx = nullptr;
    QStringList *header = new QStringList();
    ConfigSetting *configSetting = new ConfigSetting(nullptr,this);
    Config *cfg = new Config();
    QThread *xlsxParserThread = nullptr;
    XlsxParser * xlsxParser = new XlsxParser();
    QThread *mailSenderThread = nullptr;
    EmailSender *mailSender = nullptr;
    ProcessWindow *processWindow = nullptr;

    void errorMessage(const QString &message);
};

#endif // MAINWINDOW_H
