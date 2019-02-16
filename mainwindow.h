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
#include <QHash>

#include <QException>

#include "xlsxdocument.h"
#include "configsetting.h"
#include "splitonlywindow.h"
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
    void sendemail();
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

    void showSplitOnly();

    void receiveMessage(const int msgType, const QString &result);

private:
    Ui::MainWindow *ui;
    QXlsx::Document *xlsx = nullptr;
    QStringList *header = new QStringList();
    ConfigSetting *configSetting = new ConfigSetting(nullptr,this);
    SplitOnlyWindow *splitOnlyWindow = new SplitOnlyWindow(nullptr,this);

    Config *cfg = new Config();
    QThread *xlsxParserThread = nullptr;
    XlsxParser * xlsxParser = nullptr;
    QThread *mailSenderThread = nullptr;
    EmailSender *mailSender = nullptr;
    ProcessWindow *processWindow = nullptr;
    QString savePath;

    void errorMessage(const QString &message);
};

#endif // MAINWINDOW_H
