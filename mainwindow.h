#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QDateTime>
#include <QDir>
#include <QErrorMessage>
#include <QFileDialog>
#include <QHash>
#include <QHashIterator>
#include <QMainWindow>
#include <QMessageBox>
#include <QThread>

#include <QException>

#include "common.h"
#include "config.h"
#include "configsetting.h"
#include "emailsender.h"
#include "excelparser.h"
#include "officehelper.h"
#include "processwindow.h"
#include "sourceexceldata.h"
#include "splitonlywindow.h"
#include "testtimeer.h"
#include "xlsxdocument.h"
namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow {
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

    void on_cancelPushButton_2_clicked();

    private:
    Ui::MainWindow *ui;
    QXlsx::Document *xlsx = nullptr;
    QStringList *header = new QStringList();
    ConfigSetting *configSetting = new ConfigSetting(nullptr, this);
    SplitOnlyWindow *splitOnlyWindow = new SplitOnlyWindow(nullptr, this);
    TestTimeer *testTimeer = new TestTimeer(nullptr, this);

    Config *cfg = new Config();
    QThread *xlsxParserThread = nullptr;
    ExcelParser *xlsxParser = nullptr;
    QThread *mailSenderThread = nullptr;
    EmailSender *mailSender = nullptr;
    ProcessWindow *processWindow = nullptr;
    QString savePath;
    QString sourcePath;

    void errorMessage(const QString &message);
};

#endif // MAINWINDOW_H
