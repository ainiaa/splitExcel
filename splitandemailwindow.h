#ifndef SPLITANDEMAILWINDOW_H
#define SPLITANDEMAILWINDOW_H

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
#include "splitsubwindow.h"
#include "testtimeer.h"
#include "xlsxdocument.h"

namespace Ui {
class SplitAndEmailWindow;
}

class SplitAndEmailWindow : public SplitSubWindow {
    Q_OBJECT

    public:
    SplitAndEmailWindow(QMainWindow *parent = nullptr);
    ~SplitAndEmailWindow();
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

    void receiveMessage(const int msgType, const QString &result);

    void on_gobackButton_clicked();

    private:
    Ui::SplitAndEmailWindow *ui;
    QXlsx::Document *xlsx = nullptr;
    QStringList *header = new QStringList();
    ConfigSetting *configSetting = new ConfigSetting(nullptr, this);
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

#endif // SPLITANDEMAILWINDOW_H
