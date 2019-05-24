#ifndef EMAILONLYWINDOW_H
#define EMAILONLYWINDOW_H

#include "common.h"
#include "config.h"
#include "configsetting.h"
#include "emailsender.h"
#include "excelparser.h"
#include "officehelper.h"
#include "processwindow.h"
#include "splitsubwindow.h"
#include "xlsxdocument.h"
namespace Ui {
class EmailOnlyWindow;
}

class EmailOnlyWindow : public SplitSubWindow {
    Q_OBJECT

    public:
    explicit EmailOnlyWindow(QMainWindow *parent = nullptr);
    ~EmailOnlyWindow();

    void doSendEmail(QString dataSheetName, QString savePath);

    signals:
    void doProcess();

    private slots:
    void on_selectFilePushButton_clicked();
    void changeGroupby(QString selectedSheetName);
    void on_savePathPushButton_clicked();
    void on_emailOnlySubmitPushButton_clicked();
    void receiveMessage(const int msgType, const QString &result);
    void on_gobackPushButton_clicked();
    void showConfigSetting();

    private:
    Ui::EmailOnlyWindow *ui;
    QXlsx::Document *xlsx = nullptr;
    QStringList *header = new QStringList();
    ConfigSetting *configSetting = new ConfigSetting(this);

    Config *cfg = new Config();
    QThread *excelParserThread = nullptr;
    ExcelParser *excelParser = nullptr;
    QThread *mailSenderThread = nullptr;
    EmailSender *mailSender = nullptr;
    ProcessWindow *processWindow = nullptr;
    QString savePath;
    QString sourcePath;
};

#endif // EMAILONLYWINDOW_H
