#ifndef CONFIGSETTING_H
#define CONFIGSETTING_H

#include "config.h"
#include "smtpclient/src/SmtpMime"
#include "splitsubwindow.h"
#include <QMainWindow>
#include <QMessageBox>
namespace Ui {
class ConfigSetting;
}

class ConfigSetting : public SplitSubWindow {
    Q_OBJECT

    public:
    explicit ConfigSetting(QMainWindow *parent = nullptr);
    ~ConfigSetting();

    void loadConfig();
    void writeConfig();

    private slots:
    void on_submitPushButton_clicked();
    void on_testPushButton_clicked();
    void on_gobackButton_clicked();

    private:
    Ui::ConfigSetting *ui;
    Config *cfg = new Config();
};

#endif // CONFIGSETTING_H
