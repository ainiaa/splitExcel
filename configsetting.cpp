#include "configsetting.h"
#include "ui_configsetting.h"

ConfigSetting::ConfigSetting(QMainWindow *parent) : SplitSubWindow(parent), ui(new Ui::ConfigSetting) {
    ui->setupUi(this);
    loadConfig();
}

ConfigSetting::~ConfigSetting() {
    delete ui;
}

//加载配置项
void ConfigSetting::loadConfig() {
    ui->serverLineEdit->setText(cfg->get("email", "server").toString());
    ui->userNameLineEdit->setText(cfg->get("email", "userName").toString());
    ui->passwordLineEdit->setText(cfg->get("email", "password").toString());
    ui->maxThreadCntComboBox->setCurrentText(cfg->get("email", "maxThreadCnt").toString());
    ui->maxProcessPerPeriodLineEdit->setText(cfg->get("email", "maxProcessPerPeriod").toString());
    int excelLibType = cfg->get("email", "excelLibType").toInt();
    if (excelLibType == 0) { //自动识别
        ui->autoCheckRadioButton->setChecked(true);
    } else if (excelLibType == 1) { //自带类库
        ui->embbedLibRadioButton->setChecked(true);
    } else if (excelLibType == 2) { // ms office
        ui->msOfficeRadioButton->setChecked(true);
    }
}

//写入配置项
void ConfigSetting::writeConfig() {
    QString server = ui->serverLineEdit->text();
    QString userName = ui->userNameLineEdit->text();
    QString password = ui->passwordLineEdit->text();
    QString defaultSender = userName;
    QString maxThreadCnt = ui->maxThreadCntComboBox->currentText();
    QString maxProcessPerPeriod = ui->maxProcessPerPeriodLineEdit->text();
    int excelLibType = 0;
    if (ui->autoCheckRadioButton->isChecked()) { //自动识别
        excelLibType = 0;
    } else if (ui->embbedLibRadioButton->isChecked()) { //自带类库
        excelLibType = 1;
    } else if (ui->msOfficeRadioButton->isChecked()) { // ms office
        excelLibType = 2;
    }
    cfg->set("email", "server", server);
    cfg->set("email", "userName", userName);
    cfg->set("email", "defaultSender", defaultSender);
    cfg->set("email", "maxThreadCnt", maxThreadCnt);
    cfg->set("email", "maxProcessPerPeriod", maxProcessPerPeriod);
    cfg->set("email", "excelLibType", excelLibType);
    if (!password.isEmpty()) {
        cfg->set("email", "password", password);
    }
}

void ConfigSetting::on_submitPushButton_clicked() {
    writeConfig();
    QMessageBox::information(this, "配置项保存成功", "配置项保存成功");
}

//测试
void ConfigSetting::on_testPushButton_clicked() {
    QString server = ui->serverLineEdit->text();
    QString userName = ui->userNameLineEdit->text();
    QString password = ui->passwordLineEdit->text();
    if (server.isEmpty()) {
        QMessageBox::critical(this, "配置项错误", "server 配置为空");
        return;
    }
    if (userName.isEmpty()) {
        QMessageBox::critical(this, "配置项错误", "userName 配置为空");
        return;
    }
    if (password.isEmpty()) {
        QMessageBox::critical(this, "配置项错误", "password 配置为空");
        return;
    }

    SmtpClient *smtpClient = new SmtpClient(server);
    smtpClient->setUser(userName);
    smtpClient->setPassword(password);

    if (!smtpClient->connectToHost()) {
        QMessageBox::critical(this, "配置项错误", "邮件服务器连接失败！！");
        return;
    }

    if (!smtpClient->login(userName, password)) {
        QMessageBox::critical(this, "配置项错误", "邮件服务器认证失败（邮件用户名或者密码错误）!!");
        return;
    }
    QMessageBox::information(this, "配置项正确", "配置正确!!");
}

void ConfigSetting::on_gobackButton_clicked() {
    this->hide();
    this->getMainWindow()->show();
}
