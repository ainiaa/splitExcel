#include "configsetting.h"
#include "ui_configsetting.h"

ConfigSetting::ConfigSetting(QWidget *parent, QMainWindow *mainWindow) :
    QMainWindow(parent),
    ui(new Ui::ConfigSetting)
{
    this->mainWindow = mainWindow;
    ui->setupUi(this);
    loadConfig();
}

ConfigSetting::~ConfigSetting()
{
    delete ui;
}


//加载配置项
void ConfigSetting::loadConfig()
{
    ui->serverLineEdit->setText(cfg->get("email","server").toString());
    ui->userNameLineEdit->setText(cfg->get("email","userName").toString());
    ui->passwordLineEdit->setText(cfg->get("email","password").toString());
    ui->maxThreadCntComboBox->setCurrentText(cfg->get("email","maxThreadCnt").toString());
}

//写入配置项
void ConfigSetting::writeConfig()
{
    QString server = ui->serverLineEdit->text();
    QString userName = ui->userNameLineEdit->text();
    QString password = ui->passwordLineEdit->text();
    QString defaultSender =userName;
    QString maxThreadCnt = ui->maxThreadCntComboBox->currentText();
    cfg->set("email","server",server);
    cfg->set("email","userName",userName);
    cfg->set("email","defaultSender",defaultSender);
    cfg->set("email","maxThreadCnt",maxThreadCnt);
    if (!password.isEmpty())
    {
        cfg->set("email","password",password);
    }
}

void ConfigSetting::on_submitPushButton_clicked()
{
    writeConfig();
    QMessageBox::information(this, "配置项保存成功", "配置项保存成功");
}

void ConfigSetting::on_gobackPushButton_clicked()
{
    this->hide();
    mainWindow->show();
}

//测试
void ConfigSetting::on_testPushButton_clicked()
{
    QString server = ui->serverLineEdit->text();
    QString userName = ui->userNameLineEdit->text();
    QString password = ui->passwordLineEdit->text();
    if (server.isEmpty())
    {
        QMessageBox::critical(this, "配置项错误", "server 配置为空");
        return;
    }
    if (userName.isEmpty())
    {
        QMessageBox::critical(this, "配置项错误", "userName 配置为空");
        return;
    }
    if (password.isEmpty())
    {
        QMessageBox::critical(this, "配置项错误", "password 配置为空");
        return;
    }

    SmtpClient *smtpClient = new SmtpClient(server);
    smtpClient->setUser(userName);
    smtpClient->setPassword(password);

    if(!smtpClient->connectToHost())
    {
        QMessageBox::critical(this, "配置项错误", "邮件服务器连接失败！！");
        return;
    }

    if (!smtpClient->login(userName, password))
    {
        QMessageBox::critical(this, "配置项错误", "邮件服务器认证失败（邮件用户名或者密码错误）!!");
        return;
    }
    QMessageBox::information(this, "配置项正确", "配置正确!!");
}
