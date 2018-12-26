#include "configsetting.h"
#include "ui_configsetting.h"

ConfigSetting::ConfigSetting(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::ConfigSetting)
{
    ui->setupUi(this);
}

ConfigSetting::~ConfigSetting()
{
    delete ui;
}
