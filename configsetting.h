#ifndef CONFIGSETTING_H
#define CONFIGSETTING_H

#include <QMainWindow>

namespace Ui {
class ConfigSetting;
}

class ConfigSetting : public QMainWindow
{
    Q_OBJECT

public:
    explicit ConfigSetting(QWidget *parent = nullptr);
    ~ConfigSetting();

private:
    Ui::ConfigSetting *ui;
};

#endif // CONFIGSETTING_H
