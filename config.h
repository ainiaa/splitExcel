#ifndef CONFIG_H
#define CONFIG_H

#include <QVariant>
#include <QSettings>

class Config
{
public:
    Config(QString qstrfilename = "");
    virtual ~Config(void);
    void set(QString,QString,QVariant);
    QString getConfigPath();
    QVariant get(QString,QString);
    void clear();
private:
    QString fileName;
    QSettings *cfg;
};
#endif // CONFIG_H
