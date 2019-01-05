#ifndef EMAILSENDERRUNABLE_H
#define EMAILSENDERRUNABLE_H
#include <QObject>
#include <QRunnable>
#include <QMetaObject>
#include <QFileInfo>
#include "smtpclient/src/SmtpMime"
#include "common.h"
class EmailSenderRunnable : public QRunnable
{
public:
    EmailSenderRunnable(QObject *mParent);
    ~EmailSenderRunnable();

    void run();

    void setID(const int &id);

    void setSendData(QString userName,QString password, QString server,QString defaultSender,QString savePath,QString key,QList<QStringList> content);

    void requestMsg(const int msgType, const QString &result);

private:
    //父对象
    QObject *mParent;
    int runnableID;
    QString defaultSender;
    QString savePath;
    QString key;
    QString userName;
    QString password;
    QString server;
    QList<QStringList> content;
};

#endif // EMAILSENDERRUNABLE_H
