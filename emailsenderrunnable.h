#ifndef EMAILSENDERRUNABLE_H
#define EMAILSENDERRUNABLE_H
#include "common.h"
#include "smtpclient/src/SmtpMime"
#include <QFileInfo>
#include <QMetaObject>
#include <QObject>
#include <QRunnable>
class EmailSenderRunnable : public QRunnable {
    public:
    EmailSenderRunnable(QObject *mParent);
    ~EmailSenderRunnable();

    void run();

    void setID(const int &id);

    void setSendData(QString userName, QString password, QString server, QString defaultSender, QString savePath, QHash<QString, QList<QStringList>> fragmentEmailData);

    void requestMsg(const int msgType, const QString &result);

    private:
    //父对象
    QObject *mParent;
    int runnableID;
    QString defaultSender;
    QString savePath;
    QString userName;
    QString password;
    QString server;
    QHash<QString, QList<QStringList>> fragmentEmailData;
};

#endif // EMAILSENDERRUNABLE_H
