#ifndef EMAILSENDER_H
#define EMAILSENDER_H

#include <QObject>
#include <QHash>
#include <QList>
#include <QString>
#include <QMessageBox>
#include <QFileInfo>
#include <QThreadPool>
#include "smtpclient/src/SmtpMime"
#include "config.h"
#include "common.h"
#include "emailsenderrunnable.h"

class EmailSender: public QObject
{
    Q_OBJECT
public:
    EmailSender(QObject* parent = nullptr);
    ~EmailSender();

    void setSendData(Config *cfg, QHash<QString, QList<QStringList>> emailQhash, QString savePath, int total);

public slots:
    void doSend();
    void stop();
    void receiveMessage(const int msgType, const QString &result);

 signals:
    void requestMsg(const int msgType, const QString &result);
private:
    Config *cfg;
    QHash<QString, QList<QStringList>> emailQhash;
    QString savePath;
    QString msg;
    int m_total_cnt;
    int m_process_cnt;
    int m_success_cnt;
    int m_failure_cnt;
    int m_receive_msg_cnt;
};
#endif // EMAILSENDER_H
