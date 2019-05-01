#ifndef EMAILSENDER_H
#define EMAILSENDER_H

#include <QObject>
#include <QHash>
#include <QList>
#include <QString>
#include <QMessageBox>
#include <QFileInfo>
#include <QThreadPool>
#include <QtCore/qmath.h>
#include <QCoreApplication>
#include<QElapsedTimer>
#include<QTimer>
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
    void doSendNew();
    void splitSendTask();
    void splitSendTaskNew();
    void stop();
    void showIdleMsg();
    void receiveMessage(const int msgType, const QString &result);

 signals:
    void requestMsg(const int msgType, const QString &result);
private:
    Config *cfg;
    QHash<QString, QList<QStringList>> emailQhash;
    QHash<QString, QList<QStringList>> currentEmailQhash;
    QString savePath;
    QString msg;
    int m_total_cnt;
    int m_process_cnt;
    int m_success_cnt;
    int m_failure_cnt;
    int m_receive_msg_cnt;

    int leftTime;
    int idleMsgShowPeriod;
    int timeUnit;
    QElapsedTimer timer;
};
#endif // EMAILSENDER_H
