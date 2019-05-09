#ifndef EMAILSENDER_H
#define EMAILSENDER_H

#include "common.h"
#include "config.h"
#include "emailsenderrunnable.h"
#include "emailtaskqueuedata.h"
#include "smtpclient/src/SmtpMime"
#include <QCoreApplication>
#include <QElapsedTimer>
#include <QFileInfo>
#include <QHash>
#include <QList>
#include <QMessageBox>
#include <QObject>
#include <QQueue>
#include <QString>
#include <QThreadPool>
#include <QTimer>
#include <QtCore/qmath.h>

class EmailSender : public QObject {
    Q_OBJECT
    public:
    EmailSender(QObject *parent = nullptr);
    ~EmailSender();

    void setSendData(Config *cfg, QHash<QString, QList<QStringList>> emailQhash, QString savePath, int total);
    void initTimer();
    void initIdleTimer();
    public slots:
    void doSend();
    void doSendWithoutQueue();
    void doSendWithQueue();
    void splitSendTask();
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
    QQueue<EmailTaskQueueData> emailTaskQueue;
    int m_total_cnt;
    int m_process_cnt;
    int m_success_cnt;
    int m_failure_cnt;
    int m_receive_msg_cnt;
    int m_current_queue_process_cnt;
    int m_current_queue_receive_cnt;
    int leftTime;

    int timeUnit;
    QElapsedTimer timer;
    bool use_queue = false;
    int periodTime;

    // idle msg 相关配置
    int idleLeftTime;
    QElapsedTimer idleTimer;
    int idleMsgShowPeriod;
    int idleTimeUnit;
};
#endif // EMAILSENDER_H
