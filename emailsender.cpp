#include "emailsender.h"

EmailSender::EmailSender(QObject *parent) : QObject(parent) {
    m_success_cnt = 0;
    m_failure_cnt = 0;
    m_process_cnt = 0;
    m_receive_msg_cnt = 0;
}
EmailSender::~EmailSender() {
    qDebug() << "EmailSender::~EmailSender";
}

void EmailSender::setSendData(Config *cfg, QHash<QString, QList<QStringList>> emailQhash, QString savePath, int total) {
    this->cfg = cfg;
    this->emailQhash = emailQhash;
    this->savePath = savePath;
    this->m_total_cnt = total;
    QString msg("发送邮件。。。 成功：%1封  失败：%2封 /共：");

    msg.append(QString::number(m_total_cnt)).append("封");
    msg.append("    补充信息: %4 ");
    this->msg = msg;
}

//发送email
void EmailSender::doSend() {
    qDebug("sendemail");
    QHashIterator<QString, QList<QStringList>> it(emailQhash);

    QString user(cfg->get("email", "userName").toString());
    QString password(cfg->get("email", "password").toString());
    QString server(cfg->get("email", "server").toString());
    QString defaultSender(cfg->get("email", "defaultSender").toString());
    int maxProcessPerPeriod(cfg->get("email", "maxProcessPerPeriod").toInt()); // 没单位时间可以执行的最大数量
    if (user.isEmpty() || password.isEmpty() || server.isEmpty() || defaultSender.isEmpty()) {
        emit requestMsg(Common::MsgTypeError, "请先正确设置配置信息");
        return;
    }

    if (maxProcessPerPeriod > 0) { //使用队列处理
        this->splitSendTask();
    } else {
        this->doSendWithoutQueue();
    }
}

void EmailSender::doSendWithoutQueue() {
    qDebug("doSendWithoutQueue");
    QHashIterator<QString, QList<QStringList>> it(emailQhash);

    QString user(cfg->get("email", "userName").toString());
    QString password(cfg->get("email", "password").toString());
    QString server(cfg->get("email", "server").toString());
    QString defaultSender(cfg->get("email", "defaultSender").toString());
    if (user.isEmpty() || password.isEmpty() || server.isEmpty() || defaultSender.isEmpty()) {
        emit requestMsg(Common::MsgTypeError, "请先正确设置配置信息");
        return;
    }
    emit requestMsg(Common::MsgTypeInfo, "准备发送邮件...");

    QThreadPool pool;
    int maxThreadCnt = cfg->get("email", "maxThreadCnt").toInt();
    if (maxThreadCnt < 1) {
        maxThreadCnt = 2;
    }

    pool.setMaxThreadCount(maxThreadCnt);

    int emailQhashCnt = emailQhash.size();
    int maxProcessCntPreThread = qCeil(emailQhashCnt * 1.0 / maxThreadCnt);
    qDebug() << "emailQhash.size():" << emailQhashCnt << " maxThreadCnt:" << maxThreadCnt << " maxProcessCntPreThread:" << maxProcessCntPreThread;
    QHash<QString, QList<QStringList>> fragmentEmailQhash;
    int runnableId = 1;
    while (it.hasNext()) {
        it.next();
        ++m_process_cnt;
        QString key = it.key();
        QList<QStringList> content = it.value();
        fragmentEmailQhash.insert(key, content);
        int mod = m_process_cnt % maxProcessCntPreThread;
        if (mod == 0 || !it.hasNext()) {
            EmailSenderRunnable *runnable = new EmailSenderRunnable(this);
            runnable->setID(runnableId++);
            runnable->setSendData(user, password, server, defaultSender, savePath, fragmentEmailQhash);
            runnable->setAutoDelete(true);
            pool.start(runnable);
            fragmentEmailQhash.clear();
        }
    }
    pool.waitForDone();
    qDebug("处理完毕！");
}

void EmailSender::splitSendTask() {
    qDebug("splitSendTask");

    QString msg("本次邮件发送共 %1 部分，当前为 第 %2 部分,包含待发送 email 数量为 %3");
    int maxProcessPerPeriod(cfg->get("email", "maxProcessPerPeriod").toInt()); // 没单位时间可以执行的最大数量
    if (maxProcessPerPeriod == 0) {
        maxProcessPerPeriod = 1000;
    }

    int splitTaskCnt = qCeil(emailQhash.size() * 1.0 / maxProcessPerPeriod);
    qDebug() << "splitSendTask"
             << " emailQhash.size() :" << emailQhash.size() << " maxProcessPerPeriod:" << maxProcessPerPeriod;
    QHashIterator<QString, QList<QStringList>> it(emailQhash);
    int emailQhashCnt = emailQhash.size();
    for (int i = 0; i < splitTaskCnt; i++) {
        int currentPartationCnt = 0;
        QHash<QString, QList<QStringList>> fragmentEmailQhash;
        while (it.hasNext()) {
            it.next();
            ++currentPartationCnt;
            QString key = it.key();
            QList<QStringList> content = it.value();
            fragmentEmailQhash.insert(key, content);
            if (currentPartationCnt == maxProcessPerPeriod || !it.hasNext()) {
                EmailTaskQueueData queueData;
                queueData.setMsg(msg.arg(splitTaskCnt).arg(i + 1).arg(fragmentEmailQhash.size()));
                queueData.setTotal(emailQhashCnt);
                queueData.setSavePath(this->savePath);
                queueData.setEmailQhash(fragmentEmailQhash);

                emailTaskQueue.enqueue(queueData);
                break;
            }
        }
    }
    this->doSendWithQueue();
    qDebug("splitSendTask 处理完毕");
}

void EmailSender::initTimer() {
    qDebug() << "initTimer";
    int timeUnit = 1000;            //时间单位
    int periodTime = timeUnit * 60; // 1min

    this->leftTime = periodTime;
    this->timeUnit = timeUnit;
    QElapsedTimer timer;
    timer.start();
    this->timer = timer;
}

void EmailSender::initIdleTimer() {
    qDebug() << "initIdleTimer";
    int timeUnit = 1000;                  //时间单位
    int periodTime = timeUnit * 60;       // 1min
    int idleMsgShowPeriod = 2 * timeUnit; // 2s一个周期

    this->idleLeftTime = periodTime;
    this->idleMsgShowPeriod = idleMsgShowPeriod;
    this->idleTimeUnit = timeUnit;
    QElapsedTimer timer;
    timer.start();
    this->idleTimer = timer;
}

//发送email(待队列)
void EmailSender::doSendWithQueue() {
    qDebug("doSendWithQueue");

    if (emailTaskQueue.isEmpty()) {
        emit requestMsg(Common::MsgTypeError, "没有邮件待发送...");
        return;
    }
    this->use_queue = true;

    EmailTaskQueueData emailTaskQueueData = emailTaskQueue.dequeue();
    QHash<QString, QList<QStringList>> currentEmailQhash = emailTaskQueueData.getEmailQhash();
    QHashIterator<QString, QList<QStringList>> it(currentEmailQhash);
    QString currentSavePath = emailTaskQueueData.getSavePath();
    this->m_current_queue_process_cnt = currentEmailQhash.size();
    this->m_current_queue_receive_cnt = 0;

    emit requestMsg(Common::MsgTypeInfo, emailTaskQueueData.getMsg());

    QString user(cfg->get("email", "userName").toString());
    QString password(cfg->get("email", "password").toString());
    QString server(cfg->get("email", "server").toString());
    QString defaultSender(cfg->get("email", "defaultSender").toString());
    if (user.isEmpty() || password.isEmpty() || server.isEmpty() || defaultSender.isEmpty()) {
        emit requestMsg(Common::MsgTypeError, "请先正确设置配置信息");
        return;
    }
    emit requestMsg(Common::MsgTypeInfo, "准备发送邮件...");

    QThreadPool pool;
    int maxThreadCnt = cfg->get("email", "maxThreadCnt").toInt();
    if (maxThreadCnt < 1) {
        maxThreadCnt = 2;
    }
    maxThreadCnt = 5; //禁用多线程

    pool.setMaxThreadCount(maxThreadCnt);

    int emailQhashCnt = currentEmailQhash.size();
    int maxProcessCntPreThread = qCeil(emailQhashCnt * 1.0 / maxThreadCnt);
    qDebug() << "currentEmailQhash.size():" << emailQhashCnt << " maxThreadCnt:" << maxThreadCnt
             << " maxProcessCntPreThread:" << maxProcessCntPreThread;
    QHash<QString, QList<QStringList>> fragmentEmailQhash;
    int runnableId = 1;
    while (it.hasNext()) {
        it.next();
        ++m_process_cnt;
        QString key = it.key();
        QList<QStringList> content = it.value();
        fragmentEmailQhash.insert(key, content);
        int mod = m_process_cnt % maxProcessCntPreThread;
        if (mod == 0 || !it.hasNext()) {
            EmailSenderRunnable *runnable = new EmailSenderRunnable(this);
            runnable->setID(runnableId++);
            runnable->setSendData(user, password, server, defaultSender, currentSavePath, fragmentEmailQhash);
            runnable->setAutoDelete(true);
            pool.start(runnable);
            fragmentEmailQhash.clear();
        }
    }
    pool.waitForDone();
    qDebug() << emailTaskQueueData.getMsg() << "处理完毕！";
}

void EmailSender::showIdleMsg() {
    QString idleMsg("达到当前处理上限，暂时进入休息状态. 剩余休息时间 %1s");

    int ela = this->idleTimer.elapsed();
    int secs = 0; // this->idleTimer.secsTo(this->idleTimer);
    int currentLeftTime = this->idleLeftTime - ela;
    qDebug() << "showIdleMsg ela:" << ela << "  secs:" << secs << " currentLeftTime:" << currentLeftTime;
    if (currentLeftTime > 0) {
        this->receiveMessage(Common::MsgTypeInfo, idleMsg.arg(currentLeftTime / this->idleTimeUnit));
        QTimer::singleShot(this->idleTimeUnit, this, SLOT(showIdleMsg()));
    } else {
        this->receiveMessage(Common::MsgTypeInfo, "休息状态结束.");
        QTimer::singleShot(this->idleTimeUnit, this, SLOT(doSendWithQueue())); //尝试继续处理
    }
}

void EmailSender::receiveMessage(const int msgType, const QString &result) {
    qDebug() << "EmailSender::receiveMessage msgType:" << QString::number(msgType).toUtf8() << " msg:" << result;
    bool continued = false;
    switch (msgType) {
        case Common::MsgTypeError:
        case Common::MsgTypeFail:
            m_failure_cnt++;
            m_receive_msg_cnt++;
            m_current_queue_receive_cnt++;
            emit requestMsg(msgType, msg.arg(m_success_cnt).arg(m_failure_cnt).arg(result));
            continued = true;
            break;
        case Common::MsgTypeSucc:
            m_success_cnt++;
            m_receive_msg_cnt++;
            m_current_queue_receive_cnt++;
            emit requestMsg(msgType, msg.arg(m_success_cnt).arg(m_failure_cnt).arg(result));
            continued = true;
            break;
        case Common::MsgTypeInfo:
        case Common::MsgTypeWarn:
        case Common::MsgTypeFinish:
        default:
            emit requestMsg(msgType, result);
            break;
    }
    if (!continued) {
        return;
    }
    if (!this->use_queue) {                                        //没有使用队列
        if (m_total_cnt > 0 && m_receive_msg_cnt == m_total_cnt) { //全部处理完毕
            emit requestMsg(Common::MsgTypeEmailSendFinish, "处理完毕！");
        }
    } else {                                                                                                 //使用队列
        if (m_current_queue_process_cnt > 0 && m_current_queue_receive_cnt == m_current_queue_process_cnt) { //当前处理完毕
            if (this->emailTaskQueue.isEmpty()) { //全部处理完毕
                emit requestMsg(Common::MsgTypeEmailSendFinish, "全部邮件处理完毕！");
            } else { //还有邮件没有处理
                /*
                if (this->timer.elapsed() > this->periodTime)
                { //当前处理时间 大于周期时间 无需等待
                    //QTimer::singleShot(0, this, SLOT(doSendWithQueue()));
                }
                else
                { //当前处理时间 大于周期时间 需要等待
                    //QTimer::singleShot(0, this, SLOT(showIdleMsg()));
                }
                //this->initIdleTimer();//初始化计时器
                //this->showIdleMsg();
                //QTimer::singleShot(0, this, SLOT(showIdleMsg()));
                */
                // QTimer::singleShot(10, this, SLOT(doSendWithQueue()));
                // doSendWithQueue();
                this->initIdleTimer(); //初始化计时器
                QTimer::singleShot(0, this, SLOT(showIdleMsg()));
            }
        }
    }
}

//停止发送email
void EmailSender::stop() {
}
