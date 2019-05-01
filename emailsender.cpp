#include "emailsender.h"

EmailSender::EmailSender(QObject *parent):QObject (parent)
{
    m_success_cnt = 0;
    m_failure_cnt = 0;
    m_process_cnt = 0;
    m_receive_msg_cnt = 0;
}
EmailSender::~EmailSender()
{
        qDebug()<< "EmailSender::~EmailSender";
}

void EmailSender::setSendData(Config *cfg, QHash<QString, QList<QStringList>> emailQhash, QString savePath, int total)
{
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
void EmailSender::doSend()
{
    qDebug("sendemail");
    QHashIterator<QString,QList<QStringList>> it(emailQhash);

    QString user(cfg->get("email","userName").toString());
    QString password(cfg->get("email","password").toString());
    QString server(cfg->get("email","server").toString());
    QString defaultSender(cfg->get("email","defaultSender").toString());
    if (user.isEmpty() || password.isEmpty() || server.isEmpty() || defaultSender.isEmpty())
    {
        emit requestMsg(Common::MsgTypeError, "请先正确设置配置信息");
        return;
    }
    emit requestMsg(Common::MsgTypeInfo, "准备发送邮件...");

    QThreadPool pool;
    int maxThreadCnt = cfg->get("email","maxThreadCnt").toInt();
    if (maxThreadCnt < 1)
    {
        maxThreadCnt = 2;
    }

    pool.setMaxThreadCount(maxThreadCnt);

    int emailQhashCnt = emailQhash.size();
    int maxProcessCntPreThread = qCeil(emailQhashCnt * 1.0 / maxThreadCnt);
    qDebug() << "emailQhash.size():" << emailQhashCnt << " maxThreadCnt:" <<maxThreadCnt <<" maxProcessCntPreThread:" << maxProcessCntPreThread;
    QHash<QString, QList<QStringList>> fragmentEmailQhash;
    int runnableId = 1;
    while (it.hasNext()) {
        it.next();
        ++m_process_cnt;
        QString key = it.key();
        QList<QStringList> content = it.value();
        fragmentEmailQhash.insert(key,content);
        int mod = m_process_cnt % maxProcessCntPreThread;
        if (mod==0 || !it.hasNext())
        {
            EmailSenderRunnable *runnable = new EmailSenderRunnable(this) ;
            runnable->setID(runnableId++);
            runnable->setSendData(user,password,server,defaultSender,savePath,fragmentEmailQhash);
            runnable->setAutoDelete(true);
            pool.start(runnable);
            fragmentEmailQhash.clear();
        }
    }
    qDebug("处理完毕！") ;
}

//发送email
void EmailSender::doSendNew()
{
    qDebug("sendemail");
    QHashIterator<QString,QList<QStringList>> it(emailQhash);

    QString user(cfg->get("email","userName").toString());
    QString password(cfg->get("email","password").toString());
    QString server(cfg->get("email","server").toString());
    QString defaultSender(cfg->get("email","defaultSender").toString());
    if (user.isEmpty() || password.isEmpty() || server.isEmpty() || defaultSender.isEmpty())
    {
        emit requestMsg(Common::MsgTypeError, "请先正确设置配置信息");
        return;
    }
    emit requestMsg(Common::MsgTypeInfo, "准备发送邮件...");

    QThreadPool pool;
    int maxThreadCnt = cfg->get("email","maxThreadCnt").toInt();
    if (maxThreadCnt < 1)
    {
        maxThreadCnt = 2;
    }

    pool.setMaxThreadCount(maxThreadCnt);

    int emailQhashCnt = emailQhash.size();
    int maxProcessCntPreThread = qCeil(emailQhashCnt * 1.0 / maxThreadCnt);
    qDebug() << "emailQhash.size():" << emailQhashCnt << " maxThreadCnt:" <<maxThreadCnt <<" maxProcessCntPreThread:" << maxProcessCntPreThread;
    QHash<QString, QList<QStringList>> fragmentEmailQhash;
    int runnableId = 1;
    while (it.hasNext()) {
        it.next();
        ++m_process_cnt;
        QString key = it.key();
        QList<QStringList> content = it.value();
        fragmentEmailQhash.insert(key,content);
        int mod = m_process_cnt % maxProcessCntPreThread;
        if (mod==0 || !it.hasNext())
        {
            EmailSenderRunnable *runnable = new EmailSenderRunnable(this) ;
            runnable->setID(runnableId++);
            runnable->setSendData(user,password,server,defaultSender,savePath,fragmentEmailQhash);
            runnable->setAutoDelete(true);
            pool.start(runnable);
            fragmentEmailQhash.clear();
        }
    }
    qDebug("处理完毕！") ;
}

void EmailSender::splitSendTask()
{
    qDebug("sendemail");
    QHashIterator<QString,QList<QStringList>> it(emailQhash);
    QString user(cfg->get("email","userName").toString());
    QString password(cfg->get("email","password").toString());
    QString server(cfg->get("email","server").toString());
    QString defaultSender(cfg->get("email","defaultSender").toString());
    int maxProcessPerPeriod(cfg->get("email", "maxProcessPerPeriod").toInt());;// 没单位时间可以执行的最大数量
    if (maxProcessPerPeriod == 0)
    {
        maxProcessPerPeriod = 1000;
    }
    if (user.isEmpty() || password.isEmpty() || server.isEmpty() || defaultSender.isEmpty())
    {
        emit requestMsg(Common::MsgTypeError, "请先正确设置配置信息");
        return;
    }
    emit requestMsg(Common::MsgTypeInfo, "准备发送邮件...");

    QThreadPool pool;
    int maxThreadCnt = cfg->get("email","maxThreadCnt").toInt();
    if (maxThreadCnt < 1)
    {
        maxThreadCnt = 2;
    }

    pool.setMaxThreadCount(maxThreadCnt);


    int splitTaskCnt =qCeil(emailQhash.size() * 1.0 / maxProcessPerPeriod);

    for (int i = 0; i < splitTaskCnt; i++)
    {
        int currentProcessCnt = 0;
        int emailQhashCnt = emailQhash.size();
        int currentCnt = qMin(emailQhashCnt,maxProcessPerPeriod);
        int maxProcessCntPreThread = qCeil(currentCnt * 1.0 / maxThreadCnt);
        qDebug() << "currentCnt:" << currentCnt << " maxThreadCnt:" <<maxThreadCnt <<" maxProcessCntPreThread:" << maxProcessCntPreThread;
        QHash<QString, QList<QStringList>> fragmentEmailQhash;
        int runnableId = 1;
        while (it.hasNext()) {
            it.next();
            ++currentProcessCnt;
            ++m_process_cnt;
            QString key = it.key();
            QList<QStringList> content = it.value();
            emailQhash.remove(key);
            fragmentEmailQhash.insert(key,content);
            int mod = m_process_cnt % maxProcessCntPreThread;
            if (mod==0 || !it.hasNext())
            {
                EmailSenderRunnable *runnable = new EmailSenderRunnable(this) ;
                runnable->setID(runnableId++);
                runnable->setSendData(user,password,server,defaultSender,savePath,fragmentEmailQhash);
                runnable->setAutoDelete(true);
                pool.start(runnable);
                fragmentEmailQhash.clear();

                if (currentProcessCnt >= currentCnt)
                {
                    int timeUnit = 1000; //时间单位
                    int periodTime = timeUnit * 60; //1min
                    int idleMsgShowPeriod = 2 * timeUnit;//2s一个周期

                    this->leftTime = periodTime;
                    QElapsedTimer timer;
                    timer.start();
                    this->idleMsgShowPeriod = idleMsgShowPeriod;
                    this->timeUnit = timeUnit;
                    this->timer = timer;
                    this->showIdleMsg();
                }
            }

        }
    }

    qDebug("处理完毕！") ;
}

void EmailSender::splitSendTaskNew()
{
    qDebug("sendemail");
    QHashIterator<QString,QList<QStringList>> it(emailQhash);

    int maxProcessPerPeriod(cfg->get("email", "maxProcessPerPeriod").toInt());;// 没单位时间可以执行的最大数量
    if (maxProcessPerPeriod == 0)
    {
        maxProcessPerPeriod = 1000;
    }

    int maxThreadCnt = cfg->get("email","maxThreadCnt").toInt();
    if (maxThreadCnt < 1)
    {
        maxThreadCnt = 2;
    }

    int splitTaskCnt =qCeil(emailQhash.size() * 1.0 / maxProcessPerPeriod);

    for (int i = 0; i < splitTaskCnt; i++)
    {
        int currentProcessCnt = 0;
        int emailQhashCnt = emailQhash.size();
        int currentCnt = qMin(emailQhashCnt,maxProcessPerPeriod);
        int maxProcessCntPreThread = qCeil(currentCnt * 1.0 / maxThreadCnt);
        qDebug() << "currentCnt:" << currentCnt << " maxThreadCnt:" <<maxThreadCnt <<" maxProcessCntPreThread:" << maxProcessCntPreThread;
        QHash<QString, QList<QStringList>> fragmentEmailQhash;
        int runnableId = 1;
        while (it.hasNext()) {
            it.next();
            ++currentProcessCnt;
            QString key = it.key();
            QList<QStringList> content = it.value();
            emailQhash.remove(key);
            fragmentEmailQhash.insert(key,content);
            int mod = m_process_cnt % maxProcessCntPreThread;
            if (currentProcessCnt >= currentCnt)
            {
                if (mod==0 || !it.hasNext())
                {
                    int timeUnit = 1000; //时间单位
                    int periodTime = timeUnit * 60; //1min
                    int idleMsgShowPeriod = 2 * timeUnit;//2s一个周期

                    this->leftTime = periodTime;
                    QElapsedTimer timer;
                    timer.start();
                    this->idleMsgShowPeriod = idleMsgShowPeriod;
                    this->timeUnit = timeUnit;
                    this->timer = timer;
                    this->showIdleMsg();
                }
            }
        }
    }

    qDebug("处理完毕！") ;
}

void EmailSender::showIdleMsg()
{
    QString idleMsg("达到当前处理上线，暂时进入休息状态. 剩余休息时间 %1 %2s");

    int ela = this->timer.elapsed();
    int currentLeftTime = this->leftTime - ela;

    this->receiveMessage(Common::MsgTypeInfo, idleMsg.arg(ela).arg(currentLeftTime));
    if (currentLeftTime > 0)
    {
        if (ela % this->idleMsgShowPeriod == 0)
        {
            this->receiveMessage(Common::MsgTypeInfo, idleMsg.arg(ela).arg(currentLeftTime));
        }
        QTimer::singleShot(100, this, SLOT(showIdleMsg()));
    }
    else
    {
        this->receiveMessage(Common::MsgTypeInfo, "休息状态结束.");
    }
}

void EmailSender::receiveMessage(const int msgType, const QString &result)
{
    qDebug() << "EmailSender::receiveMessage msgType:" << QString::number(msgType).toUtf8() <<" msg:"<<result;
    switch (msgType)
    {
    case Common::MsgTypeError:
        m_failure_cnt++;
        m_receive_msg_cnt++;
        emit requestMsg(msgType, msg.arg(m_success_cnt).arg(m_failure_cnt).arg(result));
        break;
    case Common::MsgTypeSucc:
        m_success_cnt++;
        m_receive_msg_cnt++;
        emit requestMsg(msgType, msg.arg(m_success_cnt).arg(m_failure_cnt).arg(result));
        break;
    case Common::MsgTypeInfo:
    case Common::MsgTypeWarn:
    case Common::MsgTypeFinish:
    default:
        emit requestMsg(msgType, result);
        break;
    }
    if (m_total_cnt > 0 && m_receive_msg_cnt == m_total_cnt)
    { //全部处理完毕
        emit requestMsg(Common::MsgTypeEmailSendFinish, "处理完毕！");
    }
}

//停止发送email
void EmailSender::stop()
{

}
