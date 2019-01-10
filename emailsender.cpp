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

    int maxProcessCntPreThread = qCeil(emailQhash.size() * 1.0 / maxThreadCnt);
    qDebug() << "emailQhash.size() :" << emailQhash.size()  << " maxThreadCnt:" <<maxThreadCnt <<" maxProcessCntPreThread:" << maxProcessCntPreThread;
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
