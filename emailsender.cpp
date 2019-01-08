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
    while (it.hasNext()) {
        it.next();
        MimeMessage mineMsg;

        //防止中文乱码
        mineMsg.setHeaderEncoding(MimePart::Encoding::Base64);
        mineMsg.setSender(new EmailAddress(defaultSender, defaultSender));

        QString key = it.key();
        QList<QStringList> content = it.value();

        EmailSenderRunnable *runnable = new EmailSenderRunnable(this);
        runnable->setID(++m_process_cnt);
        runnable->setSendData(user,password,server,defaultSender,savePath,key,content);
        runnable->setAutoDelete(true);

        pool.start(runnable);
        if (m_process_cnt % maxThreadCnt == 0)
        {
            pool.waitForDone();
        }
    }
    if (pool.activeThreadCount() > 0)
    {
        pool.waitForDone();
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
