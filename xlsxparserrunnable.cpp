#include "xlsxparserrunnable.h"

XlsxParserRunnable::XlsxParserRunnable(QObject *parent)
{
    this->mParent = parent;
}

XlsxParserRunnable::~XlsxParserRunnable()
{
    runnableID = 0;
}

void XlsxParserRunnable::setID(const int &id)
{
    runnableID = id;
}

void XlsxParserRunnable::setSplitData(QString groupByText,QString dataSheetName, QString emailSheetName, QString savePath)
{
    qDebug("setSplitData") ;
    this->groupByText = groupByText;
    this->dataSheetName = dataSheetName;
    this->emailSheetName = emailSheetName;
    this->savePath = savePath;
}

void XlsxParserRunnable::run()
{
    qDebug("XlsxParserRunnable::run start") ;
    requestMsg(Common::MsgTypeInfo, "XlsxParserRunnable::run");

    qDebug("EmailSenderRunnable::run end") ;
}

void XlsxParserRunnable::requestMsg(const int msgType, const QString &msg)
{
    QMetaObject::invokeMethod(mParent, "receiveMessage", Qt::QueuedConnection, Q_ARG(int,msgType),Q_ARG(QString, msg));
}

