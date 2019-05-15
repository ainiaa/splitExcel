#include "adoconnection.h"
#include "ado.h"
#include <QTimer>
#include <QtDebug>

AdoConnection::AdoConnection(QObject *parent) : QObject(parent) {

    timer = new QTimer(this);
    Q_CHECK_PTR(timer);
    connect(timer, SIGNAL(timeout()), this, SLOT(disconnect()));

    object = new QAxObject(this);
    object->setControl("ADODB.Connection");        /*创建ADODB.Connection对象*/
    object->setProperty("ConnectionTimeout", 300); /*设置超时时间确保连接成功*/
    connect(object, SIGNAL(exception(int, const QString &, const QString &, const QString &)), this,
            SLOT(exception(int, const QString &, const QString &, const QString &)));
}

void AdoConnection::exception(int code, const QString &source, const QString &desc, const QString &help) {
    /*输出异常或错误信息*/
    qDebug() << "Code:       " << code;
    qDebug() << "Source:     " << source;
    qDebug() << "Description:" << desc;
    qDebug() << "Help:       " << help;
}
bool AdoConnection::open(const QString &connectString) {
    if (isOpen())
        return true;

    openString = connectString; /*设置连接字符串*/
    HRESULT hr = object->dynamicCall("Open(QString,QString,QString,int)", connectString, "", "", adConnectUnspecified).toInt(); /*连接到数据库*/
    return SUCCEEDED(hr);
}

bool AdoConnection::open() {
    if (openString.isEmpty())
        return false;

    bool ret = open(openString);
    if (timer && timer->isActive())
        timer->stop();

    return ret;
}

bool AdoConnection::execute(const QString &sql) {
    if (!open())
        return false;

    /*执行SQL语句*/
    HRESULT hr = object->dynamicCall("Execute(QString)", sql).toInt();

    return SUCCEEDED(hr);
}

void AdoConnection::disconnect() {
    if (isOpen())
        object->dynamicCall("Close"); /*关闭数据库连接*/

    if (timer)
        timer->stop();
    qDebug() << "AdoConnection::disconnect()";
}

void AdoConnection::close() {
    if (timer) {
        if (timer->isActive())
            timer->stop();
        timer->start(5000);
    } else
        disconnect();
}

QVariant AdoConnection::connection() {
    return object->asVariant();
}

bool AdoConnection::isOpen() const {
    return (bool)(object->property("State").toInt() != adStateClosed);
}
