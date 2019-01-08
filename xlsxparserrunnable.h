#ifndef XLSXPARSERRUNNABLE_H
#define XLSXPARSERRUNNABLE_H
#include <QObject>
#include <QRunnable>
#include <QMetaObject>
#include <QFileInfo>
#include "smtpclient/src/SmtpMime"
#include "common.h"
class XlsxParserRunnable : public QRunnable
{
public:
    XlsxParserRunnable(QObject *mParent);
    ~XlsxParserRunnable();

    void run();

    void setID(const int &id);

    void setSplitData(QString groupByText,QString dataSheetName, QString emailSheetName, QString savePath);

    void requestMsg(const int msgType, const QString &result);

private:
    //父对象
    QObject *mParent;
    int runnableID;
    QString groupByText;
    QString dataSheetName;
    QString emailSheetName;
    QString savePath;
};

#endif // XLSXPARSERRUNNABLE_H
