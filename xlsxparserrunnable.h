#ifndef XLSXPARSERRUNNABLE_H
#define XLSXPARSERRUNNABLE_H
#include <QObject>
#include <QRunnable>
#include <QMetaObject>
#include <QFileInfo>
#include <QDateTime>

#include "xlsxdocument.h"
#include "xlsxworkbook.h"
#include "common.h"
#include "config.h"
#include "xlsxcelllocation.h"
#include "xlsxformat.h"
#include "xlsxformat_p.h"

class XlsxParserRunnable : public QRunnable
{
public:
    XlsxParserRunnable(QObject *mParent);
    ~XlsxParserRunnable();

    void run();

    void setID(const int &id);

    void setSplitData(QXlsx::Document *xlsx, QString selectedSheetName,QString key, QList<int> contentList, QString savePath,int m_total);

    void requestMsg(const int msgType, const QString &result);

    void writeXls(QString selectedSheetName,QHash<QString, QList<int>> qHash, QString savePath);
    void writeXlsHeader(QXlsx::Document *currXls,QString selectedSheetName);

    QXlsx::Format copyFormat(QXlsx::Format sourceCellFormat);
    void convertToColName(int data, QString &res);
    QString to26AlphabetString(int data);

private:
    //父对象
    QObject *mParent;
    int runnableID;
    int m_total;
    QString key;
    QList<int> contentList;
    QString savePath;
    QXlsx::Document *xlsx;
    QString selectedSheetName;
};

#endif // XLSXPARSERRUNNABLE_H
