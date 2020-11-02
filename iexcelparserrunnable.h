#ifndef IEXCELPARSERRUNNABLE_H
#define IEXCELPARSERRUNNABLE_H

#include "sourceExceldata.h"
#include <QObject>
#include <QRunnable>

class IExcelParserRunnable : public QRunnable {
    public:
    virtual void setID(const int &id) = 0;

    virtual void
    setSplitData(QString sourcePath, QString selectedSheetName, QHash<QString, QList<int>> fragmentDataQhash, QString savePath, int m_total) = 0;
    virtual void
    setSplitData(SourceData *sourceXmlData, QString selectedSheetName, QHash<QString, QList<int>> fragmentDataQhash, int m_total) = 0;

    virtual void
    setSplitData(SourceData *sourceXmlData, QString selectedSheetName, QHash<QString, QList<QList<QVariant>>> fragmentDataQhash, int m_total) = 0;

    virtual void requestMsg(const int msgType, const QString &result) = 0;
};

#endif // IEXCELPARSERRUNNABLE_H
