#ifndef SOURCEXMLDATA_H
#define SOURCEXMLDATA_H

#include <QObject>

class SourceXmlData : public QObject {
    Q_OBJECT
    public:
    SourceXmlData(QObject *parent = nullptr);

    void setDataSheetIndex(int dataSheetIndex);
    int getDataSheetIndex();
    void setEmailSheetIndex(int emailSheetIndex);
    int getEmailSheetIndex();
    void setSheetCnt(int sheetCnt);
    int getSheetCnt();
    void setGroupByText(QString groupByText);
    QString getGroupByText();
    void setDataSheetName(QString dataSheetName);
    QString getDataSheetName();
    void setEmailSheetName(QString emailSheetName);
    QString getEmailSheetName();
    void setSavePath(QString savePath);
    QString getSavePath();
    void setSourcePath(QString sourcePath);
    QString getSourcePath();

    private:
    int dataSheetIndex;
    int emailSheetIndex;
    int sheetCnt;
    QString groupByText;
    QString dataSheetName;
    QString emailSheetName;
    QString savePath;
    QString sourcePath;
};

#endif // SOURCEXMLDATA_H
