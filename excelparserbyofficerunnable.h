#ifndef PARSERBYOFFICER_H
#define PARSERBYOFFICER_H

#include <QAxObject>
#include <QDateTime>
#include <QDir>
#include <QFileInfo>
#include <QMetaObject>
#include <QObject>

#include "common.h"
#include "config.h"
#include "excelbase.h"
#include "excelparser.h"
#include "iexcelparserrunnable.h"
#include "officehelper.h"
#include "sourceexceldata.h"

class ExcelParserByOfficeRunnable : public IExcelParserRunnable {
    public:
    ExcelParserByOfficeRunnable(QObject *mParent);
    ~ExcelParserByOfficeRunnable() override;

    void run() override;
    void setID(const int &id) override;
    void setSplitData(SourceExcelData *sourceXmlData, QString selectedSheetName, QHash<QString, QList<int>> fragmentDataQhash, int m_total) override;
    void setSplitData(QString sourcePath, QString selectedSheetName, QHash<QString, QList<int>> fragmentDataQhash, QString savePath, int m_total) override;
    void requestMsg(const int msgType, const QString &result) override;
    void writeXls(QString selectedSheetName, QHash<QString, QList<int>> qHash, QString savePath);
    void processByOffice(QString key, QList<int> contentList);
    void processSourceFile();
    bool static copyFileToPath(QString sourceDir, QString toDir, bool coverFileIfExist);
    void generateTplXls();
    void static processSourceFile(SourceExcelData *sourceExcelData, QString selectedSheetName);
    void static generateTplXls(SourceExcelData *sourceExcelData, int selectedSheetIndex);
    void static generateTplXls(SourceExcelData *sourceExcelData, QString selectedSheetName);
    void static freeExcel(QAxObject *excel);

    private:
    //父对象
    QString sourcePath;
    QObject *mParent;
    int runnableID;
    int m_total;
    QString key;
    QString savePath;
    QString selectedSheetName;
    QHash<QString, QList<int>> fragmentDataQhash;
    QXlsx::Document *xlsx;
    bool installedExcelApp;

    int sourceRowStart;           // 起始行数
    int sourceColStart;           // 起始列数
    QString sourceMinAlphabetCol; //最小列（字符）
    QString sourceMaxAlphabetCol; //最大列（字符）
    int sourceRowCnt;             //最大行数
    int sourceColCnt;             //最大列数
    int sourceWorkSheetCnt;       // sheets数量
    int selectedSheetIndex;       //选中索引
    QString tplXlsPath;
    QString deleteRangeFormat;

    SourceExcelData *sourceXmlData;
};

#endif // PARSERBYOFFICER_H
