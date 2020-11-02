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

class ParserByMSRunnable : public IExcelParserRunnable {
    public:
    ParserByMSRunnable(QObject *mParent);
    ~ParserByMSRunnable() override;

    void run() override;
    void setID(const int &id) override;
    void setSplitData(SourceData *sourceExcelData, QString selectedSheetName, QHash<QString, QList<int>> fragmentDataQhash, int m_total) override;
    void setSplitData(QString sourcePath, QString selectedSheetName, QHash<QString, QList<int>> fragmentDataQhash, QString savePath, int m_total) override;
    void setSplitData(SourceData *sourceXmlData, QString selectedSheetName, QHash<QString, QList<QList<QVariant>>> fragmentDataQhash, int m_total) override;
    void requestMsg(const int msgType, const QString &result) override;
    void writeXls(QString selectedSheetName, QHash<QString, QList<int>> qHash, QString savePath);
    void doProcess(QString key, QList<int> contentList);
    void doProcess(QString key, QList<QList<QVariant>> contentList);
    void doFilter(QString key);
    void processSourceFile();
    bool static copyFileToPath(QString sourceDir, QString toDir, bool coverFileIfExist);
    void generateTplXls();
    void static processSourceFile(SourceData *sourceExcelData, QString selectedSheetName);
    void static sortXls();
    void static generateTplXls(SourceData *sourceExcelData, int selectedSheetIndex);
    void static generateTplXls(SourceData *sourceExcelData, QString selectedSheetName);
    void static freeExcel(QAxObject *excel);

    private:
    //父对象
    QString sourcePath; //源Excel文件路径
    QObject *mParent;
    int runnableID; //当前ID
    int m_total;    //
    QString groupByText;
    QString key;
    QString savePath;          //拆分excel保存路径
    QString selectedSheetName; // data sheet 名称
    QHash<QString, QList<int>> fragmentDataQhash;
    QHash<QString, QList<QList<QVariant>>> fragmentDataQhash2;
    int dataType =0;
    QXlsx::Document *xlsx;
    bool installedExcelApp;

    int groupByIndex;             //排序列索引
    int sourceRowStart;           // 起始行数
    int sourceColStart;           // 起始列数
    QString sourceMinAlphabetCol; //最小列（字符）
    QString sourceMaxAlphabetCol; //最大列（字符）
    int sourceRowCnt;             //最大行数
    int sourceColCnt;             //最大列数
    int sourceWorkSheetCnt;       // sheets数量
    int selectedSheetIndex;       //选中索引
    QString tplXlsPath;           //模板路径
    QString deleteRangeFormat;

    SourceData *sourceExcelData; //原文件数据

    void run0();
    void run1();
};

#endif // PARSERBYOFFICER_H
