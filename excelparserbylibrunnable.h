#ifndef XLSXPARSERRUNNABLE_H
#define XLSXPARSERRUNNABLE_H
#include <QAxObject>
#include <QDateTime>
#include <QDir>
#include <QFileInfo>
#include <QMetaObject>
#include <QObject>
#include <QRunnable>

#include "common.h"
#include "config.h"
#include "excelbase.h"
#include "qt_windows.h"
#include "xlsxcelllocation.h"
#include "xlsxdocument.h"
#include "xlsxformat.h"
#include "xlsxformat_p.h"
#include "xlsxworkbook.h"

class ExcelParserByLibRunnable : public QRunnable {
    public:
    ExcelParserByLibRunnable(QObject *mParent);
    ~ExcelParserByLibRunnable();

    void run();

    void setID(const int &id);

    void setSplitData(QString sourcePath, QString selectedSheetName, QHash<QString, QList<int>> fragmentDataQhash, QString savePath, int m_total);

    void requestMsg(const int msgType, const QString &result);

    void writeXls(QString selectedSheetName, QHash<QString, QList<int>> qHash, QString savePath);
    void writeXlsHeader(QXlsx::Document *currXls, QString selectedSheetName);

    QXlsx::Format copyFormat(QXlsx::Format sourceCellFormat);
    void convertToColName(int data, QString &res);
    QString to26AlphabetString(int data);

    void processByOffice(QString key, QList<int> contentList);
    void processByQxls(QString key, QList<int> contentList);
    bool isInstalledExcelApp();
    void processSourceFile();

    bool copyFileToPath(QString sourceDir, QString toDir, bool coverFileIfExist);
    void generateTplXls();

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
};

#endif // XLSXPARSERRUNNABLE_H
