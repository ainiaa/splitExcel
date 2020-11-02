#ifndef EXCELPARSER_H
#define EXCELPARSER_H
#include "common.h"
#include "config.h"
#include "parserbylibrunnable.h"
#include "parserbymsrunnable.h"
#include "officehelper.h"
#include "sourceexceldata.h"
#include "xlsxdocument.h"
#include <QDateTime>
#include <QDebug>
#include <QFileDialog>
#include <QObject>
#include <QThreadPool>
#include <QtCore/qmath.h>

class ExcelParser : public QObject {
    Q_OBJECT
    public:
    ExcelParser(QObject *parent = nullptr);
    ~ExcelParser();
    QString openFile(QWidget *dlgParent);
    void setSplitData(Config *cfg, SourceData *sourceData);

    QHash<QString, QList<QStringList>> readEmailXls(QString groupByText, QString selectedSheetName);
    QHash<QString, QList<int>> readDataXls(QString groupByText, QString selectedSheetName);
    QHash<QString, QList<QStringList>> getEmailData();
    void writeXlsByMS(QString selectedSheetName, QHash<QString, QList<int>> qHash);
    void writeXlsByMS(QString selectedSheetName, QHash<QString, QList<QList<QVariant>>> qHash);
    void writeXlsByLib(QString selectedSheetName, QHash<QString, QList<int>> qHash);

    QStringList *getSheetHeader(QString selectedSheetName);

    bool selectSheet(const QString &name);
    QXlsx::CellRange dimension();
    QStringList getSheetNames();

    QHash<QString, QList<QList<QVariant>>> readDataXls(int groupByIndex, int selectedSheetIndex);

    void freeExcel(QAxObject *excel);
    QVariant readAll();
    void readAll(QList<QList<QVariant> > &cells);
    void castListListVariant2Variant(const QList<QList<QVariant> > &cells, QVariant &res);
    void castVariant2ListListVariant(const QVariant &var, QList<QList<QVariant> > &res);

    public slots:
    void receiveMessage(const int msgType, const QString &result);
    void doParse();
    signals:
    void requestMsg(const int msgType, const QString &result);

    private:
    Config *cfg;
    QString sourcePath;
    QXlsx::Document *xlsx;
    QStringList *header = new QStringList();
    QString groupByText = nullptr;
    QString dataSheetName = nullptr;
    int dataSheetIndex = 0;
    int groupByIndex = 0;
    QString emailSheetName = nullptr;
    QString savePath = nullptr;
    QHash<QString, QList<QStringList>> emailQhash;

    QString msg;
    int m_total_cnt;
    int m_process_cnt;
    int m_success_cnt;
    int m_failure_cnt;
    int m_receive_msg_cnt;

    bool isInstalledOffice;
    SourceData *sourceData = nullptr;
};
#endif // EXCELPARSER_H
