#ifndef XLSPARSER_H
#define XLSPARSER_H
#include <QDateTime>
#include <QDebug>
#include <QFileDialog>
#include <QObject>
#include <QThreadPool>
#include <QtCore/qmath.h>

#include "common.h"
#include "config.h"
#include "excelparserbylibrunnable.h"
#include "excelparserbyofficerunnable.h"
#include "officehelper.h"
#include "sourceexceldata.h"
#include "xlsxdocument.h"

class ExcelParser : public QObject {
    Q_OBJECT
    public:
    ExcelParser(QObject *parent = nullptr);
    ~ExcelParser();
    QString openFile(QWidget *dlgParent);
    void setSplitData(Config *cfg, QString groupByText, QString dataSheetName, QString emailSheetName, QString savePath);
    void setSplitData(Config *cfg, SourceExcelData *sourceExcelData);

    QHash<QString, QList<QStringList>> readEmailXls(QString groupByText, QString selectedSheetName);
    QHash<QString, QList<int>> readDataXls(QString groupByText, QString selectedSheetName);
    QHash<QString, QList<QStringList>> readXlsData(QString groupByText, QString selectedSheetName);
    QHash<QString, QList<QStringList>> getEmailData();
    void writeXlsByOffice(QString selectedSheetName, QHash<QString, QList<int>> qHash);
    void writeXlsByLib(QString selectedSheetName, QHash<QString, QList<int>> qHash);

    QStringList *getSheetHeader(QString selectedSheetName);

    bool selectSheet(const QString &name);
    QXlsx::CellRange dimension();
    QStringList getSheetNames();

    public slots:
    void receiveMessage(const int msgType, const QString &result);
    void doSplit();
    signals:
    void requestMsg(const int msgType, const QString &result);

    private:
    Config *cfg;
    QString sourcePath;
    QXlsx::Document *xlsx;
    QStringList *header = new QStringList();
    QString groupByText = nullptr;
    QString dataSheetName = nullptr;
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
    SourceExcelData *sourceExcelData = nullptr;
};
#endif // XLSPARSER_H
