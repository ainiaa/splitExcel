#ifndef XLSXPARSERBYOFFICE_H
#define XLSXPARSERBYOFFICE_H

#include <QObject>
#include <QFileDialog>
#include <QDateTime>
#include <QDebug>
#include <QThreadPool>
#include <QtCore/qmath.h>

#include "common.h"
#include "xlsxdocument.h"
#include "config.h"
#include "xlsxparserrunnable.h"
class XlsxParserByOffice : public QObject
{
Q_OBJECT
public:
XlsxParserByOffice(QObject *parent = nullptr);
~XlsxParserByOffice();
QString openFile(QWidget *dlgParent);
void setSplitData(Config *cfg, QString groupByText, QString dataSheetName, QString emailSheetName, QString savePath);

QHash<QString, QList<QStringList>> readEmailXls(QString groupByText, QString selectedSheetName);
QHash<QString, QList<int>> readDataXls(QString groupByText, QString selectedSheetName);
QHash<QString, QList<QStringList>> readXlsData(QString groupByText, QString selectedSheetName);
QHash<QString, QList<QStringList>> getEmailData();
void writeXls(QString selectedSheetName,QHash<QString, QList<int>> qHash, QString savePath);

QStringList* getSheetHeader(QString selectedSheetName);

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
    QXlsx::Document* xlsx;
    QStringList *header = new QStringList();
    QString groupByText;
    QString dataSheetName;
    QString emailSheetName;
    QString savePath;
    QHash<QString, QList<QStringList>> emailQhash;

    QString msg;
    int m_total_cnt;
    int m_process_cnt;
    int m_success_cnt;
    int m_failure_cnt;
    int m_receive_msg_cnt;
};

#endif // XLSXPARSERBYOFFICE_H
