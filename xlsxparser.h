#ifndef XLSPARSER_H
#define XLSPARSER_H
#include <QObject>
#include <QFileDialog>
#include <QDateTime>
#include <QDebug>

#include "common.h"
#include "xlsxdocument.h"

class XlsxParser : public QObject
{
    Q_OBJECT
public:
    XlsxParser(QObject* parent = nullptr);
    ~XlsxParser();
    QString openFile(QWidget *dlgParent);
    void setSplitData(QString groupByText,QString dataSheetName, QString emailSheetName, QString savePath);

    QHash<QString, QList<QStringList>> readEmailXls(QString groupByText, QString selectedSheetName, bool isEmail);
    QHash<QString, QList<int>> readDataXls(QString groupByText, QString selectedSheetName);
    QHash<QString, QList<QStringList>> getEmailData();
    void writeXls(QString selectedSheetName,QHash<QString, QList<int>> qHash, QString savePath);
    void sendemail(QHash<QString, QList<QStringList>> qHash, QString savePath);
    void writeXlsHeader(QXlsx::Document *currXls,QString selectedSheetName);

    QStringList* getSheetHeader(QString selectedSheetName);

    bool selectSheet(const QString &name);
    QXlsx::CellRange dimension();
    QStringList getSheetNames();


    void convertToColName(int data, QString &res);
    QString to26AlphabetString(int data);
public slots:
    void receiveMessage(const int msgType, const QString &result);
    void doSplit();
signals:
   void requestMsg(const int msgType, const QString &result);
private:
    QObject *mParent;
    QXlsx::Document *xlsx = nullptr;
    QStringList *header = new QStringList();
    QString groupByText;
    QString dataSheetName;
    QString emailSheetName;
    QString savePath;
    QHash<QString, QList<QStringList>> emailQhash;
};
#endif // XLSPARSER_H
