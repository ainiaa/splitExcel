#include "excelparserbyado.h"

ExcelParserByADO::ExcelParserByADO(QObject *parent) : QObject(parent) {
}

void ExcelParserByADO::parse() {
    QAxObject *conn = new QAxObject(this);
    conn->setControl("ADODB.Connection"); /*创建ADODB.Connection对象*/
    QString dns = "provider=Microsoft.ACE.OLEDB.15.0;extended properties='excel 15.0;hdr=yes;imex=1';data source=%1";

    QString path = "d:/tmp/test.xlsx";
    path = QDir::toNativeSeparators(path);
    //    conn->dynamicCall("Open(QString dns)", dns.arg(path));
    //    QString stringSql = "select * from [%1$] where %2 = '%3'";
    //    QString dateSql = "select * from [%1$] where %2 = #%3#";
    //    QString sql = stringSql.arg("data").arg("用工模式").arg("风澜");

    AdoConnection *adoConnection = new AdoConnection;
    QString openString = dns.arg(path);

    if (!adoConnection->open(openString)) {
        return;
    }
    QString sql = QString("select * from %1").arg("data");
    AdoRecordset recordset(adoConnection);
    recordset.open(sql);

    while (recordset.next()) {
        int count = recordset.fieldCount();
        for (int i = 0; i < count; i++) {
            qDebug() << recordset.fieldName(i) << recordset.fieldValue(i);
        }
    }

    adoConnection->close();
    qDebug() << "close";
}
