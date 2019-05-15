#include "adorecordset.h"
#include "ado.h"
#include "adoconnection.h"
AdoRecordset::AdoRecordset(const QVariant &connection, QObject *parent) : QObject(parent), dbConnection(connection) {
    object = new QAxObject(this);
    object->setControl("ADODB.Recordset");
}

AdoRecordset::AdoRecordset(AdoConnection *adoConnection, QObject *parent)
: QObject(parent), dbConnection(adoConnection->connection()) {
    object = new QAxObject(this);
    object->setControl("ADODB.Recordset");
}

bool AdoRecordset::open(const QString &sql) {
    HRESULT hr =
    object->dynamicCall("Open(QString,QVariant,int,int,int)", sql, dbConnection, adOpenStatic, adLockOptimistic, adCmdText).toInt();
    initial = true;
    return SUCCEEDED(hr);
}

int AdoRecordset::recordCount() const {
    return object->property("RecordCount").toUInt();
}

bool AdoRecordset::next() {
    bool ret = false;

    fieldNames.clear();
    fieldValues.clear();
    if (object->property("RecordCount").toInt() < 1)
        return false;

    if (initial) {
        ret = move(First);
        initial = false;
    } else
        ret = move(Next);
    if (object->property("EOF").toBool())
        return false;

    QAxObject *adoFields = object->querySubObject("Fields");
    if (adoFields) {
        int count = adoFields->property("Count").toInt();
        for (int i = 0; i < count; i++) {
            QAxObject *adoField = adoFields->querySubObject("Item(int)", i);
            if (adoField) {
                fieldNames += adoField->property("Name").toString();
                fieldValues += adoField->property("Value");
                ADO_DELETE(adoField);
            }
        }
        ADO_DELETE(adoFields);
    }
    return ret;
}

bool AdoRecordset::move(MoveAction action) {
    static const char *actions[] = { "MoveFirst(void)", "MoveNext(void)", "MovePrevious(void)", "MoveLast(void)" };

    HRESULT hr = object->dynamicCall(actions[action]).toInt();

    return SUCCEEDED(hr);
}

int AdoRecordset::fieldCount() const {
    return fieldNames.count();
}

QString AdoRecordset::fieldName(int index) const {
    if (index < 0 || index >= fieldNames.count())
        return QString::null;

    return fieldNames[index];
}

QVariant AdoRecordset::fieldValue(int index) const {
    if (index < 0 || index >= fieldValues.count())
        return QVariant();

    return fieldValues[index];
}

void AdoRecordset::close() {
    object->dynamicCall("Close");
}
