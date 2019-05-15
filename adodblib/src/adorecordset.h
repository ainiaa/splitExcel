#ifndef ADORECORDSET_H
#define ADORECORDSET_H

#include <QObject>
#include <QVariant>

class QAxObject;

class AdoConnection;

class AdoRecordset : public QObject {
    Q_OBJECT
    public:
    enum MoveAction { First = 0, Next, Previous, Last };
    AdoRecordset(const QVariant &dbConnection, QObject *parent = nullptr);
    AdoRecordset(AdoConnection *adoConnection, QObject *parent = nullptr);
    bool open(const QString &sql);
    void close();
    bool next();

    int recordCount() const;
    int fieldCount() const;
    QString fieldName(int index) const;
    QVariant fieldValue(int index) const;

    protected:
    bool move(MoveAction action);

    private:
    bool initial;
    QVariant dbConnection;
    QList<QString> fieldNames;
    QList<QVariant> fieldValues;
    QAxObject *object;
};
#endif // ADORECORDSET_H
