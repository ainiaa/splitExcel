#ifndef EMAILTASKQUEUEDATA_H
#define EMAILTASKQUEUEDATA_H
#include <QHash>
#include <QString>

class EmailTaskQueueData {
    public:
    EmailTaskQueueData();

    void setMsg(QString msg);
    QString getMsg();
    void setEmailQhash(QHash<QString, QList<QStringList>> emailQhash);
    QHash<QString, QList<QStringList>> getEmailQhash();
    void setSavePath(QString savePath);
    QString getSavePath();
    void setTotal(int total);
    int getTotal();

    private:
    QString msg;
    QHash<QString, QList<QStringList>> emailQhash;
    QString savePath;
    int total;
};

#endif // EMAILTASKQUEUEDATA_H
