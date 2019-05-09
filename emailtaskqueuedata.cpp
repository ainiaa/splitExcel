#include "emailtaskqueuedata.h"

EmailTaskQueueData::EmailTaskQueueData() {
}

void EmailTaskQueueData::setMsg(QString msg) {
    this->msg = msg;
}
QString EmailTaskQueueData::getMsg() {
    return this->msg;
}
void EmailTaskQueueData::setEmailQhash(QHash<QString, QList<QStringList>> emailQhash) {
    this->emailQhash = emailQhash;
}

QHash<QString, QList<QStringList>> EmailTaskQueueData::getEmailQhash() {
    return this->emailQhash;
}
void EmailTaskQueueData::setSavePath(QString savePath) {
    this->savePath = savePath;
}
QString EmailTaskQueueData::getSavePath() {
    return this->savePath;
}
void EmailTaskQueueData::setTotal(int total) {
    this->total = total;
}
int EmailTaskQueueData::getTotal() {
    return this->total;
}
