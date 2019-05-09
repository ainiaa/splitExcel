#include "sourcexmldata.h"

SourceXmlData::SourceXmlData(QObject *parent) : QObject(parent) {
}

void SourceXmlData::setDataSheetIndex(int dataSheetIndex) {
    this->dataSheetIndex = dataSheetIndex;
}
int SourceXmlData::getDataSheetIndex() {
    return this->dataSheetIndex;
}
void SourceXmlData::setEmailSheetIndex(int emailSheetIndex) {
    this->emailSheetIndex = emailSheetIndex;
}
int SourceXmlData::getEmailSheetIndex() {
    return this->emailSheetIndex;
}
void SourceXmlData::setSheetCnt(int sheetCnt) {
    this->sheetCnt = sheetCnt;
}
int SourceXmlData::getSheetCnt() {
    return this->sheetCnt;
}
void SourceXmlData::setGroupByText(QString groupByText) {
    this->groupByText = groupByText;
}
QString SourceXmlData::getGroupByText() {
    return this->groupByText;
}
void SourceXmlData::setDataSheetName(QString dataSheetName) {
    this->dataSheetName = dataSheetName;
}
QString SourceXmlData::getDataSheetName() {
    return this->dataSheetName;
}
void SourceXmlData::setEmailSheetName(QString emailSheetName) {
    this->emailSheetName = emailSheetName;
}
QString SourceXmlData::getEmailSheetName() {
    return this->emailSheetName;
}
void SourceXmlData::setSavePath(QString savePath) {
    this->savePath = savePath;
}
QString SourceXmlData::getSavePath() {
    return this->savePath;
}

void SourceXmlData::setSourcePath(QString sourcePath) {
    this->sourcePath = sourcePath;
}
QString SourceXmlData::getSourcePath() {
    return this->sourcePath;
}
