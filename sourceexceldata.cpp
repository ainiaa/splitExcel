#include "sourceexceldata.h"

SourceExcelData::SourceExcelData(QObject *parent) : QObject(parent) {
}

void SourceExcelData::setDataSheetIndex(int dataSheetIndex) {
    this->dataSheetIndex = dataSheetIndex;
}
int SourceExcelData::getDataSheetIndex() {
    return this->dataSheetIndex;
}
void SourceExcelData::setEmailSheetIndex(int emailSheetIndex) {
    this->emailSheetIndex = emailSheetIndex;
}
int SourceExcelData::getEmailSheetIndex() {
    return this->emailSheetIndex;
}
void SourceExcelData::setSheetCnt(int sheetCnt) {
    this->sheetCnt = sheetCnt;
}
int SourceExcelData::getSheetCnt() {
    return this->sheetCnt;
}
void SourceExcelData::setGroupByText(QString groupByText) {
    this->groupByText = groupByText;
}
QString SourceExcelData::getGroupByText() {
    return this->groupByText;
}
void SourceExcelData::setDataSheetName(QString dataSheetName) {
    this->dataSheetName = dataSheetName;
}
QString SourceExcelData::getDataSheetName() {
    return this->dataSheetName;
}
void SourceExcelData::setEmailSheetName(QString emailSheetName) {
    this->emailSheetName = emailSheetName;
}
QString SourceExcelData::getEmailSheetName() {
    return this->emailSheetName;
}
void SourceExcelData::setSavePath(QString savePath) {
    this->savePath = savePath;
}
QString SourceExcelData::getSavePath() {
    return this->savePath;
}

void SourceExcelData::setSourcePath(QString sourcePath) {
    this->sourcePath = sourcePath;
}
QString SourceExcelData::getSourcePath() {
    return this->sourcePath;
}

void SourceExcelData::setOpType(int opType) {
    this->opType = opType;
}
int SourceExcelData::getOpType() {
    return this->opType;
}
