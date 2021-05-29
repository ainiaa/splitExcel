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

void SourceExcelData::setGroupByIndex(int groupByIndex) {
    this->groupByIndex = groupByIndex;
}

int SourceExcelData::getGroupByIndex() {
    return this->groupByIndex;
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

void SourceExcelData::setSourceRowStart(int sourceRowStart) {
    this->sourceRowStart = sourceRowStart;
}

int SourceExcelData::getSourceRowStart() {
    return this->sourceRowStart;
}
void SourceExcelData::setSourceColStart(int sourceColStart) {
    this->sourceColStart = sourceColStart;
}
int SourceExcelData::setSourceColStart() {
    return this->sourceColStart;
}
void SourceExcelData::setSourceRowCnt(int sourceRowCnt) {
    this->sourceRowCnt = sourceRowCnt;
}
int SourceExcelData::getSourceRowCnt() {
    return this->sourceRowCnt;
}
void SourceExcelData::setSourceColCnt(int sourceColCnt) {
    this->sourceColCnt = sourceColCnt;
}
int SourceExcelData::getSourceColCnt() {
    return this->sourceColCnt;
}
void SourceExcelData::setSourceMinAlphabetCol(QString sourceMinAlphabetCol) {
    this->sourceMinAlphabetCol = sourceMinAlphabetCol;
}
QString SourceExcelData::getSourceMinAlphabetCol() {
    return this->sourceMinAlphabetCol;
}
void SourceExcelData::setSourceMaxAlphabetCol(QString sourceMaxAlphabetCol) {
    this->sourceMaxAlphabetCol = sourceMaxAlphabetCol;
}

QString SourceExcelData::getSourceMaxAlphabetCol() {
    return this->sourceMaxAlphabetCol;
}
void SourceExcelData::setTplXlsPath(QString tplXlsPath) {
    this->tplXlsPath = tplXlsPath;
}

QString SourceExcelData::getTplXlsPath() {
    return this->tplXlsPath;
}

QHash<QString, QString> SourceExcelData::getPasswordData() {
    return this->passwordData;
}
void SourceExcelData::setPasswordData(QHash<QString, QString> passwordData) {
    this->passwordData = passwordData;
}

void SourceExcelData::setNeedPassword(bool needPassword){
    this->needPassword = needPassword;
}
bool SourceExcelData::getNeedPassword() {
    return this->needPassword;
}

void SourceExcelData::setPasswordDataSheetName(QString passwordDataSheetName) {
    this->passwordDataSheetName = passwordDataSheetName;
}
QString SourceExcelData::getPasswordDataSheetName(){
    return this->passwordDataSheetName;
}
