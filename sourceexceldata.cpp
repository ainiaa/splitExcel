#include "sourceexceldata.h"

SourceData::SourceData(QObject *parent) : QObject(parent) {
}

void SourceData::setDataSheetIndex(int dataSheetIndex) {
    this->dataSheetIndex = dataSheetIndex;
}
int SourceData::getDataSheetIndex() {
    return this->dataSheetIndex;
}
void SourceData::setEmailSheetIndex(int emailSheetIndex) {
    this->emailSheetIndex = emailSheetIndex;
}
int SourceData::getEmailSheetIndex() {
    return this->emailSheetIndex;
}
void SourceData::setSheetCnt(int sheetCnt) {
    this->sheetCnt = sheetCnt;
}
int SourceData::getSheetCnt() {
    return this->sheetCnt;
}
void SourceData::setGroupByText(QString groupByText) {
    this->groupByText = groupByText;
}
QString SourceData::getGroupByText() {
    return this->groupByText;
}

void SourceData::setGroupByIndex(int groupByIndex) {
    this->groupByIndex = groupByIndex;
}

int SourceData::getGroupByIndex() {
    return this->groupByIndex;
}

void SourceData::setDataSheetName(QString dataSheetName) {
    this->dataSheetName = dataSheetName;
}
QString SourceData::getDataSheetName() {
    return this->dataSheetName;
}
void SourceData::setEmailSheetName(QString emailSheetName) {
    this->emailSheetName = emailSheetName;
}
QString SourceData::getEmailSheetName() {
    return this->emailSheetName;
}
void SourceData::setSavePath(QString savePath) {
    this->savePath = savePath;
}
QString SourceData::getSavePath() {
    return this->savePath;
}

void SourceData::setSourcePath(QString sourcePath) {
    this->sourcePath = sourcePath;
}
QString SourceData::getSourcePath() {
    return this->sourcePath;
}

void SourceData::setOpType(int opType) {
    this->opType = opType;
}
int SourceData::getOpType() {
    return this->opType;
}

void SourceData::setSourceRowStart(int sourceRowStart) {
    this->sourceRowStart = sourceRowStart;
}

int SourceData::getSourceRowStart() {
    return this->sourceRowStart;
}
void SourceData::setSourceColStart(int sourceColStart) {
    this->sourceColStart = sourceColStart;
}
int SourceData::setSourceColStart() {
    return this->sourceColStart;
}
void SourceData::setSourceRowCnt(int sourceRowCnt) {
    this->sourceRowCnt = sourceRowCnt;
}
int SourceData::getSourceRowCnt() {
    return this->sourceRowCnt;
}
void SourceData::setSourceColCnt(int sourceColCnt) {
    this->sourceColCnt = sourceColCnt;
}
int SourceData::getSourceColCnt() {
    return this->sourceColCnt;
}
void SourceData::setSourceMinAlphabetCol(QString sourceMinAlphabetCol) {
    this->sourceMinAlphabetCol = sourceMinAlphabetCol;
}
QString SourceData::getSourceMinAlphabetCol() {
    return this->sourceMinAlphabetCol;
}
void SourceData::setSourceMaxAlphabetCol(QString sourceMaxAlphabetCol) {
    this->sourceMaxAlphabetCol = sourceMaxAlphabetCol;
}

QString SourceData::getSourceMaxAlphabetCol() {
    return this->sourceMaxAlphabetCol;
}
void SourceData::setTplXlsPath(QString tplXlsPath) {
    this->tplXlsPath = tplXlsPath;
}

QString SourceData::getTplXlsPath() {
    return this->tplXlsPath;
}
