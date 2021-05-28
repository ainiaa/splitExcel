#include "excelparser.h"

ExcelParser::ExcelParser(QObject *parent) : QObject(parent) {
    m_success_cnt = 0;
    m_failure_cnt = 0;
    m_process_cnt = 0;
    m_receive_msg_cnt = 0;
    isInstalledOffice = OfficeHelper::isInstalledExcelApp();
}

ExcelParser::~ExcelParser() {
    qDebug() << "~XlsxParser start";

    qDebug() << "~XlsxParser end";
}
void ExcelParser::setSplitData(Config *cfg, QString groupByText, QString dataSheetName, QString emailSheetName, QString passwordDataSheetName, QString savePath) {
    this->cfg = cfg;
    this->groupByText = groupByText;
    this->dataSheetName = dataSheetName;
    this->emailSheetName = emailSheetName;
    this->passwordDataSheetName = passwordDataSheetName;
    this->savePath = savePath;
}

void ExcelParser::setSplitData(Config *cfg, SourceExcelData *sourceExcelData) {
    this->cfg = cfg;
    this->groupByText = sourceExcelData->getGroupByText();
    this->dataSheetName = sourceExcelData->getDataSheetName();
    this->emailSheetName = sourceExcelData->getEmailSheetName();
    this->passwordDataSheetName = sourceExcelData->getPasswordDataSheetName();
    this->savePath = sourceExcelData->getSavePath();
    this->sourceExcelData = sourceExcelData;
}
QStringList ExcelParser::getSheetNames() {
    QStringList sheetNames;
    if (nullptr != xlsx) {
        sheetNames = xlsx->sheetNames();
    }
    return sheetNames;
}

bool ExcelParser::selectSheet(const QString &name) {
    return xlsx->selectSheet(name);
}

QXlsx::CellRange ExcelParser::dimension() {
    return xlsx->dimension();
}

QStringList *ExcelParser::getSheetHeader(QString selectedSheetName) {
    QStringList *currentHeader = new QStringList();
    QXlsx::CellRange range;
    xlsx->selectSheet(selectedSheetName);
    range = xlsx->dimension();
    int colCount = range.columnCount();

    for (int colum = 1; colum <= colCount; ++colum) {
        QXlsx::Cell *cell = xlsx->cellAt(1, colum);
        if (cell) {
            currentHeader->append(cell->value().toString());
        }
    }
    return currentHeader;
}

//打开文件对话框
QString ExcelParser::openFile(QWidget *dlgParent) {
    QString path = QFileDialog::getOpenFileName(dlgParent, tr("Open Excel"), ".", tr("excel(*.xlsx)"));
    if (path.length() == 0) {
        return "";
    } else {
        this->sourcePath = path;
        xlsx = new QXlsx::Document(path);
        return path;
    }
}

void ExcelParser::receiveMessage(const int msgType, const QString &result) {
    qDebug() << "XlsxParser::receiveMessage msgType:" << QString::number(msgType).toUtf8() << " msg:" << result;
    switch (msgType) {
        case Common::MsgTypeError:
            m_failure_cnt++;
            m_receive_msg_cnt++;
            emit requestMsg(msgType, result);
            break;
        case Common::MsgTypeSucc:
            m_success_cnt++;
            m_receive_msg_cnt++;
            emit requestMsg(msgType, result);
            break;
        case Common::MsgTypeInfo:
        case Common::MsgTypeWarn:
        case Common::MsgTypeFinish:
        default:
            emit requestMsg(msgType, result);
            break;
    }
    if (m_total_cnt > 0 && m_receive_msg_cnt == m_total_cnt) { //全部处理完毕
        emit requestMsg(Common::MsgTypeWriteXlsxFinish, "excel文件拆分完毕！");
        if (this->sourceExcelData->getOpType() == SourceExcelData::OperateType::EmailOnlyType ||
            this->sourceExcelData->getOpType() == SourceExcelData::OperateType::SplitAndEmailType) {
            emit requestMsg(Common::MsgTypeStartSendEmail, "开始发送email");
        }
    }
}

//拆分excel文件
void ExcelParser::doParse() {
    qDebug() << "doSplit start";
    if (this->sourceExcelData->getNeedPassword()) {
        QHash<QString, QString> passwordData = readPasswordDataXls(passwordDataSheetName);
        this->sourceExcelData->setPasswordData(passwordData);
    }
    if (this->sourceExcelData->getOpType() == SourceExcelData::OperateType::SplitAndEmailType ||
        this->sourceExcelData->getOpType() == SourceExcelData::OperateType::EmailOnlyType) {
        qDebug() << "doSplit readEmailXls";
        //读取email
        emailQhash = readEmailXls(groupByText, emailSheetName);
        if (emailQhash.size() < 1) {
            emit requestMsg(Common::MsgTypeFail, "没有email数据");
            return;
        }
        if (this->sourceExcelData->getOpType() == SourceExcelData::OperateType::EmailOnlyType) {
            emit requestMsg(Common::MsgTypeStart, QString::number(emailQhash.size()));
            emit requestMsg(Common::MsgTypeStartSendEmail, "开始发送email");
            return;
        }
    }

    if (this->sourceExcelData->getOpType() == SourceExcelData::OperateType::SplitAndEmailType ||
        this->sourceExcelData->getOpType() == SourceExcelData::OperateType::SplitOnlyType) {
        qDebug() << "doSplit readDataXls";
        //读取excel数据
        emit requestMsg(Common::MsgTypeInfo, "开始读取excel文件信息");
        QHash<QString, QList<int>> dataQhash = readDataXls(groupByText, dataSheetName);
        if (this->sourceExcelData->getOpType() == SourceExcelData::OperateType::SplitOnlyType) {
            emit requestMsg(Common::MsgTypeStart, QString::number(dataQhash.size()));
        } else {
            emit requestMsg(Common::MsgTypeStart, QString::number(dataQhash.size() + emailQhash.size()));
        }
        if (dataQhash.size() < 1) {
            emit requestMsg(Common::MsgTypeFail, "没有data数据！！");
            return;
        }
        //写excel
        emit requestMsg(Common::MsgTypeInfo, "开始拆分excel并生成新的excel文件");
        m_total_cnt = dataQhash.size();
        qDebug() << "doSplit writeXls";
        int excelLibType = cfg->get("email", "excelLibType").toInt();
        bool useMsOffice = false;
        if (excelLibType == 0) { //自动识别
            if (this->isInstalledOffice) {
                useMsOffice = true;
            }
        } else if (excelLibType == 1) { //自带类库
            useMsOffice = false;
        } else if (excelLibType == 3) { //使用MS office
            useMsOffice = true;
        }

        if (useMsOffice) {
            writeXlsByOffice(dataSheetName, dataQhash);
        } else {
            writeXlsByLib(dataSheetName, dataQhash);
        }
    }
}

QHash<QString, QList<QStringList>> ExcelParser::getEmailData() {
    return emailQhash;
}

//读取xls
QHash<QString, QList<int>> ExcelParser::readDataXls(QString groupByText, QString selectedSheetName) {
    QXlsx::CellRange range;
    xlsx->selectSheet(selectedSheetName);
    range = xlsx->dimension();
    int rowCount = range.rowCount();
    int colCount = range.columnCount();

    QHash<QString, QList<int>> qHash;
    int groupBy = 0;
    for (int colum = 1; colum <= colCount; ++colum) {
        QXlsx::Cell *cell = xlsx->cellAt(1, colum);
        QXlsx::Format format = cell->format();
        if (cell) {
            if (groupByText == cell->value().toString()) {
                groupBy = colum;
                break;
            }
        }
    }
    if (groupBy == 0) { //没有对应的分组
        emit requestMsg(Common::MsgTypeError, "分组列“" + groupByText + "” 不存在");
        return qHash;
    }

    for (int row = 2; row <= rowCount; ++row) {
        QString groupByValue;
        QXlsx::Cell *cell = xlsx->cellAt(row, groupBy);
        if (cell) {
            groupByValue = cell->value().toString().trimmed();
        }
        if (groupByValue.isNull() || groupByValue.isEmpty()) {
            emit requestMsg(Common::MsgTypeError, "第" + QString::number(row, 10) + "” 为空，请删除所有空白行后再进行操作！");
            return qHash;
        }

        QList<int> qlist = qHash.take(groupByValue);
        qlist.append(row);
        qHash.insert(groupByValue, qlist);
    }
    return qHash;
}

//读取xls
QHash<QString, QList<QStringList>> ExcelParser::readEmailXls(QString groupByText, QString selectedSheetName) {
    QXlsx::CellRange range;
    xlsx->selectSheet(selectedSheetName);
    range = xlsx->dimension();
    int rowCount = range.rowCount();
    int colCount = range.columnCount();

    QHash<QString, QList<QStringList>> qhash;
    int groupBy = 0;
    for (int colum = 1; colum <= colCount; ++colum) {
        QXlsx::Cell *cell = xlsx->cellAt(1, colum);
        QXlsx::Format format = cell->format();
        if (cell) {
            if (groupByText == cell->value().toString().trimmed()) {
                groupBy = colum;
                break;
            }
        }
    }
    if (groupBy == 0) { //没有对应的分组
        emit requestMsg(Common::MsgTypeError, "分组列“" + groupByText + "” 不存在");
        return qhash;
    }

    for (int row = 2; row <= rowCount; ++row) {
        QString groupByValue;
        QStringList items;
        QXlsx::Cell *cell = xlsx->cellAt(row, groupBy);
        if (cell) {
            groupByValue = cell->value().toString().trimmed();
        }
        if (groupByValue.isNull() || groupByValue.isEmpty()) {
            continue;
        }
        for (int colum = 1; colum <= colCount; ++colum) {
            QXlsx::Cell *cell = xlsx->cellAt(row, colum);
            QString currentCellValue;
            if (cell) {
                if (cell->isDateTime()) {
                    currentCellValue = cell->dateTime().toString("yyyy/MM/dd");
                } else {
                    currentCellValue = cell->value().toString().trimmed();
                    if (currentCellValue.isNull() || currentCellValue.isEmpty()) {
                        continue;
                    }
                }
                items.append(currentCellValue);
            }
        }
        if (items.size() < Common::EmailColumnMinCnt) {
            emit requestMsg(Common::MsgTypeError,
                            "email第" + QString::number(row) + "行数据列小于" + QString::number(Common::EmailColumnMinCnt));
            qhash.clear();
            return qhash;
        }
        QList<QStringList> qlist = qhash.take(groupByValue);
        qlist.append(items);
        qhash.insert(groupByValue, qlist);
    }
    return qhash;
}


QHash<QString, QString> ExcelParser::readPasswordDataXls(QString selectedSheetName) {
    QXlsx::CellRange range;
    xlsx->selectSheet(selectedSheetName);
    range = xlsx->dimension();
    int rowCount = range.rowCount();

    QHash<QString, QString> qhash;

    for (int row = 2; row <= rowCount; ++row) {
        QString groupByValue;
        QString password;
        QXlsx::Cell *groupBycell = xlsx->cellAt(row, 1);
        if (groupBycell) {
            groupByValue = groupBycell->value().toString().trimmed();
        }
        if (groupByValue.isNull() || groupByValue.isEmpty()) {
            continue;
        }
        QXlsx::Cell *passwordCell = xlsx->cellAt(row, 2);
        if (passwordCell) {
            password = passwordCell->value().toString().trimmed();
        }
        qhash.insert(groupByValue, password);
    }
    return qhash;
}


//写xls(基于 MS office)
void ExcelParser::writeXlsByOffice(QString selectedSheetName, QHash<QString, QList<int>> qHash) {
    QHashIterator<QString, QList<int>> it(qHash);
    QThreadPool pool;
    int maxThreadCnt = cfg->get("email", "maxThreadCnt").toInt();
    if (maxThreadCnt < 1) {
        maxThreadCnt = 4;
    }

    int maxProcessCntPreThread = qCeil(qHash.size() * 1.0 / maxThreadCnt);
    qDebug() << "qHash.size() :" << qHash.size() << " maxThreadCnt:" << maxThreadCnt << " maxProcessCntPreThread:" << maxProcessCntPreThread;
    QHash<QString, QList<int>> fragmentDataQhash;

    pool.setMaxThreadCount(maxThreadCnt);
    int totalCnt = qHash.size();
    int runnableId = 1;
    ExcelParserByOfficeRunnable::processSourceFile(this->sourceExcelData, selectedSheetName);
    while (it.hasNext()) {
        it.next();
        QString key = it.key();
        QList<int> content = it.value();
        fragmentDataQhash.insert(key, content);
        int mod = m_process_cnt % maxProcessCntPreThread;
        if (mod == 0 || !it.hasNext()) {
            IExcelParserRunnable *runnable = new ExcelParserByOfficeRunnable(this);
            runnable->setID(runnableId++);
            runnable->setSplitData(this->sourceExcelData, selectedSheetName, fragmentDataQhash, totalCnt);
            pool.start(runnable);
            fragmentDataQhash.clear();
        }
    }
    pool.waitForDone();
}

void ExcelParser::writeXlsByLib(QString selectedSheetName, QHash<QString, QList<int>> qHash) {
    QHashIterator<QString, QList<int>> it(qHash);
    QThreadPool pool;
    int maxThreadCnt = cfg->get("email", "maxThreadCnt").toInt();
    if (maxThreadCnt < 1) {
        maxThreadCnt = 2;
    }
    maxThreadCnt = 1; //多线程处理xlsx会出现段错误 ???

    int maxProcessCntPreThread = qCeil(qHash.size() * 1.0 / maxThreadCnt);
    qDebug() << "qHash.size() :" << qHash.size() << " maxThreadCnt:" << maxThreadCnt << " maxProcessCntPreThread:" << maxProcessCntPreThread;
    QHash<QString, QList<int>> fragmentDataQhash;

    pool.setMaxThreadCount(maxThreadCnt);
    int totalCnt = qHash.size();
    int runnableId = 1;
    while (it.hasNext()) {
        it.next();
        QString key = it.key();
        QList<int> content = it.value();
        fragmentDataQhash.insert(key, content);
        int mod = m_process_cnt % maxProcessCntPreThread;
        if (mod == 0 || !it.hasNext()) {
            IExcelParserRunnable *runnable = new ExcelParserByLibRunnable(this);
            runnable->setID(runnableId++);
            runnable->setSplitData(this->sourceExcelData, selectedSheetName, fragmentDataQhash, totalCnt);
            pool.start(runnable);
            fragmentDataQhash.clear();
            pool.waitForDone();
        }
    }
}
