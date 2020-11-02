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

void ExcelParser::setSplitData(Config *cfg, SourceData *sourceData) {
    this->cfg = cfg;
    this->groupByText = sourceData->getGroupByText();
    this->dataSheetName = sourceData->getDataSheetName();
    this->emailSheetName = sourceData->getEmailSheetName();
    this->savePath = sourceData->getSavePath();
    this->dataSheetIndex = sourceData->getDataSheetIndex();
    this->groupByIndex = sourceData->getGroupByIndex();
    this->sourceData = sourceData;
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
        if (this->sourceData->getOpType() == SourceData::OperateType::EmailOnlyType ||
                this->sourceData->getOpType() == SourceData::OperateType::SplitAndEmailType) {
            emit requestMsg(Common::MsgTypeStartSendEmail, "开始发送email");
        }
    }
}

/**
 * @brief ExcelParser::doParse 拆分excel文件
 */
void ExcelParser::doParse() {
    qDebug() << "doSplit start";
    if (this->sourceData->getOpType() == SourceData::OperateType::SplitAndEmailType ||
            this->sourceData->getOpType() == SourceData::OperateType::EmailOnlyType) { // 需要发送email
        qDebug() << "doSplit readEmailXls";
        //读取email
        emailQhash = readEmailXls(groupByText, emailSheetName);
        if (emailQhash.size() < 1) {
            emit requestMsg(Common::MsgTypeFail, "没有email数据");
            return;
        }
        if (this->sourceData->getOpType() == SourceData::OperateType::EmailOnlyType) {
            emit requestMsg(Common::MsgTypeStart, QString::number(emailQhash.size()));
            emit requestMsg(Common::MsgTypeStartSendEmail, "开始发送email");
            return;
        }
    }

    if (this->sourceData->getOpType() == SourceData::OperateType::SplitAndEmailType ||
            this->sourceData->getOpType() == SourceData::OperateType::SplitOnlyType) { // 需要拆分excel
        qDebug() << "doSplit readDataXls";

        emit requestMsg(Common::MsgTypeInfo, "开始读取excel文件信息");//读取excel数据
        int excelLibType = cfg->get("email", "excelLibType").toInt();
        bool useMs = false;
        if (excelLibType == 0) { //自动识别
            if (this->isInstalledOffice) {
                useMs = true;
            }
        } else if (excelLibType == 1) { // 自带类库
            useMs = false;
        } else if (excelLibType == 3) { //使用MS
            useMs = true;
        }
        if (useMs) { // 使用MS
            QHash<QString, QList<QList<QVariant>>> dataQhash2 = readDataXls(groupByIndex,dataSheetIndex);
            if (this->sourceData->getOpType() == SourceData::OperateType::SplitOnlyType) {
                emit requestMsg(Common::MsgTypeStart, QString::number(dataQhash2.size()));
            } else {
                emit requestMsg(Common::MsgTypeStart, QString::number(dataQhash2.size() + emailQhash.size()));
            }
            if (dataQhash2.size() < 1) {
                emit requestMsg(Common::MsgTypeFail, "没有data数据！！");
                return;
            }
            //写excel
            emit requestMsg(Common::MsgTypeInfo, "开始拆分excel并生成新的excel文件");
            m_total_cnt = dataQhash2.size();
            qDebug() << "doSplit writeXls";

            writeXlsByMS(dataSheetName, dataQhash2);
        } else { // 自带类库
            QHash<QString, QList<int>> dataQhash = readDataXls(groupByText, dataSheetName);
            if (this->sourceData->getOpType() == SourceData::OperateType::SplitOnlyType) {
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
QHash<QString, QList<QList<QVariant>>> ExcelParser::readDataXls(int groupByIndex, int selectedSheetIndex) {
    QVariant all = ExcelParser::readAll();

    QList<QList<QVariant>> allData;
    castVariant2ListListVariant(all, allData);
    QHash<QString, QList<QList<QVariant>>> qHash;
    QString groupByValue;
    for (int row = 1;row <allData.size();row++) {
        groupByValue = allData[row][groupByIndex].toString();
        if (groupByValue.isNull() || groupByValue.isEmpty()) {
            emit requestMsg(Common::MsgTypeError, "第" + QString::number(row, 10) + "” 为空，请删除所有空白行后再进行操作！");
            return qHash;
        }
        QList<QList<QVariant>> qlist = qHash.take(groupByValue);
        qlist.append(allData[row]);
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

//写xls(基于 MS office)
void ExcelParser::writeXlsByMS(QString selectedSheetName, QHash<QString, QList<int>> qHash) {
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
    ParserByMSRunnable::processSourceFile(this->sourceData, selectedSheetName);
    while (it.hasNext()) {
        it.next();
        QString key = it.key();
        QList<int> content = it.value();
        fragmentDataQhash.insert(key, content);
        int mod = m_process_cnt % maxProcessCntPreThread;
        if (mod == 0 || !it.hasNext()) {
            IExcelParserRunnable *runnable = new ParserByMSRunnable(this);
            runnable->setID(runnableId++);
            runnable->setSplitData(this->sourceData, selectedSheetName, fragmentDataQhash, totalCnt);
            pool.start(runnable);
            fragmentDataQhash.clear();
        }
    }
    pool.waitForDone();
}

/**
 * @brief ExcelParser::writeXlsByMS 写xls(基于 MS office)
 * @param selectedSheetName 被选中sheet名称
 * @param qHash excel数据
 */
void ExcelParser::writeXlsByMS(QString selectedSheetName, QHash<QString, QList<QList<QVariant>>> qHash) {
    QHashIterator<QString, QList<QList<QVariant>>> it(qHash);
    QThreadPool pool;
    int maxThreadCnt = cfg->get("email", "maxThreadCnt").toInt();
    if (maxThreadCnt < 1) { // 默认开10个线程
        maxThreadCnt = 10;
    }

    int maxProcessCntPreThread = qCeil(qHash.size() * 1.0 / maxThreadCnt);
    qDebug() << "qHash.size() :" << qHash.size() << " maxThreadCnt:" << maxThreadCnt << " maxProcessCntPreThread:" << maxProcessCntPreThread;
    QHash<QString, QList<QList<QVariant>>> fragmentQhash;

    pool.setMaxThreadCount(maxThreadCnt);
    int totalCnt = qHash.size();
    int runnableId = 1;
    //生成模板文件 方便多线程填写数据
    ParserByMSRunnable::generateTplXls(this->sourceData, this->sourceData->getDataSheetIndex() + 1); //生成模板文件（只有表头数据）
    while (it.hasNext()) {
        it.next();
        QString key = it.key(); // 当前分组
        QList<QList<QVariant>> content = it.value(); // 当前分组所包含数据（可能是多条数据）
        fragmentQhash.insert(key, content);
        int mod = m_process_cnt % maxProcessCntPreThread;
        if (mod == 0 || !it.hasNext()) {
            IExcelParserRunnable *runnable = new ParserByMSRunnable(this);
            runnable->setID(runnableId++);
            runnable->setSplitData(this->sourceData, selectedSheetName, fragmentQhash, totalCnt);
            pool.start(runnable);
            fragmentQhash.clear();
        }
    }
    pool.waitForDone();
}

void ExcelParser::freeExcel(QAxObject *excel) {
    if (!excel->isNull()) {
        excel->dynamicCall("Quit()");
        delete excel;
        excel = nullptr;
    }
}


/**
 * @brief ExcelParser::readAll 读取整个sheet
 * @return
 */
QVariant ExcelParser::readAll() {
    QVariant var;
    QString xlsName = this->sourceData->getSourcePath();
#if defined(Q_OS_WIN)
    QAxObject *excel = new QAxObject("Excel.Application");  //连接Excel控件
    excel->dynamicCall("SetVisible (bool Visible)", false); //不显示窗体
    excel->setProperty("DisplayAlerts", false); //不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示
    excel->setProperty("EnableEvents", false); //没有这个 很容易报错  QAxBase: Error calling IDispatch member Open: Unknown error

    qDebug() << "ExcelParserByOfficeRunnable::processByOffice xlsName=" << xlsName;
    QAxObject *workbooks = excel->querySubObject("WorkBooks"); //获取工作簿集合
    QAxObject *workbook = workbooks->querySubObject("Open(const QString&, QVariant)", xlsName, 0);
    QAxObject *worksheet = workbook->querySubObject("WorkSheets(int)", 1);
    if (worksheet != nullptr && ! worksheet->isNull()){
        QAxObject *usedRange = worksheet->querySubObject("UsedRange");
        if(nullptr == usedRange || usedRange->isNull()) {
            return var;
        }
        var = usedRange->dynamicCall("Value()");
        delete usedRange;
    }
    ExcelParser::freeExcel(excel);
#endif
    return var;
}

/**
 * @brief ExcelParser::readAll 读取整个sheet的数据，并放置到cells中
 * @param cells
 */
void ExcelParser::readAll(QList<QList<QVariant> > &cells)
{
#if defined(Q_OS_WIN)
    ExcelParser::castVariant2ListListVariant(readAll(),cells);
#else
    Q_UNUSED(cells);
#endif

}

/**
 * @brief ExcelParser::castListListVariant2Variant QList<QList<QVariant> >转换为QVariant
 * @param cells
 * @param res
 */
void ExcelParser::castListListVariant2Variant(const QList<QList<QVariant> > &cells, QVariant &res)
{
    QVariantList vars;
    const int rows = cells.size();
    for(int i=0;i<rows;++i) {
        vars.append(QVariant(cells[i]));
    }
    res = QVariant(vars);
}

/**
 * @brief ExcelParser::castVariant2ListListVariant 把QVariant转为QList<QList<QVariant> >
 * @param var
 * @param res
 */
void ExcelParser::castVariant2ListListVariant(const QVariant &var, QList<QList<QVariant> > &res)
{
    QVariantList varRows = var.toList();
    if(varRows.isEmpty()){
        return;
    }
    const int rowCount = varRows.size();
    QVariantList rowData;
    for(int i=0;i<rowCount;++i) {
        rowData = varRows[i].toList();
        res.push_back(rowData);
    }
}

/**
 * @brief ExcelParser::writeXlsByLib
 * @param selectedSheetName
 * @param qHash
 */
void ExcelParser::writeXlsByLib(QString selectedSheetName, QHash<QString, QList<int>> qHash) {
    QHashIterator<QString, QList<int>> it(qHash);
    QThreadPool pool;
    int maxThreadCnt = cfg->get("email", "maxThreadCnt").toInt();
    if (maxThreadCnt < 1) {
        maxThreadCnt = 2;
    }
    maxThreadCnt = 1; //多线程处理xlsx会出现段错误 QXLS 不支持多线程

    int maxProcessCntPreThread = qCeil(qHash.size() * 1.0 / maxThreadCnt);
    qDebug() << "qHash.size() :" << qHash.size() << " maxThreadCnt:" << maxThreadCnt << " maxProcessCntPreThread:" << maxProcessCntPreThread;
    QHash<QString, QList<int>> fragmentDataQhash;

    pool.setMaxThreadCount(maxThreadCnt);
    int totalCnt = qHash.size();
    int runnableId = 1;
    while (it.hasNext()){
        it.next();
        QString key = it.key();
        QList<int> content = it.value();
        fragmentDataQhash.insert(key, content);
        int mod = m_process_cnt % maxProcessCntPreThread;
        if (mod == 0 || !it.hasNext()) {
            IExcelParserRunnable *runnable = new ParserByLibRunnable(this);
            runnable->setID(runnableId++);
            runnable->setSplitData(this->sourceData, selectedSheetName, fragmentDataQhash, totalCnt);
            pool.start(runnable);
            fragmentDataQhash.clear();
            pool.waitForDone();
        }
    }
}
