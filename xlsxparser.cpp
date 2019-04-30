#include "xlsxparser.h"

XlsxParser::XlsxParser(QObject *parent):QObject (parent)
{
    m_success_cnt = 0;
    m_failure_cnt = 0;
    m_process_cnt = 0;
    m_receive_msg_cnt = 0;
}

XlsxParser::~XlsxParser()
{
    qDebug() << "~XlsxParser start";

    qDebug() << "~XlsxParser end";
}
void XlsxParser::setSplitData(Config *cfg, QString groupByText, QString dataSheetName, QString emailSheetName, QString savePath)
{
    this->cfg = cfg;
    this->groupByText = groupByText;
    this->dataSheetName = dataSheetName;
    this->emailSheetName = emailSheetName;
    this->savePath = savePath;
}

QStringList XlsxParser::getSheetNames()
{
    QStringList sheetNames;
    if (nullptr != xlsx)
    {
        sheetNames = xlsx->sheetNames();
    }
    return sheetNames;
}

bool XlsxParser::selectSheet(const QString &name)
{
    return xlsx->selectSheet(name);
}

QXlsx::CellRange XlsxParser::dimension()
{
    return xlsx->dimension();
}

QStringList* XlsxParser::getSheetHeader(QString selectedSheetName)
{
    QStringList *currentHeader = new QStringList();
    QXlsx::CellRange range;
    xlsx->selectSheet(selectedSheetName);
    range = xlsx->dimension();
    int colCount = range.columnCount();

    for (int colum=1; colum<=colCount; ++colum)
    {
        QXlsx::Cell *cell =xlsx->cellAt(1, colum);
        if (cell)
        {
            currentHeader->append(cell->value().toString());
        }
    }
    return currentHeader;
}


//打开文件对话框
QString XlsxParser::openFile(QWidget *dlgParent)
{
    QString path = QFileDialog::getOpenFileName(dlgParent, tr("Open Excel"), ".", tr("excel(*.xlsx)"));
    if(path.length() == 0)
    {
        return "";
    }
    else
    {
        this->sourcePath = path;
        xlsx = new QXlsx::Document(path);
        return path;
    }
}


void XlsxParser::receiveMessage(const int msgType, const QString &result)
{
    qDebug() << "XlsxParser::receiveMessage msgType:" << QString::number(msgType).toUtf8() <<" msg:"<<result;
    switch (msgType)
    {
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
    if (m_total_cnt > 0 && m_receive_msg_cnt == m_total_cnt)
    { //全部处理完毕
        emit requestMsg(Common::MsgTypeWriteXlsxFinish, "excel文件拆分完毕！");
    }
}

//拆分excel文件
void XlsxParser::doSplit()
{
    qDebug() << "doSplit";
    if (nullptr !=emailSheetName )
    {
        qDebug() << "doSplit readEmailXls";
        //读取email
        emailQhash = readEmailXls(groupByText, emailSheetName);
        if (emailQhash.size() < 1)
        {
            emit requestMsg(Common::MsgTypeFail, "没有email数据");
            return;
        }
    }

    qDebug() << "doSplit readDataXls";
    //读取excel数据
    emit requestMsg(Common::MsgTypeInfo, "开始读取excel文件信息");
    QHash<QString, QList<int>> dataQhash = readDataXls(groupByText, dataSheetName);
    if (dataQhash.size() < 1)
    {
        emit requestMsg(Common::MsgTypeFail, "没有data数据！！");
        return;
    }

    //写excel
    emit requestMsg(Common::MsgTypeInfo, "开始拆分excel并生成新的excel文件");
    m_total_cnt = dataQhash.size();
    qDebug() << "doSplit writeXls";
    writeXls(dataSheetName, dataQhash, savePath);
}

QHash<QString, QList<QStringList>> XlsxParser::getEmailData()
{
    return emailQhash;
}

//读取xls
QHash<QString, QList<int>> XlsxParser::readDataXls(QString groupByText, QString selectedSheetName)
{
    QXlsx::CellRange range;
    xlsx->selectSheet(selectedSheetName);
    range = xlsx->dimension();
    int rowCount = range.rowCount();
    int colCount = range.columnCount();

    QHash<QString, QList<int>> qHash;
    int groupBy = 0;
    for (int colum=1; colum<=colCount; ++colum)
    {
        QXlsx::Cell *cell =xlsx->cellAt(1, colum);
        QXlsx::Format format = cell->format();
        if (cell)
        {
            if (groupByText ==cell->value().toString())
            {
                groupBy = colum;
                break;
            }
        }
    }
    if (groupBy == 0)
    {//没有对应的分组
        emit requestMsg(Common::MsgTypeError, "分组列“" + groupByText + "” 不存在");
        return qHash;
    }

    for (int row = 2;row <= rowCount;++row)
    {
        QString groupByValue;
        QXlsx::Cell *cell =xlsx->cellAt(row, groupBy);
        if (cell)
        {
            groupByValue = cell->value().toString();
        }

        QList<int> qlist = qHash.take(groupByValue);
        qlist.append(row);
        qHash.insert(groupByValue,qlist);
    }
    return qHash;
}

//读取xls
QHash<QString, QList<QStringList>> XlsxParser::readEmailXls(QString groupByText, QString selectedSheetName)
{
    QXlsx::CellRange range;
    xlsx->selectSheet(selectedSheetName);
    range = xlsx->dimension();
    int rowCount = range.rowCount();
    int colCount = range.columnCount();

    QHash<QString, QList<QStringList>> qhash;
    int groupBy = 0;
    for (int colum=1; colum<=colCount; ++colum)
    {
        QXlsx::Cell *cell =xlsx->cellAt(1, colum);
        QXlsx::Format format = cell->format();
        if (cell)
        {
            if (groupByText ==cell->value().toString())
            {
                groupBy = colum;
                break;
            }
        }
    }
    if (groupBy == 0)
    {//没有对应的分组
        emit requestMsg(Common::MsgTypeError, "分组列“" + groupByText + "” 不存在");
        return qhash;
    }

    for (int row = 2;row <= rowCount;++row)
    {
        QString groupByValue;
        QStringList items;
        QXlsx::Cell *cell =xlsx->cellAt(row, groupBy);
        if (cell)
        {
            groupByValue = cell->value().toString();
        }
        if (groupByValue.isNull() || groupByValue.isEmpty())
        {
            continue;
        }
        for (int colum=1; colum<=colCount; ++colum)
        {
            QXlsx::Cell *cell =xlsx->cellAt(row, colum);
            if (cell)
            {
                if (cell->isDateTime())
                {
                    items.append(cell->dateTime().toString("yyyy/MM/dd"));
                }
                else
                {
                    items.append(cell->value().toString());
                }
            }
        }
        if ( items.size() < Common::EmailColumnMinCnt)
        {
            emit requestMsg(Common::MsgTypeError, "email第" +QString::number(row) + "行数据列小于" + QString::number(Common::EmailColumnMinCnt));
            qhash.clear();
            return qhash;
        }
        QList<QStringList> qlist = qhash.take(groupByValue);
        qlist.append(items);
        qhash.insert(groupByValue,qlist);
    }
    return qhash;
}

//写xls
void XlsxParser::writeXls(QString selectedSheetName, QHash<QString, QList<int>> qHash, QString savePath)
{
    QHashIterator<QString,QList<int>> it(qHash);
    QThreadPool pool;
    int maxThreadCnt = cfg->get("email","maxThreadCnt").toInt();
    if (maxThreadCnt < 1)
    {
        maxThreadCnt = 2;
    }
    maxThreadCnt = 1; //多线程处理xlsx会出现段错误 ???

    int maxProcessCntPreThread = qCeil(qHash.size() * 1.0 / maxThreadCnt);
    qDebug() << "qHash.size() :" << qHash.size()  << " maxThreadCnt:" <<maxThreadCnt <<" maxProcessCntPreThread:" << maxProcessCntPreThread;
    QHash<QString, QList<int>> fragmentDataQhash;

    pool.setMaxThreadCount(maxThreadCnt);
    int totalCnt =qHash.size();
    int runnableId = 1;
    while (it.hasNext()) {
        it.next();
        QString key = it.key();
        QList<int> content = it.value();
//        XlsxParserRunnable *runnable = new XlsxParserRunnable(this);
//        runnable->setID(++m_process_cnt);
//        runnable->setSplitData(sourcePath, selectedSheetName, key, content, savePath, totalCnt);
//        pool.start(runnable);

//        if (m_process_cnt % maxThreadCnt == 0)
//        {
//            pool.waitForDone();
//        }
        fragmentDataQhash.insert(key,content);
        int mod = m_process_cnt % maxProcessCntPreThread;
        if (mod==0 || !it.hasNext())
        {
            XlsxParserRunnable *runnable = new XlsxParserRunnable(this) ;
            runnable->setID(runnableId++);
            runnable->setSplitData(sourcePath, selectedSheetName, fragmentDataQhash, savePath, totalCnt);
            pool.start(runnable);
            fragmentDataQhash.clear();
            pool.waitForDone();
        }
    }
}
