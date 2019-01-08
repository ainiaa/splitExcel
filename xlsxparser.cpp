#include "xlsxparser.h"

XlsxParser::XlsxParser(QObject *parent):mParent(parent)
{
    m_success_cnt = 0;
    m_failure_cnt = 0;
    m_process_cnt = 0;
    m_receive_msg_cnt = 0;
}

XlsxParser::~XlsxParser()
{
    qDebug() << "~XlsxParser start";
    if (xlsx != nullptr)
    {
        delete xlsx;
    }
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
    if (xlsx != nullptr)
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
        xlsx = new QXlsx::Document (path);
        return path;
    }
}


void XlsxParser::receiveMessage(const int msgType, const QString &result)
{
    qDebug() << "EmailSender::receiveMessage msgType:" << QString::number(msgType).toUtf8() <<" msg:"<<result;
    switch (msgType)
    {
    case Common::MsgTypeError:
        m_failure_cnt++;
        m_receive_msg_cnt++;
        emit requestMsg(msgType, msg.arg(m_success_cnt).arg(m_failure_cnt).arg(m_process_cnt).arg(result));
        break;
    case Common::MsgTypeSucc:
        m_success_cnt++;
        m_receive_msg_cnt++;
        emit requestMsg(msgType, msg.arg(m_success_cnt).arg(m_failure_cnt).arg(m_process_cnt).arg(result));
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
    //读取excel数据
    emit requestMsg(Common::MsgTypeInfo, "开始读取excel文件信息");
    QHash<QString, QList<int>> dataQhash = readDataXls(groupByText, dataSheetName);
    if (dataQhash.size() < 1)
    {
        emit requestMsg(Common::MsgTypeFail, "没有data数据！！");
        return;
    }
    emailQhash = readEmailXls(groupByText, emailSheetName);
    if (emailQhash.size() < 1)
    {
        emit requestMsg(Common::MsgTypeFail, "没有email数据");
        return;
    }

    //写excel
    emit requestMsg(Common::MsgTypeInfo, "开始拆分excel并生成新的excel文件");
    m_total_cnt = dataQhash.size();
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
    pool.setMaxThreadCount(maxThreadCnt);
    while (it.hasNext()) {
        it.next();
        QString key = it.key();
        QList<int> content = it.value();
        XlsxParserRunnable *runnable = new XlsxParserRunnable(this);
        runnable->setID(++m_process_cnt);
        runnable->setSplitData(xlsx, selectedSheetName, key, content, savePath, qHash.size());
        runnable->setAutoDelete(true);

        pool.start(runnable);

        if (m_process_cnt % maxThreadCnt == 0)
        {
            pool.waitForDone();
        }
    }
    if (pool.activeThreadCount() > 0)
    {
        pool.waitForDone();
    }
}

void  XlsxParser::writeXlsHeader(QXlsx::Document *xls, QString selectedSheetName)
{
    QXlsx::CellRange range;
    xlsx->selectSheet(selectedSheetName);
    range = xlsx->dimension();
    int colCount = range.columnCount();

    for (int column=1; column<=colCount; ++column)
    {
        QXlsx::Cell *sourceCell =xlsx->cellAt(1, column);
        if (sourceCell)
        {
            QString columnName;
            convertToColName(column,columnName) ;
            QString cell;
            QXlsx::Format newFormat = copyFormat(sourceCell->format());
            cell.append(columnName).append(QString::number(1));
            xls->write(cell, sourceCell->value(), newFormat);
            xls->setColumnFormat(column,xlsx->columnFormat(column));
            xls->setColumnWidth(column,xlsx->columnWidth(column));
        }
    }
    xls->setRowFormat(1,xlsx->rowFormat(1));
    xls->setRowHeight(1,xlsx->rowHeight(1));
}


QXlsx::Format XlsxParser::copyFormat(QXlsx::Format format)
{
    QXlsx::Format newFomat = format;
    newFomat.setFont(format.font());
    newFomat.setFontBold(format.fontBold());
    newFomat.setFontName(format.fontName());
    newFomat.setFontSize(format.fontSize());
    newFomat.setFontColor(format.fontColor());
    newFomat.setFontItalic(format.fontItalic());
    newFomat.setFontUnderline(format.fontUnderline());

    newFomat.setTextWarp(format.textWrap());

    newFomat.setNumberFormat(format.numberFormat());

    newFomat.setVerticalAlignment(format.verticalAlignment());
    newFomat.setHorizontalAlignment(format.horizontalAlignment());

    newFomat.setPatternBackgroundColor(format.patternBackgroundColor());
    newFomat.setPatternForegroundColor(format.patternForegroundColor());

    newFomat.setTopBorderStyle(format.topBorderStyle());
    newFomat.setTopBorderColor(format.topBorderColor());
    newFomat.setBottomBorderStyle(format.bottomBorderStyle());
    newFomat.setBottomBorderColor(format.bottomBorderColor());

    return newFomat;
}

///
/// \brief 把列数转换为excel的字母列号
/// \param data 大于0的数
/// \return 字母列号，如1->A 26->Z 27 AA
///
void XlsxParser::convertToColName(int data, QString &res)
{
    int tempData = data / 26;
    int mode = data % 26;
    if(tempData > 0 && mode > 0)
    {
        convertToColName(mode,res);
        convertToColName(tempData,res);
    }
    else
    {
        res=(to26AlphabetString(data)+res);
    }
}
///
/// \brief 数字转换为26字母
///
/// 1->A 26->Z
/// \param data
/// \return
///
QString XlsxParser::to26AlphabetString(int data)
{
    QChar ch = data + 0x40;//A对应0x41
    return QString(ch);
}
