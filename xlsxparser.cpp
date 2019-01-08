#include "xlsxparser.h"

XlsxParser::XlsxParser(QObject *parent):mParent(parent)
{

}

XlsxParser::~XlsxParser()
{

}
void XlsxParser::setSplitData(QString groupByText, QString dataSheetName, QString emailSheetName, QString savePath)
{
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
        emit requestMsg(msgType, result);
        break;
    case Common::MsgTypeSucc:
        emit requestMsg(msgType,result);
        break;
    case Common::MsgTypeInfo:
    case Common::MsgTypeWarn:
    case Common::MsgTypeFinish:
    default:
        emit requestMsg(msgType, result);
        break;
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
    emailQhash = readEmailXls(groupByText, emailSheetName, true);
    if (emailQhash.size() < 1)
    {
        emit requestMsg(Common::MsgTypeFail, "没有email数据");
        return;
    }

    //写excel
    emit requestMsg(Common::MsgTypeInfo, "开始拆分excel并生成新的excel文件");
    writeXls(dataSheetName,dataQhash,savePath);
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
QHash<QString, QList<QStringList>> XlsxParser::readEmailXls(QString groupByText, QString selectedSheetName, bool isEmail)
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
        if (isEmail && items.size() < Common::EmailColumnMinCnt)
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
    QXlsx::CellRange range;
    xlsx->selectSheet(selectedSheetName);
    range = xlsx->dimension();
    int rowCount = range.rowCount();
    int colCount = range.columnCount();

    qDebug("writeXls");
    QHashIterator<QString,QList<int>> it(qHash);
    QString startMsg("开始拆分excel并生成新的excel文件: %1/");
    QString endMsg("完成拆分excel并生成新的excel文件: %1/");
    startMsg.append(QString::number(qHash.size()));
    endMsg.append(QString::number(qHash.size()));
    int count = 1;
    while (it.hasNext()) {
        it.next();

        emit requestMsg(Common::MsgTypeInfo,startMsg.arg(QString::number(count)));
        QString key = it.key();
        QList<int> content = it.value();
        QXlsx::Document currXls;

        //写表头
        writeXlsHeader(&currXls, selectedSheetName);

        QString xlsName;
        int rows =  content.size();
        xlsName.append(savePath).append("\\").append(key).append(".xlsx");

        for(int row = 0; row < rows;row++)
        {
            int dataRow = content.at(row);
            QXlsx::CellReference cellReference = "";
            QVariant value = xlsx->read(cellReference);

            for (int column=1; column<=colCount; ++column)
            {
                QXlsx::Cell *sourceCell =xlsx->cellAt(dataRow, column);
                if (sourceCell)
                {
                    QString columnName;
                    convertToColName(column,columnName) ;
                    QString cell;
                    QXlsx::Format format = sourceCell->format();
                    QXlsx::Format newfmt;

                    cell.append(columnName).append(QString::number(row+2));
                    newfmt.setFont(format.font());
                    newfmt.setFontBold(format.fontBold());
                    newfmt.setFontName(format.fontName());
                    newfmt.setFontSize(format.fontSize());
                    newfmt.setFontColor(format.fontColor());
                    newfmt.setFontItalic(format.fontItalic());
                    newfmt.setFontUnderline(format.fontUnderline());

                    newfmt.setTextWarp(format.textWrap());

                    newfmt.setNumberFormat(format.numberFormat());

                    newfmt.setVerticalAlignment(format.verticalAlignment());
                    newfmt.setHorizontalAlignment(format.horizontalAlignment());

                    newfmt.setPatternBackgroundColor(format.patternBackgroundColor());
                    newfmt.setPatternForegroundColor(format.patternForegroundColor());

                    if (sourceCell->isDateTime() && !sourceCell->value().isNull())
                    {
                        currXls.write(cell,sourceCell->dateTime(),newfmt);
                    }
                    else
                    {
                        currXls.write(cell,sourceCell->value(), newfmt);
                    }
                }
            }
        }
        currXls.saveAs(xlsName);
        emit requestMsg(Common::MsgTypeInfo,endMsg.arg(QString::number(count)));
        count++;
    }
    emit requestMsg(Common::MsgTypeInfo,"拆分excel结束");
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
            QXlsx::Format sourceFormat = sourceCell->format();

            QXlsx::Format format = sourceCell->format();
            QXlsx::Format newfmt = format;
            newfmt.setFont(format.font());
            newfmt.setFontBold(format.fontBold());
            newfmt.setFontName(format.fontName());
            newfmt.setFontSize(format.fontSize());
            newfmt.setFontColor(format.fontColor());
            newfmt.setFontItalic(format.fontItalic());
            newfmt.setFontUnderline(format.fontUnderline());
            newfmt.setTextWarp(format.textWrap());
            newfmt.setNumberFormat(format.numberFormat());
            newfmt.setVerticalAlignment(format.verticalAlignment());
            newfmt.setHorizontalAlignment(format.horizontalAlignment());
            newfmt.setPatternBackgroundColor(format.patternBackgroundColor());
            newfmt.setPatternForegroundColor(format.patternForegroundColor());

            cell.append(columnName).append(QString::number(1));
            qDebug() << "cell:" <<cell << " column:" <<column;
            xls->write(cell, sourceCell->value(), newfmt);
            xls->setColumnFormat(column,xlsx->columnFormat(column));
            xls->setColumnWidth(column,xlsx->columnWidth(column));
        }
    }
    xls->setRowFormat(1,xlsx->rowFormat(1));
    xls->setRowHeight(1,xlsx->rowHeight(1));

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
