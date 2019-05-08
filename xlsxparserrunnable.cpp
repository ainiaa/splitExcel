#include "xlsxparser.h"
#include "xlsxparserrunnable.h"

XlsxParserRunnable::XlsxParserRunnable(QObject *parent)
{
    this->mParent = parent;
    this->installedExcelApp = isInstalledExcelApp();
}

XlsxParserRunnable::~XlsxParserRunnable()
{
    runnableID = 0;
}

void XlsxParserRunnable::setID(const int &id)
{
    runnableID = id;
}

void XlsxParserRunnable::setSplitData(QString sourcePath, QString selectedSheetName, QHash<QString, QList<int>> fragmentDataQhash, QString savePath, int total)
{
    qDebug("setSplitData") ;
    this->xlsx = new QXlsx::Document(sourcePath);
    this->sourcePath = sourcePath;
    this->selectedSheetName = selectedSheetName;
    this->fragmentDataQhash = fragmentDataQhash;
    this->savePath = savePath;
    this->m_total = total;
}


void XlsxParserRunnable::run()
{
    qDebug("XlsxParserRunnable::run start") ;
    QString startMsg("开始拆分excel并生成新的excel文件: %1/");
    QString endMsg("完成拆分excel并生成新的excel文件: %1/");
    startMsg.append(QString::number(m_total));
    endMsg.append(QString::number(m_total));
    requestMsg(Common::MsgTypeInfo, startMsg.arg(QString::number(runnableID)));
    QHashIterator<QString,QList<int>> it(fragmentDataQhash);
    while (it.hasNext()) {
        it.next();
        QString key = it.key();
        QList<int> contentList = it.value();
        contentList.insert(0,1);

        if (this->installedExcelApp)
        {//安装了 offcie
            this->processByOffice(key, contentList);
        }
        else
        {
            this->processByQxls(key, contentList);
        }
    }
    requestMsg(Common::MsgTypeSucc, endMsg.arg(QString::number(runnableID)));
    qDebug("XlsxParserRunnable::run end") ;
}

bool XlsxParserRunnable::isInstalledExcelApp()
{
    QAxObject excel("Excel.Application");
    if (excel.isNull())
    {
        return false;
    }
    return true;
}

void XlsxParserRunnable::processByOffice(QString key, QList<int> contentList)
{
    QString xlsName;
    xlsName.append(savePath).append(QDir::separator()).append(key).append(".xlsx");
    copyFileToPath(sourcePath,xlsName,true);

    QAxObject *excel = new QAxObject("Excel.Application");//连接Excel控件
    excel->dynamicCall("SetVisible (bool Visible)",false);//不显示窗体
    excel->setProperty("DisplayAlerts", true);//不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示
    excel->setProperty("EnableEvents",false); //没有这个 很容易报错  QAxBase: Error calling IDispatch member Open: Unknown error

    QAxObject *workbooks = excel->querySubObject("WorkBooks");//获取工作簿集合
    QAxObject *workbook = workbooks->querySubObject("Open(const QString&, QVariant)", xlsName, 0);
    QAxObject * worksheets = workbook->querySubObject("WorkSheets");// 获取打开的excel文件中所有的工作sheet
    int intWorkSheet = worksheets->property("Count").toInt();
    qDebug() << QString("Excel文件中表的个数: %1").arg(QString::number(intWorkSheet));
    QAxObject* worksheet = nullptr;
    for (int i = 1; i <= intWorkSheet;i++)
    {
        QAxObject *currentWorkSheet = workbook->querySubObject("WorkSheets(int)", i);  //Sheets(int)也可换用Worksheets(int)
        QString currentWorkSheetName = currentWorkSheet->property("Name").toString();  //获取工作表名称
        if (currentWorkSheetName != this->selectedSheetName)
        {
            qDebug()<<"delete work_sheet:"<<currentWorkSheetName;
            currentWorkSheet->dynamicCall("Delete()");
        }
        else
        {
            worksheet = currentWorkSheet;
        }
    }

    if (worksheet == nullptr)
    {
        qDebug()<<"worksheet is null:";
        return;
    }
    QAxObject* usedrange = worksheet->querySubObject("UsedRange"); // sheet范围
    int intRowStart = usedrange->property("Row").toInt(); // 起始行数
    int intColStart = usedrange->property("Column").toInt();  // 起始列数

    QAxObject *rows, *columns;
    rows = usedrange->querySubObject("Rows");  // 行
    columns = usedrange->querySubObject("Columns");  // 列

    int intRow = rows->property("Count").toInt(); // 行数
    int intCol = columns->property("Count").toInt();  // 列数

    QString alphabetColStart;
    QString alphabetCol;
    ExcelBase::convertToColName(intRowStart,alphabetColStart);
    ExcelBase::convertToColName(intCol,alphabetCol);

    qDebug() << " intRowStart:" <<intRowStart<< " intColStart:" <<intColStart<< " intRow:" <<intRow<< " intCol:" <<intCol << " alphabetColStart:" <<alphabetColStart << " alphabetCol:" << alphabetCol;

    QString rangeFormat = "A%1:%2%3";
    for(int row= 1; row <= intRow;row++)
    {
        if (!contentList.contains(row))
        {
            QAxObject* range = worksheet->querySubObject("Range(QVariant)", rangeFormat.arg(row).arg(alphabetCol).arg(row)); //获取单元格
            range->dynamicCall("Delete()");
        }
        else
        {
            qDebug() << "left row:" << row;
        }
    }

    workbook->dynamicCall("Save()");//保存
    workbook->dynamicCall("Close()");//关闭工作簿
    excel->dynamicCall("Quit()");//关闭excel
    delete excel;
    excel=nullptr;
}
void XlsxParserRunnable::processByQxls(QString key, QList<int> contentList)
{
    QString xlsName;
    xlsName.append(savePath).append(QDir::separator()).append(key).append(".xlsx");
    copyFileToPath(sourcePath,xlsName,true);

    QXlsx::Document* currXls = new QXlsx::Document(xlsName);

    currXls->fliterRows(contentList);//删除多余的行  多线程的生活会出现问题。

    //删除多余的sheet start
    QStringList sheetNames = currXls->sheetNames();
    for (int i = 0;i<sheetNames.size();i++)
    {
        QString currSheetName = sheetNames.at(i);
        if (selectedSheetName == currSheetName)
        {
            continue;
        }
        currXls->deleteSheet(currSheetName);
    }
    //删除多余的sheet end


    //尝试修复行高不正确的问题
    QXlsx::CellRange cellRange = currXls->dimension();
    for (int row_num = cellRange.firstRow(); row_num <= cellRange.lastRow(); row_num++) {
        double rowHeigh = currXls->rowHeight(row_num);
        if (rowHeigh < 14)
        {
            currXls->setRowHeight(row_num, 18);
        }
    }

    currXls->save();
    delete currXls;
}

// old  start

/*
void XlsxParserRunnable::run()
{
    qDebug("XlsxParserRunnable::run start") ;
    requestMsg(Common::MsgTypeInfo, "XlsxParserRunnable::run");

    QXlsx::CellRange range;
    xlsx->selectSheet(selectedSheetName);
    range = xlsx->dimension();
    int colCount = range.columnCount();

    QString startMsg("开始拆分excel并生成新的excel文件: %1/");
    QString endMsg("完成拆分excel并生成新的excel文件: %1/");
    startMsg.append(QString::number(m_total));
    endMsg.append(QString::number(m_total));
    requestMsg(Common::MsgTypeInfo, startMsg.arg(QString::number(runnableID)));

    QHashIterator<QString,QList<int>> it(fragmentDataQhash);
    while (it.hasNext()) {
        it.next();
        QString key = it.key();
        QList<int> contentList = it.value();
        QXlsx::Document currXls;

        //写表头
        writeXlsHeader(&currXls, selectedSheetName);
        QString xlsName;
        int rows =  contentList.size();
        xlsName.append(savePath).append("\\").append(key).append(".xlsx");
        qDebug() << "xlsName:" << xlsName;

        for(int row = 0; row < rows;row++)
        {
            int dataRow = contentList.at(row);
            int newRow = row + 2;
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
                    QXlsx::Format newFormat = copyFormat(sourceCell->format());
                    cell.append(columnName).append(QString::number(newRow));
                    if (sourceCell->isDateTime() && !sourceCell->value().isNull())
                    {
                        currXls.write(cell,sourceCell->dateTime(), newFormat);
                    }
                    else
                    {
                        currXls.write(cell,sourceCell->value(), newFormat);
                    }
                    currXls.setColumnFormat(column, xlsx->columnFormat(column));
                    currXls.setColumnWidth(column, xlsx->columnWidth(column));
                }
            }
            currXls.setRowFormat(newRow, xlsx->rowFormat(newRow));
            currXls.setRowHeight(newRow, xlsx->rowHeight(newRow));
        }
        currXls.saveAs(xlsName);
        requestMsg(Common::MsgTypeSucc, endMsg.arg(QString::number(runnableID)));
    }

    qDebug("EmailSenderRunnable::run end") ;
}
*/

void  XlsxParserRunnable::writeXlsHeader(QXlsx::Document *xls, QString selectedSheetName)
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


QXlsx::Format XlsxParserRunnable::copyFormat(QXlsx::Format format)
{
    QXlsx::Format newFomat = format;
    newFomat.setFont(format.font());
    newFomat.setFontBold(format.fontBold());
    newFomat.setFontName(format.fontName());
    newFomat.setFontSize(format.fontSize());
    newFomat.setFontColor(format.fontColor());
    newFomat.setFontItalic(format.fontItalic());
    newFomat.setFontUnderline(format.fontUnderline());
    newFomat.setFontScript(format.fontScript());
    newFomat.setFontOutline(format.fontOutline());

    newFomat.setTextWarp(format.textWrap());

    newFomat.setHidden(format.hidden());

    newFomat.setIndent(format.indent());
    newFomat.setLocked(format.locked());
    newFomat.setXfIndex(format.xfIndex());
    newFomat.setDxfIndex(format.dxfIndex());
    newFomat.setRotation(format.rotation());
    newFomat.setFillIndex(format.fillIndex());
    newFomat.setFillPattern(format.fillPattern());
    newFomat.setBorderIndex(format.borderIndex());
    newFomat.setShrinkToFit(format.shrinkToFit());
    newFomat.setNumberFormat(format.numberFormat());
    newFomat.setRightBorderColor(format.rightBorderColor());
    newFomat.setRightBorderStyle(format.rightBorderStyle());
    newFomat.setTopBorderColor(format.topBorderColor());
    newFomat.setTopBorderStyle(format.topBorderStyle());
    newFomat.setLeftBorderColor(format.leftBorderColor());
    newFomat.setLeftBorderStyle(format.leftBorderStyle());
    newFomat.setBottomBorderColor(format.bottomBorderColor());
    newFomat.setBottomBorderStyle(format.bottomBorderStyle());
    newFomat.setNumberFormatIndex(format.numberFormatIndex());
    newFomat.setDiagonalBorderType(format.diagonalBorderType());
    newFomat.setDiagonalBorderColor(format.diagonalBorderColor());
    newFomat.setDiagonalBorderStyle(format.diagonalBorderStyle());

    newFomat.setVerticalAlignment(format.verticalAlignment());
    newFomat.setHorizontalAlignment(format.horizontalAlignment());

    newFomat.setPatternBackgroundColor(format.patternBackgroundColor());
    newFomat.setPatternForegroundColor(format.patternForegroundColor());

    return newFomat;
}

//old end


void XlsxParserRunnable::requestMsg(const int msgType, const QString &msg)
{
     qobject_cast<XlsxParser *>(mParent)->receiveMessage(msgType,msg.arg(msgType));
    //QMetaObject::invokeMethod(mParent, "receiveMessage", Qt::QueuedConnection, Q_ARG(int,msgType),Q_ARG(QString, msg));//不能及时返回信息
}

///
/// \brief 把列数转换为excel的字母列号
/// \param data 大于0的数
/// \return 字母列号，如1->A 26->Z 27 AA
///
void XlsxParserRunnable::convertToColName(int data, QString &res)
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
QString XlsxParserRunnable::to26AlphabetString(int data)
{
    QChar ch = data + 0x40;//A对应0x41
    return QString(ch);
}

bool XlsxParserRunnable::copyFileToPath(QString sourceDir ,QString toDir, bool coverFileIfExist)
{
    sourceDir = QDir::toNativeSeparators(sourceDir);
    toDir = QDir::toNativeSeparators(toDir);

    if (sourceDir == toDir){
        return true;
    }
    if (!QFile::exists(sourceDir)){
        return false;
    }
    QDir *createfile     = new QDir();
    bool exist = createfile->exists(toDir);
    if (exist){
        if(coverFileIfExist){
            createfile->remove(toDir);
        }
    }//end if

    if(!QFile::copy(sourceDir, toDir))
    {
        return false;
    }
    return true;
}

