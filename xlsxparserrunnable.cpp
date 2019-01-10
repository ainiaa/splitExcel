#include "xlsxparserrunnable.h"

XlsxParserRunnable::XlsxParserRunnable(QObject *parent)
{
    this->mParent = parent;
}

XlsxParserRunnable::~XlsxParserRunnable()
{
    runnableID = 0;
}

void XlsxParserRunnable::setID(const int &id)
{
    runnableID = id;
}

void XlsxParserRunnable::setSplitData(QXlsx::Document *xlsx, QString selectedSheetName, QString key, QList<int> contentList, QString savePath, int total)
{
    qDebug("setSplitData") ;
    this->xlsx = xlsx;
    this->selectedSheetName = selectedSheetName;
    this->key = key;
    this->contentList = contentList;
    this->savePath = savePath;
    this->m_total = total;
    contentList.insert(0,1);
    this->contentList = contentList;
}

void XlsxParserRunnable::run()
{
    qDebug("XlsxParserRunnable::run start") ;

    QXlsx::CellRange range;
    xlsx->selectSheet(selectedSheetName);
    range = xlsx->dimension();
    int colCount = range.columnCount();

    QString startMsg("开始拆分excel并生成新的excel文件: %1/");
    QString endMsg("完成拆分excel并生成新的excel文件: %1/");
    startMsg.append(QString::number(m_total));
    endMsg.append(QString::number(m_total));
    requestMsg(Common::MsgTypeInfo, startMsg.arg(QString::number(runnableID)));

    QString xlsName;
   // int rows =  contentList.size();
    xlsName.append(savePath).append("\\").append(key).append(".xlsx");
    xlsx->saveAs(xlsName);

    QXlsx::Document currXls(xlsName);

    currXls.fliterRows(contentList);
    //写表头
    //writeXlsHeader(&currXls, selectedSheetName);

    //QXlsx::Workbook* workbook = xlsx->workbook();
    //QXlsx::Worksheet* worksheet = (QXlsx::Worksheet*)xlsx->workbook()->activeSheet();
    // get full cells of sheet
//    int maxRow = -1;
//    int maxCol = -1;
//    QVector<QXlsx::CellLocation> clList = worksheet->getFullCells( &maxRow, &maxCol );

//    QList<int> rows2;
//    rows2.append(1);
//    rows2.append(2);
//    rows2.append(4);

    //QXlsx::Worksheet* newsheet = worksheet->copy("copied sheet", 1,rows2);



    //xlsx->copySheet("data"," copy new data");
   // xlsx->save();
//    for(int row = 0; row < rows;row++)
//    {
//        int dataRow = contentList.at(row);
//        int newRow = row + 2;

//        for (int column=1; column<=colCount; ++column)
//        {
//            QXlsx::Cell *sourceCell =xlsx->cellAt(dataRow, column);
//            if (sourceCell)
//            {
////                QString columnName;
////                convertToColName(column,columnName) ;
////                QString cell;
////                QXlsx::Format newFormat = copyFormat(sourceCell->format());
////                cell.append(columnName).append(QString::number(newRow));
////                if (sourceCell->isDateTime() && !sourceCell->value().isNull())
////                {
////                    currXls.write(cell,sourceCell->dateTime(), newFormat);
////                }
////                else
////                {
////                    currXls.write(cell,sourceCell->value(), newFormat);
////                }
//                QString cell;
//                QString columnName;
//                QXlsx::Format newFormat;
////                newFormat.mergeFormat(sourceCell->format());

////                QMapIterator<int, QVariant> it(sourceCell->format().getFormatPrivate()->properties);
////                while(it.hasNext()) {
////                    it.next();
////                    newFormat.setProperty(it.key(), it.value());
////                }

//                convertToColName(column,columnName) ;
//                cell.append(columnName).append(QString::number(newRow));
//                newFormat = copyFormat(sourceCell->format());
//                currXls.write(cell,sourceCell->value(), newFormat);
//                currXls.setColumnFormat(column, xlsx->columnFormat(column));
//                currXls.setColumnWidth(column, xlsx->columnWidth(column));
//            }
//        }
//        currXls.setRowFormat(newRow, xlsx->rowFormat(newRow));
//        currXls.setRowHeight(newRow, xlsx->rowHeight(newRow));
//    }
//    currXls.saveAs(xlsName);
    QString sheetName = xlsx->workbook()->activeSheet()->sheetName();
    QStringList sheetNames = currXls.sheetNames();

    for (int i = 0;i<sheetNames.size();i++)
    {
        QString currSheetName = sheetNames.at(i);
        if (sheetName == currSheetName)
        {
            continue;
        }
        currXls.deleteSheet(currSheetName);
    }
    currXls.save();
    requestMsg(Common::MsgTypeSucc, endMsg.arg(QString::number(runnableID)));

    qDebug("EmailSenderRunnable::run end") ;
}

void XlsxParserRunnable::requestMsg(const int msgType, const QString &msg)
{
    QMetaObject::invokeMethod(mParent, "receiveMessage", Qt::QueuedConnection, Q_ARG(int,msgType),Q_ARG(QString, msg));
}


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

