#include "excelparserbylibrunnable.h"
#include "excelparser.h"

ExcelParserByLibRunnable::ExcelParserByLibRunnable(QObject *parent) {
    this->mParent = parent;
}

ExcelParserByLibRunnable::~ExcelParserByLibRunnable() {
    runnableID = 0;
}

void ExcelParserByLibRunnable::setID(const int &id) {
    runnableID = id;
}

void ExcelParserByLibRunnable::setSplitData(QString sourcePath,
                                            QString selectedSheetName,
                                            QHash<QString, QList<int>> fragmentDataQhash,
                                            QString savePath,
                                            int total) {
    qDebug("setSplitData");
    this->xlsx = new QXlsx::Document(sourcePath);
    this->sourcePath = sourcePath;
    this->selectedSheetName = selectedSheetName;
    this->fragmentDataQhash = fragmentDataQhash;
    this->savePath = savePath;
    this->m_total = total;
}

void ExcelParserByLibRunnable::run() {
    qDebug("XlsxParserRunnable::run start");
    QString startMsg("开始拆分excel并生成新的excel文件: %1/");
    QString endMsg("完成拆分excel并生成新的excel文件: %1/");
    startMsg.append(QString::number(m_total));
    endMsg.append(QString::number(m_total));
    requestMsg(Common::MsgTypeInfo, startMsg.arg(QString::number(runnableID)));
    QHashIterator<QString, QList<int>> it(fragmentDataQhash);

    while (it.hasNext()) {
        it.next();
        QString key = it.key();
        QList<int> contentList = it.value();
        contentList.insert(0, 1);

        this->processByQxls(key, contentList);
    }
    requestMsg(Common::MsgTypeSucc, endMsg.arg(QString::number(runnableID)));
    qDebug("XlsxParserRunnable::run end");
}

void ExcelParserByLibRunnable::processByQxls(QString key, QList<int> contentList) {
    QString xlsName;
    xlsName.append(savePath).append(QDir::separator()).append(key).append(".xlsx");
    copyFileToPath(sourcePath, xlsName, true);

    QXlsx::Document *currXls = new QXlsx::Document(xlsName);

    currXls->fliterRows(contentList); //删除多余的行  多线程的生活会出现问题。

    //删除多余的sheet start
    QStringList sheetNames = currXls->sheetNames();
    for (int i = 0; i < sheetNames.size(); i++) {
        QString currSheetName = sheetNames.at(i);
        if (selectedSheetName == currSheetName) {
            continue;
        }
        currXls->deleteSheet(currSheetName);
    }
    //删除多余的sheet end

    //尝试修复行高不正确的问题
    QXlsx::CellRange cellRange = currXls->dimension();
    for (int row_num = cellRange.firstRow(); row_num <= cellRange.lastRow(); row_num++) {
        double rowHeigh = currXls->rowHeight(row_num);
        if (rowHeigh < 14) {
            currXls->setRowHeight(row_num, 18);
        }
    }

    currXls->save();
    delete currXls;
}

QXlsx::Format ExcelParserByLibRunnable::copyFormat(QXlsx::Format format) {
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

void ExcelParserByLibRunnable::writeXlsHeader(QXlsx::Document *xls, QString selectedSheetName) {
    QXlsx::CellRange range;
    xlsx->selectSheet(selectedSheetName);
    range = xlsx->dimension();
    int colCount = range.columnCount();

    for (int column = 1; column <= colCount; ++column) {
        QXlsx::Cell *sourceCell = xlsx->cellAt(1, column);
        if (sourceCell) {
            QString columnName;
            OfficeHelper::convertToColName(column, columnName);
            QString cell;
            QXlsx::Format newFormat = copyFormat(sourceCell->format());
            cell.append(columnName).append(QString::number(1));
            xls->write(cell, sourceCell->value(), newFormat);
            xls->setColumnFormat(column, xlsx->columnFormat(column));
            xls->setColumnWidth(column, xlsx->columnWidth(column));
        }
    }
    xls->setRowFormat(1, xlsx->rowFormat(1));
    xls->setRowHeight(1, xlsx->rowHeight(1));
}

// old end

void ExcelParserByLibRunnable::requestMsg(const int msgType, const QString &msg) {
    qobject_cast<ExcelParser *>(mParent)->receiveMessage(msgType, msg.arg(msgType));
    // QMetaObject::invokeMethod(mParent, "receiveMessage", Qt::QueuedConnection, Q_ARG(int,msgType),Q_ARG(QString, msg));//不能及时返回信息
}

bool ExcelParserByLibRunnable::copyFileToPath(QString sourceDir, QString toDir, bool coverFileIfExist) {
    sourceDir = QDir::toNativeSeparators(sourceDir);
    toDir = QDir::toNativeSeparators(toDir);

    if (sourceDir == toDir) {
        return true;
    }
    if (!QFile::exists(sourceDir)) {
        return false;
    }
    QDir *createfile = new QDir();
    bool exist = createfile->exists(toDir);
    if (exist) {
        if (coverFileIfExist) {
            createfile->remove(toDir);
        }
    } // end if

    if (!QFile::copy(sourceDir, toDir)) {
        return false;
    }
    return true;
}

void ExcelParserByLibRunnable::processSourceFile() {
}
void ExcelParserByLibRunnable::generateTplXls() {
}
