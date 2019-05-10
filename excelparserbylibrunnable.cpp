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

void ExcelParserByLibRunnable::setSplitData(SourceExcelData *sourceExcelData,
                                            QString selectedSheetName,
                                            QHash<QString, QList<int>> fragmentDataQhash,
                                            int m_total) {
    qDebug("ExcelParserByLibRunnable::setSplitData with SourceExcelData");

    this->xlsx = new QXlsx::Document(sourcePath);
    this->sourcePath = sourceExcelData->getSourcePath();
    this->savePath = sourceExcelData->getSavePath();
    this->m_total = m_total;
    this->selectedSheetName = selectedSheetName;
    this->fragmentDataQhash = fragmentDataQhash;
    this->sourceExcelData = sourceExcelData;
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

// old end

void ExcelParserByLibRunnable::requestMsg(const int msgType, const QString &msg) {
    qobject_cast<ExcelParser *>(mParent)->receiveMessage(msgType, msg);
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
