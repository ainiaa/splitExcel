#include "excelparserbyofficerunnable.h"
#include "excelparserbyoffice.h"

ExcelParserByOfficeRunnable::ExcelParserByOfficeRunnable(QObject *parent) {
    this->mParent = parent;
}

ExcelParserByOfficeRunnable::~ExcelParserByOfficeRunnable() {
    runnableID = 0;
}

void ExcelParserByOfficeRunnable::setID(const int &id) {
    runnableID = id;
}

void ExcelParserByOfficeRunnable::setSplitData(SourceExcelData *sourceXmlData,
                                               QString selectedSheetName,
                                               QHash<QString, QList<int>> fragmentDataQhash,
                                               int m_total) {
    qDebug("XlsxParserByOfficeRunnable::setSplitData w ith SourceExcelData");

    this->sourcePath = sourceXmlData->getSourcePath();
    this->savePath = sourceXmlData->getSavePath();
    this->m_total = m_total;
    this->selectedSheetName = selectedSheetName;
    this->fragmentDataQhash = fragmentDataQhash;
    this->processSourceFile(); //处理源文件
}

void ExcelParserByOfficeRunnable::run() {
    qDebug("XlsxParserByOfficeRunnable::run start");
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

        this->processByOffice(key, contentList);
    }
    requestMsg(Common::MsgTypeSucc, endMsg.arg(QString::number(runnableID)));
    qDebug("XlsxParserRunnable::run end");
}

void ExcelParserByOfficeRunnable::processByOffice(QString key, QList<int> contentList) {
    QString xlsName;
    xlsName.append(savePath).append(QDir::separator()).append(key).append(".xlsx");
    copyFileToPath(tplXlsPath, xlsName, true);
    HRESULT hres = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    QAxObject *excel = new QAxObject("Excel.Application");  //连接Excel控件
    excel->dynamicCall("SetVisible (bool Visible)", false); //不显示窗体
    excel->setProperty("DisplayAlerts", false); //不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示
    excel->setProperty("EnableEvents", false); //没有这个 很容易报错  QAxBase: Error calling IDispatch member Open: Unknown error

    QAxObject *workbooks = excel->querySubObject("WorkBooks"); //获取工作簿集合
    QAxObject *workbook = workbooks->querySubObject("Open(const QString&, QVariant)", xlsName, 0);
    QAxObject *worksheet = workbook->querySubObject("WorkSheets(int)", 1);

    // start
    QList<QString> deleteRange;
    int intContentList = contentList.size();

    QString deleteRangeFormat = "A%1:%2%3";
    int latestStart = sourceRowStart;
    for (int i = 0; i < intContentList; i++) {
        int max = contentList.takeAt(0);
        if (latestStart == max) {
            latestStart = max + 1;
            continue;
        } else {
            deleteRange << deleteRangeFormat.arg(latestStart).arg(sourceMaxAlphabetCol).arg(max - 1);
            latestStart = max + 1;
        }
    }
    if (latestStart < sourceRowCnt) {
        deleteRange << deleteRangeFormat.arg(latestStart).arg(sourceMaxAlphabetCol).arg(sourceRowCnt);
    }
    for (int i = deleteRange.size() - 1; i >= 0; i--) {
        QString deleteRangeContent = deleteRange.at(i);
        // qDebug() << " deleteRangeContent["<< i <<"] = " << deleteRangeContent;
        QAxObject *range = worksheet->querySubObject("Range(QVariant)", deleteRangeContent); //获取单元格
        range->dynamicCall("Delete()");
        // qDebug() << " deleteRangeContent:" << deleteRangeContent;
    }
    // end

    // workbook->dynamicCall("SaveAs(const QString&)", xlsName);//保存至filepath，注意一定要用QDir::toNativeSeparators将路径中的"/"转换为"\"，不然一定保存不了。
    workbook->dynamicCall("Save()");
    workbook->dynamicCall("Close()");  //关闭工作簿
    workbooks->dynamicCall("Close()"); //关闭工作簿
    excel->dynamicCall("Quit()");      //关闭excel
    delete excel;
    excel = nullptr;
    // CoUninitialize();
}

void ExcelParserByOfficeRunnable::requestMsg(const int msgType, const QString &msg) {
    qobject_cast<ExcelParser *>(mParent)->receiveMessage(msgType, msg.arg(msgType));
    // QMetaObject::invokeMethod(mParent, "receiveMessage", Qt::QueuedConnection, Q_ARG(int,msgType),Q_ARG(QString, msg));//不能及时返回信息
}

bool ExcelParserByOfficeRunnable::copyFileToPath(QString sourceDir, QString toDir, bool coverFileIfExist) {
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

void ExcelParserByOfficeRunnable::processSourceFile() {
    HRESULT hres = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    QAxObject *excel = new QAxObject("Excel.Application");  //连接Excel控件
    excel->dynamicCall("SetVisible (bool Visible)", false); //不显示窗体
    excel->setProperty("DisplayAlerts", false); //不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示
    excel->setProperty("EnableEvents", false); //没有这个 很容易报错  QAxBase: Error calling IDispatch member Open: Unknown error
    qDebug() << "this->sourcePath" << this->sourcePath;
    QAxObject *workbooks = excel->querySubObject("WorkBooks"); //获取工作簿集合
    QAxObject *workbook = workbooks->querySubObject("Open(const QString&, QVariant)", this->sourcePath, 0);
    QAxObject *worksheets = workbook->querySubObject("WorkSheets"); // 获取打开的excel文件中所有的工作sheet
    this->sourceWorkSheetCnt = worksheets->property("Count").toInt();

    qDebug() << QString("Excel文件中表的个数: %1").arg(QString::number(sourceWorkSheetCnt));
    QAxObject *worksheet = nullptr;
    for (int i = 1; i <= sourceWorkSheetCnt; i++) {
        QAxObject *currentWorkSheet = workbook->querySubObject("WorkSheets(int)", i); // Sheets(int)也可换用Worksheets(int)
        QString currentWorkSheetName = currentWorkSheet->property("Name").toString(); //获取工作表名称
        QString message = QString("sheet ") + QString::number(i, 10) + QString(" name");
        qDebug() << message << currentWorkSheetName;
        qDebug() << "this->selectedSheetName:" << this->selectedSheetName;
        if (currentWorkSheetName == this->selectedSheetName) {
            worksheet = currentWorkSheet;
            this->selectedSheetIndex = i;
            break;
        }
    }

    if (worksheet == nullptr) {
        qDebug() << "worksheet is null";
        return;
    }

    QAxObject *usedRange = worksheet->querySubObject("UsedRange"); // sheet范围
    this->sourceRowStart = usedRange->property("Row").toInt();     // 起始行数
    this->sourceColStart = usedRange->property("Column").toInt();  // 起始列数
    QAxObject *rows, *columns;
    rows = usedRange->querySubObject("Rows");       // 行
    columns = usedRange->querySubObject("Columns"); // 列

    this->sourceRowCnt = rows->property("Count").toInt();    // 行数
    this->sourceColCnt = columns->property("Count").toInt(); // 列数

    ExcelBase::convertToColName(sourceRowStart, sourceMinAlphabetCol);
    ExcelBase::convertToColName(sourceColCnt, sourceMaxAlphabetCol);

    qDebug() << " sourceRowStart:" << sourceRowStart << " sourceColStart:" << sourceColStart << " sourceRowCnt:" << sourceRowCnt
             << " sourceColCnt:" << sourceColCnt << " sourceMinAlphabetCol:" << sourceMinAlphabetCol
             << " sourceMaxAlphabetCol:" << sourceMaxAlphabetCol;
    // CoUninitialize();
    this->generateTplXls(); //生成模板文件
}

void ExcelParserByOfficeRunnable::generateTplXls() {
    QString xlsName;
    xlsName.append(savePath).append(QDir::separator()).append("sourceTplXlsx.xlsx");
    copyFileToPath(sourcePath, xlsName, true);

    HRESULT hres = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    QAxObject *excel = new QAxObject("Excel.Application");  //连接Excel控件
    excel->dynamicCall("SetVisible (bool Visible)", false); //不显示窗体
    excel->setProperty("DisplayAlerts", false); //不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示
    excel->setProperty("EnableEvents", false); //没有这个 很容易报错  QAxBase: Error calling IDispatch member Open: Unknown error

    QAxObject *workbooks = excel->querySubObject("WorkBooks"); //获取工作簿集合
    QAxObject *workbook = workbooks->querySubObject("Open(const QString&, QVariant)", xlsName, 0);
    for (int i = 1; i <= sourceWorkSheetCnt; i++) {
        if (i != selectedSheetIndex) {                                                    //删除sheet
            QAxObject *currentWorkSheet = workbook->querySubObject("WorkSheets(int)", i); // Sheets(int)也可换用Worksheets(int)
            currentWorkSheet->dynamicCall("Delete()");
        }
    }

    workbook->dynamicCall("Save()");
    workbook->dynamicCall("Close()");  //关闭工作簿
    workbooks->dynamicCall("Close()"); //关闭工作簿
    excel->dynamicCall("Quit()");      //关闭excel
    delete excel;
    excel = nullptr;
    // CoUninitialize();
    this->tplXlsPath = xlsName;
}

void ExcelParserByOfficeRunnable::setSplitData(QString sourcePath,
                                               QString selectedSheetName,
                                               QHash<QString, QList<int>> fragmentDataQhash,
                                               QString savePath,
                                               int m_total) {
    this->sourcePath = sourcePath;
    this->selectedSheetName = selectedSheetName;
    this->fragmentDataQhash = fragmentDataQhash;
    this->savePath = savePath;
    this->m_total = m_total;
}
