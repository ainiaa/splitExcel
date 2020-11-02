#include "parserbymsrunnable.h"
#include "excelparserbyms.h"

ParserByMSRunnable::ParserByMSRunnable(QObject *parent) {
    this->mParent = parent;
}

ParserByMSRunnable::~ParserByMSRunnable() {
    runnableID = 0;
}

void ParserByMSRunnable::setID(const int &id) {
    runnableID = id;
}

//设置拆分相关数据
void ParserByMSRunnable::setSplitData(SourceData *sourceExcelData,
                                               QString selectedSheetName,
                                               QHash<QString, QList<int>> fragmentDataQhash,
                                               int m_total) {
    qDebug("XlsxParserByOfficeRunnable::setSplitData with SourceExcelData");

    this->sourcePath = sourceExcelData->getSourcePath();
    this->savePath = sourceExcelData->getSavePath();
    this->m_total = m_total;
    this->selectedSheetName = selectedSheetName;
    this->fragmentDataQhash = fragmentDataQhash;
    this->tplXlsPath = sourceExcelData->getTplXlsPath();
    this->sourceRowStart = sourceExcelData->getSourceRowStart();
    this->sourceMaxAlphabetCol = sourceExcelData->getSourceMaxAlphabetCol();
    this->sourceRowCnt = sourceExcelData->getSourceRowCnt();
    this->groupByText = sourceExcelData->getGroupByText();
    this->groupByIndex = sourceExcelData->getGroupByIndex();
    this->dataType = 0;
}

void ParserByMSRunnable::setSplitData(SourceData *sourceExcelData, QString selectedSheetName, QHash<QString, QList<QList<QVariant>>> fragmentDataQhash, int m_total){
    qDebug("XlsxParserByOfficeRunnable::setSplitData with SourceExcelData");

    this->sourcePath = sourceExcelData->getSourcePath();
    this->savePath = sourceExcelData->getSavePath();
    this->m_total = m_total;
    this->selectedSheetName = selectedSheetName;
    this->fragmentDataQhash2 = fragmentDataQhash;
    this->tplXlsPath = sourceExcelData->getTplXlsPath();
    this->sourceRowStart = sourceExcelData->getSourceRowStart();
    this->sourceMaxAlphabetCol = sourceExcelData->getSourceMaxAlphabetCol();
    this->sourceRowCnt = sourceExcelData->getSourceRowCnt();
    this->groupByText = sourceExcelData->getGroupByText();
    this->groupByIndex = sourceExcelData->getGroupByIndex();
    this->dataType = 1;
}


void ParserByMSRunnable::run() {
    if(this->dataType ==0) {
        this->run0();
    } else {
        this->run1();
    }
}

void ParserByMSRunnable::run0() {
    qDebug("XlsxParserByOfficeRunnable::run start");
    QString startMsg("【*__* 开始 ===】拆分excel : %1/%2  分组项 【%3】");
    QString endMsg("【=== 完成  ^_^】拆分excel : %1/%2  分组项 【%3】");

    QHashIterator<QString, QList<int>> it(fragmentDataQhash);
    CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    while (it.hasNext()) {
        it.next();
        QString key = it.key();
        requestMsg(Common::MsgTypeInfo,
                   startMsg.arg(QString::asprintf("%04d", runnableID)).arg(QString::asprintf("%04d", m_total)).arg(key));
        QList<int> contentList = it.value();
        contentList.insert(0, 1);

        this->doProcess(key, contentList);
        requestMsg(Common::MsgTypeSucc,
                   endMsg.arg(QString::asprintf("%04d", runnableID)).arg(QString::asprintf("%04d", m_total)).arg(key));
    }
    CoUninitialize();

    qDebug("XlsxParserRunnable::run end");
}

void ParserByMSRunnable::run1() {
    qDebug("XlsxParserByOfficeRunnable::run start");
    QString startMsg("【*__* 开始 ===】拆分excel : %1/%2  分组项 【%3】");
    QString endMsg("【=== 完成  ^_^】拆分excel : %1/%2  分组项 【%3】");

    QHashIterator<QString, QList<QList<QVariant>>> it(fragmentDataQhash2);
    CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    while (it.hasNext()) {
        it.next();
        QString key = it.key();
        requestMsg(Common::MsgTypeInfo,
                   startMsg.arg(QString::asprintf("%04d", runnableID)).arg(QString::asprintf("%04d", m_total)).arg(key));
        QList<QList<QVariant>> contentList = it.value();
        //contentList.insert(0, 1);

        this->doProcess(key, contentList);
        requestMsg(Common::MsgTypeSucc,
                   endMsg.arg(QString::asprintf("%04d", runnableID)).arg(QString::asprintf("%04d", m_total)).arg(key));
    }
    CoUninitialize();

    qDebug("XlsxParserRunnable::run end");
}

//拆分
void ParserByMSRunnable::doProcess(QString key, QList<int> contentList) {
    QString xlsName;
    xlsName.append(savePath).append(QDir::separator()).append(key).append(".xlsx");
    bool cpret = copyFileToPath(tplXlsPath, xlsName, true);
    if (!cpret) {
        qDebug() << " copyFileToPath failure."
                 << "source:" << savePath << " dist:" << tplXlsPath;
    }
    QAxObject *excel = new QAxObject("Excel.Application");  //连接Excel控件
    excel->dynamicCall("SetVisible (bool Visible)", false); //不显示窗体
    excel->setProperty("DisplayAlerts", false); //不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示
    excel->setProperty("EnableEvents", false); //没有这个 很容易报错  QAxBase: Error calling IDispatch member Open: Unknown error

    qDebug() << "ExcelParserByOfficeRunnable::processByOffice xlsName=" << xlsName;
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
        QAxObject *range = worksheet->querySubObject("Range(QVariant)", deleteRangeContent); //获取单元格
        range->dynamicCall("Delete()");
    }
    // end

    // workbook->dynamicCall("SaveAs(const QString&)", xlsName);//保存至filepath，注意一定要用QDir::toNativeSeparators将路径中的"/"转换为"\"，不然一定保存不了。
    workbook->dynamicCall("Save()");
    workbook->dynamicCall("Close()");  //关闭工作簿
    workbooks->dynamicCall("Close()"); //关闭工作簿
    //    CoUninitialize();
    ParserByMSRunnable::freeExcel(excel);

    this->doFilter(key);
}

//拆分
void ParserByMSRunnable::doProcess(QString key, QList<QList<QVariant>> contentList) {
    QString xlsName;
    xlsName.append(savePath).append(QDir::separator()).append(key).append(".xlsx");
    bool cpret = copyFileToPath(tplXlsPath, xlsName, true);
    if (!cpret) {
        qDebug() << " copyFileToPath failure."
                 << "source:" << savePath << " dist:" << tplXlsPath;
    }
    QAxObject *excel = new QAxObject("Excel.Application");  //连接Excel控件
    excel->dynamicCall("SetVisible (bool Visible)", false); //不显示窗体
    excel->setProperty("DisplayAlerts", false); //不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示
    excel->setProperty("EnableEvents", false); //没有这个 很容易报错  QAxBase: Error calling IDispatch member Open: Unknown error

    qDebug() << "ExcelParserByOfficeRunnable::processByOffice xlsName=" << xlsName;
    QAxObject *workbooks = excel->querySubObject("WorkBooks"); //获取工作簿集合
    QAxObject *workbook = workbooks->querySubObject("Open(const QString&)", xlsName);
    QAxObject *worksheet = workbook->querySubObject("WorkSheets(int)", 1);

    // start

    QString rangeFormat = "A2:%2%3";
    ExcelBase::convertToColName(contentList.at(0).size(), sourceMaxAlphabetCol);
    QString rangeStr = rangeFormat.arg(sourceMaxAlphabetCol).arg(QString::asprintf("%d",contentList.size()+1));
    QAxObject *range = worksheet->querySubObject("Range(QVariant)", rangeStr); //获取单元格
    qDebug() << "rangeStr:" << rangeStr;
    if(nullptr == range || range->isNull()) {
        return;
    }
    bool succ = false;
    QVariant var;
    ExcelBase::castListListVariant2Variant(contentList,var);
    succ = range->setProperty("Value", var);
    if (!succ) { //设置失败，尝试调用SetValue
        range->dynamicCall("SetValue(const QVariant& value)", var);
        qDebug() << "range->setProperty(\"Value\", var) success";
    } else {
        qDebug() << "range->setProperty(\"Value\", var) falure";
    }
    // end

    // workbook->dynamicCall("SaveAs(const QString&)", xlsName);//保存至filepath，注意一定要用QDir::toNativeSeparators将路径中的"/"转换为"\"，不然一定保存不了。
    workbook->dynamicCall("Save()");
    workbook->dynamicCall("Close()");  //关闭工作簿
    workbooks->dynamicCall("Close()"); //关闭工作簿
    //    CoUninitialize();
    delete range;
    ParserByMSRunnable::freeExcel(excel);

}


void ParserByMSRunnable::doFilter(QString key) {
    QString xlsName;
    xlsName.append(savePath).append(QDir::separator()).append(key).append(".xlsx");
    QAxObject *excel = new QAxObject("Excel.Application");  //连接Excel控件
    excel->dynamicCall("SetVisible (bool Visible)", false); //不显示窗体
    excel->setProperty("DisplayAlerts", false); //不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示
    excel->setProperty("EnableEvents", false); //没有这个 很容易报错  QAxBase: Error calling IDispatch member Open: Unknown error

    qDebug() << "ExcelParserByOfficeRunnable::processByOffice xlsName=" << xlsName;
    QAxObject *workbooks = excel->querySubObject("WorkBooks"); //获取工作簿集合
    QAxObject *workbook = workbooks->querySubObject("Open(const QString&, QVariant)", xlsName, 0);
    QAxObject *worksheet = workbook->querySubObject("WorkSheets(int)", 1);

    workbooks = excel->querySubObject("WorkBooks"); //获取工作簿集合
    workbook = workbooks->querySubObject("Open(const QString&, QVariant)", xlsName, 0);
    worksheet = workbook->querySubObject("WorkSheets(int)", 1);
    QAxObject *usedRange = worksheet->querySubObject("UsedRange"); // sheet范围
    QAxObject *rows, *columns;
    rows = usedRange->querySubObject("Rows");       // 行
    columns = usedRange->querySubObject("Columns"); // 列
    int leftRowCount = rows->property("Count").toInt();

    qDebug() << key << "row:" << leftRowCount << "  col:" << columns << "\r\n";

    QString deleteRangeFormat = "A%1:%2%3";
    QString currentCellFormat = "%1,%2";
    int groupByColNum = groupByIndex + 1;
    for (int i = 2; i <= leftRowCount; i++) {
        QAxObject *cell = worksheet->querySubObject("Cells(int, int)", i, groupByColNum);
        QString cellValue = cell->dynamicCall("Value()").toString();
        QString currentCell = currentCellFormat.arg(i).arg(groupByColNum);
        qDebug() << key << " " << currentCell << " groupBy:" << key << "  current:" << cellValue;
        if (cellValue.isEmpty()) { //已经到空行了 直接结束
            qDebug() << key << i << " is empty\r\n";
            break;
        }
        cellValue = cellValue.trimmed();
        if (cellValue != key) { //当前列不是所需数据（漏网之鱼）删除
            QString deleteRangeContent = deleteRangeFormat.arg(i).arg(sourceMaxAlphabetCol).arg(i);
            QAxObject *range = worksheet->querySubObject("Range(QVariant)", deleteRangeContent); //获取单元格
            range->dynamicCall("Delete()");
            qDebug() << key << " (" << deleteRangeContent << " -- 【" << currentCell << ":" << cellValue << "】) is deleted";
        }
    }
    qDebug() << "xlsName" << key << "rows:" << rows->property("Count").toInt() << "  columns:" << columns->property("Count").toInt();
    workbook->dynamicCall("Save()");
    workbook->dynamicCall("Close()");  //关闭工作簿
    workbooks->dynamicCall("Close()"); //关闭工作簿
    ParserByMSRunnable::freeExcel(excel);
}

void ParserByMSRunnable::requestMsg(const int msgType, const QString &msg) {
    qobject_cast<ExcelParser *>(mParent)->receiveMessage(msgType, msg);
    // QMetaObject::invokeMethod(mParent, "receiveMessage", Qt::QueuedConnection, Q_ARG(int,msgType),Q_ARG(QString, msg));//不能及时返回信息
}

bool ParserByMSRunnable::copyFileToPath(QString sourceDir, QString toDir, bool coverFileIfExist) {
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

void ParserByMSRunnable::processSourceFile() {
    CoInitializeEx(nullptr, COINIT_MULTITHREADED);
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
    ParserByMSRunnable::freeExcel(excel);
    this->generateTplXls(); //生成模板文件
}

void ParserByMSRunnable::processSourceFile(SourceData *sourceData, QString selectedSheetName) {
    CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    QAxObject *excel = new QAxObject("Excel.Application");  //连接Excel控件
    excel->dynamicCall("SetVisible (bool Visible)", false); //不显示窗体
    excel->setProperty("DisplayAlerts", false); //不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示
    excel->setProperty("EnableEvents", false); //没有这个 很容易报错  QAxBase: Error calling IDispatch member Open: Unknown error
    qDebug() << "ExcelParserByOfficeRunnable::processSourceFile:: with SourceExcelData: " << sourceData->getSourcePath();
    qDebug() << " new source path";

    QAxObject *workbooks = excel->querySubObject("WorkBooks"); //获取工作簿集合
    QAxObject *workbook = workbooks->querySubObject("Open(const QString&, QVariant)", sourceData->getSourcePath(), 0);
    QAxObject *worksheets = workbook->querySubObject("WorkSheets"); // 获取打开的excel文件中所有的工作sheet
    sourceData->setSheetCnt(worksheets->property("Count").toInt());

    qDebug() << QString("Excel文件中表的个数: %1").arg(QString::number(sourceData->getSheetCnt()));
    QAxObject *worksheet = nullptr;
    int selectedSheetIndex = 0;
    for (int i = 1; i <= sourceData->getSheetCnt(); i++) {
        QAxObject *currentWorkSheet = workbook->querySubObject("WorkSheets(int)", i); // Sheets(int)也可换用Worksheets(int)
        QString currentWorkSheetName = currentWorkSheet->property("Name").toString(); //获取工作表名称
        QString message = QString("sheet ") + QString::number(i, 10) + QString(" name");
        qDebug() << message << currentWorkSheetName;
        qDebug() << "this->selectedSheetName:" << selectedSheetName;
        if (currentWorkSheetName == selectedSheetName) {
            worksheet = currentWorkSheet;
            selectedSheetIndex = i;
            break;
        }
    }

    if (worksheet == nullptr) {
        qDebug() << "worksheet is null";
        return;
    }

    QAxObject *usedRange = worksheet->querySubObject("UsedRange");             // sheet范围
    sourceData->setSourceRowStart(usedRange->property("Row").toInt());    // 起始行数
    sourceData->setSourceColStart(usedRange->property("Column").toInt()); // 起始列数
    QAxObject *rows, *columns;
    rows = usedRange->querySubObject("Rows");       // 行
    columns = usedRange->querySubObject("Columns"); // 列

    sourceData->setSourceRowCnt(rows->property("Count").toInt());    // 行数
    sourceData->setSourceColCnt(columns->property("Count").toInt()); // 列数
    QString sourceMinAlphabetCol;
    QString sourceMaxAlphabetCol;
    ExcelBase::convertToColName(sourceData->getSourceRowStart(), sourceMinAlphabetCol);
    ExcelBase::convertToColName(sourceData->getSourceColCnt(), sourceMaxAlphabetCol);
    sourceData->setSourceMinAlphabetCol(sourceMinAlphabetCol);
    sourceData->setSourceMaxAlphabetCol(sourceMaxAlphabetCol);

    qDebug() << " sourceRowStart:" << sourceData->getSourceRowStart() << " sourceColStart:" << sourceData->getSourceColCnt()
             << " sourceRowCnt:" << sourceData->getSourceRowCnt() << " sourceColCnt:" << sourceData->getSourceColCnt()
             << " sourceMinAlphabetCol:" << sourceMinAlphabetCol << " sourceMaxAlphabetCol:" << sourceMaxAlphabetCol;
    // CoUninitialize();
    ParserByMSRunnable::freeExcel(excel);
    ParserByMSRunnable::generateTplXls(sourceData, selectedSheetIndex); //生成模板文件
}
void ParserByMSRunnable::freeExcel(QAxObject *excel) {
    if (!excel->isNull()) {
        excel->dynamicCall("Quit()");
        delete excel;
        excel = nullptr;
    }
}

void ParserByMSRunnable::generateTplXls(SourceData *sourceExcelData, QString selectedSheetName) {
    QString xlsName;
    xlsName.append(sourceExcelData->getSavePath()).append(QDir::separator()).append("sourceTplXlsx.xlsx");
    copyFileToPath(sourceExcelData->getSourcePath(), xlsName, true);

    CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    QAxObject *excel = new QAxObject("Excel.Application");  //连接Excel控件
    excel->dynamicCall("SetVisible (bool Visible)", false); //不显示窗体
    excel->setProperty("DisplayAlerts", false); //不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示
    excel->setProperty("EnableEvents", false); //没有这个 很容易报错  QAxBase: Error calling IDispatch member Open: Unknown error

    qDebug() << "ExcelParserByOfficeRunnable::generateTplXls:: with SourceExcelData: xlsName=" << xlsName;

    QAxObject *workbooks = excel->querySubObject("WorkBooks"); //获取工作簿集合
    QAxObject *workbook = workbooks->querySubObject("Open(const QString&, QVariant)", xlsName, 0);
    int sourceWorkSheetCnt = sourceExcelData->getSheetCnt();
    for (int i = 1; i <= sourceWorkSheetCnt; i++) {
        QAxObject *currentWorkSheet = workbook->querySubObject("WorkSheets(int)", i);
        QString currentWorkSheetName = currentWorkSheet->property("Name").toString(); //获取工作表名称
        if (currentWorkSheetName != selectedSheetName) {                              //删除sheet
            currentWorkSheet->dynamicCall("Delete()");
        }
    }


    workbook->dynamicCall("Save()");
    workbook->dynamicCall("Close()");  //关闭工作簿
    workbooks->dynamicCall("Close()"); //关闭工作簿
    // CoUninitialize();
    ParserByMSRunnable::freeExcel(excel);
    sourceExcelData->setTplXlsPath(xlsName);
}

void ParserByMSRunnable::generateTplXls(SourceData *sourceData, int selectedSheetIndex) {
    QString xlsName;
    xlsName.append(sourceData->getSavePath()).append(QDir::separator()).append("sourceTplXlsx.xlsx");
    copyFileToPath(sourceData->getSourcePath(), xlsName, true);

    CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    QAxObject *excel = new QAxObject("Excel.Application");  //连接Excel控件
    excel->dynamicCall("SetVisible (bool Visible)", false); //不显示窗体
    excel->setProperty("DisplayAlerts", false); //不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示
    excel->setProperty("EnableEvents", false); //没有这个 很容易报错  QAxBase: Error calling IDispatch member Open: Unknown error

    qDebug() << "ExcelParserByOfficeRunnable::generateTplXls:: with SourceExcelData: xlsName=" << xlsName;

    QAxObject *workbooks = excel->querySubObject("WorkBooks"); //获取工作簿集合
    QAxObject *workbook = workbooks->querySubObject("Open(const QString&, QVariant)", xlsName, 0);
    int sourceWorkSheetCnt = sourceData->getSheetCnt();
    for (int i = 1; i <= sourceWorkSheetCnt; i++) {
        QAxObject *currentWorkSheet = workbook->querySubObject("WorkSheets(int)", i); // Sheets(int)也可换用Worksheets(int)
        QString currentWorkSheetName = currentWorkSheet->property("Name").toString(); //获取工作表名称
        if (i != selectedSheetIndex) {                                                //删除sheet
            currentWorkSheet->dynamicCall("Delete()");
        } else {
            //删除除了表头之外的所有数据
            QAxObject *usedRange = currentWorkSheet->querySubObject("UsedRange");
            QAxObject *rows, *columns;
            int rowCnt = 0;
            int colCnt = 0;
            rows = usedRange->querySubObject("Rows");       // 行
            columns = usedRange->querySubObject("Columns"); // 列
            rowCnt = rows->property("Count").toInt();// 行数
            colCnt = columns->property("Count").toInt();// 列数
            QString sourceMaxAlphabetCol;

            ExcelBase::convertToColName(colCnt, sourceMaxAlphabetCol);

            // 需要删除的 range 为 A2:$sourceMaxAlphabetCol$rowCnt
            QString rangef("A2:%1%2");
            QString deleteRangeContent;
            deleteRangeContent = rangef.arg(sourceMaxAlphabetCol).arg(QString::asprintf("%d",rowCnt));
            QAxObject *range = currentWorkSheet->querySubObject("Range(QVariant)", deleteRangeContent); //获取单元格
            range->dynamicCall("Delete()");
            delete range;
        }
    }


    //排序
    //    QString groupByFieldNum;
    //    ExcelBase::convertToColName(sourceExcelData->getGroupByIndex() + 1, groupByFieldNum);
    //    QString rangFormat = "Range(A1:%1%2)";
    //    QString rang = rangFormat.arg(sourceExcelData->getSourceMaxAlphabetCol()).arg(sourceExcelData->getSourceRowCnt());
    //    QString sortFieldFormat = "Range(%1%2)";
    //    QString sortField = sortFieldFormat.arg(groupByFieldNum).arg("1");
    //    qDebug() << "rang:" << rang << "      sortField:" << sortField << " abcol:" << sourceExcelData->getSourceMaxAlphabetCol()
    //             << " col:" << sourceExcelData->getSourceColCnt();
    //    QAxObject *currentWorkSheet = workbook->querySubObject("WorkSheets(int)", selectedSheetIndex); // Sheets(int)也可换用Worksheets(int)
    //    currentWorkSheet->querySubObject(rang.toUtf8())
    //    ->dynamicCall("Sort(Key1:=QAxObject*,int)", currentWorkSheet->querySubObject(sortField.toUtf8())->asVariant(), 2);

    workbook->dynamicCall("Save()");
    workbook->dynamicCall("Close()");  //关闭工作簿
    workbooks->dynamicCall("Close()"); //关闭工作簿
    // CoUninitialize();
    ParserByMSRunnable::freeExcel(excel);
    sourceData->setTplXlsPath(xlsName);
}

void ParserByMSRunnable::generateTplXls() {
    QString xlsName;
    xlsName.append(savePath).append(QDir::separator()).append("sourceTplXlsx.xlsx");
    copyFileToPath(sourcePath, xlsName, true);

    CoInitializeEx(nullptr, COINIT_MULTITHREADED);
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
    // CoUninitialize();
    ParserByMSRunnable::freeExcel(excel);
    this->tplXlsPath = xlsName;
}

void ParserByMSRunnable::setSplitData(QString sourcePath,
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
