#include "splitonlywindow.h"
#include "ui_splitonlywindow.h"

SplitOnlyWindow::SplitOnlyWindow(QMainWindow *parent) : SplitSubWindow(parent), ui(new Ui::SplitOnlyWindow) {
    ui->setupUi(this);

    setFixedSize(this->width(), this->height());

    excelParser = new ExcelParser();
    connect(ui->actionConfig_Setting, SIGNAL(triggered()), this, SLOT(showConfigSetting()));
}

SplitOnlyWindow::~SplitOnlyWindow() {
    delete ui;
}

void SplitOnlyWindow::on_selectFilePushButton_clicked() {
    QString path = excelParser->openFile(this);
    if (path.length() > 0) { //选择了excel文件
        this->sourcePath = QDir::toNativeSeparators(path);
        ui->xlsObjLineEdit->setText(path);

        //获得所有的sheets
        QStringList sheetNames = excelParser->getSheetNames();
        ui->dataComboBox->addItems(sheetNames);
        ui->dataComboBox->setCurrentIndex(0);

        //获得表头数据
        changeGroupby(sheetNames.at(0));
    }
}

//修改分组下拉选择框的选项
void SplitOnlyWindow::changeGroupby(QString selectedSheetName) {
    header->clear();
    ui->groupByComboBox->clear();
    header = excelParser->getSheetHeader(selectedSheetName);
    ui->groupByComboBox->addItems(*header);
}

void SplitOnlyWindow::on_savePathPushButton_clicked() {
    QString path = QFileDialog::getExistingDirectory(this, tr("Save New Excels"), ".");
    if (path.length() > 0) {
        ui->savePathLineEdit->setText(path);
        savePath = QDir::toNativeSeparators(path);
    }
}

void SplitOnlyWindow::on_splitOnlySubmitPushButton_clicked() {
    ui->splitOnlySubmitPushButton->setDisabled(true);

    if (ui->xlsObjLineEdit->text().isEmpty()) {
        QMessageBox::information(this, "Save Path Error", "请选择待拆分的excel");
        ui->splitOnlySubmitPushButton->setDisabled(false);
        return;
    }
    if (savePath.length() == 0) {
        QMessageBox::information(this, "Save Path Error", "请选择保存路径");
        ui->splitOnlySubmitPushButton->setDisabled(false);
        return;
    }

    if (processWindow == nullptr) {
        processWindow = new ProcessWindow();
    } else {
        processWindow->clearProcessText();
    }

    processWindow->setWindowModality(Qt::WindowModality::ApplicationModal);
    processWindow->show();

    //拆分
    doSplitXls(ui->dataComboBox->currentText(), savePath);
}

//拆分
void SplitOnlyWindow::doSplitXls(QString dataSheetName, QString savePath) {
    if (excelParserThread != nullptr) {
        return;
    }
    QString groupByText = ui->groupByComboBox->currentText(); //分组列 内容
    int groupByIndex = ui->groupByComboBox->currentIndex();   //分组列 索引（从 0 开始)
    int dataSheetIndex = ui->dataComboBox->currentIndex();    //数据所在 sheet 索引 （从 0 开始)
    int sheetCnt = ui->dataComboBox->count();                 //总sheet个数

    //设置excel相关数据
    SourceExcelData *sourceExcelData = new SourceExcelData();
    sourceExcelData->setSavePath(savePath);
    sourceExcelData->setDataSheetName(dataSheetName);
    sourceExcelData->setDataSheetIndex(dataSheetIndex);
    sourceExcelData->setSheetCnt(sheetCnt);
    sourceExcelData->setOpType(SourceExcelData::OperateType::SplitOnlyType);
    sourceExcelData->setGroupByText(groupByText);
    sourceExcelData->setGroupByIndex(groupByIndex);
    sourceExcelData->setSourcePath(this->sourcePath);

    //设置excel 解析相关事件
    excelParserThread = new QThread();
    excelParser->setSplitData(cfg, sourceExcelData);
    excelParser->moveToThread(excelParserThread);
    connect(excelParserThread, &QThread::finished, excelParserThread, &QObject::deleteLater);
    connect(excelParserThread, &QThread::finished, excelParser, &QObject::deleteLater);
    connect(this, &SplitOnlyWindow::doProcess, excelParser, &ExcelParser::doParse);         //触发excel解析
    connect(excelParser, &ExcelParser::requestMsg, this, &SplitOnlyWindow::receiveMessage); //传递消息
    excelParserThread->start();                                                             //开启线程

    emit doProcess(); //主线程通过信号换起子线程的槽函数
}

//接受子线程的消息
void SplitOnlyWindow::receiveMessage(const int msgType, const QString &msg) {
    qDebug() << "SplitOnlyWindow::receiveMessage msgType:" << QString::number(msgType).toUtf8() << " msg:" << msg;
    int percent = 0;
    switch (msgType) {
        case Common::MsgTypeStart: //处理开始
            qDebug() << "Common::MsgTypeStart total_cnt: " << msg;
            this->setTotalCnt(msg.toInt());
            processWindow->startTimer();
            break;
        case Common::MsgTypeError: //发生错误
            ui->splitOnlySubmitPushButton->setDisabled(false);
            processWindow->setProcessText(msg);
            processWindow->setProcessPercent(this->incrAndCalcPercent(1));
            QMessageBox::critical(this, "Error", msg);
            break;
        case Common::MsgTypeWriteXlsxFinish: //新excel写入完成
            ui->splitOnlySubmitPushButton->setDisabled(false);
            processWindow->setProcessText(msg);
            processWindow->setProcessPercent(100);
            processWindow->stopTimer();
            break;
        case Common::MsgTypeSucc: //处理成功
            processWindow->setProcessText(msg);
            percent = this->incrAndCalcPercent(1);
            qDebug() << "Common::MsgTypeSucc percent: " << percent;
            processWindow->setProcessPercent(percent);
            break;
        case Common::MsgTypeInfo: //一般消息
        case Common::MsgTypeWarn: //警告消息
        default:
            processWindow->setProcessText(msg);
            break;
    }
}

//返回
void SplitOnlyWindow::on_gobackPushButton_clicked() {
    this->getMainWindow()->show();
    this->hide();
}

//显示配置UI
void SplitOnlyWindow::showConfigSetting() {
    this->hide();
    configSetting->show();
}

void SplitOnlyWindow::on_gobackPushButton_2_clicked() {
    if (excelParserThread != nullptr) {
        return;
    }
    QString groupByText = ui->groupByComboBox->currentText();    //分组列 内容
    int groupByIndex = ui->groupByComboBox->currentIndex();      //分组列 索引（从 0 开始)
    int dataSheetIndex = ui->dataComboBox->currentIndex();       //数据所在 sheet 索引 （从 0 开始)
    QString selectedSheetName = ui->dataComboBox->currentText(); //数据所在 sheet名称
    int sheetCnt = ui->dataComboBox->count();                    //总sheet个数

    //设置excel相关数据
    SourceExcelData *sourceExcelData = new SourceExcelData();
    sourceExcelData->setSavePath(savePath);
    sourceExcelData->setDataSheetName(selectedSheetName);
    sourceExcelData->setDataSheetIndex(dataSheetIndex);
    sourceExcelData->setSheetCnt(sheetCnt);
    sourceExcelData->setOpType(SourceExcelData::OperateType::SplitOnlyType);
    sourceExcelData->setGroupByText(groupByText);
    sourceExcelData->setGroupByIndex(groupByIndex);
    sourceExcelData->setSourcePath(this->sourcePath);

    CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    QAxObject *excel = new QAxObject("Excel.Application");  //连接Excel控件
    excel->dynamicCall("SetVisible (bool Visible)", false); //不显示窗体
    excel->setProperty("DisplayAlerts", false); //不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示
    excel->setProperty("EnableEvents", false); //没有这个 很容易报错  QAxBase: Error calling IDispatch member Open: Unknown error
    qDebug() << "ExcelParserByOfficeRunnable::processSourceFile:: with SourceExcelData: " << sourceExcelData->getSourcePath();
    qDebug() << " new source path"
             << "";
    QAxObject *workbooks = excel->querySubObject("WorkBooks"); //获取工作簿集合
    QAxObject *workbook = workbooks->querySubObject("Open(const QString&, QVariant)", sourceExcelData->getSourcePath(), 0);
    QAxObject *worksheets = workbook->querySubObject("WorkSheets"); // 获取打开的excel文件中所有的工作sheet
    sourceExcelData->setSheetCnt(worksheets->property("Count").toInt());

    qDebug() << QString("Excel文件中表的个数: %1").arg(QString::number(sourceExcelData->getSheetCnt()));
    QAxObject *worksheet = nullptr;
    int selectedSheetIndex = 0;
    for (int i = 1; i <= sourceExcelData->getSheetCnt(); i++) {
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
    sourceExcelData->setSourceRowStart(usedRange->property("Row").toInt());    // 起始行数
    sourceExcelData->setSourceColStart(usedRange->property("Column").toInt()); // 起始列数
    QAxObject *rows, *columns;
    rows = usedRange->querySubObject("Rows");       // 行
    columns = usedRange->querySubObject("Columns"); // 列

    sourceExcelData->setSourceRowCnt(rows->property("Count").toInt());    // 行数
    sourceExcelData->setSourceColCnt(columns->property("Count").toInt()); // 列数
    QString sourceMinAlphabetCol;
    QString sourceMaxAlphabetCol;
    ExcelBase::convertToColName(sourceExcelData->getSourceRowStart(), sourceMinAlphabetCol);
    ExcelBase::convertToColName(sourceExcelData->getSourceColCnt(), sourceMaxAlphabetCol);
    sourceExcelData->setSourceMinAlphabetCol(sourceMinAlphabetCol);
    sourceExcelData->setSourceMaxAlphabetCol(sourceMaxAlphabetCol);

    qDebug() << " sourceRowStart:" << sourceExcelData->getSourceRowStart() << " sourceColStart:" << sourceExcelData->getSourceColCnt()
             << " sourceRowCnt:" << sourceExcelData->getSourceRowCnt() << " sourceColCnt:" << sourceExcelData->getSourceColCnt()
             << " sourceMinAlphabetCol:" << sourceMinAlphabetCol << " sourceMaxAlphabetCol:" << sourceMaxAlphabetCol;
    // CoUninitialize();
    ExcelParserByOfficeRunnable::freeExcel(excel);
    ExcelParserByOfficeRunnable::generateTplXls(sourceExcelData, selectedSheetIndex); //生成模板文件
}
