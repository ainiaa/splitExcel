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
        QStringList passwordDatas;
        passwordDatas.insert(0,"--请选择--");
        passwordDatas.append(sheetNames);
        ui->passwordComboBox->addItems(passwordDatas);
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
    int passwordSheetIndex = ui->passwordComboBox->currentIndex(); //password所在 sheet 索引 （从 0 开始)
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
    sourceExcelData->setNeedPassword(passwordSheetIndex > 0);
    sourceExcelData->setPasswordDataSheetName(ui->passwordComboBox->currentText());

    //设置excel 解析相关事件
    excelParser->setSplitData(cfg, sourceExcelData);
	
	excelParserThread = new QThread();
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
