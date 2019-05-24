#include "emailonlywindow.h"
#include "ui_emailonlywindow.h"

EmailOnlyWindow::EmailOnlyWindow(QMainWindow *parent) : SplitSubWindow(parent), ui(new Ui::EmailOnlyWindow) {
    ui->setupUi(this);

    setFixedSize(this->width(), this->height());

    excelParser = new ExcelParser();
    connect(ui->actionConfig_Setting, SIGNAL(triggered()), this, SLOT(showConfigSetting()));
}

EmailOnlyWindow::~EmailOnlyWindow() {
    delete ui;
}

void EmailOnlyWindow::on_selectFilePushButton_clicked() {
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
void EmailOnlyWindow::changeGroupby(QString selectedSheetName) {
    header->clear();
    ui->groupByComboBox->clear();
    header = excelParser->getSheetHeader(selectedSheetName);
    ui->groupByComboBox->addItems(*header);
}

void EmailOnlyWindow::on_savePathPushButton_clicked() {
    QString path = QFileDialog::getExistingDirectory(this, tr("Save New Excels"), ".");
    if (path.length() > 0) {
        ui->savePathLineEdit->setText(path);
        savePath = QDir::toNativeSeparators(path);
    }
}

void EmailOnlyWindow::on_submitPushButton_clicked() {
    ui->submitPushButton->setDisabled(true);

    if (ui->xlsObjLineEdit->text().isEmpty()) {
        QMessageBox::information(this, "Save Path Error", "请选择待拆分的excel");
        ui->submitPushButton->setDisabled(false);
        return;
    }
    if (savePath.length() == 0) {
        QMessageBox::information(this, "Save Path Error", "请选择附件路径");
        ui->submitPushButton->setDisabled(false);
        return;
    }

    if (processWindow == nullptr) {
        processWindow = new ProcessWindow();
    } else {
        processWindow->clearProcessText();
    }
    processWindow->setWindowModality(Qt::WindowModality::ApplicationModal);
    processWindow->show();

    //发送email
    doSendEmail(ui->dataComboBox->currentText(), savePath);
}

//发送email
void EmailOnlyWindow::doSendEmail(QString emailSheetName, QString savePath) {
    if (excelParserThread != nullptr) {
        return;
    }
    QString groupByText = ui->groupByComboBox->currentText();
    int dataSheetIndex = ui->dataComboBox->currentIndex();
    int sheetCnt = ui->dataComboBox->count();
    SourceExcelData *sourceExcelData = new SourceExcelData();
    sourceExcelData->setSavePath(savePath);
    sourceExcelData->setEmailSheetName(emailSheetName);
    sourceExcelData->setSheetCnt(sheetCnt);
    sourceExcelData->setEmailSheetIndex(dataSheetIndex);
    sourceExcelData->setOpType(SourceExcelData::OperateType::EmailOnlyType);
    sourceExcelData->setGroupByText(groupByText);
    sourceExcelData->setSourcePath(this->sourcePath);
    excelParserThread = new QThread();
    excelParser->setSplitData(cfg, sourceExcelData);
    excelParser->moveToThread(excelParserThread);
    connect(excelParserThread, &QThread::finished, excelParserThread, &QObject::deleteLater);
    connect(excelParserThread, &QThread::finished, excelParser, &QObject::deleteLater);
    connect(this, &EmailOnlyWindow::doProcess, excelParser, &ExcelParser::doParse);
    connect(excelParser, &ExcelParser::requestMsg, this, &EmailOnlyWindow::receiveMessage);
    excelParserThread->start();

    emit doProcess(); //主线程通过信号换起子线程的槽函数
}

//接受子线程的消息
void EmailOnlyWindow::receiveMessage(const int msgType, const QString &msg) {
    qDebug() << "EmailOnlyWindow::receiveMessage msgType:" << QString::number(msgType).toUtf8() << " msg:" << msg;
    switch (msgType) {
        case Common::MsgTypeError:
            ui->submitPushButton->setDisabled(false);
            processWindow->setProcessText(msg);
            QMessageBox::critical(this, "Error", msg);
            break;
        case Common::MsgTypeWriteXlsxFinish:
            processWindow->setProcessText(msg);
            ui->submitPushButton->setDisabled(false);
            break;
        case Common::MsgTypeEmailSendFinish:
            ui->submitPushButton->setDisabled(false);
            processWindow->setProcessText(msg);
            mailSenderThread = nullptr;
            break;
        case Common::MsgTypeSucc:
        case Common::MsgTypeInfo:
        case Common::MsgTypeWarn:
        default:
            processWindow->setProcessText(msg);
            break;
    }
}

void EmailOnlyWindow::on_gobackPushButton_clicked() {
    this->hide();
    this->getMainWindow()->show();
}

//显示配置UI
void EmailOnlyWindow::showConfigSetting() {
    this->hide();
    configSetting->show();
}
