#include "splitandemailwindow.h"
#include "ui_splitandemailwindow.h"

SplitAndEmailWindow::SplitAndEmailWindow(QMainWindow *parent)
: SplitSubWindow(parent), ui(new Ui::SplitAndEmailWindow), xlsx(nullptr) {
    ui->setupUi(this);
    setWindowFlags(windowFlags() & ~Qt::WindowMaximizeButtonHint);
    setFixedSize(this->width(), this->height());

    excelParser = new ExcelParser();
    connect(ui->actionConfig_Setting, SIGNAL(triggered()), this, SLOT(showConfigSetting()));
}

SplitAndEmailWindow::~SplitAndEmailWindow() {
    qDebug() << "SplitAndEmailWindow::~SplitAndEmailWindow start";
    delete ui;

    if (mailSender) {
        delete mailSender;
    }
    if (mailSenderThread) {
        mailSenderThread->quit();
        mailSenderThread->wait();
        delete mailSenderThread;
    }

    if (excelParser) {
        delete excelParser;
    }
    if (excelParserThread) {
        excelParserThread->quit();
        excelParserThread->wait();
        delete excelParserThread;
    }

    if (processWindow) {
        delete processWindow;
    }

    if (xlsx) {
        delete xlsx;
    }
    if (header) {
        delete header;
    }

    if (configSetting) {
        delete configSetting;
    }

    if (cfg) {
        delete cfg;
    }
    qDebug() << "SplitAndEmailWindow::~SplitAndEmailWindow end";
}

//选择xlsx文件
void SplitAndEmailWindow::on_selectFilePushButton_clicked() {
    QString path = excelParser->openFile(this);
    if (path.length() > 0) { //选择了excel文件
        ui->xlsObjLineEdit->setText(path);
        this->sourcePath = QDir::toNativeSeparators(path);
        //获得所有的sheets
        QStringList sheetNames = excelParser->getSheetNames();
        ui->dataComboBox->addItems(sheetNames);
        ui->dataComboBox->setCurrentIndex(0);
        ui->emailComboBox->addItems(sheetNames);
        ui->emailComboBox->setCurrentIndex(1);

        QStringList passwordDatas;
        passwordDatas.insert(0,"--请选择--");
        passwordDatas.append(sheetNames);
        ui->passwordComboBox->addItems(passwordDatas);

        //获得表头数据
        changeGroupby(sheetNames.at(0));
    }
}

//修改分组下拉选择框的选项
void SplitAndEmailWindow::changeGroupby(QString selectedSheetName) {
    header->clear();
    ui->groupByComboBox->clear();
    header = excelParser->getSheetHeader(selectedSheetName);
    ui->groupByComboBox->addItems(*header);
}

//接受子线程的消息
void SplitAndEmailWindow::receiveMessage(const int msgType, const QString &msg) {
    qDebug() << "SplitAndEmailWindow::receiveMessage msgType:" << QString::number(msgType).toUtf8() << " msg:" << msg;
    switch (msgType) {
        case Common::MsgTypeStart:
            qDebug() << "Common::MsgTypeStart total_cnt: " << msg;
            this->setTotalCnt(msg.toInt());
            processWindow->startTimer();
            break;
        case Common::MsgTypeError:
            ui->submit2PushButton->setDisabled(false);

            processWindow->setProcessText(msg);
            processWindow->setProcessPercent(this->incrAndCalcPercent(1));
            ui->statusBar->showMessage(msg);

            QMessageBox::critical(this, "Error", msg);
            // errorMessage(msg);
            break;
        case Common::MsgTypeWriteXlsxFinish:
            processWindow->setProcessText(msg);
            processWindow->setProcessPercent(this->incrAndCalcPercent(1));
            ui->statusBar->showMessage(msg);
            break;
        case Common::MsgTypeStartSendEmail:
            //发送email
            sendemail();
            break;
        case Common::MsgTypeEmailSendFinish:
            ui->submit2PushButton->setDisabled(false);
            ui->statusBar->showMessage(msg);
            processWindow->setProcessText(msg);
            processWindow->setProcessPercent(100);
            processWindow->stopTimer();
            mailSenderThread = nullptr;
            break;
        case Common::MsgTypeSucc:
            ui->statusBar->showMessage(msg);
            processWindow->setProcessText(msg);
            processWindow->setProcessPercent(this->incrAndCalcPercent(1));
            break;
        case Common::MsgTypeInfo:
        case Common::MsgTypeWarn:
        default:
            ui->statusBar->showMessage(msg);
            processWindow->setProcessText(msg);
            break;
    }
}

//选择新Xlsx文件保存目录
void SplitAndEmailWindow::on_savePathPushButton_clicked() {
    QString path = QFileDialog::getExistingDirectory(this, tr("Save New Excels"), ".");
    if (path.length() > 0) {
        ui->savePathLineEdit->setText(path);
        savePath = QDir::toNativeSeparators(path);
    }
}

//拆分 && 后续操作(发送email)
void SplitAndEmailWindow::doSplitXls(QString dataSheetName, QString emailSheetName, QString savePath) {
    qDebug() << "SplitAndEmailWindow::doSplitXls";
    if (excelParserThread != nullptr) {
        qDebug() << "SplitAndEmailWindow::doSplitXls doing";
        return;
    }
    QString groupByText = ui->groupByComboBox->currentText();
    int groupByIndex = ui->groupByComboBox->currentIndex(); //分组列 索引（从 0 开始)

    int dataSheetIndex = ui->dataComboBox->currentIndex();
    int emailSheetIndex = ui->emailComboBox->currentIndex();
    int passwordSheetIndex = ui->passwordComboBox->currentIndex();

    int sheetCnt = ui->dataComboBox->count();
    //设置excel相关数据
    SourceExcelData *sourceExcelData = new SourceExcelData();
    sourceExcelData->setSavePath(savePath);
    sourceExcelData->setDataSheetName(dataSheetName);
    sourceExcelData->setDataSheetIndex(dataSheetIndex);
    sourceExcelData->setSheetCnt(sheetCnt);
    sourceExcelData->setEmailSheetName(emailSheetName);
    sourceExcelData->setEmailSheetIndex(emailSheetIndex);
    sourceExcelData->setOpType(SourceExcelData::OperateType::SplitAndEmailType);
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
    connect(this, &SplitAndEmailWindow::doProcess, excelParser, &ExcelParser::doParse);
    connect(excelParser, &ExcelParser::requestMsg, this, &SplitAndEmailWindow::receiveMessage);
    excelParserThread->start();

    emit doProcess(); //主线程通过信号换起子线程的槽函数
}

//发送邮件
//@see https://blog.csdn.net/czyt1988/article/details/71194457
void SplitAndEmailWindow::sendemail() {
    QHash<QString, QList<QStringList>> emailQhash = excelParser->getEmailData();
    int total = emailQhash.size();
    if (mailSenderThread != nullptr) {
        return;
    }
    mailSenderThread = new QThread();
    mailSender = new EmailSender();
    mailSender->setSendData(cfg, emailQhash, savePath, total);
    mailSender->moveToThread(mailSenderThread);
    connect(mailSenderThread, &QThread::finished, mailSenderThread, &QObject::deleteLater);
    connect(mailSenderThread, &QThread::finished, mailSender, &QObject::deleteLater);
    connect(this, &SplitAndEmailWindow::doSend, mailSender, &EmailSender::doSend);
    connect(mailSender, &EmailSender::requestMsg, this, &SplitAndEmailWindow::receiveMessage);
    mailSenderThread->start();

    emit doSend(); //主线程通过信号换起子线程的槽函数
}

void SplitAndEmailWindow::errorMessage(const QString &message) {
    QErrorMessage err(this);

    err.showMessage(message);

    err.exec();
}

//修改分组
void SplitAndEmailWindow::on_dataComboBox_currentTextChanged(const QString &arg1) {
    changeGroupby(arg1);
}

//显示配置UI
void SplitAndEmailWindow::showConfigSetting() {
    this->hide();
    configSetting->show();
}

void SplitAndEmailWindow::on_gobackPushButton_clicked() {
    this->getMainWindow()->show();
    this->hide();
}

void SplitAndEmailWindow::on_submit2PushButton_clicked() {
    ui->submit2PushButton->setDisabled(true);

    QString server = cfg->get("email", "server").toString();
    if (server.isEmpty()) {
        QMessageBox::information(this, "Setting Error", "请先配置邮件相关配置");
        ui->submit2PushButton->setDisabled(false);
        return;
    }

    if (ui->xlsObjLineEdit->text().isEmpty()) {
        QMessageBox::information(this, "Save Path Error", "请选择待拆分的excel");
        ui->submit2PushButton->setDisabled(false);
        return;
    }
    if (savePath.length() == 0) {
        QMessageBox::information(this, "Save Path Error", "请选择保存路径");
        ui->submit2PushButton->setDisabled(false);
        return;
    }

    if (processWindow == nullptr) {
        processWindow = new ProcessWindow(this);
    } else {
        processWindow->clearProcessText();
    }
    processWindow->setWindowModality(Qt::WindowModality::ApplicationModal);
    processWindow->show();

    qDebug() << "拆分 && 后续操作(发送email)";
    //拆分 && 后续操作(发送email)
    doSplitXls(ui->dataComboBox->currentText(), ui->emailComboBox->currentText(), savePath);
}
