#include "mainwindow.h"
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow),
    xlsx(nullptr)
{
    ui->setupUi(this);
    setWindowFlags(windowFlags()& ~Qt::WindowMaximizeButtonHint);
    setFixedSize(this->width(), this->height());

    xlsxParser = new XlsxParser();
    connect(ui->actionConfig_Setting, SIGNAL(triggered()), this, SLOT(showConfigSetting()));
    connect(ui->actionSplit_Only, SIGNAL(triggered()), this, SLOT(showSplitOnly()));
}

MainWindow::~MainWindow()
{
    qDebug() << "MainWindow::~MainWindow start";
    delete ui;

    if (mailSender)
    {
        delete mailSender;
    }
    if (mailSenderThread)
    {
        mailSenderThread->quit();
        mailSenderThread->wait();
        delete mailSenderThread;
    }

    if (xlsxParser)
    {
        delete xlsxParser;
    }
    if (xlsxParserThread)
    {
        xlsxParserThread->quit();
        xlsxParserThread->wait();
        delete xlsxParserThread;
    }



    if (processWindow)
    {
        delete processWindow;
    }

    if (xlsx)
    {
        delete xlsx;
    }
    if (header)
    {
        delete header;
    }

    if (configSetting)
    {
        delete configSetting;
    }

    if (cfg)
    {
        delete cfg;
    }
    qDebug() << "MainWindow::~MainWindow end";
}

//选择xlsx文件
void MainWindow::on_selectFilePushButton_clicked()
{
    QString path = xlsxParser->openFile(this);
    if (path.length() > 0)
    {//选择了excel文件
        ui->xlsObjLineEdit->setText(path);

        //获得所有的sheets
        QStringList sheetNames = xlsxParser->getSheetNames();
        ui->dataComboBox->addItems(sheetNames);
        ui->dataComboBox->setCurrentIndex(0);
        ui->emailComboBox->addItems(sheetNames);
        ui->emailComboBox->setCurrentIndex(1);

        //获得表头数据
        changeGroupby(sheetNames.at(0));
    }
}

//修改分组下拉选择框的选项
void MainWindow::changeGroupby(QString selectedSheetName)
{
    header->clear();
    ui->groupByComboBox->clear();
    header = xlsxParser->getSheetHeader(selectedSheetName);
    ui->groupByComboBox->addItems(*header);
}

//接受子线程的消息
void MainWindow::receiveMessage(const int msgType, const QString &msg)
{
    qDebug() << "MainWindow::receiveMessage msgType:" << QString::number(msgType).toUtf8() <<" msg:"<<msg;
    switch (msgType)
    {
    case Common::MsgTypeError:
        ui->submitPushButton->setDisabled(false);

        processWindow->setProcessText(msg);
        ui->statusBar->showMessage(msg);

        QMessageBox::critical(this, "Error", msg);
        //errorMessage(msg);
        break;
    case Common::MsgTypeWriteXlsxFinish:
        processWindow->setProcessText(msg);
        ui->statusBar->showMessage(msg);
        //发送email
        sendemail();
        break;
    case Common::MsgTypeEmailSendFinish:
        ui->submitPushButton->setDisabled(false);
        ui->statusBar->showMessage(msg);
        processWindow->setProcessText(msg);
        mailSenderThread = nullptr;
        break;
    case Common::MsgTypeSucc:
    case Common::MsgTypeInfo:
    case Common::MsgTypeWarn:
    default:
        ui->statusBar->showMessage(msg);
        processWindow->setProcessText(msg);
        break;
    }
}

//选择新Xlsx文件保存目录
void MainWindow::on_savePathPushButton_clicked()
{
    QString path = QFileDialog::getExistingDirectory(this, tr("Save New Excels"), ".");
    if(path.length() > 0)
    {
        ui->savePathLineEdit->setText(path);
        savePath = QDir::toNativeSeparators(path);
    }
}

//确定拆分
void MainWindow::on_submitPushButton_clicked()
{
    ui->submitPushButton->setDisabled(true);

    QString server = cfg->get("email","server").toString();
    if (server.isEmpty())
    {
        QMessageBox::information(this, "Setting Error", "请先配置邮件相关配置");
         ui->submitPushButton->setDisabled(false);
        return;
    }

    if (ui->xlsObjLineEdit->text().isEmpty())
    {
        QMessageBox::information(this, "Save Path Error", "请选择待拆分的excel");
        ui->submitPushButton->setDisabled(false);
        return;
    }
    if (savePath.length() == 0)
    {
        QMessageBox::information(this, "Save Path Error", "请选择保存路径");
        ui->submitPushButton->setDisabled(false);
        return;
    }

    if (processWindow == nullptr)
    {
        processWindow = new ProcessWindow();
    }
    else
    {
        processWindow->clearProcessText();
    }
    processWindow->setWindowModality(Qt::WindowModality::ApplicationModal);
    processWindow->show();

    qDebug() << "拆分 && 后续操作(发送email)";
    //拆分 && 后续操作(发送email)
    doSplitXls(ui->dataComboBox->currentText(), ui->emailComboBox->currentText(), savePath);
}

//拆分 && 后续操作(发送email)
void MainWindow::doSplitXls(QString dataSheetName, QString emailSheetName, QString savePath)
{
    if(xlsxParserThread != nullptr)
    {
        return;
    }
    QString groupByText = ui->groupByComboBox->currentText();
    xlsxParserThread= new QThread();
    xlsxParser->setSplitData(cfg, groupByText, dataSheetName, emailSheetName, savePath);
    xlsxParser->moveToThread(xlsxParserThread);
    connect(xlsxParserThread, &QThread::finished, xlsxParserThread, &QObject::deleteLater);
    connect(xlsxParserThread, &QThread::finished, xlsxParser, &QObject::deleteLater);
    connect(this, &MainWindow::doSplit, xlsxParser, &XlsxParser::doSplit);
    connect(xlsxParser, &XlsxParser::requestMsg, this, &MainWindow::receiveMessage);
    xlsxParserThread->start();

    emit doSplit(); //主线程通过信号换起子线程的槽函数
}

//发送邮件
//@see https://blog.csdn.net/czyt1988/article/details/71194457
void MainWindow::sendemail()
{
    QHash<QString, QList<QStringList>> emailQhash = xlsxParser->getEmailData();
    int total = emailQhash.size();
    if(mailSenderThread != nullptr)
    {
        return;
    }
    mailSenderThread= new QThread();
    mailSender = new EmailSender();
    mailSender->setSendData(cfg,emailQhash,savePath,total);
    mailSender->moveToThread(mailSenderThread);
    connect(mailSenderThread, &QThread::finished, mailSenderThread, &QObject::deleteLater);
    connect(mailSenderThread, &QThread::finished, mailSender, &QObject::deleteLater);
    connect(this, &MainWindow::doSend, mailSender, &EmailSender::doSend);
    connect(mailSender, &EmailSender::requestMsg, this, &MainWindow::receiveMessage);
    mailSenderThread->start();

    emit doSend(); //主线程通过信号换起子线程的槽函数
}

void MainWindow::errorMessage(const QString &message)
{
    QErrorMessage err (this);

    err.showMessage(message);

    err.exec();
}

//修改分组
void MainWindow::on_dataComboBox_currentTextChanged(const QString &arg1)
{
    changeGroupby(arg1);
}

//显示配置UI
void MainWindow::showConfigSetting()
{
    this->hide();
    configSetting->show();
}

//显示配置UI
void MainWindow::showSplitOnly()
{
    this->hide();
    splitOnlyWindow->show();
}
