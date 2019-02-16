#include "splitonlywindow.h"
#include "ui_splitonlywindow.h"

SplitOnlyWindow::SplitOnlyWindow(QWidget *parent,QMainWindow *mainWindow) :
    QMainWindow(parent),
    ui(new Ui::SplitOnlyWindow)
{
    this->mainWindow = mainWindow;
    ui->setupUi(this);

    setFixedSize(this->width(), this->height());

    xlsxParser = new XlsxParser();
}

SplitOnlyWindow::~SplitOnlyWindow()
{
    delete ui;
}

void SplitOnlyWindow::on_selectFilePushButton_clicked()
{
    QString path = xlsxParser->openFile(this);
    if (path.length() > 0)
    {//选择了excel文件
        ui->xlsObjLineEdit->setText(path);

        //获得所有的sheets
        QStringList sheetNames = xlsxParser->getSheetNames();
        ui->dataComboBox->addItems(sheetNames);
        ui->dataComboBox->setCurrentIndex(0);

        //获得表头数据
        changeGroupby(sheetNames.at(0));
    }
}

//修改分组下拉选择框的选项
void SplitOnlyWindow::changeGroupby(QString selectedSheetName)
{
    header->clear();
    ui->groupByComboBox->clear();
    header = xlsxParser->getSheetHeader(selectedSheetName);
    ui->groupByComboBox->addItems(*header);
}

void SplitOnlyWindow::on_savePathPushButton_clicked()
{
    QString path = QFileDialog::getExistingDirectory(this, tr("Save New Excels"), ".");
    if(path.length() > 0)
    {
        ui->savePathLineEdit->setText(path);
        savePath = QDir::toNativeSeparators(path);
    }
}

void SplitOnlyWindow::on_submitPushButton_clicked()
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

    //拆分
    doSplitXls(ui->dataComboBox->currentText(), savePath);
}

//拆分
void SplitOnlyWindow::doSplitXls(QString dataSheetName, QString savePath)
{
    if(xlsxParserThread != nullptr)
    {
        return;
    }
    QString groupByText = ui->groupByComboBox->currentText();
    xlsxParserThread= new QThread();
    xlsxParser->setSplitData(cfg, groupByText, dataSheetName, nullptr, savePath);
    xlsxParser->moveToThread(xlsxParserThread);
    connect(xlsxParserThread, &QThread::finished, xlsxParserThread, &QObject::deleteLater);
    connect(xlsxParserThread, &QThread::finished, xlsxParser, &QObject::deleteLater);
    connect(this, &SplitOnlyWindow::doSplit, xlsxParser, &XlsxParser::doSplit);
    connect(xlsxParser, &XlsxParser::requestMsg, this, &SplitOnlyWindow::receiveMessage);
    xlsxParserThread->start();

    emit doSplit(); //主线程通过信号换起子线程的槽函数
}

//接受子线程的消息
void SplitOnlyWindow::receiveMessage(const int msgType, const QString &msg)
{
    qDebug() << "SplitOnlyWindow::receiveMessage msgType:" << QString::number(msgType).toUtf8() <<" msg:"<<msg;
    switch (msgType)
    {
    case Common::MsgTypeError:
        ui->submitPushButton->setDisabled(false);

        processWindow->setProcessText(msg);

        QMessageBox::critical(this, "Error", msg);
        //errorMessage(msg);
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

void SplitOnlyWindow::on_gobackPushButton_clicked()
{
    this->hide();
    mainWindow->show();
}
