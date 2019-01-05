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

    connect(ui->actionConfig_Setting, SIGNAL(triggered()), this, SLOT(showConfigSetting()));
}

MainWindow::~MainWindow()
{
    delete ui;
    if (mailsenderThread)
    {
        mailsenderThread->quit();
    }
    mailsenderThread->wait();
    delete processWindow;
    delete mailsender;
    delete xlsx;
    delete header;
    delete configSetting;
    delete cfg;
    delete mailsenderThread;

    qDebug() << "end destroy widget";
}

//选择文件
void MainWindow::on_selectFilePushButton_clicked()
{
    QString path = MainWindow::openFile();
    if (path.length() > 0)
    {//选择了excel文件
        ui->xlsObjLineEdit->setText(path);
        xlsx = new QXlsx::Document (path);

        //获得所有的sheets
        QStringList sheetNames = xlsx->sheetNames();
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

    QXlsx::CellRange range;
    xlsx->selectSheet(selectedSheetName);
    range = xlsx->dimension();
    int colCount = range.columnCount();

    for (int colum=1; colum<=colCount; ++colum)
    {
        QXlsx::Cell *cell =xlsx->cellAt(1, colum);
        if (cell)
        {
            header->append(cell->value().toString());
        }
    }
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

        errorMessage(msg);
        break;
    case Common::MsgTypeEmailSendFinish:
        ui->submitPushButton->setDisabled(false);
        ui->statusBar->showMessage(msg);
        processWindow->setProcessText(msg);
        mailsenderThread = nullptr;
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

//打开文件对话框
QString MainWindow::openFile()
{
    QString path = QFileDialog::getOpenFileName(this, tr("Open Excel"), ".", tr("excel(*.xlsx)"));
    if(path.length() == 0)
    {
        return "";
    }
    else
    {
        return path;
    }
}

//选择保存目录
void MainWindow::on_savePathPushButton_clicked()
{
    QString path = QFileDialog::getExistingDirectory(this, tr("Save New Excels"), ".");
    if(path.length() > 0)
    {
        ui->savePathLineEdit->setText(path);
    }
}

//确定拆分
void MainWindow::on_submitPushButton_clicked()
{
    ui->submitPushButton->setDisabled(true);

    QString server = cfg->Get("email","server").toString();
    if (server.isEmpty())
    {
        QMessageBox::information(this, "Setting Error", "请先配置邮件相关配置");
    }
    QString savePath = ui->savePathLineEdit->text();
    if (xlsx == nullptr)
    {
        QMessageBox::information(this, "Save Path Error", "请选择待拆分的excel");
    }
    if (savePath.length() == 0)
    {
        QMessageBox::information(this, "Save Path Error", "请选择保存路径");
    }

    int groupbyindex = ui->groupByComboBox->currentIndex() + 1;

    savePath = savePath.replace("/","\\");

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

    //拆分 && 后续操作(发送email)
    doSplitXls(groupbyindex,ui->dataComboBox->currentText(),ui->emailComboBox->currentText(),savePath);
}

//拆分 && 后续操作(发送email)
void MainWindow::doSplitXls(int groupby, QString dataSheetName, QString emailSheetName, QString savePath)
{
    //读取excel数据
    ui->statusBar->showMessage("开始读取excel文件信息");
    receiveMessage(Common::MsgTypeInfo,"开始读取excel文件信息");
    QHash<QString, QList<QStringList>> dataQhash = readXls(groupby,dataSheetName);
    QHash<QString, QList<QStringList>> emailQhash = readXls(groupby,emailSheetName);

    //写excel
    receiveMessage(Common::MsgTypeInfo,"开始拆分excel并生成新的excel文件");
    ui->statusBar->showMessage("开始拆分excel并生成新的excel文件");
    writeXls(dataQhash,savePath);

    //发送email
    sendemail(emailQhash,savePath);
}

//发送邮件
//@see https://blog.csdn.net/czyt1988/article/details/71194457
void MainWindow::sendemail(QHash<QString, QList<QStringList>> emailQhash, QString savePath)
{
    int total = emailQhash.size();
    if(mailsenderThread != nullptr)
    {
        return;
    }
    mailsenderThread= new QThread();
    mailsender = new EmailSender();
    mailsender->setSendData(cfg,emailQhash,savePath,total);
    mailsender->moveToThread(mailsenderThread);
    connect(mailsenderThread,&QThread::finished,mailsenderThread,&QObject::deleteLater);
    connect(mailsenderThread,&QThread::finished,mailsender,&QObject::deleteLater);
    connect(this,&MainWindow::doSend,mailsender,&EmailSender::doSend);
    connect(mailsender,&EmailSender::requestMsg,this,&MainWindow::receiveMessage);
    mailsenderThread->start();

    emit doSend(); //主线程通过信号换起子线程的槽函数
}

//读取xls
QHash<QString, QList<QStringList>> MainWindow::readXls(int groupby, QString selectedSheetName)
{
    QXlsx::CellRange range;
    xlsx->selectSheet(selectedSheetName);
    range = xlsx->dimension();
    int rowCount = range.rowCount();
    int colCount = range.columnCount();

    QHash<QString, QList<QStringList>> qhash;
    for (int row = 2;row <= rowCount;++row)
    {
        QString groupbyvalue;
        QStringList items;
        QXlsx::Cell *cell =xlsx->cellAt(row, groupby);
        if (cell)
        {
            groupbyvalue = cell->value().toString();
        }
        for (int colum=1; colum<=colCount; ++colum)
        {
            QXlsx::Cell *cell =xlsx->cellAt(row, colum);
            if (cell)
            {
                if (cell->isDateTime())
                {
                    items.append(cell->dateTime().toString("yyyy/MM/dd"));
                }
                else
                {
                    items.append(cell->value().toString());
                }
            }
        }
        QList<QStringList> qlist = qhash.take(groupbyvalue);
        qlist.append(items);
        qhash.insert(groupbyvalue,qlist);
    }
    return qhash;
}

//写xls
void MainWindow::writeXls(QHash<QString, QList<QStringList>> qhash, QString savePath)
{
    qDebug("writeXls");
    QHashIterator<QString,QList<QStringList>> it(qhash);
    QString startMsg("开始拆分excel并生成新的excel文件: %1/");
    QString endMsg("完成拆分excel并生成新的excel文件: %1/");
    startMsg.append(QString::number(qhash.size()));
    endMsg.append(QString::number(qhash.size()));

    int count = 1;
    while (it.hasNext()) {
        it.next();

        receiveMessage(Common::MsgTypeInfo,startMsg.arg(QString::number(count)));
        QString key = it.key();
        QList<QStringList> content = it.value();
        QXlsx::Document currXls;

        //写表头
        writeXlsHeader(&currXls);

        QString xlsName;
        int rows =  content.size();
        xlsName.append(savePath).append("\\").append(key).append(".xlsx");

        for(int row = 0; row < rows;row++)
        {
            QStringList qsl = content.at(row);
            int columns = qsl.size();
            for (int column =0;column < columns;column++)
            {
                QString columnName;
                convertToColName(column+1,columnName) ;
                QString cell;
                cell.append(columnName).append(QString::number(row+2));
                currXls.write(cell,qsl.at(column));
            }
        }
        currXls.saveAs(xlsName);
        receiveMessage(Common::MsgTypeInfo,endMsg.arg(QString::number(count)));
        count++;
    }
    ui->statusBar->showMessage("拆分excel结束");
    receiveMessage(Common::MsgTypeInfo,"拆分excel结束");
}

void  MainWindow::writeXlsHeader(QXlsx::Document *xls)
{
    int columns =header->size();
    for (int column =0;column < columns;column++)
    {
        QString columnName;
        convertToColName(column+1,columnName) ;
        QString cell;
        cell.append(columnName).append(QString::number(1));
        xls->write(cell,header->at(column));
    }
}

///
/// \brief 把列数转换为excel的字母列号
/// \param data 大于0的数
/// \return 字母列号，如1->A 26->Z 27 AA
///
void MainWindow::convertToColName(int data, QString &res)
{
    int tempData = data / 26;
    if(tempData > 0)
    {
        int mode = data % 26;
        convertToColName(mode,res);
        convertToColName(tempData,res);
    }
    else
    {
        res=(to26AlphabetString(data)+res);
    }
}
///
/// \brief 数字转换为26字母
///
/// 1->A 26->Z
/// \param data
/// \return
///
QString MainWindow::to26AlphabetString(int data)
{
    QChar ch = data + 0x40;//A对应0x41
    return QString(ch);
}

void MainWindow::errorMessage(const QString &message)
{
    QErrorMessage err (this);

    err.showMessage(message);

    err.exec();
}

void MainWindow::on_dataComboBox_currentTextChanged(const QString &arg1)
{
    changeGroupby(arg1);
}

void MainWindow::showConfigSetting()
{
    this->hide();
    configSetting->show();
}
