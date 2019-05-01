#include "testtimeer.h"
#include "ui_testtimeer.h"

TestTimeer::TestTimeer(QWidget *parent,QMainWindow *mainWindow) :
    QMainWindow(parent),
    ui(new Ui::TestTimeer)
{
    this->mainWindow = mainWindow;
    ui->setupUi(this);
}


//接受子线程的消息
void TestTimeer::receiveMessage(const int msgType, const QString &msg)
{
    //qDebug() << "SplitOnlyWindow::receiveMessage msgType:" << QString::number(msgType).toUtf8() <<" msg:"<<msg;
    switch (msgType)
    {
    case Common::MsgTypeError:

        processWindow->setProcessText(msg);

        break;
    case Common::MsgTypeWriteXlsxFinish:
        processWindow->setProcessText(msg);
        break;
    case Common::MsgTypeEmailSendFinish:
        processWindow->setProcessText(msg);

        break;
    case Common::MsgTypeSucc:
    case Common::MsgTypeInfo:
    case Common::MsgTypeWarn:
    default:
        processWindow->setProcessText(msg);
        break;
    }
}

TestTimeer::~TestTimeer()
{
    delete ui;
}

void TestTimeer::showIdleMsg()
{
    QString idleMsg("达到当前处理上线，暂时进入休息状态. 剩余休息时间 %1 %2s");

    int ela = this->timer.elapsed();
    int currentLeftTime = this->leftTime - ela;

    this->receiveMessage(Common::MsgTypeInfo, idleMsg.arg(ela).arg(currentLeftTime));
    if (currentLeftTime > 0)
    {
        if (ela % this->idleMsgShowPeriod == 0)
        {
            this->receiveMessage(Common::MsgTypeInfo, idleMsg.arg(ela).arg(currentLeftTime));
        }
        QTimer::singleShot(100, this, SLOT(showIdleMsg()));
    }
    else
    {
        this->receiveMessage(Common::MsgTypeInfo, "休息结束.");
    }
}

void TestTimeer::on_pushButton_clicked()
{
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

    int timeUnit = 1000; //时间单位
    int periodTime = timeUnit * 60; //1min
    int idleMsgShowPeriod = 2 * timeUnit;//2s一个周期

    this->leftTime = periodTime;
    QElapsedTimer timer;
    timer.start();
    this->idleMsgShowPeriod = idleMsgShowPeriod;
    this->timeUnit = timeUnit;
    this->timer = timer;
    this->showIdleMsg();
}
