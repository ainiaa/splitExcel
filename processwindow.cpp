#include "processwindow.h"
#include "qdebug.h"
#include "ui_processwindow.h"

ProcessWindow::ProcessWindow(QWidget *parent) : QMainWindow(parent), ui(new Ui::ProcessWindow) {
    ui->setupUi(this);
    setWindowFlags(Qt::Dialog);
    setFixedSize(this->width(), this->height());

    this->timer = new QTimer();
    this->timeRecord = new QTime(0, 0, 0); //初始化时间
    connect(timer, SIGNAL(timeout()), this, SLOT(updateTime()));
}

ProcessWindow::~ProcessWindow() {
    delete ui;
}

void ProcessWindow::setProcessText(QString text) {
    QDateTime currentTime = QDateTime::currentDateTime();
    QString strCurrentTime = currentTime.toString("hh:mm:ss");
    ui->textEdit->append(strCurrentTime + "  " + text);
}

void ProcessWindow::setProcessPercent(int percent) {
    ui->progressBar->setValue(percent);
}

void ProcessWindow::updateTime() {
    *timeRecord = this->timeRecord->addSecs(1);
    QString txt = timeRecord->toString("hh:mm:ss");
    ui->timerLabel->setText(txt);
    qDebug() << "ProcessWindow::updateTime txt:" << txt;
}

void ProcessWindow::startTimer() {
    this->timer->start(1000);
    qDebug() << "ProcessWindow::startTimer";
}
void ProcessWindow::stopTimer() {
    this->timer->stop();
    this->timeRecord->setHMS(0, 0, 0);
}

void ProcessWindow::clearProcessText() {
    ui->textEdit->clear();
}
