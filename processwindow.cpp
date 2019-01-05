#include "processwindow.h"
#include "ui_processwindow.h"

ProcessWindow::ProcessWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::ProcessWindow)
{
    ui->setupUi(this);
    setWindowFlags(Qt::Dialog);
    setFixedSize(this->width(), this->height());
}

ProcessWindow::~ProcessWindow()
{
    delete ui;
}

void ProcessWindow::setProcessText(QString text)
{
    QDateTime currentTime = QDateTime::currentDateTime();
    QString strCurrentTime = currentTime.toString("hh:mm:ss");
    ui->textEdit->append(strCurrentTime +"  " + text);
}

void ProcessWindow::clearProcessText()
{
    ui->textEdit->clear();
}


