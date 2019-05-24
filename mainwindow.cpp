#include "mainwindow.h"
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent) : QMainWindow(parent), ui(new Ui::MainWindow) {
    ui->setupUi(this);
    setWindowFlags(windowFlags() & ~Qt::WindowMaximizeButtonHint);
    setFixedSize(this->width(), this->height());

    connect(ui->actionConfig_Setting, SIGNAL(triggered()), this, SLOT(showConfigSetting()));
}

MainWindow::~MainWindow() {
    qDebug() << "MainWindow::~MainWindow start";
    delete ui;
    if (this->emailOnlyWindow != nullptr) {
        delete this->emailOnlyWindow;
        this->emailOnlyWindow = nullptr;
    }
    if (this->splitOnlyWindow != nullptr) {
        delete this->splitOnlyWindow;
        this->splitOnlyWindow = nullptr;
    }
    if (this->splitAndEmailWindow != nullptr) {
        delete this->splitAndEmailWindow;
        this->splitAndEmailWindow = nullptr;
    }
    if (this->configSetting != nullptr) {
        delete this->configSetting;
        this->configSetting = nullptr;
    }
    qDebug() << "MainWindow::~MainWindow end";
}

void MainWindow::on_splitOnlyPushButton_clicked() {
    if (this->splitOnlyWindow == nullptr) {
        this->splitOnlyWindow = new SplitOnlyWindow(this);
    }

    this->splitOnlyWindow->show();
    this->hide();
}

void MainWindow::on_EmailOnlyPushButton_clicked() {

    if (this->emailOnlyWindow == nullptr) {
        this->emailOnlyWindow = new EmailOnlyWindow(this);
    }
    this->emailOnlyWindow->show();
    this->hide();
}

void MainWindow::on_SplitAndEmailPushButton_clicked() {

    if (this->splitAndEmailWindow == nullptr) {
        this->splitAndEmailWindow = new SplitAndEmailWindow(this);
    }
    this->splitAndEmailWindow->show();
    this->hide();
}

void MainWindow::showConfigSetting() {
    configSetting->show();
    this->hide();
}
