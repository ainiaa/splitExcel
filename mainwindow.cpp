#include "mainwindow.h"
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent) : QMainWindow(parent), ui(new Ui::MainWindow) {
    ui->setupUi(this);
    setWindowFlags(windowFlags() & ~Qt::WindowMaximizeButtonHint);
    setFixedSize(this->width(), this->height());

    connect(ui->actionConfig_Setting, SIGNAL(triggered()), this, SLOT(showConfigSetting()));
    connect(ui->actionSplit_Only, SIGNAL(triggered()), this, SLOT(showSplitOnly()));
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
    qDebug() << "MainWindow::~MainWindow end";
}

void MainWindow::on_splitOnlyPushButton_clicked() {
    this->hide();
    if (this->splitOnlyWindow == nullptr) {
        this->splitOnlyWindow = new SplitOnlyWindow(this);
    }

    this->splitOnlyWindow->show();
}

void MainWindow::on_EmailOnlyPushButton_clicked() {
    this->hide();
    if (this->emailOnlyWindow == nullptr) {
        this->emailOnlyWindow = new EmailOnlyWindow(this);
    }
    this->emailOnlyWindow->show();
}

void MainWindow::on_SplitAndEmailPushButton_clicked() {
    this->hide();
    if (this->splitAndEmailWindow == nullptr) {
        this->splitAndEmailWindow = new SplitAndEmailWindow(this);
    }
    this->splitAndEmailWindow->show();
}
