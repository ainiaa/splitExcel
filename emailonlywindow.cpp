#include "emailonlywindow.h"
#include "ui_emailonlywindow.h"

EmailOnlyWindow::EmailOnlyWindow(QMainWindow *parent) : SplitSubWindow(parent), ui(new Ui::EmailOnlyWindow) {
    ui->setupUi(this);
}

EmailOnlyWindow::~EmailOnlyWindow() {
    delete ui;
}
