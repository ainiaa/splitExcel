#include "emailonlywindow.h"
#include "ui_emailonlywindow.h"

EmailOnlyWindow::EmailOnlyWindow(QWidget *parent) :
QMainWindow(parent),
ui(new Ui::EmailOnlyWindow)
{
ui->setupUi(this);
}

EmailOnlyWindow::~EmailOnlyWindow()
{
delete ui;
}
