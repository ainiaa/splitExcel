#include "splitsubwindow.h"
#include <QCloseEvent>

SplitSubWindow::SplitSubWindow(QMainWindow *parent) : QMainWindow(parent) {
    this->mainWindow = parent;
}

void SplitSubWindow::closeEvent(QCloseEvent *e) {

    e->ignore();

    this->hide();
    this->getMainWindow()->show();
}

QMainWindow *SplitSubWindow::getMainWindow() {
    return this->mainWindow;
}
