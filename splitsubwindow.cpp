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

int SplitSubWindow::getTotalCnt() {
    return this->total_cnt;
}

void SplitSubWindow::setTotalCnt(int cnt) {
    this->total_cnt = cnt;
}

int SplitSubWindow::getProcessCnt() {
    return this->process_cnt;
}

void SplitSubWindow::setProcessCnt(int cnt) {
    this->process_cnt = cnt;
}

void SplitSubWindow::incrProcessCnt(int step) {
    this->process_cnt += step;
}

int SplitSubWindow::calcProcessPercent() {
    if (this->total_cnt > 0) {
        return this->process_cnt * 100 / this->total_cnt;
    }
    return 0;
}

int SplitSubWindow::incrAndCalcPercent(int step) {
    this->incrProcessCnt(step);
    return this->calcProcessPercent();
}
