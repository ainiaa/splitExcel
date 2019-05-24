#ifndef SPLITSUBWINDOW_H
#define SPLITSUBWINDOW_H

#include <QMainWindow>

class SplitSubWindow : public QMainWindow {
    Q_OBJECT
    public:
    explicit SplitSubWindow(QMainWindow *parent = nullptr);
    QMainWindow *getMainWindow();
    int getTotalCnt();
    void setTotalCnt(int cnt);
    int getProcessCnt();
    void setProcessCnt(int cnt);
    void incrProcessCnt(int step);
    int incrAndCalcPercent(int step);
    int calcProcessPercent();

    signals:

    public slots:

    protected:
    void closeEvent(QCloseEvent *event);

    private:
    QMainWindow *mainWindow;
    int total_cnt = 0;
    int process_cnt = 0;
};

#endif // SPLITSUBWINDOW_H
