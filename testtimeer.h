#ifndef TESTTIMEER_H
#define TESTTIMEER_H

#include "common.h"
#include "processwindow.h"
#include <QElapsedTimer>
#include <QMainWindow>
#include <QTimer>

namespace Ui {
class TestTimeer;
}

class TestTimeer : public QMainWindow {
    Q_OBJECT

    public:
    explicit TestTimeer(QWidget *parent = nullptr, QMainWindow *mainWindow = nullptr);
    ~TestTimeer();
    void receiveMessage(const int msgType, const QString &result);

    public slots:
    void showIdleMsg();

    private slots:
    void on_pushButton_clicked();

    private:
    Ui::TestTimeer *ui;
    ProcessWindow *processWindow = nullptr;
    QMainWindow *mainWindow;
    int leftTime;
    int idleMsgShowPeriod;
    int timeUnit;
    QElapsedTimer timer;
};

#endif // TESTTIMEER_H
