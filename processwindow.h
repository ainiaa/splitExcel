#ifndef PROCESSWINDOW_H
#define PROCESSWINDOW_H

#include <QDateTime>
#include <QErrorMessage>
#include <QMainWindow>
#include <QTime>
#include <QTimer>
namespace Ui {
class ProcessWindow;
}

class ProcessWindow : public QMainWindow {
    Q_OBJECT

    public:
    explicit ProcessWindow(QWidget *parent = nullptr);
    ~ProcessWindow();

    void setProcessText(QString text);
    void setProcessPercent(int percent);
    void startTimer();
    void stopTimer();
    void clearProcessText();

    public slots:
    void updateTime();

    private:
    Ui::ProcessWindow *ui;
    QTimer *timer;
    QTime *timeRecord;
};

#endif // PROCESSWINDOW_H
