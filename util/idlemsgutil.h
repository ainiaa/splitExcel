#ifndef IDLEMSGUTIL_H
#define IDLEMSGUTIL_H

#include <QObject>
#include <QElapsedTimer>

class IdleMsgUtil : public QObject
{
Q_OBJECT
public:
explicit IdleMsgUtil(QObject *parent = nullptr);

void setTimeUnit(int timeUnit);
int getTimeUnit();
void setLeftTime(int leftTime);
int getLeftTime();
void setProcessInterval(int processInterval);
int getProcessInterval();
void setIdleTimer(QElapsedTimer idleTimer);
QElapsedTimer getIdleTimer();

signals:

public slots:
    void startIdleMsg();
    void processIdleMsg();
    void finishIdleMsg();

private:
    int timeUnit; //时间单位
    int leftTime;//剩余时间
    int processInterval;//处理间隔
    QElapsedTimer idleTimer;//timer
};

#endif // IDLEMSGUTIL_H
