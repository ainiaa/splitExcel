#include "idlemsgutil.h"

IdleMsgUtil::IdleMsgUtil(QObject *parent) : QObject(parent)
{

}

void IdleMsgUtil::startIdleMsg()
{

}

void IdleMsgUtil::processIdleMsg()
{

}


void IdleMsgUtil::finishIdleMsg()
{

}


void IdleMsgUtil::setTimeUnit(int timeUnit)
{
    this->timeUnit = timeUnit;
}

int IdleMsgUtil::getTimeUnit()
{
    return this->timeUnit;
}
void IdleMsgUtil::setLeftTime(int leftTime)
{
    this->leftTime = leftTime;
}
int IdleMsgUtil::getLeftTime()
{
    return this->leftTime;
}
void IdleMsgUtil::setProcessInterval(int processInterval)
{
    this->processInterval = processInterval;
}
int IdleMsgUtil::getProcessInterval()
{
    return this->processInterval;
}
void IdleMsgUtil::setIdleTimer(QElapsedTimer idleTimer)
{
    this->idleTimer = idleTimer;
}

QElapsedTimer IdleMsgUtil::getIdleTimer()
{
    return this->idleTimer;
}
