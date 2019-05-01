#ifndef EMAILTASKQUEUE_H
#define EMAILTASKQUEUE_H
#include <QStringList>
#include "emailtaskqueuedata.h"

class EmailTaskQueue
{
public:
    EmailTaskQueue();
    void setTaskData(EmailTaskQueueData data);
    EmailTaskQueueData  getTaskData();
private:
    EmailTaskQueueData emailKeyList;
};

#endif // EMAILTASKQUEUE_H
