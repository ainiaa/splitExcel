#include "emailtaskqueue.h"

EmailTaskQueue::EmailTaskQueue()
{

}

void EmailTaskQueue::setEmailKeyList(QStringList emailKeyList)
{
    this->emailKeyList = emailKeyList;
}


QStringList  EmailTaskQueue::getEmailKeyList()
{
    return this->emailKeyList;
}
