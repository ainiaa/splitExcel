#ifndef OFFICEHELPER_H
#define OFFICEHELPER_H

#include<QAxObject>
#include<QDir>
#include "qdebug.h"
#include "xlsxdocument.h"
#include "qstringutil.h"
#include "excelbase.h"

class OfficeHelper
{
public:
OfficeHelper();
void process();
};

#endif // OFFICEHELPER_H
