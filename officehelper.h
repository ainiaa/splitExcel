#ifndef OFFICEHELPER_H
#define OFFICEHELPER_H

#include "excelbase.h"
#include "qdebug.h"
#include "qstringutil.h"
#include "xlsxdocument.h"
#include <QAxObject>
#include <QDir>

class OfficeHelper {
    public:
    OfficeHelper();
    void process();
    bool static isInstalledExcelApp();
};

#endif // OFFICEHELPER_H
