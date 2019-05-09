#ifndef OFFICEHELPER_H
#define OFFICEHELPER_H

#include "excelbase.h"
#include "qdebug.h"
#include "qstringutil.h"
#if defined(Q_OS_WIN)
#include "qt_windows.h"
#endif
#include "xlsxdocument.h"
#include <QAxObject>
#include <QDir>

class OfficeHelper {
    public:
    OfficeHelper();
    bool static isInstalledExcelApp();
    void static convertToColName(int data, QString &res);
    QString static to26AlphabetString(int data);
};

#endif // OFFICEHELPER_H
