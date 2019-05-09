#include "officehelper.h"

OfficeHelper::OfficeHelper() {
}

/**
 * 是否安装了office
 * @brief OfficeHelper::isInstalledExcelApp
 * @return
 */
bool OfficeHelper::isInstalledExcelApp() {
    QAxObject *excel = new QAxObject("Excel.Application");
    if (excel->isNull()) {
        return false;
    }
    delete excel;
    excel = nullptr;
    return true;
}
