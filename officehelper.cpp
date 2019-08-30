#include "officehelper.h"

OfficeHelper::OfficeHelper() {
}

/**
 * 是否安装了office
 * @brief OfficeHelper::isInstalledExcelApp
 * @return
 */
bool OfficeHelper::isInstalledExcelApp() {
    CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    QAxObject *excel = new QAxObject("Excel.Application");
    if (excel->isNull()) {
        return false;
    }
    excel->dynamicCall("Quit()");
    delete excel;
    excel = nullptr;
    // CoUninitialize();
    return true;
}

///
/// \brief 把列数转换为excel的字母列号
/// \param data 大于0的数
/// \return 字母列号，如1->A 26->Z 27 AA
///
void OfficeHelper::convertToColName(int data, QString &res) {
    int tempData = data / 26;
    int mode = data % 26;
    if (tempData > 0 && mode > 0) {
        convertToColName(mode, res);
        convertToColName(tempData, res);
    } else {
        res = (to26AlphabetString(data) + res);
    }
}
///
/// \brief 数字转换为26字母
///
/// 1->A 26->Z
/// \param data
/// \return
///
QString OfficeHelper::to26AlphabetString(int data) {
    QChar ch = data + 0x40; // A对应0x41
    return QString(ch);
}
