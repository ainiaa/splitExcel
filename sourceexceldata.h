#ifndef SOURCEXMLDATA_H
#define SOURCEXMLDATA_H

#include <QObject>

class SourceExcelData : public QObject {
    Q_OBJECT

    public:
    SourceExcelData(QObject *parent = nullptr);
    enum OperateType { SplitAndEmailType = 1, SplitOnlyType };

    void setDataSheetIndex(int dataSheetIndex);
    int getDataSheetIndex();
    void setEmailSheetIndex(int emailSheetIndex);
    int getEmailSheetIndex();
    void setSheetCnt(int sheetCnt);
    int getSheetCnt();
    void setGroupByText(QString groupByText);
    QString getGroupByText();
    void setDataSheetName(QString dataSheetName);
    QString getDataSheetName();
    void setEmailSheetName(QString emailSheetName);
    QString getEmailSheetName();
    void setSavePath(QString savePath);
    QString getSavePath();
    void setSourcePath(QString sourcePath);
    QString getSourcePath();
    void setOpType(int opType);
    int getOpType();

    private:
    int dataSheetIndex = 0;           //数据sheet index
    int emailSheetIndex = 0;          // email sheet index
    int sheetCnt = 0;                 // sheet 总数
    QString groupByText = nullptr;    //分组名
    QString dataSheetName = nullptr;  //数据sheet title
    QString emailSheetName = nullptr; //数据sheet title
    QString savePath = nullptr;       // excel文件保存路径
    QString sourcePath = nullptr;     //源excel文件路径

    int opType;
};

#endif // SOURCEXMLDATA_H
