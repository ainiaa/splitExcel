#ifndef SOURCEXMLDATA_H
#define SOURCEXMLDATA_H

#include <QObject>

class SourceExcelData : public QObject {
    Q_OBJECT

    public:
    SourceExcelData(QObject *parent = nullptr);
    enum OperateType { SplitAndEmailType = 1, SplitOnlyType, EmailOnlyType };

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

    void setSourceMinAlphabetCol(QString sourceMinAlphabetCol);
    QString getSourceMinAlphabetCol();
    void setSourceMaxAlphabetCol(QString sourceMaxAlphabetCol);
    QString getSourceMaxAlphabetCol();
    void setTplXlsPath(QString tplXlsPath);
    QString getTplXlsPath();
    void setSourceRowStart(int sourceRowStart);
    int getSourceRowStart();
    void setSourceColStart(int sourceColStart);
    int setSourceColStart();
    void setSourceRowCnt(int sourceRowCnt);
    int getSourceRowCnt();
    void setSourceColCnt(int sourceColCnt);
    int getSourceColCnt();

    private:
    int dataSheetIndex = 0;           //数据sheet index
    int emailSheetIndex = 0;          // email sheet index
    int sheetCnt = 0;                 // sheet 总数
    QString groupByText = nullptr;    //分组名
    QString dataSheetName = nullptr;  //数据sheet title
    QString emailSheetName = nullptr; //数据sheet title
    QString savePath = nullptr;       // excel文件保存路径
    QString sourcePath = nullptr;     //源excel文件路径

    int sourceRowStart = 0;       //起始行数
    int sourceColStart = 0;       //起始列数
    int sourceRowCnt = 0;         //行数
    int sourceColCnt = 0;         //列数
    QString sourceMinAlphabetCol; //
    QString sourceMaxAlphabetCol; //
    QString tplXlsPath;

    int opType;
};

#endif // SOURCEXMLDATA_H
