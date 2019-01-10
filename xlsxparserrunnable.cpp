#include "xlsxparserrunnable.h"

XlsxParserRunnable::XlsxParserRunnable(QObject *parent)
{
    this->mParent = parent;
}

XlsxParserRunnable::~XlsxParserRunnable()
{
    runnableID = 0;
}

void XlsxParserRunnable::setID(const int &id)
{
    runnableID = id;
}

void XlsxParserRunnable::setSplitData(QString sourcePath, QString selectedSheetName, QString key, QList<int> contentList, QString savePath, int total)
{
    qDebug("setSplitData") ;
    this->sourcePath = sourcePath;
    this->selectedSheetName = selectedSheetName;
    this->key = key;
    this->contentList = contentList;
    this->savePath = savePath;
    this->m_total = total;
    contentList.insert(0,1);
    this->contentList = contentList;
}

void XlsxParserRunnable::run()
{
    qDebug("XlsxParserRunnable::run start") ;
    QXlsx::Document xlsx(sourcePath);
    QXlsx::CellRange range;
    xlsx.selectSheet(selectedSheetName);

    QString startMsg("开始拆分excel并生成新的excel文件: %1/");
    QString endMsg("完成拆分excel并生成新的excel文件: %1/");
    startMsg.append(QString::number(m_total));
    endMsg.append(QString::number(m_total));
    requestMsg(Common::MsgTypeInfo, startMsg.arg(QString::number(runnableID)));

    QString xlsName;
   // int rows =  contentList.size();
    xlsName.append(savePath).append(QDir::separator()).append(key).append(".xlsx");
    //xlsx.saveAs(xlsName);
    copyFileToPath(sourcePath,xlsName,true);

    QXlsx::Document currXls(xlsName);

    currXls.fliterRows(contentList);//删除多余的行

    //删除多余的sheet start
    QStringList sheetNames = currXls.sheetNames();
    for (int i = 0;i<sheetNames.size();i++)
    {
        QString currSheetName = sheetNames.at(i);
        if (selectedSheetName == currSheetName)
        {
            continue;
        }
        currXls.deleteSheet(currSheetName);
    }
    //删除多余的sheet end
    currXls.save();
    requestMsg(Common::MsgTypeSucc, endMsg.arg(QString::number(runnableID)));

    qDebug("EmailSenderRunnable::run end") ;
}

void XlsxParserRunnable::requestMsg(const int msgType, const QString &msg)
{
    QMetaObject::invokeMethod(mParent, "receiveMessage", Qt::QueuedConnection, Q_ARG(int,msgType),Q_ARG(QString, msg));
}

///
/// \brief 把列数转换为excel的字母列号
/// \param data 大于0的数
/// \return 字母列号，如1->A 26->Z 27 AA
///
void XlsxParserRunnable::convertToColName(int data, QString &res)
{
    int tempData = data / 26;
    int mode = data % 26;
    if(tempData > 0 && mode > 0)
    {
        convertToColName(mode,res);
        convertToColName(tempData,res);
    }
    else
    {
        res=(to26AlphabetString(data)+res);
    }
}
///
/// \brief 数字转换为26字母
///
/// 1->A 26->Z
/// \param data
/// \return
///
QString XlsxParserRunnable::to26AlphabetString(int data)
{
    QChar ch = data + 0x40;//A对应0x41
    return QString(ch);
}

bool XlsxParserRunnable::copyFileToPath(QString sourceDir ,QString toDir, bool coverFileIfExist)
{
    sourceDir = QDir::toNativeSeparators(sourceDir);
    toDir = QDir::toNativeSeparators(toDir);

    if (sourceDir == toDir){
        return true;
    }
    if (!QFile::exists(sourceDir)){
        return false;
    }
    QDir *createfile     = new QDir();
    bool exist = createfile->exists(toDir);
    if (exist){
        if(coverFileIfExist){
            createfile->remove(toDir);
        }
    }//end if

    if(!QFile::copy(sourceDir, toDir))
    {
        return false;
    }
    return true;
}

