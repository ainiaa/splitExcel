#include "xlsxparser.h"
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

void XlsxParserRunnable::setSplitData(QString sourcePath, QString selectedSheetName, QHash<QString, QList<int>> fragmentDataQhash, QString savePath, int total)
{
    qDebug("setSplitData") ;
    this->sourcePath = sourcePath;
    this->selectedSheetName = selectedSheetName;
    this->fragmentDataQhash = fragmentDataQhash;
    this->savePath = savePath;
    this->m_total = total;
}

void XlsxParserRunnable::run()
{
    qDebug("XlsxParserRunnable::run start") ;
    QString startMsg("开始拆分excel并生成新的excel文件: %1/");
    QString endMsg("完成拆分excel并生成新的excel文件: %1/");
    startMsg.append(QString::number(m_total));
    endMsg.append(QString::number(m_total));
     requestMsg(Common::MsgTypeInfo, startMsg.arg(QString::number(runnableID)));
    QHashIterator<QString,QList<int>> it(fragmentDataQhash);
    while (it.hasNext()) {
        it.next();
        QString key = it.key();
        QList<int> contentList = it.value();
        contentList.insert(0,1);

        QString xlsName;
       // int rows =  contentList.size();
        xlsName.append(savePath).append(QDir::separator()).append(key).append(".xlsx");
        //xlsx.saveAs(xlsName);
        copyFileToPath(sourcePath,xlsName,true);

        QXlsx::Document* currXls = new QXlsx::Document(xlsName);

        currXls->fliterRows(contentList);//删除多余的行  多线程的生活会出现问题。

        //删除多余的sheet start
        QStringList sheetNames = currXls->sheetNames();
        for (int i = 0;i<sheetNames.size();i++)
        {
            QString currSheetName = sheetNames.at(i);
            if (selectedSheetName == currSheetName)
            {
                continue;
            }
            currXls->deleteSheet(currSheetName);
        }
        //删除多余的sheet end

        //尝试修复行高不正确的问题
        QXlsx::CellRange cellRange = currXls->dimension();
        for (int row_num = cellRange.firstRow(); row_num <= cellRange.lastRow(); row_num++) {
            double rowHeigh = currXls->rowHeight(row_num);
            if (rowHeigh < 14)
            {
                currXls->setRowHeight(row_num, 18);
            }
        }

        currXls->save();
        delete currXls;
    }
    requestMsg(Common::MsgTypeSucc, endMsg.arg(QString::number(runnableID)));
    qDebug("XlsxParserRunnable::run end") ;
}

void XlsxParserRunnable::requestMsg(const int msgType, const QString &msg)
{
     qobject_cast<XlsxParser *>(mParent)->receiveMessage(msgType,msg.arg(msgType));
    //QMetaObject::invokeMethod(mParent, "receiveMessage", Qt::QueuedConnection, Q_ARG(int,msgType),Q_ARG(QString, msg));//不能及时返回信息
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

