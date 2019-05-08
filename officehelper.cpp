#include "officehelper.h"

OfficeHelper::OfficeHelper()
{

}

/*
void OfficeHelper::process()
{
    QString filepath= "D:\\www\\qt\\build-splitExcel-Desktop_Qt_5_12_0_MinGW_64_bit-Release\\补贴拆分测试-190506.xlsx";//获取保存路径
         QList<QVariant> allRowsData;//保存所有行数据
         allRowsData.clear();
     //    mLstData.append(QVariant(12));
             if(!filepath.isEmpty()){
                 QAxObject *excel = new QAxObject("Excel.Application");//连接Excel控件
                 excel->dynamicCall("SetVisible (bool Visible)",false);//不显示窗体
                 excel->setProperty("DisplayAlerts", true);//不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示
                 QAxObject *workbooks = excel->querySubObject("WorkBooks");//获取工作簿集合
                 workbooks->dynamicCall("Add");//新建一个工作簿
                 QAxObject *workbook = excel->querySubObject("ActiveWorkBook");//获取当前工作簿
                 QAxObject *worksheets = workbook->querySubObject("Sheets");//获取工作表集合
                 QAxObject *worksheet = worksheets->querySubObject("Item(int)",1);//获取工作表集合的工作表1，即sheet1

                     for(int row = 1; row <= 1000; row++)
                         {
                         QList<QVariant> aRowData;//保存一行数据
                        for(int column = 1; column <= 2; column++)
                             {
                                 aRowData.append(QVariant(row*column));
                             }
                         allRowsData.append(QVariant(aRowData));
                     }

                    QAxObject *range = worksheet->querySubObject("Range(const QString )", "A1:B1000");
                 range->dynamicCall("SetValue(const QVariant&)",QVariant(allRowsData));//存储所有数据到 excel 中,批量操作,速度极快
                 range->querySubObject("Font")->setProperty("Size", 30);//设置字号

                    QAxObject *cell = worksheet->querySubObject("Range(QVariant, QVariant)","A1");//获取单元格
                 cell = worksheet->querySubObject("Cells(int, int)", 1, 1);//等同于上一句
                 cell->dynamicCall("SetValue(const QVariant&)",QVariant(123));//存储一个 int 数据到 excel 的单元格中
                 cell->dynamicCall("SetValue(const QVariant&)",QVariant("abc"));//存储一个 string 数据到 excel 的单元格中

                     QString str = cell->dynamicCall("Value2()").toString();//读取单元格中的值
                     qDebug() << "The value of cell is " << str;

                     worksheet->querySubObject("Range(const QString&)", "1:1")->setProperty("RowHeight", 60);//调整第一行行高

                     filepath = "D:\\www\\qt\\build-splitExcel-Desktop_Qt_5_12_0_MinGW_64_bit-Release\\补贴拆分测试-190506.new.xlsx";
                     workbook->dynamicCall("SaveAs(const QString&)",QDir::toNativeSeparators(filepath));//保存至filepath，注意一定要用QDir::toNativeSeparators将路径中的"/"转换为"\"，不然一定保存不了。
                 workbook->dynamicCall("Close()");//关闭工作簿
                 excel->dynamicCall("Quit()");//关闭excel
                 delete excel;
                 excel=nullptr;
             }
}
*/

/*)
void OfficeHelper::process()
{
    // step1：连接控件
    QAxObject excel;
    excel.setControl("Microsoft.Office.Interop.Excel");  // 连接Excel控件
    excel.dynamicCall("SetVisible (bool Visible)", "false"); // 不显示窗体
    excel.setProperty("DisplayAlerts", false);  // 不显示任何警告信息。如果为true, 那么关闭时会出现类似"文件已修改，是否保存"的提示

    // step2: 打开工作簿
    QAxObject* workbooks = excel.querySubObject("WorkBooks"); // 获取工作簿集合
  // 打开工作簿方式一：新建
  //    workbooks->dynamicCall("Add"); // 新建一个工作簿
  //    QAxObject* workbook = excel->querySubObject("ActiveWorkBook"); // 获取当前工作簿
    // 打开工作簿方式二：打开现成
    QString sourceName = "D:\\tmp\\test.xlsx";
    QAxObject* workbook = workbooks->querySubObject("Open(const QString&)",sourceName); // 从控件lineEdit获取文件名

    // step3: 打开sheet
    QAxObject* worksheet = workbook->querySubObject("WorkSheets(int)", 1); // 获取工作表集合的工作表1， 即sheet1


    // step4: 获取行数，列数
    QAxObject* usedrange = worksheet->querySubObject("UsedRange"); // sheet范围
    int intRowStart = usedrange->property("Row").toInt(); // 起始行数
    int intColStart = usedrange->property("Column").toInt();  // 起始列数

    QAxObject *rows, *columns;
    rows = usedrange->querySubObject("Rows");  // 行
    columns = usedrange->querySubObject("Columns");  // 列

    int intRow = rows->property("Count").toInt(); // 行数
    int intCol = columns->property("Count").toInt();  // 列数

    // step5: 读和写
    // 读方式一（坐标）：
    //QAxObject* cell = worksheet->querySubObject("Cells(int, int)", i, j);  //获单元格值
    //qDebug() << i << j << cell->dynamicCall("Value2()").toString();
    // 读方式二（行列名称）：
    //QString X = "A" + QString::number(i + 1); //设置要操作的单元格，A1
    //QAxObject* cellX = worksheet->querySubObject("Range(QVariant, QVariant)", X); //获取单元格
    //qDebug() << cellX->dynamicCall("Value2()").toString();

    // 写方式：
   // cellX->dynamicCall("SetValue(conts QVariant&)", 100); // 设置单元格的值

    // step6: 保存文件
    // 方式一：保存当前文件
    workbook->dynamicCall("Save()");  //保存文件
    workbook->dynamicCall("Close(Boolean)", false);  //关闭文件
    // 方式二：另存为
    QString fileName ="D:\\tmp\\test-abc.xlsx";
    workbook->dynamicCall("SaveAs(const QString&)", QDir::toNativeSeparators(fileName)); //保存到filepath
    // 注意一定要用QDir::toNativeSeparators, 将路径中的"/"转换为"\", 不然一定保存不了
    workbook->dynamicCall("Close (Boolean)", false);  //关闭文件
}
*/
/*
void OfficeHelper::process()
{
    QString filepath= "D:\\www\\qt\\build-splitExcel-Desktop_Qt_5_12_0_MinGW_64_bit-Release\\补贴拆分测试-190506.xlsx";
    QString filepath2= "D:\\www\\qt\\build-splitExcel-Desktop_Qt_5_12_0_MinGW_64_bit-Release\\补贴拆分测试-190506-abc.xlsx";
    QXlsx::Document* xlsx = new QXlsx::Document(filepath);

    xlsx->saveAs(filepath2);
    delete xlsx;
    xlsx = nullptr;
}*/

void OfficeHelper::process()
{
    QString filePath= "D:\\www\\qt\\build-splitExcel-Desktop_Qt_5_12_0_MinGW_64_bit-Release\\补贴拆分测试-190506.xlsx";//获取保存路径
    QList<QVariant> allRowsData;//保存所有行数据
    allRowsData.clear();
    if(!filePath.isEmpty()){
        QAxObject *excel = new QAxObject("Excel.Application");//连接Excel控件
        excel->dynamicCall("SetVisible (bool Visible)",false);//不显示窗体
        excel->setProperty("DisplayAlerts",  false);//不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示
        excel->setProperty("EnableEvents",false); //没有这个 很容易报错  QAxBase: Error calling IDispatch member Open: Unknown error

        QAxObject *workbooks = excel->querySubObject("WorkBooks");//获取工作簿集合
        QAxObject *workbook = workbooks->querySubObject("Open(const QString&, QVariant)", filePath, 0);
        // 获取打开的excel文件中所有的工作sheet
        QAxObject * worksheets = workbook->querySubObject("WorkSheets");
        //—————————————————Excel文件中表的个数:——————————————————
        int iWorkSheet = worksheets->property("Count").toInt();
        qDebug() << QString("Excel文件中表的个数: %1").arg(QString::number(iWorkSheet));
        QAxObject* worksheet = nullptr;
        for (int i = 1; i <= iWorkSheet;i++)
        {
            QAxObject *work_sheet = workbook->querySubObject("WorkSheets(int)", i);  //Sheets(int)也可换用Worksheets(int)
            QString work_sheet_name = work_sheet->property("Name").toString();  //获取工作表名称
            QString message = QString("sheet ")+QString::number(i, 10)+ QString(" name");
            qDebug()<<message<<work_sheet_name;
            if (work_sheet_name != "data")
            {
                qDebug()<<"delete work_sheet:"<<work_sheet_name;
                work_sheet->dynamicCall("Delete()");
            }
            else
            {
                worksheet = work_sheet;
            }
        }

        // ————————————————获取第n个工作表 querySubObject("Item(int)", n);——————————
       // QAxObject * worksheet = worksheets->querySubObject("Item(int)", 1);//本例获取第一个，最后参数填1

        // step3: 打开sheet
        //QAxObject* worksheet = workbook->querySubObject("WorkSheets(int)", 1); // 获取工作表集合的工作表1， 即sheet1
        if (worksheet == nullptr)
        {
            qDebug()<<"worksheet is null";
            return;
        }

        QAxObject* usedrange = worksheet->querySubObject("UsedRange"); // sheet范围
        int intRowStart = usedrange->property("Row").toInt(); // 起始行数
        int intColStart = usedrange->property("Column").toInt();  // 起始列数

        QAxObject *rows, *columns;
        rows = usedrange->querySubObject("Rows");  // 行
        columns = usedrange->querySubObject("Columns");  // 列

        int intRow = rows->property("Count").toInt(); // 行数
        int intCol = columns->property("Count").toInt();  // 列数

        QString alphabetColStart;
        QString alphabetCol;
        ExcelBase::convertToColName(intRowStart,alphabetColStart);
        ExcelBase::convertToColName(intCol,alphabetCol);

        qDebug() << " intRowStart:" <<intRowStart<< " intColStart:" <<intColStart<< " intRow:" <<intRow<< " intCol:" <<intCol << " alphabetColStart:" <<alphabetColStart << " alphabetCol:" << alphabetCol;

        QString X = "A" + QString::number(10 + 1); //设置要操作的单元格，A1
        QString rangeFormat = "A%1:AM%2";
//        for(int row= 1; row <= intRow;row++)
//        {
//            if (row % 5 == 0)
//            {
//                QAxObject* range = worksheet->querySubObject("Range(QVariant)", rangeFormat.arg(row).arg(row)); //获取单元格
//                //qDebug() << " Delete row:" << row;
//                range->dynamicCall("Delete()");
//            }
//        }

        filePath = "D:\\www\\qt\\build-splitExcel-Desktop_Qt_5_12_0_MinGW_64_bit-Release\\补贴拆分测试-190506.test-dd.xlsx";
        workbook->dynamicCall("SaveAs(const QString&)",QDir::toNativeSeparators(filePath));//保存至filepath，注意一定要用QDir::toNativeSeparators将路径中的"/"转换为"\"，不然一定保存不了。
        //workbook->dynamicCall("Save()");
        workbook->dynamicCall("Close()");//关闭工作簿
        workbooks->dynamicCall("Close()");//关闭工作簿
        excel->dynamicCall("Quit()");//关闭excel
        delete excel;
        excel=nullptr;
        qDebug() << "finished";
    }
}

/*
void OfficeHelper::process()
{
    QString filePath= "D:\\www\\qt\\build-splitExcel-Desktop_Qt_5_12_0_MinGW_64_bit-Release\\补贴拆分测试-190506.xlsx";//获取保存路径
    QString filePath2= "D:\\www\\qt\\build-splitExcel-Desktop_Qt_5_12_0_MinGW_64_bit-Release\\补贴拆分测试-190506-excel.xlsx";//获取保存路径
    ExcelBase excel;
    excel.open(filePath);

    excel.saveAs(filePath2);
    excel.close();
}
*/
