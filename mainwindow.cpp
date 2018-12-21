#include "mainwindow.h"
#include "ui_mainwindow.h"


MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow),
    xlsx(nullptr)
{
    ui->setupUi(this);

}

MainWindow::~MainWindow()
{
    delete ui;
}

//选择文件
void MainWindow::on_selectFilePushButton_clicked()
{
    QString path = MainWindow::openFile();
    if (path.length() > 0)
    {//选择了excel文件
        ui->xlsObjLineEdit->setText(path);
        //QMessageBox::information(this, "Title", path, QMessageBox::Yes | QMessageBox::No, QMessageBox::Yes);
        xlsx = new QXlsx::Document (path);

        //获得表头数据
        QXlsx::CellRange range;
        range = xlsx->dimension();
        int colCount = range.columnCount();

        for (int colum=1; colum<=colCount; ++colum) {
            QXlsx::Cell *cell =xlsx->cellAt(1, colum);
            if (cell) {
                qDebug(cell->value().toByteArray());
                header->append(cell->value().toString());
            }
        }
        ui->groupByComboBox->addItems(*header);

        //获得所有的sheets
        QStringList sheetNames = xlsx->sheetNames();
        ui->dataComboBox->addItems(sheetNames);
        ui->dataComboBox->setCurrentIndex(0);
        ui->emailComboBox->addItems(sheetNames);
        ui->emailComboBox->setCurrentIndex(1);
    }
}

//打开文件对话框
QString MainWindow::openFile()
{
    //    QFileDialog* dlg = new QFileDialog(this);
    //    dlg->setWindowTitle("选择xls文件...");
    //    dlg->setDirectory(".");
    //    dlg->setFileMode(QFileDialog::ExistingFile);
    //    dlg->setNameFilter( tr("Excel Files(*.xls)"));
    //    dlg->setViewMode(QFileDialog::Detail);

    QString path = QFileDialog::getOpenFileName(this, tr("Open Excel"), ".", tr("excel(*.xlsx)"));
    if(path.length() == 0) {
        return "";
    } else {
        return path;
    }
}

//选择保存目录
void MainWindow::on_savePathPushButton_clicked()
{
    QString path = QFileDialog::getExistingDirectory(this, tr("Save New Excels"), ".");
    if(path.length() > 0) {
        ui->savePathLineEdit->setText(path);
    }
}

//确定拆分
void MainWindow::on_submitPushButton_clicked()
{
    QString savePath = ui->savePathLineEdit->text();
    if (xlsx == nullptr)
    {
        QMessageBox::information(this, "Save Path Error", "请选择待拆分的excel");
    }
    if (savePath.length() == 0)
    {
        QMessageBox::information(this, "Save Path Error", "请选择保存路径");
    }
    QString groupby;
    int groupbyindex =ui->groupByComboBox->currentIndex() + 1;
    groupby.append(QString::number(groupbyindex));
    groupby.append(":").append(ui->groupByComboBox->currentText());
    //QMessageBox::information(this, "分组", groupby);

    //拆分
    doSplitXls(groupbyindex,savePath);
}

//拆分
void MainWindow::doSplitXls(int groupby, QString savePath)
{
    //读取excel数据
    QHash<QString, QList<QStringList>> qhash = readXls(groupby);

    //写excel
    writeXls(qhash,savePath);
}

//读取xls
QHash<QString, QList<QStringList>> MainWindow::readXls(int groupby)
{
    QXlsx::CellRange range;
    range = xlsx->dimension();
    int rowCount = range.rowCount();
    int colCount = range.columnCount();

    QHash<QString, QList<QStringList>> qhash;
    for (int row = 2;row <= rowCount;++row)
    {
        QString groupbyvalue;
        QStringList items;
        QXlsx::Cell *cell =xlsx->cellAt(row, groupby);
        if (cell)
        {
            groupbyvalue = cell->value().toString();
        }
        for (int colum=1; colum<=colCount; ++colum)
        {
            QXlsx::Cell *cell =xlsx->cellAt(row, colum);
            if (cell)
            {
                // qDebug(cell->value().toByteArray());
                if (cell->isDateTime())
                {
                    items.append(cell->dateTime().toString("yyyy/MM/dd"));
                }
                else
                {
                    items.append(cell->value().toString());
                }
            }
        }
        QList<QStringList> qlist = qhash.take(groupbyvalue);
        qlist.append(items);
        qhash.insert(groupbyvalue,qlist);
    }
    return qhash;
}

//写xls
void MainWindow::writeXls(QHash<QString, QList<QStringList>> qhash, QString savePath)
{
    qDebug("writeXls");
    QHashIterator<QString,QList<QStringList>> it(qhash);
    while (it.hasNext()) {
        it.next();
        //        qDebug("key:");
        QString key = it.key();
        //        qDebug(key.toUtf8());
        QList<QStringList> content = it.value();
        QXlsx::Document currXls;

        //写表头
        writeXlsHeader(&currXls);

        QString xlsName;
        int rows =  content.size();
        xlsName.append(savePath).append("\\").append(key).append("-").append(QString::number(rows)).append(".xlsx");
        //        qDebug("key:");
        //        qDebug(xlsName.toUtf8());
        //        qDebug("contentsize:");
        //        qDebug(QString::number(rows).toUtf8());
        for(int row = 0; row < rows;row++)
        {
//            qDebug("current row:");
//            qDebug(QString::number(row).toUtf8());
            QStringList qsl = content.at(row);
            int columns = qsl.size();
            for (int column =0;column < columns;column++)
            {
                QString columnName;
                convertToColName(column+1,columnName) ;
                QString cell;
                cell.append(columnName).append(QString::number(row+2));
                qDebug("cell:");
                qDebug(cell.toUtf8());
                qDebug("cell-content:");
                qDebug(qsl.at(column).toUtf8());
                currXls.write(cell,qsl.at(column));
            }
        }
        currXls.saveAs(xlsName);
    }
}

void  MainWindow::writeXlsHeader(QXlsx::Document *xls)
{
    int columns =header->size();
    for (int column =0;column < columns;column++)
    {
        QString columnName;
        convertToColName(column+1,columnName) ;
        QString cell;
        cell.append(columnName).append(QString::number(1));
        xls->write(cell,header->at(column));
    }
}

///
/// \brief 把列数转换为excel的字母列号
/// \param data 大于0的数
/// \return 字母列号，如1->A 26->Z 27 AA
///
void MainWindow::convertToColName(int data, QString &res)
{
    Q_ASSERT(data>0 && data<65535);
    int tempData = data / 26;
    if(tempData > 0)
    {
        int mode = data % 26;
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
QString MainWindow::to26AlphabetString(int data)
{
    QChar ch = data + 0x40;//A对应0x41
    return QString(ch);
}
