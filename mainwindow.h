#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QDateTime>
#include <QDir>
#include <QErrorMessage>
#include <QFileDialog>
#include <QHash>
#include <QHashIterator>
#include <QMainWindow>
#include <QMessageBox>
#include <QThread>

#include <QException>

#include "common.h"
#include "config.h"
#include "configsetting.h"
#include "emailonlywindow.h"
#include "emailsender.h"
#include "excelparser.h"
#include "officehelper.h"
#include "processwindow.h"
#include "sourceexceldata.h"
#include "splitandemailwindow.h"
#include "splitonlywindow.h"
#include "testtimeer.h"
#include "xlsxdocument.h"
namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow {
    Q_OBJECT

    public:
    explicit MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

    signals:

    private slots:

    void on_splitOnlyPushButton_clicked();

    void on_EmailOnlyPushButton_clicked();

    void on_SplitAndEmailPushButton_clicked();

    private:
    Ui::MainWindow *ui;
    SplitOnlyWindow *splitOnlyWindow = nullptr;
    SplitAndEmailWindow *splitAndEmailWindow = nullptr;
    EmailOnlyWindow *emailOnlyWindow = nullptr;
    void errorMessage(const QString &message);
};

#endif // MAINWINDOW_H
