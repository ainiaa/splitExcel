#ifndef PROCESSWINDOW_H
#define PROCESSWINDOW_H

#include <QMainWindow>
#include <QDateTime>
#include <QErrorMessage>
namespace Ui {
class ProcessWindow;
}

class ProcessWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit ProcessWindow(QWidget *parent = nullptr);
    ~ProcessWindow();

    void setProcessText(QString text);
    void clearProcessText();

private:
    Ui::ProcessWindow *ui;
};

#endif // PROCESSWINDOW_H
