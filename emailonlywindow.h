#ifndef EMAILONLYWINDOW_H
#define EMAILONLYWINDOW_H

#include <QMainWindow>

namespace Ui {
class EmailOnlyWindow;
}

class EmailOnlyWindow : public QMainWindow
{
Q_OBJECT

public:
explicit EmailOnlyWindow(QWidget *parent = nullptr);
~EmailOnlyWindow();

private:
Ui::EmailOnlyWindow *ui;
};

#endif // EMAILONLYWINDOW_H
