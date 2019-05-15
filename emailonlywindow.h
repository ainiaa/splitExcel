#ifndef EMAILONLYWINDOW_H
#define EMAILONLYWINDOW_H

#include "splitsubwindow.h"
#include <QMainWindow>

namespace Ui {
class EmailOnlyWindow;
}

class EmailOnlyWindow : public SplitSubWindow {
    Q_OBJECT

    public:
    explicit EmailOnlyWindow(QMainWindow *parent = nullptr);
    ~EmailOnlyWindow();

    private:
    Ui::EmailOnlyWindow *ui;
};

#endif // EMAILONLYWINDOW_H
