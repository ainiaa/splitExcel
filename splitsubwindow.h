#ifndef SPLITSUBWINDOW_H
#define SPLITSUBWINDOW_H

#include <QMainWindow>

class SplitSubWindow : public QMainWindow {
    Q_OBJECT
    public:
    explicit SplitSubWindow(QMainWindow *parent = nullptr);
    QMainWindow *getMainWindow();
    signals:

    public slots:

    protected:
    void closeEvent(QCloseEvent *event);

    private:
    QMainWindow *mainWindow;
};

#endif // SPLITSUBWINDOW_H
