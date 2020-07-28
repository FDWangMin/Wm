#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>

class QActionGroup;

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();

    void init();
    void initTabWidgets();
    void initFileOperationActions();
    void initCheckedActions();

private:
    Ui::MainWindow *ui;

    QActionGroup *m_actGroup;
};

#endif // MAINWINDOW_H
