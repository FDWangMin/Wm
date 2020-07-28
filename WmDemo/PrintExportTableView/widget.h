#ifndef WIDGET_H
#define WIDGET_H

#include <QWidget>
#include <QFile>
#include "jtablemodelviewdelegate.h"

namespace Ui {
class Widget;
}

class Widget : public QWidget
{
    Q_OBJECT

public:
    explicit Widget(QWidget *parent = 0);
    ~Widget();

    void initTableModelView();

    void inputExcel(QFile &file, QString data);



private slots:
    void on_pb_exportCSV_clicked();

    void on_pb_print_clicked();

    void on_pb_exportXLS_clicked();

    void on_pushButton_clicked();

private:
    Ui::Widget *ui;
    JStandardItemModel  *m_tbModel;
    JTableView          *m_tbView;
};

#endif // WIDGET_H
