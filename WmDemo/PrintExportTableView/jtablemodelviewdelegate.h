#ifndef JTABLEMODELVIEWDELEGATE_H
#define JTABLEMODELVIEWDELEGATE_H

#include <QStandardItemModel>
#include <QStandardItem>
#include <QTableView>
#include <QItemDelegate>
#include <QPair>

class JStandardItem : public QStandardItem
{
    JStandardItem();
    JStandardItem(const QString &text);
    JStandardItem(const QIcon &icon, const QString &text);
    JStandardItem(int rows, int columns = 1);
    virtual ~JStandardItem();

};

class JStandardItemModel : public QStandardItemModel
{
public:
    JStandardItemModel(QWidget *parent = Q_NULLPTR);
    JStandardItemModel(int rows, int columns, QObject *parent = Q_NULLPTR);
    ~JStandardItemModel();

private:
    int m_endItemSum;
    QPair<int, int> m_endItemXY;
};

class JTableView : public QTableView
{
public:
    JTableView(QWidget *parent = Q_NULLPTR);
    ~JTableView();
};

class JItemDelegate : public QItemDelegate
{
public:
    JItemDelegate(QWidget *parent = Q_NULLPTR);
    ~JItemDelegate();
};


#endif // JTABLEMODELVIEWDELEGATE_H
