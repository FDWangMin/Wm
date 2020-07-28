#include "jtablemodelviewdelegate.h"

//JStandardItem
JStandardItem::JStandardItem()
{

}

JStandardItem::JStandardItem(const QString &text)
    : QStandardItem(text)
{

}

JStandardItem::JStandardItem(const QIcon &icon, const QString &text)
    : QStandardItem(icon, text)
{

}

JStandardItem::JStandardItem(int rows, int columns)
    : QStandardItem(rows, columns)
{

}

JStandardItem::~JStandardItem()
{

}

//bool JStandardItem::operator<(const QStandardItem& other) const
//{
//    const QVariant left = data(Qt::DisplayRole), right = other.data(Qt::DisplayRole);
//    //   第1到2列，全部采用浮点数的大小排序
//    if (column() == other.column() && other.column() >= 1 && other.column() <= 2)
//    {
//        return left.toDouble() < right.toDouble();
//    }
//    return QStandardItem::operator<(other);
//}

//JStandardItemModel
JStandardItemModel::JStandardItemModel(QWidget *parent) :
    QStandardItemModel(parent)
{

}

JStandardItemModel::JStandardItemModel(int rows, int columns, QObject *parent) :
    QStandardItemModel(rows, columns, parent)
{

}

JStandardItemModel::~JStandardItemModel()
{

}

//JTableView
JTableView::JTableView(QWidget *parent)
{

}

JTableView::~JTableView()
{

}

//JItemDelegate
JItemDelegate::JItemDelegate(QWidget *parent)
{

}

JItemDelegate::~JItemDelegate()
{

}


