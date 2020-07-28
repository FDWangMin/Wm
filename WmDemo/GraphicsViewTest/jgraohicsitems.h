#ifndef JGRAOHICSITEMS_H
#define JGRAOHICSITEMS_H

#include <QGraphicsRectItem>
#include <QGraphicsItemGroup>
#include <QGraphicsProxyWidget>
#include <QGraphicsSceneDragDropEvent>

#include <QPen>
#include <QBrush>
#include <QLabel>
#include <QDebug>

class QGraphicsGridLayout;
class QGraphicsProxyWidget;
class QPushButton;
class QGraphicsView;

class JLabel : public QLabel
{
    Q_OBJECT
public:
    explicit JLabel(QWidget *parent=Q_NULLPTR, Qt::WindowFlags f=Qt::WindowFlags());
    explicit JLabel(const QString &text, QWidget *parent=Q_NULLPTR, Qt::WindowFlags f=Qt::WindowFlags());

    void init();

//    void paintEvent(QPaintEvent * event);

//    void mouseMoveEvent(QMouseEvent * event);
    void mousePressEvent(QMouseEvent * event);
//    void mouseDoubleClickEvent(QMouseEvent * event);
    void mouseReleaseEvent(QMouseEvent * event);

    void dragEnterEvent(QDragEnterEvent *event);
    void dragMoveEvent(QDragMoveEvent *event);
//    void dragLeaveEvent(QDragLeaveEvent *event);
    void dropEvent(QDropEvent *event);

    int getIndex()
    {
        return m_index;
    }
    void setIndex(int index)
    {
        m_index = index;
    }

private:
    QPushButton * m_hoverBtn;
    int     m_index;
    bool    m_selectStatus;
};

class JGraphicsWidget : public QGraphicsWidget
{
public:
    JGraphicsWidget(QGraphicsItem *parent = Q_NULLPTR, Qt::WindowFlags wFlags = Qt::WindowFlags());

//    void dragEnterEvent(QGraphicsSceneDragDropEvent *event);
//    void dragLeaveEvent(QGraphicsSceneDragDropEvent *event);
//    void dragMoveEvent(QGraphicsSceneDragDropEvent *event);
//    void dropEvent(QGraphicsSceneDragDropEvent *event);
};

class JGraphicsRectItem : public QGraphicsRectItem
{
public:
    JGraphicsRectItem(QGraphicsItem *parent = Q_NULLPTR);
    JGraphicsRectItem(const QRectF &rect, QGraphicsItem *parent = Q_NULLPTR);
    JGraphicsRectItem(qreal x, qreal y, qreal w, qreal h, QGraphicsItem *parent = Q_NULLPTR);

    enum SplitScreenMode
    {
        //...
        OneScreen,
        FourSplit,
        NineSplit
        //...
    };

    enum
    {
        JydSplitType = UserType+1
    };

    void init();
    void setPenBrush();
    void splitScreen(SplitScreenMode mode);

    void setGraphicsView(QGraphicsView *view)
    {
        m_view = view;
    }

    void testPlay();

    int type()
    {
        return JydSplitType;
    }

    void paint(QPainter *painter, const QStyleOptionGraphicsItem *option, QWidget *widget);
    void mousePressEvent(QGraphicsSceneMouseEvent *event);
    void mouseMoveEvent(QGraphicsSceneMouseEvent *event);
    void mouseReleaseEvent(QGraphicsSceneMouseEvent *event);
    void mouseDoubleClickEvent(QGraphicsSceneMouseEvent *event);

    void dragEnterEvent(QGraphicsSceneDragDropEvent *event);
    void dragMoveEvent(QGraphicsSceneDragDropEvent *event);
    void dragLeaveEvent(QGraphicsSceneDragDropEvent *event);
    void dropEvent(QGraphicsSceneDragDropEvent *event);

private:
    bool m_mouseButtonPressed;
    QPen m_selectPen;
    QPen m_borderPen;
    QBrush m_backgroundBrush;

    QGraphicsGridLayout *m_gridLayout;
    JGraphicsWidget *m_centerWidget;
    QMap<int, QGraphicsProxyWidget*> m_widgetsMap;

    QList<JLabel*> m_list;

    QGraphicsView *m_view;
};



class JGraphicsProxyWidget : public QGraphicsProxyWidget
{
public:
    JGraphicsProxyWidget(QGraphicsItem *parent = Q_NULLPTR, Qt::WindowFlags wFlags = Qt::WindowFlags()):
        QGraphicsProxyWidget(parent, wFlags)
    {
        setAcceptDrops(true);
        setAcceptedMouseButtons(Qt::LeftButton);
        setFlag(QGraphicsItem::ItemIsSelectable, true);
    }

//    void dragEnterEvent(QGraphicsSceneDragDropEvent *event)
//    {
////        event->accept();
//        event->acceptProposedAction();
//        qDebug() << "JGraphicsProxyWidget::acceptDrops() = " << acceptDrops();
//        event->setAccepted(true);
//        qDebug() << "JGraphicsProxyWidget::dragEnterEvent";
//        QGraphicsProxyWidget::dragEnterEvent(event);
//    }
//    void dragLeaveEvent(QGraphicsSceneDragDropEvent *event)
//    {
//        qDebug() << "JGraphicsProxyWidget::dragLeaveEvent";
//        QGraphicsProxyWidget::dragLeaveEvent(event);
//    }
//    void dragMoveEvent(QGraphicsSceneDragDropEvent *event)
//    {
//        qDebug() << "JGraphicsProxyWidget::dragMoveEvent";
//        QGraphicsProxyWidget::dragMoveEvent(event);
//    }
//    void dropEvent(QGraphicsSceneDragDropEvent *event)
//    {
//        qDebug() << "JGraphicsProxyWidget::dropEvent(QGraphicsSceneDragDropEvent *event)";
//        QGraphicsProxyWidget::dropEvent(event);
//    }
};

#endif // JGRAOHICSITEMS_H
