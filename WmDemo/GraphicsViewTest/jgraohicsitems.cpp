#include "jgraohicsitems.h"
#include "jgraphicsscene.h"

#include <QGraphicsSceneMouseEvent>
#include <QStyleOptionGraphicsItem>
#include <QGraphicsGridLayout>
#include <QGraphicsProxyWidget>

#include <QPen>
#include <QDebug>
#include <QPainter>
#include <QLabel>
#include <QMovie>
#include <QPushButton>
#include <QMimeData>

JLabel::JLabel(QWidget *parent, Qt::WindowFlags f):QLabel(parent,f),
    m_hoverBtn(NULL),
    m_index(0),
    m_selectStatus(false)
{
    init();
}

JLabel::JLabel(const QString &text, QWidget *parent, Qt::WindowFlags f):QLabel(text, parent, f),
    m_hoverBtn(NULL),
    m_index(0),
    m_selectStatus(false)
{
    init();
}

void JLabel::init()
{
//    setMouseTracking(true);
    setAcceptDrops(true);
}

//void JLabel::mouseMoveEvent(QMouseEvent *event)
//{
//    if (m_hoverBtn == NULL)
//    {
//        m_hoverBtn = new QPushButton("hoverButton",this);
//        m_hoverBtn->setWindowFlags(windowFlags() | Qt::Tool);
//    }

//    if (!rect().contains(event->pos()))
//    {
//        m_hoverBtn->hide();
//        return;
//    }

//    m_hoverBtn->move(x()/2, y()/2);
//    m_hoverBtn->show();
//}

void JLabel::mousePressEvent(QMouseEvent *event)
{
    qDebug() << "void JLabel::mousePressEvent(QMouseEvent *event)";
    QLabel::mousePressEvent(event);
}

void JLabel::mouseReleaseEvent(QMouseEvent *event)
{
    qDebug() << "void JLabel::mouseReleaseEvent(QMouseEvent *event)";
    QLabel::mouseReleaseEvent(event);
}

void JLabel::dragEnterEvent(QDragEnterEvent *event)
{
    qDebug() << "JLabel::acceptDrops() = " << acceptDrops();
    event->setDropAction(Qt::CopyAction);
    event->accept();
    qDebug() << "void JLabel::dragEnterEvent(QDragEnterEvent *event)";
//    QLabel::dragEnterEvent(event);
}

void JLabel::dragMoveEvent(QDragMoveEvent *event)
{
    event->setDropAction(Qt::CopyAction);
    event->accept();
//    QLabel::dragMoveEvent(event);
}

//void JLabel::dragLeaveEvent(QDragLeaveEvent *event)
//{
//    QLabel::dragLeaveEvent(event);
//}

void JLabel::dropEvent(QDropEvent *event)
{
    qDebug() << "JLabel::dropEvent(QDropEvent *event)";
    event->setDropAction(Qt::CopyAction);
    event->accept();
//    QLabel::dropEvent(event);
}

JGraphicsRectItem::JGraphicsRectItem(QGraphicsItem *parent):QGraphicsRectItem(parent),
    m_mouseButtonPressed(false),
    m_centerWidget(NULL),
    m_gridLayout(NULL),
    m_view(NULL)
{
    init();
}

JGraphicsRectItem::JGraphicsRectItem(const QRectF &rect, QGraphicsItem *parent):
    QGraphicsRectItem(rect, parent),
    m_mouseButtonPressed(false),
    m_centerWidget(NULL),
    m_gridLayout(NULL),
    m_view(NULL)
{
    init();
}

JGraphicsRectItem::JGraphicsRectItem(qreal x, qreal y, qreal w, qreal h, QGraphicsItem *parent):
    QGraphicsRectItem(x, y, w, h, parent),
    m_mouseButtonPressed(false),
    m_centerWidget(NULL),
    m_gridLayout(NULL),
    m_view(NULL)
{
    init();
}

void JGraphicsRectItem::init()
{
    m_centerWidget = new JGraphicsWidget(this);
    m_centerWidget->setGeometry(this->boundingRect().adjusted(2,2,-20,-20));
//    m_centerWidget->setAcceptDrops(true);
//    m_centerWidget->setAcceptedMouseButtons(Qt::LeftButton);
//    m_centerWidget->setAcceptHoverEvents(true);
    m_gridLayout = new QGraphicsGridLayout;

    m_centerWidget->setAutoFillBackground(true);
    m_centerWidget->setLayout(m_gridLayout);

    setPenBrush();
    setFlag(QGraphicsItem::ItemIsSelectable);// 可选择
setFlag(QGraphicsItem::ItemIsMovable);
    setAcceptDrops(true);
    setAcceptedMouseButtons(Qt::LeftButton);
//    setAcceptHoverEvents(true);
//    setAcceptedMouseButtons(Qt::LeftButton);
}

void JGraphicsRectItem::setPenBrush()
{
    // 画笔 - 边框色
    QPen sp = pen();
    sp.setWidth(1);
    sp.setColor(Qt::red);
    sp.setStyle(Qt::DashLine);
    m_selectPen = sp;

    QPen bp = pen();
    bp.setWidth(2);
    bp.setColor(Qt::gray);
    m_borderPen = bp;
    setPen(m_borderPen);

    // 画刷 - 背景色
    m_backgroundBrush = QBrush(/*QColor(0, 160, 230)*/);
    setBrush(m_backgroundBrush);
}

void JGraphicsRectItem::paint(QPainter *painter, const QStyleOptionGraphicsItem *option, QWidget *widget)
{
    QStyleOptionGraphicsItem jOption;
    jOption.init(widget);
    if (option->state & QStyle::State_Selected)
        jOption.state = QStyle::State_None;

    QGraphicsRectItem::paint(painter, &jOption, widget);

    if (option->state & QStyle::State_Selected)
    {
        qreal itemPenWidth = pen().widthF();
        const qreal pad = itemPenWidth / 2;

        // 绘制选中边线
        painter->setPen(m_selectPen);
        painter->drawRect(boundingRect().adjusted(pad, pad, -pad, -pad));
    }
}

void JGraphicsRectItem::splitScreen(JGraphicsRectItem::SplitScreenMode mode)
{
    switch (mode)
    {
    case OneScreen:
    {
        if (!m_list.isEmpty())
        {
            foreach (JLabel *l, m_list) {
                l->deleteLater();
            }
            m_list.clear();
        }
        JLabel *label = new JLabel(QString("JYD_One"));
        m_list.append(label);
        label->setPixmap(QPixmap(":/JYD_logo.png"));
        label->setStyleSheet("border:1px solid black");
        label->setAcceptDrops(true);
        label->setAlignment(Qt::AlignCenter);
        JGraphicsProxyWidget *gProxyWidget = new JGraphicsProxyWidget(this);
        gProxyWidget->setWidget(label);
        qDebug() << ">>>gProxyWidget->acceptDrops():" << gProxyWidget->acceptDrops();
//        gProxyWidget->setAcceptedMouseButtons(Qt::LeftButton);
//        gProxyWidget->setAcceptHoverEvents(true);
        m_gridLayout->addItem(gProxyWidget,0,0);
        gProxyWidget->show();
        break;
    }
    case FourSplit:
    {
        if (!m_list.isEmpty())
        {
            foreach (JLabel *l, m_list) {
                l->deleteLater();
            }
            m_list.clear();
        }
        for (int i = 0; i < 2; i++)
        {
            for (int j = 0; j < 2; j++)
            {
                JLabel *label = new JLabel(QString("JYD_%1").arg(QString::number((i+1)*(j+1))));
                m_list.append(label);
                label->setPixmap(QPixmap(":/JYD_logo.png"));
                label->setStyleSheet("border:1px solid black");
//                label->setAcceptDrops(true);
                label->setAlignment(Qt::AlignCenter);
                JGraphicsProxyWidget *gProxyWidget = new JGraphicsProxyWidget(this);
                gProxyWidget->setWidget(label);
//                gProxyWidget->setAcceptedMouseButtons(Qt::LeftButton);
//                gProxyWidget->setAcceptHoverEvents(true);
                gProxyWidget->show();
                m_gridLayout->addItem(gProxyWidget,i,j);
            }
        }
        break;
    }
    case NineSplit:
    {
        if (!m_list.isEmpty())
        {
            foreach (JLabel *l, m_list) {
                l->deleteLater();
            }
            m_list.clear();
        }
        for (int i = 0; i < 3; i++)
        {
            for (int j = 0; j < 3; j++)
            {
                JLabel *label = new JLabel(QString("JYD_%1").arg(QString::number((i+1)*(j+1))));
                m_list.append(label);
                label->setStyleSheet("border:1px solid black");
                label->setPixmap(QPixmap(":/JYD_logo.png"));
//                label->setAcceptDrops(true);
                label->setAlignment(Qt::AlignCenter);
                JGraphicsProxyWidget *gProxyWidget = new JGraphicsProxyWidget(this);
                gProxyWidget->setWidget(label);
//                gProxyWidget->setAcceptedMouseButtons(Qt::LeftButton);
//                gProxyWidget->setAcceptHoverEvents(true);
                gProxyWidget->show();
                m_gridLayout->addItem(gProxyWidget,i,j);
            }
        }
        break;
    }
    //...
    default:
        break;
    }

}

void JGraphicsRectItem::testPlay()
{
    static bool which = false;
    QString fileName;
    if (which)
        fileName = ":/test1.gif";
    else
        fileName = ":/test2.gif";
    QMovie *move = new QMovie(fileName);
    m_list.first()->setMovie(move);
    move->setParent(m_list.first());
    move->start();
    which = !which;
}

void JGraphicsRectItem::mousePressEvent(QGraphicsSceneMouseEvent *event)
{
    m_mouseButtonPressed = true;
    qDebug() << "-=-=-=JGraphicsRectItem::mousePressEvent=-=-=-";
    QGraphicsRectItem::mousePressEvent(event);
}

void JGraphicsRectItem::mouseMoveEvent(QGraphicsSceneMouseEvent *event)
{
    qDebug() << "JGraphicsRectItem::mouseMoveEvent";
    QGraphicsRectItem::mouseMoveEvent(event);
}

void JGraphicsRectItem::mouseReleaseEvent(QGraphicsSceneMouseEvent *event)
{
    m_mouseButtonPressed = false;
    qDebug() << "JGraphicsRectItem::mouseReleaseEvent";
    QGraphicsRectItem::mouseReleaseEvent(event);
}

void JGraphicsRectItem::mouseDoubleClickEvent(QGraphicsSceneMouseEvent *event)
{
    qDebug() << "JGraphicsRectItem::mouseDoubleClickEvent";
    QGraphicsRectItem::mouseDoubleClickEvent(event);
}

void JGraphicsRectItem::dragEnterEvent(QGraphicsSceneDragDropEvent *event)
{
//    if (event->mimeData()->hasText())
//    {
//        event->setAccepted(true);
//        update();
//    }
//    else
//    {
//        event->setAccepted(false);
//    }
    qDebug() << "JGraphicsRectItem::dragEnterEvent(QGraphicsSceneDragDropEvent *event)";
    QGraphicsRectItem::dragEnterEvent(event);
}

void JGraphicsRectItem::dragMoveEvent(QGraphicsSceneDragDropEvent *event)
{
    qDebug() << "JGraphicsRectItem::dragMoveEvent(QGraphicsSceneDragDropEvent *event)";
    QGraphicsRectItem::dragMoveEvent(event);
}

void JGraphicsRectItem::dragLeaveEvent(QGraphicsSceneDragDropEvent *event)
{
    qDebug() << "JGraphicsRectItem::dragLeaveEvent(QGraphicsSceneDragDropEvent *event)";
    QGraphicsRectItem::dragLeaveEvent(event);
}

void JGraphicsRectItem::dropEvent(QGraphicsSceneDragDropEvent *event)
{
//    if (event->mimeData()->hasText())
//    {
//        update();
    QByteArray encoded = event->mimeData()->data("application/x-qabstractitemmodeldatalist");
    QDataStream stream(&encoded, QIODevice::ReadOnly);
    QMap<int, QVariant> v;
    int r, c;
    stream >> r >> c >> v;
    qDebug() << "event->mimeData()->hasText():" << r << c << v ;
//    }
//    else
//    {
//        QGraphicsRectItem::dropEvent(event);
//    }
    qDebug() << "JGraphicsRectItem::dropEvent";
    QGraphicsRectItem::dropEvent(event);
}

JGraphicsWidget::JGraphicsWidget(QGraphicsItem *parent, Qt::WindowFlags wFlags):
    QGraphicsWidget(parent, wFlags)
{

}

//void JGraphicsWidget::dragEnterEvent(QGraphicsSceneDragDropEvent *event)
//{
//    qDebug() << "void JGraphicsWidget::dragEnterEvent(QGraphicsSceneDragDropEvent *event)";
//    QGraphicsWidget::dragEnterEvent(event);
//}

//void JGraphicsWidget::dragLeaveEvent(QGraphicsSceneDragDropEvent *event)
//{
//    qDebug() << "void JGraphicsWidget::dragLeaveEvent(QGraphicsSceneDragDropEvent *event)";
//    QGraphicsWidget::dragLeaveEvent(event);
//}

//void JGraphicsWidget::dragMoveEvent(QGraphicsSceneDragDropEvent *event)
//{
//    qDebug() << "void JGraphicsWidget::dragMoveEvent(QGraphicsSceneDragDropEvent *event)";
//    QGraphicsWidget::dragMoveEvent(event);
//}

//void JGraphicsWidget::dropEvent(QGraphicsSceneDragDropEvent *event)
//{
//    qDebug() << "void JGraphicsWidget::dropEvent(QGraphicsSceneDragDropEvent *event)";
//    QGraphicsWidget::dropEvent(event);
//}
