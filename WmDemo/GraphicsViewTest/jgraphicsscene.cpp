#include "jgraphicsscene.h"

#include <QEvent>
#include <QDebug>
#include <QGraphicsSceneMouseEvent>
#include <QGraphicsItem>

JGraphicsScene::JGraphicsScene(QObject *parent):QGraphicsScene(parent)
{

}

JGraphicsScene::JGraphicsScene(const QRectF &sceneRect, QObject *parent):
    QGraphicsScene(sceneRect, parent)
{
    init();
}

JGraphicsScene::JGraphicsScene(qreal x, qreal y, qreal width, qreal height, QObject *parent):
    QGraphicsScene(x, y, width, height, parent)
{
    init();
}

void JGraphicsScene::init()
{

}

bool JGraphicsScene::event(QEvent *event)
{
//    qDebug() << "JGraphicsScene::event(QEvent *event)" << event->type();
    return QGraphicsScene::event(event);
}

void JGraphicsScene::dragEnterEvent(QGraphicsSceneDragDropEvent *event)
{
    event->acceptProposedAction();
    qDebug() << "JGraphicsScene::dragEnterEvent(QGraphicsSceneDragDropEvent *event)";
    QGraphicsScene::dragEnterEvent(event);
}

void JGraphicsScene::dragMoveEvent(QGraphicsSceneDragDropEvent *event)
{
    qDebug() << "JGraphicsScene::dragMoveEvent(QGraphicsSceneDragDropEvent *event)";
    QGraphicsScene::dragMoveEvent(event);
}

void JGraphicsScene::dragLeaveEvent(QGraphicsSceneDragDropEvent *event)
{
    qDebug() << "JGraphicsScene::dragLeaveEvent(QGraphicsSceneDragDropEvent *event)";
    QGraphicsScene::dragLeaveEvent(event);
}

void JGraphicsScene::dropEvent(QGraphicsSceneDragDropEvent *event)
{
    qDebug() << "JGraphicsScene::dropEvent(QGraphicsSceneDragDropEvent *event)";
    QGraphicsScene::dropEvent(event);
}

void JGraphicsScene::mousePressEvent(QGraphicsSceneMouseEvent *event)
{
    qDebug() << "===JGraphicsScene::mousePressEvent(QGraphicsSceneMouseEvent *event)===\n"
             << event->button()
             << event->buttonDownPos(event->button())
             << event->buttonDownScenePos(event->button())
             << event->buttonDownScreenPos(event->button());

    QGraphicsScene::mousePressEvent(event);

//    if (event->button() == Qt::LeftButton)
//    {
//        if(items(event->scenePos()).count() > 0)
//        {
//            QGraphicsItem *selectItem = items(event->scenePos()).first();
//            foreach (QGraphicsItem *item, selectedItems())
//            {
//                if (item == selectItem)
//                    continue;
//                item->setSelected(false);
//            }
//            qDebug() << "selectItem->setSelected(!selectItem->isSelected());" << "!" << selectItem->isSelected();
//            selectItem->setSelected(!selectItem->isSelected());
//        }
//        else
//        {
//            foreach (QGraphicsItem *item, items())
//            {
//                item->setSelected(false);
//            }
//        }
//    }
}

void JGraphicsScene::drawBackground(QPainter *painter, const QRectF &rect)
{
    QGraphicsScene::drawBackground(painter, rect);
}

void JGraphicsScene::drawForeground(QPainter *painter, const QRectF &rect)
{
    QGraphicsScene::drawForeground(painter, rect);
}

void JGraphicsScene::drawItems(QPainter *painter, int numItems, QGraphicsItem *items[], const QStyleOptionGraphicsItem options[], QWidget *widget)
{
    QGraphicsScene::drawItems(painter, numItems, items, options, widget);
}
