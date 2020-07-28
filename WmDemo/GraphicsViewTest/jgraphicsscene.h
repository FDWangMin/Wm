#ifndef JGRAPHICSSCENE_H
#define JGRAPHICSSCENE_H

#include <QGraphicsScene>

class JGraphicsScene : public QGraphicsScene
{
public:
    JGraphicsScene(QObject *parent = Q_NULLPTR);
    JGraphicsScene(const QRectF &sceneRect, QObject *parent = Q_NULLPTR);
    JGraphicsScene(qreal x, qreal y, qreal width, qreal height, QObject *parent = Q_NULLPTR);

    void init();

protected:
    bool event(QEvent *event);
//    bool eventFilter(QObject *watched, QEvent *event);
//    void contextMenuEvent(QGraphicsSceneContextMenuEvent *event);
    void dragEnterEvent(QGraphicsSceneDragDropEvent *event);
    void dragMoveEvent(QGraphicsSceneDragDropEvent *event);
    void dragLeaveEvent(QGraphicsSceneDragDropEvent *event);
    void dropEvent(QGraphicsSceneDragDropEvent *event);
//    void focusInEvent(QFocusEvent *event);
//    void focusOutEvent(QFocusEvent *event);
//    void helpEvent(QGraphicsSceneHelpEvent *event);
//    void keyPressEvent(QKeyEvent *event);
//    void keyReleaseEvent(QKeyEvent *event);
    void mousePressEvent(QGraphicsSceneMouseEvent *event);
//    void mouseMoveEvent(QGraphicsSceneMouseEvent *event);
//    void mouseReleaseEvent(QGraphicsSceneMouseEvent *event);
//    void mouseDoubleClickEvent(QGraphicsSceneMouseEvent *event);
//    void wheelEvent(QGraphicsSceneWheelEvent *event);
//    void inputMethodEvent(QInputMethodEvent *event);

    void drawBackground(QPainter *painter, const QRectF &rect);
    void drawForeground(QPainter *painter, const QRectF &rect);
    void drawItems(QPainter *painter, int numItems,
                           QGraphicsItem *items[],
                           const QStyleOptionGraphicsItem options[],
                           QWidget *widget = Q_NULLPTR);
};

#endif // JGRAPHICSSCENE_H
