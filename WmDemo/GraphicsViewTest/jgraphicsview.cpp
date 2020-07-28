#include "jgraphicsview.h"

#include <QWheelEvent>
#include <QDebug>
#include <qmath.h>
#include <QGraphicsItem>

JGraphicsView::JGraphicsView(QWidget *parent):QGraphicsView(parent),
    m_tileWH(50),
    m_backgroundScale(false)
{
    init();
}

JGraphicsView::JGraphicsView(QGraphicsScene *scene, QWidget *parent):
    m_tileWH(50),
    m_backgroundScale(false)
{
    init();
}

JGraphicsView::~JGraphicsView()
{

}

void JGraphicsView::init()
{
    setAcceptDrops(true);
    setTransformationAnchor(QGraphicsView::QGraphicsView::AnchorViewCenter);
    setDragMode(QGraphicsView::NoDrag);
    setViewportUpdateMode(QGraphicsView::FullViewportUpdate);
    setFrameShape(QFrame::NoFrame);
    setRenderHint(QPainter::SmoothPixmapTransform);
}

void JGraphicsView::cleanItems()
{
    scene()->clear();
    resetTransform();
}

void JGraphicsView::setBackgroundStyle(JGraphicsView::BackgroundStyle style, QColor color)
{
    switch (style)
    {
    case JGraphicsView::NoStyle:
    {
        QPixmap tilePixmap(m_tileWH, m_tileWH);
        tilePixmap.fill(Qt::white);
        QPainter tilePainter(&tilePixmap);
        tilePainter.setPen(Qt::white);
        tilePainter.drawRect(0, 0, m_tileWH, m_tileWH);
        tilePainter.end();
        setBackgroundBrush(tilePixmap);
        break;
    }
    case JGraphicsView::ColorGridStyle: GridStyle:
    {
        QPixmap tilePixmap(m_tileWH, m_tileWH);
        tilePixmap.fill(Qt::white);
        QPainter tilePainter(&tilePixmap);
        tilePainter.setPen(color);
        tilePainter.drawRect(0, 0, m_tileWH, m_tileWH);
        tilePainter.end();
        setBackgroundBrush(tilePixmap);
        break;
    }
    case JGraphicsView::BrushGridStyle:
    {
        // Prepare background check-board pattern
        int pixmapWh = m_tileWH*2;
        QPixmap tilePixmap(pixmapWh, pixmapWh);
        tilePixmap.fill(Qt::white);
        QPainter tilePainter(&tilePixmap);
        tilePainter.fillRect(0, 0, m_tileWH, m_tileWH, color);
        tilePainter.fillRect(m_tileWH, m_tileWH, m_tileWH, m_tileWH, color);
        tilePainter.end();
        setBackgroundBrush(tilePixmap);
        break;
    }
    case JGraphicsView::PureColorStyle:
    {
        QPixmap tilePixmap(m_tileWH, m_tileWH);
        tilePixmap.fill(color);
        setBackgroundBrush(tilePixmap);
    }
    default:
        break;
    }
}

void JGraphicsView::setBackgroundTileWH(int wh)
{
    m_tileWH = wh;
}

void JGraphicsView::setBackgroundScale(bool scale)
{
    m_backgroundScale = scale;
}

void JGraphicsView::wheelEvent(QWheelEvent *event)
{
    qDebug() << "--------------------------------------------------\n";
    qDebug() << "delta()=" << event->delta() << " angleDelta()=" << event->angleDelta() << " we->pixelDelta()=" << event->pixelDelta();
    qreal factor = qPow(1.2, event->delta() / 240.0);
    qDebug() << "factor = qPow(1.2, we->delta() / 240.0) = " << qPow(1.2, event->delta() / 240.0);
    qreal currentScale = transform().m11();
    qDebug() << "currentScale=m_graphicsView->transform().m11()=" << transform().m11();
    qreal newScale = currentScale * factor;
    qDebug() << "newScale = currentScale * factor =" << currentScale * factor;;
    qreal actualFactor = factor;
    // cap to 0.1 - 10
    if (newScale > 10)
        actualFactor = 10./currentScale;
    else if (newScale < 0.1)
        actualFactor = 0.1/currentScale;
    scale(actualFactor, actualFactor);
//            m_graphicsView->rotate(15);
//            m_graphicsView->shear(0.01,0.01);
//            m_graphicsView->translate(100,0);
//    qDebug() << "actualFactor=" << actualFactor << "\n";
//            m_graphicsView->fitInView(QRectF(0,0,(1920*2)*m_scale,(1080*2)*m_scale));
}

void JGraphicsView::drawBackground(QPainter *painter, const QRectF &rect)
{
    if (m_backgroundScale)
        return QGraphicsView::drawBackground(painter, rect);

    painter->save();
    painter->resetTransform();
    painter->drawTiledPixmap(viewport()->rect(), backgroundBrush().texture());
    painter->restore();
}

void JGraphicsView::mousePressEvent(QMouseEvent *event)
{
    qDebug() << "---JGraphicsView::mousePressEvent---" << event->pos() << event->screenPos() ;
    QGraphicsView::mousePressEvent(event);
}

void JGraphicsView::dragEnterEvent(QDragEnterEvent *event)
{
    qDebug() << "JGraphicsView::dragEnterEvent(QDragEnterEvent *event)";
    event->acceptProposedAction();
    QGraphicsView::dragEnterEvent(event);
}

void JGraphicsView::dragLeaveEvent(QDragLeaveEvent *event)
{
    qDebug() << "JGraphicsView::dragLeaveEvent(QDragEnterEvent *event)";
    QGraphicsView::dragLeaveEvent(event);
}

void JGraphicsView::dragMoveEvent(QDragMoveEvent *event)
{
//    event->acceptProposedAction();
    qDebug() << "JGraphicsView::dragMoveEvent(QDragEnterEvent *event)";
    QGraphicsView::dragMoveEvent(event);
}

void JGraphicsView::dropEvent(QDropEvent *event)
{
    qDebug() << "JGraphicsView::dropEvent(QDragEnterEvent *event)";
    QGraphicsView::dropEvent(event);
}
