#ifndef JGRAPHICSVIEW_H
#define JGRAPHICSVIEW_H

#include <QGraphicsView>

class JGraphicsView : public QGraphicsView
{
public:
    enum BackgroundStyle
    {
        NoStyle,
        ColorGridStyle,
        BrushGridStyle,
        PureColorStyle
    };

    JGraphicsView(QWidget *parent = Q_NULLPTR);
    JGraphicsView(QGraphicsScene *scene, QWidget *parent = Q_NULLPTR);
    ~JGraphicsView();

    void init();
    void cleanItems();

    void setBackgroundStyle(BackgroundStyle style=NoStyle, QColor color=QColor(Qt::white));
    void setBackgroundTileWH(int wh);
    void setBackgroundScale(bool scale);

    void wheelEvent(QWheelEvent *event);
    void drawBackground(QPainter *painter, const QRectF &rect);

//    bool event(QEvent *event) Q_DECL_OVERRIDE;
//    bool viewportEvent(QEvent *event) Q_DECL_OVERRIDE;

//    void focusInEvent(QFocusEvent *event) Q_DECL_OVERRIDE;
//    bool focusNextPrevChild(bool next) Q_DECL_OVERRIDE;
//    void focusOutEvent(QFocusEvent *event) Q_DECL_OVERRIDE;
//    void keyPressEvent(QKeyEvent *event) Q_DECL_OVERRIDE;
//    void keyReleaseEvent(QKeyEvent *event) Q_DECL_OVERRIDE;
//    void mouseDoubleClickEvent(QMouseEvent *event) Q_DECL_OVERRIDE;
    void mousePressEvent(QMouseEvent *event);
//    void mouseMoveEvent(QMouseEvent *event) Q_DECL_OVERRIDE;
//    void mouseReleaseEvent(QMouseEvent *event) Q_DECL_OVERRIDE;

    void dragEnterEvent(QDragEnterEvent *event);
    void dragLeaveEvent(QDragLeaveEvent *event);
    void dragMoveEvent(QDragMoveEvent *event);
    void dropEvent(QDropEvent *event);

//    void paintEvent(QPaintEvent *event) Q_DECL_OVERRIDE;
//    void resizeEvent(QResizeEvent *event) Q_DECL_OVERRIDE;
//    void scrollContentsBy(int dx, int dy) Q_DECL_OVERRIDE;
//    void showEvent(QShowEvent *event) Q_DECL_OVERRIDE;
//    void inputMethodEvent(QInputMethodEvent *event) Q_DECL_OVERRIDE;

//    void drawBackground(QPainter *painter, const QRectF &rect);
//    void drawForeground(QPainter *painter, const QRectF &rect);
//    void drawItems(QPainter *painter, int numItems,
//                           QGraphicsItem *items[],
//                           const QStyleOptionGraphicsItem options[]);


private:
    int m_tileWH;
    bool m_backgroundScale;

};

#endif // JGRAPHICSVIEW_H
