#ifndef QAX_MS_EXECEL_IO_H
#define QAX_MS_EXECEL_IO_H

#include <QObject>

class QAxObject;

class MsExcelIO : public QObject
{
    Q_OBJECT
public:

    MsExcelIO(){}
    ~MsExcelIO(){}

    static QVariant readExcel(QString fileName, int sheetID);
    static bool writeExcel(QString fileName, int sheetID, QVariant var);

    static void castVariant2listListVariant(const QVariant &var, QList<QList<QVariant> > &res);
    static void castListListVariant2Variant(const QList<QList<QVariant> > &cells, QVariant &res);
    static void convertToColName(int data, QString &res);
    static QString to26AlphabetString(int data);

    static QScopedPointer<QAxObject> m_excel;
};

#endif // QAX_MS_EXECEL_IO_H
