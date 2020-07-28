#include "widget.h"
#include "ui_widget.h"
#include <QFileDialog>
#include <QTextDecoder>
#include <QtPrintSupport/QPrinter>
#include <QtPrintSupport/QPrintDialog>
#include <QPainter>
#include <QMessageBox>
#include <QAxObject>
#include <QDebug>

Widget::Widget(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::Widget)
{
    ui->setupUi(this);
    initTableModelView();
}

Widget::~Widget()
{
    delete ui;
}

void Widget::initTableModelView()
{
    m_tbModel = new JStandardItemModel(this);
    m_tbModel->setColumnCount(5);
    m_tbModel->setRowCount(100);

    m_tbView = ui->tableView;
    m_tbView->setModel(m_tbModel);
}

void Widget::on_pb_print_clicked()
{
    QPrinter printerPixmap;
    QPixmap pixmap = QPixmap::grabWidget(ui->tableView, ui->tableView->rect());  //获取界面的图片

    QPrintDialog print(&printerPixmap, this);
    if (print.exec())
    {
        QPainter painterPixmap;
        painterPixmap.begin(&printerPixmap);
        QRect rect = painterPixmap.viewport();
        int x = rect.width() / pixmap.width();
        int y = rect.height() / pixmap.height();
        painterPixmap.scale(x, y);
        painterPixmap.drawPixmap(0, 0, pixmap);
        painterPixmap.end();
    }
}

void Widget::inputExcel(QFile &file, QString data)
{


}

void Widget::on_pb_exportCSV_clicked()
{

    QString fileName = QFileDialog::getSaveFileName(this, tr("Save File"), "", tr("file ( *.csv)"));
    if(fileName == "")
         return;

    QFile file;
    file.setFileName(fileName);
    file.open(QIODevice::WriteOnly);

    QTextCodec *code;
    code = QTextCodec::codecForName("UTF-8");

    for(int i = 0; i < m_tbModel->rowCount(); i++)
    {
        for(int j = 0; j < m_tbModel->columnCount(); j++)
        {
            if (!(m_tbModel->data(m_tbModel->index(i,j)).isNull()))
            {
                std::string strCountBuffer = code->fromUnicode(m_tbModel->item(i,j)->text()).data();
                file.write(strCountBuffer.c_str(), qstrlen(strCountBuffer.c_str()));
            }
            file.write(",");
        }
        file.write("\n");
    }

    file.close();
}

void Widget::on_pb_exportXLS_clicked()
{
//    QMessageBox::about(this, "导出XLS", "该功能暂未实现");

/*    // step1：连接控件
    QAxObject* excel = new QAxObject();
    excel->setControl("Excel.Application");  // 连接Excel控件
    excel->dynamicCall("SetVisible (bool Visible)", "false"); // 不显示窗体
    excel->setProperty("DisplayAlerts", false);  // 不显示任何警告信息。如果为true, 那么关闭时会出现类似"文件已修改，是否保存"的提示

    // step2: 打开工作簿
    QAxObject* workbook = excel->querySubObject("WorkBooks"); // 获取工作簿集合
    // 打开工作簿方式一：新建
    workbook->dynamicCall("Add"); // 新建一个工作簿

//    QAxObject* workbook = excel->querySubObject("ActiveWorkBook"); // 获取当前工作簿
    // 打开工作簿方式二：打开现成
//    QAxObject* workbook = workbooks->querySubObject("Open(const QString&)", ("C:/Users/lixc/Desktop/tt2.xlsx")); // 从控件lineEdit获取文件名

    // step3: 打开sheet
    QAxObject* worksheet = workbook->querySubObject("WorkSheets(int)", 1); // 获取工作表集合的工作表1， 即sheet1


    // step4: 获取行数，列数
    QAxObject* usedrange = worksheet->querySubObject("UsedRange"); // sheet范围
    int intRowStart = usedrange->property("Row").toInt(); // 起始行数   为1
    int intColStart = usedrange->property("Column").toInt();  // 起始列数 为1

    QAxObject *rows, *columns;
    rows = usedrange->querySubObject("Rows");  // 行
    columns = usedrange->querySubObject("Columns");  // 列

    int intRow = rows->property("Count").toInt(); // 行数
    int intCol = columns->property("Count").toInt();  // 列数
     qDebug()<<"intRowStart:"<<intRowStart<<"\t intColStart"<<intColStart;
    qDebug()<<"intRow"<<intRow<<"\t intCol"<<intCol;
    // step5: 读和写
    // 读方式一（坐标）：
//    for(int i=intRowStart;i<intRow+intRowStart;i++)
//    {
//        for(int j=intColStart;j<intCol+intColStart;j++)
//        {
//            QAxObject* cell = worksheet->querySubObject("Cells(int, int)", i, j);  //获单元格值
//            qDebug() << i << j << cell->dynamicCall("Value2()").toString();
//        }
//    }

//    // 读方式二（行列名称）：
    QString X = "A2"; //设置要操作的单元格，A1
    QAxObject* cellX = worksheet->querySubObject("Range(QVariant, QVariant)", X); //获取单元格
    qDebug() << cellX->dynamicCall("Value2()").toString();

//    // 写方式：
    cellX->dynamicCall("SetValue(conts QVariant&)", 100); // 设置单元格的值
    QAxObject *cell_5_6 = worksheet->querySubObject("Cells(int,int)", 5, 6);
    cell_5_6->setProperty("Value2", "Java");
//    // step6: 保存文件
//    // 方式一：保存当前文件
      workbook->dynamicCall("Save()");  //保存文件
      //workbook->dynamicCall("Close(Boolean)", false);  //关闭文件
//    // 方式二：另存为
   //QString fileName = QFileDialog::getSaveFileName(NULL, QStringLiteral("保存文件"), QStringLiteral("excel名称"), QStringLiteral("EXCEL(*.xlsx)"));
    QString fileName=QStringLiteral("C:/Users/lixc/Desktop/excel名称.xlsx");
    workbook->dynamicCall("SaveAs(const QString&)", QDir::toNativeSeparators(fileName)); //保存到filepath
    // 注意一定要用QDir::toNativeSeparators, 将路径中的"/"转换为"\", 不然一定保存不了
    workbook->dynamicCall("Close (Boolean)", false);  //关闭文件

    delete excel;
*/
    QString filepath=QFileDialog::getSaveFileName(this,tr("Save orbit"),".",tr("Microsoft Office 2007 (*.xlsx)"));//获取保存路径
    if(!filepath.isEmpty()){
        QAxObject *excel = new QAxObject(this);
        excel->setControl("Excel.Application");//连接Excel控件
        excel->dynamicCall("SetVisible (bool Visible)","false");//不显示窗体
        excel->setProperty("DisplayAlerts", false);//不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示

        QAxObject *workbooks = excel->querySubObject("WorkBooks");//获取工作簿集合
        workbooks->dynamicCall("Add");//新建一个工作簿
        QAxObject *workbook = excel->querySubObject("ActiveWorkBook");//获取当前工作簿
        QAxObject *worksheets = workbook->querySubObject("Sheets");//获取工作表集合
        QAxObject *worksheet = worksheets->querySubObject("Item(int)",1);//获取工作表集合的工作表1，即sheet1
        QAxObject *cellX,*cellY;
        for(int i=0;i<m_tbModel->rowCount();i++){
            QString X="A"+QString::number(i+1);//设置要操作的单元格，如A1
            QString Y="B"+QString::number(i+1);
            cellX = worksheet->querySubObject("Range(QVariant, QVariant)",X);//获取单元格
            cellY = worksheet->querySubObject("Range(QVariant, QVariant)",Y);
            if (i > 2)
                break;
            cellX->dynamicCall("SetValue(const QVariant&)",QVariant(m_tbModel->item(i,0)->text()));//设置单元格的值
            cellY->dynamicCall("SetValue(const QVariant&)",QVariant(m_tbModel->item(i,1)->text()));
        }

        workbook->dynamicCall("SaveAs(const QString&)",QDir::toNativeSeparators(filepath));//保存至filepath，注意一定要用QDir::toNativeSeparators将路径中的"/"转换为"\"，不然一定保存不了。
        workbook->dynamicCall("Close()");//关闭工作簿
        excel->dynamicCall("Quit()");//关闭excel
        delete excel;
        excel=NULL;
    }
}

void Widget::on_pushButton_clicked()
{
//    QStandardItem *item = new QStandardItem;
//    QList<QStandardItem *> itemList;
//    static int j = 0;
//    qDebug() << "on_pushButton_clicked()" << j;;
//    for (int i = 0; i <5; i++)
//    {
//        itemList.append(new QStandardItem(QString::number(i+(j++))));
//    }
//    item->appendRow(itemList);
//    m_tbModel->appendRow(item);
////    m_tbModel->appendRow(itemList);
}
