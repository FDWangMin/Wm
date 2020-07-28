#include "mainwindow.h"
#include <QApplication>
#include <QtGui>
#include <QAxObject>
#include <QFileDialog>
#include <QObject>
#include<QDebug>

#include "qax_ms_execel_io.h"

void test()
{
    // step1：连接控件
    QAxObject* excel = new QAxObject();
    excel->setControl("Excel.Application");  // 连接Excel控件
    excel->dynamicCall("SetVisible (bool Visible)", "false"); // 不显示窗体
    excel->setProperty("DisplayAlerts", false);  // 不显示任何警告信息。如果为true, 那么关闭时会出现类似"文件已修改，是否保存"的提示

    // step2: 打开工作簿
    QAxObject* workbooks = excel->querySubObject("WorkBooks"); // 获取工作簿集合
    // 打开工作簿方式一：新建
//    workbooks->dynamicCall("Add"); // 新建一个工作簿

//    QAxObject* workbook = excel->querySubObject("ActiveWorkBook"); // 获取当前工作簿
    // 打开工作簿方式二：打开现成
    QAxObject* workbook = workbooks->querySubObject("Open(const QString&)", ("C:/Users/lixc/Desktop/tt2.xlsx")); // 从控件lineEdit获取文件名

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
}
//void test2()
//{
//       QAxObject excel("Excel.Application");//连接Excel控件
//       excel.setProperty("Visible", false);// 不显示窗体
//       excel->setProperty("DisplayAlerts", false);  // 不显示任何警告信息。如果为true, 那么关闭时会出现类似"文件已修改，是否保存"的提示
//       QAxObject *work_books = excel.querySubObject("WorkBooks");// 获取工作簿集合



//       work_books->dynamicCall("Open(const QString&)", "d:\\XSpeed产品测试管理表.xls");
//       excel.setProperty("Caption", "Qt Excel");
//       QAxObject *work_book = excel.querySubObject("ActiveWorkBook");
//       QAxObject *work_sheets = work_book->querySubObject("Sheets");  //Sheets也可换用WorkSheets
//       QAxObject *work_sheet = work_sheets->querySubObject("Item(1)",3);
//       //操作单元格（第2行第2列）
//       QAxObject *cell = work_sheet->querySubObject("Cells(int,int)", 2, 2);
//       cell->setProperty("Value", "Java C++ C# PHP Perl Python Delphi Ruby");  //设置单元格值
//       cell->setProperty("RowHeight", 50);  //设置单元格行高
//       cell->setProperty("ColumnWidth", 30);  //设置单元格列宽
//       cell->setProperty("HorizontalAlignment", -4108); //左对齐（xlLeft）：-4131  居中（xlCenter）：-4108  右对齐（xlRight）：-4152
//       cell->setProperty("VerticalAlignment", -4108);  //上对齐（xlTop）-4160 居中（xlCenter）：-4108  下对齐（xlBottom）：-4107
//       cell->setProperty("WrapText", true);  //内容过多，自动换行
//       cell->dynamicCall("ClearContents()");  //清空单元格内容
//       QAxObject* interior = cell->querySubObject("Interior");
//       interior->setProperty("Color", QColor(0, 255, 0));   //设置单元格背景色（绿色）
//       QAxObject* border = cell->querySubObject("Borders");
//       border->setProperty("Color", QColor(0, 0, 255));   //设置单元格边框色（蓝色）
//       QAxObject *font = cell->querySubObject("Font");  //获取单元格字体
//       font->setProperty("Name", QStringLiteral("华文彩云"));  //设置单元格字体
//       font->setProperty("Bold", true);  //设置单元格字体加粗
//       font->setProperty("Size", 20);  //设置单元格字体大小
//       font->setProperty("Italic", true);  //设置单元格字体斜体
//       font->setProperty("Underline", 2);  //设置单元格下划线
//       font->setProperty("Color", QColor(255, 0, 0));  //设置单元格字体颜色（红色）
//       //设置单元格内容，并合并单元格（第5行第3列-第8行第5列）
//       QAxObject *cell_5_6 = work_sheet->querySubObject("Cells(int,int)", 2, 2);
//       cell_5_6->setProperty("Value2", "Java");  //设置单元格值
//       QAxObject *cell_8_5 = work_sheet->querySubObject("Cells(int,int)", 8, 5);
//       cell_8_5->setProperty("Value2", "C++dd");
//       QString merge_cell;
//       merge_cell.append(QChar(3 - 1 + 'A'));  //初始列
//       merge_cell.append(QString::number(5));  //初始行
//       merge_cell.append(":");
//       merge_cell.append(QChar(5 - 1 + 'A'));  //终止列
//       merge_cell.append(QString::number(8));  //终止行
//       QAxObject *merge_range = work_sheet->querySubObject("Range(const QString&)", merge_cell);
//       merge_range->setProperty("HorizontalAlignment", -4108);
//       merge_range->setProperty("VerticalAlignment", -4108);
//       merge_range->setProperty("WrapText", true);
//       merge_range->setProperty("MergeCells", true);  //合并单元格
//       //merge_range->setProperty("MergeCells", false);  //拆分单元格
//       work_book->dynamicCall("SaveAs(const QString&)", "E:\\test2.xlsx");  //另存为另一个文件
//       //work_book->dynamicCall("Save()");  //保存文件（为了对比test与下面的test2文件，这里不做保存操作） work_book->dynamicCall("SaveAs(const QString&)", "E:\\test2.xlsx");  //另存为另一个文件
//       work_book->dynamicCall("Close(Boolean)", false);  //关闭文件
//       excel.dynamicCall("Quit(void)");  //退出
//       QAxObject *excel = NULL;
//       QAxObject *workbooks = NULL;
//       QAxObject *workbook = NULL;
//       excel = new QAxObject("Excel.Application");
//       if (!excel)
//       {
//           QMessageBox::critical(this, "错误信息", "EXCEL对象丢失");
//           return;
//       }
//       excel->dynamicCall("SetVisible(bool)", false);
//       workbooks = excel->querySubObject("WorkBooks");
//       workbook = workbooks->querySubObject("Open(QString, QVariant)", QString(tr("d:\\XSpeed产品测试管理表.xls")));
//       QAxObject * worksheet = workbook->querySubObject("WorkSheets(int)", 1);//打开第一个sheet
//       //QAxObject * worksheet = workbook->querySubObject("WorkSheets");//获取sheets的集合指针
//       //int intCount = worksheet->property("Count").toInt();//获取sheets的数量
//       QAxObject *cell = worksheet->querySubObject("Cells(int,int)",6,3);
//       cell->setProperty("Value2","测试管理");
//       workbook->dynamicCall("Save()");
//       QAxObject *cell2 = worksheet->querySubObject("Cells(int,int)",7,3);
//       cell2->setProperty("Value2","方法");
//       workbook->dynamicCall("Save()");

//       QAxObject * usedrange = worksheet->querySubObject("UsedRange");//获取该sheet的使用范围对象
//       QAxObject * rows = usedrange->querySubObject("Rows");
//       QAxObject * columns = usedrange->querySubObject("Columns");
//       //获取行数和列数
//       int intRowStart = usedrange->property("Row").toInt();
//       int intColStart = usedrange->property("Column").toInt();
//       int intCols = columns->property("Count").toInt();
//       int intRows = rows->property("Count").toInt();
//       //获取excel内容
//       for (int i = intRowStart; i < intRowStart + intRows; i++)  //行
//       {
//           for (int j = intColStart; j < intColStart + intCols; j++)  //列
//           {
//               QAxObject * cell = worksheet->querySubObject("Cells(int,int)", i, j );  //获取单元格
//               // qDebug() << i << j << cell->property("Value");         //*****************************出问题!!!!!!
//               qDebug() << i << j <<cell->dynamicCall("Value2()").toString(); //正确
//           }
//       }
//       workbook->dynamicCall("Close (Boolean)", false);
//       //同样，设置值，也用dynamimcCall("SetValue(const QVariant&)", QVariant(QString("Help!")))这样才成功的。。
//        excel->dynamicCall("Quit(void)");
//       //excel->dynamicCall("Quit (void)");
//       delete excel;//一定要记得删除，要不线程中会一直打开excel.exe
//}


//int main(int argc, char *argv[])
//{

//    QApplication a(argc, argv);
//    QString strMessage1 = QString::fromLocal8Bit("我是UTF8编码的文件：");
//    QString strMessage2 = QStringLiteral("我是UTF8编码的文件：");
//    QString strMessage3 = QString::fromWCharArray(L"我是UTF8编码的文件：");
//    qDebug() << strMessage1;
//    qDebug() << strMessage2;
//    qDebug() << strMessage3;
//    test();
//    return 0;
//    //MainWindow w;
//    //w.show();
////    QTextCodec::setCodecForTr(QTextCodec::codecForName("GBK"));
////    QTextCodec::setCodecForLocale(QTextCodec::codecForName("GBK"));
////    QTextCodec::setCodecForCStrings(QTextCodec::codecForName("GBK"));
//    //return a.exec();
//}

QScopedPointer<QAxObject> MsExcelIO::m_excel;

//更多参考：
//https://github.com/qtcn/tianchi/blob/v0.0.2-build20130701/include/tianchi/file/tcmsexcel.h
//
QVariant MsExcelIO::readExcel(QString fileName, int sheetID)
{
    QVariant var;
    QElapsedTimer timer;
    timer.restart();

    //m_excel
    if (m_excel.isNull() || m_excel->isNull())
    {
        qDebug() << "--- init m_excel---";
        m_excel.reset(new QAxObject("Excel.Application")); //加载Excel驱动
        qDebug()<<"Excel.Application time:"<<timer.elapsed()<<"ms";timer.restart();
        m_excel->dynamicCall("SetVisible (bool Visible)", "false"); //不显示窗体
        qDebug()<<"SetVisible time:"<<timer.elapsed()<<"ms";timer.restart();

        m_excel->setProperty("Visible", false); //不显示Excel界面，如果为true会看到启动的Excel界面
        qDebug()<<"Visible time:"<<timer.elapsed()<<"ms";timer.restart();
        m_excel->setProperty("DisplayAlerts", false);// 不显示任何警告信息。如果为true, 那么关闭时会出现类似"文件已修改，是否保存"的提示
        qDebug()<<"DisplayAlerts time:"<<timer.elapsed()<<"ms";timer.restart();
    }

    QAxObject* pWorkBooks = m_excel->querySubObject("WorkBooks");
    qDebug()<<"WorkBooks time:"<<timer.elapsed()<<"ms";timer.restart();

    pWorkBooks->dynamicCall("Open (const QString&)", fileName);//打开指定文件
    qDebug()<<"Open time:"<<timer.elapsed()<<"ms";timer.restart();

    QAxObject* pWorkBook = m_excel->querySubObject("ActiveWorkBook");
    qDebug()<<"ActiveWorkBook time:"<<timer.elapsed()<<"ms";timer.restart();

    QAxObject* pWorkSheets = pWorkBook->querySubObject("Sheets");//获取工作表
    qDebug()<<"Sheets time:"<<timer.elapsed()<<"ms";timer.restart();

    int nSheetCount = pWorkSheets->property("Count").toInt();  //获取工作表的数目
    qDebug()<<"Count time:"<<timer.elapsed()<<"ms";timer.restart();
    if(nSheetCount > 0)
    {
        sheetID > 0 ? sheetID : sheetID=1;
        QAxObject* pWorkSheet = pWorkBook->querySubObject("Sheets(int)", sheetID);//获取对应ID表

        if (pWorkSheet != NULL && ! pWorkSheet->isNull())
        {
            QAxObject *usedRange = pWorkSheet->querySubObject("UsedRange");
            if(NULL == usedRange || usedRange->isNull())
            {
                return var;
            }
            var = usedRange->dynamicCall("Value");
            usedRange->deleteLater();
        }
    }
    qDebug()<<"read time:"<<timer.elapsed()<<"ms";timer.restart();

    pWorkBooks->dynamicCall("Close()");
    qDebug()<<"Close time:"<<timer.elapsed()<<"ms";timer.restart();

    return var;
}

bool MsExcelIO::writeExcel(QString fileName, int sheetID, QVariant var)
{
    QList<QList<QVariant>> listListVar;
    qDebug() << "MsExcelIO::writeExcel" << var;
    castVariant2listListVariant(var, listListVar);
    if(listListVar.size() <= 0)
        return false;

    QElapsedTimer timer;
    timer.restart();
    //m_excel
    if (m_excel.isNull() || m_excel->isNull())
    {
        qDebug() << "--- init m_excel---";
        m_excel.reset(new QAxObject("Excel.Application")); //加载Excel驱动
        qDebug()<<"Excel.Application time:"<<timer.elapsed()<<"ms";timer.restart();
        m_excel->dynamicCall("SetVisible (bool Visible)", "false"); //不显示窗体
        qDebug()<<"SetVisible time:"<<timer.elapsed()<<"ms";timer.restart();

        m_excel->setProperty("Visible", false); //不显示Excel界面，如果为true会看到启动的Excel界面
        qDebug()<<"Visible time:"<<timer.elapsed()<<"ms";timer.restart();
        m_excel->setProperty("DisplayAlerts", false);//不显示任何警告信息。如果为true, 那么关闭时会出现类似"文件已修改，是否保存"的提示
        qDebug()<<"DisplayAlerts time:"<<timer.elapsed()<<"ms";timer.restart();
    }

    QAxObject* pWorkBooks = m_excel->querySubObject("WorkBooks");
    qDebug()<<"WorkBooks time:"<<timer.elapsed()<<"ms";timer.restart();

    pWorkBooks->dynamicCall("Add");
    QAxObject* pWorkBook   = m_excel->querySubObject("ActiveWorkBook");
    qDebug()<<"ActiveWorkBook time:"<<timer.elapsed()<<"ms";timer.restart();
    QAxObject* pWorkSheet = pWorkBook->querySubObject("Sheets(int)", sheetID);
    qDebug()<<"Sheets time:"<<timer.elapsed()<<"ms";timer.restart();
//    QAxObject* pWorkSheet = pWorkBook->querySubObject("WorkSheets");

    int row = listListVar.size();
    int col = listListVar.at(0).size();
    QString rangStr;
    convertToColName(col,rangStr);
    rangStr += QString::number(row);
    rangStr = "A1:" + rangStr;
    qDebug() << "rangStr:" << rangStr;

    QAxObject *range = pWorkSheet->querySubObject("Range(const QString&)",rangStr);

//    castListListVariant2Variant(listListVar,var);
    bool succ = false;
    succ = range->setProperty("Value", var);
    range->deleteLater();

    QString strPath = fileName.replace('/','\\');
    qDebug()<<strPath;
    pWorkBook->dynamicCall("SaveAs(const QString&,int,const QString&,const QString&,bool,bool)", strPath
                         ,56,QString(""),QString(""),false,false);

    pWorkBooks->dynamicCall("Close");

    return succ;
}

void MsExcelIO::castVariant2listListVariant(const QVariant &var, QList<QList<QVariant> > &res)
{
    QVariantList varRows = var.toList();
    if(varRows.isEmpty())
    {
        return;
    }
    const int rowCount = varRows.size();
    QVariantList rowData;
    for(int i=0;i<rowCount;++i)
    {
        rowData = varRows[i].toList();
        res.append(rowData);
    }
}

void MsExcelIO::castListListVariant2Variant(const QList<QList<QVariant> > &cells, QVariant &res)
{
    QVariantList vars;
    const int rows = cells.size();
    for(int i=0;i<rows;++i)
    {
        vars.append(QVariant(cells[i]));
    }
    res = QVariant(vars);
}

void MsExcelIO::convertToColName(int data, QString &res)
{
    Q_ASSERT(data>0 && data<65535);
    int tempData = data / 26;
    if(tempData > 0)
    {
        int mode = data % 26;
        convertToColName(mode,res);
        convertToColName(tempData,res);
    }
    else
    {
        res=(to26AlphabetString(data)+res);
    }
}

QString MsExcelIO::to26AlphabetString(int data)
{
    QChar ch = data + 0x40;//A对应0x41
    return QString(ch);
}

