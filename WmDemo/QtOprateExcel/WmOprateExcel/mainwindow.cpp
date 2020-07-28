#include "mainwindow.h"
#include "ui_mainwindow.h"

#include <QFileDialog>
#include <QDebug>

#include "qax_ms_execel_io.h"

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    init();
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::init()
{
    initTabWidgets();
    initFileOperationActions();
    initCheckedActions();
}

void MainWindow::initTabWidgets()
{
    ui->tableWidget_open->setRowCount(100);
    ui->tableWidget_open->setColumnCount(100);
    ui->tableWidget_save->setRowCount(10);
    ui->tableWidget_save->setColumnCount(10);
}

void MainWindow::initFileOperationActions()
{
    connect(ui->action_save, &QAction::triggered, [this](){
        QString excelName = QFileDialog::getSaveFileName(NULL, "解析Excel", "", "Excel(*.xls *.xlsx *.csv)");
        if (excelName.isEmpty())
            return;

        QVariant variant;
        QList<QList<QVariant>> listList;
        QTableWidget *tbw = ui->tableWidget_save;

        for (int i = 0; i < tbw->rowCount(); i++)
        {
            QList<QVariant> list;
            for (int j = 0; j < tbw->columnCount(); j++)
            {
                QString text;
                tbw->item(i, j) == NULL ? text = "" : text = tbw->item(i, j)->text();
                list.append(QVariant(text));
            }
            listList.append(list);
        }
        qDebug() << "writeExcel : " << listList;
//        variant.setValue<QList<QList<QVariant>>>(listList);
        MsExcelIO::castListListVariant2Variant(listList, variant);

        qDebug() << "writeExcel : " << variant;
        MsExcelIO::writeExcel(excelName, 1, variant);
    });

    connect(ui->action_open, &QAction::triggered, [this](){
        QString excelName = QFileDialog::getOpenFileName(NULL, "解析Excel", "", "Excel(*.xls *.xlsx *.csv)");
        if (excelName.isEmpty())
            return;

        QVariant variant = MsExcelIO::readExcel(excelName, 1);
        QList<QList<QVariant>> listList;
        MsExcelIO::castVariant2listListVariant(variant, listList);
        QTableWidget *tbw = ui->tableWidget_open;
        tbw->clear();
        for (int i = 0; i < listList.size(); i++)
        {
            for (int j = 0; j < listList.at(i).size(); j++)
            {
                qDebug() << listList[i][j].type() << listList[i][j].typeName() << listList[i][j];
                tbw->setItem(i, j, new QTableWidgetItem(listList[i][j].toString()));
            }
        }
    });
}

void MainWindow::initCheckedActions()
{
    m_actGroup = new QActionGroup(this);
    m_actGroup->addAction(ui->actionMicroSoftExcel);
    m_actGroup->addAction(ui->actionThirdExcel);
    m_actGroup->addAction(ui->actionWPSExcel);

    foreach(QAction *act, m_actGroup->actions())
    {
        act->setCheckable(true);
    }
    m_actGroup->actions().first()->setChecked(true);
}
