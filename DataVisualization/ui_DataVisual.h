/********************************************************************************
** 
********************************************************************************/

#ifndef UI_DATAVISUALIZATION_H
#define UI_DATAVISUALIZATION_H

#include <QtCore/QVariant>
#include <QtWidgets/QApplication>
#include <QtWidgets/QMainWindow>
#include <QtWidgets/QMenuBar>
#include <QtWidgets/QStatusBar>
#include <QtWidgets/QToolBar>
#include <QtWidgets/QToolButton>
#include <QtWidgets/QWidget>

QT_BEGIN_NAMESPACE

class Ui_DataVisualizationClass
{
public:
    QMenuBar *menuBar;
	QMenu* menuFile;
	QMenu* menuEdit;

    QToolBar *mainToolBar;
    QWidget *centralWidget;
    QStatusBar *statusBar;
	QToolButton* pActionOpenBar;
	QAction* actionOpenFile;
	QAction* actionOpenChart;

    void setupUi(QMainWindow *DataVisualizationClass)
    {
        if (DataVisualizationClass->objectName().isEmpty())
            DataVisualizationClass->setObjectName(QString::fromUtf8("DataVisualizationClass"));
        DataVisualizationClass->resize(600, 400);

		actionOpenFile = new QAction(DataVisualizationClass);
		actionOpenFile->setObjectName(QString::fromUtf8("actionOpen"));
		actionOpenChart = new QAction(DataVisualizationClass);
		actionOpenChart->setObjectName(QString::fromUtf8("actionChart"));

        menuBar = new QMenuBar(DataVisualizationClass);
        menuBar->setObjectName(QString::fromUtf8("menuBar"));
        DataVisualizationClass->setMenuBar(menuBar);

		menuFile = new QMenu(menuBar);
		menuFile->setObjectName(QString::fromUtf8("menuFile"));
		menuFile->addAction(actionOpenFile);
		menuBar->addAction(menuFile->menuAction());

		menuEdit = new QMenu(menuBar);
		menuEdit->setObjectName(QString::fromUtf8("menuEdit"));
		menuEdit->addAction(actionOpenChart);
		menuBar->addAction(menuEdit->menuAction());


        mainToolBar = new QToolBar(DataVisualizationClass);
        mainToolBar->setObjectName(QString::fromUtf8("mainToolBar"));
        DataVisualizationClass->addToolBar(mainToolBar);

		pActionOpenBar = new QToolButton(mainToolBar);
		pActionOpenBar->setIcon(QIcon("./Resources/open.png"));
		pActionOpenBar->setToolButtonStyle(Qt::ToolButtonTextUnderIcon);
		pActionOpenBar->setText(QStringLiteral("打开"));
		mainToolBar->addWidget(pActionOpenBar);

        centralWidget = new QWidget(DataVisualizationClass);
        centralWidget->setObjectName(QString::fromUtf8("centralWidget"));
        DataVisualizationClass->setCentralWidget(centralWidget);

        statusBar = new QStatusBar(DataVisualizationClass);
        statusBar->setObjectName(QString::fromUtf8("statusBar"));
        DataVisualizationClass->setStatusBar(statusBar);

        retranslateUi(DataVisualizationClass);

        QMetaObject::connectSlotsByName(DataVisualizationClass);
    } // setupUi

    void retranslateUi(QMainWindow *DataVisualizationClass)
    {
        DataVisualizationClass->setWindowTitle(QApplication::translate("DataVisualizationClass", "DataVisualization", nullptr));
		actionOpenFile->setText(QString::fromLocal8Bit(std::string("打开").data()));
		actionOpenChart->setText(QString::fromLocal8Bit(std::string("图表").data()));
		
		menuFile->setTitle(QString::fromLocal8Bit(std::string("文件").data()));
		menuEdit->setTitle(QString::fromLocal8Bit(std::string("视图").data()));
	} // retranslateUi

};

namespace Ui {
    class DataVisualizationClass: public Ui_DataVisualizationClass {};
} // namespace Ui

QT_END_NAMESPACE

#endif // UI_DATAVISUALIZATION_H
