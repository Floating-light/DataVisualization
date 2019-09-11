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
#include <qcombobox.h>
#include <QHBoxLayout>
#include <QVBoxLayout>
#include <QLabel>
#include <QLineEdit>
#include <QListWidget>
#include "MyCombbox.h"

QT_BEGIN_NAMESPACE

struct ItemSelectCombox
{
	QLabel* label;
	MyCombbox* content;
	QListWidget * alist;
};

class Ui_DataVisualizationClass
{
public:
    QMenuBar *menuBar;
	QMenu* menuFile;
	QMenu* menuEdit;
	QMenu* menuExport;

    QToolBar *mainToolBar;
    QWidget *centralWidget;
    QStatusBar *statusBar;

	QAction* actionOpenFile;
	QAction* actionOpenChart;
	QAction* histogramChart;
	QAction* scatterChart;
	QAction* lineChart;
	QAction* calculate;
	QAction* saveFile;
	QAction* templateExport;
	QAction* savePic;

	std::vector<QAction*> exportActions;

	QHBoxLayout* topCombox;
	QHBoxLayout* viewLayout;

	QVBoxLayout* mainVLayout;

	QLineEdit* chartLineEdit;

	QList< ItemSelectCombox*> itemSelect;

    void setupUi(QMainWindow *DataVisualizationClass)
    {
        if (DataVisualizationClass->objectName().isEmpty())
            DataVisualizationClass->setObjectName(QString::fromUtf8("DataVisualizationClass"));
        DataVisualizationClass->resize(600, 400);

		actionOpenFile = new QAction(DataVisualizationClass);
		actionOpenFile->setObjectName(QString::fromUtf8("actionOpen"));
		actionOpenChart = new QAction(DataVisualizationClass);
		actionOpenChart->setObjectName(QString::fromUtf8("actionChart"));
		histogramChart = new QAction(DataVisualizationClass);
		scatterChart = new QAction(DataVisualizationClass);
		lineChart = new QAction(DataVisualizationClass);
		calculate = new QAction(DataVisualizationClass);
		saveFile = new QAction(DataVisualizationClass);
		templateExport = new QAction(DataVisualizationClass);
		savePic = new QAction(DataVisualizationClass);

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
		menuEdit->addAction(histogramChart);
		menuEdit->addAction(scatterChart);
		menuEdit->addAction(lineChart);

		menuBar->addAction(menuEdit->menuAction());

		menuExport = new QMenu(menuBar);
		menuExport->setObjectName(QString::fromUtf8("menuEdit"));
		menuExport->addAction(templateExport);
		menuExport->addAction(savePic);

		menuBar->addAction(menuExport->menuAction());

        mainToolBar = new QToolBar(DataVisualizationClass);
        mainToolBar->setObjectName(QString::fromUtf8("mainToolBar"));
        DataVisualizationClass->addToolBar(mainToolBar);

		mainToolBar->setToolButtonStyle(Qt::ToolButtonTextUnderIcon);
		mainToolBar->addAction(actionOpenFile);
		mainToolBar->addAction(histogramChart);
		mainToolBar->addAction(lineChart);
		mainToolBar->addAction(scatterChart);
		mainToolBar->addAction(calculate);
		mainToolBar->addAction(saveFile);
		mainToolBar->addAction(templateExport);
		mainToolBar->addAction(savePic);

		topCombox = new QHBoxLayout();
		topCombox->setAlignment(Qt::AlignLeft);

		viewLayout = new QHBoxLayout();

		chartLineEdit = new QLineEdit();
		chartLineEdit->setPlaceholderText(QStringLiteral("请输入图表所需要展示的表头"));
		chartLineEdit->setText(QStringLiteral("19年1月供货,19年2月供货;1月环比对比,2月环比对比;环比对比图"));

		mainVLayout = new QVBoxLayout();
		mainVLayout->addWidget(chartLineEdit);
		mainVLayout->addLayout(topCombox);
		mainVLayout->addLayout(viewLayout);

        centralWidget = new QWidget(DataVisualizationClass);
        centralWidget->setObjectName(QString::fromUtf8("centralWidget"));
        DataVisualizationClass->setCentralWidget(centralWidget);
		centralWidget->setLayout(mainVLayout);
		
		

		//ItemSelectCombox* current = new ItemSelectCombox();
		//itemSelect.push_back(current);
		//current->label = new QLabel(QStringLiteral("营销操盘方:"));
		//current->content = new QComboBox();
		//current->content->addItems(QStringList{ QStringLiteral("中南") ,QStringLiteral("新城"),
		//	QStringLiteral("碧桂园") });
		////test->setSizeAdjustPolicy();
		//topCombox->addWidget(current->label);
		//topCombox->addWidget(current->content);

        statusBar = new QStatusBar(DataVisualizationClass);
        statusBar->setObjectName(QString::fromUtf8("statusBar"));
        DataVisualizationClass->setStatusBar(statusBar);

        retranslateUi(DataVisualizationClass);

        QMetaObject::connectSlotsByName(DataVisualizationClass);
    } // setupUi

    void retranslateUi(QMainWindow *DataVisualizationClass)
    {
		DataVisualizationClass->setWindowIcon(QIcon("./Resources/data.png"));
        DataVisualizationClass->setWindowTitle(QApplication::translate("DataVisualizationClass", "DataVisualization", nullptr));
		actionOpenFile->setText(QString::fromLocal8Bit(std::string("打开").data()));
		actionOpenFile->setIcon(QIcon("./Resources/file.png"));

		actionOpenChart->setText(QString::fromLocal8Bit(std::string("图表").data()));
		actionOpenChart->setIcon(QIcon("./Resources/all.png"));

		histogramChart->setText(QStringLiteral("柱状图"));
		histogramChart->setIcon(QIcon("./Resources/histogram.png"));
		

		scatterChart->setText(QStringLiteral("散点图"));
		scatterChart->setIcon(QIcon("./Resources/scatter.png"));

		lineChart->setText(QStringLiteral("折线图"));
		lineChart->setIcon(QIcon("./Resources/line.png"));

		calculate->setText(QStringLiteral("计算"));
		calculate->setIcon(QIcon("./Resources/caclute.png"));//saveFile

		saveFile->setText(QStringLiteral("保存"));
		saveFile->setIcon(QIcon("./Resources/save.png"));//saveFiletemplateExport

		templateExport->setText(QStringLiteral("模板导出"));
		templateExport->setIcon(QIcon("./Resources/export.png"));//saveFile

		savePic->setText(QStringLiteral("保存图片"));
		savePic->setIcon(QIcon("./Resources/download.png"));//savePic

		menuFile->setTitle(QString::fromLocal8Bit(std::string("文件").data()));
		menuEdit->setTitle(QString::fromLocal8Bit(std::string("视图").data()));
		menuExport->setTitle(QString::fromLocal8Bit(std::string("导出").data()));
	} // retranslateUi

};

namespace Ui {
    class DataVisualizationClass: public Ui_DataVisualizationClass {};
} // namespace Ui

QT_END_NAMESPACE

#endif // UI_DATAVISUALIZATION_H
