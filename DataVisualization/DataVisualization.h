#pragma once
#include <vector>
#include <QtWidgets/QMainWindow>
#include "ui_DataVisual.h"
#include <ActiveQt/qaxobject.h>
#include <QTableWidget>
#include <QHBoxLayout>
#include <QtCharts/QChartGlobal>

#include <QtCharts/QChartView>
#include <QtCharts/QPieSeries>
#include <QtCharts/QPieSlice>
#include <QtCharts/QAbstractBarSeries>
#include <QtCharts/QPercentBarSeries>
#include <QtCharts/QStackedBarSeries>
#include <QtCharts/QBarSeries>
#include <QtCharts/QBarSet>
#include <QtCharts/QLineSeries>
#include <QtCharts/QSplineSeries>
#include <QtCharts/QScatterSeries>
#include <QtCharts/QAreaSeries>
#include <QtCharts/QLegend>
#include <QtWidgets/QGridLayout>
#include <QtWidgets/QFormLayout>
#include <QtWidgets/QComboBox>
#include <QtWidgets/QSpinBox>
#include <QtWidgets/QCheckBox>
#include <QtWidgets/QGroupBox>
#include <QtWidgets/QLabel>
#include <QtCore/QRandomGenerator>
#include <QtCharts/QBarCategoryAxis>
#include <QtWidgets/QApplication>
#include <QtCharts/QValueAxis>
#include <QHeaderView>
QT_BEGIN_NAMESPACE
class QComboBox;
class QCheckBox;
class Ui_ThemeWidgetForm;
QT_END_NAMESPACE

QT_CHARTS_BEGIN_NAMESPACE
class QChartView;
class QChart;
QT_CHARTS_END_NAMESPACE

typedef QPair<QPointF, QString> Data;
typedef QList<Data> DataList;
typedef QList<DataList> DataTable;

QT_CHARTS_USE_NAMESPACE

typedef QPair<QPointF, QString> Data;
typedef QList<Data> DataList;
typedef QList<DataList> DataTable;

constexpr auto excelFilePath = "C:/fileds.xlsx";

class DataVisualization : public QMainWindow
{
	Q_OBJECT

public:
	DataVisualization(QWidget *parent = Q_NULLPTR);

	void displayData(const QList<QList<QVariant>>& data,const  std::vector<QString>& headers);
	void displayData(const std::vector<std::vector<QVariant>>& data,int beginRow ,int headerRow);

private:
	Ui::DataVisualizationClass ui;
	DataTable m_dataTable;
	QChartView*chartView ;
	QTableWidget* tableWidget;
	QHBoxLayout* hLayout;
	void DataVisualization::Read_Excel(const QString PATH, const QString FILENAME, 
		const int SHEETNUM, const QString RANGE, const int INVALIDROW,
		const int TOTALCOLNUM, std::vector<QString>& RESULT);

	void readAll(const QString Path);

	void import();

	QChart* createLineChart() const;

	DataTable generateRandomData(int listCount, int valueMax, int valueCount)const;
	void addData(int column);

	//excel host
	QAxObject* excel;

	//get a work book bind with a excel file.
	QAxObject* getWorkBooks(const QString& );

	//get number i sheet
	QAxObject* getSheet(QAxObject* workBook, int number);

	//do some test
	void standardTest();
private slots:
	void headerClicked(int);

};
