#pragma once
#include <vector>
#include <algorithm>
#include <limits>
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

const std::vector<QString> selectHeaderName{QStringLiteral("事业部\n（住开/商开）"),
    QStringLiteral("大区\n（南/中/北）"),QStringLiteral("业态"),
	QStringLiteral("城市环线") };

class DataVisualization : public QMainWindow
{
	Q_OBJECT

public:
	DataVisualization(QWidget *parent = Q_NULLPTR);

	std::vector<bool> rowChecked;

	void displayData(const QList<QList<QVariant>>& data,const  std::vector<QString>& headers);
	void displayData(const std::vector<std::vector<QVariant>>& data,int beginRow ,int headerRow);
	void displaySelectRow(const std::vector<int>& rowsNumber);
	std::vector<int>  getSelectRowNumber(const std::vector<QString>& headType, const std::vector<QString>& traget);
	
	//std::vector<int>  selectRowNumber(const std::vector<QString>& headType, const std::vector<QString>& traget);
	//get unique item QString vector under header name
	void uniqueItem(const QString& headerName, std::vector<QString>& items);

	void addSelectCombox(const QString& headerName);
	ItemSelectCombox* createSelectCombox(const QString& headerName);
	
	void filterItem();

	void displayScatterChart();
	void displayLineChart();
	void displayBarChart();
public slots:
	void buttonPress();
	void checkBoxchange(int state);
	void comboxChanged(const QString& text);
private:

	int count = 0;
	Ui::DataVisualizationClass ui;
	DataTable m_dataTable;
	QChartView*chartView ;
	QTableWidget* tableWidget;
	QHBoxLayout* hLayout;

	std::map<QString, std::vector<double>> scatterData;

	QChart* createLineChart() const;
	QChart* createScatterChart();
	QChart* createScatterChartTwo();
	QChart* createBarChart();
	DataTable generateRandomData(int listCount, int valueMax, int valueCount)const;
	void addData(int column);
	void addSelectedRowColumData(int column);
	void updataChartData();
	//return -1 if not find, begin with 0
	int headerString2ColumnNumber(const QString& headerName);
private slots:
	void headerClicked(int);

};
