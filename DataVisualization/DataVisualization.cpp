#include "DataVisualization.h"

DataVisualization::DataVisualization(QWidget *parent)
	: QMainWindow(parent)
{
	ui.setupUi(this);

	tableWidget = new QTableWidget(this);


	ui.viewLayout->addWidget(tableWidget);
	/*hLayout = new QHBoxLayout();
	hLayout->addWidget(tableWidget);
	ui.centralWidget->setLayout(hLayout);*/

	//chart view 
	chartView = new QChartView(ui.centralWidget);
	//chartView->setChart();
	//ui.boxLayout->setDirection(QBoxLayout::LeftToRight);
	ui.viewLayout->addWidget(chartView);
	//hLayout->addWidget(chartView);

	//connect(newAct, &QAction::triggered, this, &MainWindow::newFile);

	//select event
	connect(tableWidget->horizontalHeader(), SIGNAL(sectionClicked(int)),this, SLOT(headerClicked(int)));

	connect(ui.actionOpenFile, &QAction::triggered, this, &DataVisualization::buttonPress);
	//connect(ui.actionOpenChart, &QAction::triggered, this, &DataVisualization::buttonPress);
	connect(ui.histogramChart, &QAction::triggered, this, &DataVisualization::displayBarChart);
	connect(ui.scatterChart, &QAction::triggered, this, &DataVisualization::displayScatterChart);
	connect(ui.lineChart, &QAction::triggered, this, &DataVisualization::displayLineChart);
}

QChart* DataVisualization::createLineChart() const
{
	int maxHoriz = INT_MIN;
	int minHoriz = INT_MAX;
	int maxVert = INT_MIN;
	int minVert = INT_MAX;
	//![1]
	QChart* chart = new QChart();
	chart->setTitle("Line chart");
	//![1]

	//![2]
	QString name("Series ");
	int nameIndex = 0;
	for (const DataList& list : m_dataTable) {
		QLineSeries* series = new QLineSeries(chart);
		for (const Data& data : list)
		{
			QPointF point = data.first;
			maxHoriz < point.x() ? maxHoriz = point.x() : (minHoriz > point.x() ? minHoriz= point.x() : true) ;
			maxVert < point.y() ? maxVert = point.y() : (minVert > point.y() ? minVert = point.y() : true) ;
			series->append(point);
		}
			
		series->setName(name + QString::number(nameIndex));
		nameIndex++;
		chart->addSeries(series);
	}
	//![2]

	//![3]
	chart->createDefaultAxes();
	chart->axes(Qt::Horizontal).first()->setRange(minHoriz-5, maxHoriz+10);
	chart->axes(Qt::Vertical).first()->setRange(minVert -5 , maxVert +10);
	//![3]
	//![4]
	// Add space to label to add space between labels and axis
	QValueAxis* axisY = qobject_cast<QValueAxis*>(chart->axes(Qt::Vertical).first());
	Q_ASSERT(axisY);
	axisY->setLabelFormat("%.1f  ");
	//![4]

	return chart;
}

QChart* DataVisualization::createScatterChart()
{
	int maxHoriz = INT_MIN;
	int minHoriz = INT_MAX;
	int maxVert = INT_MIN;
	int minVert = INT_MAX;
	// scatter chart  scatterData
	QChart* chart = new QChart();
	chart->setTitle("Scatter chart");
	QString name("Series ");
	int nameIndex = 0;
	auto iter1 = scatterData.cbegin();
	auto iter2 = ++scatterData.cbegin();
	std::vector<double> vec1 = iter1->second;
	std::vector<double> vec2 = iter2->second;
	QScatterSeries* series = new QScatterSeries(chart);
	for (int i = 0; i < vec1.size(); ++i)
	{
		series->append(QPointF(vec1[i], vec2[i]));
		maxHoriz < vec1[i] ? (maxHoriz = vec1[i]) : (minHoriz > vec1[i] ? minHoriz = vec1[i] : true);
		maxVert < vec2[i] ?( maxVert = vec2[i]) : (minVert > vec2[i] ? minVert = vec2[i] : true);
	}
	series->setName(iter2->first + ":" + iter1->first);
	chart->addSeries(series);

	chart->createDefaultAxes();
	chart->axes(Qt::Horizontal).first()->setRange(minHoriz - 1, maxHoriz + 1);
	chart->axes(Qt::Vertical).first()->setRange(minVert - 1, maxVert + 1);
	// Add space to label to add space between labels and axis
	QValueAxis * axisY = qobject_cast<QValueAxis*>(chart->axes(Qt::Vertical).first());
	Q_ASSERT(axisY);
	axisY->setLabelFormat("%.1f  ");

	QValueAxis* axisX = qobject_cast<QValueAxis*>(chart->axes(Qt::Horizontal).first());
	Q_ASSERT(axisX);
	axisX->setLabelFormat("%.1f  ");
	return chart;
}

QChart* DataVisualization::createScatterChartTwo()
{
	updataChartData();
	Q_ASSERT(m_dataTable.size() > 1);
	// scatter chart
	QChart* chart = new QChart();
	chart->setTitle(QStringLiteral("ɢ��ͼ"));

	DataList dataList_x = m_dataTable[0];
	DataList dataList_y = m_dataTable[1];

	QScatterSeries* series = new QScatterSeries(chart);
	for (int i = 0; i < dataList_x.size(); ++i)
	{
		dataList_x.at(i).first.y();
		series->append(QPointF(dataList_x.at(i).first.y(), dataList_y.at(i).first.y()));
	}
	series->setName(QStringLiteral("ɢ��"));
	chart->addSeries(series);

	chart->createDefaultAxes();
	chart->axes(Qt::Horizontal).first()->setRange(-100, 100);
	chart->axes(Qt::Vertical).first()->setRange(-100, 100);
	// Add space to label to add space between labels and axis
	QValueAxis * axisY = qobject_cast<QValueAxis*>(chart->axes(Qt::Vertical).first());
	Q_ASSERT(axisY);
	axisY->setLabelFormat("%.1f  ");
	return chart;
}

QChart* DataVisualization::createBarChart()
{
	QChart* chart = new QChart();
	chart->setTitle("Bar chart");
	int valueMax = std::numeric_limits<int>::min();
	int valueMin = std::numeric_limits<int>::max();
	QStackedBarSeries* series = new QStackedBarSeries(chart);
	for (int i(0); i < m_dataTable.count(); i++) {
		QBarSet* set = new QBarSet("Bar set " + QString::number(i));
		for (const Data& data : m_dataTable[i])
		{
			*set << data.first.y();
			data.first.y() > valueMax ? (valueMax = data.first.y()) : true;
			data.first.y() < valueMin ? (valueMin = data.first.y()) : true;
		}
		series->append(set);
	}
	chart->addSeries(series);

	chart->createDefaultAxes();
	chart->axes(Qt::Vertical).first()->setRange(valueMin, valueMax*2+10);
	// Add space to label to add space between labels and axis
	QValueAxis * axisY = qobject_cast<QValueAxis*>(chart->axes(Qt::Vertical).first());
	Q_ASSERT(axisY);
	axisY->setLabelFormat("%.1f  ");

	return chart;
}

void DataVisualization::displayLineChart()
{
	m_dataTable.clear();
	updataChartData();
	QChart* currentChart = createLineChart();
	QChart* previous = chartView->chart();
	chartView->setChart(currentChart);
	if (previous)
		delete previous;
}

void DataVisualization::displayScatterChart()
{
	m_dataTable.clear();
	updataChartData();
	QChart * currentChart = createScatterChartTwo();
	QChart* previous = chartView->chart();
	chartView->setChart(currentChart);
	if (previous)
		delete previous;
}

void DataVisualization::displayBarChart()
{
	m_dataTable.clear();
	updataChartData();
	QChart* currentChart = createBarChart();
	QChart* previous = chartView->chart();
	chartView->setChart(currentChart);
	if (previous)
		delete previous;
}

DataTable DataVisualization::generateRandomData(int listCount, int valueMax, int valueCount) const
{
	DataTable dataTable;
	// generate random data
	for (int i(0); i < listCount; i++) {
		DataList dataList;
		qreal yValue(0);
		for (int j(0); j < valueCount; j++) {
			yValue = yValue + QRandomGenerator::global()->bounded(valueMax / (qreal)valueCount);
			QPointF value((j + QRandomGenerator::global()->generateDouble()) * ((qreal)valueMax / (qreal)valueCount),
				yValue);
			QString label = "Slice " + QString::number(i) + ":" + QString::number(j);
			dataList << Data(value, label);
		}
		dataTable << dataList;
	}

	return dataTable;
}

void DataVisualization::addData(int column)
{
	DataList dataList;

	for (int i = 0; i < tableWidget->rowCount(); i++)
	{
		auto item = tableWidget->item(i, column);
		if (item == nullptr)
		{
			continue;
		}
		dataList << Data(QPointF(i, item->text().toDouble()),
			QString::number(column) + ":" + QString::number(i));
	}
	m_dataTable << dataList;
}

void DataVisualization::addSelectedRowColumData(int column)
{
	DataList dataList;

	for (int i = 1; i < tableWidget->rowCount(); i++)
	{
		if (tableWidget->isRowHidden(i) || !rowChecked[i])
		{
			continue;
		}
		auto item = tableWidget->item(i, column);
		if (item == nullptr)
		{
			dataList << Data(QPointF(i, 0),
				QString::number(column) + ":" + QString::number(i));
			continue;
		}
		dataList << Data(QPointF(i, item->text().toDouble()),
			QString::number(column) + ":" + QString::number(i));
	}
	m_dataTable << dataList;
}

void DataVisualization::updataChartData()
{
	QString content = ui.chartLineEdit->text();
	content.replace("��", ",");
	QStringList stringList = content.split(',');
	for (QString s : stringList)
	{
		int column = headerString2ColumnNumber(s);
		if (column == -1)
		{
			printf("can't find column : %s\n", qPrintable(s));
			continue;
		}
		addSelectedRowColumData(column);
	}
}

void DataVisualization::headerClicked(int i ) {
	printf("headerClicked index : %d", i);
	addData(i);
	QChart* previous = chartView->chart();
	chartView->setChart(createLineChart());
	if(previous)
	    delete previous;
}

void DataVisualization::displayData(const QList<QList<QVariant>>& data, 
	const  std::vector<QString>& names)
{
	int rows = data.size();
	int columns = data.at(0).size();
	tableWidget->clear();
	tableWidget->setRowCount(rows);
	tableWidget->setColumnCount(columns);

	QStringList headers;
	for (auto head: names)
	{
		headers << head;
	}
	tableWidget->setHorizontalHeaderLabels(headers);
	for (int i = 0; i < data.size(); i++)
	{
		QList<QVariant> allEnvDataList_i = data[i];//��i�е�����
		//printf("\n");
		for (int j = 0; j < allEnvDataList_i.size(); j++)
		{
			QString tempvalue = allEnvDataList_i[j].toString();
			//printf("%s, ", qPrintable(tempvalue));
			QTableWidgetItem* item = new QTableWidgetItem(tempvalue);
			tableWidget->setItem(i, j, item);
		}
	}
}

void DataVisualization::displayData(const std::vector<std::vector<QVariant>>& data,
	int beginRow, int headerRow)
{
	int rows = data.size() - beginRow +1 ;
	rowChecked = std::vector<bool>(rows, true);
	int columns = data.at(0).size() + 1;
	tableWidget->clear();
	tableWidget->setRowCount(rows);
	//tableWidget->setRowCount(21);
	tableWidget->setColumnCount(columns);

	//tableWidget->setHorizontalHeader();

	//QHeaderView* firstHeader = new QHeaderView(Qt::Horizontal, tableWidget);
	//firstHeader->setItemDelegate();

	//tableWidget->setHorizontalHeader(firstHeader);

	QStringList headers;
	headers << "";
	for (auto head : data[headerRow])
	{
		headers << head.toString();
	}
	tableWidget->setHorizontalHeaderLabels(headers);

	for (int i = beginRow; i < data.size(); i++)
	//for (int i = beginRow; i < beginRow + 20; i++)
	{
	    std::vector<QVariant>  allEnvDataList_i = data[i];//��i�е�����
		// set check box
		QCheckBox* CheckBox = new QCheckBox(tableWidget);
		CheckBox->setFixedSize(QSize(39,35));
		CheckBox->setCheckState(Qt::Checked);
		CheckBox->setWhatsThis(QString::number(i - beginRow));
		connect(CheckBox, SIGNAL(stateChanged(int)), this, SLOT(checkBoxchange(int)));
		tableWidget->setCellWidget(i - beginRow, 0, CheckBox);//row�У�0��

		for (int j = 1; j < allEnvDataList_i.size(); j++)
		{
			QString tempvalue = allEnvDataList_i[j].toString();
			//printf("%s, ", qPrintable(tempvalue));
			QTableWidgetItem* item = new QTableWidgetItem(tempvalue);
			tableWidget->setItem(i - beginRow + 1, j, item);
		}
	}
    //create combobox
	for (const QString& name : selectHeaderName)
	{
		int column = headerString2ColumnNumber(name);
		ItemSelectCombox* combo = createSelectCombox(name);
		tableWidget->setCellWidget(0, column, combo->content);//row�У�0��
	}
}

//traverse all row to updata filter view
void DataVisualization::displaySelectRow(const std::vector<int>& rowsNumber)
{
	for (int i = 1; i < tableWidget->rowCount(); ++i)
	{
		tableWidget->setRowHidden(i,
			(std::find(rowsNumber.cbegin(), rowsNumber.cend(), i) == rowsNumber.cend()));
	}
}

std::vector<int>  DataVisualization::getSelectRowNumber(const std::vector<QString>& headType,
	const std::vector<QString>& traget)
{
	Q_ASSERT(headType.size() == traget.size());
	int tragetNumber = headType.size();
	std::vector<int> tragetRowNumberCount(tableWidget->rowCount(), 0);
	for (int i = 0; i < headType.size(); ++i)
	{
		if (traget[i] == QStringLiteral("ȫ��"))
		{
			--tragetNumber;
			continue;
		}
		int column = headerString2ColumnNumber(headType[i]);
		if (column != -1)
		{
			for (int row = 0; row < tableWidget->rowCount(); ++row)
			{
				auto currentItem  = tableWidget->item(row, column);
				if (currentItem == nullptr)
					continue;
				if (currentItem->text() == traget[i])
				{
					++tragetRowNumberCount[row];
				}
			}
		}
		else
		{
			printf("Can not find column :%s\n", qPrintable(headType[i]));
			--tragetNumber;
		}
	}
	std::vector<int> result;
	for (int index = 0; index < tragetRowNumberCount.size(); ++index)
	{
		if (tragetRowNumberCount[index] == tragetNumber)
		{
			result.push_back(index);
		}
	}
	return result;
}

int DataVisualization::headerString2ColumnNumber(const QString& headerName)
{
	for (int i = 0; i < tableWidget->columnCount(); ++i)
	{
		if (tableWidget->horizontalHeaderItem(i)->text() == headerName)
		{
			return i;
		}
	}
	return -1;
}

void DataVisualization::checkBoxchange(int state)
{
	QCheckBox* check = (QCheckBox*)sender();
	int row = check->whatsThis().toInt();
	rowChecked[row] = (state == Qt::Checked);
}

void DataVisualization::buttonPress()
{
	/*std::vector<int> result = getSelectRowNumber(std::vector<QString>{QString::fromLocal8Bit(std::string("Ӫ�����̷�").data())},
		std::vector<QString>{QString::fromLocal8Bit(std::string("�³�").data())});
	displaySelectRow(result);*/
	/*if (count == 0)
	{
		addSelectCombox(QStringLiteral("��ҵ��\n��ס��/�̿���"));
		addSelectCombox(QStringLiteral("Ӫ�����̷�"));
		addSelectCombox(QStringLiteral("���л���"));
	}
	else
	{
		filterItem();
	}
	++count;*/
	/*scatterData = std::map<QString, std::vector<double>>
	{ {QStringLiteral("x��"),std::vector<double >{2, 3,4,5,6.5,1.2,2.6,7.5}},
	  {QStringLiteral("y��"),std::vector<double >{3, 1.2,4.7,5.3,6.5,1.2,7.8,7.5}}
	};*/
	
	//QChart* currentChart = createScatterChart();
}

void DataVisualization::uniqueItem(const QString& headerName, std::vector<QString>& items)
{
	int column = headerString2ColumnNumber(headerName);
	if (column != -1)
	{
		int rowCount = tableWidget->rowCount();
		for (int row = 0; row < rowCount; ++row)
		{
			auto currentItem = tableWidget->item(row, column);
			if (currentItem == nullptr)
				continue;
			if (std::find(items.cbegin(), items.cend(), currentItem->text()) == items.cend())
			{
				items.push_back(currentItem->text());
			}
		}
	}
	else
	{
		printf("Can not find column :%s\n", qPrintable(headerName));
	}
	
}

void DataVisualization::addSelectCombox(const QString& headerName)
{
	std::vector<QString> uniqueNames;
	uniqueItem(headerName, uniqueNames);

	ItemSelectCombox* view = new ItemSelectCombox();
	view->label = new QLabel(headerName);
	view->content = new QComboBox();
	view->content->addItem(QStringLiteral("ȫ��"));
	for (const QString& name : uniqueNames)
	{
		view->content->addItem(name);
	}
	ui.itemSelect.push_back(view);
	ui.topCombox->addWidget(view->label);
	ui.topCombox->addWidget(view->content);
}

ItemSelectCombox* DataVisualization::createSelectCombox(const QString& headerName)
{
	std::vector<QString> uniqueNames;
	uniqueItem(headerName, uniqueNames);

	ItemSelectCombox* view = new ItemSelectCombox();
	view->label = new QLabel(headerName);
	QComboBox* cccombox = new QComboBox();
	view->content = cccombox;
	
	view->content->addItem(QStringLiteral("ȫ��"));
	for (const QString& name : uniqueNames)
	{
		view->content->addItem(name);
	}
	view->content->setObjectName(headerName);
	ui.itemSelect.push_back(view);
	connect(cccombox, SIGNAL(currentTextChanged(const QString&)), this, SLOT(comboxChanged(const QString&)));
	/*ui.topCombox->addWidget(view->label);
	ui.topCombox->addWidget(view->content);*/
	return view;
}
void DataVisualization::comboxChanged(const QString& text)
{
	//QComboBox* combox = (QComboBox*)sender();
	filterItem();
}

void DataVisualization::filterItem()
{
	std::vector<QString> headerName;
	std::vector<QString> filterName;

	for (int i = 0; i < ui.itemSelect.size(); ++i)
	{
		ItemSelectCombox* selected = ui.itemSelect.at(i);
		headerName.push_back(selected->label->text());
		filterName.push_back(selected->content->currentText());
	}

	std::vector<int> rows = getSelectRowNumber(headerName, filterName);
	displaySelectRow(rows);
}