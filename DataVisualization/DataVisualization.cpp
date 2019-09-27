#include <QTimer>
#include "DataVisualization.h"
#include "charttip.h"

DataVisualization::DataVisualization(QWidget* parent)
	: QMainWindow(parent),
	excelServer(nullptr),
	worrkbook(nullptr),
	usedrange(nullptr),
	startrow(4),
    endrow(100)
{
	ui.setupUi(this);

	tableWidget = new QTableWidget(this);
	ui.viewLayout->addWidget(tableWidget);

	//chart view 
	chartView = new QChartView(ui.centralWidget);
	ui.viewLayout->addWidget(chartView);

	connect(ui.actionOpenFile, &QAction::triggered, this, &DataVisualization::openFile);
	//connect(ui.actionOpenChart, &QAction::triggered, this, &DataVisualization::openFile);
	connect(ui.histogramChart, &QAction::triggered, this, &DataVisualization::displayBarChart);
	connect(ui.scatterChart, &QAction::triggered, this, &DataVisualization::displayScatterChart);
	connect(ui.lineChart, &QAction::triggered, this, &DataVisualization::displayLineChart);
	connect(ui.calculate, &QAction::triggered, this, &DataVisualization::excute);
	connect(ui.saveFile, &QAction::triggered, this, &DataVisualization::saveFile);
	connect(ui.templateExport, &QAction::triggered, this, &DataVisualization::templateExport);
	connect(ui.savePic, &QAction::triggered, this, &DataVisualization::picExport);
}

DataVisualization::~DataVisualization()
{
	
	if (excelServer)
	{
		//saveFile();
		excelServer->freeExcel();
	}
	    
}

void DataVisualization::saveFile()
{
	if (excelServer != nullptr)
	{
		printf("Save excel file ...\n");
		QVariant var;
		excelServer->castSheetVector2Variant(var);
		usedrange->setProperty("Value", var);
		worrkbook->dynamicCall("Save()");

		QMessageBox *box = new QMessageBox(this);
		box->setIcon(QMessageBox::Information);
		box->setWindowTitle(QStringLiteral("提示"));
		box->setText(QStringLiteral("保存<font color='red'>成功</font>"));
		QTimer::singleShot(1000, box, SLOT(accept())); //也可将accept改为close
		box->exec();
	}
	else
	{
		QMessageBox *box = new QMessageBox(this);
		box->setIcon(QMessageBox::Information);
		box->setWindowTitle(QStringLiteral("提示"));
		box->setText(QStringLiteral("请先打开文件"));
		QTimer::singleShot(1000, box, SLOT(accept())); //也可将accept改为close
		box->exec();
	}
}

void DataVisualization::openFile()
{
	QString filePath = QFileDialog::getOpenFileName();
	printf("file path : %s\n", qPrintable(filePath));

	//excel server
	excelServer = new ExcelDataServer();

	worrkbook = excelServer->openExcelFile(filePath);
	if (worrkbook == NULL)
	{
		printf("open file failed : %s, %p\n", qPrintable(filePath), worrkbook);
		return;
	}
	else
	{
		printf("open success\n");
	}
	printf("Excel initialization...\n");
	QAxObject* worksheets = worrkbook->querySubObject("WorkSheets");
	QAxObject* sheet = excelServer->getSheet(worrkbook, 1);//updata work sheets.
	//QAxObject* sheet = excelServer->getNamedSheet(worksheets,
		//QString::fromLocal8Bit(std::string("��������").data()));
	excelServer->setCurrentWorksheet(sheet);

	//add a new sheet
	//QAxObject * newSheets = excelServer->addSheet(excelServer->getCurrentWorkSheets(), QString("selectedTest"));
	usedrange = sheet->querySubObject("UsedRange");
	excelServer->setAllData(usedrange);
	int columsNumber = excelServer->getColumsNumber();
	int rowsNumber = excelServer->getRowsNumber();

	startrow = 4;
	endrow = rowsNumber;
	//���㿪ʼ��������
	excelServer->setBeginEndRow(startrow, endrow);

	//initialize calculate server
	proService = service(startrow, endrow);
	//report = proService.getReport();
	initExportenu();

	QVariantList resultColum3;
	excelServer->getRowData(sheet, 3, resultColum3);//��ȡ�����е�ֵ

	printf("Excel initialization complete\n");
	printf("Analysis data...\n");
	displayData(excelServer->sheetContent, 3, 2);
	QMessageBox *box = new QMessageBox(this);
	box->setIcon(QMessageBox::Information);
	box->setWindowTitle(QStringLiteral("提示"));
	box->setText(QStringLiteral("打开<font color='red'>成功</font>"));
	QTimer::singleShot(1000, box, SLOT(accept())); //也可将accept改为close
	box->exec();
}

void DataVisualization::excute()
{
	if (excelServer == nullptr)
	{
		QMessageBox::about(this, QStringLiteral("提示"), QStringLiteral("请先打开文件.\n"));
		return;
	}
	printf("Excute calculator... ...\n");
	
    proService.confirm(excelServer);
	printf("Complete...\n");
	printf("updata table view ...\n");
	updataContent(excelServer->sheetContent, 3, 2);
	printf("updata complete...\n");

	QMessageBox *box = new QMessageBox(this);
	box->setIcon(QMessageBox::Information);
	box->setWindowTitle(QStringLiteral("提示"));
	box->setText(QStringLiteral("计算<font color='red'>完成</font>"));
	QTimer::singleShot(1000, box, SLOT(accept())); //也可将accept改为close
	box->exec();
}

QChart* DataVisualization::createLineChart() const
{
	QString content = ui.chartLineEdit->text();
	content.replace("，", ",");
	QStringList stringList = content.split(',');

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
			
		series->setName(stringList[nameIndex]);
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
	Q_ASSERT(m_dataTable.size() > 1);
	// scatter chart
	QString content = ui.chartLineEdit->text();
	QStringList stringLists = content.split(';');
	//QString title = stringLists[1];
	QChart* chart = new QChart();
	//chart->setTitle(title);
	chart->legend()->hide();
	DataList dataList_x = m_dataTable[0];
	DataList dataList_y = m_dataTable[1];

	mytips.clear();
	double valueMaxX = std::numeric_limits<double>::min();
	double valueMinX = std::numeric_limits<double>::max();

	double valueMaxY = std::numeric_limits<double>::min();
	double valueMinY = std::numeric_limits<double>::max();

	QScatterSeries* series = new QScatterSeries(chart);
	for (int i = 0; i < dataList_x.size(); ++i)
	{
		dataList_x.at(i).first.y() > valueMaxX ? (valueMaxX = dataList_x.at(i).first.y()) : true;
		dataList_x.at(i).first.y()< valueMinX ? (valueMinX = dataList_x.at(i).first.y()) : true;

		dataList_y.at(i).first.y() > valueMaxY ? (valueMaxY = dataList_y.at(i).first.y()) : true;
		dataList_y.at(i).first.y()< valueMinY ? (valueMinY = dataList_y.at(i).first.y()) : true;

		series->append(QPointF(dataList_x.at(i).first.y(), dataList_y.at(i).first.y()));
		ChartTip* m_tooltip = new ChartTip(chart);
		mytips.append(m_tooltip);
		m_tooltip->setText(QString(dataList_x.at(i).second));
		QPointF point(dataList_x.at(i).first.y(), dataList_y.at(i).first.y());
		m_tooltip->setAnchor(point);
	}	

	chart->addSeries(series);

	QValueAxis * m_typeAxisY = new QValueAxis;
	QValueAxis * m_typeAxisX = new QValueAxis;

	double absy = abs(valueMaxY) > abs(valueMinY) ? abs(valueMaxY) : abs(valueMinY);
	double absx = abs(valueMaxX) > abs(valueMinX) ? abs(valueMaxX) : abs(valueMinX);

	m_typeAxisY->setRange(-absy*1.5, absy*1.5);
	m_typeAxisX->setRange(-absx*1.5, absx*1.5);

	//m_typeAxisY->setGridLineVisible(false);   //隐藏背景网格Y轴框线
	//m_typeAxisX->setGridLineVisible(false);   //隐藏背景网格X轴框线
	m_typeAxisX->setVisible(false);
	m_typeAxisY->setVisible(false);

	chart->setAxisX(m_typeAxisX, series);
	chart->setAxisY(m_typeAxisY, series);
	
	//for (ChartTip* tmp : mytips) {
	//	tmp->updateGeometry();
	//	tmp->show();
	//}
	
	//chart->createDefaultAxes();
	chart->setAnimationOptions(QChart::SeriesAnimations);
	chart->setBackgroundVisible(false);
	//chart->setDropShadowEnabled(false);
	return chart;
}

QChart* DataVisualization::createBarChart()
{
	QString content = ui.chartLineEdit->text();
	QStringList stringLists = content.split(';');
	if (stringLists.size() != 3) {
		QMessageBox::about(this, QStringLiteral("提示"), QStringLiteral("输入格式错误"));
		QChart* chartfalse = new QChart();
		return chartfalse;
	}
	QStringList field = stringLists[0].split(',');
	QStringList map = stringLists[1].split(',');
	if (field.size() != map.size()) {
		QMessageBox::about(this, QStringLiteral("提示"), QStringLiteral("字段数量不匹配"));
		QChart* chartfalse = new QChart();
		return chartfalse;
	}
	QString title = stringLists[2];
	QChart* chart = new QChart();
	int valueMax = std::numeric_limits<int>::min();
	int valueMin = std::numeric_limits<int>::max();
	//QStackedBarSeries* series = new QStackedBarSeries(chart);
	QBarSeries* series = new QBarSeries(chart);
	for (int i(0); i < m_dataTable.size(); i++) {
		QBarSet* set = new QBarSet(map[i]);
		for (const Data& data : m_dataTable[i])
		{
			*set << data.first.y();
			data.first.y() > valueMax ? (valueMax = data.first.y()) : true;
			data.first.y() < valueMin ? (valueMin = data.first.y()) : true;
		}
		series->append(set);
	}
	chart->addSeries(series);
	chart->setTitle(title);
	chart->setAnimationOptions(QChart::SeriesAnimations);

	series->setLabelsPosition(QAbstractBarSeries::LabelsInsideEnd); // 设置数据系列标签的位置于数据柱内测上方
	series->setLabelsVisible(true);

	QStringList categories;
	for(int tmp=0;tmp< m_dataTable[0].size();tmp++)
		categories << m_dataTable[0][tmp].second;
	QBarCategoryAxis *axis = new QBarCategoryAxis();
	chart->createDefaultAxes();
	axis->append(categories);
	chart->setAxisX(axis, series);

	chart->legend()->setVisible(true); //设置图例为显示状态
	chart->legend()->setAlignment(Qt::AlignBottom);
	return chart;
}

void DataVisualization::displayLineChart()
{
	if (excelServer == nullptr)
	{
		QMessageBox::about(this, QStringLiteral("提示"), QStringLiteral("请先打开文件.\n"));
		return;
	}
	//m_dataTable.clear();
	//updataChartData();
	//QChart* currentChart = createLineChart();
	//QChart* previous = chartView->chart();
	//chartView->setChart(currentChart);
	//if (previous)
		//delete previous;
}

void DataVisualization::displayScatterChart()
{
	if (excelServer == nullptr)
	{
		QMessageBox::about(this, QStringLiteral("提示"), QStringLiteral("请先打开文件.\n"));
		return;
	}
	m_dataTable.clear();
	bool state = updataChartDataForScatter();
	if (state) {
		QChart * currentChart = createScatterChartTwo();
		QChart* previous = chartView->chart();
		chartView->setChart(currentChart);
		chartView->setStyleSheet(R"(QGraphicsView{ background-image:url(:/DataVisualization/Resources/back.jpg);})");
		for (ChartTip* tmp : mytips) {
			tmp->updateGeometry();
		}
		if (previous)
			delete previous;
	}
}

void DataVisualization::displayBarChart()
{
	if (excelServer == nullptr)
	{
		QMessageBox::about(this, QStringLiteral("提示"), QStringLiteral("请先打开文件.\n"));
		return;
	}
	m_dataTable.clear();
	bool state = updataChartData();
	if (state) {
		QChart* currentChart = createBarChart();
		QChart* previous = chartView->chart();
		chartView->setChart(currentChart);
		if (previous)
			delete previous;
	}
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
				QString(tableWidget->item(i, 1)->text()));
			continue;
		}
		dataList << Data(QPointF(i, item->text().toDouble()),
			QString(tableWidget->item(i, 1)->text()));
	}
	m_dataTable << dataList;
}

void DataVisualization::addSelectedRowColumDataForScatter(int column)
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
				QString(tableWidget->item(i, 4)->text()));
			continue;
		}
		dataList << Data(QPointF(i, item->text().toDouble()),
			QString(tableWidget->item(i, 4)->text()));
	}
	m_dataTable << dataList;
}

bool DataVisualization::updataChartData()
{
	QString content = ui.chartLineEdit->text();
	//content = content.replace(QChar('，'), QChar(','));
	//content = content.replace(QChar('；'), QChar(';'));
	QStringList stringLists = content.split(';');
	QStringList stringList = stringLists[0].split(',');
	for (QString s : stringList)
	{
		int column = headerString2ColumnNumber(s);
		if (column == -1)
		{
			QMessageBox::about(this, QStringLiteral("提示"), QStringLiteral("表中无“") + s + QStringLiteral("”字段"));
			return false;
		}
		addSelectedRowColumData(column);
		
	}
	return true;
}

bool DataVisualization::updataChartDataForScatter()
{
	QString content = ui.chartLineEdit->text();
	//content = content.replace(QChar('，'), QChar(','));
	//content = content.replace(QChar('；'), QChar(';'));
	QStringList stringLists = content.split(';');
	QStringList stringList = stringLists[0].split(',');
	for (QString s : stringList)
	{
		int column = headerString2ColumnNumber(s);
		if (column == -1)
		{
			QMessageBox::about(this, QStringLiteral("提示"), QStringLiteral("表中无“")+ s+ QStringLiteral("”字段"));
			return false;
		}
		addSelectedRowColumDataForScatter(column);
	}
	return true;
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
		QList<QVariant> allEnvDataList_i = data[i];//第i行的数据
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

void DataVisualization::updataContent(const std::vector<std::vector<QVariant>>& data,
	int beginRow, int headerRow)
{
	int rows = data.size() - beginRow + 1;
	int columns = data.at(0).size() + 1;

	for (int i = beginRow; i < data.size(); i++)
		//for (int i = beginRow; i < beginRow + 20; i++)
	{
		std::vector<QVariant>  allEnvDataList_i = data[i];//第i行的数据

		for (int j = 0; j < allEnvDataList_i.size(); j++)
		{
			QString tempvalue = allEnvDataList_i[j].toString();
			//printf("%s, ", qPrintable(tempvalue));
			auto item = tableWidget->item(i - beginRow + 1, j + 1);
			if (item == nullptr)
			{
				continue;
			}
			item->setText(tempvalue);
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
	    std::vector<QVariant>  allEnvDataList_i = data[i];//第i行的数据
		// set check box
		QCheckBox* CheckBox = new QCheckBox(tableWidget);
		CheckBox->setFixedSize(QSize(39,35));
		CheckBox->setCheckState(Qt::Checked);
		CheckBox->setWhatsThis(QString::number(i - beginRow));
		connect(CheckBox, SIGNAL(stateChanged(int)), this, SLOT(checkBoxchange(int)));
		tableWidget->setCellWidget(i - beginRow, 0, CheckBox);//row行，0列

		for (int j = 0; j < allEnvDataList_i.size(); j++)
		{
			QString tempvalue = allEnvDataList_i[j].toString();
			//printf("%s, ", qPrintable(tempvalue));
			QTableWidgetItem* item = new QTableWidgetItem(tempvalue);
			tableWidget->setItem(i - beginRow + 1, j + 1, item);
		}
	}
    //create combobox
	ui.itemSelect.clear();//清楚原有表的combobox
	int columnlist = 0; //反向搜索的索引
	for (const QString& name : selectHeaderName)
	{
		int column = headerString2ColumnNumber(name);
		if (column != -1) {
			ItemSelectCombox* combo = createSelectCombox(name, columnlist);
			tableWidget->setCellWidget(0, column, combo->content);//row行，0列
			columnlist++;
		}	
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
		int column = headerString2ColumnNumber(headType[i]);
		if (column != -1)
		{
			for (int row = 0; row < tableWidget->rowCount(); ++row)
			{
				auto currentItem  = tableWidget->item(row, column);
				if (currentItem == nullptr)
					continue;
				QStringList propertys = traget[i].split(";");
				for (int temp = 0; temp<propertys.size(); temp++)
				{
					if (currentItem->text() == propertys[temp]) {
						++tragetRowNumberCount[row];
						break;
					}					
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

//to do
ItemSelectCombox* DataVisualization::createSelectCombox(const QString& headerName, int v)
{
	
	std::vector<QString> uniqueNames;
	uniqueItem(headerName, uniqueNames);

	ItemSelectCombox* view = new ItemSelectCombox();
	view->label = new QLabel(headerName);
	QListWidget* listWidget = new QListWidget;
	view->alist = listWidget;
	MyCombbox* cccombox =new MyCombbox(v);
	view->content = cccombox;
	for (const QString& name : uniqueNames)
	{
		QListWidgetItem* item = new QListWidgetItem(name);
		item->setCheckState(Qt::Unchecked);
		listWidget->addItem(item);
	}
	//QListWidgetItem* item = listWidget->item(0);
	//if (item) {
	//	item->setCheckState(Qt::Checked);
	//}
	//设置数据源到显示控件中
	cccombox->setModel(listWidget->model());
	cccombox->setView(listWidget);
	
	cccombox->setEditable(true);
	QLineEdit* lineEdit = cccombox->lineEdit();
	if (lineEdit) {
		lineEdit->setReadOnly(true);
	}
	connect(cccombox, SIGNAL(activated(int)), this, SLOT(onCurrentIndexChanged(int)));

	//view->content->addItem(QStringLiteral("全部"));
	//for (const QString& name : uniqueNames)
	//{
	//	view->content->addItem(name);
	//}
	ui.itemSelect.push_back(view);
	//connect(cccombox, SIGNAL(currentTextChanged(const QString&)), this, SLOT(comboxChanged(const QString&)));
	/*ui.topCombox->addWidget(view->label);
	ui.topCombox->addWidget(view->content);*/
	return view;
}

void DataVisualization::onCurrentIndexChanged(int index) {
	//获取当前点击的对象
	QObject *object = QObject::sender();
	MyCombbox *cccombox = static_cast<MyCombbox *>(object);
	int fullIndex = cccombox->getIndex();
	QListWidget*listWidget = ui.itemSelect[fullIndex]->alist;
	//获取当前点击的对象
	QListWidgetItem* item = listWidget->item(index);
	if (!item)
		return;

	//更新当前点击对象的选中状态
	if (item->checkState() == Qt::Unchecked)
	{
		item->setCheckState(Qt::Checked);
	}
	else
	{
		item->setCheckState(Qt::Unchecked);
	}

	//循环获取所有选中状态的对象显示文字
	QString text;
	for (int row = 0, rows = listWidget->count(); row < rows; ++row)
	{
		QListWidgetItem* item = listWidget->item(row);
		if (item && item->checkState() == Qt::Checked)
		{
			text.append(item->text() + ";");
		}
	}
	//更新显示控件的文字
	cccombox->lineEdit()->setText(text.mid(0, text.size() - 1));

	std::vector<QString> headerName;
	std::vector<QString> filterName;

	for (int i = 0; i < ui.itemSelect.size(); ++i)
	{
		ItemSelectCombox* selected = ui.itemSelect.at(i);
		headerName.push_back(selected->label->text());
		filterName.push_back(selected->content->lineEdit()->text());
	}
	std::vector<int> rows = getSelectRowNumber(headerName, filterName);
	displaySelectRow(rows);
}

void DataVisualization::templateExport()
{
	if (excelServer == nullptr)
	{
		QMessageBox::about(this, QStringLiteral("提示"), QStringLiteral("请先打开文件.\n"));
		return;
	}
	QString filePath = QFileDialog::getOpenFileName();
	printf("file path : %s\n", qPrintable(filePath));
	if (filePath == "")
	{
		printf("file path is empty %s\n");
		return;
	}
	excelServer->templateExport(filePath, 3);
	printf("Export complete.\n");
	QMessageBox *box = new QMessageBox(this);
	box->setIcon(QMessageBox::Information);
	box->setWindowTitle(QStringLiteral("提示"));
	box->setText(QStringLiteral("模板输出<font color='red'>完成</font>"));
	QTimer::singleShot(1000, box, SLOT(accept())); //也可将accept改为close
	box->exec();

}

void DataVisualization::picExport() {

	QScreen * screen = QGuiApplication::primaryScreen();
	QPixmap p = chartView->grab(); 
	QImage image = p.toImage();
	image.save("./pics/chart.png");
	QMessageBox *box = new QMessageBox(this);
	box->setIcon(QMessageBox::Information);
	box->setWindowTitle(QStringLiteral("提示"));
	box->setText(QStringLiteral("保存图片<font color='red'>完成</font>"));
	QTimer::singleShot(1000, box, SLOT(accept())); //也可将accept改为close
	box->exec();
}

//obsolete
void DataVisualization::comboxChanged(const QString&)
{
	//QComboBox* combox = (QComboBox*)sender();
	filterItem();
}

//obsolete
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

//obsolete
void DataVisualization::initExportenu()
{
	auto iter = report.cbegin();

	while (iter != report.cend())
	{
		QAction* exportAction = new QAction();
		exportAction->setText(iter.key());
		ui.menuExport->addAction(exportAction);
		ui.exportActions.push_back(exportAction);
		connect(exportAction, &QAction::triggered, this, &DataVisualization::doExport);
		++iter;
	}
}

//obsolete
void DataVisualization::doExport()
{
	QAction* action = (QAction*)sender();
	printf("Export something...%s\n", qPrintable(action->text()));
	QStringList exportColumnList = report[action->text()];
	QList<QList<QVariant>> exportData;

	QList<QVariant> currentRow;
	for (auto s : exportColumnList)
	{
		currentRow << QVariant(s);
	}
	exportData << currentRow;

	std::vector<std::vector<QVariant>>& allContent = excelServer->sheetContent;
	for (int i = startrow - 1; i < allContent.size(); ++i)
	{
		if (tableWidget->isRowHidden(i - startrow + 2) || !rowChecked[i - startrow + 2])
			continue;
		QList<QVariant> currentRow;

		for (QString columnName : exportColumnList)
		{
			currentRow << allContent[i][excelServer->nameToSubScript[columnName]];
		}
		exportData << currentRow;
	}

	excelServer->exportSheet(exportData, action->text());
}

//obsolete
void DataVisualization::addSelectCombox(const QString& headerName)
{
	std::vector<QString> uniqueNames;
	uniqueItem(headerName, uniqueNames);

	ItemSelectCombox* view = new ItemSelectCombox();
	view->label = new QLabel(headerName);
	//view->content = new MyCombbox();
	view->content->addItem(QStringLiteral("全部"));
	for (const QString& name : uniqueNames)
	{
		view->content->addItem(name);
	}
	ui.itemSelect.push_back(view);
	ui.topCombox->addWidget(view->label);
	ui.topCombox->addWidget(view->content);
}

//obsolete
void DataVisualization::buttonPress()
{
	/*std::vector<int> result = getSelectRowNumber(std::vector<QString>{QString::fromLocal8Bit(std::string("营销操盘方").data())},
	std::vector<QString>{QString::fromLocal8Bit(std::string("新城").data())});
	displaySelectRow(result);*/
	/*if (count == 0)
	{
	addSelectCombox(QStringLiteral("事业部\n（住开/商开）"));
	addSelectCombox(QStringLiteral("营销操盘方"));
	addSelectCombox(QStringLiteral("城市环线"));
	}
	else
	{
	filterItem();
	}
	++count;*/
	/*scatterData = std::map<QString, std::vector<double>>
	{ {QStringLiteral("x轴"),std::vector<double >{2, 3,4,5,6.5,1.2,2.6,7.5}},
	{QStringLiteral("y轴"),std::vector<double >{3, 1.2,4.7,5.3,6.5,1.2,7.8,7.5}}
	};*/

	//QChart* currentChart = createScatterChart();
}

//obsolete
void DataVisualization::headerClicked(int i) {
	printf("headerClicked index : %d", i);
	addData(i);
	QChart* previous = chartView->chart();
	chartView->setChart(createLineChart());
	if (previous)
		delete previous;
}
