#include "DataVisualization.h"

DataVisualization::DataVisualization(QWidget *parent)
	: QMainWindow(parent)
{
	ui.setupUi(this);

	tableWidget = new QTableWidget(this);

	hLayout = new QHBoxLayout();
	hLayout->addWidget(tableWidget);
	ui.centralWidget->setLayout(hLayout);

	//chart view 
	chartView = new QChartView(this);
	//chartView->setChart();
	hLayout->addWidget(chartView);

	//select event

	connect(tableWidget->horizontalHeader(), SIGNAL(sectionClicked(int)),this, SLOT(headerClicked(int)));
}



void DataVisualization::Read_Excel(const QString PATH, const QString FILENAME,
	const int SHEETNUM, const QString RANGE, const int INVALIDROW, 
	const int TOTALCOLNUM, std::vector<QString>& RESULT)
//�������ͣ�·�����ļ������ڼ���sheet����ȡ��Χ����ʽΪA1:B������Ч�����������粻��Ҫ��title�ȣ�
//����ȡ��Χ��������������һ��QString��vector��
{
	QString pathandfilename = excelFilePath;
	QAxObject excel("Excel.Application");
	excel.setProperty("Visible", false); //���ش򿪵�excel�ļ�����
	QAxObject* workbooks = excel.querySubObject("WorkBooks");
	QAxObject* workbook = workbooks->querySubObject("Open(QString, QVariant)", pathandfilename); //���ļ�
	QAxObject* worksheet = workbook->querySubObject("WorkSheets(int)", SHEETNUM); //���ʵ�SHEETNUM��������
	QAxObject* usedrange = worksheet->querySubObject("UsedRange");
	QAxObject* rows = usedrange->querySubObject("Rows");
	int rownum = rows->property("Count").toInt(); //��ȡ����

	QString Range = RANGE + QString::number(rownum); 
	QAxObject* allEnvData = worksheet->querySubObject("Range(QString)", Range); //��ȡ��Χ
	//QAxObject* allEnvData = worksheet->querySubObject("Range(QString)", "A1:B2"); //��ȡ��Χ
	QVariant allEnvDataQVariant = allEnvData->property("Value");//��ȡ���е�ֵ

	QVariantList allEnvDataList = allEnvDataQVariant.toList();//ת��Ϊlist

	for (int i = 0; i < rownum - INVALIDROW; i++)
	{
		QVariantList allEnvDataList_i = allEnvDataList[i].toList();//��i�е�����
		printf("\n");
		for (int j = 0; j < TOTALCOLNUM; j++)
		{
			QString tempvalue = allEnvDataList_i[j].toString();
			printf("%s, ", qPrintable(tempvalue));
			RESULT.push_back(tempvalue);
		}
	}
	workbooks->dynamicCall("Close()");
	excel.dynamicCall("Quit()");
}

void DataVisualization::readAll(const QString Path)
{
	QString pathandfilename = Path;
	QAxObject excel("Excel.Application");
	excel.setProperty("Visible", false); //���ش򿪵�excel�ļ�����
	QAxObject* workbooks = excel.querySubObject("WorkBooks");//�ɴ򿪶��excel

	QAxObject* workbook = workbooks->querySubObject("Open(QString, QVariant)", pathandfilename); //���ļ�
	QAxObject* worksheet = workbook->querySubObject("WorkSheets(int)", 1); //���ʵ�SHEETNUM��������
	
	QAxObject* cellA = worksheet->querySubObject("Range(QVariant, QVariant)", "A8");
	cellA->dynamicCall("SetValue(const QVariant&)", QVariant("1+1"));//���õ�Ԫ���ֵ
	
	//pWorkbook->dynamicCall("SaveAs(QString)", "E:\\456.xlsx");���Ϊ

	QAxObject* usedrange = worksheet->querySubObject("UsedRange");//��������

	QAxObject* rows = usedrange->querySubObject("Rows");
    int rownum = rows->property("Count").toInt(); //��ȡ����Columns

    QAxObject* cols = usedrange->querySubObject("Columns");
    int columns = cols->property("Count").toInt(); //��ȡ����Columns

	QVariant allEnvDataQVariant = usedrange->dynamicCall("Value");//->dynamicCall("Value");
	QVariantList allEnvDataList = allEnvDataQVariant.toList();//ת��Ϊlist

	
	tableWidget->setRowCount(rownum);
	tableWidget->setColumnCount(columns);
	QStringList headers;
	for (int c = 0; c < columns; ++c)
	{
		headers << QString(char(65 + c));
	}

	tableWidget->setHorizontalHeaderLabels(headers);
	for (int i = 0; i < allEnvDataList.size(); i++)
	{
		QVariantList allEnvDataList_i = allEnvDataList[i].toList();//��i�е�����
		//printf("\n");
		for (int j = 0; j < allEnvDataList_i.size(); j++)
		{
			QString tempvalue = allEnvDataList_i[j].toString();
			//printf("%s, ", qPrintable(tempvalue));
			QTableWidgetItem* item = new QTableWidgetItem(tempvalue);
			tableWidget->setItem(i, j, item);
		}
	}

	workbooks->dynamicCall("Close()");
	excel.dynamicCall("Quit()");
}

//��̬����excel����
void DataVisualization::import()
{

	QAxObject* excel = new QAxObject("Excel.Application");
	excel->setProperty("Visible", false);
	QAxObject* workbooks = excel->querySubObject("WorkBooks");
	workbooks->dynamicCall("Open (const QString&)", excelFilePath);
	QAxObject* workbook = excel->querySubObject("ActiveWorkBook");//��ȡ�������

	QAxObject* worksheets = workbook->querySubObject("Sheets");

	int sheetcount = worksheets->property("Count").toInt();  //��ȡ��������Ŀ
	int * rowNum = new int[sheetcount];
	int *colNum = new int[sheetcount];
	QTableWidget** table = new QTableWidget*[sheetcount];
	QString* worksheetname = new QString[sheetcount];
	for (int k = 0; k < sheetcount; k++) {
		table[k] = new QTableWidget;

		//��õ�һ��excel���
		QAxObject* worksheet = workbook->querySubObject("Worksheets(int)", k + 1);
		QAxObject* range = worksheet->querySubObject("UsedRange");

		worksheetname[k] = worksheet->property("Name").toString();

		//���excel��������
		QAxObject* rows = range->querySubObject("Rows");
		rowNum[k] = rows->property("Count").toInt();

		QAxObject* columns = range->querySubObject("Columns");
		colNum[k] = columns->property("Count").toInt();

		//��ȡexcel����ʾ�������
		QString txt;
		table[k]->setRowCount(rowNum[k]);
		table[k]->setColumnCount(colNum[k]);
		QVariant cell = range->dynamicCall("Value");
		QVariantList row = cell.value<QVariantList>();
		for (int i = 0; i != row.size(); i++) {
			QVariantList col = row[i].value<QVariantList>();
			for (int j = 0; j != col.size(); j++) {
				txt = col[j].toString();
				QTableWidgetItem* item = new QTableWidgetItem(txt);
				table[k]->setItem(i, j, item);
				//����ӳ�䵽�ṹ����
			}
		}

		//���ɱ༭
		table[k]->setEditTriggers(QAbstractItemView::NoEditTriggers);
		//tabWidget->addTab(table[k], worksheetname[k]);
	}
	//currentsheet = 0;

	//�رղ��˳�
	workbook->dynamicCall("Close(Boolean)", false);
	excel->dynamicCall("Quit(void)");
}

//QAxObject* rows = usedrange->querySubObject("Rows");
//int rownum = rows->property("Count").toInt(); //��ȡ����Columns
//
//QAxObject* cols = usedrange->querySubObject("Columns");
//int columns = cols->property("Count").toInt(); //��ȡ����Columns
//
//QVariant allEnvDataQVariant = usedrange->dynamicCall("Value");//->dynamicCall("Value");
////QVariant allEnvDataQVariant = allEnvData->property("Value");//��ȡ���е�ֵ

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
		dataList << Data(QPointF(i, tableWidget->item(i, column)->text().toDouble()),
			QString::number(column) + ":" + QString::number(i));
	}
	m_dataTable << dataList;
}

void DataVisualization::headerClicked(int i ) {
	printf("headerClicked index : %d", i);
	addData(i);
	QChart* previous = chartView->chart();
	chartView->setChart(createLineChart());
	if(previous)
	    delete previous;
}

//get a work book bind with a excel file.
QAxObject* DataVisualization::getWorkBooks(const QString& excelPath)
{
	QString pathandfilename = excelPath;

	excel = new  QAxObject("Excel.Application");

	excel->setProperty("Visible", false); //���ش򿪵�excel�ļ�����
	QAxObject* workbooks = excel->querySubObject("WorkBooks");//�ɴ򿪶��excel

	return workbooks->querySubObject("Open(QString, QVariant)", pathandfilename); //���ļ�
}

//get number i sheet
QAxObject* DataVisualization::getSheet(QAxObject* workBook, int number)
{
	return workBook->querySubObject("WorkSheets(int)", number); //���ʵ�SHEETNUM��������
}

//do some test
void DataVisualization::standardTest()
{
	QAxObject* workBook = getWorkBooks(excelFilePath);
	QAxObject* workSheet = getSheet(workBook, 1);
	QAxObject* cellA = workSheet->querySubObject("Range(QVariant, QVariant)", "A8");
	cellA->dynamicCall("SetValue(const QVariant&)", QVariant("1+1"));//���õ�Ԫ���ֵ

	workBook->dynamicCall("Close()");
	excel->dynamicCall("Quit()");
}

void DataVisualization::displayData(const QList<QList<QVariant>>& data, const  std::vector<QString>& names)
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

void DataVisualization::displayData(const std::vector<std::vector<QVariant>>& data, int beginRow, int headerRow)
{
	int rows = data.size() - beginRow;
	int columns = data.at(0).size();
	tableWidget->clear();
	tableWidget->setRowCount(rows);
	tableWidget->setColumnCount(columns);

	QStringList headers;
	for (auto head : data[headerRow])
	{
		headers << head.toString();
	}
	tableWidget->setHorizontalHeaderLabels(headers);
	for (int i = beginRow; i < data.size(); i++)
	{
	    std::vector<QVariant>  allEnvDataList_i = data[i];//��i�е�����
		//printf("\n");
		for (int j = 0; j < allEnvDataList_i.size(); j++)
		{
			QString tempvalue = allEnvDataList_i[j].toString();
			//printf("%s, ", qPrintable(tempvalue));
			QTableWidgetItem* item = new QTableWidgetItem(tempvalue);
			tableWidget->setItem(i - beginRow, j, item);
		}
	}
}