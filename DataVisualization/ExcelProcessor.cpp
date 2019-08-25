#include "ExcelProcessor.h"



ExcelProcessor::ExcelProcessor()
{
	initExcelAppAndWorksheet();
	currentWorkbook = addExistExcelFile(excelFilePath1);

	//currentWorksheet = currentWorkbook->querySubObject("WorkSheets(int)", 1); //���ʵ�SHEETNUM��������
	currentWorksheets = currentWorkbook->querySubObject("Sheets");
	currentWorksheet = currentWorksheets->querySubObject("Item(int)", 1);
	
	QAxObject* usedrange = currentWorksheet->querySubObject("UsedRange");
	QAxObject* rows = usedrange->querySubObject("Rows");
	int rownum = rows->property("Count").toInt(); //��ȡ����
	//printf("the row number : %d\n", rownum);
	//QVariant var;
	//createColumData(var, std::vector<QString>{"Count1", "Count2", "Count3", "Count4", "Count5", "Count6", "Count7", "Count8"});
	//castListListVariant2Variant(var, QList<QList<QVariant>>{ {"Count1"}, { "Count2" }, { "Count3" }, { "Count4" }, { "Count5" }, { "Count6" }, { "Count7" }, { "Count8" } });
	//writeColum(currentWorksheet, "E3:E10",var);
	writeAllData(currentWorksheet,"Z", QList<QList<QVariant>>{ {"Count1"}, { "Count2" },
		{ "Count3" }, { "Count4" }, { "Count5" }, { "Count6" }, { "Count7" }, { "Count8" },
		{ "Count3" }, { "Count4" }, { "Count5" }, { "Count6" }, { "Count7" }, { "Count24" },
		{ "Count25" }, { "Count26" }, { "Count27" }, { "Count28" }, { "Count29" }, { "Count30" }
	});
	//TestRecordExporter* test = new TestRecordExporter();
	//test->writeToSheet(QList<QList<QVariant>>{ {"Count1"}, { "Count2" }, { "Count3" }, { "Count4" }, { "Count5" }, { "Count6" }, { "Count7" }, { "Count8" } }, currentWorksheet);


	//QAxObject* allEnvData = currentWorksheet->querySubObject("Range(QString)", "A1:A" + QString::number(rownum));
	//QVariant allEnvDataQVariant;
	//getColum(currentWorksheet, "A", allEnvDataQVariant);
	//QVariantList listVariant = allEnvDataQVariant.toList();

	
	//��i��
	//for (int i = 0; i < listVariant.size(); ++i)
	//{
	//	QVariantList lastList= listVariant[i].toList();
	//	for (int j = 0; j < lastList.size(); ++j)
	//	{
	//		//printf("(%d, %d)--->%s\n", i, j, qPrintable(lastList[j].toString()));
	//		QString temp = lastList[j].toString();
	//		//printf("(%d, %d)--->%s\n", i, j, qPrintable(temp));
	//		qDebug() << qPrintable(temp);
	//	}
	//}
	

	currentWorkbook->dynamicCall("Save()");
	freeExcel();
}


ExcelProcessor::~ExcelProcessor()
{
}

void ExcelProcessor::initExcelAppAndWorksheet()
{
	excelApp = new QAxObject("Excel.Application");
	excelApp->setProperty("Visible", false); //���ش򿪵�excel�ļ�����
	excelApp->setProperty("DisplayAlerts", false);//����ʾ����
	excelWorkbooks = excelApp->querySubObject("WorkBooks");//�ɴ򿪶��excel
	
}

//uppercase ,successful
void ExcelProcessor::getColum(QAxObject* sheet, QString colum, QVariant& data)
{
	QAxObject* usedrange = sheet->querySubObject("UsedRange");
	QAxObject* rows = usedrange->querySubObject("Rows");
	int rownum = rows->property("Count").toInt(); //��ȡ����
	//printf("the row number : %d\n", rownum);


	QAxObject* allEnvData = currentWorksheet->querySubObject("Range(QString)", colum+"1:"+colum + QString::number(rownum));
	data = allEnvData->property("Value");
}

void ExcelProcessor::writeColum(QAxObject* sheet, QString range, const QVariant& data)
{
	QAxObject* allEnvData = sheet->querySubObject("Range(QString)", range);
	allEnvData->setProperty("Value", data);
}

void ExcelProcessor::writeAllData(QAxObject* sheet, const QString& colum, const QList<QList<QVariant>>& res)
{
	int row = res.size();
	int col = res.at(0).size();
	QVariant var;
	castListListVariant2Variant(var, res);
	writeColum(currentWorksheet, colum + "9:"+colum+QString::number(row + 8), var);
}

void ExcelProcessor::createColumData(QVariant& variant, const std::vector<QString>& vectorData)
{
	QVariantList vList;
	for (int i = 0; i < vectorData.size(); ++i)
	{
		vList.append(QVariantList{ QVariant(vectorData[i]) });
	}
	variant = QVariant(vList);

}

void ExcelProcessor::castListListVariant2Variant(QVariant& var, const QList<QList<QVariant>>& res)
{
	QVariant temp = QVariant(QVariantList());
	QVariantList record;

	int listSize = res.size();
	for (int i = 0; i < listSize; ++i)
	{
		temp = res.at(i);
		record << temp;
	}
	temp = record;
	var = temp;
}


//QAxObject* allEnvData = worksheet->querySubObject("Range(QString)", Range); //��ȡ��Χ

void ExcelProcessor::appendExcelWorksheet(const QString& sheetName)
{

}

//��������д����
void ExcelProcessor::writeCell(QAxObject* currentSheet, int row, int column, const QString& dataStr)
{
	QAxObject* range = currentSheet->querySubObject("Cells(int,int)", row, column);
	range->dynamicCall("Value", dataStr);
}
/*for (int i = 0; i < 10; ++i)
{
	QAxObject* cellA = currentWorksheet->querySubObject("Range(QVariant, QVariant)", "B" + QString::number(i + 10));
	cellA->dynamicCall("SetValue(const QVariant&)", QVariant(i));//���õ�Ԫ���ֵ
}*/

//����Excel
void ExcelProcessor::saveExcel(const QString& fileName)
{

}



//�ͷ�Excel����
void ExcelProcessor::freeExcel()
{
	excelWorkbooks->dynamicCall("Close()");
	excelApp->dynamicCall("Quit()");
}




