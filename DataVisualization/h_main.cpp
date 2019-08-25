#include "DataVisualization.h"
#include "ExcelProcessor.h"
#include <QtWidgets/QApplication>
#include "data/ExcelDataServer.h"
#include <algorithm>
#include <QDir>
const QString filepath = "C:\\Users\\liuyu\\Desktop\\DataVisualization\\x64\\Debug\\nownewtest.xlsx";

int main(int argc, char *argv[])
{
	//QString excelPath;
	//if (argc < 1)
	//{
	//	printf("Please input excel file name.\n", argv[0]);
	//	return 0;
	//}
	//else {
	//	QDir dir;
	//	if (!dir.exists(argv[1]))
	//	{
	//		printf("Can't find file : %s\n", argv[1]);
	//	}
	//	excelPath = dir.absoluteFilePath(*(argv + 1));
	//	//excelPath = dir.filePath(*(argv + 1));
	//}
	QApplication a(argc, argv);
	ExcelDataServer* excelServer = new ExcelDataServer();
	//printf("open file : %s\n", qPrintable(excelPath));
	QAxObject*  worrkbook = excelServer->openExcelFile(filepath);
	if (worrkbook == NULL) 
		printf("open file failed : %s, %p\n", qPrintable(filepath), worrkbook);

	QAxObject*  sheet = excelServer->getSheet(worrkbook, 1);//updata work sheets.
	
	excelServer->setCurrentWorksheet(sheet);

	//add a new sheet
	QAxObject* newSheets = excelServer->addSheet(excelServer->getCurrentWorkSheets(), QString("TestSheet1"));

	excelServer->setAllData(sheet);
	int columsNumber = excelServer->getColumsNumber();
	int rowsNumber = excelServer->getRowsNumber();

	//计算开始、结束行
	excelServer->setBeginEndRow(9, 20);
	QVariantList resultColum3;
	excelServer->getRowData(sheet, 3, resultColum3);//获取第三行的值

	excelServer->appendColums2Sheet(std::vector<QString>{
		QString::fromLocal8Bit(std::string("业态").data()),
		QString::fromLocal8Bit(std::string("城市环线").data())}, newSheets);

	//计算
	//excelServer->calculator(QString::fromLocal8Bit(std::string("18年年期初库存+(业态+营销操盘方)*(测试+股权比例/首开日期)+18年年度供货").data()),
		//QString::fromLocal8Bit(std::string("城市环线").data()));

	//begin arithmetic operation
	//std::vector<double> resultm;
	//std::vector<double> mid{5,1,7,1,9,1,2,2,1,4,4,1};
	//excelServer->arithmeticOperation("H", "I", result, '*');
	//excelServer->arithmeticOperation(mid, "I", result, '/');
	//excelServer->arithmeticOperation("I", mid, resultm, '/');
	//std::vector<double> result;
	//excelServer->arithmeticOperation(resultm, std::vector<double>(resultm.size(), 100), result, '+');
	//excelServer->writeColumData(sheet, "L", result);

	worrkbook->dynamicCall("Save()");
	excelServer->freeExcel();
	a.quit();
	return 0;
}

//QAxObject* cellA22 = sheet->querySubObject("Range(QVariant, QVariant)", "A22");
//QAxObject* cellA23 = sheet->querySubObject("Range(QVariant, QVariant)", "A23");
////cellA->dynamicCall("SetValue(const QVariant&)", QVariant("1+1"));//设置单元格的值
//QVariant resA22 = cellA22->dynamicCall("Value");
////QVariant resA23 = cellA23->property("Value");
////QVariant resA23 = cellA23->property("Value");
//QVariant resA23 = cellA23->dynamicCall("Value");
//
//resA22.isValid() ? printf("A22 isValid\n") : printf("A22 is not Valid\n");
//resA23.isValid() ? printf("A23 isValid\n") : printf("A23 is not Valid\n");
//
//printf("A22 to string : %s, to double : %f. \n", qPrintable(resA22.toString()), resA22.toDouble());
//printf("A23 to string : %s, to double : %f. \n", qPrintable(resA23.toString()), resA23.toDouble());
//resA23.toDouble() == 0 ? printf("is equal.\n") : printf("is not equal.\n");
