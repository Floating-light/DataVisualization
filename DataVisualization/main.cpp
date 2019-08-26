#include <QtWidgets/QApplication>
#include <QFile>
#include <qdebug.h>
#include "QDate"
#include <string>
#include<vector>
#include<stack>
#include <cstdlib>
#include "data/ExcelDataServer.h"
#include <algorithm>
#include <QDir>
#include "service.h"

#include "DataVisualization.h"
const QString filepath = "C:\\Users\\PIS\\Desktop\\DataVisualization\\DataVisualization\\input\\datasj.xlsx";

using namespace std;

/*

QMap<QString, QStringList> simpleLoad;
QMap<QString, QStringList> complexLoad;
QStringList simpleTable;
QStringList complexTable;

void loadSimpleAndComplex() {
	QFile file(QString::fromLocal8Bit(std::string("C:\\Users\\PIS\\Desktop\\simpleload.txt").data()));
	if (!file.open(QIODevice::ReadOnly | QIODevice::Text))
	{
		qDebug() << "Can't open simple file!";
	}
	while (!file.atEnd())
	{
		QByteArray line = file.readLine();
		QString str(line);
		QStringList list = str.split(" ");
		QString key = list[0];
		list.pop_front();
		simpleLoad.insert(key, list);
	}

	QFile file(QString::fromLocal8Bit(std::string("C:\\Users\\PIS\\Desktop\\complexload.txt").data()));
	if (!file.open(QIODevice::ReadOnly | QIODevice::Text))
	{
		qDebug() << "Can't open simple file!";
	}
	while (!file.atEnd())
	{
		QByteArray line = file.readLine();
		QString str(line);
		QStringList list = str.split(" ");
		QString key = list[0];
		list.pop_front();
		complexLoad.insert(key, list);
	}

	simpleTable = simpleLoad.keys();
	complexTable = complexLoad.keys();
}

void geneSimpleTable(QString v) {
	QStringList rec = simpleLoad.value(v);
	geneSimpleExcel(rec,v);
}


void geneComplexTable(QString v) {
	QStringList rec = complexLoad.value(v);
	QVector<>getExcel(rec);
}
*/


int main(int argc, char *argv[])
{
	QApplication a(argc, argv);
	ExcelDataServer* excelServer = new ExcelDataServer();
	//printf("open file : %s\n", qPrintable(excelPath));
	QAxObject* worrkbook = excelServer->openExcelFile(filepath);
	if (worrkbook == NULL)
		printf("open file failed : %s, %p\n", qPrintable(filepath), worrkbook);
	QAxObject* worksheets = worrkbook->querySubObject("WorkSheets");
	QAxObject * sheet = excelServer->getSheet(worrkbook, 1);//updata work sheets.
	//QAxObject* sheet = excelServer->getNamedSheet(worksheets,
		//QString::fromLocal8Bit(std::string("测试数据").data()));
	excelServer->setCurrentWorksheet(sheet);

	//add a new sheet
	//QAxObject * newSheets = excelServer->addSheet(excelServer->getCurrentWorkSheets(), QString("selectedTest"));
	QAxObject* usedrange = sheet->querySubObject("UsedRange");
	excelServer->setAllData(usedrange);
	int columsNumber = excelServer->getColumsNumber();
	int rowsNumber = excelServer->getRowsNumber();

	int startrow = 4;
	int endrow = rowsNumber;
	//计算开始、结束行
	excelServer->setBeginEndRow(startrow , endrow);

	QVariantList resultColum3;
	excelServer->getRowData(sheet, 3, resultColum3);//获取第三行的值

	/*QAxObject* cellA23 = sheet->querySubObject("Range(QVariant, QVariant)", "A23");
	QVariant resA23 = cellA23->property("Value");
	printf("data : %s\n", qPrintable(resA23.toString()));*/
	/******************************************************************/
	/*QList<QList<QVariant>> result;
	std::vector<QString> selected{
		QString::fromLocal8Bit(std::string("营销操盘方").data()),
			QString::fromLocal8Bit(std::string("首开日期").data()),
			QString::fromLocal8Bit(std::string("城市环线").data()) };

	excelServer->selectWhere(selected,
		QString::fromLocal8Bit(std::string("股权比例").data()),
		QString::fromLocal8Bit(std::string("13").data()), result);*/
	/*result.push_front(QList<QVariant>{
		QString::fromLocal8Bit(std::string("营销操盘方").data()),
		QString::fromLocal8Bit(std::string("首开日期").data()),
		QString::fromLocal8Bit(std::string("城市环线").data())});*/

	/*DataVisualization widget;
	widget.displayData(excelServer->sheetContent, 3, 2);
	widget.show();*/

	//excelServer->writeArea(newSheets, result);
	/******************************************************************/
	service proService(startrow, endrow);
	proService.confirm(excelServer);
	//write all data
	QVariant var;
	excelServer->castSheetVector2Variant(var);
	usedrange->setProperty("Value", var);

	worrkbook->dynamicCall("Save()");
	excelServer->freeExcel();
	a.exit();
	return 0;
}






