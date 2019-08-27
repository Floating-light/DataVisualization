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
const QString filepath = "C:\\source\\DataVisualization\\DataVisualization\\input\\new1-8.xlsx";

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
	//ExcelDataServer* excelServer = new ExcelDataServer();
	////printf("open file : %s\n", qPrintable(excelPath));
	//printf("Open file : %s\n", qPrintable(filepath));
	//QAxObject* worrkbook = excelServer->openExcelFile(filepath);
	//if (worrkbook == NULL)
	//	printf("open file failed : %s, %p\n", qPrintable(filepath), worrkbook);
	//else
	//{
	//	printf("open success\n");
	//}
	//QAxObject* worksheets = worrkbook->querySubObject("WorkSheets");
	//QAxObject * sheet = excelServer->getSheet(worrkbook, 1);//updata work sheets.
	////QAxObject* sheet = excelServer->getNamedSheet(worksheets,
	//	//QString::fromLocal8Bit(std::string("��������").data()));
	//excelServer->setCurrentWorksheet(sheet);

	////add a new sheet
	////QAxObject * newSheets = excelServer->addSheet(excelServer->getCurrentWorkSheets(), QString("selectedTest"));
	//QAxObject* usedrange = sheet->querySubObject("UsedRange");
	//excelServer->setAllData(usedrange);
	//int columsNumber = excelServer->getColumsNumber();
	//int rowsNumber = excelServer->getRowsNumber();

	//int startrow = 4;
	//int endrow = rowsNumber;
	////���㿪ʼ��������
	//excelServer->setBeginEndRow(startrow , endrow);

	//QVariantList resultColum3;
	//excelServer->getRowData(sheet, 3, resultColum3);//��ȡ�����е�ֵ

	DataVisualization widget;
	//widget.displayData(excelServer->sheetContent, 3, 2);
	widget.show();

	/******************************************************************/
	//core calculate 
	/*calHistary(QString::fromLocal8Bit(std::string("./input/1��ʷ��-����.txt").data()), excelServer);
	
	calCurrent(QString::fromLocal8Bit(std::string("./input/2��ǰ��-����.txt").data()), excelServer);
	
	calYear(QString::fromLocal8Bit(std::string("./input/3���?����.txt").data()), excelServer);

	calHistary(QString::fromLocal8Bit(std::string("./input/4��ʷ��-���?txt").data()), excelServer);

	calCurrent(QString::fromLocal8Bit(std::string("./input/5��ǰ��-���?txt").data()), excelServer);

	//6���ӹ�ʽ-���?txt
	//���ǩԼ��ɣ���ҵ���棩
	yearPredict3(excelServer);
	//���ȫ�ھ�������ɣ���ҵ���棩
	yearPredict1(excelServer);
	//���Ȩ��ھ�������ɣ���ҵ����?	yearPredict5(excelServer);
	//�ۺ�ȫ�ھ�������
	predictcurrentMonth3(excelServer);
	//�ۺ�Ȩ��ھ�������?	predictcurrentMonth4(excelServer);

	//7���ӹ�ʽ - Ԥ�м���.txt
	//��Ŀ����
	predictcurrentMonth1(excelServer);
	//Ԥ��
	predictcurrentMonth2(excelServer);

	calCurrent(QString::fromLocal8Bit(std::string("./input/8��ǰ��-����ͻ���Ԥ��?txt").data()), excelServer);
	calCurrent(QString::fromLocal8Bit(std::string("./input/9��ǰ��-����.txt").data()), excelServer);
	
	//���ǩԼ��ɣ����ư棩
	yearPredict4(excelServer);
	//���ȫ�ھ�������ɣ����ư棩
	yearPredict2(excelServer);
	//���Ȩ��ھ�������ɣ����ư�?	yearPredict6(excelServer);

	calYear(QString::fromLocal8Bit(std::string("./input/11���?����.txt").data()), excelServer);*/
	
	//*************************************************************************************************



	//service proService(startrow, endrow);
	//proService.confirm(excelServer);

	//service proService(startrow, endrow);
	//QMap<QString, QStringList> report = proService.getReport();
	//proService.confirm(excelServer);


	//write all data
	/*QVariant var;
	excelServer->castSheetVector2Variant(var);
	usedrange->setProperty("Value", var);

	worrkbook->dynamicCall("Save()");
	excelServer->freeExcel();*/
	
	return a.exec();
}






