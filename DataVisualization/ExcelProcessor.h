#pragma once
#include <iostream>
#include <vector>

#include <QDir>
#include <QVariant>
#include <qdebug.h>
#include <ActiveQt/qaxobject.h>
#include "TestWrite.h"
//const QString excelFilePath1 = "C:/fileds.xlsx";
const QString excelFilePath1 = "C:\\Users\\liuyu\\Documents\\fileds.xlsx";
class ExcelProcessor
{
public:
	ExcelProcessor();
	~ExcelProcessor();
private:
	QAxObject* excelApp;

	QAxObject* excelWorkbooks;
	//excelWorkbooks->dynamicCall("Add");//添加工作簿
	//int count = excelWorkbooks->property("Count").toInt();//获取工作簿数目
	//excelWorkbook = excelWorkbooks->querySubObject("Item(int)",1);//获取第一个的工作簿
	
	QAxObject* currentWorkbook;

	//excelWorksheets = excelWorkbook->querySubObject("Sheets");//获取工作表管理器
	QAxObject* currentWorksheets;
	QAxObject* currentWorksheet;

	void initExcelAppAndWorksheet();

	//add excel file to work books
	inline QAxObject* addExistExcelFile(const QString& filePath);
	inline QAxObject* getSpecifyItem(QAxObject* parent, int number);

	void appendExcelWorksheet(const QString& sheetName);
	void writeCell(QAxObject* currentSheet,int row, int column, const QString& dataStr);//向工作表中写数据
	void saveExcel(const QString& fileName);//保存Excel
	void freeExcel();//释放Excel对象
	
	//uppercase
	void getColum(QAxObject* sheet, QString range, QVariant& data);
	void writeColum(QAxObject* sheet, QString range, const QVariant& data);
	void writeAllData(QAxObject* sheet,const QString& colum, const QList<QList<QVariant>>& res);

	void createColumData(QVariant& variant, const std::vector<QString>& vectorData);

	void castListListVariant2Variant(QVariant& var, const QList<QList<QVariant>>& res);
};

inline QAxObject* ExcelProcessor::addExistExcelFile(const QString& filePath)
{
	return excelWorkbooks->querySubObject("Open(QString, QVariant)", filePath);
}

inline QAxObject* ExcelProcessor::getSpecifyItem(QAxObject* parent, int number)
{
	return parent->querySubObject("Item(int)", 1);
}


