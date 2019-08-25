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
	//excelWorkbooks->dynamicCall("Add");//��ӹ�����
	//int count = excelWorkbooks->property("Count").toInt();//��ȡ��������Ŀ
	//excelWorkbook = excelWorkbooks->querySubObject("Item(int)",1);//��ȡ��һ���Ĺ�����
	
	QAxObject* currentWorkbook;

	//excelWorksheets = excelWorkbook->querySubObject("Sheets");//��ȡ�����������
	QAxObject* currentWorksheets;
	QAxObject* currentWorksheet;

	void initExcelAppAndWorksheet();

	//add excel file to work books
	inline QAxObject* addExistExcelFile(const QString& filePath);
	inline QAxObject* getSpecifyItem(QAxObject* parent, int number);

	void appendExcelWorksheet(const QString& sheetName);
	void writeCell(QAxObject* currentSheet,int row, int column, const QString& dataStr);//��������д����
	void saveExcel(const QString& fileName);//����Excel
	void freeExcel();//�ͷ�Excel����
	
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


