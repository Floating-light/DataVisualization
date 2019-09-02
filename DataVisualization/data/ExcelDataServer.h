#pragma once
#include <vector>
#include <limits>
#include <map>
#include "data/DataServer.h"
#include <QVector>
#include <QChar>
#include <QMessageBox>

enum SummaryType
{
	Grounp,
	BU, //business unit
	Region,
	Company
};

const std::map<int, QColor> color{
	{-1, QColor(196, 215, 155)},
	{0, QColor(83, 141, 213)},
	{1, QColor(141, 180, 226)},
	{2, QColor(220, 230, 241)},
};

class ExcelDataServer : public DataServer
{
public:
	ExcelDataServer();
	~ExcelDataServer();

	//first : rows number, second : color type;
	std::map<int, int> colorRows;

	QVariant allData;
	std::vector<std::vector<QVariant>> sheetContent;
	void setAllData(QAxObject* sheet);
	std::map<QString, QString> nameToColum;
	std::map<QString, int> nameToSubScript;
	QString beginRow = "9";
	QString endRow = "16";

	QAxObject* getNamedSheet(QAxObject* sheets, const QString& name);

	//arithmetic operation = - * /
	//std::vector<double>(length, num);
	inline void arithmeticOperation(const QString& colum1, const QString& colum2,
		std::vector<double>& result, char oper);
	inline void arithmeticOperation(const std::vector<double>& colum1, const QString& colum2,
		std::vector<double>& result, char oper);
	inline void arithmeticOperation(const QString& colum1, const std::vector<double>& colum2,
		std::vector<double>& result, char oper);
	inline void arithmeticOperation(const std::vector<double>& colum1, const std::vector<double>& colum2,
		std::vector<double>& result, char oper);

	void operation(const std::vector<double>& colum1, const std::vector<double>& colum2,
		std::vector<double>& result,  const char oper);

	inline void ExcelDataServer::arithmeticOperation(const QString& colum1, double colum2,
		std::vector<double>& result, char oper);
	inline void ExcelDataServer::arithmeticOperation(double colum2, const QString& colum1,
		std::vector<double>& result, char oper);

	//
	inline void setCurrentWorksheet(QAxObject* sheet);
	
	//设置计算的开始与结束行
	inline void setBeginEndRow(int begin, int end);
	//return opened excel work book 
	QAxObject* openExcelFile(const QString& fileName);
	//get the number i worksheet
	QAxObject* getSheet(QAxObject* workbook, int i);
	void getColum(QAxObject* sheet, QString colum, QVariant& data);
	void getColum(QAxObject* sheet, QString colum, std::vector<double>& data);
	void getColum(QAxObject* sheet, int column, std::vector<double>& data);

	void writeArea(QAxObject* tragetSheet, const QList<QList<QVariant>>& res);
	void writeAllData(QAxObject* sheet, const QString& colum, const QList<QList<QVariant>>& res);
	void writeColumData(QAxObject* sheet, const QString& colum, const std::vector<double>& res);
	void writeColumData(const QString& colum, const std::vector<double>& res);
	void writeColumData(int column, const std::vector<double>& res);
	void freeExcel();
	int getRowsNumber();
	int getColumsNumber();
	//获取取某行所有数据
	bool getRowData(QAxObject* sheet, int rowNumber, QVariantList& result);

	//列计算
	void calculator(const QString& expression, const QString& output);

	//创建一个新的sheet,并返回
	QAxObject* addSheet(QAxObject* workSheets,const QString& name );
	inline QAxObject* getCurrentWorkSheets();

	//将字母列号转为数字下标
	int alphabet2Int(const QString& alp);
	void int2Alphabet(int number, QString& alphabet);
	//void wirte2Sheet(const std::vector<QString>& headNames, QAxObject* sheet);
	void appendColums2Sheet(const std::vector<QString>& headNames, QAxObject* sheet);
	//void appendRows2Sheet(const std::vector<QString>& headNames, QAxObject* sheet);

	QVariant getCellData(const QString& name, int row);
	void writedata(QString data, QString c, int r);
	void writedata(int data, QString c, int r);
	void writedata(QVariant data, QString c, int r);

	void selectWhere(const std::vector<QString>& selectedName,
		const QString& whereName, const QString& whereValue, QList<QList<QVariant>>& result);

	void castSheetVector2Variant(QVariant& var);

	void exportSheet(const QList<QList<QVariant>>& exportData,  const QString& sheetName);

	void templateExport(const QString& templatePath,  int headerRow);
	void getColumnSpecifyData(const QVariantList& exportHeader, QList<QList<QVariant>>& exportData);
	QList<QVariant> getInsertRow(const std::vector<QVariant>& cache, int changedIndex, int columnNumber);
	std::vector<double> excuteSummation(const QList<QList<QVariant>>& exportData,const std::vector<int>& sumColumn,
		int changedHeader, const QString& name);
	std::vector<double> mergeSummation(const QList<QList<QVariant>>& exportData,
		const std::vector<int>& sumColumn, int changedHeader,const QString& name);
	bool isPureDigit(const QString& str);
	void sumSkipColumn(const std::vector<QVariant>&checkColumn, const std::vector<int>& exportIndexs, std::vector<int>& sumColumn);

	void setRowColor(QAxObject* sheet, int columns);
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

	int checkProority(QChar c1, QChar c2);

	void initExcelAppAndWorksheet();

	void writeColum(QAxObject* sheet, QString range, const QVariant& data);
	void castListListVariant2Variant(QVariant& var, const QList<QList<QVariant>>& res);
	void castDoubleVector2Variant(QVariant& var, const std::vector<double>& res);
	
};

inline void ExcelDataServer::arithmeticOperation(const std::vector<double>& colum1,
	const std::vector<double>& colum2, std::vector<double>& result, char oper)
{
	operation(colum1, colum2,result, oper);
}

inline void ExcelDataServer::arithmeticOperation(const QString& colum1, const QString& colum2,
	std::vector<double>& result, char oper)
{
	std::vector<double> columVector1;
	getColum(currentWorksheet, colum1 + beginRow +":" + colum1 + endRow, columVector1);

	std::vector<double> columVector2;
	getColum(currentWorksheet, colum2 + beginRow + ":" + colum2 + endRow, columVector2);
	operation(columVector1, columVector2, result, oper);
}

inline void ExcelDataServer::arithmeticOperation(const QString& colum1, const std::vector<double>& colum2,
	std::vector<double>& result, char oper)
{
	std::vector<double> columVector1;
	getColum(currentWorksheet, colum1 + beginRow + ":" + colum1 + endRow, columVector1);
	operation(columVector1, colum2, result, oper);
}

inline void ExcelDataServer::arithmeticOperation(const std::vector<double>& colum1,
	const QString& colum2,std::vector<double>& result, char oper)
{
	std::vector<double> columVector2;
	getColum(currentWorksheet, colum2 + beginRow + ":" + colum2 + endRow, columVector2);
	operation(colum1, columVector2, result, oper);
}

inline void ExcelDataServer::arithmeticOperation(const QString& colum1, double colum2,
	std::vector<double>& result, char oper)
{
	std::vector<double> columVector1;
	getColum(currentWorksheet, colum1 + beginRow + ":" + colum1 + endRow, columVector1);
	operation(columVector1, std::vector<double>(columVector1.size(), colum2), result, oper);
}

inline void ExcelDataServer::arithmeticOperation(double colum1, const QString& colum2,
	std::vector<double>& result, char oper)
{
	std::vector<double> columVector2;
	getColum(currentWorksheet, colum2 + beginRow + ":" + colum2 + endRow, columVector2);
	operation(std::vector<double>(columVector2.size(), colum1), columVector2, result, oper);
}

inline void ExcelDataServer::setCurrentWorksheet(QAxObject* sheet)
{
	currentWorksheet = sheet;
}

inline void ExcelDataServer::setBeginEndRow(int begin, int end)
{
	beginRow = QString::number(begin);
	endRow = QString::number(end);
}

inline QAxObject* ExcelDataServer::getCurrentWorkSheets()
{
	return currentWorksheets;
}