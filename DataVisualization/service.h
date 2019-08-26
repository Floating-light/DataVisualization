#pragma once
#include <QString>
#include <QStringList>
#include <QFile>
#include "data/ExcelDataServer.h"
#include <qdebug.h>
using namespace std;

class service
{
public:

	service();
	service(int,int);

	void confirm(ExcelDataServer*);

	QString replace(QStringList & v, int month, int year);
	QString expand(QString v, int month, int year);
	void calCurrent(QString filepath, ExcelDataServer* dataServer);
	void calHistary(QString filepath, ExcelDataServer* dataServer);

	void predictcurrentMonth1(ExcelDataServer* server);
	void predictcurrentMonth2(ExcelDataServer* server);
	void predictcurrentMonth3(ExcelDataServer* server);
	void predictcurrentMonth4(ExcelDataServer* server);

	void yearPredict1(ExcelDataServer* server);
	void yearPredict2(ExcelDataServer* server);
	void yearPredict3(ExcelDataServer* server);
	void yearPredict4(ExcelDataServer* server);
	void yearPredict5(ExcelDataServer* server);
	void yearPredict6(ExcelDataServer* server);
	void calYear(QString filepath, ExcelDataServer* dataServer);

	void setStartRow(int);
	void setEndRow(int);
	int getStartRow();
	int getEndRow();

	~service();

private:
	int startRow;
	int endRow;
};

