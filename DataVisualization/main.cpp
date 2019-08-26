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

#include "DataVisualization.h"
const QString filepath = "C:\\source\\DataVisualization\\DataVisualization\\input\\new1-8.xlsx";

using namespace std;

QDate D1 = QDate::currentDate();
int startYear = 19;
int currentYear = D1.year() - 2000;//获取年
int currentMonth = D1.month();
const QString constYear = QString::fromLocal8Bit(std::string("年").data());
const QString constMonth = QString::fromLocal8Bit(std::string("月").data());

static int startrow = 9;
static int endrow = 1000;

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

QString replace(QStringList & v, int month, int year) {
	for (int i = 0; i < v.size(); i++)
	{
		if (v[i].startsWith("(") || v[i].startsWith(")")) {
			;
		}
		else if (!v[i][0].isDigit() && !v[i].startsWith("/") && !v[i].startsWith("*") && !v[i].startsWith("+") && !v[i].startsWith("-")) {
			if(v[i].startsWith("!"))
			{
				v[i] = v[i].mid(1, -1);
			}
			else if (v[i].startsWith("@")) {
				v[i] = QString::number(year) + constYear + v[i].mid(1,-1);
			}
			else if (v[i].startsWith("#")) {
				if (month == 1 && year == startYear) {
					return "";
				}
				else {
					v[i] = QString::number(year) + constYear + QString::number(month-1) + constMonth + v[i].mid(1, -1);
				}
			}
			else if (v[i].startsWith("$")) {
				if (month == 1) {
					v[i] = QString::number(year) + constYear + QString::number(month) + constMonth + v[i].mid(1, -1);
				}
				else {
					QString temp = "(";
					for (int tempi = 1; tempi <= month; tempi++) {
						temp.append(QString::number(year) + constYear + QString::number(tempi) + constMonth + v[i].mid(1, -1) + "+");
					}
					temp.chop(1);
					v[i] = temp + ")";
				}	
			}
			else if(v[i].startsWith("%")){
				QString temp = "(";
				for (int tempi = month-2; tempi <= month; tempi++) {
					temp.append(QString::number(year) + constYear + QString::number(i) + constMonth + v[i].mid(1, -1) + "+");
				}
				temp.chop(1);
				v[i] = temp + ")/3";
			}
			else {
				v[i] = QString::number(year) + constYear + QString::number(month) + constMonth + v[i];
			}
		}
	}
	return v.join("");
}

QString expand(QString v,int month,int year) {
	QStringList list = v.split(" ");
	return replace(list, month,year);
}

void calCurrent(QString filepath, ExcelDataServer* dataServer) {
	QFile file(filepath);
	if (!file.open(QIODevice::ReadOnly | QIODevice::Text))
	{
		qDebug() << "Can't open the file!";
	}
	while (!file.atEnd())
	{
		QByteArray line = file.readLine();
		QString str(line);
		//printf("%s", qPrintable(str));

		QStringList list = str.split("=");
		QString rec = list[1].trimmed();
		QString label = list[0].trimmed();
		QString longstring = expand(rec, currentMonth, currentYear);
		if (!longstring.isEmpty()){
			printf("Process ... ...\n");
			printf("\t%s\t", qPrintable(longstring));	
			printf("%s\n", qPrintable(QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + label));
			dataServer->calculator(longstring, QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + label);
			printf("Process ... ... done\n");
		}
	}
}

void calHistary(QString filepath, ExcelDataServer* dataServer){
	QFile file(filepath);
	if (!file.open(QIODevice::ReadOnly | QIODevice::Text))
	{
		qDebug() << "Can't open the file!";
	}
	while (!file.atEnd())
	{
		QByteArray line = file.readLine();
		QString str(line);
		//printf("%s", qPrintable(str));

		QStringList list = str.split("=");
		QString rec = list[1].trimmed();
		QString label = list[0].trimmed();
		for (int i = 1; i <= 7; i++) {
			QString longstring = expand(rec, i, currentYear);
			if (!longstring.isEmpty()) {
				printf("Process ... ...\n");
				printf("\t%s\t", qPrintable(longstring));
				printf("%s\n", qPrintable(QString::number(currentYear) + constYear + QString::number(i) + constMonth + label));
				dataServer->calculator(longstring, QString::number(currentYear) + constYear + QString::number(i) + constMonth + label );
				printf("Process ... ... done\n");
			}
		}
	}
}

void predictcurrentMonth1(ExcelDataServer* server) {
	int count_firstOpen = 0;
	int count_pulus = 0;
	int count_continue = 0;
	int count_null = 0;
	for (int i = startrow; i < endrow; i++) {
		QString ndqydc = server->getCellData(QString::number(currentYear) + constYear
			+ QString::fromLocal8Bit(std::string("年度签约达成").data()), i).toString();

		QString dyqyydc = server->getCellData(QString::number(currentYear) + constYear
			+ QString::number(currentMonth) + constMonth + 
			QString::fromLocal8Bit(std::string("签约已达成").data()), i).toString();

		QString dyqckc = server->getCellData(QString::number(currentYear) + constYear + 
			QString::number(currentMonth) + constMonth +
			QString::fromLocal8Bit(std::string("期初库存").data()), i).toString();

		QString dygh = server->getCellData(QString::number(currentYear) + constYear +
			QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("供货").data()), i).toString();
		if (ndqydc.toFloat() - dyqyydc.toFloat() == 0 && dyqckc.toFloat() == 0 && dygh.toFloat() > 0)
		{
			++count_firstOpen;
			server->writedata(QVariant(QString::fromLocal8Bit(std::string("首开类").data())),
				QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth +
				QString::fromLocal8Bit(std::string("项目类型").data()), i);
		}
		else if (dyqckc.toFloat() > 0 && dygh.toFloat() > 0) {
			++count_pulus;
			server->writedata(QVariant(QString::fromLocal8Bit(std::string("加推类").data())),
				QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + 
				QString::fromLocal8Bit(std::string("项目类型").data()), i);
		}
		else if (ndqydc.toFloat() - dyqyydc.toFloat() > 0 || dyqckc.toFloat() > 0 && dygh.toFloat() == 0) {
			++count_continue;
			server->writedata(QVariant(QString::fromLocal8Bit(std::string("续销类").data())),
				QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth +
				QString::fromLocal8Bit(std::string("项目类型").data()), i);
		}
		else
		{
			++count_null;
		}
	}
	printf("--------->>first open： %d, plus push：%d, continue sell： %d, null： %d\n", count_firstOpen, count_pulus, count_continue, count_null);
}
//八月预判
void predictcurrentMonth2(ExcelDataServer* server) {
	for (int i = startrow; i < endrow; i++) {
		QString xmlx = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
			+ constMonth + QString::fromLocal8Bit(std::string("项目类型").data()), i).toString();
		if (xmlx.compare(QString::fromLocal8Bit(std::string("首开类").data())) == 0) {
			QString syb = server->getCellData(QString::fromLocal8Bit(std::string("事业部（住开/商开）").data()), i).toString();
			if (syb.compare(QString::fromLocal8Bit(std::string("商开").data())) == 0) {
				QString yt = server->getCellData(QString::fromLocal8Bit(std::string("业态").data()), i).toString();
				if (yt.compare(QString::fromLocal8Bit(std::string("住宅").data())) == 0) {
					int hz = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth) 
						+ constMonth + QString::fromLocal8Bit(std::string("可售").data()), i).toInt();
					int result = hz * 0.7;
					server->writedata(QVariant(result < 50000 ? result : 50000), QString::number(currentYear) + constYear
						+ QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), i);
				}
				else if (yt.compare(QString::fromLocal8Bit(std::string("商业").data())) == 0) {
					int hz = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth) 
						+ constMonth + QString::fromLocal8Bit(std::string("可售").data()), i).toInt();
					int result = hz * 0.65;
					server->writedata(QVariant(result < 50000 ? result : 50000), QString::number(currentYear) + constYear
						+ QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), i);
				}
				else if (yt.compare(QString::fromLocal8Bit(std::string("公寓/办公").data())) == 0) {
					int hz = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
						+ constMonth + QString::fromLocal8Bit(std::string("可售").data()), i).toInt();
					int result = hz * 0.35;
					server->writedata(QVariant(result < 50000 ? result : 50000), QString::number(currentYear) + constYear
						+ QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), i);
				}
			}
			else if (syb.compare(QString::fromLocal8Bit(std::string("住开").data())) == 0) {
				QString cshx = server->getCellData(QString::fromLocal8Bit(std::string("城市环线").data()), i).toString();
				if (cshx.compare(QString::fromLocal8Bit(std::string("一线").data())) == 0
					|| cshx.compare(QString::fromLocal8Bit(std::string("二线").data())) == 0
					|| cshx.compare(QString::fromLocal8Bit(std::string("三线").data())) == 0)
				{
					int hz = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth) 
						+ constMonth + QString::fromLocal8Bit(std::string("可售").data()), i).toInt();
					int result = hz * 0.6;
					server->writedata(QVariant(result < 50000 ? result : 50000), QString::number(currentYear) + constYear
						+ QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), i);
				}
				else if (cshx.compare(QString::fromLocal8Bit(std::string("四线").data())) == 0 
				     || cshx.compare(QString::fromLocal8Bit(std::string("五线").data())) == 0)
				{
					int hz = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
						+ constMonth + QString::fromLocal8Bit(std::string("可售").data()), i).toInt();
					int result = hz * 0.35;
					server->writedata(QVariant(result < 12000 ? result : 12000), QString::number(currentYear) + constYear
						+ QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), i);
				}
			}
		}
		else if (xmlx.compare(QString::fromLocal8Bit(std::string("续销类").data())) == 0) {
			QString cshx = server->getCellData(QString::fromLocal8Bit(std::string("首开日期").data()), i).toString();
			/*if (cshx == "" || cshx == "/")
				continue;*/
			int fmonth = 0;
		    int fyear = 0;
			QStringList ymd = cshx.split("T");
			if (ymd.size() == 1)
			{
				fmonth = 12;
				fyear = 18;
			}
			else
			{
				QString date = ymd[0];
				QStringList alist = date.split("-");

				fmonth = alist[1].toInt();
				fyear = alist[0].toInt() - 2000;

			}
			if (currentMonth + (currentYear - fyear) * 12 - fmonth <= 3) {
				double radio = 0;
				QString syb = server->getCellData(QString::fromLocal8Bit(std::string("事业部（住开/商开）").data()), i).toString();
				if (syb.compare(QString::fromLocal8Bit(std::string("住开").data())) == 0)
					radio = 0.5;
				else if (syb.compare(QString::fromLocal8Bit(std::string("商开").data())) == 0) {
					QString yt = server->getCellData(QString::fromLocal8Bit(std::string("业态").data()), i).toString();
					if (yt.compare(QString::fromLocal8Bit(std::string("住宅").data()))) {
						radio = 0.6;
					}
					else if (yt.compare(QString::fromLocal8Bit(std::string("商铺").data()))) {
						radio = 0.22;
					}
					else if (yt.compare(QString::fromLocal8Bit(std::string("公寓/办公").data()))) {
						radio = 0.5;
					}
				}

				int ret1 = server->getCellData(QString::number(fyear) + constYear + QString::number(fmonth) + constMonth
					+ QString::fromLocal8Bit(std::string("签约已达成").data()), i).toInt();
				int ret2 = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth) 
					+ constMonth + QString::fromLocal8Bit(std::string("可售").data()), i).toInt();
				int ret3 = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
					+ constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), i).toInt();
				int retf = int(max(ret3, int(min(ret1 * radio, ret2 * 0.8))));
				server->writedata(QVariant(retf), QString::number(currentYear) + constYear + QString::number(currentMonth)
					+ constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), i);
			}
			else {
				int count = 3;
				int remain = 3;
				int output;
				int dyear = currentYear;
				int dmonth = currentMonth;
				while (remain > 0) {
					if (dmonth > 1) {
						dmonth--;
						int ret1 = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
							+ constMonth + QString::fromLocal8Bit(std::string("供货").data()), i).toInt();
						if (ret1 == 0)
						{
							int ret2 = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
								+ constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), i).toInt();
							if (ret2 == 0) {
								count--;
							}
							else {
								output += ret2;
							}
							remain -= 1;
						}
					}
					else {
						dmonth = 12;
						dyear -= 1;
						int ret1 = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth) 
							+ constMonth + QString::fromLocal8Bit(std::string("供货").data()), i).toInt();
						if (ret1 == 0)
						{
							int ret2 = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
								+ constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), i).toInt();
							if (ret2 == 0) {
								count--;
							}
							else {
								output += ret2;
							}
							remain -= 1;
						}
						

					}
				}
				int dyhz = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth) + constMonth
					+ QString::fromLocal8Bit(std::string("供货").data()), i).toInt();
				int dyqyydc = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth) + constMonth 
					+ QString::fromLocal8Bit(std::string("签约已达成").data()), i).toInt();
				if (count == 0)
				server->writedata(QVariant(dyqyydc), QString::number(currentYear) + constYear + QString::number(currentMonth)
					+ constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), i);
				else
				server->writedata(QVariant(max(dyqyydc, min(dyhz, int(output / count)))), QString::number(currentYear) + constYear
					+ QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), i);
			}
		}
		else if (xmlx.compare(QString::fromLocal8Bit(std::string("加推类").data())) == 0) {
			QString cshx = server->getCellData(QString::fromLocal8Bit(std::string("首开日期").data()), i).toString();
			/*if (cshx == "")
				continue;*/
			int fmonth = 0;
			int fyear = 0;
			QStringList ymd = cshx.split("T");
			if (ymd.size() == 1)
			{
				fmonth = 12;
				fyear = 18;
			}
			else
			{
				QString date = ymd[0];
				QStringList alist = date.split("-");

				fmonth = alist[1].toInt();
				fyear = alist[0].toInt() - 2000;

			}
			if (currentMonth + (currentYear - fyear) * 12 - fmonth <= 3) {
				double radio = 0;
				QString syb = server->getCellData(QString::fromLocal8Bit(std::string("事业部（住开/商开）").data()), i).toString();
				if (syb.compare(QString::fromLocal8Bit(std::string("住开").data())) == 0)
					radio = 0.5;
				else if (syb.compare(QString::fromLocal8Bit(std::string("商开").data())) == 0) {
					QString yt = server->getCellData(QString::fromLocal8Bit(std::string("业态").data()), i).toString();
					if (yt.compare(QString::fromLocal8Bit(std::string("住宅").data()))) {
						radio = 0.6;
					}
					else if (yt.compare(QString::fromLocal8Bit(std::string("商铺").data()))) {
						radio = 0.22;
					}
					else if (yt.compare(QString::fromLocal8Bit(std::string("公寓/办公").data()))) {
						radio = 0.5;
					}
				}

				int ret1 = server->getCellData(QString::number(fyear) + constYear + QString::number(fmonth) + constMonth
					+ QString::fromLocal8Bit(std::string("签约已达成").data()), i).toInt();
				int ret2 = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
					+ constMonth + QString::fromLocal8Bit(std::string("可售").data()), i).toInt();
				int ret3 = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
					+ constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), i).toInt();
				int retf = int(max(ret3, int(min(ret1 * radio, ret2 * 0.8))));
				server->writedata(QVariant(retf), QString::number(currentYear) + constYear + QString::number(currentMonth) +
					constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), i);
			}
			else {
				int diff = 3;
				int output;
				int dyear = currentYear;
				int dmonth = currentMonth;
				bool isok = false;
				while (diff > 0) {
					if (dmonth > 1) {
						dmonth--;
						int ret1 = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
							+ constMonth + QString::fromLocal8Bit(std::string("供货").data()), i).toInt();
						if (ret1 != 0) {
							isok = 1;
							break;
						}
						diff--;
					}
					else {
						dmonth = 12;
						dyear -= 1;
						int ret1 = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
							+ constMonth + QString::fromLocal8Bit(std::string("供货").data()), i).toInt();
						if (ret1 != 0) {
							isok = 1;
							break;
						}
						diff--;
					}
				}
				if (isok) {
					int dyhz = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
						+ constMonth + QString::fromLocal8Bit(std::string("可售").data()), i).toInt();
					int qhljz;
					int wyyjll;
					int dyear = currentYear;
					int dmonth = currentMonth;
					int count = 3;
					int calcount = 3;
					int output;
					while (calcount--) {
						if (currentMonth > 1) {
							dmonth -= 1;
							double qhl = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth) 
								+ constMonth + QString::fromLocal8Bit(std::string("去化率").data()), i).toDouble();
							if (qhl < 0.00001)
								count -= 1;
							else {
								output += qhl;
							}
						}
						else {
							dmonth = 12;
							dyear--;
							double qhl = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
								+ constMonth + QString::fromLocal8Bit(std::string("去化率").data()), i).toDouble();
							if (qhl < 0.00001)
								count -= 1;
							else {
								output += qhl;
							}
						}
					}
					if (count > 0) {
						qhljz = output / count;
					}
					else {
						qhljz = 0;
					}

					int circletime = 7;
					dmonth = currentMonth;
					dyear = currentYear;
					int maxdata;
					int mindata;
					int alldata = 0;
					if (dmonth > 1) {
						dmonth -= 1;
						int firstdata = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
							+ constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), i).toInt();
						maxdata = firstdata;
						mindata = firstdata;
						alldata += firstdata;
					}
					else {
						dmonth = 12;
						dyear -= 1;
						int firstdata = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
							+ constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), i).toInt();
						maxdata = firstdata;
						mindata = firstdata;
						alldata += firstdata;
					}

					for (int tempi = 0; tempi < circletime - 1; tempi++) {
						if (dmonth > 1) {
							dmonth -= 1;
							int data = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
								+ constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), i).toInt();
							maxdata = maxdata > data ? maxdata : data;
							mindata = mindata < data ? mindata : data;
							alldata += data;
						}
						else {
							dmonth = 12;
							dyear -= 1;
							int data = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth) 
								+ constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), i).toInt();
							maxdata = maxdata > data ? maxdata : data;
							mindata = mindata < data ? mindata : data;
							alldata += data;
						}
					}
					wyyjll = (alldata - maxdata - mindata) / (circletime - 2);
					int dyqyxdc = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
						+ constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), i).toInt();
					server->writedata(QVariant(max(min(min(dyhz * qhljz, wyyjll), dyhz), dyqyxdc)), QString::number(currentYear)
						+ constYear + QString::number(currentMonth) + constMonth
						+ QString::fromLocal8Bit(std::string("预判计算").data()), i);
				}
				else {
					dyear = currentYear;
					dmonth = currentMonth;
					int yjll;
					double radio;
					int dyhz;
					int dyks;
					int output;
					QString syb;
					int alldata = 0;
					int ncount = 3;
					for (int tmpi = 0; tmpi < ncount; tmpi++) {
						if (dmonth > 1) {
							dmonth -= 1;
							int data = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
								+ constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), i).toInt();
							alldata += data;
						}
						else {
							dmonth = 12;
							dyear -= 1;
							int data = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
								+ constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), i).toInt();
							alldata += data;
						}
					}
					yjll = int(alldata / 3);
					dyhz = server->getCellData(QString::number(dyear) + constYear + QString::number(currentMonth) + constMonth
						+ QString::fromLocal8Bit(std::string("供货").data()), i).toInt();
					dyks = server->getCellData(QString::number(dyear) + constYear + QString::number(currentMonth) + constMonth
						+ QString::fromLocal8Bit(std::string("可售").data()), i).toInt();
					syb = server->getCellData(QString::fromLocal8Bit(std::string("事业部（住开/商开）").data()), i).toString();
					if (syb.compare(QString::fromLocal8Bit(std::string("住开").data())) == 0)
						radio = 0.5;
					else if (syb.compare(QString::fromLocal8Bit(std::string("商开").data())) == 0) {
						QString yt = server->getCellData(QString::fromLocal8Bit(std::string("业态").data()), i).toString();
						if (yt.compare(QString::fromLocal8Bit(std::string("住宅").data()))) {
							radio = 0.6;
						}
						else if (yt.compare(QString::fromLocal8Bit(std::string("商铺").data()))) {
							radio = 0.22;
						}
						else if (yt.compare(QString::fromLocal8Bit(std::string("公寓/办公").data()))) {
							radio = 0.5;
						}
					}
					output = yjll + radio * dyhz;

					int dyqyydc = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
						+ constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), i).toInt();

					dyear = currentYear;
					dmonth = currentMonth;
					int maxqydc = 0;
					ncount = 7;
					for (int tmpi = 0; tmpi < ncount; tmpi++) {
						if (dmonth > 1) {
							dmonth -= 1;
							int data = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
								+ constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), i).toInt();
							maxqydc = maxqydc > data ? maxqydc : data;
						}
						else {
							dmonth = 12;
							dyear -= 1;
							int data = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
								+ constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), i).toInt();
							maxqydc = maxqydc > data ? maxqydc : data;
						}
					}
					server->writedata(QVariant(max(min(min(output, maxqydc), dyks), dyqyydc)), QString::number(currentYear) + constYear
						+ QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), i);
				}
			}
		}
		else
		{
		    printf("Warning : row %d missing->%s.\n", i, qPrintable(xmlx));
		}
	}
}

void predictcurrentMonth3(ExcelDataServer* server) {
	for (int tempi = startrow; tempi <= endrow; tempi++) {
		int sumdata1 = 0;
		for (int i = 1; i <= currentMonth; i++) {
			sumdata1 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("全口径回笼已达成").data()), tempi).toInt();
		}
		int qkj = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("期初全口径应收款").data()), tempi).toInt();
		int sumdata2 = 0;
		if (currentMonth == 1) {
			sumdata2 += server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("签约指标").data()), tempi).toInt();
		}
		else {
			for (int i = 1; i < currentMonth; i++) {
				sumdata2 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), tempi).toInt();
			}
			sumdata2 += server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("签约指标").data()), tempi).toInt();
		}
		if ((qkj + sumdata2) == 0)
		{
			server->writedata(QVariant("/"), QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("综合全口径回笼率").data()), tempi);
		}
		else
		{
			server->writedata(QVariant(int(sumdata1 / (qkj + sumdata2))), QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("综合全口径回笼率").data()), tempi);
		}
		
	}
}

void predictcurrentMonth4(ExcelDataServer* server) {
	for (int tempi = startrow; tempi <= endrow; tempi++) {
		int sumdata1 = 0;
		for (int i = 1; i <= currentMonth; i++) {
			sumdata1 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("权益口径回笼已达成").data()), tempi).toInt();
		}
		int qkj = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("期初权益口径应收款").data()), tempi).toInt();
		int sumdata2 = 0;
		if (currentMonth == 1) {
			sumdata2 += server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("签约指标").data()), tempi).toInt();
		}
		else {
			for (int i = 1; i < currentMonth; i++) {
				sumdata2 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), tempi).toInt();
			}
			sumdata2 += server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("签约指标").data()), tempi).toInt();
		}
		double gqbl = server->getCellData(QString::fromLocal8Bit(std::string("股权比例").data()), tempi).toInt();
		
		if ((qkj + gqbl * sumdata2) == 0)
		{
			server->writedata(QVariant("/"), QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("综合权益口径回笼率").data()), tempi);
		}
		else
		{
			server->writedata(QVariant(int(sumdata1 / (qkj + gqbl * sumdata2))), QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("综合权益口径回笼率").data()), tempi);
		}
	}
}

void yearPredict1(ExcelDataServer* server) {
	for (int tempi = startrow; tempi <= endrow; tempi++) {
		if (currentMonth == 1) {
			int data = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("全口径预估达成（指标排摸）").data()), tempi).toInt();
			server->writedata(QVariant(data), QString::number(currentYear) + constYear + QString::fromLocal8Bit(std::string("年度全口径回笼达成（事业部版）").data()), tempi);
		}
		else {
			int sumdata = 0;
			for (int i = 1; i < currentMonth; i++)
			{
				sumdata += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("全口径回笼已达成").data()), tempi).toInt();
			}
			sumdata += server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("全口径预估达成（指标排摸）").data()), tempi).toInt();

			server->writedata(QVariant(sumdata), QString::number(currentYear) + constYear + QString::fromLocal8Bit(std::string("年度全口径回笼达成（事业部版）").data()), tempi);
		}
	}
}

void yearPredict2(ExcelDataServer* server) {
	for (int tempi = startrow; tempi <= endrow; tempi++) {
		if (currentMonth == 1) {
			int data = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("回笼预估达成（趋势版）").data()), tempi).toInt();
			server->writedata(QVariant(data), QString::number(currentYear) + constYear + QString::fromLocal8Bit(std::string("年度全口径回笼达成（趋势版）").data()), tempi);
		}
		else {
			int sumdata = 0;
			for (int i = 1; i < currentMonth; i++)
			{
				sumdata += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("全口径回笼已达成").data()), tempi).toInt();
			}
			sumdata += server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("回笼预估达成（趋势版）").data()), tempi).toInt();

			server->writedata(QVariant(sumdata), QString::number(currentYear) + constYear + QString::fromLocal8Bit(std::string("年度全口径回笼达成（趋势版）").data()), tempi);
		}
	}
}

void yearPredict3(ExcelDataServer* server) {
	for (int tempi = startrow; tempi <= endrow; tempi++) {
		if (currentMonth == 1) {
			int data = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("签约预估达成（指标排摸）").data()), tempi).toInt();
			server->writedata(QVariant(data), QString::number(currentYear) + constYear + QString::fromLocal8Bit(std::string("年度签约达成（事业部版）").data()), tempi);
		}
		else {
			int sumdata = 0;
			for (int i = 1; i < currentMonth; i++)
			{
				sumdata += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), tempi).toInt();
			}
			sumdata += server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("签约预估达成（指标排摸）").data()), tempi).toInt();

			server->writedata(QVariant(sumdata), QString::number(currentYear) + constYear + QString::fromLocal8Bit(std::string("年度签约达成（事业部版）").data()), tempi);
		}
	}
}

void yearPredict4(ExcelDataServer* server) {
	for (int tempi = startrow; tempi <= endrow; tempi++) {
		if (currentMonth == 1) {
			int data = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), tempi).toInt();
			server->writedata(QVariant(data), QString::number(currentYear) + constYear + QString::fromLocal8Bit(std::string("年度签约达成（趋势版）").data()), tempi);
		}
		else {
			int sumdata = 0;
			for (int i = 1; i < currentMonth; i++)
			{
				sumdata += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), tempi).toInt();
			}
			sumdata += server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), tempi).toInt();

			server->writedata(QVariant(sumdata), QString::number(currentYear) + constYear + QString::fromLocal8Bit(std::string("年度签约达成（趋势版）").data()), tempi);
		}
	}
}

void yearPredict5(ExcelDataServer* server) {
	for (int tempi = startrow; tempi <= endrow; tempi++) {
		int data = server->getCellData(QString::number(currentYear) + constYear + QString::fromLocal8Bit(std::string("年度全口径回笼达成（事业部版）").data()), tempi).toInt();
		int gqbl = server->getCellData(QString::fromLocal8Bit(std::string("股权比例").data()), tempi).toInt();
		server->writedata(QVariant(int(gqbl * data)), QString::number(currentYear) + constYear + QString::fromLocal8Bit(std::string("年度权益口径回笼达成（事业部版）").data()), tempi);
	}
}

void yearPredict6(ExcelDataServer* server) {
	for (int tempi = startrow; tempi <= endrow; tempi++) {
		if (currentMonth == 1) {
			int data = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("回笼指标（趋势版）").data()), tempi).toInt();
			server->writedata(QVariant(data), QString::number(currentYear) + constYear + QString::fromLocal8Bit(std::string("年度权益口径回笼达成（趋势版）").data()), tempi);
		}
		else {
			int sumdata = 0;
			for (int i = 1; i < currentMonth; i++)
			{
				sumdata += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("权益口径回笼已达成").data()), tempi).toInt();
			}
			sumdata += server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("回笼指标（趋势版）").data()), tempi).toInt();

			server->writedata(QVariant(sumdata), QString::number(currentYear) + constYear + QString::fromLocal8Bit(std::string("年度权益口径回笼达成（趋势版）").data()), tempi);
		}
	}
}

void calYear(QString filepath, ExcelDataServer* dataServer) {
	QFile file(filepath);
	if (!file.open(QIODevice::ReadOnly | QIODevice::Text))
	{
		qDebug() << "Can't open the file!";
	}
	while (!file.atEnd())
	{
		QByteArray line = file.readLine();
		QString str(line);

		QStringList list = str.split("=");
		QString rec = list[1].trimmed();
		QString label = list[0].trimmed();

		QString longstring = expand(rec, currentMonth, currentYear);
		if (!longstring.isEmpty()) {
			printf("%s\t", qPrintable(longstring));
			printf("%s\n", qPrintable(QString::number(currentYear) + constYear + label));
			dataServer->calculator(longstring, QString::number(currentYear) + constYear + label);
		}
	}
}

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

	startrow = 4;
	endrow = rowsNumber;
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
	//19年1月期初库存+19年1月供货     19年1月可售
	calHistary(QString::fromLocal8Bit(std::string("./input/1历史月-基础.txt").data()), excelServer);
	
	calCurrent(QString::fromLocal8Bit(std::string("./input/2当前月-基础.txt").data()), excelServer);
	
	calYear(QString::fromLocal8Bit(std::string("./input/3年度-基础.txt").data()), excelServer);

	calHistary(QString::fromLocal8Bit(std::string("./input/4历史月-年度.txt").data()), excelServer);

	calCurrent(QString::fromLocal8Bit(std::string("./input/5当前月-年度.txt").data()), excelServer);

	//6复杂公式-年度.txt
	//年度签约达成（事业部版）
	yearPredict3(excelServer);
	//年度全口径回笼达成（事业部版）
	yearPredict1(excelServer);
	//年度权益口径回笼达成（事业部版）
	yearPredict5(excelServer);
	//综合全口径回笼率
	predictcurrentMonth3(excelServer);
	//综合权益口径回笼率
	predictcurrentMonth4(excelServer);

	//7复杂公式 - 预判计算.txt
	//项目类型
	predictcurrentMonth1(excelServer);
	//预测
	predictcurrentMonth2(excelServer);

	calCurrent(QString::fromLocal8Bit(std::string("./input/8当前月-差异和回笼预判.txt").data()), excelServer);
	calCurrent(QString::fromLocal8Bit(std::string("./input/9当前月-趋势.txt").data()), excelServer);
	
	//年度签约达成（趋势版）
	yearPredict4(excelServer);
	//年度全口径回笼达成（趋势版）
	yearPredict2(excelServer);
	//年度权益口径回笼达成（趋势版）
	yearPredict6(excelServer);

	calYear(QString::fromLocal8Bit(std::string("./input/11年度-趋势.txt").data()), excelServer);
	
	//******************************************************************************************8*******

	//write all data
	QVariant var;
	excelServer->castSheetVector2Variant(var);
	usedrange->setProperty("Value", var);

	worrkbook->dynamicCall("Save()");
	excelServer->freeExcel();
	a.exit();
	return 0;
}






