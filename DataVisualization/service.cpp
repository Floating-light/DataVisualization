#include "service.h"
#include <QString>
#include <QDate>

QDate D1 = QDate::currentDate();
int startYear = 19;
const int currentYear = D1.year() - 2000;//获取年
const int currentMonth = D1.month();
const QString constYear = QString::fromLocal8Bit(std::string("年").data());
const QString constMonth = QString::fromLocal8Bit(std::string("月").data());

service::service()
{
}

service::service(int start, int end)
{
	startRow = start;
	endRow = end;
}

service::~service()
{

}

//-------------------------------------基础方法------------------------------------------

int service::getEndRow() {
	return endRow;
}

int service::getStartRow() {
	return startRow;
}

void service::setEndRow(int v) {
	endRow = v;
}

void service::setStartRow(int v) {
	startRow = v;
}

//--------------------------------------报表打入----------------------------------------
void service::loadReport() {


}


//--------------------------------------工具方法------------------------------------------
//工具方法 字符替换
QString service::replace(QStringList & v, int month, int year) {
	for (int i = 0; i < v.size(); i++)
	{
		if (v[i].startsWith("(") || v[i].startsWith(")")) {
			;
		}
		else if (!v[i][0].isDigit() && !v[i].startsWith("/") && !v[i].startsWith("*") && !v[i].startsWith("+") && !v[i].startsWith("-")) {
			if (v[i].startsWith("!"))
			{
				v[i] = v[i].mid(1, -1);
			}
			else if (v[i].startsWith("@")) {
				v[i] = QString::number(year) + constYear + v[i].mid(1, -1);
			}
			else if (v[i].startsWith("#")) {
				if (month == 1 && year == startYear) {
					return "";
				}
				else {
					v[i] = QString::number(year) + constYear + QString::number(month - 1) + constMonth + v[i].mid(1, -1);
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
			else if (v[i].startsWith("&")) {
				QString temp = "(";
				for (int tempi = 1; tempi <= 12; tempi++) {
					temp.append(QString::number(year) + constYear + QString::number(tempi) + constMonth + v[i].mid(1, -1) + "+");
				}
				temp.chop(1);
				v[i] = temp + ")";
			}
			else if (v[i].startsWith("%")) {
				QString temp = "(";
				for (int tempi = month - 2; tempi <= month; tempi++) {
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

//工具方法 字符展开
QString service::expand(QString v, int month, int year) {
	QStringList list = v.split(" ");
	return replace(list, month, year);
}

//当前月任务计算
void service::calCurrent(QString filepath, ExcelDataServer* dataServer) {
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
		if (!longstring.isEmpty()) {
			printf("Process ... ...\n");
			printf("\t%s\t", qPrintable(longstring));
			printf("%s\n", qPrintable(QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + label));
			dataServer->calculator(longstring, QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + label);
			printf("Process ... ... done\n");
		}
	}
}

//历史月任务计算
void service::calHistary(QString filepath, ExcelDataServer* dataServer) {
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
		for (int i = 1; i <= currentMonth-1; i++) {
			QString longstring = expand(rec, i, currentYear);
			if (!longstring.isEmpty()) {
				printf("Process ... ...\n");
				printf("\t%s\t", qPrintable(longstring));
				printf("%s\n", qPrintable(QString::number(currentYear) + constYear + QString::number(i) + constMonth + label));
				dataServer->calculator(longstring, QString::number(currentYear) + constYear + QString::number(i) + constMonth + label);
				printf("Process ... ... done\n");
			}
		}
	}
}

//年任务计算
void service::calYear(QString filepath, ExcelDataServer* dataServer) {
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

//未来月计算任务
void service::calFutureMonth(QString filepath,ExcelDataServer* dataServer) {
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
		for (int i = currentMonth + 1; i <= 12; i++) {
			QString longstring = expand(rec, i, currentYear);
			if (!longstring.isEmpty()) {
				printf("Process ... ...\n");
				printf("\t%s\t", qPrintable(longstring));
				printf("%s\n", qPrintable(QString::number(currentYear) + constYear + QString::number(i) + constMonth + label));
				dataServer->calculator(longstring, QString::number(currentYear) + constYear + QString::number(i) + constMonth + label);
				printf("Process ... ... done\n");
			}
		}
	}
}


//----------------------------------当前月预判--------------------------------
//当前月项目类型
void service::predictcurrentMonth1( ExcelDataServer* server) {
	int count_firstOpen = 0;
	int count_pulus = 0;
	int count_continue = 0;
	int count_null = 0;
	for (int i = this->startRow; i < this->endRow; i++) {
		double ndqydc = server->getCellData(QString::number(currentYear) + constYear
			+ QString::fromLocal8Bit(std::string("年度签约已达成").data()), i).toDouble();

		double dyqyydc = server->getCellData(QString::number(currentYear) + constYear
			+ QString::number(currentMonth) + constMonth +
			QString::fromLocal8Bit(std::string("签约已达成").data()), i).toDouble();

		double dyqckc = server->getCellData(QString::number(currentYear) + constYear +
			QString::number(currentMonth) + constMonth +
			QString::fromLocal8Bit(std::string("期初库存").data()), i).toDouble();

		double dygh = server->getCellData(QString::number(currentYear) + constYear +
			QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("供货").data()), i).toDouble();
		if (ndqydc - dyqyydc == 0 && dyqckc == 0 && dygh> 0)
		{
			++count_firstOpen;
			server->writedata(QVariant(QString::fromLocal8Bit(std::string("首开类").data())),
				QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth +
				QString::fromLocal8Bit(std::string("项目类型").data()), i);
		}
		else if (dyqckc > 0 && dygh> 0) {
			++count_pulus;
			server->writedata(QVariant(QString::fromLocal8Bit(std::string("加推类").data())),
				QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth +
				QString::fromLocal8Bit(std::string("项目类型").data()), i);
		}
		else if (ndqydc - dyqyydc > 0 && dygh == 0 || dyqckc > 0 && dygh == 0) {
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

//当前月预判计算
void service::predictcurrentMonth2(ExcelDataServer* server) {
	for (int i = this->startRow; i < this->endRow; i++) {
		QString xmlx = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
			+ constMonth + QString::fromLocal8Bit(std::string("项目类型").data()), i).toString();
		if (xmlx.compare(QString::fromLocal8Bit(std::string("首开类").data())) == 0) {
			QString syb = server->getCellData(QString::fromLocal8Bit(std::string("事业部（住开/商开）").data()), i).toString();
			if (syb.compare(QString::fromLocal8Bit(std::string("商开").data())) == 0) {
				QString yt = server->getCellData(QString::fromLocal8Bit(std::string("业态").data()), i).toString();
				if (yt.compare(QString::fromLocal8Bit(std::string("住宅").data())) == 0) {
					double hz = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
						+ constMonth + QString::fromLocal8Bit(std::string("可售").data()), i).toDouble();
					double result = hz * 0.7;
					server->writedata(QVariant(result < 50000 ? result : 50000), QString::number(currentYear) + constYear
						+ QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), i);
				}
				else if (yt.compare(QString::fromLocal8Bit(std::string("商铺").data())) == 0) {
					double hz = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
						+ constMonth + QString::fromLocal8Bit(std::string("可售").data()), i).toDouble();
					double result = hz * 0.65;
					server->writedata(QVariant(result < 50000 ? result : 50000), QString::number(currentYear) + constYear
						+ QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), i);
				}
				else if (yt.compare(QString::fromLocal8Bit(std::string("公寓/办公").data())) == 0) {
					double hz = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
						+ constMonth + QString::fromLocal8Bit(std::string("可售").data()), i).toDouble();
					double result = hz * 0.35;
					server->writedata(QVariant(result < 50000 ? result : 50000), QString::number(currentYear) + constYear
						+ QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), i);
				}
				else {
					double hz = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
						+ constMonth + QString::fromLocal8Bit(std::string("可售").data()), i).toDouble();
					double result = hz * 0.35;
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
					double hz = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
						+ constMonth + QString::fromLocal8Bit(std::string("可售").data()), i).toDouble();
					double result = hz * 0.6;
					server->writedata(QVariant(result < 50000 ? result : 50000), QString::number(currentYear) + constYear
						+ QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), i);
				}
				else if (cshx.compare(QString::fromLocal8Bit(std::string("四线").data())) == 0
					|| cshx.compare(QString::fromLocal8Bit(std::string("五线").data())) == 0)
				{
					double hz = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
						+ constMonth + QString::fromLocal8Bit(std::string("可售").data()), i).toDouble();
					double result = hz * 0.35;
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
					else if (yt.compare(QString::fromLocal8Bit(std::string("商铺").data())) == 0) {
						radio = 0.22;
					}
					else if (yt.compare(QString::fromLocal8Bit(std::string("公寓/办公").data())) == 0) {
						radio = 0.5;
					}
					else {
						radio = 0.35;
					}
				}
				double ret1 = server->getCellData(QString::number(fyear) + constYear + QString::number(fmonth) + constMonth
					+ QString::fromLocal8Bit(std::string("签约已达成").data()), i).toDouble();
				double ret2 = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
					+ constMonth + QString::fromLocal8Bit(std::string("可售").data()), i).toDouble();
				double ret3 = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
					+ constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), i).toDouble();
				double retf = max(ret3, min(ret1 * radio, ret2 * 0.8));
				server->writedata(QVariant(retf), QString::number(currentYear) + constYear + QString::number(currentMonth)
					+ constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), i);
			}
			else {
				int count = 0;
				int remain = 3;
				double output = 0;
				int dyear = currentYear;
				int dmonth = currentMonth;
				while (remain > 0) {
					dmonth--;
					if (dmonth == 0 ||dmonth+(currentYear-fyear)*12-fmonth<=0)
						break;
					double ret1 = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
						+ constMonth + QString::fromLocal8Bit(std::string("供货").data()), i).toDouble();
					if (ret1 == 0)
					{
						double ret2 = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
							+ constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), i).toDouble();
						if (ret2 == 0) {
							;
						}
						else {
							output += ret2;
							count++;	
						}
						remain -= 1;
					}
				}
				double dyhz = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth
					+ QString::fromLocal8Bit(std::string("可售").data()), i).toDouble();
				double dyqyydc = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth
					+ QString::fromLocal8Bit(std::string("签约已达成").data()), i).toDouble();
				if (count == 0)
					server->writedata(QVariant(dyqyydc), QString::number(currentYear) + constYear + QString::number(currentMonth)
						+ constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), i);
				else
					server->writedata(QVariant(max(dyqyydc, min(dyhz, output / count))), QString::number(currentYear) + constYear
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
					if (yt.compare(QString::fromLocal8Bit(std::string("住宅").data()))==0) {
						radio = 0.6;
					}
					else if (yt.compare(QString::fromLocal8Bit(std::string("商铺").data()))==0) {
						radio = 0.22;
					}
					else if (yt.compare(QString::fromLocal8Bit(std::string("公寓/办公").data()))==0) {
						radio = 0.5;
					}
					else {
						radio = 0.35;
					}
				}

				double ret1 = server->getCellData(QString::number(fyear) + constYear + QString::number(fmonth) + constMonth
					+ QString::fromLocal8Bit(std::string("签约已达成").data()), i).toDouble();
				double ret2 = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
					+ constMonth + QString::fromLocal8Bit(std::string("可售").data()), i).toDouble();
				double ret3 = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
					+ constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), i).toDouble();
				double retf = max(ret3, min(ret1 * radio, ret2 * 0.8));
				server->writedata(QVariant(retf), QString::number(currentYear) + constYear + QString::number(currentMonth) +
					constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), i);
			}
			else {
				int diff = 3;
				double output;
				int dyear = currentYear;
				int dmonth = currentMonth;
				bool isok = false;
				while (diff > 0) {
					dmonth--;
					double ret1 = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
						+ constMonth + QString::fromLocal8Bit(std::string("供货").data()), i).toDouble();
					if (ret1 != 0) {
						isok = true;
						break;
					}
					diff--;
				}
				if (isok) {
					double dyhz = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
						+ constMonth + QString::fromLocal8Bit(std::string("可售").data()), i).toDouble();
					double qhljz;
					double wyyjll;
					int dyear = currentYear;
					int dmonth = currentMonth;
					int count = 0;
					int calcount = 3;
					double output;
					while (calcount>0) {
						dmonth -= 1;
						double qhl = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
							+ constMonth + QString::fromLocal8Bit(std::string("去化率").data()), i).toDouble();
						if (qhl < 0.00001)
						{
							;
						}
						else {
							output += qhl;
							count++;
						}
						calcount--;
					}
					if (count > 0) {
						qhljz = output / count;
					}
					else {
						qhljz = 0;
					}

					int circletime;
					if(currentYear==fyear)
						circletime = currentMonth - fmonth;
					else
						circletime = currentMonth - 1;
					dmonth = currentMonth;
					dyear = currentYear;
					double maxdata = std::numeric_limits<double>::min();
					double mindata = std::numeric_limits<double>::max();
					double alldata = 0;
					//当前月-1
					for (int tempi = 0; tempi < circletime; tempi++) {
						dmonth -= 1;
						double data = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
							+ constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), i).toDouble();
						maxdata = maxdata > data ? maxdata : data;
						mindata = mindata < data ? mindata : data;
						alldata += data;
					}
					wyyjll = (alldata - maxdata - mindata) / (circletime - 2);
					double dyqyxdc = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
						+ constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), i).toDouble();
					server->writedata(QVariant(max(min(min(dyhz * qhljz, wyyjll), dyhz), dyqyxdc)), QString::number(currentYear)
						+ constYear + QString::number(currentMonth) + constMonth
						+ QString::fromLocal8Bit(std::string("预判计算").data()), i);
				}
				else {
					dyear = currentYear;
					dmonth = currentMonth;
					double yjll;
					double radio;
					double dyhz;
					double dyks;
					double output;
					QString syb;
					double alldata = 0;
					int ncount = 3;
					for (int tmpi = 0; tmpi < ncount; tmpi++) {
						dmonth -= 1;
						double data = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
							+ constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), i).toDouble();
						alldata += data;
					}
					yjll = alldata / 3;

					dyhz = server->getCellData(QString::number(dyear) + constYear + QString::number(currentMonth) + constMonth
						+ QString::fromLocal8Bit(std::string("供货").data()), i).toDouble();
					dyks = server->getCellData(QString::number(dyear) + constYear + QString::number(currentMonth) + constMonth
						+ QString::fromLocal8Bit(std::string("可售").data()), i).toDouble();
					syb = server->getCellData(QString::fromLocal8Bit(std::string("事业部（住开/商开）").data()), i).toString();
					if (syb.compare(QString::fromLocal8Bit(std::string("住开").data())) == 0)
					{
						QString cshx = server->getCellData(QString::fromLocal8Bit(std::string("城市环线").data()), i).toString();
						if (cshx.compare(QString::fromLocal8Bit(std::string("一线").data())) == 0
							|| cshx.compare(QString::fromLocal8Bit(std::string("二线").data())) == 0
							|| cshx.compare(QString::fromLocal8Bit(std::string("三线").data())) == 0) {
							radio = 0.6;
						}
						else if (cshx.compare(QString::fromLocal8Bit(std::string("四线").data())) == 0
							|| cshx.compare(QString::fromLocal8Bit(std::string("五线").data())) == 0) {
							radio = 0.35;
						}
					}
					else if (syb.compare(QString::fromLocal8Bit(std::string("商开").data())) == 0) {
						QString yt = server->getCellData(QString::fromLocal8Bit(std::string("业态").data()), i).toString();
						if (yt.compare(QString::fromLocal8Bit(std::string("住宅").data()))==0) {
							radio = 0.7;
						}
						else if (yt.compare(QString::fromLocal8Bit(std::string("商铺").data()))==0) {
							radio = 0.65;
						}
						else if (yt.compare(QString::fromLocal8Bit(std::string("公寓/办公").data()))==0) {
							radio = 0.35;
						}
						else {
							radio = 0.35;
						}
					}
					output = yjll + radio * dyhz;

					double dyqyydc = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
						+ constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), i).toDouble();

					dyear = currentYear;
					dmonth = currentMonth;
					double maxqydc = 0;
					ncount = currentMonth - 1;
					for (int tmpi = 0; tmpi < ncount; tmpi++) {
						dmonth -= 1;
						double data = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
							+ constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), i).toDouble();
						maxqydc = maxqydc > data ? maxqydc : data;
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

//当前月综合全口径回笼率
void service::predictcurrentMonth3(ExcelDataServer* server) {
	for (int tempi = this->startRow; tempi <= this->endRow; tempi++) {
		double sumdata1 = 0;
		for (int i = 1; i <= currentMonth; i++) {
			sumdata1 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("全口径回笼已达成").data()), tempi).toDouble();
		}
		double qkj = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("期初全口径应收款").data()), tempi).toDouble();
		double sumdata2 = 0;
		if (currentMonth == 1) {
			sumdata2 += server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("签约指标").data()), tempi).toDouble();
		}
		else {
			for (int i = 1; i < currentMonth; i++) {
				sumdata2 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), tempi).toDouble();
			}
			sumdata2 += server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("签约指标").data()), tempi).toDouble();
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

//当前月综合权益口径回笼率
void service::predictcurrentMonth4(ExcelDataServer* server) {
	for (int tempi = this->startRow; tempi <= this->endRow; tempi++) {
		double sumdata1 = 0;
		for (int i = 1; i <= currentMonth; i++) {
			sumdata1 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("权益口径回笼已达成").data()), tempi).toDouble();
		}
		double qkj = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("期初权益口径应收款").data()), tempi).toDouble();
		double sumdata2 = 0;
		if (currentMonth == 1) {
			sumdata2 += server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("签约指标").data()), tempi).toDouble();
		}
		else {
			for (int i = 1; i < currentMonth; i++) {
				sumdata2 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), tempi).toDouble();
			}
			sumdata2 += server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("签约指标").data()), tempi).toDouble();
		}
		double gqbl = server->getCellData(QString::fromLocal8Bit(std::string("股权比例").data()), tempi).toDouble();

		if ((qkj + gqbl * sumdata2) == 0)
		{
			server->writedata(QVariant("/"), QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("综合权益口径回笼率").data()), tempi);
		}
		else
		{
			server->writedata(QVariant(sumdata1 / (qkj + gqbl * sumdata2)), QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("综合权益口径回笼率").data()), tempi);
		}
	}
}



//---------------------------------年度计算------------------------------------
//年度全口径回笼达成（事业部版）
void service::yearPredict1(ExcelDataServer* server) {
	for (int tempi = this->startRow; tempi <= this->endRow; tempi++) {
		double dataAll = 0;
		for (int i = 1; i<currentMonth; i++)
		{
			dataAll += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("全口径回笼已达成").data()), tempi).toDouble();
		}
		for (int i = currentMonth; i <= 12; i++)
		{
			dataAll += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("全口径预估达成（指标排摸）").data()), tempi).toDouble();
		}
		server->writedata(QVariant(dataAll), QString::number(currentYear) + constYear + QString::fromLocal8Bit(std::string("年度全口径回笼达成（事业部版）").data()), tempi);
	}
}

//年度全口径回笼达成（趋势版）
void service::yearPredict2(ExcelDataServer* server) {
	for (int tempi = this->startRow; tempi <= this->endRow; tempi++) {
		double dataAll = 0;
		for (int i = 1; i<currentMonth; i++)
		{
			dataAll += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("全口径回笼已达成").data()), tempi).toDouble();
		}
		for (int i = currentMonth; i <= 12; i++)
		{
			dataAll += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("回笼预估达成（趋势版）").data()), tempi).toDouble();
		}
		server->writedata(QVariant(dataAll), QString::number(currentYear) + constYear + QString::fromLocal8Bit(std::string("年度全口径回笼达成（趋势版）").data()), tempi);
	}
}

//年度签约达成（事业部版）
void service::yearPredict3(ExcelDataServer* server) {
	for (int tempi = this->startRow; tempi <= this->endRow; tempi++) {
		double dataAll = 0;
		for(int i = 1;i<currentMonth;i++)
		{
			dataAll += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), tempi).toDouble();
		}
		for (int i = currentMonth; i <= 12; i++)
		{
			dataAll += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("签约预估达成（指标排摸）").data()), tempi).toDouble();
		}
		server->writedata(QVariant(dataAll), QString::number(currentYear) + constYear + QString::fromLocal8Bit(std::string("年度签约达成（事业部版）").data()), tempi);
	}
}

//年度签约达成（趋势版）
void service::yearPredict4(ExcelDataServer* server) {
	for (int tempi = this->startRow; tempi <= this->endRow; tempi++) {
		double dataAll = 0;
		for (int i = 1; i<currentMonth; i++)
		{
			dataAll += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), tempi).toDouble();
		}
		for (int i = currentMonth; i <= 12; i++)
		{
			dataAll += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), tempi).toDouble();
		}
		server->writedata(QVariant(dataAll), QString::number(currentYear) + constYear + QString::fromLocal8Bit(std::string("年度签约达成（趋势版）").data()), tempi);
	}
}

//年度权益口径回笼达成（事业部版）
void service::yearPredict5(ExcelDataServer* server) {
	for (int tempi = this->startRow; tempi <= this->endRow; tempi++) {
		double dataAll = 0;
		for (int i = 1; i<currentMonth; i++)
		{
			dataAll += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("权益口径回笼已达成").data()), tempi).toDouble();
		}
		for (int i = currentMonth; i <= 12; i++)
		{
			dataAll += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("权益口径预估达成（指标排摸）").data()), tempi).toDouble();
		}
		server->writedata(QVariant(dataAll), QString::number(currentYear) + constYear + QString::fromLocal8Bit(std::string("年度权益口径回笼达成（事业部版）").data()), tempi);
	}
}

//年度权益口径回笼达成（趋势版）
void service::yearPredict6(ExcelDataServer* server) {
	for (int tempi = this->startRow; tempi <= this->endRow; tempi++) {
		double dataAll = 0;
		for (int i = 1; i<currentMonth; i++)
		{
			dataAll += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("权益口径回笼已达成").data()), tempi).toDouble();
		}
		for (int i = currentMonth; i <= 12; i++)
		{
			dataAll += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("回笼指标（趋势版）").data()), tempi).toDouble();
		}
		server->writedata(QVariant(dataAll), QString::number(currentYear) + constYear + QString::fromLocal8Bit(std::string("年度权益口径回笼达成（趋势版）").data()), tempi);
	}
}


//-------------------------------------未来月--------------------------------------
//未来月综合权益口径回笼率
void service::predictNextMonth1(ExcelDataServer* server) {
	for (int nextmonth = currentMonth + 1; nextmonth <= 12; nextmonth++) {
		for (int tempi = this->startRow; tempi <= this->endRow; tempi++) {
			double sumdata1 = 0;
			for (int i = 1; i <= currentMonth; i++) {
				sumdata1 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("权益口径回笼已达成").data()), tempi).toDouble();
			}
			for (int i = currentMonth+1; i <= 12; i++) {
				sumdata1 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("回笼指标（趋势版）").data()), tempi).toDouble();
			}
			double qkj = server->getCellData(QString::number(currentYear) + constYear + QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("期初权益口径应收款").data()), tempi).toDouble();
			double sumdata2 = 0;
			for (int i = 1; i < currentMonth; i++) {
				sumdata2 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), tempi).toDouble();
			}
			for (int i = currentMonth; i <= 12; i++) {
				sumdata2 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("签约指标").data()), tempi).toDouble();
			}
			double gqbl = server->getCellData(QString::fromLocal8Bit(std::string("股权比例").data()), tempi).toDouble();

			if ((qkj + sumdata2*gqbl) == 0)
			{
				server->writedata(QVariant("/"), QString::number(currentYear) + constYear + QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("综合权益口径回笼率").data()), tempi);
			}
			else
			{
				server->writedata(QVariant(sumdata1 / (qkj + sumdata2*gqbl)), QString::number(currentYear) + constYear + QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("综合权益口径回笼率").data()), tempi);
			}
		}
	}
}

//未来月权益口径回笼进度比
void service::predictNextMonth2(ExcelDataServer* server) {
	for (int nextmonth = currentMonth + 1; nextmonth <= 12; nextmonth++) {
		for (int tempi = this->startRow; tempi <= this->endRow; tempi++) {
			double sumdata1 = 0;
			for (int i = 1; i <currentMonth; i++) {
				sumdata1 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("权益口径回笼已达成").data()), tempi).toDouble();
			}
			for (int i = currentMonth; i <= 12; i++) {
				sumdata1 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("权益口径回笼指标").data()), tempi).toDouble();
			}
			double ndqykj = server->getCellData(QString::number(currentYear) + constYear  + QString::fromLocal8Bit(std::string("年度权益口径回笼指标").data()), tempi).toDouble();
			server->writedata(QVariant(sumdata1 / ndqykj), QString::number(currentYear) + constYear + QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("权益口径回笼进度比").data()), tempi);
		}
	}
}

//未来月综合全口径回笼率
void service::predictNextMonth3(ExcelDataServer* server) {
	for (int nextmonth = currentMonth + 1; nextmonth <= 12; nextmonth++) {
		for (int tempi = this->startRow; tempi <= this->endRow; tempi++) {
			double sumdata1 = 0;
			for (int i = 1; i <= currentMonth; i++) {
				sumdata1 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("全口径回笼已达成").data()), tempi).toDouble();
			}
			for (int i = currentMonth+1; i <= 12; i++) {
				sumdata1 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("回笼预估达成（趋势版）").data()), tempi).toDouble();
			}
			double qkj = server->getCellData(QString::number(currentYear) + constYear + QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("初全口径应收款").data()), tempi).toDouble();
			double sumdata2 = 0;
			for (int i = 1; i <currentMonth; i++) {
				sumdata2 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), tempi).toDouble();
			}
			for (int i = currentMonth; i <= 12; i++) {
				sumdata2 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("签约指标").data()), tempi).toDouble();
			}
			if ((qkj + sumdata2) == 0)
			{
				server->writedata(QVariant("/"), QString::number(currentYear) + constYear + QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("综合全口径回笼率").data()), tempi);
			}
			else
			{
				server->writedata(QVariant(int(sumdata1 / (qkj + sumdata2))), QString::number(currentYear) + constYear + QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("综合全口径回笼率").data()), tempi);
			}
		}
	}
}

//未来月全口径回笼进度比
void service::predictNextMonth4(ExcelDataServer* server) {
	for (int nextmonth = currentMonth + 1; nextmonth <= 12; nextmonth++) {
		for (int tempi = this->startRow; tempi <= this->endRow; tempi++) {
			double sumdata1 = 0;
			for (int i = 1; i <currentMonth; i++) {
				sumdata1 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("全口径回笼已达成").data()), tempi).toDouble();
			}
			for (int i = currentMonth; i <= 12; i++) {
				sumdata1 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("全口径回笼指标").data()), tempi).toDouble();
			}
			double ndqykj = server->getCellData(QString::number(currentYear) + constYear + QString::fromLocal8Bit(std::string("年度全口径回笼指标").data()), tempi).toDouble();
			server->writedata(QVariant(sumdata1 / ndqykj), QString::number(currentYear) + constYear + QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("全口径回笼进度比").data()), tempi);
		}
	}
}

//未来月回笼预估达成（趋势版）
void service::predictNextMonth5(ExcelDataServer* server) {
	for (int nextmonth = currentMonth + 1; nextmonth <= 12; nextmonth++) {
		for (int tempi = this->startRow; tempi <= this->endRow; tempi++) {
			double ypjs = server->getCellData(QString::number(currentYear) + constYear + QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), tempi).toDouble();
			int count = 0;
			double alldata = 0;
			for (int i = currentMonth - 1; i > currentMonth-4; i--) {
				double hlqyb = server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("回笼签约比").data()), tempi).toDouble();
				if (hlqyb > 0.00001) {
					alldata += hlqyb;
					count++;
				}
			}
			if (count == 0) {
				server->writedata(QVariant(0), QString::number(currentYear) + constYear + QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("回笼预估达成（趋势版）").data()), tempi);
			}
			else {
				server->writedata(QVariant(ypjs*alldata/count), QString::number(currentYear) + constYear + QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("回笼预估达成（趋势版）").data()), tempi);
			}
		}
	}
}

//未来月项目类型
void service::predictNextMonth6(ExcelDataServer* server) {
	for (int nextmonth = currentMonth + 1; nextmonth <= 12; nextmonth++) {
		for (int tempi = this->startRow; tempi <= this->endRow; tempi++) {
			double ndqydc = server->getCellData(QString::number(currentYear) + constYear
				+ QString::fromLocal8Bit(std::string("年度签约已达成").data()), tempi).toDouble();

			double dyqyydc = server->getCellData(QString::number(currentYear) + constYear
				+ QString::number(nextmonth) + constMonth +
				QString::fromLocal8Bit(std::string("签约已达成").data()), tempi).toDouble();

			double dyqckc = server->getCellData(QString::number(currentYear) + constYear +
				QString::number(nextmonth) + constMonth +
				QString::fromLocal8Bit(std::string("期初库存").data()), tempi).toDouble();

			double dygh = server->getCellData(QString::number(currentYear) + constYear +
				QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("供货").data()), tempi).toDouble();
			if (ndqydc - dyqyydc == 0 && dyqckc == 0 && dygh> 0)
			{
				server->writedata(QVariant(QString::fromLocal8Bit(std::string("首开类").data())),
					QString::number(currentYear) + constYear + QString::number(nextmonth) + constMonth +
					QString::fromLocal8Bit(std::string("项目类型").data()), tempi);
			}
			else if (dyqckc > 0 && dygh> 0) {
				server->writedata(QVariant(QString::fromLocal8Bit(std::string("加推类").data())),
					QString::number(currentYear) + constYear + QString::number(nextmonth) + constMonth +
					QString::fromLocal8Bit(std::string("项目类型").data()), tempi);
			}
			else if (ndqydc - dyqyydc > 0 && dygh == 0 || dyqckc > 0 && dygh == 0) {
				server->writedata(QVariant(QString::fromLocal8Bit(std::string("续销类").data())),
					QString::number(currentYear) + constYear + QString::number(nextmonth) + constMonth +
					QString::fromLocal8Bit(std::string("项目类型").data()), tempi);
			}
		}
	}
}

//未来月签约完成年度进度比
void service::predictNextMonth7(ExcelDataServer* server) {
	for (int nextmonth = currentMonth + 1; nextmonth <= 12; nextmonth++) {
		for (int tempi = this->startRow; tempi <= this->endRow; tempi++) {
			double sumdata1 = 0;
			for (int i = 1; i <currentMonth; i++) {
				sumdata1 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), tempi).toDouble();
			}
			for (int i = currentMonth; i <= 12; i++) {
				sumdata1 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("签约指标").data()), tempi).toDouble();
			}
			double ndqyzb = server->getCellData(QString::number(currentYear) + constYear + QString::fromLocal8Bit(std::string("年度签约指标").data()), tempi).toDouble();
			server->writedata(QVariant(sumdata1 / ndqyzb), QString::number(currentYear) + constYear + QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("签约完成年度进度比").data()), tempi);
		}
	}
}

//未来月预判计算
void service::predictNextMonth8(ExcelDataServer* server) {
	for (int nextmonth = currentMonth + 1; nextmonth <= 12; nextmonth++) {
		for (int tempi = this->startRow; tempi <= this->endRow; tempi++) {
			QString xmlx = server->getCellData(QString::number(currentYear) + constYear + QString::number(nextmonth)
				+ constMonth + QString::fromLocal8Bit(std::string("项目类型").data()), tempi).toString();
			if (xmlx.compare(QString::fromLocal8Bit(std::string("首开类").data())) == 0) {
				QString syb = server->getCellData(QString::fromLocal8Bit(std::string("事业部（住开/商开）").data()), tempi).toString();
				if (syb.compare(QString::fromLocal8Bit(std::string("商开").data())) == 0) {
					QString yt = server->getCellData(QString::fromLocal8Bit(std::string("业态").data()), tempi).toString();
					if (yt.compare(QString::fromLocal8Bit(std::string("住宅").data())) == 0) {
						double hz = server->getCellData(QString::number(currentYear) + constYear + QString::number(nextmonth)
							+ constMonth + QString::fromLocal8Bit(std::string("可售").data()), tempi).toDouble();
						double result = hz * 0.7;
						server->writedata(QVariant(result < 50000 ? result : 50000), QString::number(currentYear) + constYear
							+ QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), tempi);
					}
					else if (yt.compare(QString::fromLocal8Bit(std::string("商铺").data())) == 0) {
						double hz = server->getCellData(QString::number(currentYear) + constYear + QString::number(nextmonth)
							+ constMonth + QString::fromLocal8Bit(std::string("可售").data()), tempi).toDouble();
						double result = hz * 0.65;
						server->writedata(QVariant(result < 50000 ? result : 50000), QString::number(currentYear) + constYear
							+ QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), tempi);
					}
					else if (yt.compare(QString::fromLocal8Bit(std::string("公寓/办公").data())) == 0) {
						double hz = server->getCellData(QString::number(currentYear) + constYear + QString::number(nextmonth)
							+ constMonth + QString::fromLocal8Bit(std::string("可售").data()), tempi).toDouble();
						double result = hz * 0.35;
						server->writedata(QVariant(result < 50000 ? result : 50000), QString::number(currentYear) + constYear
							+ QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), tempi);
					}
					else {
						double hz = server->getCellData(QString::number(currentYear) + constYear + QString::number(nextmonth)
							+ constMonth + QString::fromLocal8Bit(std::string("可售").data()), tempi).toDouble();
						double result = hz * 0.35;
						server->writedata(QVariant(result < 50000 ? result : 50000), QString::number(currentYear) + constYear
							+ QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), tempi);
					}
				}
				else if (syb.compare(QString::fromLocal8Bit(std::string("住开").data())) == 0) {
					QString cshx = server->getCellData(QString::fromLocal8Bit(std::string("城市环线").data()), tempi).toString();
					if (cshx.compare(QString::fromLocal8Bit(std::string("一线").data())) == 0
						|| cshx.compare(QString::fromLocal8Bit(std::string("二线").data())) == 0
						|| cshx.compare(QString::fromLocal8Bit(std::string("三线").data())) == 0)
					{
						double hz = server->getCellData(QString::number(currentYear) + constYear + QString::number(nextmonth)
							+ constMonth + QString::fromLocal8Bit(std::string("可售").data()), tempi).toDouble();
						double result = hz * 0.6;
						server->writedata(QVariant(result < 50000 ? result : 50000), QString::number(currentYear) + constYear
							+ QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), tempi);
					}
					else if (cshx.compare(QString::fromLocal8Bit(std::string("四线").data())) == 0
						|| cshx.compare(QString::fromLocal8Bit(std::string("五线").data())) == 0)
					{
						double hz = server->getCellData(QString::number(currentYear) + constYear + QString::number(nextmonth)
							+ constMonth + QString::fromLocal8Bit(std::string("可售").data()), tempi).toDouble();
						double result = hz * 0.35;
						server->writedata(QVariant(result < 12000 ? result : 12000), QString::number(currentYear) + constYear
							+ QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), tempi);
					}
				}
			}
			else if (xmlx.compare(QString::fromLocal8Bit(std::string("续销类").data())) == 0) {
				QString cshx = server->getCellData(QString::fromLocal8Bit(std::string("首开日期").data()), tempi).toString();
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
				if (nextmonth + (currentYear - fyear) * 12 - fmonth <= 3) {
					double radio = 0;
					QString syb = server->getCellData(QString::fromLocal8Bit(std::string("事业部（住开/商开）").data()), tempi).toString();
					if (syb.compare(QString::fromLocal8Bit(std::string("住开").data())) == 0)
						radio = 0.5;
					else if (syb.compare(QString::fromLocal8Bit(std::string("商开").data())) == 0) {
						QString yt = server->getCellData(QString::fromLocal8Bit(std::string("业态").data()), tempi).toString();
						if (yt.compare(QString::fromLocal8Bit(std::string("住宅").data()))== 0) {
							radio = 0.6;
						}
						else if (yt.compare(QString::fromLocal8Bit(std::string("商铺").data()))== 0) {
							radio = 0.22;
						}
						else if (yt.compare(QString::fromLocal8Bit(std::string("公寓/办公").data()))== 0) {
							radio = 0.5;
						}
						else {
							radio = 0.35;
						}
					}
					double ret1;
					if(fmonth<currentMonth)
						ret1 = server->getCellData(QString::number(fyear) + constYear + QString::number(fmonth) + constMonth
						+ QString::fromLocal8Bit(std::string("签约已达成").data()), tempi).toDouble();
					else
						ret1 = server->getCellData(QString::number(fyear) + constYear + QString::number(fmonth) + constMonth
							+ QString::fromLocal8Bit(std::string("预判计算").data()), tempi).toDouble();
					double ret2 = server->getCellData(QString::number(currentYear) + constYear + QString::number(nextmonth)
						+ constMonth + QString::fromLocal8Bit(std::string("可售").data()), tempi).toDouble();
					double retf = min(ret1 * radio, ret2 * 0.8);
					server->writedata(QVariant(retf), QString::number(currentYear) + constYear + QString::number(nextmonth)
						+ constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), tempi);
				}
				else {
					int count = 0;
					int remain = 3;
					double output = 0;
					int dyear = currentYear;
					int dmonth = nextmonth;
					while (remain > 0) {
						dmonth--;
						if (dmonth == 0 || dmonth+12*(currentYear-fyear) <= fmonth)
							break;
						double ret1 = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
							+ constMonth + QString::fromLocal8Bit(std::string("供货").data()), tempi).toDouble();
						if (ret1 == 0)
						{
							double ret2;
							if(dmonth<currentMonth)
								ret2 = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
								+ constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), tempi).toDouble();
							else
								ret2 = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
									+ constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), tempi).toDouble();
							if (ret2 == 0) {
								;
							}
							else {
								output += ret2;
								count++;
							}
							remain -= 1;
						}
					}
					double dyhz = server->getCellData(QString::number(currentYear) + constYear + QString::number(nextmonth) + constMonth
						+ QString::fromLocal8Bit(std::string("可售").data()), tempi).toDouble();
					if (count == 0)
						server->writedata(QVariant(0), QString::number(currentYear) + constYear + QString::number(nextmonth)
							+ constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), tempi);
					else
						server->writedata(QVariant(min(dyhz, output / count)), QString::number(currentYear) + constYear
							+ QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), tempi);
				}
			}
			else if (xmlx.compare(QString::fromLocal8Bit(std::string("加推类").data())) == 0) {
				QString cshx = server->getCellData(QString::fromLocal8Bit(std::string("首开日期").data()), tempi).toString();
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
				if (nextmonth + (currentYear - fyear) * 12 - fmonth <= 3) {
					double radio = 0;
					QString syb = server->getCellData(QString::fromLocal8Bit(std::string("事业部（住开/商开）").data()), tempi).toString();
					if (syb.compare(QString::fromLocal8Bit(std::string("住开").data())) == 0)
						radio = 0.5;
					else if (syb.compare(QString::fromLocal8Bit(std::string("商开").data())) == 0) {
						QString yt = server->getCellData(QString::fromLocal8Bit(std::string("业态").data()), tempi).toString();
						if (yt.compare(QString::fromLocal8Bit(std::string("住宅").data())) ==0) {
							radio = 0.6;
						}
						else if (yt.compare(QString::fromLocal8Bit(std::string("商铺").data())) == 0) {
							radio = 0.22;
						}
						else if (yt.compare(QString::fromLocal8Bit(std::string("公寓/办公").data())) ==0) {
							radio = 0.5;
						}
						else {
							radio = 0.35;
						}
					}
					double ret1;
					if(currentMonth<=fmonth && currentYear == fyear)
						ret1 = server->getCellData(QString::number(fyear) + constYear + QString::number(fmonth) + constMonth
						+ QString::fromLocal8Bit(std::string("预判计算").data()), tempi).toDouble();
					else
						ret1 = server->getCellData(QString::number(fyear) + constYear + QString::number(fmonth) + constMonth
							+ QString::fromLocal8Bit(std::string("签约已达成").data()), tempi).toDouble();
					double ret2 = server->getCellData(QString::number(currentYear) + constYear + QString::number(nextmonth)
						+ constMonth + QString::fromLocal8Bit(std::string("可售").data()), tempi).toDouble();
					double retf =  min(ret1 * radio, ret2 * 0.8);
					server->writedata(QVariant(retf), QString::number(currentYear) + constYear + QString::number(nextmonth) +
						constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), tempi);
				}
				else {
					int diff = 3;
					double output;
					int dyear = currentYear;
					int dmonth = nextmonth;
					bool isok = false;
					while (diff > 0) {
						dmonth--;
						double ret1 = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
							+ constMonth + QString::fromLocal8Bit(std::string("供货").data()), tempi).toDouble();
						if (ret1 != 0) {
							isok = true;
							break;
						}
						diff--;
					}
					if (isok) {
						double dyhz = server->getCellData(QString::number(currentYear) + constYear + QString::number(nextmonth)
							+ constMonth + QString::fromLocal8Bit(std::string("可售").data()), tempi).toDouble();
						double qhljz;
						double wyyjll;
						int dyear = currentYear;
						int dmonth = nextmonth;
						int count = 0;
						int calcount = 3;
						double output=0;
						while (calcount>0) {
							dmonth -= 1;
							double qhl;
							if(dmonth<currentMonth)
								qhl =server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
								+ constMonth + QString::fromLocal8Bit(std::string("去化率").data()), tempi).toDouble();
							else
							{ 
								double ypjs = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
									+ constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), tempi).toDouble();
								double ks = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
								+ constMonth + QString::fromLocal8Bit(std::string("可售").data()), tempi).toDouble();
								qhl = ypjs/ks;
							}
							if (qhl < 0.00001)
							{
								;
							}
							else {
								output += qhl;
								count++;
							}
							calcount--;
						}
						if (count > 0) {
							qhljz = output / count;
						}
						else {
							qhljz = 0;
						}

						int circletime;
						if(currentYear==fyear)
							circletime = nextmonth- fmonth;
						else
							circletime = nextmonth - 1;
						dmonth = nextmonth;
						dyear = currentYear;
						double maxdata= std::numeric_limits<double>::min();
						double mindata = std::numeric_limits<double>::max();
						double alldata = 0;
						//当前月-1
						for (int cireclei = 0; cireclei < circletime ; cireclei++) {
							dmonth -= 1;
							double data;
							if(dmonth<currentMonth)
								data = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
									+ constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), tempi).toDouble();
							else {
								data = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
									+ constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), tempi).toDouble();
							}
							if (data > maxdata)
								maxdata = data;
							if (data < mindata)
								mindata = data;
							alldata += data;	
						}
						wyyjll = (alldata - maxdata - mindata) / (circletime - 2);
						server->writedata(QVariant(min(min(dyhz * qhljz, wyyjll), dyhz)), QString::number(currentYear)
							+ constYear + QString::number(nextmonth) + constMonth
							+ QString::fromLocal8Bit(std::string("预判计算").data()), tempi);
					}
					else {
						dyear = currentYear;
						dmonth = nextmonth;
						double yjll;
						double radio;
						double dyhz;
						double dyks;
						double output;
						QString syb;
						double alldata = 0;
						int ncount = 3;
						for (int tmpi = 0; tmpi < ncount; tmpi++) {
							dmonth -= 1;
							double data;
							if(dmonth<currentMonth)
								data = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
								+ constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), tempi).toDouble();
							else
								data = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
									+ constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), tempi).toDouble();
							alldata += data;
						}
						yjll = alldata / 3;

						dyhz = server->getCellData(QString::number(dyear) + constYear + QString::number(nextmonth) + constMonth
							+ QString::fromLocal8Bit(std::string("供货").data()), tempi).toDouble();
						dyks = server->getCellData(QString::number(dyear) + constYear + QString::number(nextmonth) + constMonth
							+ QString::fromLocal8Bit(std::string("可售").data()), tempi).toDouble();
						syb = server->getCellData(QString::fromLocal8Bit(std::string("事业部（住开/商开）").data()), tempi).toString();
						if (syb.compare(QString::fromLocal8Bit(std::string("住开").data())) == 0)
						{
							QString cshx = server->getCellData(QString::fromLocal8Bit(std::string("城市环线").data()), tempi).toString();
							if (cshx.compare(QString::fromLocal8Bit(std::string("一线").data())) == 0
								|| cshx.compare(QString::fromLocal8Bit(std::string("二线").data())) == 0
								|| cshx.compare(QString::fromLocal8Bit(std::string("三线").data())) == 0) {
								radio = 0.6;
							}
							else if (cshx.compare(QString::fromLocal8Bit(std::string("四线").data())) == 0
								|| cshx.compare(QString::fromLocal8Bit(std::string("五线").data())) == 0) {
								radio = 0.35;
							}
						}
						else if (syb.compare(QString::fromLocal8Bit(std::string("商开").data())) == 0) {
							QString yt = server->getCellData(QString::fromLocal8Bit(std::string("业态").data()), tempi).toString();
							if (yt.compare(QString::fromLocal8Bit(std::string("住宅").data()))==0) {
								radio = 0.7;
							}
							else if (yt.compare(QString::fromLocal8Bit(std::string("商铺").data()))==0) {
								radio = 0.65;
							}
							else if (yt.compare(QString::fromLocal8Bit(std::string("公寓/办公").data()))==0) {
								radio = 0.35;
							}
							else {
								radio = 0.35;
							}
						}
						output = yjll + radio * dyhz;

						dyear = currentYear;
						dmonth = nextmonth;
						double maxqydc = 0;
						ncount = nextmonth - 1;
						for (int tmpi = 0; tmpi < ncount; tmpi++) {
							dmonth -= 1;
							double data;
							if(dmonth<currentMonth)
								data = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
								+ constMonth + QString::fromLocal8Bit(std::string("签约已达成").data()), tempi).toDouble();
							else
								data = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
									+ constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), tempi).toDouble();
							maxqydc = maxqydc > data ? maxqydc : data;
						}
						server->writedata(QVariant(min(min(output, maxqydc), dyks)), QString::number(currentYear) + constYear
							+ QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("预判计算").data()), tempi);
					}
				}
			}
			else
			{
				printf("Warning : row %d missing->%s.\n", tempi, qPrintable(xmlx));
			}
		}
	}
}

//任务提交
void service::confirm(ExcelDataServer* excelServer) {
	//19年1月期初库存+19年1月供货     19年1月可售
	calHistary(QString::fromLocal8Bit(std::string("./input/1历史月-基础.txt").data()), excelServer);

	calCurrent(QString::fromLocal8Bit(std::string("./input/2当前月-基础.txt").data()), excelServer);

	calFutureMonth(QString::fromLocal8Bit(std::string("./input/3未来月-基础.txt").data()), excelServer);

	calYear(QString::fromLocal8Bit(std::string("./input/4年度-基础.txt").data()), excelServer);

	calHistary(QString::fromLocal8Bit(std::string("./input/5历史月-年度.txt").data()), excelServer);
	
	calCurrent(QString::fromLocal8Bit(std::string("./input/6当前月-年度.txt").data()), excelServer);


	//未来月综合全口径回笼率
	predictNextMonth3(excelServer);
	//未来月综合权益口径回笼率
	predictNextMonth1(excelServer);
	//未来月签约完成年度进度比
	predictNextMonth7(excelServer);
	//未来月全口径回笼进度比
	predictNextMonth4(excelServer);
	//未来月权益口径回笼进度比
	predictNextMonth2(excelServer);


	//年度签约达成（事业部版）
	yearPredict3(excelServer);
	//年度全口径回笼达成（事业部版）
	yearPredict1(excelServer);
	//年度权益口径回笼达成（事业部版）
	yearPredict5(excelServer);
	//（当前月）综合全口径回笼率
	predictcurrentMonth3(excelServer);
	//（当前月）综合权益口径回笼率
	predictcurrentMonth4(excelServer);


	//项目类型
	predictcurrentMonth1(excelServer);
	//预测
	predictcurrentMonth2(excelServer);


	calCurrent(QString::fromLocal8Bit(std::string("./input/10-当前月-差异和回笼预判.txt").data()), excelServer);
	calCurrent(QString::fromLocal8Bit(std::string("./input/11当前月-趋势.txt").data()), excelServer);


	//项目类型(0)
	predictNextMonth6(excelServer);
	//预判计算(1)
	predictNextMonth8(excelServer);
	//回笼预估达成（趋势版）
	predictNextMonth5(excelServer);


	calFutureMonth(QString::fromLocal8Bit(std::string("./input/13未来月-趋势.txt").data()), excelServer);


	//年度签约达成（趋势版）
	yearPredict4(excelServer);
	//年度全口径回笼达成（趋势版）
	yearPredict2(excelServer);
	//年度权益口径回笼达成（趋势版）
	yearPredict6(excelServer);


	calYear(QString::fromLocal8Bit(std::string("./input/15年度-趋势.txt").data()), excelServer);
	//******************************************************************************************8*******
}
