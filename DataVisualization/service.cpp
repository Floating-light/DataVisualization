#include "service.h"
#include <QString>
#include <QDate>

QDate D1 = QDate::currentDate();
int startYear = 19;
const int currentYear = D1.year() - 2000;//��ȡ��
const int currentMonth = D1.month();
const QString constYear = QString::fromLocal8Bit(std::string("��").data());
const QString constMonth = QString::fromLocal8Bit(std::string("��").data());

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

//-------------------------------------��������------------------------------------------

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

//--------------------------------------�������----------------------------------------
void service::loadReport() {


}


//--------------------------------------���߷���------------------------------------------
//���߷��� �ַ��滻
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

//���߷��� �ַ�չ��
QString service::expand(QString v, int month, int year) {
	QStringList list = v.split(" ");
	return replace(list, month, year);
}

//��ǰ���������
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

//��ʷ���������
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

//���������
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

//δ���¼�������
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


//----------------------------------��ǰ��Ԥ��--------------------------------
//��ǰ����Ŀ����
void service::predictcurrentMonth1( ExcelDataServer* server) {
	int count_firstOpen = 0;
	int count_pulus = 0;
	int count_continue = 0;
	int count_null = 0;
	for (int i = this->startRow; i < this->endRow; i++) {
		double ndqydc = server->getCellData(QString::number(currentYear) + constYear
			+ QString::fromLocal8Bit(std::string("���ǩԼ�Ѵ��").data()), i).toDouble();

		double dyqyydc = server->getCellData(QString::number(currentYear) + constYear
			+ QString::number(currentMonth) + constMonth +
			QString::fromLocal8Bit(std::string("ǩԼ�Ѵ��").data()), i).toDouble();

		double dyqckc = server->getCellData(QString::number(currentYear) + constYear +
			QString::number(currentMonth) + constMonth +
			QString::fromLocal8Bit(std::string("�ڳ����").data()), i).toDouble();

		double dygh = server->getCellData(QString::number(currentYear) + constYear +
			QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("����").data()), i).toDouble();
		if (ndqydc - dyqyydc == 0 && dyqckc == 0 && dygh> 0)
		{
			++count_firstOpen;
			server->writedata(QVariant(QString::fromLocal8Bit(std::string("�׿���").data())),
				QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth +
				QString::fromLocal8Bit(std::string("��Ŀ����").data()), i);
		}
		else if (dyqckc > 0 && dygh> 0) {
			++count_pulus;
			server->writedata(QVariant(QString::fromLocal8Bit(std::string("������").data())),
				QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth +
				QString::fromLocal8Bit(std::string("��Ŀ����").data()), i);
		}
		else if (ndqydc - dyqyydc > 0 && dygh == 0 || dyqckc > 0 && dygh == 0) {
			++count_continue;
			server->writedata(QVariant(QString::fromLocal8Bit(std::string("������").data())),
				QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth +
				QString::fromLocal8Bit(std::string("��Ŀ����").data()), i);
		}
		else
		{
			++count_null;
		}
	}
	printf("--------->>first open�� %d, plus push��%d, continue sell�� %d, null�� %d\n", count_firstOpen, count_pulus, count_continue, count_null);
}

//��ǰ��Ԥ�м���
void service::predictcurrentMonth2(ExcelDataServer* server) {
	for (int i = this->startRow; i < this->endRow; i++) {
		QString xmlx = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
			+ constMonth + QString::fromLocal8Bit(std::string("��Ŀ����").data()), i).toString();
		if (xmlx.compare(QString::fromLocal8Bit(std::string("�׿���").data())) == 0) {
			QString syb = server->getCellData(QString::fromLocal8Bit(std::string("��ҵ����ס��/�̿���").data()), i).toString();
			if (syb.compare(QString::fromLocal8Bit(std::string("�̿�").data())) == 0) {
				QString yt = server->getCellData(QString::fromLocal8Bit(std::string("ҵ̬").data()), i).toString();
				if (yt.compare(QString::fromLocal8Bit(std::string("סլ").data())) == 0) {
					double hz = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
						+ constMonth + QString::fromLocal8Bit(std::string("����").data()), i).toDouble();
					double result = hz * 0.7;
					server->writedata(QVariant(result < 50000 ? result : 50000), QString::number(currentYear) + constYear
						+ QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("Ԥ�м���").data()), i);
				}
				else if (yt.compare(QString::fromLocal8Bit(std::string("����").data())) == 0) {
					double hz = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
						+ constMonth + QString::fromLocal8Bit(std::string("����").data()), i).toDouble();
					double result = hz * 0.65;
					server->writedata(QVariant(result < 50000 ? result : 50000), QString::number(currentYear) + constYear
						+ QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("Ԥ�м���").data()), i);
				}
				else if (yt.compare(QString::fromLocal8Bit(std::string("��Ԣ/�칫").data())) == 0) {
					double hz = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
						+ constMonth + QString::fromLocal8Bit(std::string("����").data()), i).toDouble();
					double result = hz * 0.35;
					server->writedata(QVariant(result < 50000 ? result : 50000), QString::number(currentYear) + constYear
						+ QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("Ԥ�м���").data()), i);
				}
				else {
					double hz = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
						+ constMonth + QString::fromLocal8Bit(std::string("����").data()), i).toDouble();
					double result = hz * 0.35;
					server->writedata(QVariant(result < 50000 ? result : 50000), QString::number(currentYear) + constYear
						+ QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("Ԥ�м���").data()), i);
				}
			}
			else if (syb.compare(QString::fromLocal8Bit(std::string("ס��").data())) == 0) {
				QString cshx = server->getCellData(QString::fromLocal8Bit(std::string("���л���").data()), i).toString();
				if (cshx.compare(QString::fromLocal8Bit(std::string("һ��").data())) == 0
					|| cshx.compare(QString::fromLocal8Bit(std::string("����").data())) == 0
					|| cshx.compare(QString::fromLocal8Bit(std::string("����").data())) == 0)
				{
					double hz = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
						+ constMonth + QString::fromLocal8Bit(std::string("����").data()), i).toDouble();
					double result = hz * 0.6;
					server->writedata(QVariant(result < 50000 ? result : 50000), QString::number(currentYear) + constYear
						+ QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("Ԥ�м���").data()), i);
				}
				else if (cshx.compare(QString::fromLocal8Bit(std::string("����").data())) == 0
					|| cshx.compare(QString::fromLocal8Bit(std::string("����").data())) == 0)
				{
					double hz = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
						+ constMonth + QString::fromLocal8Bit(std::string("����").data()), i).toDouble();
					double result = hz * 0.35;
					server->writedata(QVariant(result < 12000 ? result : 12000), QString::number(currentYear) + constYear
						+ QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("Ԥ�м���").data()), i);
				}
			}
		}
		else if (xmlx.compare(QString::fromLocal8Bit(std::string("������").data())) == 0) {
			QString cshx = server->getCellData(QString::fromLocal8Bit(std::string("�׿�����").data()), i).toString();
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
				QString syb = server->getCellData(QString::fromLocal8Bit(std::string("��ҵ����ס��/�̿���").data()), i).toString();
				if (syb.compare(QString::fromLocal8Bit(std::string("ס��").data())) == 0)
					radio = 0.5;
				else if (syb.compare(QString::fromLocal8Bit(std::string("�̿�").data())) == 0) {
					QString yt = server->getCellData(QString::fromLocal8Bit(std::string("ҵ̬").data()), i).toString();
					if (yt.compare(QString::fromLocal8Bit(std::string("סլ").data()))) {
						radio = 0.6;
					}
					else if (yt.compare(QString::fromLocal8Bit(std::string("����").data())) == 0) {
						radio = 0.22;
					}
					else if (yt.compare(QString::fromLocal8Bit(std::string("��Ԣ/�칫").data())) == 0) {
						radio = 0.5;
					}
					else {
						radio = 0.35;
					}
				}
				double ret1 = server->getCellData(QString::number(fyear) + constYear + QString::number(fmonth) + constMonth
					+ QString::fromLocal8Bit(std::string("ǩԼ�Ѵ��").data()), i).toDouble();
				double ret2 = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
					+ constMonth + QString::fromLocal8Bit(std::string("����").data()), i).toDouble();
				double ret3 = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
					+ constMonth + QString::fromLocal8Bit(std::string("ǩԼ�Ѵ��").data()), i).toDouble();
				double retf = max(ret3, min(ret1 * radio, ret2 * 0.8));
				server->writedata(QVariant(retf), QString::number(currentYear) + constYear + QString::number(currentMonth)
					+ constMonth + QString::fromLocal8Bit(std::string("Ԥ�м���").data()), i);
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
						+ constMonth + QString::fromLocal8Bit(std::string("����").data()), i).toDouble();
					if (ret1 == 0)
					{
						double ret2 = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
							+ constMonth + QString::fromLocal8Bit(std::string("ǩԼ�Ѵ��").data()), i).toDouble();
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
					+ QString::fromLocal8Bit(std::string("����").data()), i).toDouble();
				double dyqyydc = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth
					+ QString::fromLocal8Bit(std::string("ǩԼ�Ѵ��").data()), i).toDouble();
				if (count == 0)
					server->writedata(QVariant(dyqyydc), QString::number(currentYear) + constYear + QString::number(currentMonth)
						+ constMonth + QString::fromLocal8Bit(std::string("Ԥ�м���").data()), i);
				else
					server->writedata(QVariant(max(dyqyydc, min(dyhz, output / count))), QString::number(currentYear) + constYear
						+ QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("Ԥ�м���").data()), i);
			}
		}
		else if (xmlx.compare(QString::fromLocal8Bit(std::string("������").data())) == 0) {
			QString cshx = server->getCellData(QString::fromLocal8Bit(std::string("�׿�����").data()), i).toString();
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
				QString syb = server->getCellData(QString::fromLocal8Bit(std::string("��ҵ����ס��/�̿���").data()), i).toString();
				if (syb.compare(QString::fromLocal8Bit(std::string("ס��").data())) == 0)
					radio = 0.5;
				else if (syb.compare(QString::fromLocal8Bit(std::string("�̿�").data())) == 0) {
					QString yt = server->getCellData(QString::fromLocal8Bit(std::string("ҵ̬").data()), i).toString();
					if (yt.compare(QString::fromLocal8Bit(std::string("סլ").data()))==0) {
						radio = 0.6;
					}
					else if (yt.compare(QString::fromLocal8Bit(std::string("����").data()))==0) {
						radio = 0.22;
					}
					else if (yt.compare(QString::fromLocal8Bit(std::string("��Ԣ/�칫").data()))==0) {
						radio = 0.5;
					}
					else {
						radio = 0.35;
					}
				}

				double ret1 = server->getCellData(QString::number(fyear) + constYear + QString::number(fmonth) + constMonth
					+ QString::fromLocal8Bit(std::string("ǩԼ�Ѵ��").data()), i).toDouble();
				double ret2 = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
					+ constMonth + QString::fromLocal8Bit(std::string("����").data()), i).toDouble();
				double ret3 = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
					+ constMonth + QString::fromLocal8Bit(std::string("ǩԼ�Ѵ��").data()), i).toDouble();
				double retf = max(ret3, min(ret1 * radio, ret2 * 0.8));
				server->writedata(QVariant(retf), QString::number(currentYear) + constYear + QString::number(currentMonth) +
					constMonth + QString::fromLocal8Bit(std::string("Ԥ�м���").data()), i);
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
						+ constMonth + QString::fromLocal8Bit(std::string("����").data()), i).toDouble();
					if (ret1 != 0) {
						isok = true;
						break;
					}
					diff--;
				}
				if (isok) {
					double dyhz = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
						+ constMonth + QString::fromLocal8Bit(std::string("����").data()), i).toDouble();
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
							+ constMonth + QString::fromLocal8Bit(std::string("ȥ����").data()), i).toDouble();
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
					//��ǰ��-1
					for (int tempi = 0; tempi < circletime; tempi++) {
						dmonth -= 1;
						double data = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
							+ constMonth + QString::fromLocal8Bit(std::string("ǩԼ�Ѵ��").data()), i).toDouble();
						maxdata = maxdata > data ? maxdata : data;
						mindata = mindata < data ? mindata : data;
						alldata += data;
					}
					wyyjll = (alldata - maxdata - mindata) / (circletime - 2);
					double dyqyxdc = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
						+ constMonth + QString::fromLocal8Bit(std::string("ǩԼ�Ѵ��").data()), i).toDouble();
					server->writedata(QVariant(max(min(min(dyhz * qhljz, wyyjll), dyhz), dyqyxdc)), QString::number(currentYear)
						+ constYear + QString::number(currentMonth) + constMonth
						+ QString::fromLocal8Bit(std::string("Ԥ�м���").data()), i);
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
							+ constMonth + QString::fromLocal8Bit(std::string("ǩԼ�Ѵ��").data()), i).toDouble();
						alldata += data;
					}
					yjll = alldata / 3;

					dyhz = server->getCellData(QString::number(dyear) + constYear + QString::number(currentMonth) + constMonth
						+ QString::fromLocal8Bit(std::string("����").data()), i).toDouble();
					dyks = server->getCellData(QString::number(dyear) + constYear + QString::number(currentMonth) + constMonth
						+ QString::fromLocal8Bit(std::string("����").data()), i).toDouble();
					syb = server->getCellData(QString::fromLocal8Bit(std::string("��ҵ����ס��/�̿���").data()), i).toString();
					if (syb.compare(QString::fromLocal8Bit(std::string("ס��").data())) == 0)
					{
						QString cshx = server->getCellData(QString::fromLocal8Bit(std::string("���л���").data()), i).toString();
						if (cshx.compare(QString::fromLocal8Bit(std::string("һ��").data())) == 0
							|| cshx.compare(QString::fromLocal8Bit(std::string("����").data())) == 0
							|| cshx.compare(QString::fromLocal8Bit(std::string("����").data())) == 0) {
							radio = 0.6;
						}
						else if (cshx.compare(QString::fromLocal8Bit(std::string("����").data())) == 0
							|| cshx.compare(QString::fromLocal8Bit(std::string("����").data())) == 0) {
							radio = 0.35;
						}
					}
					else if (syb.compare(QString::fromLocal8Bit(std::string("�̿�").data())) == 0) {
						QString yt = server->getCellData(QString::fromLocal8Bit(std::string("ҵ̬").data()), i).toString();
						if (yt.compare(QString::fromLocal8Bit(std::string("סլ").data()))==0) {
							radio = 0.7;
						}
						else if (yt.compare(QString::fromLocal8Bit(std::string("����").data()))==0) {
							radio = 0.65;
						}
						else if (yt.compare(QString::fromLocal8Bit(std::string("��Ԣ/�칫").data()))==0) {
							radio = 0.35;
						}
						else {
							radio = 0.35;
						}
					}
					output = yjll + radio * dyhz;

					double dyqyydc = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth)
						+ constMonth + QString::fromLocal8Bit(std::string("ǩԼ�Ѵ��").data()), i).toDouble();

					dyear = currentYear;
					dmonth = currentMonth;
					double maxqydc = 0;
					ncount = currentMonth - 1;
					for (int tmpi = 0; tmpi < ncount; tmpi++) {
						dmonth -= 1;
						double data = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
							+ constMonth + QString::fromLocal8Bit(std::string("ǩԼ�Ѵ��").data()), i).toDouble();
						maxqydc = maxqydc > data ? maxqydc : data;
					}
					server->writedata(QVariant(max(min(min(output, maxqydc), dyks), dyqyydc)), QString::number(currentYear) + constYear
						+ QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("Ԥ�м���").data()), i);
				}
			}
		}
		else
		{
			printf("Warning : row %d missing->%s.\n", i, qPrintable(xmlx));
		}
	}
}

//��ǰ���ۺ�ȫ�ھ�������
void service::predictcurrentMonth3(ExcelDataServer* server) {
	for (int tempi = this->startRow; tempi <= this->endRow; tempi++) {
		double sumdata1 = 0;
		for (int i = 1; i <= currentMonth; i++) {
			sumdata1 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("ȫ�ھ������Ѵ��").data()), tempi).toDouble();
		}
		double qkj = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("�ڳ�ȫ�ھ�Ӧ�տ�").data()), tempi).toDouble();
		double sumdata2 = 0;
		if (currentMonth == 1) {
			sumdata2 += server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("ǩԼָ��").data()), tempi).toDouble();
		}
		else {
			for (int i = 1; i < currentMonth; i++) {
				sumdata2 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("ǩԼ�Ѵ��").data()), tempi).toDouble();
			}
			sumdata2 += server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("ǩԼָ��").data()), tempi).toDouble();
		}
		if ((qkj + sumdata2) == 0)
		{
			server->writedata(QVariant("/"), QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("�ۺ�ȫ�ھ�������").data()), tempi);
		}
		else
		{
			server->writedata(QVariant(int(sumdata1 / (qkj + sumdata2))), QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("�ۺ�ȫ�ھ�������").data()), tempi);
		}
	}
}

//��ǰ���ۺ�Ȩ��ھ�������
void service::predictcurrentMonth4(ExcelDataServer* server) {
	for (int tempi = this->startRow; tempi <= this->endRow; tempi++) {
		double sumdata1 = 0;
		for (int i = 1; i <= currentMonth; i++) {
			sumdata1 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("Ȩ��ھ������Ѵ��").data()), tempi).toDouble();
		}
		double qkj = server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("�ڳ�Ȩ��ھ�Ӧ�տ�").data()), tempi).toDouble();
		double sumdata2 = 0;
		if (currentMonth == 1) {
			sumdata2 += server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("ǩԼָ��").data()), tempi).toDouble();
		}
		else {
			for (int i = 1; i < currentMonth; i++) {
				sumdata2 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("ǩԼ�Ѵ��").data()), tempi).toDouble();
			}
			sumdata2 += server->getCellData(QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("ǩԼָ��").data()), tempi).toDouble();
		}
		double gqbl = server->getCellData(QString::fromLocal8Bit(std::string("��Ȩ����").data()), tempi).toDouble();

		if ((qkj + gqbl * sumdata2) == 0)
		{
			server->writedata(QVariant("/"), QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("�ۺ�Ȩ��ھ�������").data()), tempi);
		}
		else
		{
			server->writedata(QVariant(sumdata1 / (qkj + gqbl * sumdata2)), QString::number(currentYear) + constYear + QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("�ۺ�Ȩ��ھ�������").data()), tempi);
		}
	}
}



//---------------------------------��ȼ���------------------------------------
//���ȫ�ھ�������ɣ���ҵ���棩
void service::yearPredict1(ExcelDataServer* server) {
	for (int tempi = this->startRow; tempi <= this->endRow; tempi++) {
		double dataAll = 0;
		for (int i = 1; i<currentMonth; i++)
		{
			dataAll += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("ȫ�ھ������Ѵ��").data()), tempi).toDouble();
		}
		for (int i = currentMonth; i <= 12; i++)
		{
			dataAll += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("ȫ�ھ�Ԥ����ɣ�ָ��������").data()), tempi).toDouble();
		}
		server->writedata(QVariant(dataAll), QString::number(currentYear) + constYear + QString::fromLocal8Bit(std::string("���ȫ�ھ�������ɣ���ҵ���棩").data()), tempi);
	}
}

//���ȫ�ھ�������ɣ����ư棩
void service::yearPredict2(ExcelDataServer* server) {
	for (int tempi = this->startRow; tempi <= this->endRow; tempi++) {
		double dataAll = 0;
		for (int i = 1; i<currentMonth; i++)
		{
			dataAll += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("ȫ�ھ������Ѵ��").data()), tempi).toDouble();
		}
		for (int i = currentMonth; i <= 12; i++)
		{
			dataAll += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("����Ԥ����ɣ����ư棩").data()), tempi).toDouble();
		}
		server->writedata(QVariant(dataAll), QString::number(currentYear) + constYear + QString::fromLocal8Bit(std::string("���ȫ�ھ�������ɣ����ư棩").data()), tempi);
	}
}

//���ǩԼ��ɣ���ҵ���棩
void service::yearPredict3(ExcelDataServer* server) {
	for (int tempi = this->startRow; tempi <= this->endRow; tempi++) {
		double dataAll = 0;
		for(int i = 1;i<currentMonth;i++)
		{
			dataAll += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("ǩԼ�Ѵ��").data()), tempi).toDouble();
		}
		for (int i = currentMonth; i <= 12; i++)
		{
			dataAll += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("ǩԼԤ����ɣ�ָ��������").data()), tempi).toDouble();
		}
		server->writedata(QVariant(dataAll), QString::number(currentYear) + constYear + QString::fromLocal8Bit(std::string("���ǩԼ��ɣ���ҵ���棩").data()), tempi);
	}
}

//���ǩԼ��ɣ����ư棩
void service::yearPredict4(ExcelDataServer* server) {
	for (int tempi = this->startRow; tempi <= this->endRow; tempi++) {
		double dataAll = 0;
		for (int i = 1; i<currentMonth; i++)
		{
			dataAll += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("ǩԼ�Ѵ��").data()), tempi).toDouble();
		}
		for (int i = currentMonth; i <= 12; i++)
		{
			dataAll += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("Ԥ�м���").data()), tempi).toDouble();
		}
		server->writedata(QVariant(dataAll), QString::number(currentYear) + constYear + QString::fromLocal8Bit(std::string("���ǩԼ��ɣ����ư棩").data()), tempi);
	}
}

//���Ȩ��ھ�������ɣ���ҵ���棩
void service::yearPredict5(ExcelDataServer* server) {
	for (int tempi = this->startRow; tempi <= this->endRow; tempi++) {
		double dataAll = 0;
		for (int i = 1; i<currentMonth; i++)
		{
			dataAll += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("Ȩ��ھ������Ѵ��").data()), tempi).toDouble();
		}
		for (int i = currentMonth; i <= 12; i++)
		{
			dataAll += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("Ȩ��ھ�Ԥ����ɣ�ָ��������").data()), tempi).toDouble();
		}
		server->writedata(QVariant(dataAll), QString::number(currentYear) + constYear + QString::fromLocal8Bit(std::string("���Ȩ��ھ�������ɣ���ҵ���棩").data()), tempi);
	}
}

//���Ȩ��ھ�������ɣ����ư棩
void service::yearPredict6(ExcelDataServer* server) {
	for (int tempi = this->startRow; tempi <= this->endRow; tempi++) {
		double dataAll = 0;
		for (int i = 1; i<currentMonth; i++)
		{
			dataAll += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("Ȩ��ھ������Ѵ��").data()), tempi).toDouble();
		}
		for (int i = currentMonth; i <= 12; i++)
		{
			dataAll += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("����ָ�꣨���ư棩").data()), tempi).toDouble();
		}
		server->writedata(QVariant(dataAll), QString::number(currentYear) + constYear + QString::fromLocal8Bit(std::string("���Ȩ��ھ�������ɣ����ư棩").data()), tempi);
	}
}


//-------------------------------------δ����--------------------------------------
//δ�����ۺ�Ȩ��ھ�������
void service::predictNextMonth1(ExcelDataServer* server) {
	for (int nextmonth = currentMonth + 1; nextmonth <= 12; nextmonth++) {
		for (int tempi = this->startRow; tempi <= this->endRow; tempi++) {
			double sumdata1 = 0;
			for (int i = 1; i <= currentMonth; i++) {
				sumdata1 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("Ȩ��ھ������Ѵ��").data()), tempi).toDouble();
			}
			for (int i = currentMonth+1; i <= 12; i++) {
				sumdata1 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("����ָ�꣨���ư棩").data()), tempi).toDouble();
			}
			double qkj = server->getCellData(QString::number(currentYear) + constYear + QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("�ڳ�Ȩ��ھ�Ӧ�տ�").data()), tempi).toDouble();
			double sumdata2 = 0;
			for (int i = 1; i < currentMonth; i++) {
				sumdata2 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("ǩԼ�Ѵ��").data()), tempi).toDouble();
			}
			for (int i = currentMonth; i <= 12; i++) {
				sumdata2 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("ǩԼָ��").data()), tempi).toDouble();
			}
			double gqbl = server->getCellData(QString::fromLocal8Bit(std::string("��Ȩ����").data()), tempi).toDouble();

			if ((qkj + sumdata2*gqbl) == 0)
			{
				server->writedata(QVariant("/"), QString::number(currentYear) + constYear + QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("�ۺ�Ȩ��ھ�������").data()), tempi);
			}
			else
			{
				server->writedata(QVariant(sumdata1 / (qkj + sumdata2*gqbl)), QString::number(currentYear) + constYear + QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("�ۺ�Ȩ��ھ�������").data()), tempi);
			}
		}
	}
}

//δ����Ȩ��ھ��������ȱ�
void service::predictNextMonth2(ExcelDataServer* server) {
	for (int nextmonth = currentMonth + 1; nextmonth <= 12; nextmonth++) {
		for (int tempi = this->startRow; tempi <= this->endRow; tempi++) {
			double sumdata1 = 0;
			for (int i = 1; i <currentMonth; i++) {
				sumdata1 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("Ȩ��ھ������Ѵ��").data()), tempi).toDouble();
			}
			for (int i = currentMonth; i <= 12; i++) {
				sumdata1 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("Ȩ��ھ�����ָ��").data()), tempi).toDouble();
			}
			double ndqykj = server->getCellData(QString::number(currentYear) + constYear  + QString::fromLocal8Bit(std::string("���Ȩ��ھ�����ָ��").data()), tempi).toDouble();
			server->writedata(QVariant(sumdata1 / ndqykj), QString::number(currentYear) + constYear + QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("Ȩ��ھ��������ȱ�").data()), tempi);
		}
	}
}

//δ�����ۺ�ȫ�ھ�������
void service::predictNextMonth3(ExcelDataServer* server) {
	for (int nextmonth = currentMonth + 1; nextmonth <= 12; nextmonth++) {
		for (int tempi = this->startRow; tempi <= this->endRow; tempi++) {
			double sumdata1 = 0;
			for (int i = 1; i <= currentMonth; i++) {
				sumdata1 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("ȫ�ھ������Ѵ��").data()), tempi).toDouble();
			}
			for (int i = currentMonth+1; i <= 12; i++) {
				sumdata1 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("����Ԥ����ɣ����ư棩").data()), tempi).toDouble();
			}
			double qkj = server->getCellData(QString::number(currentYear) + constYear + QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("��ȫ�ھ�Ӧ�տ�").data()), tempi).toDouble();
			double sumdata2 = 0;
			for (int i = 1; i <currentMonth; i++) {
				sumdata2 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("ǩԼ�Ѵ��").data()), tempi).toDouble();
			}
			for (int i = currentMonth; i <= 12; i++) {
				sumdata2 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("ǩԼָ��").data()), tempi).toDouble();
			}
			if ((qkj + sumdata2) == 0)
			{
				server->writedata(QVariant("/"), QString::number(currentYear) + constYear + QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("�ۺ�ȫ�ھ�������").data()), tempi);
			}
			else
			{
				server->writedata(QVariant(int(sumdata1 / (qkj + sumdata2))), QString::number(currentYear) + constYear + QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("�ۺ�ȫ�ھ�������").data()), tempi);
			}
		}
	}
}

//δ����ȫ�ھ��������ȱ�
void service::predictNextMonth4(ExcelDataServer* server) {
	for (int nextmonth = currentMonth + 1; nextmonth <= 12; nextmonth++) {
		for (int tempi = this->startRow; tempi <= this->endRow; tempi++) {
			double sumdata1 = 0;
			for (int i = 1; i <currentMonth; i++) {
				sumdata1 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("ȫ�ھ������Ѵ��").data()), tempi).toDouble();
			}
			for (int i = currentMonth; i <= 12; i++) {
				sumdata1 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("ȫ�ھ�����ָ��").data()), tempi).toDouble();
			}
			double ndqykj = server->getCellData(QString::number(currentYear) + constYear + QString::fromLocal8Bit(std::string("���ȫ�ھ�����ָ��").data()), tempi).toDouble();
			server->writedata(QVariant(sumdata1 / ndqykj), QString::number(currentYear) + constYear + QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("ȫ�ھ��������ȱ�").data()), tempi);
		}
	}
}

//δ���»���Ԥ����ɣ����ư棩
void service::predictNextMonth5(ExcelDataServer* server) {
	for (int nextmonth = currentMonth + 1; nextmonth <= 12; nextmonth++) {
		for (int tempi = this->startRow; tempi <= this->endRow; tempi++) {
			double ypjs = server->getCellData(QString::number(currentYear) + constYear + QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("Ԥ�м���").data()), tempi).toDouble();
			int count = 0;
			double alldata = 0;
			for (int i = currentMonth - 1; i > currentMonth-4; i--) {
				double hlqyb = server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("����ǩԼ��").data()), tempi).toDouble();
				if (hlqyb > 0.00001) {
					alldata += hlqyb;
					count++;
				}
			}
			if (count == 0) {
				server->writedata(QVariant(0), QString::number(currentYear) + constYear + QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("����Ԥ����ɣ����ư棩").data()), tempi);
			}
			else {
				server->writedata(QVariant(ypjs*alldata/count), QString::number(currentYear) + constYear + QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("����Ԥ����ɣ����ư棩").data()), tempi);
			}
		}
	}
}

//δ������Ŀ����
void service::predictNextMonth6(ExcelDataServer* server) {
	for (int nextmonth = currentMonth + 1; nextmonth <= 12; nextmonth++) {
		for (int tempi = this->startRow; tempi <= this->endRow; tempi++) {
			double ndqydc = server->getCellData(QString::number(currentYear) + constYear
				+ QString::fromLocal8Bit(std::string("���ǩԼ�Ѵ��").data()), tempi).toDouble();

			double dyqyydc = server->getCellData(QString::number(currentYear) + constYear
				+ QString::number(nextmonth) + constMonth +
				QString::fromLocal8Bit(std::string("ǩԼ�Ѵ��").data()), tempi).toDouble();

			double dyqckc = server->getCellData(QString::number(currentYear) + constYear +
				QString::number(nextmonth) + constMonth +
				QString::fromLocal8Bit(std::string("�ڳ����").data()), tempi).toDouble();

			double dygh = server->getCellData(QString::number(currentYear) + constYear +
				QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("����").data()), tempi).toDouble();
			if (ndqydc - dyqyydc == 0 && dyqckc == 0 && dygh> 0)
			{
				server->writedata(QVariant(QString::fromLocal8Bit(std::string("�׿���").data())),
					QString::number(currentYear) + constYear + QString::number(nextmonth) + constMonth +
					QString::fromLocal8Bit(std::string("��Ŀ����").data()), tempi);
			}
			else if (dyqckc > 0 && dygh> 0) {
				server->writedata(QVariant(QString::fromLocal8Bit(std::string("������").data())),
					QString::number(currentYear) + constYear + QString::number(nextmonth) + constMonth +
					QString::fromLocal8Bit(std::string("��Ŀ����").data()), tempi);
			}
			else if (ndqydc - dyqyydc > 0 && dygh == 0 || dyqckc > 0 && dygh == 0) {
				server->writedata(QVariant(QString::fromLocal8Bit(std::string("������").data())),
					QString::number(currentYear) + constYear + QString::number(nextmonth) + constMonth +
					QString::fromLocal8Bit(std::string("��Ŀ����").data()), tempi);
			}
		}
	}
}

//δ����ǩԼ�����Ƚ��ȱ�
void service::predictNextMonth7(ExcelDataServer* server) {
	for (int nextmonth = currentMonth + 1; nextmonth <= 12; nextmonth++) {
		for (int tempi = this->startRow; tempi <= this->endRow; tempi++) {
			double sumdata1 = 0;
			for (int i = 1; i <currentMonth; i++) {
				sumdata1 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("ǩԼ�Ѵ��").data()), tempi).toDouble();
			}
			for (int i = currentMonth; i <= 12; i++) {
				sumdata1 += server->getCellData(QString::number(currentYear) + constYear + QString::number(i) + constMonth + QString::fromLocal8Bit(std::string("ǩԼָ��").data()), tempi).toDouble();
			}
			double ndqyzb = server->getCellData(QString::number(currentYear) + constYear + QString::fromLocal8Bit(std::string("���ǩԼָ��").data()), tempi).toDouble();
			server->writedata(QVariant(sumdata1 / ndqyzb), QString::number(currentYear) + constYear + QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("ǩԼ�����Ƚ��ȱ�").data()), tempi);
		}
	}
}

//δ����Ԥ�м���
void service::predictNextMonth8(ExcelDataServer* server) {
	for (int nextmonth = currentMonth + 1; nextmonth <= 12; nextmonth++) {
		for (int tempi = this->startRow; tempi <= this->endRow; tempi++) {
			QString xmlx = server->getCellData(QString::number(currentYear) + constYear + QString::number(nextmonth)
				+ constMonth + QString::fromLocal8Bit(std::string("��Ŀ����").data()), tempi).toString();
			if (xmlx.compare(QString::fromLocal8Bit(std::string("�׿���").data())) == 0) {
				QString syb = server->getCellData(QString::fromLocal8Bit(std::string("��ҵ����ס��/�̿���").data()), tempi).toString();
				if (syb.compare(QString::fromLocal8Bit(std::string("�̿�").data())) == 0) {
					QString yt = server->getCellData(QString::fromLocal8Bit(std::string("ҵ̬").data()), tempi).toString();
					if (yt.compare(QString::fromLocal8Bit(std::string("סլ").data())) == 0) {
						double hz = server->getCellData(QString::number(currentYear) + constYear + QString::number(nextmonth)
							+ constMonth + QString::fromLocal8Bit(std::string("����").data()), tempi).toDouble();
						double result = hz * 0.7;
						server->writedata(QVariant(result < 50000 ? result : 50000), QString::number(currentYear) + constYear
							+ QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("Ԥ�м���").data()), tempi);
					}
					else if (yt.compare(QString::fromLocal8Bit(std::string("����").data())) == 0) {
						double hz = server->getCellData(QString::number(currentYear) + constYear + QString::number(nextmonth)
							+ constMonth + QString::fromLocal8Bit(std::string("����").data()), tempi).toDouble();
						double result = hz * 0.65;
						server->writedata(QVariant(result < 50000 ? result : 50000), QString::number(currentYear) + constYear
							+ QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("Ԥ�м���").data()), tempi);
					}
					else if (yt.compare(QString::fromLocal8Bit(std::string("��Ԣ/�칫").data())) == 0) {
						double hz = server->getCellData(QString::number(currentYear) + constYear + QString::number(nextmonth)
							+ constMonth + QString::fromLocal8Bit(std::string("����").data()), tempi).toDouble();
						double result = hz * 0.35;
						server->writedata(QVariant(result < 50000 ? result : 50000), QString::number(currentYear) + constYear
							+ QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("Ԥ�м���").data()), tempi);
					}
					else {
						double hz = server->getCellData(QString::number(currentYear) + constYear + QString::number(nextmonth)
							+ constMonth + QString::fromLocal8Bit(std::string("����").data()), tempi).toDouble();
						double result = hz * 0.35;
						server->writedata(QVariant(result < 50000 ? result : 50000), QString::number(currentYear) + constYear
							+ QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("Ԥ�м���").data()), tempi);
					}
				}
				else if (syb.compare(QString::fromLocal8Bit(std::string("ס��").data())) == 0) {
					QString cshx = server->getCellData(QString::fromLocal8Bit(std::string("���л���").data()), tempi).toString();
					if (cshx.compare(QString::fromLocal8Bit(std::string("һ��").data())) == 0
						|| cshx.compare(QString::fromLocal8Bit(std::string("����").data())) == 0
						|| cshx.compare(QString::fromLocal8Bit(std::string("����").data())) == 0)
					{
						double hz = server->getCellData(QString::number(currentYear) + constYear + QString::number(nextmonth)
							+ constMonth + QString::fromLocal8Bit(std::string("����").data()), tempi).toDouble();
						double result = hz * 0.6;
						server->writedata(QVariant(result < 50000 ? result : 50000), QString::number(currentYear) + constYear
							+ QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("Ԥ�м���").data()), tempi);
					}
					else if (cshx.compare(QString::fromLocal8Bit(std::string("����").data())) == 0
						|| cshx.compare(QString::fromLocal8Bit(std::string("����").data())) == 0)
					{
						double hz = server->getCellData(QString::number(currentYear) + constYear + QString::number(nextmonth)
							+ constMonth + QString::fromLocal8Bit(std::string("����").data()), tempi).toDouble();
						double result = hz * 0.35;
						server->writedata(QVariant(result < 12000 ? result : 12000), QString::number(currentYear) + constYear
							+ QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("Ԥ�м���").data()), tempi);
					}
				}
			}
			else if (xmlx.compare(QString::fromLocal8Bit(std::string("������").data())) == 0) {
				QString cshx = server->getCellData(QString::fromLocal8Bit(std::string("�׿�����").data()), tempi).toString();
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
					QString syb = server->getCellData(QString::fromLocal8Bit(std::string("��ҵ����ס��/�̿���").data()), tempi).toString();
					if (syb.compare(QString::fromLocal8Bit(std::string("ס��").data())) == 0)
						radio = 0.5;
					else if (syb.compare(QString::fromLocal8Bit(std::string("�̿�").data())) == 0) {
						QString yt = server->getCellData(QString::fromLocal8Bit(std::string("ҵ̬").data()), tempi).toString();
						if (yt.compare(QString::fromLocal8Bit(std::string("סլ").data()))== 0) {
							radio = 0.6;
						}
						else if (yt.compare(QString::fromLocal8Bit(std::string("����").data()))== 0) {
							radio = 0.22;
						}
						else if (yt.compare(QString::fromLocal8Bit(std::string("��Ԣ/�칫").data()))== 0) {
							radio = 0.5;
						}
						else {
							radio = 0.35;
						}
					}
					double ret1;
					if(fmonth<currentMonth)
						ret1 = server->getCellData(QString::number(fyear) + constYear + QString::number(fmonth) + constMonth
						+ QString::fromLocal8Bit(std::string("ǩԼ�Ѵ��").data()), tempi).toDouble();
					else
						ret1 = server->getCellData(QString::number(fyear) + constYear + QString::number(fmonth) + constMonth
							+ QString::fromLocal8Bit(std::string("Ԥ�м���").data()), tempi).toDouble();
					double ret2 = server->getCellData(QString::number(currentYear) + constYear + QString::number(nextmonth)
						+ constMonth + QString::fromLocal8Bit(std::string("����").data()), tempi).toDouble();
					double retf = min(ret1 * radio, ret2 * 0.8);
					server->writedata(QVariant(retf), QString::number(currentYear) + constYear + QString::number(nextmonth)
						+ constMonth + QString::fromLocal8Bit(std::string("Ԥ�м���").data()), tempi);
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
							+ constMonth + QString::fromLocal8Bit(std::string("����").data()), tempi).toDouble();
						if (ret1 == 0)
						{
							double ret2;
							if(dmonth<currentMonth)
								ret2 = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
								+ constMonth + QString::fromLocal8Bit(std::string("ǩԼ�Ѵ��").data()), tempi).toDouble();
							else
								ret2 = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
									+ constMonth + QString::fromLocal8Bit(std::string("Ԥ�м���").data()), tempi).toDouble();
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
						+ QString::fromLocal8Bit(std::string("����").data()), tempi).toDouble();
					if (count == 0)
						server->writedata(QVariant(0), QString::number(currentYear) + constYear + QString::number(nextmonth)
							+ constMonth + QString::fromLocal8Bit(std::string("Ԥ�м���").data()), tempi);
					else
						server->writedata(QVariant(min(dyhz, output / count)), QString::number(currentYear) + constYear
							+ QString::number(nextmonth) + constMonth + QString::fromLocal8Bit(std::string("Ԥ�м���").data()), tempi);
				}
			}
			else if (xmlx.compare(QString::fromLocal8Bit(std::string("������").data())) == 0) {
				QString cshx = server->getCellData(QString::fromLocal8Bit(std::string("�׿�����").data()), tempi).toString();
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
					QString syb = server->getCellData(QString::fromLocal8Bit(std::string("��ҵ����ס��/�̿���").data()), tempi).toString();
					if (syb.compare(QString::fromLocal8Bit(std::string("ס��").data())) == 0)
						radio = 0.5;
					else if (syb.compare(QString::fromLocal8Bit(std::string("�̿�").data())) == 0) {
						QString yt = server->getCellData(QString::fromLocal8Bit(std::string("ҵ̬").data()), tempi).toString();
						if (yt.compare(QString::fromLocal8Bit(std::string("סլ").data())) ==0) {
							radio = 0.6;
						}
						else if (yt.compare(QString::fromLocal8Bit(std::string("����").data())) == 0) {
							radio = 0.22;
						}
						else if (yt.compare(QString::fromLocal8Bit(std::string("��Ԣ/�칫").data())) ==0) {
							radio = 0.5;
						}
						else {
							radio = 0.35;
						}
					}
					double ret1;
					if(currentMonth<=fmonth && currentYear == fyear)
						ret1 = server->getCellData(QString::number(fyear) + constYear + QString::number(fmonth) + constMonth
						+ QString::fromLocal8Bit(std::string("Ԥ�м���").data()), tempi).toDouble();
					else
						ret1 = server->getCellData(QString::number(fyear) + constYear + QString::number(fmonth) + constMonth
							+ QString::fromLocal8Bit(std::string("ǩԼ�Ѵ��").data()), tempi).toDouble();
					double ret2 = server->getCellData(QString::number(currentYear) + constYear + QString::number(nextmonth)
						+ constMonth + QString::fromLocal8Bit(std::string("����").data()), tempi).toDouble();
					double retf =  min(ret1 * radio, ret2 * 0.8);
					server->writedata(QVariant(retf), QString::number(currentYear) + constYear + QString::number(nextmonth) +
						constMonth + QString::fromLocal8Bit(std::string("Ԥ�м���").data()), tempi);
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
							+ constMonth + QString::fromLocal8Bit(std::string("����").data()), tempi).toDouble();
						if (ret1 != 0) {
							isok = true;
							break;
						}
						diff--;
					}
					if (isok) {
						double dyhz = server->getCellData(QString::number(currentYear) + constYear + QString::number(nextmonth)
							+ constMonth + QString::fromLocal8Bit(std::string("����").data()), tempi).toDouble();
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
								+ constMonth + QString::fromLocal8Bit(std::string("ȥ����").data()), tempi).toDouble();
							else
							{ 
								double ypjs = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
									+ constMonth + QString::fromLocal8Bit(std::string("Ԥ�м���").data()), tempi).toDouble();
								double ks = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
								+ constMonth + QString::fromLocal8Bit(std::string("����").data()), tempi).toDouble();
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
						//��ǰ��-1
						for (int cireclei = 0; cireclei < circletime ; cireclei++) {
							dmonth -= 1;
							double data;
							if(dmonth<currentMonth)
								data = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
									+ constMonth + QString::fromLocal8Bit(std::string("ǩԼ�Ѵ��").data()), tempi).toDouble();
							else {
								data = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
									+ constMonth + QString::fromLocal8Bit(std::string("Ԥ�м���").data()), tempi).toDouble();
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
							+ QString::fromLocal8Bit(std::string("Ԥ�м���").data()), tempi);
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
								+ constMonth + QString::fromLocal8Bit(std::string("ǩԼ�Ѵ��").data()), tempi).toDouble();
							else
								data = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
									+ constMonth + QString::fromLocal8Bit(std::string("Ԥ�м���").data()), tempi).toDouble();
							alldata += data;
						}
						yjll = alldata / 3;

						dyhz = server->getCellData(QString::number(dyear) + constYear + QString::number(nextmonth) + constMonth
							+ QString::fromLocal8Bit(std::string("����").data()), tempi).toDouble();
						dyks = server->getCellData(QString::number(dyear) + constYear + QString::number(nextmonth) + constMonth
							+ QString::fromLocal8Bit(std::string("����").data()), tempi).toDouble();
						syb = server->getCellData(QString::fromLocal8Bit(std::string("��ҵ����ס��/�̿���").data()), tempi).toString();
						if (syb.compare(QString::fromLocal8Bit(std::string("ס��").data())) == 0)
						{
							QString cshx = server->getCellData(QString::fromLocal8Bit(std::string("���л���").data()), tempi).toString();
							if (cshx.compare(QString::fromLocal8Bit(std::string("һ��").data())) == 0
								|| cshx.compare(QString::fromLocal8Bit(std::string("����").data())) == 0
								|| cshx.compare(QString::fromLocal8Bit(std::string("����").data())) == 0) {
								radio = 0.6;
							}
							else if (cshx.compare(QString::fromLocal8Bit(std::string("����").data())) == 0
								|| cshx.compare(QString::fromLocal8Bit(std::string("����").data())) == 0) {
								radio = 0.35;
							}
						}
						else if (syb.compare(QString::fromLocal8Bit(std::string("�̿�").data())) == 0) {
							QString yt = server->getCellData(QString::fromLocal8Bit(std::string("ҵ̬").data()), tempi).toString();
							if (yt.compare(QString::fromLocal8Bit(std::string("סլ").data()))==0) {
								radio = 0.7;
							}
							else if (yt.compare(QString::fromLocal8Bit(std::string("����").data()))==0) {
								radio = 0.65;
							}
							else if (yt.compare(QString::fromLocal8Bit(std::string("��Ԣ/�칫").data()))==0) {
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
								+ constMonth + QString::fromLocal8Bit(std::string("ǩԼ�Ѵ��").data()), tempi).toDouble();
							else
								data = server->getCellData(QString::number(dyear) + constYear + QString::number(dmonth)
									+ constMonth + QString::fromLocal8Bit(std::string("Ԥ�м���").data()), tempi).toDouble();
							maxqydc = maxqydc > data ? maxqydc : data;
						}
						server->writedata(QVariant(min(min(output, maxqydc), dyks)), QString::number(currentYear) + constYear
							+ QString::number(currentMonth) + constMonth + QString::fromLocal8Bit(std::string("Ԥ�м���").data()), tempi);
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

//�����ύ
void service::confirm(ExcelDataServer* excelServer) {
	//19��1���ڳ����+19��1�¹���     19��1�¿���
	calHistary(QString::fromLocal8Bit(std::string("./input/1��ʷ��-����.txt").data()), excelServer);

	calCurrent(QString::fromLocal8Bit(std::string("./input/2��ǰ��-����.txt").data()), excelServer);

	calFutureMonth(QString::fromLocal8Bit(std::string("./input/3δ����-����.txt").data()), excelServer);

	calYear(QString::fromLocal8Bit(std::string("./input/4���-����.txt").data()), excelServer);

	calHistary(QString::fromLocal8Bit(std::string("./input/5��ʷ��-���.txt").data()), excelServer);
	
	calCurrent(QString::fromLocal8Bit(std::string("./input/6��ǰ��-���.txt").data()), excelServer);


	//δ�����ۺ�ȫ�ھ�������
	predictNextMonth3(excelServer);
	//δ�����ۺ�Ȩ��ھ�������
	predictNextMonth1(excelServer);
	//δ����ǩԼ�����Ƚ��ȱ�
	predictNextMonth7(excelServer);
	//δ����ȫ�ھ��������ȱ�
	predictNextMonth4(excelServer);
	//δ����Ȩ��ھ��������ȱ�
	predictNextMonth2(excelServer);


	//���ǩԼ��ɣ���ҵ���棩
	yearPredict3(excelServer);
	//���ȫ�ھ�������ɣ���ҵ���棩
	yearPredict1(excelServer);
	//���Ȩ��ھ�������ɣ���ҵ���棩
	yearPredict5(excelServer);
	//����ǰ�£��ۺ�ȫ�ھ�������
	predictcurrentMonth3(excelServer);
	//����ǰ�£��ۺ�Ȩ��ھ�������
	predictcurrentMonth4(excelServer);


	//��Ŀ����
	predictcurrentMonth1(excelServer);
	//Ԥ��
	predictcurrentMonth2(excelServer);


	calCurrent(QString::fromLocal8Bit(std::string("./input/10-��ǰ��-����ͻ���Ԥ��.txt").data()), excelServer);
	calCurrent(QString::fromLocal8Bit(std::string("./input/11��ǰ��-����.txt").data()), excelServer);


	//��Ŀ����(0)
	predictNextMonth6(excelServer);
	//Ԥ�м���(1)
	predictNextMonth8(excelServer);
	//����Ԥ����ɣ����ư棩
	predictNextMonth5(excelServer);


	calFutureMonth(QString::fromLocal8Bit(std::string("./input/13δ����-����.txt").data()), excelServer);


	//���ǩԼ��ɣ����ư棩
	yearPredict4(excelServer);
	//���ȫ�ھ�������ɣ����ư棩
	yearPredict2(excelServer);
	//���Ȩ��ھ�������ɣ����ư棩
	yearPredict6(excelServer);


	calYear(QString::fromLocal8Bit(std::string("./input/15���-����.txt").data()), excelServer);
	//******************************************************************************************8*******
}
