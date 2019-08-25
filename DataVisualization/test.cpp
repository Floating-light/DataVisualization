#include <QtCore/QCoreApplication>
#include <QFile>
#include <qdebug.h>
#include "QDate"
#include <string>
#include<vector>
#include<stack>
using namespace std;

QDate D1 = QDate::currentDate();
int startyear = 2019;
int currentyear = D1.year();//获取年
int currentmonth = D1.month();
const QString constyear = QString::fromLocal8Bit(std::string("年").data());
const QString constmonth = QString::fromLocal8Bit(std::string("月").data());

//替换单个公式中的 年 月
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
				v[i] = QString::number(year) + constyear + v[i].mid(1,-1);
			}
			else if (v[i].startsWith("#")) {
				if (month == 1 && year == startyear) {
					return "";
				}
				else {
					v[i] = QString::number(year) + constyear + QString::number(month-1) + constmonth + v[i].mid(1, -1);
				}
			}
			else if (v[i].startsWith("$")) {
				if (month == 1) {
					v[i] = QString::number(year) + constyear + QString::number(month) + constmonth + v[i].mid(1, -1);
				}
				else {
					QString temp = "(";
					for (int tempi = 1; tempi <= month; tempi++) {
						temp.append(QString::number(year) + constyear + QString::number(tempi) + constmonth + v[i].mid(1, -1) + "+");
					}
					temp.chop(1);
					v[i] = temp + ")";
				}	
			}
			else if(v[i].startsWith("%")){
				QString temp = "(";
				for (int tempi = month-2; tempi <= month; tempi++) {
					temp.append(QString::number(year) + constyear + QString::number(i) + constmonth + v[i].mid(1, -1) + "+");
				}
				temp.chop(1);
				v[i] = temp + ")/3";
			}
			else {
				v[i] = QString::number(year) + constyear + QString::number(month) + constmonth + v[i];
			}
		}
	}
	return v.join("");
}

//
QString expand(QString v,int month,int year) {
	QStringList list = v.split(" ");
	return replace(list, month,year);
}

void calCurrent(QString filepath) {
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
		QString longstring = expand(rec, currentmonth, currentyear);
		if (!longstring.isEmpty()){
			printf("%s\t", qPrintable(longstring));	
			printf("%s\n", qPrintable(QString::number(currentyear) + constyear + QString::number(currentmonth) + constmonth + label));
		}
	}
}

void calHistary(QString filepath){
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
			QString longstring = expand(rec, i, currentyear);
			if (!longstring.isEmpty()) {
				printf("%s\t", qPrintable(longstring));
				printf("%s\n", qPrintable(QString::number(currentyear) + constyear + QString::number(i) + constmonth + label));
			}
		}
	}
}

void tocal(QString expression,QString name) {
	
	stack<QChar> opera;
	vector<int> numcnt;
	QString s1;//后缀表达式
			  //中缀表达式转后缀表达式
	QVector<QChar> oprates{ '-','+','*','/','(',')' };
	for (int i = 0; i<expression.size(); i++)
	{
		//is a word
		if(!oprates.contains(expression[i]))
		{
			int tmp = 0;
			while(!oprates.contains(expression[i]))
			{
				tmp++;
				s1 += expression[i];
				i++;
			}
			i--;
			numcnt.push_back(tmp);
		}
		else if (expression[i] == '-' || expression[i] == '+')
		{
			if (expression[i] == '-' && (expression[i - 1] == '(' || expression[i - 1] == '[' || expression[i - 1] == '{'))
				s1 += '0';
			while (!opera.empty() && (opera.top() == '*' || opera.top() == '/' || opera.top() == '+' || opera.top() == '-'))
			{
				s1 += opera.top();
				opera.pop();
			}
			opera.push(expression[i]);
		}
		else if (expression[i] == '*' || expression[i] == '/')
		{
			while (!opera.empty() && (opera.top() == '*' || opera.top() == '/'))
			{
				s1 += opera.top();
				opera.pop();
			}
			opera.push(expression[i]);
		}
		else if (expression[i] == '(' || expression[i] == '[' || expression[i] == '{')
			opera.push(expression[i]);
		else if (expression[i] == ')')
		{
			while (opera.top() != '(')
			{
				s1 += opera.top();
				opera.pop();
			}
			opera.pop();
		}
		else if (expression[i] == ']')
		{
			while (opera.top() != '[')
			{
				s1 += opera.top();
				opera.pop();
			}
			opera.pop();
		}
		else if (expression[i] == '}')
		{
			while (opera.top() != '{')
			{
				s1 += opera.top();
				opera.pop();
			}
			opera.pop();
		}
		else
			qDebug() << "Invalid input!" << endl;
	}
	while (!opera.empty())
	{
		s1 += opera.top();
		opera.pop();
	}
	//计算后缀表达式的值
	stack<int> nums;
	int ind = 0;
	for (int i = 0; i<s1.size(); i++)
	{
		if (s1[i] >= '0'&&s1[i] <= '9')
		{
			int total = 0;
			while (numcnt[ind]--)
				total = 10 * total + (s1[i++] - '0');
			i--;
			nums.push(total);
			ind++;
		}
		else
		{
			int tmp1 = nums.top();
			nums.pop();
			int tmp2 = nums.top();
			nums.pop();
			if (s1[i] == '+')
				nums.push(tmp2 + tmp1);
			else if (s1[i] == '-')
				nums.push(tmp2 - tmp1);
			else if (s1[i] == '*')
				nums.push(tmp2*tmp1);
			else
				nums.push(tmp2 / tmp1);
		}
	}
	nums.top();
}

//int main(int argc, char *argv[])
//{
//	QCoreApplication a(argc, argv);
//
//	calHistary(QString::fromLocal8Bit(std::string("C:\\Users\\PIS\\Desktop\\历史月.txt").data()));
//
//	printf("%s\n", "--------------------");
//	calCurrent(QString::fromLocal8Bit(std::string("C:\\Users\\PIS\\Desktop\\当前月.txt").data()));
//	return a.exec();
//}