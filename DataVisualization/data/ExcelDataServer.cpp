#include "data/ExcelDataServer.h"
#include <stack>
#include <math.h>
ExcelDataServer::ExcelDataServer()
{
	initExcelAppAndWorksheet();
}
ExcelDataServer::~ExcelDataServer()
{

}

void ExcelDataServer::initExcelAppAndWorksheet()
{
	excelApp = new QAxObject("Excel.Application");
	//excelApp->setProperty("Visible", false); //���ش򿪵�excel�ļ�����
	excelApp->setProperty("DisplayAlerts", false);//����ʾ����
	excelWorkbooks = excelApp->querySubObject("WorkBooks");//�ɴ򿪶��excel
}

//�ͷ�Excel����
void ExcelDataServer::freeExcel()
{
	excelWorkbooks->dynamicCall("Close()");
	excelApp->dynamicCall("Quit()");
}

QAxObject* ExcelDataServer::openExcelFile(const QString& filePath)
{
	return excelWorkbooks->querySubObject("Open(QString, QVariant)", filePath);
}

QAxObject* ExcelDataServer::getSheet(QAxObject* workbook, int i)
{
	QAxObject* sheets = workbook->querySubObject("Sheets");
	currentWorksheets = sheets;
	return sheets->querySubObject("Item(int)", i);
}

//range: "A9:A100"
void ExcelDataServer::getColum(QAxObject* sheet, QString colum, QVariant& data)
{
	QAxObject* allEnvData = currentWorksheet->querySubObject("Range(QString)", colum);
	data = allEnvData->property("Value");
}

void ExcelDataServer::getColum(QAxObject* sheet, QString colum, std::vector<double>& data)
{
	QAxObject* allEnvData = sheet->querySubObject("Range(QString)", colum);
	
	QVariant variantData = allEnvData->property("Value");
	QVariantList listVariant = variantData.toList();

	//��i��
	for (int i = 0; i < listVariant.size(); ++i)
	{
		QVariantList lastList= listVariant[i].toList();
		for (int j = 0; j < lastList.size(); ++j)
		{
			if (lastList[j].toString() == "/")
			{
				data.push_back(0);
			}
			else
			{
				data.push_back(lastList[j].toDouble());
			}
			
		}
	}
}

//fast version
void ExcelDataServer::getColum(QAxObject* sheet, int column, std::vector<double>& data)
{
	int end = endRow.toInt();
	int begin = beginRow.toInt() - 1;
	for (int r = begin; r < end; ++r)
	{
		
		if (sheetContent[r][column].toString() == "/")
		{
			data.push_back(0);
		}
		else
		{
			data.push_back(sheetContent[r][column].toDouble());
		}
	}
}

void ExcelDataServer::writeAllData(QAxObject* sheet, const QString& colum, const QList<QList<QVariant>>& res)
{
	int row = res.size();
	int col = res.at(0).size();
	QVariant var;
	castListListVariant2Variant(var, res);
	writeColum(sheet, colum + "9:" + colum + QString::number(row + 8), var);

}

void ExcelDataServer::writeArea(QAxObject* tragetSheet, const QList<QList<QVariant>>& res)
{
	int row = res.size();
	int col = res.at(0).size();

	QString rangStr;
	int2Alphabet(col, rangStr);
	rangStr += QString::number(row);
	rangStr = "A1:" + rangStr;

	QAxObject* range = tragetSheet->querySubObject("Range(const QString&)", rangStr);
	if (NULL == range || range->isNull())
	{
		return ;
	}
	bool succ = false;
	QVariant var;
	castListListVariant2Variant(var, res);
	succ = range->setProperty("Value", var);
	delete range;
}

void ExcelDataServer::writeColumData(QAxObject* sheet, const QString& colum, const std::vector<double>& res)
{
	int row = res.size();
	QVariant var;
	castDoubleVector2Variant(var, res);
	writeColum(sheet, colum + beginRow+":" + colum + QString::number(row + beginRow.toInt() - 1), var);
}

void ExcelDataServer::writeColumData(const QString& colum, const std::vector<double>& res)
{
	int row = res.size();
	QVariant var;
	castDoubleVector2Variant(var, res);
	writeColum(currentWorksheet, colum + beginRow + ":" + colum + QString::number(row + beginRow.toInt() - 1), var);
}

//fast write column
void ExcelDataServer::writeColumData(int column, const std::vector<double>& res)
{
	int begin = beginRow.toInt() - 1;
	int end = endRow.toInt();
	for (int r = begin; r < end; ++r)
	{
		if ((res[r - begin] != std::numeric_limits<double>::max()))
			sheetContent[r][column] = QVariant(res[r-begin]);
		else
		{
			sheetContent[r][column] = QVariant("/");
		}
	}
}

void ExcelDataServer::castListListVariant2Variant(QVariant& var, const QList<QList<QVariant>>& res)
{
	QVariant temp = QVariant(QVariantList());
	QVariantList record;

	int listSize = res.size();
	for (int i = 0; i < listSize; ++i)
	{
		temp = res.at(i);
		record << temp;
	}
	temp = record;
	var = temp;
}

void ExcelDataServer::castSheetVector2Variant(QVariant& var)
{
	QVariant temp = QVariant(QVariantList());
	QVariantList recordRows;
	for (int r = 0; r < sheetContent.size(); ++r)
	{
		QVariantList currentRow;
		for (int c = 0; c < sheetContent[r].size(); ++c)
		{
			currentRow.push_back(sheetContent[r][c]);
		}
		temp = currentRow;
		recordRows << temp;
	}
	temp = recordRows;
	var = temp;
}

void ExcelDataServer::castDoubleVector2Variant(QVariant& var, const std::vector<double>& res)
{
	QVariant temp = QVariant(QVariantList());
	QVariantList record;

	int listSize = res.size();

	for (int i = 0; i < listSize; ++i)
	{
		if((res.at(i) != std::numeric_limits<double>::max()))
			temp = QList<QVariant>{ res.at(i)};
		else
		{
			temp = QList<QVariant>{ "/" };
		}
		record << temp;
	}
	temp = record;
	var = temp;
}

void ExcelDataServer::writeColum(QAxObject* sheet, QString range, const QVariant& data)
{
	QAxObject* allEnvData = sheet->querySubObject("Range(QString)", range);
	allEnvData->setProperty("Value", data);
}

int ExcelDataServer::getRowsNumber()
{
	QVariantList eachRow = allData.toList();
	return eachRow.size(); //��ȡ����
	
}

bool ExcelDataServer::getRowData(QAxObject* sheet, int rowNumber, QVariantList& result)
{
	QVariantList allEnvDataList = allData.toList();//ת��Ϊlist
	result = allEnvDataList.at(rowNumber - 1).toList();
	
	for (int i = 0; i < result.size(); ++i)
	{
		QString pre("");
		if(i > 25)
		    pre = QString(char(64 + i / 26));

		QString c = QString(pre + char(65 + i % 26));
		QString name = result.at(i).toString();
		name.remove(QRegExp("\\s"));
		nameToColum[name] = c;
		nameToSubScript[name] = i;
	}

	return true; 
}

int ExcelDataServer::getColumsNumber()
{
	QVariantList eachRow = allData.toList();

	return eachRow[0].toList().size(); //��ȡ����
}

void ExcelDataServer::setAllData(QAxObject* usedrange)
{
	allData = usedrange->dynamicCall("Value");
	QVariantList  rows = allData.toList();
	for (int r = 0; r < rows.size(); ++r)
	{
		std::vector<QVariant> currentRow;
		QVariantList sourceRow = rows.at(r).toList();
		for (int c = 0; c < sourceRow.size(); ++c)
		{
			currentRow.push_back(sourceRow.at(c));
		}
		sheetContent.push_back(currentRow);
	}
}

void ExcelDataServer::operation(const std::vector<double>& colum1, const std::vector<double>& colum2,
	std::vector<double>& result, const char oper)
{
	switch (oper)
	{
	case '+':
		for (int i = 0; i < colum1.size(); ++i)
		{
			if ((colum1[i] == std::numeric_limits<double>::max()) || (colum2[i] == std::numeric_limits<double>::max()))
				result.push_back(std::numeric_limits<double>::max());
			else
				result.push_back(colum1[i] + colum2[i]);
		}
		break;
	case'-':
		for (int i = 0; i < colum1.size(); ++i)
		{
			if ((colum1[i] == std::numeric_limits<double>::max()) || (colum2[i] == std::numeric_limits<double>::max()))
				result.push_back(std::numeric_limits<double>::max());
			else
				result.push_back(colum1[i] - colum2[i]);
		}
		break;
	case '*':
		for (int i = 0; i < colum1.size(); ++i)
		{
			if ((colum1[i] == std::numeric_limits<double>::max()) || (colum2[i] == std::numeric_limits<double>::max()))
				result.push_back(std::numeric_limits<double>::max());
			else
				result.push_back(colum1[i] * colum2[i]);
		}
		break;
	case '/':
		for (int i = 0; i < colum1.size(); ++i)
		{
			if (colum2[i] == 0)
			{
				result.push_back(std::numeric_limits<double>::max());
				continue;
			}
			if ((colum1[i] == std::numeric_limits<double>::max()) || (colum2[i] == std::numeric_limits<double>::max()))
				result.push_back(std::numeric_limits<double>::max());
			else
				result.push_back(colum1[i] / colum2[i]);
		}
		break;
	default:
		break;
	}
}

void ExcelDataServer::calculator(const QString& exp, const QString& output)
{
	//reference: https://blog.csdn.net/qq_36236235/article/details/80086779
	std::stack<std::vector<double>> numberStack;
	std::stack<QChar> operatorStack;

	QVector<QChar> oprates{ '-','+','*','/','(',')' };

	QChar current;
	for (int i = 0; i < exp.size(); ++i)
	{
		current = exp.at(i);
		if (!oprates.contains(current))
		{
			int count = 0;
			QString temp;
			while (i < exp.size() && !oprates.contains(exp.at(i)))
			{
				temp += exp.at(i);
				++i;
			}
			--i;
			std::vector<double> tempVec;
			if (temp.size() == 1)
			{
				tempVec = std::vector<double>(endRow.toInt() - beginRow.toInt() + 1, temp.toDouble());
			}
			else
			{
				/*auto itr = nameToColum.find(temp);
				if (itr == nameToColum.end())
				{
					printf("\tCan not find %s \n", qPrintable(temp));
					return;
				}
				getColum(currentWorksheet, itr->second + beginRow + ":" + itr->second + endRow, tempVec);*/

				auto iter = nameToSubScript.find(temp);
				if (iter == nameToSubScript.end())
				{
					printf("\tCan not find %s \n", qPrintable(temp));
					return;
				}
				getColum(currentWorksheet, iter->second, tempVec);
			}
			numberStack.push(tempVec);
			continue;
		}
		else//operator
		{
			//is a empty operator stack
			if (operatorStack.empty())
			{
				operatorStack.push(current);
			}
			else //not empty
			{
				QChar pre = operatorStack.top();
				if (checkProority(current, pre) == 1)//current > pre 
				{
					if (current == ')')
					{
						std::vector<double> d1;
						std::vector<double> d2;
						while (pre != '(')
						{
							operatorStack.pop();
							d1 = numberStack.top();
							numberStack.pop();
							d2 = numberStack.top();
							numberStack.pop();
							std::vector<double> result;
							arithmeticOperation(d2, d1, result, pre.toLatin1());
							numberStack.push(result);
							pre = operatorStack.top();
						}
						operatorStack.pop();
					}
					else
					{
						operatorStack.push(current); //tested<<<===
					}
				}
				//current opertor prority =  pre   tested<<<<<<============
				else if (checkProority(current, pre) == 0)
				{
					operatorStack.pop();
					std::vector<double> d1 = numberStack.top();
					numberStack.pop();
					std::vector<double> d2 = numberStack.top();
					numberStack.pop();
					std::vector<double> result;
					arithmeticOperation(d2, d1, result, pre.toLatin1());
					numberStack.push(result);
					operatorStack.push(current);
				}
				else//current < prority
				{
					if (pre == '(')
						operatorStack.push(current);
					else
					{
						while (checkProority(current, pre) <= 0 )
						{
							pre = operatorStack.top();
							if (pre == "(")
							{
								break;
							}
							operatorStack.pop();
							std::vector<double> d1 = numberStack.top();
							numberStack.pop();
							std::vector<double> d2 = numberStack.top();
							numberStack.pop();
							std::vector<double> result;
							arithmeticOperation(d2, d1, result, pre.toLatin1());
							numberStack.push(result);
							if (operatorStack.empty())
								break;
							pre = operatorStack.top();
						}
						operatorStack.push(current);
					}
				}
			}
		}
	}
	while (!operatorStack.empty())
	{
		QChar oper = operatorStack.top();
		operatorStack.pop();
		std::vector<double> d1 = numberStack.top();
		numberStack.pop();
		std::vector<double> d2 = numberStack.top();
		numberStack.pop();
		std::vector<double> result;
		arithmeticOperation(d2, d1, result, oper.toLatin1());
		numberStack.push(result);
	}
	/*auto tempIter = nameToColum.find(output);
	if (tempIter == nameToColum.end())
	{
		printf("Can not find output column: %s\n", qPrintable(output));
		return;
	}
	writeColumData(currentWorksheet, nameToColum[output], numberStack.top());*/

	auto iter = nameToSubScript.find(output);
	if (iter == nameToSubScript.end())
	{
		printf("Can not find output column: %s\n", qPrintable(output));
		return;
	}
	writeColumData(nameToSubScript[output], numberStack.top());
}

int ExcelDataServer::checkProority(QChar s1, QChar s2)
{
	if (s1 == '+' || s1 == '-') {
		if (s2 == '+' || s2 == '-') {
			return 0;
		}
		else {
			return -1;
		}
	}
	else if (s1 == '*' || s1 == '/') {
		if (s2 == '+' || s2 == '-') {
			return 1;
		}
		else if (s2 == '*' || s2 == '/') {
			return 0;
		}
		else {
			return -1;
		}
	}
	else {
		return 1;
	}
}

QAxObject* ExcelDataServer::addSheet(QAxObject* workSheets, const QString& name)
{
	int sheet_count = workSheets->property("Count").toInt();  //��ȡ��������Ŀ
	QAxObject* last_sheet = workSheets->querySubObject("Item(int)", sheet_count);
	QAxObject* work_sheet = workSheets->querySubObject("Add(QVariant)", last_sheet->asVariant());
	last_sheet->dynamicCall("Move(QVariant)", work_sheet->asVariant());

	work_sheet->setProperty("Name", name);  //���ù���������
	return work_sheet;
}

//65 -- 90
int ExcelDataServer::alphabet2Int(const QString& alp)
{
	auto itr = alp.rbegin();
	int result = 0;
	int i = 0;
	while (itr != alp.rend())
	{
		result += pow(26, i)*((*itr).toLatin1() - 64);
		++itr;
		++i;
	}
	return result - 1;
}

//number > 0
void ExcelDataServer::int2Alphabet(int number, QString& alphabet)
{
	Q_ASSERT(number > 0);
	int tempData = number / 26;
	if (tempData > 0)
	{
		int mode = number % 26;
		int2Alphabet(mode, alphabet);
		int2Alphabet(tempData, alphabet);
	}
	else
	{
		alphabet += char(64 + number);
	}
}

void ExcelDataServer::appendColums2Sheet(const std::vector<QString>& headNames,
	QAxObject* sheet)
{
	QAxObject* usedrange = sheet->querySubObject("UsedRange");
	QAxObject* rows = usedrange->querySubObject("Rows");
	int rownum = rows->property("Count").toInt(); //��ȡ����

	QAxObject* colums = usedrange->querySubObject("Columns");
	int colnum = rows->property("Count").toInt(); //��ȡ����

	for (int i = 0; i < headNames.size(); ++i)
	{
		int currentColnum = colnum + i + 1;
		QString pre("");
		if (currentColnum > 25)
			pre = QString(char(64 + currentColnum / 26));

		QString toColnumStr = QString(pre + char(65 + currentColnum % 26));
		QVariant data;
		getColum(currentWorksheet, nameToColum[headNames[i]] + beginRow + ":" + nameToColum[headNames[i]] + endRow, data);
		QVariantList list = data.toList();
		list.push_front(QVariant(QVariantList{ headNames[i] }));
		writeColum(sheet, toColnumStr + "1:" + toColnumStr + QString::number(list.size()), list);
	}
}

void ExcelDataServer::selectWhere( const std::vector<QString>& selectedName,
	const QString& whereName, const QString& whereValue, QList<QList<QVariant>>& result)
{
	QVariantList rows = allData.toList();
	for (int i = 0; i < rows.size(); ++i)
	{
		QVariantList currentColumn = rows.at(i).toList();
		QString tempValue = currentColumn.at(nameToSubScript[whereName]).toString();
		if (tempValue == whereValue)
		{
			QList<QVariant> temp;
			for (const QString& field : selectedName)
			{
				temp.push_back(currentColumn.at(nameToSubScript[field]));
			}
			result.push_back(temp);
		}
	}
}
QAxObject* ExcelDataServer::getNamedSheet(QAxObject* sheets, const QString& name)
{
	//QAxObject* worksheets = workbook->querySubObject("WorkSheets");
	int count = sheets->property("Count").toInt();
	for (int i = 0; i < count; ++i)
	{
		QAxObject* currentSheet = sheets->querySubObject("Item(int)", i+1);
		QString currentName = currentSheet->property("Name").toString();
		if (currentName == name)
			return currentSheet;
	}
	printf("Can not find sheet : %s\s", name);
	return nullptr;
}

QVariant ExcelDataServer::getCellData(const QString& name, int row)
{
	//auto iter1 = nameToColum.find(name);
	//if (iter1 == nameToColum.end())
	//	return QVariant();
	//QAxObject* cellA22 = currentWorksheet->querySubObject("Range(QVariant, QVariant)", iter1->second + QString::number(row));
	//return cellA22->dynamicCall("Value");

	auto iter = nameToSubScript.find(name);
	if (iter == nameToSubScript.end())
		return QVariant();
	return sheetContent[row - 1][iter->second];
	/*QVariantList rowsData = allData.toList();
	QVariantList tragetRow = rowsData.at(row-1).toList();

	return tragetRow.at(iter->second);*/
}

void ExcelDataServer::writedata(QString data, QString c, int r)
{
	auto iter = nameToColum.find(c);
	if (iter == nameToColum.end())
	{
		printf("can not find column : %s\n", qPrintable(c));
		return;
	}

	QAxObject* cell = currentWorksheet->querySubObject("Range(QVariant, QVariant)",
		iter->second + QString::number(r));
	cell->dynamicCall("SetValue(const QVariant&)", QVariant(data));//���õ�Ԫ���ֵ
}
void ExcelDataServer::writedata(int data, QString c, int r)
{
	auto iter = nameToColum.find(c);
	if (iter == nameToColum.end())
	{
		printf("can not find column : %s\n", qPrintable(c));
		return;
	}

	QAxObject * cell = currentWorksheet->querySubObject("Range(QVariant, QVariant)",
		iter->second + QString::number(r));
	cell->dynamicCall("SetValue(const QVariant&)", QVariant(data));//���õ�Ԫ���ֵ
}

void ExcelDataServer::writedata(QVariant data, QString c, int r)
{
	auto iter = nameToSubScript.find(c);
	if (iter == nameToSubScript.end())
	{
		printf("can not find column : %s\n", qPrintable(c));
		return;
	}
	sheetContent[r - 1][iter->second] = data;
	/*QVariantList rowsData = sheetContent.toList();
	QVariantList tragetRow = rowsData.at(r - 1).toList();
	tragetRow[iter->second] = data;*/

	//auto iter = nameToColum.find(c);
	//if (iter == nameToColum.end())
	//{
	//	printf("can not find column : %s\n", qPrintable(c));
	//	return;
	//}

	//QAxObject* cell = currentWorksheet->querySubObject("Range(QVariant, QVariant)",
	//	iter->second + QString::number(r));
	//cell->dynamicCall("SetValue(const QVariant&)", data);//���õ�Ԫ���ֵ
}