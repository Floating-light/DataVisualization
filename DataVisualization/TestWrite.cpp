#include <QAxObject>
#include <QDebug>
#include <QElapsedTimer>
#include <QMessageBox>  

#include "TestWrite.h"

TestRecordExporter::TestRecordExporter(QObject* parent)
	: QObject(parent)
{
}

TestRecordExporter::~TestRecordExporter()
{

}

bool TestRecordExporter::writeToFile(QString savePath, QList<SingleTestResult>& data)
{
	bool write_success = false;

	QAxObject excel("Excel.Application");
	excel.dynamicCall("SetVisible (bool Visible)", "false");
	excel.setProperty("DisplayAlerts", false);

	QAxObject* workbooks = excel.querySubObject("WorkBooks");
	if (!workbooks)
	{
		QMessageBox::information(NULL, "Export Error ", "No Microsoft Excel ", QMessageBox::Yes);
		return false;
	}


	workbooks->dynamicCall("Add");

	QAxObject* workbook = excel.querySubObject("ActiveWorkBook");
	if (!workbook) return false;

	QAxObject* worksheets = workbook->querySubObject("Sheets");
	if (!worksheets) return false;

	int worksheetsNum = data.size();
	int currentSheetsNum = worksheets->property("Count").toInt();
	if (worksheetsNum > currentSheetsNum)
	{
		for (int i = currentSheetsNum; i < worksheetsNum; ++i)
		{
			worksheets->dynamicCall("Add");
		}
	}

	//exportTofile
	int sheetsNum = worksheets->property("Count").toInt();
	for (int i = 1; i <= sheetsNum; ++i)
	{
		QAxObject* worksheet = worksheets->querySubObject("Item(int)", i);
		SingleTestResult temp = data.takeFirst();
		worksheet->setProperty("Name", temp.testName);
		write_success = writeToSheet(temp.Data, worksheet);
	}
	//save
	savePath.replace('/', '\\');
	workbook->dynamicCall("SaveAs(const QString&)", savePath);
	workbook->dynamicCall("Close(Boolean)", false);
	excel.dynamicCall("Quit(void)");

	return write_success;
}

bool TestRecordExporter::writeToSheet(QList<QList<QVariant>>& data, QAxObject* worksheet)
{
	int row = data.size();
	int col = data.at(0).size();
	QString rangStr;
	convertToColName(col, rangStr);
	rangStr += QString::number(row);
	rangStr = "A1:" + rangStr;
	if (!worksheet)
		return false;
	QAxObject* range = worksheet->querySubObject("Range(const QString&)", rangStr);
	if (NULL == range || range->isNull())
		return false;
	QVariant var;
	castListListVariant2Variant(var, data);
	range->setProperty("Value", var);
	delete range;
	return true;
}


//1->A 26->Z 27->AA
void TestRecordExporter::convertToColName(int data, QString & res)
{
	Q_ASSERT(data > 0 && data < 65535);
	int tempData = data / 26;
	if (tempData > 0)
	{
		int mode = data % 26;
		convertToColName(mode, res);
		convertToColName(tempData, res);
	}
	else
	{
		res = (to26AlphabetString(data) + res);
	}
}

QString TestRecordExporter::to26AlphabetString(int data)
{
	QChar ch = data + 0x40;//A∂‘”¶0x41
	return QString(ch);
}

void TestRecordExporter::castListListVariant2Variant(QVariant & var, const QList<QList<QVariant>> & res)
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