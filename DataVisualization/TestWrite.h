#pragma once
#include <QObject>
#include <ActiveQt/qaxobject.h>
class SingleTestResult
{
public:
	QString testName;
	QList<QList<QVariant>> Data;
};

class TestRecordExporter : public QObject
{
	Q_OBJECT

public:
	TestRecordExporter(QObject* parent = 0);
	~TestRecordExporter();
	bool writeToFile(QString savePath, QList<SingleTestResult>& data);
	bool writeToSheet(QList<QList<QVariant>>& data, QAxObject* worksheet);

private:
	
	void convertToColName(int data, QString& res);
	QString to26AlphabetString(int data);
	void castListListVariant2Variant(QVariant& var, const QList<QList<QVariant> >& res);



private:

};
