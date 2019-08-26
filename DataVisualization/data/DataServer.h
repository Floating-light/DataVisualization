#pragma once
#include <ActiveQt/qaxobject.h>

class DataServer
{
public:
	virtual ~DataServer()
	{

	}

	//header: "A9:A100"
	virtual void getColumBeginWith(QAxObject*, const QString& header) {};
	virtual void writeColumBeginWith(QAxObject*, const QString& header) {};

};