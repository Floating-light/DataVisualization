#pragma once
#include <QComboBox>

class MyCombbox:  public QComboBox
{
public:
	MyCombbox(int);
	~MyCombbox();
	int getIndex();
private:
	int index;
};

