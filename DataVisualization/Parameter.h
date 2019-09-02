#pragma once
#include <QSettings>

class Parameter
{
public:
	Parameter();
	~Parameter();

	void loadMessage();

	double getFirstOpenCommerceOpenResidence();
	double getFirstOpenCommerceOpenShop();
	double getFirstOpenComemrceOpenOffice();
	double getFirstOpenComemrceOpenElse();
	double getFirstOpenResidenceOpenUp3();
	double getFirstOpenResidenceOpenDown4();

	double getContinueResidence();
	double getContinueCommerceResidence();
	double getContinueCommerceShop();
	double getContinueCommerceOffice();
	double getContinueCommerceElse();

	double getPlusLess3Residence();
	double getPlusLess3CommerceResidence();
	double getPlusLess3CommerceShop();
	double getPlusLess3CommerceOffice();
	double getPlusLess3CommerceElse();

	double getPlusMore3ResidenceCityGre3();
	double getPlusMore3ResidenceCityLess3();
	double getPlusMore3CommerceResidence();
	double getPlusMore3CommerceStore();
	double getPlusMore3CommerceOffice();
	double getPlusMore3CommerceElse();

private:
	//首开类商开
	double firstOpenCommerceOpenResidence;
	double firstOpenCommerceOpenShop;
	double firstOpenComemrceOpenOffice;
	double firstOpenComemrceOpenElse;
	//首开类住开
	double firstOpenResidenceOpenUp3;
	double firstOpenResidenceOpenDown4;
	
	//续销类住开
	double continueResidence;
	//续销类商开
	double continueCommerceResidence;
	double continueCommerceShop;
	double continueCommerceOffice;
	double continueCommerceElse;

	//加推类小于3住开
	double plusLess3Residence;
	//加推类小于3商开
	double plusLess3CommerceResidence;
	double plusLess3CommerceShop;
	double plusLess3CommerceOffice;
	double plusLess3CommerceElse;
	//加推大于3城市环线
	double plusMore3ResidenceCityGre3;
	double plusMore3ResidenceCityLess3;
	double plusMore3CommerceResidence;
	double plusMore3CommerceStore;
	double plusMore3CommerceOffice;
	double plusMore3CommerceElse;
};

