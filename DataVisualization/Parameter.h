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
	//�׿����̿�
	double firstOpenCommerceOpenResidence;
	double firstOpenCommerceOpenShop;
	double firstOpenComemrceOpenOffice;
	double firstOpenComemrceOpenElse;
	//�׿���ס��
	double firstOpenResidenceOpenUp3;
	double firstOpenResidenceOpenDown4;
	
	//������ס��
	double continueResidence;
	//�������̿�
	double continueCommerceResidence;
	double continueCommerceShop;
	double continueCommerceOffice;
	double continueCommerceElse;

	//������С��3ס��
	double plusLess3Residence;
	//������С��3�̿�
	double plusLess3CommerceResidence;
	double plusLess3CommerceShop;
	double plusLess3CommerceOffice;
	double plusLess3CommerceElse;
	//���ƴ���3���л���
	double plusMore3ResidenceCityGre3;
	double plusMore3ResidenceCityLess3;
	double plusMore3CommerceResidence;
	double plusMore3CommerceStore;
	double plusMore3CommerceOffice;
	double plusMore3CommerceElse;
};

