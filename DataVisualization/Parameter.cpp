#include "Parameter.h"


Parameter::Parameter()
{
	loadMessage();
}

void Parameter::loadMessage() {
	QSettings set("./input/preload.ini", QSettings::IniFormat);

	set.beginGroup("AppConfig");

	firstOpenCommerceOpenResidence = set.value("firstOpenCommerceOpenResidence").toDouble();
	firstOpenCommerceOpenShop = set.value("firstOpenCommerceOpenShop").toDouble();
	firstOpenComemrceOpenOffice = set.value("firstOpenComemrceOpenOffice").toDouble();
	firstOpenComemrceOpenElse = set.value("firstOpenComemrceOpenElse").toDouble();
	firstOpenResidenceOpenUp3 = set.value("firstOpenResidenceOpenUp3").toDouble();
	firstOpenResidenceOpenDown4 = set.value("firstOpenResidenceOpenDown4").toDouble();

	continueResidence = set.value("continueResidence").toDouble();
	continueCommerceResidence = set.value("continueCommerceResidence").toDouble();
	continueCommerceShop = set.value("continueCommerceShop").toDouble();
	continueCommerceOffice = set.value("continueCommerceOffice").toDouble();
	continueCommerceElse = set.value("continueCommerceElse").toDouble();

	plusLess3Residence = set.value("plusLess3Residence").toDouble();
	plusLess3CommerceResidence = set.value("plusLess3CommerceResidence").toDouble();
	plusLess3CommerceShop = set.value("plusLess3CommerceShop").toDouble();
	plusLess3CommerceOffice = set.value("plusLess3CommerceOffice").toDouble();
	plusLess3CommerceElse = set.value("plusLess3CommerceElse").toDouble();

	plusMore3ResidenceCityGre3 = set.value("plusMore3ResidenceCityGre3").toDouble();
	plusMore3ResidenceCityLess3 = set.value("plusMore3ResidenceCityLess3").toDouble();
	plusMore3CommerceResidence = set.value("plusMore3CommerceResidence").toDouble();
	plusMore3CommerceStore = set.value("plusMore3CommerceStore").toDouble();
	plusMore3CommerceOffice = set.value("plusMore3CommerceOffice").toDouble();
	plusMore3CommerceElse = set.value("plusMore3CommerceElse").toDouble();

	set.endGroup();
}

double Parameter::getFirstOpenCommerceOpenResidence(){
	return firstOpenCommerceOpenResidence;
}

double Parameter::getFirstOpenCommerceOpenShop(){
	return firstOpenCommerceOpenShop;
}
double Parameter::getFirstOpenComemrceOpenOffice(){
	return firstOpenComemrceOpenOffice;
}
double Parameter::getFirstOpenComemrceOpenElse() {
	return firstOpenComemrceOpenElse;
}
double Parameter::getFirstOpenResidenceOpenUp3(){
	return firstOpenResidenceOpenUp3;
}
double Parameter::getFirstOpenResidenceOpenDown4(){
	return firstOpenResidenceOpenDown4;
}
double Parameter::getContinueResidence() {
	return continueResidence;
}
double Parameter::getContinueCommerceResidence() {
	return continueCommerceResidence;
}
double Parameter::getContinueCommerceShop() {
	return continueCommerceShop;
}
double Parameter::getContinueCommerceOffice() {
	return continueCommerceOffice;
}
double Parameter::getContinueCommerceElse() {
	return continueCommerceElse;
}
double Parameter::getPlusLess3Residence() {
	return plusLess3Residence;
}
double Parameter::getPlusLess3CommerceResidence() {
	return plusLess3CommerceResidence;
}
double Parameter::getPlusLess3CommerceShop() {
	return plusLess3CommerceShop;
}
double Parameter::getPlusLess3CommerceOffice() {
	return plusLess3CommerceOffice;
}
double Parameter::getPlusLess3CommerceElse() {
	return plusLess3CommerceElse;
}
double Parameter::getPlusMore3ResidenceCityGre3() {
	return plusMore3ResidenceCityGre3;
}
double Parameter::getPlusMore3ResidenceCityLess3() {
	return plusMore3ResidenceCityLess3;
}
double Parameter::getPlusMore3CommerceResidence() {
	return plusMore3CommerceResidence;
}
double Parameter::getPlusMore3CommerceStore() {
	return plusMore3CommerceStore;
}
double Parameter::getPlusMore3CommerceOffice() {
	return plusMore3CommerceOffice;
}
double Parameter::getPlusMore3CommerceElse() {
	return plusMore3CommerceElse;
}

Parameter::~Parameter()
{
}
