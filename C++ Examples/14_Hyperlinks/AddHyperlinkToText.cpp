#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"CommonTemplate1.xlsx";
	wstring outputFile = output_path + L"AddHyperlinkToText.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Add url link
	HyperLink* UrlLink = dynamic_cast<HyperLink*>(sheet->GetHyperLinks()->Add(sheet->GetRange(L"D10")));
	UrlLink->SetTextToDisplay(sheet->GetRange(L"D10")->GetText());
	UrlLink->SetType(HyperLinkType::Url);
	UrlLink->SetAddress(L"http://en.wikipedia.org/wiki/Chicago");

	//Add email link
	XlsHyperLink* MailLink = sheet->GetHyperLinks()->Add(sheet->GetRange(L"E10"));
	MailLink->SetTextToDisplay(sheet->GetRange(L"E10")->GetText());
	MailLink->SetType(HyperLinkType::Url);
	MailLink->SetAddress(L"mailto:Amor.Aqua@gmail.com");

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}