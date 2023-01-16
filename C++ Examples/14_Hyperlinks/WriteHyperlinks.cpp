#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"WriteHyperlinks.xlsx";
	wstring outputFile = output_path + L"WriteHyperlinks.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	sheet->GetRange(L"B9")->SetText(L"Home page");
	HyperLink* hylink1 = dynamic_cast<HyperLink*>(sheet->GetHyperLinks()->Add(sheet->GetRange(L"B10")));
	hylink1->SetType(HyperLinkType::Url);
	hylink1->SetAddress(L"(http://www.e-iceblue.com)");

	sheet->GetRange(L"B11")->SetText(L"Support");
	HyperLink* hylink2 = dynamic_cast<HyperLink*>(sheet->GetHyperLinks()->Add(sheet->GetRange(L"B12")));
	hylink2->SetType(HyperLinkType::Url);
	hylink2->SetAddress(L"mailto:support@e-iceblue.com");

	sheet->GetRange(L"B13")->SetText(L"Forum");
	HyperLink* hylink3 = dynamic_cast<HyperLink*>(sheet->GetHyperLinks()->Add(sheet->GetRange(L"B14")));
	hylink3->SetType(HyperLinkType::Url);
	hylink3->SetAddress(L"https://www.e-iceblue.com/forum/");

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}