#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"LinkToExternalFile.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	CellRange* range = sheet->GetRange(1, 1);

	//Add hyperlink in the range
	HyperLink* hyperlink = dynamic_cast<HyperLink*>(sheet->GetHyperLinks()->Add(range));

	//Set the link type
	hyperlink->SetType(HyperLinkType::File);

	//Set the display text
	hyperlink->SetTextToDisplay(L"Link To External File");

	//Set file address
	hyperlink->SetAddress((input_path + L"SampeB_4.xlsx").c_str());

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}