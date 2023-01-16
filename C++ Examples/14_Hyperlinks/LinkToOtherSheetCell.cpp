#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"LinkToExternalFile.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	CellRange* range = sheet->GetRange(L"A1");

	//Add hyperlink in the range
	HyperLink* hyperlink = dynamic_cast<HyperLink*>(sheet->GetHyperLinks()->Add(range));

	//Set the link type
	hyperlink->SetType(HyperLinkType::Workbook);

	//Set the display text
	hyperlink->SetTextToDisplay(L"Link to Sheet2 cell C5");

	//Set the address
	hyperlink->SetAddress(L"Sheet2!C5");

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}