#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile = input_path + L"WorksheetSample1.xlsx";
	std::wstring outputFile = output_path + L"SetPageBreak.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Set Excel Page Break Horizontally
	(dynamic_cast<HPageBreaksCollection*>(sheet->GetHPageBreaks()))->Add(sheet->GetRange(L"A8"));
	(dynamic_cast<HPageBreaksCollection*>(sheet->GetHPageBreaks()))->Add(sheet->GetRange(L"A14"));

	//Set Excel Page Break Vertically
	//sheet.VPageBreaks.Add(sheet.Range["B1"]);
	//sheet.VPageBreaks.Add(sheet.Range["C1"]);

	//Set view mode to Preview mode
	workbook->GetWorksheets()->Get(0)->SetViewMode(ViewMode::Preview);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}