#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile = input_path + L"PageBreak.xlsx";
	std::wstring outputFile = output_path + L"RemovePageBreak.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Clear all the vertical page breaks
	dynamic_cast<VPageBreaksCollection*>(sheet->GetVPageBreaks())->Clear();

	//Remove the firt horizontal Page Break
	dynamic_cast<HPageBreaksCollection*>(sheet->GetHPageBreaks())->RemoveAt(0);

	//Set the ViewMode as Preview to see how the page breaks work
	sheet->SetViewMode(ViewMode::Preview);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}