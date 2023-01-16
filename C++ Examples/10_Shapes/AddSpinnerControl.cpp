#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"ExcelSample_N1.xlsx";
	wstring outputFile = output_path + L"AddSpinnerControl.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Set text for range C11
	sheet->GetRange(L"C11")->SetText(L"Value:");
	sheet->GetRange(L"C11")->GetStyle()->GetFont()->SetIsBold(true);

	//Set value for range B10
	sheet->GetRange(L"C12")->SetValue(0);

	//Add spinner control
	ISpinnerShape* spinner = sheet->GetSpinnerShapes()->AddSpinner(12, 4, 20, 20);
	spinner->SetLinkedCell(sheet->GetRange(L"C12"));
	spinner->SetMin(0);
	spinner->SetMax(100);
	spinner->SetIncrementalChange(5);
	spinner->SetDisplay3DShading(true);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
