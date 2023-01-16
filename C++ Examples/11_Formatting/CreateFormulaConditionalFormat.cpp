#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Template_Xls_6.xlsx";
	wstring outputFile = output_path + L"CreateFormulaConditionalFormat.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);
	XlsRange* range = sheet->GetColumns()->GetItem(0);

	//Set the conditional formatting formula and apply the rule to the chosen cell range.
	XlsConditionalFormats* xcfs = sheet->GetConditionalFormats()->Add();
	xcfs->AddRange(range);
	IConditionalFormat* conditional = xcfs->AddCondition();
	conditional->SetFormatType(ConditionalFormatType::Formula);
	conditional->SetFirstFormula(L"=($A1<$B1)");
	conditional->SetBackKnownColor(ExcelColors::Yellow);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}