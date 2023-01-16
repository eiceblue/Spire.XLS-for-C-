#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetConditionalFormatFormula.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Add ConditionalFormat
	XlsConditionalFormats* xcfs = sheet->GetConditionalFormats()->Add();

	//Define the range
	xcfs->AddRange(sheet->GetRange(L"B5"));

	//Add condition
	IConditionalFormat* format = xcfs->AddCondition();
	format->SetFormatType(ConditionalFormatType::CellValue);

	//If greater than 1000
	format->SetFirstFormula(L"1000");
	format->SetOperator(ComparisonOperatorType::Greater);
	format->SetBackColor(Spire::Common::Color::GetOrange());

	sheet->GetRange(L"B1")->SetNumberValue(40);
	sheet->GetRange(L"B2")->SetNumberValue(500);
	sheet->GetRange(L"B3")->SetNumberValue(300);
	sheet->GetRange(L"B4")->SetNumberValue(400);

	//Set a SUM formula for B5
	sheet->GetRange(L"B5")->SetFormula(L"=SUM(B1:B4)");

	//Add text
	sheet->GetRange(L"C5")->SetText(L"If Sum of B1:B4 is greater than 1000, B5 will have orange background.");

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}