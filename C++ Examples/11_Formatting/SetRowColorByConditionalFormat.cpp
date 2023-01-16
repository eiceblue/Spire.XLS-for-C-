#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Template_Xls_4.xlsx";
	wstring outputFile = output_path + L"SetRowColorByConditionalFormat.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Select the range that you want to format.
	CellRange* dataRange = sheet->GetAllocatedRange();

	//Set conditional formatting.
	XlsConditionalFormats* xcfs = sheet->GetConditionalFormats()->Add();
	xcfs->AddRange(dataRange);
	IConditionalFormat* format1 = xcfs->AddCondition();
	//Determines the cells to format.
	format1->SetFirstFormula(L"=MOD(ROW(),2)=0");
	//Set conditional formatting type
	format1->SetFormatType(ConditionalFormatType::Formula);
	//Set the color.
	format1->SetBackColor(Spire::Common::Color::GetLightSeaGreen());

	//Set the backcolor of the odd rows as Yellow.
	XlsConditionalFormats* xcfs1 = sheet->GetConditionalFormats()->Add();
	xcfs1->AddRange(dataRange);
	IConditionalFormat* format2 = xcfs->AddCondition();
	format2->SetFirstFormula(L"=MOD(ROW(),2)=1");
	format2->SetFormatType(ConditionalFormatType::Formula);
	format2->SetBackColor(Spire::Common::Color::GetYellow());

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}