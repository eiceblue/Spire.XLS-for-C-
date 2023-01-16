#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring output_path = OUTPUTPATH;
	std::wstring outputFile = output_path + L"FillDataInWorksheet.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Fill data
	sheet->GetRange(L"A1")->GetStyle()->GetFont()->SetIsBold(true);
	sheet->GetRange(L"B1")->GetStyle()->GetFont()->SetIsBold(true);
	sheet->GetRange(L"C1")->GetStyle()->GetFont()->SetIsBold(true);
	sheet->GetRange(L"A1")->SetText(L"Month");
	sheet->GetRange(L"A2")->SetText(L"January");
	sheet->GetRange(L"A3")->SetText(L"February");
	sheet->GetRange(L"A4")->SetText(L"March");
	sheet->GetRange(L"A5")->SetText(L"April");
	sheet->GetRange(L"B1")->SetText(L"Payments");
	sheet->GetRange(L"B2")->SetNumberValue(251);
	sheet->GetRange(L"B3")->SetNumberValue(515);
	sheet->GetRange(L"B4")->SetNumberValue(454);
	sheet->GetRange(L"B5")->SetNumberValue(874);
	sheet->GetRange(L"C1")->SetText(L"Sample");
	sheet->GetRange(L"C2")->SetText(L"Sample1");
	sheet->GetRange(L"C3")->SetText(L"Sample2");
	sheet->GetRange(L"C4")->SetText(L"Sample3");
	sheet->GetRange(L"C5")->SetText(L"Sample4");

	//Set width for the second column
	sheet->SetColumnWidth(2, 10);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}