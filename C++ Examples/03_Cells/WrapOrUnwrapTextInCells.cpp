#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring outputFolder = OUTPUTPATH;
	wstring outputFile = outputFolder + L"WrapOrUnwrapTextInCells_out.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Wrap the excel text
	sheet->GetRange(L"C1")->SetText(L"e-iceblue is in facebook and welcome to like us");
	sheet->GetRange(L"C1")->GetStyle()->SetWrapText(true);
	sheet->GetRange(L"D1")->SetText(L"e-iceblue is in twitter and welcome to follow us");
	sheet->GetRange(L"D1")->GetStyle()->SetWrapText(true);

	//Unwrap the excel text
	sheet->GetRange(L"C2")->SetText(L"http://www.facebook.com/pages/e-iceblue/139657096082266");
	sheet->GetRange(L"C2")->GetStyle()->SetWrapText(false);
	sheet->GetRange(L"D2")->SetText(L"https://twitter.com/eiceblue");
	sheet->GetRange(L"D2")->GetStyle()->SetWrapText(false);

	//Set the text color of Range["C1:D1"]
	sheet->GetRange(L"C1:D1")->GetStyle()->GetFont()->SetSize(15);
	sheet->GetRange(L"C1:D1")->GetStyle()->GetFont()->SetColor(Spire::Common::Color::GetBlue());
	//Set the text color of Range["C2:D2"]
	sheet->GetRange(L"C2:D2")->GetStyle()->GetFont()->SetSize(15);
	sheet->GetRange(L"C2:D2")->GetStyle()->GetFont()->SetColor(Spire::Common::Color::GetDeepSkyBlue());

	//Save to file
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}