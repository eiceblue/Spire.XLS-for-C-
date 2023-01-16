#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"SetBorder.xlsx";
	wstring outputFile = output_path + L"SetBorder.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Get the cell range where you want to apply border style
	CellRange* cr = sheet->GetRange(sheet->GetFirstRow(), sheet->GetFirstColumn(), sheet->GetLastRow(), sheet->GetLastColumn());

	//Apply border style 
	cr->GetBorders()->SetLineStyle(LineStyleType::Double);
	IBorder* border = cr->GetBorders()->Get(BordersLineType::DiagonalDown);
	border->SetLineStyle(LineStyleType::None);
	cr->GetBorders()->Get(BordersLineType::DiagonalUp)->SetLineStyle(LineStyleType::None);
	cr->GetBorders()->SetColor(Spire::Common::Color::GetCadetBlue());

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}