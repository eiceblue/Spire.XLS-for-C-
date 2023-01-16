#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"CreateTable.xlsx";
	wstring outputFile = output_path + L"FindAndReplaceData_result.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Find the "Brazil" string
	auto ranges = sheet->FindAllString(L"Area", false, false);

	//Traverse the found ranges
	for (int i = 0; i < ranges->GetCount(); i++)
	{
		CellRange* cr = ranges->GetItem(i);
		//Replace it with "China"
		cr->SetText(L"Area Code");
		//Highlight the color
		cr->GetStyle()->SetColor(Spire::Common::Color::GetYellow());
	}

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();

}

