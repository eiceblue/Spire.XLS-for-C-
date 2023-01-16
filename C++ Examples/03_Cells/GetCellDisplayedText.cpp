#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"GetCellDisplayedText_result.txt";
	wfstream ofs;

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Set value for B8
	CellRange* cell = sheet->GetRange(L"B8");
	cell->SetNumberValue(0.012345);

	//Set the cell style
	CellStyle* style = cell->GetStyle();
	style->SetNumberFormat(L"0.00");

	//Get the cell value
	wstring cellValue = cell->GetValue();

	//Get the displayed text of the cell
	wstring displayedText = cell->GetDisplayedText();

	//Create StringBuilder to save 
	wstring* content = new wstring();

	//Set string format for displaying
	wstring result = L"B8 Value: " + cellValue + L"\nB8 displayed text: " + displayedText;

	//Add result string to StringBuilder
	content->append(result);

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << *content << endl;
	ofs.close();
	workbook->Dispose();
}

