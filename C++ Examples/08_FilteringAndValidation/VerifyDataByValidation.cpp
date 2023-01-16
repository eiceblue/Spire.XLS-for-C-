#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Sample.xlsx";
    	wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"VerifyDataByValidation.txt";
	wfstream ofs;

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Cell B4 has the Decimal Validation
	CellRange* cell = sheet->GetRange(L"B4");

	//Get the valditation of this cell
	Validation* validation = cell->GetDataValidation();

	//Get the specified data range
	double minimum = std::stod(validation->GetFormula1());
	double maximum = std::stod(validation->GetFormula2());

	//Create StringBuilder to save 
	wstring* content = new wstring();

	//Set different numbers for the cell
	for (int i = 5; i < 100; i = i + 40)
	{
		cell->SetNumberValue(i);
		std::wstring result = L"";
		//Verify 
		if (cell->GetNumberValue() < minimum || cell->GetNumberValue() > maximum)
		{
			//Set string format for displaying
			result = L"Is input " + std::to_wstring(i) + L" a valid value for this Cell: false \n";
		}
		else
		{
			//Set string format for displaying
			result = L"Is input " + std::to_wstring(i) + L" a valid value for this Cell: true \n";
		}
		//Add result string to StringBuilder
		content->append(result);
	}

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << *content << endl;
	ofs.close();
	workbook->Dispose();
}
