#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"ReadFormulas.xlsx";
	wstring outputFile = output_path + L"ReadFormulas.txt";
	wfstream ofs;

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	wstring formula = sheet->GetRange(L"C14")->GetFormula();
	wstring value = to_wstring(sheet->GetRange(L"C14")->GetFormulaNumberValue());

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << "Formula��" << formula << "\r\n" << "Value��" << value << endl;

	ofs.close();
	workbook->Dispose();
}