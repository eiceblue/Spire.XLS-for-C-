#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"AddATotalRowToTable.xlsx";
	wstring outputFile = output_path + L"AddATotalRowToTable_result.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Create a table with the data from the specific cell range.
	IListObject* table = sheet->GetListObjects()->Create(L"Table", sheet->GetRange(L"A1:D4"));

	//Display total row.
	table->SetDisplayTotalRow(true);

	//Add a total row.
	Spire::Common::IList<IListObjectColumn>* list = table->GetColumns();
	list->GetItem(0)->SetTotalsRowLabel(L"Total");
	list->GetItem(1)->SetTotalsCalculation(ExcelTotalsCalculation::Sum);
	list->GetItem(2)->SetTotalsCalculation(ExcelTotalsCalculation::Sum);
	list->GetItem(3)->SetTotalsCalculation(ExcelTotalsCalculation::Sum);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

