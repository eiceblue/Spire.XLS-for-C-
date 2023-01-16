#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring outputFile = output_path + L"CopySheetToAnotherXlsFile.xlsx";
	std::wstring outputFile_ = output_path + L"sourceFile.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Put some data into header rows (A1:A4)
	for (int i = 1; i < 6; i++)
	{
		wstring range = L"A" + to_wstring(i);
		wstring text = L"Header Row " + to_wstring(i);
		sheet->GetRange(range.c_str())->SetText(text.c_str());
	}

	//Put some detail data (A5:A99)
	for (int i = 5; i < 100; i++)
	{
		wstring range = L"A" + to_wstring(i);
		wstring text = L"Detail Row " + to_wstring(i);
		sheet->GetRange(range.c_str())->SetText(text.c_str());
	}
	//Define a pagesetup object based on the first worksheet.
	PageSetup* pageSetup = sheet->GetPageSetup();
	//The first five rows are repeated in each page. It can be seen in print preview.
	pageSetup->SetPrintTitleRows(L"$1:$5");
	//Create another Workbook.
	Workbook* workbook1 = new Workbook();
	//Get the first worksheet in the book.
	Worksheet* sheet1 = workbook1->GetWorksheets()->Get(0);
	//Copy worksheet to destination worsheet in another Excel file.
	sheet1->CopyFrom(sheet);

	//Save the document
	workbook->SaveToFile(outputFile_.c_str(), ExcelVersion::Version2013);
	workbook1->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}