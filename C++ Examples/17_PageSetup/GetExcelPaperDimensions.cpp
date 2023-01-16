#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring output = OUTPUTPATH;
	wstring outputFile = output + L"GetExcelPaperDimensions.txt";
	wfstream ofs;

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	wstring* content = new wstring();

	//Get the dimensions of A2 paper.
	sheet->GetPageSetup()->SetPaperSize(PaperSizeType::A2Paper);
	content->append(L"A2Paper: " + to_wstring(sheet->GetPageSetup()->GetPageWidth()) + L" x " + to_wstring(sheet->GetPageSetup()->GetPageHeight()) + L"\n");

	//Get the dimensions of A3 paper.
	sheet->GetPageSetup()->SetPaperSize(PaperSizeType::PaperA3);
	content->append(L"PaperA3: " + to_wstring(sheet->GetPageSetup()->GetPageWidth()) + L" x " + to_wstring(sheet->GetPageSetup()->GetPageHeight()) + L"\n");

	//Get the dimensions of A4 paper.
	sheet->GetPageSetup()->SetPaperSize(PaperSizeType::PaperA4);
	content->append(L"PaperA4: " + to_wstring(sheet->GetPageSetup()->GetPageWidth()) + L" x " + to_wstring(sheet->GetPageSetup()->GetPageHeight()) + L"\n");

	//Get the dimensions of paper letter.
	sheet->GetPageSetup()->SetPaperSize(PaperSizeType::PaperLetter);
	content->append(L"PaperLetter: " + to_wstring(sheet->GetPageSetup()->GetPageWidth()) + L" x " + to_wstring(sheet->GetPageSetup()->GetPageHeight()) + L"\n");

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << *content << endl;
	ofs.close();
	workbook->Dispose();
}