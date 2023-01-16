#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring inputFolder = DATAPATH;
	wstring outputFolder = OUTPUTPATH;
	wstring inputFile = inputFolder + L"WorksheetSample1.xlsx";
	wstring outputFile = outputFolder + L"AddCommentWithAuthor_out.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Get the range that will add comment
	CellRange* range = sheet->GetRange(L"C1");

	//Set the author and comment content
	wstring author = L"E-iceblue";
	wstring text = L"This is demo to show how to add a comment with editable Author property.";

	//Add comment to the range and set properties
	ExcelComment* comment = range->AddExcelComment();
	comment->SetWidth(200);
	comment->SetVisible(true);
	comment->SetText(author.empty() ? text.c_str() : (author + L":\n" + text).c_str());

	//Set the font of the author
	IFont* font = range->GetWorksheet()->GetWorkbook()->CreateFont();
	font->SetFontName(L"Tahoma");
	font->SetKnownColor(ExcelColors::Black);
	font->SetIsBold(true);
	comment->GetRichText()->SetFont(0, author.length(), font);

	//Save to file
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}