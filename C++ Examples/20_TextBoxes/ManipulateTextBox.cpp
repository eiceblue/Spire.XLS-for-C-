#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring output = OUTPUTPATH;
	wstring inputFile = fn + L"ManipulateTextBoxControl.xlsx";
	wstring outputFile = output + L"ManipulateTextBox.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Get the first textbox
	ITextBox* tb = sheet->GetTextBoxes()->Get(0);

	//Change the text of textbox
	tb->SetText(L"Spire.XLS for C++");

	//Set the alignment of textbox as center
	tb->SetHAlignment(CommentHAlignType::Center);
	tb->SetVAlignment(CommentVAlignType::Center);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}