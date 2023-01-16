#include "pch.h"
using namespace Spire::Xls;

int main() {
    wstring output_path = OUTPUTPATH;
    wstring outputFile = output_path + L"SetPositionAndAlignment.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Set two font styles which will be used in comments
	ExcelFont* font1 = workbook->CreateExcelFont();
	font1->SetFontName(L"Calibri");
	font1->SetColor(Spire::Common::Color::GetFirebrick());
	font1->SetIsBold(true);
	font1->SetSize(12);
	ExcelFont* font2 = workbook->CreateExcelFont();
	font2->SetFontName(L"Calibri");
	font2->SetColor(Spire::Common::Color::GetBlue());
	font2->SetSize(12);
	font2->SetIsBold(true);

	//Add comment 1 and set its size, text, position and alignment
	sheet->GetRange(L"G5")->SetText(L"Spire.XLS");
	IComment* Comment1 = sheet->GetRange(L"G5")->GetComment();
	Comment1->SetIsVisible(true);
	Comment1->SetHeight(150);
	Comment1->SetWidth(300);
	Comment1->GetRichText()->SetText(L"Spire.XLS for .Net:\nStandalone Excel component to meet your needs for conversion, data manipulation, charts in workbook etc. ");
	Comment1->GetRichText()->SetFont(0, 19, font1);
	Comment1->SetTextRotation(TextRotationType::LeftToRight);
	//Set the position of Comment
	Comment1->SetTop(20);
	Comment1->SetLeft(40);
	//Set the alignment of text in Comment
	Comment1->SetVAlignment(CommentVAlignType::Center);
	Comment1->SetHAlignment(CommentHAlignType::Justified);

	//Add comment2 and set its size, text, position and alignment for comparison
	sheet->GetRange(L"D14")->SetText(L"E-iceblue");
	IComment* Comment2 = sheet->GetRange(L"D14")->GetComment();
	Comment2->SetIsVisible(true);
	Comment2->SetHeight(150);
	Comment2->SetWidth(300);
	Comment2->GetRichText()->SetText(L"About E-iceblue: \nWe focus on providing excellent office components for developers to operate Word, Excel, PDF, and PowerPoint documents.");
	Comment2->SetTextRotation(TextRotationType::LeftToRight);
	Comment2->GetRichText()->SetFont(0, 16, font2);
	//Set the position of Comment
	Comment2->SetTop(170);
	Comment2->SetLeft(450);
	//Set the alignment of text in Comment
	Comment2->SetVAlignment(CommentVAlignType::Top);
	Comment2->SetHAlignment(CommentHAlignType::Justified);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
