#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"DataValidation.xlsx";
    	wstring output_path = OUTPUTPATH;
   	wstring outputFile = output_path + L"DataValidation_out.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Decimal DataValidation
	sheet->GetRange(L"B11")->SetText(L"Input Number(3-6):");
	CellRange* rangeNumber = sheet->GetRange(L"B12");

	//Set the operator for the data validation.
	rangeNumber->GetDataValidation()->SetCompareOperator(ValidationComparisonOperator::Between);

	//Set the value or expression associated with the data validation.
	rangeNumber->GetDataValidation()->SetFormula1(L"3");

	//The value or expression associated with the second part of the data validation.
	rangeNumber->GetDataValidation()->SetFormula2(L"6");

	//Set the data validation type.
	rangeNumber->GetDataValidation()->SetAllowType(CellDataType::Decimal);

	//Set the data validation error message.
	rangeNumber->GetDataValidation()->SetErrorMessage(L"Please input correct number!");

	//Enable the error.
	rangeNumber->GetDataValidation()->SetShowError(true);
	rangeNumber->GetStyle()->SetKnownColor(ExcelColors::Gray25Percent);

	//Date DataValidation
	sheet->GetRange(L"B14")->SetText(L"Input Date:");
	CellRange* rangeDate = sheet->GetRange(L"B15");
	rangeDate->GetDataValidation()->SetAllowType(CellDataType::Date);
	rangeDate->GetDataValidation()->SetCompareOperator(ValidationComparisonOperator::Between);
	rangeDate->GetDataValidation()->SetFormula1(L"1/1/1970");
	rangeDate->GetDataValidation()->SetFormula2(L"12/31/1970");
	rangeDate->GetDataValidation()->SetErrorMessage(L"Please input correct date!");
	rangeDate->GetDataValidation()->SetShowError(true);
	rangeDate->GetDataValidation()->SetAlertStyle(AlertStyleType::Warning);
	rangeDate->GetStyle()->SetKnownColor(ExcelColors::Gray25Percent);

	//TextLength DataValidation
	sheet->GetRange(L"B17")->SetText(L"Input Text:");
	CellRange* rangeTextLength = sheet->GetRange(L"B18");
	rangeTextLength->GetDataValidation()->SetAllowType(CellDataType::TextLength);
	rangeTextLength->GetDataValidation()->SetCompareOperator(ValidationComparisonOperator::LessOrEqual);
	rangeTextLength->GetDataValidation()->SetFormula1(L"5");
	rangeTextLength->GetDataValidation()->SetErrorMessage(L"Enter a Valid String!");
	rangeTextLength->GetDataValidation()->SetShowError(true);
	rangeTextLength->GetDataValidation()->SetAlertStyle(AlertStyleType::Stop);
	rangeTextLength->GetStyle()->SetKnownColor(ExcelColors::Gray25Percent);

	sheet->AutoFitColumn(2);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
