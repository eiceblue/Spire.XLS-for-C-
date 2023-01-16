#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;

	for (int n = 0; n < 50; n++)
	{
		Workbook* workbook = new Workbook();
		workbook->CreateEmptySheets(5);
		for (int i = 0; i < 5; i++)
		{
			Worksheet* sheet = workbook->GetWorksheets()->Get(i);
			sheet->SetName((L"Sheet" + std::to_wstring(i)).c_str());
			for (int row = 1; row <= 150; row++)
			{
				for (int col = 1; col <= 50; col++)
				{
					sheet->GetRange(row, col)->SetText((L"row" + std::to_wstring(row) + L" col" + std::to_wstring(col)).c_str());
				}

			}
		}
		workbook->SaveToFile((output_path + L"Workbook" + std::to_wstring(n) + L".xlsx").c_str(), ExcelVersion::Version2010);
		workbook->Dispose();
	}
}

