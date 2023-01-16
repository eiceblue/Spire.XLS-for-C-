#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output = OUTPUTPATH;
	wstring outputFile = output + L"GetTextBoxByName.txt";
	wfstream ofs;

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Insert a TextBox
	sheet->GetRange(L"A2")->SetText(L"Name£º");
	ITextBoxShape* textBox = sheet->GetTextBoxes()->AddTextBox(2, 2, 18, 65);

	//Set the name 
	textBox->SetName(L"FirstTextBox");

	//Set string text for TextBox 
	textBox->SetText(L"Spire.XLS for C++ is a professional Excel  C++ API that can be used to create, read, write and convert Excel files in any type of C++ application.\n");

	//Get the TextBox by the name
	ITextBoxShape* FindTextBox = sheet->GetTextBoxes()->Get(L"FirstTextBox");

	//Get the TextBox text 
	wstring text = FindTextBox->GetText();

	//Create StringBuilder to save 
	wstring* content = new wstring();

	wstring name = textBox->GetName();
	//Set string format for displaying
	wstring result = L"The text of \"" + name +L"\" is :" + text;

	//Add result string to StringBuilder
	content->append(result);

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << *content << endl;
	ofs.close();
	workbook->Dispose();
}