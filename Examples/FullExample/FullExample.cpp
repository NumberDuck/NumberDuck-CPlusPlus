#include "../../NumberDuck.hpp"

#include <stdio.h>
#include <string>

using namespace NumberDuck;

int main(int argc, char **argv)
{
	printf("Full Example\n");
	printf("General cool things!\n\n");

	// This example runs through various API calls to give you an idea of all the varied and wonderful things NumberDuck can do.

	// First lets create a new Workbook and set the name of the worksheet in it.
	Workbook* pWorkbook = new Workbook(Workbook::License::AGPL);
	Worksheet* pWorksheet = pWorkbook->GetWorksheetByIndex(0);
	pWorksheet->SetName("Synergy");

	// Let's set a cool title in big letters so people know we mean business.
	// Firstly we create the cell style, making sure we use a serious business font.
	// Then we get which cell we want to edit, apply the cell style and set the contents.
	Style* pStyle = pWorkbook->CreateStyle();
	pStyle->GetFont()->SetName("Comic Sans MS");
	pStyle->GetFont()->SetSize(96);
	Color* pColor = new Color(0, 144, 255);
	pStyle->GetFont()->GetColor(true)->SetFromColor(pColor);
	delete pColor;
	pStyle->GetFont()->SetBold(true);

	Cell* pCell = pWorksheet->GetCellByAddress("A1");
	pCell->SetString("Serious Business Inc.");
	pCell->SetStyle(pStyle);

	// Repeat the process for our cool slogan. Note that we have to create a new style.
	// If we edited the existing style that would change the title text as well!
	// When we grab our cell to write this time, we do so using the coordinates instead of the A1 reference style.
	pStyle = pWorkbook->CreateStyle();
	pStyle->GetFont()->SetName("Times New Roman");
	pStyle->GetFont()->SetItalic(true);

	pCell = pWorksheet->GetCell(0,1);     // note that we are using an x,y reference here
	pCell->SetString("Because business is Serious Business!");
	pCell->SetStyle(pStyle);

	// Let's create a new worksheet and throw some numbers into it, using the boring old default style, since accountants are boring and old.
	// Note that it's the same function name to set a string or a number.
	// We also set the column to be 120 pixels wide so our text does not get cut off.
	pWorksheet = pWorkbook->CreateWorksheet();
	pWorksheet->SetName("Paradigms");

	pWorksheet->SetColumnWidth(0, 120);
	pWorksheet->GetCell(0,0)->SetString("Total Business:");
	pWorksheet->GetCell(1,0)->SetFloat(999);

	// Now we just save to disk so that we can email our cool spreadsheet to all the investors we are courting.
	pWorkbook->Save("FullExample.xls", Workbook::FileType::XLS);

	delete pWorkbook;

	// Just for fun, and to demo more features, let's load that spreadsheet we just saved and read some data.
	Workbook* pWorkbookIn = new Workbook(Workbook::License::AGPL);
	if (pWorkbookIn->Load("FullExample.xls"))
	{
		Worksheet* pWorksheetIn = pWorkbookIn->GetWorksheetByIndex(0);
		Cell* pCellIn = pWorksheetIn->GetCell(0,0);
		printf("Cell Contents: %s\n", pCellIn->GetString());
	}

	delete pWorkbookIn;
	return 0;
}
