# Number Duck
v3.0.2 [J160]
Copyright (C) 2012-2025, File Scribe
https://numberduck.com

## 🦆 About
Number Duck is a programming library for developers to read and write Microsoft Excel compatible spreadsheets from a variety of languages.

See NumberDuck.html for API details, or check the example folders.

## 🚧 Installation
Since Number Duck is delivered as source, you can just drop the `NumberDuck.cpp` and `NumberDuck.hpp` files into your project.

## 💾 Saving and Loading a Spreadsheet
Here is a simple example of writing to a spreadsheet, and then reading it back in.

```cpp
#include "../NumberDuck.hpp"

#include <stdio.h>

using namespace NumberDuck;

int main(int argc, char **argv)
{
	printf("Simple Example\n");
	printf("Create a spreadsheet!\n\n");

	Workbook* pWorkbook = new Workbook(Workbook::License::AGPL);
	Worksheet* pWorksheet = pWorkbook->GetWorksheetByIndex(0);

	Cell* pCell = pWorksheet->GetCellByAddress("A1");
	pCell->SetString("Totally cool spreadsheet!");

	pWorksheet->GetCell(1,1)->SetFloat(3.1417f);

	pWorkbook->Save("SimpleExample.xls", Workbook::FileType::XLS);
	delete pWorkbook;

	Workbook* pWorkbookIn = new Workbook(Workbook::License::AGPL);
	if (pWorkbookIn->Load("SimpleExample.xls"))
	{
		Worksheet* pWorksheetIn = pWorkbookIn->GetWorksheetByIndex(0);
		Cell* pCellIn = pWorksheetIn->GetCell(0,0);
		printf("Cell Contents: %s\n", pCellIn->GetString());
	}

	delete pWorkbookIn;

	return 0;
}
```

More how to code in the example folders above, or at https://numberduck.com/docs

## 👮 License
Number Duck is dual licensed as Open Source (AGPLv3) and commercial closed source.
Closed source licenses may be purchased from https://numberduck.com

Additional third party libraries (TinyXML2, Miniz) are licensed seperately, see License-ThirdParty.txt for more information.