# Number Duck
v3.0.7 [J16]

Copyright (C) 2012-2025, File Scribe

https://numberduck.com

## 🦆 About
Number Duck is a programming library for developers to read and write Microsoft Excel compatible spreadsheets from a variety of languages.

See NumberDuck.html for API details, or check the example folders.

## 🚧 Installation
There are currently three ways to build Number Duck.

### Source
Since Number Duck is delivered as source, you can just drop the `NumberDuck.cpp` and `NumberDuck.hpp` files into your project.

### CMake
Alternatively you can use CMake to build and install a static library locally, by calling CMake with the `SIMPLE_OUTPUT` flag, eg:

```
cmake -S . -B _build -DSIMPLE_OUTPUT=true
cmake --build _build
sudo cmake --install _build
```

This will build Number Duck (in the `_build` directory) and copy it to the default search paths for includes and libraries, so then you can just `#include <NumberDuck.hpp>` in your code and link to `NumberDuck`, eg:

```
g++ SimpleExample.cpp -lNumberDuck
```

### CPM.cmake

Number Duck has been tested with the [CPM.cmake](https://github.com/cpm-cmake/CPM.cmake) dependancy manager and works successfully.

```
CPMAddPackage("gh:NumberDuck/NumberDuck-CPlusPlus@3.0.7")
````

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