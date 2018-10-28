# QSimpleXlsxWriter

## Introduction

- Use SimpleXlsxWriter in Qt. 
- [SimpleXlsxWriter](https://sourceforge.net/projects/simplexlsx/)
	- C++ library for creating XLSX files for MS Excel 2007 and above.
	- Brought to you by: oxod, programmeralex
	- The main feature of this library is that it uses C++ standard file streams. On the one hand it results in almost unnoticeable memory and CPU resources consumption while processing (that may be very useful at saving a large data arrays), but on the other hand it makes unfeasible to edit data that were written. Hence, if using this library the structure of the future report should be known enough.

## Hello World (HelloQSimpleXlsxWriter)	

```cpp
	#include <iostream>
	#include <stdio.h>
	#include <stdlib.h>
	#include <time.h>

    #include <Xlsx/Workbook.h>

    using namespace std;
    using namespace SimpleXlsx;

    int main()
    {
       srand(time(NULL));
       const int colNum = 20;
       const int rowNum = 10;

       CWorkbook book;

       {    // Creating a simple data sheet
            CWorksheet &sheet = book.AddSheet(_T("New sheet simple"));

            vector<CellDataDbl> data;   // (data:style_index)
            CellDataDbl def;
            def.style_id = 0;       // 0,1 - default styles
            for (int i = 0; i < colNum; i++) {
                def.value = (double)(rand() % 100) / 101.0;
                data.push_back(def);
            }

            for (int i = 0; i < rowNum; i++)
                sheet.AddRow(data);
       }

       {    // Creating data sheet with a frozen pane
            CWorksheet &sheet = book.AddSheet(_T("New sheet with frozen pane"), 0, 1);

            vector<CellDataStr> dataStr;// (data:style_index)
            CellDataStr col;

            col.value = _T("Frozen pane header");
            col.style_id = 0;       // 0,1 - default styles
            dataStr.push_back(col);
            sheet.AddRow(dataStr);
            sheet.MergeCells(CellCoord(1, 0), CellCoord(1, colNum-1)); // merge first row

            vector<CellDataInt> data;   // (data:style_index)
            CellDataInt def;
            def.style_id = 0;       // 0,1 - default styles
            for (int i = 0; i < colNum; i++) {
                def.value = (double)(rand() % 100);
                data.push_back(def);
            }

            for (int i = 0; i < rowNum; i++)
                sheet.AddRow(data);
       }

       {    // Creating data sheet with col and row specified sizes
            std::vector<ColumnWidth> colWidths;
            for (int i = 0; i < 50; i++) {
                ColumnWidth colWidth;
                colWidth.colFrom = colWidth.colTo = i;
                colWidth.width = i + 10;
                colWidths.push_back(colWidth);
            }

            CWorksheet &sheet = book.AddSheet(_T("New sheet with column and row specified sizes"), colWidths);

            Style style;
            style.wrapText = true;
            int style_index = book.m_styleList.Add(style);

            vector<CellDataStr> data;   // (data:style_index)
            CellDataStr def;
            def.style_id = style_index; // 0,1 - default styles
            data.push_back(def);
                TCHAR szText[30] = { 0 };
            for (int i = 0; i < colNum; i++) {
                _stprintf(szText, _T("some\nmultirow\ntext %d_%d"), i+10, i);
                def.value = szText;
                data.push_back(def);
            }

            for (int i = 0; i < rowNum; i++)
                sheet.AddRow(data, 0, 45);
       }

        {   // Creating data sheet with styled cells
            CWorksheet &sheet = book.AddSheet(_T("New sheet with styled cells"));

            Style style;
            style.fill.patternType = PATTERN_NONE;
            style.font.size = 14;
            style.font.theme = true;
            style.font.attributes = FONT_BOLD;
            style.horizAlign = ALIGN_H_RIGHT;
            style.vertAlign = ALIGN_V_CENTER;
            int style_index_1 = book.m_styleList.Add(style);

            style.fill.patternType = PATTERN_NONE;
            style.font.size = 14;
            style.font.theme = true;
            style.font.attributes = FONT_ITALIC;
            style.horizAlign = ALIGN_H_LEFT;
            style.vertAlign = ALIGN_V_CENTER;
            int style_index_2 = book.m_styleList.Add(style);

            vector<CellDataFlt> data;
            CellDataFlt def;
            for (int i = 0; i < colNum; i++) {
                def.value = (double)(rand() % 100) / 101.0;
                if (i % 2 == 0) def.style_id = style_index_1;
                else        def.style_id = style_index_2;

                data.push_back(def);
            }

            for (int i = 0; i < rowNum; i++)
                sheet.AddRow(data, 5);
       }

       bool bRes = book.Save(_T("MyBook.xlsx"));
       if (bRes)   cout << "The book has been saved successfully";
       else        cout << "The book saving has been failed";

       return 0;
    }
```

## License and links
- QSimpleXlsxWriter is under MIT license.
- SimpleXlsxWriter is under zlib license.
- zlib is under zlib license.
