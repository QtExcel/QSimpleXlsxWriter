# QSimpleXlsxWriter

[![Build Status](https://travis-ci.com/j2doll/QSimpleXlsxWriter.svg?branch=master)](https://travis-ci.com/j2doll/QSimpleXlsxWriter)

## Introduction

- Use SimpleXlsxWriter in Qt. 
- [SimpleXlsxWriter](https://sourceforge.net/projects/simplexlsx/) is C++ library for creating XLSX files for MS Excel 2007 and above.
	- Brought to you by: [oxod](https://sourceforge.net/u/oxod/), [programmeralex](https://sourceforge.net/u/programmeralex/)
	- The main feature of this library is that it uses C++ standard file streams. On the one hand it results in almost unnoticeable memory and CPU resources consumption while processing (that may be very useful at saving a large data arrays), but on the other hand it makes unfeasible to edit data that were written. Hence, if using this library the structure of the future report should be known enough.

## Hello World (Simple.cpp)

```cpp
#include <iostream>
#include <math.h>
#include <stdio.h>
#include <stdlib.h>
#include <time.h>

#include <Xlsx/Workbook.h>

using namespace SimpleXlsx;

int main( int argc, char * argv[] )
{
    ( void )argc; ( void )argv;

    setlocale( LC_ALL, "" );

    time_t CurrentTime = time( NULL );

    CWorkbook book( "Incognito" );

    std::vector<ColumnWidth> ColWidth;
    ColWidth.push_back( ColumnWidth( 0, 3, 20 ) );
    CWorksheet & Sheet = book.AddSheet( "Unicode", ColWidth );

    Style style;
    style.horizAlign = ALIGN_H_CENTER;
    style.font.attributes = FONT_BOLD;
    size_t CenterStyleIndex = book.AddStyle( style );

    Sheet.BeginRow();
    Sheet.AddCell( "Common test of Unicode support", CenterStyleIndex );
    Sheet.MergeCells( CellCoord( 1, 0 ), CellCoord( 1, 3 ) );
    Sheet.EndRow();

    Font TmpFont = book.GetFonts().front();
    TmpFont.attributes = FONT_ITALIC;
    Comment Com;
    Com.x = 250;
    Com.y = 100;
    Com.width = 100;
    Com.height = 30;
    Com.cellRef = CellCoord( 8, 1 );
    Com.isHidden = false;
    Com.AddContent( TmpFont, "Comment with custom style" );
    Sheet.AddComment( Com );

    Sheet.BeginRow();
    Sheet.AddCell( "English language" );
    Sheet.AddCell( "English language" );
    Sheet.EndRow();

    Sheet.BeginRow();
    Sheet.AddCell( "Russian language" );
    Sheet.AddCell( L"Русский язык" );
    Sheet.EndRow();

    Sheet.BeginRow();
    Sheet.AddCell( "Chinese language" );
    Sheet.AddCell( L"中文" );
    Sheet.EndRow();

    Sheet.BeginRow();
    Sheet.AddCell( "French language" );
    Sheet.AddCell( L"le français" );
    Sheet.EndRow();

    Sheet.BeginRow();
    Sheet.AddCell( "Arabic language" );
    Sheet.AddCell( L"العَرَبِيَّة‎‎" );
    Sheet.EndRow();

    Sheet.BeginRow();
    Sheet.EndRow();

    style.fill.patternType = PATTERN_NONE;
    style.font.theme = true;
    style.horizAlign = ALIGN_H_RIGHT;
    style.vertAlign = ALIGN_V_CENTER;

    style.numFormat.numberStyle = NUMSTYLE_MONEY;

    size_t MoneyStyleIndex = book.AddStyle( style );

    Sheet.BeginRow();
    Sheet.AddCell( "Money symbol" );
    Sheet.AddCell( 123.45, MoneyStyleIndex );
    Sheet.EndRow();

    style.numFormat.numberStyle = NUMSTYLE_DATETIME;
    size_t DateTimeStyleIndex = book.AddStyle( style );

    Sheet.BeginRow();
    Sheet.AddCell( "Write date/time" );
    Sheet.AddCell( CurrentTime, DateTimeStyleIndex );
    Sheet.EndRow();

    style.numFormat.formatString = "hh:mm:ss";
    size_t CustomDateTimeStyleIndex = book.AddStyle( style );
    Sheet.BeginRow();
    Sheet.AddCell( "Custom date/time" );
    Sheet.AddCell( CurrentTime, CustomDateTimeStyleIndex );
    Sheet.EndRow();

    Sheet.BeginRow();
    Sheet.EndRow();

    Style stPanel;
    stPanel.border.top.style = BORDER_THIN;
    stPanel.border.bottom.color = "FF000000";
    stPanel.fill.patternType = PATTERN_SOLID;
    stPanel.fill.fgColor = "FFCCCCFF";
    size_t PanelStyleIndex = book.AddStyle( stPanel );
    Sheet.BeginRow();
    Sheet.AddCell( "Cells with border", PanelStyleIndex );
    Sheet.AddCell( "", PanelStyleIndex );
    Sheet.AddCell( "", PanelStyleIndex );
    Sheet.AddCell( "", PanelStyleIndex );
    Sheet.EndRow();

    if( book.Save( "Simple.xlsx" ) ) std::cout << "The book has been saved successfully" << std::endl;
    else std::cout << "The book saving has been failed" << std::endl;

    return 0;
}
```

## License and links
- QSimpleXlsxWriter is under MIT license. [https://github.com/j2doll/QSimpleXlsxWriter](https://github.com/j2doll/QSimpleXlsxWriter)
- SimpleXlsxWriter is under zlib license. [https://sourceforge.net/projects/simplexlsx/](https://sourceforge.net/projects/simplexlsx/)
- zlib is under zlib license. [https://zlib.net/](https://zlib.net/)

## Similar projects

### :star: <b>QXlsx</b> [https://github.com/j2doll/QXlsx](https://github.com/j2doll/QXlsx)

<p align="center"><img src="https://github.com/j2doll/QXlsx/raw/master/markdown.data/QXlsx-Desktop.png"></p>

- QXlsx is excel file(*.xlsx) reader/writer library.
- Because QtXlsx is no longer supported(2014), I created a new project that is based on QtXlsx. (2017-)
- Development language of QXlsx is C++. (with Qt)
- You don't need to use static library or dynamic shared object using QXlsx.

### :star: <b>Qxlnt</b> [https://github.com/j2doll/Qxlnt](https://github.com/j2doll/Qxlnt)

<p align="center"><img src="https://github.com/j2doll/Qxlnt/raw/master/markdown-data/Concept-QXlnt.jpg"></p>

- Qxlnt is a helper project that allows xlnt to be used in Qt.
- xlnt is a excellent C++ library for using xlsx Excel files. :+1:
- I was looking for a way to make it easy to use in Qt. Of course, cmake is compatible with Qt, but it is not convenient to use. So I created Qxlnt.

### :star: <b>Qlibxlsxwriter</b> [https://github.com/j2doll/Qlibxlsxwriter](https://github.com/j2doll/Qlibxlsxwriter)

<p align="center"><img src="https://github.com/j2doll/Qlibxlsxwriter/raw/master/markdown.data/logo.png"></p>

- Qlibxlsxwriter is a helper project that allows libxlsxwriter to be used in Qt.
- libxlsxwriter is a C library for creating Excel XLSX files. :+1:
