/*
  SimpleXlsxWriter
  Copyright (C) 2012-2020 Pavel Akimov <oxod.pavel@gmail.com>, Alexandr Belyak <programmeralex@bk.ru>

  This software is provided 'as-is', without any express or implied
  warranty. In no event will the authors be held liable for any damages
  arising from the use of this software.

  Permission is granted to anyone to use this software for any purpose,
  including commercial applications, and to alter it and redistribute it
  freely, subject to the following restrictions:

  1. The origin of this software must not be misrepresented; you must not
     claim that you wrote the original software. If you use this software
     in a product, an acknowledgment in the product documentation would be
     appreciated but is not required.
  2. Altered source versions must be plainly marked as such, and must not be
     misrepresented as being the original software.
  3. This notice may not be removed or altered from any source distribution.
*/

#include "Chartsheet.h"

#include "Chart.h"
#include "Drawing.h"
#include "Worksheet.h"
#include "XlsxHeaders.h"

#include "../PathManager.hpp"
#include "../XMLWriter.hpp"

namespace SimpleXlsx
{
    // ****************************************************************************
    /// @brief  The class constructor
    /// @param  index index of a sheet to be created (for example, chart1.xml)
    /// @param	chart is a reference for correspond chart
    /// @param	drawing is a references for correspond drawing
    /// @param  Parent is a pointer to parent CWorkbook
    /// @return no
// ****************************************************************************
const UniString &CChartsheet::GetTitle() const
{
    return m_Chart.GetTitle();
}

CChartsheet::CChartsheet( size_t index, CChart & chart, CDrawing & drawing, PathManager & pathmanager ) :
        CSheet( index ), m_Chart( chart ), m_Drawing( drawing ), m_pathManager( pathmanager )
    {
    }

    // ****************************************************************************
    /// @brief  The class destructor (virtual)
    /// @return no
    // ****************************************************************************
    CChartsheet::~CChartsheet()
    {
    }

    // ****************************************************************************
    /// @brief  For each chart sheet it creates and saves 2 files:
    ///         sheetXX.xml, sheetXX.xml.rels,
    /// @return Boolean result of the operation
    // ****************************************************************************
    bool CChartsheet::Save()
    {
        {
            // [- /xl/chartsheets/_rels/sheetX.xml.rels
            std::stringstream FileName;
            FileName << "/xl/chartsheets/_rels/sheet" << m_index << ".xml.rels";

            std::stringstream Target;
            Target << "../drawings/drawing" << m_Drawing.GetIndex() << ".xml";

            XMLWriter xmlw( m_pathManager.RegisterXML( FileName.str() ) );
            xmlw.Tag( "Relationships" ).Attr( "xmlns", ns_relationships );
            xmlw.TagL( "Relationship" ).Attr( "Id", "rId1" ).Attr( "Type", type_drawing ).Attr( "Target", Target.str() ).EndL();

            xmlw.End( "Relationships" );
            // /xl/chartsheets/_rels/sheetX.xml.rels -]
        }

        {
            // [- /xl/chartsheets/sheetX.xml
            std::stringstream FileName;
            FileName << "/xl/chartsheets/sheet" << m_index << ".xml";

            XMLWriter xmlw( m_pathManager.RegisterXML( FileName.str() ) );
            xmlw.Tag( "chartsheet" ).Attr( "xmlns", ns_book ).Attr( "xmlns:r", ns_book_r ).TagL( "sheetPr" ).EndL();

            xmlw.Tag( "sheetViews" );
            xmlw.TagL( "sheetView" ).Attr( "zoomScale", 85 ).Attr( "workbookViewId", 0 ).Attr( "zoomToFit", 1 ).EndL();
            xmlw.End( "sheetViews" );

            xmlw.TagL( "pageMargins" ).Attr( "left", 0.7 ).Attr( "right", 0.7 ).Attr( "top", 0.75 );
            xmlw.Attr( "bottom", 0.75 ).Attr( "header", 0.3 ).Attr( "footer", 0.3 ).EndL();
            xmlw.TagL( "drawing" ).Attr( "r:id", "rId1" ).EndL();

            xmlw.End( "chartsheet" );
            // /xl/chartsheets/sheetX.xml -]
        }
        return true;
    }

}	// namespace SimpleXlsx
