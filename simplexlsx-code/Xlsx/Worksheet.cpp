/*
  SimpleXlsxWriter
  Copyright (C) 2012-2018 Pavel Akimov <oxod.pavel@gmail.com>, Alexandr Belyak <programmeralex@bk.ru>

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

#include <stdlib.h>
#include <limits>
#include <iomanip>

#include "Worksheet.h"
#include "XlsxHeaders.h"
#include "Drawing.h"

#include "../PathManager.hpp"
#include "../XMLWriter.hpp"

namespace SimpleXlsx
{
    // ****************************************************************************
    /// @brief  Appends another a group of cells into a row
    /// @param  data template data value
    /// @param	style style index
    /// @return no
    // ****************************************************************************
    template<typename T>
    static void AddCellRoutine_zzz( T data, size_t style, const std::string & CellCoord, XMLWriter & xmlw )
    {
        xmlw.Tag( "c" ).Attr( "r", CellCoord );
        if( style != 0 )    // default style is not necessary to sign explicitly
            xmlw.Attr( "s", style );
        xmlw.TagOnlyContent( "v", data ).End( "c" );
    }


    // ****************************************************************************
    /// @brief      The class constructor
    /// @param      index index of a sheet to be created (for example, sheet1.xml)
    /// @return     no
    /// @note	    The constructor creates an instance with specified sheetX.xml filename
    ///             and without any frozen panes
    // ****************************************************************************
    CWorksheet::CWorksheet( size_t index, CDrawing & drawing, PathManager & pathmanager ) : m_pathManager( pathmanager ), m_Drawing( drawing )
    {
        std::vector<ColumnWidth> colWidths;
        Init( index, 0, 0, colWidths );
    }

    // ****************************************************************************
    /// @brief      The class constructor to create sheet with frozen pane
    /// @param      index index of a sheet to be created (for example, sheet1.xml)
    /// @param      width frozen pane with
    /// @param      height frozen pane height
    /// @return     no
    // ****************************************************************************
    CWorksheet::CWorksheet( size_t index, uint32_t width, uint32_t height,
                            CDrawing & drawing, PathManager & pathmanager ) : m_pathManager( pathmanager ), m_Drawing( drawing )
    {
        std::vector<ColumnWidth> colWidths;
        Init( index, width, height, colWidths );
    }

    // ****************************************************************************
    /// @brief      The class constructor
    /// @param      index index of a sheet to be created (for example, sheet1.xml)
    /// @param		colWidths list of pairs colNumber:Width
    /// @return     no
    /// @note	    The constructor creates an instance with specified sheetX.xml filename
    ///             and without any frozen panes
    // ****************************************************************************
    CWorksheet::CWorksheet( size_t index, const std::vector<ColumnWidth> & colWidths,
                            CDrawing & drawing, PathManager & pathmanager ) : m_pathManager( pathmanager ), m_Drawing( drawing )
    {
        Init( index, 0, 0, colWidths );
    }

    // ****************************************************************************
    /// @brief      The class constructor to create sheet with frozen pane
    /// @param      index index of a sheet to be created (for example, sheet1.xml)
    /// @param      width frozen pane with
    /// @param      height frozen pane height
    /// @param		colWidths list of pairs colNumber:Width
    /// @return     no
    // ****************************************************************************
    CWorksheet::CWorksheet( size_t index, uint32_t width, uint32_t height, const std::vector<ColumnWidth> & colWidths,
                            CDrawing & drawing, PathManager & pathmanager ) : m_pathManager( pathmanager ), m_Drawing( drawing )
    {
        Init( index, width, height, colWidths );
    }

    // ****************************************************************************
    /// @brief  The class destructor (virtual)
    /// @return no
    // ****************************************************************************
    CWorksheet::~CWorksheet()
    {
        if( m_XMLWriter != NULL ) delete m_XMLWriter;
    }

    // ****************************************************************************
    /// @brief  Initializes internal variables and creates a header of a sheet xml tree
    /// @param  index index of a sheet to be created (for example, sheet1.xml)
    /// @param  frozenWidth frozen pane with
    /// @param  frozenHeight frozen pane height
    /// @param	colWidths list of pairs colNumber:Width
    /// @return no
    // ****************************************************************************
    void CWorksheet::Init( size_t index, uint32_t frozenWidth, uint32_t frozenHeight, const std::vector<ColumnWidth> & colWidths )
    {
        m_isOk = true;
        m_row_opened = false;
        m_current_column = 0;
        m_offset_column = 0;
        m_index = index;
        m_title = "Sheet 1";
        m_withFormula = false;
        m_withComments = false;
        m_calcChain.clear();
        m_sharedStrings = NULL;
        m_comments = NULL;
        m_mergedCells.clear();
        m_row_index = 0;
        m_page_orientation = PAGE_PORTRAIT;

        std::stringstream FileName;
        FileName << "/xl/worksheets/sheet" << m_index << ".xml";
        m_XMLWriter = new XMLWriter( m_pathManager.RegisterXML( FileName.str() ) );
        if( ( m_XMLWriter == NULL ) || ! m_XMLWriter->IsOk() )
        {
            m_isOk = false;
            return;
        }

        m_XMLWriter->Tag( "worksheet" ).Attr( "xmlns", ns_book ).Attr( "xmlns:r", ns_book_r ).Attr( "xmlns:mc", ns_mc ).Attr( "mc:Ignorable", "x14ac" ).Attr( "xmlns:x14ac", ns_x14ac );
        m_XMLWriter->TagL( "dimension" ).Attr( "ref", "A1" ).EndL();
        m_XMLWriter->Tag( "sheetViews" ).Tag( "sheetView" ).Attr( "tabSelected", 0 ).Attr( "workbookViewId", 0 );
        if( frozenWidth != 0 || frozenHeight != 0 )
            AddFrozenPane( frozenWidth, frozenHeight );
        m_XMLWriter->End( "sheetView" ).End( "sheetViews" );

        m_XMLWriter->TagL( "sheetFormatPr" ).Attr( "defaultRowHeight", 15 ).Attr( "x14ac:dyDescent", 0.25 ).EndL();
        if( ! colWidths.empty() )
        {
            m_XMLWriter->Tag( "cols" );
            for( std::vector<ColumnWidth>::const_iterator it = colWidths.begin(); it != colWidths.end(); it++ )
                m_XMLWriter->TagL( "col" ).Attr( "min", it->colFrom + 1 ).Attr( "max", it->colTo + 1 ).Attr( "width", it->width ).EndL();
            m_XMLWriter->End( "cols" );
        }
        m_XMLWriter->Tag( "sheetData" );    // open sheetData tag
    }

    // ****************************************************************************
    /// @brief	Generates a header for another row
    /// @param	height row height (default if 0)
    /// @return	no
    // ****************************************************************************
    void CWorksheet::BeginRow( double height )
    {
        if( m_row_opened ) m_XMLWriter->End( "row" );
        m_XMLWriter->Tag( "row" ).Attr( "r", ++m_row_index ).Attr( "x14ac:dyDescent", 0.25 );

        if( height > 0 ) m_XMLWriter->Attr( "ht", height ).Attr( "customHeight", 1 );

        m_current_column = 0;
        m_row_opened = true;

    }

    // ****************************************************************************
    /// @brief	Add string-formatted cell with specified style
    /// @param	data reference to data
    /// @return	no
    // ****************************************************************************
    void CWorksheet::AddCell( const std::string & value, size_t style_id )
    {
        if( value != "" )
        {
            const std::string & szCoord = CellCoord( m_row_index, m_offset_column + m_current_column ).ToString();
            m_XMLWriter->Tag( "c" ).Attr( "r", szCoord );

            if( style_id != 0 ) m_XMLWriter->Attr( "s", style_id );  // default style is not necessary to sign explisitly

            if( value[0] == '=' )
            {
                m_XMLWriter->TagOnlyContent( "f", value.substr( 1, value.length() - 1 ) );

                m_withFormula = true;
                m_calcChain.push_back( szCoord );
            }
            else
            {
                if( m_sharedStrings != NULL )
                {
                    uint64_t str_index = 0;
                    std::map<std::string, uint64_t>::iterator it = m_sharedStrings->find( value );
                    if( it == m_sharedStrings->end() )
                    {
                        str_index = m_sharedStrings->size();
                        ( *m_sharedStrings )[ value ] = str_index;
                    }
                    else
                    {
                        str_index = it->second;
                    }
                    m_XMLWriter->Attr( "t", "s" ).TagOnlyContent( "v", str_index );
                }
                else
                {
                    m_XMLWriter->TagOnlyContent( "v", value );
                }
            }
            m_XMLWriter->End( "c" );
        }
        ///  empty cell with style   ---
        else if( style_id != 0 )
        {
            const std::string & szCoord = CellCoord( m_row_index, m_offset_column + m_current_column ).ToString();
            m_XMLWriter->Tag( "c" ).Attr( "r", szCoord ).Attr( "s", style_id ).End( "c" );
        };
        ///  empty cell with style   ---
        m_current_column++;
    }

    // ****************************************************************************
    /// @brief	Add time-formatted cell with specified style
    /// @param	data reference to data
    /// @return	no
    // ****************************************************************************
    void CWorksheet::AddCell( time_t value, size_t style_id )
    {
        const int64_t secondsFrom1900to1970 = 2208988800u;
        const double excelOneSecond = 0.0000115740740740741;

        struct tm * t = localtime( & value );

        time_t timeSinceEpoch = t->tm_sec + t->tm_min * 60 + t->tm_hour * 3600 + t->tm_yday * 86400 +
                                ( t->tm_year - 70 ) * 31536000 + ( ( t->tm_year - 69 ) / 4 ) * 86400 -
                                ( ( t->tm_year - 1 ) / 100 ) * 86400 + ( ( t->tm_year + 299 ) / 400 ) * 86400;

        double CalcedValue = excelOneSecond * ( secondsFrom1900to1970 + timeSinceEpoch ) + 2;

        std::streamsize OldPrecision = m_XMLWriter->SetFloatPrecision( std::numeric_limits<double>::digits10 );
        //AddCellRoutine( CalcedValue, style_id );
        AddCellRoutine_zzz( CalcedValue, style_id, GetCellCoordStrAndIncColumn(), * m_XMLWriter );
        m_XMLWriter->SetFloatPrecision( OldPrecision );
    }

    void CWorksheet::AddCell( int32_t value, size_t style_id )
    {
        AddCellRoutine_zzz( value, style_id, GetCellCoordStrAndIncColumn(), * m_XMLWriter );
    }

    void CWorksheet::AddCell( uint32_t value, size_t style_id )
    {
        AddCellRoutine_zzz( value, style_id, GetCellCoordStrAndIncColumn(), * m_XMLWriter );
    }

    void CWorksheet::AddCell( uint64_t value, size_t style_id )
    {
        AddCellRoutine_zzz( value, style_id, GetCellCoordStrAndIncColumn(), * m_XMLWriter );
    }

    void CWorksheet::AddCell( double value, size_t style_id )
    {
        AddCellRoutine_zzz( value, style_id, GetCellCoordStrAndIncColumn(), * m_XMLWriter );
    }

    void CWorksheet::AddCell( float value, size_t style_id )
    {
        AddCellRoutine_zzz( value, style_id, GetCellCoordStrAndIncColumn(), * m_XMLWriter );
    }

    // ****************************************************************************
    /// @brief	Closes previously began row
    /// @return	no
    // ****************************************************************************
    void CWorksheet::EndRow()
    {
        if( ! m_row_opened ) return;
        m_XMLWriter->End( "row" );
        m_row_opened = false;
    }


    // ****************************************************************************
    /// @brief  Internal initializatino method adds frozen pane`s information into sheet
    /// @param  width frozen pane width (in number of cells)
    /// @param  height frozen pane height (in number of cells)
    /// @return no
    // ****************************************************************************
    void CWorksheet::AddFrozenPane( uint32_t width, uint32_t height )
    {
        m_XMLWriter->Tag( "pane" );
        if( width != 0 ) m_XMLWriter->Attr( "xSplit", width );
        if( height != 0 ) m_XMLWriter->Attr( "ySplit", height );

        const std::string & szCoord = CellCoord( height + 1, width ).ToString();
        m_XMLWriter->Attr( "topLeftCell", szCoord );

        const char * ActivePane = "activePane";
        if( ( width != 0 ) && ( height != 0 ) ) m_XMLWriter->Attr( ActivePane, "bottomRight" );
        else if( ( width == 0 ) && ( height != 0 ) ) m_XMLWriter->Attr( ActivePane, "bottomLeft" );
        else if( ( width != 0 ) && ( height == 0 ) ) m_XMLWriter->Attr( ActivePane, "topRight" );

        m_XMLWriter->Attr( "state", "frozen" ).End( "pane" );

        if( ( width > 0 ) && ( height > 0 ) )
        {
            const std::string & szCoordBL = CellCoord( height + 1, width - 1 ).ToString();
            const std::string & szCoordTR = CellCoord( height, width ).ToString();
            m_XMLWriter->TagL( "selection" ).Attr( "pane", "topRight" ).Attr( "activeCell", szCoordTR ).Attr( "sqref", szCoordTR ).EndL();
            m_XMLWriter->TagL( "selection" ).Attr( "pane", "bottomLeft" ).Attr( "activeCell", szCoordBL ).Attr( "sqref", szCoordBL ).EndL();
            m_XMLWriter->TagL( "selection" ).Attr( "pane", "bottomRight" ).Attr( "activeCell", szCoord ).Attr( "sqref", szCoord ).EndL();
        }
        else if( width > 0 )
        {
            m_XMLWriter->TagL( "selection" ).Attr( "pane", "topRight" ).Attr( "activeCell", szCoord ).Attr( "sqref", szCoord ).EndL();
        }
        else if( height > 0 )
            m_XMLWriter->TagL( "selection" ).Attr( "pane", "bottomLeft" ).Attr( "activeCell", szCoord ).Attr( "sqref", szCoord ).EndL();
    }

    void CWorksheet::AddRowHeader( std::size_t Size, double Height )
    {
        std::stringstream Spans;
        Spans << m_offset_column + 1 << ':' << Size + m_offset_column + 1;
        m_XMLWriter->Tag( "row" ).Attr( "r", ++m_row_index ).Attr( "spans", Spans.str() ).Attr( "x14ac:dyDescent", 0.25 );
        if( Height > 0 )
            m_XMLWriter->Attr( "ht", Height ).Attr( "customHeight", 1 );
    }

    void CWorksheet::AddRowFooter() const
    {
        m_XMLWriter->End( "row" );
    }

    // ****************************************************************************
    /// @brief  Appends merged cells range into the sheet
    /// @param  cellFrom (row value from 1, col value from 0)
    /// @param  cellTo (row value from 1, col value from 0)
    /// @return no
    // ****************************************************************************
    void CWorksheet::MergeCells( CellCoord cellFrom, CellCoord cellTo )
    {
        if( ( cellFrom.row == 0 ) || ( cellTo.row == 0 ) ) return;
        m_mergedCells.push_back( cellFrom.ToString() + ':' + cellTo.ToString() );
    }

    // ****************************************************************************
    ///	@brief	Receives next to write cell`s coordinates
    /// @param	currCell (row value from 1, col value from 0)
    /// @return	no
    // ****************************************************************************
    void CWorksheet::GetCurrentCellCoord( CellCoord & currCell ) const
    {
        currCell.row = m_row_index;
        currCell.col = m_current_column;
    }

    // ****************************************************************************
    /// @brief  Saves current xml document into a file with preset name
    /// @return Boolean result of the operation
    // ****************************************************************************
    bool CWorksheet::Save()
    {
        m_XMLWriter->End( "sheetData" );    // close sheetData tag

        if( ! m_mergedCells.empty() )
        {
            m_XMLWriter->Tag( "mergeCells" ).Attr( "count", m_mergedCells.size() );
            for( std::list<std::string>::const_iterator it = m_mergedCells.begin(); it != m_mergedCells.end(); it++ )
                m_XMLWriter->TagL( "mergeCell" ).Attr( "ref", * it ).EndL();
            m_XMLWriter->End( "mergeCells" );
        }
        std::string sOrient;
        if( m_page_orientation == PAGE_PORTRAIT ) sOrient = "portrait";
        else if( m_page_orientation == PAGE_LANDSCAPE ) sOrient = "landscape";

        m_XMLWriter->TagL( "pageMargins" ).Attr( "left", 0.7 ).Attr( "right", 0.7 ).Attr( "top", 0.75 );
        /*                  */m_XMLWriter->Attr( "bottom", 0.75 ).Attr( "header", 0.3 ).Attr( "footer", 0.3 ).EndL();
        m_XMLWriter->TagL( "pageSetup" ).Attr( "paperSize", 9 ).Attr( "orientation", sOrient ).EndL();  // A4 paper size

        size_t rId = 1;
        if( ! m_Drawing.IsEmpty() )
        {
            std::stringstream rIdStream;
            rIdStream << "rId" << rId;
            m_XMLWriter->TagL( "drawing" ).Attr( "r:id", rIdStream.str() ).EndL();
            rId++;
        }
        if( m_withComments )
        {
            std::stringstream rIdStream;
            rIdStream << "rId" << rId;
            m_XMLWriter->TagL( "legacyDrawing" ).Attr( "r:id", rIdStream.str() ).EndL();
            rId += 2;
        }

        m_XMLWriter->End( "worksheet" );

        // by deleting the stream the end of file writes and closes the stream
        delete m_XMLWriter;
        m_XMLWriter = NULL;

        if( ( rId != 1 ) && ! SaveSheetRels() ) return false;

        m_isOk = false;
        return true;
    }

    // ****************************************************************************
    /// @brief  Saves current sheet relations file
    /// @return no
    // ****************************************************************************
    bool CWorksheet::SaveSheetRels()
    {
        // [- zip/xl/worksheets/_rels/sheetN.xml.rels
        std::stringstream FileName;
        FileName << "/xl/worksheets/_rels/sheet" << m_index << ".xml.rels";

        XMLWriter xmlw( m_pathManager.RegisterXML( FileName.str() ) );
        xmlw.Tag( "Relationships" ).Attr( "xmlns", ns_relationships );
        size_t rId = 1;
        if( ! m_Drawing.IsEmpty() )
        {
            std::stringstream rIdStream, Drawing;
            rIdStream << "rId" << rId;
            Drawing << "../drawings/drawing" << m_index << ".xml";
            xmlw.TagL( "Relationship" ).Attr( "Id", rIdStream.str() ).Attr( "Type", type_drawing ).Attr( "Target", Drawing.str() ).EndL();
            rId++;
        }
        if( m_withComments )
        {
            std::stringstream Vml, Comments, rIdVml, rIdComments;
            Vml << "../drawings/vmlDrawing" << m_index << ".vml";
            Comments << "../comments" << m_index << ".xml";
            rIdVml << "rId" << rId;
            rIdComments << "rId" << rId + 1;
            rId += 2;
            xmlw.TagL( "Relationship" ).Attr( "Id", rIdVml.str() ).Attr( "Type", type_vml ).Attr( "Target", Vml.str() ).EndL();
            xmlw.TagL( "Relationship" ).Attr( "Id", rIdComments.str() ).Attr( "Type", type_comments ).Attr( "Target", Comments.str() ).EndL();
        }

        xmlw.End( "Relationships" );

        // zip/xl/worksheets/_rels/sheetN.xml.rels -]
        return true;
    }

}	// namespace SimpleXlsx
