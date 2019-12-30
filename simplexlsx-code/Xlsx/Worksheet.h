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

#ifndef XLSX_WORKSHEET_H
#define XLSX_WORKSHEET_H

#include <list>
#include <map>
#include <string>
#include <vector>

#include "SimpleXlsxDef.h"

namespace SimpleXlsx
{
class CDrawing;

class PathManager;
class XMLWriter;

// ****************************************************************************
/// @brief	The class CWorksheet is used for creation and population
///         .xlsx file sheet table with data
// ****************************************************************************
class CWorksheet
{
    public:
        enum EPageOrientation
        {
            PAGE_PORTRAIT = 0,
            PAGE_LANDSCAPE
        };

    private:
        XMLWriter       *       m_XMLWriter;        ///< xml output stream
        std::vector<std::string>m_calcChain;        ///< list of cells with formulae
        std::map<std::string, uint64_t> * m_sharedStrings; ///< pointer to the list of string supposed to be into shared area
        std::vector<Comment> *	m_comments;         ///< pointer to the vector of comments
        std::list<std::string>  m_mergedCells;	///< list of merged cells` ranges (e.g. A1:B2)
        size_t					m_index;            ///< current sheet number
        UniString             	m_title;            ///< page title
        bool                    m_withFormula;      ///< indicates whether the sheet contains formulae
        bool					m_withComments;		///< indicates whether the sheet contains any comments
        bool                    m_isOk;             ///< indicates initialization successfulness
        bool                    m_isDataPresented;  ///< indicates whether the sheet contains a data
        uint32_t				m_row_index;        ///< since data add row-by-row it contains current row to write
        bool					m_row_opened;		///< indicates whether row tag is opened
        uint32_t				m_current_column;	///< used at separate row generation - last cell column number to be added
        uint32_t				m_offset_column;	///< used at entire row addition (implicit parameter for AddCell method)

        EPageOrientation		m_page_orientation;	///< defines page orientation for printing

        PathManager      &      m_pathManager;      ///< reference to XML PathManager
        CDrawing        &       m_Drawing;          ///< Reference to drawing object

    public:
        // *INDENT-OFF*   For AStyle tool

        // @section    SEC_INTERNAL Interclass internal interface methods
        inline bool     IsThereComment() const      { return m_withComments; }
        inline bool     IsThereFormula() const      { return m_withFormula; }
        inline size_t   GetIndex() const            { return m_index; }
        inline const CWorksheet & GetCalcChain( std::vector<std::string> & chain ) const  { chain = m_calcChain; return * this; }

        // @section    SEC_USER User interface
        inline const UniString & GetTitle() const                           { return m_title; }
        inline CWorksheet & SetTitle( const UniString & title )             { if( ! title.empty() ) m_title = title; return * this; }

        inline CWorksheet & SetPageOrientation( EPageOrientation orient )   { m_page_orientation = orient; return * this; }

        // *INDENT-ON*   For AStyle tool

        CWorksheet & AddComment( const Comment & comment )
        {
            if( m_comments == NULL )
                return * this;
            m_comments->push_back( comment );
            m_comments->back().sheetIndex = m_index;
            m_withComments = true;
            return * this;
        }

        // *INDENT-OFF*   For AStyle tool

        CWorksheet & BeginRow( double height = 0 );
        CWorksheet & EndRow();

        inline CWorksheet & AddCell()                                       { m_current_column++; return * this; }
        inline CWorksheet & AddEmptyCells( uint32_t Count )                 { m_current_column += Count; return * this; }

        CWorksheet & AddCell( const std::string & value, size_t style_id = 0 );
        inline CWorksheet & AddCell( const CellDataStr & data )                 { return AddCell( data.value, data.style_id ); }
        CWorksheet & AddCell( const char * value, size_t style_id = 0 )         { return AddCell( std::string( value ), style_id ); }
        CWorksheet & AddCell( const std::wstring & value, size_t style_id = 0 ) { return AddCell( UTF8Encoder::From_wstring( value ), style_id ); }
        inline CWorksheet & AddCells( const std::vector<CellDataStr> & data );

        CWorksheet & AddCell( const CellDataTime & data );
        inline CWorksheet & AddCells( const std::vector<CellDataTime> & data );

        CWorksheet & AddCell( int32_t value, size_t style_id = 0 );
        inline CWorksheet & AddCell( const CellDataInt & data )                 { return AddCell( data.value, data.style_id ); }
        inline CWorksheet & AddCells( const std::vector<CellDataInt> & data )   { return AddCellsTempl( data ); }

        CWorksheet & AddCell( uint32_t value, size_t style_id = 0 );
        inline CWorksheet & AddCell( const CellDataUInt & data )                { return AddCell( data.value, data.style_id ); }
        inline CWorksheet & AddCells( const std::vector<CellDataUInt> & data )  { return AddCellsTempl( data ); }

        CWorksheet & AddCell( int64_t value, size_t style_id = 0 );
        CWorksheet & AddCell( uint64_t value, size_t style_id = 0 );

        CWorksheet & AddCell( float value, size_t style_id = 0 );
        inline CWorksheet & AddCell( const CellDataFlt & data )                 { return AddCell( data.value, data.style_id ); }
        inline CWorksheet & AddCells( const std::vector<CellDataFlt> & data )   { return AddCellsTempl( data ); }

        CWorksheet & AddCell( double value, size_t style_id = 0 );
        inline CWorksheet & AddCell( const CellDataDbl & data )                 { return AddCell( data.value, data.style_id ); }
        inline CWorksheet & AddCells( const std::vector<CellDataDbl> & data )   { return AddCellsTempl( data ); }

        CWorksheet & AddRow( const std::vector<CellDataStr> & data, uint32_t offset = 0, double height = 0.0 )  { return AddRowTempl( data, offset, height ); }
        CWorksheet & AddRow( const std::vector<CellDataTime> & data, uint32_t offset = 0, double height = 0.0 ) { return AddRowTempl( data, offset, height ); }
        CWorksheet & AddRow( const std::vector<CellDataInt> & data, uint32_t offset = 0, double height = 0.0 )  { return AddRowTempl( data, offset, height ); }
        CWorksheet & AddRow( const std::vector<CellDataUInt> & data, uint32_t offset = 0, double height = 0.0 ) { return AddRowTempl( data, offset, height ); }
        CWorksheet & AddRow( const std::vector<CellDataDbl> & data, uint32_t offset = 0, double height = 0.0 )  { return AddRowTempl( data, offset, height ); }
        CWorksheet & AddRow( const std::vector<CellDataFlt> & data, uint32_t offset = 0, double height = 0.0 )  { return AddRowTempl( data, offset, height ); }

        CWorksheet & MergeCells( CellCoord cellFrom, CellCoord cellTo );

        const CWorksheet & GetCurrentCellCoord( CellCoord & currCell ) const;
        inline uint32_t CurrentRowIndex() const     { return m_row_index; }
        inline uint32_t CurrentColumnIndex() const  { return m_current_column; }

        // *INDENT-ON*   For AStyle tool

        bool IsOk() const
        {
            return m_isOk;
        }

    protected:
        CWorksheet( size_t index, CDrawing & drawing, PathManager & pathmanager );
        CWorksheet( size_t index, const std::vector<ColumnWidth> & colHeights, CDrawing & drawing, PathManager & pathmanager );
        CWorksheet( size_t index, uint32_t width, uint32_t height, CDrawing & drawing, PathManager & pathmanager );
        CWorksheet( size_t index, uint32_t width, uint32_t height, const std::vector<ColumnWidth> & colHeights,
                    CDrawing & drawing, PathManager & pathmanager );
        virtual ~CWorksheet();

    private:
        //Disable copy and assignment
        CWorksheet( const CWorksheet & );
        CWorksheet & operator=( const CWorksheet & );

        bool Save();

            // *INDENT-OFF*   For AStyle tool
            inline void     SetSharedStr( std::map<std::string, uint64_t> * share ) { m_sharedStrings = share; }
            inline void     SetComments( std::vector<Comment> * share )             { m_comments = share; }
            // *INDENT-ON*   For AStyle tool

        void Init( size_t index, uint32_t frozenWidth, uint32_t frozenHeight, const std::vector<ColumnWidth> & colHeights );
        void AddFrozenPane( uint32_t width, uint32_t height );

        template<typename T>
        CWorksheet & AddCellsTempl( const std::vector<T> & data );

        void AddRowHeader( std::size_t Size, double Height );
        void AddRowFooter() const;

        template<typename T>
        CWorksheet & AddRowTempl( const std::vector<T> & data, uint32_t offset, double height );

        bool SaveSheetRels();

        inline std::string GetCellCoordStrAndIncColumn()
        {
            std::string Result = CellCoord( m_row_index, m_offset_column + m_current_column ).ToString();
            m_current_column++;
            return Result;
        }

        friend class CWorkbook;
};

template<typename T>
inline CWorksheet & CWorksheet::AddCellsTempl( const std::vector<T> & data )
{
    for( typename std::vector<T>::const_iterator it = data.begin(); it != data.end(); it++ )
        AddCell( it->value, it->style_id );
    return * this;
}

// ****************************************************************************
/// @brief  Appends another row into the sheet
/// @param  data reference to the vector of  <T>
/// @param  offset the offset from the row begining (0 by default)
/// @param	height row height (default if 0)
/// @return Reference to this object
// ****************************************************************************
template<typename T>
CWorksheet & CWorksheet::AddRowTempl( const std::vector<T> & data, uint32_t offset, double height )
{
    m_offset_column = offset;
    AddRowHeader( data.size(), height );
    m_current_column = 0;
    AddCells( data );
    AddRowFooter();

    m_offset_column = 0;
    return * this;
}

// ****************************************************************************
/// @brief	Adds a group of cells into a row
/// @param  data reference to the vector of CellDataStr
/// @return	Reference to this object
// ****************************************************************************
inline CWorksheet & CWorksheet::AddCells( const std::vector<CellDataStr> & data )
{
    for( size_t i = 0; i < data.size(); i++ )
        AddCell( data[i] );
    return * this;
}

// ****************************************************************************
/// @brief	Adds a group of cells into a row
/// @param  data reference to the vector of CellDataTime
/// @return	Reference to this object
// ****************************************************************************
inline CWorksheet & CWorksheet::AddCells( const std::vector<CellDataTime> & data )
{
    for( size_t i = 0; i < data.size(); i++ )
        AddCell( data[i] );
    return * this;
}
} // namespace SimpleXlsx

#endif	// XLSX_WORKSHEET_H
