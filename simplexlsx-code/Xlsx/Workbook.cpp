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
#include <locale>
#include <algorithm>
#include <string.h>
#include <errno.h>

#ifdef _WIN32
#include <windows.h>	// for ZIP
#include <direct.h>
#include <iostream>
#else
#include <unistd.h>
#include <sys/stat.h>
#include <sys/types.h>
#include <sys/syscall.h>
#endif  // _WIN32_WINNT

#include "Workbook.h"

#include "../Zip/zip.h"

#include "Chart.h"
#include "Drawing.h"
#include "XlsxHeaders.h"

#include "../PathManager.hpp"
#include "../XMLWriter.hpp"

namespace SimpleXlsx
{
    // ****************************************************************************
    //
    // CWorksheet class implementation
    //
    // ****************************************************************************

    // ****************************************************************************
    /// @brief  The class default constructor
    /// @return no
    // ****************************************************************************
    CWorkbook::CWorkbook( const UniString & UserName ) : m_UserName( UserName )
    {
        m_commLastId = 0;
        m_sheetId = 1;
        m_activeSheetIndex = 0;

        m_chartsheets.clear();
        m_worksheets.clear();
        m_sharedStrings.clear();

        Style style;
        style.numFormat.id = 0;
        m_styleList.Add( style ); // default style

        style.numFormat.id = 1;
        style.fill.patternType = PATTERN_GRAY_125;
        m_styleList.Add( style ); // default style

        std::stringstream TmpStringStream;
#ifdef _WIN32
        m_temp_path = getenv( "TEMP" );
        TmpStringStream << "xlsx_" << GetCurrentThreadId() << "_" << std::clock();
#else

        const int EnvVarCount = 4;
        const char * EnvVar[ EnvVarCount ] = { "TMPDIR", "TMP",  "TEMP", "TEMPDIR" };
        for( int i = 0; i < EnvVarCount; i++ )
        {
            const char * Ptr = getenv( EnvVar[ i ] );
            if( Ptr != NULL )
            {
                m_temp_path = Ptr;
                break;
            }
        }
        if( m_temp_path.empty() )
            m_temp_path = "/tmp";
        TmpStringStream << "xlsx_" << std::abs( syscall( SYS_gettid ) ) << "_" << std::clock();
#endif
        if( m_temp_path[ m_temp_path.size() - 1 ] != '/' )
            m_temp_path += "/";
        m_temp_path += TmpStringStream.str();

        //Check the UserName
        if( m_UserName.empty() )
        {
#ifdef _WIN32
            m_UserName = _wgetenv( L"USERNAME" );
#else
            m_UserName = getenv( "USERNAME" );
#endif
            if( m_UserName.empty() ) m_UserName = "Unknown";
        }

        m_pathManager = new PathManager( m_temp_path );
    }

    // ****************************************************************************
    /// @brief  The class destructor
    /// @return no
    // ****************************************************************************
    CWorkbook::~CWorkbook()
    {
        for( std::vector<CWorksheet *>::const_iterator it = m_worksheets.begin(); it != m_worksheets.end(); it++ )
            delete * it;
        for( std::vector<CChartsheet *>::const_iterator it = m_chartsheets.begin(); it != m_chartsheets.end(); it++ )
            delete * it;
        for( std::vector<CChart *>::const_iterator it = m_charts.begin(); it != m_charts.end(); it++ )
            delete * it;
        for( std::vector<CDrawing *>::const_iterator it = m_drawings.begin(); it != m_drawings.end(); it++ )
            delete * it;
        for( std::vector<CImage *>::const_iterator it = m_images.begin(); it != m_images.end(); it++ )
            delete * it;

        delete m_pathManager;
    }

    // ****************************************************************************
    /// @brief  Adds another chart into the existing worksheet
    /// @param  sheet existing worksheet
    /// @param  TopLeft top left point for the chart
    /// @param  BottomRight bottom right point for the chart
    /// @return Reference to a newly created object
    // ****************************************************************************
    CChart & CWorkbook::AddChart( CWorksheet & sheet, DrawingPoint TopLeft, DrawingPoint BottomRight, EChartTypes type )
    {
        assert( type != CHART_NONE );
        CChart * chart = new CChart( m_charts.size() + 1, type, * m_pathManager );
        m_charts.push_back( chart );
        sheet.m_Drawing.AppendChart( chart, TopLeft, BottomRight );
        return * chart;
    }

    // ****************************************************************************
    /// @brief  Adds image into the data sheet. Supported image formats: gif, jpg, png, tif.
    /// @brief  If the same image is added several times, then in XLSX will be copied only once.
    /// @param  sheet existing worksheet
    /// @param  filename path to the image file
    /// @param  TopLeft top left point for the chart
    /// @param  BottomRight bottom right point for the chart
    /// @return True if success. False if the file is unavailable or the format does not supported.
    // ****************************************************************************
    bool CWorkbook::AddImage( CWorksheet & sheet, const std::string & filename, DrawingPoint TopLeft, DrawingPoint BottomRight )
    {
        CImage * image = CreateImage( filename );
        if( image != NULL )
        {
            sheet.m_Drawing.AppendImage( image, TopLeft, BottomRight );
        }
        return image != NULL;
    }
    bool CWorkbook::AddImage( CWorksheet & sheet, const std::wstring & filename, DrawingPoint TopLeft, DrawingPoint BottomRight )
    {
        return AddImage( sheet, PathManager::PathEncode( filename ), TopLeft, BottomRight );
    }

    // ****************************************************************************
    /// @brief  Adds image into the data sheet. Supported image formats (file extension): gif, jpg, jpeg, png, tif, tiff.
    /// @brief  If the same image is added several times, then in XLSX will be copied only once.
    /// @param  sheet existing worksheet
    /// @param  filename path to the image file
    /// @param  TopLeft top left point for the chart
    /// @param  XScale the scale along the X-axis, in percent
    /// @param  YScale the scale along the Y-axis, in percent
    /// @return True if success. False if the file is unavailable or the format does not supported.
    bool CWorkbook::AddImage( CWorksheet & sheet, const std::string & filename, DrawingPoint TopLeft, uint16_t XScale, uint16_t YScale )
    {
        CImage * image = CreateImage( filename );
        if( image != NULL )
        {
            sheet.m_Drawing.AppendImage( image, TopLeft, XScale, YScale );
        }
        return image != NULL;
    }
    bool CWorkbook::AddImage( CWorksheet & sheet, const std::wstring & filename, DrawingPoint TopLeft, uint16_t XScale, uint16_t YScale )
    {
        return AddImage( sheet, PathManager::PathEncode( filename ), TopLeft, XScale, YScale );
    }

    // ****************************************************************************
    /// @brief  Saves workbook into the specified file
    /// @param  name full path to the file
    /// @return Boolean result of the operation
    // ****************************************************************************
    bool CWorkbook::Save( const std::string & filename )
    {
        if( !SaveCore() || !SaveApp() || !SaveContentType() || !SaveTheme() ||
                !SaveComments() || !SaveSharedStrings() || !SaveStyles() || !SaveWorkbook() )
            return false;

        for( std::vector<CWorksheet *>::const_iterator it = m_worksheets.begin(); it != m_worksheets.end(); it++ )
            if( ( * it )->Save() == false ) return false;

        for( std::vector<CChartsheet *>::const_iterator it = m_chartsheets.begin(); it != m_chartsheets.end(); it++ )
            if( ( * it )->Save() == false ) return false;
        for( std::vector<CChart *>::const_iterator it = m_charts.begin(); it != m_charts.end(); it++ )
            if( ( * it )->Save() == false ) return false;
        for( std::vector<CDrawing *>::const_iterator it = m_drawings.begin(); it != m_drawings.end(); it++ )
            if( ( * it )->Save() == false ) return false;

        bool bRetCode = true;

        HZIP hZip = CreateZip( filename.c_str(), NULL ); // create .zip without encryption
        if( hZip != 0 )
        {
            std::vector< std::string >::const_iterator end_it = m_pathManager->ContentFiles().end();
            for( std::vector< std::string >::const_iterator it = m_pathManager->ContentFiles().begin(); it != end_it; it++ )
            {
                const std::string & File = * it;
                std::string Path = m_temp_path + File;
                ZRESULT res = ZipAdd( hZip, File.c_str() + 1, Path.c_str() );
                if( res != ZR_OK )
                {
                    bRetCode = false;
                    break;
                }
            }
            CloseZip( hZip );
        }
        else bRetCode = false;

        m_pathManager->ClearTemp();
        return bRetCode;
    }
    bool CWorkbook::Save( const std::wstring & filename )
    {
        return Save( PathManager::PathEncode( filename ) );
    }

    // ****************************************************************************
    /// @brief  Adds another data sheet into the workbook
    /// @param  title chart page title
    /// @return Reference to a newly created object
    // ****************************************************************************
    CWorksheet & CWorkbook::CreateSheet( const UniString & title )
    {
        CWorksheet * sheet = new CWorksheet( m_sheetId++, * CreateDrawing(), * m_pathManager );
        return InitWorkSheet( sheet, title );
    }

    // ****************************************************************************
    /// @brief  Adds another data sheet into the workbook
    /// @param  title chart page title
    /// @param	colWidths list of pairs colNumber:Width
    /// @return Reference to a newly created object
    // ****************************************************************************
    CWorksheet & CWorkbook::CreateSheet( const UniString & title, std::vector<ColumnWidth> & colWidths )
    {
        CWorksheet * sheet = new CWorksheet( m_sheetId++, colWidths, * CreateDrawing(), * m_pathManager );
        return InitWorkSheet( sheet, title );
    }

    // ****************************************************************************
    /// @brief  Adds another data sheet with a frozen pane into the workbook
    /// @param  title chart page title
    /// @param  frozenWidth frozen pane width (in number of cells from 0)
    /// @param  frozenHeight frozen pane height (in number of cells from 0)
    /// @return Reference to a newly created object
    // ****************************************************************************
    CWorksheet & CWorkbook::CreateSheet( const UniString & title, uint32_t frozenWidth, uint32_t frozenHeight )
    {
        CWorksheet * sheet = new CWorksheet( m_sheetId++, frozenWidth, frozenHeight, * CreateDrawing(), * m_pathManager );
        return InitWorkSheet( sheet, title );
    }

    // ****************************************************************************
    /// @brief  Adds another data sheet with a frozen pane into the workbook
    /// @param  title chart page title
    /// @param  frozenWidth frozen pane width (in number of cells from 0)
    /// @param  frozenHeight frozen pane height (in number of cells from 0)
    /// @param	colWidths list of pairs colNumber:Width
    /// @return Reference to a newly created object
    // ****************************************************************************
    CWorksheet & CWorkbook::CreateSheet( const UniString & title, uint32_t frozenWidth, uint32_t frozenHeight, std::vector<ColumnWidth> & colWidths )
    {
        CWorksheet * sheet = new CWorksheet( m_sheetId++, frozenWidth, frozenHeight, colWidths, * CreateDrawing(), * m_pathManager );
        return InitWorkSheet( sheet, title );
    }

    CWorksheet & CWorkbook::InitWorkSheet( CWorksheet * sheet, const UniString & title )
    {
        sheet->SetTitle( title );
        sheet->SetSharedStr( & m_sharedStrings );
        sheet->SetComments( & m_comments );
        m_worksheets.push_back( sheet );
        return * sheet;
    }

    CChartsheet & CWorkbook::CreateChartSheet( const UniString & title, EChartTypes type )
    {
        CChart * chart = new CChart( m_charts.size() + 1, type, * m_pathManager );
        chart->SetTitle( title );
        m_charts.push_back( chart );

        CDrawing * drawing = CreateDrawing();
        drawing->AppendChart( chart );

        CChartsheet * chartsheet = new CChartsheet( m_sheetId++, * chart, * drawing, * m_pathManager );
        m_chartsheets.push_back( chartsheet );

        return * chartsheet;
    }

    CDrawing * CWorkbook::CreateDrawing()
    {
        CDrawing * drawing = new CDrawing( m_drawings.size() + 1, * m_pathManager );
        m_drawings.push_back( drawing );
        return drawing;
    }

    //Reverse for other byte order
    static inline uint16_t ReverseUInt16( const uint16_t * Value )
    {
        const uint8_t * Ptr = reinterpret_cast< const uint8_t * >( Value );
        return ( uint16_t( Ptr[ 0 ] ) << 8 ) | Ptr[ 1 ];
    }
    static inline uint32_t ReverseUInt32( const uint32_t * Value )
    {
        const uint8_t * Ptr = reinterpret_cast< const uint8_t * >( Value );
        return ( uint32_t( Ptr[ 0 ] ) << 24 ) | ( uint32_t( Ptr[ 1 ] ) << 16 ) | ( uint32_t( Ptr[ 2 ] ) << 8 ) | Ptr[ 3 ];
    }
    static bool GetDimensionFromGIF( std::ifstream & fstream, uint16_t & Width, uint16_t & Height )
    {
        const uint8_t BufSize = 10;
        char Buf[ BufSize ];
        fstream.read( Buf, BufSize );
        Width = * reinterpret_cast< uint16_t * >( & Buf[ 6 ] );
        Height = * reinterpret_cast< uint16_t * >( & Buf[ 8 ] );
        return ( fstream.gcount() == BufSize ) && ( Buf[ 0 ] == 'G' ) && ( Buf[ 1 ] == 'I' ) && ( Buf[ 2 ] == 'F' );
    }
    static bool GetDimensionFromJPEG( std::ifstream & fstream, uint16_t & Width, uint16_t & Height )
    {
        //See: http://vip.sugovica.hu/Sardi/kepnezo/JPEG%20File%20Layout%20and%20Format.htm
        const uint8_t BufSize = 7;
        char Buf[ BufSize ];
        //Header
        fstream.read( Buf, 2 );
        if( ( fstream.gcount() != 2 ) || ( Buf[ 0 ] != char( 0xFF ) ) || ( Buf[ 1 ] != char( 0xD8 ) ) )
            return false;
        //Search SOF0 (Start Of Frame 0) marker in segments
        const uint8_t SegmentHeaderSize = 2 * sizeof( uint8_t ) + sizeof( uint16_t );
        while( ! fstream.eof() )
        {
            fstream.read( Buf, SegmentHeaderSize );
            if( ( fstream.gcount() != SegmentHeaderSize ) || ( Buf[ 0 ] != char( 0xFF ) ) ) //EOF or invalid marker
                return false;
            if( Buf[ 1 ] == char( 0xC0 ) )
            {
                const uint8_t ElapsedSegmentSize = sizeof( uint8_t ) + 2 * sizeof( uint16_t );
                fstream.read( Buf, ElapsedSegmentSize );
                Width = ReverseUInt16( reinterpret_cast< uint16_t * >( & Buf[ 3 ] ) );
                Height = ReverseUInt16( reinterpret_cast< uint16_t * >( & Buf[ 1 ] ) );
                return fstream.gcount() == ElapsedSegmentSize;
            }
            else    //To next segment
            {
                uint16_t SegmentSize = ReverseUInt16( reinterpret_cast< uint16_t * >( & Buf[ 2 ] ) );
                assert( SegmentSize > sizeof( SegmentSize ) );
                fstream.seekg( SegmentSize - sizeof( SegmentSize ), fstream.cur );  //SegmentSize includes self size
            }
        }
        return false;
    }
    static bool GetDimensionFromPNG( std::ifstream & fstream, uint16_t & Width, uint16_t & Height )
    {
        const uint8_t ValidHeaderSize = 8 + 8;
        static const char ValidHeader[ ValidHeaderSize ] = { char( 0x89 ), 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A,
                                                             0x00, 0x00, 0x00, 0x0D, 'I', 'H', 'D', 'R'
                                                           };
        const uint8_t BufSize = 8 + 8 + 8;  // Header + IHDR section name + dimension
        char Buf[ BufSize ];
        fstream.read( Buf, BufSize );
        if( ( fstream.gcount() != BufSize ) || ( strncmp( Buf, ValidHeader, ValidHeaderSize ) != 0 ) )
            return false;
        const uint32_t TmpWidth = ReverseUInt32( reinterpret_cast< uint32_t * >( & Buf[ 16 ] ) );
        const uint32_t TmpHeight = ReverseUInt32( reinterpret_cast< uint32_t * >( & Buf[ 20 ] ) );
        Width = uint16_t( TmpWidth );
        Height = uint16_t( TmpHeight );
        return ( TmpWidth <= 0xFFFF ) && ( TmpHeight <= 0xFFFF );
    }
    static bool GetDimensionFromTIFF( std::ifstream & fstream, uint16_t & Width, uint16_t & Height )
    {
        //See: http://www.fileformat.info/format/tiff/corion.htm
        const uint8_t BufSize = 16;
        char Buf[ BufSize ];
        fstream.read( Buf, 8 );
        if( fstream.gcount() != 8 )
            return false;
        bool NeedReverse = false;
        if( ( Buf[ 0 ] == 'I' ) && ( Buf[ 1 ] == 'I' ) && ( Buf[ 2 ] == 42 )  && ( Buf[ 3 ] == 0 ) )        //Intel byte order
            NeedReverse = false;
        else if( ( Buf[ 0 ] == 'M' ) && ( Buf[ 1 ] == 'M' ) && ( Buf[ 2 ] == 0 )  && ( Buf[ 3 ] == 42 ) )   //Motorola byte order
            NeedReverse = true;
        else return false;  //Wrong format
        //Process file - search dimension
        uint32_t DirOffset = * reinterpret_cast< uint32_t * >( & Buf[ 4 ] );
        while( true )
        {
            fstream.seekg( NeedReverse ? ReverseUInt32( & DirOffset ) : DirOffset );    //Seek to image file directory
            uint16_t EntryNum;  //Number of entries
            fstream.read( reinterpret_cast< char * >( & EntryNum ), sizeof( EntryNum ) );
            if( fstream.gcount() != sizeof( EntryNum ) )
                return false;
            if( NeedReverse )
                EntryNum = ReverseUInt16( & EntryNum );
            int Counter = 0;
            for( int i = 0; i < EntryNum; i++ )
            {
                const uint8_t FieldSize = 2 * sizeof( uint16_t ) + 2 * sizeof( uint32_t );
                fstream.read( Buf, FieldSize );
                if( fstream.gcount() != FieldSize )
                    return false;
                uint16_t Tag = * reinterpret_cast< uint16_t * >( & Buf[ 0 ] );      //Field tag
                uint32_t Offset = * reinterpret_cast< uint32_t * >( & Buf[ 8 ] );   //Data offset of the field
                if( NeedReverse )
                {
                    Tag = ReverseUInt16( & Tag );
                    Offset = ReverseUInt32( &Offset );
                }
                if( Tag == 0x0100 )         //ImageWidth
                {
                    assert( Offset <= 0xFFFF );
                    if( Offset <= 0xFFFF )
                        Width = uint16_t( Offset );
                    if( ++Counter == 2 )        //Width and height found
                        return true;
                }
                else if( Tag == 0x0101 )    //ImageLength
                {
                    assert( Offset <= 0xFFFF );
                    if( Offset <= 0xFFFF )
                        Height = uint16_t( Offset );
                    if( ++Counter == 2 )        //Width and height found
                        return true;
                }
            }
            //Possible exist next image file directory.
            //Offset of next IFD in file, 0 if none follow.
            fstream.read( reinterpret_cast< char * >( & DirOffset ), sizeof( DirOffset ) );
            if( NeedReverse )
                DirOffset = ReverseUInt32( & DirOffset );
            if( ( fstream.gcount() != sizeof( DirOffset ) ) || ( DirOffset == 0 ) )
                return false;
        }
        return false;
    }
    struct TSupportedImages // Supported image files
    {
        std::string            FileExt;
        CImage::ImageType   ImageType;
        bool ( * FuncPtr )( std::ifstream &, uint16_t &, uint16_t & );
    };
    static const TSupportedImages KnownExt[] =
    {
        { ".gif",     CImage::gif,      & GetDimensionFromGIF   },
        { ".jpg",     CImage::jpg,      & GetDimensionFromJPEG  },
        { ".jpeg",    CImage::jpeg,     & GetDimensionFromJPEG  },
        { ".png",     CImage::png,      & GetDimensionFromPNG   },
        { ".tif",     CImage::tif,      & GetDimensionFromTIFF  },
        { ".tiff",    CImage::tiff,     & GetDimensionFromTIFF  },
        { "",         CImage::unknown,  NULL }
    };

    CImage * CWorkbook::CreateImage( const std::string & filename )
    {
        CImage * image = NULL;
        for( std::vector< CImage * >::const_iterator it = m_images.begin(); it != m_images.end(); it++ )
            if( ( * it )->LocalPath == filename )
            {
                image = * it;
                break;
            }
        if( image == NULL ) //New image
        {
            size_t LastPoint = filename.find_last_of( '.' );
            if( LastPoint == filename.npos )    //No extension
                return NULL;
            std::string Ext = filename.substr( LastPoint );
            std::transform( Ext.begin(), Ext.end(), Ext.begin(), ::tolower );
            std::ifstream imfile( filename.c_str(), std::ios::binary );
            if( ! imfile.is_open() )            //File not exist or busy
                return NULL;
            //Check for correct extension and format, extract image dimension
            uint16_t ImWidth = 0, ImHeight = 0;
            const TSupportedImages * Ptr = KnownExt;
            while( ( Ptr->ImageType != CImage::unknown ) && ( Ptr->FileExt != Ext ) )
                Ptr++;
            if( ( Ptr->ImageType == CImage::unknown ) || ! Ptr->FuncPtr( imfile, ImWidth, ImHeight )
                    || ( ImWidth == 0 ) || ( ImHeight == 0 ) )
                return NULL;

            std::stringstream IntFileName;
            IntFileName << "image" << m_images.size() + 1 << Ext;
            image = new CImage( filename, IntFileName.str(), Ptr->ImageType, ImWidth, ImHeight );
            if( ! m_pathManager->RegisterImage( filename, "/xl/media/" + image->InternalName ) )
            {
                delete image;
                return NULL;
            }
            else m_images.push_back( image );   //File successfully copied to temporary folder
        }
        return image;
    }

    // ****************************************************************************
    /// @brief  ...
    /// @return Boolean result of the operation
    // ****************************************************************************
    bool CWorkbook::SaveCore()
    {
        {
            std::time_t t = std::time( NULL );
            const size_t MAX_USER_TIME_LENGTH   =   32;
            char UserTime[ MAX_USER_TIME_LENGTH ] = { 0 };
            std::strftime( UserTime, MAX_USER_TIME_LENGTH, "%Y-%m-%dT%H:%M:%SZ", std::localtime( & t ) ) ;

            // [- zip/docProps/core.xml
            XMLWriter xmlw( m_pathManager->RegisterXML( "/docProps/core.xml" ) );
            xmlw.Tag( "cp:coreProperties" ).Attr( "xmlns:cp", ns_cp ).Attr( "xmlns:dc", ns_dc );
            xmlw.Attr( "xmlns:dcterms", ns_dcterms ).Attr( "xmlns:dcmitype", ns_dcmitype ).Attr( "xmlns:xsi", ns_xsi );
            xmlw.TagOnlyContent( "dc:creator", m_UserName );
            xmlw.TagOnlyContent( "cp:lastModifiedBy", m_UserName );
            xmlw.Tag( "dcterms:created" ).Attr( "xsi:type", xsi_type ).Cont( UserTime ).End();
            xmlw.Tag( "dcterms:modified" ).Attr( "xsi:type", xsi_type ).Cont( UserTime ).End();
            xmlw.End( "cp:coreProperties" );
            // zip/docProps/core.xml -]
        }
        {
            // [- zip/_rels/.rels
            XMLWriter xmlw( m_pathManager->RegisterXML( "/_rels/.rels" ) );
            xmlw.Tag( "Relationships" ).Attr( "xmlns", ns_relationships );
            const char * Node = "Relationship";
            xmlw.TagL( Node ).Attr( "Id", "rId3" ).Attr( "Type", type_app ).Attr( "Target", "docProps/app.xml" ).EndL();
            xmlw.TagL( Node ).Attr( "Id", "rId2" ).Attr( "Type", type_core ).Attr( "Target", "docProps/core.xml" ).EndL();
            xmlw.TagL( Node ).Attr( "Id", "rId1" ).Attr( "Type", type_book ).Attr( "Target", "xl/workbook.xml" ).EndL();
            xmlw.End( "Relationships" );
            // zip/_rels/.rels -]
        }
        return true;
    }

    // ****************************************************************************
    /// @brief  ...
    /// @return Boolean result of the operation
    // ****************************************************************************
    bool CWorkbook::SaveContentType()
    {
        // [- zip/[Content_Types].xml
        XMLWriter xmlw( m_pathManager->RegisterXML( "/[Content_Types].xml" ) );
        xmlw.Tag( "Types" ).Attr( "xmlns", ns_content_types );
        xmlw.TagL( "Default" ).Attr( "Extension", "rels" ).Attr( "ContentType", content_rels ).EndL();
        xmlw.TagL( "Default" ).Attr( "Extension", "xml" ).Attr( "ContentType", content_xml ).EndL();
        AddImagesExtensions( xmlw );
        xmlw.TagL( "Override" ).Attr( "PartName", "/xl/workbook.xml" ).Attr( "ContentType", content_book ).EndL();

        bool bFormula = false;
        for( std::vector<CWorksheet *>::const_iterator it = m_worksheets.begin(); it != m_worksheets.end(); it++ )
        {
            std::stringstream PropValue;
            PropValue << "/xl/worksheets/sheet" << ( * it )->GetIndex() << ".xml";
            xmlw.TagL( "Override" ).Attr( "PartName", PropValue.str() ).Attr( "ContentType", content_sheet ).EndL();
            if( ( * it )->IsThereFormula() ) bFormula = true;
        }
        if( bFormula )
        {
            xmlw.TagL( "Override" ).Attr( "PartName", "/xl/calcChain.xml" ).Attr( "ContentType", content_chain ).EndL();
            if( ! SaveChain() ) return false;
        }

        for( std::vector<CChartsheet *>::const_iterator it = m_chartsheets.begin(); it != m_chartsheets.end(); it++ )
        {
            std::stringstream PropValue;
            PropValue << "/xl/chartsheets/sheet" << ( * it )->GetIndex() << ".xml";
            xmlw.TagL( "Override" ).Attr( "PartName", PropValue.str() ).Attr( "ContentType", content_chartsheet ).EndL();
        }

        xmlw.TagL( "Override" ).Attr( "PartName", "/xl/theme/theme1.xml" ).Attr( "ContentType", content_theme ).EndL();
        xmlw.TagL( "Override" ).Attr( "PartName", "/xl/styles.xml" ).Attr( "ContentType", content_styles ).EndL();

        if( ! m_sharedStrings.empty() )
            xmlw.TagL( "Override" ).Attr( "PartName", "/xl/sharedStrings.xml" ).Attr( "ContentType", content_sharedStr ).EndL();

        for( std::vector<CDrawing *>::const_iterator it = m_drawings.begin(); it != m_drawings.end(); it++ )
            if( ( * it )->IsEmpty() == false )
            {
                std::stringstream PropValue;
                PropValue << "/xl/drawings/drawing" << ( * it )->GetIndex() << ".xml";
                xmlw.TagL( "Override" ).Attr( "PartName", PropValue.str() ).Attr( "ContentType", content_drawing ).EndL();
            }
        for( std::vector<CChart *>::const_iterator it = m_charts.begin(); it != m_charts.end(); it++ )
        {
            std::stringstream PropValue;
            PropValue << "/xl/charts/chart" << ( * it )->GetIndex() << ".xml";
            xmlw.TagL( "Override" ).Attr( "PartName", PropValue.str() ).Attr( "ContentType", content_chart ).EndL();
        }

        if( ! m_comments.empty() )
        {
            xmlw.TagL( "Default" ).Attr( "Extension", "vml" ).Attr( "ContentType", content_vml ).EndL();

            for( std::vector<CWorksheet *>::const_iterator it = m_worksheets.begin(); it != m_worksheets.end(); it++ )
            {
                if( ( * it )->IsThereComment() )
                {
                    std::stringstream Temp;
                    Temp << "/xl/comments" << ( * it )->GetIndex() << ".xml";
                    xmlw.TagL( "Override" ).Attr( "PartName", Temp.str() ).Attr( "ContentType", content_comment ).EndL();
                }
            }
        }

        xmlw.TagL( "Override" ).Attr( "PartName", "/docProps/core.xml" ).Attr( "ContentType", content_core ).EndL();
        xmlw.TagL( "Override" ).Attr( "PartName", "/docProps/app.xml" ).Attr( "ContentType", content_app ).EndL();

        xmlw.End( "Types" );
        // zip/[ContentTypes].xml -]
        return true;
    }

    // ****************************************************************************
    /// @brief  ...
    /// @return Boolean result of the operation
    // ****************************************************************************
    bool CWorkbook::SaveApp()
    {
        // [- zip/docProps/app.xml
        size_t nSheets = m_worksheets.size();
        size_t nCharts = m_chartsheets.size();
        size_t nVectorSize = ( nCharts > 0 ) ? 4 : 2;

        XMLWriter xmlw( m_pathManager->RegisterXML( "/docProps/app.xml" ) );
        xmlw.Tag( "Properties" ).Attr( "xmlns", ns_doc_prop ).Attr( "xmlns:vt", ns_vt );
        xmlw.TagOnlyContent( "Application", "Microsoft Excel" );
        xmlw.TagOnlyContent( "DocSecurity", 0 );
        xmlw.TagOnlyContent( "ScaleCrop", "false" );

        xmlw.Tag( "HeadingPairs" ).Tag( "vt:vector" ).Attr( "size", nVectorSize ).Attr( "baseType", "variant" );
        xmlw.Tag( "vt:variant" ).TagOnlyContent( "vt:lpstr", "Worksheets" ).End( "vt:variant" );
        xmlw.Tag( "vt:variant" ).TagOnlyContent( "vt:i4", nSheets ).End( "vt:variant" );
        if( nCharts > 0 )
        {
            xmlw.Tag( "vt:variant" ).TagOnlyContent( "vt:lpstr", "Diagramms" ).End( "vt:variant" );
            xmlw.Tag( "vt:variant" ).TagOnlyContent( "vt:i4", nCharts ).End( "vt:variant" );
        }
        xmlw.End( "vt:vector" ).End( "HeadingPairs" );

        xmlw.Tag( "TitlesOfParts" ).Tag( "vt:vector" ).Attr( "size", nSheets + nCharts ).Attr( "baseType", "lpstr" );
        for( size_t i = 0; i < nSheets; i++ )
            xmlw.TagOnlyContent( "vt:lpstr", m_worksheets[ i ]->GetTitle() );
        for( size_t i = 0; i < nCharts; i++ )
            //xmlw.TagOnlyContent( "vt:lpstr", m_chartsheets[ i ]->Chart().GetTitle() );
            xmlw.TagOnlyContent( "vt:lpstr", m_chartsheets[ i ]->Chart().GetTitle() );
        xmlw.End( "vt:vector" ).End( "TitlesOfParts" );

        xmlw.TagOnlyContent( "Company", "" );
        xmlw.TagOnlyContent( "LinksUpToDate", "false" );
        xmlw.TagOnlyContent( "SharedDoc", "false" );
        xmlw.TagOnlyContent( "HyperlinksChanged", "false" );
        xmlw.TagOnlyContent( "AppVersion", 14.03 );         //Microsoft Excel 2010
        xmlw.End( "Properties" );
        // zip/docProps/app.xml -]
        return true;
    }

    // ****************************************************************************
    /// @brief  ...
    /// @return Boolean result of the operation
    // ****************************************************************************
    bool CWorkbook::SaveTheme()
    {
        // [- zip/xl/theme/theme1.xml
        XMLWriter xmlw( m_pathManager->RegisterXML( "/xl/theme/theme1.xml" ) );
        xmlw.Tag( "a:theme" ).Attr( "xmlns:a", ns_a ).Attr( "name", "Office Theme" );
        xmlw.Tag( "a:themeElements" );

        xmlw.Tag( "a:clrScheme" ).Attr( "name", "Office" );
        xmlw.Tag( "a:dk1" ).TagL( "a:sysClr" ).Attr( "val", "windowText" ).Attr( "lastClr", "000000" ).EndL().End();
        xmlw.Tag( "a:lt1" ).TagL( "a:sysClr" ).Attr( "val", "window" ).Attr( "lastClr", "FFFFFF" ).EndL().End();
        xmlw.Tag( "a:dk2" ).TagL( "a:srgbClr" ).Attr( "val", "1F497D" ).EndL().End();
        xmlw.Tag( "a:lt2" ).TagL( "a:srgbClr" ).Attr( "val", "EEECE1" ).EndL().End();
        xmlw.Tag( "a:accent1" ).TagL( "a:srgbClr" ).Attr( "val", "4F81BD" ).EndL().End();
        xmlw.Tag( "a:accent2" ).TagL( "a:srgbClr" ).Attr( "val", "C0504D" ).EndL().End();
        xmlw.Tag( "a:accent3" ).TagL( "a:srgbClr" ).Attr( "val", "9BBB59" ).EndL().End();
        xmlw.Tag( "a:accent4" ).TagL( "a:srgbClr" ).Attr( "val", "8064A2" ).EndL().End();
        xmlw.Tag( "a:accent5" ).TagL( "a:srgbClr" ).Attr( "val", "4BACC6" ).EndL().End();
        xmlw.Tag( "a:accent6" ).TagL( "a:srgbClr" ).Attr( "val", "F79646" ).EndL().End();
        xmlw.Tag( "a:hlink" ).TagL( "a:srgbClr" ).Attr( "val", "0000FF" ).EndL().End();
        xmlw.Tag( "a:folHlink" ).TagL( "a:srgbClr" ).Attr( "val", "800080" ).EndL().End();
        xmlw.End( "a:clrScheme" );

        const char * szAFont = "a:font";
        const char * szScript = "script";
        const char * szTypeface = "typeface";
        xmlw.Tag( "a:fontScheme" ).Attr( "name", "Office" );
        for( int i = 0; i < 2; i++ )
        {
            if( i == 0 ) xmlw.Tag( "a:majorFont" );
            else xmlw.Tag( "a:minorFont" );

            xmlw.TagL( "a:latin" ).Attr( szTypeface, "Cambria" ).EndL();
            xmlw.TagL( "a:ea" ).Attr( szTypeface, "" ).EndL();
            xmlw.TagL( "a:cs" ).Attr( szTypeface, "" ).EndL();

            xmlw.TagL( szAFont ).Attr( szScript, "Arab" ).Attr( szTypeface, "Times New Roman" ).EndL();
            xmlw.TagL( szAFont ).Attr( szScript, "Herb" ).Attr( szTypeface, "Times New Roman" ).EndL();
            xmlw.TagL( szAFont ).Attr( szScript, "Thai" ).Attr( szTypeface, "Tahoma" ).EndL();
            xmlw.TagL( szAFont ).Attr( szScript, "Ethi" ).Attr( szTypeface, "Nyala" ).EndL();
            xmlw.TagL( szAFont ).Attr( szScript, "Beng" ).Attr( szTypeface, "Vrinda" ).EndL();
            xmlw.TagL( szAFont ).Attr( szScript, "Gujr" ).Attr( szTypeface, "Shruti" ).EndL();
            xmlw.TagL( szAFont ).Attr( szScript, "Khmr" ).Attr( szTypeface, "MoolBoran" ).EndL();
            xmlw.TagL( szAFont ).Attr( szScript, "Knda" ).Attr( szTypeface, "Tunga" ).EndL();
            xmlw.TagL( szAFont ).Attr( szScript, "Guru" ).Attr( szTypeface, "Raavi" ).EndL();
            xmlw.TagL( szAFont ).Attr( szScript, "Cans" ).Attr( szTypeface, "Euphemia" ).EndL();
            xmlw.TagL( szAFont ).Attr( szScript, "Cher" ).Attr( szTypeface, "Plantagenet Cherokee" ).EndL();
            xmlw.TagL( szAFont ).Attr( szScript, "Yiii" ).Attr( szTypeface, "Microsoft Yi Baiti" ).EndL();
            xmlw.TagL( szAFont ).Attr( szScript, "Tibt" ).Attr( szTypeface, "Microsoft Himalaya" ).EndL();
            xmlw.TagL( szAFont ).Attr( szScript, "Thaa" ).Attr( szTypeface, "MV Boli" ).EndL();
            xmlw.TagL( szAFont ).Attr( szScript, "Deva" ).Attr( szTypeface, "Mangal" ).EndL();
            xmlw.TagL( szAFont ).Attr( szScript, "Telu" ).Attr( szTypeface, "Gautami" ).EndL();
            xmlw.TagL( szAFont ).Attr( szScript, "Taml" ).Attr( szTypeface, "Latha" ).EndL();
            xmlw.TagL( szAFont ).Attr( szScript, "Syrc" ).Attr( szTypeface, "Estrangelo Edessa" ).EndL();
            xmlw.TagL( szAFont ).Attr( szScript, "Orya" ).Attr( szTypeface, "Kalinga" ).EndL();
            xmlw.TagL( szAFont ).Attr( szScript, "Mlym" ).Attr( szTypeface, "Kartika" ).EndL();
            xmlw.TagL( szAFont ).Attr( szScript, "Laoo" ).Attr( szTypeface, "DokChampa" ).EndL();
            xmlw.TagL( szAFont ).Attr( szScript, "Sinh" ).Attr( szTypeface, "Iskoola Pota" ).EndL();
            xmlw.TagL( szAFont ).Attr( szScript, "Mong" ).Attr( szTypeface, "Mongolian Baiti" ).EndL();
            xmlw.TagL( szAFont ).Attr( szScript, "Viet" ).Attr( szTypeface, "Times New Roman" ).EndL();
            xmlw.TagL( szAFont ).Attr( szScript, "Uigh" ).Attr( szTypeface, "Microsoft Uighur" ).EndL();
            xmlw.TagL( szAFont ).Attr( szScript, "Geor" ).Attr( szTypeface, "Sylfaen" ).EndL();

            xmlw.End(); //For "a:majorFont" and "a:minorFont"
        }
        xmlw.End( "a:fontScheme" );

        xmlw.Tag( "a:fmtScheme" ).Attr( "name", "Office" );

        xmlw.Tag( "a:fillStyleLst" );
        xmlw.Tag( "a:solidFill" ).TagL( "a:schemeClr" ).Attr( "val", "phClr" ).EndL().End();
        xmlw.Tag( "a:gradFill" ).Attr( "rotWithShape", 1 );
        xmlw.Tag( "a:gsLst" );
        /**/xmlw.Tag( "a:gs" ).Attr( "pos", 0 ).Tag( "a:schemeClr" ).Attr( "val", "phClr" );
        /**/xmlw.TagL( "a:tint" ).Attr( "val", 50000 ).EndL().TagL( "a:satMod" ).Attr( "val", 300000 ).EndL().End().End();
        /**/xmlw.Tag( "a:gs" ).Attr( "pos", 35000 ).Tag( "a:schemeClr" ).Attr( "val", "phClr" );
        /**/xmlw.TagL( "a:tint" ).Attr( "val", 37000 ).EndL().TagL( "a:satMod" ).Attr( "val", 300000 ).EndL().End().End();
        /**/xmlw.Tag( "a:gs" ).Attr( "pos", 100000 ).Tag( "a:schemeClr" ).Attr( "val", "phClr" );
        /**/xmlw.TagL( "a:tint" ).Attr( "val", 15000 ).EndL().TagL( "a:satMod" ).Attr( "val", 350000 ).EndL().End().End();
        xmlw.End( "a:gsLst" );
        xmlw.TagL( "a:lin" ).Attr( "ang", 16200000 ).Attr( "scaled", 1 ).EndL();
        xmlw.End( "a:gradFill" );
        xmlw.Tag( "a:gradFill" ).Attr( "rotWithShape", 1 );
        xmlw.Tag( "a:gsLst" );
        /**/xmlw.Tag( "a:gs" ).Attr( "pos", 0 ).Tag( "a:schemeClr" ).Attr( "val", "phClr" );
        /**/xmlw.TagL( "a:shade" ).Attr( "val", 51000 ).EndL().TagL( "a:satMod" ).Attr( "val", 130000 ).EndL().End().End();
        /**/xmlw.Tag( "a:gs" ).Attr( "pos", 80000 ).Tag( "a:schemeClr" ).Attr( "val", "phClr" );
        /**/xmlw.TagL( "a:shade" ).Attr( "val", 93000 ).EndL().TagL( "a:satMod" ).Attr( "val", 130000 ).EndL().End().End();
        /**/xmlw.Tag( "a:gs" ).Attr( "pos", 100000 ).Tag( "a:schemeClr" ).Attr( "val", "phClr" );
        /**/xmlw.TagL( "a:shade" ).Attr( "val", 94000 ).EndL().TagL( "a:satMod" ).Attr( "val", 135000 ).EndL().End().End();
        xmlw.End( "a:gsLst" );
        xmlw.TagL( "a:lin" ).Attr( "ang", 16200000 ).Attr( "scaled", 0 ).EndL();
        xmlw.End( "a:gradFill" );
        xmlw.End( "a:fillStyleLst" );

        xmlw.Tag( "a:lnStyleLst" );
        xmlw.Tag( "a:ln" ).Attr( "w", 9525 ).Attr( "cap", "flat" ).Attr( "cmpd", "sng" ).Attr( "algn", "ctr" );
        /**/xmlw.Tag( "a:solidFill" ).Tag( "a:schemeClr" ).Attr( "val", "phClr" );
        /**/xmlw.TagL( "a:shade" ).Attr( "val", 95000 ).EndL().TagL( "a:satMod" ).Attr( "val", 105000 ).EndL().End().End();
        /**/xmlw.TagL( "a:prstDash" ).Attr( "val", "solid" ).EndL();
        xmlw.End( "a:ln" );
        xmlw.Tag( "a:ln" ).Attr( "w", 25400 ).Attr( "cap", "flat" ).Attr( "cmpd", "sng" ).Attr( "algn", "ctr" );
        /**/xmlw.Tag( "a:solidFill" ).TagL( "a:schemeClr" ).Attr( "val", "phClr" ).EndL().End();
        /**/xmlw.TagL( "a:prstDash" ).Attr( "val", "solid" ).EndL();
        xmlw.End( "a:ln" );
        xmlw.Tag( "a:ln" ).Attr( "w", 38100 ).Attr( "cap", "flat" ).Attr( "cmpd", "sng" ).Attr( "algn", "ctr" );
        /**/xmlw.Tag( "a:solidFill" ).TagL( "a:schemeClr" ).Attr( "val", "phClr" ).EndL().End();
        /**/xmlw.TagL( "a:prstDash" ).Attr( "val", "solid" ).EndL();
        xmlw.End( "a:ln" );
        xmlw.End( "a:lnStyleLst" );

        xmlw.Tag( "a:effectStyleLst" );
        xmlw.Tag( "a:effectStyle" ).Tag( "a:effectLst" );
        /**/xmlw.Tag( "a:outerShdw" ).Attr( "blurRad", 40000 ).Attr( "dist", 20000 ).Attr( "dir", 5400000 ).Attr( "rotWithShape", 0 );
        /**/xmlw.Tag( "a:srgbClr" ).Attr( "val", "000000" ).TagL( "a:alpha" ).Attr( "val", 38000 ).EndL().End().End();
        xmlw.End( "a:effectLst" ).End( "a:effectStyle" );
        xmlw.Tag( "a:effectStyle" ).Tag( "a:effectLst" );
        /**/xmlw.Tag( "a:outerShdw" ).Attr( "blurRad", 40000 ).Attr( "dist", 23000 ).Attr( "dir", 5400000 ).Attr( "rotWithShape", 0 );
        /**/xmlw.Tag( "a:srgbClr" ).Attr( "val", "000000" ).TagL( "a:alpha" ).Attr( "val", 35000 ).EndL().End().End();
        xmlw.End( "a:effectLst" ).End( "a:effectStyle" );
        xmlw.Tag( "a:effectStyle" ).Tag( "a:effectLst" );
        /**/xmlw.Tag( "a:outerShdw" ).Attr( "blurRad", 40000 ).Attr( "dist", 23000 ).Attr( "dir", 5400000 ).Attr( "rotWithShape", 0 );
        /**/xmlw.Tag( "a:srgbClr" ).Attr( "val", "000000" ).TagL( "a:alpha" ).Attr( "val", 35000 ).EndL().End().End();
        xmlw.End( "a:effectLst" );
        xmlw.Tag( "a:scene3d" );
        /**/xmlw.Tag( "a:camera" ).Attr( "prst", "orthographicFront" );
        /**/xmlw.TagL( "a:rot" ).Attr( "lat", 0 ).Attr( "lon", 0 ).Attr( "rev", 0 ).EndL().End();
        /**/xmlw.Tag( "a:lightRig" ).Attr( "rig", "threePt" ).Attr( "dir", "t" );
        /**/xmlw.TagL( "a:rot" ).Attr( "lat", 0 ).Attr( "lon", 0 ).Attr( "rev", 1200000 ).EndL().End();
        xmlw.End( "a:scene3d" );
        xmlw.Tag( "a:sp3d" ).Tag( "a:bevelT" ).Attr( "w", 63500 ).Attr( "h", 25400 ).End().End();
        xmlw.End( "a:effectStyle" );
        xmlw.End( "a:effectStyleLst" );

        xmlw.Tag( "a:bgFillStyleLst" );
        xmlw.Tag( "a:solidFill" ).TagL( "a:schemeClr" ).Attr( "val", "phClr" ).EndL().End();
        xmlw.Tag( "a:gradFill" ).Attr( "rotWithShape", 1 ).Tag( "a:gsLst" );
        /**/xmlw.Tag( "a:gs" ).Attr( "pos", 0 ).Tag( "a:schemeClr" ).Attr( "val", "phClr" );
        /**/xmlw.TagL( "a:tint" ).Attr( "val", 40000 ).EndL().TagL( "a:satMod" ).Attr( "val", 350000 ).EndL().End().End();
        /**/xmlw.Tag( "a:gs" ).Attr( "pos", 40000 ).Tag( "a:schemeClr" ).Attr( "val", "phClr" );
        /**/xmlw.TagL( "a:tint" ).Attr( "val", 45000 ).EndL().TagL( "a:shade" ).Attr( "val", 99000 ).EndL();
        /**/xmlw.TagL( "a:satMod" ).Attr( "val", 350000 ).EndL().End().End();
        /**/xmlw.Tag( "a:gs" ).Attr( "pos", 100000 ).Tag( "a:schemeClr" ).Attr( "val", "phClr" );
        /**/xmlw.TagL( "a:shade" ).Attr( "val", 20000 ).EndL().TagL( "a:satMod" ).Attr( "val", 255000 ).EndL().End().End();
        xmlw.End( "a:gsLst" );
        xmlw.Tag( "a:path" ).Attr( "path", "circle" );
        /**/xmlw.TagL( "a:fillToRect" ).Attr( "l", 50000 ).Attr( "t", -80000 ).Attr( "r", 50000 ).Attr( "b", 180000 ).EndL();
        xmlw.End( "a:path" ).End( "a:gradFill" );
        xmlw.Tag( "a:gradFill" ).Attr( "rotWithShape", 1 ).Tag( "a:gsLst" );
        /**/xmlw.Tag( "a:gs" ).Attr( "pos", 0 ).Tag( "a:schemeClr" ).Attr( "val", "phClr" );
        /**/xmlw.TagL( "a:tint" ).Attr( "val", 80000 ).EndL().TagL( "a:satMod" ).Attr( "val", 300000 ).EndL().End().End();
        /**/xmlw.Tag( "a:gs" ).Attr( "pos", 100000 ).Tag( "a:schemeClr" ).Attr( "val", "phClr" );
        /**/xmlw.TagL( "a:shade" ).Attr( "val", 30000 ).EndL().TagL( "a:satMod" ).Attr( "val", 200000 ).EndL().End().End();
        xmlw.End( "a:gsLst" );
        xmlw.Tag( "a:path" ).Attr( "path", "circle" );
        /**/xmlw.TagL( "a:fillToRect" ).Attr( "l", 50000 ).Attr( "t", 50000 ).Attr( "r", 50000 ).Attr( "b", 50000 ).EndL();
        xmlw.End( "a:path" ).End( "a:gradFill" );

        xmlw.End( "a:bgFillStyleLst" );

        xmlw.End( "a:fmtScheme" );

        xmlw.End( "a:themeElements" );
        xmlw.TagL( "a:objectDefaults" ).EndL();
        xmlw.TagL( "a:extraClrSchemeLst" ).EndL();
        xmlw.End( "a:theme" );
        // zip/xl/theme/theme1.xml -]
        return true;
    }

    // ****************************************************************************
    /// @brief  ...
    /// @return Boolean result of the operation
    // ****************************************************************************
    bool CWorkbook::SaveStyles()
    {
        // [- zip/xl/styles.xml
        XMLWriter xmlw( m_pathManager->RegisterXML( "/xl/styles.xml" ) );
        xmlw.Tag( "styleSheet" ).Attr( "xmlns", ns_book ).Attr( "xmlns:mc", ns_mc ).Attr( "mc:Ignorable", "x14ac" ).Attr( "xmlns:x14ac", ns_x14ac );

        AddNumberFormats( xmlw );
        AddFonts( xmlw );
        AddFills( xmlw );
        AddBorders( xmlw );

        xmlw.Tag( "cellStyleXfs" ).Attr( "count", 1 );
        xmlw.TagL( "xf" ).Attr( "numFmtId", 0 ).Attr( "fontId", 0 ).Attr( "fillId", 0 ).Attr( "borderId", 0 ).EndL();
        xmlw.End( "cellStyleXfs" );

        const std::vector<std::vector<size_t> > & styleIndexes = m_styleList.GetIndexes();
        xmlw.Tag( "cellXfs" ).Attr( "count", styleIndexes.size() );
        const std::vector< std::pair<std::pair<EAlignHoriz, EAlignVert>, bool> > & styleAligns = m_styleList.GetPositions();
        assert( styleIndexes.size() == styleAligns.size() );
        for( size_t i = 0; i < styleIndexes.size(); i++ )
        {
            std::vector<size_t> index = styleIndexes[ i ];
            std::pair<std::pair<EAlignHoriz, EAlignVert>, bool> align = styleAligns[ i ];

            xmlw.Tag( "xf" ).Attr( "numFmtId", index[ StyleList::STYLE_LINK_NUM_FORMAT ] );
            xmlw.Attr( "fontId", index[ StyleList::STYLE_LINK_FONT ] );
            xmlw.Attr( "fillId", index[ StyleList::STYLE_LINK_FILL ] );
            xmlw.Attr( "borderId", index[ StyleList::STYLE_LINK_BORDER ] );

            if( index[StyleList::STYLE_LINK_FONT] != 0 )        xmlw.Attr( "applyFont", 1 );
            if( index[StyleList::STYLE_LINK_FILL] != 0 )        xmlw.Attr( "applyFill", 1 );
            if( index[StyleList::STYLE_LINK_BORDER] != 0 )      xmlw.Attr( "applyBorder", 1 );
            if( index[StyleList::STYLE_LINK_NUM_FORMAT ] != 0 ) xmlw.Attr( "applyNumberFormat", 1 );

            if( ( align.first.first != ALIGN_H_NONE ) || ( align.first.second != ALIGN_V_NONE ) || align.second )
            {
                xmlw.TagL( "alignment" );
                switch( align.first.first )
                {
                    case ALIGN_H_LEFT   :	xmlw.Attr( "horizontal", "left" );      break;
                    case ALIGN_H_CENTER :   xmlw.Attr( "horizontal", "center" );    break;
                    case ALIGN_H_RIGHT  :	xmlw.Attr( "horizontal", "right" );     break;
                    case ALIGN_H_NONE   :
                        /*default:*/ 			break;
                }
                switch( align.first.second )
                {
                    case ALIGN_V_BOTTOM :	xmlw.Attr( "vertical", "bottom" );  break;
                    case ALIGN_V_CENTER :	xmlw.Attr( "vertical", "center" );  break;
                    case ALIGN_V_TOP    :	xmlw.Attr( "vertical", "top" );    break;
                    case ALIGN_V_NONE   :
                        /*default:*/ 				break;
                }
                if( align.second ) xmlw.Attr( "wrapText", 1 );
                xmlw.EndL();
            }
            xmlw.End( "xf" );
        }
        xmlw.End( "cellXfs" );

        xmlw.Tag( "cellStyles" ).Attr( "count", 1 );
        xmlw.TagL( "cellStyle" ).Attr( "name", "Normal" ).Attr( "xfId", 0 ).Attr( "builtinId", 0 ).EndL();
        xmlw.End( "cellStyles" );

        xmlw.TagL( "dxfs" ).Attr( "count", 0 ).EndL();
        xmlw.TagL( "tableStyles" ).Attr( "count", 0 ).Attr( "defaultTableStyle", "TableStyleMedium2" );
        xmlw.Attr( "defaultPivotStyle", "PivotStyleLight16" ).EndL();

        xmlw.Tag( "extLst" );
        xmlw.Tag( "ext" ).Attr( "uri", "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" ).Attr( "xmlns:x14", ns_x14 );
        xmlw.TagL( "x14:slicerStyles" ).Attr( "defaultSlicerStyle", "SlicerStyleLight1" ).EndL();
        xmlw.End( "ext" );
        xmlw.End( "extLst" );

        xmlw.End( "styleSheet" );
        // zip/xl/styles.xml -]
        return true;
    }

    // ****************************************************************************
    /// @brief  Appends number format section into styles file
    /// @param  stream reference to xml stream
    /// @return no
    // ****************************************************************************
    void CWorkbook::AddNumberFormats( XMLWriter & xmlw ) const
    {
        const std::vector<NumFormat> & nums = m_styleList.GetNumFormats();

        size_t built_in_formats = 0;
        for( std::vector<NumFormat>::const_iterator it = nums.begin(); it != nums.end(); it++ )
            if( it->id < StyleList::BUILT_IN_STYLES_NUMBER )
                built_in_formats++;

        if( nums.size() - built_in_formats == 0 ) return;

        xmlw.Tag( "numFmts" ).Attr( "count", nums.size() - built_in_formats );
        for( std::vector<NumFormat>::const_iterator it = nums.begin(); it != nums.end(); it++ )
        {
            if( it->id < StyleList::BUILT_IN_STYLES_NUMBER ) continue;
            xmlw.Tag( "numFmt" ).Attr( "numFmtId", it->id ).Attr( "formatCode", GetFormatCodeString( * it ) ).End();
        }
        xmlw.End( "numFmts" );
    }

    // ****************************************************************************
    /// @brief  Converts numeric format object into its string representation
    /// @param  fmt reference to format to be converted
    /// @return String format code
    // ****************************************************************************
    std::string CWorkbook::GetFormatCodeString( const NumFormat & fmt )
    {
        if( fmt.formatString != "" ) return fmt.formatString;

        bool addNegative = false;
        bool addZero = false;

        if( fmt.positiveColor != NUMSTYLE_COLOR_DEFAULT )
        {
            addNegative = addZero = true;
        }

        if( fmt.negativeColor != NUMSTYLE_COLOR_DEFAULT )
        {
            addNegative = true;
        }

        if( fmt.zeroColor != NUMSTYLE_COLOR_DEFAULT )
        {
            addZero = true;
        }

        std::string thousandPrefix;
        if( fmt.showThousandsSeparator ) thousandPrefix = "#,##";

        static std::string currency = CurrencySymbol();

        std::string resCode;
        std::string affix;
        std::string digits = "0";
        if( fmt.numberOfDigitsAfterPoint != 0 )
        {
            digits.append( "." );
            digits.append( fmt.numberOfDigitsAfterPoint, '0' );
        }

        switch( fmt.numberStyle )
        {
            case NUMSTYLE_EXPONENTIAL:
                affix = "E+00";
                break;
            case NUMSTYLE_PERCENTAGE:
                affix = "%";
                break;
            case NUMSTYLE_FINANCIAL:
                //_-* #,##0.00"$"_-;\-* #,##0.00"$"_-;_-* "-"??"$"_-;_-@_-

                resCode = 	GetFormatCodeColor( fmt.positiveColor ) + "_-* " + thousandPrefix + digits + currency + "_-;" +
                            GetFormatCodeColor( fmt.negativeColor ) + "\\-* " + thousandPrefix + digits + currency + "_-;" +
                            "_-* &quot;-&quot;??" + currency + "_-;_-@_-";
                break;
            case NUMSTYLE_MONEY:
                affix = " " + currency;
                break;
            case NUMSTYLE_DATETIME:
                resCode = "yyyy.mm.dd hh:mm:ss";
                break;
            case NUMSTYLE_DATE:
                resCode = "yyyy.mm.dd";
                break;
            case NUMSTYLE_TIME:
                resCode = "hh:mm:ss";
                break;

            case NUMSTYLE_GENERAL:
            case NUMSTYLE_NUMERIC:
                /*default:*/
                affix = "";
                break;
        }

        if( fmt.numberStyle == NUMSTYLE_GENERAL || fmt.numberStyle == NUMSTYLE_NUMERIC ||
                fmt.numberStyle == NUMSTYLE_EXPONENTIAL || fmt.numberStyle == NUMSTYLE_PERCENTAGE ||
                fmt.numberStyle == NUMSTYLE_MONEY )
        {

            resCode = GetFormatCodeColor( fmt.positiveColor ) + thousandPrefix + digits + affix;
            if( addNegative )
            {
                resCode += ";" + GetFormatCodeColor( fmt.negativeColor ) + "\\-" + thousandPrefix + digits + affix;
            }
            if( addZero )
            {
                if( addNegative == false ) resCode += ";";
                resCode += ";" + GetFormatCodeColor( fmt.zeroColor ) + "0" + affix;
            }
        }

        return resCode;
    }

    // ****************************************************************************
    /// @brief  Converts numeric format color into its string representation
    /// @param  color color code
    /// @return String color code
    // ****************************************************************************
    std::string CWorkbook::GetFormatCodeColor( ENumericStyleColor color )
    {
        switch( color )
        {
            case NUMSTYLE_COLOR_BLACK   :	return "[BLACK]";
            case NUMSTYLE_COLOR_GREEN   :	return "[Green]";
            case NUMSTYLE_COLOR_WHITE   :	return "[White]";
            case NUMSTYLE_COLOR_BLUE    :	return "[Blue]";
            case NUMSTYLE_COLOR_MAGENTA :   return "[Magenta]";
            case NUMSTYLE_COLOR_YELLOW  :	return "[Yellow]";
            case NUMSTYLE_COLOR_CYAN    :	return "[Cyan]";
            case NUMSTYLE_COLOR_RED     :	return "[Red]";

            case NUMSTYLE_COLOR_DEFAULT :
                /*default:*/                        return "";
        }
        return "";
    }

    std::string CWorkbook::CurrencySymbol()
    {
#ifdef _WIN32
        //Work directly with WinAPI because MinGW not working properly with std::locale...
        setlocale( LC_ALL, "" );
        struct lconv * lc = localeconv();
        if( lc == NULL ) return "";
        int WideBufSize = MultiByteToWideChar( CP_ACP, 0, lc->currency_symbol, -1, 0, 0 );
        if( WideBufSize == 0 ) return "";
        std::vector<WCHAR> WideBuf( WideBufSize );
        WideBufSize = MultiByteToWideChar( CP_ACP, 0, lc->currency_symbol, -1, WideBuf.data(), WideBufSize );
        if( WideBufSize == 0 ) return "";

        int Utf8BufSize = WideCharToMultiByte( CP_UTF8, 0, WideBuf.data(), WideBufSize, NULL, 0, NULL, NULL );
        if( Utf8BufSize == 0 ) return "";
        std::vector<char> SingleBuf( Utf8BufSize );
        Utf8BufSize = WideCharToMultiByte( CP_UTF8, 0, WideBuf.data(), WideBufSize, SingleBuf.data(), Utf8BufSize, NULL, NULL );
        if( Utf8BufSize == 0 ) return "";
        return SingleBuf.data();
#else
        std::locale loc( "" );
        std::wstring CurrencySymbol = std::use_facet<std::moneypunct<wchar_t> >( loc ).curr_symbol();
        return UTF8Encoder::From_wstring( CurrencySymbol );
#endif
    }

    // ****************************************************************************
    /// @brief  Appends fonts section into styles file
    /// @param  stream reference to xml stream
    /// @return no
    // ****************************************************************************
    void CWorkbook::AddFonts( XMLWriter & xmlw ) const
    {
        const int32_t defaultCharset = 204;
        const std::vector<Font> & fonts = m_styleList.GetFonts();
        xmlw.Tag( "fonts" ).Attr( "count", fonts.size() );
        for( std::vector<Font>::const_iterator it = fonts.begin(); it != fonts.end(); it++ )
        {
            xmlw.Tag( "font" );
            AddFontInfo( xmlw, * it, "name", defaultCharset );
            xmlw.End( "font" );
        }
        xmlw.End( "fonts" );
    }

    // ****************************************************************************
    /// @brief  Appends fills section into styles file
    /// @param  stream reference to xml stream
    /// @return no
    // ****************************************************************************
    void CWorkbook::AddFills( XMLWriter & xmlw ) const
    {
        const std::vector<Fill> & fills = m_styleList.GetFills();

        xmlw.Tag( "fills" ).Attr( "count", fills.size() );
        const char * strPattern = NULL;
        for( std::vector<Fill>::const_iterator it = fills.begin(); it != fills.end(); it++ )
        {
            xmlw.Tag( "fill" ).Tag( "patternFill" );
            switch( it->patternType )
            {
                case PATTERN_DARK_DOWN      :   strPattern = "darkDown";        break;
                case PATTERN_DARK_GRAY      :   strPattern = "darkGray";        break;
                case PATTERN_DARK_GRID      :   strPattern = "darkGrid";        break;
                case PATTERN_DARK_HORIZ     :   strPattern = "darkHorizontal";  break;
                case PATTERN_DARK_TRELLIS   :   strPattern = "darkTrellis";     break;
                case PATTERN_DARK_UP        :   strPattern = "darkUp";          break;
                case PATTERN_DARK_VERT      :   strPattern = "darkVertical";    break;
                case PATTERN_GRAY_0625      :   strPattern = "gray0625";        break;
                case PATTERN_GRAY_125       :   strPattern = "gray125";         break;
                case PATTERN_LIGHT_DOWN     :   strPattern = "lightDown";       break;
                case PATTERN_LIGHT_GRAY     :   strPattern = "lightGray";       break;
                case PATTERN_LIGHT_GRID     :   strPattern = "lightGrid";       break;
                case PATTERN_LIGHT_HORIZ    :   strPattern = "lightHorizontal"; break;
                case PATTERN_LIGHT_TRELLIS  :   strPattern = "lightTrellis";    break;
                case PATTERN_LIGHT_UP       :   strPattern = "lightUp";         break;
                case PATTERN_LIGHT_VERT     :   strPattern = "lightVertical";   break;
                case PATTERN_MEDIUM_GRAY    :   strPattern = "mediumGray";      break;
                case PATTERN_NONE           :   strPattern = "none";            break;
                case PATTERN_SOLID          :   strPattern = "solid";           break;
            }
            xmlw.Attr( "patternType", strPattern );
            if( it->bgColor != "" ) xmlw.TagL( "bgColor" ).Attr( "rgb", it->bgColor ).EndL();
            if( it->fgColor != "" ) xmlw.TagL( "fgColor" ).Attr( "rgb", it->fgColor ).EndL();
            xmlw.End( "patternFill" ).End( "fill" );
        }
        xmlw.End( "fills" );
    }

    // ****************************************************************************
    /// @brief  Appends borders section into styles file
    /// @param  stream reference to xml stream
    /// @return no
    // ****************************************************************************
    void CWorkbook::AddBorders( XMLWriter & xmlw ) const
    {
        const std::vector<Border> & borders = m_styleList.GetBorders();
        xmlw.Tag( "borders" ).Attr( "count", borders.size() );
        for( std::vector<Border>::const_iterator it = borders.begin(); it != borders.end(); it++ )
        {
            xmlw.Tag( "border" );
            if( it->isDiagonalUp ) xmlw.Attr( "diagonalUp", 1 );
            if( it->isDiagonalDown ) xmlw.Attr( "diagonalDown", 1 );
            AddBorder( xmlw, "left", it->left );
            AddBorder( xmlw, "right", it->right );
            AddBorder( xmlw, "top", it->top );
            AddBorder( xmlw, "bottom", it->bottom );
            AddBorder( xmlw, "diagonal", it->diagonal );
            xmlw.End( "border" );
        }
        xmlw.End( "borders" );
    }

    // ****************************************************************************
    /// @brief  Appends border item section into borders set
    /// @param  stream reference to xml stream
    /// @param	borderName border name
    /// @param	border style set
    /// @return no
    // ****************************************************************************
    void CWorkbook::AddBorder( XMLWriter & xmlw, const char * borderName, Border::BorderItem border ) const
    {
        const char * sStyle = NULL;
        switch( border.style )
        {
            case BORDER_THIN                :	sStyle = "thin";                break;
            case BORDER_MEDIUM              :	sStyle = "medium";              break;
            case BORDER_DASHED              :	sStyle = "dashed";              break;
            case BORDER_DOTTED              :	sStyle = "dotted";              break;
            case BORDER_THICK               :	sStyle = "thick";               break;
            case BORDER_DOUBLE              :	sStyle = "double";              break;
            case BORDER_HAIR                :	sStyle = "hair";                break;
            case BORDER_MEDIUM_DASHED       :	sStyle = "mediumDashed";        break;
            case BORDER_DASH_DOT            :	sStyle = "dashDot";             break;
            case BORDER_MEDIUM_DASH_DOT     :	sStyle = "mediumDashDot";       break;
            case BORDER_DASH_DOT_DOT        :	sStyle = "dashDotDot";          break;
            case BORDER_MEDIUM_DASH_DOT_DOT :   sStyle = "mediumDashDotDot";    break;
            case BORDER_SLANT_DASH_DOT      :   sStyle = "slantDashDot";        break;

            case BORDER_NONE:
                /*default:*/						break;
        }
        xmlw.Tag( borderName );
        if( sStyle != NULL )
        {
            xmlw.Attr( "style", sStyle );
            xmlw.Tag( "color" );
            if( border.color != "" ) xmlw.Attr( "rgb", border.color );
            else xmlw.Attr( "indexed", 64 );
            xmlw.End( "color" );
        }
        xmlw.End( borderName );
    }

    void CWorkbook::AddFontInfo( XMLWriter & xmlw, const Font & font, const char * FontTagName, int32_t Charset ) const
    {
        int32_t attributes = font.attributes;
        if( attributes & FONT_BOLD )        xmlw.TagL( "b" ).EndL();
        if( attributes & FONT_ITALIC )      xmlw.TagL( "i" ).EndL();
        if( attributes & FONT_UNDERLINED )  xmlw.TagL( "u" ).EndL();
        if( attributes & FONT_STRIKE )      xmlw.TagL( "strike" ).EndL();
        if( attributes & FONT_OUTLINE )     xmlw.TagL( "outline" ).EndL();
        if( attributes & FONT_SHADOW )      xmlw.TagL( "shadow" ).EndL();
        if( attributes & FONT_CONDENSE )    xmlw.TagL( "condense" ).EndL();
        if( attributes & FONT_EXTEND )      xmlw.TagL( "extend" ).EndL();

        xmlw.TagL( "sz" ).Attr( "val", font.size ).EndL();
        xmlw.TagL( FontTagName ).Attr( "val", font.name ).EndL();
        xmlw.TagL( "charset" ).Attr( "val", Charset ).EndL();
        if( font.theme || ( font.color == "" ) ) xmlw.TagL( "color" ).Attr( "theme", 1 ).EndL();
        else xmlw.TagL( "color" ).Attr( "rgb", font.color ).EndL();
    }

    void CWorkbook::AddImagesExtensions( XMLWriter & xmlw ) const
    {
        bool gif = false, jpg = false, jpeg = false, png = false, tif = false, tiff = false;
        for( std::vector< CImage * >::const_iterator it = m_images.begin(); it != m_images.end(); it++ )
            switch( ( * it )->Type )
            {
                case CImage::unknown    :   assert( false ); break;
                case CImage::gif        :   gif = true; break;
                case CImage::jpg        :   jpg = true; break;
                case CImage::jpeg       :   jpeg = true; break;
                case CImage::png        :   png = true; break;
                case CImage::tif        :   tif = true; break;
                case CImage::tiff       :   tiff = true; break;
            }
        if( gif ) xmlw.TagL( "Default" ).Attr( "Extension", "gif" ).Attr( "ContentType", "image/gif" ).EndL();
        if( jpg ) xmlw.TagL( "Default" ).Attr( "Extension", "jpg" ).Attr( "ContentType", "image/jpeg" ).EndL();
        if( jpeg ) xmlw.TagL( "Default" ).Attr( "Extension", "jpeg" ).Attr( "ContentType", "image/jpeg" ).EndL();
        if( png ) xmlw.TagL( "Default" ).Attr( "Extension", "png" ).Attr( "ContentType", "image/png" ).EndL();
        if( tif ) xmlw.TagL( "Default" ).Attr( "Extension", "tif" ).Attr( "ContentType", "image/tiff" ).EndL();
        if( tiff ) xmlw.TagL( "Default" ).Attr( "Extension", "tiff" ).Attr( "ContentType", "image/tiff" ).EndL();
    }

    // ****************************************************************************
    /// @brief  ...
    /// @return Boolean result of the operation
    // ****************************************************************************
    bool CWorkbook::SaveChain()
    {
        // [- zip/xl/calcChain.xml
        XMLWriter xmlw( m_pathManager->RegisterXML( "/xl/calcChain.xml" ) );
        xmlw.Tag( "calcChain" ).Attr( "xmlns", ns_book );
        for( std::vector<CWorksheet *>::const_iterator it = m_worksheets.begin(); it != m_worksheets.end(); it++ )
        {
            const CWorksheet * const Worksheet = * it;
            if( Worksheet->IsThereFormula() )
            {
                std::vector<std::string> vstrChain;
                Worksheet->GetCalcChain( vstrChain );
                for( size_t j = 0; j < vstrChain.size(); j++ )
                {
                    xmlw.TagL( "c" );
                    xmlw.Attr( "r", vstrChain[j] );
                    if( j == 0 ) xmlw.Attr( "i", Worksheet->GetIndex() );
                    xmlw.EndL();
                }
            }
        }
        xmlw.End( "calcChain" );
        // zip/xl/calcChain.xml -]
        return true;
    }

    // ****************************************************************************
    /// @brief  ...
    /// @return Boolean result of the operation
    // ****************************************************************************
    bool CWorkbook::SaveComments()
    {
        if( m_comments.empty() ) return true;

        std::vector<Comment *> sheet_comments;
        std::vector<std::vector<Comment *> > comments;
        std::sort( m_comments.begin(), m_comments.end() );

        size_t active_sheet = m_comments[0].sheetIndex;
        for( size_t i = 0; i < m_comments.size(); i++ )
        {
            if( m_comments[i].sheetIndex == active_sheet )
            {
                sheet_comments.push_back( &m_comments[i] );
            }
            else
            {
                active_sheet = m_comments[i].sheetIndex;
                comments.push_back( sheet_comments );
                sheet_comments.clear();

                i--;
            }
        }

        comments.push_back( sheet_comments );
        sheet_comments.clear();
        for( std::vector<std::vector<Comment *> >::const_iterator it = comments.begin(); it != comments.end(); it++ )
            if( ! SaveCommentList( * it ) ) return false;
        return true;
    }

    // ****************************************************************************
    /// @brief  ...
    /// @param	comments comment list on the sheet
    /// @return Boolean result of the operation
    // ****************************************************************************
    bool CWorkbook::SaveCommentList( const std::vector<Comment *> & comments )
    {
        assert( ! comments.empty() );
        {
            // [- zip/xl/commentsN.xml
            std::stringstream FileName;
            FileName << "/xl/comments" << comments[ 0 ]->sheetIndex << ".xml";

            XMLWriter xmlw( m_pathManager->RegisterXML( FileName.str() ) );
            xmlw.Tag( "comments" ).Attr( "xmlns", ns_book );
            xmlw.Tag( "authors" ).TagOnlyContent( "author", m_UserName ).End().Tag( "commentList" );
            for( std::vector<Comment *>::const_iterator it = comments.begin(); it != comments.end(); it++ )
                AddComment( xmlw, * * it );
            xmlw.End( "commentList" ).End( "comments" );
            // zip/xl/commentsN.xml -]
        }
        {
            // [- zip/xl/drawings/vmlDrawingN.xml
            std::stringstream FileName;
            FileName << "/xl/drawings/vmlDrawing" << comments[ 0 ]->sheetIndex << ".vml";

            XMLWriter xmlw( m_pathManager->RegisterXML( FileName.str() ) );
            xmlw.Tag( "xml" ).Attr( "xmlns:v", "urn:schemas-microsoft-com:vml" );
            xmlw.Attr( "xmlns:o", "urn:schemas-microsoft-com:office:office" );
            xmlw.Attr( "xmlns:x", "urn:schemas-microsoft-com:office:excel" );
            xmlw.Tag( "o:shapelayout" ).Attr( "v:ext", "edit" );
            xmlw.TagL( "o:idmap" ).Attr( "v:ext", "edit" ).Attr( "data", 1 ).EndL();
            xmlw.End( "o:shapelayout" );

            xmlw.Tag( "v:shapetype" ).Attr( "id", "_x0000_t202" ).Attr( "coordsize", "21600,21600" );
            xmlw.Attr( "o:spt", 202 ).Attr( "path", "m,l,21600r21600,l21600,xe" );
            xmlw.TagL( "v:stroke" ).Attr( "joinstyle", "miter" ).EndL();
            xmlw.TagL( "v:path" ).Attr( "gradientshapeok", "t" ).Attr( "o:connecttype", "rect" ).EndL();
            xmlw.End( "v:shapetype" );

            for( std::vector<Comment *>::const_iterator it = comments.begin(); it != comments.end(); it++ )
                AddCommentDrawing( xmlw, * * it );
            xmlw.End( "xml" );
            // zip/xl/drawings/vmlDrawingN.xml -]
        }
        return true;
    }

    // ****************************************************************************
    /// @brief  ...
    /// @param  stream reference to xml stream
    /// @param	comments comment list on the sheet
    /// @return Boolean result of the operation
    // ****************************************************************************
    void CWorkbook::AddComment( XMLWriter & xmlw, const Comment & comment ) const
    {
        xmlw.Tag( "comment" ).Attr( "ref", comment.cellRef.ToString() ).Attr( "authorId", 0 ).Tag( "text" );
        for( std::list<std::pair<Font, UniString> >::const_iterator it = comment.contents.begin(); it != comment.contents.end(); it++ )
        {
            xmlw.Tag( "r" ).Tag( "rPr" );
            AddFontInfo( xmlw, it->first, "rFont", 1 );
            xmlw.End( "rPr" ).TagOnlyContent( "t", it->second ).End( "r" );
        }
        xmlw.End( "text" ).End( "comment" );
    }

    // ****************************************************************************
    /// @brief  ...
    /// @param  stream reference to xml stream
    /// @param	comments comment list on the sheet
    /// @return Boolean result of the operation
    // ****************************************************************************
    void CWorkbook::AddCommentDrawing( XMLWriter & xmlw, const Comment & comment )
    {
        std::stringstream IdValue, StyleValue;
        IdValue << "_x0000_s" << 1000u + ( ++m_commLastId );

        if( ( comment.x >= 0 ) && ( comment.y >= 0 ) && ( comment.width > 0 ) && ( comment.height > 0 ) )
        {
            StyleValue << "position:absolute; margin-left:" << comment.x << "pt;margin-top:" << comment.y;
            StyleValue << "pt;width:" << comment.width << "pt;height:" << comment.height << "pt; z-index:2";
        }
        if( comment.isHidden ) StyleValue << ";visibility:hidden";

        bool wrapText = false;
        for( std::list<std::pair<Font, UniString> >::const_iterator it = comment.contents.begin(); it != comment.contents.end(); it++ )
            //if( it->second.find( "\n" ) != std::string::npos )
            if( static_cast< std::string >( it->second ).find( "\n" ) != std::string::npos )
            {
                wrapText = true;
                break;
            }

        if( wrapText ) StyleValue << ";mso-wrap-style:tight";

        xmlw.Tag( "v:shape" ).Attr( "id", IdValue.str() ).Attr( "type", "#_x0000_t202" ).Attr( "style", StyleValue.str() );
        /*             */xmlw.Attr( "fillcolor", comment.fillColor ).Attr( "o:insetmode", "auto" );
        xmlw.TagL( "v:fill" ).Attr( "color2", comment.fillColor ).EndL();
        xmlw.TagL( "v:shadow" ).Attr( "on", "t" ).Attr( "color", "black" ).Attr( "obscured", "t" ).EndL();
        xmlw.TagL( "v:path" ).Attr( "o:connecttype", "none" ).EndL();

        xmlw.Tag( "v:textbox" ).Attr( "style", "mso-direction-alt:auto" );
        xmlw.TagL( "div" ).Attr( "style", "text-align:left" ).EndL();
        xmlw.End( "v:textbox" );

        xmlw.Tag( "x:ClientData" ).Attr( "ObjectType", "Note" );
        xmlw.TagL( "x:MoveWithCells" ).EndL();
        xmlw.TagL( "x:SizeWithCells" ).EndL();
        xmlw.TagOnlyContent( "x:AutoFill", "False" );
        xmlw.TagOnlyContent( "x:Row", comment.cellRef.row - 1 );
        xmlw.TagOnlyContent( "x:Column", comment.cellRef.col );
        xmlw.End( "x:ClientData" );

        xmlw.End( "v:shape" );
    }

    // ****************************************************************************
    /// @brief  ...
    /// @return Boolean result of the operation
    // ****************************************************************************
    bool CWorkbook::SaveSharedStrings()
    {
        // [- zip/xl/sharedStrings.xml
        if( m_sharedStrings.empty() ) return true;

        XMLWriter xmlw( m_pathManager->RegisterXML( "/xl/sharedStrings.xml" ) );
        xmlw.Tag( "sst" ).Attr( "xmlns", ns_book ).Attr( "count", m_sharedStrings.size() ).Attr( "uniqueCount", m_sharedStrings.size() );

        std::vector< std::pair<const std::string, uint64_t> *> pointers_to_hash;
        pointers_to_hash.resize( m_sharedStrings.size() );
        for( std::map<std::string, uint64_t>::iterator it = m_sharedStrings.begin(); it != m_sharedStrings.end(); it++ )
            pointers_to_hash[it->second] = & ( * it );

        std::vector< std::pair<const std::string, uint64_t> *>::const_iterator it = pointers_to_hash.begin();
        for( ; it != pointers_to_hash.end(); it++ )
            xmlw.Tag( "si" ).TagOnlyContent( "t", ( * it )->first ).End( "si" );

        xmlw.End( "sst" );
        // zip/xl/sharedStrings.xml -]
        return true;
    }

    // ****************************************************************************
    /// @brief  ...
    /// @return Boolean result of the operation
    // ****************************************************************************
    bool CWorkbook::SaveWorkbook()
    {
        char szId[ 16 ] = { 0 };
        {
            // [- zip/xl/_rels/workbook.xml.rels
            XMLWriter xmlw( m_pathManager->RegisterXML( "/xl/_rels/workbook.xml.rels" ) );
            xmlw.Tag( "Relationships" ).Attr( "xmlns", ns_relationships );

            bool bFormula = false;
            for( std::vector<CWorksheet *>::const_iterator it = m_worksheets.begin(); it != m_worksheets.end(); it++ )
            {
                //sprintf( szId, "rId%zu", ( * it )->GetIndex() );
                sprintf( szId, "rId%u", unsigned( ( * it )->GetIndex() ) );
                std::stringstream PropValue;
                PropValue << "worksheets/sheet" << ( * it )->GetIndex() << ".xml";
                xmlw.TagL( "Relationship" ).Attr( "Id", szId ).Attr( "Type", type_sheet ).Attr( "Target", PropValue.str() ).EndL();
                if( ( * it )->IsThereFormula() ) bFormula = true;
            }
            for( std::vector<CChartsheet *>::const_iterator it = m_chartsheets.begin(); it != m_chartsheets.end(); it++ )
            {
                //sprintf( szId, "rId%zu", ( * it )->GetIndex() );
                sprintf( szId, "rId%u", unsigned( ( * it )->GetIndex() ) );
                std::stringstream PropValue;
                PropValue << "chartsheets/sheet" << ( * it )->GetIndex() << ".xml";
                xmlw.TagL( "Relationship" ).Attr( "Id", szId ).Attr( "Type", type_chartsheet ).Attr( "Target", PropValue.str() ).EndL();
            }
            size_t id = m_sheetId;
            if( bFormula )
            {
                //sprintf( szId, "rId%zu", id++ );
                sprintf( szId, "rId%u", unsigned( id++ ) );
                xmlw.TagL( "Relationship" ).Attr( "Id", szId ).Attr( "Type", type_chain ).Attr( "Target", "calcChain.xml" ).EndL();
            }
            if( ! m_sharedStrings.empty() )
            {
                //sprintf( szId, "rId%zu", id++ );
                sprintf( szId, "rId%u", unsigned( id++ ) );
                xmlw.TagL( "Relationship" ).Attr( "Id", szId ).Attr( "Type", type_sharedStr ).Attr( "Target", "sharedStrings.xml" ).EndL();
            }
            //sprintf( szId, "rId%zu", id++ );
            sprintf( szId, "rId%u", unsigned( id++ ) );
            xmlw.TagL( "Relationship" ).Attr( "Id", szId ).Attr( "Type", type_style ).Attr( "Target", "styles.xml" ).EndL();
            //sprintf( szId, "rId%zu", id++ );
            sprintf( szId, "rId%u", unsigned( id++ ) );
            xmlw.TagL( "Relationship" ).Attr( "Id", szId ).Attr( "Type", type_theme ).Attr( "Target", "theme/theme1.xml" ).EndL();

            xmlw.End( "Relationships" );
            // zip/xl/_rels/workbook.xml.rels -]
        }
        {
            // [- zip/xl/workbook.xml
            XMLWriter xmlw( m_pathManager->RegisterXML( "/xl/workbook.xml" ) );
            xmlw.Tag( "workbook" ).Attr( "xmlns", ns_book ).Attr( "xmlns:r", ns_book_r );
            xmlw.TagL( "fileVersion" ).Attr( "appName", "xl" ).Attr( "lastEdited", 5 );
            /*                  */xmlw.Attr( "lowestEdited", 5 ).Attr( "rupBuild", 9303 ).EndL();
            xmlw.TagL( "workbookPr" ).Attr( "codeName", "ThisWorkbook" ).Attr( "defaultThemeVersion", 124226 ).EndL();
            xmlw.Tag( "bookViews" );
            xmlw.TagL( "workbookView" ).Attr( "xWindow", 270 ).Attr( "yWindow", 630 );
            /*                   */xmlw.Attr( "windowWidth", 24615 ).Attr( "windowHeight", 11445 );
            /*                   */xmlw.Attr( "activeTab", m_activeSheetIndex < m_sheetId - 1 ? m_activeSheetIndex : 0 ).EndL();
            xmlw.End( "bookViews" );

            if( ! m_worksheets.empty() || ! m_chartsheets.empty() )
            {
                xmlw.Tag( "sheets" );
                //Sheets ordering
                for( size_t i = 1; i < m_sheetId; i++ )
                {
                    //sprintf( szId, "rId%zu", i );
                    sprintf( szId, "rId%u", unsigned( i ) );
                    bool Found = false;
                    for( std::vector<CWorksheet *>::const_iterator it = m_worksheets.begin(); it != m_worksheets.end(); it++ )
                        if( i == ( * it )->GetIndex() )
                        {
                            Found = true;
                            xmlw.TagL( "sheet" ).Attr( "name", ( * it )->GetTitle() ).Attr( "sheetId", i ).Attr( "r:id", szId ).EndL();
                            break;
                        }
                    if( Found ) continue;
                    for( std::vector<CChartsheet *>::const_iterator it = m_chartsheets.begin(); it != m_chartsheets.end(); it++ )
                        if( i == ( * it )->GetIndex() )
                        {
                            xmlw.TagL( "sheet" ).Attr( "name", ( * it )->Chart().GetTitle() ).Attr( "sheetId", i ).Attr( "r:id", szId ).EndL();
                            break;
                        }
                }
                xmlw.End( "sheets" );
            }
            else return false;
            xmlw.TagL( "calcPr" ).Attr( "calcId", 124519 ).EndL().End( "workbook" );
            // zip/xl/workbook.xml -]
        }
        return true;
    }

}	// namespace SimpleXlsx
