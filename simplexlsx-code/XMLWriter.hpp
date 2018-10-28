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

#ifndef XMLWRITER_H
#define XMLWRITER_H

#include <cassert>
#include <iostream>
#include <ostream>
#include <stack>
#include <string>

namespace SimpleXlsx
{

    class XMLWriter
    {
        public:
            inline XMLWriter( const std::string & FileName ) : m_TagOpen( false ), m_SelfClosed( true )
            {
                Init( FileName );
            }

            inline ~XMLWriter()
            {
                EndAll();
                DebugCheckIsLightTagOpened();
            }

            inline bool IsOk() const
            {
                return m_OStream.is_open();
            }

            //Returns the current precision of floating point
            std::streamsize GetFloatPrecision()
            {
                return m_OStream.precision();
            }

            //Set the current precision for floating point.
            //Returns the precision before the call this function.
            std::streamsize SetFloatPrecision( std::streamsize NewPrecision )
            {
                return m_OStream.precision( NewPrecision );
            }

            //Light version without using stack of Tag Names.
            //No internal elements/tags or content string. Only attributes accepted.
            //Must be used with EndL.
            inline XMLWriter & TagL( const char * TagName )
            {
                assert( TagName != NULL );
                DebugCheckAndIncLightTag();
                CloseOpenedTag();
                m_OStream << '<' << TagName;
                m_TagOpen = true;
                m_SelfClosed = false;
                return * this;
            }

            //Light version without using stack of Tag Names.
            //No internal elements/tags or content string. Only attributes accepted.
            //Must be used with TagL.
            inline XMLWriter & EndL()
            {
                DebugCheckAndDecLightTag();
                m_OStream << "/>";
                m_TagOpen = false;
                return * this;
            }

            inline XMLWriter & Tag( const char * TagName )
            {
                assert( TagName != NULL );
                CloseOpenedTag();
                DebugCheckIsLightTagOpened();
                m_OStream << '<' << TagName;
                m_TagOpen = true;
                m_SelfClosed = true;
                m_Tags.push( TagName );
                return * this;
            }

            inline XMLWriter & End( const char * TagName = NULL )
            {
                assert( ! m_Tags.empty() );
                DebugCheckIsLightTagOpened();
                if( m_SelfClosed ) m_OStream << "/>";
                else m_OStream << "</" << m_Tags.top() << '>';
#ifndef NDEBUG
                if( TagName != NULL )
                {
                    if( m_Tags.top() != TagName )
                        std::cerr << "Wrong TagName for End: " << TagName << ". Wanted: " << m_Tags.top() << std::endl;
                    assert( m_Tags.top() == TagName );
                }
#else
                ( void )TagName;
#endif
                m_Tags.pop();
                m_TagOpen = false;
                m_SelfClosed = false;
                return * this;
            }

            inline XMLWriter & EndAll()
            {
                while( ! m_Tags.empty() )
                    this->End();
                return * this;
            }

            inline XMLWriter & TagOnlyContent( const char * TagName, const char * ContentString )
            {
                assert( TagName != NULL );
                CloseOpenedTag();
                DebugCheckIsLightTagOpened();
                m_OStream << '<' << TagName << '>';
                WriteStringEscape( ContentString );
                m_OStream << "</" << TagName << '>';
                m_SelfClosed = false;
                return * this;
            }

            inline XMLWriter & TagOnlyContent( const char * TagName, const std::string  & ContentString )
            {
                return TagOnlyContent( TagName, ContentString.c_str() );
            }

            //TagOnlyContent() template for all streamable types
            template <typename _T>
            inline XMLWriter & TagOnlyContent( const char * TagName, _T Value )
            {
                CloseOpenedTag();
                DebugCheckIsLightTagOpened();
                m_OStream << '<' << TagName << '>' << Value << "</" << TagName << '>';
                m_SelfClosed = false;
                return *this;
            }

            template <typename _T>
            inline XMLWriter & TagOnlyContent( const char * TagName, _T * Value );

            //Fast writing an attribute for the current Tag
            inline XMLWriter & Attr( const char * AttrName, const char * String )
            {
                assert( AttrName != NULL );
                assert( m_TagOpen );
                m_OStream << ' ' << AttrName << "=\"" << String << '"';
                return * this;
            }

            //Write an attribute for the current Tag
            inline XMLWriter & Attr( const char * AttrName, char * String )
            {
                return AttrInt( AttrName, String );
            }

            //Attr() overload for std::string type
            inline XMLWriter & Attr( const char * AttrName, const std::string & String )
            {
                return AttrInt( AttrName, String.c_str() );
            }

            //Attr() function template for all streamable types
            template <typename _T>
            inline XMLWriter & Attr( const char * AttrName, _T Value )
            {
                assert( AttrName != NULL );
                assert( m_TagOpen );
                m_OStream.put( ' ' );
                m_OStream << AttrName << "=\"" << Value;
                m_OStream.put( '"' );
                return *this;
            }

            template <typename _T>
            inline XMLWriter & Attr( const char * AttrName, _T * Value );

            //Write an content for the current Tag
            inline XMLWriter & Cont( const char * String )
            {
                CloseOpenedTag();
                DebugCheckIsLightTagOpened();
                WriteStringEscape( String );
                m_SelfClosed = false;
                return * this;
            }

            //Write an content for the current Tag
            inline XMLWriter & Cont( char * String )
            {
                return Cont( static_cast<const char *>( String ) );
            }

            //Content() overload for std::string type
            inline XMLWriter & Cont( const std::string & String )
            {
                return Cont( String.c_str() );
            }

            //Content() template for all streamable types
            template <typename _T>
            inline XMLWriter & Cont( _T Value )
            {
                CloseOpenedTag();
                DebugCheckIsLightTagOpened();
                m_OStream << Value;
                m_SelfClosed = false;
                return *this;
            }

            template <typename _T>
            inline XMLWriter & Cont( _T * Value );

        private:
            bool                    m_TagOpen, m_SelfClosed;
            std::ofstream           m_OStream;
            std::stack<std::string> m_Tags;

            inline void Init( const std::string & FileName )
            {
                assert( ! FileName.empty() );
#ifndef NDEBUG
                m_LightTagCounter = 0;
#endif
                m_OStream.open( FileName.c_str(), std::ios_base::out );
                m_OStream.imbue( std::locale( "C" ) );
                m_OStream << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
            }

            inline void CloseOpenedTag()
            {
                if( ! m_TagOpen ) return;
                m_OStream.put( '>' );
                m_TagOpen = false;
            }

            inline void WriteStringEscape( const char * String )
            {
                for( ; * String; String++ )
                    switch( * String )
                    {
                        case '&'    :   m_OStream << "&amp;";        break;
                        case '<'    :   m_OStream << "&lt;";         break;
                        case '>'    :   m_OStream << "&gt;";         break;
                        case '\''   :   m_OStream << "&apos;";       break;
                        case '"'    :   m_OStream << "&quot;";       break;
                        default     :   m_OStream.put( * String );   break;
                    }
            }

            //Write an attribute for the current Tag
            inline XMLWriter & AttrInt( const char * AttrName, const char * String )
            {
                assert( AttrName != NULL );
                assert( m_TagOpen );
                m_OStream << ' ' << AttrName << "=\"";
                WriteStringEscape( String );
                m_OStream.put( '"' );
                return * this;
            }

            //Debug version for checking TagL and EndL
#ifdef NDEBUG
            inline void DebugCheckAndIncLightTag()      {}
            inline void DebugCheckAndDecLightTag()      {}
            inline void DebugCheckIsLightTagOpened()    {}
#else
            intptr_t    m_LightTagCounter;

            inline void DebugCheckAndIncLightTag()
            {
                assert( m_LightTagCounter == 0 );
                m_LightTagCounter++;
            }

            inline void DebugCheckAndDecLightTag()
            {
                assert( m_LightTagCounter == 1 );
                m_LightTagCounter--;
            }
            inline void DebugCheckIsLightTagOpened()
            {
                assert( m_LightTagCounter == 0 );
            }
#endif
    };
}

#endif // XMLWRITER_H
