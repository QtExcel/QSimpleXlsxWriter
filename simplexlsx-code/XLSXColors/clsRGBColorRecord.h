#ifndef CLSRGBCOLORRECORD_H
#define CLSRGBCOLORRECORD_H
#include <string>
namespace SimpleXlsx
{
class clsRGBColorRecord
{
    public:
        clsRGBColorRecord();
        clsRGBColorRecord( const unsigned char gs) { Set(gs);};
        clsRGBColorRecord( const unsigned char r, const unsigned char g, const unsigned char b) { Set(r,g,b);};
        clsRGBColorRecord( const double gs){ Set(gs);};
        clsRGBColorRecord( const char (&sharpn)[7]){ Set(sharpn);};
        virtual ~clsRGBColorRecord();
        clsRGBColorRecord(const clsRGBColorRecord& other);
        clsRGBColorRecord& operator=(const clsRGBColorRecord& other);
        const char * Get() const {return colstr;};
        void Set(const unsigned char r, const unsigned char g, const unsigned char b);
        void Set(const unsigned char chargs);
        void Set(const double gs);
        void Set(const char (&sharpn)[7]);
    protected:
        char colstr[7];
        void Copy(const clsRGBColorRecord& other);
    private:
};
};
#endif // CLSRGBCOLORRECORD_H
