
// The following ifdef block is the standard way of creating macros which make exporting 
// from a DLL simpler. All files within this DLL are compiled with the BIG_DLL_EXPORTS
// symbol defined on the command line. this symbol should not be defined on any project
// that uses this DLL. This way any other project whose source files include this file see 
// BIG_DLL_API functions as being imported from a DLL, wheras this DLL sees symbols
// defined with this macro as being exported.
#ifdef BIG_DLL_EXPORTS
#define BIG_DLL_API __declspec(dllexport)
#else
#define BIG_DLL_API __declspec(dllimport)
#endif

extern BIG_DLL_API int nBIG_DLL;

int __stdcall Encode(char * userdata,char * header,struct RAW_SECTOR_MODE1 * rmode);

int __stdcall Decode(struct RAW_SECTOR_MODE1 * rmode,int D0);

int __stdcall Fast(char * bmp,char * bDib,int IM_HEIGHT,int IM_WIDTH,int Width,int Height);

int __stdcall ModC(int   , int );

int __stdcall ShiftLeftC(int *, int );

int __stdcall ShiftRightC(int *, int );

int __stdcall MultiplayC(int , int );

int __stdcall ModA(unsigned int  value,unsigned int  count);