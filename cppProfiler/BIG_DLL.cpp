// BIG_DLL.cpp : Defines the entry point for the DLL application.
//

#include "stdafx.h"
#include "BIG_DLL.h"


BOOL APIENTRY DllMain( HANDLE hModule, 
                       DWORD  ul_reason_for_call, 
                       LPVOID lpReserved
					 )
{
    switch (ul_reason_for_call)
	{
		case DLL_PROCESS_ATTACH:
		case DLL_THREAD_ATTACH:
		case DLL_THREAD_DETACH:
		case DLL_PROCESS_DETACH:
			break;
    }
    return TRUE;
}


// This is an example of an exported variable
BIG_DLL_API int nBIG_DLL=0;



// exported function fast
int __stdcall Fast(char * bmp,char * bDib,int IM_HEIGHT,int IM_WIDTH,int Width,int Height)
{
	int y=0,x=0,c=0;
	double yy=0,xx=0;
    int n = 0;
	int buf=0;

    for (y=(IM_HEIGHT - 1); y>=0 ;y--){
      yy=y*Height;
	  yy=yy/IM_HEIGHT;

      for (x=0;x<=(IM_WIDTH-1);x++){
        xx = x*Width;
		xx = xx/IM_WIDTH;

        for (c=0;c<=2;c++) {
          bmp[n] =bDib[((int)xx*3 + c)*(Height+1)+(int)yy];
          n++;
        }
	  }
    }
	return n;
}

/*int __stdcall ModA(unsigned int  value,unsigned  int  count)
{
	//x=(a1*xprev+13849) % 65536;
	return ((count*value+13849) % 65536);
}*/


// exported function Encode.
/*int __stdcall Encode(char * userdata,char * header,struct RAW_SECTOR_MODE1 * rmode)
{
	typedef int (__cdecl *GENECC)(char *,char *,struct RAW_SECTOR_MODE1 *);
    
	HINSTANCE hDLL;					  // Handle to DLL

	GENECC GenECCAndEDC_Mode1;	      // Function pointer

	int a;

	hDLL = LoadLibrary("ECC.dll");
	if (!hDLL) return -5;

	//--------exporting functions------------
    GenECCAndEDC_Mode1 = (GENECC)GetProcAddress(hDLL,"GenECCAndEDC_Mode1");

	a = GenECCAndEDC_Mode1(userdata, header, rmode);

	return a;
}*/


// exported function Decode.
/*int __stdcall Decode(struct RAW_SECTOR_MODE1 * rmode,int D0)
{
	typedef int (__cdecl *CHECKSECTOR)(struct RAW_SECTOR *,int);
    
	HINSTANCE hDLL;                   // Handle to DLL

	CHECKSECTOR CheckSector;	      // Function pointer

	int a;

	hDLL = LoadLibrary("ECC.dll");
	if (!hDLL) return -5;

	//--------exporting functions------------
    CheckSector = (CHECKSECTOR)GetProcAddress(hDLL,"CheckSector");

	a=CheckSector((struct RAW_SECTOR*)rmode,D0);

	return a;

}*/