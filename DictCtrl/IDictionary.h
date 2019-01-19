#ifndef __IDictionary_H__
#define __IDictionary_H__

#include "Unknwn.h"

typedef unsigned short *String;

// {54BF6568-1007-11D1-B0AA-444553540000}
extern "C" const GUID IID_Dictionary;

class IDictionary : public IUnknown 
{
	public : 
		virtual BOOL __stdcall Initialize() = 0;
		virtual BOOL __stdcall LoadLibrary(LPCWSTR) = 0;
		virtual BOOL __stdcall InsertWord(BSTR, BSTR) = 0;
		virtual void __stdcall DeleteWord(BSTR) = 0;
		virtual BOOL __stdcall LookupWord(BSTR, BSTR *) = 0;
		virtual BOOL __stdcall RestoreLibrary(BSTR) = 0;
		virtual void __stdcall FreeLibrary() = 0;
};

#endif // __IDictionary_H__
