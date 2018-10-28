// DictCtrl.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include <afx.h>
#include "windows.h"
#include <stdio.h>
#include <comutil.h>
#include <string>
#include "IDictionary.h"
#include "ISpellCheck.h"
using namespace std;
// {54BF6567-1007-11D1-B0AA-444553540000}
extern "C" const GUID CLSID_Dictionary = 
		{ 0x54bf6567, 0x1007, 0x11d1,
		{ 0xb0, 0xaa, 0x44, 0x45, 0x53, 0x54, 0x00, 0x00} } ;

extern "C" const GUID IID_Dictionary = 
		{ 0x54bf6568, 0x1007, 0x11d1,
		{ 0xb0, 0xaa, 0x44, 0x45, 0x53, 0x54, 0x00, 0x00} } ;

extern "C" const GUID IID_SpellCheck = 
		{ 0x54bf6569, 0x1007, 0x11d1,
		{ 0xb0, 0xaa, 0x44, 0x45, 0x53, 0x54, 0x00, 0x00} } ;




//wstring×ª»»³Éstring
string WChar2Ansi(LPCWSTR pwszSrc)
{	
	int nLen = WideCharToMultiByte(CP_ACP, 0, pwszSrc, -1, NULL, 0, NULL, NULL);
	if (nLen <= 0) return std::string(""); 
	char* pszDst = new char[nLen];
	if (NULL == pszDst) return std::string(""); 
	WideCharToMultiByte(CP_ACP, 0, pwszSrc, -1, pszDst, nLen, NULL, NULL);
	pszDst[nLen - 1] = 0; 	std::string strTemp(pszDst);
	delete[] pszDst; 
	return strTemp;
}


LPCWSTR String2LPCWSTR(CString str)
{
	USES_CONVERSION;
	LPCWSTR lp_str = A2CW(W2A(str));
	return lp_str;
}




int main(int argc, char* argv[])
{
	IUnknown *pUnknown;
	IDictionary *pDictionary;
	ISpellCheck *pSpellCheck;
	String stringResult;
	BOOL bResult;
	HRESULT hResult;

	if (CoInitialize(NULL) != S_OK) {
		printf("Initialize COM library failed!\n");
		return -1;
	}

	GUID dictionaryCLSID;
	hResult = ::CLSIDFromProgID(L"Dictionary.Object", &dictionaryCLSID);
	if (hResult != S_OK) 
	{
		printf("Can't find the dictionary CLSID!\n");
		return -2;
	}
	
	hResult = CoCreateInstance(dictionaryCLSID, NULL,
		CLSCTX_INPROC_SERVER, IID_IUnknown, (void **)&pUnknown);
	if (hResult != S_OK) 
	{
		printf("Create object failed!\n");
		return -2;
	}

	hResult = pUnknown->QueryInterface(IID_Dictionary, (void **)&pDictionary);
	if (hResult != S_OK) {
		pUnknown->Release();
		printf("QueryInterface IDictionary failed!\n");
		return -3;
	}
	//2018.10.8---------------------------------------------------------------
	CString str1("animal.dict");
	USES_CONVERSION;
	LPCWSTR lp_str = A2CW(W2A(str1));

	bResult = pDictionary->LoadLibrary((String)lp_str);//2018.10.8
	if (bResult) {
		String stringResult;
		bResult = pDictionary->LookupWord((String)String2LPCWSTR(L"tiger"), &stringResult);
		
		if (bResult) {
			char *pTiger = _com_util::ConvertBSTRToString((BSTR)stringResult);
			printf("find the word \"tiger\" -- %s\n", pTiger);
			delete pTiger;
		}

		pDictionary->InsertWord((String)String2LPCWSTR(L"elephant"), (String)String2LPCWSTR(L"Ïó"));
		bResult = pDictionary->LookupWord((String)String2LPCWSTR(L"elephant"), &stringResult);
		if (bResult) {

			pDictionary->RestoreLibrary((String)String2LPCWSTR(L"animal1.dict"));
		}
	} else {
		printf("Load Library \"animal.dict\"\n");
	}
	
	hResult = pDictionary->QueryInterface(IID_SpellCheck, (void **)&pSpellCheck);
	pDictionary->Release();
	if (hResult != S_OK) {
		pUnknown->Release();
		printf("QueryInterface IDictionary failed!\n");
		return -4;
	}

	bResult = pSpellCheck->CheckWord((String)String2LPCWSTR(L"lion"), &stringResult);
	if (bResult) {
		printf("Word \"lion\" spelling right.\n");
	} else {
		char *pLion = _com_util::ConvertBSTRToString((BSTR)stringResult);
		printf("Word \"lion\" spelling is wrong. Maybe it is %s.\n", pLion);
		delete pLion;
	}
	bResult = pSpellCheck->CheckWord((String)String2LPCWSTR(L"dot"), &stringResult);
	if (bResult) {
		printf("Word \"dot\" spelling right.\n");
	} else {
		char *pDot = _com_util::ConvertBSTRToString((BSTR)stringResult);
		printf("Word \"dot\" spelling is wrong. Maybe it is %s.\n", pDot);
		delete pDot;
	}

	pSpellCheck->Release();
	if (pUnknown->Release()== 0) 
		printf("The reference count of dictionary object is zero.");

	CoUninitialize();
	return 0;
}
