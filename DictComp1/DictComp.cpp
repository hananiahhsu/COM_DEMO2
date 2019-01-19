// DictComp.cpp : Defines the entry point for the DLL application.
//

//#include "stdafx.h"
#include <afx.h>
#include <comutil.h>
#include <stdio.h>
#include "factory.h"
#include "objbase.h"
#include "olectl.h"

#include "DictComp.h"

#include "registry.h"
#include <cstring>
#include <exception>
#include <string>
using namespace std;
ULONG    g_LockNumber = 0;
ULONG    g_DictionaryNumber = 0;
HANDLE	 g_hModule;

// {54BF6567-1007-11D1-B0AA-444553540000}
//extern "C" const GUID CLSID_Dictionary = 
//		{ 0x54bf6567, 0x1007, 0x11d1,
//		{ 0xb0, 0xaa, 0x44, 0x45, 0x53, 0x54, 0x00, 0x00} } ;

// {222B68F6-2B02-41AA-BA5E-8EFED4B72040}
extern "C" const GUID CLSID_Dictionary =
{ 0x222b68f6, 0x2b02, 0x41aa, 
{ 0xba, 0x5e, 0x8e, 0xfe, 0xd4, 0xb7, 0x20, 0x40 } };


extern "C" const GUID IID_Dictionary = 
		{ 0x54bf6568, 0x1007, 0x11d1,
		{ 0xb0, 0xaa, 0x44, 0x45, 0x53, 0x54, 0x00, 0x00} } ;

extern "C" const GUID IID_SpellCheck = 
		{ 0x54bf6569, 0x1007, 0x11d1,
		{ 0xb0, 0xaa, 0x44, 0x45, 0x53, 0x54, 0x00, 0x00} } ;



BOOL APIENTRY DllMain( HANDLE hModule, 
                       DWORD  ul_reason_for_call, 
                       LPVOID lpReserved
					 )
{
    g_hModule = hModule;
	return TRUE;
}


extern "C" HRESULT __stdcall DllGetClassObject(const CLSID& clsid, const IID& iid, void **ppv)
{
	if (clsid == CLSID_Dictionary ) {
		
		CDictionaryFactory *pFactory = new CDictionaryFactory;
		
		if (pFactory == NULL) {
			return E_OUTOFMEMORY ;
		}
		
		HRESULT result = pFactory->QueryInterface(iid, ppv);

		return result;
	} else {
		return CLASS_E_CLASSNOTAVAILABLE;
	}
}

extern "C" HRESULT __stdcall DllCanUnloadNow(void)
{
	if ((g_DictionaryNumber == 0) && (g_LockNumber == 0))
		return S_OK;
	else
		return S_FALSE;
}




char* ConvertLPWSTRToLPSTR(LPWSTR lpwszStrIn)
{
	LPSTR pszOut = NULL;
	try
	{
		if (lpwszStrIn != NULL)
		{
			int nInputStrLen = wcslen(lpwszStrIn);

			// Double NULL Termination  
			int nOutputStrLen = WideCharToMultiByte(CP_ACP, 0, lpwszStrIn, nInputStrLen, NULL, 0, 0, 0) + 2;
			pszOut = new char[nOutputStrLen];

			if (pszOut)
			{
				memset(pszOut, 0x00, nOutputStrLen);
				WideCharToMultiByte(CP_ACP, 0, lpwszStrIn, nInputStrLen, pszOut, nOutputStrLen, 0, 0);
			}
		}
	}
	catch (std::exception e)
	{
	}

	return pszOut;
}

/***********************************************************************
* 函数： CString2char 
* 描述：将CString转换为 char*
***********************************************************************/
//char* CString2char(CString &str)
//{
//	int len = str.GetLength();
//	char* chRtn = (char*)malloc((len * 2 + 1) * sizeof(char));//CString的长度中汉字算一个长度 
//	memset(chRtn, 0, 2 * len + 1);
//	USES_CONVERSION;
//	strcpy((LPSTR)chRtn, OLE2A(str.LockBuffer()));
//	return chRtn;
//}


//void TCHAR2Char(const TCHAR* tchar, char* _char)
//{
//	int iLength = 0;
//	iLength = WideCharToMultiByte(CP_ACP, 0, tchar, -1, NULL, 0, NULL, NULL);
//	WideCharToMultiByte(CP_ACP, 0, tchar, -1, _char, iLength, NULL, NULL);
//}


//char* CStringToChar(CString str)
//{
//	int n = str.GetLength();
//	int len = WideCharToMultiByte(CP_ACP, 0, str, str.GetLength(), NULL, 0, NULL, NULL);
//	char * pFileName = new char[len + 1];
//	WideCharToMultiByte(CP_ACP, 0, str, str.GetLength(), pFileName, len, NULL, NULL);
//	pFileName[len + 1] = '\0';
//	return pFileName;
//}

//void CStringToChar(CString str, char ch[])
//{
//	char* tmpch;
//	int wLen = WideCharToMultiByte(CP_ACP, 0, str, -1, NULL, 0, NULL, NULL);
//	tmpch = new char[wLen + 1];
//	WideCharToMultiByte(CP_ACP, 0, str, -1, tmpch, wLen, NULL, NULL);
//	for (int i = 0; i < wLen; ++i)
//		ch[i] = tmpch[i];
//}



//ok
//char* UsrCStringToChar(CString str)
//{
//	int n = str.GetLength(); 
//	int len = WideCharToMultiByte(CP_ACP, 0, str, n, NULL, 0, NULL, NULL); 
//	char *pChar = new char[len + 1]; //以字节为单位
//	WideCharToMultiByte(CP_ACP, 0, str, n, pChar, len, NULL, NULL); //宽字节编码转换成多字节编码
//	pChar[len + 1] = '\0'; //多字节字符以'\0'结束
//	return pChar;
//}



// ANSI To UNCODE转换
CString AnsiToUnicode(char * szAnsi, int len)
{
	CString str;
	//预转换，得到所需空间的大小
	int wcsLen;
	if (len > 0)
		wcsLen = len;
	else
		wcsLen = ::MultiByteToWideChar(CP_ACP, NULL, szAnsi, strlen(szAnsi), NULL, 0);
	//分配空间要给'\0'留个空间，MultiByteToWideChar不会给'\0'空间
	wchar_t* wszString = new wchar_t[wcsLen + 1];
	//转换
	::MultiByteToWideChar(CP_ACP, NULL, szAnsi, strlen(szAnsi), wszString, wcsLen);
	//最后加上'\0'
	wszString[wcsLen] = '\0'; 
	str = wszString;
	delete wszString;
	return str;
}



//
// Server registration
//
extern "C" HRESULT __stdcall DllRegisterServer()
{
	TCHAR szModule[1024];
	char  *pChPath=NULL;
	CString str_path;

	DWORD dwResult = GetModuleFileName((HMODULE)g_hModule, szModule, 1024);
	if (dwResult == 0)
	{
		MessageBox(NULL,L"REG FAILED",L"",MB_OK);
		return SELFREG_E_CLASS;
	}

	CString sMoudle(szModule);
	MessageBox(NULL, sMoudle, L"sMoudle", MB_OK);

	USES_CONVERSION;
	std::string s(W2A(sMoudle));
	CString str(s.c_str());
	MessageBox(NULL, str, L"---pChPath---", MB_OK);



	if (S_OK != RegisterServer(CLSID_Dictionary,
		s.c_str(),
		"Dictionary.Object",
		"Dictionary Component",
		NULL))
	{
		MessageBox(NULL, L"REG SERVER FAILED", L"", MB_OK);
	}
}


//
// Server unregistration
//
extern "C" HRESULT __stdcall DllUnregisterServer()
{
	return UnregisterServer(CLSID_Dictionary,
	                        "Dictionary.Object",NULL);
}

// class CDictionary implementation

CDictionary::CDictionary()
{
	m_Ref = 0;
	m_nWordNumber = 0;
	m_nStructNumber = 0;
	m_pData = NULL;
	g_DictionaryNumber ++ ;
}

CDictionary::~CDictionary()
{
	if (m_pData != NULL) 
	{
		delete [] m_pData;
	}
}

HRESULT  CDictionary::QueryInterface(const IID& iid, void **ppv)
{
	if ( iid == IID_IUnknown )
	{
		*ppv = (IDictionary *) this ;
				((IDictionary *)(*ppv))->AddRef() ;
	} else if ( iid == IID_Dictionary ) 
	{
		*ppv = (IDictionary *) this ;
				((IDictionary *)(*ppv))->AddRef() ;
	} else if ( iid == IID_SpellCheck ) 
	{
		*ppv = (ISpellCheck *) this ;
				((ISpellCheck *)(*ppv))->AddRef() ;
	}
	else
	{
		*ppv = NULL;
		return E_NOINTERFACE ;
	}
	return S_OK;
}

ULONG	  CDictionary::AddRef()
{
	m_Ref ++;
	return  (ULONG) m_Ref;
}

ULONG	  CDictionary::Release()
{
	m_Ref --;
	if (m_Ref == 0 ) {
		g_DictionaryNumber -- ;
		delete this;
		return 0;
	}
	return  (ULONG) m_Ref;
}

BOOL CDictionary::Initialize()
{
	m_nWordNumber = 0;
	m_nStructNumber = 0;
	if (m_pData != NULL) 
	{
		delete [] m_pData;
	}
	m_pData = NULL;
	return TRUE;
}

BOOL CDictionary::LoadLibrary(LPCWSTR filename)
{
	BSTR bstrText = ::SysAllocString((BSTR)filename);
	char *pFileName = _com_util::ConvertBSTRToString((BSTR)filename);
	FILE *fp;
	if( (fp = fopen( pFileName, "rt" )) == NULL ) {
		printf("Open dictionary file : %s failed.\n", pFileName);
		delete pFileName;
		return FALSE;
	}
	char LineBuffer[128];
	if (feof(fp)) {
		printf("It is a null file!\n");
		fclose(fp);
		delete pFileName;
		return FALSE; 
	}
	if (fgets(LineBuffer, 128, fp) == NULL) {
		printf("Read TotalNumber failed!\n");
		fclose(fp);
		delete pFileName;
		return FALSE; 
	}
	
	int nTotalNumber = 0;
	sscanf_s(LineBuffer, "%d", &nTotalNumber);
	if ( (nTotalNumber < 1) && (nTotalNumber > 5000) ) {
		printf("The Number of words is invalid!\n");
		fclose(fp);
		delete pFileName;
		return FALSE; 
	}

	Initialize();
	m_nStructNumber = nTotalNumber+100;
	m_pData = new DictWord[m_nStructNumber];
	char c = '\0';
	memcpy(m_pData->wordForLang1,&c,1);
	memcpy(m_pData->wordForLang2,&c,1);
	m_nWordNumber = 0;
	
	while(!feof(fp)) {
		if (fgets(LineBuffer, MaxWordLength, fp) == NULL) {
			printf("Read the first string failed!\n");
			break;
		}
		sscanf(LineBuffer, "%s", m_pData[m_nWordNumber].wordForLang1);
		if (fgets(LineBuffer, MaxWordLength, fp) == NULL) {
			printf("Read the second string failed!\n");
			break;
		}
		sscanf(LineBuffer, "%s", m_pData[m_nWordNumber].wordForLang2);
		m_nWordNumber ++;
		if (m_nWordNumber == nTotalNumber)
			break;
		if (m_nWordNumber > m_nStructNumber)
			break;
	}	
	
	fclose(fp);
	delete pFileName;
	return TRUE;
}

BOOL CDictionary::InsertWord(BSTR word1, BSTR word2)
{
	char *pWord1, *pWord2;
	if (m_nWordNumber < m_nStructNumber) {
		pWord1 = _com_util::ConvertBSTRToString((BSTR)word1);
		pWord2 = _com_util::ConvertBSTRToString((BSTR)word2);

		if (strlen(pWord1) > MaxWordLength)
			*(pWord1+MaxWordLength-1) = '\0';
		if (strlen(pWord2) > MaxWordLength)
			*(pWord2+MaxWordLength-1) = '\0';
		sscanf(pWord1, "%s", m_pData[m_nWordNumber].wordForLang1);//2019.1.19
		sscanf(pWord2, "%s", m_pData[m_nWordNumber].wordForLang2);//2019.1.19
		//strcpy(m_pData[m_nWordNumber].wordForLang1, pWord1);
		//strcpy(m_pData[m_nWordNumber].wordForLang2, pWord2);
		m_nWordNumber ++ ;
		delete pWord1;
		delete pWord2;
		return TRUE;
	}
	return FALSE;
}

void CDictionary::DeleteWord(BSTR word)
{
	char *pWord = _com_util::ConvertBSTRToString((BSTR)word);
	char *pUpperWord = strupr(pWord);
	for (int i = 0; i < m_nWordNumber; i++)
	{
		char *tmpWord = strupr(m_pData[i].wordForLang1);
		if (strcmp(tmpWord, pWord) == 0) {
			for(int j = i + 1; j < m_nWordNumber; j++) {
				strcpy( m_pData[j].wordForLang1, m_pData[j + 1].wordForLang1);
				strcpy( m_pData[j].wordForLang2, m_pData[j + 1].wordForLang2);
			}
			m_nWordNumber ++ ;
			break;
		}
	}
	delete pWord;
}

BOOL CDictionary::LookupWord(BSTR word, BSTR *resultWord)
{
	char *pWord = _com_util::ConvertBSTRToString((BSTR)word);
	char *pUpperWord = strupr(pWord);
	for (int i = 0; i < m_nWordNumber; i++)
	{
		char *tmpWord = strupr(m_pData[i].wordForLang1);
		if (strcmp(tmpWord, pWord) == 0) {
			*resultWord = (BSTR)(_com_util::ConvertStringToBSTR(m_pData[i].wordForLang2));
			delete pWord;
			return TRUE;
		}
	}
	*resultWord = NULL;
	delete pWord;
	return FALSE;
}

BOOL CDictionary::RestoreLibrary(BSTR filename)
{
	char *pFileName = _com_util::ConvertBSTRToString((BSTR)filename);
	FILE *fp;
	if( (fp = fopen( pFileName, "wt" )) == NULL ) {
		printf("Open dictionary file : %s failed.\n", pFileName);
		delete pFileName;
		return FALSE;
	}
	char LineBuffer[128];
	sprintf(LineBuffer, "%d\n", m_nWordNumber);
	if (fputs(LineBuffer, fp) == EOF) {
		printf("Write TotalNumber failed!\n");
		fclose(fp);
		delete pFileName;
		return FALSE; 
	}
	
	for(int i = 0; i < m_nWordNumber; i ++ ) {
		if (fputs(m_pData[i].wordForLang1, fp) == EOF) {
			printf("Write the first string failed!\n");
			fclose(fp);
			delete pFileName;
			return FALSE; 
		}
		fputs("\n", fp);
		if (fputs(m_pData[i].wordForLang2, fp) == EOF) {
			printf("Write the first string failed!\n");
			fclose(fp);
			delete pFileName;
			return FALSE; 
		}
		fputs("\n", fp);
	}
	
	fclose(fp);
	delete pFileName;
	return TRUE;
}

void CDictionary::FreeLibrary()
{
	Initialize();
}

	
BOOL CDictionary::CheckWord (BSTR word, BSTR *resultWord)
{
	char *pWord = _com_util::ConvertBSTRToString((BSTR)word);
	char *pUpperWord = strupr(pWord);
	char *pMinMaxWord, *pMaxMinWord;
	int	 nMinIndex = -1, nMaxIndex = -1;
	pMinMaxWord = pMaxMinWord = NULL;
	for (int i = 0; i < m_nWordNumber; i++)
	{
		char *tmpWord = strupr(m_pData[i].wordForLang1);
		if (strcmp(tmpWord, pWord) == 0) {
			delete pWord;
			return TRUE;
		} else if (strcmp(tmpWord, pWord) < 0) {
			if ((pMinMaxWord == NULL) || (strcmp(tmpWord, pMinMaxWord) > 0)) 
			{
				pMinMaxWord = tmpWord;
				nMinIndex = i;
			}
		} else {
			if ((pMaxMinWord == NULL) || (strcmp(tmpWord, pMaxMinWord) < 0)) 
			{
				pMaxMinWord = tmpWord;
				nMaxIndex = i;
			}
		}
	}
	
	*resultWord = NULL;
	if (nMinIndex != -1)
		*resultWord = (BSTR)_com_util::ConvertStringToBSTR(m_pData[nMinIndex].wordForLang1);
	else if (nMaxIndex != -1)
		*resultWord = (BSTR)_com_util::ConvertStringToBSTR(m_pData[nMaxIndex].wordForLang1);
	delete pWord;
	return FALSE;
}