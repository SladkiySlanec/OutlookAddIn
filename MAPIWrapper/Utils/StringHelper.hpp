#pragma once
#include "windef.h"
#include "MAPIDefS.h"
#include "stdio.h"
#include "..\Errors\checks.h"
#include "..\SmartPointers\SmartPtr.hpp"
class StringHelper
{
	StringHelper(void){}
	~StringHelper(void){}
public:
	static SBinary HexStringToSBinary(LPWSTR lpHexString)
	{
		check_arg(lpHexString != NULL);

		WCHAR *pos = lpHexString;
		SBinary bin={0};
		bin.cb = wcslen(lpHexString)/2;
		bin.lpb = new byte[bin.cb];

		for(ULONG count = 0; count < bin.cb; count++) 
		{
			swscanf_s(pos, L"%2hhx", &bin.lpb[count]);
			pos += 2;
		}
		return bin;
	}

	 
	static LPWSTR SBinaryToHexString(LPSBinary pBinary)
	{
		ULONG    cb=pBinary->cb;
		BSTR    strReturn=pBinary ? SysAllocStringLen(NULL, pBinary->cb * 2): NULL;
		LPBYTE    pb=pBinary->lpb;
		LPWSTR    psz=strReturn;

		if(psz)
		{
			while(cb--)
			{
				int n1 =*pb & 0x0f;
				int n2 =(*pb >> 4)& 0x0f;

				*psz++ =(BYTE)(n2 +(n2 >= 10 ?(L'A' - 10): L'0'));
				*psz++ =(BYTE)(n1 +(n1 >= 10 ?(L'A' - 10): L'0'));
				++pb;
			}
			*psz=0;
		}
		return strReturn;
		//check_arg(pBinary != NULL);
		//check_arg(pBinary->lpb != NULL);

		//LPSTR pBuff = new char[pBinary->cb*2+1];
		//SmartPtr<char> ptr(pBuff);
		//char* pos = pBuff;
		//for (ULONG i = 0; i < pBinary->cb; i++)
		//{
		//	pos += sprintf_s(pos, pBinary->cb*2-i, "%02X", pBinary->lpb[i]);
		//}

		//return CComBSTR(pBuff).Detach();
	}
};