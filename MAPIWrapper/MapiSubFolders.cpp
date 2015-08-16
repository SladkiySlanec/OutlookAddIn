// MapiSubFolders.cpp : Implementation of CMapiSubFolders

#include "stdafx.h"
#include "MapiSubFolders.h"
#include "atlsafe.h"
#include ".\Utils\StringHelper.hpp"
// CMapiSubFolders

void CMapiSubFolders::Init(IMAPITable * pTable)
{
	m_ptrTable = pTable;
}
void CMapiSubFolders::SetSeekRowAndColumns()
{
	if(m_isSetSeekRowAndColumns)
		return;
	HRESULT hr = E_FAIL;
	//THROW_MAPI_ERR(pRoot, hr);
	hr = m_ptrTable->SeekRow(BOOKMARK_BEGINNING,0,NULL);
	check_mapi_err(hr);

	SizedSPropTagArray(2, sptCols) = { 2, { PR_ENTRYID, PR_DISPLAY_NAME } };
	SPropTagArray * pTagArray = new SPropTagArray();
	pTagArray->cValues = 2;
	ULONG* pUl = new ULONG[2];


    hr = m_ptrTable->SetColumns((LPSPropTagArray)&sptCols , 0);
	m_isSetSeekRowAndColumns = true;
}

LPSAFEARRAY CreateSafeArray(VARTYPE vtType, ULONG ulSize, UINT uDims = 1, LONG lBound = 0)
{
	SAFEARRAYBOUND rgsabound;
    rgsabound.cElements = ulSize;
    rgsabound.lLbound = lBound;

	return ::SafeArrayCreate(vtType, uDims, &rgsabound);
}


STDMETHODIMP CMapiSubFolders::GetNextItemPro(VARIANT ppsaMyArray, ULONG* size, VARIANT *buffer)
{
	//SetSeekRowAndColumns();
	HRESULT hr = S_OK;
	VARTYPE varType = ppsaMyArray.vt;
	V_ARRAY(&ppsaMyArray);

	LONG lbound = 0;
	hr = SafeArrayGetLBound(ppsaMyArray.parray, 1, &lbound);

	LONG ubound = 0;
	hr = SafeArrayGetUBound(ppsaMyArray.parray, 1, &ubound);
	ULONG cRows = ubound - lbound + 1; 

	ULONG* pUlong = NULL;
	hr = SafeArrayAccessData(ppsaMyArray.parray, (void HUGEP* FAR*)&pUlong);
	ULONG* pTagsToShow = new ULONG[cRows+1];
	pTagsToShow[0] = cRows;
	memcpy_s(&pTagsToShow[1], cRows * sizeof(ULONG), pUlong, cRows * sizeof(ULONG));

	SafeArrayUnaccessData (ppsaMyArray.parray);

	try
	{
		hr = m_ptrTable->SeekRow(BOOKMARK_BEGINNING,0,NULL);
		check_mapi_err(hr);

		hr = m_ptrTable->SetColumns((LPSPropTagArray)pTagsToShow , 0);


		SmartSRowSet m_pRows;
		hr = m_ptrTable->QueryRows(255, 0, &m_pRows);

		CComBSTR DisplayName = NULL;
		
		CComVariant folders = CreateSafeArray( VT_VARIANT, m_pRows->cRows);

		for(LONG i = 0; i<(LONG)m_pRows->cRows; i++ )
		{
			//CComVariant currentFolder = CreateSafeArray( VT_VARIANT, 1);
			CComVariant setOfFolderProps = CreateSafeArray( VT_VARIANT, 2);
			
			CComVariant propTag = CreateSafeArray( VT_I4, m_pRows->aRow[i].cValues);
			CComVariant propValues = CreateSafeArray( VT_VARIANT, m_pRows->aRow[i].cValues);

			for(LONG j = 0; j<(LONG)m_pRows->aRow[i].cValues; j++)
			{

				::SafeArrayPutElement(propTag.parray, &j, &m_pRows->aRow[i].lpProps[j].ulPropTag);
				CComVariant temp;
				switch (PROP_TYPE(m_pRows->aRow[i].lpProps[j].ulPropTag))
				{
				case PT_STRING8:
					temp.vt = VT_BSTR;
					temp.bstrVal = ::SysAllocString(m_pRows->aRow[i].lpProps[j].Value.lpszW);
					break;
				case PT_UNICODE:
					temp.vt = VT_BSTR;
					temp.bstrVal = ::SysAllocString(m_pRows->aRow[i].lpProps[j].Value.lpszW);
					break;
				case PT_LONG:
					temp.vt = VT_I4;
					temp.lVal = m_pRows->aRow[i].lpProps[j].Value.l;
					break;
				case PT_BOOLEAN:
					temp.vt = VT_BOOL;
					temp.boolVal = m_pRows->aRow[i].lpProps[j].Value.b ? VARIANT_TRUE:VARIANT_FALSE ;
					break;
				case PT_BINARY:
					temp.vt = VT_BSTR;
					temp.bstrVal = StringHelper::SBinaryToHexString(&m_pRows->aRow[i].lpProps[j].Value.bin);
					break;
				}
				::SafeArrayPutElement(propValues.parray, &j, &temp);
			}

			LONG index = 0;
			::SafeArrayPutElement(setOfFolderProps.parray, &index, &propTag);
			index++;
			::SafeArrayPutElement(setOfFolderProps.parray, &index, &propValues);
			//::SafeArrayPutElement(currentFolder.parray, 0, &setOfFolderProps);
			::SafeArrayPutElement(folders.parray, &i, &setOfFolderProps);
		}
		 folders.Detach(buffer);
	}
	catch(...)
	{
		m_pRows = NULL;
	}
	

	return hr;
}


STDMETHODIMP CMapiSubFolders::GetNextItem(BSTR* DisplayName)
{
	SetSeekRowAndColumns();

	HRESULT hr = E_FAIL;
	try
	{
		if(m_pRows == NULL || m_counter == 255)
		{
			hr = m_ptrTable->QueryRows(255, 0, &m_pRows);
			m_counter = 0;
		}

		if(m_pRows->cRows == 0)
			return S_OK;

		for(int j = 0; j<2; j++)
		{
			if(m_pRows->aRow[m_counter].lpProps[j].ulPropTag == PR_DISPLAY_NAME)
			{
				*DisplayName = ::SysAllocString(m_pRows->aRow[m_counter].lpProps[j].Value.lpszW);
				m_counter++;
				return S_OK;
			}
		}
	}
	catch(...)
	{
		m_pRows = NULL;
	}
	

	return hr;
}


STDMETHODIMP CMapiSubFolders::get_Count(ULONG* pVal)
{
	return m_ptrTable->GetRowCount(0, pVal);
}
