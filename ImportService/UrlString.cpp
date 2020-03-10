/*****************************************************************************
Module :     UrlString.cpp
Notices:     Written 2006 by Stephane Erhardt
Description: CPP URL Encoder/Decoder
*****************************************************************************/
#include "stdafx.h"
#include "UrlString.h"

/*****************************************************************************/
CUrlString::CUrlString()
{
	m_csUnsafe = _T("%=\"<>\\^[]`+$,@:;/!#?&'");
	for(int iChar = 1; iChar < 33; iChar++)
		m_csUnsafe += (char)iChar;
	for(int iChar = 124; iChar < 256; iChar++)
		m_csUnsafe += (char)iChar;
}

/*****************************************************************************/
CString CUrlString::Encode(CString csDecoded)
{
	CString csCharEncoded, csCharDecoded;

	CString csEncoded = csDecoded;

	for(int iPos = 0; iPos < m_csUnsafe.GetLength(); iPos++)
	{
		csCharEncoded.Format(_T("%%%02X"), m_csUnsafe[iPos]);
		csCharDecoded = m_csUnsafe[iPos];
		csEncoded.Replace(csCharDecoded, csCharEncoded);
	}
	return csEncoded;
}

/*****************************************************************************/
CString CUrlString::Decode(CString csEncoded)
{
	CString csUnsafeEncoded = Encode(m_csUnsafe);
	CString csDecoded = csEncoded;
	CString csCharEncoded, csCharDecoded;

	for(int iPos = 0; iPos < csUnsafeEncoded.GetLength(); iPos += 3)
	{
		csCharEncoded = csUnsafeEncoded.Mid(iPos, 3);
		csCharDecoded = (TCHAR)_tcstol(csUnsafeEncoded.Mid(iPos + 1, 2), NULL, 16);
		csDecoded.Replace(csCharEncoded, csCharDecoded);
	}
	return csDecoded;
}