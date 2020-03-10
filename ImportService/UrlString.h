/*****************************************************************************
Module :     UrlString.h
Notices:     Written 2006 by Stephane Erhardt
Description: H URL Encoder/Decoder
*****************************************************************************/
#pragma once

class CUrlString
{
private:
	CString m_csUnsafe;

public:
	CUrlString();
	virtual ~CUrlString() { };
	CString Encode(CString csDecoded);
	CString Decode(CString csEncoded);
};


