#include "StdAfx.h"
#include <direct.h>
#include <winnetwk.h>
#include <wininet.h>
#include "UrlString.h"
#include "Defs.h"
#include "Utils.h"
#include "Prefs.h"
#include "Ping.h"
#include "Resource.h"
#include "EmailClient.h"

#include <boost/regex.hpp>
#include <boost/regex/v4/fileiter.hpp>

extern CPrefs g_prefs;


BOOL CALLBACK EnableAllChildren(HWND hwnd, LPARAM lparam)
{
	::EnableWindow(hwnd, (BOOL)lparam);
	return TRUE;
}


void EnableAllControls(HWND hDlg, BOOL bEnable)
{
	::EnumChildWindows(hDlg, EnableAllChildren, (LPARAM)bEnable);
}

CUtils::CUtils(void)
{
}


CUtils::~CUtils(void)
{
}

CString CUtils::GetComputerName()
{
	TCHAR buf[MAX_PATH];
	DWORD nBufLen = MAX_PATH;
	::GetComputerName(buf, &nBufLen);

	CString s(buf);

	return s;
}

CString CUtils::GetLastWin32Error()
{
	TCHAR szBuf[4096];
	LPVOID lpMsgBuf;
	DWORD dw = GetLastError();

	FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM,
		NULL,
		dw,
		MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
		(LPTSTR)&lpMsgBuf,
		0, NULL);

	wsprintf(szBuf, _T("%s (errorcode %d)"), lpMsgBuf, dw);

	CString s(szBuf);
	LocalFree(lpMsgBuf);

	return s;
}

CString CUtils::LoadResourceString(int nID)
{
	CString s = _T("");
	s.LoadString(nID);

	return s;
}

CString CUtils::Int2String(int n)
{
	CString s;
	s.Format(_T("%d"), n);
	return s;
}

CString CUtils::Bigint2String(__int64 n)
{
	CString s;
	s.Format(_T("%I64d"), n);
	return s;
}

CString CUtils::Double2String(double f)
{
	CString s;
	s.Format(_T("%.4f"), f);
	return s;
}

CString CUtils::Double2String(double f, int decimals)
{
	CString s;
	if (decimals == 1)
		s.Format(_T("%.1f"), f);
	else if (decimals == 2)
		s.Format(_T("%.2f"), f);
	else if (decimals == 3)
		s.Format(_T("%.3f"), f);
	else
		s.Format(_T("%.4f"), f);
	return s;
}

std::wstring  CUtils::Double2StringW(double f, int decimals)
{
	CString s = Double2String(f, decimals);
	CStringW ws(s);
	std::wstring wws(ws);
	return wws;
}


void CUtils::StripTrailingSpaces(TCHAR *szString, TCHAR *szStringDest)
{
	TCHAR *p = szString;
	TCHAR *pdest = szStringDest;

	BOOL	bProcessed = FALSE;

	while (*p != 0) {
		BOOL hasCR = FALSE;
		BOOL hasLF = FALSE;

		TCHAR *pline = p;
		TCHAR *peoline = p;
		while (*pline != 0 && *pline != 0x0D && *pline != 0x0A)
			pline++;
		if (*pline == 0)		// end of buffer
			break;

		if (*pline == 0x0D)					// CR+LF
			hasCR = TRUE;

		if (*pline == 0x0A)					// LF
			hasLF = TRUE;

		peoline = pline;

		pline--;
		while (*pline == 0x20 && pline > p)
			pline--;

		if (peoline != pline + 1)
			bProcessed = TRUE;

		//	if (pline == p)			// all spaces on this line..?
		//		break;

			// pline ptr is now at the last non-space in line
			// copy contents from srat of line until pline ptr.
		while (p <= pline) {
			*pdest++ = *p++;
		}

		if (hasCR)
			*pdest++ = 0x0D;

		*pdest++ = 0x0A;


		p = peoline;

		while (*p != 0 && (*p == 0x0D || *p == 0x0A))
			p++;

	}

	*pdest = 0;
}

int CUtils::GetFileAge(CString sFileName)
{
	if (FileExist(sFileName) == FALSE)
		return 100000;
	CTime t1 = GetWriteTime(sFileName);
	CTime t2 = CTime::GetCurrentTime();
	CTimeSpan ts = t2 - t1;
	if (ts.GetTotalHours() > 0)
		return (int)ts.GetTotalHours();

	return 100000;
}


int CUtils::GetFileAgeHours(CString sFileName)
{
	if (FileExist(sFileName) == FALSE)
		return 100000;
	CTime t1 = GetWriteTime(sFileName);
	CTime t2 = CTime::GetCurrentTime();
	CTimeSpan ts = t2 - t1;
	if (ts.GetTotalHours() > 0)
		return (int)ts.GetTotalHours();

	return -1;
}

int CUtils::GetFileAgeMinutes(CString sFileName)
{
	if (FileExist(sFileName) == FALSE)
		return -1;
	CTime t1 = GetWriteTime(sFileName);
	CTime t2 = CTime::GetCurrentTime();
	CTimeSpan ts = t2 - t1;
	if (ts.GetTotalMinutes() > 0)
		return (int)ts.GetTotalMinutes();

	return -1;
}

BOOL CUtils::RemoveReadOnlyAttribute(CString sFullPath)
{
	DWORD dwAttrs = ::GetFileAttributes(sFullPath);
	if (dwAttrs != INVALID_FILE_ATTRIBUTES) {
		if ((dwAttrs & FILE_ATTRIBUTE_READONLY)) {
			::SetFileAttributes(sFullPath, dwAttrs & ~FILE_ATTRIBUTE_READONLY);
		}
	}

	// Test if succedded
	dwAttrs = ::GetFileAttributes(sFullPath);
	return (dwAttrs & FILE_ATTRIBUTE_READONLY) == 0;
}

BOOL CUtils::LockCheck(CString sFileToLock, BOOL bSimpleCheck)
{
	// Appempt to lock and read from  file
	OVERLAPPED		Overlapped;
	DWORD			BytesRead, dwSizeHigh;
	char			readbuffer[4097];
	HANDLE Hndl = ::CreateFile(sFileToLock, GENERIC_READ | GENERIC_WRITE, 0, NULL, OPEN_EXISTING, 0, 0);

	if (Hndl == INVALID_HANDLE_VALUE)
		return FALSE;

	if (bSimpleCheck == FALSE) {
		DWORD dwSizeLow = ::GetFileSize(Hndl, &dwSizeHigh);
		if (dwSizeLow == -1) {
			::CloseHandle(Hndl);
			return FALSE;
		}

		if (!::LockFileEx(Hndl, LOCKFILE_FAIL_IMMEDIATELY | LOCKFILE_EXCLUSIVE_LOCK, 0,
			dwSizeLow, dwSizeHigh, &Overlapped)) {
			::CloseHandle(Hndl);
			return FALSE;
		}

		::UnlockFileEx(Hndl, 0, dwSizeLow, dwSizeHigh, &Overlapped);
	}

	if (::ReadFile(Hndl, readbuffer, 4096, &BytesRead, NULL)) {
		::CloseHandle(Hndl);
		return TRUE;
	}

	::CloseHandle(Hndl);
	return FALSE;
}

BOOL CUtils::SoftLockCheck(CString sFileToLock)
{
	// Appempt to lock and read from  file
	DWORD			BytesRead;
	TCHAR			readbuffer[4097];

	HANDLE Hndl = ::CreateFile(sFileToLock, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, 0, 0);

	if (Hndl == INVALID_HANDLE_VALUE)
		return FALSE;

	if (ReadFile(Hndl, readbuffer, 4096, &BytesRead, NULL)) {
		CloseHandle(Hndl);
		return TRUE;
	}

	CloseHandle(Hndl);
	return FALSE;
}

BOOL CUtils::StableTimeCheck(CString sFileToTest, int nStableTimeSec)
{
	HANDLE	Hndl;
	FILETIME WriteTime, LastWriteTime;
	DWORD dwFileSize = 0, dwLastFileSize = 0, dwFileSizeHigh = 0, dwLastFileSizeHigh = 0;

	if (nStableTimeSec == 0)
		return TRUE;

	Hndl = ::CreateFile(sFileToTest, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, 0, 0);
	if (Hndl == INVALID_HANDLE_VALUE)
		return FALSE;

	::GetFileTime(Hndl, NULL, NULL, &WriteTime);
	dwFileSize = ::GetFileSize(Hndl, &dwFileSizeHigh);

	::CloseHandle(Hndl);

	::Sleep(nStableTimeSec * 1000);

	Hndl = CreateFile(sFileToTest, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, 0, 0);
	if (Hndl == INVALID_HANDLE_VALUE)
		return FALSE;

	GetFileTime(Hndl, NULL, NULL, &LastWriteTime);
	dwLastFileSize = ::GetFileSize(Hndl, &dwLastFileSizeHigh);

	CloseHandle(Hndl);

	return (WriteTime.dwLowDateTime == LastWriteTime.dwLowDateTime && WriteTime.dwHighDateTime == LastWriteTime.dwHighDateTime && dwFileSize == dwLastFileSize);
}

DWORD CUtils::GetFileSize(CString sFile)
{
	DWORD	dwFileSizeLow = 0, dwFileSizeHigh = 0;

	HANDLE	Hndl = ::CreateFile(sFile, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, 0, 0);
	if (Hndl == INVALID_HANDLE_VALUE)
		return 0;

	dwFileSizeLow = ::GetFileSize(Hndl, &dwFileSizeHigh);
	::CloseHandle(Hndl);

	return dwFileSizeLow;
}


CTime CUtils::GetWriteTime(CString sFile)
{
	FILETIME WriteTime, LocalWriteTime;
	SYSTEMTIME SysTime;
	CTime t(1975, 1, 1, 0, 0, 0);
	BOOL ok = TRUE;

	HANDLE Hndl = ::CreateFile(sFile, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, 0, 0);
	if (Hndl == INVALID_HANDLE_VALUE)
		return t;

	ok = ::GetFileTime(Hndl, NULL, NULL, &WriteTime);

	CloseHandle(Hndl);
	if (ok) {
		::FileTimeToLocalFileTime(&WriteTime, &LocalWriteTime);
		::FileTimeToSystemTime(&LocalWriteTime, &SysTime);
		CTime t2((int)SysTime.wYear, (int)SysTime.wMonth, (int)SysTime.wDay, (int)SysTime.wHour, (int)SysTime.wMinute, (int)SysTime.wSecond);
		t = t2;
	}

	return t;
}

BOOL CUtils::FileExist(CString sFile)
{
	HANDLE hFind;
	WIN32_FIND_DATA  ff32;
	BOOL ret = FALSE;

	hFind = ::FindFirstFile(sFile, &ff32);
	if (hFind != INVALID_HANDLE_VALUE) {
		if (!(ff32.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY))
			ret = TRUE;
	}

	FindClose(hFind);
	return ret;
}


BOOL CUtils::DirectoryExist(CString sDir)
{
	HANDLE hFind;
	WIN32_FIND_DATA  ff32;
	BOOL ret = FALSE;

	hFind = ::FindFirstFile(sDir, &ff32);
	if (hFind != INVALID_HANDLE_VALUE) {
		if (ff32.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
			ret = TRUE;
	}

	FindClose(hFind);
	return ret;
}


BOOL CUtils::CheckFolder(CString sFolder)
{
	TCHAR	szCurrentDir[_MAX_PATH];
	BOOL	ret = TRUE;

	::GetCurrentDirectory(_MAX_PATH, szCurrentDir);

	if ((_tchdir(sFolder.GetBuffer(255))) == -1)
		ret = FALSE;

	sFolder.ReleaseBuffer();

	::SetCurrentDirectory(szCurrentDir);

	return ret;
}

BOOL CUtils::CheckFolderWithPing(CString sFolder)
{
	TCHAR	szCurrentDir[_MAX_PATH];
	BOOL	ret = TRUE;
	CString sServerName = _T("");

	if (g_prefs.m_bypassping)
		return TRUE;

	// Resolve mapped drive connection
	if (sFolder.Mid(1, 1) == _T(":")) {
		TCHAR szRemoteName[MAX_PATH];
		DWORD lpnLength = MAX_PATH;

		if (::WNetGetConnection(sFolder.Mid(0, 2), szRemoteName, &lpnLength) == NO_ERROR)
			sFolder = szRemoteName;
	}


	if (sFolder.Mid(0, 2) == _T("\\\\"))
		sServerName = sFolder.Mid(2);

	int n = sServerName.Find(_T("\\"));
	if (n != -1)
		sServerName = sServerName.Left(n);

	BOOL 	   bSuccess = FALSE;
	if (sServerName != _T("")) {
		WSADATA wsa;
		if (WSAStartup(MAKEWORD(2, 0), &wsa) != 0)
			return FALSE;

		CPing p;
		CPingReply pr;
		bSuccess = p.PingUsingWinsock(sServerName, pr, 30, 1000, 32, 0, FALSE);
		WSACleanup();
	}

	if (sServerName != "" && bSuccess == FALSE)
		return FALSE;

	::GetCurrentDirectory(_MAX_PATH, szCurrentDir);

	if ((_tchdir(sFolder.GetBuffer(255))) == -1)
		ret = FALSE;

	sFolder.ReleaseBuffer();

	::SetCurrentDirectory(szCurrentDir);

	return ret;
}

BOOL CUtils::Reconnect(CString sFolder, CString sUser, CString sPW)
{
	if (g_prefs.m_bypassreconnect == BYPASSPING_ALWAYS)
		return TRUE;

	if (sUser == "" && g_prefs.m_bypassreconnect == BYPASSPING_CURRENTUSER)
		return TRUE;

	if (CheckFolderWithPing(sFolder) == FALSE)
		return FALSE;
	NETRESOURCE nr;
	nr.dwScope = NULL;
	nr.dwDisplayType = RESOURCEDISPLAYTYPE_GENERIC;
	nr.dwUsage = RESOURCEDISPLAYTYPE_GENERIC;  //RESOURCEUSAGE_CONNECTABLE 
	nr.dwType = RESOURCETYPE_DISK;
	nr.lpLocalName = NULL;
	nr.lpComment = NULL;
	nr.lpProvider = NULL;
	nr.lpRemoteName = sFolder.GetBuffer(260);
	DWORD dwResult = WNetAddConnection2(&nr, // NETRESOURCE from enumeration 
		sUser != _T("") ? sPW.GetBuffer(200) : NULL,                  // no password 
		sUser != _T("") ? (LPCTSTR)sUser.GetBuffer(200) : NULL,                  // logged-in user 
		CONNECT_UPDATE_PROFILE);       // update profile with connect information 
	sFolder.ReleaseBuffer();
	sUser.ReleaseBuffer();
	sPW.ReleaseBuffer();

	// Try alternative connection method...
	if (dwResult != NO_ERROR && sUser != _T("")) {
		if (sPW == _T(""))
			sPW = _T("\"\"");
		CString commandLine = _T("NET USE \\\\") + sFolder + _T(" ") + sPW + _T(" /USER:") + sUser;
		return	DoCreateProcess(commandLine, 300, TRUE);
	}

	return CheckFolderWithPing(sFolder);
}

CString CUtils::GetExtension(CString sFileName)
{
	int n = sFileName.ReverseFind(_T('.'));
	if (n == -1)
		return _T("");
	return sFileName.Mid(n + 1);
}

CString CUtils::GetFileName(CString sFullName)
{
	return GetFileName(sFullName, FALSE);
}

CString CUtils::GetFileName(CString sFullName, BOOL bExcludeExtension)
{
	if (bExcludeExtension) {
		int m = sFullName.ReverseFind(_T('.'));
		if (m != -1)
			sFullName = sFullName.Left(m);
	}
	int n = sFullName.ReverseFind(_T('\\'));
	if (n == -1)
		return sFullName;
	return sFullName.Mid(n + 1);
}

CString CUtils::GetChangedExtension(CString sFullName, CString sNewExtension)
{
	return GetFilePath(sFullName) + _T("\\") + GetFileName(sFullName, TRUE) + sNewExtension;
}


CString CUtils::GetFilePath(CString sFullName)
{
	int n = sFullName.ReverseFind(_T('\\'));
	if (n == -1)
		return sFullName;
	return sFullName.Left(n);
}

CString CUtils::GetModuleLoadPath()
{
	TCHAR ModName[_MAX_PATH];

	if (!::GetModuleFileName(NULL, ModName, sizeof(ModName)))
		return CString(_T(""));


	CString LoadPath(ModName);

	int Idx = LoadPath.ReverseFind(_T('\\'));

	if (Idx == -1)
		return CString(_T(""));

	LoadPath = LoadPath.Mid(0, Idx);

	return LoadPath;
}

CString  CUtils::GetFirstFile(CString sSpoolFolder, CString sSearchMask)
{
	WIN32_FIND_DATA	fdata;
	HANDLE			fHandle;
	TCHAR			szNameSearch[MAX_PATH];

	CString sFoundFile = _T("");
	if (CheckFolder(sSpoolFolder) == FALSE)
		return _T("");

	if (sSpoolFolder.Right(1) != "\\")
		sSpoolFolder += _T("\\");

	_stprintf(szNameSearch, _T("%s%s"), (LPCTSTR)sSpoolFolder, (LPCTSTR)sSearchMask);
	fHandle = ::FindFirstFile(szNameSearch, &fdata);
	if (fHandle == INVALID_HANDLE_VALUE) {
		// All empty
		return _T("");
	}

	do {
		if (!(fdata.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)) {
			sFoundFile.Format(_T("%s\\%s"), (LPCTSTR)sSpoolFolder, fdata.cFileName);
			break;
		}
		if (!::FindNextFile(fHandle, &fdata))
			break;
	} while (TRUE);

	::FindClose(fHandle);

	return sFoundFile;
}

int CUtils::StringSplitter(CString sInput, CString sSeparators, CStringArray &sArr)
{
	sArr.RemoveAll();

	int nElements = 0;
	CString s = sInput;
	s.Trim();

	int v = s.FindOneOf(sSeparators);
	BOOL isLast = FALSE;
	do {
		nElements++;
		if (v != -1)
			sArr.Add(s.Left(v));
		else {
			sArr.Add(s);
			isLast = TRUE;
		}
		if (isLast == FALSE) {
			s = s.Mid(v + 1);
			v = s.FindOneOf(sSeparators);
		}
	} while (isLast == FALSE);

	return nElements;
}

BOOL CUtils::IsInArray(CUIntArray *nArr, int nValueToFind)
{
	for (int i = 0; i < nArr->GetCount(); i++) {
		if (nArr->GetAt(i) == nValueToFind) {
			return TRUE;
		}
	}

	return FALSE;
}

BOOL CUtils::IsInArray(CStringArray *sArr, CString sValueToFind)
{
	for (int i = 0; i < sArr->GetCount(); i++) {
		if (sValueToFind.CompareNoCase(sArr->GetAt(i)) == 0) {
			return TRUE;
		}
	}

	return FALSE;
}


CString CUtils::IntArrayToString(int nArr[], int nMaxItems)
{
	CString s, t;

	for (int i = 0; i < nMaxItems; i++) {
		t.Format(_T("%d"), nArr[i]);
		s += t;
		if (i < nMaxItems - 1)
			s += _T(",");
	}
	return s;
}

CString CUtils::FloatArrayToString(double nArr[], int nMaxItems)
{
	CString s, t;

	for (int i = 0; i < nMaxItems; i++) {
		t.Format(_T("%.6f"), nArr[i]);
		s += t;
		if (i < nMaxItems - 1)
			s += _T(",");
	}
	return s;
}

CString CUtils::StringArrayToString(CString nArr[], int nMaxItems)
{
	CString s;

	for (int i = 0; i < nMaxItems; i++) {
		s += nArr[i];
		if (i < nMaxItems - 1)
			s += _T(",");
	}
	return s;
}

CString CUtils::GrepFile(CString sFileName, TCHAR *szSearchString)
{
	HANDLE	Hndl;
	TCHAR	infbuf[8092], *pinfbuf, mask[40];
	DWORD	nBytesRead;
	CString sRet = _T("");

	if ((Hndl = ::CreateFile(sFileName, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, 0, 0)) == INVALID_HANDLE_VALUE)
		return sRet;

	if (!::ReadFile(Hndl, infbuf, 8092, &nBytesRead, NULL)) {
		::CloseHandle(Hndl);
		return sRet;
	}

	::CloseHandle(Hndl);
	infbuf[nBytesRead == 8092 ? 8091 : nBytesRead] = 0;

	//	ToUpper(infbuf);
	_tcscpy(mask, szSearchString);
	//	ToUpper(mask);

	if ((pinfbuf = _tcsstr(infbuf, mask)) == NULL)
		return sRet;

	// Search string found - isolate

	TCHAR *pp = pinfbuf + _tcslen(mask);

	TCHAR *p = pp;

	while (*p && *p != 13 && *p != 10)	// scan to eol
		p++;

	if (!*p && p == pp)
		return sRet;
	--p;
	while (*p == _T(' '))					// ignore trailing spaces
		--p;
	++p;
	*p = 0;

	sRet = pp;

	return sRet;
}


BOOL CUtils::TryMatchExpression(CString sMatchExpression, CString sInputString, BOOL bPartialMatch)
{
	boost::regex	e;
	unsigned int	match_flag = bPartialMatch ? boost::match_partial : boost::match_default;
	BOOL			ret;

#ifdef UNICODE
	std::wstring wMatch(sMatchExpression);
	const std::string match(wMatch.begin(), wMatch.end());
	std::wstring wInput(sInputString);
	const std::string input(wInput.begin(), wInput.end());
#else
	TCHAR szMatch[1024], szInput[1024];
	_tcscpy(szMatch, sMatchExpression);
	_tcscpy(szInput, sInputString);
	std::string match = szMatch;
	std::string input = szInput;
#endif

	try {

		e.assign(match);
		ret = boost::regex_match(input, e, bPartialMatch ? boost::match_partial | boost::match_default : boost::match_default);
	}
	catch (std::exception& e)
	{
#ifdef UNICODE
		Logprintf(_T("ERROR in regular expression"));
#else
		Logprintf(_T("ERROR in regular expression %s"), e.what());
#endif
		return FALSE;
	}

	return ret;
}

BOOL CUtils::TryMatchExpressionEx(CString sMatchExpression, CString sInputString, BOOL bPartialMatch, TCHAR *szErrorMessage)
{
	_tcscpy(szErrorMessage, _T(""));

	boost::regex	e;
	unsigned int	match_flag = bPartialMatch ? boost::match_partial : boost::match_default;
	BOOL			ret;

#ifdef UNICODE
	std::wstring wMatch(sMatchExpression);
	const std::string match(wMatch.begin(), wMatch.end());
	std::wstring wInput(sInputString);
	const std::string input(wInput.begin(), wInput.end());
#else
	TCHAR szMatch[1024], szInput[1024];
	_tcscpy(szMatch, sMatchExpression);
	_tcscpy(szInput, sInputString);
	std::string match = szMatch;
	std::string input = szInput;
#endif


	try {
		e.assign(match);
	}
	catch (std::exception& e)
	{
		_tcscpy(szErrorMessage, _T("ERROR in regular expression "));
#ifndef UNICODE
		_tcscat(szErrorMessage, e.what());
#endif
		Logprintf(_T("%s"), szErrorMessage);
		return FALSE;
	}
	try {
		ret = boost::regex_match(input, e, bPartialMatch ? boost::match_partial | boost::match_default : boost::match_default);
	}
	catch (std::exception& e)
	{
		_tcscpy(szErrorMessage, _T("ERROR in regular expression "));
#ifndef UNICODE
		_tcscat(szErrorMessage, e.what());
#endif
		Logprintf(_T("%s"), szErrorMessage);
		return FALSE;
	}

	return ret;
}


CString CUtils::FormatExpressionEx(CString sMatchExpression, CString sFormatExpression, CString sInputString, BOOL bPartialMatch, TCHAR *szErrorMessage, BOOL *bMatched)
{
	_tcscpy(szErrorMessage, _T(""));
	*bMatched = TRUE;
	boost::regex	e;
	unsigned int				match_flag = bPartialMatch ? boost::match_partial : boost::match_default;
	unsigned int				flags = match_flag | boost::format_perl;
	CString			retSt = sInputString;


#ifdef UNICODE
	std::wstring wMatch(sMatchExpression);
	const std::string match(wMatch.begin(), wMatch.end());
	std::wstring wInput(sInputString);
	const std::string input(wInput.begin(), wInput.end());
	std::wstring wFormat(sFormatExpression);
	const std::string format(wFormat.begin(), wFormat.end());
#else
	TCHAR szMatch[1024], szInput[1024], szFormat[1024];
	_tcscpy(szMatch, sMatchExpression);
	_tcscpy(szInput, sInputString);
	_tcscpy(szFormat, sFormatExpression);
	std::string match = szMatch;
	std::string input = szInput;
	std::string format = szFormat;
#endif

	try {
		e.assign(match);
	}
	catch (std::exception& e)
	{
		_tcscpy(szErrorMessage, _T("ERROR in regular expression "));
#ifndef UNICODE
		_tcscat(szErrorMessage, e.what());
#endif
		*bMatched = FALSE;
		return retSt;
	}

	try {
		BOOL ret = boost::regex_match(input, e, bPartialMatch ? boost::match_partial | boost::match_default : boost::match_default);
		*bMatched = ret;
		if (ret == TRUE) {

			std::string r = boost::regex_merge(input, e, format, boost::format_perl | (bPartialMatch ? boost::match_partial | boost::match_default : boost::match_default));

#ifdef UNICODE
			CStringA retStA = r.data();
			CString wretSt(retStA);
			retSt = wretSt;
#else
			retSt = r.data();
#endif
		}
	}
	catch (std::exception& e)
	{
		_tcscpy(szErrorMessage, _T("ERROR in regular expression "));
#ifndef UNICODE
		_tcscat(szErrorMessage, e.what());
#endif
		*bMatched = FALSE;
		return retSt;
	}

	return retSt;
}

CString CUtils::FormatExpression(CString sMatchExpression, CString sFormatExpression, CString sInputString, BOOL bPartialMatch)
{
	boost::regex	e;
	unsigned int				match_flag = bPartialMatch ? boost::match_partial : boost::match_default;
	unsigned int				flags = match_flag | boost::format_perl;
	CString			retSt = sInputString;

#ifdef UNICODE
	std::wstring wMatch(sMatchExpression);
	const std::string match(wMatch.begin(), wMatch.end());
	std::wstring wInput(sInputString);
	const std::string input(wInput.begin(), wInput.end());
	std::wstring wFormat(sFormatExpression);
	const std::string format(wFormat.begin(), wFormat.end());
#else
	TCHAR szMatch[1024], szInput[1024], szFormat[1024];
	_tcscpy(szMatch, sMatchExpression);
	_tcscpy(szInput, sInputString);
	_tcscpy(szFormat, sFormatExpression);
	std::string match = szMatch;
	std::string input = szInput;
	std::string format = szFormat;
#endif


	try {
		e.assign(match);
	}
	catch (std::exception& e) {
#ifdef UNICODE
		Logprintf(_T("ERROR in regular expression "));
#else
		Logprintf(_T("ERROR in regular expression %s"), e.what());
#endif
		return retSt;
	}

	try {
		BOOL ret = boost::regex_match(input, e, bPartialMatch ? boost::match_partial | boost::match_default : boost::match_default);

		if (ret == TRUE) {
			std::string r = boost::regex_merge(input, e, format, boost::format_perl | (bPartialMatch ? boost::match_partial | boost::match_default : boost::match_default));
#ifdef UNICODE
			CStringA retStA = r.data();
			CString wretSt(retStA);
			retSt = wretSt;
#else
			retSt = r.data();
#endif
		}
	}
	catch (std::exception& e)
	{
		Logprintf(_T("ERROR in regular expression %s"), e.what());
		return retSt;
	}
	return retSt;
}

void CUtils::TruncateLogFile(CString sFile, DWORD dwMaxSize)
{
	DWORD dwCurrentSize = GetFileSize(sFile);

	if (dwCurrentSize == 0)
		return;

	if (dwCurrentSize > dwMaxSize) {
		::MoveFileEx(sFile, GetFilePath(sFile) + _T("\\") + GetFileName(sFile, TRUE) + _T("2.") + GetExtension(sFile), MOVEFILE_COPY_ALLOWED | MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH);
	}
}

void CUtils::Logprintf(const TCHAR *msg, ...)
{
	TCHAR	szLogLine[16000], szFinalLine[16000];
	va_list	ap;
	DWORD	n, nBytesWritten;
	SYSTEMTIME	ltime;

	//if (g_prefs.m_logToFile == FALSE)
	//	return;

	TruncateLogFile(g_prefs.m_logfolder + _T("\\ImportService.log"), 10000000);

	HANDLE hFile = ::CreateFile(g_prefs.m_logfolder + _T("\\ImportService.log"), GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
	if (hFile == INVALID_HANDLE_VALUE)
		return;

	// Seek to end of file
	::SetFilePointer(hFile, 0, NULL, FILE_END);

	va_start(ap, msg);
	n = _vstprintf(szLogLine, msg, ap);
	va_end(ap);
	szLogLine[n] = _T('\0');

	::GetLocalTime(&ltime);
	_stprintf(szFinalLine, _T("[%.2d-%.2d %.2d:%.2d:%.2d.%.3d] %s\r\n"), (int)ltime.wDay, (int)ltime.wMonth, (int)ltime.wHour, (int)ltime.wMinute, (int)ltime.wSecond, (int)ltime.wMilliseconds, szLogLine);

	::WriteFile(hFile, szFinalLine, (DWORD)_tcsclen(szFinalLine), &nBytesWritten, NULL);

	::CloseHandle(hFile);


}

int CUtils::ScanDirMaxFiles(CString sSpoolFolder, CString sSearchMask, CStringArray &sFoundList, int nMaxFiles, BOOL bIgnoreHiddenFiles)
{
	WIN32_FIND_DATA	fdata;
	HANDLE			fHandle;
	TCHAR			NameSearch[MAX_PATH], FoundFile[MAX_PATH];
	int				nFiles = 0;

	sFoundList.RemoveAll();

	if (CheckFolderWithPing(sSpoolFolder) == FALSE)
		return -1;

	if (sSpoolFolder.Right(1) != "\\")
		sSpoolFolder += _T("\\");
	if (sSearchMask == _T(""))
		sSearchMask = _T("*.*");

	_stprintf(NameSearch, _T("%s%s"), (LPCTSTR)sSpoolFolder, (LPCTSTR)sSearchMask);
	fHandle = ::FindFirstFile(NameSearch, &fdata);
	if (fHandle == INVALID_HANDLE_VALUE) {
		// All empty
		return 0;
	}

	// Got something
	do {

		if (!(fdata.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) && (!bIgnoreHiddenFiles || !(fdata.dwFileAttributes & FILE_ATTRIBUTE_HIDDEN)) && fdata.nFileSizeLow > 0) {
			_stprintf(FoundFile, _T("%s\\%s"), (LPCTSTR)sSpoolFolder, fdata.cFileName);
			HANDLE Hndl = CreateFile(FoundFile, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, 0, 0);
			if (Hndl != INVALID_HANDLE_VALUE) {
				CloseHandle(Hndl);

				CString sFile = fdata.cFileName;
				sFoundList.Add(sFile);
				nFiles++;
			}
		}

		if (!::FindNextFile(fHandle, &fdata))
			break;

	} while (nFiles < nMaxFiles);

	::FindClose(fHandle);

	return nFiles;
}

int CUtils::ScanSubDirNames(CString sSpoolFolder, CStringArray &sFoundList, int nMaxFolders)
{
	WIN32_FIND_DATA	fdata;
	HANDLE			fHandle;
	TCHAR			NameSearch[MAX_PATH];
	int				nFolders = 0;

	sFoundList.RemoveAll();

	if (CheckFolderWithPing(sSpoolFolder) == FALSE)
		return -1;

	if (sSpoolFolder.Right(1) != _T("\\"))
		sSpoolFolder += _T("\\");

	_stprintf(NameSearch, _T("%s*.*"), (LPCTSTR)sSpoolFolder);
	fHandle = ::FindFirstFile(NameSearch, &fdata);
	if (fHandle == INVALID_HANDLE_VALUE) {
		// All empty
		return 0;
	}

	// Got something
	do {

		if ((fdata.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)) {

			CString s(fdata.cFileName);
			if (s != _T("..") && s != _T(".")) {
				sFoundList.Add(s);
				nFolders++;
			}
		}

		if (!::FindNextFile(fHandle, &fdata))
			break;

	} while (nFolders < nMaxFolders);

	::FindClose(fHandle);

	return nFolders;
}


int CUtils::ScanDirMaxFilesEx(CString sSpoolFolder, CString sSearchMask, FILEINFOSTRUCT aFileList[], int nMaxFiles)
{
	return ScanDirMaxFilesEx(sSpoolFolder, sSearchMask, aFileList, nMaxFiles, FOLDERSCANORDER_FIFO, FALSE, FALSE);
}

int CUtils::ScanDirMaxFilesEx(CString sSpoolFolder, CString sSearchMask, FILEINFOSTRUCT aFileList[], int nMaxFiles, int nFolderScanOrder, BOOL bIgnoreHiddenFiles)
{
	return ScanDirMaxFilesEx(sSpoolFolder, sSearchMask, aFileList, nMaxFiles, nFolderScanOrder, bIgnoreHiddenFiles, FALSE);
}

int CUtils::ScanDirMaxFilesEx(CString sSpoolFolder, CString sSearchMask, FILEINFOSTRUCT aFileList[], int nMaxFiles, int nFolderScanOrder, BOOL bIgnoreHiddenFiles, BOOL bGetLogInfo)
{
	WIN32_FIND_DATA	fdata;
	HANDLE			fHandle;
	TCHAR			NameSearch[MAX_PATH], FoundFile[MAX_PATH];
	int				nFiles = 0;
	FILETIME		WriteTime, CreateTime, LocalWriteTime;
	SYSTEMTIME		SysTime;
	BOOL			sortOnCreateTime = TRUE;

	if (CheckFolderWithPing(sSpoolFolder) == FALSE)
		return -1;

	if (sSpoolFolder.Right(1) != "\\")
		sSpoolFolder += _T("\\");

	_stprintf(NameSearch, _T("%s%s"), (LPCTSTR)sSpoolFolder, (LPCTSTR)sSearchMask);
	fHandle = ::FindFirstFile(NameSearch, &fdata);
	if (fHandle == INVALID_HANDLE_VALUE) {
		return 0;
	}

	do {

		if (!(fdata.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) && (!bIgnoreHiddenFiles || !(fdata.dwFileAttributes & FILE_ATTRIBUTE_HIDDEN)) && fdata.nFileSizeLow > 0) {
			_stprintf(FoundFile, _T("%s\\%s"), (LPCTSTR)sSpoolFolder, fdata.cFileName);

			BOOL bSkipFile = FALSE;

			CString sExt = _T("");
			CString sFileName = fdata.cFileName;
			int n = sFileName.ReverseFind(_T('.'));
			if (n == -1)
				sExt = _T("");
			sExt = sFileName.Mid(n + 1);

			if (sExt != "") {
				//				if (g_prefs.m_ignoreextension != "") {
					//				if (sExt.CompareNoCase(g_prefs.m_ignoreextension) == 0)
						//				bSkipFile = TRUE;
							//	}
			}

			if (bSkipFile == FALSE) {
				HANDLE Hndl = ::CreateFile(FoundFile, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, 0, 0);
				if (Hndl != INVALID_HANDLE_VALUE) {
					int ok = ::GetFileTime(Hndl, &CreateTime, NULL, &WriteTime);

					::CloseHandle(Hndl);

					if (ok) {
						CString sFile = fdata.cFileName;
						aFileList[nFiles].sFileName = sFile;
						aFileList[nFiles].nFileSize = fdata.nFileSizeLow;

						// Store create time for FIFO sorting

						FileTimeToLocalFileTime(sortOnCreateTime ? &CreateTime : &WriteTime, &LocalWriteTime);
						FileTimeToSystemTime(&LocalWriteTime, &SysTime);
						CTime t((int)SysTime.wYear, (int)SysTime.wMonth, (int)SysTime.wDay, (int)SysTime.wHour, (int)SysTime.wMinute, (int)SysTime.wSecond);
						aFileList[nFiles].tJobTime = t;

						// Store last write time for stable time checker
						FileTimeToLocalFileTime(&WriteTime, &LocalWriteTime);
						FileTimeToSystemTime(&LocalWriteTime, &SysTime);
						CTime t2((int)SysTime.wYear, (int)SysTime.wMonth, (int)SysTime.wDay, (int)SysTime.wHour, (int)SysTime.wMinute, (int)SysTime.wSecond);
						aFileList[nFiles].tWriteTime = t2;

						nFiles++;
					}
				}
			}
		}

		if (!::FindNextFile(fHandle, &fdata))
			break;

	} while (nFiles < nMaxFiles);

	::FindClose(fHandle);

	if (nFolderScanOrder != FOLDERSCANORDER_ALPHABETIC) {
		// Sort found files on create-time
		CTime		tTmpTimeValueCTime;
		CString 	sTmpFileName;
		DWORD		nTmpFileSize;
		for (int i = 0; i < nFiles - 1; i++) {
			for (int j = i + 1; j < nFiles; j++) {
				if (aFileList[i].tJobTime > aFileList[j].tJobTime) {

					// Swap elements
					sTmpFileName = aFileList[i].sFileName;
					aFileList[i].sFileName = aFileList[j].sFileName;
					aFileList[j].sFileName = sTmpFileName;

					tTmpTimeValueCTime = aFileList[i].tJobTime;
					aFileList[i].tJobTime = aFileList[j].tJobTime;
					aFileList[j].tJobTime = tTmpTimeValueCTime;

					nTmpFileSize = aFileList[i].nFileSize;
					aFileList[i].nFileSize = aFileList[j].nFileSize;
					aFileList[j].nFileSize = nTmpFileSize;

					tTmpTimeValueCTime = aFileList[i].tWriteTime;
					aFileList[i].tWriteTime = aFileList[j].tWriteTime;
					aFileList[j].tWriteTime = tTmpTimeValueCTime;
				}
			}
		}
	}

	return nFiles;
}

int CUtils::ScanDirMaxFilesEx(CString sSpoolFolder, CString sSearchMask, FILEINFOSTRUCT aFileList[], int nMaxFiles, CString sIgnoreMask,
	BOOL bIgnoreHidden, BOOL bKillZeroByteFiles, CString sRegExpSearchMask, int nLockCheckMode)
{
	WIN32_FIND_DATA	fdata;
	HANDLE			fHandle;
	TCHAR			NameSearch[MAX_PATH], FoundFile[MAX_PATH];
	int				nFiles = 0;
	FILETIME		WriteTime, CreateTime, LocalWriteTime;
	SYSTEMTIME		SysTime;

	if (CheckFolderWithPing(sSpoolFolder) == FALSE)
		return -1;

	if (sSpoolFolder.Right(1) != _T("\\"))
		sSpoolFolder += _T("\\");
	if (sSearchMask == _T(""))
		sSearchMask = _T("*.*");

	_stprintf(NameSearch, _T("%s%s"), (LPCTSTR)sSpoolFolder, (LPCTSTR)sSearchMask);
	fHandle = ::FindFirstFile(NameSearch, &fdata);
	if (fHandle == INVALID_HANDLE_VALUE) {
		// All empty
		return 0;
	}

	// Got something
	do {

		if (!(fdata.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)) {

			if (bIgnoreHidden == FALSE ||
				(bIgnoreHidden && (fdata.dwFileAttributes & FILE_ATTRIBUTE_HIDDEN) == 0) ||
				(bIgnoreHidden && (fdata.dwFileAttributes & FILE_ATTRIBUTE_SYSTEM) == 0)) {

				BOOL bFilterOK = TRUE;

				if (sIgnoreMask != _T("") && sIgnoreMask != _T("*") && sIgnoreMask != _T("*.*")) {
					CString s1 = GetExtension(fdata.cFileName);
					CString s2 = GetExtension(sIgnoreMask);
					if (s1.CompareNoCase(s2) == 0)
						bFilterOK = FALSE;
				}

				CString sFile = fdata.cFileName;
				if (sFile.CompareNoCase(_T(".DS_Store")) == 0 || sFile.CompareNoCase(_T("Thumbs.db")) == 0)
					bFilterOK = FALSE;

				if (bFilterOK && sRegExpSearchMask != _T("")) {
					bFilterOK = FALSE;
					sFile.MakeUpper();
					bFilterOK = TryMatchExpression(sRegExpSearchMask, fdata.cFileName, FALSE);
				}

				if (bFilterOK) {

					_stprintf(FoundFile, _T("%s\\%s"), (LPCTSTR)sSpoolFolder, fdata.cFileName);
					HANDLE Hndl = INVALID_HANDLE_VALUE;
					if (nLockCheckMode == LOCKCHECKMODE_READ)
						Hndl = ::CreateFile(FoundFile, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, 0, 0);
					else if (nLockCheckMode != LOCKCHECKMODE_NONE)
						Hndl = ::CreateFile(FoundFile, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, 0, 0);

					if (Hndl != INVALID_HANDLE_VALUE || nLockCheckMode == LOCKCHECKMODE_NONE) {

						int ok = TRUE;

						if (nLockCheckMode != LOCKCHECKMODE_NONE) {
							ok = ::GetFileTime(Hndl, &CreateTime, NULL, &WriteTime);
							::CloseHandle(Hndl);
						}
						else {
							CreateTime.dwHighDateTime = fdata.ftCreationTime.dwHighDateTime;
							CreateTime.dwLowDateTime = fdata.ftCreationTime.dwLowDateTime;
							WriteTime.dwHighDateTime = fdata.ftLastWriteTime.dwHighDateTime;
							WriteTime.dwLowDateTime = fdata.ftLastWriteTime.dwLowDateTime;
						}

						if (bKillZeroByteFiles) {
							if (fdata.nFileSizeLow == 0 && fdata.nFileSizeHigh == 0 && GetFileAgeMinutes(FoundFile) > 5) {
								::DeleteFile(FoundFile);
								ok = FALSE;
							}
						}

						if (ok && fdata.nFileSizeLow == 0 && fdata.nFileSizeHigh == 0)
							ok = FALSE;

						if (ok) {
							aFileList[nFiles].sFolder = sSpoolFolder;
							aFileList[nFiles].sFileName = fdata.cFileName;
							aFileList[nFiles].nFileSize = fdata.nFileSizeLow;
							aFileList[nFiles].tJobStartTime = CTime::GetCurrentTime();

							// Store create time for FIFO sorting

							FileTimeToLocalFileTime(&CreateTime, &LocalWriteTime);
							FileTimeToSystemTime(&LocalWriteTime, &SysTime);
							CTime t((int)SysTime.wYear, (int)SysTime.wMonth, (int)SysTime.wDay, (int)SysTime.wHour, (int)SysTime.wMinute, (int)SysTime.wSecond);
							aFileList[nFiles].tJobTime = t;

							// Store last write time for stable time checker

							FileTimeToLocalFileTime(&WriteTime, &LocalWriteTime);
							FileTimeToSystemTime(&LocalWriteTime, &SysTime);
							CTime t2((int)SysTime.wYear, (int)SysTime.wMonth, (int)SysTime.wDay, (int)SysTime.wHour, (int)SysTime.wMinute, (int)SysTime.wSecond);
							aFileList[nFiles++].tWriteTime = t2;

						}
					}
				}
			}
		}

		if (!::FindNextFile(fHandle, &fdata))
			break;

	} while (nFiles < nMaxFiles);

	::FindClose(fHandle);

	// Sort found files on create-time
	CTime		tTmpTimeValueCTime;
	CString 	sTmpFileName;
	DWORD		nTmpFileSize;
	for (int i = 0; i < nFiles - 1; i++) {
		for (int j = i + 1; j < nFiles; j++) {
			if (aFileList[i].tJobTime > aFileList[j].tJobTime) {

				// Swap elements
				sTmpFileName = aFileList[i].sFileName;
				aFileList[i].sFileName = aFileList[j].sFileName;
				aFileList[j].sFileName = sTmpFileName;

				tTmpTimeValueCTime = aFileList[i].tJobTime;
				aFileList[i].tJobTime = aFileList[j].tJobTime;
				aFileList[j].tJobTime = tTmpTimeValueCTime;

				nTmpFileSize = aFileList[i].nFileSize;
				aFileList[i].nFileSize = aFileList[j].nFileSize;
				aFileList[j].nFileSize = nTmpFileSize;

				tTmpTimeValueCTime = aFileList[i].tWriteTime;
				aFileList[i].tWriteTime = aFileList[j].tWriteTime;
				aFileList[j].tWriteTime = tTmpTimeValueCTime;

				tTmpTimeValueCTime = aFileList[i].tJobStartTime;
				aFileList[i].tJobStartTime = aFileList[j].tJobStartTime;
				aFileList[j].tJobStartTime = tTmpTimeValueCTime;
			}
		}
	}

	return nFiles;
}

int CUtils::ScanDirCountWithSubfolder(CString sSpoolFolder, CString sSearchMask)
{
	DWORD dwArchived;
	CStringArray sDirList;
	int count = ScanDirCount(sSpoolFolder, sSearchMask, dwArchived, FALSE);
	if (count < 0)
		return -1;

	int nSubdirs = ScanSubDirNames(sSpoolFolder, sDirList, 100);
	for (int i = 0; i < sDirList.GetCount(); i++) {
		if (sDirList[i] != _T(".") && sDirList[i] != _T("..")) {
			int subcount = ScanDirCount(sSpoolFolder + _T("\\") + sDirList[i], sSearchMask, dwArchived, TRUE);
			if (subcount > 0)
				count += subcount;
		}
	}
	return count;

}

BOOL	CUtils::DeleteEmptyDirectory(CString sFolder)
{
	DWORD dwArchived;
	if (ScanDirCount(sFolder, _T("*.*"), dwArchived, TRUE) == 0)
		return ::RemoveDirectory(sFolder);

	return FALSE;
}

int CUtils::ScanDirCount(CString sSpoolFolder, CString sSearchMask)
{
	DWORD dwArchived;
	return ScanDirCount(sSpoolFolder, sSearchMask, dwArchived, FALSE);
}

int CUtils::ScanDirCount(CString sSpoolFolder, CString sSearchMask, DWORD &dwArchived, BOOL bIgnoreHiddenFiles)
{
	WIN32_FIND_DATA	fdata;
	HANDLE			fHandle;
	TCHAR			NameSearch[MAX_PATH];
	int				nFiles = 0;
	dwArchived = 0;

	if (CheckFolderWithPing(sSpoolFolder) == FALSE)
		return -1;

	if (sSpoolFolder.Right(1) != "\\")
		sSpoolFolder += _T("\\");
	if (sSearchMask == _T(""))
		sSearchMask = _T("*.*");

	_stprintf(NameSearch, _T("%s%s"), (LPCTSTR)sSpoolFolder, (LPCTSTR)sSearchMask);
	fHandle = ::FindFirstFile(NameSearch, &fdata);
	if (fHandle == INVALID_HANDLE_VALUE) {
		return 0;
	}

	// Got something
	do {
		if (!(fdata.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) && (!bIgnoreHiddenFiles || !(fdata.dwFileAttributes & FILE_ATTRIBUTE_HIDDEN)) && fdata.nFileSizeLow > 0) {

			CString s(fdata.cFileName);
			s.MakeLower();
			if (s[0] != _T('.')) {

				if (s != _T("thumbs.db")) {
					nFiles++;

					if (fdata.dwFileAttributes & FILE_ATTRIBUTE_ARCHIVE)
						dwArchived++;
				}
			}
		}

		if (!::FindNextFile(fHandle, &fdata))
			break;
	} while (TRUE);

	::FindClose(fHandle);

	return nFiles;
}


int CUtils::DirectoryList(CString sSpoolFolder, FILEINFOSTRUCT aFileList[], int nMaxDirs)
{
	WIN32_FIND_DATA	fdata;
	HANDLE			fHandle;
	int				nDirs = 0;

	if (CheckFolderWithPing(sSpoolFolder) == FALSE)
		return -1;

	sSpoolFolder += _T("\\*.*");

	fHandle = ::FindFirstFile(sSpoolFolder, &fdata);
	if (fHandle == INVALID_HANDLE_VALUE) {
		// All empty
		return 0;
	}

	// Got something
	do {

		if ((fdata.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) && fdata.cFileName[0] != _T('.')) {
			aFileList[nDirs++].sFileName = fdata.cFileName;
		}

		if (!::FindNextFile(fHandle, &fdata))
			break;

	} while (nDirs < nMaxDirs);

	::FindClose(fHandle);

	return nDirs;
}

CString CUtils::Time2String(CTime tm)
{
	CString s;
	s.Format(_T("%.4d-%.2d-%.2d %.2d:%.2d:%.2d"), tm.GetYear(), tm.GetMonth(), tm.GetDay(), tm.GetHour(), tm.GetMinute(), tm.GetSecond());

	return s;
}

CString CUtils::GenerateTimeStamp(BOOL bIncludemillisecs)
{
	CString s;
	SYSTEMTIME ltime;
	::GetLocalTime(&ltime);
	if (bIncludemillisecs)
		s.Format(_T("%.4d%.2d%.2d%.2d%.2d%2d%.3d"), (int)ltime.wYear, (int)ltime.wMonth, (int)ltime.wDay, (int)ltime.wHour, (int)ltime.wMinute, (int)ltime.wSecond, (int)ltime.wMilliseconds);
	else
		s.Format(_T("%.4d%.2d%.2d%.2d%.2d%2d"), (int)ltime.wYear, (int)ltime.wMonth, (int)ltime.wDay, (int)ltime.wHour, (int)ltime.wMinute, (int)ltime.wSecond);
	return s;
}

// 12.34000
// 01234
double CUtils::SafeStringToDouble(CString s)
{
	s += _T("000");
	int n = _tstoi(s);
	s.Replace(_T(","), _T("."));
	int i = s.Find(_T("."));
	if (i == -1)
		return (double)n;

	int m = _tstoi(s.Mid(n + 1, 3)); // 1/1000-parts
	return (double)n + (double)m / 1000.0;
}

CTime CUtils::String2Time(CString s)
{
	// 2010-01-22 23:24:25

	CTime tm(1975, 1, 1, 0, 0, 0);

	if (s.GetLength() < 19)
		return tm;

	int y = _tstoi(s.Mid(0, 4));
	int m = _tstoi(s.Mid(5, 2));
	int d = _tstoi(s.Mid(8, 2));
	int h = _tstoi(s.Mid(11, 2));
	int mi = _tstoi(s.Mid(14, 2));
	int sec = _tstoi(s.Mid(17, 2));

	try {
		CTime tm2(y, m, d, h, mi, sec);
		tm = tm2;
	}
	catch (CException *ex)
	{
	}

	return tm;
}

int CUtils::DoCreateProcess(CString sCmdLine)
{
	return DoCreateProcess(sCmdLine, g_prefs.m_processtimeout, TRUE);
}

int CUtils::DoCreateProcess(CString sCmdLine, int nTimeout, BOOL bBlocking)
{
	DWORD dwExitCode;
	DWORD dwPID;
	return DoCreateProcess(sCmdLine, nTimeout, bBlocking, dwExitCode, dwPID);
}


BOOL CUtils::DoCreateProcessEx(CString sCmdLine, int nTimeout)
{
	STARTUPINFO			si;
	PROCESS_INFORMATION pi;

	ZeroMemory(&si, sizeof(si));
	si.cb = sizeof(si);

	si.dwFlags = STARTF_USESHOWWINDOW;
	si.wShowWindow = SW_HIDE;

	// Start the child process. 
	TCHAR szCmdLine[MAX_PATH];
	strcpy(szCmdLine, sCmdLine);
	if (!::CreateProcess(NULL,	// No module name (use command line). 
		szCmdLine,				// Command line. 
		NULL,					// Process handle not inheritable. 
		NULL,					// Thread handle not inheritable. 
		FALSE,					// Set handle inheritance to FALSE. 
		0,						// No creation flags. 
		NULL,					// Use parent's environment block. 
		NULL,					// Use parent's starting directory. 
		&si,					// Pointer to STARTUPINFO structure.
		&pi)) {				// Pointer to PROCESS_INFORMATION structure.
		CString sMsg;
		sMsg.Format("Create Process for external script file failed. (%s)", szCmdLine);
		//AfxMessageBox(sMsg);
		return FALSE;
	}

	::WaitForInputIdle(pi.hProcess, nTimeout > 0 ? nTimeout * 1000 : INFINITE);
	::WaitForSingleObject(pi.hProcess, nTimeout > 0 ? nTimeout * 1000 : INFINITE);
	::CloseHandle(pi.hProcess);
	::CloseHandle(pi.hThread);
	return TRUE;
}

//extern THREADINFO threadInfo[16];

int CUtils::DoCreateProcess(CString sCmdLine, int nTimeout, BOOL bBlocking, DWORD &dwExitCode, DWORD &dwProcessID)
{

	STARTUPINFO			si;
	PROCESS_INFORMATION pi;
	dwExitCode = 0;
	dwProcessID = 0;

	ZeroMemory(&si, sizeof(si));
	si.cb = sizeof(si);
	si.dwFlags = STARTF_USESHOWWINDOW;
	si.wShowWindow = SW_HIDE;

	ZeroMemory(&pi, sizeof(pi));

	// Start the child process. 
	TCHAR szCmdLine[MAX_PATH * 4];
	_tcscpy(szCmdLine, sCmdLine);
	if (!::CreateProcess(NULL,	// No module name (use command line). 
		szCmdLine,				// Command line. 
		NULL,					// Process handle not inheritable. 
		NULL,					// Thread handle not inheritable. 
		FALSE,					// Set handle inheritance to FALSE. 
		0,						// No creation flags. 
		NULL,					// Use parent's environment block. 
		NULL,					// Use parent's starting directory. 
		&si,					// Pointer to STARTUPINFO structure.
		&pi)) {				// Pointer to PROCESS_INFORMATION structure.
		CString sMsg;
		sMsg.Format(_T("Create Process for external script file failed. (%s)"), szCmdLine);
		//AfxMessageBox(sMsg);
		return FALSE;
	}

	dwProcessID = pi.dwProcessId;

	DWORD retWaitForInputIdle = 0;
	DWORD retWaitForSingleObject = 0;
	if (bBlocking) {

		retWaitForInputIdle = ::WaitForInputIdle(pi.hProcess, nTimeout > 0 ? nTimeout * 1000 : INFINITE);
		retWaitForSingleObject = ::WaitForSingleObject(pi.hProcess, nTimeout > 0 ? nTimeout * 1000 : INFINITE);
		::GetExitCodeProcess(pi.hProcess, &dwExitCode);
	}

	int ret = TRUE;
	if (nTimeout > 0 && (retWaitForInputIdle == WAIT_TIMEOUT || retWaitForSingleObject == WAIT_TIMEOUT))
		ret = -1;

	::CloseHandle(pi.hProcess);
	::CloseHandle(pi.hThread);
	return ret;
}

BOOL CUtils::CreateProcessGetTextResult(CString sCmdLine, int nTimeout, CString &strResult, DWORD &dwExitCode, DWORD &dwProcessID)
{

	dwProcessID = 0;
	strResult = _T(""); // Contains result of cmdArg.
	dwExitCode = 0;

	HANDLE hChildStdoutRd; // Read-side, used in calls to ReadFile() to get child's stdout output.
	HANDLE hChildStdoutWr; // Write-side, given to child process using si struct.

	 // Create security attributes to create pipe.
	SECURITY_ATTRIBUTES saAttr = { sizeof(SECURITY_ATTRIBUTES) };
	saAttr.bInheritHandle = TRUE; // Set the bInheritHandle flag so pipe handles are inherited by child process. Required.
	saAttr.lpSecurityDescriptor = NULL;

	// Create a pipe to get results from child's stdout.
	// I'll create only 1 because I don't need to pipe to the child's stdin.

	Logprintf(_T("INFO: calling CreatePipe .."));
	if (!CreatePipe(&hChildStdoutRd, &hChildStdoutWr, &saAttr, 0)) {
		return FALSE;
	}
	Logprintf(_T("INFO: CreatePipe OK"));

	STARTUPINFO			si;
	PROCESS_INFORMATION pi = { 0 };

	ZeroMemory(&si, sizeof(si));
	si.cb = sizeof(si);
	si.dwFlags = STARTF_USESHOWWINDOW | STARTF_USESTDHANDLES;;
	si.wShowWindow = SW_HIDE;
	si.hStdOutput = hChildStdoutWr; // Requires STARTF_USESTDHANDLES in dwFlags.
	si.hStdError = hChildStdoutWr; // Requires STARTF_USESTDHANDLES in dwFlags.

	// Start the child process. 
	TCHAR szCmdLine[MAX_PATH * 4];
	_tcscpy(szCmdLine, sCmdLine);
	if (!::CreateProcess(NULL,	// No module name(use command line). 
		szCmdLine,				// Command line. 
		NULL,					// Process handle not inheritable. 
		NULL,					// Thread handle not inheritable. 
		TRUE,					// Set handle inheritance to FALSE. 
		CREATE_NEW_CONSOLE,		// No creation flags. 
		NULL,					// Use parent's environment block. 
		NULL,					// Use parent's starting directory. 
		&si,					// Pointer to STARTUPINFO structure.
		&pi)) {				// Pointer to PROCESS_INFORMATION structure.
		return FALSE;
	}

	Logprintf(_T("INFO: CreateProcess fired .."));

	dwProcessID = pi.dwProcessId;

	DWORD retWaitForInputIdle = ::WaitForInputIdle(pi.hProcess, nTimeout > 0 ? nTimeout * 1000 : INFINITE);
	DWORD retWaitForSingleObject = ::WaitForSingleObject(pi.hProcess, nTimeout > 0 ? nTimeout * 1000 : INFINITE);

	Logprintf(_T("INFO: Getting exit code.."));
	::GetExitCodeProcess(pi.hProcess, &dwExitCode);

	int ret = TRUE;
	if (nTimeout > 0 && (retWaitForInputIdle == WAIT_TIMEOUT || retWaitForSingleObject == WAIT_TIMEOUT))
		ret = -1;

	Logprintf(_T("INFO: Terminating process.."));

	::TerminateProcess(pi.hProcess, 0);

	if (!::CloseHandle(hChildStdoutWr))
		return FALSE;

	Logprintf(_T("INFO: Waiting for result: .."));

	for (;;) {
		DWORD dwRead;
		CHAR chBuf[1025];

		// Read chunk of text from pipe that is the standard output for child process.
		BOOL done = !::ReadFile(hChildStdoutRd, chBuf, 1024, &dwRead, NULL) || dwRead == 0;

		if (done)
			break;

		// Append result to string.
		strResult += CString(chBuf, dwRead);

		// Limit text output...
		if (strResult.GetLength() >= 10000)
			break;
	}

	// Close process and thread handles.
	::CloseHandle(hChildStdoutRd);

	::CloseHandle(pi.hProcess);
	::CloseHandle(pi.hThread);

	Logprintf(_T("INFO: Process result: %s"), strResult);

	return ret;
}


int CUtils::DeleteFiles(CString sSpoolFolder, CString sSearchMask)
{
	return DeleteFiles(sSpoolFolder, sSearchMask, 0);
}

int CUtils::DeleteFiles(CString sSpoolFolder, CString sSearchMask, int sMaxAgeHours)
{
	WIN32_FIND_DATA	fdata;
	HANDLE			fHandle;
	TCHAR			NameSearch[MAX_PATH], FoundFile[MAX_PATH];
	int				nFiles = 0;
	CStringArray	arr;
	FILETIME ftm;
	FILETIME ltm;
	SYSTEMTIME  stm;

	CTime tNow = CTime::GetCurrentTime();

	if (CheckFolder(sSpoolFolder) == FALSE)
		return -1;

	if (sSpoolFolder.Right(1) != _T("\\"))
		sSpoolFolder += _T("\\");

	_stprintf(NameSearch, _T("%s%s"), (LPCTSTR)sSpoolFolder, (LPCTSTR)sSearchMask);
	fHandle = ::FindFirstFile(NameSearch, &fdata);
	if (fHandle == INVALID_HANDLE_VALUE) {
		// All empty
		return 0;
	}

	do {
		if (!(fdata.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)) {
			_stprintf(FoundFile, _T("%s\\%s"), (LPCTSTR)sSpoolFolder, fdata.cFileName);
			HANDLE Hndl = CreateFile(FoundFile, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, 0, 0);
			if (Hndl != INVALID_HANDLE_VALUE) {
				CloseHandle(Hndl);
				ftm = fdata.ftLastWriteTime;
				::FileTimeToLocalFileTime(&ftm, &ltm);
				::FileTimeToSystemTime(&ltm, &stm);

				CTime tFileAge((int)stm.wYear, (int)stm.wMonth, (int)stm.wDay, (int)stm.wHour, (int)stm.wMinute, (int)stm.wSecond);
				CTimeSpan ts = tNow - tFileAge;
				if (sMaxAgeHours == 0 || ts.GetTotalHours() >= sMaxAgeHours) {
					arr.Add(fdata.cFileName);
					nFiles++;
				}
			}
		}

		if (!::FindNextFile(fHandle, &fdata))
			break;

	} while (TRUE);

	::FindClose(fHandle);

	for (int i = 0; i < arr.GetCount(); i++) {
		::DeleteFile(sSpoolFolder + arr.ElementAt(i));
		::DeleteFile(sSpoolFolder + _T("logs\\") + arr.ElementAt(i) + _T(".txt"));
	}
	return nFiles;
}


int CUtils::CopyFiles(CString sSpoolFolder, CString sSearchMask, CString sDestFolder)
{
	WIN32_FIND_DATA	fdata;
	HANDLE			fHandle;
	TCHAR			NameSearch[MAX_PATH], FoundFile[MAX_PATH], DestFoundFile[MAX_PATH];
	int nFiles = 0;

	if (CheckFolder(sSpoolFolder) == FALSE)
		return -1;

	if (sSpoolFolder.Right(1) != _T("\\"))
		sSpoolFolder += _T("\\");
	if (sDestFolder.Right(1) != _T("\\"))
		sDestFolder += _T("\\");

	_stprintf(NameSearch, _T("%s%s"), (LPCTSTR)sSpoolFolder, (LPCTSTR)sSearchMask);
	fHandle = ::FindFirstFile(NameSearch, &fdata);
	if (fHandle == INVALID_HANDLE_VALUE) {
		// All empty
		return 0;
	}

	// Got something
	do {
		if (!(fdata.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)) {
			_stprintf(FoundFile, _T("%s%s"), (LPCTSTR)sSpoolFolder, fdata.cFileName);
			HANDLE Hndl = CreateFile(FoundFile, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, 0, 0);
			BOOL ok = FALSE;
			if (Hndl != INVALID_HANDLE_VALUE) {
				CloseHandle(Hndl);
				ok = TRUE;

			}
			if (ok) {
				_stprintf(DestFoundFile, _T("%s%s"), (LPCTSTR)sDestFolder, fdata.cFileName);
				::CopyFile(FoundFile, DestFoundFile, FALSE);
				nFiles++;
			}
		}

		if (!::FindNextFile(fHandle, &fdata))
			break;

	} while (TRUE);

	::FindClose(fHandle);

	return nFiles;
}



int CUtils::CRC32(CString sFileName)
{
	HANDLE	Hndl;
	BYTE	buf[8192];
	DWORD	nBytesRead;

	if ((Hndl = ::CreateFile(sFileName, GENERIC_READ, FILE_SHARE_WRITE | FILE_SHARE_READ, NULL, OPEN_EXISTING, 0, 0)) == INVALID_HANDLE_VALUE)
		return 0;

	DWORD sum = 0;
	do {
		if (!::ReadFile(Hndl, buf, 8092, &nBytesRead, NULL)) {
			::CloseHandle(Hndl);
			return 0;
		}
		if (nBytesRead > 0) {
			for (int i = 0; i < (int)nBytesRead; i++) {
				sum += buf[i];
				// just let it overflow..
			}
		}

	} while (nBytesRead > 0);

	::CloseHandle(Hndl);

	int nCRC = sum;
	return nCRC;
}

BOOL CUtils::SetFileToCurrentTime(CString sFile)
{
	FILETIME ft;
	SYSTEMTIME st;
	BOOL ret;

	HANDLE hFile = ::CreateFile(sFile, FILE_WRITE_ATTRIBUTES, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, 0, 0);

	if (hFile == INVALID_HANDLE_VALUE)
		return FALSE;

	::GetSystemTime(&st);              // gets current time
	::SystemTimeToFileTime(&st, &ft);  // converts to file time format

	ret = ::SetFileTime(hFile, &ft, (LPFILETIME)NULL, &ft);

	::CloseHandle(hFile);

	return ret;
}



// Function name   : GetOwner
// Description     : Determines the 'Owner' of a given file or folder.
// Return type     : UINT is S_OK if successful; A Win32 'ERROR_' value otherwise.
// Argument        : LPCWSTR szFileOrFolderPathName is the fully qualified path of the file or folder to examine.
// Argument        : LPWSTR pUserNameBuffer is a pointer to a buffer used to contain the resulting 'Owner' string.
// Argument        : int nSizeInBytes is the size of the buffer.
UINT CUtils::GetOwner(LPCTSTR szFileOrFolderPathName, LPTSTR pUserNameBuffer, int nSizeInBytes)
{
	// 1) Validate the path:
	// 1.1) Length should not be 0.
	// 1.2) Path must point to an existing file or folder.
	if (!_tcslen(szFileOrFolderPathName) || !PathFileExists(szFileOrFolderPathName))
		return ERROR_INVALID_PARAMETER;

	// 2) Validate the buffer:
	// 2.1) Size must not be 0.
	// 2.2) Pointer must not be NULL.
	if (nSizeInBytes <= 0 || pUserNameBuffer == NULL)
		return ERROR_INVALID_PARAMETER;

	// 3) Convert the path to UNC if it is not already UNC so that we can extract a machine name from it:
	// 3.1) Use a big buffer... some OS's can have a path that is 32768 chars in length.
	TCHAR szUNCPathName[32767] = { 0 };
	// 3.2) If path is not UNC...
	if (!PathIsUNC(szFileOrFolderPathName)) {
		// 3.3) Mask the big buffer into a UNIVERSAL_NAME_INFO.
		DWORD dwUniSize = 32767 * sizeof(TCHAR);
		UNIVERSAL_NAME_INFO* pUNI = (UNIVERSAL_NAME_INFO*)szUNCPathName;
		// 3.4) Attempt to get the UNC version of the path into the big buffer.
		if (!WNetGetUniversalName(szFileOrFolderPathName, UNIVERSAL_NAME_INFO_LEVEL, pUNI, &dwUniSize)) {
			// 3.5) If successful, copy the UNC version into the buffer.
			lstrcpy(szUNCPathName, pUNI->lpUniversalName); // You must extract from this offset as the buffer has UNI overhead.
		}
		else {
			// 3.6) If not successful, copy the original path into the buffer.
			lstrcpy(szUNCPathName, szFileOrFolderPathName);
		}
	}
	else {
		// 3.7) Path is already UNC, copy the original path into the buffer.
		lstrcpy(szUNCPathName, szFileOrFolderPathName);
	}

	// 4) If path is UNC (will not be the case for local physical drive paths) we want to extract the machine name:
	// 4.1) Use a buffer bug enough to hold a machine name per Win32.
	TCHAR szMachineName[MAX_COMPUTERNAME_LENGTH + 1] = { 0 };
	// 4.2) If path is UNC...
	if (PathIsUNC(szUNCPathName)) {
		// 4.3) Use PathFindNextComponent() to skip past the double backslashes.
		LPTSTR lpMachineName = PathFindNextComponent(szUNCPathName);
		// 4.4) Walk the the rest of the path to find the end of the machine name.
		int nPos = 0;
		LPTSTR lpNextSlash = lpMachineName;
		while ((lpNextSlash[0] != L'\\') && (lpNextSlash[0] != L'\0')) {
			nPos++;
			lpNextSlash++;
		}
		// 4.5) Copyt the machine name into the buffer.
		lstrcpyn(szMachineName, lpMachineName, nPos + 1);
	}

	// 5) Derive the 'Owner' by getting the owner's Security ID from a Security Descriptor associated with the file or folder indicated in the path.
	// 5.1) Get a security descriptor for the file or folder that contains the Owner Security Information.
	// 5.1.1) Use GetFileSecurity() with some null params to get the required buffer size.
	// 5.1.2) We don't really care about the return value.
	// 5.1.3) The error code must be ERROR_INSUFFICIENT_BUFFER for use to continue.
	unsigned long   uSizeNeeded = 0;
	GetFileSecurity(szUNCPathName, OWNER_SECURITY_INFORMATION, 0, 0, &uSizeNeeded);
	UINT uRet = GetLastError();
	if (uRet == ERROR_INSUFFICIENT_BUFFER && uSizeNeeded) {
		uRet = S_OK; // Clear the ERROR_INSUFFICIENT_BUFFER

		// 5.2) Allocate the buffer for the Security Descriptor, check for out of memory
		LPBYTE lpSecurityBuffer = (LPBYTE)malloc(uSizeNeeded * sizeof(BYTE));
		if (!lpSecurityBuffer) {
			return ERROR_NOT_ENOUGH_MEMORY;
		}

		// 5.2) Get the Security Descriptor that contains the Owner Security Information into the buffer, check for errors
		if (!GetFileSecurity(szUNCPathName, OWNER_SECURITY_INFORMATION, lpSecurityBuffer, uSizeNeeded, &uSizeNeeded)) {
			free(lpSecurityBuffer);
			return GetLastError();
		}

		// 5.3) Get the the owner's Security ID (SID) from the Security Descriptor, check for errors
		PSID pSID = NULL;
		BOOL bOwnerDefaulted = FALSE;
		if (!GetSecurityDescriptorOwner(lpSecurityBuffer, &pSID, &bOwnerDefaulted)) {
			free(lpSecurityBuffer);
			return GetLastError();
		}

		// 5.4) Get the size of the buffers needed for the owner information (domain and name)
		// 5.4.1) Use LookupAccountSid() with buffer sizes set to zero to get the required buffer sizes.
		// 5.4.2) We don't really care about the return value.
		// 5.4.3) The error code must be ERROR_INSUFFICIENT_BUFFER for use to continue.
		LPTSTR			pName = NULL;
		LPTSTR			pDomain = NULL;
		unsigned long   uNameLen = 0;
		unsigned long   uDomainLen = 0;
		SID_NAME_USE    sidNameUse = SidTypeUser;
		LookupAccountSid(szMachineName, pSID, pName, &uNameLen, pDomain, &uDomainLen, &sidNameUse);
		uRet = GetLastError();
		if ((uRet == ERROR_INSUFFICIENT_BUFFER) && uNameLen && uDomainLen) {
			uRet = S_OK; // Clear the ERROR_INSUFFICIENT_BUFFER

			// 5.5) Allocate the required buffers, check for out of memory
			pName = (LPTSTR)malloc(uNameLen * sizeof(TCHAR));
			pDomain = (LPTSTR)malloc(uDomainLen * sizeof(TCHAR));
			if (!pName || !pDomain) {
				free(lpSecurityBuffer);
				return ERROR_NOT_ENOUGH_MEMORY;
			}

			// 5.6) Get domain and username
			if (!LookupAccountSid(szMachineName, pSID, pName, &uNameLen, pDomain, &uDomainLen, &sidNameUse)) {
				free(pName);
				free(pDomain);
				free(lpSecurityBuffer);
				return GetLastError();
			}

			// 5.7) Build the owner string from the domain and username if there is enough room in the buffer.
			if (nSizeInBytes > ((uNameLen + uDomainLen + 2) * sizeof(TCHAR))) {
				lstrcpy(pUserNameBuffer, pDomain);
				//	lstrcat(pUserNameBuffer, L"\\");
				lstrcat(pUserNameBuffer, _T("\\"));
				lstrcat(pUserNameBuffer, pName);
			}
			else {
				uRet = ERROR_INSUFFICIENT_BUFFER;
			}

			// 5.8) Release memory
			free(pName);
			free(pDomain);
		}
		// 5.9) Release memory
		free(lpSecurityBuffer);
	}
	return uRet;
}

/*
CString CUtils::GetProcessOwner(CString sFileOrFolderPathName)
{
	CString sBuf = "";
	CString sCmd;
	DWORD dwBytesWritten;

	::DeleteFile("c:\\processinfo.txt");
	sCmd.Format("%s\\Handle.exe -a \"%s\" > c:\\processinfo.txt", g_prefs.m_appPath, sFileOrFolderPathName);

	HANDLE hFile = CreateFile(_T("c:\\gethandle.bat"),
		 GENERIC_WRITE, FILE_SHARE_READ,
		 NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);

	if (hFile == INVALID_HANDLE_VALUE)
	   return "";

	WriteFile(hFile, sCmd, sCmd.GetLength(), &dwBytesWritten, NULL);
	CloseHandle(hFile);


	DoCreateProcess("c:\\gethandle.bat", 120, TRUE);

	for (int i=0; i<500; i++) {
		Sleep(100);
		if (FileExist("c:\\processinfo.txt") && GetFileSize("c:\\processinfo.txt") > 0)
			break;
	}

	TRY {

		CTextFileRead file("c:\\processinfo.txt");

		//Read the file to a Unicode-string.
		wstring temp;
		file.Read(temp);
		file.Close();
		string simple;
		CTextFileBase::ConvertWcharToString(temp.c_str(), simple);
		sBuf = simple.c_str();

	}
	CATCH_ALL(e) {}
	END_CATCH_ALL

	return sBuf;
}
*/

int CUtils::ZipAllFiles(CString sSpoolFolder, CString sSearchMask, CString sZipName)
{
	/*	WIN32_FIND_DATA	fdata;
		HANDLE			fHandle;
		TCHAR			NameSearch[MAX_PATH], FoundFile[MAX_PATH];
		int nFiles = 0;

		if (CheckFolder(sSpoolFolder) == FALSE)
			return -1;

		if (sSpoolFolder.Right(1) != "\\")
			sSpoolFolder += _T("\\");

		CZipArchive m_zip;
		::DeleteFile(sZipName);
		m_zip.Open(sZipName, CZipArchive::zipCreate);

		_stprintf(NameSearch, "%s%s",(LPCTSTR)sSpoolFolder, (LPCTSTR)sSearchMask);
		fHandle = ::FindFirstFile(NameSearch, &fdata);
		if (fHandle == INVALID_HANDLE_VALUE) {
			// All empty
			return 0;
		}

		// Got something
		do {
			if (!(fdata.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY )) {
				_stprintf(FoundFile, "%s%s", (LPCTSTR)sSpoolFolder, fdata.cFileName);

				if (m_zip.AddNewFile(FoundFile, 5, true))
					nFiles++;
			}

			if (!::FindNextFile(fHandle, &fdata))
				break;

		} while (TRUE);

		::FindClose(fHandle);

		m_zip.Close();*/

		//return nFiles;

	return 0;
}

CString CUtils::CutMultiLine(CString s)
{
	int n = s.Find('\r');
	if (n != -1)
		return s.Left(n);
	n = s.Find('\n');
	if (n != -1)
		return s.Left(n);

	return s;
}


BOOL CUtils::SendMail(CString sMailSubject, CString sMailBody, CString sMailTo)
{
	CEmailClient email;
	if (sMailTo == _T(""))
		sMailTo = g_prefs.m_mailto;
	email.m_SMTPuseSSL = g_prefs.m_mailuseSSL;
	return email.SendMail(g_prefs.m_mailserver, g_prefs.m_mailport, g_prefs.m_mailusername, g_prefs.m_mailpassword, g_prefs.m_mailfrom, 
		sMailTo, g_prefs.m_mailcc, sMailSubject, sMailBody, false);
}



int CUtils::FreeDiskSpaceMB(CString sPath)
{
	ULARGE_INTEGER	i64FreeBytesToCaller;
	ULARGE_INTEGER	i64TotalBytes;
	ULARGE_INTEGER	i64FreeBytes;

	BOOL fResult = GetDiskFreeSpaceEx(sPath,
		(PULARGE_INTEGER)&i64FreeBytesToCaller,
		(PULARGE_INTEGER)&i64TotalBytes,
		(PULARGE_INTEGER)&i64FreeBytes);

	if (fResult) {
		//		double ft = i64TotalBytes.HighPart; 
		//		ft *= MAXDWORD;
		//		ft += i64TotalBytes.LowPart;

		double ff = i64FreeBytes.HighPart;
		ff *= MAXDWORD;
		ff += i64FreeBytes.LowPart;

		double fMBfree = ff / (1024 * 1024);
		//		TCHAR sz[30];
		//		sprintf(sz, "%d MB free",(DWORD)fMBfree);
		//		double ratio = 100 - 100.0 * ff/ft;

		return (int)fMBfree;
	}

	return -1;

}



int CUtils::GetFolderList(CString sSpoolFolder, CStringArray &sList)
{
	WIN32_FIND_DATA	fdata;
	HANDLE			fHandle;
	TCHAR			NameSearch[MAX_PATH];
	int				nFolders = 0;

	sList.RemoveAll();

	if (CheckFolderWithPing(sSpoolFolder) == FALSE)
		return -1;

	if (sSpoolFolder.Right(1) != _T("\\"))
		sSpoolFolder += _T("\\");

	_stprintf(NameSearch, _T("%s\\*.*"), (LPCTSTR)sSpoolFolder);
	fHandle = ::FindFirstFile(NameSearch, &fdata);
	if (fHandle == INVALID_HANDLE_VALUE) {
		// All empty
		return 0;
	}

	// Got something
	do {

		CString sFolder = fdata.cFileName;

		if ((fdata.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) && sFolder != _T(".") && sFolder != _T("..")) {

			sList.Add(sFolder);
			nFolders++;
		}

		if (!::FindNextFile(fHandle, &fdata))
			break;

	} while (TRUE);

	::FindClose(fHandle);

	return nFolders;
}

BOOL CUtils::Load(CString  csFileName, CString &csDoc)
{
	CFile file;
	if (!file.Open(csFileName, CFile::modeRead))
		return FALSE;
	int nLength = (int)file.GetLength();

#if defined(_UNICODE)
	// Allocate Buffer for UTF-8 file data
	unsigned char* pBuffer = new unsigned char[nLength + 1];
	nLength = file.Read(pBuffer, nLength);
	pBuffer[nLength] = _T('\0');

	// Convert file from UTF-8 to Windows UNICODE (AKA UCS-2)
	int nWideLength = MultiByteToWideChar(CP_UTF8, 0, (const char*)pBuffer, nLength, NULL, 0);
	nLength = MultiByteToWideChar(CP_UTF8, 0, (const char*)pBuffer, nLength,
		csDoc.GetBuffer(nWideLength), nWideLength);
	ASSERT(nLength == nWideLength);
	delete[] pBuffer;
#else
	nLength = file.Read(csDoc.GetBuffer(nLength), nLength);
#endif
	csDoc.ReleaseBuffer(nLength);
	file.Close();

	return TRUE;
}

BOOL CUtils::Save(CString  csFileName, CString csDoc)
{
	int nLength = csDoc.GetLength();
	CFile file;
	if (!file.Open(csFileName, CFile::modeWrite | CFile::modeCreate))
		return FALSE;
#if defined( _UNICODE )
	int nUTF8Len = WideCharToMultiByte(CP_UTF8, 0, csDoc, nLength, NULL, 0, NULL, NULL);
	char* pBuffer = new char[nUTF8Len + 1];
	nLength = WideCharToMultiByte(CP_UTF8, 0, csDoc, nLength, pBuffer, nUTF8Len + 1, NULL, NULL);
	file.Write(pBuffer, nLength);
	delete pBuffer;
#else
	file.Write((LPCTSTR)csDoc, nLength);
#endif
	file.Close();
	return TRUE;
}

CString CUtils::CTime2String(CTime tm)
{
	CString strTime;
	strTime.Format(_T("%.2d-%.2d %.2d:%.2d:%.2d"), tm.GetDay(), tm.GetMonth(), tm.GetHour(), tm.GetMinute(), tm.GetSecond());

	return strTime;
}




CString CUtils::CDate2String(CTime tm)
{
	CString strTime;
	strTime.Format(_T("%.4d-%.2d-%.2d"), tm.GetYear(), tm.GetMonth(), tm.GetDay());

	return strTime;
}

int  CUtils::GetLogFileStatus(CString sFileName, CString &sErrorMessage, BOOL &bHasErrorSeverityHit, CString &sHitErrorMessage)
{
	sErrorMessage = _T("");
	sHitErrorMessage = _T("");

	bHasErrorSeverityHit = FALSE;

	CString sBuffer;
	if (Load(sFileName, sBuffer) == FALSE)
		return -1;

	int n = sBuffer.Find(_T("Hit with severity \"Error\""));
	if (n != -1) {
		bHasErrorSeverityHit = TRUE;

		int m = sBuffer.Find(_T("\n"), n);
		if (m != -1)
			sHitErrorMessage = sBuffer.Mid(n, m - n + 1);

		sHitErrorMessage.Replace(_T("\n"), _T(""));
		sHitErrorMessage.Replace(_T("\r"), _T(""));
	}

	if (sBuffer.Find(_T("DONE")) != -1)
		return TRUE;

	n = sBuffer.Find(_T("ERROR:"));
	if (n != -1) {
		int m = sBuffer.Find(_T("\n"), n);
		if (m != -1)
			sErrorMessage = sBuffer.Mid(n, m - n + 1);

		sErrorMessage.Replace(_T("\n"), _T(""));
		sErrorMessage.Replace(_T("\r"), _T(""));
		return FALSE;
	}

	if (sBuffer.Find(_T("DONE")) != -1)
		return TRUE;

	return -1;
}

int  CUtils::GetLogFileStatus(CString sFileName, CString &sErrorMessage)
{
	sErrorMessage = _T("");

	CString sBuffer;
	if (Load(sFileName, sBuffer) == FALSE)
		return -1;


	if (sBuffer.Find(_T("DONE")) != -1)
		return TRUE;

	int n = sBuffer.Find(_T("ERROR:"));
	if (n != -1) {
		int m = sBuffer.Find(_T("\n"), n);
		if (m != -1)
			sErrorMessage = sBuffer.Mid(n, m - n + 1);

		sErrorMessage.Replace(_T("\n"), _T(""));
		sErrorMessage.Replace(_T("\r"), _T(""));
		return FALSE;
	}

	if (sBuffer.Find(_T("DONE")) != -1)
		return TRUE;

	return -1;
}

CString CUtils::SanitizeFileName(CString sFileName)
{
	sFileName.Replace(_T("%"), _T(""));
	sFileName.Replace(_T(":"), _T(""));
	sFileName.Replace(_T("*"), _T(""));
	sFileName.Replace(_T("?"), _T(""));
	sFileName.Replace(_T("|"), _T(""));
	sFileName.Replace(_T("\""), _T(""));

	return sFileName;
}

int CUtils::GetFileTypeFromExtension(CString sFileName)
{
	sFileName.MakeLower();

	if (sFileName.Find(_T(".pdf")) != -1)
		return FILETYPE_PDF;

	if (sFileName.Find(_T(".tif")) != -1)
		return FILETYPE_TIFF;

	return FILETYPE_OTHER;
}


BOOL CUtils::MoveFile(CString sSource, CString sDest)
{
	::DeleteFile(sDest);

	if (::MoveFileEx(sSource, sDest, MOVEFILE_COPY_ALLOWED | MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH) == FALSE) {
		::Sleep(1000);
		if (::MoveFileEx(sSource, sDest, MOVEFILE_COPY_ALLOWED | MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH) == FALSE) {
			TCHAR szMsg[1024];
			_stprintf(szMsg, _T("ERROR: Unable to move file %s to %s - %s"), (LPCTSTR)sSource, (LPCTSTR)sDest, (LPCTSTR)GetLastWin32Error());
			Logprintf(szMsg);
			return FALSE;
		}
	}
	return TRUE;
}

BOOL CUtils::CopyFile(CString sSource, CString sDest)
{
	if (::CopyFile(sSource, sDest, FALSE) == FALSE) {
		::Sleep(1000);
		if (::CopyFile(sSource, sDest, FALSE) == FALSE) {
			TCHAR szMsg[1024];
			_stprintf(szMsg, _T("ERROR: Unable to copy file %s to %s - %s"), (LPCTSTR)sSource, (LPCTSTR)sDest, (LPCTSTR)GetLastWin32Error());
			Logprintf(szMsg);
			return FALSE;
		}
	}
	return TRUE;
}


CString CUtils::Date2String(CTime tDate, CString sDateFormat)
{
	CString s;
	sDateFormat.MakeUpper();
	CString sDate = sDateFormat;

	s.Format(_T("%.2d"), tDate.GetDay());
	sDate.Replace(_T("DD"), s);
	s.Format(_T("%.2d"), tDate.GetMonth());
	sDate.Replace(_T("MM"), s);
	s.Format(_T("%.4d"), tDate.GetYear());
	sDate.Replace(_T("YYYY"), s);
	s.Format(_T("%.2d"), tDate.GetYear() - 2000);
	sDate.Replace(_T("YY"), s);

	return sDate;
}

void CUtils::KillProcess(DWORD dwProcessID, CString sExeFile)
{
	CString sKillCommand;
	if (dwProcessID > 0)
		sKillCommand.Format(_T("taskkill /PID %d"), dwProcessID);
	else
		sKillCommand.Format(_T("taskkill /F /IM \"%s\""), GetFileName(sExeFile));

	DoCreateProcess(sKillCommand, 120, TRUE);
	::Sleep(10 * 1000);
}


CString CUtils::SanitizeMessage(CString sMsg)
{
	sMsg.Replace(_T("<"), _T("&lt;"));
	sMsg.Replace(_T(">"), _T("&gt;"));
	sMsg.Replace(_T("&"), _T("&amp;"));
	sMsg.Replace(_T("'"), _T("&apos;"));
	sMsg.Replace(_T("\""), _T("&quot;"));
	sMsg.Trim();

	return sMsg;
}

int CUtils::String2Lines(CString sDoc, CStringArray &aLines)
{
	aLines.RemoveAll();
	StringSplitter(sDoc, _T("\n"), aLines);
	for (int i = 0; i < aLines.GetCount(); i++) {
		aLines[i].Replace(_T("\n"), _T(""));
		aLines[i].Replace(_T("\r"), _T(""));
		aLines[i].Trim();
	}

	return aLines.GetCount();

}

CString CUtils::DecodeString(CString s)
{
	CString sOut = s;

	sOut.Replace(_T("&#10;"), _T(" "));
	sOut.Replace(_T("&#13;"), _T(""));
	sOut.Replace(_T("&#9;"), _T("  "));
	sOut.Replace(_T("&quot;"), _T("\""));
	sOut.Replace(_T("Ã¦"), _T("æ"));
	sOut.Replace(_T("Ã¸"), _T("ø"));
	sOut.Replace(_T("Ã¥"), _T("å"));
	sOut.Replace(_T("Ã†"), _T("Æ"));
	sOut.Replace(_T("Ã˜"), _T("Ø"));
	sOut.Replace(_T("Ã…"), _T("Å"));
	sOut.Replace(_T("â€æ"), _T("\""));
	sOut.Replace(_T("â€"), _T("\""));
	sOut.Replace(_T("â‚¬"), _T("€"));
	return sOut;
}



double CUtils::GethighDPIScalingFactor()
{

	//DWORD systemdpi = GetDpiForSystem();
	HDC screen = GetDC(0);
	int systemdpi = GetDeviceCaps(screen, LOGPIXELSX);
	ReleaseDC(0, screen);
	if (systemdpi > 0)
		return (double)systemdpi / 96.0;
	return 1.0;
}

std::wstring CUtils::charToWstring(TCHAR *szStr)
{
#ifdef UNICODE
	std::wstring wstr(szStr);
	return wstr;
#else
	std::wstring wstr(szStr, szStr + _tcslen(szStr));
	return wstr;
#endif
}

std::wstring CUtils::charToWstring(CString sStr)
{

#ifdef UNICODE
	std::wstring wstr(sStr);
#else
	TCHAR szStr[1024];
	_tcscpy(szStr, sStr);
	std::wstring wstr(szStr, szStr + _tcslen(szStr));
#endif

	return wstr;
}

CString  CUtils::WstringTochar(std::wstring wstr)
{
#ifdef UNICODE
	CString cs(wstr.c_str());
	return cs;
#else
	std::string str(CW2A(wstr.c_str()));
	CString cs(str.c_str());
	return cs;
#endif

}

void CUtils::WriteIniFileString(CString sIniFile, CString sSection, CString sElement, CString sValue)
{
	::WritePrivateProfileString((LPCTSTR)sSection, LPCTSTR(sElement), LPCTSTR(sValue), (LPCTSTR)sIniFile);
}

void CUtils::WriteIniFileInt(CString sIniFile, CString sSection, CString sElement, int nValue)
{
	::WritePrivateProfileString((LPCTSTR)sSection, LPCTSTR(sElement), LPCTSTR(Int2String(nValue)), (LPCTSTR)sIniFile);
}

CString CUtils::ReadIniFileString(CString sIniFile, CString sSection, CString sElement, CString sDefaultValue)
{
	TCHAR	Tmp[_MAX_PATH];

	::GetPrivateProfileString((LPCTSTR)sSection, LPCTSTR(sElement), (LPCTSTR)sDefaultValue, Tmp, _MAX_PATH, (LPCTSTR)sIniFile);

	CString s(Tmp);
	return s;
}

int CUtils::ReadIniFileInt(CString sIniFile, CString sSection, CString sElement, int nDefaultValue)
{
	TCHAR	Tmp[_MAX_PATH];

	::GetPrivateProfileString((LPCTSTR)sSection, LPCTSTR(sElement), Int2String(nDefaultValue), Tmp, _MAX_PATH, (LPCTSTR)sIniFile);
	return _tstoi(Tmp);

}

double CUtils::ReadIniFileDouble(CString sIniFile, CString sSection, CString sElement, double fDefaultValue)
{
	TCHAR	Tmp[_MAX_PATH];

	::GetPrivateProfileString((LPCTSTR)sSection, LPCTSTR(sElement), Double2String(fDefaultValue), Tmp, _MAX_PATH, (LPCTSTR)sIniFile);
	return _tstof(Tmp);
}

CString CUtils::GetTempFolder()
{
	TCHAR lpTempPathBuffer[MAX_PATH];

	if (GetTempPath(MAX_PATH, lpTempPathBuffer) == 0)
		return  _T("c:\\temp");
	CString s(lpTempPathBuffer);

	return s;
}


std::wstring CUtils::LoadUtf8FileToString(const std::wstring& filename)
{
	struct _stat fileinfo;
	_wstat(filename.c_str(), &fileinfo);
	size_t filesize = fileinfo.st_size;

	std::wstring buffer;            // stores file contents
	FILE* f = _wfopen(filename.c_str(), L"rtS, ccs=UTF-8");

	// Failed to open file
	if (f == NULL)
	{
		// ...handle some error...
		return buffer;
	}

	// Read entire file contents in to memory
	if (filesize > 0)
	{
		buffer.resize(filesize);
		size_t wchars_read = fread(&(buffer.front()), sizeof(wchar_t), filesize, f);
		buffer.resize(wchars_read);
		buffer.shrink_to_fit();
	}

	fclose(f);

	return buffer;
}

int CUtils::wchar2char(const wchar_t * src, char * dest, size_t dest_len)
{
	size_t i = 0;
	wchar_t code;

	while (src[i] != '\0' && i < (dest_len - 1)) {
		code = src[i];
		if (code < 128)
			dest[i] = char(code);
		else {
			dest[i] = '?';
			if (code >= 0xD800 && code <= 0xD8FF)
				// lead surrogate, skip the next code unit, which is the trail
				i++;
		}
		i++;
	}

	dest[i] = '\0';

	return i - 1;
}


void CUtils::strcpyex(TCHAR *szDst, TCHAR *szSrc, int maxsize)
{
	_tcsncpy(szDst, szSrc, maxsize);
	if (strlen(szSrc) > maxsize)
		*(szDst + maxsize - 1) = 0;
}

void CUtils::strcpyexx(TCHAR *szDst, CString sSrc, int maxsize)
{
	_tcsncpy(szDst, sSrc, maxsize);
	if (sSrc.GetLength() >= maxsize)
		*(szDst + maxsize - 1) = 0;
}

CTime CUtils::Translate2Time(CString sTime)
{
	CTime t(1975, 1, 1, 0, 0, 0);

	if (sTime.GetLength() < 19)
		return t;

	int y = atoi(sTime.Mid(0, 4));
	int m = atoi(sTime.Mid(5, 2));
	int d = atoi(sTime.Mid(8, 2));
	int h = atoi(sTime.Mid(11, 2));
	int mi = atoi(sTime.Mid(14, 2));
	int s = atoi(sTime.Mid(17, 2));
	try {
		CTime t1(y, m, d, h, mi, s);
		t = t1;
	}
	catch (CException *ex)
	{
	}

	return t;
}

void CUtils::Translate2TimeEx(CString sTime, SYSTEMTIME &t)
{

	t.wYear = 1975;
	t.wMonth = 1;
	t.wDay = 1;
	t.wHour = 0;
	t.wMinute = 0;
	t.wSecond = 0;
	t.wMilliseconds = 0;

	//YYYY-MM-DDTHH:NN:SS.ZZZ 
	if (sTime.GetLength() < 19)
		return;

	t.wYear = atoi(sTime.Mid(0, 4));
	t.wMonth = atoi(sTime.Mid(5, 2));
	t.wDay = atoi(sTime.Mid(8, 2));
	t.wHour = atoi(sTime.Mid(11, 2));
	t.wMinute = atoi(sTime.Mid(14, 2));
	t.wSecond = atoi(sTime.Mid(17, 2));
	if (sTime.GetLength() == 23)
		t.wMilliseconds = atoi(sTime.Mid(20, 3));
	else if (sTime.GetLength() == 22)
		t.wMilliseconds = atoi(sTime.Mid(20, 2));


}


CTime CUtils::ParseDateTime(CString sDate, CString sDateFormat)
{
	int n = 0;
	CString sDay = "", sMonth = "", sYear = "";
	CString sHour = "", sMinute = "", sSecond = "";
	CTime today = CTime::GetCurrentTime();
	CTime tErrorDate = CTime(1975, 1, 1, 0, 0, 0);
	CTime tDate = tErrorDate;

	while (n < sDateFormat.GetLength() && n < sDate.GetLength()) {
		switch (sDateFormat[n]) {
		case 'D':
			sDay += sDate[n];
			break;
		case 'M':
			sMonth += sDate[n];
			break;
		case 'Y':
			sYear += sDate[n];
			break;
		case 'h':
			sHour += sDate[n];
			break;
		case 'm':
			sMinute += sDate[n];
			break;
		case 's':
			sSecond += sDate[n];
			break;
		default:
			break;
		}
		n++;
	}
	if (sHour == "" || atoi(sHour) > 23)
		sHour = "0";
	if (sMinute == "" || atoi(sMinute) > 59)
		sMinute = "0";
	if (sSecond == "" || atoi(sSecond) > 59)
		sSecond = "0";

	if (sMonth == "99") {
		sMonth = "";
		sDateFormat = "DD";
	}

	if (sDay == "" || atoi(sDay) == 0) {
		sDay = "01";
	}

	// Month not in date format - default to current month (or roll over to next month if last day in month)	
	if (sMonth == "" || atoi(sMonth) == 0) {

		// Day is after current day
		if (atoi(sDay) >= today.GetDay()) {
			// Day after current date-day (future) - use current month/year
			sMonth.Format("%.2d", today.GetMonth());
			sYear.Format("%.4d", today.GetYear());
		}
		else {

			// Day before current day - use next month
			if (today.GetMonth() < 12) {
				// Same year as now
				sMonth.Format("%.2d", today.GetMonth() + 1);
				sYear.Format("%.4d", today.GetYear());
			}
			else {
				// Next year
				sMonth = "01";
				sYear.Format("%.4d", today.GetYear() + 1);
			}
		}
	}

	// Year not in date format - default to current year (or roll over if 31/12+1)

	if (sYear == "" || atoi(sYear) == 0) {
		try {
			CTime tNewdate = CTime(today.GetYear(), atoi(sMonth), atoi(sDay), 23, 59, 59);
			if (tNewdate >= today) {
				sYear.Format("%.4d", today.GetYear());
			}
			else {
				sYear.Format("%.4d", today.GetYear() + 1);
			}
		}
		catch (CException *ex) {
			return tErrorDate;
		}
	}


	if (atoi(sYear) < 100)
		sYear = "20" + sYear;


	try {
		CTime  tm(atoi(sYear), atoi(sMonth), atoi(sDay), atoi(sHour), atoi(sMinute), atoi(sSecond));
		tDate = tm;
	}
	catch (CException *ex) {
		return tErrorDate;
	}

	return tDate;

}

BOOL CUtils::AddToCStringArray(CStringArray &sa, CString s)
{
	BOOL bFound = FALSE;
	for (int i = 0; i < sa.GetCount(); i++) {
		if (sa[i] == s) {
			bFound = TRUE;
			break;
		}
	}
	if (bFound == FALSE)
		sa.Add(s);

	return bFound == FALSE;
}
