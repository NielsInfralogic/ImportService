#pragma once

#include "Defs.h"
#include <string>

class CUtils
{
public:
	CUtils(void);
	~CUtils(void);

	CString GetLastWin32Error();
	CString LoadResourceString(int nID);
	CString		GenerateTimeStamp(BOOL bIncludemillisecs);

	CString Int2String(int n);
	CString Bigint2String(__int64 n);
	CString Double2String(double f);
	CString Double2String(double f, int decimals);
	std::wstring Double2StringW(double f, int decimals);
	double SafeStringToDouble(CString s);

	CString GetTempFolder();

	void	StripTrailingSpaces(TCHAR *szString, TCHAR *szStringDest);
	int		GetFileAge(CString sFileName);
	int		GetFileAgeMinutes(CString sFileName);
	int		GetFileAgeHours(CString sFileName);

	BOOL    RemoveReadOnlyAttribute(CString sFullPath);
	BOOL	LockCheck(CString sFileToLock, BOOL bSimpleCheck);
	BOOL	SoftLockCheck(CString sFileToLock);
	BOOL	StableTimeCheck(CString sFileToTest, int nStableTimeSec);
	DWORD	GetFileSize(CString sFile);
	CTime	GetWriteTime(CString sFile);
	BOOL	FileExist(CString sFile);
	BOOL	DirectoryExist(CString sDir);
	BOOL	CheckFolder(CString sFolder);
	BOOL	CheckFolderWithPing(CString sFolder);
	BOOL	Reconnect(CString sFolder, CString sUser, CString sPW);
	CString GetExtension(CString sFileName);
	CString GetFileName(CString sFullName);
	CString	GetFileName(CString sFullName, BOOL bExcludeExtension);
	CString	GetFilePath(CString sFullName);
	CString GetChangedExtension(CString sFullName, CString sNewExtension);

	BOOL	IsInArray(CStringArray *sArr, CString sValueToFind);
	BOOL	IsInArray(CUIntArray *nArr, int nValueToFind);
	CString	GetModuleLoadPath();
	CString GetFirstFile(CString sSpoolFolder, CString sSearchMask);
	int		StringSplitter(CString sInput, CString sSeparators, CStringArray &sArr);
	CString	IntArrayToString(int nArr[], int nMaxItems);
	CString	FloatArrayToString(double nArr[], int nMaxItems);
	CString	StringArrayToString(CString nArr[], int nMaxItems);
	CString	GrepFile(CString sFileName, TCHAR *szSearchString);

	BOOL	TryMatchExpression(CString sMatchExpression, CString sInputString, BOOL bPartialMatch);
	BOOL	TryMatchExpressionEx(CString sMatchExpression, CString sInputString, BOOL bPartialMatch, TCHAR *szErrorMessage);
	CString FormatExpressionEx(CString sMatchExpression, CString sFormatExpression, CString sInputString, BOOL bPartialMatch, TCHAR *szErrorMessage, BOOL *bMatched);
	CString FormatExpression(CString sMatchExpression, CString sFormatExpression, CString sInputString, BOOL bPartialMatch);

	void	TruncateLogFile(CString sFile, DWORD dwMaxSize);
	void	Logprintf(const TCHAR *msg, ...);


	int		ScanDirMaxFiles(CString sSpoolFolder, CString sSearchMask, CStringArray &sFoundList, int nMaxFiles, BOOL bIgnoreHiddenFiles);
	int		ScanDirMaxFilesEx(CString sSpoolFolder, CString sSearchMask, FILEINFOSTRUCT aFileList[], int nMaxFiles);
	int		ScanDirMaxFilesEx(CString sSpoolFolder, CString sSearchMask, FILEINFOSTRUCT aFileList[], int nMaxFiles, int nFolderScanOrder, BOOL bIgnoreHiddenFiles);
	int		ScanDirMaxFilesEx(CString sSpoolFolder, CString sSearchMask, FILEINFOSTRUCT aFileList[], int nMaxFiles, CString sIgnoreMask, BOOL bIgnoreHidden, BOOL bKillZeroByteFiles, CString sRegExpSearchMask, int nLockCheckMode);
	int		ScanDirMaxFilesEx(CString sSpoolFolder, CString sSearchMask, FILEINFOSTRUCT aFileList[], int nMaxFiles, int nFolderScanOrder, BOOL bIgnoreHiddenFiles, BOOL bGetLogInfo);

	int		ScanDirCount(CString sSpoolFolder, CString sSearchMask);
	int		ScanDirCount(CString sSpoolFolder, CString sSearchMask, DWORD &dwArchived, BOOL bIgnoreHiddenFiles);
	int		ScanSubDirNames(CString sSpoolFolder, CStringArray &sFoundList, int nMaxFolders);
	int		ScanDirCountWithSubfolder(CString sSpoolFolder, CString sSearchMask);

	int		DirectoryList(CString sSpoolFolder, FILEINFOSTRUCT aFileList[], int nMaxDirs);
	CString	Time2String(CTime tm);
	CTime	String2Time(CString s);
	BOOL	DoCreateProcessEx(CString sCmdLine, int nTimeout);
	int		DoCreateProcess(CString sCmdLine);
	int     DoCreateProcess(CString sCmdLine, int nTimeout, BOOL bBlocking);
	int     DoCreateProcess(CString sCmdLine, int nTimeout, BOOL bBlocking, DWORD &dwExitCode, DWORD &dwProcessID);
	BOOL	CreateProcessGetTextResult(CString sCmdLine, int nTimeout, CString &strResult, DWORD &dwExitCode, DWORD &dwProcessID);

	void	KillProcess(DWORD dwProcessID, CString sExeFile);

	int		DeleteFiles(CString sSpoolFolder, CString sSearchMask);
	int		DeleteFiles(CString sSpoolFolder, CString sSearchMask, int sMaxAgeHours);
	int		CopyFiles(CString sSpoolFolder, CString sSearchMask, CString sDestFolder);
	int		CRC32(CString sFileName);
	BOOL	SetFileToCurrentTime(CString sFile);

	UINT	GetOwner(LPCTSTR szFileOrFolderPathName, LPTSTR pUserNameBuffer, int nSizeInBytes);
	BOOL	SendMail(CString sMailSubject, CString sMailBody, CString sMailTo);


	int		ZipAllFiles(CString sSpoolFolder, CString sSearchMask, CString sZipName);
	CString CutMultiLine(CString s);

	int		FreeDiskSpaceMB(CString sPath);


	int		GetFolderList(CString sSpoolFolder, CStringArray &sList);


	BOOL	Load(CString  csFileName, CString &csDoc);
	BOOL	Save(CString  csFileName, CString csDoc);

	CString CTime2String(CTime tm);
	CString CDate2String(CTime tm);

	CString	SanitizeFileName(CString sFileName);
	CString		SanitizeMessage(CString sMsg);
	int		GetFileTypeFromExtension(CString sFileName);

	BOOL MoveFile(CString sSource, CString sDest);
	BOOL CopyFile(CString sSource, CString sDest);

	CString Date2String(CTime tDate, CString sDateFormat);

	BOOL	DeleteEmptyDirectory(CString sFolder);

	int		GetLogFileStatus(CString sFileName, CString &sErrorMessage);
	int		GetLogFileStatus(CString sFileName, CString &sErrorMessage, BOOL &bHasErrorSeverityHit, CString &sHitErrorMessage);

	int		String2Lines(CString sDoc, CStringArray &aLines);

	CString DecodeString(CString s);


	double	GethighDPIScalingFactor();


	std::wstring charToWstring(TCHAR *szStr);
	int wchar2char(const wchar_t * src, char * dest, size_t dest_len);

	std::wstring charToWstring(CString sStr);
	CString WstringTochar(std::wstring wstr);

	void	WriteIniFileString(CString sIniFile, CString sSection, CString sElement, CString sValue);
	void	WriteIniFileInt(CString sIniFile, CString sSection, CString sElement, int nValue);
	CString ReadIniFileString(CString sIniFile, CString sSection, CString sElement, CString sDefaultValue);
	int ReadIniFileInt(CString sIniFile, CString sSection, CString sElement, int nDefaultValue);
	double ReadIniFileDouble(CString sIniFile, CString sSection, CString sElement, double fDefaultValue);
	std::wstring LoadUtf8FileToString(const std::wstring& filename);

	void strcpyex(TCHAR *szDst, TCHAR *szSrc, int maxsize);
	void strcpyexx(TCHAR *szDst, CString sSrc, int maxsize);

	CTime Translate2Time(CString sTime);
	void Translate2TimeEx(CString sTime, SYSTEMTIME &t);

	CTime	ParseDateTime(CString sDate, CString sDateFormat);
	BOOL	AddToCStringArray(CStringArray &sa, CString s);

	CString GetComputerName();
};



