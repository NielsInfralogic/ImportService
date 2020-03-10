#pragma once

class CPPIPrefs : public CObject
{
public:
	CPPIPrefs();
	virtual ~CPPIPrefs();
	void  LoadIniFile(CString sIninFile);
	CTime PPIDateToDateTime(CString sPPIDate);

	CString m_archivefolder;
	CString m_fixededition;
	CString m_fixedsection;
};

