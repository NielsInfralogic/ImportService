#include "stdafx.h"
#include "Defs.h"
#include "PlanDataDefs.h"
#include "Prefs.h"
#include "Utils.h"
#include "Markup.h"
#include "ImportService.h"
#include "Registry.h"
#include "PPIPrefs.h"

CPPIPrefs::CPPIPrefs()
{
}

CPPIPrefs::~CPPIPrefs()
{

}

void CPPIPrefs::LoadIniFile(CString sIniFile)
{

	::GetPrivateProfileString("Setup", "ArchiveFolder", "", m_archivefolder.GetBuffer(255), 255, sIniFile);
	m_archivefolder.ReleaseBuffer();

	::GetPrivateProfileString("Setup", "FixedEdition", "1_1", m_fixededition.GetBuffer(255), 255, sIniFile);
	m_fixededition.ReleaseBuffer();

	::GetPrivateProfileString("Setup", "FixedSection", "1", m_fixededition.GetBuffer(255), 255, sIniFile);
	m_fixedsection.ReleaseBuffer();

}

CTime CPPIPrefs::PPIDateToDateTime(CString sPPIDate)
{
	if (sPPIDate.GetLength() != 8)
		return 	CTime(1975, 1, 1);
	

	return CTime(_tstoi(sPPIDate.Mid(0, 4)), _tstoi(sPPIDate.Mid(4, 2)), _tstoi(sPPIDate.Mid(6, 2)));
}
