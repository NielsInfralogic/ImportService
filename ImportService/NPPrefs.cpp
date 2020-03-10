#include "stdafx.h"
#include "Defs.h"
#include "Prefs.h"
#include "Utils.h"
#include "Markup.h"
#include "ImportService.h"
#include "Registry.h"
#include "NPPrefs.h"

CNPPrefs::CNPPrefs()
{
}

CNPPrefs::~CNPPrefs()
{

}

void CNPPrefs::LoadIniFile(CString sIniFile)
{
	TCHAR Tmp[MAX_PATH], Tmp2[MAX_PATH];
	int i = 0;

	sprintf(Tmp2, "ArchiveFolder%d", i + 1);
	GetPrivateProfileString("Setup", Tmp2, "", m_archivefolder.GetBuffer(255), 255, sIniFile);
	m_archivefolder.ReleaseBuffer();


	sprintf(Tmp2, "HeaderContents%d", i + 1);
	GetPrivateProfileString("Setup", Tmp2, "%P,%D", m_headercontents.GetBuffer(255), 255, sIniFile);
	m_headercontents.ReleaseBuffer();

	sprintf(Tmp2, "SubHeaderContents%d", i + 1);
	GetPrivateProfileString("Setup", Tmp2, "%E", m_subheadercontents.GetBuffer(255), 255, sIniFile);
	m_subheadercontents.ReleaseBuffer();

	sprintf(Tmp2, "SubHeaderTrigger%d", i + 1);
	GetPrivateProfileString("Setup", Tmp2, "", m_subheadertriggerexpr.GetBuffer(255), 255, sIniFile);
	m_subheadertriggerexpr.ReleaseBuffer();

	sprintf(Tmp2, "DataRowContents%d", i + 1);
	GetPrivateProfileString("Setup", Tmp2, "%N", m_datarowcontents.GetBuffer(255), 255, sIniFile);
	m_datarowcontents.ReleaseBuffer();

	sprintf(Tmp2, "MatchExpressionHeader%d", i + 1);
	GetPrivateProfileString("Setup", Tmp2, "", m_matchexpressionheader.GetBuffer(255), 255, sIniFile);
	m_matchexpressionheader.ReleaseBuffer();

	sprintf(Tmp2, "FormatExpressionHeader%d", i + 1);
	GetPrivateProfileString("Setup", Tmp2, "", m_formatexpressionheader.GetBuffer(255), 255, sIniFile);
	m_formatexpressionheader.ReleaseBuffer();

	sprintf(Tmp2, "MatchExpressionTitle%d", i + 1);
	GetPrivateProfileString("Setup", Tmp2, "", m_matchexpressiontop.GetBuffer(255), 255, sIniFile);
	m_matchexpressiontop.ReleaseBuffer();

	sprintf(Tmp2, "FormatExpressionTitle%d", i + 1);
	GetPrivateProfileString("Setup", Tmp2, "", m_formatexpressiontop.GetBuffer(255), 255, sIniFile);
	m_formatexpressiontop.ReleaseBuffer();

	sprintf(Tmp2, "MatchExpressionPreparse%d", i + 1);
	GetPrivateProfileString("Setup", Tmp2, "", m_preparsematchexpression.GetBuffer(255), 255, sIniFile);
	m_preparsematchexpression.ReleaseBuffer();

	sprintf(Tmp2, "FormatExpressionPreparse%d", i + 1);
	GetPrivateProfileString("Setup", Tmp2, "", m_preparseformatexpression.GetBuffer(255), 255, sIniFile);
	m_preparseformatexpression.ReleaseBuffer();


	sprintf(Tmp2, "NumberOfExpressions%d", i + 1);
	GetPrivateProfileString("Setup", Tmp2, "0", Tmp, 255, sIniFile);
	m_numberofexpressions = atoi(Tmp);


	for (int j = 0; j < m_numberofexpressions; j++) {

		sprintf(Tmp2, "MatchExpressionBody%d_%d", i + 1, j + 1);
		GetPrivateProfileString("Setup", Tmp2, "", m_matchexpressionbody[j].GetBuffer(255), 255, sIniFile);
		m_matchexpressionbody[j].ReleaseBuffer();

		sprintf(Tmp2, "FormatExpressionBody%d_%d", i + 1, j + 1);
		GetPrivateProfileString("Setup", Tmp2, "", m_formatexpressionbody[j].GetBuffer(255), 255, sIniFile);
		m_formatexpressionbody[j].ReleaseBuffer();

	}

	sprintf(Tmp2, "DateFormat%d", i + 1);
	GetPrivateProfileString("Setup", Tmp2, "YYYYMMDD", m_dateformat.GetBuffer(255), 255, sIniFile);
	m_dateformat.ReleaseBuffer();

	sprintf(Tmp2, "TimeFormat%d", i + 1);
	GetPrivateProfileString("Setup", Tmp2, "YYYYMMDDhhmmss", m_timeformat.GetBuffer(255), 255, sIniFile);
	m_timeformat.ReleaseBuffer();

	sprintf(Tmp2, "DefaultLocation%d", i + 1);
	GetPrivateProfileString("Setup", Tmp2, "Default", m_defaultlocation.GetBuffer(255), 255, sIniFile);
	m_defaultlocation.ReleaseBuffer();

	sprintf(Tmp2, "DefaultZone%d", i + 1);
	GetPrivateProfileString("Setup", Tmp2, "1", m_defaultzone.GetBuffer(255), 255, sIniFile);
	m_defaultzone.ReleaseBuffer();

	sprintf(Tmp2, "DefaultEdition%d", i + 1);
	GetPrivateProfileString("Setup", Tmp2, "1", m_defaultedition.GetBuffer(255), 255, sIniFile);
	m_defaultedition.ReleaseBuffer();

	sprintf(Tmp2, "FilenameFilterExpression%d", i + 1);
	GetPrivateProfileString("Setup", Tmp2, "1", m_filenamefilterexpression.GetBuffer(255), 255, sIniFile);
	m_filenamefilterexpression.ReleaseBuffer();

	sprintf(Tmp2, "MatchExpressionFilename%d", i + 1);
	GetPrivateProfileString("Setup", Tmp2, "", m_matchexpressionfilename.GetBuffer(255), 255, sIniFile);
	m_matchexpressionfilename.ReleaseBuffer();

	sprintf(Tmp2, "FormatExpressionFilename%d", i + 1);
	GetPrivateProfileString("Setup", Tmp2, "", m_formatexpressionfilename.GetBuffer(255), 255, sIniFile);
	m_formatexpressionfilename.ReleaseBuffer();

	sprintf(Tmp2, "FilenameContents%d", i + 1);
	GetPrivateProfileString("Setup", Tmp2, "", m_filenamecontents.GetBuffer(255), 255, sIniFile);
	m_filenamecontents.ReleaseBuffer();

	sprintf(Tmp2, "NumberOfZoneLocations%d", i + 1);
	GetPrivateProfileString("Setup", Tmp2, "0", Tmp, 255, sIniFile);
	m_numberofzonelocations = atoi(Tmp);

	for (int zn = 0; zn < m_numberofzonelocations; zn++) {
		sprintf(Tmp2, "ZoneLocation%d_%d", i + 1, zn + 1);
		GetPrivateProfileString("Setup", Tmp2, "", Tmp, 255, sIniFile);
		CString s = Tmp;
		CString sZone, sLocation;
		int m = s.Find(";");
		if (m != -1) {
			sLocation = s.Mid(m + 1);
			sZone = s.Left(m);
		}

		m_zonelocations[zn].m_zone = sZone;
		m_zonelocations[zn].m_location = sLocation;

	}

	GetPrivateProfileString("Setup", "ZeroExtend", "1", Tmp, 255, sIniFile);
	m_zeroextend = atoi(Tmp);

	GetPrivateProfileString("Setup", "MultipleRegexTests", "2", Tmp, 255, sIniFile);
	m_multipleregextests = atoi(Tmp);

	GetPrivateProfileString("Setup", "PagenamePrefixFromComment", "0", Tmp, 255, sIniFile);
	m_pagenameprefixfromcomment = atoi(Tmp);

	GetPrivateProfileString("Setup", "DayeToKeepImportData", "2", Tmp, 255, sIniFile);
	m_daystokeepdata = atoi(Tmp);


	GetPrivateProfileString("ExtraEditions", "NumberOfEntries", "0", Tmp, 255, sIniFile);
	int n = atoi(Tmp);

	m_ExtraEditionList.RemoveAll();

	for (int i = 0; i < n; i++) {
		sprintf(Tmp2, "Publication%d", i + 1);
		GetPrivateProfileString("ExtraEditions", Tmp2, "", Tmp, 255, sIniFile);
		CString s(Tmp);

		sprintf(Tmp2, "MinEditions%d", i + 1);
		GetPrivateProfileString("ExtraEditions", Tmp2, "", Tmp, 255, sIniFile);

		NPEXTRAEDITION *pItem = new NPEXTRAEDITION();
		pItem->m_Publication = s;
		pItem->m_mineditions = Tmp;
		m_ExtraEditionList.Add(*pItem);
	}
}