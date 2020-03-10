
#include "stdafx.h"
#include "Defs.h"
#include "Prefs.h"
#include "Utils.h"
#include "Markup.h"
#include "ImportService.h"
#include "Registry.h"

extern CUtils g_util;

CPrefs::CPrefs()
{
	
	m_publicationlock = FALSE;
	m_title = "";
	m_getpubfromDB = FALSE;

	m_XMLLogFile = m_apppath + _T("\\Log.xml");
	m_lockFile = m_apppath + _T("\\Running");

	m_lockHandle = ::CreateFile(m_lockFile, 0, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);

	m_searchmask = _T("*.*");
	m_stabletime = 1;
	m_polltime = 2;
	m_bSortOnCreateTime = TRUE;
	m_useretryservice = TRUE;

	m_serverName = g_util.GetComputerName();
	m_serverShare = _T("CCData");
	m_storagePath = _T("\\\\") + m_serverName + _T("\\") + m_serverShare + _T("\\CCfilesHires");
	m_previewPath = _T("\\\\") + m_serverName + _T("\\") + m_serverShare + _T("\\CCpreviews");
	m_thumbnailPath = _T("\\\\") + m_serverName + _T("\\") + m_serverShare + _T("\\CCthumbnails");
	m_logFilePath = "c:\\temp";
	m_configPath = _T("\\\\") + m_serverName + _T("\\") + m_serverShare + _T("\\CCconfig");
	m_errorPath = _T("\\\\") + m_serverName + _T("\\") + m_serverShare + _T("\\CCerrorfiles");
	m_MainLocationID = 1;
	m_copyToWeb = FALSE;
	m_copyWithFtp = FALSE;
	m_webPath = _T("");
	m_webFTPserver = _T("");
	m_webFTPfolder = _T("");
	m_webFTPuser = _T("");
	m_webFTPpassword = _T("");
	m_webFTPport = 21;

	m_doWarn = TRUE;
	m_logToFile = TRUE;

	m_username = _T("admin");
	m_locationname = _T("Default");

	m_QueryBackoffTime = 1000;	// ms
	m_nQueryRetries = 2;

	m_customizedmarkgroup = 0;

	m_nologin = FALSE;
	m_isadmin = FALSE;
	m_tLastLogin = CTime::GetCurrentTime();

	m_runpostprocedure3 = FALSE;
	m_runpostprocedureProduction2 = FALSE;
}

CPrefs::~CPrefs()
{
	::CloseHandle(m_lockHandle);
	DeleteFile(m_lockFile);
}

int CPrefs::GetProofID(CString s)
{
	for (int i = 0; i < m_ProofList.GetCount(); i++) {
		if (s.CompareNoCase(m_ProofList[i].m_name) == 0)
			return m_ProofList[i].m_ID;
	}
	return 0;
}

CString CPrefs::GetProofName(int nID)
{
	for (int i = 0; i < m_ProofList.GetCount(); i++) {
		if (m_ProofList[i].m_ID == nID)
			return m_ProofList[i].m_name;
	}
	return _T("");
}

int CPrefs::GetPageformatFromPublication(int nPublicationID)
{
	for (int i = 0; i < m_PublicationList.GetCount(); i++) {
		if (nPublicationID == m_PublicationList[i].m_ID) {
			return m_PublicationList[i].m_PageFormatID;
		}
	}

	return 0;
}

void CPrefs::FlushPublicationName(CString sName)
{
	for (int i = 0; i < m_PublicationList.GetCount(); i++) {
		if (sName.CompareNoCase(m_PublicationList[i].m_name) == 0) {
			m_PublicationList.RemoveAt(i);
			return;
		}
	}
}

int CPrefs::GetPublicationID(CString s)
{
	CString sErrorMessage;
	if (m_getpubfromDB) {
		int nID = m_DB.GetPublicationID(s, sErrorMessage);
		// Known ID?
	/*	BOOL existingID = FALSE;
		for (int i=0; i<m_PublicationList.GetCount(); i++) {
			if (s.CompareNoCase(m_PublicationList[i].m_name) == 0 && m_PublicationList[i].m_ID == nID) {
				existingID = TRUE;
				break;
			}
		}

		if (existingID == FALSE) {
			;
		}
*/
		return nID;
	}
	else {

		for (int i = 0; i < m_PublicationList.GetCount(); i++) {
			if (s.CompareNoCase(m_PublicationList[i].m_name) == 0 || s.CompareNoCase(m_PublicationList[i].m_alias) == 0)
				return m_PublicationList[i].m_ID;
		}

		return m_DB.LoadNewPublicationID(s, sErrorMessage);
	}
}

BOOL CPrefs::ChangePublicationName(CString sLongName, CString sAbbr)
{
	CString sErrorMessage;
	if (m_DB.UpdatePublicationName(sAbbr, sLongName, sErrorMessage) == FALSE)
		return FALSE;
	if (sAbbr != sLongName)
		m_DB.UpdateAliasName("Publication", sLongName, sAbbr, sErrorMessage);

	return TRUE;
}

int CPrefs::AddNewPublication(CString sLongName, CString sAbbr)
{
	CString sErrorMessage;
	int nID = m_DB.RetrieveMaxValueCount("PublicationID", "PublicationNames", sErrorMessage);
	if (nID < 0)
		nID = 0;
	if (!m_DB.InsertPublicationNames(nID + 1, sLongName, 0, FALSE, 0.0, GetProofID(m_proofer), 0, m_defaultapprovalnewpublication, sErrorMessage)) {
		return 0;
	}

	m_DB.DeleteInputAlias("Publication", sLongName, sErrorMessage);

	if (sAbbr != sLongName)
		m_DB.UpdateAliasName("Publication", sLongName, sAbbr, sErrorMessage);

	PUBLICATIONSTRUCT *item = new PUBLICATIONSTRUCT();
	item->m_ID = nID + 1;
	item->m_name = sLongName;
	item->m_PageFormatID = 0;

	m_PublicationList.Add(*item);
	return nID + 1;
}

int CPrefs::AddNewEdition(CString s, CString sAbbr)
{
	CString sErrorMessage;
	int nID = m_DB.RetrieveMaxValueCount("EditionID", "EditionNames", sErrorMessage);
	if (nID < 0)
		nID = 0;
	if (!m_DB.InsertEditionName(nID + 1, s, 1, 1, sErrorMessage)) {
		return 0;
	}
	EDITIONSTRUCT *item = new EDITIONSTRUCT();
	item->m_ID = nID + 1;
	item->m_name = s;
	item->m_iscommon = TRUE;
	item->m_level = 1;
	item->m_subofeditionID = 1;
	m_EditionList.Add(*item);
	return nID + 1;
}

int CPrefs::AddNewSection(CString s, CString sAbbr)
{
	CString sErrorMessage;
	int nID = m_DB.RetrieveMaxValueCount("SectionID", "SectionNames", sErrorMessage);
	if (nID < 0)
		nID = 0;
	if (!m_DB.InsertDBName("SectionNames", "SectionID", "Name", nID + 1, s, sErrorMessage)) {
		return 0;
	}
	ITEMSTRUCT *item = new ITEMSTRUCT();
	item->m_ID = nID + 1;
	item->m_name = s;
	m_SectionList.Add(*item);
	return nID + 1;
}


int CPrefs::GetSectionID(CString s)
{
	CString sErrorMessage;
	if (m_getpubfromDB) {
		return m_DB.GetSectionID(s, sErrorMessage);
	}
	else {
		for (int i = 0; i < m_SectionList.GetCount(); i++) {
			if (s.CompareNoCase(m_SectionList[i].m_name) == 0)
				return m_SectionList[i].m_ID;
		}

		return m_DB.LoadNewSectionID(s, sErrorMessage);
	}

	/*	if (m_allowunknownnames) {

			TCHAR szDbErrMsg[1024];
			int nID = m_DB.RetrieveMaxValueCount("SectionID", "SectionNames", szDbErrMsg);
			if (nID<0)
				nID = 0;
			if (!m_DB.InsertDBName("SectionNames", "SectionID", "Name",  nID+1, s, szDbErrMsg )) {
				return 0;
			}
			ITEMSTRUCT *item = new ITEMSTRUCT();
			item->m_ID = nID+1;
			item->m_name = s;
			m_SectionList.Add(*item);
			return nID+1;

		}
		return 0;	*/
}

int CPrefs::GetEditionID(CString s)
{
	CString sErrorMessage;
	if (m_getpubfromDB) {
		return m_DB.GetEditionID(s, sErrorMessage);
	}
	else {
		for (int i = 0; i < m_EditionList.GetCount(); i++) {
			if (s.CompareNoCase(m_EditionList[i].m_name) == 0)
				return m_EditionList[i].m_ID;
		}
		return m_DB.LoadNewEditionID(s, sErrorMessage);
	}

	/*
		if (m_allowunknownnames) {

			TCHAR szDbErrMsg[1024];
			int nID = m_DB.RetrieveMaxValueCount("EditionID", "EditionNames", szDbErrMsg);
			if (nID<0)
				nID = 0;
			if (!m_DB.InsertEditionName( nID+1, s, 1, 1, szDbErrMsg )) {
				return 0;
			}
			EDITIONSTRUCT *item = new EDITIONSTRUCT();
			item->m_ID = nID+1;
			item->m_name = s;
			item->m_iscommon = TRUE;
			item->m_level = 1;
			item->m_subofeditionID = 1;
			m_EditionList.Add(*item);
			return nID+1;
		}
		return 0;	*/
}

int CPrefs::GetIssueID(CString s)
{
	for (int i = 0; i < m_IssueList.GetCount(); i++) {
		if (s.CompareNoCase(m_IssueList[i].m_name) == 0)
			return m_IssueList[i].m_ID;
	}
	if (m_allowunknownnames) {

		CString sErrorMessage;
		int nID = m_DB.RetrieveMaxValueCount("IssueID", "IssueNames", sErrorMessage);
		if (nID < 0)
			nID = 0;
		if (!m_DB.InsertDBName("IssueNames", "IssueID", "Name", nID + 1, s, sErrorMessage)) {
			return 0;
		}
		ITEMSTRUCT *item = new ITEMSTRUCT();
		item->m_ID = nID + 1;
		item->m_name = s;
		m_IssueList.Add(*item);
		return nID + 1;

	}

	return 0;
}

int CPrefs::GetTemplateID(CString s)
{
	for (int i = 0; i < m_TemplateList.GetCount(); i++) {
		if (s.CompareNoCase(m_TemplateList[i].m_name) == 0)
			return m_TemplateList[i].m_ID;
	}

	CString sErrorMessage;
	m_DB.LoadTemplateList(sErrorMessage);

	for (int i = 0; i < m_TemplateList.GetCount(); i++) {
		if (s.CompareNoCase(m_TemplateList[i].m_name) == 0)
			return m_TemplateList[i].m_ID;
	}

	return 0;
}


int CPrefs::GetTemplateIndex(CString s)
{
	for (int i = 0; i < m_TemplateList.GetCount(); i++) {
		if (s.CompareNoCase(m_TemplateList[i].m_name) == 0)
			return i;
	}
	return 0;
}

int CPrefs::GetTemplatePageRotation(int nID)
{
	for (int i = 0; i < m_TemplateList.GetCount(); i++) {
		if (m_TemplateList[i].m_ID == nID) {
			return m_TemplateList[i].m_incomingrotationeven;
		}
	}
	return 0;
}

int CPrefs::GetPressIDFromTemplateID(int nID, int *nPagesAcross, int *nPagesDown)
{
	for (int i = 0; i < m_TemplateList.GetCount(); i++) {
		if (m_TemplateList[i].m_ID == nID) {
			*nPagesAcross = m_TemplateList[i].m_pagesacross;
			*nPagesDown = m_TemplateList[i].m_pagesdown;
			return m_TemplateList[i].m_pressID;
		}
	}
	return 0;
}

int CPrefs::GetPagesOnPlateFromTemplateID(int nID)
{
	for (int i = 0; i < m_TemplateList.GetCount(); i++) {
		if (m_TemplateList[i].m_ID == nID) {
			return m_TemplateList[i].m_pagesacross *  m_TemplateList[i].m_pagesdown;
		}
	}
	return 0;
}

int CPrefs::GetPressIDFromTemplateID(int nID)
{
	for (int i = 0; i < m_TemplateList.GetCount(); i++) {
		if (m_TemplateList[i].m_ID == nID) {
			return m_TemplateList[i].m_pressID;
		}
	}
	return 0;
}


PRESSSTRUCT *CPrefs::GetPressStruct(int nID)
{
	for (int i = 0; i < m_PressList.GetCount(); i++) {
		if (m_PressList[i].m_ID == nID)
			return &m_PressList[i];
	}
	return NULL;
}

int CPrefs::GetDeviceID(CString s)
{
	for (int i = 0; i < m_DeviceList.GetCount(); i++) {
		if (s.CompareNoCase(m_DeviceList[i].m_name) == 0)
			return m_DeviceList[i].m_ID;
	}
	CString sErrorMessage;
	m_DB.LoadDeviceList(sErrorMessage);
	for (int i = 0; i < m_DeviceList.GetCount(); i++) {
		if (s.CompareNoCase(m_DeviceList[i].m_name) == 0)
			return m_DeviceList[i].m_ID;
	}

	return 0;
}

int CPrefs::GetPressID(CString s)
{
	for (int i = 0; i < m_PressList.GetCount(); i++) {
		if (s.CompareNoCase(m_PressList[i].m_name) == 0)
			return m_PressList[i].m_ID;
	}
	CString sErrorMessage;
	m_DB.LoadPressList(sErrorMessage);
	for (int i = 0; i < m_PressList.GetCount(); i++) {
		if (s.CompareNoCase(m_PressList[i].m_name) == 0)
			return m_PressList[i].m_ID;
	}
	return 0;
}

int CPrefs::GetPressIndex(int nID)
{
	for (int i = 0; i < m_PressList.GetCount(); i++) {
		if (m_PressList[i].m_ID == nID)
			return i;
	}
	CString sErrorMessage;
	m_DB.LoadPressList(sErrorMessage);
	for (int i = 0; i < m_PressList.GetCount(); i++) {
		if (m_PressList[i].m_ID == nID)
			return i;
	}
	return -1;
}

CString CPrefs::GetDeviceName(int nID)
{
	for (int i = 0; i < m_DeviceList.GetCount(); i++) {
		if (m_DeviceList[i].m_ID == nID)
			return m_DeviceList[i].m_name;
	}
	return _T("");
}

CString CPrefs::GetTemplateName(int nID)
{
	for (int i = 0; i < m_TemplateList.GetCount(); i++) {
		if (m_TemplateList[i].m_ID == nID)
			return m_TemplateList[i].m_name;
	}
	return _T("");
}

CString CPrefs::GetPublicationName(int nID)
{
	CString sErrorMessage;

	for (int i = 0; i < m_PublicationList.GetCount(); i++) {
		if (m_PublicationList[i].m_ID == nID)
			return m_PublicationList[i].m_name;
	}

	return m_DB.LoadNewPublicationName(nID, sErrorMessage);
}

PUBLICATIONSTRUCT *CPrefs::GetPublicationStruct(int nID)
{
	CString sErrorMessage;

	if (m_getpubfromDB)
		return m_DB.LoadPublicationStruct(nID, sErrorMessage);
	else {
		for (int i = 0; i < m_PublicationList.GetCount(); i++) {
			if (m_PublicationList[i].m_ID == nID)
				return &m_PublicationList[i];
		}

		if (m_DB.LoadNewPublicationName(nID, sErrorMessage) > 0) {
			for (int i = 0; i < m_PublicationList.GetCount(); i++) {
				if (m_PublicationList[i].m_ID == nID)
					return &m_PublicationList[i];
			}
		}
	}
	return NULL;
}

PUBLICATIONPRESSDEFAULTSSTRUCT *CPrefs::GetPublicationPressDefaultStruct(int nPublicationID, int nPressID)
{
	PUBLICATIONSTRUCT *pPublication = GetPublicationStruct(nPublicationID);
	if (pPublication == NULL)
		return NULL;
	for (int i = 0; i < MAXPRESSES; i++) {
		if (pPublication->m_PressDefaults[i].m_PressID == nPressID)
			return &pPublication->m_PressDefaults[i];
	}

	return NULL;

}

CString CPrefs::GetSectionName(int nID)
{
	for (int i = 0; i < m_SectionList.GetCount(); i++) {
		if (m_SectionList[i].m_ID == nID)
			return m_SectionList[i].m_name;
	}
	return _T("");
}

CString CPrefs::GetEditionName(int nID)
{
	for (int i = 0; i < m_EditionList.GetCount(); i++) {
		if (m_EditionList[i].m_ID == nID)
			return m_EditionList[i].m_name;
	}

	return _T("");
}

CString CPrefs::GetIssueName(int nID)
{
	for (int i = 0; i < m_IssueList.GetCount(); i++) {
		if (m_IssueList[i].m_ID == nID)
			return m_IssueList[i].m_name;
	}
	return _T("");
}

CString CPrefs::GetPressName(int nID)
{
	for (int i = 0; i < m_PressList.GetCount(); i++) {
		if (m_PressList[i].m_ID == nID)
			return m_PressList[i].m_name;
	}
	return _T("");
}

int CPrefs::GetPressLocationIDFromID(int nID)
{
	for (int i = 0; i < m_PressList.GetCount(); i++) {
		if (m_PressList[i].m_ID == nID)
			return m_PressList[i].m_locationID;
	}
	return 0;
}

int CPrefs::GetMarkGroupID(CString sName)
{
	for (int i = 0; i < m_MarkGroupList.GetCount(); i++) {
		if (m_MarkGroupList[i].m_name == sName)
			return m_MarkGroupList[i].m_ID;
	}
	return 0;
}

CString CPrefs::GetMarkGroupName(int nID)
{
	for (int i = 0; i < m_MarkGroupList.GetCount(); i++) {
		if (m_MarkGroupList[i].m_ID == nID)
			return m_MarkGroupList[i].m_name;
	}
	return "";
}

int CPrefs::GetColorID(CString s)
{
	for (int i = 0; i < m_ColorList.GetCount(); i++) {
		if (s.CompareNoCase(m_ColorList[i].m_name) == 0)
			return m_ColorList[i].m_ID;
	}
	CString sErrorMessage;
	m_DB.LoadColorList(sErrorMessage);
	for (int i = 0; i < m_ColorList.GetCount(); i++) {
		if (s.CompareNoCase(m_ColorList[i].m_name) == 0)
			return m_ColorList[i].m_ID;
	}
	return 0;
}

CString CPrefs::GetColorName(int nID)
{
	for (int i = 0; i < m_ColorList.GetCount(); i++) {
		if (m_ColorList[i].m_ID == nID)
			return m_ColorList[i].m_name;
	}
	CString sErrorMessage;
	m_DB.LoadColorList(sErrorMessage);
	for (int i = 0; i < m_ColorList.GetCount(); i++) {
		if (m_ColorList[i].m_ID == nID)
			return m_ColorList[i].m_name;
	}
	return "";
}


BOOL CPrefs::IsBlackColor(int nID)
{
	for (int i = 0; i < m_ColorList.GetCount(); i++) {
		if (m_ColorList[i].m_ID == nID) {
			if (m_ColorList[i].m_k > 0 && m_ColorList[i].m_c == 0 && m_ColorList[i].m_m == 0 && m_ColorList[i].m_y == 0)
				return TRUE;
			else
				return FALSE;
		}
	}
	return FALSE;
}

BOOL CPrefs::IsCyanColor(int nID)
{
	for (int i = 0; i < m_ColorList.GetCount(); i++) {
		if (m_ColorList[i].m_ID == nID) {
			if (m_ColorList[i].m_c > 0 && m_ColorList[i].m_m == 0 && m_ColorList[i].m_y == 0 && m_ColorList[i].m_k == 0)
				return TRUE;
			else
				return FALSE;
		}
	}
	return FALSE;
}

BOOL CPrefs::IsMagentaColor(int nID)
{
	for (int i = 0; i < m_ColorList.GetCount(); i++) {
		if (m_ColorList[i].m_ID == nID) {
			if (m_ColorList[i].m_m > 0 && m_ColorList[i].m_c == 0 && m_ColorList[i].m_y == 0 && m_ColorList[i].m_k == 0)
				return TRUE;
			else
				return FALSE;
		}
	}
	return FALSE;
}

BOOL CPrefs::IsYellowColor(int nID)
{
	for (int i = 0; i < m_ColorList.GetCount(); i++) {
		if (m_ColorList[i].m_ID == nID) {
			if (m_ColorList[i].m_y > 0 && m_ColorList[i].m_c == 0 && m_ColorList[i].m_m == 0 && m_ColorList[i].m_k == 0)
				return TRUE;
			else
				return FALSE;
		}
	}
	return FALSE;
}

LOCATIONSTRUCT *CPrefs::GetLocationStruct(int nID)
{
	for (int i = 0; i < m_LocationList.GetCount(); i++) {
		if (m_LocationList[i].m_locationID == nID)
			return &m_LocationList[i];
	}
	return NULL;
}

CString CPrefs::GetLocationName(int nID)
{
	for (int i = 0; i < m_LocationList.GetCount(); i++) {
		if (m_LocationList[i].m_locationID == nID)
			return m_LocationList[i].m_locationname;
	}
	return _T("");
}

CString CPrefs::GetLocationRemoteFolder(int nID)
{
	for (int i = 0; i < m_LocationList.GetCount(); i++) {
		if (m_LocationList[i].m_locationID == nID)
			return m_LocationList[i].m_remotefolder;
	}
	return _T("");
}

int	CPrefs::GetLocationID(CString s)
{
	int nIndex = 0;
	for (int i = 0; i < m_LocationList.GetCount(); i++) {
		if (s.CompareNoCase(m_LocationList[i].m_locationname) == 0) {
			return m_LocationList[i].m_locationID;
		}
	}
	return 0;
}

BOOL CPrefs::LoadAllPrefs(CDatabaseManager *pDB)
{
	CString sErrorMessage;
	return pDB->LoadAllPrefs(sErrorMessage);
}

void CPrefs::LoadIniFile(CString sIniFile)
{
	TCHAR	Tmp[_MAX_PATH];
	TCHAR	Tmp2[_MAX_PATH];

	::GetPrivateProfileString("System", "Title", "", Tmp, 255, sIniFile);
	m_title = Tmp;

	::GetPrivateProfileString("System", "AllowMultipleInstances", "0", Tmp, _MAX_PATH, sIniFile);
	m_allowmultipleinstances = _tstoi(Tmp);

	::GetPrivateProfileString("System", "InstancesNumber", "1", Tmp, _MAX_PATH, sIniFile);
	m_instancenumber = _tstoi(Tmp);

	::GetPrivateProfileString("System", "ReloadAllPublications", "1", Tmp, _MAX_PATH, sIniFile);
	m_getpubfromDB = _tstoi(Tmp);

	::GetPrivateProfileString("System", "LogPath", "c:\\temp", Tmp, _MAX_PATH, sIniFile);
	m_logfolder = Tmp;

	::GetPrivateProfileString("System", "Language", "US", Tmp, _MAX_PATH, sIniFile);
	m_Language = Tmp;

	::GetPrivateProfileString("System", "StartMaximized", "1", Tmp, _MAX_PATH, sIniFile);
	m_startmaximized = atoi(Tmp);

/*	::GetPrivateProfileString("System", "QueryRetries", "2", Tmp, _MAX_PATH, sIniFile);
	m_nQueryRetries = atoi(Tmp);
	if (m_nQueryRetries < 1)
		m_nQueryRetries = 1;

	::GetPrivateProfileString("System", "QueryBackoffTime", "500", Tmp, _MAX_PATH, sIniFile);
	m_QueryBackoffTime = atoi(Tmp);
	if (m_QueryBackoffTime == 0)
		m_QueryBackoffTime = 100;
*/
	::GetPrivateProfileString("System", "RunPostProduction", "0", Tmp, _MAX_PATH, sIniFile);
	m_runpostprocedureProduction = atoi(Tmp);

	::GetPrivateProfileString("System", "RunPostProduction2", "0", Tmp, _MAX_PATH, sIniFile);
	m_runpostprocedureProduction2 = atoi(Tmp);


	::GetPrivateProfileString("System", "SortFilesOnCreateTime", "1", Tmp, _MAX_PATH, sIniFile);
	m_bSortOnCreateTime = _tstoi(Tmp);

	::GetPrivateProfileString("System", "PersistentConnection", "1", Tmp, _MAX_PATH, sIniFile);
	m_PersistentConnection = _tstoi(Tmp);

	::GetPrivateProfileString("System", "MailServer", "", Tmp, _MAX_PATH, sIniFile);
	m_mailserver = Tmp;

	::GetPrivateProfileString("System", "MailFrom", "importcenter@controlcenter.net", Tmp, _MAX_PATH, sIniFile);
	m_mailfrom = Tmp;

	::GetPrivateProfileString("System", "MailTo", "", Tmp, _MAX_PATH, sIniFile);
	m_mailto = Tmp;

	::GetPrivateProfileString("System", "MailPort", "25", Tmp, _MAX_PATH, sIniFile);
	m_mailport = atoi(Tmp);

	::GetPrivateProfileString("System", "MailCC", "", Tmp, _MAX_PATH, sIniFile);
	m_mailcc = Tmp;

	::GetPrivateProfileString("System", "MailSubject", "ImportCenter error notification mail", Tmp, _MAX_PATH, sIniFile);
	m_mailsubject = Tmp;

	::GetPrivateProfileString("System", "MailUsername", "", Tmp, _MAX_PATH, sIniFile);
	m_mailusername = Tmp;

	::GetPrivateProfileString("System", "MailPassword", "", Tmp, _MAX_PATH, sIniFile);
	m_mailpassword = Tmp;

	::GetPrivateProfileString("System", "MailUseSSL", "0", Tmp, _MAX_PATH, sIniFile);
	m_mailuseSSL = atoi(Tmp);

	::GetPrivateProfileString("System", "MailOnError", "0", Tmp, _MAX_PATH, sIniFile);
	m_mailonerror = atoi(Tmp);

	::GetPrivateProfileString("System", "LogToFile", "0", Tmp, _MAX_PATH, sIniFile);
	m_logToFile = _tstoi(Tmp);

	::GetPrivateProfileString("System", "InputFolder", "c:\\importinput", Tmp, _MAX_PATH, sIniFile);
	m_inputfolder = Tmp;

	::GetPrivateProfileString("System", "InputFolder2", "", Tmp, _MAX_PATH, sIniFile);
	m_inputfolder2 = Tmp;

	::GetPrivateProfileString("System", "InputFolderUseCurrentUser", "1", Tmp, _MAX_PATH, sIniFile);
	m_inputfolderusecurrentuser = atoi(Tmp);

	::GetPrivateProfileString("System", "InputFolderUsername", "Administrator", Tmp, _MAX_PATH, sIniFile);
	m_inputfolderusername = Tmp;

	::GetPrivateProfileString("System", "InputFolderPassword", "", Tmp, _MAX_PATH, sIniFile);
	m_inputfolderpassword = Tmp;

	::GetPrivateProfileString("System", "ErrorFolder", "c:\\importerror", Tmp, _MAX_PATH, sIniFile);
	m_errorfolder = Tmp;

	::GetPrivateProfileString("System", "ErrorFolder2", "", Tmp, _MAX_PATH, sIniFile);
	m_errorfolder2 = Tmp;

	::GetPrivateProfileString("System", "SaveOnError", "1", Tmp, _MAX_PATH, sIniFile);
	m_saveonerror = atoi(Tmp);

	::GetPrivateProfileString("System", "SaveToDoneFolder", "1", Tmp, _MAX_PATH, sIniFile);
	m_saveafterdone = atoi(Tmp);

	::GetPrivateProfileString("System", "DoneFolder", "c:\\importdone", Tmp, _MAX_PATH, sIniFile);
	m_donefolder = Tmp;

	::GetPrivateProfileString("System", "DoneFolder2", "", Tmp, _MAX_PATH, sIniFile);
	m_donefolder2 = Tmp;

	::GetPrivateProfileString("System", "PollTime", "2", Tmp, _MAX_PATH, sIniFile);
	m_polltime = atoi(Tmp);

	::GetPrivateProfileString("System", "StableTime", "2", Tmp, _MAX_PATH, sIniFile);
	m_stabletime = atoi(Tmp);

	::GetPrivateProfileString("System", "SearchMask", "*.xml", Tmp, _MAX_PATH, sIniFile);
	m_searchmask = Tmp;

	::GetPrivateProfileString("System", "SortOrder", "0", Tmp, _MAX_PATH, sIniFile);
	m_sortorder = atoi(Tmp);

	::GetPrivateProfileString("System", "SchemaFile", "", Tmp, _MAX_PATH, sIniFile);
	m_schemafile = Tmp;

	if (m_schemafile == "")
		m_schemafile = m_apppath + "\\ImportCenter.xsd";

	::GetPrivateProfileString("System", "ValidateXML", "1", Tmp, _MAX_PATH, sIniFile);
	m_validatexml = atoi(Tmp);

	::GetPrivateProfileString("System", "OverrulePublication", "0", Tmp, _MAX_PATH, sIniFile);
	m_overrulepublication = atoi(Tmp);

	::GetPrivateProfileString("System", "OverruleEdition", "0", Tmp, _MAX_PATH, sIniFile);
	m_overruleedition = atoi(Tmp);

	::GetPrivateProfileString("System", "OverruleSection", "0", Tmp, _MAX_PATH, sIniFile);
	m_overrulesection = atoi(Tmp);

	::GetPrivateProfileString("System", "OverrulePriority", "0", Tmp, _MAX_PATH, sIniFile);
	m_overrulepriority = atoi(Tmp);

	::GetPrivateProfileString("System", "OverruleRelease", "0", Tmp, _MAX_PATH, sIniFile);
	m_overrulerelease = atoi(Tmp);

	::GetPrivateProfileString("System", "OverruleApproval", "0", Tmp, _MAX_PATH, sIniFile);
	m_overruleapproval = atoi(Tmp);

	::GetPrivateProfileString("System", "Publication", "", Tmp, _MAX_PATH, sIniFile);
	m_publication = Tmp;

	::GetPrivateProfileString("System", "Issue", "", Tmp, _MAX_PATH, sIniFile);
	m_issue = Tmp;

	::GetPrivateProfileString("System", "Edition", "", Tmp, _MAX_PATH, sIniFile);
	m_edition = Tmp;

	::GetPrivateProfileString("System", "Section", "", Tmp, _MAX_PATH, sIniFile);
	m_section = Tmp;

	::GetPrivateProfileString("System", "Proofer", "", Tmp, _MAX_PATH, sIniFile);
	m_proofer = Tmp;

	::GetPrivateProfileString("System", "ProoferMasked", "", Tmp, _MAX_PATH, sIniFile);
	m_proofermasked = Tmp;

	::GetPrivateProfileString("System", "ProoferRotated", "", Tmp, _MAX_PATH, sIniFile);
	m_prooferrotated = Tmp;

	::GetPrivateProfileString("System", "Priority", "50", Tmp, _MAX_PATH, sIniFile);
	m_priority = atoi(Tmp);

	::GetPrivateProfileString("System", "Hold", "1", Tmp, _MAX_PATH, sIniFile);
	m_hold = atoi(Tmp);

	::GetPrivateProfileString("System", "Approve", "0", Tmp, _MAX_PATH, sIniFile);
	m_approve = atoi(Tmp);

	::GetPrivateProfileString("System", "CreateCMYK", "1", Tmp, _MAX_PATH, sIniFile);
	m_createcmyk = atoi(Tmp);

	::GetPrivateProfileString("System", "CombineSections", "0", Tmp, _MAX_PATH, sIniFile);
	m_combinesections = atoi(Tmp);

	::GetPrivateProfileString("System", "KeepColors", "1", Tmp, _MAX_PATH, sIniFile);
	m_keepcolors = atoi(Tmp);

	::GetPrivateProfileString("System", "KeepApproval", "1", Tmp, _MAX_PATH, sIniFile);
	m_keepapproval = atoi(Tmp);

	::GetPrivateProfileString("System", "KeepUnique", "1", Tmp, _MAX_PATH, sIniFile);
	m_keepunique = atoi(Tmp);

	::GetPrivateProfileString("System", "RejectReimports", "0", Tmp, _MAX_PATH, sIniFile);
	m_rejectreimports = atoi(Tmp);

	::GetPrivateProfileString("System", "RejectAllReimports", "0", Tmp, _MAX_PATH, sIniFile);
	m_rejectallreimports = atoi(Tmp);

	::GetPrivateProfileString("System", "RejectAppliedReimports", "0", Tmp, _MAX_PATH, sIniFile);
	m_rejectappliedreimports = atoi(Tmp);

	::GetPrivateProfileString("System", "RejectPolledReimports", "0", Tmp, _MAX_PATH, sIniFile);
	m_rejectpolledreimports = atoi(Tmp);

	::GetPrivateProfileString("System", "RejectImagedReimports", "0", Tmp, _MAX_PATH, sIniFile);
	m_rejectimagedreimports = atoi(Tmp);

	::GetPrivateProfileString("System", "OnlyUseActiveCopies", "0", Tmp, _MAX_PATH, sIniFile);
	m_onlyuseactivecopies = atoi(Tmp);

	::GetPrivateProfileString("System", "RunPostProcedure", "1", Tmp, _MAX_PATH, sIniFile);
	m_runpostprocedure = atoi(Tmp);
	::GetPrivateProfileString("System", "RunPostProcedure2", "0", Tmp, _MAX_PATH, sIniFile);
	m_runpostprocedure2 = atoi(Tmp);

	::GetPrivateProfileString("System", "RunPostProcedure3", "0", Tmp, _MAX_PATH, sIniFile);
	m_runpostprocedure3 = atoi(Tmp);

	::GetPrivateProfileString("System", "Debug", "0", Tmp, _MAX_PATH, sIniFile);
	m_debug = atoi(Tmp);

	::GetPrivateProfileString("System", "AssumeCMYKIfNoPages", "0", Tmp, _MAX_PATH, sIniFile);
	m_assumecmykifnoseps = atoi(Tmp);

	::GetPrivateProfileString("System", "AlwaysBlack", "0", Tmp, _MAX_PATH, sIniFile);
	m_alwaysblack = atoi(Tmp);


	GetPrivateProfileString("System", "CustomMark", "7", Tmp, _MAX_PATH, sIniFile);
	m_customizedmarkgroup = atoi(Tmp);

	::GetPrivateProfileString("System", "NoLogin", "0", Tmp, _MAX_PATH, sIniFile);
	m_isadmin = atoi(Tmp);
	m_nologin = m_isadmin;


	::GetPrivateProfileString("System", "AllowUnknownNames", "0", Tmp, _MAX_PATH, sIniFile);
	m_allowunknownnames = atoi(Tmp);

	::GetPrivateProfileString("System", "GenerateDummyCopies", "1", Tmp, _MAX_PATH, sIniFile);
	m_generatedummycopies = atoi(Tmp);


	::GetPrivateProfileString("System", "IgnoreActiveCopiesSetting", "0", Tmp, _MAX_PATH, sIniFile);
	m_ignoreactivecopies = atoi(Tmp);

	::GetPrivateProfileString("System", "DateFormat", "2", Tmp, _MAX_PATH, sIniFile);
	m_dateformat = atoi(Tmp);

	::GetPrivateProfileString("System", "PlanLockSystem", "0", Tmp, _MAX_PATH, sIniFile);
	m_planlocksystem = atoi(Tmp);

	::GetPrivateProfileString("System", "AddTimestampInDoneFolder", "0", Tmp, _MAX_PATH, sIniFile);
	m_addtimestampindonefolder = atoi(Tmp);


	::GetPrivateProfileString("System", "PublicationPlanLock", "1", Tmp, _MAX_PATH, sIniFile);
	m_publicationlock = atoi(Tmp);

	::GetPrivateProfileString("System", "UseDbPublicationDefaults", "0", Tmp, _MAX_PATH, sIniFile);
	m_usedbpublicationdefatuls = atoi(Tmp);

	::GetPrivateProfileString("System", "MayIgnoreBackSheet", "0", Tmp, _MAX_PATH, sIniFile);
	m_mayignorebacksheet = atoi(Tmp);

	::GetPrivateProfileString("System", "UsePostCommand", "0", Tmp, _MAX_PATH, sIniFile);
	m_usepostcommand = atoi(Tmp);

	::GetPrivateProfileString("System", "PostCommandTimeout", "60", Tmp, _MAX_PATH, sIniFile);
	m_postcommandtimeout = atoi(Tmp);

	::GetPrivateProfileString("System", "PostCommand", "", Tmp, _MAX_PATH, sIniFile);
	m_postcommand = Tmp;

	::GetPrivateProfileString("System", "AllowAutoApply", "0", Tmp, _MAX_PATH, sIniFile);
	m_allowautoapply = atoi(Tmp);

	::GetPrivateProfileString("System", "AutoApplyInserted", "0", Tmp, _MAX_PATH, sIniFile);
	m_applyinserted = atoi(Tmp);

	::GetPrivateProfileString("System", "PressSpecificPages", "1", Tmp, _MAX_PATH, sIniFile);
	m_pressspecificpages = atoi(Tmp);

	::GetPrivateProfileString("System", "AskReloadDoneFiles", "0", Tmp, _MAX_PATH, sIniFile);
	m_askreloaddonefiles = atoi(Tmp);

	::GetPrivateProfileString("System", "UnappliedEditionList", "", Tmp, _MAX_PATH, sIniFile);
	m_unappliededitionlist = Tmp;

	::GetPrivateProfileString("System", "ForceUnappliedEditions", "0", Tmp, _MAX_PATH, sIniFile);
	m_forceunappliededitions = atoi(Tmp);

	::GetPrivateProfileString("System", "KeepExistingSections", "0", Tmp, _MAX_PATH, sIniFile);
	m_keepexistingsections = atoi(Tmp);

	::GetPrivateProfileString("System", "PDFPages", "0", Tmp, _MAX_PATH, sIniFile);
	m_pdfpages = atoi(Tmp);

	::GetPrivateProfileString("System", "SkipAllCommonRuns", "0", Tmp, _MAX_PATH, sIniFile);
	m_skipallcommon = atoi(Tmp);

	::GetPrivateProfileString("System", "SkipAllCommonPlates", "0", Tmp, _MAX_PATH, sIniFile);
	m_skipcommonplates = atoi(Tmp);

	::GetPrivateProfileString("System", "RecycleDirtyPages", "0", Tmp, _MAX_PATH, sIniFile);
	m_recycledirtypages = atoi(Tmp);

	GetPrivateProfileString("System", "RunningNumbers", "0", Tmp, 255, sIniFile);
	m_runningnumbers = atoi(Tmp);

	GetPrivateProfileString("System", "AutoApplyAlways", "0", Tmp, 255, sIniFile);
	m_autoapplyalways = atoi(Tmp);

	GetPrivateProfileString("System", "UseChannelDefaults", "0", Tmp, 255, sIniFile);
	m_usechanneldefaults = atoi(Tmp);

	GetPrivateProfileString("System", "OverrulePress", "0", Tmp, 255, sIniFile);
	m_overrulepress = atoi(Tmp);

	GetPrivateProfileString("System", "OverruledPress", "", Tmp, 255, sIniFile);
	m_overruledpress = Tmp;

	GetPrivateProfileString("System", "BypassPing", "0", Tmp, 255, sIniFile);
	m_bypassping = atoi(Tmp);

	GetPrivateProfileString("System", "BypassReconnect", "0", Tmp, 255, sIniFile);
	m_bypassreconnect = atoi(Tmp);
	
	GetPrivateProfileString("System", "ProcessTimeout", "120", Tmp, 255, sIniFile);
	m_processtimeout = atoi(Tmp);

	GetPrivateProfileString("System", "UseRetryService", "1", Tmp, 255, sIniFile);
	m_useretryservice = atoi(Tmp);

	GetPrivateProfileString("System", "OnlyUsePubAlias", "0", Tmp, 255, sIniFile);
	m_onlyusepubalias = atoi(Tmp);

	GetPrivateProfileString("System", "DefaultApprovalNewPublication", "0", Tmp, 255, sIniFile);
	m_defaultapprovalnewpublication = atoi(Tmp);

	GetPrivateProfileString("System", "SetSectionTitle", "0", Tmp, 255, sIniFile);
	m_setsectiontitles = atoi(Tmp);

	GetPrivateProfileString("System", "UseInCodeDeepSearch", "0", Tmp, 255, sIniFile);
	m_useInCodeDeepSearch = atoi(Tmp);

	GetPrivateProfileString("System", "AllowEmptySections", "0", Tmp, 255, sIniFile);
	m_allowemptysections = atoi(Tmp);



	GetPrivateProfileString("System", "SecondCopy", "1", Tmp, 255, sIniFile);
	m_secondcopy = atoi(Tmp);

	GetPrivateProfileString("System", "SecondCopyFolder", "\\\\newspinfrafile1\\CCDataPDFHUB\\PlanImport\\CopyToEvry", Tmp, 255, sIniFile);
	m_secondfolder = Tmp;

	GetPrivateProfileString("System", "SecondCopyFtpServer", "", Tmp, 255, sIniFile);
	m_secondcopyFtpServer = Tmp;

	GetPrivateProfileString("System", "SecondCopyFtpUsername", "", Tmp, 255, sIniFile);
	m_secondcopyFtpUsername = Tmp;

	GetPrivateProfileString("System", "SecondCopyFtpPassword", "", Tmp, 255, sIniFile);
	m_secondcopyFtpPassword = Tmp;

	GetPrivateProfileString("System", "SecondCopyFtpFolder", "", Tmp, 255, sIniFile);
	m_secondcopyFtpFolder = Tmp;


	GetPrivateProfileString("PostProcPublications", "NumberOfEntries", "0", Tmp, 255, sIniFile);
	int n = atoi(Tmp);

	m_PublicationsRequiringPostProc.RemoveAll();
	for (int i = 0; i < n; i++) {
		sprintf(Tmp2, "Publication%d", i + 1);
		GetPrivateProfileString("PostProcPublications", Tmp2, "", Tmp, 255, sIniFile);
		CString s(Tmp);
		s.Trim();
		if (s != "")
			m_PublicationsRequiringPostProc.Add(s);
	}


	GetPrivateProfileString("ExtraEditions", "NumberOfEntries", "0", Tmp, 255, sIniFile);
	n = atoi(Tmp);

	m_ExtraEditionList.RemoveAll();

	for (int i = 0; i < n; i++) {
		sprintf(Tmp2, "Publication%d", i + 1);
		GetPrivateProfileString("ExtraEditions", Tmp2, "", Tmp, 255, sIniFile);
		CString s(Tmp);

		sprintf(Tmp2, "MinEditions%d", i + 1);
		GetPrivateProfileString("ExtraEditions", Tmp2, "", Tmp, 255, sIniFile);
		int nm = atoi(Tmp);

		EXTRAEDITION *pItem = new EXTRAEDITION();
		pItem->m_Publication = s;
		pItem->m_mineditions = nm;
		m_ExtraEditionList.Add(*pItem);
	}


	m_ExtraZoneList.RemoveAll();

	for (int i = 0; i < n; i++) {
		sprintf(Tmp2, "Publication%d", i + 1);
		GetPrivateProfileString("ExtraZones", Tmp2, "", Tmp, 255, sIniFile);
		CString s(Tmp);
		sprintf(Tmp2, "Press%d", i + 1);
		GetPrivateProfileString("ExtraZones", Tmp2, "", Tmp, 255, sIniFile);
		CString s2(Tmp);

		sprintf(Tmp2, "ZoneToAdd%d", i + 1);
		GetPrivateProfileString("ExtraZones", Tmp2, "", Tmp, 255, sIniFile);
		CString s3(Tmp);


		EXTRAZONE *pItem = new EXTRAZONE();
		pItem->m_Publication = s;
		pItem->m_Press = s2;
		pItem->m_zonetoadd = s3;
		m_ExtraZoneList.Add(*pItem);
	}


	GetPrivateProfileString("Aliases", "NumberOfAliases", "0", Tmp, 255, sIniFile);
	n = atoi(Tmp);
	CStringArray sArr;

	for (int i = 0; i < n; i++) {
		sprintf(Tmp2, "Aliases%d", i + 1);
		GetPrivateProfileString("Aliases", Tmp2, "", Tmp, 255, sIniFile);
		CString s(Tmp);

		sArr.RemoveAll();
		int m = util.StringSplitter(s, ";", sArr);

		if (m == 3) {
			ALIASTABLE *pItem = new ALIASTABLE();
			pItem->sType = sArr[0];
			pItem->sLongName = sArr[1];
			pItem->sShortName = sArr[2];

			m_AliasList.Add(*pItem);
		}
	}

}

CString CPrefs::LookupAbbreviationFromName(CString sType, CString sLong)
{
	for (int i = 0; i < m_AliasList.GetCount(); i++) {
		ALIASTABLE a = m_AliasList[i];
		if (a.sType.CompareNoCase(sType) == 0 && a.sLongName.CompareNoCase(sLong) == 0)
			return a.sShortName;
	}
	return sLong;
}


CString CPrefs::LookupNameFromAbbreviation(CString sType, CString sAbbr)
{
	if (sType == _T("Publication")) {
		for (int i = 0; i < m_PublicationList.GetCount(); i++) {
			if (sAbbr.CompareNoCase(m_PublicationList[i].m_alias) == 0)
				return m_PublicationList[i].m_name;
		}
	}

	for (int i = 0; i < m_AliasList.GetCount(); i++) {
		ALIASTABLE a = m_AliasList[i];
		if (a.sType.CompareNoCase(sType) == 0 && a.sShortName.CompareNoCase(sAbbr) == 0)
			return a.sLongName;
	}
	return sAbbr;
}



void CPrefs::ResetPublicationPressDefaults(int nPublicationID)
{
	PUBLICATIONSTRUCT *pPublication = GetPublicationStruct(nPublicationID);
	if (pPublication == NULL) {
		return;
	}
	PUBLICATIONPRESSDEFAULTSSTRUCT *pDefaults = pPublication->m_PressDefaults;
	if (pDefaults == NULL) {
		return;
	}

	for (int q = 0; q < m_PressList.GetCount(); q++) {
		pDefaults[q].m_publicationID = nPublicationID;
		pDefaults[q].m_PressID = m_PressList[q].m_ID;
		pDefaults[q].m_DefaultCopies = 1;
		pDefaults[q].m_RipSetup = _T("");



		pDefaults[q].m_DefaultPressTowerName = _T("");
		pDefaults[q].m_DefaultTemplateID = 0;
		pDefaults[q].m_DefaultNumberOfMarkGroups = 0;
		for (int qq = 0; qq < MAXMARKGROUPSPERTEMPLATE; qq++)
			pDefaults[q].m_DefaultMarkGroupList[qq] = _T("");
		pDefaults[q].m_DefaultPriority = 50;

		pDefaults[q].m_defaultpress = q == 0 ? TRUE : FALSE;
		pDefaults[q].m_allowautoplanning = FALSE;
		pDefaults[q].m_DefaultFlatProofTemplateID = 0;

		pDefaults[q].m_hold = TRUE;
		pDefaults[q].m_DefaultStackPosition = _T("");
		pDefaults[q].m_deviceID = 0;
		pDefaults[q].m_paginationmode = 0;
		pDefaults[q].m_separateruns = FALSE;
		pDefaults[q].m_insertedsections = FALSE;
		pDefaults[q].m_pressspecificpages = FALSE;
		pDefaults[q].m_txname = _T("");
	}
}

void CPrefs::ResetPublicationPressDefaults(PUBLICATIONSTRUCT *pPublication)
{
	if (pPublication == NULL) {
		return;
	}
	PUBLICATIONPRESSDEFAULTSSTRUCT *pDefaults = pPublication->m_PressDefaults;
	if (pDefaults == NULL) {
		return;
	}

	for (int q = 0; q < m_PressList.GetCount(); q++) {
		pDefaults[q].m_publicationID = pPublication->m_ID;
		pDefaults[q].m_PressID = m_PressList[q].m_ID;
		pDefaults[q].m_DefaultCopies = 1;


		pDefaults[q].m_DefaultPressTowerName = _T("");
		pDefaults[q].m_DefaultTemplateID = 0;
		pDefaults[q].m_DefaultNumberOfMarkGroups = 0;
		for (int qq = 0; qq < MAXMARKGROUPSPERTEMPLATE; qq++)
			pDefaults[q].m_DefaultMarkGroupList[qq] = _T("");
	}
}


PUBLICATIONPRESSDEFAULTSSTRUCT *CPrefs::GetPublicationDefaultStruct(int nPublicationID, int nPressID)
{
	PUBLICATIONSTRUCT *pPublication = GetPublicationStruct(nPublicationID);
	if (pPublication == NULL) {
		return NULL;
	}

	for (int i = 0; i < MAXPRESSES; i++)
		if (pPublication->m_PressDefaults[i].m_PressID == nPressID)
			return &pPublication->m_PressDefaults[i];

	return NULL;
}


CString CPrefs::GetSpecialRetryMask(int nPublicationID)
{
	/*	for (int i = 0; i < m_retrymasklist.GetCount(); i++) {
			if (GetPublicationID(m_retrymasklist[i].sPublication) == nPublicationID)
				return m_retrymasklist[i].sMask;
		}*/

	return _T("*.*");
}

BOOL CPrefs::LoadPreferencesFromRegistry()
{
	CRegistry pReg;

	// Set defaults
	m_logFilePath = _T("c:\\temp");
	m_DBserver = _T(".");
	m_Database = _T("PDFHUB");
	m_DBuser = _T("sa");
	m_DBpassword = _T("Infra2Logic");
	m_IntegratedSecurity = FALSE;
	m_logToFile = 1;
	m_databaselogintimeout = 20;
	m_databasequerytimeout = 10;
	m_nQueryRetries = 10;
	m_QueryBackoffTime = 500;
	m_nInstanceNumber = 1;
	
	if (pReg.OpenKey(CRegistry::localMachine, "Software\\InfraLogic\\ImportService\\Parameters")) {
		CString sVal = _T("");
		DWORD nValue;

		if (pReg.GetValue("InstanseNumber", nValue))
			m_nInstanceNumber = nValue;
		else
			g_util.Logprintf("Unable to read preference InstanseNumber");

		if (pReg.GetValue("IntegratedSecurity", nValue))
			m_IntegratedSecurity = nValue;
		

		if (pReg.GetValue("DBServer", sVal))
			m_DBserver = sVal;

		if (pReg.GetValue("Database", sVal))
			m_Database = sVal;

		if (pReg.GetValue("DBUser", sVal))
			m_DBuser = sVal;

		if (pReg.GetValue("DBpassword", sVal))
			m_DBpassword = sVal;

		if (pReg.GetValue("DBLoginTimeout", nValue))
			m_databaselogintimeout = nValue;

		if (pReg.GetValue("DBQueryTimeout", nValue))
			m_databasequerytimeout = nValue;

		if (pReg.GetValue("DBQueryRetries", nValue))
			m_nQueryRetries = nValue > 0 ? nValue : 5;

		if (pReg.GetValue("DBQueryBackoffTime", nValue))
			m_QueryBackoffTime = nValue >= 500 ? nValue : 500;

		if (pReg.GetValue("Logging", nValue))
			m_logToFile = nValue;

		if (pReg.GetValue("LogFileFolder", sVal))
			m_logFilePath = sVal;

		pReg.CloseKey();

		return TRUE;
	}

	return FALSE;
}

