
#include "Stdafx.h"
#include "ImportService.h"
#include "Defs.h"
#include "PlanDataDefs.h"
#include "Prefs.h"	
#include "Utils.h"	

extern CPrefs g_prefs;
extern CUtils g_util;

CDatabaseManager::CDatabaseManager(void)
{
	m_DBopen = FALSE;
	m_pDB = NULL;

	m_DBserver = _T(".");
	m_Database = _T("PDFHUB");
	m_DBuser = "saxxxxxxxxx";
	m_DBpassword = "xxxxxx";
	m_IntegratedSecurity = FALSE;
}

CDatabaseManager::~CDatabaseManager(void)
{
	ExitDB();
	if (m_pDB != NULL)
		delete m_pDB;
}

BOOL CDatabaseManager::InitDB(CString sDBserver, CString sDatabase, CString sDBuser, CString sDBpassword, BOOL bIntegratedSecurity, CString &sErrorMessage)
{
	m_DBserver = sDBserver;
	m_Database = sDatabase;
	m_DBuser = sDBuser;
	m_DBpassword = sDBpassword;
	m_IntegratedSecurity = bIntegratedSecurity;

	return InitDB(sErrorMessage);
}

int CDatabaseManager::InitDB(CString &sErrorMessage)
{
	sErrorMessage.Format("");
	if (m_pDB) {
		if (m_pDB->IsOpen() == FALSE) {
			try {
				m_DBopen = FALSE;
				m_pDB->Close();
			}
			catch (CDBException* e) {
				// So what..! Go on!;
			}
		}
	}

	if (!m_PersistentConnection)
		ExitDB();

	if (m_DBopen)
		return TRUE;

	if (m_DBserver == _T("") || m_Database == _T("") || m_DBuser == _T("")) {
		sErrorMessage.Format("Empty server, database or username not allowed");
		return FALSE;
	}

	if (m_pDB == NULL)
		m_pDB = new CDatabase;

	m_pDB->SetLoginTimeout(g_prefs.m_databaselogintimeout);
	m_pDB->SetQueryTimeout(g_prefs.m_databasequerytimeout);

	CString sConnectStr = _T("Driver={SQL Server}; Server=") + m_DBserver + _T("; ") +
		_T("Database=") + m_Database + _T("; ");

	if (m_IntegratedSecurity)
		sConnectStr += _T(" Integrated Security=True;");
	else
		sConnectStr += _T("USER=") + m_DBuser + _T("; PASSWORD=") + m_DBpassword + _T(";");

	try {
		if (!m_pDB->OpenEx((LPCTSTR)sConnectStr, CDatabase::noOdbcDialog)) {
			sErrorMessage.Format(_T("Error connecting to database with connection string '%s'"), (LPCSTR)sConnectStr);
			return FALSE;
		}
	}
	catch (CDBException* e) {
		sErrorMessage.Format(_T("Error connecting to database - %s (%s)"), (LPCSTR)e->m_strError, (LPCSTR)sConnectStr);
		e->Delete();
		return FALSE;
	}

	m_DBopen = TRUE;
	return TRUE;
}

void CDatabaseManager::ExitDB()
{
	if (!m_DBopen)
		return;

	if (m_pDB)
		m_pDB->Close();

	m_DBopen = FALSE;

	return;
}

BOOL CDatabaseManager::IsOpen()
{
	return m_DBopen;
}


BOOL CDatabaseManager::RegisterService(CString &sErrorMessage)
{
	CString sSQL, s;
	sErrorMessage.Format("");

	if (InitDB(sErrorMessage) == FALSE)
		return -1;

	CRecordset Rs(m_pDB);

	TCHAR buf[MAX_PATH];
	DWORD buflen = MAX_PATH;
	::GetComputerName(buf, &buflen);

	sSQL.Format("{CALL spRegisterService ('ImportService', %d, %d, '%s',-1,'','','')}", g_prefs.m_instancenumber, SERVICETYPE_PLANIMPORT, buf);

	try {
		m_pDB->ExecuteSQL(sSQL);
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Insert failed - %s", (LPCSTR)e->m_strError);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!; 
		}
		return FALSE;
	}
	return TRUE;
}


BOOL CDatabaseManager::UpdateService(int nCurrentState, CString sCurrentJob, CString sLastError, CString &sErrorMessage)
{
	return UpdateService(nCurrentState, sCurrentJob, sLastError, _T(""), sErrorMessage);
}

BOOL CDatabaseManager::UpdateService(int nCurrentState, CString sCurrentJob, CString sLastError, CString sAddedLogData, CString &sErrorMessage)
{
	CString sSQL;

	TCHAR buf[MAX_PATH];
	DWORD buflen = MAX_PATH;
	::GetComputerName(buf, &buflen);
	sErrorMessage = _T("");

	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;
	sSQL.Format("{CALL spRegisterService ('ImportService', %d, %d, '%s',%d,'%s','%s','%s')}", g_prefs.m_instancenumber, SERVICETYPE_PLANIMPORT, buf, nCurrentState, sCurrentJob, sLastError, sAddedLogData);
	//g_util.Logprintf("DEBUG: %s", sSQL);
	try {
		m_pDB->ExecuteSQL(sSQL);
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}

	return TRUE;
}

BOOL CDatabaseManager::InsertLogEntry(int nEvent, CString sSource, CString sFileName, CString sMessage, int nMasterCopySeparationSet, int nVersion, int nMiscInt, CString sMiscString, CString &sErrorMessage)
{
	CString sSQL;

	sErrorMessage = _T("");
	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	if (sMessage == "")
		sMessage = "Ok";

	sSQL.Format("{CALL [spAddServiceLogEntry] ('%s',%d,%d, '%s','%s', '%s',%d,%d,%d,'%s')}",
		_T("ImportService"), g_prefs.m_instancenumber, nEvent, sSource, sFileName, sMessage, nMasterCopySeparationSet, nVersion, nMiscInt, sMiscString);

	try {
		m_pDB->ExecuteSQL(sSQL);
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Insert failed - %s", (LPCSTR)e->m_strError);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}

	return TRUE;
}

BOOL CDatabaseManager::LoadImportConfigurations(int nInstanceNumber, CString &sErrorMessage)
{
	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	CRecordset Rs(m_pDB);
	CString sSQL, s;

	g_prefs.m_ImportConfigurations.RemoveAll();


	sSQL.Format("SELECT [ImportID],[Name],[Type],[InputFolder],[DoneFolder],[ErrorFolder],LogFolder,ConfigFile,ConfigFile2,CopyFolder,SendErrorEmail,EmailReceiver,ConfigChangeTime FROM ImportConfigurations WHERE (OwnerInstance=%d OR OwnerInstance=0) ORDER BY ImportID", nInstanceNumber);

	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("Query failed - %s", (LPCSTR)sSQL);
			return FALSE;
		}

		while (!Rs.IsEOF()) {
			int f = 0;
			IMPORTCONFIGURATION *pItem = new IMPORTCONFIGURATION();

			Rs.GetFieldValue((short)f++, s);
			pItem->nImportID = atoi(s);
			Rs.GetFieldValue((short)f++, s);
			pItem->sImportName = s;
			Rs.GetFieldValue((short)f++, s);
			pItem->nType = atoi(s);
			Rs.GetFieldValue((short)f++, s);
			pItem->sInputFolder = s;

			Rs.GetFieldValue((short)f++, s);
			pItem->sDoneFolder = s;
			Rs.GetFieldValue((short)f++, s);
			pItem->sErrorFolder = s;
			Rs.GetFieldValue((short)f++, s);
			pItem->sLogFolder = s;
			Rs.GetFieldValue((short)f++, s);
			pItem->sConfigFile = s;
			Rs.GetFieldValue((short)f++, s);
			pItem->sConfigFile2 = s;
			Rs.GetFieldValue((short)f++, s);
			pItem->sCopyFolder = s;
			Rs.GetFieldValue((short)f++, s);
			pItem->bSendErrorEmail = atoi(s);
			Rs.GetFieldValue((short)f++, s);
			pItem->sEmailReceivers = s;

			//ConfigChangeTim

			g_prefs.m_ImportConfigurations.Add(*pItem);
			Rs.MoveNext();
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		e->Delete();
		Rs.Close();

		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}
	return TRUE;


}

BOOL CDatabaseManager::LoadConfigIniFile(int nInstanceNumber, CString sFileName, CString sFileName2, CString &sErrorMessage)
{
	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return -1;

	CRecordset Rs(m_pDB);
	CString sSQL, s;
	sSQL.Format("SELECT ConfigData,ConfigData2 FROM ServiceConfigurations WHERE ServiceName='ImportService' AND InstanceNumber=%d", nInstanceNumber);

	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("Query failed - %s", (LPCSTR)sSQL);
			return FALSE;
		}

		if (!Rs.IsEOF()) {
			Rs.GetFieldValue((short)0, s);
			s.Trim();
			if (sFileName != "" && s != "")
				g_util.Save(sFileName, s);
			Rs.GetFieldValue((short)1, s);
			s.Trim();
			if (sFileName2 != "" && s != "")
				g_util.Save(sFileName, s);
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		e->Delete();
		Rs.Close();

		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}
	return TRUE;

}

BOOL CDatabaseManager::LoadSTMPSetup(CString &sErrorMessage)
{
	sErrorMessage = _T("");

	CString sSQL = _T("SELECT TOP 1 SMTPServer,SMTPPort, SMTPUserName,SMTPPassword,UseSSL,SMTPConnectionTimeout,SMTPFrom,SMTPCC,SMTPTo,SMTPSubject FROM SMTPPreferences");
	CString s;
	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	CRecordset Rs(m_pDB);

	try {
		if (!Rs.Open(CRecordset::snapshot, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("Query failed - %s", sSQL);
			return FALSE;
		}

		if (!Rs.IsEOF()) {
			CString s;
			int fld = 0;
			Rs.GetFieldValue((short)fld++, g_prefs.m_mailserver);
			Rs.GetFieldValue((short)fld++, s);
			g_prefs.m_mailport = atoi(s);

			Rs.GetFieldValue((short)fld++, g_prefs.m_mailusername);
			Rs.GetFieldValue((short)fld++, g_prefs.m_mailpassword);
			Rs.GetFieldValue((short)fld++, s);
			g_prefs.m_mailuseSSL = atoi(s);
			Rs.GetFieldValue((short)fld++, s);
			g_prefs.m_mailtimeout = atoi(s);

			Rs.GetFieldValue((short)fld++, g_prefs.m_mailfrom);
			Rs.GetFieldValue((short)fld++, g_prefs.m_mailcc);
			Rs.GetFieldValue((short)fld++, g_prefs.m_mailto);
			Rs.GetFieldValue((short)fld++, g_prefs.m_mailsubject);

		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("Query failed - %s (%s)", e->m_strError, sSQL);
		LogprintfDB("Error: %s   (%s)", e->m_strError, sSQL);

		e->Delete();
		Rs.Close();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
			return FALSE;
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}

	return TRUE;

}


BOOL CDatabaseManager::LoadAllPrefs(CString &sErrorMessage)
{
	if (LoadGeneralPreferences(sErrorMessage) == FALSE)
		return FALSE;

	LoadSTMPSetup(sErrorMessage);

	if (LoadPressList(sErrorMessage) == FALSE)
		return FALSE;

	if (LoadTemplateList(sErrorMessage) == FALSE)
		return FALSE;

	if (LoadMarkGroupList(sErrorMessage) == FALSE)
		return FALSE;

	g_prefs.m_PublicationList.RemoveAll();
	if (RetrievePublicationList(sErrorMessage) == FALSE)
		return FALSE;

	g_prefs.m_SectionList.RemoveAll();
	if (LoadDBNameList("SectionNames", &g_prefs.m_SectionList, sErrorMessage) == FALSE)
		return FALSE;

	g_prefs.m_IssueList.RemoveAll();
	if (LoadDBNameList("IssueNames", &g_prefs.m_IssueList, sErrorMessage) == FALSE)
		return FALSE;

	if (LoadEditionNameList(sErrorMessage) == FALSE)
		return FALSE;

	if (LoadLocationList(sErrorMessage) < 0)
		return FALSE;

	if (LoadColorList(sErrorMessage) < 0)
		return FALSE;

	if (LoadProofProcessList(sErrorMessage) == FALSE)
		return FALSE;

	if (LoadDeviceList(sErrorMessage) == FALSE)
		return FALSE;

	if (RetrievePageFormatList(sErrorMessage) == FALSE)
		return FALSE;

	return TRUE;
}

BOOL CDatabaseManager::LoadDBNameList(CString sIDtable, ITEMLIST *v, CString &sErrorMessage)
{
	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return -1;

	CRecordset Rs(m_pDB);

	CString sSQL;
	sSQL.Format("SELECT * FROM %s", sIDtable);

	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("Query failed - %s", (LPCSTR)sSQL);
			return FALSE;
		}

		while (!Rs.IsEOF()) {
			ITEMSTRUCT *pItem = new ITEMSTRUCT;
			CString s;
			Rs.GetFieldValue((short)0, s);
			pItem->m_ID = atoi(s);
			Rs.GetFieldValue((short)1, pItem->m_name);
			v->Add(*pItem);
			Rs.MoveNext();
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		e->Delete();
		Rs.Close();

		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}
	return TRUE;
}

BOOL CDatabaseManager::LoadEditionNameList(CString &sErrorMessage)
{
	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	CRecordset Rs(m_pDB);

	g_prefs.m_EditionList.RemoveAll();
	CString sSQL;
	sSQL = _T("SELECT EditionID,Name,IsCommonEdition,SubOfEditionID FROM EditionNames");

	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("Query failed - %s", (LPCSTR)sSQL);
			return FALSE;
		}

		while (!Rs.IsEOF()) {
			EDITIONSTRUCT *pItem = new EDITIONSTRUCT;
			CString s;
			Rs.GetFieldValue((short)0, s);
			pItem->m_ID = atoi(s);
			Rs.GetFieldValue((short)1, pItem->m_name);

			Rs.GetFieldValue((short)2, s);
			pItem->m_iscommon = atoi(s);
			Rs.GetFieldValue((short)3, s);
			pItem->m_subofeditionID = atoi(s);
			g_prefs.m_EditionList.Add(*pItem);
			Rs.MoveNext();
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		e->Delete();
		Rs.Close();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}
	return TRUE;
}


int CDatabaseManager::LoadLocationList(CString &sErrorMessage)
{
	int foundJob = 0;

	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return -1;

	g_prefs.m_LocationList.RemoveAll();
	CRecordset Rs(m_pDB);

	CString sSQL = _T("SELECT LocationID,Name,RemoteFolder,BackupRemoteFolder,UseFTP, FTPServer, FTPusername, FTPpassword, FTPfolder, FTPport FROM LocationNames WHERE LocationID>0 AND RemoteFolder<>'' AND RemoteFolder<>'\'");
	try {
		if (!Rs.Open(CRecordset::snapshot, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("Query failed - %s", (LPCSTR)sSQL);
			return -1;
		}

		while (!Rs.IsEOF()) {
			CString s;
			int fld = 0;
			LOCATIONSTRUCT *pItem = new LOCATIONSTRUCT();

			Rs.GetFieldValue((short)fld++, s);
			pItem->m_locationID = atoi(s);

			Rs.GetFieldValue((short)fld++, pItem->m_locationname);
			Rs.GetFieldValue((short)fld++, pItem->m_remotefolder);
			if (pItem->m_remotefolder.Right(1) != "\\")
				pItem->m_remotefolder += _T("\\");
			Rs.GetFieldValue((short)fld++, pItem->m_backupremotefolder);
			if (pItem->m_backupremotefolder.Right(1) != "\\")
				pItem->m_backupremotefolder += _T("\\");

			Rs.GetFieldValue((short)fld++, s);
			pItem->m_useftp = atoi(s) > 0;

			Rs.GetFieldValue((short)fld++, pItem->m_ftpserver);
			Rs.GetFieldValue((short)fld++, pItem->m_ftpusername);
			Rs.GetFieldValue((short)fld++, pItem->m_ftppassword);
			Rs.GetFieldValue((short)fld++, pItem->m_ftpfolder);
			Rs.GetFieldValue((short)fld++, s);
			pItem->m_ftpport = atoi(s);

			g_prefs.m_LocationList.Add(*pItem);
			foundJob++;

			Rs.MoveNext();
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("Query failed - %s (%s)", (LPCSTR)e->m_strError, (LPCSTR)sSQL);
		e->Delete();
		Rs.Close();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
			return -1;
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return -1;
	}
	return foundJob;
}

BOOL CDatabaseManager::LoadTemplateList(CString &sErrorMessage)
{

	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	g_prefs.m_TemplateList.RemoveAll();
	CRecordset Rs(m_pDB);

	CString sSQL = _T("SELECT T.TemplateID, T.TemplateName, T.PressID,T.PagesAcross,T.PagesDown,T.IncomingPageRotationEven,T.IncomingPageRotationOdd FROM TemplateConfigurations AS T ORDER BY T.PressID, T.TemplateName");

	try {
		if (!Rs.Open(CRecordset::snapshot, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("Query failed - %s", (LPCSTR)sSQL);
			return FALSE;
		}

		while (!Rs.IsEOF()) {
			CString s;
			int fld = 0;
			TEMPLATESTRUCT *pItem = new TEMPLATESTRUCT();
			Rs.GetFieldValue((short)fld++, s);
			pItem->m_ID = atoi(s);
			Rs.GetFieldValue((short)fld++, pItem->m_name);
			Rs.GetFieldValue((short)fld++, s);
			pItem->m_pressID = atoi(s);
			Rs.GetFieldValue((short)fld++, s);
			pItem->m_pagesacross = atoi(s);
			Rs.GetFieldValue((short)fld++, s);
			pItem->m_pagesdown = atoi(s);
			Rs.GetFieldValue((short)fld++, s);
			pItem->m_incomingrotationeven = atoi(s);
			Rs.GetFieldValue((short)fld++, s);
			pItem->m_incomingrotationodd = atoi(s);

			g_prefs.m_TemplateList.Add(*pItem);
			Rs.MoveNext();
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("Query failed - %s (%s)", (LPCSTR)e->m_strError, (LPCSTR)sSQL);
		LogprintfDB("Error : %s   (%s)", e->m_strError, sSQL);
		e->Delete();
		Rs.Close();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
			return FALSE;
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}

	for (int i = 0; i < g_prefs.m_TemplateList.GetCount(); i++) {
		sSQL.Format("SELECT MarkGroupName FROM MarkGroupNames WHERE TemplateID = %d", g_prefs.m_TemplateList[i].m_ID);

		for (int j = 0; j < MAXMARKGROUPS; j++)
			g_prefs.m_TemplateList[i].m_markgroups[j] = _T("");

		try {
			if (!Rs.Open(CRecordset::snapshot, sSQL, CRecordset::readOnly)) {
				sErrorMessage.Format("Query failed - %s", (LPCSTR)sSQL);
				return FALSE;
			}

			int j = 0;
			while (!Rs.IsEOF() && j < MAXMARKGROUPS) {
				CString s;
				Rs.GetFieldValue((short)0, s);
				g_prefs.m_TemplateList[i].m_markgroups[j] = s;

				Rs.MoveNext();
			}
			Rs.Close();
		}
		catch (CDBException* e) {
			sErrorMessage.Format("Query failed - %s (%s)", (LPCSTR)e->m_strError, (LPCSTR)sSQL);
			e->Delete();
			Rs.Close();
			try {
				m_DBopen = FALSE;
				m_pDB->Close();
				return FALSE;
			}
			catch (CDBException* e) {
				// So what..! Go on!;
			}
			return FALSE;
		}

	}

	for (int i = 0; i < g_prefs.m_TemplateList.GetCount(); i++)
		g_prefs.m_TemplateList[i].m_pageformatID = RetrieveTemplatePageFormatName(g_prefs.m_TemplateList[i].m_ID, sErrorMessage);

	return TRUE;
}

int CDatabaseManager::LoadColorList(CString &sErrorMessage)
{
	BOOL foundJob = FALSE;

	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return -1;

	g_prefs.m_ColorList.RemoveAll();
	CRecordset Rs(m_pDB);

	CString sSQL = _T("SELECT ColorID,ColorName,C,M,Y,K,ColorOrder FROM ColorNames");
	try {
		if (!Rs.Open(CRecordset::snapshot, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("Query failed - %s", (LPCSTR)sSQL);
			return -1;
		}

		while (!Rs.IsEOF()) {
			CString s;
			CColorTableItem *pItem = new CColorTableItem();
			Rs.GetFieldValue((short)0, s);
			pItem->m_ID = atoi(s);
			Rs.GetFieldValue((short)1, pItem->m_name);

			Rs.GetFieldValue((short)2, s);
			pItem->m_c = atoi(s);
			Rs.GetFieldValue((short)3, s);
			pItem->m_m = atoi(s);
			Rs.GetFieldValue((short)4, s);
			pItem->m_y = atoi(s);
			Rs.GetFieldValue((short)5, s);
			pItem->m_k = atoi(s);
			Rs.GetFieldValue((short)6, s);
			pItem->m_colororder = atoi(s);

			g_prefs.m_ColorList.Add(*pItem);
			foundJob++;

			Rs.MoveNext();
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("Query failed - %s (%s)", (LPCSTR)e->m_strError, (LPCSTR)sSQL);
		LogprintfDB("Error : %s   (%s)", e->m_strError, sSQL);
		e->Delete();
		Rs.Close();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
			return -1;
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return -1;
	}
	return foundJob;
}

BOOL CDatabaseManager::LoadMarkGroupList(CString &sErrorMessage)
{
	sErrorMessage = _T("");
	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	CRecordset Rs(m_pDB);
	g_prefs.m_MarkGroupList.RemoveAll();
	CString sSQL;
	sSQL.Format("SELECT DISTINCT MarkGroupID,MarkGroupName FROM MarkGroupNames");

	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)sSQL);
			return FALSE;
		}

		while (!Rs.IsEOF()) {
			ITEMSTRUCT *pItem = new ITEMSTRUCT;
			CString s;
			Rs.GetFieldValue((short)0, s);
			pItem->m_ID = atoi(s);
			Rs.GetFieldValue((short)1, pItem->m_name);
			g_prefs.m_MarkGroupList.Add(*pItem);
			Rs.MoveNext();
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error : %s   (%s)", e->m_strError, sSQL);
		e->Delete();
		Rs.Close();

		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}
	return TRUE;
}

BOOL CDatabaseManager::LoadProofProcessList(CString &sErrorMessage)
{
	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	CRecordset Rs(m_pDB);

	g_prefs.m_ProofList.RemoveAll();

	CString sSQL, s;
	int fld = 0;

	sSQL = _T("SELECT ProofID,ProofName FROM ProofConfigurations");
	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)sSQL);
			return FALSE;
		}

		while (!Rs.IsEOF()) {
			ITEMSTRUCT *pItem = new ITEMSTRUCT();

			Rs.GetFieldValue((short)0, s);
			pItem->m_ID = atoi(s);
			Rs.GetFieldValue((short)1, pItem->m_name);
			g_prefs.m_ProofList.Add(*pItem);
			Rs.MoveNext();
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		e->Delete();
		Rs.Close();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}
	return TRUE;
}

BOOL CDatabaseManager::LoadDeviceList(CString &sErrorMessage)
{
	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	CRecordset Rs(m_pDB);

	g_prefs.m_DeviceList.RemoveAll();

	CString sSQL, s;
	int fld = 0;

	sSQL = _T("SELECT DeviceID,DeviceName,LocationID FROM DeviceConfigurations");
	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)sSQL);
			return FALSE;
		}

		while (!Rs.IsEOF()) {
			DEVICESTRUCT *pItem = new DEVICESTRUCT();

			Rs.GetFieldValue((short)0, s);
			pItem->m_ID = atoi(s);
			Rs.GetFieldValue((short)1, pItem->m_name);
			Rs.GetFieldValue((short)2, s);
			pItem->m_locationID = atoi(s);
			g_prefs.m_DeviceList.Add(*pItem);
			Rs.MoveNext();
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error : %s   (%s)", e->m_strError, sSQL);
		e->Delete();
		Rs.Close();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}
	return TRUE;
}

BOOL CDatabaseManager::LoadPressList(CString &sErrorMessage)
{
	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	CRecordset Rs(m_pDB);

	g_prefs.m_PressList.RemoveAll();

	CString sSQL, s;
	int fld = 0;

	sSQL = _T("SELECT PressID,PressName,LocationID FROM PressNames");
	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)sSQL);
			return FALSE;
		}

		while (!Rs.IsEOF()) {
			PRESSSTRUCT *pItem = new PRESSSTRUCT();

			Rs.GetFieldValue((short)0, s);
			pItem->m_ID = atoi(s);
			Rs.GetFieldValue((short)1, pItem->m_name);
			Rs.GetFieldValue((short)2, s);
			pItem->m_locationID = atoi(s);
			g_prefs.m_PressList.Add(*pItem);
			Rs.MoveNext();
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error : %s   (%s)", e->m_strError, sSQL);
		e->Delete();
		Rs.Close();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}

	for (int i = 0; i < g_prefs.m_PressList.GetCount(); i++)
		RetrievePressTowerList(g_prefs.m_PressList[i].m_ID, sErrorMessage);

	return TRUE;
}

BOOL CDatabaseManager::RetrievePressTowerList(int nPressID, CString &sErrorMessage)
{
	int ret = 0;
	CString sSQL;

	sErrorMessage = _T("");
	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	CRecordset Rs(m_pDB);

	sSQL.Format("SELECT DISTINCT TowerName FROM PressTowerNames WHERE PressID=%d", nPressID);

	PRESSSTRUCT *pPress = g_prefs.GetPressStruct(nPressID);

	if (pPress == NULL)
		return FALSE;

	pPress->m_numberoftowers = 0;

	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)sSQL);
			return FALSE;
		}

		while (!Rs.IsEOF()) {
			CString s;
			Rs.GetFieldValue((short)0, s);
			pPress->m_towerlist[pPress->m_numberoftowers++] = s;

			Rs.MoveNext();

		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error : %s   (%s)", e->m_strError, sSQL);
		e->Delete();
		Rs.Close();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}


	return TRUE;
}


BOOL CDatabaseManager::LoadGeneralPreferences(CString &sErrorMessage)
{
	sErrorMessage = _T("");
	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	CRecordset Rs(m_pDB);

	CString sSQL = _T("SELECT TOP 1 ServerName,ServerShare,ServerFilePath,ServerPreviewPath,ServerThumbnailPath,ServerLogPath,ServerConfigPath,ServerErrorPath,MainLocationID,PopupWarnings,LogEvents,UseInches,MinDiskSpace,EnableAutoCleanup,AutoCleanupDays,ServerUseCurrentUser,ServerUserName,ServerPassword FROM GeneralPreferences");

	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)sSQL);
			return FALSE;
		}

		if (!Rs.IsEOF()) {
			CString s;
			int fld = 0;
			Rs.GetFieldValue((short)fld++, g_prefs.m_serverName);
			Rs.GetFieldValue((short)fld++, g_prefs.m_serverShare);
			Rs.GetFieldValue((short)fld++, g_prefs.m_storagePath);
			if (g_prefs.m_storagePath.Right(1) != "\\")
				g_prefs.m_storagePath += _T("\\");
			Rs.GetFieldValue((short)fld++, g_prefs.m_previewPath);
			if (g_prefs.m_previewPath.Right(1) != "\\")
				g_prefs.m_previewPath += _T("\\");
			Rs.GetFieldValue((short)fld++, g_prefs.m_thumbnailPath);
			if (g_prefs.m_thumbnailPath.Right(1) != "\\")
				g_prefs.m_thumbnailPath += _T("\\");
			Rs.GetFieldValue((short)fld++, g_prefs.m_logFilePath);
			if (g_prefs.m_logFilePath.Right(1) != "\\")
				g_prefs.m_logFilePath += _T("\\");
			Rs.GetFieldValue((short)fld++, g_prefs.m_configPath);
			if (g_prefs.m_configPath.Right(1) != "\\")
				g_prefs.m_configPath += _T("\\");
			Rs.GetFieldValue((short)fld++, g_prefs.m_errorPath);
			if (g_prefs.m_errorPath.Right(1) != "\\")
				g_prefs.m_errorPath += _T("\\");

			Rs.GetFieldValue((short)fld++, s);
			g_prefs.m_mainlocationID = atoi(s);

			Rs.GetFieldValue((short)fld++, s);
			g_prefs.m_doWarn = atoi(s);

			Rs.GetFieldValue((short)fld++, s);
			g_prefs.m_logActions = atoi(s);

			Rs.GetFieldValue((short)fld++, s);
			g_prefs.m_useInches = atoi(s);

			Rs.GetFieldValue((short)fld++, s);
			g_prefs.m_minDiskSpaceMB = atoi(s);

			Rs.GetFieldValue((short)fld++, s);
			g_prefs.m_enableautodelete = atoi(s);

			Rs.GetFieldValue((short)fld++, s);
			g_prefs.m_autodeletedays = atoi(s);
			Rs.GetFieldValue((short)fld++, s);
			g_prefs.m_serverusecurrentuser = atoi(s);
			Rs.GetFieldValue((short)fld++, g_prefs.m_serverusername);
			Rs.GetFieldValue((short)fld++, g_prefs.m_serverpassword);
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error : %s   (%s)", e->m_strError, sSQL);
		e->Delete();
		Rs.Close();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}
	return TRUE;
}


int CDatabaseManager::RetrieveTemplatePageFormatName(int nTemplateID, CString &sErrorMessage)
{
	int nPageFormatID = 0;
	sErrorMessage = _T("");
	if (InitDB(sErrorMessage) == FALSE)
		return 0;

	CRecordset Rs(m_pDB);

	CString sSQL, s;
	sSQL.Format("SELECT PageFormatID FROM TemplatePageFormats WHERE TemplateID=%d", nTemplateID);

	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)sSQL);
			return 0;
		}

		if (!Rs.IsEOF()) {
			Rs.GetFieldValue((short)0, s);
			nPageFormatID = atoi(s);
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		LogprintfDB("Error : %s   (%s)", e->m_strError, sSQL);
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		e->Delete();
		Rs.Close();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return 0;
	}
	return nPageFormatID;


}

BOOL CDatabaseManager::RetrievePageFormatList(CString &sErrorMessage)
{

	sErrorMessage = _T("");
	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	CRecordset Rs(m_pDB);
	CString sSQL, s;

	sSQL.Format("SELECT DISTINCT PageFormatID,PageFormatName,Width,Height,Bleed,SnapModeEven,SnapModeOdd FROM PageFormatNames");

	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)sSQL);
			return FALSE;
		}

		g_prefs.m_PageFormatList.RemoveAll();

		while (!Rs.IsEOF()) {
			PAGEFORMATSTRUCT *pItem = new PAGEFORMATSTRUCT;
			Rs.GetFieldValue((short)0, s);
			pItem->m_ID = atoi(s);
			Rs.GetFieldValue((short)1, pItem->m_name);
			Rs.GetFieldValue((short)2, s);
			pItem->m_formatw = atof(s);
			Rs.GetFieldValue((short)3, s);
			pItem->m_formath = atof(s);
			Rs.GetFieldValue((short)4, s);
			pItem->m_bleed = atof(s);

			Rs.GetFieldValue((short)5, s);
			pItem->m_snapmodeeven = atoi(s);
			Rs.GetFieldValue((short)6, s);
			pItem->m_snapmodeodd = atoi(s);

			g_prefs.m_PageFormatList.Add(*pItem);

			Rs.MoveNext();
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error : %s   (%s)", e->m_strError, sSQL);
		e->Delete();
		Rs.Close();

		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}
	return TRUE;
}


BOOL CDatabaseManager::RetrievePublicationList(CString &sErrorMessage)
{
	sErrorMessage = _T("");
	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	CRecordset Rs(m_pDB);
	CString sSQL, s;

	sSQL.Format("SELECT DISTINCT PublicationID,Name,PageFormatID,TrimToFormat,LatestHour,DefaultProofID,DefaultHardProofID,DefaultApprove, UploadFolder,InputAlias FROM PublicationNames ORDER BY Name");

	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)sSQL);
			return FALSE;
		}

		while (!Rs.IsEOF()) {
			PUBLICATIONSTRUCT *pItem = new PUBLICATIONSTRUCT;
			Rs.GetFieldValue((short)0, s);
			pItem->m_ID = atoi(s);

			Rs.GetFieldValue((short)1, pItem->m_name);

			Rs.GetFieldValue((short)2, s);
			pItem->m_PageFormatID = atoi(s);
			Rs.GetFieldValue((short)3, s);
			pItem->m_TrimToFormat = atoi(s);
			Rs.GetFieldValue((short)4, s);
			pItem->m_LatestHour = atof(s);

			Rs.GetFieldValue((short)5, s);
			pItem->m_ProofID = atoi(s);
			Rs.GetFieldValue((short)6, s);
			pItem->m_HardProofID = atoi(s);
			Rs.GetFieldValue((short)7, s);
			pItem->m_Approve = atoi(s);

			Rs.GetFieldValue((short)8, pItem->m_uploadfolder);

			pItem->m_annumtext = _T("");
			pItem->m_releasedays = 0;
			pItem->m_releasetimehour = 0;
			pItem->m_releasetimeminute = 0;
			pItem->m_autoapprove = FALSE;
			pItem->m_rejectsubeditionpages = FALSE;

			int q = pItem->m_uploadfolder.Find(";");
			if (q != -1) {
				CStringArray sArr;
				q = g_util.StringSplitter(pItem->m_uploadfolder, ";", sArr);
				if (sArr.GetCount() > 0)
					pItem->m_uploadfolder = sArr[0];
				if (sArr.GetCount() > 1)
					pItem->m_annumtext = sArr[1];
				if (sArr.GetCount() > 2)
					pItem->m_releasedays = atoi(sArr[2]);
				if (sArr.GetCount() > 3)
					pItem->m_releasetimehour = atoi(sArr[3]);
				if (sArr.GetCount() > 4)
					pItem->m_releasetimeminute = atoi(sArr[4]);

				if (sArr.GetCount() > 5)
					pItem->m_autoapprove = atoi(sArr[5]);
				if (sArr.GetCount() > 6)
					pItem->m_rejectsubeditionpages = atoi(sArr[6]);
			}
			Rs.GetFieldValue((short)9, pItem->m_alias);

			g_prefs.m_PublicationList.Add(*pItem);

			Rs.MoveNext();
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error : %s   (%s)", e->m_strError, sSQL);
		e->Delete();
		Rs.Close();

		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}



	return TRUE;
}


BOOL CDatabaseManager::RetrievePublicationPressDefaults(PUBLICATIONSTRUCT *pPublication, CString &sErrorMessage)
{

	BOOL hasRipSetup = FieldExists("PublicationTemplates", "RipSetup", sErrorMessage);
	BOOL hasStackPosition = FieldExists("PublicationTemplates", "StackPosition", sErrorMessage);
	BOOL hasMiscInt = FieldExists("PublicationTemplates", "MiscInt", sErrorMessage);
	BOOL hasDeviceID = FieldExists("PublicationTemplates", "DeviceID", sErrorMessage);
	BOOL hasHold = FieldExists("PublicationTemplates", "Hold", sErrorMessage);
	BOOL hasPagination = FieldExists("PublicationTemplates", "PaginationMode", sErrorMessage);

	sErrorMessage = _T("");
	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	CRecordset Rs(m_pDB);
	CString sSQL, s;

	if (hasPagination)
		sSQL.Format("SELECT DISTINCT PressID,TemplateID,Copies,Markgroups,FanOut,Priority,ISNULL(RipSetup,''),ISNULL(StackPosition,''),ISNULL(MiscInt,0),ISNULL(MiscString,''),ISNULL(DeviceID,0),ISNULL(Hold,0),ISNULL(TxName,''),ISNULL(InsertedSections,0), ISNULL(PressSpecific,0),ISNULL(PaginationMode,0),ISNULL(SeparateRuns,0) FROM PublicationTemplates WHERE PublicationID=%d ORDER BY PressID", pPublication->m_ID);
	else if (hasHold)
		sSQL.Format("SELECT DISTINCT PressID,TemplateID,Copies,Markgroups,FanOut,Priority,ISNULL(RipSetup,''),ISNULL(StackPosition,''),ISNULL(MiscInt,0),ISNULL(MiscString,''),ISNULL(DeviceID,0),ISNULL(Hold,0),ISNULL(TxName,'') FROM PublicationTemplates WHERE PublicationID=%d ORDER BY PressID", pPublication->m_ID);
	else if (hasDeviceID)
		sSQL.Format("SELECT DISTINCT PressID,TemplateID,Copies,Markgroups,FanOut,Priority,ISNULL(RipSetup,''),ISNULL(StackPosition,''),ISNULL(MiscInt,0),ISNULL(MiscString,''),ISNULL(DeviceID,0) FROM PublicationTemplates WHERE PublicationID=%d ORDER BY PressID", pPublication->m_ID);
	else if (hasMiscInt)
		sSQL.Format("SELECT DISTINCT PressID,TemplateID,Copies,Markgroups,FanOut,Priority,ISNULL(RipSetup,''),ISNULL(StackPosition,''),ISNULL(MiscInt,0),ISNULL(MiscString,'') FROM PublicationTemplates WHERE PublicationID=%d ORDER BY PressID", pPublication->m_ID);
	else if (hasStackPosition)
		sSQL.Format("SELECT DISTINCT PressID,TemplateID,Copies,Markgroups,FanOut,Priority,ISNULL(RipSetup,''),ISNULL(StackPosition,'') FROM PublicationTemplates WHERE PublicationID=%d ORDER BY PressID", pPublication->m_ID);
	else if (hasRipSetup)
		sSQL.Format("SELECT DISTINCT PressID,TemplateID,Copies,Markgroups,FanOut,Priority,ISNULL(RipSetup,'') FROM PublicationTemplates WHERE PublicationID=%d ORDER BY PressID", pPublication->m_ID);
	else
		sSQL.Format("SELECT DISTINCT PressID,TemplateID,Copies,Markgroups,FanOut,Priority FROM PublicationTemplates WHERE PublicationID=%d ORDER BY PressID", pPublication->m_ID);

	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)sSQL);
			return FALSE;
		}

		for (int i = 0; i < g_prefs.m_PressList.GetCount(); i++) {
			pPublication->m_PressDefaults[i].m_PressID = g_prefs.m_PressList[i].m_ID;
			pPublication->m_PressDefaults[i].m_defaultpress = 0;
			pPublication->m_PressDefaults[i].m_DefaultTemplateID = 0;
			pPublication->m_PressDefaults[i].m_DefaultFlatProofTemplateID = 0;
			pPublication->m_PressDefaults[i].m_deviceID = 0;
			pPublication->m_PressDefaults[i].m_DefaultNumberOfMarkGroups = 0;
			pPublication->m_PressDefaults[i].m_DefaultPressTowerName = _T("");
			pPublication->m_PressDefaults[i].m_DefaultCopies = 1;
			pPublication->m_PressDefaults[i].m_DefaultPriority = 50;
			pPublication->m_PressDefaults[i].m_RipSetup = _T("");
			pPublication->m_PressDefaults[i].m_DefaultStackPosition = _T("");
			pPublication->m_PressDefaults[i].m_allowautoplanning = FALSE;
			pPublication->m_PressDefaults[i].m_txname = _T("");
			pPublication->m_PressDefaults[i].m_hold = FALSE;
			pPublication->m_PressDefaults[i].m_publicationID = pPublication->m_ID;
			pPublication->m_PressDefaults[i].m_paginationmode = 0;
			pPublication->m_PressDefaults[i].m_insertedsections = FALSE;
			pPublication->m_PressDefaults[i].m_separateruns = FALSE;
			pPublication->m_PressDefaults[i].m_pressspecificpages = FALSE;
			pPublication->m_PressDefaults[i].m_pressnotused = FALSE;

		}


		CStringArray arrMarkGroups;

		while (!Rs.IsEOF()) {

			Rs.GetFieldValue((short)0, s);
			if (atoi(s) <= 0 || atoi(s) > 10000)
				Rs.MoveNext();

			int nPressID = atoi(s);
			int nPressIdx = 0;
			for (int i = 0; i < g_prefs.m_PressList.GetCount(); i++) {
				if (nPressID == g_prefs.m_PressList[i].m_ID) {
					nPressIdx = i;
					break;
				}
			}


			pPublication->m_PressDefaults[nPressIdx].m_defaultpress = nPressID;
			pPublication->m_PressDefaults[nPressIdx].m_PressID = atoi(s);

			Rs.GetFieldValue((short)1, s);
			if (atoi(s) < 0 || atoi(s) > 10000)
				s = "0";
			pPublication->m_PressDefaults[nPressIdx].m_DefaultTemplateID = atoi(s);

			Rs.GetFieldValue((short)2, s);
			if (atoi(s) < 0 || atoi(s) > 10000)
				s = "1";
			pPublication->m_PressDefaults[nPressIdx].m_DefaultCopies = atoi(s);

			Rs.GetFieldValue((short)3, s);
			g_util.StringSplitter(s, ",", arrMarkGroups);

			pPublication->m_PressDefaults[nPressIdx].m_DefaultNumberOfMarkGroups = (int)arrMarkGroups.GetCount();
			for (int i = 0; i < arrMarkGroups.GetCount(); i++)
				pPublication->m_PressDefaults[nPressIdx].m_DefaultMarkGroupList[i] = g_prefs.GetMarkGroupName(atoi(arrMarkGroups[i]));

			Rs.GetFieldValue((short)4, s);
			pPublication->m_PressDefaults[nPressIdx].m_DefaultPressTowerName = s;

			Rs.GetFieldValue((short)5, s);
			if (atoi(s) < 0 || atoi(s) > 10000)
				s = "50";

			pPublication->m_PressDefaults[nPressIdx].m_DefaultPriority = atoi(s);

			if (hasRipSetup) {
				Rs.GetFieldValue((short)6, s);
				pPublication->m_PressDefaults[nPressIdx].m_RipSetup = s;
			}
			else
				pPublication->m_PressDefaults[nPressIdx].m_RipSetup = _T("");

			if (hasStackPosition) {
				Rs.GetFieldValue((short)7, s);
				pPublication->m_PressDefaults[nPressIdx].m_DefaultStackPosition = s;
			}
			else
				pPublication->m_PressDefaults[nPressIdx].m_DefaultStackPosition = _T("");

			if (hasMiscInt) {
				Rs.GetFieldValue((short)8, s);
				if (atoi(s) < 0 || atoi(s) > 10000)
					s = "0";

				pPublication->m_PressDefaults[nPressIdx].m_allowautoplanning = (atoi(s) & 0x01) > 0 ? TRUE : FALSE;
				pPublication->m_PressDefaults[nPressIdx].m_pressnotused = (atoi(s) & 0x02) > 0 ? TRUE : FALSE;
				Rs.GetFieldValue((short)9, s);
				if (atoi(s) < 0 || atoi(s) > 10000)
					s = "0";
				pPublication->m_PressDefaults[nPressIdx].m_DefaultFlatProofTemplateID = atoi(s);
			}
			else {
				pPublication->m_PressDefaults[nPressIdx].m_allowautoplanning = FALSE;
				pPublication->m_PressDefaults[nPressIdx].m_DefaultFlatProofTemplateID = 0;
				pPublication->m_PressDefaults[nPressIdx].m_pressnotused = FALSE;
			}

			if (hasDeviceID) {
				Rs.GetFieldValue((short)10, s);
				if (atoi(s) < 0 || atoi(s) > 10000)
					s = "0";
				pPublication->m_PressDefaults[nPressIdx].m_deviceID = atoi(s);
			}
			else
				pPublication->m_PressDefaults[nPressIdx].m_deviceID = 0;


			if (hasHold) {
				Rs.GetFieldValue((short)11, s);
				if (atoi(s) < 0 || atoi(s) > 10000)
					s = "0";
				pPublication->m_PressDefaults[nPressIdx].m_hold = atoi(s);
				Rs.GetFieldValue((short)12, s);
				pPublication->m_PressDefaults[nPressIdx].m_txname = s;
			}
			else {
				pPublication->m_PressDefaults[nPressIdx].m_txname = _T("");
				pPublication->m_PressDefaults[nPressIdx].m_hold = FALSE;
			}

			if (hasPagination) {
				Rs.GetFieldValue((short)13, s);
				if (atoi(s) < 0 || atoi(s) > 10000)
					s = "0";
				pPublication->m_PressDefaults[nPressIdx].m_insertedsections = atoi(s);
				Rs.GetFieldValue((short)14, s);
				if (atoi(s) < 0 || atoi(s) > 10000)
					s = "0";
				pPublication->m_PressDefaults[nPressIdx].m_pressspecificpages = atoi(s);
				Rs.GetFieldValue((short)15, s);
				if (atoi(s) < 0 || atoi(s) > 10000)
					s = "0";
				pPublication->m_PressDefaults[nPressIdx].m_paginationmode = atoi(s);
				Rs.GetFieldValue((short)16, s);
				if (atoi(s) < 0 || atoi(s) > 10000)
					s = "0";
				pPublication->m_PressDefaults[nPressIdx].m_separateruns = atoi(s);
			}
			else {
				pPublication->m_PressDefaults[nPressIdx].m_insertedsections = FALSE;
				pPublication->m_PressDefaults[nPressIdx].m_pressspecificpages = FALSE;
				pPublication->m_PressDefaults[nPressIdx].m_paginationmode = 0;
				pPublication->m_PressDefaults[nPressIdx].m_separateruns = FALSE;
			}
			Rs.MoveNext();
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error : %s   (%s)", e->m_strError, sSQL);
		e->Delete();
		Rs.Close();

		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}
	return TRUE;
}

CString CDatabaseManager::LoadNewPublicationName(int nID, CString &sErrorMessage)
{
	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return _T("");

	CRecordset Rs(m_pDB);

	CString sSQL, s;
	CString sPub = "";
	PUBLICATIONSTRUCT *pItem = NULL;

	sSQL.Format("SELECT Name,PageFormatID,TrimToFormat,LatestHour,DefaultProofID,DefaultHardProofID,DefaultApprove,UploadFolder,InputAlias FROM PublicationNames WHERE PublicationID=%d", nID);

	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("Query failed - %s", (LPCSTR)sSQL);
			return "";
		}

		if (!Rs.IsEOF()) {
			pItem = new PUBLICATIONSTRUCT;
			Rs.GetFieldValue((short)0, sPub);
			pItem->m_ID = nID;
			pItem->m_name = sPub;

			Rs.GetFieldValue((short)1, s);
			pItem->m_PageFormatID = atoi(s);
			Rs.GetFieldValue((short)2, s);
			pItem->m_TrimToFormat = atoi(s);
			Rs.GetFieldValue((short)3, s);
			pItem->m_LatestHour = atof(s);

			Rs.GetFieldValue((short)4, s);
			pItem->m_ProofID = atoi(s);
			Rs.GetFieldValue((short)5, s);
			pItem->m_HardProofID = atoi(s);
			Rs.GetFieldValue((short)6, s);
			pItem->m_Approve = atoi(s);
			Rs.GetFieldValue((short)7, pItem->m_uploadfolder);

			pItem->m_annumtext = _T("");
			pItem->m_releasedays = 0;
			pItem->m_releasetimehour = 0;
			pItem->m_releasetimeminute = 0;
			pItem->m_autoapprove = FALSE;
			pItem->m_rejectsubeditionpages = FALSE;

			int q = pItem->m_uploadfolder.Find(";");
			if (q != -1) {
				CStringArray sArr;
				q = g_util.StringSplitter(pItem->m_uploadfolder, ";", sArr);
				if (sArr.GetCount() > 0)
					pItem->m_uploadfolder = sArr[0];
				if (sArr.GetCount() > 1)
					pItem->m_annumtext = sArr[1];
				if (sArr.GetCount() > 2)
					pItem->m_releasedays = atoi(sArr[2]);
				if (sArr.GetCount() > 3)
					pItem->m_releasetimehour = atoi(sArr[3]);
				if (sArr.GetCount() > 4)
					pItem->m_releasetimeminute = atoi(sArr[4]);

				if (sArr.GetCount() > 5)
					pItem->m_autoapprove = atoi(sArr[5]);
				if (sArr.GetCount() > 6)
					pItem->m_rejectsubeditionpages = atoi(sArr[6]);
			}
			Rs.GetFieldValue((short)8, pItem->m_alias);

			g_prefs.m_PublicationList.Add(*pItem);
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error: %s   (%s)", e->m_strError, sSQL);

		e->Delete();
		Rs.Close();

		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return "";
	}

	if (nID > 0 && pItem != NULL) {
		g_prefs.ResetPublicationPressDefaults(nID);
		RetrievePublicationPressDefaults(pItem, sErrorMessage);
	}

	return sPub;
}

PUBLICATIONSTRUCT *CDatabaseManager::LoadPublicationStruct(int nID, CString &sErrorMessage)
{
	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return NULL;

	CRecordset Rs(m_pDB);

	CString sSQL, s;
	CString sPub = "";
	PUBLICATIONSTRUCT *pItem = NULL;

	sSQL.Format("SELECT Name,PageFormatID,TrimToFormat,LatestHour,DefaultProofID,DefaultHardProofID,DefaultApprove,UploadFolder,InputAlias FROM PublicationNames WHERE PublicationID=%d", nID);

	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("Query failed - %s", (LPCSTR)sSQL);
			return NULL;
		}

		if (!Rs.IsEOF()) {
			pItem = new PUBLICATIONSTRUCT;
			Rs.GetFieldValue((short)0, sPub);
			pItem->m_ID = nID;
			pItem->m_name = sPub;

			Rs.GetFieldValue((short)1, s);
			pItem->m_PageFormatID = atoi(s);
			Rs.GetFieldValue((short)2, s);
			pItem->m_TrimToFormat = atoi(s);
			Rs.GetFieldValue((short)3, s);
			pItem->m_LatestHour = atof(s);

			Rs.GetFieldValue((short)4, s);
			pItem->m_ProofID = atoi(s);
			Rs.GetFieldValue((short)5, s);
			pItem->m_HardProofID = atoi(s);
			Rs.GetFieldValue((short)6, s);
			pItem->m_Approve = atoi(s);
			Rs.GetFieldValue((short)7, pItem->m_uploadfolder);

			pItem->m_annumtext = _T("");
			pItem->m_releasedays = 0;
			pItem->m_releasetimehour = 0;
			pItem->m_releasetimeminute = 0;
			pItem->m_autoapprove = FALSE;
			pItem->m_rejectsubeditionpages = FALSE;

			int q = pItem->m_uploadfolder.Find(";");
			if (q != -1) {
				CStringArray sArr;
				q = g_util.StringSplitter(pItem->m_uploadfolder, ";", sArr);
				if (sArr.GetCount() > 0)
					pItem->m_uploadfolder = sArr[0];
				if (sArr.GetCount() > 1)
					pItem->m_annumtext = sArr[1];
				if (sArr.GetCount() > 2)
					pItem->m_releasedays = atoi(sArr[2]);
				if (sArr.GetCount() > 3)
					pItem->m_releasetimehour = atoi(sArr[3]);
				if (sArr.GetCount() > 4)
					pItem->m_releasetimeminute = atoi(sArr[4]);

				if (sArr.GetCount() > 5)
					pItem->m_autoapprove = atoi(sArr[5]);
				if (sArr.GetCount() > 6)
					pItem->m_rejectsubeditionpages = atoi(sArr[6]);
			}
			Rs.GetFieldValue((short)8, pItem->m_alias);
	
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error: %s   (%s)", e->m_strError, sSQL);

		e->Delete();
		Rs.Close();

		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return NULL;
	}

	if (nID > 0 && pItem != NULL) {
		g_prefs.ResetPublicationPressDefaults(pItem);
		RetrievePublicationPressDefaults(pItem, sErrorMessage);
	}

	return pItem;
}



int CDatabaseManager::GetPublicationID(CString sName, CString &sErrorMessage)
{

	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return 0;

	CRecordset Rs(m_pDB);

	CString sSQL, s;
	int nID = 0;

	sSQL.Format("SELECT PublicationID FROM PublicationNames WHERE ([Name]='%s' OR InputAlias='%s')", sName, sName);
	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("Query failed - %s", (LPCSTR)sSQL);
			return 0;
		}

		if (!Rs.IsEOF()) {
			Rs.GetFieldValue((short)0, s);
			nID = atoi(s);
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error: %s   (%s)", e->m_strError, sSQL);

		e->Delete();
		Rs.Close();

		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return 0;
	}

	return nID;
}

int CDatabaseManager::GetEditionID(CString sName, CString &sErrorMessage)
{

	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return 0;

	CRecordset Rs(m_pDB);

	CString sSQL, s;
	int nID = 0;

	sSQL.Format("SELECT EditionID FROM EditionNames WHERE [Name]='%s'", sName);
	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("Query failed - %s", (LPCSTR)sSQL);
			return 0;
		}

		if (!Rs.IsEOF()) {
			Rs.GetFieldValue((short)0, s);
			nID = atoi(s);
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error: %s   (%s)", e->m_strError, sSQL);

		e->Delete();
		Rs.Close();

		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return 0;
	}

	return nID;
}

int CDatabaseManager::GetSectionID(CString sName, CString &sErrorMessage)
{

	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return 0;

	CRecordset Rs(m_pDB);

	CString sSQL, s;
	int nID = 0;

	sSQL.Format("SELECT SectionID FROM SectionNames WHERE [Name]='%s'", sName);
	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("Query failed - %s", (LPCSTR)sSQL);
			return 0;
		}

		if (!Rs.IsEOF()) {
			Rs.GetFieldValue((short)0, s);
			nID = atoi(s);
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error: %s   (%s)", e->m_strError, sSQL);

		e->Delete();
		Rs.Close();

		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return 0;
	}

	return nID;
}

int CDatabaseManager::LoadNewPublicationID(CString sName, CString &sErrorMessage)
{
	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return 0;

	CRecordset Rs(m_pDB);

	CString sSQL, s;
	int nID = 0;
	PUBLICATIONSTRUCT *pItem = NULL;

	sSQL.Format("SELECT PublicationID,PageFormatID,TrimToFormat,LatestHour,DefaultProofID,DefaultHardProofID,DefaultApprove,UploadFolder,InputAlias FROM PublicationNames WHERE ([Name]='%s' OR InputAlias='%s')", sName, sName);

	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("Query failed - %s", (LPCSTR)sSQL);
			return 0;
		}

		if (!Rs.IsEOF()) {
			pItem = new PUBLICATIONSTRUCT;
			Rs.GetFieldValue((short)0, s);
			nID = atoi(s);
			pItem->m_ID = nID;
			pItem->m_name = sName;

			Rs.GetFieldValue((short)1, s);
			pItem->m_PageFormatID = atoi(s);
			Rs.GetFieldValue((short)2, s);
			pItem->m_TrimToFormat = atoi(s);
			Rs.GetFieldValue((short)3, s);
			pItem->m_LatestHour = atof(s);

			Rs.GetFieldValue((short)4, s);
			pItem->m_ProofID = atoi(s);
			Rs.GetFieldValue((short)5, s);
			pItem->m_HardProofID = atoi(s);
			Rs.GetFieldValue((short)6, s);
			pItem->m_Approve = atoi(s);
			Rs.GetFieldValue((short)7, pItem->m_uploadfolder);

			pItem->m_annumtext = _T("");
			pItem->m_releasedays = 0;
			pItem->m_releasetimehour = 0;
			pItem->m_releasetimeminute = 0;
			pItem->m_autoapprove = FALSE;
			pItem->m_rejectsubeditionpages = FALSE;

			int q = pItem->m_uploadfolder.Find(";");
			if (q != -1) {
				CStringArray sArr;
				q = g_util.StringSplitter(pItem->m_uploadfolder, ";", sArr);
				if (sArr.GetCount() > 0)
					pItem->m_uploadfolder = sArr[0];
				if (sArr.GetCount() > 1)
					pItem->m_annumtext = sArr[1];
				if (sArr.GetCount() > 2)
					pItem->m_releasedays = atoi(sArr[2]);
				if (sArr.GetCount() > 3)
					pItem->m_releasetimehour = atoi(sArr[3]);
				if (sArr.GetCount() > 4)
					pItem->m_releasetimeminute = atoi(sArr[4]);

				if (sArr.GetCount() > 5)
					pItem->m_autoapprove = atoi(sArr[5]);
				if (sArr.GetCount() > 6)
					pItem->m_rejectsubeditionpages = atoi(sArr[6]);
			}
			Rs.GetFieldValue((short)8, pItem->m_alias);


			g_prefs.m_PublicationList.Add(*pItem);
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error: %s   (%s)", e->m_strError, sSQL);

		e->Delete();
		Rs.Close();

		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return 0;
	}

	if (nID > 0 && pItem != NULL) {
		g_prefs.ResetPublicationPressDefaults(nID);
		RetrievePublicationPressDefaults(pItem, sErrorMessage);		
	}

	return nID;
}

int CDatabaseManager::LoadNewEditionID(CString sName, CString &sErrorMessage)
{
	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return 0;

	CRecordset Rs(m_pDB);

	CString sSQL, s;
	int nID = 0;

	sSQL.Format("SELECT EditionID FROM EditionNames WHERE [Name]='%s'", sName);

	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("Query failed - %s", (LPCSTR)sSQL);
			return 0;
		}

		if (!Rs.IsEOF()) {
			EDITIONSTRUCT *pItem = new EDITIONSTRUCT;
			Rs.GetFieldValue((short)0, s);
			nID = atoi(s);
			pItem->m_ID = nID;
			pItem->m_name = sName;
			pItem->m_iscommon = FALSE;
			pItem->m_level = 1;
			pItem->m_subofeditionID = 1;
			g_prefs.m_EditionList.Add(*pItem);
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error: %s   (%s)", e->m_strError, sSQL);

		e->Delete();
		Rs.Close();

		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return 0;
	}

	return nID;
}

int CDatabaseManager::LoadNewSectionID(CString sName, CString &sErrorMessage)
{
	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return 0;

	CRecordset Rs(m_pDB);

	CString sSQL, s;
	int nID = 0;

	sSQL.Format("SELECT SectionID FROM SectionNames WHERE [Name]='%s'", sName);

	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("Query failed - %s", (LPCSTR)sSQL);
			return 0;
		}

		if (!Rs.IsEOF()) {
			ITEMSTRUCT *pItem = new ITEMSTRUCT;
			Rs.GetFieldValue((short)0, s);
			nID = atoi(s);
			pItem->m_ID = nID;
			pItem->m_name = sName;
			g_prefs.m_SectionList.Add(*pItem);
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error: %s   (%s)", e->m_strError, sSQL);

		e->Delete();
		Rs.Close();

		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return 0;
	}

	return nID;
}


BOOL CDatabaseManager::InsertDBName(CString sNameTable, CString sIDFieldName, CString sNameFieldName, int nID, CString sName, CString &sErrorMessage)
{
	CString sSQL;

	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return -1;

	sSQL.Format("INSERT INTO %s (%s, %s) VALUES (%d, '%s')",
		sNameTable, sIDFieldName, sNameFieldName, nID, sName);
	try {
		m_pDB->ExecuteSQL(sSQL);
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Insert failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error : %s   (%s)", e->m_strError, sSQL);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}
	UpdateNotificationTable(sNameTable, sErrorMessage);
	return TRUE;
}

BOOL CDatabaseManager::InsertPublicationNames(int nPublicationID, CString sName, int nPageFormatID, BOOL bTrimToPageformat, double fLatestHours, int nProofID, int nHardProofID, int nApproveMode, CString &sErrorMessage)
{
	CString sSQL;

	sSQL.Format("INSERT INTO PublicationNames (PublicationID, Name,PageFormatID,TrimToFormat,LatestHour,DefaultProofID,DefaultHardProofID,DefaultApprove) VALUES (%d,'%s',%d,%d,%.4f,%d,%d,%d)",
		nPublicationID, sName, nPageFormatID, bTrimToPageformat, fLatestHours, nProofID, nHardProofID, nApproveMode);

	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	try {
		m_pDB->ExecuteSQL(sSQL);
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Delete failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error: %s   (%s)", e->m_strError, sSQL);

		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}

	sSQL.Format("UPDATE ChangeNotification SET ChangeID = ChangeID+1 WHERE TableName='PublicationNames'");
	try {
		m_pDB->ExecuteSQL(sSQL);
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error: %s   (%s)", e->m_strError, sSQL);

		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}
	return TRUE;
}


BOOL CDatabaseManager::InsertEditionName(int nID, CString sNewName, BOOL isCommon, int MasterID, CString &sErrorMessage)
{
	CString sSQL;
	sSQL.Format("INSERT INTO EditionNames (EditionID,Name,IsCommonEdition,SubOfEditionID) VALUES (%d,'%s',%d, %d)",
		nID, sNewName, isCommon, MasterID);

	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	try {
		m_pDB->ExecuteSQL(sSQL);
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error : %s   (%s)", e->m_strError, sSQL);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}
	UpdateNotificationTable(_T("EditionNames"), sErrorMessage);
	return TRUE;
}

int CDatabaseManager::RetrieveMaxValueCount(CString sID, CString sIDtable, CString &sErrorMessage)
{
	int nMax = 0;

	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return -1;

	CRecordset Rs(m_pDB);

	CString sSQL;
	sSQL.Format("SELECT MAX(%s) FROM %s", sID, sIDtable);

	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)sSQL);
			return -1;
		}

		if (!Rs.IsEOF()) {
			CString s;
			Rs.GetFieldValue((short)0, s);
			nMax = atoi(s);
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error : %s   (%s)", e->m_strError, sSQL);
		e->Delete();
		Rs.Close();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return -1;
	}
	return nMax;
}

int CDatabaseManager::RetrieveCustomerID(CString sCustomerName, CString &sErrorMessage)
{
	int nID = 0;

	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return -1;

	CRecordset Rs(m_pDB);

	CString sSQL;
	sSQL.Format("SELECT CustomerID FROM CustomerNames WHERE CustomerName='%s'", sCustomerName);

	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)sSQL);
			return -1;
		}

		if (!Rs.IsEOF()) {
			CString s;
			Rs.GetFieldValue((short)0, s);
			nID = atoi(s);
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error : %s   (%s)", e->m_strError, sSQL);
		e->Delete();
		Rs.Close();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return -1;
	}
	return nID;
}

int CDatabaseManager::InsertCustomerName(CString sCustomerName, CString sCustomerAlias, CString &sErrorMessage)
{
	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return -1;

	int nID = RetrieveNextValueCount("CustomerID", "CustomerNames", sErrorMessage);
	if (nID < 0)
		return -1;

	if (nID == 0)
		nID = 1;

	CString sSQL;
	sSQL.Format("INSERT INTO CustomerNames (CustomerID,CustomerName,CustomerAlias,CustomerEmail) VALUES (%d,'%s','%s','')", nID, sCustomerName, sCustomerAlias);

	BOOL bSuccess = FALSE;
	int nRetries = g_prefs.m_nQueryRetries;

	while (!bSuccess && nRetries-- > 0) {

		try {
			if (InitDB(sErrorMessage) == FALSE) {
				Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
				continue;
			}

			m_pDB->ExecuteSQL(sSQL);
			bSuccess = TRUE;
		}
		catch (CDBException* e) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
			LogprintfDB("Error Try %d - : %s   (%s)", nRetries, e->m_strError, sSQL);
			e->Delete();
			try {
				m_DBopen = FALSE;
				m_pDB->Close();
			}
			catch (CDBException* e) {
				// So what..! Go on!;
			}
			Sleep(g_prefs.m_QueryBackoffTime);
		}
	}

	if (bSuccess)
		UpdateNotificationTable("CustomerNames", sErrorMessage);

	return bSuccess ? nID : -1;
}

BOOL CDatabaseManager::DeleteInputAlias(CString sType, CString sShortName, CString &sErrorMessage)
{
	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return -1;

	CString sSQL;
	sSQL.Format("DELETE FROM InputAliases WHERE Type='%s' AND ShortName='%s'", sType, sShortName);

	BOOL bSuccess = FALSE;
	int nRetries = g_prefs.m_nQueryRetries;

	while (!bSuccess && nRetries-- > 0) {

		try {
			if (InitDB(sErrorMessage) == FALSE) {
				Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
				continue;
			}

			m_pDB->ExecuteSQL(sSQL);
			bSuccess = TRUE;
		}
		catch (CDBException* e) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
			LogprintfDB("Error Try %d - : %s   (%s)", nRetries, e->m_strError, sSQL);
			e->Delete();
			try {
				m_DBopen = FALSE;
				m_pDB->Close();
			}
			catch (CDBException* e) {
				// So what..! Go on!;
			}
			Sleep(g_prefs.m_QueryBackoffTime);
		}
	}

	if (bSuccess)
		UpdateNotificationTable("InputAlias", sErrorMessage);

	return bSuccess;
}




int CDatabaseManager::RetrieveNextValueCount(CString sID, CString sIDtable, CString &sErrorMessage)
{
	int nMax = 0;

	sErrorMessage = _T("");
	if (InitDB(sErrorMessage) == FALSE)
		return -1;

	CRecordset Rs(m_pDB);

	CString sSQL;

	sSQL.Format("SELECT TOP 1 MIN(C1.%s+1) FROM %s AS C1 WHERE NOT EXISTS (SELECT C2.%s FROM %s AS C2 WHERE C1.%s+1=C2.%s)",
		sID, sIDtable, sID, sIDtable, sID, sID);
	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)sSQL);
			return -1;
		}

		if (!Rs.IsEOF()) {
			CString s;
			Rs.GetFieldValue((short)0, s);
			nMax = atoi(s);
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error : %s   (%s)", e->m_strError, sSQL);
		e->Delete();
		Rs.Close();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return -1;
	}
	return nMax;
}

BOOL CDatabaseManager::UpdatePublicationName(CString sNameBefore, CString sNewName, CString &sErrorMessage)
{
	CString sSQL;
	CUtils util;

	sSQL.Format("UPDATE PublicationNames SET [Name]='%s' WHERE [Name]='%s'", sNewName, sNameBefore);
	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	try {
		m_pDB->ExecuteSQL(sSQL);
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error: %s   (%s)", e->m_strError, sSQL);

		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}

	return TRUE;
}


BOOL CDatabaseManager::UpdateAliasName(CString sTypeName, CString sLongName, CString sShortName, CString &sErrorMessage)
{
	CString sSQL;
	CUtils util;

	sSQL.Format("DELETE FROM InputAliases WHERE Type='%s' AND LongName='%s'", sTypeName, sLongName);

	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	try {
		m_pDB->ExecuteSQL(sSQL);
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Delete failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error: %s   (%s)", e->m_strError, sSQL);

		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}

	if (sShortName == _T(""))
		return TRUE;

	CStringArray sArr;

	int n = util.StringSplitter(sShortName, _T(","), sArr);
	for (int i = 0; i < n; i++) {
		sSQL.Format("INSERT INTO InputAliases (Type,LongName,ShortName) VALUES ('%s','%s','%s')", sTypeName, sLongName, sArr[i]);

		try {
			m_pDB->ExecuteSQL(sSQL);
		}
		catch (CDBException* e) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Insert failed - %s", (LPCSTR)e->m_strError);
			LogprintfDB("Error: %s   (%s)", e->m_strError, sSQL);

			e->Delete();
			try {
				m_DBopen = FALSE;
				m_pDB->Close();
			}
			catch (CDBException* e) {
				// So what..! Go on!;
			}
			return FALSE;
		}
	}
	UpdateNotificationTable(_T("InputAliases"), sErrorMessage);

	return TRUE;
}


BOOL CDatabaseManager::UpdateNotificationTable(CString sTableName, CString &sErrorMessage)
{
	CString sSQL;
	sSQL.Format("UPDATE ChangeNotification SET ChangeID=ChangeID+1 WHERE TableName='%s'", sTableName);

	BOOL bSuccess = FALSE;
	int nRetries = g_prefs.m_nQueryRetries;

	while (!bSuccess && nRetries-- > 0) {

		sErrorMessage.Format("");
		if (InitDB(sErrorMessage) == FALSE) {
			Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
			continue;
		}

		try {
			m_pDB->ExecuteSQL(sSQL);
			bSuccess = TRUE;
		}
		catch (CDBException* e) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Delete failed - %s", (LPCSTR)e->m_strError);
			LogprintfDB("Error : %s   (%s)", e->m_strError, sSQL);
			e->Delete();
			try {
				m_DBopen = FALSE;
				m_pDB->Close();
			}
			catch (CDBException* e) {
				// So what..! Go on!;
			}
			Sleep(g_prefs.m_QueryBackoffTime);
		}

	}
	return nRetries;
}

BOOL CDatabaseManager::DeleteProductionPagesEx(CString sOrderNumber, CString &sErrorMessage)
{
	sErrorMessage.Format("");

	CString sSQL;
	sSQL.Format("{CALL spImportCenterCleanProduction2 ('%s')}", sOrderNumber);

	BOOL bSuccess = FALSE;
	int nRetries = g_prefs.m_nQueryRetries;

	while (!bSuccess && nRetries-- > 0) {

		try {
			if (InitDB(sErrorMessage) == FALSE) {
				Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
				continue;
			}

			m_pDB->ExecuteSQL(sSQL);
			bSuccess = TRUE;
		}
		catch (CDBException* e) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
			LogprintfDB("Error Try %d - : %s   (%s)", nRetries, e->m_strError, sSQL);
			e->Delete();
			try {
				m_DBopen = FALSE;
				m_pDB->Close();
			}
			catch (CDBException* e) {
				// So what..! Go on!;
			}
			Sleep(g_prefs.m_QueryBackoffTime);
		}
	}

	return bSuccess;
}

BOOL CDatabaseManager::DeleteProductionPages(int nProductionID, BOOL bRecyclePages, CString &sErrorMessage)
{
	sErrorMessage.Format("");

	CString sSQL, sInputID, s1, s2, s3, s4, s5, sFile;
	CString sDestFile;
	CString sFileName;

	sSQL.Format("SELECT P.MasterCopySeparationSet,P.ColorID,P.LocationID,P.InputID,P.Filename FROM PageTable AS P WITH (NOLOCK) WHERE P.UniquePage=1 AND P.ProductionID=%d AND P.Dirty=1", nProductionID);

	BOOL bSuccess = FALSE;
	int nRetries = g_prefs.m_nQueryRetries;

	if (g_prefs.m_debug == 3)
		g_util.Logprintf("DeleteProductionPages() called:  %s\n", sSQL);

	while (!bSuccess && nRetries-- > 0) {
		if (InitDB(sErrorMessage) == FALSE) {
			Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
			continue;
		}
		CRecordset Rs(m_pDB);

		try {
			if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
				sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)sSQL);
				Sleep(g_prefs.m_QueryBackoffTime);
				continue;
			}

			while (!Rs.IsEOF()) {
				Rs.GetFieldValue((short)0, s1);
				Rs.GetFieldValue((short)1, s2);
				Rs.GetFieldValue((short)2, s3);
				Rs.GetFieldValue((short)3, sInputID);
				Rs.GetFieldValue((short)4, sFileName);

				sFile.Format("%s\\%s.%s", g_prefs.m_storagePath, s1, g_prefs.GetColorName(atoi(s2)));

				if (bRecyclePages) {
					sDestFile.Format("%s\\%s\\%s", g_prefs.m_errorfolder, sInputID, sFileName);
					CopyFile(sFile, sDestFile, FALSE);
				}
				/*				::DeleteFile(sFile);

								sFile.Format("%s\\%s.jpg",g_prefs.m_previewPath, s1);
								::DeleteFile(sFile);

								sFile.Format("%s\\%s_%s.jpg",g_prefs.m_previewPath, s1, g_prefs.GetColorName(atoi(s2)));
								::DeleteFile(sFile);

								sFile.Format("%s\\%s.jpg",g_prefs.m_thumbnailPath, s1);
								::DeleteFile(sFile);

								g_prefs.GetLocationRemoteFolder(atoi(s3));
								sFile.Format("%s\\%s.%s",g_prefs.GetLocationRemoteFolder(atoi(s3)), s1, g_prefs.GetColorName(atoi(s2)));
								::DeleteFile(sFile);
				*/
				Rs.MoveNext();
			}
			Rs.Close();
		}
		catch (CDBException* e) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
			LogprintfDB("Error Try %d - : %s   (%s)", nRetries, (LPCSTR)e->m_strError, (LPCSTR)sSQL);
			e->Delete();
			try {
				m_DBopen = FALSE;
				m_pDB->Close();
			}
			catch (CDBException* e) {
				// So what..! Go on!;
			}
			Sleep(g_prefs.m_QueryBackoffTime);
			continue;
		}

		sSQL.Format("{CALL spImportCenterCleanProduction (%d)}", nProductionID);
		try {
			m_pDB->ExecuteSQL(sSQL);
			bSuccess = TRUE;
			break;
		}
		catch (CDBException* e) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
			LogprintfDB("Error Try %d - : %s   (%s)", nRetries, e->m_strError, sSQL);
			e->Delete();
			try {
				m_DBopen = FALSE;
				m_pDB->Close();
			}
			catch (CDBException* e) {
				// So what..! Go on!;
			}
			Sleep(g_prefs.m_QueryBackoffTime);
		}
	}

	return bSuccess;
}



BOOL CDatabaseManager::RecycleProductionPages(int nProductionID, CString &sErrorMessage)
{
	sErrorMessage.Format("");

	CString sSQL, sInputID, s1, s2, s3, s4, s5, sFile;
	CString sDestFile;
	CString sFileName;

	sSQL.Format("SELECT P.MasterCopySeparationSet,P.ColorID,P.LocationID,P.InputID,P.Filename FROM PageTable AS P WITH (NOLOCK) WHERE P.UniquePage=1 AND P.ProductionID=%d AND P.Dirty=1", nProductionID);

	BOOL bSuccess = FALSE;
	int nRetries = g_prefs.m_nQueryRetries;


	while (!bSuccess && nRetries-- > 0) {
		if (InitDB(sErrorMessage) == FALSE) {
			Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
			continue;
		}
		CRecordset Rs(m_pDB);

		try {
			if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
				sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)sSQL);
				Sleep(g_prefs.m_QueryBackoffTime);
				continue;
			}

			while (!Rs.IsEOF()) {
				Rs.GetFieldValue((short)0, s1);
				Rs.GetFieldValue((short)1, s2);
				Rs.GetFieldValue((short)2, s3);
				Rs.GetFieldValue((short)3, sInputID);
				Rs.GetFieldValue((short)4, sFileName);

				sFile.Format("%s\\%s.%s", g_prefs.m_storagePath, s1, g_prefs.GetColorName(atoi(s2)));

				sDestFile.Format("%s\\%s\\%s", g_prefs.m_errorfolder, sInputID, sFileName);
				CopyFile(sFile, sDestFile, FALSE);

				Rs.MoveNext();
			}
			Rs.Close();
		}
		catch (CDBException* e) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
			LogprintfDB("Error Try %d - : %s   (%s)", nRetries, e->m_strError, sSQL);
			e->Delete();
			try {
				m_DBopen = FALSE;
				m_pDB->Close();
			}
			catch (CDBException* e) {
				// So what..! Go on!;
			}
			Sleep(g_prefs.m_QueryBackoffTime);
			continue;
		}
	}

	return bSuccess;
}



BOOL CDatabaseManager::PostImport(int nPressRunID, CString &sErrorMessage)
{
	sErrorMessage.Format("");

	CString sSQL;
	sSQL.Format("{CALL spImportCenterPressRunCustom (%d)}", nPressRunID);

	BOOL bSuccess = FALSE;
	int nRetries = g_prefs.m_nQueryRetries;

	if (g_prefs.m_debug >= 3)
		g_util.Logprintf("\nQuery for post phase %s\n", sSQL);

	while (!bSuccess && nRetries-- > 0) {

		try {
			if (InitDB(sErrorMessage) == FALSE) {
				Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
				continue;
			}

			m_pDB->ExecuteSQL(sSQL);
			bSuccess = TRUE;
		}
		catch (CDBException* e) {
			CString s = e->m_strError;
			if (s.GetLength() >= MAX_ERRMSG)
				s = s.Left(MAX_ERRMSG - 1);
			sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)s);
			LogprintfDB("Error Try %d - : %s   (%s)", nRetries, e->m_strError, sSQL);
			e->Delete();
			try {
				m_DBopen = FALSE;
				m_pDB->Close();
			}
			catch (CDBException* e) {
				// So what..! Go on!;
			}
			Sleep(g_prefs.m_QueryBackoffTime);
		}
	}

	return bSuccess;
}


// DEEP SEARCH ROUTINE BMA...!

BOOL CDatabaseManager::PostImport2Version2(int nPressRunID, CString &sErrorMessage)
{
	if (g_prefs.m_debug >= 3)
		g_util.Logprintf("\nStarting deep search - PostImport2Version2\n");

	int nProductionID = 0;
	int nPublicationID = 0;
	int nSectionID = 0;
	CTime tPubDate = CTime(1975, 1, 1);
	CString sPlanPageNameTest = "";
	CDBVariant DBvar;

	sErrorMessage.Format("");

	CString sSQL, s;

	sSQL.Format("SELECT TOP 1 ProductionID,PublicationID,SectionID,PubDate FROM PageTable WITH (NOLOCK) WHERE PressRunID=%d ORDER BY SectionID", nPressRunID);


	if (g_prefs.m_debug >= 3)
		g_util.Logprintf("\nPostImport2Version2() phase 1 - %s\n", sSQL);

	BOOL bSuccess = FALSE;
	int nRetries = g_prefs.m_nQueryRetries;

	// STEP 1 - Get usual parameters...

	while (!bSuccess && nRetries-- > 0) {
		if (InitDB(sErrorMessage) == FALSE) {
			Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
			continue;
		}
		CRecordset Rs(m_pDB);

		try {
			if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
				sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)sSQL);
				Sleep(g_prefs.m_QueryBackoffTime);
				continue;
			}

			if (!Rs.IsEOF()) {
				Rs.GetFieldValue((short)0, s);
				nProductionID = atoi(s);
				Rs.GetFieldValue((short)1, s);
				nPublicationID = atoi(s);
				Rs.GetFieldValue((short)2, s);
				nSectionID = atoi(s);

				Rs.GetFieldValue((short)3, DBvar, SQL_C_TIMESTAMP);
				try {
					CTime t(DBvar.m_pdate->year, DBvar.m_pdate->month, DBvar.m_pdate->day, 0, 0, 0);
					tPubDate = t;
				}
				catch (CException* e) {

				}

			}
			Rs.Close();
		}
		catch (CDBException* e) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
			LogprintfDB("Error Try %d - : %s   (%s)", nRetries, e->m_strError, sSQL);
			e->Delete();
			try {
				m_DBopen = FALSE;
				m_pDB->Close();
			}
			catch (CDBException* e) {
				// So what..! Go on!;
			}
			Sleep(g_prefs.m_QueryBackoffTime);
			continue;
		}
	}

	if (nProductionID == 0)
		return TRUE;

	// STEP 2:

	sSQL.Format("SELECT P1.PlanPageName FROM PageTable P1 WITH (NOLOCK) WHERE P1.PressRunID=%d AND P1.PlanPageName IN (SELECT P2.PlanPageName FROM PageTable P2 WITH (NOLOCK) WHERE P2.UniquePage=1 AND P2.ProductionID<>%d AND P2.PubDate='%s' AND (P2.PublicationID<>%d OR P2.SectionID<>%d))",
		nPressRunID, nProductionID, TranslateDate(tPubDate), nPublicationID, nSectionID);

	if (g_prefs.m_debug >= 3)
		g_util.Logprintf("\nPostImport2Version2() phase 2 - %s\n", sSQL);

	bSuccess = FALSE;
	nRetries = g_prefs.m_nQueryRetries;

	while (!bSuccess && nRetries-- > 0) {
		if (InitDB(sErrorMessage) == FALSE) {
			Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
			continue;
		}
		CRecordset Rs(m_pDB);

		try {
			if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
				sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)sSQL);
				Sleep(g_prefs.m_QueryBackoffTime);
				continue;
			}

			if (!Rs.IsEOF()) {
				Rs.GetFieldValue((short)0, sPlanPageNameTest);


			}
			Rs.Close();
		}
		catch (CDBException* e) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
			LogprintfDB("Error Try %d - : %s   (%s)", nRetries, e->m_strError, sSQL);
			e->Delete();
			try {
				m_DBopen = FALSE;
				m_pDB->Close();
			}
			catch (CDBException* e) {
				// So what..! Go on!;
			}
			Sleep(g_prefs.m_QueryBackoffTime);
			continue;
		}
	}

	if (sPlanPageNameTest == "")
		return bSuccess;

	// STEP 3 loop the loop....

	PLANPAGETABLELIST PlanPageNameList;
	PlanPageNameList.RemoveAll();

	sSQL.Format("SELECT DISTINCT P1.PlanPageName, P1.CopySeparationSet, P1.PageName, P1.SectionID,P1.Status, P1.ProofStatus,P1.Approved, P1.InputTime, P1.Version,P1.InputID,P1.FileServer,P1.FileName,P1.InputProcessID,P1.EmailStatus,P1.Active,P1.PageType,P1.ColorID, P1.UniquePage FROM PageTable P1 WITH (NOLOCK) WHERE P1.UniquePage=1 AND P1.PubDate='%s' AND P1.ProductionID<>%d AND (P1.PublicationID<>%d OR P1.SectionID<>%d) AND P1.PlanPageName IN (SELECT P2.PlanPageName FROM PageTable P2 WITH (NOLOCK) WHERE P2.PressRunID=%d)",
		TranslateDate(tPubDate), nProductionID, nPublicationID, nSectionID, nPressRunID);

	if (g_prefs.m_debug >= 3)
		g_util.Logprintf("\nPostImport2Version2() phase 3 - %s\n", sSQL);

	bSuccess = FALSE;
	nRetries = g_prefs.m_nQueryRetries;

	while (!bSuccess && nRetries-- > 0) {
		if (InitDB(sErrorMessage) == FALSE) {
			Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
			continue;
		}
		CRecordset Rs(m_pDB);

		try {
			if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
				sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)sSQL);
				Sleep(g_prefs.m_QueryBackoffTime);
				continue;
			}

			while (!Rs.IsEOF()) {
				PLANPAGETABLESTRUCT *pItem = new PLANPAGETABLESTRUCT();
				int idx = 0;

				Rs.GetFieldValue((short)idx++, s);
				strcpy(pItem->PlanPageName, s);

				Rs.GetFieldValue((short)idx++, s);
				pItem->MasterCopySeparationSet = atoi(s);

				Rs.GetFieldValue((short)idx++, s);
				strcpy(pItem->PageName, s);

				Rs.GetFieldValue((short)idx++, s);
				pItem->SectionID = atoi(s);

				Rs.GetFieldValue((short)idx++, s);
				pItem->Status = atoi(s);

				Rs.GetFieldValue((short)idx++, s);
				pItem->ProofStatus = atoi(s);

				Rs.GetFieldValue((short)idx++, s);
				pItem->Approved = atoi(s);

				Rs.GetFieldValue((short)idx++, DBvar, SQL_C_TIMESTAMP);
				try {
					CTime t(DBvar.m_pdate->year, DBvar.m_pdate->month, DBvar.m_pdate->day, 0, 0, 0);
					pItem->InputTime = t;
				}
				catch (CException* e) {
					pItem->InputTime = CTime(1975, 1, 1);
				}

				Rs.GetFieldValue((short)idx++, s);
				pItem->Version = atoi(s);

				Rs.GetFieldValue((short)idx++, s);
				pItem->InputID = atoi(s);

				Rs.GetFieldValue((short)idx++, s);
				strcpy(pItem->FileServer, s);

				Rs.GetFieldValue((short)idx++, s);
				strcpy(pItem->FileName, s);

				Rs.GetFieldValue((short)idx++, s);
				pItem->InputProcessID = atoi(s);

				Rs.GetFieldValue((short)idx++, s);
				pItem->EmailStatus = atoi(s);

				Rs.GetFieldValue((short)idx++, s);
				pItem->Active = atoi(s);

				Rs.GetFieldValue((short)idx++, s);
				pItem->PageType = atoi(s);

				Rs.GetFieldValue((short)idx++, s);
				pItem->ColorID = atoi(s);

				Rs.GetFieldValue((short)idx++, s);
				pItem->UniquePage = atoi(s);

				PlanPageNameList.Add(*pItem);

				Rs.MoveNext();
			}
			Rs.Close();
		}
		catch (CDBException* e) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
			LogprintfDB("Error Try %d - : %s   (%s)", nRetries, e->m_strError, sSQL);
			e->Delete();
			try {
				m_DBopen = FALSE;
				m_pDB->Close();
			}
			catch (CDBException* e) {
				// So what..! Go on!;
			}
			Sleep(g_prefs.m_QueryBackoffTime);
			continue;
		}
	}

	for (int i = 0; i < PlanPageNameList.GetCount(); i++) {
		PLANPAGETABLESTRUCT *pItem = &PlanPageNameList[i];

		// Re-image if required
		if (pItem->Status > 30)
			pItem->Status = 30;

		// Force page on this pressrun
		if (pItem->UniquePage == 1)
			pItem->UniquePage = 2;

		sSQL.Format("UPDATE PageTable SET MasterCopySeparationSet=%d,Status=%d,ProofStatus=%d,Approved=%d,InputTime='%s',Version=%d,InputID=%d,FileServer='%s',Filename='%s',InputProcessID=%d, EmailStatus=%d,Active=%d,PageType=%d,UniquePage=%d WHERE PlanPageName='%s' AND PressRunID=%d AND ColorID=%d",
			pItem->MasterCopySeparationSet,
			pItem->Status,
			pItem->ProofStatus,
			pItem->Approved,
			TranslateDate(pItem->InputTime),
			pItem->Version,
			pItem->InputID,
			pItem->FileServer,
			pItem->FileName,
			pItem->InputProcessID,
			pItem->EmailStatus,
			pItem->Active,
			pItem->PageType,
			pItem->UniquePage,

			pItem->PlanPageName,
			nPressRunID,
			pItem->ColorID);

		bSuccess = FALSE;
		nRetries = g_prefs.m_nQueryRetries;


		if (g_prefs.m_debug >= 3 && i == 0)
			g_util.Logprintf("\nPostImport2Version2() phase 4 - %s\n", sSQL);


		while (!bSuccess && nRetries-- > 0) {

			try {
				if (InitDB(sErrorMessage) == FALSE) {
					Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
					continue;
				}

				m_pDB->ExecuteSQL(sSQL);
				bSuccess = TRUE;
			}
			catch (CDBException* e) {
				CString s = e->m_strError;
				if (s.GetLength() >= MAX_ERRMSG)
					s = s.Left(MAX_ERRMSG - 1);
				sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)s);
				LogprintfDB("Error Try %d - : %s   (%s)", nRetries, e->m_strError, sSQL);
				e->Delete();
				try {
					m_DBopen = FALSE;
					m_pDB->Close();
				}
				catch (CDBException* e) {
					// So what..! Go on!;
				}
				Sleep(g_prefs.m_QueryBackoffTime);
			}
		}
	}


	return bSuccess;
}


BOOL CDatabaseManager::PostImport2(int nPressRunID, CString &sErrorMessage)
{
	sErrorMessage.Format("");

	if (g_prefs.m_useInCodeDeepSearch)
		return PostImport2Version2(nPressRunID, sErrorMessage);

	CString sSQL;
	sSQL.Format("{CALL spImportCenterPressRunCustom2 (%d)}", nPressRunID);

	BOOL bSuccess = FALSE;
	int nRetries = g_prefs.m_nQueryRetries;

	if (g_prefs.m_debug >= 3)
		g_util.Logprintf("\nQuery for post 2 phase %s\n", sSQL);

	while (!bSuccess && nRetries-- > 0) {

		try {
			if (InitDB(sErrorMessage) == FALSE) {
				Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
				continue;
			}

			m_pDB->ExecuteSQL(sSQL);
			bSuccess = TRUE;
		}
		catch (CDBException* e) {
			CString s = e->m_strError;
			if (s.GetLength() >= MAX_ERRMSG)
				s = s.Left(MAX_ERRMSG - 1);
			sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)s);
			LogprintfDB("Error Try %d - : %s   (%s)", nRetries, e->m_strError, sSQL);
			e->Delete();
			try {
				m_DBopen = FALSE;
				m_pDB->Close();
			}
			catch (CDBException* e) {
				// So what..! Go on!;
			}
			Sleep(g_prefs.m_QueryBackoffTime);
		}
	}

	return bSuccess;
}


BOOL CDatabaseManager::RestoreEditionDirtyFlag(int nPublicationID, CTime tPubDate, int nEditionID, int nPressID, CString &sErrorMessage)
{
	sErrorMessage.Format("");

	CString sSQL;
	if (nPressID == 0)
		sSQL.Format("UPDATE PageTable SET Dirty=0 WHERE PublicationID=%d AND PubDate='%s' AND EditionID=%d", nPublicationID, TranslateDate(tPubDate), nEditionID);
	else
		sSQL.Format("UPDATE PageTable SET Dirty=0 WHERE PublicationID=%d AND PubDate='%s' AND EditionID=%d AND PressID=%d", nPublicationID, TranslateDate(tPubDate), nEditionID, nPressID);
	BOOL bSuccess = FALSE;
	int nRetries = g_prefs.m_nQueryRetries;

	if (g_prefs.m_debug >= 3)
		g_util.Logprintf("\nQuery for RestoreEditionDirtyFlag() %s\n", sSQL);

	while (!bSuccess && nRetries-- > 0) {

		try {
			if (InitDB(sErrorMessage) == FALSE) {
				Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
				continue;
			}

			m_pDB->ExecuteSQL(sSQL);
			bSuccess = TRUE;
		}
		catch (CDBException* e) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
			LogprintfDB("Error Try %d - : %s   (%s)", nRetries, e->m_strError, sSQL);
			e->Delete();
			try {
				m_DBopen = FALSE;
				m_pDB->Close();
			}
			catch (CDBException* e) {
				// So what..! Go on!;
			}
			Sleep(g_prefs.m_QueryBackoffTime);
		}
	}

	return bSuccess;
}


BOOL CDatabaseManager::SetPressRunPageFormat(int nPressRunID, int nPageFormatID, CString &sErrorMessage)
{
	sErrorMessage.Format("");

	CString sSQL;
	sSQL.Format("UPDATE PressRunID SET MiscInt1=%d WHERE PressRunID=%d", nPageFormatID, nPressRunID);

	BOOL bSuccess = FALSE;
	int nRetries = g_prefs.m_nQueryRetries;

	if (g_prefs.m_debug >= 3)
		g_util.Logprintf("\nQuery for SetPressRunPageFormat() %s\n", sSQL);

	while (!bSuccess && nRetries-- > 0) {

		try {
			if (InitDB(sErrorMessage) == FALSE) {
				Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
				continue;
			}

			m_pDB->ExecuteSQL(sSQL);
			bSuccess = TRUE;
		}
		catch (CDBException* e) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
			LogprintfDB("Error Try %d - : %s   (%s)", nRetries, e->m_strError, sSQL);
			e->Delete();
			try {
				m_DBopen = FALSE;
				m_pDB->Close();
			}
			catch (CDBException* e) {
				// So what..! Go on!;
			}
			Sleep(g_prefs.m_QueryBackoffTime);
		}
	}

	return bSuccess;
}


BOOL CDatabaseManager::AutoApplyProduction(int nProductionID, BOOL bInsertedSections, int nPaginationMode, int nSeparateRuns, CString &sErrorMessage)
{
	sErrorMessage.Format("");

	CString sSQL;
	sSQL.Format("{CALL spApplyProduction (%d,%d,0,0,'',0,0,%d,%d)}", nProductionID, bInsertedSections, nPaginationMode, nSeparateRuns);


	BOOL bSuccess = FALSE;
	int nRetries = g_prefs.m_nQueryRetries;

	g_util.Logprintf("INFO: %s", sSQL);

	while (!bSuccess && nRetries-- > 0) {

		try {
			if (InitDB(sErrorMessage) == FALSE) {
				Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
				continue;
			}

			m_pDB->ExecuteSQL(sSQL);
			bSuccess = TRUE;
		}
		catch (CDBException* e) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
			LogprintfDB("Error Try %d - : %s   (%s)", nRetries, e->m_strError, sSQL);
			e->Delete();
			try {
				m_DBopen = FALSE;
				m_pDB->Close();
			}
			catch (CDBException* e) {
				// So what..! Go on!;
			}
			Sleep(g_prefs.m_QueryBackoffTime);
		}
	}

	return bSuccess;
}


BOOL CDatabaseManager::PostProduction(int nProductionID, CString &sErrorMessage)
{
	sErrorMessage.Format("");

	CString sSQL;
	sSQL.Format("{CALL spImportCenterProductionCustom (%d)}", nProductionID);

	BOOL bSuccess = FALSE;
	int nRetries = g_prefs.m_nQueryRetries;

	g_util.Logprintf("INFO: %s", sSQL);

	while (!bSuccess && nRetries-- > 0) {

		try {
			if (InitDB(sErrorMessage) == FALSE) {
				Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
				continue;
			}

			m_pDB->ExecuteSQL(sSQL);
			bSuccess = TRUE;
		}
		catch (CDBException* e) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
			LogprintfDB("Error Try %d - : %s   (%s)", nRetries, e->m_strError, sSQL);
			e->Delete();
			try {
				m_DBopen = FALSE;
				m_pDB->Close();
			}
			catch (CDBException* e) {
				// So what..! Go on!;
			}
			Sleep(g_prefs.m_QueryBackoffTime);
		}
	}

	return bSuccess;
}

BOOL CDatabaseManager::PostProduction2(int nProductionID, CString &sErrorMessage)
{
	sErrorMessage.Format("");

	CString sSQL;
	sSQL.Format("{CALL spImportCenterProductionCustom2 (%d)}", nProductionID);

	BOOL bSuccess = FALSE;
	int nRetries = g_prefs.m_nQueryRetries;

	g_util.Logprintf("INFO: %s", sSQL);

	while (!bSuccess && nRetries-- > 0) {

		try {
			if (InitDB(sErrorMessage) == FALSE) {
				Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
				continue;
			}

			m_pDB->ExecuteSQL(sSQL);
			bSuccess = TRUE;
		}
		catch (CDBException* e) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
			LogprintfDB("Error Try %d - : %s   (%s)", nRetries, e->m_strError, sSQL);
			e->Delete();
			try {
				m_DBopen = FALSE;
				m_pDB->Close();
			}
			catch (CDBException* e) {
				// So what..! Go on!;
			}
			Sleep(g_prefs.m_QueryBackoffTime);
		}
	}

	return bSuccess;
}


BOOL CDatabaseManager::UnApplyEdition(int nProductionID, int nEditionID, CString &sErrorMessage)
{
	sErrorMessage.Format("");

	CString sSQL;
	sSQL.Format("UPDATE ProductionNames SET PlanType=0 WHERE ProductionID=%d AND EXISTS (SELECT DISTINCT EditionID FROM PageTable AS P WITH (NOLOCK) WHERE P.ProductionID=%d AND P.EditionID=%d)", nProductionID, nProductionID, nEditionID);

	BOOL bSuccess = FALSE;
	int nRetries = g_prefs.m_nQueryRetries;

	g_util.Logprintf("INFO: %s", sSQL);

	while (!bSuccess && nRetries-- > 0) {

		try {
			if (InitDB(sErrorMessage) == FALSE) {
				Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
				continue;
			}

			m_pDB->ExecuteSQL(sSQL);
			bSuccess = TRUE;
		}
		catch (CDBException* e) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
			LogprintfDB("Error Try %d - : %s   (%s)", nRetries, e->m_strError, sSQL);
			e->Delete();
			try {
				m_DBopen = FALSE;
				m_pDB->Close();
			}
			catch (CDBException* e) {
				// So what..! Go on!;
			}
			Sleep(g_prefs.m_QueryBackoffTime);
		}
	}

	if (bSuccess == FALSE)
		return FALSE;

	sSQL.Format("UPDATE PressRunID SET PlanType=0 WHERE PressRunID.PressRunID IN (SELECT DISTINCT PressRunID FROM PageTable AS P WITH (NOLOCK) WHERE P.ProductionID=%d AND P.EditionID=%d)", nProductionID, nEditionID);

	bSuccess = FALSE;
	nRetries = g_prefs.m_nQueryRetries;

	g_util.Logprintf("INFO: %s", sSQL);

	while (!bSuccess && nRetries-- > 0) {

		try {
			if (InitDB(sErrorMessage) == FALSE) {
				Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
				continue;
			}

			m_pDB->ExecuteSQL(sSQL);
			bSuccess = TRUE;
		}
		catch (CDBException* e) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
			LogprintfDB("Error Try %d - : %s   (%s)", nRetries, e->m_strError, sSQL);
			e->Delete();
			try {
				m_DBopen = FALSE;
				m_pDB->Close();
			}
			catch (CDBException* e) {
				// So what..! Go on!;
			}
			Sleep(g_prefs.m_QueryBackoffTime);
		}
	}



	return bSuccess;
}

BOOL CDatabaseManager::ResetAllPollLocks(CString &sErrorMessage)
{
	sErrorMessage.Format("");
	CString sSQL = "UPDATE PageTable SET HardProofStatus=0";

	BOOL bSuccess = FALSE;
	int nRetries = g_prefs.m_nQueryRetries;

	while (!bSuccess && nRetries-- > 0) {

		try {
			if (InitDB(sErrorMessage) == FALSE) {
				Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
				continue;
			}

			m_pDB->ExecuteSQL(sSQL);
			bSuccess = TRUE;
		}
		catch (CDBException* e) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
			LogprintfDB("Error Try %d - : %s   (%s)", nRetries, e->m_strError, sSQL);
			e->Delete();
			try {
				m_DBopen = FALSE;
				m_pDB->Close();
			}
			catch (CDBException* e) {
				// So what..! Go on!;
			}
			Sleep(g_prefs.m_QueryBackoffTime);
		}
	}

	return bSuccess;
}

BOOL CDatabaseManager::DeleteAllDirtyPages(int nProductionID, CString &sErrorMessage)
{
	sErrorMessage.Format("");

	CString sSQL;

	sSQL.Format("{CALL spImportCenterCleanDirtyPages (%d)}", nProductionID);


	BOOL bSuccess = FALSE;
	int nRetries = g_prefs.m_nQueryRetries;

	if (g_prefs.m_debug >= 3)
		g_util.Logprintf("DeleteAllDirtyPages() called: %s\n", sSQL);
	while (!bSuccess && nRetries-- > 0) {

		try {
			if (InitDB(sErrorMessage) == FALSE) {
				Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
				continue;
			}

			m_pDB->ExecuteSQL(sSQL);
			bSuccess = TRUE;
		}
		catch (CDBException* e) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
			LogprintfDB("Error Try %d - : %s   (%s)", nRetries, e->m_strError, sSQL);
			e->Delete();
			try {
				m_DBopen = FALSE;
				m_pDB->Close();
			}
			catch (CDBException* e) {
				// So what..! Go on!;
			}
			Sleep(g_prefs.m_QueryBackoffTime);
		}
	}

	return bSuccess;
}

int CDatabaseManager::CreateProductionID(int &nPressID, int nPublicationID, CTime tPubDate, int nNumberOfPages, int nNumberOfSections, int nNumberOfEditions, int nPlannedApproval, int nPlannedHold, int &nExistingProduction, int nWeekNumber, int nPlanType, int nRipPressID, CString &sErrorMessage)
{
	int nProductionID = -1;
	nExistingProduction = FALSE;

	CString sSQL;

	sSQL.Format("{CALL spImportCenterAddProduction (%d,'%s',%d,%d,%d,%d,%d,%d,%d,%d,0,%d,%d)}", nPublicationID, TranslateDate(tPubDate), nNumberOfPages, nNumberOfSections, nNumberOfEditions, nPlannedApproval, nPlannedHold, nPressID, g_prefs.m_keeppress, nWeekNumber, nPlanType, nRipPressID);

	if (g_prefs.m_debug >= 3)
		g_util.Logprintf("\nQuery for production name creation %s\n", sSQL);

	sErrorMessage.Format("");


	BOOL bSuccess = FALSE;
	int nRetries = g_prefs.m_nQueryRetries;


	while (!bSuccess && nRetries-- > 0) {

		if (InitDB(sErrorMessage) == FALSE) {
			Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
			continue;
		}
		CRecordset Rs(m_pDB);


		try {
			if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
				sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)sSQL);
				return FALSE;
			}

			if (!Rs.IsEOF()) {
				CString s;
				Rs.GetFieldValue((short)0, s);
				nProductionID = atoi(s);
				Rs.GetFieldValue((short)1, s);
				nExistingProduction = atoi(s);

				Rs.GetFieldValue((short)2, s);
				nPressID = atoi(s);

			}
			Rs.Close();
			bSuccess = TRUE;
			break;
		}
		catch (CDBException* e) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
			LogprintfDB("Error Try %d - : %s   (%s)", nRetries, e->m_strError, sSQL);
			e->Delete();
			try {
				m_DBopen = FALSE;
				m_pDB->Close();
			}
			catch (CDBException* e) {
				// So what..! Go on!;
			}
			::Sleep(g_prefs.m_QueryBackoffTime);
		}
	}
	return bSuccess ? nProductionID : -1;
}



BOOL CDatabaseManager::QueryProduction(int nPublicationID, CTime tPubDate, int nIssueID, int nEditionID, int nSectionID, int *nPressID,
	int *nNumberOfPage, int *nNumberOfPagesPolled, int *nNumberOfPagesImaged, int *nTemplateID, int nWeekNumber, BOOL bKeepExistingSections, CString &sErrorMessage)
{
	*nTemplateID = 1;
	*nNumberOfPage = 0;
	*nNumberOfPagesPolled = 0;
	*nNumberOfPagesImaged = 0;

	CString sSQL;

	sSQL.Format("{CALL spImportCenterQueryProduction (%d,'%s',%d,%d,%d,%d,%d,%d,%d)}",
		nPublicationID,
		TranslateDate(tPubDate),
		nIssueID,
		nEditionID,
		nSectionID,
		*nPressID,
		g_prefs.m_keeppress,
		nWeekNumber,
		bKeepExistingSections/*g_prefs.m_keepexistingsections*/);

	sErrorMessage.Format("");

	BOOL bSuccess = FALSE;
	int nRetries = g_prefs.m_nQueryRetries;


	while (!bSuccess && nRetries-- > 0) {

		if (InitDB(sErrorMessage) == FALSE) {
			Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
			continue;
		}
		CRecordset Rs(m_pDB);

		try {
			if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
				sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)sSQL);
				return FALSE;
			}

			if (!Rs.IsEOF()) {
				CString s;

				Rs.GetFieldValue((short)0, s);
				*nNumberOfPage = atoi(s);
				Rs.GetFieldValue((short)1, s);
				*nTemplateID = atoi(s);

				Rs.GetFieldValue((short)2, s);
				*nPressID = atoi(s);

				Rs.GetFieldValue((short)3, s);
				*nNumberOfPagesPolled = atoi(s);

				Rs.GetFieldValue((short)4, s);
				*nNumberOfPagesImaged = atoi(s);
			}
			Rs.Close();
			bSuccess = TRUE;
			break;
		}
		catch (CDBException* e) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
			LogprintfDB("Error Try %d - : %s   (%s)", nRetries, e->m_strError, sSQL);
			e->Delete();
			try {
				m_DBopen = FALSE;
				m_pDB->Close();
			}
			catch (CDBException* e) {
				// So what..! Go on!;
			}
			Sleep(g_prefs.m_QueryBackoffTime);
		}
	}
	return bSuccess;
}

BOOL CDatabaseManager::SetTimedEdition(int nPressRunID, int nFromTimedEditionID, int nToTimedEditionID, int nEditionOrderNumber, int nZoneMasterEditionID, CString &sErrorMessage)
{
	BOOL bSuccess = FALSE;
	int nRetries = g_prefs.m_nQueryRetries;

	CString sSQL;

	sSQL.Format("UPDATE PressRunID SET TimedEditionFrom=%d, TimedEditionTo=%d,FromZone=%d,ToZone=%d WHERE PressRunID=%d",
		nFromTimedEditionID, nToTimedEditionID, nEditionOrderNumber, nZoneMasterEditionID, nPressRunID);

	if (g_prefs.m_debug >= 3)
		g_util.Logprintf("\nQuery for second phase %s\n", sSQL);

	while (!bSuccess && nRetries-- > 0) {

		if (InitDB(sErrorMessage) == FALSE) {
			Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
			continue;
		}

		try {
			m_pDB->ExecuteSQL(sSQL);
			bSuccess = TRUE;
			break;
		}
		catch (CDBException* e) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
			LogprintfDB("Error Try %d - : %s   (%s)", nRetries, e->m_strError, sSQL);
			e->Delete();
			try {
				m_DBopen = FALSE;
				m_pDB->Close();
			}
			catch (CDBException* e) {
				// So what..! Go on!;
			}
			Sleep(g_prefs.m_QueryBackoffTime);
		}
	}
	return bSuccess;
}

BOOL CDatabaseManager::SetAutoRelease(int nPressRunID, BOOL bAutoRelease, BOOL bRejectSubEditionPages, CString &sErrorMessage)
{
	BOOL bSuccess = FALSE;
	int nRetries = g_prefs.m_nQueryRetries;

	CString sSQL;

	int n = bAutoRelease + 2 * bRejectSubEditionPages;
	sSQL.Format("UPDATE PressRunID SET MiscInt2=%d WHERE PressRunID=%d", n, nPressRunID);

	if (g_prefs.m_debug >= 3)
		g_util.Logprintf("\nQuery for second phase %s\n", sSQL);

	while (!bSuccess && nRetries-- > 0) {

		if (InitDB(sErrorMessage) == FALSE) {
			Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
			continue;
		}

		try {
			m_pDB->ExecuteSQL(sSQL);
			bSuccess = TRUE;
			break;
		}
		catch (CDBException* e) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
			LogprintfDB("Error Try %d - : %s   (%s)", nRetries, e->m_strError, sSQL);
			e->Delete();
			try {
				m_DBopen = FALSE;
				m_pDB->Close();
			}
			catch (CDBException* e) {
				// So what..! Go on!;
			}
			Sleep(g_prefs.m_QueryBackoffTime);
		}
	}
	return bSuccess;
}

BOOL CDatabaseManager::GetNextPressRun(int nPublicationID, CTime tPubDate, int nIssueID, int nEditionID, int nSectionID, int nPressID, int *nPressRunID, int nSequenceNumber, int nWeekNumber, CString &sErrorMessage)
{
	*nPressRunID = 1;
	CString sSQL;

	sSQL.Format("{CALL spImportCenterGetPressRun2 (%d,'%s',%d,%d,%d,%d,%d,%d,%d,%d)}",
		nPublicationID,
		TranslateDate(tPubDate),
		nIssueID,
		nEditionID,
		nSectionID,
		nPressID,
		g_prefs.m_combinesections,
		g_prefs.m_keeppress,
		nSequenceNumber,
		nWeekNumber);

	sErrorMessage.Format("");

	BOOL bSuccess = FALSE;
	int nRetries = g_prefs.m_nQueryRetries;


	while (!bSuccess && nRetries-- > 0) {

		if (InitDB(sErrorMessage) == FALSE) {
			Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
			continue;
		}
		CRecordset Rs(m_pDB);

		try {
			if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
				sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)sSQL);
				return FALSE;
			}

			if (!Rs.IsEOF()) {
				CString s;

				Rs.GetFieldValue((short)0, s);
				*nPressRunID = atoi(s);
			}
			Rs.Close();
			bSuccess = TRUE;
			break;
		}
		catch (CDBException* e) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
			LogprintfDB("Error Try %d - : %s   (%s)", nRetries, e->m_strError, sSQL);
			e->Delete();
			try {
				m_DBopen = FALSE;
				m_pDB->Close();
			}
			catch (CDBException* e) {
				// So what..! Go on!;
			}
			Sleep(g_prefs.m_QueryBackoffTime);
		}
	}
	return bSuccess;
}

BOOL CDatabaseManager::InsertSeparation(CPageTableEntry *pItem, int *nCopySeparationSet, int *nCopyFlatSeparationSet, int nNumberOfColors, BOOL bFirstPagePosition, int nCopies, int nCopiesToProduce,CString &sErrorMessage)
{
	sErrorMessage.Format("");

	if (g_prefs.m_generatedummycopies == FALSE && nCopiesToProduce == 0)
		return TRUE;

	if (g_prefs.m_generatedummycopies == FALSE && nCopiesToProduce < nCopies)
		nCopies = nCopiesToProduce;

	CString sSQL;

	CString sPlateName;
	sPlateName.Format("%d", pItem->m_sheetnumber * 2 + pItem->m_sheetside - 1);

	sSQL.Format("{CALL spImportCenterAddSeparation2 (%d, %d, %d, %d, %d, %d,%d,%d,%d, %d,%d,%d,%d, '%s','%s', %d,%d,%d,%d, %d,%d,%d,%d,%d,%d, '%s',  %d,%d,%d,%d,%d,%d, '%s','%s','%s','%s','%s',  %d,%d, '%s',  %d,%d,%d,%d, %f,  '%s',%d,%d,'%s','%s',%d,%d,%d,%d,%d,%d,  '%s', %d,  %d,%d,%d,%d,  '%s','%s','%s',  '%s',0,0,0,0,0)}",
		pItem->m_pagecountchange,
		g_prefs.m_keepcolors,
		g_prefs.m_keepapproval,
		g_prefs.m_keepunique,
		TRUE,
		bFirstPagePosition,
		*nCopyFlatSeparationSet,
		nNumberOfColors,
		nCopies,

		pItem->m_publicationID,
		pItem->m_sectionID,
		pItem->m_editionID,
		pItem->m_issueID,

		TranslateDate(pItem->m_pubdate),
		pItem->m_pagename,

		pItem->m_colorID,
		pItem->m_templateID,
		pItem->m_proofID,
		pItem->m_deviceID,

		pItem->m_version,
		pItem->m_pagination,
		pItem->m_approved,
		pItem->m_hold,
		pItem->m_active,
		pItem->m_priority,

		pItem->m_pagepositions,

		pItem->m_pagetype,
		pItem->m_pagesonplate,
		pItem->m_sheetnumber,
		pItem->m_sheetside,
		pItem->m_pressID,
		pItem->m_presssectionnumber,

		pItem->m_sortingposition,
		pItem->m_presstower,
		pItem->m_presszone,
		pItem->m_presscylinder,
		pItem->m_presshighlow,

		pItem->m_productionID,
		pItem->m_pressrunID,

		pItem->m_planpagename,

		pItem->m_issuesequencenumber,
		pItem->m_uniquepage,
		pItem->m_locationID,
		pItem->m_flatproofID,

		(float)pItem->m_creep,

		pItem->m_markgroups,

		pItem->m_pageindex,
		pItem->m_hardproofID,

		TranslateDate(pItem->m_deadline),

		pItem->m_comment,

		pItem->m_mastereditionID,
		1, //g_prefs.GetLocationID(g_prefs.m_masterlocation),
		pItem->m_colorIndex,
		nCopiesToProduce,
		pItem->m_priority,						// Output prio
		g_prefs.m_keeppress,					// @PressChange

		TranslateTimeEx(pItem->m_presstime),		// @PressTime


		pItem->m_customerID,					// @CustomerID

		pItem->m_autorelease,				    // @Miscint1 = autorelease (0/1) + extragutter*1000
		pItem->m_weekreference,					// @Miscint2 = weeknumber	
		0,										// @Miscint3 = CRC32 
		0,										// @Miscint4 = poll lock		

		sPlateName,								// @MiscString1  = PlateName (default is 2*SheetNumber+Sheetside)
		pItem->m_formID,						// @MiscString2  = InkSystemID, EAE formID
		pItem->m_plateID,						// @MiscString3  = EAE PlateID, ReferenceID
	
		"1975-01-01T00:00:00"					// @MiscDate

	);

	if (g_prefs.m_debug > 1 && *nCopySeparationSet == 0)
		g_util.Logprintf("\nQuery for first entry %s\n", sSQL);
	if (g_prefs.m_debug == 3)
		g_util.Logprintf("\nQuery for adding record %s\n", sSQL);


	BOOL bSuccess = FALSE;
	int nRetries = g_prefs.m_nQueryRetries;



	while (!bSuccess && nRetries-- > 0) {


		if (InitDB(sErrorMessage) == FALSE) {
			Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
			continue;
		}
		CRecordset Rs(m_pDB);

		try {
			if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
				Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
				continue;
			}

			if (!Rs.IsEOF()) {
				CString s;
				Rs.GetFieldValue((short)0, s);
				*nCopySeparationSet = atoi(s);
				Rs.GetFieldValue((short)1, s);
				*nCopyFlatSeparationSet = atoi(s);
				Rs.GetFieldValue((short)2, s);
				int nUpdated = atoi(s);

				if (g_prefs.m_debug == 3)
					g_util.Logprintf("\nSeparation was %s (CopySeparationSet=%d, CopyFlatSeparationSet=%d)\n", nUpdated ? "updated" : "inserted", *nCopySeparationSet, *nCopyFlatSeparationSet);
			}
			Rs.Close();
			bSuccess = TRUE;
			break;
		}
		catch (CDBException* e) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
			LogprintfDB("Error Try %d - : %s   (%s)", nRetries, e->m_strError, sSQL);
			e->Delete();
			try {
				m_DBopen = FALSE;
				m_pDB->Close();
			}
			catch (CDBException* e) {
				// So what..! Go on!;
			}
			Sleep(g_prefs.m_QueryBackoffTime);
		}
	}
	return bSuccess;
}

BOOL CDatabaseManager::InsertSeparationPhase2(CPageTableEntry *pItem, BOOL bPressSpecific, CString &sErrorMessage)
{
	sErrorMessage.Format("");

	CString sSQL;
	
		sSQL.Format("{CALL spImportCenterAddSeparationPhase2 (%d, %d, %d, %d, '%s','%s', %d, %d, %d,%d,%d,'%s',%d)}",
				pItem->m_publicationID,
				pItem->m_sectionID,
				pItem->m_editionID,
				pItem->m_issueID,

				TranslateDate(pItem->m_pubdate),
				pItem->m_pagename,

				pItem->m_colorID,

				pItem->m_pagetype,
				pItem->m_uniquepage,
				pItem->m_mastereditionID,
				pItem->m_pressID,
				pItem->m_planpagename,
				bPressSpecific);
		

	BOOL bSuccess = FALSE;
	int nRetries = g_prefs.m_nQueryRetries;

	if (g_prefs.m_debug >= 3)
		g_util.Logprintf("\nQuery for second phase %s\n", sSQL);

	while (!bSuccess && nRetries-- > 0) {

		if (InitDB(sErrorMessage) == FALSE) {
			Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
			continue;
		}

		try {
			m_pDB->ExecuteSQL(sSQL);
			bSuccess = TRUE;
			break;
		}
		catch (CDBException* e) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
			LogprintfDB("Error Try %d - : %s   (%s)", nRetries, e->m_strError, sSQL);
			e->Delete();
			try {
				m_DBopen = FALSE;
				m_pDB->Close();
			}
			catch (CDBException* e) {
				// So what..! Go on!;
			}
			Sleep(g_prefs.m_QueryBackoffTime);
		}
	}
	return bSuccess;
}

BOOL CDatabaseManager::KillDirtyflags(CString &sErrorMessage)
{
	sErrorMessage.Format("");

	CString sSQL;
	sSQL = "UPDATE PageTable SET Dirty=0";

	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	try {
		m_pDB->ExecuteSQL(sSQL);
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error : %s   (%s)", e->m_strError, sSQL);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}

	KillUnusedProductions(sErrorMessage);
	return TRUE;
}

BOOL CDatabaseManager::UpdatePressRun(int nPressRunID, CString sPlanID, CString sPlanName, CString sInkComment, CString &sErrorMessage)
{
	sErrorMessage.Format("");
	CString sSQL;
	if (sInkComment != "")
		sSQL.Format("UPDATE PressRunID SET Comment='%s', OrderNumber='%s', InkComment='%s' WHERE PressRunID=%d", sPlanID, sPlanName, sInkComment, nPressRunID);
	else
		sSQL.Format("UPDATE PressRunID SET Comment='%s', OrderNumber='%s' WHERE PressRunID=%d", sPlanID, sPlanName, nPressRunID);

	try {
		if (InitDB(sErrorMessage) == FALSE)
			return FALSE;

		m_pDB->ExecuteSQL(sSQL);
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error : %s   (%s)", e->m_strError, sSQL);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}

	return TRUE;

}

BOOL CDatabaseManager::GetProductionOrderNumber(int nProductionID, CString &sOrderNumber, CString &sErrorMessage)
{
	sOrderNumber = _T("");
	CString sSQL;
	CString s;

	sSQL.Format("SELECT TOP 1 OrderNumber FROM ProductionNames WHERE ProductionID=%d", nProductionID);

	sErrorMessage.Format("");

	BOOL bSuccess = FALSE;
	int nRetries = g_prefs.m_nQueryRetries;


	while (!bSuccess && nRetries-- > 0) {

		if (InitDB(sErrorMessage) == FALSE) {
			Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
			continue;
		}
		CRecordset Rs(m_pDB);

		try {
			if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
				sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)sSQL);
				return FALSE;
			}

			if (!Rs.IsEOF()) {


				Rs.GetFieldValue((short)0, s);
				sOrderNumber = s;
			}
			Rs.Close();
			bSuccess = TRUE;
			break;
		}
		catch (CDBException* e) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
			LogprintfDB("Error Try %d - : %s   (%s)", nRetries, e->m_strError, sSQL);
			e->Delete();
			try {
				m_DBopen = FALSE;
				m_pDB->Close();
			}
			catch (CDBException* e) {
				// So what..! Go on!;
			}
			Sleep(g_prefs.m_QueryBackoffTime);
		}
	}
	return bSuccess;
}



BOOL CDatabaseManager::GetPublicationChannels(int nPublicationID, CUIntArray &aChannels, CString &sErrorMessage)
{
	aChannels.RemoveAll();
	CString sSQL;
	CString s;

	sSQL.Format("SELECT DISTINCT ChannelID FROM PublicationChannels WHERE PublicationID=%d", nPublicationID);

	sErrorMessage.Format("");

	BOOL bSuccess = FALSE;
	int nRetries = g_prefs.m_nQueryRetries;


	while (!bSuccess && nRetries-- > 0) {

		if (InitDB(sErrorMessage) == FALSE) {
			Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
			continue;
		}
		CRecordset Rs(m_pDB);

		try {
			if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
				sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)sSQL);
				return FALSE;
			}

			while (!Rs.IsEOF()) {
				Rs.GetFieldValue((short)0, s);
				aChannels.Add(atoi(s));
				Rs.MoveNext();
			}
			Rs.Close();
			bSuccess = TRUE;
			break;
		}
		catch (CDBException* e) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
			LogprintfDB("Error Try %d - : %s   (%s)", nRetries, e->m_strError, sSQL);
			e->Delete();
			try {
				m_DBopen = FALSE;
				m_pDB->Close();
			}
			catch (CDBException* e) {
				// So what..! Go on!;
			}
			Sleep(g_prefs.m_QueryBackoffTime);
		}
	}
	return bSuccess;
}



BOOL CDatabaseManager::UpdateProductionOrderNumber(int nProductionID, CString sOrderNumber, CString &sErrorMessage)
{
	sErrorMessage.Format("");
	CString sSQL;

	sSQL.Format("UPDATE ProductionNames SET OrderNumber='%s' WHERE ProductionID=%d", sOrderNumber, nProductionID);

	try {
		if (InitDB(sErrorMessage) == FALSE)
			return FALSE;

		m_pDB->ExecuteSQL(sSQL);
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error : %s   (%s)", e->m_strError, sSQL);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}

	return TRUE;

}


BOOL CDatabaseManager::UpdateProductionType(int nProductionID, int nPlanType, CString &sErrorMessage)
{
	sErrorMessage.Format("");
	CString sSQL;
	sSQL.Format("UPDATE ProductionNames SET PlanType=%d WHERE ProductionID=%d", nPlanType, nProductionID);

	try {
		if (InitDB(sErrorMessage) == FALSE)
			return FALSE;

		m_pDB->ExecuteSQL(sSQL);
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error : %s   (%s)", e->m_strError, sSQL);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}

	return TRUE;

}


BOOL CDatabaseManager::KillUnusedProductions(CString &sErrorMessage)
{
	sErrorMessage.Format("");

	CString sSQL;
	sSQL = "DELETE ProductionNames WHERE ProductionID NOT IN (SELECT DISTINCT ProductionID FROM PageTable)";

	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	try {
		m_pDB->ExecuteSQL(sSQL);
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Delete failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error : %s   (%s)", e->m_strError, sSQL);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}
	return TRUE;
}


BOOL CDatabaseManager::KillUnusedPressRuns(CString &sErrorMessage)
{
	sErrorMessage.Format("");

	CString sSQL;
	sSQL = "DELETE PressRunID WHERE PressRunID NOT IN (SELECT DISTINCT PressRunID FROM PageTable)";

	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	try {
		m_pDB->ExecuteSQL(sSQL);
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Delete failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error : %s   (%s)", e->m_strError, sSQL);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}
	return TRUE;
}



BOOL CDatabaseManager::LoadUser(CString sUserName, CString sPassword, int *nReturnCode, BOOL *bIsAdmin, CString &sErrorMessage)
{
	CString sSQL;

	*nReturnCode = USER_UNKNOWNUSER;
	*bIsAdmin = FALSE;

	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	CRecordset Rs(m_pDB);

	sSQL.Format("SELECT U.Password, U.AccountEnabled, G.IsAdmin, G.MayConfigure FROM UserNames AS U INNER JOIN UserGroupNames AS G On G.UserGroupID=U.UserGroupID WHERE U.Username='%s'", sUserName);

	CString sThisPassword;
	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)sSQL);
			return FALSE;
		}

		if (!Rs.IsEOF()) {
			CString s;
			*nReturnCode = USER_PASSWORDWRONG;
			Rs.GetFieldValue((short)0, sThisPassword);
			Rs.GetFieldValue((short)1, s);
			if (atoi(s) == 0) {
				*nReturnCode = USER_ACCOUNTDISABLED;
			}
			else {
				Rs.GetFieldValue((short)2, s);
				BOOL l_bIsAdmin = atoi(s);
				Rs.GetFieldValue((short)3, s);
				BOOL l_bMayGonfigure = atoi(s);
				*bIsAdmin = (l_bIsAdmin || l_bMayGonfigure) ? TRUE : FALSE;
				if (sThisPassword == sPassword)
					*nReturnCode = USER_USEROK;
			}
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		e->Delete();
		Rs.Close();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}
	return TRUE;
}


BOOL CDatabaseManager::PlanLock(int bRequestedPlanLock, int *bCurrentPlanLock, TCHAR *szClientName, TCHAR *szClientTime, CString &sErrorMessage)
{
	CString sSQL;
	strcpy(szClientTime, "");
	strcpy(szClientName, "");
	*bCurrentPlanLock = -1;

	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	CRecordset Rs(m_pDB);

	CString sClientName;
	TCHAR buf[260];
	DWORD len = 260;
	::GetComputerName(buf, &len);
	sSQL.Format("{CALL spPlanningLock (%d,'IMPORTCENTER%d %s')}", bRequestedPlanLock, g_prefs.m_instancenumber, buf);
	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)sSQL);
			return FALSE;
		}

		if (!Rs.IsEOF()) {
			CString s;
			Rs.GetFieldValue((short)0, s);
			*bCurrentPlanLock = atoi(s);
			Rs.GetFieldValue((short)1, s);
			strcpy(szClientName, s);

			CDBVariant DBvar;
			Rs.GetFieldValue((short)2, DBvar, SQL_C_TIMESTAMP);
			sprintf(szClientTime, "%.4d-%.2d-%.2d %.2d:%.2d:%.2d", DBvar.m_pdate->year, DBvar.m_pdate->month, DBvar.m_pdate->day, DBvar.m_pdate->hour, DBvar.m_pdate->minute, DBvar.m_pdate->second);
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		e->GetErrorMessage(buf, 260);
		e->Delete();
		Rs.Close();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}
	return TRUE;
}



void CDatabaseManager::LogprintfDB(const TCHAR *msg, ...)
{
	TCHAR	szLogLine[16000], szFinalLine[16000];
	va_list	ap;
	DWORD	n, nBytesWritten;
	SYSTEMTIME	ltime;


	HANDLE hFile = ::CreateFile(g_prefs.m_logfolder + _T("\\ImportServiceDBerror.log"), GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
	if (hFile == INVALID_HANDLE_VALUE)
		return;

	// Seek to end of file
	::SetFilePointer(hFile, 0, NULL, FILE_END);

	va_start(ap, msg);
	n = vsprintf(szLogLine, msg, ap);
	va_end(ap);
	szLogLine[n] = '\0';

	::GetLocalTime(&ltime);
	_stprintf(szFinalLine, "[%.2d-%.2d %.2d:%.2d:%.2d.%.3d] %s\r\n", (int)ltime.wDay, (int)ltime.wMonth, (int)ltime.wHour, (int)ltime.wMinute, (int)ltime.wSecond, (int)ltime.wMilliseconds, szLogLine);

	::WriteFile(hFile, szFinalLine, (DWORD)_tcsclen(szFinalLine), &nBytesWritten, NULL);

	::CloseHandle(hFile);

#ifdef _DEBUG
	TRACE(szFinalLine);
#endif
}

/////////////////////////////////////////////////
// Translate CTime to SQL Server DATETIME type
/////////////////////////////////////////////////

CString CDatabaseManager::TranslateDate(CTime tDate)
{
	CString dd, mm, yy, yysmall;

	dd.Format("%.2d", tDate.GetDay());
	mm.Format("%.2d", tDate.GetMonth());
	yy.Format("%.4d", tDate.GetYear());
	yysmall.Format("%.2d", 2000 - tDate.GetYear());

	if (g_prefs.m_dateformat == 2)
		return yy + "-" + mm + "-" + dd + "T00:00:00";
	else
		return mm + "/" + dd + "/" + yy;
}

CString CDatabaseManager::TranslateTimeEx(SYSTEMTIME tDate)
{
	CString dd, mm, yy, yysmall, hour, min, sec, msec;

	dd.Format("%.2d", tDate.wDay);
	mm.Format("%.2d", tDate.wMonth);
	yy.Format("%.4d", tDate.wYear);
	yysmall.Format("%.2d", 2000 - tDate.wYear);

	hour.Format("%.2d", tDate.wHour);
	min.Format("%.2d", tDate.wMinute);
	sec.Format("%.2d", tDate.wSecond);
	msec.Format("%.3d", tDate.wMilliseconds);


	if (g_prefs.m_dateformat == 2) {
		if (tDate.wMilliseconds != 0)
			return yy + "-" + mm + "-" + dd + "T" + hour + ":" + min + ":" + sec + "." + msec;
		else
			return yy + "-" + mm + "-" + dd + "T" + hour + ":" + min + ":" + sec;

	}
	else {
		if (tDate.wMilliseconds != 0)
			return mm + "/" + dd + "/" + yy + " " + hour + ":" + min + ":" + sec + "." + msec;
		else
			return mm + "/" + dd + "/" + yy + " " + hour + ":" + min + ":" + sec;
	}
}


CString CDatabaseManager::TranslateTime(CTime tDate)
{
	CString dd, mm, yy, yysmall, hour, min, sec;

	dd.Format("%.2d", tDate.GetDay());
	mm.Format("%.2d", tDate.GetMonth());
	yy.Format("%.4d", tDate.GetYear());
	yysmall.Format("%.2d", 2000 - tDate.GetYear());

	hour.Format("%.2d", tDate.GetHour());
	min.Format("%.2d", tDate.GetMinute());
	sec.Format("%.2d", tDate.GetSecond());


	if (g_prefs.m_dateformat == 2)
		return yy + "-" + mm + "-" + dd + "T" + hour + ":" + min + ":" + sec;
	else
		return mm + "/" + dd + "/" + yy + " " + hour + ":" + min + ":" + sec;
}


int CDatabaseManager::FieldExists(CString sTableName, CString sFieldName, CString &sErrorMessage)
{
	CString sSQL, s;
	BOOL bFound = FALSE;
	sErrorMessage = _T("");
	if (InitDB(sErrorMessage) == FALSE)
		return -1;

	CRecordset Rs(m_pDB);
	sSQL.Format("{ CALL sp_columns ('%s') }", sTableName);

	try {
		if (!Rs.Open(CRecordset::snapshot, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("Query failed - %s", (LPCSTR)sSQL);
			return FALSE;
		}
		while (!Rs.IsEOF()) {

			Rs.GetFieldValue((short)3, s);

			if (s.CompareNoCase(sFieldName) == 0) {
				bFound = TRUE;
				break;
			}
			Rs.MoveNext();
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return -1;
	}

	return bFound ? 1 : 0;
}

int CDatabaseManager::TableExists(CString sTableName, CString &sErrorMessage)
{
	CString sSQL, s;
	BOOL bFound = FALSE;
	sErrorMessage = _T("");
	if (InitDB(sErrorMessage) == FALSE)
		return -1;

	CRecordset Rs(m_pDB);
	sSQL.Format("{ CALL sp_tables ('%s') }", sTableName);

	try {
		if (!Rs.Open(CRecordset::snapshot, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("Query failed - %s", (LPCSTR)sSQL);
			return FALSE;
		}
		if (!Rs.IsEOF()) {

			bFound = TRUE;
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return -1;
	}

	return bFound ? 1 : 0;
}

int CDatabaseManager::StoredProcedureExists(CString sProcedureName, CString &sErrorMessage)
{
	CString sSQL, s;
	BOOL bFound = FALSE;
	sErrorMessage = _T("");
	if (InitDB(sErrorMessage) == FALSE)
		return -1;

	CRecordset Rs(m_pDB);
	sSQL.Format("{ CALL sp_stored_procedures ('%s') }", sProcedureName);

	try {
		if (!Rs.Open(CRecordset::snapshot, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("Query failed - %s", (LPCSTR)sSQL);
			return FALSE;
		}
		if (!Rs.IsEOF()) {

			bFound = TRUE;
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return -1;
	}

	return bFound ? 1 : 0;
}

int CDatabaseManager::StoredProcParameterExists(CString sSPname, CString sParameterName, CString &sErrorMessage)
{
	CString sSQL, s;
	BOOL bFound = FALSE;
	sErrorMessage = _T("");
	if (InitDB(sErrorMessage) == FALSE)
		return -1;

	CRecordset Rs(m_pDB);
	sSQL.Format("{ CALL sp_sproc_columns ('%s') }", sSPname);

	try {
		if (!Rs.Open(CRecordset::snapshot, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("Query failed - %s", (LPCSTR)sSQL);
			return FALSE;
		}
		while (!Rs.IsEOF()) {

			Rs.GetFieldValue((short)3, s);

			if (s.CompareNoCase(sParameterName) == 0) {
				bFound = TRUE;
				break;
			}
			Rs.MoveNext();
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return -1;
	}

	return bFound ? 1 : 0;
}

BOOL CDatabaseManager::InsertPlanLogEntry(int nProcessID, int nEventCode, CString sProductionName, CString sPageDescription, int nProductionID, CString sSender, CString &sErrorMessage)
{
	__int64 sep = GetFirstSeparation(nProductionID, sErrorMessage);

	CString sSQL;

	sErrorMessage = _T("");

	sSQL.Format("INSERT INTO Log (EventTime, ProcessID,Event,Separation,FlatSeparation, ErrorMsg,[FileName],Version,MiscInt,MiscString) VALUES (GETDATE(), %d,%d,%I64d,0,'%s','%s',1,%d,'%s')", nProcessID, nEventCode, sep, sPageDescription, sProductionName, nProductionID, sSender);


	if (InitDB(sErrorMessage) == FALSE) {
		return FALSE;
	}
	try {
		m_pDB->ExecuteSQL(sSQL);

	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}


	return TRUE;
}

__int64 CDatabaseManager::GetFirstSeparation(int nProductionID, CString &sErrorMessage)
{
	CString sSQL;

	__int64 n64Sep = 0;

	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return 0;

	CRecordset Rs(m_pDB);

	sSQL.Format("SELECT TOP 1 Separation FROM PageTable WITh (NOLOCK) WHERE ProductionID=%d", nProductionID);

	CString sThisPassword;
	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)sSQL);
			return 0;
		}

		if (!Rs.IsEOF()) {
			CString s;
			Rs.GetFieldValue((short)0, s);
			n64Sep = _atoi64(s);

		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		e->Delete();
		Rs.Close();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return 0;
	}
	return n64Sep;
}

int CDatabaseManager::FindPageFormat(double fSizeX, double fSizeY, double fBleed, int nSnapEven, int nSnapOdd, CString &sErrorMessage)
{
	int nID = 0;

	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return -1;

	CRecordset Rs(m_pDB);

	CString sSQL;
	sSQL.Format("SELECT PageFormatID FROM PageFormatNames WHERE Width=%.3f AND Height=%.3f AND Bleed=%.3f AND SnapModeEven=%d AND SnapModeOdd=%d", fSizeX, fSizeY, fBleed, nSnapEven, nSnapOdd);

	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)sSQL);
			return -1;
		}

		if (!Rs.IsEOF()) {
			CString s;
			Rs.GetFieldValue((short)0, s);
			nID = atoi(s);
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error : %s   (%s)", e->m_strError, sSQL);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return -1;
	}
	return nID;
}


int CDatabaseManager::FindPageFormat(double fSizeX, double fSizeY, CString &sErrorMessage)
{
	int nID = 0;

	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return -1;

	CRecordset Rs(m_pDB);

	CString sSQL;
	sSQL.Format("SELECT PageFormatID FROM PageFormatNames WHERE Width=%.3f AND Height=%.3f", fSizeX, fSizeY);

	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)sSQL);
			return -1;
		}

		if (!Rs.IsEOF()) {
			CString s;
			Rs.GetFieldValue((short)0, s);
			nID = atoi(s);
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error : %s   (%s)", e->m_strError, sSQL);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return -1;
	}
	return nID;
}


int CDatabaseManager::FindPageFormat(CString sPageFormatName, CString &sErrorMessage)
{
	int nID = 0;

	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return -1;

	CRecordset Rs(m_pDB);

	CString sSQL;
	sSQL.Format("SELECT PageFormatID FROM PageFormatNames WHERE PageFormatName='%s'", sPageFormatName);

	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)sSQL);
			return -1;
		}

		if (!Rs.IsEOF()) {
			CString s;
			Rs.GetFieldValue((short)0, s);
			nID = atoi(s);
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error : %s   (%s)", e->m_strError, sSQL);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return -1;
	}

	if (nID == 0 && sPageFormatName.Find("X") != -1) {
		// Try alternative way
		sPageFormatName.Replace(" ", "");
		// 260x360
		// 0123456
		CString sXdim = sPageFormatName.Left(sPageFormatName.Find("X"));
		CString sYdim = sPageFormatName.Mid(sPageFormatName.Find("X") + 1);
		nID = FindPageFormat(atof(sXdim.Trim()), atof(sYdim.Trim()), sErrorMessage);

	}

	return nID;
}



int CDatabaseManager::InsertPageFormat(CString sName, double fSizeX, double fSizeY, double fBleed,
	int nSnapEven, int nSnapOdd, int nType, CString &sErrorMessage)
{
	int nPageFormatID = 0;
	CString sSQL;
	CString sPageFormatName;

	sErrorMessage.Format("");

	if (sName == "")
	{
		if (fBleed > 0)
			sPageFormatName.Format("%.1f x %.1f - %.1f bleed", fSizeX, fSizeY);
		else
			sPageFormatName.Format("%.1f x %.1f ", fSizeX, fSizeY);
		if (nSnapEven == SNAP_CENTER)
			sPageFormatName += " centered";
		if (nType > 0)
			sPageFormatName += " to gutter";
	}

	nPageFormatID = RetrieveNextValueCount("PageFormatID", "PageFormatNames", sErrorMessage);
	sSQL.Format("INSERT INTO PageFormatNames (PageFormatID,PageFormatName, Width, Height,Bleed,Flag,ImageWidth,ImageHeight,ImageOffsetX,ImageOffsetY,ShiftImageLeftPage,ShiftImageRightPage,OffsetEvenX,OffsetEvenY,OffsetOddX,OffsetOddY,SnapModeEven,SnapModeOdd) VALUES (%d,'%s',%f,%f,%f, %d,0,0,0,0,0,0,0,0,0,0,%d,%d)", nPageFormatID, sPageFormatName, fSizeX, fSizeY, fBleed, nType, nSnapEven, nSnapOdd);

	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	try {
		m_pDB->ExecuteSQL(sSQL);
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Insert failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error : %s   (%s)", e->m_strError, sSQL);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {			// So what..! Go on!;
		}
		return FALSE;
	}
	return nPageFormatID;
}

BOOL CDatabaseManager::PublicationPlanLock(int nPublicationID, CTime tPubDate, int bRequestedPlanLock, int *bCurrentPlanLock, TCHAR *szClientName, TCHAR *szClientTime, CString &sErrorMessage)
{
	CString sSQL;
	strcpy(szClientTime, "");
	strcpy(szClientName, "");
	*bCurrentPlanLock = PLANLOCK_UNKNOWN;

	if (nPublicationID == 0)
		return TRUE;

	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	CRecordset Rs(m_pDB);

	CString sClientName;
	TCHAR buf[260];
	DWORD len = 260;
	::GetComputerName(buf, &len);
	sSQL.Format("{CALL spPublicationLock (%d, '%s',%d,'IMPORTCENTER%d %s')}", nPublicationID, TranslateDate(tPubDate), bRequestedPlanLock, g_prefs.m_instancenumber, buf);
	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)sSQL);
			return FALSE;
		}

		if (!Rs.IsEOF()) {
			CString s;
			Rs.GetFieldValue((short)0, s);
			*bCurrentPlanLock = atoi(s);
			Rs.GetFieldValue((short)1, s);
			strcpy(szClientName, s);

			CDBVariant DBvar;
			Rs.GetFieldValue((short)2, DBvar, SQL_C_TIMESTAMP);
			sprintf(szClientTime, "%.4d-%.2d-%.2d %.2d:%.2d:%.2d", DBvar.m_pdate->year, DBvar.m_pdate->month, DBvar.m_pdate->day, DBvar.m_pdate->hour, DBvar.m_pdate->minute, DBvar.m_pdate->second);
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error : %s   (%s)", e->m_strError, sSQL);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}
	return TRUE;
}


BOOL CDatabaseManager::AddRetryRequest(int nProductionID, CString &sErrorMessage)
{
	CString sSQL;

	sErrorMessage = _T("");

	sSQL.Format("INSERT INTO AutoRetryQueue SELECT %d,0,'',GETDATE()", nProductionID);

	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	try {
		m_pDB->ExecuteSQL(sSQL);

	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Insert failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error : %s   (%s)", e->m_strError, sSQL);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}


	return TRUE;
}

int CDatabaseManager::GetPlanType(int nPublicationID, CTime tPubDate, int nPressID, int &nExstingProductionID, CString &sErrorMessage)
{
	CString sSQL, s;

	int nPlanType = PLANTYPE_UNKNOWN;
	nExstingProductionID = 0;

	sSQL.Format("SELECT TOP 1 PN.PlanType,PN.ProductionID FROM ProductionNames PN INNER JOIN PageTable P WITH (NOLOCK) ON P.ProductionID=PN.ProductionID WHERE P.PubDate='%s' AND PublicationID=%d AND PressID=%d",
		TranslateDate(tPubDate), nPublicationID, nPressID);

	sErrorMessage.Format("");

	BOOL bSuccess = FALSE;
	int nRetries = g_prefs.m_nQueryRetries;
	int nCount = 0;

	while (!bSuccess && nRetries-- > 0) {

		if (InitDB(sErrorMessage) == FALSE) {
			Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
			continue;
		}
		CRecordset Rs(m_pDB);

		try {
			if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
				sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)sSQL);
				return FALSE;
			}

			if (!Rs.IsEOF()) {
				Rs.GetFieldValue((short)0, s);
				nPlanType = atoi(s);
				Rs.GetFieldValue((short)1, s);
				nExstingProductionID = atoi(s);
			}
			Rs.Close();
			bSuccess = TRUE;
			break;
		}
		catch (CDBException* e) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
			LogprintfDB("Error Try %d - : %s   (%s)", nRetries, e->m_strError, sSQL);
			e->Delete();
			try {
				m_DBopen = FALSE;
				m_pDB->Close();
			}
			catch (CDBException* e) {
				// So what..! Go on!;
			}
			Sleep(g_prefs.m_QueryBackoffTime);
		}
	}

	return bSuccess ? nPlanType : PLANTYPE_ERROR;
}


BOOL CDatabaseManager::GetPressRunDetails(int nPressRunID, int &nProductionID, int &nPressID,
	CTime &tPubDate, int &nPublicationID, int &nEditionID, CUIntArray &aSectionID, CString &sErrorMessage)
{
	nProductionID = 0;
	nPressID = 0;
	tPubDate = CTime(1975, 1, 1, 0, 0, 0);
	nPublicationID = 0;
	nEditionID = 0;
	aSectionID.RemoveAll();

	CString sSQL, s;

	if (nPressRunID == 0)
		return FALSE;

	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	CRecordset Rs(m_pDB);

	sSQL.Format("SELECT DISTINCT PressID,PublicationID,PubDate,EditionID,SectionID,ProductionID FROM PageTable WITH (NOLOCK) WHERE PressRunID=%d ORDER BY SectionID", nPressRunID);
	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)sSQL);
			return FALSE;
		}

		while (!Rs.IsEOF()) {
			CString s;
			Rs.GetFieldValue((short)0, s);
			nPressID = atoi(s);

			Rs.GetFieldValue((short)1, s);
			nPublicationID = atoi(s);

			CDBVariant DBvar;
			Rs.GetFieldValue((short)2, DBvar, SQL_C_TIMESTAMP);
			tPubDate = CTime(DBvar.m_pdate->year, DBvar.m_pdate->month, DBvar.m_pdate->day, 0, 0, 0);

			Rs.GetFieldValue((short)3, s);
			nEditionID = atoi(s);

			Rs.GetFieldValue((short)4, s);
			aSectionID.Add(atoi(s));

			Rs.GetFieldValue((short)5, s);
			nProductionID = atoi(s);

			Rs.MoveNext();
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error : %s   (%s)", e->m_strError, sSQL);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}

	return TRUE;
}




BOOL CDatabaseManager::GetPagesInPressRun(int nPressRunID, int nSectionID, CStringArray &aPages, CString &sErrorMessage)
{
	aPages.RemoveAll();

	CString sSQL, s;

	if (nPressRunID == 0)
		return FALSE;

	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	CRecordset Rs(m_pDB);

	sSQL.Format("SELECT DISTINCT PageName FROM PageTable WITH (NOLOCK) WHERE PressRunID=%d AND SectionID=%d ORDER BY PageName", nPressRunID, nSectionID);
	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)sSQL);
			return FALSE;
		}

		while (!Rs.IsEOF()) {
			CString s;
			Rs.GetFieldValue((short)0, s);
			aPages.Add(s);

			Rs.MoveNext();
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error : %s   (%s)", e->m_strError, sSQL);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}

	return TRUE;
}




BOOL CDatabaseManager::GetOtherUniqueProductionID(int nPressID,
	CTime tPubDate, int nPublicationID, int nEditionID, int nSectionID, int &nOtherProductionID, CString &sErrorMessage)
{
	nOtherProductionID = 0;
	CString sSQL, s;

	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	CRecordset Rs(m_pDB);

	sSQL.Format("SELECT TOP 1 ProductionID FROM PageTable WITH (NOLOCK) WHERE PubDate='%s' AND PublicationID=%d AND EditionID=%d AND PressID <> %d AND SectionID=%d AND UniquePage=1",
		TranslateDate(tPubDate), nPublicationID, nEditionID, nPressID, nSectionID);
	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)sSQL);
			return FALSE;
		}

		if (!Rs.IsEOF()) {
			CString s;
			Rs.GetFieldValue((short)0, s);
			nOtherProductionID = atoi(s);
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error : %s   (%s)", e->m_strError, sSQL);
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}

	return TRUE;
}


BOOL CDatabaseManager::ImportCenterPressRunCustomPage(int nPressRunID, int nOtherProductionID, int EditionID, int SectionID, CString sPageName, CString &sErrorMessage)
{
	CString sSQL;

	sErrorMessage = _T("");

	sSQL.Format("{ CALL spImportCenterPressRunCustomPage (%d,%d,%d,%d,'%s') }", nPressRunID, nOtherProductionID, EditionID, SectionID, sPageName);

	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	try {
		m_pDB->ExecuteSQL(sSQL);

	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Insert failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error : %s   (%s)", e->m_strError, sSQL);
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}

	return TRUE;
}

BOOL CDatabaseManager::LinkMasterCopySepSetsToOtherPress(int nPressRunID, CString &sErrorMessage)
{
	g_util.Logprintf("INFO: LinkMasterCopySepSetsToOtherPress called for pressrun %d..", nPressRunID);
	int nPublicationID;
	CTime tPubDate;
	int nPressID = 0;
	int nProductionID = 0;
	int nEditionID = 0;
	CUIntArray  aSectionID;
	aSectionID.RemoveAll();
	int nOtherProductionID = 0;

	BOOL bSuccess = FALSE;
	int nRetries = g_prefs.m_nQueryRetries;
	while (bSuccess == FALSE && --nRetries >= 0) {
		bSuccess = GetPressRunDetails(nPressRunID, nProductionID, nPressID, tPubDate, nPublicationID, nEditionID, aSectionID, sErrorMessage);

		if (bSuccess == FALSE)
			Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
	}

	if (bSuccess == FALSE)
		return FALSE;

	bSuccess = FALSE;
	nRetries = g_prefs.m_nQueryRetries;
	if (aSectionID.GetCount() > 0) {
		while (bSuccess == FALSE && --nRetries >= 0) {
			bSuccess = GetOtherUniqueProductionID(nPressID,
				tPubDate, nPublicationID, nEditionID, aSectionID.GetAt(0), nOtherProductionID, sErrorMessage);
			if (bSuccess == FALSE)
				Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
		}
	}

	if (bSuccess == FALSE)
		return FALSE;

	if (nOtherProductionID == 0)
		return TRUE;

	g_util.Logprintf("INFO: LinkMasterCopySepSetsToOtherPress - Other (unique) ProductionID detected - %d", nOtherProductionID);

	CStringArray aPageNames;

	for (int nSec = 0; nSec < aSectionID.GetCount(); nSec++) {
		aPageNames.RemoveAll();

		bSuccess = FALSE;
		nRetries = g_prefs.m_nQueryRetries;
		while (bSuccess == FALSE && --nRetries >= 0) {
			bSuccess = GetPagesInPressRun(nPressRunID, aSectionID[nSec], aPageNames, sErrorMessage);
			if (bSuccess == FALSE)
				Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
		}

		if (bSuccess == FALSE)
			return FALSE;

		for (int nPage = 0; nPage < aPageNames.GetCount(); nPage++) {
			bSuccess = FALSE;
			nRetries = g_prefs.m_nQueryRetries;
			while (bSuccess == FALSE && --nRetries >= 0) {

				g_util.Logprintf("INFO: spImportCenterPressRunCustomPage (%d, %d, %d, %d, %s)", nPressRunID, nOtherProductionID, nEditionID, aSectionID[nSec], aPageNames[nPage]);

				bSuccess = ImportCenterPressRunCustomPage(nPressRunID, nOtherProductionID, nEditionID, aSectionID[nSec], aPageNames[nPage], sErrorMessage);
				if (bSuccess == FALSE)
					Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
			}

			if (bSuccess == FALSE)
				return FALSE;
		}
	}

	return FALSE;
}

/*
extern LOGLIST g_loglist;
BOOL CDatabaseManager::GetLog(CString &sErrorMessage)
{
	g_loglist.RemoveAll();

	CString sSQL = "select Eventtime, Ec.EventName,filename,log.ErrorMsg,miscstring from log inner join eventcodes ec on log.event=ec.EventNumber where event>=900 order by eventtime desc";

	sErrorMessage = _T("");
	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	CRecordset Rs(m_pDB);
	CDBVariant DBvar;

	try {
		if (!Rs.Open(CRecordset::snapshot, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("Query failed - %s", (LPCSTR)sSQL);
			return FALSE;
		}
		while (!Rs.IsEOF()) {
			LOGITEM	*p = new LOGITEM();

			Rs.GetFieldValue((short)0, DBvar, SQL_C_TIMESTAMP);
			p->sTime.Format("%.4d-%.2d-%.2d %.2d:%.2d:%.2d", DBvar.m_pdate->year, DBvar.m_pdate->month, DBvar.m_pdate->day, DBvar.m_pdate->hour, DBvar.m_pdate->minute, DBvar.m_pdate->second);

			Rs.GetFieldValue((short)1, p->sEvent);
			Rs.GetFieldValue((short)2, p->sProduct);
			Rs.GetFieldValue((short)3, p->sComment);
			Rs.GetFieldValue((short)4, p->sProcess);

			g_loglist.Add(*p);

			Rs.MoveNext();
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}


	return TRUE;

}
*/

BOOL CDatabaseManager::AddRetryRequestFileCenter(int nProductionID, CString &sErrorMessage)
{
	return AddRetryRequestFileCenter(nProductionID, _T(""), 0, sErrorMessage);
}

BOOL CDatabaseManager::AddRetryRequestFileCenter(int nProductionID, CString sMask, CString &sErrorMessage)
{
	return AddRetryRequestFileCenter(nProductionID, sMask, 0, sErrorMessage);
}

BOOL CDatabaseManager::AddRetryRequestFileCenter(int nProductionID, CString sMask, BOOL bProcessOnlyUnprocessedFiles, CString &sErrorMessage)
{
	CString sSQL;

	sErrorMessage = _T("");

	sSQL.Format("INSERT INTO AutoRetryQueueFileCenter SELECT %d,%d,'%s',GETDATE()", nProductionID, bProcessOnlyUnprocessedFiles, sMask);

	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	try {
		m_pDB->ExecuteSQL(sSQL);

	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Insert failed - %s", (LPCSTR)e->m_strError);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}


	return TRUE;
}

int CDatabaseManager::GetPageCountInSection(int nPublicationID, CTime tPubDate, int nEditionID, int nSectionID, CString &sErrorMessage)
{
	CUtils util;
	CString sSQL;
	int nNumberOfPages = 0;
	BOOL bSuccess = FALSE;
	int nRetries = g_prefs.m_nQueryRetries;

	sErrorMessage.Format("");

	sSQL.Format("SELECT COUNT (DISTINCT PageName) FROM PageTable WITH (NOLOCK) WHERE PublicationID=%d AND PubDate='%s' AND EditionID=%d AND SectionID=%d AND Active>0 AND PageType<=2", nPublicationID, TranslateDate(tPubDate), nEditionID, nSectionID);

	while (!bSuccess && nRetries-- > 0) {

		if (InitDB(sErrorMessage) == FALSE) {
			Sleep(g_prefs.m_QueryBackoffTime); // Time between retries
			continue;
		}
		CRecordset Rs(m_pDB);

		try {
			if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
				sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)sSQL);
				return -1;
			}

			if (!Rs.IsEOF()) {
				CString s;

				Rs.GetFieldValue((short)0, s);
				nNumberOfPages = atoi(s);

			}
			Rs.Close();
			bSuccess = TRUE;
			break;
		}
		catch (CDBException* e) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
			LogprintfDB("Error Try %d - : %s   (%s)", nRetries, e->m_strError, sSQL);
			e->Delete();
			try {
				m_DBopen = FALSE;
				m_pDB->Close();
			}
			catch (CDBException* e) {
				// So what..! Go on!;
			}
			Sleep(g_prefs.m_QueryBackoffTime);
		}
	}

	return bSuccess ? nNumberOfPages : -1;
}



BOOL CDatabaseManager::BypassRetryRequest(int nProductionID, BOOL &bBypassRetry, CString &sErrorMessage)
{
	CUtils util;
	bBypassRetry = FALSE;

	sErrorMessage.Format("");
	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	CRecordset Rs(m_pDB);

	CString sSQL;
	sSQL.Format("{ CALL spImportCenterBypassRetryRequest (%d)}", nProductionID);
	util.Logprintf("INFO: spImportCenterBypassRetryRequest %d called..", nProductionID);



	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("Query failed - %s", (LPCSTR)sSQL);
			return FALSE;
		}

		if (!Rs.IsEOF()) {
			CString s;
			Rs.GetFieldValue((short)0, s);
			bBypassRetry = atoi(s);
			util.Logprintf("INFO: spImportCenterBypassRetryRequest %d returned %d", nProductionID, bBypassRetry);
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", (LPCSTR)e->m_strError);
		LogprintfDB("Error : %s   (%s)", e->m_strError, sSQL);
		e->Delete();
		Rs.Close();

		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}

	return TRUE;
}


BOOL CDatabaseManager::NPDeleteIllegalLocationData(CString &sErrorMessage)
{
	CString sSQL;
	sErrorMessage = _T("");


	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	sSQL.Format("DELETE FROM ImportTable2 WHERE Location IN ('P1','P2','P3','P4','P5','P6','P7','P8','P9')");
	try {
		m_pDB->ExecuteSQL(sSQL);
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", e->m_strError);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}

	return TRUE;
}


BOOL CDatabaseManager::NPChangeSeqNumber(CString sPublication, CTime tPubDate, CString sLocation, int nSeq, CString &sErrorMessage)
{
	CString sSQL;
	sErrorMessage = _T("");

	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	sSQL.Format("UPDATE ImportTable2 SET SequenceNumber=%d WHERE Publication='%s' AND PubDate='%s' AND Location='%s' ", nSeq, sPublication, TranslateDate(tPubDate), sLocation);
	try {
		m_pDB->ExecuteSQL(sSQL);
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", e->m_strError);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}

	return TRUE;
}




BOOL CDatabaseManager::NPDeleteOldData(int nDaysToKeep, CString &sErrorMessage)
{
	CString sSQL;
	sErrorMessage = _T("");

	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	if (nDaysToKeep <= 0)
		sSQL.Format("DELETE FROM ImportTable2");
	else
		sSQL.Format("DELETE FROM ImportTable2 WHERE DATEDIFF(day, PubDate, GETDATE()) > %d", nDaysToKeep);
	try {
		m_pDB->ExecuteSQL(sSQL);
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", e->m_strError);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}

	return TRUE;
}

int CDatabaseManager::NPLookupPageEntry(CString sPublication, CTime tPubDate, CString sLocation, CString sEdition, CString sZone, CString sSection, CString sPageName, CString &sErrorMessage)
{
	CString sSQL, s;
	sErrorMessage = _T("");
	int nRet = 0;

	sSQL.Format("SELECT PageName FROM ImportTable2 WHERE Publication='%s' AND PubDate='%s' AND Location='%s' AND Edition='%s' AND Zone='%s' AND Section='%s' AND PageName='%s'",
		sPublication, TranslateDate(tPubDate), sLocation, sEdition, sZone, sSection, sPageName);

	if (InitDB(sErrorMessage) == FALSE)
		return -1;

	CRecordset Rs(m_pDB);


	try {
		if (!Rs.Open(CRecordset::snapshot, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("Query failed - %s", sSQL);

			return -1;
		}
		if (!Rs.IsEOF()) {
			Rs.GetFieldValue((short)0, s);
			nRet = 1;

		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s - %s", e->m_strError, sSQL);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return -1;
	}

	return nRet;
}

BOOL CDatabaseManager::NPResetPages(CString &sErrorMessage)
{
	CString sSQL;
	sErrorMessage = _T("");

	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	sSQL.Format("UPDATE ImportTable2 SET MiscInt4=0");

	try {
		m_pDB->ExecuteSQL(sSQL);
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Insert/update failed - %s - %s", e->m_strError, sSQL);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}

	return TRUE;
}

BOOL CDatabaseManager::NPDeletePages(CString sLocation, CString sPublication, CTime tPubDate, CString sZone, CString sEdition, CString sSection, CString &sErrorMessage)
{
	CString sSQL;

	sErrorMessage =  _T("");
	
	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	sSQL.Format("DELETE ImportTable2 WHERE Publication='%s' AND PubDate='%s' AND Location='%s' AND Edition='%s' AND Zone='%s' AND Section='%s'",
		sPublication, TranslateDate(tPubDate), sLocation, sEdition, sZone, sSection);

	if (sLocation == "" && sEdition == "" && sZone == "" && sSection == "")
		sSQL.Format("DELETE ImportTable2 WHERE Publication='%s' AND PubDate='%s'", sPublication, TranslateDate(tPubDate));

	try {
		m_pDB->ExecuteSQL(sSQL);
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Insert/update failed - %s - %s", e->m_strError, sSQL);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}

	return TRUE;
}

BOOL CDatabaseManager::NPAddPage(CString sPublication, CTime tPubDate, CString sProdcutionName, int nSeqNo, CString sLocation, CString sEdition,
	CString sZone, CString sSection,
	CString sPageName, int nPagination, CString sPlanPagename,
	CString sComment, CString sColors, CString sVersion, CTime tDeadLine,
	CString &sErrorMessage, int &bAdded)
{
	return NPAddPage(sPublication, tPubDate, sProdcutionName, nSeqNo, sLocation, sEdition,
		sZone, sSection,
		sPageName, nPagination, sPlanPagename,
		sComment, sColors, sVersion, tDeadLine,
		0, 0, 0, 0, "", "", "", "", CTime(1975, 1, 1, 0, 0, 0),
		sErrorMessage, bAdded);
}

BOOL CDatabaseManager::NPAddPage(CString sPublication, CTime tPubDate, CString sProdcutionName, int nSeqNo, CString sLocation, CString sEdition,
	CString sZone, CString sSection,
	CString sPageName, int nPagination, CString sPlanPagename,
	CString sComment, CString sColors, CString sVersion, CTime tDeadLine,
	int nMiscInt1, int nMiscInt2, int nMiscInt3, int nMiscInt4, CString sMiscString1, CString sMiscString2, CString sMiscString3, CString sMiscString4, CTime tMiscDate,
	CString &sErrorMessage, int &bAdded)
{
	CString sSQL;
	bAdded = FALSE;

	nMiscInt4 = 1;	// Recently added flag..

	if (sProdcutionName == "")
		sProdcutionName.Format("%s-%s-%.2d%.2d-%s_%s-%s", sLocation, sPublication, tPubDate.GetDay(), tPubDate.GetMonth(),
			sZone, sEdition, sSection);

	BOOL bNewEntry = TRUE;
	if (NPLookupPageEntry(sPublication, tPubDate, sLocation, sEdition, sZone, sSection, sPageName, sErrorMessage) > 0)
		bNewEntry = FALSE;


	sErrorMessage = _T("");

	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	if (bNewEntry)
		sSQL.Format("INSERT INTO ImportTable2 (Publication, PubDate, ProductionName, SequenceNumber, Location, Edition, Zone, Section, PageName, Pagination, PlanPageName,Comment,Colors,Version,DeadLine,MiscInt1,MiscInt2,MiscInt3,MiscInt4,MiscString1,MiscString2,MiscString3,MiscString4,MiscDate) VALUES ('%s', '%s','%s',%d,'%s','%s','%s','%s','%s',%d,'%s','%s','%s','%s','%s',%d,%d,%d,%d,'%s','%s','%s','%s','%s')",
			sPublication, TranslateDate(tPubDate), sProdcutionName, nSeqNo, sLocation, sEdition, sZone, sSection, sPageName, nPagination, sPlanPagename, sComment, sColors, sVersion, TranslateTime(tDeadLine), nMiscInt1, nMiscInt2, nMiscInt3, nMiscInt4, sMiscString1, sMiscString2, sMiscString3, sMiscString4, TranslateTime(tMiscDate));
	else
		sSQL.Format("UPDATE ImportTable2 SET PlanPageName='%s', Comment='%s', Colors='%s', Version='%s',DeadLine='%s',MiscInt1=%d,MiscInt2=%d,MiscInt3=%d,MiscInt4=%d,MiscString1='%s',MiscString2='%s',MiscString3='%s',MiscString4='%s',MiscDate='%s' WHERE Publication='%s' AND PubDate='%s' AND Location='%s' AND Edition='%s' AND Zone='%s' AND Section='%s' AND PageName='%s'",

			sPlanPagename, sComment, sColors, sVersion, TranslateTime(tDeadLine), nMiscInt1, nMiscInt2, nMiscInt3, nMiscInt4, sMiscString1, sMiscString2, sMiscString3, sMiscString4, TranslateTime(tMiscDate),
			sPublication, TranslateDate(tPubDate), sLocation, sEdition, sZone, sSection, sPageName);
	try {
		m_pDB->ExecuteSQL(sSQL);
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Insert/update failed - %s - %s", e->m_strError, sSQL);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}

	bAdded = bNewEntry;
	return TRUE;

}

BOOL CDatabaseManager::NPGetPageCount(CString sPublication, CTime tPubDate, CString sEdition, CString sSection, int &nPageCount, CString &sErrorMessage)
{

	nPageCount = 0;
	CString sSQL, s, s2;


	CString sZone = "";
	if (sEdition != "") {
		int m = sEdition.Find("_");
		if (m != -1) {
			sZone = sEdition.Left(m);
			sEdition = sEdition.Mid(m + 1);
		}
	}

	if (sEdition == "")
		sSQL.Format("SELECT COUNT(DISTINCT PageName) FROM ImportTable2 WHERE Publication='%s' AND PubDate='%s' AND Section='%s' AND MiscInt4=1", sPublication, TranslateDate(tPubDate), sSection);
	else
		sSQL.Format("SELECT COUNT(DISTINCT PageName) FROM ImportTable2 WHERE Publication='%s' AND PubDate='%s' AND Section='%s' AND Zone='%s' AND Edition='%s' AND MiscInt4=1", sPublication, TranslateDate(tPubDate), sSection, sZone, sEdition);

	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	CRecordset Rs(m_pDB);


	try {
		if (!Rs.Open(CRecordset::snapshot, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("Query failed - %s", sSQL);

			return FALSE;
		}
		if (!Rs.IsEOF()) {
			Rs.GetFieldValue((short)0, s);
			nPageCount = atoi(s);

		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s - %s", e->m_strError, sSQL);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {

		}
		return FALSE;
	}

	return TRUE;
}

BOOL CDatabaseManager::NPGetPageTablePageCount(CString sPublication, CTime tPubDate, CString sEdition, CString sSection, int &nPageCount, BOOL bCountAll, CString &sErrorMessage)
{

	int nPublicationID = GetPublicationID(sPublication, sErrorMessage);
	if (nPublicationID <= 0)
		return FALSE;

	int nSectionID = GetSectionID(sSection, sErrorMessage);
	if (nSectionID <= 0)
		return FALSE;

	int nEditionID = GetEditionID(sEdition, sErrorMessage);

	nPageCount = 0;
	CString sSQL, s;


	if (nEditionID == 0)
		sSQL.Format("SELECT COUNT(DISTINCT PageName) FROM PageTable WITH (NOLOCK) WHERE PublicationID=%d AND PubDate='%s' AND SectionID=%d AND Dirty=0 AND PageType<=2 AND UniquePage=1", nPublicationID, TranslateDate(tPubDate), nSectionID);
	else {
		if (bCountAll)
			sSQL.Format("SELECT COUNT(DISTINCT PageName) FROM PageTable WITH (NOLOCK) WHERE PublicationID=%d AND PubDate='%s' AND SectionID=%d AND EditionID=%d AND Dirty=0 AND PageType<=2", nPublicationID, TranslateDate(tPubDate), nSectionID, nEditionID);
		else
			sSQL.Format("SELECT COUNT(DISTINCT PageName) FROM PageTable WITH (NOLOCK) WHERE PublicationID=%d AND PubDate='%s' AND SectionID=%d AND EditionID=%d AND Dirty=0 AND PageType<=2 AND UniquePage=1", nPublicationID, TranslateDate(tPubDate), nSectionID, nEditionID);
	}
	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	CRecordset Rs(m_pDB);


	try {
		if (!Rs.Open(CRecordset::snapshot, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("Query failed - %s", sSQL);

			return FALSE;
		}
		if (!Rs.IsEOF()) {
			Rs.GetFieldValue((short)0, s);
			nPageCount = atoi(s);

		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s - %s", e->m_strError, sSQL);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {

		}
		return FALSE;
	}

	return TRUE;
}



BOOL CDatabaseManager::NPGetSequenceNumberCount(CString sPublication, CTime tPubDate, CString sLocation, int &nSequeceNumbers, CString &sErrorMessage)
{
	CUtils util;

	nSequeceNumbers = 0;
	CString sSQL, s, s2;
	if (sLocation == "")
		sSQL.Format("SELECT COUNT(DISTINCT SequenceNumber) FROM ImportTable2 WHERE Publication='%s' AND PubDate='%s' ", sPublication, TranslateDate(tPubDate));
	else
		sSQL.Format("SELECT COUNT(DISTINCT SequenceNumber) FROM ImportTable2 WHERE Publication='%s' AND PubDate='%s' AND Location='%s'", sPublication, TranslateDate(tPubDate), sLocation);
	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	CRecordset Rs(m_pDB);


	try {
		if (!Rs.Open(CRecordset::snapshot, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("Query failed - %s", sSQL);

			return FALSE;
		}
		if (!Rs.IsEOF()) {
			Rs.GetFieldValue((short)0, s);
			nSequeceNumbers = atoi(s);

		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s - %s", e->m_strError, sSQL);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {

		}
		return FALSE;
	}

	return TRUE;
}


BOOL CDatabaseManager::NPGetInfoFromImportTable(int nMode, CString sPublication, CTime tPubDate, CString sLocation, CString sZoneEdition, CStringArray &aList, CString &sErrorMessage)
{
	CUtils util;
	aList.RemoveAll();

	CString sZone = "";
	CString sEdition = sZoneEdition;
	if (sEdition != "") {
		int n = sEdition.Find("_");
		if (n != -1) {
			sZone = sEdition.Left(n);
			sEdition = sEdition.Mid(n + 1);
		}
	}
	CString sSQL, s, s2;
	sErrorMessage = _T("");

	if (nMode == IMPORTTABLEQUERYMODE_EDITIONS)
		sSQL.Format("SELECT DISTINCT Zone,Edition,SequenceNumber FROM ImportTable2 WHERE Publication='%s' AND PubDate='%s' ORDER BY SequenceNumber,Edition", sPublication, TranslateDate(tPubDate));
	else if (nMode == IMPORTTABLEQUERYMODE_LOCATIONS)
		sSQL.Format("SELECT DISTINCT Location,SequenceNumber FROM ImportTable2 WHERE Publication='%s' AND PubDate='%s'  ORDER BY SequenceNumber", sPublication, TranslateDate(tPubDate));
	else if (nMode == IMPORTTABLEQUERYMODE_EDITIONS_ON_LOCATION)
		sSQL.Format("SELECT DISTINCT Zone,Edition,SequenceNumber FROM ImportTable2 WHERE Publication='%s' AND PubDate='%s' AND Location='%s' ORDER BY SequenceNumber,Edition", sPublication, TranslateDate(tPubDate), sLocation);
	else if (nMode == IMPORTTABLEQUERYMODE_LOCATIONS_FOR_EDITION)
		sSQL.Format("SELECT DISTINCT Location FROM ImportTable2 WHERE Publication='%s' AND PubDate='%s' AND Edition='%s' AND Zone='%s'", sPublication, TranslateDate(tPubDate), sEdition, sZone);
	else if (nMode == IMPORTTABLEQUERYMODE_SECTIONS)
		sSQL.Format("SELECT DISTINCT Section FROM ImportTable2 WHERE Publication='%s' AND PubDate='%s'", sPublication, TranslateDate(tPubDate));
	else
		return FALSE;

	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	CRecordset Rs(m_pDB);

	try {
		if (!Rs.Open(CRecordset::snapshot, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("Query failed - %s", sSQL);

			return FALSE;
		}
		while (!Rs.IsEOF()) {
			Rs.GetFieldValue((short)0, s);
			if (nMode == IMPORTTABLEQUERYMODE_EDITIONS || nMode == IMPORTTABLEQUERYMODE_EDITIONS_ON_LOCATION) {
				Rs.GetFieldValue((short)1, s2);
				s = s + "_" + s2;
			}

			util.AddToCStringArray(aList, s);

			Rs.MoveNext();
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s - %s", e->m_strError, sSQL);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {

		}
		return FALSE;
	}

	return TRUE;
}


BOOL CDatabaseManager::NPGetInfoFromImportTableEx(int nMode, CString sPublication, CTime tPubDate, CString sLocation, CString sZoneEdition, CStringArray &aList, CString &sErrorMessage)
{
	CUtils util;
	aList.RemoveAll();

	CString sZone = "";
	CString sEdition = sZoneEdition;
	if (sEdition != "") {
		int n = sEdition.Find("_");
		if (n != -1) {
			sZone = sEdition.Left(n);
			sEdition = sEdition.Mid(n + 1);
		}
	}
	CString sSQL, s, s2;
	sErrorMessage = _T("");

	if (nMode == IMPORTTABLEQUERYMODE_EDITIONS)
		sSQL.Format("SELECT DISTINCT Zone,Edition,SequenceNumber FROM ImportTable2 WHERE Publication='%s' AND PubDate='%s' AND MiscInt4=1 ORDER BY SequenceNumber,Edition", sPublication, TranslateDate(tPubDate));
	else if (nMode == IMPORTTABLEQUERYMODE_LOCATIONS)
		sSQL.Format("SELECT DISTINCT Location,SequenceNumber FROM ImportTable2 WHERE Publication='%s' AND PubDate='%s' AND MiscInt4=1 ORDER BY SequenceNumber", sPublication, TranslateDate(tPubDate));
	else if (nMode == IMPORTTABLEQUERYMODE_EDITIONS_ON_LOCATION)
		sSQL.Format("SELECT DISTINCT Zone,Edition,SequenceNumber FROM ImportTable2 WHERE Publication='%s' AND PubDate='%s' AND Location='%s' AND MiscInt4=1 ORDER BY SequenceNumber,Edition", sPublication, TranslateDate(tPubDate), sLocation);
	else if (nMode == IMPORTTABLEQUERYMODE_LOCATIONS_FOR_EDITION)
		sSQL.Format("SELECT DISTINCT Location FROM ImportTable2 WHERE Publication='%s' AND PubDate='%s' AND Edition='%s' AND Zone='%s' AND MiscInt4=1", sPublication, TranslateDate(tPubDate), sEdition, sZone);
	else if (nMode == IMPORTTABLEQUERYMODE_SECTIONS)
		sSQL.Format("SELECT DISTINCT Section FROM ImportTable2 WHERE Publication='%s' AND PubDate='%s' AND MiscInt4=1", sPublication, TranslateDate(tPubDate));
	else
		return FALSE;

	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	CRecordset Rs(m_pDB);


	try {
		if (!Rs.Open(CRecordset::snapshot, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("Query failed - %s", sSQL);
			return FALSE;
		}
		while (!Rs.IsEOF()) {
			Rs.GetFieldValue((short)0, s);
			if (nMode == IMPORTTABLEQUERYMODE_EDITIONS || nMode == IMPORTTABLEQUERYMODE_EDITIONS_ON_LOCATION) {
				Rs.GetFieldValue((short)1, s2);
				s = s + "_" + s2;
			}

			util.AddToCStringArray(aList, s);

			Rs.MoveNext();
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s - %s", e->m_strError, sSQL);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {

		}
		return FALSE;
	}

	return TRUE;
}

BOOL CDatabaseManager::NPGetTimedMasterEditionOnLocation(CString sPublication, CTime tPubDate, CString sLocation, CString &sEdition, CString &sZone, int &nSequenceNumber, CString &sErrorMessage)
{
	CString sSQL, s;
	sErrorMessage = _T("");

	nSequenceNumber = 0;
	sEdition = "";
	sZone = "";

	sSQL.Format("SELECT DISTINCT Zone,SequenceNumber FROM ImportTable2 WHERE Publication='%s' AND PubDate='%s' AND Location='%s' AND Edition='1' ORDER BY SequenceNumber", sPublication, TranslateDate(tPubDate), sLocation);

	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	CRecordset Rs(m_pDB);

	try {
		if (!Rs.Open(CRecordset::snapshot, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("Query failed - %s", sSQL);
			return FALSE;
		}
		if (!Rs.IsEOF()) {
			Rs.GetFieldValue((short)0, s);
			sEdition = s;

			Rs.GetFieldValue((short)1, s);
			sZone = s;

			Rs.GetFieldValue((short)2, s);
			nSequenceNumber = atoi(s);
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s - %s", e->m_strError, sSQL);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {

		}
		return FALSE;
	}

	return TRUE;
}

BOOL CDatabaseManager::NPGetZoneMasterEditionOnLocation(CString sPublication, CTime tPubDate, CString sLocation, CString &sZone, CString &sErrorMessage)
{
	CString sSQL, s;
	sErrorMessage =  _T("");

	sZone = "";


	sSQL.Format("SELECT TOP 1 Zone FROM ImportTable2 WHERE Publication='%s' AND PubDate='%s' AND Location='%s' AND Edition='1' AND SequenceNumber=1", sPublication, TranslateDate(tPubDate), sLocation);

	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	CRecordset Rs(m_pDB);


	try {
		if (!Rs.Open(CRecordset::snapshot, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("Query failed - %s", sSQL);

			return FALSE;
		}
		if (!Rs.IsEOF()) {
			Rs.GetFieldValue((short)0, s);
			sZone = s;
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s - %s", e->m_strError, sSQL);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {

		}
		return FALSE;
	}

	return TRUE;
}

BOOL CDatabaseManager::NPGetZoneMaster(CString sPublication, CTime tPubDate, CString &sLocation, CString &sZone, CString &sErrorMessage)
{

	CString sSQL, s;
	sErrorMessage =  _T("");

	sZone = "";
	sLocation = "";

	if (sLocation != "")
		sSQL.Format("SELECT TOP 1 Zone,Location FROM ImportTable2 WHERE Publication='%s' AND PubDate='%s' AND Edition='1' AND SequenceNumber=1 AND sLocation='%s'", sPublication, TranslateDate(tPubDate), sLocation);
	else
		sSQL.Format("SELECT TOP 1 Zone,Location FROM ImportTable2 WHERE Publication='%s' AND PubDate='%s' AND Edition='1' AND SequenceNumber=1", sPublication, TranslateDate(tPubDate));
	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	CRecordset Rs(m_pDB);


	try {
		if (!Rs.Open(CRecordset::snapshot, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("Query failed - %s", sSQL);

			return FALSE;
		}
		if (!Rs.IsEOF()) {
			Rs.GetFieldValue((short)0, s);
			sZone = s;
			Rs.GetFieldValue((short)1, s);
			sLocation = s;
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s - %s", e->m_strError, sSQL);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {

		}
		return FALSE;
	}

	return TRUE;
}

BOOL CDatabaseManager::NPGetSlaveZonesOnLocation(CString sPublication, CTime tPubDate, CString sLocation, CString sMasterZone,
	CStringArray &sZones, CString &sErrorMessage)
{
	CUtils util;

	sZones.RemoveAll();

	CString sSQL, s;
	sErrorMessage =  _T("");

	sSQL.Format("SELECT DISTINCT Zone,SequenceNumber FROM ImportTable2 WHERE Publication='%s' AND PubDate='%s' AND Location='%s' AND Zone<>'%s' ORDER BY SequenceNumber",
		sPublication, TranslateDate(tPubDate), sLocation, sMasterZone);

	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	CRecordset Rs(m_pDB);

	try {
		if (!Rs.Open(CRecordset::snapshot, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("Query failed - %s", sSQL);
			return FALSE;
		}
		while (!Rs.IsEOF()) {

			Rs.GetFieldValue((short)1, s);
			util.AddToCStringArray(sZones, s);

			Rs.MoveNext();
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s - %s", e->m_strError, sSQL);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {

		}
		return FALSE;
	}

	return TRUE;
}

BOOL CDatabaseManager::NPGetSlaveZones(CString sPublication, CTime tPubDate, CString sMasterZone, CString sLocation,
	CStringArray &sZones, CString &sErrorMessage)
{
	CUtils util;

	sZones.RemoveAll();

	CString sSQL, s;
	sErrorMessage = _T("");

	if (sLocation == "")
		sSQL.Format("SELECT DISTINCT Zone,SequenceNumber FROM ImportTable2 WHERE Publication='%s' AND PubDate='%s' AND Zone<>'%s' ORDER BY SequenceNumber",
			sPublication, TranslateDate(tPubDate), sMasterZone);
	else
		sSQL.Format("SELECT DISTINCT Zone,SequenceNumber FROM ImportTable2 WHERE Publication='%s' AND PubDate='%s' AND Zone<>'%s' AND Location='%s' ORDER BY SequenceNumber",
			sPublication, TranslateDate(tPubDate), sMasterZone, sLocation);

	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	CRecordset Rs(m_pDB);

	try {
		if (!Rs.Open(CRecordset::snapshot, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("Query failed - %s", sSQL);

			return FALSE;
		}
		while (!Rs.IsEOF()) {

			Rs.GetFieldValue((short)1, s);
			util.AddToCStringArray(sZones, s);

			Rs.MoveNext();
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s - %s", e->m_strError, sSQL);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {

		}
		return FALSE;
	}

	return TRUE;
}

BOOL CDatabaseManager::NPGetTimedSlaveEditionsOnZoneLocation(CString sPublication, CTime tPubDate, CString sLocation, CString sMasterZone, CString sMasterEdition,
	CStringArray &sEditions, CString &sErrorMessage)
{
	CUtils util;

	sEditions.RemoveAll();

	CString sSQL, s;
	sErrorMessage =  _T("");

	if (sLocation != "")
		sSQL.Format("SELECT DISTINCT Edition FROM ImportTable2 WHERE Publication='%s' AND PubDate='%s' AND Location='%s' AND Zone='%s' AND Edition<>'%s' ORDER BY Edition",
			sPublication, TranslateDate(tPubDate), sLocation, sMasterZone, sMasterEdition);
	else
		sSQL.Format("SELECT DISTINCT Edition FROM ImportTable2 WHERE Publication='%s' AND PubDate='%s' AND Zone='%s' AND Edition<>'%s' ORDER BY Edition",
			sPublication, TranslateDate(tPubDate), sMasterZone, sMasterEdition);

	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	CRecordset Rs(m_pDB);

	try {
		if (!Rs.Open(CRecordset::snapshot, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("Query failed - %s", sSQL);

			return FALSE;
		}
		while (!Rs.IsEOF()) {
			Rs.GetFieldValue((short)0, s);
			util.AddToCStringArray(sEditions, s);

			Rs.MoveNext();
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s - %s", e->m_strError, sSQL);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {

		}
		return FALSE;
	}

	return TRUE;
}

BOOL CDatabaseManager::NPLoadImportTable(PlanDataEdition *editem, CString sPublication, CTime tPubDate, CString sLocation, CString sEdition, int &nPagesLoaded, CString &sErrorMessage)
{
	nPagesLoaded = 0;
	CString sSQL, s, sColors;
	sErrorMessage = _T("");
	CString sZone = "";
	if (sEdition != "") {
		int n = sEdition.Find("_");
		if (n != -1) {
			sZone = sEdition.Left(n);
			sEdition = sEdition.Mid(n + 1);
		}
	}

	if (editem == NULL)
		return FALSE;

	if (sLocation != "")
		sSQL.Format("SELECT DISTINCT Section,PageName, Pagination, PlanPageName, Comment, Colors,Version,SequenceNumber FROM ImportTable2 WHERE Publication='%s' AND PubDate='%s'  AND Edition='%s' AND Zone='%s' AND Location='%s'  ORDER BY SequenceNumber, Section, Pagination", sPublication, TranslateDate(tPubDate), sEdition, sZone, sLocation);
	else
		sSQL.Format("SELECT DISTINCT Section,PageName, Pagination, PlanPageName, Comment, Colors,Version,SequenceNumber FROM ImportTable2 WHERE Publication='%s' AND PubDate='%s'  AND Edition='%s' AND Zone='%s' ORDER BY SequenceNumber, Section, Pagination", sPublication, TranslateDate(tPubDate), sEdition, sZone);

	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	CRecordset Rs(m_pDB);
	try {
		if (!Rs.Open(CRecordset::snapshot, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("Query failed - %s", sSQL);

			return FALSE;
		}

		while (!Rs.IsEOF()) {
			CString sectionName;
			Rs.GetFieldValue((short)0, sectionName);


			PlanDataSection *secitem = editem->GetSectionObject(sectionName);
			if (secitem == NULL) {
				secitem = new PlanDataSection(sectionName);
				editem->sectionList->Add(*secitem);
				secitem->pagesInSection = 0;
			}


			//			Rs.GetFieldValue((short)0, s);
			//			secitem->pagesInSection = atoi(s);

			PlanDataPage *pPage = new PlanDataPage();
			pPage->approve = TRUE;
			pPage->pageType = PAGETYPE_NORMAL;
			pPage->hold = FALSE;
			pPage->priority = 50;
			pPage->version = 1;
			pPage->uniquePage = PAGEUNIQUE_UNIQUE;
			pPage->masterEdition = editem->editionName;
			pPage->pageID = "";
			pPage->masterPageID = "";


			Rs.GetFieldValue((short)1, s);
			pPage->pageName = s;
			Rs.GetFieldValue((short)2, s);
			pPage->pagination = atoi(s);

			pPage->pageIndex = pPage->pagination;
			Rs.GetFieldValue((short)3, s);
			pPage->fileName = s;
			Rs.GetFieldValue((short)4, s);
			pPage->comment = s;
			Rs.GetFieldValue((short)5, sColors);
			if (sColors == "")
				sColors = "CMYK";
			pPage->miscstring1 = sColors;

			Rs.GetFieldValue((short)6, s);
			if (atoi(s) > 0)
				pPage->version = atoi(s);

			// Add
			if (sColors == _T("PDF")) {
				PlanDataPageSeparation *sep = new PlanDataPageSeparation("PDF");
				pPage->colorList->Add(*sep);
			}
			else {
				for (int color = 0; color < sColors.GetLength(); color++) {
					PlanDataPageSeparation *sep = new PlanDataPageSeparation(sColors.Mid(color, 1));
					pPage->colorList->Add(*sep);
				}
			}

			secitem->pageList->Add(*pPage);
			nPagesLoaded++;
			secitem->pagesInSection++;

			Rs.MoveNext();
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", e->m_strError);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {

		}
		return FALSE;
	}

	return TRUE;

}


CString CDatabaseManager::LoadSpecificAlias(CString sType, CString sShortName, CString &sErrorMessage)
{

	CString sLongName = sShortName;

	if (sShortName == "")
		return "";


	sErrorMessage = _T("");
	if (InitDB(sErrorMessage) == FALSE)
		return sShortName;

	CRecordset Rs(m_pDB);

	CString sSQL;
	sSQL.Format("SELECT Longname FROM InputAliases WHERE Type='%s' AND ShortName='%s'", sType, sShortName);

	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", sSQL);
			return FALSE;
		}

		if (!Rs.IsEOF()) {
			Rs.GetFieldValue((short)0, sLongName);
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", e->m_strError);
		//		LogprintfDB("Error: %s   (%s)", e->m_strError, sSQL);

		e->Delete();
		Rs.Close();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return sLongName;
	}
	return sLongName;
}

extern PPINAMETRANSLATIONLIST g_ppitranslations;
BOOL CDatabaseManager::LoadPPITanslations(CString &sErrorMessage)
{

	g_ppitranslations.RemoveAll();

	sErrorMessage = _T("");
	if (InitDB(sErrorMessage) == FALSE)
		return FALSE;

	CRecordset Rs(m_pDB);

	CString sSQL;
	sSQL.Format("SELECT DISTINCT PPIProduct,PPIEdition,Publication FROM PPINames");

	try {
		if (!Rs.Open(CRecordset::snapshot | CRecordset::forwardOnly, sSQL, CRecordset::readOnly)) {
			sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", sSQL);
			return FALSE;
		}

		while (!Rs.IsEOF()) {
			PPINAMETRANSLATION *pItem = new PPINAMETRANSLATION();

			Rs.GetFieldValue((short)0, pItem->m_PPIProduct.MakeUpper());
			Rs.GetFieldValue((short)1, pItem->m_PPIEdition.MakeUpper());
			Rs.GetFieldValue((short)2, pItem->m_Publication.MakeUpper());
			g_ppitranslations.Add(*pItem);

			Rs.MoveNext();
		}
		Rs.Close();
	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Query failed - %s", e->m_strError);
		//		LogprintfDB("Error: %s   (%s)", e->m_strError, sSQL);

		e->Delete();
		Rs.Close();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}
	return TRUE;
}

BOOL CDatabaseManager::InsertFileSendRequest(CString sSourceFile, CString &sErrorMessage)
{
	DWORD ticks = GetTickCount();
	__int64 n64Guid = (__int64)ticks;

	CString sSQL;

	sErrorMessage = _T("");

	sSQL.Format("INSERT INTO [SendRequestQueue] (GUID, UseFTP, FTPServer,FTPUsername,FTPPassword,FTPFolder,FtpUsePASV,SMBFolder,EventTime,SourceFile,DeleteFileAfter) VALUES (%I64d, 1,'%s', '%s', '%s', '%s',1, '',GETDATE(), '%s',1)",
		n64Guid, g_prefs.m_secondcopyFtpServer, g_prefs.m_secondcopyFtpUsername, g_prefs.m_secondcopyFtpPassword, g_prefs.m_secondcopyFtpFolder, sSourceFile);


	if (InitDB(sErrorMessage) == FALSE) {
		return FALSE;
	}
	try {
		m_pDB->ExecuteSQL(sSQL);

	}
	catch (CDBException* e) {
		sErrorMessage.Format("ERROR (DATABASEMGR): Update failed - %s", (LPCSTR)e->m_strError);
		e->Delete();
		try {
			m_DBopen = FALSE;
			m_pDB->Close();
		}
		catch (CDBException* e) {
			// So what..! Go on!;
		}
		return FALSE;
	}


	return TRUE;
}
