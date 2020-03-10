#pragma once

#include "stdafx.h"
#include <afxmt.h>
#include "Defs.h"
#include "DatabaseManager.h"
#include "Utils.h"
// CPrefs command target


class CPrefs : public CObject
{
public:
	CPrefs();
	virtual ~CPrefs();

	CUtils util;
	TCHAR m_szDBServer[260];

	CString m_configfile;


	BOOL	m_publicationlock;
	BOOL	m_addtimestampindonefolder;

	BOOL	m_defaultapprovalnewpublication;

	// From ini-file
	BOOL	m_sortorder;
	CString		m_title;
	CString		m_Language;

	int			m_runningnumbers;

	int			m_QueryBackoffTime;
	int			m_nQueryRetries;
	// database connection for admin.

	CDatabaseManager m_DB;


	CTime	m_inputstoptime;
	BOOL	m_bypassping;
	BOOL    m_bypassreconnect;
	int		m_processtimeout;
	BOOL	m_getpubfromDB;

	BOOL	m_bSortOnCreateTime;


	CString m_DBserver;
	CString m_Database;
	CString m_DBuser;
	CString m_DBpassword;
	BOOL m_IntegratedSecurity;

	BOOL	m_PersistentConnection;
	int		m_databaselogintimeout;
	int		m_databasequerytimeout;

	CString m_inputfolder;
	CString m_inputfolder2;
	BOOL	m_inputfolderusecurrentuser;
	CString m_inputfolderusername;
	CString m_inputfolderpassword;

	CString m_logfolder;
	CString m_logfolder2;
	BOOL	m_saveonerror;
	CString m_errorfolder;
	CString m_errorfolder2;
	BOOL	m_saveafterdone;
	CString m_donefolder;
	CString m_donefolder2;

	int		m_polltime;
	int		m_stabletime;
	CString m_searchmask;
	CString m_schemafile;
	BOOL	m_validatexml;
	BOOL	m_overrulepublication;
	BOOL	m_overruleissue;
	BOOL	m_overruleedition;
	BOOL	m_overrulesection;
	BOOL	m_overrulepriority;
	BOOL	m_overrulerelease;
	BOOL	m_overruleapproval;
	CString m_publication;
	CString m_issue;
	CString m_edition;
	CString m_section;


	CString m_proofer;
	CString m_proofermasked;
	int		m_priority;
	int		m_hold;
	int		m_approve;

	BOOL	m_createcmyk;
	BOOL	m_combinesections;
	BOOL	m_keepcolors;
	BOOL	m_keepapproval;
	BOOL	m_keeppress;
	BOOL	m_keepunique;
	BOOL	m_rejectreimports;

	// From DB GeneralPreferences
	CString m_serverName;
	CString m_serverShare;
	CString m_storagePath; // CCFileHires
	CString m_previewPath;
	CString m_thumbnailPath;
	CString m_logFilePath;
	CString m_configPath;
	CString m_errorPath;
	int		m_MainLocationID;
	BOOL	m_copyToWeb;
	BOOL	m_copyWithFtp;
	CString m_webPath;
	CString m_webFTPserver;
	CString m_webFTPfolder;
	CString m_webFTPuser;
	CString m_webFTPpassword;
	UINT	m_webFTPport;
	int		m_autodeletedays;
	BOOL	m_enableautodelete;
	BOOL	m_logActions;
	int		m_mainlocationID;
	DWORD	m_minDiskSpaceMB;
	CString m_serverusername;
	CString m_serverpassword;
	BOOL	m_serverusecurrentuser;
	BOOL	m_useInches;			// Not used
	BOOL	m_doWarn;
	BOOL	m_logToFile;

	BOOL	m_onlyusepubalias;

	// Internal prefs

	CString m_apppath;

	IMPORTCONFIGURATIOnLIST m_ImportConfigurations;
	// Structures to hold config data
	ITEMLIST			m_ProofList;
	LOCATIONLIST		m_LocationList;
	CColorTableList		m_ColorList;
	PUBLICATIONLIST		m_PublicationList;
	ITEMLIST			m_SectionList;
	EDITIONLIST			m_EditionList;
	ITEMLIST			m_IssueList;
	ITEMLIST			m_MarkGroupList;
	DEVICELIST			m_DeviceList;
	TEMPLATELIST		m_TemplateList;
	PRESSLIST			m_PressList;

	PAGEFORMATLIST		m_PageFormatList;

	EXTRAEDITIONLIST	m_ExtraEditionList;
	EXTRAZONELIST		m_ExtraZoneList;

	BOOL	m_stampInputFiles;
	CString m_XMLLogFile;

	CString m_lockFile;
	HANDLE	m_lockHandle;


	int		m_cleaninterval;


	DWORD	m_maxlogfilesize;

	CString m_username;
	BOOL	m_sectionnameprefix;

	int		m_nProcessID_Import;

	CString m_locationname;
	int		m_resamplesleeptime;

	CString	m_deletefileswithextension;

	int		m_startmaximized;

	CString m_GrayColor;


	BOOL	m_debug;


	BOOL	m_alwaysblack;

	CString m_prooferrotated;

	int		m_customizedmarkgroup;


	BOOL	m_isadmin;
	BOOL	m_nologin;
	CTime	m_tLastLogin;


	BOOL	m_allowunknownnames;
	BOOL	m_generatedummycopies;
	BOOL	m_ignoreactivecopies;
	int		m_dateformat;

	BOOL	m_runpostprocedure;
	BOOL	m_runpostprocedure2;
	BOOL	m_runpostprocedure3;

	CString m_mailserver;
	CString m_mailusername;
	CString m_mailpassword;
	BOOL	m_mailuseSSL;
	CString m_mailfrom;
	CString m_mailto;
	CString m_mailcc;
	CString m_mailsubject;
	BOOL	m_mailonerror;
	int		m_mailport;
	int     m_mailtimeout;

	BOOL	m_planlocksystem;

	BOOL	m_allowmultipleinstances;
	int		m_instancenumber;

	BOOL	m_rejectappliedreimports;
	BOOL	m_rejectpolledreimports;
	BOOL	m_rejectimagedreimports;

	BOOL	m_usedbpublicationdefatuls;

	BOOL	m_mayignorebacksheet;

	BOOL	m_onlyuseactivecopies;

	BOOL	m_usepostcommand;
	int		m_postcommandtimeout;
	CString m_postcommand;

	BOOL	m_askreloaddonefiles;

	ALIASTABLELIST		m_AliasList;

	BOOL	m_allowautoapply;
	BOOL	m_autoapplyalways;
	BOOL	m_applyinserted;

	BOOL	m_pressspecificpages;
	CString m_unappliededitionlist;
	BOOL m_forceunappliededitions;
	BOOL	m_keepexistingsections;

	BOOL	m_assumecmykifnoseps;
	BOOL	m_pdfpages;
	BOOL	m_skipallcommon;
	BOOL	m_skipcommonplates;

	BOOL	m_recycledirtypages;
	BOOL	m_usechanneldefaults;

	BOOL	m_overrulepress;
	CString m_overruledpress;

	BOOL	m_runpostprocedureProduction;

	BOOL	m_useretryservice;

	BOOL	m_rejectallreimports;

	BOOL	m_setsectiontitles;
	BOOL	m_runpostprocedureProduction2;

	CStringArray m_PublicationsRequiringPostProc;

	BOOL	m_useInCodeDeepSearch;

	BOOL	m_allowemptysections;

	int		m_nInstanceNumber;

	CString m_copyfolder;
	CString m_copyfolder2;

	BOOL m_senderrormail;
	BOOL m_senderrormail2;
	CString m_emailreceivers;
	CString m_emailreceivers2;

	BOOL m_secondcopy;
	CString m_secondfolder;
	CString m_secondcopyFtpServer;
	CString m_secondcopyFtpUsername;
	CString m_secondcopyFtpPassword;
	CString m_secondcopyFtpFolder;
public:
	CString LookupNameFromAbbreviation(CString sType, CString sAbbr);
	CString LookupAbbreviationFromName(CString sType, CString sLong);
	BOOL	LoadAllPrefs(CDatabaseManager *pDB);
	BOOL	LoadPreferencesFromRegistry();
	void	LoadIniFile(CString sIniFile);

	void	ResetPublicationPressDefaults(int nPublicationID);

	LOCATIONSTRUCT *GetLocationStruct(int nID);
	int		GetLocationID(CString s);
	CString GetLocationName(int nID);
	CString GetLocationRemoteFolder(int nID);

	int		GetProofID(CString s);
	CString	GetProofName(int nID);

	int		GetColorID(CString s);
	CString	GetColorName(int nID);

	int		GetPageformatFromPublication(int nID);

	void	ResetPublicationPressDefaults(PUBLICATIONSTRUCT *pPublication);
	void	FlushPublicationName(CString sName);
	BOOL	IsBlackColor(int nID);
	BOOL	IsYellowColor(int nID);
	BOOL	IsCyanColor(int nID);
	BOOL	IsMagentaColor(int nID);


	int		GetPublicationID(CString s);
	PUBLICATIONSTRUCT *GetPublicationStruct(int nID);

	PRESSSTRUCT *GetPressStruct(int nID);
	int		GetSectionID(CString s);
	int		GetEditionID(CString s);
	int		GetIssueID(CString s);
	int		GetTemplateID(CString s);
	int		GetPressIDFromTemplateID(int nID, int *nPagesAcross, int *nPagesDown);
	int		GetTemplatePageRotation(int nID);
	int		GetPagesOnPlateFromTemplateID(int nID);
	int		GetPressIDFromTemplateID(int nID);
	int		GetDeviceID(CString s);
	int		GetTemplateIndex(CString s);
	int		GetPressID(CString s);
	int		GetPressLocationIDFromID(int nID);
	CString GetPublicationName(int nID);
	CString GetSectionName(int nID);
	CString GetEditionName(int nID);
	CString GetIssueName(int nID);

	CString GetTemplateName(int nID);
	CString GetDeviceName(int nID);
	CString GetPressName(int nID);

	int		GetPressIndex(int nID);


	int		GetMarkGroupID(CString sName);
	CString	GetMarkGroupName(int nID);

	int		AddNewPublication(CString s, CString sAbbr);
	int		AddNewSection(CString s, CString sAbbr);
	int		AddNewEdition(CString s, CString sAbbr);
	BOOL	ChangePublicationName(CString sLongName, CString sAbbr);

	PUBLICATIONPRESSDEFAULTSSTRUCT *GetPublicationPressDefaultStruct(int nPublicationID, int nPressID);
	PUBLICATIONPRESSDEFAULTSSTRUCT *GetPublicationDefaultStruct(int nPublicationID, int nPressID);


	CString GetSpecialRetryMask(int nPublicationID);
	BOOL  IsSpecialRetryProduct(int nPublicationID);




};


