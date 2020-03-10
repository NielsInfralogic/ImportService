#pragma once
#include "StdAfx.h"


/////////////////////////////////////////
// Process types
/////////////////////////////////////////

#define MAX_ERRMSG 2048


#define SERVICETYPE_PLANIMPORT 1
#define SERVICETYPE_FILEIMPORT 2
#define SERVICETYPE_PROCESSING 3
#define SERVICETYPE_EXPORT     4
#define SERVICETYPE_DATABASE   5
#define SERVICETYPE_FILESERVER 6
#define SERVICETYPE_MAINTENANCE 7


#define SERVICESTATUS_STOPPED 0
#define SERVICESTATUS_RUNNING 1
#define SERVICESTATUS_HASERROR 2


#define EVENTCODE_PLANERROR   996
#define EVENTCODE_PLANCREATED 990
#define EVENTCODE_PLANCHANGED 992
#define EVENTCODE_PLANDELETED 999



#define BYPASSPING_CURRENTUSER	1
#define BYPASSPING_ALWAYS		2

#define FOLDERSCANORDER_ALPHABETIC	0
#define FOLDERSCANORDER_FIFO		1
#define FOLDERSCANORDER_REGEXPRIO	2
#define FOLDERSCANORDER_REGEXRENAME	3

#define LOCKCHECKMODE_NONE	0
#define LOCKCHECKMODE_READ	1
#define LOCKCHECKMODE_READWRITE	2
#define LOCKCHECKMODE_RANGELOCK	3

#define FILETYPE_TIFF 0
#define FILETYPE_PDF 1
#define FILETYPE_OTHER 2

#define MAXFILES 1000

#define PLANTYPE_ERROR		-3
#define PLANTYPE_UNKNOWN	-2
#define PLANTYPE_INPROGRESS	-1
#define PLANTYPE_UNAPPLIED	0
#define PLANTYPE_APPLIED	1


#define PLANLOCK_LOCK		1
#define PLANLOCK_NOLOCK		0
#define PLANLOCK_UNLOCK		0
#define PLANLOCK_UNKNOWN	-1

#define	PARSEFILE_ERROR		0
#define	PARSEFILE_OK		1
#define	PARSEFILE_ABORT		2
#define PARSEFILE_OK_PAGESONLY 3
#define PARSEFILE_UNKNOWNPRESS 4

#define	IMPORTPAGES_DONE	1
#define	IMPORTPAGES_ERROR	0
#define	IMPORTPAGES_DBERROR	-1
#define	IMPORTPAGES_NOEDITION -2
#define IMPORTPAGES_REJECTEDREIMPORT	-3

#define IMPORTPAGES_REJECTEDAPPLIEDREIMPORT	-4
#define IMPORTPAGES_REJECTEDPOLLEDREIMPORT	-5
#define IMPORTPAGES_REJECTEDIMAGEDREIMPORT	-6

#define IMPORTPAGES_REJECTEDALLREIMPORT -7

#define EXISTINGPRODUCTION_NOPLAN			0
#define EXISTINGPRODUCTION_UNAPPLIEDPLAN	1
#define EXISTINGPRODUCTION_APPLIEDPLAN		2
#define	IMPORTPAGES_OK		1




typedef struct FILEINFOSTRUCTTAG {
	CString	sFileName;
	CString	sFolder;
	CTime	tJobTime;
	CTime tJobStartTime;
	CTime	tWriteTime;
	DWORD	nFileSize;
} FILEINFOSTRUCT;

#define MAXISSUES	1
#ifdef SMALLBUILD
#define MAXFILES 10000
#define MAXEDITIONS	16
#define MAXSECTIONS 16
#define MAXPAGES	300
#define MAXSHEETS	200
#define MAXCOLORS	5
#define MAXMARKGROUPSPERTEMPLATE 50
#define MAXTOWERS 16
#define MAXPRESSES 16
#define MAXPAGENAME 48
#define MAXPAGEID	16
#define MAXCOMMENT  64
#define MAXFORMIDSTRING 64
#define MAXSORTINGPOS	64
#else
#define MAXFILES 1000
#define MAXEDITIONS	200
#define MAXZONES		20
#define MAXSECTIONS 8
#define MAXPAGES	512
#define MAXSHEETS	256
#define MAXCOLORS	6
#define MAXMARKGROUPSPERTEMPLATE 100
#define MAXTOWERS 24
#define MAXPRESSES 32
#define MAXPAGENAME 48
#define MAXPAGEID	16
#define MAXCOMMENT  64
#define MAXFORMIDSTRING 64
#define MAXSORTINGPOS	64
#endif

#define MAXPRESSSTRING 8
#define MAXPRESSTOWERSTRING 24
#define MAXPAGEPOSITIONS 16
#define MAXGROUPNAMES	64
#define MAXPAGEIDLEN	8
#define MAXMARKGROUPS	32
#define MAXPOSTALURL	128

#define STATUSNUMBER_MISSING			0
#define STATUSNUMBER_POLLING			5
#define STATUSNUMBER_POLLINGERROR		6
#define STATUSNUMBER_POLLED				10
#define STATUSNUMBER_RESAMPLING			15
#define STATUSNUMBER_RESAMPLINGERROR	16
#define STATUSNUMBER_READY				20
#define STATUSNUMBER_TRANSMITTING		25
#define STATUSNUMBER_TRANSMISSIONERROR	26
#define STATUSNUMBER_TRANSMITTED		30
#define STATUSNUMBER_ASSEMBLING			35
#define STATUSNUMBER_ASSEMBLINGERROR	36
#define STATUSNUMBER_ASSEMBLED			40
#define STATUSNUMBER_IMAGING			45
#define STATUSNUMBER_IMAGINGERROR		46
#define STATUSNUMBER_IMAGED				50
#define STATUSNUMBER_VERIFYING			55
#define STATUSNUMBER_VERIFYERROR		56
#define STATUSNUMBER_VERIFIED			60
#define STATUSNUMBER_KILLED				99

#define PROOFSTATUSNUMBER_NOTPROOFED	0
#define PROOFSTATUSNUMBER_PROOFING		5
#define PROOFSTATUSNUMBER_PROOFERROR	6
#define PROOFSTATUSNUMBER_PROOFED		10

#define APPROVAL_AUTO					-1
#define APPROVAL_NOTAPPROVED			0
#define APPROVAL_APPROVED				1
#define APPROVAL_REJECTED				2

#define HOLDSTATE_RELEASED				0
#define HOLDSTATE_HELD					1

#define PAGETYPE_NORMAL					0
#define PAGETYPE_PANORAMA				1
#define PAGETYPE_ANTIPANORAMA			2
#define PAGETYPE_DUMMY					3

#define COLORID_DUMMY					6



#define IMPORTTABLEQUERYMODE_EDITIONS				0
#define IMPORTTABLEQUERYMODE_LOCATIONS				1
#define IMPORTTABLEQUERYMODE_EDITIONS_ON_LOCATION	2
#define IMPORTTABLEQUERYMODE_LOCATIONS_FOR_EDITION	3
#define IMPORTTABLEQUERYMODE_EDITIONIDLOCATIONS		4
#define IMPORTTABLEQUERYMODE_EDITIONIDLOCATIONS_FOR_EDITION 5
#define IMPORTTABLEQUERYMODE_SECTIONS 6


#define PAGETYPE_NORMAL			0
#define PAGETYPE_PANORAMA		1
#define PAGETYPE_ANTIPANORAMA	2
#define PAGETYPE_DUMMY			3

#define PAGEUNIQUE_UNIQUE		1
#define PAGEUNIQUE_COMMON		0
#define PAGEUNIQUE_FORCED		2

/////////////////////////////////////////
// Internal definitions for user login
/////////////////////////////////////////

#define USER_UNKNOWNUSER				1
#define USER_PASSWORDWRONG				2
#define USER_USEROK						3
#define USER_ACCOUNTDISABLED			4

#define UNIQUEPAGE_UNIQUE		1
#define UNIQUEPAGE_COMMON		0		
#define UNIQUEPAGE_FORCE		2

#define EVENTCODE_PLANCREATE	990
#define EVENTCODE_PLANAPPLIED	991
#define EVENTCODE_PLANEDITED	992
#define EVENTCODE_PLANERROR		996
#define EVENTCODE_PLANDELETE	999

#define SNAP_UL		0
#define SNAP_UR		1
#define SNAP_LL		2
#define SNAP_LR		3
#define SNAP_UC		4
#define SNAP_LC		5
#define SNAP_CL		6
#define SNAP_CR		7
#define SNAP_CENTER	8

#define IMPORTCONFIGURATIONTYPE_STANDARD 0
#define IMPORTCONFIGURATIONTYPE_NEWSPILOT 1
#define IMPORTCONFIGURATIONTYPE_PPI		 2


typedef struct {
	int nImportID;
	CString sImportName;
	int nType;
	CString sInputFolder;
	CString sDoneFolder;
	CString sErrorFolder;
	CString sLogFolder;
	CString sConfigFile;
	CString sConfigFile2;
	CString sCopyFolder;
	BOOL bSendErrorEmail;
	CString sEmailReceivers;
} IMPORTCONFIGURATION;

typedef CArray <IMPORTCONFIGURATION, IMPORTCONFIGURATION&> IMPORTCONFIGURATIOnLIST;


typedef struct ITEMSTRUCTTAG {
	int m_ID;
	CString m_name;
} ITEMSTRUCT;

typedef CArray <ITEMSTRUCT, ITEMSTRUCT&> ITEMLIST;

typedef CArray <LPVOID, LPVOID&> LPVOIDLIST;
typedef CArray <int, int> intArray;

typedef struct  {
	int		m_ID;
	CString	m_name;
	BOOL	m_iscommon;
	int 	m_subofeditionID;
	int		m_level;
} EDITIONSTRUCT;
typedef CArray <EDITIONSTRUCT, EDITIONSTRUCT&> EDITIONLIST;

typedef struct DEVICESTRUCTTAG {
	int		m_ID;
	CString	m_name;
	BOOL	m_locationID;
} DEVICESTRUCT;
typedef CArray <DEVICESTRUCT, DEVICESTRUCT&> DEVICELIST;

typedef struct EXTRAEDITIONTAG {
	CString		m_Publication;
	int		m_mineditions;
} EXTRAEDITION;
typedef CArray <EXTRAEDITION, EXTRAEDITION&> EXTRAEDITIONLIST;

typedef struct EXTRAZONETAG {
	CString		m_Publication;
	CString     m_Press;
	CString     m_zonetoadd;
} EXTRAZONE;
typedef CArray <EXTRAZONE, EXTRAZONE&> EXTRAZONELIST;

/*
typedef struct PRESSSTRUCTTAG {
	int		m_ID;
	CString	m_name;
	BOOL	m_locationID;
} PRESSSTRUCT;
typedef CArray <PRESSSTRUCT,PRESSSTRUCT&> PRESSLIST;
*/

class ALIASTABLE {
public:
	ALIASTABLE() { sType = _T("Color"); sLongName = _T(""); sShortName = _T(""); };
	CString	sType;
	CString sLongName;
	CString sShortName;
};
typedef CArray <ALIASTABLE, ALIASTABLE&> ALIASTABLELIST;

class PRESSSTRUCT {
public:
	PRESSSTRUCT() {
		m_ID = 0;
		m_name = _T("");
		m_locationID = 0;
		m_numberoftowers = 0;
		for (int i = 0; i < MAXTOWERS; i++)
			m_towerlist[i] = _T("");
	}


	int		m_ID;
	CString	m_name;
	int		m_locationID;
	int		m_numberoftowers;
	CString m_towerlist[MAXTOWERS];
};
typedef CArray <PRESSSTRUCT, PRESSSTRUCT&> PRESSLIST;


class LOCATIONSTRUCT {
public:
	LOCATIONSTRUCT() {
		m_locationID = 0;
		m_locationname = _T("");
		m_remotefolder = _T("");
		m_backupremotefolder = _T("");
		m_useftp = FALSE;
		m_ftpserver = _T("");
		m_ftpusername = _T("");
		m_ftppassword = _T("");
		m_ftpfolder = _T("");
		m_ftpport = 21;
	};

	int		m_locationID;
	CString m_locationname;
	CString m_remotefolder;
	CString m_backupremotefolder;
	BOOL	m_useftp;
	CString m_ftpserver;
	CString m_ftpusername;
	CString m_ftppassword;
	CString m_ftpfolder;
	DWORD	m_ftpport;
};

typedef CArray <LOCATIONSTRUCT, LOCATIONSTRUCT&> LOCATIONLIST;


#define TRIMMODE_NONE	0
#define TRIMMODE_TRIM  4
#define TRIMMODE_SCALE 8


class PAGEFORMATSTRUCT {
public:
	PAGEFORMATSTRUCT() {
		m_name = _T(""); m_ID = 0; m_formatw = 0.0; m_formath = 0.0; m_bleed = 0.0; m_trimmode = TRIMMODE_NONE; m_snapmodeeven = 7; m_snapmodeodd = 6;
	};

	int	m_ID;
	CString m_name;
	double m_formatw;
	double m_formath;
	double m_bleed;
	int m_snapmodeeven;
	int m_snapmodeodd;
	int m_trimmode;
};

typedef CArray <PAGEFORMATSTRUCT, PAGEFORMATSTRUCT&> PAGEFORMATLIST;

typedef struct PUBLICATIONPRESSDEFAULTSSTRUCTTAG {
	int			m_publicationID;
	int			m_PressID;
	int			m_DefaultTemplateID;
	CString		m_DefaultPressTowerName;
	int			m_DefaultCopies;
	int			m_DefaultNumberOfMarkGroups;
	CString		m_DefaultMarkGroupList[MAXMARKGROUPSPERTEMPLATE];
	int			m_DefaultPriority;
	CString		m_RipSetup;
	CString		m_DefaultStackPosition;
	BOOL		m_defaultpress;
	BOOL		m_allowautoplanning;
	int			m_DefaultFlatProofTemplateID;
	int			m_deviceID;
	BOOL		m_hold;
	CString		m_txname;
	BOOL		m_insertedsections;
	BOOL		m_pressspecificpages;
	int			m_paginationmode;
	BOOL		m_separateruns;
	BOOL		m_pressnotused;
} PUBLICATIONPRESSDEFAULTSSTRUCT;

typedef struct PUBLICATIONSTRUCTTAG {
	int			m_ID;
	CString		m_name;
	int			m_PageFormatID;
	int			m_TrimToFormat;
	double		m_LatestHour;
	int			m_ProofID;
	int			m_HardProofID;
	int			m_Approve;
	CString		m_uploadfolder;
	CString		m_annumtext;
	int			m_releasedays;
	int			m_releasetimehour;
	int			m_releasetimeminute;
	int			m_inputpriority;
	BOOL		m_unappliedplanrequired;
	BOOL		m_autoapprove;
	BOOL		m_rejectsubeditionpages;
	CString     m_alias;
	PUBLICATIONPRESSDEFAULTSSTRUCT m_PressDefaults[MAXPRESSES];
} PUBLICATIONSTRUCT;

typedef CArray <PUBLICATIONSTRUCT, PUBLICATIONSTRUCT&> PUBLICATIONLIST;



class TEMPLATESTRUCT {
public:
	TEMPLATESTRUCT() {
		m_ID = 0;
		m_name = _T("");
		m_pressID = 0;
		m_pagesacross = 1;
		m_pagesdown = 1;
		m_incomingrotationeven = 0;
		m_incomingrotationodd = 0;
		m_pageformatID = 0;
		for (int i = 0; i < MAXMARKGROUPS; i++)
			m_markgroups[i] = "";
	};

	int		m_ID;
	CString m_name;
	int		m_pressID;
	int		m_pagesacross;
	int		m_pagesdown;
	int		m_incomingrotationeven;
	int		m_incomingrotationodd;
	int		m_pageformatID;
	CString m_markgroups[MAXMARKGROUPS];
};

typedef CArray <TEMPLATESTRUCT, TEMPLATESTRUCT&> TEMPLATELIST;


class CColorTableItem {
public:
	CColorTableItem() {
		m_ID = 0;
		m_name = _T("");
		m_c = 0;
		m_m = 0;
		m_y = 0;
		m_k = 0;
		m_colororder = 0;
	};

	int			m_ID;
	CString		m_name;
	UINT		m_c;
	UINT		m_m;
	UINT		m_y;
	UINT		m_k;
	int			m_colororder;
};

typedef CArray <CColorTableItem, CColorTableItem&> CColorTableList;


class CPageTableEntry {
public:
	CPageTableEntry() { Init(); }

	void Init()
	{
		m_copyseparationset = 1;
		m_separationset = 101;
		m_separation = 10101;
		m_copyflatseparationset = 1;
		m_flatseparationset = 101;
		m_flatseparation = 10101;
		m_status = STATUSNUMBER_MISSING;
		m_externalstatus = 0;

		m_weekreference = 0;
		m_publicationID = 0;
		m_sectionID = 0;
		m_editionID = 0;
		m_issueID = 0;
		m_pubdate = CTime::GetCurrentTime();	// Default to neext day
		m_pubdate += CTimeSpan(1, 0, 0, 0);
		_tcscpy(m_pagename, _T("1"));
		m_colorID = 1;
		m_templateID = 0;
		m_proofID = 0;
		m_deviceID = 0;
		m_version = 0;
		m_layer = 1;
		m_copynumber = 1;
		m_pagination = 0;
		m_approved = APPROVAL_NOTAPPROVED;
		m_hold = FALSE;
		m_active = TRUE;
		m_priority = 50;
		m_pageposition = 1;
		m_pagetype = PAGETYPE_NORMAL;
		m_pagesonplate = 1;
		m_sheetnumber = 1;
		m_sheetside = 0;
		m_pressID = 0;
		m_presssectionnumber = 1;
		_tcscpy(m_sortingposition, _T(""));
		_tcscpy(m_presstower, _T("1"));
		_tcscpy(m_presszone, _T("1"));
		_tcscpy(m_presscylinder, _T("1"));
		_tcscpy(m_formID, _T(""));
		_tcscpy(m_plateID, _T(""));

		m_productionID = 0;
		m_pressrunID = 0;
		m_proofstatus = 0;
		m_inkstatus = 0;
		_tcscpy(m_planpagename, _T(""));
		m_issuesequencenumber = 1;
		m_mastercopyseparationset = 1;
		m_uniquepage = TRUE;
		m_locationID = 0;
		m_flatproofID = 0;
		m_flatproofstatus = 0;
		m_creep = 0.0;
		_tcscpy(m_markgroups, _T(""));
		m_pageindex = 1;
		m_gutterimage = FALSE;
		m_outputversion = 0;
		m_hardproofID = 0;
		m_hardproofstatus = 0;
		_tcscpy(m_fileserver, _T(""));
		m_dirty = FALSE;
		_tcscpy(m_comment, _T(""));
		m_deadline = CTime(1975, 1, 1, 0, 0, 0);
		_tcscpy(m_pagepositions, _T("1"));
		m_mastereditionID = 0;
		m_colorIndex = 1;
		m_pagecountchange = TRUE;
		m_customerID = 0;
		m_autorelease = FALSE;
	};

	int		m_copyseparationset;
	int		m_separationset;
	int		m_separation;
	int		m_copyflatseparationset;
	int		m_flatseparationset;
	int		m_flatseparation;

	int		m_status;
	int		m_externalstatus;
	int		m_publicationID;
	int		m_sectionID;
	int		m_editionID;
	int		m_issueID;
	CTime	m_pubdate;

	TCHAR	m_pagename[50];	// page number!
	int		m_colorID;
	int		m_templateID;
	int		m_proofID;
	int		m_deviceID;
	int		m_version;
	int		m_layer;
	int		m_copynumber;
	int		m_pagination;
	int		m_approved;
	BOOL	m_hold;
	BOOL	m_active;
	int		m_priority;
	int		m_pageposition;	// NOT USED!
	int		m_pagetype;
	int		m_pagesonplate;
	int		m_sheetnumber;
	int		m_sheetside;
	int		m_pressID;
	int		m_presssectionnumber;
	TCHAR	m_sortingposition[MAXSORTINGPOS];
	TCHAR	m_presstower[24];
	TCHAR	m_presszone[24];
	TCHAR	m_presscylinder[24];
	TCHAR	m_presshighlow[16];
	int		m_productionID;
	int		m_pressrunID;
	int		m_proofstatus;
	int		m_inkstatus;
	TCHAR	m_planpagename[128];
	int		m_issuesequencenumber;
	int		m_mastercopyseparationset;
	int		m_uniquepage;
	int		m_locationID;
	int		m_flatproofID;
	int		m_flatproofstatus;
	double	m_creep;
	TCHAR	m_markgroups[200];
	int		m_pageindex;
	int		m_gutterimage;
	int		m_outputversion;
	int		m_hardproofID;
	int		m_hardproofstatus;
	TCHAR	m_fileserver[50];
	int		m_dirty;

	TCHAR	m_comment[100];
	CTime	m_deadline;
	TCHAR	m_pagepositions[50];

	int		m_mastereditionID;
	int		m_colorIndex;
	BOOL	m_pagecountchange;
	int		m_weekreference;

	TCHAR	m_formID[64];
	TCHAR	m_plateID[64];

	int		m_customerID;
	//CTime	m_presstime;
	SYSTEMTIME m_presstime;
	BOOL	m_autorelease;

	int		m_miscint2;
	CString	m_miscstring2;
	CString m_miscstring3;
};

typedef CArray <CPageTableEntry, CPageTableEntry&> CPageTableList;



class PLANSTRUCT {
public:
	PLANSTRUCT() { };

	void Init() {
		m_runpostproc = FALSE;
		m_version = 0;
		m_press = _T("");
		m_rippress = _T("");
		m_pressID = 0;				// Found via template
		m_rippressID = 0;
		m_updatetime = _T("");
		m_sender = _T("");
		m_publicationID = 0;
		m_weekreference = 0;
		m_pubdate = CTime(1975, 1, 1, 0, 0, 0);
		m_numberofissues = 0;
		m_productionID = 0;
		m_NumberOfPresses = 0;
		m_planName = "";
		m_planID = "";
		m_customer = "";
		m_customeralias = "";

		m_usemask = FALSE;
		m_masksizex = 0.0;
		m_masksizey = 0.0;
		m_masksnapeven = SNAP_CR;
		m_masksnapodd = SNAP_CL;
		m_maskbleed = 0.0;
		m_maskshowspread = TRUE;

		for (int i = 0; i < MAXPRESSES; i++)
			m_productionIDperpress[i] = 0;
		for (int i = 0; i < MAXISSUES; i++) {
			m_numberofeditions = 0;
			m_issueIDlist[i] = 0;
			for (int j = 0; j < MAXEDITIONS; j++) {
				m_numberofpresses[i][j] = 1;
				m_numberofsections[i][j] = 0;

				m_editionIDlist[i][j] = 0;
				m_numberofsheets[i][j] = 0;
				m_pressIDlist[i][j] = 0;
				m_allnonuniqueepages[i][j] = FALSE;
				m_numberofpapercopiesedition[i][j] = 0;
				m_editionnochange[i][j] = FALSE;
				m_editionistimed[i][j] = FALSE;
				m_editiontimedfromID[i][j] = 0;
				m_editiontimedtoID[i][j] = 0;
				m_editionordernumber[i][j] = 1;
				m_editionmasterzoneeditionID[i][j] = 0;
				strcpy(m_editioncomment[i][j], "");

				for (int k = 0; k < MAXPRESSES; k++) {
					m_pressesinedition[i][j][k] = 0;
					m_numberofpapercopies[i][j][k] = 0;
					m_numberofcopies[i][j][k] = 1;
					m_autorelease[i][j][k] = FALSE;
					strcpy(m_postalurl[i][j][k], "");
					strcpy(m_presstime[i][j][k], "");
				}

				for (int k = 0; k < MAXSECTIONS; k++) {
					m_numberofpages[i][j][k] = 0;
					m_sectionIDlist[i][j][k] = 0;
					strcpy(m_sectionTitlelist[i][j][k], "");
					strcpy(m_sectionPageFormat[i][j][k], "");
					m_sectionallcommon[i][j][k] = FALSE;
					for (int m = 0; m < MAXPAGES; m++) {
						m_numberofpagecolors[i][j][k][m] = 0;
						//	m_pagemasteredition[i][j][k][m] = 0;
						strcpy(m_pagenames[i][j][k][m], "");
						strcpy(m_pagefilenames[i][j][k][m], "");
						strcpy(m_pagemiscstring1[i][j][k][m], "");
						strcpy(m_pagemiscstring2[i][j][k][m], "");


						strcpy(m_pageIDs[i][j][k][m], "");
						strcpy(m_pagemasterIDs[i][j][k][m], "");
						m_pagetypes[i][j][k][m] = PAGETYPE_NORMAL;
						m_pageindex[i][j][k][m] = 0;
						m_pagepagination[i][j][k][m] = 1;
						m_pageposition[i][j][k][m] = 1;
						strcpy(m_pagecomments[i][j][k][m], "");
						m_pageversion[i][j][k][m] = 0;
						m_pageapproved[i][j][k][m] = 0;
						m_pagehold[i][j][k][m] = TRUE;
						m_pagepriority[i][j][k][m] = 50;
						m_pageunique[i][j][k][m] = UNIQUEPAGE_UNIQUE;
						m_pagemiscint[i][j][k][m] = 0;

						for (int n = 0; n < MAXCOLORS; n++) {
							m_pagecolorIDlist[i][j][k][m][n] = 0;
							m_pagecolorActivelist[i][j][k][m][n] = 0;
						}
					}
				}
				for (int k = 0; k < MAXSHEETS; k++) {
					m_sheettemplateID[i][j][k] = 0;
					m_sheetbacktemplateID[i][j][k] = 0;
					strcpy(m_sheetmarkgroup[i][j][k], "");
					m_sheetpagesonplate[i][j][k] = 1;
					m_sheetignoreback[i][j][k] = FALSE;
					m_pressSectionNumber[i][j][k] = 0;

					for (int m = 0; m < 2; m++) {
						m_sheetdefaultsectionID[i][j][k][m] = 0;
						m_sheetmaxcolors[i][j][k][m] = 0;
						strcpy(m_sheetsortingposition[i][j][k][m], "");
						m_sheetactivecopies[i][j][k][m] = 1;

						strcpy(m_sheetpresstower[i][j][k][m], "");
						strcpy(m_sheetpresszone[i][j][k][m], "");
						strcpy(m_sheetpresshighlow[i][j][k][m], "");
						m_sheetallcommonpages[i][j][k][m] = FALSE;
						for (int n = 0; n < MAXPAGEPOSITIONS; n++) {
							m_sheetpageposition[i][j][k][m][n] = 1;
							strcpy(m_sheetpageID[i][j][k][m][n], "");
							strcpy(m_sheetmasterpageID[i][j][k][m][n], "");

							strcpy(m_oldsheetpagename[i][j][k][m][n], "");
							m_oldsheetsectionID[i][j][k][m][n] = 0;

						}
						for (int n = 0; n < MAXCOLORS; n++) {
							strcpy(m_sheetpresscylinder[i][j][k][m][n], "");
							strcpy(m_sheetformID[i][j][k][m][n], "");
							strcpy(m_sheetplateID[i][j][k][m][n], "");
							strcpy(m_sheetsortingpositioncolorspecific[i][j][k][m][n], "");
						}
					}
				}
			}
		}
		m_pagespersheetside = 1;
	};

	int		m_version;
	BOOL	m_runpostproc;

	CString m_press;
	CString m_rippress;
	CString	m_updatetime;
	CString m_sender;
	int		m_publicationID;
	CTime   m_pubdate;
	int		m_productionID;
	int		m_pressID;
	int     m_rippressID;
	CString m_planName;
	CString m_planID;

	int		m_NumberOfPresses;
	int		m_productionIDperpress[MAXPRESSES];
	int		m_weekreference;

	CString	m_customer;
	CString	m_customeralias;

	BOOL	m_usemask;
	double	m_masksizex;
	double	m_masksizey;
	int		m_masksnapeven;
	int		m_masksnapodd;
	double  m_maskbleed;
	double  m_maskshowspread;

	//#pragma pack(1)
	BYTE		m_numberofissues;
	BYTE		m_numberofeditions;
	BYTE		m_numberofpresses[MAXISSUES][MAXEDITIONS];
	DWORD		m_numberofpapercopiesedition[MAXISSUES][MAXEDITIONS];
	BYTE		m_pressesinedition[MAXISSUES][MAXEDITIONS][MAXPRESSES];
	BYTE		m_numberofcopies[MAXISSUES][MAXEDITIONS][MAXPRESSES];
	DWORD		m_numberofpapercopies[MAXISSUES][MAXEDITIONS][MAXPRESSES];
	DWORD		m_autorelease[MAXISSUES][MAXEDITIONS][MAXPRESSES];
	TCHAR		m_postalurl[MAXISSUES][MAXEDITIONS][MAXPRESSES][MAXPOSTALURL];
	TCHAR		m_presstime[MAXISSUES][MAXEDITIONS][MAXPRESSES][32];

	BYTE		m_numberofsections[MAXISSUES][MAXEDITIONS];
	WORD		m_numberofpages[MAXISSUES][MAXEDITIONS][MAXSECTIONS];
	BYTE		m_numberofpagecolors[MAXISSUES][MAXEDITIONS][MAXSECTIONS][MAXPAGES];
	BYTE		m_pagespersheetside;

	WORD		m_issueIDlist[MAXISSUES];
	BYTE		m_editionnochange[MAXISSUES][MAXEDITIONS];
	WORD		m_editionIDlist[MAXISSUES][MAXEDITIONS];
	BYTE		m_editionistimed[MAXISSUES][MAXEDITIONS];
	WORD		m_editiontimedfromID[MAXISSUES][MAXEDITIONS];
	WORD		m_editiontimedtoID[MAXISSUES][MAXEDITIONS];
	WORD		m_editionmasterzoneeditionID[MAXISSUES][MAXEDITIONS];
	BYTE		m_editionordernumber[MAXISSUES][MAXEDITIONS];
	TCHAR		m_editioncomment[MAXISSUES][MAXEDITIONS][MAXCOMMENT];

	WORD		m_pressIDlist[MAXISSUES][MAXEDITIONS];

	BYTE		m_allnonuniqueepages[MAXISSUES][MAXEDITIONS];

	WORD		m_sectionIDlist[MAXISSUES][MAXEDITIONS][MAXSECTIONS];
	TCHAR		m_sectionTitlelist[MAXISSUES][MAXEDITIONS][MAXSECTIONS][MAXCOMMENT];
	BYTE		m_sectionallcommon[MAXISSUES][MAXEDITIONS][MAXSECTIONS];
	TCHAR		m_sectionPageFormat[MAXISSUES][MAXEDITIONS][MAXSECTIONS][MAXCOMMENT];

	//WORD		m_pagemasteredition[MAXISSUES][MAXEDITIONS][MAXSECTIONS][MAXPAGES];
	BYTE		m_pagetypes[MAXISSUES][MAXEDITIONS][MAXSECTIONS][MAXPAGES];
	WORD		m_pagepagination[MAXISSUES][MAXEDITIONS][MAXSECTIONS][MAXPAGES];
	BYTE		m_pageposition[MAXISSUES][MAXEDITIONS][MAXSECTIONS][MAXPAGES];
	TCHAR		m_pagenames[MAXISSUES][MAXEDITIONS][MAXSECTIONS][MAXPAGES][MAXPAGENAME];
	TCHAR		m_pagefilenames[MAXISSUES][MAXEDITIONS][MAXSECTIONS][MAXPAGES][MAXPAGENAME];
	TCHAR		m_pageIDs[MAXISSUES][MAXEDITIONS][MAXSECTIONS][MAXPAGES][MAXPAGEIDLEN];
	TCHAR		m_pagemasterIDs[MAXISSUES][MAXEDITIONS][MAXSECTIONS][MAXPAGES][MAXPAGEIDLEN];
	TCHAR		m_pagecomments[MAXISSUES][MAXEDITIONS][MAXSECTIONS][MAXPAGES][MAXCOMMENT];
	WORD		m_pageindex[MAXISSUES][MAXEDITIONS][MAXSECTIONS][MAXPAGES];

	WORD		m_pagemiscint[MAXISSUES][MAXEDITIONS][MAXSECTIONS][MAXPAGES];
	TCHAR		m_pagemiscstring1[MAXISSUES][MAXEDITIONS][MAXSECTIONS][MAXPAGES][MAXCOMMENT];
	TCHAR		m_pagemiscstring2[MAXISSUES][MAXEDITIONS][MAXSECTIONS][MAXPAGES][MAXCOMMENT];

	BYTE		m_pagecolorIDlist[MAXISSUES][MAXEDITIONS][MAXSECTIONS][MAXPAGES][MAXCOLORS];
	BYTE		m_pagecolorActivelist[MAXISSUES][MAXEDITIONS][MAXSECTIONS][MAXPAGES][MAXCOLORS];
	BYTE		m_pageversion[MAXISSUES][MAXEDITIONS][MAXSECTIONS][MAXPAGES];
	char		m_pageapproved[MAXISSUES][MAXEDITIONS][MAXSECTIONS][MAXPAGES];
	BYTE		m_pagehold[MAXISSUES][MAXEDITIONS][MAXSECTIONS][MAXPAGES];
	BYTE		m_pagepriority[MAXISSUES][MAXEDITIONS][MAXSECTIONS][MAXPAGES];
	BYTE		m_pageunique[MAXISSUES][MAXEDITIONS][MAXSECTIONS][MAXPAGES];

	WORD		m_numberofsheets[MAXISSUES][MAXEDITIONS];
	BYTE		m_sheettemplateID[MAXISSUES][MAXEDITIONS][MAXSHEETS];
	BYTE		m_sheetbacktemplateID[MAXISSUES][MAXEDITIONS][MAXSHEETS];

	TCHAR		m_sheetmarkgroup[MAXISSUES][MAXEDITIONS][MAXSHEETS][MAXGROUPNAMES];
	BYTE		m_sheetpagesonplate[MAXISSUES][MAXEDITIONS][MAXSHEETS];
	BYTE		m_sheetignoreback[MAXISSUES][MAXEDITIONS][MAXSHEETS];
	BYTE		m_pressSectionNumber[MAXISSUES][MAXEDITIONS][MAXSHEETS];


	TCHAR		m_sheetsortingpositioncolorspecific[MAXISSUES][MAXEDITIONS][MAXSHEETS][2][MAXCOLORS][MAXFORMIDSTRING];
	TCHAR		m_sheetpresscylinder[MAXISSUES][MAXEDITIONS][MAXSHEETS][2][MAXCOLORS][MAXPRESSSTRING];
	TCHAR		m_sheetformID[MAXISSUES][MAXEDITIONS][MAXSHEETS][2][MAXCOLORS][MAXFORMIDSTRING];
	TCHAR		m_sheetplateID[MAXISSUES][MAXEDITIONS][MAXSHEETS][2][MAXCOLORS][MAXFORMIDSTRING];
	TCHAR		m_sheetpresstower[MAXISSUES][MAXEDITIONS][MAXSHEETS][2][MAXPRESSTOWERSTRING];
	TCHAR		m_sheetpresszone[MAXISSUES][MAXEDITIONS][MAXSHEETS][2][MAXPRESSSTRING];
	TCHAR		m_sheetpresshighlow[MAXISSUES][MAXEDITIONS][MAXSHEETS][2][MAXPRESSSTRING];

	TCHAR		m_sheetsortingposition[MAXISSUES][MAXEDITIONS][MAXSHEETS][2][MAXSORTINGPOS];
	BYTE		m_sheetpageposition[MAXISSUES][MAXEDITIONS][MAXSHEETS][2][MAXPAGEPOSITIONS];
	BYTE		m_sheetmaxcolors[MAXISSUES][MAXEDITIONS][MAXSHEETS][2];
	BYTE		m_sheetallcommonpages[MAXISSUES][MAXEDITIONS][MAXSHEETS][2];

	BYTE		m_sheetactivecopies[MAXISSUES][MAXEDITIONS][MAXSHEETS][2];
	TCHAR		m_sheetpageID[MAXISSUES][MAXEDITIONS][MAXSHEETS][2][MAXPAGEPOSITIONS][MAXPAGEIDLEN];
	TCHAR		m_sheetmasterpageID[MAXISSUES][MAXEDITIONS][MAXSHEETS][2][MAXPAGEPOSITIONS][MAXPAGEIDLEN];

	TCHAR		m_oldsheetpagename[MAXISSUES][MAXEDITIONS][MAXSHEETS][2][MAXPAGEPOSITIONS][MAXPAGEIDLEN];
	BYTE		m_oldsheetsectionID[MAXISSUES][MAXEDITIONS][MAXSHEETS][2][MAXPAGEPOSITIONS];

	BYTE		m_sheetdefaultsectionID[MAXISSUES][MAXEDITIONS][MAXSHEETS][2];




	//#pragma pack(8)

};

typedef struct {
	CString sTime;
	CString sEvent;
	CString sProduct;
	CString sComment;
	CString sProcess;
} LOGITEM;

typedef CArray <LOGITEM, LOGITEM&> LOGLIST;



typedef struct {
	TCHAR	PlanPageName[100];
	int		MasterCopySeparationSet;
	TCHAR	PageName[20];
	int		SectionID;
	int		Status;
	int		ProofStatus;
	int		Approved;
	CTime	InputTime;
	int		Version;
	int		InputID;
	TCHAR	FileServer[100];
	TCHAR	FileName[260];
	int		InputProcessID;
	int		EmailStatus;
	int		Active;
	int		PageType;
	int		ColorID;
	int		UniquePage;
} PLANPAGETABLESTRUCT;

typedef CArray <PLANPAGETABLESTRUCT, PLANPAGETABLESTRUCT&> PLANPAGETABLELIST;


typedef struct PPINAMETRANSLATION {
	CString		m_PPIProduct;
	CString		m_PPIEdition;
	CString		m_Publication;
} PPINAMETRANSLATION;
typedef CArray <PPINAMETRANSLATION, PPINAMETRANSLATION&> PPINAMETRANSLATIONLIST;