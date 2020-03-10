
#include "stdafx.h"
#include <afxmt.h>
#include <afxtempl.h>
#include <direct.h>
#include <iostream>
#include <winnetwk.h>
#include "DatabaseManager.h"
#include "Defs.h"
#include "Utils.h"
#include "Prefs.h"
#include "Markup.h"
#include "ParseXML.h"
#include "WorkerThread.h"
#include "Utils.h"

extern CPrefs	g_prefs;
extern CUtils g_util;
extern CDatabaseManager m_pollDB;

BOOL IsInternalXML(CString sFullInputPath, CString &sErrorMessage)
{
	CString str;
	CFile file;
	CString s;

	sErrorMessage = _T("");

	if (!file.Open(sFullInputPath, CFile::modeRead)) {
		sErrorMessage.Format("Cannot open XML file %s for reading", (LPCSTR)sFullInputPath);
		return FALSE;
	}

	sFullInputPath = g_util.GetFileName(sFullInputPath);

	int nFileLen = (int)file.GetLength();

	// Allocate buffer for binary file data
	BYTE* pBuffer = new BYTE[nFileLen + 2];
	nFileLen = file.Read(pBuffer, nFileLen);
	file.Close();
	pBuffer[nFileLen] = '\0';
	pBuffer[nFileLen + 1] = '\0'; // in case 2-byte encoded

	CString csNotes, csText;
	if (pBuffer[0] == 0xff && pBuffer[1] == 0xfe) {
		csText = (LPCWSTR)(&pBuffer[2]);
		csNotes += _T("File starts with hex FFFE, assumed to be wide char format. ");
	}
	else {
		csText = (LPCSTR)pBuffer;
	}
	delete[] pBuffer;

	// If it is too short, assume it got truncated due to non-text content
	if (csText.GetLength() < nFileLen / 2 - 20) {
		sErrorMessage.Format("Error converting file %s to string (may contain binary data)", (LPCSTR)sFullInputPath);
		return FALSE;
	}

	// Parse
#ifdef MARKUP_MSXML
	CMarkupMSXML xml;
#elif defined( MARKUP_STL )
	CMarkupSTL xml;
#else
	CMarkup xml;
#endif

	// Display results
	if (xml.SetDoc(csText) == FALSE) {
#ifdef MARKUP_STL
		CString csError = xml.GetError().c_str();
#else
		CString csError = xml.GetError();
#endif
		sErrorMessage.Format("Error parsing file %s - %s", (LPCSTR)sFullInputPath, (LPCSTR)csError);
		return FALSE;
	}

	bool hasPlanElement = xml.FindElem(_T("Plan"));
	xml.IntoElem();
	bool hasPublicationElement = xml.FindElem(_T("Publication"));


	return hasPlanElement && hasPublicationElement;
}

int ParseXML(CString sFullInputPath, PLANSTRUCT *plan, int &nPagesImported, CString &sErrMsg, CString &sInfo,
	CString &sPressName, BOOL &bIgnorePostCommand, BOOL &bKeepExistingSections)
{
	CString str;
	CFile file;
	CString sErrorMessage;
	bIgnorePostCommand = FALSE;
	bKeepExistingSections = FALSE;
	nPagesImported = 0;
	sInfo = _T("");
	sErrMsg = _T("");
	sPressName = _T("");

	BOOL bPagesOnly = FALSE;
	PUBLICATIONPRESSDEFAULTSSTRUCT *pPublicationDefaults = NULL;

	int copyseparationset = 1;
	int nPageIDAutoGen = 1;
	if (!file.Open(sFullInputPath, CFile::modeRead)) {
		sErrMsg.Format("Cannot open XML file %s for reading", sFullInputPath);
		return PARSEFILE_ERROR;
	}

	sFullInputPath = g_util.GetFileName(sFullInputPath);

	int nFileLen = (int)file.GetLength();

	// Allocate buffer for binary file data
	unsigned char* pBuffer = new unsigned char[nFileLen + 2];
	nFileLen = file.Read(pBuffer, nFileLen);
	file.Close();
	pBuffer[nFileLen] = '\0';
	pBuffer[nFileLen + 1] = '\0'; // in case 2-byte encoded

	CString csNotes, csText;
	if (pBuffer[0] == 0xff && pBuffer[1] == 0xfe) {
		csText = (LPCWSTR)(&pBuffer[2]);
		csNotes += _T("File starts with hex FFFE, assumed to be wide char format. ");
	}
	else {
		csText = (LPCSTR)pBuffer;
	}
	delete[] pBuffer;

	// If it is too short, assume it got truncated due to non-text content
	if (csText.GetLength() < nFileLen / 2 - 20) {
		sErrMsg.Format("Error converting file %s to string (may contain binary data)", sFullInputPath);
		return PARSEFILE_ERROR;
	}

	// Parse
#ifdef MARKUP_MSXML
	CMarkupMSXML xml;
#elif defined( MARKUP_STL )
	CMarkupSTL xml;
#else
	CMarkup xml;
#endif

	// Display results
	if (xml.SetDoc(csText) == FALSE) {
#ifdef MARKUP_STL
		CString csError = xml.GetError().c_str();
#else
		CString csError = xml.GetError();
#endif
		sErrMsg.Format("Error parsing file %s - %s", sFullInputPath, csError);
		return PARSEFILE_ERROR;
	}

	if (xml.FindElem(_T("Plan")) == false) {
		sErrMsg.Format("Error parsing file %s - <Plan> element not found", sFullInputPath);
		return PARSEFILE_ERROR;
	}

	plan->m_version = atoi(xml.GetAttrib("Version"));
	plan->m_runpostproc = atoi(xml.GetAttrib("RunPostProcedure"));

	bIgnorePostCommand = atoi(xml.GetAttrib("IgnorePostCommand"));

	plan->m_planID = xml.GetAttrib("ID");
	plan->m_planName = xml.GetAttrib("Name");
	plan->m_updatetime = xml.GetAttrib("UpdateTime");
	plan->m_sender = xml.GetAttrib("Sender");

	CString sCommand = xml.GetAttrib("Command");
	if (sCommand == "DeletePlan") {
		return PARSEFILE_ABORT;
	}

	if (sCommand == "AppendPlan")
		bKeepExistingSections = TRUE;

	xml.IntoElem();

	/////////////////////////////////////////

	if (xml.FindElem(_T("Publication")) == false) {
		sErrMsg.Format("Error parsing file %s - <Publication> element not found", sFullInputPath);
		return PARSEFILE_ERROR;
	}


	///////////////////////////
	// Store publication date
	///////////////////////////

	str = xml.GetAttrib("PubDate");
	int year = atoi(str.Mid(0, 4));
	int month = atoi(str.Mid(5, 2));
	int day = atoi(str.Mid(8, 2));

	if (year < 2000 || year > 3000 || month < 1 || month > 12 || day < 1 || day > 31) {
		sErrMsg.Format("Error parsing file %s - PubDate attribute '%s' out of range", sFullInputPath, str);
		return PARSEFILE_ERROR;
	}
	try {
		plan->m_pubdate = CTime(year, month, day, 0, 0, 0);
	}
	catch (CException *e)
	{
		sErrMsg.Format("Error parsing file %s - PubDate attribute '%s' out of range", sFullInputPath, str);
		return PARSEFILE_ERROR;
	}

	plan->m_weekreference = atoi(xml.GetAttrib("WeekReference"));

	plan->m_customer = xml.GetAttrib("Customer");
	plan->m_customeralias = xml.GetAttrib("CustomerAlias");

	plan->m_usemask = atoi(xml.GetAttrib("ShowMask"));
	plan->m_masksizex = atof(xml.GetAttrib("MaskWidth"));
	plan->m_masksizey = atof(xml.GetAttrib("MaskHeight"));
	plan->m_maskbleed = atof(xml.GetAttrib("MaskBleed"));
	CString s = xml.GetAttrib("MaskSnap");

	plan->m_maskshowspread = TRUE;
	plan->m_masksnapeven = SNAP_CR;
	plan->m_masksnapodd = SNAP_CL;
	if (s.CompareNoCase("center") == 0) {
		plan->m_masksnapeven = SNAP_CENTER;
		plan->m_masksnapodd = SNAP_CENTER;
	}

	// Disable this!!
	plan->m_usemask = FALSE;
	plan->m_masksizex = 0;
	plan->m_masksizey = 0;
	plan->m_maskbleed = 0;

	///////////////////////////
	// Store publication name
	///////////////////////////

	str = xml.GetAttrib("Name");
	CString str2 = xml.GetAttrib("Alias");

	if (g_prefs.m_onlyusepubalias)
		str = str2;

	for (int i = 0; i < g_prefs.m_PublicationsRequiringPostProc.GetCount(); i++) {
		if (str.CompareNoCase(g_prefs.m_PublicationsRequiringPostProc[i]) == 0) {
			plan->m_runpostproc = TRUE;
			bIgnorePostCommand = FALSE;
		}
		if (str2.CompareNoCase(g_prefs.m_PublicationsRequiringPostProc[i]) == 0) {
			plan->m_runpostproc = TRUE;
			bIgnorePostCommand = FALSE;
		}
	}

	if (g_prefs.m_overrulepublication)
		plan->m_publicationID = g_prefs.GetPublicationID(g_prefs.m_publication);
	else {
		plan->m_publicationID = g_prefs.GetPublicationID(str);

		// Retry after full cache reload..
		if (plan->m_publicationID == 0) {
			g_prefs.LoadAllPrefs(&m_pollDB);
			plan->m_publicationID = g_prefs.GetPublicationID(str);
		}

		// Try alias..
		if (plan->m_publicationID == 0)
		{
			plan->m_publicationID = g_prefs.GetPublicationID(str2);
			if (plan->m_publicationID > 0)
			{
				// Got alias
				g_prefs.ChangePublicationName(str, str2);
				PUBLICATIONSTRUCT *pPub = g_prefs.GetPublicationStruct(plan->m_publicationID);
				pPub->m_name = str;
			}
		}

		if (plan->m_publicationID == 0 && g_prefs.m_allowunknownnames) {
			plan->m_publicationID = g_prefs.AddNewPublication(str, str2);
		}

		if (plan->m_publicationID == 0) {
			sErrMsg.Format("Error parsing file %s - unknown publication name %s", sFullInputPath, str);
			return PARSEFILE_ERROR;
		}
	}

	xml.IntoElem();

	if (xml.FindElem(_T("Issues")) == false) {
		sErrMsg.Format("Error parsing file %s - <Issues> element not found", sFullInputPath);
		return PARSEFILE_ERROR;
	}

	xml.IntoElem();

	plan->m_numberofissues = 0;
	while (xml.FindElem(_T("Issue"))) {

		///////////////////////////
		// Store issue name
		///////////////////////////

		str = xml.GetAttrib("Name");
		if (g_prefs.m_overruleissue)
			plan->m_issueIDlist[plan->m_numberofissues] = g_prefs.GetIssueID(g_prefs.m_issue);
		else {
			plan->m_issueIDlist[plan->m_numberofissues] = g_prefs.GetIssueID(str);
			if (plan->m_issueIDlist[plan->m_numberofissues] == 0) {
				sErrMsg.Format("Error parsing file %s - unknown issue name %s", sFullInputPath, str);
				return PARSEFILE_ERROR;
			}
		}

		xml.IntoElem();	// in Issue

		if (xml.FindElem(_T("Editions")) == false) {
			sErrMsg.Format("Error parsing file %s - <Editions> element not found", sFullInputPath);
			return PARSEFILE_ERROR;
		}

		xml.IntoElem();	// in Editions
		int issues = plan->m_numberofissues;
		plan->m_numberofeditions = 0;
		int timedEditionNumber = 1;
		while (xml.FindElem(_T("Edition"))) {

			///////////////////////////
			// Store edition name
			///////////////////////////
			int eds = plan->m_numberofeditions;
			str = xml.GetAttrib("NoChange");
			plan->m_editionnochange[issues][eds] = (str == "true" || str == "1");

			plan->m_numberofpapercopiesedition[issues][eds] = atoi(xml.GetAttrib("EditionCopies"));

			str = xml.GetAttrib("Name");
			if (g_prefs.m_overruleedition)
				plan->m_editionIDlist[issues][eds] = g_prefs.GetEditionID(g_prefs.m_edition);
			else {
				plan->m_editionIDlist[issues][eds] = g_prefs.GetEditionID(str);


				if (plan->m_editionIDlist[issues][eds] == 0)
				{
					str = g_prefs.LookupNameFromAbbreviation("Edition", str);
					plan->m_editionIDlist[issues][eds] = g_prefs.GetEditionID(str);
				}


				if (plan->m_editionIDlist[issues][eds] == 0 && g_prefs.m_allowunknownnames) {
					plan->m_editionIDlist[issues][eds] = g_prefs.AddNewEdition(str, "");
				}

				if (plan->m_editionIDlist[issues][eds] == 0) {
					sErrMsg.Format("Error parsing file %s - unknown edition name %s", sFullInputPath, str);
					return PARSEFILE_ERROR;
				}
			}

			// Timed edition management
			CString s = xml.GetAttrib("IsTimed");
			s.MakeLower();
			plan->m_editionistimed[issues][eds] = s == "1" || s == "true";

			plan->m_editiontimedfromID[issues][eds] = g_prefs.GetEditionID(xml.GetAttrib("TimedFrom"));
			plan->m_editiontimedtoID[issues][eds] = g_prefs.GetEditionID(xml.GetAttrib("TimedTo"));
			plan->m_editionordernumber[issues][eds] = atoi(xml.GetAttrib("EditionOrderNumber"));
			_tcscpy(plan->m_editioncomment[issues][eds], xml.GetAttrib("EditionComment"));
			plan->m_editionmasterzoneeditionID[issues][eds] = g_prefs.GetEditionID(xml.GetAttrib("ZoneMasterEdition"));
			if (plan->m_editionistimed[issues][eds] && plan->m_editionordernumber[issues][eds] == 0)
				plan->m_editionordernumber[issues][eds] = timedEditionNumber++;

			// If this edition has no FROM and TO editions it is not timed..
			//		TimedFrom	TimedTo		EditionOrder	ZoneMasterEdition
			// A1		0			A2			1				0
			// A2		A1			A3			2				0
			// A3		A2			0			3				0
			// B1 (sub) 0			B2			1				A1
			// B2 (sub) B1			0			2				A2

			if (plan->m_editiontimedfromID[issues][eds] == 0 && plan->m_editiontimedtoID[issues][eds] == 0)
				plan->m_editionistimed[issues][eds] = FALSE;

			// Reset edition order number if we are out of a timed edition sequence
			if (plan->m_editionistimed[issues][eds] == FALSE || plan->m_editiontimedfromID[issues][eds] == 0)
				timedEditionNumber = 1;

			// Old compatibility hack..

			CString sEditionPressOldStyle = xml.GetAttrib("Press");
			sEditionPressOldStyle = g_prefs.LookupNameFromAbbreviation("Press", sEditionPressOldStyle);
			int nCopiesOldStyle = atoi(xml.GetAttrib("Copies"));

			xml.IntoElem();	// Into edition

			if (xml.FindElem(_T("IntendedPresses"))) {

				xml.IntoElem();	// in IntendedPresses

				int nPressesInEdition = 0;
				while (xml.FindElem(_T("IntendedPress"))) {
					str = xml.GetAttrib("Name");

					/*				if (g_prefs.m_overruletemplate) {
										int nPressID = g_prefs.GetPressIDFromTemplateID(g_prefs.GetTemplateID(g_prefs.m_template));
										if (nPressID > 0)
											str = g_prefs.GetPressName(nPressID);
									}
					*/



					if (g_prefs.m_overrulepress)
						str = g_prefs.m_overruledpress;

					str = g_prefs.LookupNameFromAbbreviation("Press", str);
					sPressName = str;

					if (g_prefs.GetPressID(str) == 0) {
						sErrMsg.Format("Error parsing file %s - <IntendedPress> %s not known", sFullInputPath, str);
						return PARSEFILE_ERROR;
					}

					str2 = xml.GetAttrib("ExpectedPressName");
					str2 = g_prefs.LookupNameFromAbbreviation("Press", str2);


					plan->m_press = str;
					plan->m_rippress = str2 != "" ? str2 : str;
					plan->m_pressesinedition[issues][eds][nPressesInEdition] = g_prefs.GetPressID(str);
					plan->m_pressIDlist[issues][eds] = g_prefs.GetPressID(str);
					plan->m_pressID = g_prefs.GetPressID(plan->m_press);
					plan->m_rippressID = g_prefs.GetPressID(plan->m_rippress);

					if (plan->m_pressIDlist[issues][eds] == 0) {
						sErrMsg.Format("Error parsing file %s - unknown press name %s", sFullInputPath, str);
						return  PARSEFILE_ERROR;;
					}

					plan->m_numberofpapercopies[issues][eds][nPressesInEdition] = atoi(xml.GetAttrib("Copies"));
					g_util.strcpyexx(plan->m_presstime[issues][eds][nPressesInEdition], xml.GetAttrib("PressTime"), 32);
					g_util.strcpyexx(plan->m_postalurl[issues][eds][nPressesInEdition], xml.GetAttrib("PostalUrl"), MAXPOSTALURL);

					CString ss = xml.GetAttrib("AutoRelease");
					ss.MakeLower();
					plan->m_autorelease[issues][eds][nPressesInEdition] = (ss == "1" || ss == "true") ? TRUE : FALSE;


					plan->m_numberofcopies[issues][eds][nPressesInEdition] = 1;
					
					
					nPressesInEdition++;
				}
				xml.OutOfElem(); //</IntendedPresses>
			}

			if (xml.FindElem(_T("Sections")) == false) {
				sErrMsg.Format("Error parsing file %s - <Sections> element not found", sFullInputPath);
				return PARSEFILE_ERROR;
			}

			xml.IntoElem();	// in Sections

			plan->m_numberofsections[issues][eds] = 0;
			while (xml.FindElem(_T("Section"))) {

				///////////////////////////
				// Store section name
				///////////////////////////
				int secs = plan->m_numberofsections[issues][eds];


				str = xml.GetAttrib("Name");
				if (g_prefs.m_overrulesection)
					plan->m_sectionIDlist[issues][eds][secs] = g_prefs.GetSectionID(g_prefs.m_section);
				else {
					plan->m_sectionIDlist[issues][eds][secs] = g_prefs.GetSectionID(str);

					if (plan->m_sectionIDlist[issues][eds][secs] == 0)
					{
						str = g_prefs.LookupNameFromAbbreviation("Section", str);
						plan->m_sectionIDlist[issues][eds][secs] = g_prefs.GetSectionID(str);
					}

					if (plan->m_sectionIDlist[issues][eds][secs] == 0 && g_prefs.m_allowunknownnames) {
						plan->m_sectionIDlist[issues][eds][secs] = g_prefs.AddNewSection(str, "");
					}

					if (plan->m_sectionIDlist[issues][eds][secs] == 0) {
						sErrMsg.Format("Error parsing file %s - unknown section name %s", sFullInputPath, str);
						return PARSEFILE_ERROR;
					}
				}

				str = xml.GetAttrib("Title");
				CString str2 = xml.GetAttrib("Format");
				if (str2 != "")
					str += ":" + str2;
				g_util.strcpyexx(plan->m_sectionTitlelist[issues][eds][secs], str, MAXCOMMENT);

				CString sPageFormatName = xml.GetAttrib("PageFormat");

				double fPageWidth = atof(xml.GetAttrib("TrimWidth"));
				double fPageHeight = atof(xml.GetAttrib("TrimHeight"));
				double fPageBleed = atof(xml.GetAttrib("TrimBleed"));
				CString sMode = xml.GetAttrib("TrimMode");
				int nTrimMode = TRIMMODE_NONE;
				if (sMode.CompareNoCase("trim") == 0)
					nTrimMode = TRIMMODE_TRIM;
				else if (sMode.CompareNoCase("scale") == 0)
					nTrimMode = TRIMMODE_SCALE;

				if (sPageFormatName != "" && fPageWidth > 0.0 && fPageHeight > 0.0) {
					
					if (m_pollDB.FindPageFormat(sPageFormatName, sErrorMessage) == 0) {
						s.Format("%dX%d", (int)fPageWidth, (int)fPageHeight);
						if (m_pollDB.FindPageFormat(s, sErrorMessage) == 1)
							sPageFormatName = s;
						else {
							if (m_pollDB.InsertPageFormat(sPageFormatName, fPageWidth, fPageHeight, fPageBleed, 7, 6, nTrimMode, sErrorMessage))
								g_util.Logprintf("INFO: New PageFormat added to system - %s, %.2f x %.2f", sPageFormatName, fPageWidth, fPageHeight);
						}
					}
				}
				g_util.strcpyexx(plan->m_sectionPageFormat[issues][eds][secs], sPageFormatName, MAXCOMMENT);

				xml.IntoElem();

				if (xml.FindElem(_T("Pages")) == false) {
					sErrMsg.Format("Error parsing file %s - <Pages> element not found", sFullInputPath);
					return PARSEFILE_ERROR;
				}

				xml.IntoElem();	// in Pages

				plan->m_numberofpages[issues][eds][secs] = 0;
				//	plan->m_allnonuniqueepages[issues][eds] = TRUE;
				plan->m_allnonuniqueepages[issues][eds] = FALSE;

				CString sPrevPageName = "";
				int nPageCounter = 0;
				while (xml.FindElem(_T("Page"))) {
					nPageCounter++;

					BOOL bConvertedToPanorama = FALSE;
					///////////////////////////
					// Store page name
					///////////////////////////
					int pages = plan->m_numberofpages[issues][eds][secs];

					CString sPageName = xml.GetAttrib("Name");
					CString sPageFileName = xml.GetAttrib("FileName");
					CString sPageType = xml.GetAttrib("PageType");
					CString sComment = xml.GetAttrib("Comment");
					int nPageType = PAGETYPE_NORMAL;

					bConvertedToPanorama = TRUE;

					if (sPageType.CollateNoCase("panorama") == 0)
						nPageType = PAGETYPE_PANORAMA;
					else if (sPageType.CollateNoCase("antipanorama") == 0)
						nPageType = PAGETYPE_ANTIPANORAMA;
					else if (sPageType.CollateNoCase("dummy") == 0 || sPageType.CollateNoCase("dinky") == 0)
						nPageType = PAGETYPE_DUMMY;
					else
						nPageType = PAGETYPE_NORMAL;

					if (sPageName == "" && nPageType != PAGETYPE_DUMMY) {
						sErrMsg.Format("Error parsing file %s - <Page> element Name attribute empty", sFullInputPath);
						return PARSEFILE_ERROR;
					}
					plan->m_pagetypes[issues][eds][secs][pages] = nPageType;

					int nPagination = atoi(xml.GetAttrib("Pagination"));

					int nPageIndex = atoi(xml.GetAttrib("PageIndex"));
					str = xml.GetAttrib("Unique");
					int nUnique = str.CompareNoCase("true") == 0 ? UNIQUEPAGE_UNIQUE : UNIQUEPAGE_COMMON;
					if (str.CompareNoCase("force") == 0)
						nUnique = UNIQUEPAGE_FORCE;

					if (nUnique != UNIQUEPAGE_COMMON)
						plan->m_allnonuniqueepages[issues][eds] = FALSE;

					plan->m_pageunique[issues][eds][secs][pages] = nUnique;

					g_util.strcpyexx(plan->m_pagecomments[issues][eds][secs][pages], sComment, MAXCOMMENT);
					g_util.strcpyexx(plan->m_pagenames[issues][eds][secs][pages], sPageName, MAXPAGENAME);
					g_util.strcpyexx(plan->m_pagefilenames[issues][eds][secs][pages], sPageFileName, MAXPAGENAME);

					if (nPagination > 0)
						plan->m_pagepagination[issues][eds][secs][pages] = nPagination;
					else if (atoi(sPageName) > 0)
						plan->m_pagepagination[issues][eds][secs][pages] = atoi(sPageName);
					else if (nPageIndex > 0)
						plan->m_pagepagination[issues][eds][secs][pages] = nPageIndex;
					else
						plan->m_pagepagination[issues][eds][secs][pages] = nPageCounter;

					if (nPageIndex > 0)
						plan->m_pageindex[issues][eds][secs][pages] = nPageIndex;
					else if (atoi(sPageName) > 0)
						plan->m_pagepagination[issues][eds][secs][pages] = atoi(sPageName);
					else
						plan->m_pagepagination[issues][eds][secs][pages] = nPageCounter;

					//	plan->m_pagemasteredition[issues][eds][secs][pages] = g_prefs.GetEditionID(sMasterEdition);
					CString sPageID = xml.GetAttrib("PageID");
					g_util.strcpyexx(plan->m_pageIDs[issues][eds][secs][pages], sPageID, MAXPAGEIDLEN);
					CString sMasterPageID = xml.GetAttrib("MasterPageID");
					if (sMasterPageID == "" && nUnique == UNIQUEPAGE_UNIQUE)
						sMasterPageID = sPageID;

					g_util.strcpyexx(plan->m_pagemasterIDs[issues][eds][secs][pages], sMasterPageID, MAXPAGEIDLEN);


					BOOL	nApproved = atoi(xml.GetAttrib("Approved"));
					if (xml.GetAttrib("Approve") != "")
						nApproved = atoi(xml.GetAttrib("Approve"));

					if (g_prefs.m_overruleapproval)
						plan->m_pageapproved[issues][eds][secs][pages] = g_prefs.m_approve ? 0 : -1;
					else
						plan->m_pageapproved[issues][eds][secs][pages] = nApproved ? 0 : -1;

					BOOL	nHold = atoi(xml.GetAttrib("Hold"));
					if (g_prefs.m_overrulerelease)
						plan->m_pagehold[issues][eds][secs][pages] = g_prefs.m_hold;
					else
						plan->m_pagehold[issues][eds][secs][pages] = nHold;

					int		nPriority = atoi(xml.GetAttrib("Priority"));
					if (nPriority < 0 || nPriority > 100) {
						sErrMsg.Format("Error parsing file %s - priority %d out of range", sFullInputPath, nPriority);
						return PARSEFILE_ERROR;
					}

					if (g_prefs.m_overrulepriority)
						plan->m_pagepriority[issues][eds][secs][pages] = g_prefs.m_priority;
					else
						plan->m_pagepriority[issues][eds][secs][pages] = nPriority;

					plan->m_pageversion[issues][eds][secs][pages] = atoi(xml.GetAttrib("Version")); ;


					plan->m_pagemiscint[issues][eds][secs][pages] = atoi(xml.GetAttrib("MiscInt"));
					g_util.strcpyexx(plan->m_pagemiscstring1[issues][eds][secs][pages], xml.GetAttrib("MiscString1"), MAXCOMMENT);
					g_util.strcpyexx(plan->m_pagemiscstring2[issues][eds][secs][pages], xml.GetAttrib("MiscString2"), MAXCOMMENT);

					sPrevPageName = sPageName;

					xml.IntoElem();

					if (xml.FindElem(_T("Separations")) == false) {
						sErrMsg.Format("Error parsing file %s - <Separations> element not found", sFullInputPath);
						return PARSEFILE_ERROR;
					}

					xml.IntoElem();	// in Separations

					plan->m_numberofpagecolors[issues][eds][secs][pages] = 0;
					while (xml.FindElem(_T("Separation"))) {

						///////////////////////////
						// Store color separation name
						///////////////////////////

						//xml.IntoElem();	
						int ncolors = plan->m_numberofpagecolors[issues][eds][secs][pages];

						CString str = xml.GetAttrib("Name");

						int nColorID = g_prefs.GetColorID(str);
						if (nColorID == 0) {
							sErrMsg.Format("Error parsing file %s - unknown color name %s", sFullInputPath, str);
							return PARSEFILE_ERROR;
						}
						plan->m_pagecolorIDlist[issues][eds][secs][pages][ncolors] = g_prefs.m_pdfpages ? g_prefs.GetColorID("PDF") : nColorID;
						plan->m_pagecolorActivelist[issues][eds][secs][pages][ncolors] = TRUE;

						plan->m_numberofpagecolors[issues][eds][secs][pages]++;
						if (g_prefs.m_pdfpages)
							break;
						//xml.OutOfElem(); //</Separation>
					}

					// Handle auto-creation of filler seps.
					if (g_prefs.m_createcmyk && plan->m_numberofpagecolors[issues][eds][secs][pages] == 1 && g_prefs.m_pdfpages == FALSE) {
						int ncolorsnow = plan->m_numberofpagecolors[issues][eds][secs][pages];

						for (int u = 0; u < ncolorsnow; u++) {
							if (g_prefs.IsBlackColor(plan->m_pagecolorIDlist[issues][eds][secs][pages][u]) == FALSE) {
								plan->m_pagecolorIDlist[issues][eds][secs][pages][plan->m_numberofpagecolors[issues][eds][secs][pages]] = g_prefs.GetColorID(_T("K"));
								plan->m_pagecolorActivelist[issues][eds][secs][pages][plan->m_numberofpagecolors[issues][eds][secs][pages]] = FALSE;
								plan->m_numberofpagecolors[issues][eds][secs][pages]++;
							}
							if (g_prefs.IsCyanColor(plan->m_pagecolorIDlist[issues][eds][secs][pages][u]) == FALSE) {
								plan->m_pagecolorIDlist[issues][eds][secs][pages][plan->m_numberofpagecolors[issues][eds][secs][pages]] = g_prefs.GetColorID(_T("C"));
								plan->m_pagecolorActivelist[issues][eds][secs][pages][plan->m_numberofpagecolors[issues][eds][secs][pages]] = FALSE;
								plan->m_numberofpagecolors[issues][eds][secs][pages]++;
							}
							if (g_prefs.IsMagentaColor(plan->m_pagecolorIDlist[issues][eds][secs][pages][u]) == FALSE) {
								plan->m_pagecolorIDlist[issues][eds][secs][pages][plan->m_numberofpagecolors[issues][eds][secs][pages]] = g_prefs.GetColorID(_T("M"));
								plan->m_pagecolorActivelist[issues][eds][secs][pages][plan->m_numberofpagecolors[issues][eds][secs][pages]] = FALSE;
								plan->m_numberofpagecolors[issues][eds][secs][pages]++;

							}
							if (g_prefs.IsYellowColor(plan->m_pagecolorIDlist[issues][eds][secs][pages][u]) == FALSE) {
								plan->m_pagecolorIDlist[issues][eds][secs][pages][plan->m_numberofpagecolors[issues][eds][secs][pages]] = g_prefs.GetColorID(_T("Y"));
								plan->m_pagecolorActivelist[issues][eds][secs][pages][plan->m_numberofpagecolors[issues][eds][secs][pages]] = FALSE;
								plan->m_numberofpagecolors[issues][eds][secs][pages]++;

							}
						}
					}

					xml.OutOfElem(); //</Separations>
					copyseparationset++;

					if (plan->m_numberofpagecolors[issues][eds][secs][pages] == 0 && g_prefs.m_assumecmykifnoseps) {
						plan->m_pagecolorIDlist[issues][eds][secs][pages][0] = g_prefs.GetColorID("C");
						plan->m_pagecolorActivelist[issues][eds][secs][pages][0] = TRUE;

						plan->m_pagecolorIDlist[issues][eds][secs][pages][1] = g_prefs.GetColorID("M");
						plan->m_pagecolorActivelist[issues][eds][secs][pages][1] = TRUE;

						plan->m_pagecolorIDlist[issues][eds][secs][pages][2] = g_prefs.GetColorID("Y");
						plan->m_pagecolorActivelist[issues][eds][secs][pages][2] = TRUE;

						plan->m_pagecolorIDlist[issues][eds][secs][pages][3] = g_prefs.GetColorID("K");
						plan->m_pagecolorActivelist[issues][eds][secs][pages][3] = TRUE;

						plan->m_numberofpagecolors[issues][eds][secs][pages] = 4;

						g_util.Logprintf("WARNING: Autogenerated CMYK for pageID %s", plan->m_pageIDs[issues][eds][secs][pages]);

					}
					else if (plan->m_numberofpagecolors[issues][eds][secs][pages] == 0) {
						sErrMsg.Format("Error parsing file %s - no <Separation> elements  found in <Separations> element", sFullInputPath);
						return PARSEFILE_ERROR;
					}
					plan->m_numberofpages[issues][eds][secs]++;
					nPagesImported++;
					xml.OutOfElem(); //</Page>

					// Cater for anti-panorama..
					/*
					if (nPageType == PAGETYPE_PANORAMA && bConvertedToPanorama == FALSE) {
						strcpyexx(plan->m_pagecomments[issues][eds][secs][pages+1], plan->m_pagecomments[issues][eds][secs][pages], MAXCOMMENT);
						CString ss = plan->m_pagenames[issues][eds][secs][pages];
						int mm = strlen(plan->m_pagenames[issues][eds][secs][pages]);
						TCHAR chfirst = plan->m_pagenames[issues][eds][secs][pages][0];
						TCHAR chlast = plan->m_pagenames[issues][eds][secs][pages][mm-1];

						if (isdigit(chfirst) && isdigit(chlast))
							sprintf(plan->m_pagenames[issues][eds][secs][pages+1],"%d",atoi(plan->m_pagenames[issues][eds][secs][pages])+1);
						else if (!isdigit(chfirst))
							sprintf(plan->m_pagenames[issues][eds][secs][pages+1],"%c%d",chfirst,atoi(plan->m_pagenames[issues][eds][secs][pages]+1)+1);
						else if (!isdigit(chlast))
							sprintf(plan->m_pagenames[issues][eds][secs][pages+1],"%d%c",atoi(plan->m_pagenames[issues][eds][secs][pages])+1,chlast);

						plan->m_pagepagination[issues][eds][secs][pages+1] = plan->m_pagepagination[issues][eds][secs][pages]+1;
						plan->m_pageunique[issues][eds][secs][pages+1] = plan->m_pageunique[issues][eds][secs][pages];
						plan->m_pagetypes[issues][eds][secs][pages+1] = PAGETYPE_ANTIPANORAMA;
						plan->m_pagemasteredition[issues][eds][secs][pages+1] = plan->m_pagemasteredition[issues][eds][secs][pages];

						plan->m_pagepriority[issues][eds][secs][pages+1] = plan->m_pagepriority[issues][eds][secs][pages];
						plan->m_pageapproved[issues][eds][secs][pages+1] = plan->m_pageapproved[issues][eds][secs][pages];
						plan->m_numberofpagecolors[issues][eds][secs][pages+1] = plan->m_numberofpagecolors[issues][eds][secs][pages];

						plan->m_pagehold[issues][eds][secs][pages+1] = plan->m_pagehold[issues][eds][secs][pages];
						plan->m_pageversion[issues][eds][secs][pages+1] = plan->m_pageversion[issues][eds][secs][pages];

						for (int ncolors=0; ncolors<plan->m_numberofpagecolors[issues][eds][secs][pages]; ncolors++) {
							plan->m_pagecolorIDlist[issues][eds][secs][pages+1][ncolors] = plan->m_pagecolorIDlist[issues][eds][secs][pages][ncolors];
							plan->m_pagecolorActivelist[issues][eds][secs][pages+1][ncolors] = FALSE;

						}
						plan->m_numberofpages[issues][eds][secs]++;
						nPagesImported++;
					}*/

				}
				xml.OutOfElem(); //</Pages>

				if (g_prefs.m_allowemptysections == FALSE) {
					if (plan->m_numberofpages[issues][eds][secs] == 0) {
						sErrMsg.Format("Error parsing file %s - no <Page> elements  found in <Pages> element", sFullInputPath);
						return PARSEFILE_ERROR;
					}
				}

				plan->m_numberofsections[issues][eds]++;
				xml.OutOfElem(); //</Section>
			}
			if (plan->m_numberofsections[issues][eds] == 0) {
				sErrMsg.Format("Error parsing file %s - no <Section> elements  found in <Sections> element", sFullInputPath);
				return PARSEFILE_ERROR;
			}
			xml.OutOfElem(); //</Sections>


			if (xml.FindElem(_T("Sheets")) == false) {
				//				sprintf(szErrMsg,"Error parsing file %s - <Sheets> element not found", sFullInputPath);

								//SetMaxPageColors(plan);

				bPagesOnly = TRUE;
			}
			else {

				xml.IntoElem();	// in Sheets

				plan->m_numberofsheets[issues][eds] = 0;
				while (xml.FindElem(_T("Sheet"))) {

					///////////////////////////
					// Store section name
					///////////////////////////

					int sheets = plan->m_numberofsheets[issues][eds];
					plan->m_sheetbacktemplateID[issues][eds][sheets] = 0;
					str = xml.GetAttrib("Template");
					if (str == _T(""))
						str = xml.GetAttrib("TemplateFront");
					str2 = xml.GetAttrib("TemplateBack");

					int nPressID = plan->m_pressIDlist[issues][eds];
					int nPressIndex = g_prefs.GetPressIndex(nPressID);
					plan->m_sheettemplateID[issues][eds][sheets] = 1;// g_prefs.GetTemplateID(g_prefs.m_template[0]);

					if (plan->m_sheetbacktemplateID[issues][eds][sheets] == 0)
						plan->m_sheetbacktemplateID[issues][eds][sheets] = plan->m_sheettemplateID[issues][eds][sheets];

					plan->m_pressID = g_prefs.GetPressIDFromTemplateID(plan->m_sheettemplateID[issues][eds][sheets]);

					int nPagesOnPlate = atoi(xml.GetAttrib("PagesOnPlate"));
					g_util.strcpyexx(plan->m_sheetmarkgroup[issues][eds][sheets], xml.GetAttrib("MarkGroups"), MAXGROUPNAMES);


					plan->m_pressSectionNumber[issues][eds][sheets] = atoi(xml.GetAttrib("PressSectionNumber"));

					plan->m_sheetpagesonplate[issues][eds][sheets] = nPagesOnPlate;
					xml.IntoElem();	// Into Sheet

					if (xml.FindElem(_T("SheetFrontItems")) == false) {
						sErrMsg.Format("Error parsing file %s - <SheetFrontItems> element not found", sFullInputPath);
						return PARSEFILE_ERROR;
					}

					g_util.strcpyexx(plan->m_sheetsortingposition[issues][eds][sheets][0], xml.GetAttrib("SortingPosition"), MAXSORTINGPOS);
					g_util.strcpyexx(plan->m_sheetpresstower[issues][eds][sheets][0], xml.GetAttrib("PressTower"), MAXPRESSTOWERSTRING);
					g_util.strcpyexx(plan->m_sheetpresszone[issues][eds][sheets][0], xml.GetAttrib("PressZone"), MAXPRESSSTRING);
					g_util.strcpyexx(plan->m_sheetpresshighlow[issues][eds][sheets][0], xml.GetAttrib("PressHighLow"), MAXPRESSSTRING);

					plan->m_sheetactivecopies[issues][eds][sheets][0] = atoi(xml.GetAttrib("ActiveCopies"));
					xml.IntoElem();	// Into SheetFrontItems

					// Set some defaults..

					for (int ii = 0; ii < nPagesOnPlate; ii++) {
						plan->m_sheetpageposition[issues][eds][sheets][0][ii] = 0;
						plan->m_sheetpageposition[issues][eds][sheets][1][ii] = 0;
						g_util.strcpyexx(plan->m_sheetpageID[issues][eds][sheets][0][ii], _T("Dummy"), MAXPAGEIDLEN);
						g_util.strcpyexx(plan->m_sheetmasterpageID[issues][eds][sheets][0][ii], _T("Dummy"), MAXPAGEIDLEN);
						g_util.strcpyexx(plan->m_sheetpageID[issues][eds][sheets][1][ii], _T("Dummy"), MAXPAGEIDLEN);
						g_util.strcpyexx(plan->m_sheetmasterpageID[issues][eds][sheets][1][ii], _T("Dummy"), MAXPAGEIDLEN);
					}

					int nPagesOnThisSide = 0;
					CString sSectionOnPlate = _T("");
					while (xml.FindElem(_T("SheetFrontItem"))) {


						g_util.strcpyexx(plan->m_oldsheetpagename[issues][eds][sheets][0][nPagesOnThisSide], xml.GetAttrib("PageName"), MAXPAGEIDLEN);
						plan->m_oldsheetsectionID[issues][eds][sheets][0][nPagesOnThisSide] = g_prefs.GetSectionID(xml.GetAttrib("Section"));

						CString sThisPageID = xml.GetAttrib("PageID");
						CString sMP = xml.GetAttrib("MasterPageID");
						g_util.strcpyexx(plan->m_sheetpageID[issues][eds][sheets][0][nPagesOnThisSide], xml.GetAttrib("PageID"), MAXPAGEIDLEN);
						g_util.strcpyexx(plan->m_sheetmasterpageID[issues][eds][sheets][0][nPagesOnThisSide], sMP, MAXPAGEIDLEN);

						int Xpos = atoi(xml.GetAttrib("PosX"));
						if (Xpos < 1 || Xpos > 16) {
							sErrMsg.Format("Error parsing file %s - PosX %d atttribute out of range ", Xpos, sFullInputPath);
							return PARSEFILE_ERROR;
						}
						int Ypos = atoi(xml.GetAttrib("PosY"));
						if (Ypos < 1 || Ypos > 16) {
							sErrMsg.Format("Error parsing file %s - PosY %d atttribute out of range ", Ypos, sFullInputPath);
							return PARSEFILE_ERROR;
						}
						plan->m_sheetpageposition[issues][eds][sheets][0][nPagesOnThisSide] = GetPagePosition(nPagesOnPlate, Xpos, Ypos);

						nPagesOnThisSide++;

						// Get a default sectionID from a non-dummy for later dummy creation
						plan->m_sheetdefaultsectionID[issues][eds][sheets][0] = GetPageSectionID(plan, eds, sThisPageID);
					}

					if (nPagesOnThisSide == 0) {
						sErrMsg.Format("Error parsing file %s - Number of <SheetFrontItem> elements must be greater than zero", sFullInputPath);
						return PARSEFILE_ERROR;
					}

					for (int ii = 0; ii < nPagesOnPlate; ii++) {
						if (plan->m_sheetpageposition[issues][eds][sheets][0][ii] == 0) {
							// Find free pageposition for dummy page

							BOOL bFreePosition = TRUE;
							int nFreePosition = 0;
							for (int jj = 1; jj <= nPagesOnPlate; jj++) {

								bFreePosition = TRUE;
								nFreePosition = 0;

								for (int iii = 0; iii < nPagesOnPlate; iii++) {
									if (plan->m_sheetpageposition[issues][eds][sheets][0][iii] == jj) {
										bFreePosition = FALSE;
										break;
									}
								}
								if (bFreePosition) {
									nFreePosition = jj;
									break;
								}
							}
							if (bFreePosition && nFreePosition > 0)
								plan->m_sheetpageposition[issues][eds][sheets][0][ii] = nFreePosition;

						}
					}

					/*if (nPagesOnThisSide > nPagesOnPlate) {
						sprintf(szErrMsg,"Error parsing file %s - Number of <SheetFrontItem> elements greater than PagesOnPlate attribute ", sFullInputPath);
						return PARSEFILE_ERROR;
					}*/

					if (xml.FindElem(_T("PressCylindersFront"))) {
						xml.IntoElem();	// Into PressCylinders
						int nNumberOfPlateColors = 0;
						int nNumberOfPlateSpotColors = 0;

						// Clear spot
						g_util.strcpyexx(plan->m_sheetpresscylinder[issues][eds][sheets][0][4], _T(""), MAXPRESSSTRING);
						g_util.strcpyexx(plan->m_sheetformID[issues][eds][sheets][0][4], _T(""), MAXFORMIDSTRING);
						g_util.strcpyexx(plan->m_sheetplateID[issues][eds][sheets][0][4], _T(""), MAXFORMIDSTRING);
						g_util.strcpyexx(plan->m_sheetsortingpositioncolorspecific[issues][eds][sheets][0][4], _T(""), MAXFORMIDSTRING);

						while (xml.FindElem(_T("PressCylinderFront"))) {
							// Set colors to fixed places - now to handle spots??
							CString sCol = xml.GetAttrib("Color");
							if (sCol != "")
								nNumberOfPlateColors++;

							if (g_prefs.m_pdfpages) {
								g_util.strcpyexx(plan->m_sheetpresscylinder[issues][eds][sheets][0][0], xml.GetAttrib("Name"), MAXPRESSSTRING);
								g_util.strcpyexx(plan->m_sheetformID[issues][eds][sheets][0][0], xml.GetAttrib("FormID"), MAXFORMIDSTRING);
								g_util.strcpyexx(plan->m_sheetplateID[issues][eds][sheets][0][0], xml.GetAttrib("PlateID"), MAXFORMIDSTRING);
								g_util.strcpyexx(plan->m_sheetsortingpositioncolorspecific[issues][eds][sheets][0][0], xml.GetAttrib("SortingPosition"), MAXFORMIDSTRING);

							}
							else {
								BOOL bIsProcessColor = FALSE;
								if (sCol.CompareNoCase("cyan") == 0 || sCol.CompareNoCase("c") == 0) {
									g_util.strcpyexx(plan->m_sheetpresscylinder[issues][eds][sheets][0][0], xml.GetAttrib("Name"), MAXPRESSSTRING);
									g_util.strcpyexx(plan->m_sheetformID[issues][eds][sheets][0][0], xml.GetAttrib("FormID"), MAXFORMIDSTRING);
									g_util.strcpyexx(plan->m_sheetplateID[issues][eds][sheets][0][0], xml.GetAttrib("PlateID"), MAXFORMIDSTRING);
									g_util.strcpyexx(plan->m_sheetsortingpositioncolorspecific[issues][eds][sheets][0][0], xml.GetAttrib("SortingPosition"), MAXFORMIDSTRING);
									bIsProcessColor = TRUE;
								}
								if (sCol.CompareNoCase("magenta") == 0 || sCol.CompareNoCase("m") == 0) {
									g_util.strcpyexx(plan->m_sheetpresscylinder[issues][eds][sheets][0][1], xml.GetAttrib("Name"), MAXPRESSSTRING);
									g_util.strcpyexx(plan->m_sheetformID[issues][eds][sheets][0][1], xml.GetAttrib("FormID"), MAXFORMIDSTRING);
									g_util.strcpyexx(plan->m_sheetplateID[issues][eds][sheets][0][1], xml.GetAttrib("PlateID"), MAXFORMIDSTRING);
									g_util.strcpyexx(plan->m_sheetsortingpositioncolorspecific[issues][eds][sheets][0][1], xml.GetAttrib("SortingPosition"), MAXFORMIDSTRING);
									bIsProcessColor = TRUE;
								}
								if (sCol.CompareNoCase("yellow") == 0 || sCol.CompareNoCase("y") == 0) {
									g_util.strcpyexx(plan->m_sheetpresscylinder[issues][eds][sheets][0][2], xml.GetAttrib("Name"), MAXPRESSSTRING);
									g_util.strcpyexx(plan->m_sheetformID[issues][eds][sheets][0][2], xml.GetAttrib("FormID"), MAXFORMIDSTRING);
									g_util.strcpyexx(plan->m_sheetplateID[issues][eds][sheets][0][2], xml.GetAttrib("PlateID"), MAXFORMIDSTRING);
									g_util.strcpyexx(plan->m_sheetsortingpositioncolorspecific[issues][eds][sheets][0][2], xml.GetAttrib("SortingPosition"), MAXFORMIDSTRING);
									bIsProcessColor = TRUE;
								}
								if (sCol.CompareNoCase("black") == 0 || sCol.CompareNoCase("k") == 0) {
									g_util.strcpyexx(plan->m_sheetpresscylinder[issues][eds][sheets][0][3], xml.GetAttrib("Name"), MAXPRESSSTRING);
									g_util.strcpyexx(plan->m_sheetformID[issues][eds][sheets][0][3], xml.GetAttrib("FormID"), MAXFORMIDSTRING);
									g_util.strcpyexx(plan->m_sheetplateID[issues][eds][sheets][0][3], xml.GetAttrib("PlateID"), MAXFORMIDSTRING);
									g_util.strcpyexx(plan->m_sheetsortingpositioncolorspecific[issues][eds][sheets][0][3], xml.GetAttrib("SortingPosition"), MAXFORMIDSTRING);
									bIsProcessColor = TRUE;
								}
								if (bIsProcessColor == FALSE) {
									g_util.strcpyexx(plan->m_sheetpresscylinder[issues][eds][sheets][0][4 + nNumberOfPlateSpotColors], xml.GetAttrib("Name"), MAXPRESSSTRING);
									g_util.strcpyexx(plan->m_sheetformID[issues][eds][sheets][0][4 + nNumberOfPlateSpotColors], xml.GetAttrib("FormID"), MAXFORMIDSTRING);
									g_util.strcpyexx(plan->m_sheetplateID[issues][eds][sheets][0][4 + nNumberOfPlateSpotColors], xml.GetAttrib("PlateID"), MAXFORMIDSTRING);
									g_util.strcpyexx(plan->m_sheetsortingpositioncolorspecific[issues][eds][sheets][0][4 + nNumberOfPlateSpotColors], xml.GetAttrib("SortingPosition"), MAXFORMIDSTRING);
									nNumberOfPlateSpotColors++;
								}
							}
						}
						xml.OutOfElem(); // </PressCylinderFront>
					}

					xml.OutOfElem(); // </SheetFrontItems>

					plan->m_sheetignoreback[issues][eds][sheets] = FALSE;

					if (xml.FindElem(_T("SheetBackItems")) == false) {
						if (g_prefs.m_mayignorebacksheet == FALSE) {
							sErrMsg.Format("Error parsing file %s - <SheetBackItems> element not found", sFullInputPath);
							return PARSEFILE_ERROR;
						}
						plan->m_sheetignoreback[issues][eds][sheets] = TRUE;
					}

					if (plan->m_sheetignoreback[issues][eds][sheets] == FALSE) {

						g_util.strcpyexx(plan->m_sheetsortingposition[issues][eds][sheets][1], xml.GetAttrib("SortingPosition"), MAXSORTINGPOS);		// Validate!!
						g_util.strcpyexx(plan->m_sheetpresstower[issues][eds][sheets][1], xml.GetAttrib("PressTower"), MAXPRESSTOWERSTRING);		// Validate!!
						g_util.strcpyexx(plan->m_sheetpresszone[issues][eds][sheets][1], xml.GetAttrib("PressZone"), MAXPRESSSTRING);		// Validate!!
						g_util.strcpyexx(plan->m_sheetpresshighlow[issues][eds][sheets][1], xml.GetAttrib("PressHighLow"), MAXPRESSSTRING);		// Validate!!1
						plan->m_sheetactivecopies[issues][eds][sheets][1] = atoi(xml.GetAttrib("ActiveCopies"));
						xml.IntoElem();	// Into SheetBackItems

						nPagesOnThisSide = 0;
						//if (g_prefs.m_SwapBack && nPagesOnPlate == 2)
						//	nPagesOnThisSide = 1;

						while (xml.FindElem(_T("SheetBackItem"))) {

							g_util.strcpyexx(plan->m_oldsheetpagename[issues][eds][sheets][1][nPagesOnThisSide], xml.GetAttrib("PageName"), MAXPAGEIDLEN);
							plan->m_oldsheetsectionID[issues][eds][sheets][1][nPagesOnThisSide] = g_prefs.GetSectionID(xml.GetAttrib("Section"));

							CString sThisPageID = xml.GetAttrib("PageID");
							g_util.strcpyexx(plan->m_sheetpageID[issues][eds][sheets][1][nPagesOnThisSide], sThisPageID, MAXPAGEIDLEN);
							g_util.strcpyexx(plan->m_sheetmasterpageID[issues][eds][sheets][1][nPagesOnThisSide], xml.GetAttrib("MasterPageID"), MAXPAGEIDLEN);
							int Xpos = atoi(xml.GetAttrib("PosX"));
							if (Xpos < 1 || Xpos > 16) {
								sErrMsg.Format("Error parsing file %s - PosX %d atttribute out of range ", Xpos, sFullInputPath);
								return PARSEFILE_ERROR;
							}

							int Ypos = atoi(xml.GetAttrib("PosY"));
							if (Ypos < 1 || Ypos > 16) {
								sErrMsg.Format("Error parsing file %s - PosY %d atttribute out of range ", Ypos, sFullInputPath);
								return PARSEFILE_ERROR;
							}
							plan->m_sheetpageposition[issues][eds][sheets][1][nPagesOnThisSide] = GetPagePosition(nPagesOnPlate, Xpos, Ypos);
							/*	if (g_prefs.m_SwapBack && nPagesOnPlate == 2)
								plan->m_sheetpageposition[issues][eds][sheets][1][nPagesOnThisSide] = GetPagePosition(nPagesOnPlate, 3 - Xpos, Ypos);

							if (g_prefs.m_SwapBack && nPagesOnPlate == 2)
								nPagesOnThisSide--;
							else*/
								nPagesOnThisSide++;

							// Get a default sectionID from a non-dummy for later dummy creation
							plan->m_sheetdefaultsectionID[issues][eds][sheets][1] = GetPageSectionID(plan, eds, sThisPageID);

						}

						if (nPagesOnThisSide == 0 && g_prefs.m_mayignorebacksheet)
							plan->m_sheetignoreback[issues][eds][sheets] = TRUE;

						if (plan->m_sheetignoreback[issues][eds][sheets] == FALSE) {

							for (int ii = 0; ii < nPagesOnPlate; ii++) {
								if (plan->m_sheetpageposition[issues][eds][sheets][1][ii] == 0) {
									// Find free pageposition for dummy page

									BOOL bFreePosition = TRUE;
									int nFreePosition = 0;
									for (int jj = 1; jj <= nPagesOnPlate; jj++) {

										bFreePosition = TRUE;
										nFreePosition = 0;

										for (int iii = 0; iii < nPagesOnPlate; iii++) {
											if (plan->m_sheetpageposition[issues][eds][sheets][1][iii] == jj) {
												bFreePosition = FALSE;
												break;
											}
										}
										if (bFreePosition) {
											nFreePosition = jj;
											break;
										}
									}
									if (bFreePosition && nFreePosition > 0)
										plan->m_sheetpageposition[issues][eds][sheets][1][ii] = nFreePosition;

								}
							}

							if (nPagesOnThisSide == 0) {
								sErrMsg.Format("Error parsing file %s - Number of <SheetBackItem> elements must be greater than zero", sFullInputPath);
								return FALSE;
							}
							/*							if (nPagesOnThisSide > nPagesOnPlate) {
															sprintf(szErrMsg,"Error parsing file %s - Number of <SheetBackItem> elements greater than PagesOnPlate attribute ", sFullInputPath);
															return PARSEFILE_ERROR;
														}
							*/
							if (xml.FindElem(_T("PressCylindersBack"))) {
								xml.IntoElem();	// Into PressCylinderBack											// Clear spot

								int nNumberOfPlateColors = 0;
								int nNumberOfPlateSpotColors = 0;
								g_util.strcpyexx(plan->m_sheetpresscylinder[issues][eds][sheets][1][4], "", MAXPRESSSTRING);
								g_util.strcpyexx(plan->m_sheetformID[issues][eds][sheets][1][4], "", MAXFORMIDSTRING);
								g_util.strcpyexx(plan->m_sheetplateID[issues][eds][sheets][1][4], "", MAXFORMIDSTRING);
								g_util.strcpyexx(plan->m_sheetsortingpositioncolorspecific[issues][eds][sheets][0][4], "", MAXFORMIDSTRING);


								while (xml.FindElem(_T("PressCylinderBack"))) {
									// Set colors to fixed places - now to handle spots??
									CString sCol = xml.GetAttrib("Color");
									if (sCol != "")
										nNumberOfPlateColors++;

									if (g_prefs.m_pdfpages) {
										g_util.strcpyexx(plan->m_sheetpresscylinder[issues][eds][sheets][1][0], xml.GetAttrib("Name"), MAXPRESSSTRING);
										g_util.strcpyexx(plan->m_sheetformID[issues][eds][sheets][1][0], xml.GetAttrib("FormID"), MAXFORMIDSTRING);
										g_util.strcpyexx(plan->m_sheetplateID[issues][eds][sheets][1][0], xml.GetAttrib("PlateID"), MAXFORMIDSTRING);
										g_util.strcpyexx(plan->m_sheetsortingpositioncolorspecific[issues][eds][sheets][1][0], xml.GetAttrib("SortingPosition"), MAXFORMIDSTRING);

									}
									else {
										BOOL bIsProcessColor = FALSE;
										if (sCol.CompareNoCase("cyan") == 0 || sCol.CompareNoCase("c") == 0) {
											g_util.strcpyexx(plan->m_sheetpresscylinder[issues][eds][sheets][1][0], xml.GetAttrib("Name"), MAXPRESSSTRING);
											g_util.strcpyexx(plan->m_sheetformID[issues][eds][sheets][1][0], xml.GetAttrib("FormID"), MAXFORMIDSTRING);
											g_util.strcpyexx(plan->m_sheetplateID[issues][eds][sheets][1][0], xml.GetAttrib("PlateID"), MAXFORMIDSTRING);
											g_util.strcpyexx(plan->m_sheetsortingpositioncolorspecific[issues][eds][sheets][1][0], xml.GetAttrib("SortingPosition"), MAXFORMIDSTRING);

											bIsProcessColor = TRUE;
										}
										if (sCol.CompareNoCase("magenta") == 0 || sCol.CompareNoCase("m") == 0) {
											g_util.strcpyexx(plan->m_sheetpresscylinder[issues][eds][sheets][1][1], xml.GetAttrib("Name"), MAXPRESSSTRING);
											g_util.strcpyexx(plan->m_sheetformID[issues][eds][sheets][1][1], xml.GetAttrib("FormID"), MAXFORMIDSTRING);
											g_util.strcpyexx(plan->m_sheetplateID[issues][eds][sheets][1][1], xml.GetAttrib("PlateID"), MAXFORMIDSTRING);
											g_util.strcpyexx(plan->m_sheetsortingpositioncolorspecific[issues][eds][sheets][1][1], xml.GetAttrib("SortingPosition"), MAXFORMIDSTRING);

											bIsProcessColor = TRUE;
										}
										if (sCol.CompareNoCase("yellow") == 0 || sCol.CompareNoCase("y") == 0) {
											g_util.strcpyexx(plan->m_sheetpresscylinder[issues][eds][sheets][1][2], xml.GetAttrib("Name"), MAXPRESSSTRING);
											g_util.strcpyexx(plan->m_sheetformID[issues][eds][sheets][1][2], xml.GetAttrib("FormID"), MAXFORMIDSTRING);
											g_util.strcpyexx(plan->m_sheetplateID[issues][eds][sheets][1][2], xml.GetAttrib("PlateID"), MAXFORMIDSTRING);
											g_util.strcpyexx(plan->m_sheetsortingpositioncolorspecific[issues][eds][sheets][1][2], xml.GetAttrib("SortingPosition"), MAXFORMIDSTRING);

											bIsProcessColor = TRUE;
										}
										if (sCol.CompareNoCase("black") == 0 || sCol.CompareNoCase("k") == 0) {
											g_util.strcpyexx(plan->m_sheetpresscylinder[issues][eds][sheets][1][3], xml.GetAttrib("Name"), MAXPRESSSTRING);
											g_util.strcpyexx(plan->m_sheetformID[issues][eds][sheets][1][3], xml.GetAttrib("FormID"), MAXFORMIDSTRING);
											g_util.strcpyexx(plan->m_sheetplateID[issues][eds][sheets][1][3], xml.GetAttrib("PlateID"), MAXFORMIDSTRING);
											g_util.strcpyexx(plan->m_sheetsortingpositioncolorspecific[issues][eds][sheets][1][3], xml.GetAttrib("SortingPosition"), MAXFORMIDSTRING);

											bIsProcessColor = TRUE;
										}
										if (bIsProcessColor == FALSE) {
											g_util.strcpyexx(plan->m_sheetpresscylinder[issues][eds][sheets][1][4 + nNumberOfPlateSpotColors], xml.GetAttrib("Name"), MAXPRESSSTRING);
											g_util.strcpyexx(plan->m_sheetformID[issues][eds][sheets][1][4 + nNumberOfPlateSpotColors], xml.GetAttrib("FormID"), MAXFORMIDSTRING);
											g_util.strcpyexx(plan->m_sheetplateID[issues][eds][sheets][1][4 + nNumberOfPlateSpotColors], xml.GetAttrib("PlateID"), MAXFORMIDSTRING);
											g_util.strcpyexx(plan->m_sheetsortingpositioncolorspecific[issues][eds][sheets][1][4 + nNumberOfPlateSpotColors], xml.GetAttrib("SortingPosition"), MAXFORMIDSTRING);

											++nNumberOfPlateSpotColors;
										}
									}
								}
								xml.OutOfElem(); // </PressCylinderBack>
							}
						}

						xml.OutOfElem(); // </SheetBackItems>

					}

					xml.OutOfElem();	// </Sheet>
					plan->m_numberofsheets[issues][eds]++;
				}


				xml.OutOfElem(); //</Sheets>
			}
			plan->m_numberofeditions++;
			xml.OutOfElem(); //</Edition>
		}

		if (plan->m_numberofeditions == 0) {
			sErrMsg.Format("Error parsing file %s - no <Edition> elements  found in <Editions> element", sFullInputPath);
			return PARSEFILE_ERROR;
		}

		xml.OutOfElem(); // </Editions>

		plan->m_numberofissues++;
		xml.OutOfElem(); // </Issue>
	}


	if (plan->m_numberofissues == 0) {
		sErrMsg.Format("Error parsing file %s - no <Issue> element found in <Issues> element", sFullInputPath);
		return PARSEFILE_ERROR;
	}

	xml.OutOfElem(); // </Issues>
	xml.OutOfElem(); // </Publication>

	if (plan->m_numberofissues > 1)
		sInfo.Format("%d issues imported - %d pages in total", plan->m_numberofissues, nPagesImported);
	else {
		if (plan->m_numberofeditions > 1)
			sInfo.Format("%d editions imported - %d pages in total", plan->m_numberofeditions, nPagesImported);
		else
			sInfo.Format("%d pages imported", nPagesImported);
	}


	if (bPagesOnly == FALSE)
		SetMaxPageColors(plan);

	return bPagesOnly == TRUE ? PARSEFILE_OK_PAGESONLY : PARSEFILE_OK;
}

int GetPagePosition(int nPagesOnPlate, int Xpos, int Ypos)
{
	switch (nPagesOnPlate) {
	case 1:
		return 1;
	case 2:
		return max(Xpos, Ypos);
	case 3:
		return Xpos;
	case 4:
		return (Ypos - 1) * 2 + Xpos;
	case 6:
		return (Ypos - 1) * 3 + Xpos;
	case 8:
	case 12:
	case 16:
		return (Ypos - 1) * 4 + Xpos;
	case 32:
	case 64:
		return (Ypos - 1) * 8 + Xpos;
	default:
		return 1;
	}
}

int GetPageColors(PLANSTRUCT *plan, CString sPageID)
{
	int nColors = 0;

	for (int ed = 0; ed < plan->m_numberofeditions; ed++) {
		for (int section = 0; section < plan->m_numberofsections[0][ed]; section++) {
			for (int page = 0; page < plan->m_numberofpages[0][ed][section]; page++) {
				CString s(plan->m_pageIDs[0][ed][section][page]);
				if (sPageID.CompareNoCase(s) == 0) {
					return  plan->m_numberofpagecolors[0][ed][section][page];
				}
			}
		}
	}

	return 0;
}

void SetMaxPageColors(PLANSTRUCT *plan)
{
	plan->m_numberofissues = 1;
	for (int ed = 0; ed < plan->m_numberofeditions; ed++) {
		for (int sheet = 0; sheet < plan->m_numberofsheets[0][ed]; sheet++) {
			for (int side = 0; side < 2; side++) {
				plan->m_sheetmaxcolors[0][ed][sheet][side] = 0;
				for (int pos = 0; pos < plan->m_sheetpagesonplate[0][ed][sheet]; pos++) {
					int nPageColors = GetPageColors(plan, plan->m_sheetpageID[0][ed][sheet][side][pos]);
					if (nPageColors > plan->m_sheetmaxcolors[0][ed][sheet][side])
						plan->m_sheetmaxcolors[0][ed][sheet][side] = nPageColors;
				}
			}
		}
	}
}

int  GetPageSectionID(PLANSTRUCT *plan, int ed, CString sPageID)
{

	int nSections = plan->m_numberofsections[0][ed] > 0 ? plan->m_numberofsections[0][ed] : 1;
	for (int section = 0; section < nSections; section++) {
		int nPages = plan->m_numberofpages[0][ed][section] > 0 ? plan->m_numberofpages[0][ed][section] : 1;
		for (int page = 0; page < plan->m_numberofpages[0][ed][section]; page++) {
			CString s(plan->m_pageIDs[0][ed][section][page]);
			if (sPageID.CompareNoCase(s) == 0) {
				return  plan->m_sectionIDlist[0][ed][section];
			}
		}
	}

	return 0;
}


void LogXMLcontent(CString sFileName, PLANSTRUCT *plan)
{

	g_util.Logprintf("Parsing result of file %s", sFileName);
	g_util.Logprintf("----------------------------------");

	g_util.Logprintf("Sender           : %s", plan->m_sender);
	g_util.Logprintf("Time             : %s", plan->m_updatetime);
	g_util.Logprintf("PlanID           : %s", plan->m_planID);
	g_util.Logprintf("PlanName         : %s", plan->m_planName);

	g_util.Logprintf("Press            : %s", g_prefs.GetPressName(plan->m_pressID));
	g_util.Logprintf("Presstime        : %s", plan->m_presstime);
	g_util.Logprintf("Version          : %d\r\n", plan->m_version);
	g_util.Logprintf("Publication      : %s (id %d)", g_prefs.GetPublicationName(plan->m_publicationID), plan->m_publicationID);
	g_util.Logprintf("Week reference   : %d", plan->m_weekreference);

	g_util.Logprintf("PubDate          : %.4d-%.2d-%.2d (YYYY-MM-DD)\r\n", plan->m_pubdate.GetYear(), plan->m_pubdate.GetMonth(), plan->m_pubdate.GetDay());
	g_util.Logprintf("Number of Issues : %d", plan->m_numberofissues);

	g_util.Logprintf("Issue            : %s (id %d)", g_prefs.GetIssueName(plan->m_issueIDlist[0]), plan->m_issueIDlist[0]);
	g_util.Logprintf("Editions in issue: %d", plan->m_numberofeditions);
	for (int edition = 0; edition < plan->m_numberofeditions; edition++) {
		g_util.Logprintf("  Edition        : %s (id %d)", g_prefs.GetEditionName(plan->m_editionIDlist[0][edition]), plan->m_editionIDlist[0][edition]);
		g_util.Logprintf("  Copies         : %d", plan->m_numberofcopies[0][edition][0]);
		g_util.Logprintf("  Sections in edition : %d", plan->m_numberofsections[0][edition]);
		for (int section = 0; section < plan->m_numberofsections[0][edition]; section++) {
			g_util.Logprintf("    Section      : %s (id %d)", g_prefs.GetSectionName(plan->m_sectionIDlist[0][edition][section]), plan->m_sectionIDlist[0][edition][section]);
			g_util.Logprintf("    Pages in section: %d", plan->m_numberofpages[0][edition][section]);
			for (int page = 0; page < plan->m_numberofpages[0][edition][section]; page++) {
				CString sColors = "";
				for (int color = 0; color < plan->m_numberofpagecolors[0][edition][section][page]; color++) {
					sColors += g_prefs.GetColorName(plan->m_pagecolorIDlist[0][edition][section][page][color]) + "  ";
				}
				g_util.Logprintf("      Pagename : %s  Pagination %d Colors (%s) PageID %s MasterPageID %s", plan->m_pagenames[0][edition][section][page],
					plan->m_pagepagination[0][edition][section][page],
					sColors, plan->m_pageIDs[0][edition][section][page], plan->m_pagemasterIDs[0][edition][section][page]);
			}
		}

		for (int sheet = 0; sheet < plan->m_numberofsheets[0][edition]; sheet++) {
			CString sPagePos = "";
			for (int pos = 0; pos < plan->m_sheetpagesonplate[0][edition][sheet]; pos++) {
				CString s;
				s.Format("%d: PageID %s MasterPageID %s", pos + 1, plan->m_sheetpageID[0][edition][sheet][0][pos], plan->m_sheetmasterpageID[0][edition][sheet][0][pos]);
				sPagePos += s;
			}
			g_util.Logprintf("    Front Sheet  : %d Pagepositions: %s   Max colors: %d\r\n", sheet + 1, sPagePos, plan->m_sheetmaxcolors[0][edition][sheet][0]);
			sPagePos = "";
			for (int pos = 0; pos < plan->m_sheetpagesonplate[0][edition][sheet]; pos++) {
				CString s;
				s.Format("%d: PageID %s MasterPageID %s", pos + 1, plan->m_sheetpageID[0][edition][sheet][1][pos], plan->m_sheetmasterpageID[0][edition][sheet][1][pos]);
				sPagePos += s;
			}
			g_util.Logprintf("    Back Sheet   : %d Pagepositions: %s  Max colors: %d", sheet + 1, sPagePos, plan->m_sheetmaxcolors[0][edition][sheet][1]);
		}
	}
}



