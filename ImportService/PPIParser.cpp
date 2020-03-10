#include "stdafx.h"
#include <afxmt.h>
#include <afxtempl.h>
#include <direct.h>
#include <iostream>
#include "Defs.h"
#include "PlanDataDefs.h"
#include "Markup.h"
#include "Utils.h"
#include "Prefs.h"
#include "PPIPrefs.h"
#include "PPIParser.h"
#include "DatabaseManager.h"

extern CPrefs g_prefs;
extern CUtils g_util;

CPPIPrefs g_ppiprefs;
PPINAMETRANSLATIONLIST g_ppitranslations;

BOOL ProcessPPIFile(CDatabaseManager &DB, CString sFullInputPath, CString &sOutputFile, CString &sErrorMessage)
{
	PlanData pageplan;
	pageplan.planName = g_util.GetFileName(sFullInputPath, true);
	pageplan.updatetime = CTime::GetCurrentTime();

	g_util.Logprintf("INFO: ");
	g_util.Logprintf("INFO: Analyzing PPI file %s..", g_util.GetFileName(sFullInputPath));

	if (g_ppiprefs.m_archivefolder != "")
		::CopyFile(sFullInputPath, g_ppiprefs.m_archivefolder + _T("\\") + g_util.GetFileName(sFullInputPath), FALSE);

	int nPagesImported = 0;
	CString sInfo = _T("");
	CString sErrorMessageParser = _T("");

	int nParseResult = PPIParseXML(sFullInputPath, pageplan, g_ppiprefs.m_fixededition, g_ppiprefs.m_fixedsection, nPagesImported, sInfo, sErrorMessageParser);

	if (nParseResult != PARSEFILE_OK) {
		sErrorMessage = sErrorMessageParser;
		return FALSE;
	}
	// Parsing went ok
	// Go find setup data for this PPI product name

	// Sort pages according to pagination
	for (int ed = 0; ed < pageplan.editionList->GetCount(); ed++) {
		PlanDataEdition *pEdition = &pageplan.editionList->GetAt(ed);
		for (int sec = 0; sec < pEdition->sectionList->GetCount(); sec++) {
			PlanDataSection *pSection = &pEdition->sectionList->GetAt(sec);
			for (int i1 = 0; i1 < pSection->pageList->GetCount() - 1; i1++) {
				for (int i2 = i1; i2 < pSection->pageList->GetCount() - 1; i2++) {
					if (pSection->pageList->GetAt(i1).pagination > pSection->pageList->GetAt(i2).pagination) {
						// swap pages
						PlanDataPage *pPageCopy = new PlanDataPage();
						pSection->pageList->GetAt(i1).ClonePage(pPageCopy);
						pSection->pageList->GetAt(i1).comment = pSection->pageList->GetAt(i2).comment;
						pSection->pageList->GetAt(i1).masterEdition = pSection->pageList->GetAt(i2).masterEdition;
						pSection->pageList->GetAt(i1).pageID = pSection->pageList->GetAt(i2).pageID;
						pSection->pageList->GetAt(i1).masterPageID = pSection->pageList->GetAt(i2).masterPageID;
						pSection->pageList->GetAt(i1).pageIndex = pSection->pageList->GetAt(i2).pageIndex;
						pSection->pageList->GetAt(i1).pagination = pSection->pageList->GetAt(i2).pagination;
						pSection->pageList->GetAt(i1).pageName = pSection->pageList->GetAt(i2).pageName;
						pSection->pageList->GetAt(i1).uniquePage = pSection->pageList->GetAt(i2).uniquePage;

						pSection->pageList->GetAt(i2).comment = pPageCopy->comment;
						pSection->pageList->GetAt(i2).masterEdition = pPageCopy->masterEdition;
						pSection->pageList->GetAt(i2).pageID = pPageCopy->pageID;
						pSection->pageList->GetAt(i2).masterPageID = pPageCopy->masterPageID;
						pSection->pageList->GetAt(i2).pageIndex = pPageCopy->pageIndex;
						pSection->pageList->GetAt(i2).pagination = pPageCopy->pagination;
						pSection->pageList->GetAt(i2).pageName = pPageCopy->pageName;
						pSection->pageList->GetAt(i2).uniquePage = pPageCopy->uniquePage;
					}
				}
			}
		}
	}


	return TRUE;
}


CString LookupPPIName(CString sPPIProduct, CString sPPIEdition)
{
	sPPIProduct.MakeUpper();
	sPPIEdition.MakeUpper();
	for(int i=0; i< g_ppitranslations.GetCount(); i++) {
		if (g_ppitranslations[i].m_PPIProduct == sPPIProduct && g_ppitranslations[i].m_PPIEdition == sPPIEdition)
			return g_ppitranslations[i].m_Publication;
		if ((g_ppitranslations[i].m_PPIProduct == _T("") || g_ppitranslations[i].m_PPIProduct == _T("*")) && g_ppitranslations[i].m_PPIEdition == sPPIEdition)
			return g_ppitranslations[i].m_Publication;
		if ((g_ppitranslations[i].m_PPIEdition == _T("") || g_ppitranslations[i].m_PPIEdition == _T("*")) && g_ppitranslations[i].m_PPIProduct == sPPIProduct)
			return g_ppitranslations[i].m_Publication;
	}
	
	return sPPIProduct;
		
}

int PPIParseXML(CString sFullInputPath, PlanData &pageplan, CString sFixedEditionNames, CString sFixedSectionName, int &nPagesImported, CString &sInfo, CString &sErrorMessage)
{
	nPagesImported = 0;
	sErrorMessage = _T("");
	sInfo = _T("");

	int nPageID = 1;


	CString str;
	CFile file;
	if (!file.Open(sFullInputPath, CFile::modeRead)) {
		sErrorMessage.Format("Cannot open XML file %s for reading", sFullInputPath);
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
		sErrorMessage.Format("Error converting file %s to string (may contain binary data)", sFullInputPath);
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
		sErrorMessage.Format("Error parsing file %s - %s", sFullInputPath, csError);
		return PARSEFILE_ERROR;
	}

	if (xml.FindElem(_T("format")) == false) {
		sErrorMessage.Format("Error parsing file %s - <format> element not found", sFullInputPath);
		return PARSEFILE_ERROR;
	}

	xml.IntoElem();

	pageplan.version = 1;
	pageplan.publicationDate = g_ppiprefs.PPIDateToDateTime(xml.GetAttrib("date"));
	CString sPPIProduct = xml.GetAttrib("product");

	

	if (xml.FindElem(_T("pages")) == false) {
		sErrorMessage.Format("Error parsing file %s - <page> element not found", sFullInputPath);
		return PARSEFILE_ERROR;
	}

	// OBS - use PPI edition names as pub..
	CString sPPIEditionName = xml.GetAttrib("edition");

	
	pageplan.publicationName = LookupPPIName(sPPIProduct, sPPIEditionName);

	pageplan.publicationComment = xml.GetAttrib("part");

	if (pageplan.publicationName == _T("") || pageplan.publicationDate.GetYear() < 2000) {
		sErrorMessage.Format("product/date attributes not found in file %s", sFullInputPath);
		return PARSEFILE_ERROR;
	}

	xml.IntoElem();

	int nPageCounter = 0;
	while (xml.FindElem(_T("page"))) {
		nPageCounter++;
		PlanDataEdition *pEdition = NULL;
		PlanDataSection *pSection = NULL;

		CString sVersion = xml.GetAttrib("version"); // edition
		CString sEdition = (sFixedEditionNames == _T("") && sVersion != _T("")) ? sVersion : sFixedEditionNames;
		
		// see if this is a new edition name
		for (int ed = 0; ed < pageplan.editionList->GetCount(); ed++) {			
		{
				if (sEdition.CompareNoCase(pageplan.editionList->GetAt(ed).editionName) == 0) {
					pEdition = &pageplan.editionList->GetAt(ed);
					if (sFixedSectionName != "")
						pSection = &pEdition->sectionList->GetAt(0);
					break;
				}
			}
		}

		if (pEdition == NULL) {
			pEdition = new PlanDataEdition(sEdition);
			pSection = new PlanDataSection("Dummy");
			PlanDataPress *pPress = new PlanDataPress("Dummy");
			pEdition->pressList->Add(*pPress);
			pEdition->sectionList->Add(*pSection);
			pageplan.editionList->Add(*pEdition);
		}
		pEdition->editionComment = sVersion;

		CString sSection = xml.GetAttrib("section"); // dummy read - section is NOT traditional section for us

		int nPagination = atoi(xml.GetAttrib("pagina")); // pagina is printed pagenumber

		if (nPagination == 0)
			continue;

		// See if page already registered (possibly with other incoming version name)
		BOOL hasExistingPage = FALSE;
		for (int pg = 0; pg < pSection->pageList->GetCount(); pg++) {
			if (pSection->pageList->GetAt(pg).pagination == nPagination) {
				hasExistingPage = TRUE;
				break;
			}
		}

		if (hasExistingPage == FALSE) {
			PlanDataPage *pPage = new PlanDataPage(g_util.Int2String(nPagination));
			PlanDataPageSeparation *pColor = new PlanDataPageSeparation("PDF");

			pPage->uniquePage = PAGEUNIQUE_UNIQUE;
			pPage->pageIndex = nPagination;
			pPage->pagination = nPagination;
			pPage->pageID = g_util.Int2String(nPageID);
			pPage->masterPageID = pPage->pageID;
			pPage->comment = sSection;			
			pPage->colorList->Add(*pColor);
			pSection->pageList->Add(*pPage);
			nPageID++;
		}
	}
	xml.OutOfElem();
	xml.OutOfElem();


	return PARSEFILE_OK;
}