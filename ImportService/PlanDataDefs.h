#pragma once
#include "stdafx.h"
#include "Utils.h"


static TCHAR szPageTypes[4][14] = { "Normal","Panorama","Antipanorama","Dummy" };
static TCHAR szPageUniques[3][8] = { "false","true","force" };

typedef struct	 {
	CString m_zone;
	CString m_location;
} ZONELOCATIONSTRUCT;

typedef struct  {
	CString		m_Publication;
	CString		m_mineditions;
} NPEXTRAEDITION;
typedef CArray <NPEXTRAEDITION, NPEXTRAEDITION&> NPEXTRAEDITIONLIST;


////////////////////////////////

/// <summary>
///  Level 5 - page separation
/// </summary>
/// 
class PlanDataPageSeparation
{
public:
	CString colorName;
public:
	PlanDataPageSeparation()
	{
	}

	PlanDataPageSeparation(CString color)
	{
		colorName = color;
	}
};

typedef CArray <PlanDataPageSeparation, PlanDataPageSeparation&> ListPlanDataPageSeparation;


/// <summary>
///  Level 4 - pages
/// </summary>
/// 
class PlanDataPage
{
public:

	CString pageName;
	CString fileName;
	CString pageID;
	int		pageType;
	int		pagination;
	int		pageIndex;
	CString comment;
	ListPlanDataPageSeparation *colorList;
	int		uniquePage;
	CString masterPageID;

	BOOL	approve;
	BOOL	hold;
	int		priority;
	int		version;
	CString masterEdition;
	CString miscstring1;
	CString miscstring2;
	int	miscint;

public:
	PlanDataPage()
	{
		pageName = "";
		fileName = "";
		pageID = "";
		pageType = PAGETYPE_NORMAL;
		pagination = 0;
		pageIndex = 0;
		comment = "";

		masterPageID = "";
		masterEdition = "";
		approve = FALSE;
		hold = FALSE;
		priority = 50;
		version = 0;
		uniquePage = PAGEUNIQUE_UNIQUE;
		miscstring1 = "";
		miscstring2 = "";
		miscint = 0;

		colorList = new ListPlanDataPageSeparation();

	}

	PlanDataPage(CString sPageName)
	{
		pageName = sPageName;
		fileName = "";
		pageID = "";
		pageType = PAGETYPE_NORMAL;
		pagination = 0;
		pageIndex = 0;
		comment = "";

		masterPageID = "";
		approve = FALSE;
		hold = FALSE;
		priority = 50;
		version = 0;
		uniquePage = PAGEUNIQUE_UNIQUE;
		masterEdition = "";
		miscstring1 = "";
		miscstring2 = "";
		miscint = 0;

		colorList = new ListPlanDataPageSeparation();
	}


	BOOL ClonePage(PlanDataPage *pCopy)
	{
		if (pCopy == NULL)
			return FALSE;

		pCopy->approve = approve;
		pCopy->comment = comment;
		pCopy->fileName = fileName;
		pCopy->hold = hold;
		pCopy->masterEdition = masterEdition;
		pCopy->masterPageID = masterPageID;
		pCopy->miscint = miscint;
		pCopy->miscstring1 = miscstring1;
		pCopy->miscstring2 = miscstring2;
		pCopy->pageID = pageID;
		pCopy->pageIndex = pageIndex;
		pCopy->pageName = pageName;
		pCopy->pageType = pageType;
		pCopy->pagination = pagination;
		pCopy->priority = priority;
		pCopy->uniquePage = uniquePage;
		pCopy->version = version;
		pCopy->colorList->RemoveAll();

		for (int color = 0; color < colorList->GetCount(); color++) {
			PlanDataPageSeparation *org = &colorList->GetAt(color);
			PlanDataPageSeparation *sep = new PlanDataPageSeparation(org->colorName);
			pCopy->colorList->Add(*sep);
		}

		return TRUE;
	}


};

typedef CArray <PlanDataPage, PlanDataPage&> ListPlanDataPage;

/// <summary>
///  Group level 3 - sections
/// </summary>
class PlanDataSection
{
public:
	CString sectionName;

	ListPlanDataPage *pageList;
	int pagesInSection;
public:
	PlanDataSection()
	{
		sectionName = "";
		pagesInSection = 0;

		//pageList.RemoveAll();
		pageList = new ListPlanDataPage();
	}

	PlanDataSection(CString sSectionName)
	{
		sectionName = sSectionName;
		pagesInSection = 0;
		//pageList.RemoveAll();
		pageList = new ListPlanDataPage();
	}

	PlanDataPage *GetPageObject(CString sPageName)
	{
		for (int i = 0; i < pageList->GetCount(); i++) {
			PlanDataPage *p = &pageList->GetAt(i);
			if (sPageName.CompareNoCase(p->pageName) == 0)
				return p;
		}

		return NULL;
	}

	PlanDataPage *GetPageObject(int nPageIndex)
	{
		for (int i = 0; i < pageList->GetCount(); i++) {
			PlanDataPage *p = &pageList->GetAt(i);
			if (nPageIndex == p->pageIndex)
				return p;
		}

		return NULL;
	}
};

typedef CArray <PlanDataSection, PlanDataSection&> ListPlanDataSection;

/// <summary>
/// Sheet level 4 - sheet press cylinder
/// </summary>
class PlanDataSheetPressCylinder
{
public:

	CString pressCylinder;
	CString colorName;
	CString formID;
	CString sortingPosition;
public:
	PlanDataSheetPressCylinder()
	{
		pressCylinder = "";
		colorName = "K";
		formID = "";
		sortingPosition = "";
	}
};

typedef CArray <PlanDataSheetPressCylinder, PlanDataSheetPressCylinder&> ListPlanDataSheetPressCylinder;

/// <summary>
/// Sheet level 3 - sheet item (page)
/// </summary>
class PlanDataSheetItem
{
public:
	int pagePositionX;
	int pagePositionY;
	// CString pageName;
	CString pageID;
	CString masterPageID;

public:
	PlanDataSheetItem()
	{
		pagePositionX = 1;
		pagePositionY = 1;
		pageID = "";
		masterPageID = "";
	}
};

typedef CArray <PlanDataSheetItem, PlanDataSheetItem&> ListPlanDataSheetItem;


/// <summary>
/// Sheet level 3 - sheet side 
/// </summary>
class PlanDataSheetSide
{
public:
	CString SortingPosition;
	CString PressTower;
	CString PressZone;
	CString PressHighLow;
	int		ActiveCopies;

	ListPlanDataSheetItem *sheetItems;
	ListPlanDataSheetPressCylinder  *pressCylinders;

public:
	PlanDataSheetSide()
	{
		SortingPosition = "";
		PressTower = "";
		PressZone = "";
		PressHighLow = "";
		ActiveCopies = 1;

		sheetItems = new ListPlanDataSheetItem();
		pressCylinders = new ListPlanDataSheetPressCylinder();
	}
};

typedef CArray <PlanDataSheetSide, PlanDataSheetSide&> ListPlanDataSheetSide;

/// <summary>
/// Sheet level 2 - sheet (both sides) 
/// </summary>
class PlanDataSheet
{
public:
	CString sheetName;
	CString templateName;
	int		pagesOnPlate;
	CString markGroups;

	BOOL	hasback;
	PlanDataSheetSide *frontSheet;
	PlanDataSheetSide *backSheet;

	int  pressSectionNumber;
	PlanDataSheet()
	{
		sheetName = "";
		templateName = "";
		pagesOnPlate = 1;
		markGroups = "";
		hasback = TRUE;

		frontSheet = new PlanDataSheetSide();
		backSheet = new PlanDataSheetSide();
		pressSectionNumber = 1;

	}
};

typedef CArray <PlanDataSheet, PlanDataSheet&> ListPlanDataSheet;


class PlanDataPress
{
public:
	CString	pressName;
	int		copies;
	CTime	pressRunTime;
	int		paperCopies;
	CString postalUrl;
	CString jdfPressName;
	BOOL	isLocalPress;
public:
	PlanDataPress() {
		pressName = "";
		copies = 1;
		pressRunTime = CTime(1975, 1, 1, 0, 0, 0);
		paperCopies = 0;
		postalUrl = "";
		jdfPressName = "";
		isLocalPress = TRUE;
	}

	PlanDataPress(CString sPressName) {
		pressName = sPressName;
		copies = 1;
		pressRunTime = CTime(1975, 1, 1, 0, 0, 0);
		paperCopies = 0;
		postalUrl = "";
		jdfPressName = "";
		isLocalPress = TRUE;

	}
};

typedef CArray <PlanDataPress, PlanDataPress&> ListPlanDataPress;


/// <summary>
///  Group level 2 - edtion
/// </summary>
class PlanDataEdition
{
public:
	CString				editionName;

	ListPlanDataPress	*pressList;
	ListPlanDataSection *sectionList;
	ListPlanDataSheet	*sheetList;

	BOOL				masterEdition;
	int					editionCopy;
	int					editionSequenceNumber;
	CString				editionComment;


	CString timedFrom;
	CString timedTo;
	BOOL	IsTimed;
	CString zoneMaster;
	CString	locationMaster;

public:
	PlanDataEdition()
	{
		editionName = "";
		editionCopy = 1;
		masterEdition = FALSE;
		pressList = new ListPlanDataPress();
		sectionList = new ListPlanDataSection();
		sheetList = new ListPlanDataSheet();
		editionSequenceNumber = 1;
		editionComment = _T("");
		IsTimed = TRUE;
		timedTo = _T("");
		timedFrom = _T("");
		zoneMaster = _T("");
		locationMaster = _T("");

	}


	PlanDataEdition(CString sEditionName)
	{
		editionName = sEditionName;
		masterEdition = FALSE;
		editionCopy = 0;

		pressList = new ListPlanDataPress();
		sectionList = new ListPlanDataSection();
		sheetList = new ListPlanDataSheet();

	}


	PlanDataPress *GetPressObject(CString sPressName)
	{
		for (int i = 0; i < pressList->GetCount(); i++) {
			PlanDataPress *p = &pressList->GetAt(i);
			if (sPressName.CompareNoCase(p->pressName) == 0)
				return p;
		}

		return NULL;
	}

	PlanDataSection *GetSectionObject(CString sSectionName)
	{
		for (int i = 0; i < sectionList->GetCount(); i++) {
			PlanDataSection *p = &sectionList->GetAt(i);
			if (sSectionName.CompareNoCase(p->sectionName) == 0)
				return p;
		}

		return NULL;
	}

	BOOL CloneEdition(PlanDataEdition *pCopy)
	{
		if (pCopy == NULL)
			return FALSE;
		pCopy->editionName = editionName;
		pCopy->editionCopy = editionCopy;
		pCopy->masterEdition = masterEdition;

		pCopy->timedFrom = timedFrom;
		pCopy->timedTo = timedTo;
		pCopy->IsTimed = IsTimed;
		pCopy->zoneMaster = zoneMaster;
		pCopy->editionComment = editionComment;

		for (int i = 0; i < pressList->GetCount(); i++) {
			PlanDataPress *pPress = &pressList->GetAt(i);
			PlanDataPress *pPressCopy = new PlanDataPress();
			pPressCopy->copies = pPress->copies;
			pPressCopy->isLocalPress = pPress->isLocalPress;
			pPressCopy->jdfPressName = pPress->jdfPressName;
			pPressCopy->paperCopies = pPress->paperCopies;
			pPressCopy->postalUrl = pPress->postalUrl;
			pPressCopy->pressName = pPress->pressName;
			pPressCopy->pressRunTime = pPress->pressRunTime;
			pCopy->pressList->Add(*pPressCopy);
		}
		/*
				for(int i=0; i<sectionList->GetCount(); i++) {
					PlanDataSection *pSection = &sectionList->GetAt(i);
					PlanDataSection *pSectionCopy = new PlanDataSection();
					pSectionCopy->pagesInSection = pSection->pagesInSection;
					pSectionCopy->sectionName = pSection->sectionName;
					pCopy->sectionList->Add(*pSectionCopy);
				}
		*/
		return TRUE;
	}



};

typedef CArray <PlanDataEdition, PlanDataEdition&> ListPlanDataEdition;

#define PLANSTATE_UNLOCKED 0
#define PLANSTATE_LOCKED   1


class PlanData
{
public:
	CString planName;
	CString planID;
	CString publicationName;
	CString publicationAlias;
	CString publicationComment;
	CTime publicationDate;
	int version;
	ListPlanDataEdition *editionList;
	int state;

	CString pagenameprefix;

	CTime   updatetime;
	CString sender;

	CStringArray arrSectionNames;
	CStringArray arrSectionLinkNames;

	CString defaultColors;

	CStringArray arrSectionColors;

	CString customername;
	CString customeralias;
public:

	PlanData()
	{
		pagenameprefix = "";
		planName = "";
		planID = "";
		publicationName = "";
		publicationAlias = "";
		publicationDate = CTime(1975, 1, 1, 0, 0, 0);
		updatetime = CTime(1975, 1, 1, 0, 0, 0);
		sender = "";
		version = 1;
		state = PLANSTATE_LOCKED;

		editionList = new ListPlanDataEdition();

		arrSectionNames.RemoveAll();
		arrSectionLinkNames.RemoveAll();
		arrSectionColors.RemoveAll();

		defaultColors = "CMYK";
		customername = "";
		customeralias = "";

		publicationComment = _T("");
	}

	CString GetRealSectionName(CString sSectionLinkName)
	{
		for (int i = 0; i < arrSectionLinkNames.GetCount(); i++) {
			if (sSectionLinkName.CompareNoCase(arrSectionLinkNames[i]) == 0)
				return arrSectionNames[i];
		}
		return "";
	}

	CString GetSectionColors(CString sSectionLinkName)
	{
		for (int i = 0; i < arrSectionLinkNames.GetCount(); i++) {
			if (sSectionLinkName.CompareNoCase(arrSectionLinkNames[i]) == 0)
				return arrSectionColors[i];
		}
		return defaultColors;
	}

	PlanDataEdition *GetEditionObject(CString sEditionName)
	{
		for (int i = 0; i < editionList->GetCount(); i++) {
			PlanDataEdition *p = &editionList->GetAt(i);
			if (sEditionName.CompareNoCase(p->editionName) == 0)
				return p;
		}

		return NULL;
	}

	PlanDataPage *FindMasterPage(CString sectionName, int pageIndex)
	{
		for (int edition = 0; edition < editionList->GetCount(); edition++) {
			PlanDataEdition *pEdition = &editionList->GetAt(edition);
			for (int section = 0; section < pEdition->sectionList->GetCount(); section++) {
				PlanDataSection *pSection = &pEdition->sectionList->GetAt(section);
				if (sectionName.CompareNoCase(pSection->sectionName) == 0) {
					for (int page = 0; page < pSection->pageList->GetCount(); page++) {
						PlanDataPage *pPage = &pSection->pageList->GetAt(page);
						if (pPage->pageIndex == pageIndex && pPage->pageName != "") {
							return pPage;
						}
					}
				}
			}
		}

		return NULL;
	}


	PlanDataPage *FindMasterPage(CString masterEditionName, CString sectionName, int pageIndex)
	{
		PlanDataEdition *pEdition = GetEditionObject(masterEditionName);
		if (pEdition == NULL)
			return NULL;

		for (int section = 0; section < pEdition->sectionList->GetCount(); section++) {
			PlanDataSection *pSection = &pEdition->sectionList->GetAt(section);
			if (sectionName.CompareNoCase(pSection->sectionName) == 0) {

				for (int page = 0; page < pSection->pageList->GetCount(); page++) {
					PlanDataPage *pPage = &pSection->pageList->GetAt(page);
					if (pPage->pageIndex == pageIndex && pPage->pageName != "") {
						return pPage;
					}
				}
			}
		}

		return NULL;
	}

	CString FindNextFreePageID()
	{
		int nMaxPageId = 0;
		for (int edition = 0; edition < editionList->GetCount(); edition++) {
			PlanDataEdition *pEdition = &editionList->GetAt(edition);
			for (int section = 0; section < pEdition->sectionList->GetCount(); section++) {
				PlanDataSection *pSection = &pEdition->sectionList->GetAt(section);
				for (int page = 0; page < pSection->pageList->GetCount(); page++) {
					PlanDataPage *pPage = &pSection->pageList->GetAt(page);
					if (atoi(pPage->pageID) > nMaxPageId) {

						nMaxPageId = atoi(pPage->pageID);
					}
				}
			}
		}

		CString s;
		s.Format("%d", nMaxPageId + 1);
		return s;
	}

	PlanDataPage *FindPageByFileName(CString fileName, CString sEditionName)
	{
		PlanDataEdition *pEdition = GetEditionObject(sEditionName);
		if (pEdition == NULL)
			return NULL;

		CString fileNameNoExt = fileName;
		fileNameNoExt.MakeUpper();

		// 0123456.pdf
		int n = fileNameNoExt.Find(".PDF");
		if (n != -1)
			fileNameNoExt = fileName.Mid(0, n);

		for (int section = 0; section < pEdition->sectionList->GetCount(); section++) {
			PlanDataSection *pSection = &pEdition->sectionList->GetAt(section);
			for (int page = 0; page < pSection->pageList->GetCount(); page++) {
				PlanDataPage *pPage = &pSection->pageList->GetAt(page);
				if (fileName.CompareNoCase(pPage->fileName) == 0) {
					return pPage;
				}
				if (fileNameNoExt.CompareNoCase(pPage->fileName) == 0) {
					return pPage;
				}
			}
		}

		return NULL;
	}

	void SortPlanPages()
	{
		for (int eds = 0; eds < editionList->GetCount(); eds++) {
			PlanDataEdition *planDataEdition = &editionList->GetAt(eds);
			for (int sections = 0; sections < planDataEdition->sectionList->GetCount(); sections++) {
				PlanDataSection *planDataSection = &planDataEdition->sectionList->GetAt(sections);

				for (int page = 0; page < planDataSection->pageList->GetCount() - 1; page++) {
					PlanDataPage *planDataPage = &planDataSection->pageList->GetAt(page);

					for (int page2 = page + 1; page2 < planDataSection->pageList->GetCount(); page2++) {
						PlanDataPage *planDataPage2 = &planDataSection->pageList->GetAt(page2);
						if (planDataPage->pagination > planDataPage2->pagination) {
							// Swap these pages
							PlanDataPage *pTemp = new PlanDataPage();
							planDataPage->ClonePage(pTemp);
							planDataPage2->ClonePage(planDataPage);
							pTemp->ClonePage(planDataPage2);
							delete pTemp;
						}
					}
				}
			}
		}
	}

	void DisposePlanData()
	{
		for (int edition = 0; edition < editionList->GetCount(); edition++) {
			PlanDataEdition *editem = &editionList->GetAt(edition);

			editem->pressList->RemoveAll();
			delete editem->pressList;

			for (int section = 0; section < editem->sectionList->GetCount(); section++) {
				PlanDataSection *secitem = &editem->sectionList->GetAt(section);
				for (int page = 0; page < secitem->pageList->GetCount(); page++) {
					PlanDataPage *pageitem = &secitem->pageList->GetAt(page);
					pageitem->colorList->RemoveAll();
					delete pageitem->colorList;
				}
				secitem->pageList->RemoveAll();
				delete secitem->pageList;
			}
			editem->sectionList->RemoveAll();
			delete editem->sectionList;

			for (int sheet = 0; sheet < editem->sheetList->GetCount(); sheet++) {
				PlanDataSheet *sheetItem = &editem->sheetList->GetAt(sheet);

				sheetItem->frontSheet->pressCylinders->RemoveAll();
				delete sheetItem->frontSheet->pressCylinders;

				sheetItem->frontSheet->sheetItems->RemoveAll();
				delete sheetItem->frontSheet->sheetItems;

				sheetItem->backSheet->sheetItems->RemoveAll();
				delete sheetItem->backSheet->sheetItems;

				sheetItem->backSheet->pressCylinders->RemoveAll();
				delete sheetItem->backSheet->pressCylinders;

				delete sheetItem->frontSheet;
				delete sheetItem->backSheet;

			}

			editem->sheetList->RemoveAll();
			delete editem->sheetList;
		}
		editionList->RemoveAll();
		delete editionList;

		planName = "";
		planID = "";
		publicationName = "";
		publicationAlias = "";
		publicationDate = CTime(1975, 1, 1, 0, 0, 0);
		version = 1;
		state = PLANSTATE_LOCKED;

		arrSectionNames.RemoveAll();
		arrSectionLinkNames.RemoveAll();
		arrSectionColors.RemoveAll();

	}

	void LogPlanData(CString sFileName)
	{

		CUtils util;

		util.Logprintf("\r\nParsing result of file %s", sFileName);
		util.Logprintf("----------------------------------");

		util.Logprintf("PlanName         : %s", planName);
		util.Logprintf("Version          : %d", version);
		util.Logprintf("Publication      : %s (alias %d)", publicationName, publicationAlias);
		util.Logprintf("PubDate          : %.4d-%.2d-%.2d (YYYY-MM-DD)", publicationDate.GetYear(), publicationDate.GetMonth(), publicationDate.GetDay());
		util.Logprintf("Editions in plan : %d", editionList->GetCount());
		for (int edition = 0; edition < editionList->GetCount(); edition++) {
			PlanDataEdition editem = editionList->GetAt(edition);
			util.Logprintf("  Edition        : %s", editem.editionName);

			CString s = "";
			for (int press = 0; press < editem.pressList->GetCount(); press++) {
				PlanDataPress pressitem = editem.pressList->GetAt(press);
				if (s != "")
					s += ", ";
				s += pressitem.pressName;
			}
			util.Logprintf("  Presses to use : %s", s);

			util.Logprintf("  Sections in edition : %d", editem.sectionList->GetCount());
			for (int section = 0; section < editem.sectionList->GetCount(); section++) {
				PlanDataSection secitem = editem.sectionList->GetAt(section);
				util.Logprintf("    Section      : %s", secitem.sectionName);
				util.Logprintf("    Pages in section: %d", secitem.pageList->GetCount());
				for (int page = 0; page < secitem.pageList->GetCount(); page++) {
					PlanDataPage pageitem = secitem.pageList->GetAt(page);
					CString sColors = "";
					for (int color = 0; color < pageitem.colorList->GetCount(); color++) {
						PlanDataPageSeparation coloritem = pageitem.colorList->GetAt(color);
						sColors += coloritem.colorName + "  ";
					}
					util.Logprintf("      Pagename : %s  Pagination %d PageIndex %d Colors (%s) PageID %s MasterPageID %s", pageitem.pageName, pageitem.pagination, pageitem.pageIndex, sColors, pageitem.pageID, pageitem.masterPageID);
				}
			}

		}
	}
};

