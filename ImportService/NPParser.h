#pragma once
#include "stdafx.h"
#include "Defs.h"
#include "PlanDataDefs.h"
#include "DatabaseManager.h"

BOOL ProcessNPFile(CDatabaseManager &DB, CString sFullInputPath, CString &sOutputFile, CString &sErrorMessage);

//int NPPreProcessFile(CString sInputFileName, CString sOutputFolder, CString sLocationFilter);
BOOL LegalPageName(CString sPageName);
int NPGenerateXML(CString sXmlFile, PlanData *plan, BOOL hasSheets, BOOL bForcePDF, BOOL bIsLast);

int NPParseLine(CString sLine,
	CString sContents, CString &sLocation,
	CString &sPublication, CString &sPubDate, CString &sZonedEdition, CString &sTimedEdition,
	CString &sSection, CString &sPageNumber, CString &sComment, CString &sVersion,
	CString &sPagination, CString &sPlanPagename, CString &sColors, CString &sDeadLine, CString &sPress);

BOOL NPPreImportFile(CDatabaseManager &DB, CString sInputFileName, CString &sErrorMessage);
BOOL NPImportFile(CDatabaseManager &DB, CString sInputFileName, CString sOutputFolder, CString &sInfo, int &nPageCount, int &nEditionCount, CString &sFinalOutputName, CString &sErrorMessage);


