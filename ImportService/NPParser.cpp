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
#include "NPPrefs.h"
#include "NPParser.h"

extern CPrefs g_prefs;
extern CUtils g_util;

CNPPrefs g_npprefs;

int g_planID = 0;


BOOL ProcessNPFile(CDatabaseManager &DB, CString sFullInputPath, CString &sOutputFile, CString &sErrorMessage)
{
	CStringArray sKnownLocations;
	sKnownLocations.Add("P1");
	sKnownLocations.Add("P2");
	sKnownLocations.Add("P3");
	sKnownLocations.Add("P4");

	int nReferenceNumber = 0;
	int nPageCount = 0;
	int nEditionCount = 0;
	g_util.Logprintf("INFO: ");
	g_util.Logprintf("INFO: Analyzing NP %s..", g_util.GetFileName(sFullInputPath));

	if (g_npprefs.m_archivefolder != "")
		::CopyFile(sFullInputPath, g_npprefs.m_archivefolder + _T("\\") + g_util.GetFileName(sFullInputPath), FALSE);

	// Housekeeping..
	DB.NPDeleteOldData(g_npprefs.m_daystokeepdata, sErrorMessage);
	DB.NPDeleteIllegalLocationData(sErrorMessage);

	if (NPPreImportFile(DB, sFullInputPath, sErrorMessage) == FALSE)
		return FALSE;

	CString sFinalFileName = "";
	CString sInfo = "";
	if (NPImportFile(DB, sFullInputPath, g_util.GetFilePath(sFullInputPath), sInfo, nPageCount, nEditionCount, sFinalFileName, sErrorMessage) == FALSE) {
		g_util.Logprintf("ERROR: Conversion failed for file %s", g_util.GetFilePath(sFullInputPath));
		g_util.Logprintf("ERROR:  %s", sInfo);
		g_util.Logprintf("ERROR:  %s", sErrorMessage);
		return FALSE;
	}

	g_util.Logprintf("INFO: Conversion succeeded %s", sInfo);

	sOutputFile = sFinalFileName;
	return TRUE;

}

/*
int NPPreProcessFile(CString sInputFileName, CString sOutputFolder, CString sLocationFilter)
{
	CString csText = "";
	int nLinesFound = 0;
	if (g_util.Load(sInputFileName, csText) == FALSE)
		return 0;

	CStringArray sLines, sLinesOutput;

	csText.Replace('\t', ' ');
	int nLines = g_util.String2Lines(csText, sLines);


	if (nLines == 0)
		return FALSE;

	// Top line holds header
	CString sLine = sLines[0];
	sLinesOutput.Add(sLines[0]);

	for (int i = 1; i < nLines; i++) {

		sLine = sLines[i];
		if (sLine.Trim() == "")
			continue;

		int n = sLine.ReverseFind('_');
		if (n != -1) {
			CString sLocs = sLine.Mid(n + 1);
			if (sLocs.Find(sLocationFilter) != -1) {
				nLinesFound++;
				sLinesOutput.Add(sLine.Left(n) + "_" + sLocationFilter);
			}
		}
	}

	if (nLinesFound > 0) {

		CString sTitle = g_util.GetFileName(sInputFileName);
		if (sTitle.ReverseFind('.'))
			sTitle = sTitle.Left(sTitle.ReverseFind('.'));
		CString sOutputName = sOutputFolder + "\\" + sTitle + "_" + sLocationFilter + ".txt";

		CString csTextOutput = "";
		for (int i = 0; i < sLinesOutput.GetCount(); i++)
			csTextOutput += sLinesOutput[i] + "\r\n";

		g_util.Save(sOutputName, csTextOutput);

	}

	return nLinesFound;
}
*/

BOOL LegalPageName(CString sPageName)
{
	if (sPageName == "" || sPageName == "0" || sPageName == "00" || sPageName == "000")
		return FALSE;
	if (sPageName.GetLength() > 1 && g_npprefs.m_pagenameprefixfromcomment>0) {
		sPageName = sPageName.Mid(1);
		if (sPageName == "" || sPageName == "0" || sPageName == "00" || sPageName == "000")
			return FALSE;
	}
	return TRUE;
}


int NPParseLine(CString sLine,
				CString sContents, CString &sLocation,
				CString &sPublication, CString &sPubDate, CString &sZonedEdition, CString &sTimedEdition,
				CString &sSection, CString &sPageNumber, CString &sComment, CString &sVersion,
				CString &sPagination, CString &sPlanPagename, CString &sColors, CString &sDeadLine, CString &sPress)
{
	CStringArray aElements;
	CStringArray aContentElements;
	sContents.Trim();
	sLine.Trim();

	sLine.Replace("\x1e", "");
	sLine.Replace("\x1f", "");

	int n = g_util.StringSplitter(sContents, ",;", aContentElements);
	int nElementsInLine = g_util.StringSplitter(sLine, ",;", aElements);
	int nHits = 0;

	for (int i = 0; i < n; i++) {

		if (aContentElements[i].Trim() == "%K" && i < nElementsInLine) {
			sComment = aElements[i].Trim();
			nHits++;
			continue;
		}

		if (aContentElements[i].Trim() == "%L" && i < nElementsInLine) {
			sLocation = aElements[i].Trim();
			nHits++;
			continue;
		}

		if (aContentElements[i].Trim() == "%A" && i < nElementsInLine) {
			sPress = aElements[i].Trim();
			nHits++;
			continue;
		}

		if (aContentElements[i].Trim() == "%P" && i < nElementsInLine) {
			sPublication = aElements[i].Trim();
			nHits++;
			continue;
		}

		if (aContentElements[i].Trim() == "%D" && i < nElementsInLine) {
			sPubDate = aElements[i].Trim();
			nHits++;
			continue;
		}

		if (aContentElements[i].Trim() == "%Z" && i < nElementsInLine) {
			sZonedEdition = aElements[i].Trim();
			nHits++;
			continue;
		}
		if (aContentElements[i].Trim() == "%E" && i < nElementsInLine) {
			sTimedEdition = aElements[i].Trim();
			nHits++;
			continue;
		}

		if (aContentElements[i].Trim() == "%S" && i < nElementsInLine) {
			sSection = aElements[i].Trim();
			nHits++;
			continue;
		}

		if (aContentElements[i].Trim() == "%3N" && i < nElementsInLine) {
			sPageNumber = aElements[i].Trim();
			int nn = atoi(sPageNumber);
			sPageNumber.Format("%.3d", nn);
			if (g_npprefs.m_pagenameprefixfromcomment)
				sPageNumber = sComment + sPageNumber;
			nHits++;
			continue;
		}

		if (aContentElements[i].Trim() == "%2N" && i < nElementsInLine) {
			sPageNumber = aElements[i].Trim();
			int nn = atoi(sPageNumber);
			sPageNumber.Format("%.2d", nn);
			if (g_npprefs.m_pagenameprefixfromcomment)
				sPageNumber = sComment + sPageNumber;

			nHits++;
			continue;
		}

		if (aContentElements[i].Trim() == "%N" && i < nElementsInLine) {
			sPageNumber = aElements[i].Trim();
			int nn = atoi(sPageNumber);
			sPageNumber.Format("%d", nn);
			if (g_npprefs.m_pagenameprefixfromcomment)
				sPageNumber = sComment + sPageNumber;
			nHits++;
			continue;
		}

		if (aContentElements[i].Trim() == "%M" && i < nElementsInLine) {
			sPagination = aElements[i].Trim();
			int nn = atoi(sPagination);
			sPagination.Format("%d", nn);
			nHits++;
			continue;
		}

		if (aContentElements[i].Trim() == "%Q" && i < nElementsInLine) {
			sPlanPagename = aElements[i].Trim();
			nHits++;
			continue;
		}


		if (aContentElements[i].Trim() == "%V" && i < nElementsInLine) {
			sVersion = aElements[i].Trim();
			nHits++;
			continue;
		}

		if (aContentElements[i].Trim() == "%C" && i < nElementsInLine) {
			sColors = aElements[i].Trim();
			nHits++;
			continue;
		}

		if (aContentElements[i].Trim() == "%X" && i < nElementsInLine) {
			sDeadLine = aElements[i].Trim();
			nHits++;
			continue;
		}

	}

	return nHits;
}

BOOL NPPreImportFile(CDatabaseManager &DB, CString sInputFileName, CString &sErrorMessage)
{

	BOOL bRet = TRUE;
	sErrorMessage = _T("");
	CString sHeaderString;
	CString sLocation = "", sPublication = "", sPubDate = "";
	CString sZone = "", sEdition = "", sVersion = "", sColor = "", sSection = "";
	CString sPageName = "", sPagination = "", sComment = "", sDeadLine = "", sPress = "";
	CString sProductionName = "", sPlanPageName = "";
	CTime tPubDate = CTime(1975, 1, 1, 0, 0, 0);
	BOOL ret = TRUE;

	//int nPagesInEdition = 0;

	//nPageCount = 0;
	//nEditionCount = 0;

	CString csText = "";
	if (g_util.Load(sInputFileName, csText) == FALSE)
		return FALSE;

	CString sFileFileTitle = g_util.GetFileName(sInputFileName);
	if (g_npprefs.m_filenamecontents != "" && g_npprefs.m_matchexpressionfilename != "" && g_npprefs.m_formatexpressionfilename != "") {

		sFileFileTitle = g_util.FormatExpression(g_npprefs.m_matchexpressionfilename, g_npprefs.m_formatexpressionfilename, sFileFileTitle, TRUE);

		if (sFileFileTitle.Trim() != "") {

			NPParseLine(sFileFileTitle, g_npprefs.m_filenamecontents, sLocation, sPublication, sPubDate, sZone, sEdition, sSection,
				sPageName, sComment, sVersion, sPagination, sPlanPageName, sColor, sDeadLine, sPress);

			if (sLocation != "") {
				CString sLocationAlias = DB.LoadSpecificAlias("Location", sLocation, sErrorMessage);
				if (sLocationAlias == sLocation)
					sLocationAlias = DB.LoadSpecificAlias("Press", sLocation, sErrorMessage);
				sLocation = sLocationAlias;
			}

			sPress = DB.LoadSpecificAlias("Press", sPress, sErrorMessage);
			sPublication = DB.LoadSpecificAlias("Publication", sPublication, sErrorMessage);
			sZone = DB.LoadSpecificAlias("Edition", sZone, sErrorMessage);
			sEdition = DB.LoadSpecificAlias("Edition", sEdition, sErrorMessage);
			sSection = DB.LoadSpecificAlias("Section", sSection, sErrorMessage);

		}
	}

	CStringArray sLines, sElements;

	csText.Replace('\t', ' ');
	int nLines = g_util.String2Lines(csText, sLines);


	if (nLines == 0)
		return FALSE;

	// Top line holds header
	CString sLine = sLines[0];

	int nTotalPages = 0;
	int nImportedPages = 0;
	CStringArray sEditionNameList;

	if (g_npprefs.m_headercontents != "" && sLine.Trim() != "") {
		sHeaderString = sLine;

		if (g_npprefs.m_matchexpressiontop != "" && g_npprefs.m_formatexpressiontop != "")
			sHeaderString = g_util.FormatExpression(g_npprefs.m_matchexpressiontop, g_npprefs.m_formatexpressiontop, sHeaderString, TRUE);

		if (sHeaderString.Trim() != "") {

			NPParseLine(sHeaderString, g_npprefs.m_headercontents, sLocation, sPublication, sPubDate, sZone, sEdition, sSection,
				sPageName, sComment, sVersion, sPagination, sPlanPageName, sColor, sDeadLine, sPress);

			CString sLocationAlias = DB.LoadSpecificAlias("Location", sLocation, sErrorMessage);
			if (sLocationAlias == sLocation)
				sLocationAlias = DB.LoadSpecificAlias("Press", sLocation, sErrorMessage);
			sLocation = sLocationAlias;

			sPublication = DB.LoadSpecificAlias("Publication", sPublication, sErrorMessage);
			sZone = DB.LoadSpecificAlias("Edition", sZone, sErrorMessage);
			sEdition = DB.LoadSpecificAlias("Edition", sEdition, sErrorMessage);
			sSection = DB.LoadSpecificAlias("Section", sSection, sErrorMessage);
			sPress = DB.LoadSpecificAlias("Press", sPress, sErrorMessage);

		}
	}

	BOOL bHasPublication = FALSE;
	BOOL bHasPress = FALSE;
	BOOL bHasLocation = FALSE;

	for (int i = 1; i < nLines; i++) {

//		double tm = 100.0*i / nLines;

		sLine = sLines[i];
		if (sLine.Trim() == "")
			continue;

		if (g_npprefs.m_subheadercontents.Trim() != "" && g_npprefs.m_subheadertriggerexpr != "") {
			if (g_util.TryMatchExpression(g_npprefs.m_subheadertriggerexpr, sLine, FALSE)) {
				// We got a sub-header line!
				sHeaderString = sLine;
				if (g_npprefs.m_matchexpressionheader != "" && g_npprefs.m_formatexpressionheader != "")
					sHeaderString = g_util.FormatExpression(g_npprefs.m_matchexpressionheader, g_npprefs.m_formatexpressionheader, sHeaderString, FALSE);
				if (sHeaderString != "") {
					if (NPParseLine(sHeaderString, g_npprefs.m_subheadercontents, sLocation, sPublication, sPubDate, sZone, sEdition,
						sSection, sPageName, sComment, sVersion, sPagination, sPlanPageName, sColor, sDeadLine, sPress) > 0) {

						CString sLocationAlias = DB.LoadSpecificAlias("Location", sLocation, sErrorMessage);
						if (sLocationAlias == sLocation)
							sLocationAlias = DB.LoadSpecificAlias("Press", sLocation, sErrorMessage);
						sLocation = sLocationAlias;

						sPublication = DB.LoadSpecificAlias("Publication", sPublication, sErrorMessage);
						sZone = DB.LoadSpecificAlias("Edition", sZone, sErrorMessage);
						sEdition = DB.LoadSpecificAlias("Edition", sEdition, sErrorMessage);
						sSection = DB.LoadSpecificAlias("Section", sSection, sErrorMessage);
						sPress = DB.LoadSpecificAlias("Press", sPress, sErrorMessage);
					}
				}
			}
			continue;
		}

		// Step 3 - analyze page line

		if (g_npprefs.m_preparsematchexpression != "" && g_npprefs.m_preparseformatexpression != "") {
			CString preparsematchexpression1 = g_npprefs.m_preparsematchexpression;
			CString preparsematchexpression2 = "";
			CString preparseformatexpression1 = g_npprefs.m_preparseformatexpression;
			CString preparseformatexpression2 = "";

			int nn = preparsematchexpression1.Find(";");

			if (nn != -1) {
				preparsematchexpression2 = preparsematchexpression1.Mid(nn + 1);
				preparsematchexpression1 = preparsematchexpression1.Left(nn);
			}
			nn = preparseformatexpression1.Find(";");
			if (nn != -1) {
				preparseformatexpression2 = preparseformatexpression1.Mid(nn + 1);
				preparseformatexpression1 = preparseformatexpression1.Left(nn);
			}

			sLine = g_util.FormatExpression(preparsematchexpression1, preparseformatexpression1, sLine, FALSE);
			if (preparsematchexpression2 != "" && preparseformatexpression2 != "")
				sLine = g_util.FormatExpression(preparsematchexpression2, preparseformatexpression2, sLine, FALSE);

		}

		CString sDataRow = sLine;

		if (g_npprefs.m_multipleregextests) {
			for (int j = 0; j < g_npprefs.m_numberofexpressions; j++) {
				BOOL bRegexHit = FALSE;
				if (g_npprefs.m_matchexpressionbody[j] != "" && g_npprefs.m_formatexpressionbody[j] != "") {
					sDataRow = g_util.FormatExpression(g_npprefs.m_matchexpressionbody[j], g_npprefs.m_formatexpressionbody[j], sLine, FALSE);

					if (sDataRow != sLine)
						bRegexHit = TRUE;
				}

				if (bRegexHit) {
					sPageName = "";
					if (sDataRow.Trim() != "") {
						if (NPParseLine(sDataRow, g_npprefs.m_datarowcontents, sLocation, sPublication, sPubDate, sZone, sEdition, sSection,
							sPageName, sComment, sVersion, sPagination, sPlanPageName, sColor, sDeadLine, sPress) > 0) {

							if (sLocation != "" && bHasLocation == FALSE) {
								CString sLocationAlias = DB.LoadSpecificAlias("Location", sLocation, sErrorMessage);
								if (sLocationAlias == sLocation)
									sLocationAlias = DB.LoadSpecificAlias("Press", sLocation, sErrorMessage);
								sLocation = sLocationAlias;
							}
							if (sLocation != "")
								bHasLocation = TRUE;

							if (sPress != "" && bHasPress == FALSE)
								sPress = DB.LoadSpecificAlias("Press", sPress, sErrorMessage);
							if (sPress != "")
								bHasPress = TRUE;

							if (sPublication != "" && bHasPublication == FALSE)
								sPublication = DB.LoadSpecificAlias("Publication", sPublication, sErrorMessage);
							if (sPublication != "")
								bHasPublication = TRUE;

							sZone = DB.LoadSpecificAlias("Edition", sZone, sErrorMessage);
							sEdition = DB.LoadSpecificAlias("Edition", sEdition, sErrorMessage);
							sSection = DB.LoadSpecificAlias("Section", sSection, sErrorMessage);
						}
					}


					if (sPubDate == "" || sPublication == "" || sPageName == "" || sPageName == "0")
						continue;

					CTime tDeadLine = CTime(1975, 1, 1, 0, 0, 0);
					if (sDeadLine != "")
						tDeadLine = g_util.ParseDateTime(sDeadLine, g_npprefs.m_timeformat);

					tPubDate = g_util.ParseDateTime(sPubDate, g_npprefs.m_dateformat);

					if (g_npprefs.m_daystokeepdata > 0) {
						DB.NPDeletePages(
							sLocation != "" ? sLocation : g_npprefs.m_defaultlocation,
							sPublication, tPubDate,
							sZone != "" ? sZone : g_npprefs.m_defaultzone,
							sEdition != "" ? sEdition : g_npprefs.m_defaultedition,
							sSection != "" ? sSection : "1", sErrorMessage);
					}
				}
			}

		}
		else {
			for (int j = 0; j < g_npprefs.m_numberofexpressions; j++) {
				if (g_npprefs.m_matchexpressionbody[j] != "" && g_npprefs.m_formatexpressionbody[j] != "") {
					sDataRow = g_util.FormatExpression(g_npprefs.m_matchexpressionbody[j], g_npprefs.m_formatexpressionbody[j], sLine, FALSE);

					if (sDataRow != sLine)
						break;
				}
			}

			sPageName = "";
			if (sDataRow.Trim() != "") {
				if (NPParseLine(sDataRow, g_npprefs.m_datarowcontents, sLocation, sPublication, sPubDate, sZone, sEdition, sSection,
					sPageName, sComment, sVersion, sPagination, sPlanPageName, sColor, sDeadLine, sPress) > 0) {

					if (sLocation != "" && bHasLocation == FALSE) {
						CString sLocationAlias = DB.LoadSpecificAlias("Location", sLocation, sErrorMessage);
						if (sLocationAlias == sLocation)
							sLocationAlias = DB.LoadSpecificAlias("Press", sLocation, sErrorMessage);
						sLocation = sLocationAlias;
					}
					if (sLocation != "")
						bHasLocation = TRUE;
					if (sPress != "" && bHasPress == FALSE)
						sPress = DB.LoadSpecificAlias("Press", sPress, sErrorMessage);
					if (sPress != "")
						bHasPress = TRUE;

					if (sPublication != "" && bHasPublication == FALSE)
						sPublication = DB.LoadSpecificAlias("Publication", sPublication, sErrorMessage);
					if (sPublication != "")
						bHasPublication = TRUE;

					sZone = DB.LoadSpecificAlias("Edition", sZone, sErrorMessage);
					sEdition = DB.LoadSpecificAlias("Edition", sEdition, sErrorMessage);
					sSection = DB.LoadSpecificAlias("Section", sSection, sErrorMessage);
				}
			}

			if (sPubDate == "" || sPublication == "" || LegalPageName(sPageName) == FALSE)
				continue;

			CTime tDeadLine = CTime(1975, 1, 1, 0, 0, 0);
			if (sDeadLine != "")
				tDeadLine = g_util.ParseDateTime(sDeadLine, g_npprefs.m_timeformat);

			tPubDate = g_util.ParseDateTime(sPubDate, g_npprefs.m_dateformat);

			if (g_npprefs.m_daystokeepdata > 0) {
				DB.NPDeletePages(
						sLocation != "" ? sLocation : g_npprefs.m_defaultlocation,
						sPublication, tPubDate,
						sZone != "" ? sZone : g_npprefs.m_defaultzone,
						sEdition != "" ? sEdition : g_npprefs.m_defaultedition,
						sSection != "" ? sSection : "1", sErrorMessage);
			}

		}

	}

	return TRUE;
}


BOOL NPImportFile(CDatabaseManager &DB, CString sInputFileName, CString sOutputFolder, CString &sInfo, int &nPageCount, int &nEditionCount, CString &sFinalOutputName, CString &sErrorMessage)
{
	BOOL bRet = TRUE;
	CString sHeaderString;
	CString sLocation = "", sPublication = "", sPubDate = "";
	CString sZone = "", sEdition = "", sVersion = "", sColor = "", sSection = "";
	CString sPageName = "", sPagination = "", sComment = "", sDeadLine = "", sPress = "";
	CString sProductionName = "", sPlanPageName = "";
	CTime tPubDate = CTime(1975, 1, 1, 0, 0, 0);
	BOOL ret = TRUE;

	int nPagesInEdition = 0;
	sInfo = _T("");
	sErrorMessage = _T("");

	nPageCount = 0;
	nEditionCount = 0;


	CString csText = "";
	if (g_util.Load(sInputFileName, csText) == FALSE) {
		sErrorMessage =_T( "XML load() failed");
		return FALSE;
	}

	CString sFileFileTitle = g_util.GetFileName(sInputFileName);
	if (g_npprefs.m_filenamecontents != "" && g_npprefs.m_matchexpressionfilename != "" && g_npprefs.m_formatexpressionfilename != "") {

		sFileFileTitle = g_util.FormatExpression(g_npprefs.m_matchexpressionfilename, g_npprefs.m_formatexpressionfilename, sFileFileTitle, TRUE);

		if (sFileFileTitle.Trim() != "") {

			NPParseLine(sFileFileTitle, g_npprefs.m_filenamecontents, sLocation, sPublication, sPubDate, sZone, sEdition, sSection,
				sPageName, sComment, sVersion, sPagination, sPlanPageName, sColor, sDeadLine, sPress);
			CString sLocationAlias = DB.LoadSpecificAlias("Location", sLocation, sErrorMessage);
			if (sLocationAlias == sLocation)
				sLocationAlias = DB.LoadSpecificAlias("Press", sLocation, sErrorMessage);
			sLocation = sLocationAlias;

			//sPress = m_DB.LoadSpecificAlias("Press", sPress, szErrorMessage);

			CString sPubcationAlias = DB.LoadSpecificAlias("Publication", sPublication, sErrorMessage);
			sPublication = sPubcationAlias;

			sZone = DB.LoadSpecificAlias("Edition", sZone, sErrorMessage);
			sEdition = DB.LoadSpecificAlias("Edition", sEdition, sErrorMessage);
			sSection = DB.LoadSpecificAlias("Section", sSection, sErrorMessage);

		}
	}

	CStringArray sLines, sElements;

	csText.Replace('\t', ' ');
	int nLines = g_util.String2Lines(csText, sLines);


	if (nLines == 0) {
		sErrorMessage = _T( "String2Lines() failed");
		return FALSE;
	}

	// Top line holds header
	CString sLine = sLines[0];

	int nTotalPages = 0;
	int nImportedPages = 0;
	CStringArray sEditionNameList;

	if (g_npprefs.m_headercontents != "" && sLine.Trim() != "") {
		sHeaderString = sLine;

		if (g_npprefs.m_matchexpressiontop != "" && g_npprefs.m_formatexpressiontop != "")
			sHeaderString = g_util.FormatExpression(g_npprefs.m_matchexpressiontop, g_npprefs.m_formatexpressiontop, sHeaderString, TRUE);

		if (sHeaderString.Trim() != "") {

			NPParseLine(sHeaderString, g_npprefs.m_headercontents, sLocation, sPublication, sPubDate, sZone, sEdition, sSection,
							sPageName, sComment, sVersion, sPagination, sPlanPageName, sColor, sDeadLine, sPress);
			CString sLocationAlias = DB.LoadSpecificAlias("Location", sLocation, sErrorMessage);
			if (sLocationAlias == sLocation)
				sLocationAlias = DB.LoadSpecificAlias("Press", sLocation, sErrorMessage);
			sLocation = sLocationAlias;
			sPress = DB.LoadSpecificAlias("Press", sPress, sErrorMessage);

			if (sPublication != "") {
				CString sPubcationAlias = DB.LoadSpecificAlias("Publication", sPublication, sErrorMessage);
				sPublication = sPubcationAlias;
			}

			sZone = DB.LoadSpecificAlias("Edition", sZone, sErrorMessage);
			sEdition = DB.LoadSpecificAlias("Edition", sEdition, sErrorMessage);
			sSection = DB.LoadSpecificAlias("Section", sSection, sErrorMessage);

		}
	}

	BOOL bHasLocation = FALSE;
	BOOL bHasPress = FALSE;
	BOOL bHasPublication = FALSE;

	for (int i = 1; i < nLines; i++) {

		double tm = 100.0*i / nLines;


		sLine = sLines[i];
		if (sLine.Trim() == "")
			continue;

		if (g_npprefs.m_subheadercontents.Trim() != "" && g_npprefs.m_subheadertriggerexpr != "") {
			if (g_util.TryMatchExpression(g_npprefs.m_subheadertriggerexpr, sLine, FALSE)) {
				// We got a sub-header line!
				sHeaderString = sLine;
				if (g_npprefs.m_matchexpressionheader != "" && g_npprefs.m_formatexpressionheader != "")
					sHeaderString = g_util.FormatExpression(g_npprefs.m_matchexpressionheader, g_npprefs.m_formatexpressionheader, sHeaderString, FALSE);
				if (sHeaderString != "") {
					if (NPParseLine(sHeaderString, g_npprefs.m_subheadercontents, sLocation, sPublication, sPubDate, sZone, sEdition,
						sSection, sPageName, sComment, sVersion, sPagination, sPlanPageName, sColor, sDeadLine, sPress) > 0) {

						if (sLocation != "" && bHasLocation == FALSE) {
							CString sLocationAlias = DB.LoadSpecificAlias("Location", sLocation, sErrorMessage);
							if (sLocationAlias == sLocation)
								sLocationAlias = DB.LoadSpecificAlias("Press", sLocation, sErrorMessage);
							sLocation = sLocationAlias;
						}
						if (sLocation != "")
							bHasLocation = TRUE;

						if (sPress != "" && bHasPress == FALSE)
							sPress = DB.LoadSpecificAlias("Press", sPress, sErrorMessage);
						if (sPress != "")
							bHasPress = TRUE;

						if (sPublication != "" && bHasPublication == FALSE) {
							CString sPubcationAlias = DB.LoadSpecificAlias("Publication", sPublication, sErrorMessage);
							sPublication = sPubcationAlias;
						}

						if (sPublication != "")
							bHasPublication = TRUE;

						sZone = DB.LoadSpecificAlias("Edition", sZone, sErrorMessage);
						sEdition = DB.LoadSpecificAlias("Edition", sEdition, sErrorMessage);
						sSection = DB.LoadSpecificAlias("Section", sSection, sErrorMessage);
					}
				}
			}
			continue;
		}

		// Step 3 - analyze page line

		if (g_npprefs.m_preparsematchexpression != "" && g_npprefs.m_preparseformatexpression != "") {
			CString preparsematchexpression1 = g_npprefs.m_preparsematchexpression;
			CString preparsematchexpression2 = "";
			CString preparseformatexpression1 = g_npprefs.m_preparseformatexpression;
			CString preparseformatexpression2 = "";

			int nn = preparsematchexpression1.Find(";");

			if (nn != -1) {
				preparsematchexpression2 = preparsematchexpression1.Mid(nn + 1);
				preparsematchexpression1 = preparsematchexpression1.Left(nn);
			}
			nn = preparseformatexpression1.Find(";");
			if (nn != -1) {
				preparseformatexpression2 = preparseformatexpression1.Mid(nn + 1);
				preparseformatexpression1 = preparseformatexpression1.Left(nn);
			}

			sLine = g_util.FormatExpression(preparsematchexpression1, preparseformatexpression1, sLine, FALSE);
			if (preparsematchexpression2 != "" && preparseformatexpression2 != "")
				sLine = g_util.FormatExpression(preparsematchexpression2, preparseformatexpression2, sLine, FALSE);

		}

		CString sDataRow = sLine;

		if (g_npprefs.m_multipleregextests) {
			for (int j = 0; j < g_npprefs.m_numberofexpressions; j++) {
				BOOL bRegexHit = FALSE;
				if (g_npprefs.m_matchexpressionbody[j] != "" && g_npprefs.m_formatexpressionbody[j] != "") {
					sDataRow = g_util.FormatExpression(g_npprefs.m_matchexpressionbody[j], g_npprefs.m_formatexpressionbody[j], sLine, FALSE);

					if (sDataRow != sLine)
						bRegexHit = TRUE;
				}

				if (bRegexHit) {
					sPageName = "";
					if (sDataRow.Trim() != "") {

						CString sThisPublication = _T("");
						CString sThisLocation = _T("");
						CString sThisPress = _T("");
						if (NPParseLine(sDataRow, g_npprefs.m_datarowcontents, sThisLocation, sThisPublication, sPubDate, sZone, sEdition, sSection,
							sPageName, sComment, sVersion, sPagination, sPlanPageName, sColor, sDeadLine, sThisPress) > 0) {

							if (sThisLocation != "" && bHasLocation == FALSE) {
								CString sLocationAlias = DB.LoadSpecificAlias("Location", sThisLocation, sErrorMessage);
								if (sLocationAlias == sThisLocation)
									sLocationAlias = DB.LoadSpecificAlias("Press", sThisLocation, sErrorMessage);
								sLocation = sLocationAlias;
							}
							if (sLocation != "")
								bHasLocation = TRUE;

							if (sThisPress != "" && bHasPress == FALSE)
								sPress = DB.LoadSpecificAlias("Press", sThisPress, sErrorMessage);
							if (sPress != "")
								bHasPress = TRUE;

							if (sThisPublication != "" && bHasPublication == FALSE) {
								CString sPubcationAlias = DB.LoadSpecificAlias("Publication", sThisPublication, sErrorMessage);
								sPublication = sPubcationAlias;
							}

							if (sPublication != "")
								bHasPublication = TRUE;

							sZone = DB.LoadSpecificAlias("Edition", sZone, sErrorMessage);
							sEdition = DB.LoadSpecificAlias("Edition", sEdition, sErrorMessage);
							sSection = DB.LoadSpecificAlias("Section", sSection, sErrorMessage);
						}
					}


					if (sPubDate == "" || sPublication == "" || LegalPageName(sPageName) == FALSE)
						continue;

					CTime tDeadLine = CTime(1975, 1, 1, 0, 0, 0);
					if (sDeadLine != "")
						tDeadLine = g_util.ParseDateTime(sDeadLine, g_npprefs.m_timeformat);

					tPubDate = g_util.ParseDateTime(sPubDate, g_npprefs.m_dateformat);

					BOOL bAdded = FALSE;

					int nPag = atoi(sPagination);
					if (nPag == 0) {
						nPag = atoi(sPageName);
						if (nPag == 0 && g_npprefs.m_pagenameprefixfromcomment)
							nPag = atoi(sPageName.Mid(1));
					}

					sColor = _T("PDF");

					DB.NPAddPage(sPublication, tPubDate, sProductionName,
								1,//atoi(sZone) > 0 ? atoi(sZone) : 1,
								sLocation != "" ? sLocation : g_npprefs.m_defaultlocation,
								sEdition != "" ? sEdition : g_npprefs.m_defaultedition,
								sZone != "" ? sZone : g_npprefs.m_defaultzone,
								sSection != "" ? sSection : "1",
								sPageName,
								nPag, //atoi(sPagination) > 0 ? atoi(sPagination) : atoi(sPageName), 
								sPlanPageName, sComment,
								_T("PDF"),
								sVersion != "" ? sVersion : "1",
								tDeadLine,
								sErrorMessage, bAdded);

					nTotalPages++;
					if (bAdded)
						nImportedPages++;

					g_util.AddToCStringArray(sEditionNameList, sZone + "_" + sEdition);
				}
			}

		}
		else {
			for (int j = 0; j < g_npprefs.m_numberofexpressions; j++) {
				if (g_npprefs.m_matchexpressionbody[j] != "" && g_npprefs.m_formatexpressionbody[j] != "") {
					sDataRow = g_util.FormatExpression(g_npprefs.m_matchexpressionbody[j], g_npprefs.m_formatexpressionbody[j], sLine, FALSE);

					if (sDataRow != sLine)
						break;
				}
			}
			sPageName = "";
			if (sDataRow.Trim() != "") {
				CString sThisPublication = _T("");
				CString sThisLocation = _T("");
				CString sThisPress = _T("");
				if (NPParseLine(sDataRow, g_npprefs.m_datarowcontents, sThisLocation, sThisPublication, sPubDate, sZone, sEdition, sSection,
					sPageName, sComment, sVersion, sPagination, sPlanPageName, sColor, sDeadLine, sThisPress) > 0) {

					if (sThisLocation != "" && bHasLocation == FALSE) {
						CString sLocationAlias = DB.LoadSpecificAlias("Location", sThisLocation, sErrorMessage);
						if (sLocationAlias == sThisLocation)
							sLocationAlias = DB.LoadSpecificAlias("Press", sThisLocation, sErrorMessage);
						sLocation = sLocationAlias;
					}
					if (sLocation != "")
						bHasLocation = TRUE;

					if (sThisPress != "" && bHasPress == FALSE)
						sPress = DB.LoadSpecificAlias("Press", sThisPress, sErrorMessage);
					if (sPress != "")
						bHasPress = TRUE;

					if (sThisPublication != "" && bHasPublication == FALSE) {
						CString sPubcationAlias = DB.LoadSpecificAlias("Publication", sThisPublication, sErrorMessage);
						sPublication = sPubcationAlias;
					}

					if (sPublication != "")
						bHasPublication = TRUE;

					sZone = DB.LoadSpecificAlias("Edition", sZone, sErrorMessage);
					sEdition = DB.LoadSpecificAlias("Edition", sEdition, sErrorMessage);
					sSection = DB.LoadSpecificAlias("Section", sSection, sErrorMessage);

				}
			}


			if (sPubDate == "" || sPublication == "" || LegalPageName(sPageName) == FALSE)
				continue;

			CTime tDeadLine = CTime(1975, 1, 1, 0, 0, 0);
			if (sDeadLine != "")
				tDeadLine = g_util.ParseDateTime(sDeadLine, g_npprefs.m_timeformat);

			tPubDate = g_util.ParseDateTime(sPubDate, g_npprefs.m_dateformat);

			BOOL bAdded = FALSE;

			int nPag = atoi(sPagination);
			if (nPag == 0) {
				nPag = atoi(sPageName);
				if (nPag == 0 && g_npprefs.m_pagenameprefixfromcomment)
					nPag = atoi(sPageName.Mid(1));
			}



			DB.NPAddPage(sPublication, tPubDate, sProductionName,
							1, //atoi(sZone) > 0 ? atoi(sZone) : 1,
							sLocation != "" ? sLocation : g_npprefs.m_defaultlocation,
							sEdition != "" ? sEdition : g_npprefs.m_defaultedition,
							sZone != "" ? sZone : g_npprefs.m_defaultzone,
							sSection != "" ? sSection : "1",
							sPageName,
							nPag, //atoi(sPagination) > 0 ? atoi(sPagination) : atoi(sPageName), 
							sPlanPageName, sComment,
							_T("PDF"),
							sVersion != "" ? sVersion : "1",
							tDeadLine,
							sErrorMessage, bAdded);

			nTotalPages++;
			if (bAdded)
				nImportedPages++;

			g_util.AddToCStringArray(sEditionNameList, sZone + "_" + sEdition);

		}

	}


	nPageCount = nTotalPages;
	nEditionCount = (int)sEditionNameList.GetCount();

	sInfo.Format("Imported %s-%s %d editions, total %d new pages", sPublication, sPubDate, sEditionNameList.GetCount(), nImportedPages);


	CStringArray saEditionList;	// hold form Zone_Ed
	CStringArray saSectionList;
	CString sError = "";

	if (g_npprefs.m_rejectidenticalplan || g_npprefs.m_rejectdifferentpagecount) {

		if (DB.NPGetInfoFromImportTableEx(IMPORTTABLEQUERYMODE_EDITIONS, sPublication, tPubDate, "", "", saEditionList, sErrorMessage) == FALSE) {
			sErrorMessage = _T("GetInfoFromImportTableEx(1) failed");
			return FALSE;
		}

		if (DB.NPGetInfoFromImportTableEx(IMPORTTABLEQUERYMODE_SECTIONS, sPublication, tPubDate, "", "", saSectionList, sErrorMessage) == FALSE) {
			sErrorMessage =  _T("GetInfoFromImportTableEx(2) failed");
			return FALSE;
		}
	}

	// Investigate if we must reject plan due to page count changes!!!

	if (g_npprefs.m_rejectidenticalplan) {

		// See if identical plan already imported - all sections/editions must match

		BOOL bAllSame = TRUE;

		sError = "";
		for (int ed = 0; ed < saEditionList.GetCount(); ed++) {
			for (int sec = 0; sec < saSectionList.GetCount(); sec++) {

				int nImportPageCount = 0; // unique pages
				if (DB.NPGetPageCount(sPublication, tPubDate, saEditionList[ed], saSectionList[sec], nImportPageCount, sErrorMessage) == FALSE) {
					sErrorMessage = _T("GetPageCount() failed");
					return FALSE;
				}

				int nPlanPageCount = 0; // unique pages
				if (DB.NPGetPageTablePageCount(sPublication, tPubDate, saEditionList[ed], saSectionList[sec], nPlanPageCount, FALSE, sErrorMessage) == FALSE) {
					sErrorMessage =  _T("GetPageTablePageCount() failed");
					return FALSE;
				}

				// Not yet imported at all?
				if (nPlanPageCount == 0) {
					bAllSame = FALSE;
					break;
				}

				// Detect if any difference
				if (nImportPageCount > 0 && nPlanPageCount > 0 && nImportPageCount != nPlanPageCount) {
					bAllSame = FALSE;
					break;
				}
			}
			if (bAllSame == FALSE)
				break;
		}
		if (bAllSame) {
			sError.Format("%s %.2d-%.2d already imported - plan is rejected", sPublication, tPubDate.GetDay(), tPubDate.GetMonth());
			//m_DB.DeletePages("", sPublication, tPubDate, "", "", "", szErrorMessage);
			sErrorMessage = sError;
			sInfo = sError;

			return -1;
		}
	}

	if (g_npprefs.m_rejectdifferentpagecount) {

		BOOL bRejectPlan = FALSE;

		sError = "";

		// Investigate main edition page count - if new import different - reject (must be deleted manually)
		for (int ed = 0; ed < saEditionList.GetCount(); ed++) {

			for (int sec = 0; sec < saSectionList.GetCount(); sec++) {

				int nImportPageCount = 0; // Unique pages!
				if (DB.NPGetPageCount(sPublication, tPubDate, saEditionList[ed], saSectionList[sec], nImportPageCount, sErrorMessage) == FALSE)
					return FALSE;

				int nPlanPageCountUnique = 0; // Unique pages!
				if (DB.NPGetPageTablePageCount(sPublication, tPubDate, saEditionList[ed], saSectionList[sec], nPlanPageCountUnique, FALSE, sErrorMessage) == FALSE)
					return FALSE;
				int nPlanPageCountAll = 0; // All pages!
				if (DB.NPGetPageTablePageCount(sPublication, tPubDate, saEditionList[ed], saSectionList[sec], nPlanPageCountAll, TRUE, sErrorMessage) == FALSE)
					return FALSE;

				// Only analyze main unique edition (where all pages unique)
				if (nPlanPageCountUnique != nPlanPageCountAll)
					continue;

				if (nImportPageCount > 0 && nPlanPageCountUnique > 0 && nImportPageCount != nPlanPageCountUnique) {
					bRejectPlan = TRUE;
					sError.Format("%s %.2d-%.2d edition %s section %s has different page count (%d vs %d) than already imported plan - plan is rejected",
											sPublication, tPubDate.GetDay(), tPubDate.GetMonth(), saEditionList[ed], saSectionList[sec], nImportPageCount, nPlanPageCountUnique);
					//m_DB.DeletePages("", sPublication, tPubDate, "", "", "", szErrorMessage);
					sErrorMessage = sError;
					sInfo = sError;

					return -1;
				}
			}
		}



	}
	// PHASE TWO
	// Load database rows into plan structure


	DB.NPResetPages(sErrorMessage);

	CStringArray saLocationList;
	if (DB.NPGetInfoFromImportTable(IMPORTTABLEQUERYMODE_LOCATIONS, sPublication, tPubDate, "", "", saLocationList, sErrorMessage) == FALSE) {
		sErrorMessage = _T("GetPageTablePageCount(3) failed");
		return FALSE;
	}

	/*
		int nSequenceNumbers = 0;
		DB.NPGetSequenceNumberCount(sPublication, tPubDate, "", nSequenceNumbers, sErrorMessage);
		if (nSequenceNumbers == 1) {
			for (int loc = 0; loc < saLocationList.GetCount(); loc++) {
				DB.NPChangeSeqNumber(sPublication, tPubDate, saLocationList[loc], loc+1, sErrorMessage);
			}
		}
	*/
	for (int loc = 0; loc < saLocationList.GetCount(); loc++) {
		BOOL bIsLast = (loc == saLocationList.GetCount() - 1);
		CString thisLocation = saLocationList[loc];
		CString sTitle;
		CTime tNow = CTime::GetCurrentTime();
		sTitle.Format("NewsPilot-%s-%s-%s-%.4d%.2d%.2d%.2d%.2d%.2d", thisLocation, sPublication, sPubDate, tNow.GetYear(), tNow.GetMonth(), tNow.GetDay(), tNow.GetHour(), tNow.GetMinute(), tNow.GetSecond());

		PlanData plan;
		plan.planID.Format("%.6d", g_planID++);
		plan.planName.Format("NewsPilot-%s-%s-%s", thisLocation, sPublication, sPubDate);
		plan.version = 1;
		plan.updatetime = CTime::GetCurrentTime();
		plan.state = 1;
		plan.publicationName = sPublication;

		plan.publicationDate = g_util.ParseDateTime(sPubDate, g_npprefs.m_dateformat);
		plan.publicationAlias = sPublication;
		plan.pagenameprefix = "";
		plan.customername = "";
		plan.customeralias = "";

		CString sMasterZone = _T("");
		CString sMasterEdition = _T("1");
		CString sMasterLocation = saLocationList[loc];

		if (DB.NPGetZoneMaster(sPublication, tPubDate, sMasterLocation, sMasterZone, sErrorMessage) == FALSE) {
			sErrorMessage = _T("GetZoneMaster() failed");
			return FALSE;
		}

		int nPagesLoaded = 0;

		CStringArray saAllZoneEditionList;
		//if (DB.NPGetInfoFromImportTable(IMPORTTABLEQUERYMODE_EDITIONS_ON_LOCATION, sPublication, tPubDate, thisLocation, "", saAllZoneEditionList, sErrorMessage) == FALSE) {
		//	return FALSE;
		//}

		if (DB.NPGetInfoFromImportTable(IMPORTTABLEQUERYMODE_EDITIONS_ON_LOCATION, sPublication, tPubDate, thisLocation, "", saAllZoneEditionList, sErrorMessage) == FALSE) {
			sErrorMessage = _T("GetInfoFromImportTable(4) failed");
			return FALSE;
		}

		if (saAllZoneEditionList.GetCount() == 0) {
			sErrorMessage = "saAllZoneEditionList,GetCount()==0";
			return FALSE;
		}


		// Create ALL edition objects first

		for (int ed = 0; ed < saAllZoneEditionList.GetCount(); ed++) {

			PlanDataEdition *editem = new PlanDataEdition(saAllZoneEditionList[ed]);

			plan.editionList->Add(*editem);

			//CStringArray saLocationListDummy;
			//if (DB.NPGetInfoFromImportTable(IMPORTTABLEQUERYMODE_LOCATIONS_FOR_EDITION, sPublication, tPubDate, "", saAllZoneEditionList[ed], saLocationListDummy, sErrorMessage) == FALSE) {
			//	return FALSE;
			//}

			//for (int loc = 0; loc < saLocationList.GetCount(); loc++) {
			PlanDataPress *pressitem = new PlanDataPress(thisLocation);
			editem->pressList->Add(*pressitem);
			// Load specific Zone_Edition (all sections)
			DB.NPLoadImportTable(editem, sPublication, tPubDate, thisLocation, saAllZoneEditionList[ed], nPagesLoaded, sErrorMessage);
			//}

		}

		// Loaded order for this location

		// 1 Zone (SequenceNumber) (first on each location is zone and time master)
		// 2 Edition (first on each location is timed master)
		// 3 Section
		// 4 Pagenumber

		// Set PageIDs for all editions
		// ALL LOADED PAGES ARE UNIQUE!!!
		int nPageID = 1;

		for (int eds = 0; eds < plan.editionList->GetCount(); eds++) {
			PlanDataEdition *planDataEdition = &plan.editionList->GetAt(eds);
			for (int sections = 0; sections < planDataEdition->sectionList->GetCount(); sections++) {
				PlanDataSection *planDataSection = &planDataEdition->sectionList->GetAt(sections);
				for (int page = 0; page < planDataSection->pageList->GetCount(); page++) {
					PlanDataPage *planDataPage = &planDataSection->pageList->GetAt(page);
					planDataPage->pageID.Format("%d", nPageID++);
					planDataPage->masterPageID = planDataPage->pageID;
					planDataPage->uniquePage = PAGEUNIQUE_UNIQUE;
				}
			}
		}

		// Add extra editions

		for (int ex = 0; ex < g_npprefs.m_ExtraEditionList.GetCount(); ex++) {
			if (sPublication.CollateNoCase(g_npprefs.m_ExtraEditionList[ex].m_Publication) == 0 && g_npprefs.m_ExtraEditionList[ex].m_mineditions != "") {
				CStringArray aExEd;
				g_util.StringSplitter(g_npprefs.m_ExtraEditionList[ex].m_mineditions, ",;", aExEd);

				for (int exed = 0; exed < aExEd.GetCount(); exed++) {

					BOOL bHasEdition = FALSE;
					for (int alled = 0; alled < saAllZoneEditionList.GetCount(); alled++) {
						if (saAllZoneEditionList[alled] == aExEd[exed]) {
							bHasEdition = TRUE;
							break;
						}
					}

					if (bHasEdition == FALSE) {

						// Add new edition as common to first edition..

						saAllZoneEditionList.Add(aExEd[exed]);
						PlanDataEdition *firstEdition = &plan.editionList->GetAt(0);
						PlanDataEdition *neweditem = new PlanDataEdition();
						firstEdition->CloneEdition(neweditem);
						neweditem->editionName = aExEd[exed];
						neweditem->editionSequenceNumber = saAllZoneEditionList.GetCount();
						neweditem->masterEdition = FALSE;
						neweditem->zoneMaster = firstEdition->editionName;
						neweditem->IsTimed = FALSE;
						neweditem->timedTo = "";
						neweditem->timedFrom = "";

						plan.editionList->Add(*neweditem);

						// Copy all sections and pages from first edition to this new common edition
						for (int mainsec = 0; mainsec < firstEdition->sectionList->GetCount(); mainsec++) {
							PlanDataSection *secitem = &firstEdition->sectionList->GetAt(mainsec);
							PlanDataSection *newsecitem = new PlanDataSection(secitem->sectionName);
							neweditem->sectionList->Add(*newsecitem);

							for (int mainpage = 0; mainpage < secitem->pageList->GetCount(); mainpage++) {
								PlanDataPage *pageitem = &secitem->pageList->GetAt(mainpage);
								PlanDataPage *newpageitem = new PlanDataPage(pageitem->pageName);
								newsecitem->pageList->Add(*newpageitem);

								newpageitem->approve = pageitem->approve;
								newpageitem->comment = pageitem->comment;
								newpageitem->fileName = pageitem->fileName;
								newpageitem->hold = pageitem->hold;
								newpageitem->masterEdition = firstEdition->editionName;
								newpageitem->masterPageID = pageitem->pageID;	// Master to 1. ed
								newpageitem->miscint = pageitem->miscint;
								newpageitem->miscstring1 = pageitem->miscstring1;
								newpageitem->miscstring2 = pageitem->miscstring2;
								newpageitem->pageIndex = pageitem->pageIndex;
								newpageitem->pageType = pageitem->pageType;
								newpageitem->pagination = pageitem->pagination;
								newpageitem->priority = pageitem->priority;
								newpageitem->uniquePage = PAGEUNIQUE_COMMON;
								newpageitem->pageID.Format("%d", nPageID++);

								CString sColors = pageitem->miscstring1;

								PlanDataPageSeparation *newcolor = new PlanDataPageSeparation("PDF");
								newpageitem->colorList->Add(*newcolor);
								


							}
						}
					}
				}
			}
		}

		PlanDataEdition *planDataMasterEdition = plan.GetEditionObject(sMasterZone + "_" + sMasterEdition);
		if (planDataMasterEdition == NULL) {
			sInfo = _T("Master zone not avaialble yet - " + sMasterZone + "_" + sMasterEdition);
			sErrorMessage = _T("Master zone not avaialble yet");
			return FALSE;
		}

		planDataMasterEdition->editionComment = sMasterZone;
		planDataMasterEdition->masterEdition = TRUE;
		planDataMasterEdition->IsTimed = FALSE;
		planDataMasterEdition->timedFrom = _T("");
		planDataMasterEdition->zoneMaster = _T("");
		planDataMasterEdition->locationMaster = _T("");;
		planDataMasterEdition->editionSequenceNumber = 1;

		CStringArray saSlaveEditions;
		if (DB.NPGetTimedSlaveEditionsOnZoneLocation(sPublication, tPubDate, sMasterLocation, sMasterZone, sMasterEdition, saSlaveEditions, sErrorMessage) == FALSE) {
			sErrorMessage = _T("GetTimedSlaveEditionsOnZoneLocation() failed");
			return FALSE;
		}

		// Create all pages for timed editions for this zone

		if (saSlaveEditions.GetCount() > 0)
			planDataMasterEdition->IsTimed = TRUE;

		CString sTimedFrom = sMasterEdition;
		PlanDataEdition *prevplanDataSlaveEdition = NULL;
		for (int subed = 0; subed < saSlaveEditions.GetCount(); subed++) {
			PlanDataEdition *planDataSlaveEdition = plan.GetEditionObject(sMasterZone + "_" + saSlaveEditions[subed]);
			if (planDataSlaveEdition == NULL)
				continue; // ERROR
			planDataSlaveEdition->editionComment = sMasterZone;
			planDataSlaveEdition->IsTimed = TRUE;
			planDataSlaveEdition->timedFrom = sMasterZone + "_" + sMasterEdition;
			planDataSlaveEdition->masterEdition = FALSE;
			planDataSlaveEdition->zoneMaster = planDataMasterEdition->editionName;
			planDataSlaveEdition->locationMaster = sMasterLocation;
			planDataSlaveEdition->editionSequenceNumber = subed + 2;

			// First sub referenced back to master
			if (subed == 0) {
				planDataMasterEdition->timedTo = sMasterZone + "_" + saSlaveEditions[subed];
				planDataMasterEdition->IsTimed = TRUE;
				planDataSlaveEdition->timedFrom = planDataMasterEdition->editionName;
			}
			// Set previous sub reference to this sub
			if (subed > 0 && prevplanDataSlaveEdition != NULL) {
				prevplanDataSlaveEdition->timedTo = sMasterZone + "_" + saSlaveEditions[subed];
				planDataSlaveEdition->timedFrom = prevplanDataSlaveEdition->editionName;
			}

			prevplanDataSlaveEdition = planDataSlaveEdition;

			for (int sections = 0; sections < planDataMasterEdition->sectionList->GetCount(); sections++) {
				PlanDataSection *planDataMasterSection = &planDataMasterEdition->sectionList->GetAt(sections);
				PlanDataSection *planDataSlaveSection = planDataSlaveEdition->GetSectionObject(planDataMasterSection->sectionName);
				if (planDataSlaveSection == NULL) {
					planDataSlaveSection = new PlanDataSection(planDataMasterSection->sectionName);
					planDataSlaveEdition->sectionList->Add(*planDataSlaveSection);
				}

				for (int page = 0; page < planDataMasterSection->pageList->GetCount(); page++) {
					PlanDataPage *planDataMasterPage = &planDataMasterSection->pageList->GetAt(page); // From

					BOOL bFound = FALSE;

					// See if page exists as unique - otherwise create as common
					for (int slavepage = 0; slavepage < planDataSlaveSection->pageList->GetCount(); slavepage++) {
						PlanDataPage *planDataSlavePage = &planDataSlaveSection->pageList->GetAt(slavepage);
						if (planDataSlavePage->pageName == planDataMasterPage->pageName) {
							bFound = TRUE;
							break;
						}
					}
					if (bFound == FALSE) {
						PlanDataPage *planDataSlavePage = new PlanDataPage();							// To
						planDataSlavePage->approve = planDataMasterPage->approve;
						planDataSlavePage->pageType = planDataMasterPage->pageType;
						planDataSlavePage->hold = planDataMasterPage->hold;
						planDataSlavePage->priority = planDataMasterPage->priority;
						planDataSlavePage->version = planDataMasterPage->version;
						planDataSlavePage->uniquePage = PAGEUNIQUE_COMMON;
						planDataSlavePage->masterEdition = planDataMasterEdition->editionName;
						planDataSlavePage->pageID.Format("%d", nPageID++);
						planDataSlavePage->masterPageID = planDataMasterPage->pageID;
						planDataSlavePage->pageName = planDataMasterPage->pageName;
						planDataSlavePage->pageIndex = planDataMasterPage->pageIndex;
						planDataSlavePage->pagination = planDataMasterPage->pagination;
						planDataSlavePage->fileName = planDataMasterPage->fileName;
						planDataSlavePage->comment = planDataMasterPage->comment;
						CString sColors = planDataMasterPage->miscstring1;
						planDataSlavePage->miscstring1 = sColors;

						PlanDataPageSeparation *sep = new PlanDataPageSeparation("PDF");
						planDataSlavePage->colorList->Add(*sep);
						
						planDataSlaveSection->pageList->Add(*planDataSlavePage);
						planDataSlaveSection->pagesInSection++;
					}
				}
			}
		}

		// Master zone complete!

	// Now iterate over sub-zones

		CStringArray saSlaveZones;
		if (DB.NPGetSlaveZones(sPublication, tPubDate, sMasterZone, thisLocation, saSlaveZones, sErrorMessage) == FALSE) {
			sErrorMessage = _T("GetSlaveZones() failed");
			return FALSE;
		}

		CString sSlaveZoneMasterEdition = "1";	// ALWAYS!
		for (int z = 0; z < saSlaveZones.GetCount(); z++) {

			PlanDataEdition *planDataSlaveEdition = plan.GetEditionObject(saSlaveZones[z] + "_" + sSlaveZoneMasterEdition);
			if (planDataSlaveEdition == NULL)
				continue;

			// Create first timed edition of slave zone
			planDataSlaveEdition->editionComment = saSlaveZones[z];
			planDataSlaveEdition->masterEdition = FALSE;
			planDataSlaveEdition->IsTimed = FALSE;
			planDataSlaveEdition->timedFrom = _T("");
			planDataSlaveEdition->zoneMaster = sMasterZone;
			planDataSlaveEdition->locationMaster = sMasterLocation;
			planDataSlaveEdition->editionSequenceNumber = 1;

			/*	if (m_DB.GetInfoFromImportTable(IMPORTTABLEQUERYMODE_LOCATIONS_FOR_EDITION, sPublication, tPubDate, "", planDataSlaveEdition->editionName, saLocationList, szErrorMessage) == FALSE) {
					return FALSE;
				}
	*/
			for (int sections = 0; sections < planDataMasterEdition->sectionList->GetCount(); sections++) {
				PlanDataSection *planDataMasterSection = &planDataMasterEdition->sectionList->GetAt(sections);
				PlanDataSection *planDataSlaveSection = planDataSlaveEdition->GetSectionObject(planDataMasterSection->sectionName);
				if (planDataSlaveSection == NULL) {
					planDataSlaveSection = new PlanDataSection(planDataMasterSection->sectionName);
					planDataSlaveEdition->sectionList->Add(*planDataSlaveSection);
				}
				for (int page = 0; page < planDataMasterSection->pageList->GetCount(); page++) {
					PlanDataPage *planDataMasterPage = &planDataMasterSection->pageList->GetAt(page); // From

					BOOL bFound = FALSE;

					// See if page exists as unique - otherwise create as common
					for (int slavepage = 0; slavepage < planDataSlaveSection->pageList->GetCount(); slavepage++) {
						PlanDataPage *planDataSlavePage = &planDataSlaveSection->pageList->GetAt(slavepage);
						if (planDataSlavePage->pageName == planDataMasterPage->pageName) {
							bFound = TRUE;
							break;
						}
					}
					if (bFound == FALSE) {
						PlanDataPage *planDataSlavePage = new PlanDataPage();							// To
						planDataSlavePage->approve = planDataMasterPage->approve;
						planDataSlavePage->pageType = planDataMasterPage->pageType;
						planDataSlavePage->hold = planDataMasterPage->hold;
						planDataSlavePage->priority = planDataMasterPage->priority;
						planDataSlavePage->version = planDataMasterPage->version;
						planDataSlavePage->uniquePage = thisLocation == sMasterLocation ? PAGEUNIQUE_COMMON : PAGEUNIQUE_FORCED;
						planDataSlavePage->masterEdition = planDataMasterEdition->editionName;
						planDataSlavePage->pageID.Format("%d", nPageID++);
						planDataSlavePage->masterPageID = planDataMasterPage->pageID;
						planDataSlavePage->pageName = planDataMasterPage->pageName;
						planDataSlavePage->pageIndex = planDataMasterPage->pageIndex;
						planDataSlavePage->pagination = planDataMasterPage->pagination;
						planDataSlavePage->fileName = planDataMasterPage->fileName;
						planDataSlavePage->comment = planDataMasterPage->comment;
						CString sColors = planDataMasterPage->miscstring1;
						planDataSlavePage->miscstring1 = sColors;

						PlanDataPageSeparation *sep = new PlanDataPageSeparation("PDF");
						planDataSlavePage->colorList->Add(*sep);
					
						planDataSlaveSection->pageList->Add(*planDataSlavePage);
						planDataSlaveSection->pagesInSection++;
					}
				}
			}

			// First edition (timed) of zones complete

			CStringArray saSlaveZoneEditions;
			if (DB.NPGetTimedSlaveEditionsOnZoneLocation(sPublication, tPubDate, thisLocation, saSlaveZones[z], sSlaveZoneMasterEdition, saSlaveZoneEditions, sErrorMessage) == FALSE)
				return FALSE;
			if (saSlaveZoneEditions.GetCount() > 0)
				planDataSlaveEdition->IsTimed = TRUE;

			// Handle timed subs of zone sub..
			PlanDataEdition *prevplanDataSlaveZoneEdition = NULL;
			for (int subed = 0; subed < saSlaveZoneEditions.GetCount(); subed++) {
				PlanDataEdition *planDataSlaveZoneEdition = plan.GetEditionObject(saSlaveZones[z] + "_" + saSlaveZoneEditions[subed]);
				if (planDataSlaveZoneEdition == NULL)
					continue; // ERROR

				planDataSlaveZoneEdition->editionComment = saSlaveZones[z];
				planDataSlaveZoneEdition->IsTimed = TRUE;
				planDataSlaveZoneEdition->timedFrom = planDataSlaveEdition->editionName;
				planDataSlaveZoneEdition->masterEdition = FALSE;
				planDataSlaveZoneEdition->zoneMaster = planDataMasterEdition->editionName;
				planDataSlaveZoneEdition->locationMaster = sMasterLocation;
				planDataSlaveZoneEdition->editionSequenceNumber = subed + 2;

				// First sub referenced back to master in zone
				if (subed == 0) {
					planDataSlaveZoneEdition->timedTo = planDataSlaveEdition->editionName;
					planDataMasterEdition->IsTimed = TRUE;
				}
				// Set previous sub reference to this sub
				if (subed > 0 && prevplanDataSlaveZoneEdition != NULL)
					prevplanDataSlaveZoneEdition->timedTo = planDataSlaveZoneEdition->editionName;

				prevplanDataSlaveZoneEdition = planDataSlaveZoneEdition;

				for (int sections = 0; sections < planDataSlaveEdition->sectionList->GetCount(); sections++) {
					PlanDataSection *planDataMasterSection = &planDataSlaveEdition->sectionList->GetAt(sections);
					PlanDataSection *planDataSlaveSection = planDataSlaveZoneEdition->GetSectionObject(planDataMasterSection->sectionName);
					if (planDataSlaveSection == NULL) {
						planDataSlaveSection = new PlanDataSection(planDataMasterSection->sectionName);
						planDataSlaveZoneEdition->sectionList->Add(*planDataSlaveSection);
					}

					for (int page = 0; page < planDataMasterSection->pageList->GetCount(); page++) {
						PlanDataPage *planDataMasterPage = &planDataMasterSection->pageList->GetAt(page); // From

						BOOL bFound = FALSE;

						// See if page exists as unique - otherwise create as common
						for (int slavepage = 0; slavepage < planDataSlaveSection->pageList->GetCount(); slavepage++) {
							PlanDataPage *planDataSlavePage = &planDataSlaveSection->pageList->GetAt(slavepage);
							if (planDataSlavePage->pageName == planDataMasterPage->pageName) {
								bFound = TRUE;
								break;
							}
						}
						if (bFound == FALSE) {
							PlanDataPage *planDataSlavePage = new PlanDataPage();							// To
							planDataSlavePage->approve = planDataMasterPage->approve;
							planDataSlavePage->pageType = planDataMasterPage->pageType;
							planDataSlavePage->hold = planDataMasterPage->hold;
							planDataSlavePage->priority = planDataMasterPage->priority;
							planDataSlavePage->version = planDataMasterPage->version;
							planDataSlavePage->uniquePage = PAGEUNIQUE_COMMON;
							planDataSlavePage->masterEdition = planDataMasterEdition->editionName;
							planDataSlavePage->pageID.Format("%d", nPageID++);
							if (planDataMasterPage->uniquePage == PAGEUNIQUE_UNIQUE)
								planDataSlavePage->masterPageID = planDataMasterPage->pageID;
							else
								planDataSlavePage->masterPageID = planDataMasterPage->masterPageID;
							planDataSlavePage->pageName = planDataMasterPage->pageName;
							planDataSlavePage->pageIndex = planDataMasterPage->pageIndex;
							planDataSlavePage->pagination = planDataMasterPage->pagination;
							planDataSlavePage->fileName = planDataMasterPage->fileName;
							planDataSlavePage->comment = planDataMasterPage->comment;
							CString sColors = planDataMasterPage->miscstring1;
							planDataSlavePage->miscstring1 = sColors;

							PlanDataPageSeparation *sep = new PlanDataPageSeparation("PDF");
							planDataSlavePage->colorList->Add(*sep);
							
							planDataSlaveSection->pageList->Add(*planDataSlavePage);
							planDataSlaveSection->pagesInSection++;
						}
					}
				}


			}
		}

		plan.SortPlanPages();

		sTitle.Format("NewsPilot-%s-%s-%.4d%.2d%.2d%.2d%.2d%.2d", sPublication, sPubDate, tNow.GetYear(), tNow.GetMonth(), tNow.GetDay(), tNow.GetHour(), tNow.GetMinute(), tNow.GetSecond());
		sFinalOutputName = sOutputFolder + _T("\\") + sTitle + _T(".xml");
		if (NPGenerateXML(sFinalOutputName, &plan, FALSE, TRUE, bIsLast) == FALSE) {
			sErrorMessage = _T("GenerateXML() failed");
			ret = FALSE;
		}


		plan.DisposePlanData();

		break;

	}


	return ret;
}

int NPGenerateXML(CString sXmlFile, PlanData *plan, BOOL hasSheets, BOOL bForcePDF, BOOL bIsLast)
{
	CUtils util;
	CString s, s2;
	CMarkup xml;
	CString sMainLocation("Default");

	BOOL bMustProduce = FALSE;
	int nPageID = 1;

	xml.SetDoc("<?xml version=\"1.0\" encoding=\"ISO-8859-1\"?>\r\n");
	xml.AddElem("Plan");
	xml.AddAttrib("Version", util.Int2String(plan->version));

	xml.AddAttrib("Command", "ForceUnapplied");


	if (bIsLast == FALSE)
		xml.AddAttrib("IgnorePostCommand", "1");

	xml.AddAttrib("ID", plan->planID);
	xml.AddAttrib("Name", plan->planName);
	CTime tNow = CTime::GetCurrentTime();
	CString sTime;
	sTime.Format("%.4d-%.2d-%.2dT%.2d:%.2d:%.2d", tNow.GetYear(), tNow.GetMonth(), tNow.GetDay(), tNow.GetHour(), tNow.GetMinute(), tNow.GetSecond());
	xml.AddAttrib("UpdateTime", sTime);
	xml.AddAttrib("Sender", "NPImporter");
	xml.AddAttrib("xmlns", "http://tempuri.org/ImportCenter.xsd");

	xml.AddChildElem("Publication");
	xml.IntoElem();

	sTime.Format("%.4d-%.2d-%.2d", plan->publicationDate.GetYear(), plan->publicationDate.GetMonth(), plan->publicationDate.GetDay());
	xml.AddAttrib("PubDate", sTime);

	int nWeekRef = 0;
	int nn = plan->pagenameprefix.FindOneOf("-_ ");
	if (nn != -1)
		if (plan->pagenameprefix[nn + 1] == '0' && plan->pagenameprefix[nn + 2] == '0')
			nWeekRef = atoi(plan->pagenameprefix.Mid(nn + 1));

	s.Format("%d", nWeekRef);
	xml.AddAttrib("WeekReference", s);

	xml.AddAttrib("Name", plan->publicationName);
	xml.AddAttrib("Alias", plan->publicationAlias);

	xml.AddAttrib("Customer", plan->customername);
	xml.AddAttrib("CustomerAlias", plan->customeralias);

	xml.AddChildElem("Issues");
	xml.IntoElem();
	xml.AddChildElem("Issue");
	xml.IntoElem();
	xml.AddAttrib("Name", "Main");
	xml.AddChildElem("Editions");
	xml.IntoElem();

	for (int eds = 0; eds < plan->editionList->GetCount(); eds++) {
		PlanDataEdition *planDataEdition = &plan->editionList->GetAt(eds);

		BOOL bMustProduceEdition = TRUE;
		/*	for (int press=0; press<planDataEdition->pressList->GetCount(); press++) {
				PlanDataPress *planDataPress = &planDataEdition->pressList->GetAt(press);
				if (sPressToUse.CompareNoCase(planDataPress->pressName) == 0) {
					bMustProduceEdition = TRUE;
					break;
				}
			} */
		if (bMustProduceEdition) {

			xml.AddChildElem("Edition");
			xml.IntoElem();
			xml.AddAttrib("Name", planDataEdition->editionName);
			CString sCopies = util.Int2String(planDataEdition->editionCopy);
			xml.AddAttrib("EditionCopies", sCopies != "0" ? sCopies : "1");
			xml.AddAttrib("IsTimed", planDataEdition->IsTimed ? "true" : "false");
			xml.AddAttrib("TimedFrom", planDataEdition->timedFrom);
			xml.AddAttrib("TimedTo", planDataEdition->timedTo);
			xml.AddAttrib("EditionOrderNumber", planDataEdition->editionSequenceNumber);
			xml.AddAttrib("EditionComment", planDataEdition->editionComment);
			xml.AddAttrib("ZoneMasterEdition", planDataEdition->masterEdition ? "" : planDataEdition->zoneMaster);

			xml.AddChildElem("IntendedPresses");
			xml.IntoElem();

			for (int press = 0; press < planDataEdition->pressList->GetCount(); press++) {
				PlanDataPress *planDataPress = &planDataEdition->pressList->GetAt(press);

				//if (sPressToUse.CompareNoCase(planDataPress->pressName) == 0) {
				bMustProduce = TRUE;
				if (planDataPress->pressName == "")
					planDataPress->pressName = sMainLocation;
				xml.AddChildElem("IntendedPress");
				xml.IntoElem();
				xml.AddAttrib("Name", planDataPress->pressName);
				xml.AddAttrib("PressTime", util.CTime2String(planDataPress->pressRunTime));
				xml.AddAttrib("PlateCopies", util.Int2String(planDataPress->copies));
				xml.AddAttrib("Copies", util.Int2String(planDataPress->paperCopies));
				xml.AddAttrib("PostalUrl", planDataPress->postalUrl);

				xml.OutOfElem();
				//	}
			}
			xml.OutOfElem();
			xml.AddChildElem("Sections");
			xml.IntoElem();

			for (int sections = 0; sections < planDataEdition->sectionList->GetCount(); sections++) {
				PlanDataSection *planDataSection = &planDataEdition->sectionList->GetAt(sections);

				xml.AddChildElem("Section");
				xml.IntoElem();
				xml.AddAttrib("Name", planDataSection->sectionName);
				xml.AddChildElem("Pages");
				xml.IntoElem();

				for (int page = 0; page < planDataSection->pageList->GetCount(); page++) {

					PlanDataPage *planDataPage = &planDataSection->pageList->GetAt(page);

					xml.AddChildElem("Page");
					xml.IntoElem();

					xml.AddAttrib("Name", planDataPage->pageName);

					xml.AddAttrib("Pagination", planDataPage->pageIndex);
					xml.AddAttrib("PageType", szPageTypes[planDataPage->pageType]);

					// Straighten out inconsistent unique flag.....
					int nUnique = planDataPage->uniquePage;
					if (planDataPage->pageID == planDataPage->masterPageID)
						nUnique = PAGEUNIQUE_UNIQUE;
					xml.AddAttrib("Unique", szPageUniques[nUnique]);

					xml.AddAttrib("PageIndex", planDataPage->pageIndex);
					xml.AddAttrib("FileName", planDataPage->fileName);
					xml.AddAttrib("Comment", planDataPage->comment);
					xml.AddAttrib("Approve", planDataPage->approve ? "1" : "0");
					xml.AddAttrib("Hold", planDataPage->hold ? "1" : "0");

					/*
					s.Format("%d", planDataPage->miscint);
					xml.AddAttrib("MiscInt", s);
					xml.AddAttrib("MiscString1",planDataPage->miscstring1);
					xml.AddAttrib("MiscString2",planDataPage->miscstring2);
					*/

					xml.AddAttrib("PageID", planDataPage->pageID);
					xml.AddAttrib("MasterPageID", planDataPage->masterPageID != "" ? planDataPage->masterPageID : planDataPage->pageID);
					xml.AddAttrib("Priority", util.Int2String(planDataPage->priority));
					xml.AddAttrib("Version", util.Int2String(planDataPage->version));

					//xml.AddAttrib("MasterEdition", sMasterEdition);

					xml.AddChildElem("Separations");
					xml.IntoElem();

					if (bForcePDF) {
						xml.AddChildElem("Separation");
						xml.IntoElem();
						xml.AddAttrib("Name", "PDF");
						xml.OutOfElem();// </Separation>

					}
					else {
						for (int col = 0; col < planDataPage->colorList->GetCount(); col++) {
							PlanDataPageSeparation *planDataPageSeparation = &planDataPage->colorList->GetAt(col);
							xml.AddChildElem("Separation");
							xml.IntoElem();
							xml.AddAttrib("Name", planDataPageSeparation->colorName);
							xml.OutOfElem();// </Separation>
						} // for (int col..
					}
					xml.OutOfElem();// </Separations>

					xml.OutOfElem();// </Page>
				} // for (int page..
				xml.OutOfElem();// </Pages>

				xml.OutOfElem();// </Section>
			} // for (int sections..

			xml.OutOfElem();// </Sections>

			if (hasSheets) {

				xml.AddChildElem("Sheets");
				xml.IntoElem();

				for (int sheets = 0; sheets < planDataEdition->sheetList->GetCount(); sheets++) {
					PlanDataSheet *planDataSheet = &planDataEdition->sheetList->GetAt(sheets);

					xml.AddChildElem("Sheet");
					xml.IntoElem();
					xml.AddAttrib("Name", planDataSheet->sheetName);
					xml.AddAttrib("Template", planDataSheet->templateName);
					xml.AddAttrib("MarkGroups", planDataSheet->markGroups);
					xml.AddAttrib("PagesOnPlate", util.Int2String(planDataSheet->pagesOnPlate));

					PlanDataSheetSide *planDataSheetSideFront = planDataSheet->frontSheet;

					xml.AddChildElem("SheetFrontItems");
					xml.IntoElem();
					xml.AddAttrib("SortingPosition", planDataSheetSideFront->SortingPosition);
					xml.AddAttrib("PressTower", planDataSheetSideFront->PressTower);
					xml.AddAttrib("PressZone", planDataSheetSideFront->PressZone);
					xml.AddAttrib("PressHighLow", planDataSheetSideFront->PressHighLow);
					xml.AddAttrib("ActiveCopies", util.Int2String(planDataSheetSideFront->ActiveCopies));

					for (int sheetitems = 0; sheetitems < planDataSheetSideFront->sheetItems->GetCount(); sheetitems++) {
						PlanDataSheetItem *planDataSheetItem = &planDataSheetSideFront->sheetItems->GetAt(sheetitems);
						xml.AddChildElem("SheetFrontItem");
						xml.IntoElem();
						xml.AddAttrib("PageID", planDataSheetItem->pageID);
						xml.AddAttrib("MasterPageID", planDataSheetItem->masterPageID);
						xml.AddAttrib("PosX", util.Int2String(planDataSheetItem->pagePositionX));
						xml.AddAttrib("PosY", util.Int2String(planDataSheetItem->pagePositionY));
						xml.OutOfElem();// </SheetFrontItem>
					}
					xml.AddChildElem("PressCylindersFront");
					xml.IntoElem();

					for (int cylinders = 0; cylinders < planDataSheetSideFront->pressCylinders->GetCount(); cylinders++) {
						PlanDataSheetPressCylinder *planDataSheetPressCylinder = &planDataSheetSideFront->pressCylinders->GetAt(cylinders);
						xml.AddChildElem("PressCylinderFront");
						xml.IntoElem();

						xml.AddAttrib("Name", planDataSheetPressCylinder->pressCylinder);
						xml.AddAttrib("Color", planDataSheetPressCylinder->colorName);
						xml.AddAttrib("FormID", planDataSheetPressCylinder->formID);
						xml.AddAttrib("SortingPosition", planDataSheetPressCylinder->sortingPosition);
						xml.OutOfElem();// </PressCylinderFront>
					} // for (int cylinders..

					xml.OutOfElem();// </PressCylindersFront>
				  // for (int sheetitems..
					xml.OutOfElem();// </SheetFrontItems>

					if (planDataSheet->hasback) {
						PlanDataSheetSide *planDataSheetSideBack = planDataSheet->backSheet;

						xml.AddChildElem("SheetBackItems");
						xml.IntoElem();
						xml.AddAttrib("SortingPosition", planDataSheetSideBack->SortingPosition);
						xml.AddAttrib("PressTower", planDataSheetSideBack->PressTower);
						xml.AddAttrib("PressZone", planDataSheetSideBack->PressZone);
						xml.AddAttrib("PressHighLow", planDataSheetSideBack->PressHighLow);
						xml.AddAttrib("ActiveCopies", util.Int2String(planDataSheetSideBack->ActiveCopies));

						for (int sheetitems = 0; sheetitems < planDataSheetSideBack->sheetItems->GetCount(); sheetitems++) {
							PlanDataSheetItem *planDataSheetItem = &planDataSheetSideBack->sheetItems->GetAt(sheetitems);
							xml.AddChildElem("SheetBackItem");
							xml.IntoElem();
							xml.AddAttrib("PageID", planDataSheetItem->pageID);
							xml.AddAttrib("MasterPageID", planDataSheetItem->masterPageID);
							xml.AddAttrib("PosX", util.Int2String(planDataSheetItem->pagePositionX));
							xml.AddAttrib("PosY", util.Int2String(planDataSheetItem->pagePositionY));
							xml.OutOfElem();// </SheetFrontItem>
						}
						xml.AddChildElem("PressCylindersBack");
						xml.IntoElem();

						for (int cylinders = 0; cylinders < planDataSheetSideBack->pressCylinders->GetCount(); cylinders++) {
							PlanDataSheetPressCylinder *planDataSheetPressCylinder = &planDataSheetSideBack->pressCylinders->GetAt(cylinders);
							xml.AddChildElem("PressCylinderBack");
							xml.IntoElem();
							xml.AddAttrib("Name", planDataSheetPressCylinder->pressCylinder);
							xml.AddAttrib("Color", planDataSheetPressCylinder->colorName);
							xml.AddAttrib("FormID", planDataSheetPressCylinder->formID);
							xml.AddAttrib("SortingPosition", planDataSheetPressCylinder->sortingPosition);
							xml.OutOfElem();// </PressCylinderBack>
						} // for (int cylinders..
						xml.OutOfElem();// </PressCylindersBack>
					  // for (int sheetitems..

						xml.OutOfElem();// </SheetBackItems>
					} // end if hasBack
					xml.OutOfElem();// </Sheet>
				} // for (int sheets..
				xml.OutOfElem();// </Sheets>
			}
			xml.OutOfElem();// </Edition>
		}

	} // for (int eds..
	xml.OutOfElem();// </Editions>
	xml.OutOfElem();// </Issue>
	xml.OutOfElem();// </Issues>
	xml.OutOfElem();// </Publication>

	CString sCmdBuffer = xml.GetDoc();
	HANDLE hHndlWrite;
	DWORD nBytesWritten;

	if (bMustProduce) {

		

		if ((hHndlWrite = ::CreateFile(sXmlFile, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL)) == INVALID_HANDLE_VALUE) {
			//AddToLog("Cannot create XML plan file "+ sXmlFile);
			CString sErr = util.GetLastWin32Error();
			util.Logprintf("ERROR: GenerateXML - CreateFile(%s) failed - %s", sXmlFile, sErr);
			return FALSE;
		}

		if (::WriteFile(hHndlWrite, sCmdBuffer, sCmdBuffer.GetLength(), &nBytesWritten, NULL) == FALSE) {
			//AddToLog("Cannot write XML plan file "+ sXmlFile);
			CloseHandle(hHndlWrite);
			util.Logprintf("ERROR: GenerateXML - WriteFile() failed");
			return FALSE;
		}
		CloseHandle(hHndlWrite);
	}
	return TRUE;
}
