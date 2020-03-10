#include "stdafx.h"
#include <afxmt.h>
#include <afxtempl.h>
#include <direct.h>
#include <iostream>
#include "DatabaseManager.h"
#include "Defs.h"
#include "Utils.h"
#include "Prefs.h"
#include "Markup.h"
#include "WorkerThread.h"
#include <winnetwk.h>
#include "ParseXML.h"
#include "ImportXMLPages.h"
#include "NPParser.h"
#include "PPIParser.h"

extern CPrefs	g_prefs;
extern CUtils g_util;
extern BOOL		g_bRunning;


BOOL	g_BreakSignal = FALSE;
BOOL	g_KillSignal = FALSE;
CTime	g_tLastPoll = CTime::GetCurrentTime();
int		g_TotalNewImports = 0;
int		g_TotalUpdateImports = 0;
int		g_TotalErrorImports = 0;

FILEINFOSTRUCT aFilesReady[MAXFILES];
CDatabaseManager m_pollDB;
CPageTableEntry pageTableEntry;
CPageTableList PageTable;


void WorkerLoop()
{
	CString sErrMsg, sInfo, s;
	PLANSTRUCT *plan = new PLANSTRUCT();

	CString sErr = _T("");

	BOOL	bDoBreak = false;
	BOOL bConncetionOK = TRUE;
	CString sErrorMessage;
	CString sPreviousImportFileName = _T("");

	int nFolderNumber = 1;

	if (m_pollDB.InitDB(g_prefs.m_DBserver, g_prefs.m_Database, g_prefs.m_DBuser, g_prefs.m_DBpassword, g_prefs.m_IntegratedSecurity, sErrorMessage) == FALSE) {
		sErr.Format("Unable to connect to database - %s", sErrorMessage);
		g_util.Logprintf("%s", sErr);
	}
	g_util.Logprintf("Pollthread starting..");
	// Reload caches..
	if (m_pollDB.LoadAllPrefs(sErrorMessage) == FALSE)
		g_util.Logprintf("ERROR: LoadAllPrefs() - %s", sErrorMessage);

	if (m_pollDB.LoadPPITanslations(sErrorMessage) == FALSE)
		g_util.Logprintf("ERROR: LoadPPITanslations() - %s", sErrorMessage);

	int nFiles = 0;
	BOOL bReConnect = FALSE;
	BOOL bUsebackup = FALSE;
	int nTick = 0;
	g_util.Logprintf("WorkerThread() - DB Preferences loaded - entering main loop...");
	BOOL	bFirstFile = TRUE;

	int nTicksSinceLastFile = 0;

	m_pollDB.UpdateService(SERVICESTATUS_RUNNING, _T(""), g_bRunning ?  _T("1") : _T("0"), sErrorMessage);

	do {
		if (g_bRunning == FALSE)
			break;

		if (bReConnect) {

			for (int i = 0; i < 40; i++) {
				::Sleep(100);
				if (g_bRunning == FALSE)
					break;
			}
			if (g_bRunning == FALSE)
				break;

			bReConnect = FALSE;

			if (m_pollDB.InitDB(g_prefs.m_DBserver, g_prefs.m_Database, g_prefs.m_DBuser, g_prefs.m_DBpassword, g_prefs.m_IntegratedSecurity, sErrorMessage) == FALSE) {
				bReConnect = TRUE;
				g_util.Logprintf("ERROR: Database error - %s", sErrorMessage);
			}
		}


		if ((g_util.Reconnect(g_prefs.m_inputfolder, g_prefs.m_inputfolderusecurrentuser ? "" : g_prefs.m_inputfolderusername, g_prefs.m_inputfolderusecurrentuser ? "" : g_prefs.m_inputfolderpassword)) == FALSE) {
			g_util.Logprintf("Input folder not found  - %s", g_prefs.m_inputfolder);
			bReConnect = TRUE;
			continue;
		}

		if (g_prefs.m_saveonerror) {
			if ((g_util.CheckFolderWithPing(g_prefs.m_errorfolder)) == FALSE) {
				g_util.Logprintf("Error folder not found  - %s", g_prefs.m_errorfolder);
				bReConnect = TRUE;
				continue;
			}
		}

		if (g_prefs.m_saveafterdone) {
			if ((g_util.CheckFolderWithPing(g_prefs.m_donefolder)) == FALSE) {
				g_util.Logprintf("Done folder not found  - %s", g_prefs.m_donefolder);
				bReConnect = TRUE;
				continue;
			}
		}

		int nPollTime = g_prefs.m_polltime;
		if (nPollTime) {
			for (int tm = 0; tm < 100; tm++) {
				if (g_bRunning == FALSE)
					break;
				::Sleep(10 * nPollTime);
			}
		}
		if (g_bRunning == FALSE)
			break;


		nTick++;
		if ((nTick) % 20 == 0)
			m_pollDB.UpdateService(SERVICESTATUS_RUNNING, _T(""), _T(""), sErrorMessage);

		CString sInputFolder = nFolderNumber == 2 && g_prefs.m_inputfolder2  != "" ? g_prefs.m_inputfolder2 : g_prefs.m_inputfolder;

		BOOL bFileInWorkFolder = FALSE;

		int	nFiles = g_util.ScanDirMaxFilesEx(sInputFolder, _T("*.*"), aFilesReady, MAXFILES);

		if (nFiles < 0) {
			bReConnect = TRUE;
		}

		// Every 50'th polling - reset any hanging polling locks..

		if (nFiles == 0) {
			nTicksSinceLastFile++;
			if (nTicksSinceLastFile > 50) {
				m_pollDB.ResetAllPollLocks(sErrorMessage);
				nTicksSinceLastFile = 0;
			}

		}


		if (nFiles <= 0) {
			nFolderNumber = nFolderNumber == 1 && g_prefs.m_inputfolder2  != "" ? 2 : 1; // toggle folder..
			continue;
		}

		nTicksSinceLastFile = 0;

		for (int it = 0; it < nFiles; it++) {

			FILEINFOSTRUCT fi = aFilesReady[it];

			if (g_bRunning == FALSE)
				break;			

			CString sFullInputPath = sInputFolder + "\\" + fi.sFileName;

			// Allready processed??
			if (g_util.FileExist(sFullInputPath) == FALSE)
				continue;

			// Still tampered with??
			if (g_util.LockCheck(sFullInputPath, TRUE) == FALSE)
				continue;

			//////////////////////////////////////
			// process this job (fi.sFileName)
			//////////////////////////////////////

			DWORD dwInitialSize = fi.nFileSize;
			CTime tInitialWriteTime = fi.tWriteTime;

			//////////////////////////////////////
			// process this job
			//////////////////////////////////////

			// Still file growth??
			if (g_prefs.m_stabletime > 0 && it == 0)
				Sleep(g_prefs.m_stabletime * 1000);

			if (g_prefs.m_stabletime > 0) {
				if (g_util.GetFileSize(sFullInputPath) != dwInitialSize)
					continue;
				if (g_util.GetWriteTime(sFullInputPath) != tInitialWriteTime)
					continue;
			}

			CString sTimeStamp;
			CTime tNow = CTime::GetCurrentTime();
			sTimeStamp.Format("%.4d%.2d%.2dT%.2d%.2d%.2d", tNow.GetYear(), tNow.GetMonth(), tNow.GetDay(), tNow.GetHour(), tNow.GetMinute(), tNow.GetSecond());
			if (g_prefs.m_saveafterdone) {
				CString sDoneFile = nFolderNumber == 2 ? g_prefs.m_donefolder2 + "\\" + fi.sFileName : g_prefs.m_donefolder + "\\" + fi.sFileName;
				if (g_prefs.m_addtimestampindonefolder) {
					CTime tNow = CTime::GetCurrentTime();
					sDoneFile.Format("%s\\%s-%s.%s", nFolderNumber == 2 ? g_prefs.m_donefolder2 : g_prefs.m_donefolder, g_util.GetFileName(fi.sFileName, TRUE), sTimeStamp, g_util.GetExtension(fi.sFileName));
				}
				::CopyFile(sFullInputPath, sDoneFile, FALSE);
			}

			if ((nFolderNumber == 1 && g_prefs.m_copyfolder != "") || (nFolderNumber == 2 && g_prefs.m_copyfolder2 != "")) {
				CString sCopyFile = nFolderNumber == 2 ? g_prefs.m_copyfolder2 + "\\" + fi.sFileName : g_prefs.m_copyfolder + "\\" + fi.sFileName;
				::CopyFile(sFullInputPath, sCopyFile, FALSE);
			}

			// Planfile to Evry...
			if (g_prefs.m_secondcopy && g_prefs.m_secondfolder) {
				::CopyFile(sFullInputPath, g_prefs.m_secondfolder + _T("\\") + fi.sFileName, FALSE);
				if (m_pollDB.InsertFileSendRequest(g_prefs.m_secondfolder + _T("\\") + fi.sFileName, sErrorMessage) == FALSE)
					g_util.Logprintf("ERROR: m_pollDB.InsertFileSendRequest() - %s", sErrorMessage);
			}


			sPreviousImportFileName = fi.sFileName;

			BOOL bIsCCxml = IsInternalXML(sFullInputPath, sErrorMessage);

			BOOL bIsNewsPilotPlan = FALSE;

			CString sErrorPath = nFolderNumber == 2 ? g_prefs.m_errorfolder2 + "\\" + fi.sFileName :
				g_prefs.m_errorfolder + "\\" + fi.sFileName;

			// Handle conversion from native formats..
			if (bIsCCxml == FALSE) {
				CString  sTempName = sFullInputPath + "_convert";
				g_util.MoveFile(sFullInputPath, sTempName);
				CString sOutputFile = _T("");
				BOOL bConvertOK;
				
				if (nFolderNumber == 1)
					bConvertOK = ProcessNPFile(m_pollDB, sTempName, sOutputFile, sErrorMessage);
				else 
					bConvertOK = ProcessPPIFile(m_pollDB, sTempName, sOutputFile, sErrorMessage);

				if (bConvertOK == FALSE || g_util.FileExist(sOutputFile)== FALSE) {
					if (g_prefs.m_saveonerror)
						::MoveFileEx(sTempName, sErrorPath, MOVEFILE_COPY_ALLOWED | MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH);
					else
						::DeleteFile(sTempName);
					g_util.Logprintf("ERROR: Could not convert file %s - %s", fi.sFileName, sErrorMessage);

					::DeleteFile(sOutputFile);
					CString sErrorMessage2 = sErrorMessage;
					m_pollDB.UpdateService(SERVICESTATUS_HASERROR, fi.sFileName, "Newspilot file parsing error", sErrorMessage);
					m_pollDB.InsertLogEntry(EVENTCODE_PLANERROR, _T("NewsPilot"), fi.sFileName, _T("Newspilot file parsing error - ") + sErrorMessage2, 0, 0, 0, _T(""), sErrorMessage);
					if (g_prefs.m_senderrormail && nFolderNumber == 1 || g_prefs.m_senderrormail2 && nFolderNumber == 2)
						g_util.SendMail("Plan import error - " + fi.sFileName, "Plan parsing error " + sErrorMessage, 
							nFolderNumber == 1 ? g_prefs.m_emailreceivers : g_prefs.m_emailreceivers2) ;
					continue;
				}
				bIsNewsPilotPlan = TRUE;
				if (nFolderNumber == 1)
					g_util.Logprintf("INFO: Newspilot file %s converted til XML", fi.sFileName);
				else 
					g_util.Logprintf("INFO: PPI file %s converted til XML", fi.sFileName);

				fi.sFileName = g_util.GetFileName(sOutputFile);
				sFullInputPath = sInputFolder + "\\" + fi.sFileName;
				::DeleteFile(sTempName);	

				// Converted ok .. proceed..
			}

			CString sSender = _T("");
			int nPageCount = 0;
			//SetCurrentInfoText(_T("Analyzing ") + fi.sFileName + _T(".."));

			::CopyFile(sFullInputPath, g_prefs.m_apppath + "\\CurrentFile.xml", FALSE);

			plan->Init();

			g_util.Logprintf("\r\nINFO: New file %s", fi.sFileName);

			CString sPress = "";

			BOOL bIgnorePostCommand = FALSE;
			BOOL bAppendProduction = FALSE;


			int ok = ParseXML(sFullInputPath, plan, nPageCount, sErrMsg, sInfo, sPress, bIgnorePostCommand, bAppendProduction);

			g_util.Logprintf("INFO: ParseXML() returned %d", ok);
			
			if (ok == PARSEFILE_ERROR || ok == PARSEFILE_UNKNOWNPRESS) {
				if (g_prefs.m_saveonerror)
					::MoveFileEx(sFullInputPath, nFolderNumber == 2 ? g_prefs.m_errorfolder2 + "\\" + fi.sFileName :
						g_prefs.m_errorfolder + "\\" + fi.sFileName, MOVEFILE_COPY_ALLOWED | MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH);
				else
					::DeleteFile(sFullInputPath);
				CString sError;
				sError.Format("%s  (%s)", sErrMsg, sInfo);
				AddLogItem("Error", fi.sFileName + " " + plan->m_planID, sPress, nPageCount, sError, sTimeStamp);
				g_util.Logprintf("ERROR: Could not validate file %s - %s", fi.sFileName, sError);
				
				g_TotalErrorImports++;
				
				if (g_prefs.m_senderrormail && nFolderNumber == 1 || g_prefs.m_senderrormail2 && nFolderNumber == 2)
					g_util.SendMail("Plan parsing error - " + fi.sFileName, "Plan parsing error " + sError,
						nFolderNumber == 1 ? g_prefs.m_emailreceivers : g_prefs.m_emailreceivers2);

				m_pollDB.UpdateService(SERVICESTATUS_HASERROR, fi.sFileName, "Plan XML parsing error", sErrorMessage);
				m_pollDB.InsertLogEntry(EVENTCODE_PLANERROR, _T("XML-plan"), fi.sFileName, _T("XML file parsing error ") + sErrMsg, 0, 0, 0, sInfo, sErrorMessage);

				continue;

			}


			int nExistingProductionID = 0;
			int nExistingPlanType = g_prefs.m_DB.GetPlanType(plan->m_publicationID, plan->m_pubdate, plan->m_pressID, nExistingProductionID, sErrorMessage);


			int nTotalPages = 0;
			for (int edition = 0; edition < plan->m_numberofeditions; edition++) {
				for (int section = 0; section < plan->m_numberofsections[0][edition]; section++) {
					nTotalPages += plan->m_numberofpages[0][edition][section];
				}
			}

			CUIntArray aChannelList;
			g_prefs.m_DB.GetPublicationChannels(plan->m_publicationID, aChannelList, sErrorMessage);
			if (aChannelList.GetCount() == 0) {
				if (g_prefs.m_saveonerror)
					::MoveFileEx(sFullInputPath, nFolderNumber == 2 ? g_prefs.m_errorfolder2 + "\\" + fi.sFileName :
						g_prefs.m_errorfolder + "\\" + fi.sFileName, MOVEFILE_COPY_ALLOWED | MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH);
				else
					::DeleteFile(sFullInputPath);
				CString sError;
				sError.Format("No channels defined for publication %s - plan file %s", g_prefs.GetPublicationName(plan->m_publicationID), fi.sFileName);
				AddLogItem("Error", fi.sFileName + " " + plan->m_planID, sPress, nPageCount, sError, sTimeStamp);
				g_util.Logprintf("ERROR: %s", sError);

				g_TotalErrorImports++;

				if (g_prefs.m_senderrormail && nFolderNumber == 1 || g_prefs.m_senderrormail2 && nFolderNumber == 2)
					g_util.SendMail("Plan parsing error - " + fi.sFileName, "Plan parsing error " + sError,
						nFolderNumber == 1 ? g_prefs.m_emailreceivers : g_prefs.m_emailreceivers2);

				m_pollDB.UpdateService(SERVICESTATUS_HASERROR, fi.sFileName, "Plan publication error", sErrorMessage);
				m_pollDB.InsertLogEntry(EVENTCODE_PLANERROR, _T("XML-plan"), fi.sFileName, _T("Plan publication error ") + sError, 0, 0, 0, sInfo, sErrorMessage);

				continue;
			}

			// Determine if all new section(s) for existing production.

			BOOL bAllNewSection = TRUE;

			if (nExistingProductionID > 0) {
				for (int edition = 0; edition < plan->m_numberofeditions; edition++) {
					for (int section = 0; section < plan->m_numberofsections[0][edition]; section++) {
						int nPageCountInSection = g_prefs.m_DB.GetPageCountInSection(plan->m_publicationID, plan->m_pubdate, plan->m_editionIDlist[0][edition], plan->m_sectionIDlist[0][edition][section], sErrorMessage);
						if (nPageCountInSection > 0) {
							bAllNewSection = FALSE;
							break;
						}
					}
				}
			}




			if (g_prefs.m_runningnumbers > 0) {

				for (int edition = 0; edition < plan->m_numberofeditions; edition++) {
					int nPag = 0;
					int nPageForward = 0;
					int nPageBackwards = nTotalPages;
					for (int section = 0; section < plan->m_numberofsections[0][edition]; section++) {

						int nMiddle = plan->m_numberofpages[0][edition][section] / 2;
						nPageBackwards -= nMiddle;
						for (int page = 0; page < plan->m_numberofpages[0][edition][section]; page++) {

							if (g_prefs.m_runningnumbers == 1)	// Continous numbering
								nPag++;

							else if (g_prefs.m_runningnumbers == 2) { // Inserted numbering

								if (page + 1 <= nMiddle)
									nPag = (page + 1) + nPageForward;
								else
									nPag = (page + 1) + nPageBackwards - nMiddle;

							}

							plan->m_pagepagination[0][edition][section][page] = (WORD)nPag;
						}
						nPageForward += nMiddle;
					}
				}
			}

			// REPAIR masterpageIDs etc (backward compatibility)
			BOOL bNoMasterIDs = FALSE;
			for (int edition = 0; edition < plan->m_numberofeditions; edition++) {
				for (int section = 0; section < plan->m_numberofsections[0][edition]; section++) {
					BOOL bAllCommon = TRUE;
					for (int page = 0; page < plan->m_numberofpages[0][edition][section]; page++) {

						if (*plan->m_pagemasterIDs[0][edition][section][page] == 0) {
							if (plan->m_pageunique[0][edition][section][page] == UNIQUEPAGE_UNIQUE || plan->m_numberofeditions == 1) {
								_tcscpy(plan->m_pagemasterIDs[0][edition][section][page], plan->m_pageIDs[0][edition][section][page]);
							}
							else {
								// common page - find master (same pagename and section in other edition)...
								BOOL bFound = FALSE;
								CString sPageNameToFind(plan->m_pagenames[0][edition][section][page]);
								WORD nSectionIDToFind = plan->m_sectionIDlist[0][edition][section];
								for (int medition = 0; medition < plan->m_numberofeditions; medition++) {
									// Only search other editions..
									if (plan->m_editionIDlist[0][edition] != plan->m_editionIDlist[0][medition]) {
										for (int msection = 0; msection < plan->m_numberofsections[0][medition]; msection++) {
											if (nSectionIDToFind == plan->m_sectionIDlist[0][medition][msection]) {
												for (int mpage = 0; mpage < plan->m_numberofpages[0][medition][msection]; mpage++) {
													CString sThisPageName(plan->m_pagenames[0][medition][msection][mpage]);
													if (sThisPageName == sPageNameToFind && plan->m_pageunique[0][medition][msection][mpage] == UNIQUEPAGE_UNIQUE) {
														_tcscpy(plan->m_pagemasterIDs[0][edition][section][page], plan->m_pageIDs[0][medition][msection][mpage]);
														bFound = TRUE;
														break;
													}
												}
											}
											if (bFound)
												break;
										}
									}
									if (bFound)
										break;
								}
							}
						}
					}
				}
			}

			if (bNoMasterIDs) {
				// Now repair sheets (backward compatibility)
				for (int edition = 0; edition < plan->m_numberofeditions; edition++) {
					for (int sheet = 0; sheet < plan->m_numberofsheets[0][edition]; sheet++) {
						for (int side = 0; side < 2; side++) {

							if (side == 1 && plan->m_sheetignoreback[0][edition][sheet])
								continue;

							for (int pos = 0; pos < plan->m_sheetpagesonplate[0][edition][sheet]; pos++) {
								CString sPageNameToFind(plan->m_oldsheetpagename[0][edition][sheet][side][pos]);
								int nSectionIDToFind = plan->m_oldsheetsectionID[0][edition][sheet][side][pos];
								if (*plan->m_sheetpageID[0][edition][sheet][side][pos] == 0 && sPageNameToFind != "" && nSectionIDToFind > 0) {
									// Find page in section/page array
									BOOL bFound = TRUE;
									for (int section = 0; section < plan->m_numberofsections[0][edition]; section++) {
										if (nSectionIDToFind == plan->m_sectionIDlist[0][edition][section]) {
											for (int page = 0; page < plan->m_numberofpages[0][edition][section]; page++) {
												CString sThisPageName(plan->m_pagenames[0][edition][section][page]);
												if (sThisPageName == sPageNameToFind) {
													_tcscpy(plan->m_sheetpageID[0][edition][sheet][side][pos], plan->m_pageIDs[0][edition][section][page]);
													_tcscpy(plan->m_sheetmasterpageID[0][edition][sheet][side][pos], plan->m_pagemasterIDs[0][edition][section][page]);
													bFound = TRUE;
													break;
												}
											}
										}
										if (bFound)
											break;
									}
								}
							}
						}
					}
				}
			}

			for (int edition = 0; edition < plan->m_numberofeditions; edition++) {

				for (int section = 0; section < plan->m_numberofsections[0][edition]; section++) {
					if (plan->m_editionistimed[0][edition] > 0 && plan->m_editiontimedtoID[0][edition] > 0) {
						plan->m_sectionallcommon[0][edition][section] = FALSE;
						continue;
					}
					plan->m_sectionallcommon[0][edition][section] = TRUE;
					for (int page = 0; page < plan->m_numberofpages[0][edition][section]; page++) {
						if (plan->m_pageunique[0][edition][section][page] == UNIQUEPAGE_UNIQUE || plan->m_numberofeditions == 1) {
							plan->m_sectionallcommon[0][edition][section] = FALSE;
							break;
						}
					}
				}
			}

			int bMinimumEditions = 1;
			for (int a = 0; a < g_prefs.m_ExtraEditionList.GetCount(); a++) {
				int nID = g_prefs.GetPublicationID(g_prefs.m_ExtraEditionList[a].m_Publication);
				if (nID == plan->m_publicationID) {
					bMinimumEditions = g_prefs.m_ExtraEditionList[a].m_mineditions;
					g_util.Logprintf("INFO: Publication %s must have %d extra editions", g_prefs.m_ExtraEditionList[a].m_Publication, bMinimumEditions);
					break;
				}
			}

			// Count editions names of format X_Y where Y is [1..9]
			CStringArray sZones;
			int nOriginalNumberOfEditions = plan->m_numberofeditions;
			for (int edition = 0; edition < plan->m_numberofeditions; edition++) {
				CString sEd = g_prefs.GetEditionName(plan->m_editionIDlist[0][edition]);
				int q = sEd.Find("_");
				if (q == -1)
					continue;
				if (atoi(sEd.Mid(q + 1)) > 0)
					sZones.Add(sEd.Left(q)); // First character of edition name X_Y
			}

			if (sZones.GetCount() > 0 && bMinimumEditions > 1) {
				int issue = 0;

				for (int zone = 0; zone < nOriginalNumberOfEditions; zone++) {
					CString sMasterEd = g_prefs.GetEditionName(plan->m_editionIDlist[0][zone]);
					int q = sMasterEd.Find("_");

					if (q == -1)
						continue;

					// Only copy from X_1 edition (zone)
					if (atoi(sMasterEd.Mid(q + 1)) != 1)
						continue;

					// Already a timed edition in plan?
					if (atoi(sMasterEd.Mid(q + 1)) >= bMinimumEditions || plan->m_editionistimed[0][zone])
						continue;

					int nSecondTimesEditionID = 0;
					for (int timeded = 2; timeded <= bMinimumEditions; timeded++) {
						CString sEd = g_util.Int2String(timeded);

						int nPrevSecondTimesEditionID = nSecondTimesEditionID;
						nSecondTimesEditionID = g_prefs.GetEditionID(sMasterEd.Left(q + 1) + sEd);

						if (nSecondTimesEditionID <= 0)
							continue;

						int edition = plan->m_numberofeditions++;

						plan->m_pressIDlist[0][edition] = plan->m_pressIDlist[0][zone];
						plan->m_allnonuniqueepages[0][edition] = TRUE;
						plan->m_numberofpapercopiesedition[0][edition] = plan->m_numberofpapercopiesedition[0][zone];
						plan->m_editionnochange[0][edition] = FALSE;

						plan->m_editionIDlist[0][edition] = nSecondTimesEditionID;
						plan->m_editiontimedfromID[0][edition] = timeded == 2 ? plan->m_editionIDlist[0][zone] : nPrevSecondTimesEditionID;
						plan->m_editionistimed[0][edition] = TRUE;
						if (timeded == 2)
							plan->m_editiontimedtoID[0][zone] = nSecondTimesEditionID;
						else { // 3..4..5..
							// Point to next timed..
							sEd = g_util.Int2String(timeded + 1);
							plan->m_editiontimedtoID[0][edition] = timeded < bMinimumEditions ? g_prefs.GetEditionID(sMasterEd.Left(q + 1) + sEd) : 0;
						}
						plan->m_editionmasterzoneeditionID[0][edition] = plan->m_editionIDlist[0][zone];
						_tcscpy(plan->m_editioncomment[0][edition], plan->m_editioncomment[0][zone]);
						plan->m_editionordernumber[0][edition] = plan->m_editionordernumber[0][zone] + 1;

						plan->m_numberofpresses[0][edition] = plan->m_numberofpresses[0][zone];
						for (int k = 0; k < plan->m_numberofpresses[0][edition]; k++) {
							plan->m_pressesinedition[0][edition][k] = plan->m_pressesinedition[0][zone][k];
							plan->m_numberofpapercopies[0][edition][k] = plan->m_numberofpapercopies[0][zone][k];
							plan->m_numberofcopies[0][edition][k] = plan->m_numberofcopies[0][zone][k];
							plan->m_autorelease[0][edition][k] = plan->m_autorelease[0][zone][k];
							_tcscpy(plan->m_postalurl[0][edition][k], plan->m_postalurl[0][zone][k]);
							_tcscpy(plan->m_presstime[0][edition][k], plan->m_presstime[0][zone][k]);
						}

						plan->m_numberofsections[0][edition] = plan->m_numberofsections[0][zone];
						for (int section = 0; section < plan->m_numberofsections[0][edition]; section++) {
							plan->m_sectionIDlist[0][edition][section] = plan->m_sectionIDlist[0][zone][section];
							plan->m_sectionallcommon[0][edition][section] = TRUE;
							plan->m_numberofpages[0][edition][section] = plan->m_numberofpages[0][zone][section];
							for (int page = 0; page < plan->m_numberofpages[0][edition][section]; page++) {

								plan->m_numberofpagecolors[0][edition][section][page] = plan->m_numberofpagecolors[0][zone][section][page];

								_tcscpy(plan->m_pagenames[0][edition][section][page], plan->m_pagenames[0][zone][section][page]);
								plan->m_pageindex[0][edition][section][page] = plan->m_pageindex[0][zone][section][page];
								plan->m_pagepagination[0][edition][section][page] = plan->m_pagepagination[0][zone][section][page];
								plan->m_pageunique[0][edition][section][page] = UNIQUEPAGE_COMMON;
								plan->m_pageapproved[0][edition][section][page] = plan->m_pageapproved[0][zone][section][page];
								_tcscpy(plan->m_pagecomments[0][edition][section][page], plan->m_pagecomments[0][zone][section][page]);
								_tcscpy(plan->m_pagefilenames[0][edition][section][page], plan->m_pagefilenames[0][zone][section][page]);
								plan->m_pageversion[0][edition][section][page] = plan->m_pageversion[0][zone][section][page];
								plan->m_pagehold[0][edition][section][page] = plan->m_pageversion[0][zone][section][page];
								plan->m_pagemiscint[0][edition][section][page] = plan->m_pagemiscint[0][zone][section][page];
								_tcscpy(plan->m_pagemiscstring1[0][edition][section][page], plan->m_pagemiscstring1[0][zone][section][page]);
								_tcscpy(plan->m_pagemiscstring2[0][edition][section][page], plan->m_pagemiscstring2[0][zone][section][page]);
								plan->m_pageposition[0][edition][section][page] = plan->m_pageposition[0][zone][section][page];
								plan->m_pagepriority[0][edition][section][page] = plan->m_pagepriority[0][zone][section][page];
								plan->m_pagetypes[0][edition][section][page] = plan->m_pagetypes[0][zone][section][page];
								_tcscpy(plan->m_pagemasterIDs[0][edition][section][page], plan->m_pagemasterIDs[0][zone][section][page]);
								sprintf(plan->m_pageIDs[0][edition][section][page], "%d", atoi(plan->m_pageIDs[0][zone][section][page]) + 1000 * (timeded - 1));
								for (int n = 0; n < plan->m_numberofpagecolors[0][edition][section][page]; n++) {
									plan->m_pagecolorIDlist[0][edition][section][page][n] = plan->m_pagecolorIDlist[0][zone][section][page][n];
									plan->m_pagecolorActivelist[0][edition][section][page][n] = plan->m_pagecolorActivelist[0][zone][section][page][n];
								}
							}
						}

						plan->m_numberofsheets[0][edition] = plan->m_numberofsheets[0][zone];

						for (int k = 0; k < plan->m_numberofsheets[0][edition]; k++) {
							plan->m_sheettemplateID[0][edition][k] = plan->m_sheettemplateID[0][zone][k];
							plan->m_sheetbacktemplateID[0][edition][k] = plan->m_sheetbacktemplateID[0][zone][k];
							_tcscpy(plan->m_sheetmarkgroup[0][edition][k], plan->m_sheetmarkgroup[0][zone][k]);
							plan->m_sheetpagesonplate[0][edition][k] = plan->m_sheetpagesonplate[0][zone][k];
							plan->m_sheetignoreback[0][edition][k] = plan->m_sheetignoreback[0][zone][k];
							plan->m_pressSectionNumber[0][edition][k] = plan->m_pressSectionNumber[0][zone][k];

							for (int m = 0; m < 2; m++) {
								plan->m_sheetdefaultsectionID[0][edition][k][m] = plan->m_sheetdefaultsectionID[0][zone][k][m];
								plan->m_sheetmaxcolors[0][edition][k][m] = plan->m_sheetmaxcolors[0][zone][k][m];
								_tcscpy(plan->m_sheetsortingposition[0][edition][k][m], plan->m_sheetsortingposition[0][zone][k][m]);
								plan->m_sheetactivecopies[0][edition][k][m] = plan->m_sheetactivecopies[0][zone][k][m];

								_tcscpy(plan->m_sheetpresstower[0][edition][k][m], plan->m_sheetpresstower[0][zone][k][m]);
								_tcscpy(plan->m_sheetpresszone[0][edition][k][m], plan->m_sheetpresszone[0][zone][k][m]);
								_tcscpy(plan->m_sheetpresshighlow[0][edition][k][m], plan->m_sheetpresshighlow[0][zone][k][m]);

								plan->m_sheetallcommonpages[0][edition][k][m] = TRUE;

								for (int n = 0; n < MAXPAGEPOSITIONS; n++) {
									plan->m_sheetpageposition[0][edition][k][m][n] = plan->m_sheetpageposition[0][zone][k][m][n];
									sprintf(plan->m_sheetpageID[0][edition][k][m][n], "%d", atoi(plan->m_sheetpageID[0][zone][k][m][n]) + 1000 * (timeded - 1));
									_tcscpy(plan->m_sheetmasterpageID[0][edition][k][m][n], plan->m_sheetmasterpageID[0][zone][k][m][n]);

									_tcscpy(plan->m_oldsheetpagename[0][edition][k][m][n], plan->m_oldsheetpagename[0][zone][k][m][n]);
									plan->m_oldsheetsectionID[0][edition][k][m][n] = plan->m_oldsheetsectionID[0][zone][k][m][n];
								}

								for (int n = 0; n < MAXCOLORS; n++) {
									_tcscpy(plan->m_sheetpresscylinder[0][edition][k][m][n], plan->m_sheetpresscylinder[0][zone][k][m][n]);
									_tcscpy(plan->m_sheetformID[0][edition][k][m][n], plan->m_sheetformID[0][zone][k][m][n]);
									_tcscpy(plan->m_sheetplateID[0][edition][k][m][n], plan->m_sheetplateID[0][zone][k][m][n]);
									_tcscpy(plan->m_sheetsortingpositioncolorspecific[0][edition][k][m][n], plan->m_sheetsortingpositioncolorspecific[0][zone][k][m][n]);
								}
							}
						}
					}
				}
			}


			BOOL bAddExtraZone = FALSE;
			CString sZoneToAdd = _T("");

			for (int a = 0; a < g_prefs.m_ExtraZoneList.GetCount(); a++) {
				int nID = g_prefs.GetPublicationID(g_prefs.m_ExtraZoneList[a].m_Publication);
				if (nID == plan->m_publicationID && g_prefs.m_ExtraZoneList[a].m_Press == sPress) {
					bAddExtraZone = TRUE;
					sZoneToAdd = g_prefs.m_ExtraZoneList[a].m_zonetoadd;
					break;
				}
			}

			if (bAddExtraZone && sZoneToAdd != "") {
				// Zone already added?
				for (int zone = 0; zone < sZones.GetCount(); zone++) {
					if (sZoneToAdd == sZones[zone]) {
						bAddExtraZone = FALSE;
						break;
					}
				}
			}

			int nSecondZoneEditionID = g_prefs.GetEditionID(sZoneToAdd + "_1");

			if (bAddExtraZone && sZoneToAdd != "" && nSecondZoneEditionID > 0) {

				int masterzone = 0;
				int edition = plan->m_numberofeditions++;

				plan->m_pressIDlist[0][edition] = plan->m_pressIDlist[0][masterzone];
				plan->m_allnonuniqueepages[0][edition] = TRUE;
				plan->m_numberofpapercopiesedition[0][edition] = plan->m_numberofpapercopiesedition[0][masterzone];
				plan->m_editionnochange[0][edition] = FALSE;

				plan->m_editionIDlist[0][edition] = nSecondZoneEditionID;
				plan->m_editiontimedfromID[0][edition] = 0;
				plan->m_editiontimedtoID[0][edition] = 0;
				plan->m_editionistimed[0][edition] = FALSE;
				plan->m_editionmasterzoneeditionID[0][edition] = plan->m_editionIDlist[0][masterzone];
				_tcscpy(plan->m_editioncomment[0][edition], plan->m_editioncomment[0][masterzone]);
				plan->m_editionordernumber[0][edition] = plan->m_editionordernumber[0][masterzone] + 1;

				plan->m_numberofpresses[0][edition] = plan->m_numberofpresses[0][masterzone];
				for (int k = 0; k < plan->m_numberofpresses[0][edition]; k++) {
					plan->m_pressesinedition[0][edition][k] = plan->m_pressesinedition[0][masterzone][k];
					plan->m_numberofpapercopies[0][edition][k] = plan->m_numberofpapercopies[0][masterzone][k];
					plan->m_numberofcopies[0][edition][k] = plan->m_numberofcopies[0][masterzone][k];
					plan->m_autorelease[0][edition][k] = plan->m_autorelease[0][masterzone][k];
					_tcscpy(plan->m_postalurl[0][edition][k], plan->m_postalurl[0][masterzone][k]);
					_tcscpy(plan->m_presstime[0][edition][k], plan->m_presstime[0][masterzone][k]);
				}

				plan->m_numberofsections[0][edition] = plan->m_numberofsections[0][masterzone];
				for (int section = 0; section < plan->m_numberofsections[0][edition]; section++) {
					plan->m_sectionIDlist[0][edition][section] = plan->m_sectionIDlist[0][masterzone][section];
					plan->m_sectionallcommon[0][edition][section] = TRUE;
					plan->m_numberofpages[0][edition][section] = plan->m_numberofpages[0][masterzone][section];
					for (int page = 0; page < plan->m_numberofpages[0][edition][section]; page++) {

						plan->m_numberofpagecolors[0][edition][section][page] = plan->m_numberofpagecolors[0][masterzone][section][page];

						_tcscpy(plan->m_pagenames[0][edition][section][page], plan->m_pagenames[0][masterzone][section][page]);
						plan->m_pageindex[0][edition][section][page] = plan->m_pageindex[0][masterzone][section][page];
						plan->m_pagepagination[0][edition][section][page] = plan->m_pagepagination[0][masterzone][section][page];
						plan->m_pageunique[0][edition][section][page] = UNIQUEPAGE_COMMON;
						plan->m_pageapproved[0][edition][section][page] = plan->m_pageapproved[0][masterzone][section][page];
						_tcscpy(plan->m_pagecomments[0][edition][section][page], plan->m_pagecomments[0][masterzone][section][page]);
						_tcscpy(plan->m_pagefilenames[0][edition][section][page], plan->m_pagefilenames[0][masterzone][section][page]);
						plan->m_pageversion[0][edition][section][page] = plan->m_pageversion[0][masterzone][section][page];
						plan->m_pagehold[0][edition][section][page] = plan->m_pageversion[0][masterzone][section][page];
						plan->m_pagemiscint[0][edition][section][page] = plan->m_pagemiscint[0][masterzone][section][page];
						_tcscpy(plan->m_pagemiscstring1[0][edition][section][page], plan->m_pagemiscstring1[0][masterzone][section][page]);
						_tcscpy(plan->m_pagemiscstring2[0][edition][section][page], plan->m_pagemiscstring2[0][masterzone][section][page]);
						plan->m_pageposition[0][edition][section][page] = plan->m_pageposition[0][masterzone][section][page];
						plan->m_pagepriority[0][edition][section][page] = plan->m_pagepriority[0][masterzone][section][page];
						plan->m_pagetypes[0][edition][section][page] = plan->m_pagetypes[0][masterzone][section][page];
						_tcscpy(plan->m_pagemasterIDs[0][edition][section][page], plan->m_pagemasterIDs[0][masterzone][section][page]);
						sprintf(plan->m_pageIDs[0][edition][section][page], "%d", atoi(plan->m_pageIDs[0][masterzone][section][page]) + 1000);
						for (int n = 0; n < plan->m_numberofpagecolors[0][edition][section][page]; n++) {
							plan->m_pagecolorIDlist[0][edition][section][page][n] = plan->m_pagecolorIDlist[0][masterzone][section][page][n];
							plan->m_pagecolorActivelist[0][edition][section][page][n] = plan->m_pagecolorActivelist[0][masterzone][section][page][n];
						}
					}
				}

				plan->m_numberofsheets[0][edition] = plan->m_numberofsheets[0][masterzone];

				for (int k = 0; k < plan->m_numberofsheets[0][edition]; k++) {
					plan->m_sheettemplateID[0][edition][k] = plan->m_sheettemplateID[0][masterzone][k];
					plan->m_sheetbacktemplateID[0][edition][k] = plan->m_sheetbacktemplateID[0][masterzone][k];
					_tcscpy(plan->m_sheetmarkgroup[0][edition][k], plan->m_sheetmarkgroup[0][masterzone][k]);
					plan->m_sheetpagesonplate[0][edition][k] = plan->m_sheetpagesonplate[0][masterzone][k];
					plan->m_sheetignoreback[0][edition][k] = plan->m_sheetignoreback[0][masterzone][k];
					plan->m_pressSectionNumber[0][edition][k] = plan->m_pressSectionNumber[0][masterzone][k];

					for (int m = 0; m < 2; m++) {
						plan->m_sheetdefaultsectionID[0][edition][k][m] = plan->m_sheetdefaultsectionID[0][masterzone][k][m];
						plan->m_sheetmaxcolors[0][edition][k][m] = plan->m_sheetmaxcolors[0][masterzone][k][m];
						_tcscpy(plan->m_sheetsortingposition[0][edition][k][m], plan->m_sheetsortingposition[0][masterzone][k][m]);
						plan->m_sheetactivecopies[0][edition][k][m] = plan->m_sheetactivecopies[0][masterzone][k][m];

						_tcscpy(plan->m_sheetpresstower[0][edition][k][m], plan->m_sheetpresstower[0][masterzone][k][m]);
						_tcscpy(plan->m_sheetpresszone[0][edition][k][m], plan->m_sheetpresszone[0][masterzone][k][m]);
						_tcscpy(plan->m_sheetpresshighlow[0][edition][k][m], plan->m_sheetpresshighlow[0][masterzone][k][m]);

						plan->m_sheetallcommonpages[0][edition][k][m] = TRUE;

						for (int n = 0; n < MAXPAGEPOSITIONS; n++) {
							plan->m_sheetpageposition[0][edition][k][m][n] = plan->m_sheetpageposition[0][masterzone][k][m][n];
							sprintf(plan->m_sheetpageID[0][edition][k][m][n], "%d", atoi(plan->m_sheetpageID[0][masterzone][k][m][n] + 1000));
							_tcscpy(plan->m_sheetmasterpageID[0][edition][k][m][n], plan->m_sheetmasterpageID[0][masterzone][k][m][n]);

							_tcscpy(plan->m_oldsheetpagename[0][edition][k][m][n], plan->m_oldsheetpagename[0][masterzone][k][m][n]);
							plan->m_oldsheetsectionID[0][edition][k][m][n] = plan->m_oldsheetsectionID[0][masterzone][k][m][n];
						}

						for (int n = 0; n < MAXCOLORS; n++) {
							_tcscpy(plan->m_sheetpresscylinder[0][edition][k][m][n], plan->m_sheetpresscylinder[0][masterzone][k][m][n]);
							_tcscpy(plan->m_sheetformID[0][edition][k][m][n], plan->m_sheetformID[0][masterzone][k][m][n]);
							_tcscpy(plan->m_sheetplateID[0][edition][k][m][n], plan->m_sheetplateID[0][masterzone][k][m][n]);
							_tcscpy(plan->m_sheetsortingpositioncolorspecific[0][edition][k][m][n], plan->m_sheetsortingpositioncolorspecific[0][masterzone][k][m][n]);
						}
					}
				}
			}

			if (g_prefs.m_skipcommonplates) {
				for (int edition = 0; edition < plan->m_numberofeditions; edition++) {
					for (int sheet = 0; sheet < plan->m_numberofsheets[0][edition]; sheet++) {
						for (int side = 0; side < 2; side++) {
							if (plan->m_editionistimed[0][edition] > 0 && plan->m_editiontimedfromID[0][edition] > 0) {
								plan->m_sheetallcommonpages[0][edition][sheet][side] = FALSE;
								continue;
							}
							plan->m_sheetallcommonpages[0][edition][sheet][side] = TRUE;
							for (int pos = 0; pos < plan->m_sheetpagesonplate[0][edition][sheet]; pos++) {
								CString sPageID(plan->m_sheetpageID[0][edition][sheet][side][pos]);
								CString sPageDummyTest = sPageID;
								sPageDummyTest.MakeUpper();

								if (sPageDummyTest.Find("DUMMY") != -1 || sPageDummyTest.Find("DINKY") != -1 || sPageDummyTest.Find("DINKEY") != -1 || sPageDummyTest == "-1")
									continue;

								// Find this PageID in the pages structure
								BOOL bFound = TRUE;

								for (int section = 0; section < plan->m_numberofsections[0][edition]; section++) {
									for (int page = 0; page < plan->m_numberofpages[0][edition][section]; page++) {
										CString sThisPageID(plan->m_pageIDs[0][edition][section][page]);

										if (sThisPageID.CompareNoCase(sPageID) == 0) {

											if (plan->m_pageunique[0][edition][section][page] > 0)
												plan->m_sheetallcommonpages[0][edition][sheet][side] = FALSE;
											bFound = TRUE;
											break;
										}
									}

									if (bFound)
										break;
								}
								if (plan->m_sheetallcommonpages[0][edition][sheet][side] == FALSE)
									break;
							}
						}
					}
				}
			}
			// From here we need exclusive pagetable insert/delete lock

			TCHAR szClientName[100];
			TCHAR szClientTime[50];
			int nWaitLoop = 0;
			CString sInfo;

			int nCurrentPlanLock = PLANLOCK_UNKNOWN;

			sSender = plan->m_sender;

			if (g_prefs.m_planlocksystem) {
				if (g_prefs.m_debug > 1)
					g_util.Logprintf("Requesting plan lock..");
				while (nCurrentPlanLock != PLANLOCK_LOCK && g_BreakSignal == FALSE) {
					if (m_pollDB.PlanLock(PLANLOCK_LOCK, &nCurrentPlanLock, szClientName, szClientTime, sErrorMessage) == FALSE) {
						sErr.Format(_T("Database error - %s"), sErrorMessage);
						g_util.Logprintf("%s", sErr);
						bReConnect = TRUE;
						m_pollDB.UpdateService(SERVICESTATUS_HASERROR, fi.sFileName, _T("Database error"), sErrorMessage);
						m_pollDB.InsertLogEntry(EVENTCODE_PLANERROR, _T("Database error"), fi.sFileName, sErr, 0, 0, 0, _T(""), sErrorMessage);
						break;
					}
					if (nCurrentPlanLock == PLANLOCK_NOLOCK) {
						if (nWaitLoop % 100 == 0 && g_prefs.m_debug > 1)
							g_util.Logprintf("Unable to accuire plan lock. Owner: %s at %s", szClientName, szClientTime);

						nWaitLoop++;
						if (nWaitLoop % 2 == 0)
							sInfo.Format("Waiting for plan lock to free (%s at %s)..", szClientName, szClientTime);
						else
							sInfo.Format(" ", szClientName, szClientTime);
						//SetCurrentInfoText(sInfo);
						::Sleep(3000);
					}
					else if (nCurrentPlanLock == PLANLOCK_LOCK) {
						sInfo.Format("Got plan lock - proceeding..");
						break;
					}

				}

				if (g_prefs.m_planlocksystem > 1 && nCurrentPlanLock == PLANLOCK_LOCK)
					m_pollDB.DeleteAllDirtyPages(plan->m_productionID, sErrorMessage);

				if (g_BreakSignal)
					break;
			}


			TCHAR szPublicationPlanClientName[100];
			TCHAR szPublicationPlanClientTime[50];
			int nCurrentPublicationPlanLock = PLANLOCK_UNKNOWN;

			nWaitLoop = 0;

			if (plan->m_publicationID > 0 && g_prefs.m_publicationlock ) {
				if (g_prefs.m_debug > 1)
					g_util.Logprintf("Requesting publication plan lock..");
			
				while (nCurrentPublicationPlanLock != PLANLOCK_LOCK && g_BreakSignal == FALSE) {
					if (m_pollDB.PublicationPlanLock(plan->m_publicationID, plan->m_pubdate, PLANLOCK_LOCK, &nCurrentPublicationPlanLock, szPublicationPlanClientName, szPublicationPlanClientTime, sErrorMessage) == FALSE) {
						bReConnect = TRUE;
						sErr = sErrorMessage;
						m_pollDB.UpdateService(SERVICESTATUS_HASERROR, fi.sFileName, _T("Database error"), sErrorMessage);
						m_pollDB.InsertLogEntry(EVENTCODE_PLANERROR, _T("Database error"), fi.sFileName, sErr, 0, 0, 0, _T(""), sErrorMessage);
						break;
					}
					if (nCurrentPublicationPlanLock == PLANLOCK_NOLOCK) {
						if (nWaitLoop % 100 == 0 && g_prefs.m_debug > 1)
							g_util.Logprintf("Unable to accuire publication plan lock. Owner: %s at %s", szPublicationPlanClientName, szPublicationPlanClientTime);
						nWaitLoop++;
						if (nWaitLoop % 2 == 0)
							sInfo.Format("Waiting for publication plan lock to free (%s at %s)..", szPublicationPlanClientName, szPublicationPlanClientTime);
						else
							sInfo.Format(" ", szPublicationPlanClientName, szPublicationPlanClientTime);
						//SetCurrentInfoText(sInfo);
						Sleep(3000);
					}
					else if (nCurrentPublicationPlanLock == PLANLOCK_LOCK) {
						sInfo.Format("Got publication plan lock - proceeding..");
						break;
					}

				}


				if (g_BreakSignal)
					break;
			}

			// Handle plan kill request

			if (ok == PARSEFILE_ABORT) {
				if (!m_pollDB.DeleteProductionPagesEx(plan->m_planName, sErrorMessage)) {
					sErr.Format(_T("ERROR: DeleteProductionPagesEx() - %s"), sErrorMessage);

					m_pollDB.UpdateService(SERVICESTATUS_HASERROR, fi.sFileName, _T("Database error"), sErrorMessage);
					m_pollDB.InsertLogEntry(EVENTCODE_PLANERROR, _T("Database error"), fi.sFileName, sErr, 0, 0, 0, _T(""), sErrorMessage);
					bReConnect = TRUE;
					//					break;
				}

				if (g_prefs.m_planlocksystem > 1 && nCurrentPlanLock == PLANLOCK_LOCK)
					m_pollDB.DeleteAllDirtyPages(plan->m_productionID, sErrorMessage);

				if (g_prefs.m_planlocksystem) {
					if (m_pollDB.PlanLock(PLANLOCK_UNLOCK, &nCurrentPlanLock, szClientName, szClientTime, sErrorMessage) == FALSE) {
						sErr.Format("ERROR: PlanLock() - %s", sErrorMessage);
						g_util.Logprintf("%s", sErr);
						m_pollDB.UpdateService(SERVICESTATUS_HASERROR, fi.sFileName, _T("Database error"), sErrorMessage);
						m_pollDB.InsertLogEntry(EVENTCODE_PLANERROR, _T("Database error"), fi.sFileName, sErr, 0, 0, 0, _T(""), sErrorMessage);
						bReConnect = TRUE;
					}
				}
				if (nCurrentPublicationPlanLock == PLANLOCK_LOCK)
					m_pollDB.PublicationPlanLock(plan->m_publicationID, plan->m_pubdate, PLANLOCK_UNLOCK, &nCurrentPublicationPlanLock, szPublicationPlanClientName, szPublicationPlanClientTime, sErrorMessage);


				s = fi.sFileName + " - plan deleted from system";
				AddLogItem("Deleted", fi.sFileName + " " + plan->m_planID, sPress, 0, s, sTimeStamp);
				::DeleteFile(sFullInputPath);


				continue;
			}

			BOOL bIsUpdate = FALSE;

			// Insert pages into database

		
			ok = ImportXMLpages(&m_pollDB, plan, bIsUpdate, sErrMsg, sInfo, (g_prefs.m_keepexistingsections || bAppendProduction));
				g_util.Logprintf("ImportXMLpages() returned %d", ok);			

			m_pollDB.ResetAllPollLocks(sErrorMessage);

			if (g_prefs.m_debug)
				LogXMLcontent(sFullInputPath, plan);

			// Report if empty import file (no database action)

			if (ok == IMPORTPAGES_NOEDITION) {
				if (g_prefs.m_planlocksystem > 1 && nCurrentPlanLock == PLANLOCK_LOCK)
					m_pollDB.DeleteAllDirtyPages(plan->m_productionID, sErrorMessage);
				if (g_prefs.m_planlocksystem) {
					if (m_pollDB.PlanLock(PLANLOCK_UNLOCK, &nCurrentPlanLock, szClientName, szClientTime, sErrorMessage) == FALSE) {
						sErr.Format("ERROR: PlanLock() - %s", sErrorMessage);
						bReConnect = TRUE;
						m_pollDB.UpdateService(SERVICESTATUS_HASERROR, fi.sFileName, _T("Database error"), sErrorMessage);
						m_pollDB.InsertLogEntry(EVENTCODE_PLANERROR, _T("Database error"), fi.sFileName, sErr, 0, 0, 0, _T(""), sErrorMessage);
					}
				}
				if (nCurrentPublicationPlanLock == PLANLOCK_LOCK)
					m_pollDB.PublicationPlanLock(plan->m_publicationID, plan->m_pubdate, PLANLOCK_UNLOCK, &nCurrentPublicationPlanLock, szPublicationPlanClientName, szPublicationPlanClientTime, sErrorMessage);
				s = fi.sFileName + _T("- No main edition found");
				g_util.Logprintf("%s", s);

				if (g_prefs.m_senderrormail && nFolderNumber == 1 || g_prefs.m_senderrormail2 && nFolderNumber == 2)
					g_util.SendMail("Plan create error - " + fi.sFileName, "Plan create error " + s,
						nFolderNumber == 1 ? g_prefs.m_emailreceivers : g_prefs.m_emailreceivers2);

				if (g_prefs.m_debug)
					g_util.Logprintf("Plan re-jected because no active editions in import file");

	
				continue;
			}

			// Report if identical import file (no database action)

			if (ok == IMPORTPAGES_REJECTEDREIMPORT || ok == IMPORTPAGES_REJECTEDAPPLIEDREIMPORT || ok == IMPORTPAGES_REJECTEDPOLLEDREIMPORT || ok == IMPORTPAGES_REJECTEDIMAGEDREIMPORT || ok == IMPORTPAGES_REJECTEDALLREIMPORT) {
				if (g_prefs.m_planlocksystem > 1 && nCurrentPlanLock == PLANLOCK_LOCK)
					m_pollDB.DeleteAllDirtyPages(plan->m_productionID, sErrorMessage);
				if (g_prefs.m_planlocksystem) {
					if (m_pollDB.PlanLock(PLANLOCK_UNLOCK, &nCurrentPlanLock, szClientName, szClientTime, sErrorMessage) == FALSE) {
						sErr.Format("ERROR: PlanLock() - %s", sErrorMessage);
						g_util.Logprintf("%s", sErr);
						bReConnect = TRUE;
					}
				}
				if (nCurrentPublicationPlanLock == PLANLOCK_LOCK)
					m_pollDB.PublicationPlanLock(plan->m_publicationID, plan->m_pubdate, PLANLOCK_UNLOCK, &nCurrentPublicationPlanLock, szPublicationPlanClientName, szPublicationPlanClientTime, sErrorMessage);
				//s.LoadString(IDS_POLL_REIMORT);	
				if (ok == IMPORTPAGES_REJECTEDREIMPORT)
					s = fi.sFileName + " - re-import rejected - same page count";
				else if (ok == IMPORTPAGES_REJECTEDAPPLIEDREIMPORT)
					s = fi.sFileName + " - re-import rejected - press plan already applied";
				else if (ok == IMPORTPAGES_REJECTEDPOLLEDREIMPORT)
					s = fi.sFileName + " - re-import rejected - pages already arrived";
				else if (ok == IMPORTPAGES_REJECTEDIMAGEDREIMPORT)
					s = fi.sFileName + " - re-import rejected - plates already imaged";
				else if (ok == IMPORTPAGES_REJECTEDALLREIMPORT)
					s = fi.sFileName + " - re-import rejected - plan already imported to specific press";

				AddLogItem("Warning", fi.sFileName + " " + plan->m_planID, sPress, nPageCount, s, sTimeStamp);
				::DeleteFile(sFullInputPath);

				if (g_prefs.m_debug)
					g_util.Logprintf("%s", s);

				m_pollDB.UpdateService(SERVICESTATUS_HASERROR, fi.sFileName, _T("Plan error"), sErrorMessage);
				m_pollDB.InsertLogEntry(EVENTCODE_PLANERROR, _T("Plan error"), fi.sFileName, s, 0, 0, 0, _T(""), sErrorMessage);

				m_pollDB.InsertPlanLogEntry(g_prefs.m_instancenumber, EVENTCODE_PLANERROR, fi.sFileName, s, 0, sSender, sErrorMessage);


				continue;
			}

			// Report if db-error occured during import
			CString sErrorMessage2;
			if (ok == IMPORTPAGES_DBERROR) {
				sErr.Format("Database error - %s", sErrorMessage);
				g_util.Logprintf("%s", sErr);
				
				if (g_prefs.m_planlocksystem > 1 && nCurrentPlanLock == PLANLOCK_LOCK)
					m_pollDB.DeleteAllDirtyPages(plan->m_productionID, sErrorMessage);
				if (g_prefs.m_planlocksystem) {
					
					if (m_pollDB.PlanLock(PLANLOCK_UNLOCK, &nCurrentPlanLock, szClientName, szClientTime, sErrorMessage2) == FALSE) {
						sErr.Format("ERROR: PlanLock() -", sErrorMessage2);
						bReConnect = TRUE;
						m_pollDB.UpdateService(SERVICESTATUS_HASERROR, fi.sFileName, _T("Database error"), sErrorMessage);
						m_pollDB.InsertLogEntry(EVENTCODE_PLANERROR, _T("Database error"), fi.sFileName, sErr, 0, 0, 0, _T(""), sErrorMessage);
					}
				}
				if (nCurrentPublicationPlanLock == PLANLOCK_LOCK)
					m_pollDB.PublicationPlanLock(plan->m_publicationID, plan->m_pubdate, PLANLOCK_UNLOCK, &nCurrentPublicationPlanLock, szPublicationPlanClientName, szPublicationPlanClientTime, sErrorMessage2);

				bReConnect = TRUE;

				if (g_prefs.m_senderrormail && nFolderNumber == 1 || g_prefs.m_senderrormail2 && nFolderNumber == 2)
					g_util.SendMail("Plan create error - " + fi.sFileName, "Plan create error " + sErr,
						nFolderNumber == 1 ? g_prefs.m_emailreceivers : g_prefs.m_emailreceivers2);

				if (g_prefs.m_debug)
					g_util.Logprintf("Plan rejected due to database problem - %s", sErrorMessage);

				if (g_prefs.m_saveonerror)
					::MoveFileEx(sFullInputPath, sErrorPath, MOVEFILE_COPY_ALLOWED | MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH);
				else
					::DeleteFile(sFullInputPath);

				m_pollDB.InsertPlanLogEntry(g_prefs.m_instancenumber, EVENTCODE_PLANERROR, fi.sFileName, "Plan rejected due to database problem", 0, sSender, sErrorMessage);

				break;
			}

			if (ok == IMPORTPAGES_ERROR) {
				if (g_prefs.m_saveonerror)
					::MoveFileEx(sFullInputPath, sErrorPath, MOVEFILE_COPY_ALLOWED | MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH);
				else
					::DeleteFile(sFullInputPath);
				CString sError;
				sError.Format("%s  (%s)", sErrMsg, sInfo);
				AddLogItem("Error", fi.sFileName + " " + plan->m_planID, sPress, nPageCount, sError, sTimeStamp);
				g_util.Logprintf("ERROR: Could not import file %s although XML validation succeeded", fi.sFileName);
				g_TotalErrorImports++;

				if (g_prefs.m_senderrormail && nFolderNumber == 1 || g_prefs.m_senderrormail2 && nFolderNumber == 2)
					g_util.SendMail("Plan create error - " + fi.sFileName, "Plan create error " + sError,
						nFolderNumber == 1 ? g_prefs.m_emailreceivers : g_prefs.m_emailreceivers2);

				if (g_prefs.m_debug)
					g_util.Logprintf("Plan rejected due to inconsistency in plan data - %s - %s", sInfo, sErrMsg);

				if (g_prefs.m_planlocksystem) {
					
					if (m_pollDB.PlanLock(PLANLOCK_UNLOCK, &nCurrentPlanLock, szClientName, szClientTime, sErrorMessage2) == FALSE) {
						g_util.Logprintf("ERROR: PlanLock() - %s", sErrorMessage2);
						bReConnect = TRUE;
					}
				}
				if (nCurrentPublicationPlanLock == PLANLOCK_LOCK)
					m_pollDB.PublicationPlanLock(plan->m_publicationID, plan->m_pubdate, PLANLOCK_UNLOCK, &nCurrentPublicationPlanLock, szPublicationPlanClientName, szPublicationPlanClientTime, sErrorMessage2);

				m_pollDB.UpdateService(SERVICESTATUS_HASERROR, fi.sFileName, _T("Plan error"), sErrorMessage);
				m_pollDB.InsertLogEntry(EVENTCODE_PLANERROR, _T("Plan error"), fi.sFileName, sError, 0, 0, 0, _T(""), sErrorMessage);

				//				continue;
			}

			BOOL bRecyclePages = (ok == IMPORTPAGES_DONE) && bIsUpdate && g_prefs.m_recycledirtypages;

			if (bRecyclePages) {
				for (int press = 0; press < plan->m_NumberOfPresses; press++) {
					if (m_pollDB.RecycleProductionPages(plan->m_productionIDperpress[press], sErrorMessage) == FALSE) {
						g_util.Logprintf("ERROR: RecycleProductionPages() - %s", sErrorMessage);
						bReConnect = TRUE;
						break;
					}
				}
			}

			if (g_prefs.m_planlocksystem) {
				if (g_prefs.m_planlocksystem > 1 && nCurrentPlanLock == PLANLOCK_LOCK)
					m_pollDB.DeleteAllDirtyPages(plan->m_productionID, sErrorMessage);
				if (m_pollDB.PlanLock(PLANLOCK_UNLOCK, &nCurrentPlanLock, szClientName, szClientTime, sErrorMessage) == FALSE) {
					g_util.Logprintf("Error: PlanLock() - %s", sErrorMessage);
					bReConnect = TRUE;
				}
			}
			if (nCurrentPublicationPlanLock == PLANLOCK_LOCK)
				m_pollDB.PublicationPlanLock(plan->m_publicationID, plan->m_pubdate, PLANLOCK_UNLOCK, &nCurrentPublicationPlanLock, szPublicationPlanClientName, szPublicationPlanClientTime, sErrorMessage);


			if (ok == IMPORTPAGES_ERROR)
				continue;

			AddLogItem("Done", fi.sFileName + " " + plan->m_planID, sPress, nPageCount, sInfo, sTimeStamp);
			g_util.Logprintf("INFO: Done importing %s. %d pages found\r", fi.sFileName, nPageCount);
			::DeleteFile(sFullInputPath);
			g_tLastPoll = CTime::GetCurrentTime();
			if (bIsUpdate)
				g_TotalUpdateImports++;
			else
				g_TotalNewImports++;

			if (g_prefs.m_usepostcommand && g_prefs.m_postcommand != "" && bIgnorePostCommand == FALSE) {

				CString sPostCommand;
				sPostCommand.Format("%s %d", g_prefs.m_postcommand, plan->m_productionID);

				//SetCurrentInfoText(_T("Running post command file.."));
				if (g_util.DoCreateProcessEx(sPostCommand, g_prefs.m_postcommandtimeout) == FALSE)
					g_util.Logprintf("ERROR: Post command file %s failed", sPostCommand);
				else
					g_util.Logprintf("INFO: Post command file %s OK", sPostCommand);
			}

			// START OF FILE RETRY LOGIC (FileCenter)

			BOOL bMustBypassRetry = FALSE;
			m_pollDB.BypassRetryRequest(plan->m_productionID, bMustBypassRetry, sErrorMessage);

			g_util.Logprintf("INFO: Starting archive file retry processing.. bIgnorePostCommand=%d,bMustBypassRetry=%d", bIgnorePostCommand, bMustBypassRetry);

			if (g_prefs.m_useretryservice && bIgnorePostCommand == FALSE && bMustBypassRetry == FALSE) {
				if (m_pollDB.AddRetryRequest(plan->m_productionID, sErrorMessage) == TRUE)
					g_util.Logprintf("INFO:  File retry request added to queue for ProductionID %d", plan->m_productionID);
				else
					g_util.Logprintf("ERROR: Failed entering  retry request to queue for ProductionID %d - %s", plan->m_productionID, sErrorMessage);

				/*BOOL bDoPDFretry = (nExistingProductionID != plan->m_productionID) || (bAllNewSection);

				if (bDoPDFretry) {					
					if (m_pollDB.AddRetryRequestFileCenter(plan->m_productionID, sErrorMessage) == TRUE)
						g_util.Logprintf("INFO: File retry (FileCenter) request added to queue for ProductionID %d", plan->m_productionID);
					else
						g_util.Logprintf("ERROR: Failed entering retry request (FileCenter) to queue for ProductionID %d - %s", plan->m_productionID, sErrorMessage);					
				}*/
			}

			CString sProductionName;
			CString sEditionString = _T("");
			CString sSectionString = _T("");
			CString sPageDescription = _T("");

			for (int edition = 0; edition < plan->m_numberofeditions; edition++) {
				if (sEditionString != _T(""))
					sEditionString += _T("&");
				sEditionString += g_prefs.GetEditionName(plan->m_editionIDlist[0][edition]);
				if (edition == 0) {
					for (int section = 0; section < plan->m_numberofsections[0][0]; section++) {
						if (sSectionString != _T(""))
							sSectionString += _T("&");

						sPageDescription.Format(_T("%s 1-%d"), g_prefs.GetSectionName(plan->m_sectionIDlist[0][0][section]),
							plan->m_numberofpages[0][0][section]);

						sSectionString += sPageDescription;
					}
				}
			}

			sProductionName.Format("%s-%.2d%.2d-%s-%s", g_prefs.GetPublicationName(plan->m_publicationID), plan->m_pubdate.GetDay(), plan->m_pubdate.GetMonth(), sEditionString, sSectionString);

			m_pollDB.UpdateService(SERVICESTATUS_RUNNING, fi.sFileName, _T("Plan created"), sErrorMessage);
			m_pollDB.InsertLogEntry(nExistingProductionID>0 ? EVENTCODE_PLANCHANGED : EVENTCODE_PLANCREATED, bIsNewsPilotPlan ?  _T("NewsPilot plan") : _T("PPI plan"), fi.sFileName, sProductionName,  0, 0, 0, sPageDescription, sErrorMessage);

			m_pollDB.InsertPlanLogEntry(g_prefs.m_instancenumber, bIsUpdate ? EVENTCODE_PLANCHANGED : EVENTCODE_PLANCREATED,
					sProductionName, sPageDescription, plan->m_productionID, sSender, sErrorMessage);

		}

	} while (g_bRunning);

	g_util.Logprintf("INFO: Polling stopped");
	m_pollDB.UpdateService(SERVICESTATUS_STOPPED, _T(""), _T(""), sErrorMessage);

	m_pollDB.ExitDB();
	delete plan;

}

void	AddLogItem(CString sStatus, CString sInputName, CString sPress, int nPages, CString sComment, CString sTimeStamp)
{
	g_util.Logprintf("File:%s Status:%s Press:%s Pages:%d Comment:%s", sInputName,sStatus,sPress,nPages,sComment);
}


