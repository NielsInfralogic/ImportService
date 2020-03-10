
#include "stdafx.h"
#include <afxmt.h>
#include <afxtempl.h>
#include "DatabaseManager.h"
#include "Defs.h"
#include "Utils.h"
#include "Prefs.h"
#include "Markup.h"
#include "ImportXMLPages.h"
#include "WorkerThread.h"

extern CPrefs	g_prefs;
extern CUtils  g_util;
extern CDatabaseManager m_pollDB;
extern CPageTableEntry pageTableEntry;
extern BOOL		g_bRunning;

int ImportXMLpages(CDatabaseManager *pDB, PLANSTRUCT *plan, BOOL &bIsUpdate, CString &sErrMsg, CString &sInfo, BOOL bKeepExistingSections)
{
	CString s, s2, s3, s4, sText;
	CPageTableEntry *pItem = &pageTableEntry;
	pItem->Init();

	PUBLICATIONSTRUCT *pPublication = g_prefs.GetPublicationStruct(plan->m_publicationID);
	CString  sErrorMessage;

	PUBLICATIONPRESSDEFAULTSSTRUCT *pPublicationDefaults = NULL;
	if (pPublication != NULL && g_prefs.m_usedbpublicationdefatuls)
		pDB->RetrievePublicationPressDefaults(pPublication, sErrorMessage);

	int copyseparationset = 0;
	int copyflatseparationset = 0;

	int nPressRunID = 0;
	int nPageIndex = 1;

	// Gather statistics for production name data
	int nNumberOfPages = 0;
	int nNumberOfSeparations = 0;
	int nNumberOfSections = 0;
	int nNumberOfEditions = 0;
	int nPlannedApproval = 0;
	int nPlannedHold = 0;
	int bExistingProduction = 0;

	int nSequenceNumber = 1;

	int nNumberOfPagesPolled = 0;
	int nNumberOfPagesImaged = 0;

	///////////////////////////////////////
	// Dig out presses used in this import
	///////////////////////////////////////

	int nPressesUsed = 0;
	int nDifferentPressIDs[MAXPRESSES*MAXEDITIONS] = { 0 };

	for (int edition = 0; edition < plan->m_numberofeditions; edition++) {
		BOOL bFound = FALSE;
		for (int press = 0; press < plan->m_numberofpresses[0][edition]; press++) {

			for (int q = 0; q < nPressesUsed; q++) {
				if (nDifferentPressIDs[q] == plan->m_pressesinedition[0][edition][press]) {
					bFound = TRUE;
					break;
				}
			}
			if (!bFound)
				nDifferentPressIDs[nPressesUsed++] = plan->m_pressesinedition[0][edition][press];
		}
	}

	///////////////////////////////////////
	// Generate/get productionID per press used
	///////////////////////////////////////

	int	 currentproductionID[32] = { 0 }; // Per press!
	plan->m_NumberOfPresses = nPressesUsed;

	///////////////////////////////////////
	// Gather statistics for production table
	///////////////////////////////////////
	for (int press = 0; press < nPressesUsed; press++) {

		nNumberOfSections = 0;
		nNumberOfEditions = 0;
		nNumberOfSeparations = 0;
		for (int edition = 0; edition < plan->m_numberofeditions; edition++) {

			for (int editionpress = 0; editionpress < plan->m_numberofpresses[0][edition]; editionpress++) {

				if (nDifferentPressIDs[press] == plan->m_pressesinedition[0][edition][editionpress]) {

					nNumberOfEditions++;
					nNumberOfPages = 0;

					int m_pageidx = 1;

					for (int section = 0; section < plan->m_numberofsections[0][edition]; section++) {
						if (g_prefs.m_skipallcommon && plan->m_sectionallcommon[0][edition][section])
							continue;

						nNumberOfSections++;
						nNumberOfPages += plan->m_numberofpages[0][edition][section];
						nPlannedHold = plan->m_pagehold[0][edition][section][0];
						nPlannedApproval = plan->m_pageapproved[0][edition][section][0];

						for (int page = 0; page < plan->m_numberofpages[0][edition][section]; page++) {
							if (plan->m_pageindex[0][edition][section][page] == 0)
								plan->m_pageindex[0][edition][section][page] = m_pageidx++;
							nNumberOfSeparations += plan->m_numberofpagecolors[0][edition][section][page];
						}
					}
				}
			}
		}

		int nPressID = nDifferentPressIDs[press];
		currentproductionID[press] = pDB->CreateProductionID(nPressID, plan->m_publicationID, plan->m_pubdate,
								nNumberOfPages, nNumberOfSections, nNumberOfEditions,
								nPlannedApproval, nPlannedHold ? 2 : 3, bExistingProduction,
								plan->m_weekreference, PLANTYPE_UNAPPLIED, plan->m_rippressID, sErrorMessage);

		g_util.Logprintf("CreateProductionID() returned ProductionID=%d for PressID=%d", currentproductionID[press], nPressID);

		if (currentproductionID[press] == -1) {
			sErrMsg.Format( "Error getting ProductionID - %s", sErrorMessage);
			return IMPORTPAGES_DBERROR;
		}

		if (g_prefs.m_rejectappliedreimports && bExistingProduction == EXISTINGPRODUCTION_APPLIEDPLAN) {
			return IMPORTPAGES_REJECTEDAPPLIEDREIMPORT;
		}

		//		nDifferentPressIDs[press] = nPressID;

		plan->m_productionIDperpress[press] = currentproductionID[press];

	}

	bIsUpdate = bExistingProduction;


	
	int nCustomerID = 0;
	if (plan->m_customer != "") {
		nCustomerID = pDB->RetrieveCustomerID(plan->m_customer, sErrorMessage);
		if (nCustomerID == 0)
			nCustomerID = pDB->InsertCustomerName(plan->m_customer, plan->m_customeralias, sErrorMessage);
		if (nCustomerID == -1)
			nCustomerID = 0;
	}

	///////////////////////////////////////
	// Get current page count for existing production (for re-import status adjust
	///////////////////////////////////////

	WORD	 currentpagespersection[MAXISSUES][MAXEDITIONS][MAXSECTIONS] = { 0 };
	WORD	 currenttemplateIDpersection[MAXISSUES][MAXEDITIONS][MAXSECTIONS] = { 0 };

	BOOL	bPageCountChanged = FALSE;
	BOOL	bReImport = bExistingProduction;
	if (bExistingProduction) {
		for (int edition = 0; edition < plan->m_numberofeditions; edition++) {
			for (int section = 0; section < plan->m_numberofsections[0][edition]; section++) {
				if (g_prefs.m_skipallcommon && plan->m_sectionallcommon[0][edition][section])
					continue;
				int nNumberOfPages, nTemplateID;
				int nPressID = plan->m_pressIDlist[0][edition];

				// Set all pages Dirty

				if (!pDB->QueryProduction(plan->m_publicationID, plan->m_pubdate, plan->m_issueIDlist[0],
					plan->m_editionIDlist[0][edition], plan->m_sectionIDlist[0][edition][section], &nPressID,
					&nNumberOfPages, &nNumberOfPagesPolled, &nNumberOfPagesImaged, &nTemplateID,
					plan->m_weekreference, bKeepExistingSections, sErrorMessage)) {
					sErrMsg.Format("Error quering existing Production - %s", sErrorMessage);
					return IMPORTPAGES_DBERROR;
				}
				plan->m_pressIDlist[0][edition] = nPressID;
				currentpagespersection[0][edition][section] = nNumberOfPages;
				currenttemplateIDpersection[0][edition][section] = nTemplateID;

				if (nNumberOfPages != plan->m_numberofpages[0][edition][section])
					bPageCountChanged = TRUE;

				if (g_prefs.m_rejectpolledreimports && nNumberOfPagesPolled > 0) {
					pDB->KillDirtyflags(sErrorMessage);
					return IMPORTPAGES_REJECTEDPOLLEDREIMPORT;
				}

				if (g_prefs.m_rejectimagedreimports && nNumberOfPagesImaged > 0) {
					pDB->KillDirtyflags(sErrorMessage);
					return IMPORTPAGES_REJECTEDIMAGEDREIMPORT;
				}

			}
		}

		if (g_prefs.m_rejectreimports && bPageCountChanged == FALSE) {
			pDB->KillDirtyflags(sErrorMessage);
			return IMPORTPAGES_REJECTEDREIMPORT;
		}
	}
	else {
		BOOL bAllNonUnique = TRUE;
		for (int issue = 0; issue < plan->m_numberofissues; issue++) {
			for (int edition = 0; edition < plan->m_numberofeditions; edition++) {
				if (plan->m_allnonuniqueepages[0][edition] == FALSE) {
					bAllNonUnique = FALSE;
					break;
				}
			}
		}
		if (bAllNonUnique) {
			// No unique pages in this file and no existing production loaded - bail out.
			pDB->KillUnusedProductions(sErrorMessage);
			return IMPORTPAGES_NOEDITION;

		}
	}

	for (int press = 0; press < nPressesUsed; press++)
		pDB->UpdateProductionType(plan->m_productionIDperpress[press], PLANTYPE_UNAPPLIED, sErrorMessage);

	copyseparationset = 0;
	int nPageCounter = 0;
	int nNumberOfPressRuns = 0;
	int nPressRunsUsed[400];
	BOOL nPressRunsUsedAutoPlan[400] = { FALSE };
	CTime tPressTime = CTime::GetCurrentTime();


	TCHAR szThisMarkGroups[200];
	_tcscpy(szThisMarkGroups, "");
	if (pPublicationDefaults != NULL) {
		_tcscpy(szThisMarkGroups, "");
		TCHAR szMarkIDList[250];
		TCHAR szMarkID[50];
		_tcscpy(szMarkIDList, "");

		for (int mks = 0; mks < pPublicationDefaults->m_DefaultNumberOfMarkGroups; mks++) {

			int nMarkID = g_prefs.GetMarkGroupID(pPublicationDefaults->m_DefaultMarkGroupList[mks]);
			if (nMarkID > 0) {
				if (strlen(szMarkIDList) > 0)
					strcat(szMarkIDList, ",");
				sprintf(szMarkID, "%d", nMarkID);

				strcat(szMarkIDList, szMarkID);
			}
		}
		_tcscpy(szThisMarkGroups, szMarkIDList);
	}

	CStringArray sSectionTitles;
	for (int pageuniquetype = 1; pageuniquetype >= 0; pageuniquetype--) {
		nPageCounter = 0;
		for (int edition = 0; edition < plan->m_numberofeditions; edition++) {

			if (plan->m_editionnochange[0][edition]) {
				pDB->RestoreEditionDirtyFlag(plan->m_publicationID, plan->m_pubdate, plan->m_editionIDlist[0][edition], 0, sErrorMessage);
				continue;
			}

			for (int press = 0; press < plan->m_numberofpresses[0][edition]; press++) {

				int side = 0;
				int sheet = 0;

				int nPressID = plan->m_pressesinedition[0][edition][press];

				if (g_prefs.m_usedbpublicationdefatuls)
					pPublicationDefaults = g_prefs.GetPublicationDefaultStruct(plan->m_publicationID, nPressID);

				int nTemplateID = 1;// FindDefaultTemplate(nPressID, plan->m_publicationID);
				//if (nTemplateID == 0) {
				////	sErrMsg.Format("No template defined for requested press");
					//continue;
				//
//			}

				pItem->m_templateID = 1;
				//if (pPublicationDefaults != NULL) {
			//		pItem->m_templateID = pPublicationDefaults->m_DefaultTemplateID;
	//			}

				pItem->m_pressID = 1;
				//pItem->m_presstime = util.Translate2Time(plan->m_presstime[0][edition][press]);
				g_util.Translate2TimeEx(plan->m_presstime[0][edition][press], pItem->m_presstime);
				pItem->m_autorelease = plan->m_autorelease[0][edition][press];

				plan->m_productionID = currentproductionID[0]; // Safeguard..
				for (int pressx = 0; pressx < nPressesUsed; pressx++) {
					if (nDifferentPressIDs[pressx] == pItem->m_pressID) {
						plan->m_productionID = currentproductionID[pressx];
						break;
					}
				}

				pItem->m_pagesonplate = 1;
				pItem->m_locationID = g_prefs.GetPressLocationIDFromID(nPressID);

				pItem->m_customerID = nCustomerID;

				for (int section = 0; section < plan->m_numberofsections[0][edition]; section++) {

					if (g_prefs.m_skipallcommon && plan->m_sectionallcommon[0][edition][section])
						continue;

					CString sThisSectionTitle;
					sThisSectionTitle.Format("%d=%s", plan->m_sectionIDlist[0][edition][section], plan->m_sectionTitlelist[0][edition][section]);
					BOOL bFoundTitle = FALSE;
					for (int sx = 0; sx < sSectionTitles.GetCount(); sx++) {
						if (sSectionTitles[sx] == sThisSectionTitle) {
							bFoundTitle = TRUE;
							break;
						}
					}
					if (bFoundTitle == FALSE)
						sSectionTitles.Add(sThisSectionTitle);
					// Restart pressrun and sheetcounter if sections are split over different runs




					if (g_prefs.m_combinesections == FALSE) {
						sheet = 0;
						side = 0;
					}

					if (g_prefs.m_combinesections == FALSE || g_prefs.m_combinesections == TRUE && section == 0) {


						if (!pDB->GetNextPressRun(plan->m_publicationID, plan->m_pubdate, plan->m_issueIDlist[0], plan->m_editionIDlist[0][edition], plan->m_sectionIDlist[0][edition][section], nPressID, &nPressRunID, nSequenceNumber++, plan->m_weekreference, sErrorMessage)) {
							sErrMsg.Format("Error quering for pressrun - %s", sErrorMessage);
							if (g_prefs.m_planlocksystem > 1)
								pDB->DeleteAllDirtyPages(plan->m_productionID, sErrorMessage);
							else
								pDB->DeleteProductionPages(pItem->m_productionID, TRUE, sErrorMessage);
							pDB->KillUnusedProductions(sErrorMessage);
							pDB->KillUnusedPressRuns(sErrorMessage);

							return IMPORTPAGES_DBERROR;
						}

						CString sPressRunComment(plan->m_editioncomment[0][edition]);
						if (plan->m_planID != "" || plan->m_planName != "" || sPressRunComment != "")
							pDB->UpdatePressRun(nPressRunID, plan->m_sender, plan->m_planName, sPressRunComment, sErrorMessage);

						int nPageFormatIDPressRun = 0;
						CString thisPageFormat(plan->m_sectionPageFormat[0][edition][section]);
						if (thisPageFormat != "")
							nPageFormatIDPressRun = pDB->FindPageFormat(thisPageFormat, sErrorMessage);

						if (nPageFormatIDPressRun >= 0)  // Also set 0!!
							pDB->SetPressRunPageFormat(nPressRunID, nPageFormatIDPressRun, sErrorMessage);

						pDB->SetTimedEdition(nPressRunID,
							plan->m_editionistimed[0][edition] > 0 ? plan->m_editiontimedfromID[0][edition] : 0,
							plan->m_editionistimed[0][edition] > 0 ? plan->m_editiontimedtoID[0][edition] : 0,
							plan->m_editionordernumber[0][edition],
							plan->m_editionmasterzoneeditionID[0][edition],
							sErrorMessage);


						BOOL bExistingPressRun = FALSE;
						for (int prrun = 0; prrun < nNumberOfPressRuns; prrun++) {
							if (nPressRunsUsed[prrun] == nPressRunID && nPressRunID > 0) {
								bExistingPressRun = TRUE;
								break;
							}
						}
						if (bExistingPressRun == FALSE) {
							nPressRunsUsed[nNumberOfPressRuns] = nPressRunID;
							nPressRunsUsedAutoPlan[nNumberOfPressRuns] = pPublicationDefaults != NULL ? pPublicationDefaults->m_allowautoplanning : FALSE;
							nNumberOfPressRuns++;
						}
					}

					pItem->m_pressrunID = nPressRunID;

					for (int page = 0; page < plan->m_numberofpages[0][edition][section]; page++) {

						pItem->m_sheetside = side;
						pItem->m_sheetnumber = sheet / 2 + 1;
						sheet++;
						side = 1 - side;

						_tcscpy(pItem->m_sortingposition, "");
						_tcscpy(pItem->m_presstower, "");
						if (pPublicationDefaults != NULL) {
							if (pPublicationDefaults->m_DefaultPressTowerName != "")
								_tcscpy(pItem->m_presstower, pPublicationDefaults->m_DefaultPressTowerName);
						}
						_tcscpy(pItem->m_presszone, "");
						_tcscpy(pItem->m_presshighlow, "");

						//int nMaxColorsOnSheet = plan->m_sheetmaxcolors[0][edition][sheet][side];

						pItem->m_pagetype = PAGETYPE_DUMMY;
						pItem->m_uniquepage = TRUE;

						pItem->m_pageposition = 1;
						sprintf(pItem->m_pagepositions, "%d", pItem->m_pageposition);

						int sThisMasterEditionID = 0;

						// Find Master edition reference for this page in page arrays

						CString sThisPageID(plan->m_pageIDs[0][edition][section][page]);
						CString sThisMasterPageID(plan->m_pagemasterIDs[0][edition][section][page]);

						BOOL bFound = TRUE;

						if (sThisPageID != sThisMasterPageID) {
							bFound = FALSE;
							for (int edx = 0; edx < plan->m_numberofeditions; edx++) {
								for (int sectionx = 0; sectionx < plan->m_numberofsections[0][edx]; sectionx++) {
									for (int pagex = 0; pagex < plan->m_numberofpages[0][edx][sectionx]; pagex++) {
										CString sPageID(plan->m_pageIDs[0][edx][sectionx][pagex]);
										if (sThisMasterPageID.CompareNoCase(sPageID) == 0) {
											sThisMasterEditionID = plan->m_editionIDlist[0][edx];
											bFound = TRUE;
											break;
										}
									}
									if (bFound)
										break;
								}
								if (bFound)
									break;
							}

							// Safeguard
							if (bFound == FALSE)
								sThisMasterPageID = sThisPageID;
						}


						if (sThisMasterEditionID == 0)
							sThisMasterEditionID = plan->m_editionIDlist[0][edition];

						pItem->m_uniquepage = plan->m_pageunique[0][edition][section][page];

						if (pageuniquetype == UNIQUEPAGE_UNIQUE)
							pItem->m_uniquepage = UNIQUEPAGE_UNIQUE;	// Phase 1
						else if (pageuniquetype == UNIQUEPAGE_COMMON && pItem->m_uniquepage == UNIQUEPAGE_UNIQUE)	// Only update uniquepage=0 or 2 in second pass
							continue;

						// For re-import and page count change - tell separation inserter to change status from imaged
						int nCurrentPagesOnPlate = g_prefs.GetPagesOnPlateFromTemplateID(currenttemplateIDpersection[0][edition][section]);

						pItem->m_pagecountchange = currentpagespersection[0][edition][section] != plan->m_numberofpages[0][edition][section];

						int nNumberOfColors = plan->m_numberofpagecolors[0][edition][section][page];

						for (int color = 0; color < nNumberOfColors; color++) {
							int nColorID = plan->m_pagecolorIDlist[0][edition][section][page][color];
							pItem->m_colorIndex = nColorID;//color+1;

							_tcscpy(pItem->m_presscylinder, "");

							// Special usage for page plans only!!
							_tcscpy(pItem->m_formID, plan->m_pagemiscstring1[0][edition][section][page]);
							_tcscpy(pItem->m_plateID, plan->m_pagemiscstring2[0][edition][section][page]);

							CString sCol = g_prefs.GetColorName(nColorID);

							BOOL bIsBlack = FALSE;
							if (sCol.CompareNoCase("black") == 0 || sCol.CompareNoCase("k") == 0 || sCol.CompareNoCase("gray") == 0 || sCol.CompareNoCase("grey") == 0)
								bIsBlack = TRUE;

							///////////////////////////
							// Add separation to internal separation list
							///////////////////////////

							pItem->m_publicationID = plan->m_publicationID;
							pItem->m_weekreference = plan->m_weekreference;
							pItem->m_sectionID = plan->m_sectionIDlist[0][edition][section];
							pItem->m_editionID = plan->m_editionIDlist[0][edition];
							pItem->m_issueID = plan->m_issueIDlist[0];
							pItem->m_pubdate = plan->m_pubdate;
							_tcscpy(pItem->m_pagename, plan->m_pagenames[0][edition][section][page]);

							/// ###########################

							pItem->m_proofID = g_prefs.GetProofID(g_prefs.m_proofer);
							if (plan->m_usemask)
								pItem->m_proofID = g_prefs.GetProofID(g_prefs.m_proofermasked);

							pItem->m_deviceID = 0;

							pItem->m_version = plan->m_pageversion[0][edition][section][page];
							pItem->m_layer = 1;
							pItem->m_pagination = plan->m_pagepagination[0][edition][section][page];

							pItem->m_approved = plan->m_pageapproved[0][edition][section][page];
							pItem->m_hold = plan->m_pagehold[0][edition][section][page];


							if (g_prefs.m_alwaysblack && !bIsBlack)
								pItem->m_active = 0;
							else
								pItem->m_active = plan->m_pagecolorActivelist[0][edition][section][page][color];

							pItem->m_priority = plan->m_pagepriority[0][edition][section][page];
							pItem->m_pagetype = plan->m_pagetypes[0][edition][section][page];
							pItem->m_presssectionnumber = nNumberOfPressRuns;
							pItem->m_productionID = plan->m_productionID;



							pItem->m_proofstatus = 0;
							pItem->m_inkstatus = 0;

							_tcscpy(pItem->m_planpagename, plan->m_pagefilenames[0][edition][section][page]);
							pItem->m_issuesequencenumber = 1; //plan->m_pagemasteredition[0][edition][section][page]; // issue+1;

							pItem->m_colorID = nColorID;

							pItem->m_flatproofID = 0;

							pItem->m_creep = 0.0;
							pItem->m_pageindex = plan->m_pageindex[0][edition][section][page];

							pItem->m_hardproofID = 0;
							_tcscpy(pItem->m_fileserver, "");

							pItem->m_deadline = CTime(1975, 1, 1, 0, 0, 0);					// To Be IMPLEMENTED		
							//pItem->m_presstime =  CTime(1975,1,1,0,0,0); 
							_tcscpy(pItem->m_comment, plan->m_pagecomments[0][edition][section][page]);

							pItem->m_mastereditionID = sThisMasterEditionID; //plan->m_pagemasteredition[0][edition][section][page];
							if (pageuniquetype == 1)
								pItem->m_mastereditionID = pItem->m_editionID;

							if (pItem->m_pagetype == PAGETYPE_ANTIPANORAMA)
								pItem->m_active = FALSE;

							int nCopiesToInsert = plan->m_numberofcopies[0][edition][0];
							int nActiveCopies = plan->m_sheetactivecopies[0][edition][sheet][side];
							if (pPublicationDefaults != NULL) {
								//if (pPublicationDefaults->m_DefaultCopies > 0) {
								//	nCopiesToInsert = pPublicationDefaults->m_DefaultCopies;
								//	nActiveCopies = pPublicationDefaults->m_DefaultCopies;
								//}

								if (pPublication != NULL) {
									pItem->m_approved = pPublication->m_Approve == 0 ? -1 : 0;

									pItem->m_proofID = plan->m_usemask ? g_prefs.GetProofID(g_prefs.m_proofermasked) : pPublication->m_ProofID;
									pItem->m_hardproofID = pPublication->m_HardProofID;

									//										CTime tm = pItem->m_pubdate;
									//										tm += CTimeSpan(-1* pPublication->m_deadline.GetDay(),  pPublication->m_deadline.GetHour(), pPublication->m_deadline.GetMinute() , 0 );
									//										pItem->m_deadline = tm;

								}
							}

							// Last catch-all safety valve....
							if (pItem->m_proofID == 0 && g_prefs.m_ProofList.GetCount() > 0)
								pItem->m_proofID = g_prefs.m_ProofList[0].m_ID;

							if (pItem->m_hardproofID == 0 && g_prefs.m_ProofList.GetCount() > 0)
								pItem->m_hardproofID = g_prefs.m_ProofList[0].m_ID;


							if (g_prefs.m_ignoreactivecopies)
								nActiveCopies = nCopiesToInsert;

							if (g_prefs.m_onlyuseactivecopies) {
								nActiveCopies = plan->m_sheetactivecopies[0][edition][sheet][side];
								nCopiesToInsert = nActiveCopies;
							}

							if (pageuniquetype == 1) { // First iteration - hardwire to unique
								if (pDB->InsertSeparation(pItem, &copyseparationset, &copyflatseparationset, nNumberOfColors, TRUE, nCopiesToInsert, nActiveCopies, sErrorMessage) == FALSE) {
									sErrMsg.Format("Error inserting separation - %s", sErrorMessage);
									if (g_prefs.m_planlocksystem > 1)
										pDB->DeleteAllDirtyPages(plan->m_productionID, sErrorMessage);
									else
										pDB->DeleteProductionPages(pItem->m_productionID, TRUE, sErrorMessage);
									pDB->KillUnusedProductions(sErrorMessage);
									pDB->KillUnusedPressRuns(sErrorMessage);
									return IMPORTPAGES_DBERROR;
								}
							}
							else {					// Second iteration - fix up unique and masterset numbers for non-unique pages
								if (pDB->InsertSeparationPhase2(pItem, g_prefs.m_pressspecificpages, sErrorMessage) == FALSE) {
									sErrMsg.Format("Error inserting separation - %s", sErrorMessage);
									if (g_prefs.m_planlocksystem > 1)
										pDB->DeleteAllDirtyPages(plan->m_productionID, sErrorMessage);
									else
										pDB->DeleteProductionPages(pItem->m_productionID, TRUE, sErrorMessage);
									pDB->KillUnusedProductions(sErrorMessage);
									pDB->KillUnusedPressRuns(sErrorMessage);
									return IMPORTPAGES_DBERROR;
								}
							}

							nPageCounter++;
						
						} // end for color

					} // end for page
				}  // end for section
			} // end for press..

			// Re-apply for plan lock for each edition
			if (g_prefs.m_planlocksystem) {
				TCHAR szClientName[100];
				TCHAR szClientTime[50];

				int nCurrentPlanLock = PLANLOCK_UNKNOWN;

				while (nCurrentPlanLock != PLANLOCK_LOCK && g_bRunning) {
					if (pDB->PlanLock(PLANLOCK_LOCK, &nCurrentPlanLock, szClientName, szClientTime, sErrorMessage) == FALSE) {
						break;
					}
					if (nCurrentPlanLock == PLANLOCK_NOLOCK) {
						Sleep(3000);
					}
					else if (nCurrentPlanLock == PLANLOCK_LOCK) {
						break;
					}
				}
			}
		} // end for edition..
	}

	/*	if (plan->m_usemask) {
			int nPageFormatID = pDB->FindPageFormat(plan->m_masksizex,plan->m_masksizey,plan->m_maskbleed, plan->m_masksnapeven, plan->m_masksnapodd, sErrorMessage);
			if (nPageFormatID == 0)
				nPageFormatID = pDB->AddPageFormat(plan->m_masksizex,plan->m_masksizey,plan->m_maskbleed, plan->m_masksnapeven, plan->m_masksnapodd, plan->m_maskshowspread, sErrorMessage);
			if (nPageFormatID > 0) {
				for (int run=0; run<nNumberOfPressRuns; run++)
					pDB->SetPressRunPageFormat(nPressRunsUsed[run], nPageFormatID, sErrorMessage);
			}
		}
	*/
	BOOL bDoApplyProduction = FALSE;
	if (g_prefs.m_usedbpublicationdefatuls) {
		for (int run = 0; run < nNumberOfPressRuns; run++) {
			if (nPressRunsUsedAutoPlan[run])
				bDoApplyProduction = TRUE;
		}
	}

	if ((bDoApplyProduction && g_prefs.m_allowautoapply) || g_prefs.m_autoapplyalways) {
		for (int press = 0; press < nPressesUsed; press++) {
			BOOL bInsertedSections = g_prefs.m_applyinserted;
			BOOL nSeparateRuns = TRUE;
			int nPaginationMode = 0;
			PUBLICATIONPRESSDEFAULTSSTRUCT *pdefaults = g_prefs.GetPublicationPressDefaultStruct(plan->m_publicationID, nDifferentPressIDs[press] > 0 ? nDifferentPressIDs[press] : 1);
			if (pdefaults != NULL) {
				bInsertedSections = pdefaults->m_insertedsections;
				nSeparateRuns = pdefaults->m_separateruns;
				nPaginationMode = pdefaults->m_paginationmode;
			}
			g_util.Logprintf("INFO: AutoApply productionID %d", plan->m_productionIDperpress[press]);
			m_pollDB.AutoApplyProduction(plan->m_productionIDperpress[press], bInsertedSections, nPaginationMode, nSeparateRuns, sErrorMessage);
		}
	}
	if (g_prefs.m_runpostprocedure) {
		for (int run = 0; run < nNumberOfPressRuns; run++) {
			pDB->PostImport(nPressRunsUsed[run], sErrorMessage);
		}
	}
	if (g_prefs.m_runpostprocedure3) {
		for (int run = 0; run < nNumberOfPressRuns; run++) {
			pDB->LinkMasterCopySepSetsToOtherPress(nPressRunsUsed[run], sErrorMessage);
		}
	}

	// Set press-specific codes (e.g. ABB codes)
	if (g_prefs.m_runpostprocedure2 && plan->m_runpostproc) {
		for (int run = 0; run < nNumberOfPressRuns; run++)
			pDB->PostImport2(nPressRunsUsed[run], sErrorMessage);
	}

	if (g_prefs.m_forceunappliededitions && g_prefs.m_unappliededitionlist.Trim() != "") {
		CStringArray sArr;
		int m = g_util.StringSplitter(g_prefs.m_unappliededitionlist, ",", sArr);
		for (int eds = 0; eds < sArr.GetCount(); eds++) {
			int id = g_prefs.GetEditionID(sArr[eds]);
			if (id > 0) {
				for (int press = 0; press < nPressesUsed; press++) {
					m_pollDB.UnApplyEdition(plan->m_productionIDperpress[press], id, sErrorMessage);
				}
			}
		}
	}

	if (g_prefs.m_runpostprocedureProduction) {
		for (int press = 0; press < nPressesUsed; press++) {
			m_pollDB.PostProduction(plan->m_productionIDperpress[press], sErrorMessage);
		}
	}

	CStringArray sExistingSectionTitles;
	CString sExistingSectionTitleString = _T("");
	m_pollDB.GetProductionOrderNumber(plan->m_productionIDperpress[0], sExistingSectionTitleString, sErrorMessage);
	g_util.Logprintf("Existing Production-ordernumber '%s'", sExistingSectionTitleString);
	g_util.Logprintf("New Production-ordernumber '%s'", sExistingSectionTitleString);


	if (sExistingSectionTitleString != "") {
		if (sExistingSectionTitleString.Right(1) == ",")
			sExistingSectionTitleString = sExistingSectionTitleString.Left(sExistingSectionTitleString.GetLength() - 1);
		g_util.StringSplitter(sExistingSectionTitleString, ",", sExistingSectionTitles);

		// Merge existing sections with new ones
		for (int i = 0; i < sExistingSectionTitles.GetCount(); i++) {
			if (sExistingSectionTitles[i].Trim() == "")
				continue;
			BOOL bFoundSection = FALSE;
			for (int j = 0; j < sSectionTitles.GetCount(); j++) {

				if (GetSectionItem(sExistingSectionTitles[i]) == GetSectionItem(sSectionTitles[j])) {
					bFoundSection = TRUE;
					break;
				}
			}
			if (bFoundSection == FALSE)
				sSectionTitles.Add(sExistingSectionTitles[i].Trim());
		}
	}


	CString sSectionTitleString = _T("");
	for (int i = 0; i < sSectionTitles.GetCount(); i++) {
		if (sSectionTitles[i] == "")
			continue;
		if (sSectionTitleString != _T(""))
			sSectionTitleString += _T(",");
		sSectionTitleString += sSectionTitles[i];
	}
	if (g_prefs.m_setsectiontitles && sSectionTitleString.GetLength() > 0) {
		if (sSectionTitleString.Right(1) == ",")
			sSectionTitleString = sSectionTitleString.Left(sSectionTitleString.GetLength() - 1);
		g_util.Logprintf("New Production-ordernumber '%s'", sSectionTitleString);

		m_pollDB.UpdateProductionOrderNumber(plan->m_productionIDperpress[0], sSectionTitleString, sErrorMessage);
	}

	if (g_prefs.m_runpostprocedureProduction2) {
		for (int press = 0; press < nPressesUsed; press++) {
			m_pollDB.PostProduction2(plan->m_productionIDperpress[press], sErrorMessage);
		}
	}

	return TRUE;
}


CString GetSectionItem(CString s)
{
	CString sN = _T("");
	if (s.GetLength() < 2)
		return s;

	int n = s.Find(_T(":"));
	if (n == -1)
		return s;
	return s.Left(n);
}

int GetSectionIndex(PLANSTRUCT *plan, int nSectionID)
{
	for (int ed = 0; ed < plan->m_numberofeditions; ed++) {
		for (int section = 0; section < plan->m_numberofsections[0][ed]; section++) {
			if (plan->m_sectionIDlist[0][ed][section] == nSectionID)
				return section;
		}
	}
	return 0;
}