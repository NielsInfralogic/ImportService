#pragma once
#include <afxdb.h>
#include "Defs.h"
#include "PlanDataDefs.h"

class CDatabaseManager
{
public:
	CDatabaseManager(void);
	virtual ~CDatabaseManager(void);

	BOOL	InitDB(CString sDBserver, CString sDatabase, CString sDBuser, CString sDBpassword, BOOL bIntegratedSecurity, CString &sErrorMessage);
	BOOL	InitDB(CString &sErrorMessage);
	void	ExitDB();
	BOOL	IsOpen();

	BOOL	RegisterService(CString &sErrorMessage);
	BOOL	UpdateService(int nCurrentState, CString sCurrentJob, CString sLastError, CString &sErrorMessage);
	BOOL	UpdateService(int nCurrentState, CString sCurrentJob, CString sLastError, CString sAddedLogData, CString &sErrorMessage);
	BOOL	InsertLogEntry(int nEvent, CString sSource, CString sFileName, CString sMessage, int nMasterCopySeparationSet, int nVersion, int nMiscInt, CString sMiscString, CString &sErrorMessage);

	BOOL	LoadImportConfigurations(int nInstanceNumber,  CString &sErrorMessage);
	BOOL	LoadConfigIniFile(int nInstanceNumber, CString sFileName, CString sFileName2, CString &sErrorMessage);
	BOOL	LoadSTMPSetup(CString &sErrorMessage);

	BOOL	LoadAllPrefs(CString &sErrorMessage);
	BOOL	LoadDBNameList(CString sIDtable, ITEMLIST *v, CString &sErrorMessage);
	CString LoadNewPublicationName(int nID, CString &sErrorMessage);
	int		LoadNewPublicationID(CString sName, CString &sErrorMessage);
	int		LoadNewSectionID(CString sName, CString &sErrorMessage);
	int		LoadNewEditionID(CString sName, CString &sErrorMessage);
	int		GetPublicationID(CString sName, CString &sErrorMessage);

	PUBLICATIONSTRUCT *LoadPublicationStruct(int nID, CString &sErrorMessage);

	BOOL	LoadEditionNameList(CString &sErrorMessage);
	int		LoadLocationList(CString &sErrorMessage);
	BOOL	LoadTemplateList(CString &sErrorMessage);
	int		LoadColorList(CString &sErrorMessage);
	BOOL	LoadMarkGroupList(CString &sErrorMessage);
	BOOL	LoadProofProcessList(CString &sErrorMessage);
	BOOL	LoadDeviceList(CString &sErrorMessage);
	BOOL	LoadPressList(CString &sErrorMessage);

	int		RetrieveTemplatePageFormatName(int nTemplateID, CString &sErrorMessage);
	BOOL	RetrievePageFormatList(CString &sErrorMessage);
	BOOL	RetrievePublicationList(CString &sErrorMessage);
	BOOL	RetrievePublicationPressDefaults(PUBLICATIONSTRUCT *pPublication, CString &sErrorMessage);

	BOOL	LoadGeneralPreferences(CString &sErrorMessage);
	BOOL	UnApplyEdition(int nProductionID, int nEditionID, CString &sErrorMessage);

	BOOL	GetPublicationChannels(int nPublicationID, CUIntArray &aChannels, CString &sErrorMessage);
	int		CreateProductionID(int &nPressID, int nPublicationID, CTime tPubDate, int nNumberOfPages, int nNumberOfSections, int nNumberOfEditions, int nPlannedApproval, int nPlannedHold, int &bExistingProduction, int nWeekNumber, int nPlanType, int nRipPressID, CString &sErrorMessage);
	BOOL	InsertSeparation(CPageTableEntry *pItem, int *nCopySeparationSet, int *nCopyFlatSeparationSet, int nNumberOfColors, BOOL bFirstPagePosition, int nCopies, int nCopiesToProduce, CString &sErrorMessage);
	BOOL	InsertSeparationPhase2(CPageTableEntry *pItem, CString &sErrorMessage);
	BOOL	InsertSeparationPhase2(CPageTableEntry *pItem, BOOL bPressSpecific, CString &sErrorMessage);
	BOOL	DeleteProductionPages(int nProductionID, BOOL bRecyclePages, CString &sErrorMessage);
	BOOL	QueryProduction(int nPublicationID, CTime tPubDate, int nIssueID, int nEditionID, int nSectionID, int *nPressID, int *nNumberOfPage, int *nNumberOfPagesPolled, int *nNumberOfPagesImaged, int *nTemplateID, int nWeekNumber, BOOL bKeepExistingSection, CString &sErrorMessage);

	BOOL	GetNextPressRun(int nPublicationID, CTime tPubDate, int nIssueID, int nEditionID, int nSectionID, int nPressID, int *nPressRunID, int nSequenceNumber, int nWeekNumber, CString &sErrorMessage);
	BOOL	UpdateAliasName(CString sTypeName, CString sLongName, CString sShortName, CString &sErrorMessage);
	BOOL	InsertPublicationNames(int nPublicationID, CString sName, int nPageFormatID, BOOL bTrimToPageformat, double fLatestHours, int nProofID, int nHardProofID, int nApproveMode, CString &sErrorMessage);
	BOOL	UpdatePublicationName(CString sNameBefore, CString sNewName, CString &sErrorMessage);


	BOOL	DeleteProductionPagesEx(CString sOrderNumber, CString &sErrorMessage);

	CString TranslateDate(CTime tDate);
	CString TranslateTime(CTime tDate);
	CString TranslateTimeEx(SYSTEMTIME tDate);

	
	BOOL	KillDirtyflags(CString &sErrorMessage);
	BOOL	KillUnusedProductions(CString &sErrorMessage);
	BOOL	KillUnusedPressRuns(CString &sErrorMessage);
	BOOL	LoadUser(CString sUserName, CString sPassword, int *nReturnCode, BOOL *bIsAdmin, CString &sErrorMessage);

	BOOL	InsertDBName(CString sNameTable, CString sIDFieldName, CString sNameFieldName, int nID, CString sName, CString &sErrorMessage);
	BOOL	InsertEditionName(int nID, CString sNewName, BOOL isCommon, int MasterID, CString &sErrorMessage);
	int		RetrieveNextValueCount(CString sID, CString sIDtable, CString &sErrorMessage);
	int		RetrieveMaxValueCount(CString sID, CString sIDtable, CString &sErrorMessage);
	BOOL	UpdateNotificationTable(CString sTableName, CString &sErrorMessage);

	BOOL	PostImport(int nPressRunID, CString &sErrorMessage);
	BOOL	UpdatePressRun(int nPressRunID, CString sPlanID, CString sPlanName, CString sInkComment, CString &sErrorMessage);

	void	LogprintfDB(const TCHAR *msg, ...);

	BOOL	UpdateProductionType(int nProductionID, int nPlanType, CString &sErrorMessage);

	int		RetrieveCustomerID(CString sCustomerName, CString &sErrorMessage);
	int		InsertCustomerName(CString sCustomerName, CString sCustomerAlias, CString &sErrorMessage);
	BOOL	DeleteInputAlias(CString sType, CString sShortName, CString &sErrorMessage);

	int		GetEditionID(CString sName, CString &sErrorMessage);
	int		GetSectionID(CString sName, CString &sErrorMessage);

	BOOL	DeleteAllDirtyPages(int nProductionID, CString &sErrorMessage);

	BOOL	PlanLock(int bRequestedPlanLock, int *bCurrentPlanLock, TCHAR *szClientName, TCHAR *szClientTime, CString &sErrorMessage);
	BOOL	PublicationPlanLock(int nPublicationID, CTime tPubDate, int bRequestedPlanLock, int *bCurrentPlanLock, TCHAR *szClientName, TCHAR *szClientTime, CString &sErrorMessage);

	BOOL	RetrievePressTowerList(int nPressID, CString &sErrorMessage);

	BOOL	SetTimedEdition(int nPressRunID, int nFromTimedEditionID, int nToTimedEditionID, int nEditionOrderNumber, int nZoneMasterEditionID, CString &sErrorMessage);
	BOOL	SetAutoRelease(int nPressRunID, BOOL bAutoRelease, BOOL bRejectSubEditionPages, CString &sErrorMessage);


	int		FieldExists(CString sTableName, CString sFieldName, CString &sErrorMessage);
	int		TableExists(CString sTableName, CString &sErrorMessage);
	int		StoredProcedureExists(CString sProcedureName, CString &sErrorMessage);
	int		StoredProcParameterExists(CString sSPname, CString sParameterName, CString &sErrorMessage);



	BOOL	AutoApplyProduction(int nProductionID, BOOL bInsertedSections, int nPaginationMode, int nSeparateRuns, CString &sErrorMessage);

	__int64 GetFirstSeparation(int nProductionID, CString &sErrorMessage);
	BOOL	InsertPlanLogEntry(int nProcessID, int nEventCode, CString sProductionName, CString sPageDescription, int nProductionID, CString sSender, CString &sErrorMessage);

	BOOL	SetPressRunPageFormat(int nPressRunID, int nPageFormatID, CString &sErrorMessage);
	int		FindPageFormat(double fSizeX, double fSizeY, double fBleed, int nSnapEven, int nSnapOdd, CString &sErrorMessage);
	int		FindPageFormat(double fSizeX, double fSizeY, CString &sErrorMessage);
	int		InsertPageFormat(CString sName, double fSizeX, double fSizeY, double fBleed,
		int nSnapEven, int nSnapOdd, int nType, CString &sErrorMessage);

	BOOL	RestoreEditionDirtyFlag(int nPublicationID, CTime tPubDate, int nEditionID, int nPressID, CString &sErrorMessage);
	BOOL	PostImport2(int nPressRunID, CString &sErrorMessage);
	BOOL	PostProduction(int nProductionID, CString &sErrorMessage);

	BOOL	RecycleProductionPages(int nProductionID, CString &sErrorMessage);

	BOOL	AddRetryRequest(int nProductionID, CString &sErrorMessage);

	BOOL	LinkMasterCopySepSetsToOtherPress(int nPressRunID, CString &sErrorMessage);
	BOOL	ImportCenterPressRunCustomPage(int nPressRunID, int nOtherProductionID, int EditionID, int SectionID, CString sPageName, CString &sErrorMessage);
	BOOL	GetOtherUniqueProductionID(int nPressID, CTime tPubDate, int nPublicationID, int nEditionID, int nSectionID, int &nOtherProductionID, CString &sErrorMessage);
	BOOL	GetPagesInPressRun(int nPressRunID, int nSectionID, CStringArray &aPages, CString &sErrorMessage);
	BOOL	GetPressRunDetails(int nPressRunID, int &nProductionID, int &nPressID, CTime &tPubDate, int &nPublicationID, int &nEditionID, CUIntArray &aSectionID, CString &sErrorMessage);
	int		GetPlanType(int nPublicationID, CTime tPubDate, int nPressID, int &nExstingProductionID, CString &sErrorMessage);
	int		GetPageCountInSection(int nPublicationID, CTime tPubDate, int nEditionID, int nSectionID, CString &sErrorMessage);

	BOOL	GetLog(CString &sErrorMessage);

	BOOL	GetProductionOrderNumber(int nProductionID, CString &sOrderNumber, CString &sErrorMessage);
	BOOL	UpdateProductionOrderNumber(int nProductionID, CString sOrderNumber, CString &sErrorMessage);

	BOOL	PostProduction2(int nProductionID, CString &sErrorMessage);
	BOOL	PostImport2Version2(int nPressRunID, CString &sErrorMessage);

	int		FindPageFormat(CString sPageFormatName, CString &sErrorMessage);
	BOOL	ResetAllPollLocks(CString &sErrorMessage);
	BOOL	BypassRetryRequest(int nProductionID, BOOL &bBypassRetry, CString &sErrorMessage);
	BOOL	AddRetryRequestFileCenter(int nProductionID, CString sMask, BOOL bProcessOnlyUnprocessedFiles, CString &sErrorMessage);
	BOOL	AddRetryRequestFileCenter(int nProductionID, CString sMask, CString &sErrorMessage);
	BOOL	AddRetryRequestFileCenter(int nProductionID, CString &sErrorMessage);

	CString LoadSpecificAlias(CString sType, CString sShortName, CString &sErrorMessage);


	///////
	BOOL	NPDeleteOldData(int nDaysToKeep, CString &sErrorMessage);
	int		NPLookupPageEntry(CString sPublication, CTime tPubDate, CString sLocation, CString sEdition, CString sZone, CString sSection, CString sPageName, CString &sErrorMessage);
	BOOL	NPResetPages(CString &sErrorMessage);
	BOOL	NPDeletePages(CString sLocation, CString sPublication, CTime tPubDate, CString sZone, CString sEdition, CString sSection, CString &sErrorMessage);
	BOOL	NPAddPage(CString sPublication, CTime tPubDate, CString sProdcutionName, int nSeqNo, CString sLocation, CString sEdition,
						CString sZone, CString sSection,
						CString sPageName, int nPagination, CString sPlanPagename,
						CString sComment, CString sColors, CString sVersion, CTime tDeadLine,
						CString &sErrorMessage, int &bAdded);
	BOOL	NPAddPage(CString sPublication, CTime tPubDate, CString sProdcutionName, int nSeqNo, CString sLocation, CString sEdition,
						CString sZone, CString sSection,
						CString sPageName, int nPagination, CString sPlanPagename,
						CString sComment, CString sColors, CString sVersion, CTime tDeadLine,
						int nMiscInt1, int nMiscInt2, int nMiscInt3, int nMiscInt4, CString sMiscString1, CString sMiscString2, CString sMiscString3, CString sMiscString4, CTime tMiscDate,
						CString &sErrorMessage, int &bAdded);

	BOOL	NPDeleteIllegalLocationData(CString &sErrorMessage);
	BOOL	NPChangeSeqNumber(CString sPublication, CTime tPubDate, CString sLocation, int nSeq, CString &sErrorMessage);
	BOOL	NPGetPageCount(CString sPublication, CTime tPubDate, CString sEdition, CString sSection, int &nPageCount, CString &sErrorMessage);
	BOOL	NPGetPageTablePageCount(CString sPublication, CTime tPubDate, CString sEdition, CString sSection, int &nPageCount, BOOL bCountAll, CString &sErrorMessage);
	BOOL	NPGetSequenceNumberCount(CString sPublication, CTime tPubDate, CString sLocation, int &nSequeceNumbers, CString &sErrorMessage);
	BOOL	NPGetInfoFromImportTable(int nMode, CString sPublication, CTime tPubDate, CString sLocation, CString sZoneEdition, CStringArray &aList, CString &sErrorMessage);
	BOOL	NPGetInfoFromImportTableEx(int nMode, CString sPublication, CTime tPubDate, CString sLocation, CString sZoneEdition, CStringArray &aList, CString &sErrorMessage);
	BOOL	NPGetTimedMasterEditionOnLocation(CString sPublication, CTime tPubDate, CString sLocation, CString &sEdition, CString &sZone, int &nSequenceNumber, CString &sErrorMessage);
	BOOL	NPGetZoneMasterEditionOnLocation(CString sPublication, CTime tPubDate, CString sLocation, CString &sZone, CString &sErrorMessage);
	BOOL	NPGetZoneMaster(CString sPublication, CTime tPubDate, CString &sLocation, CString &sZone, CString &sErrorMessage);
	BOOL	NPGetSlaveZonesOnLocation(CString sPublication, CTime tPubDate, CString sLocation, CString sMasterZone,
																		CStringArray &sZones, CString &sErrorMessage);
	BOOL	NPGetSlaveZones(CString sPublication, CTime tPubDate, CString sMasterZone, CString sLocation,
																		CStringArray &sZones, CString &sErrorMessage);
	BOOL	NPGetTimedSlaveEditionsOnZoneLocation(CString sPublication, CTime tPubDate, CString sLocation, CString sMasterZone, CString sMasterEdition,
																		CStringArray &sEditions, CString &sErrorMessage);
	BOOL	NPLoadImportTable(PlanDataEdition *editem, CString sPublication, CTime tPubDate, CString sLocation, CString sEdition, int &nPagesLoaded, CString &sErrorMessage);

	BOOL	LoadPPITanslations(CString &sErrorMessage);

	BOOL	InsertFileSendRequest(CString sSourceFile, CString &sErrorMessage);
private:
	BOOL	m_DBopen;
	CDatabase *m_pDB;

	BOOL	m_IntegratedSecurity;
	CString m_DBserver;
	CString m_Database;
	CString m_DBuser;
	CString m_DBpassword;
	BOOL	m_PersistentConnection;
};
