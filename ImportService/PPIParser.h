#pragma once

BOOL ProcessPPIFile(CDatabaseManager &DB, CString sFullInputPath, CString &sOutputFile, CString &sErrorMessage);
int PPIParseXML(CString sFullInputPath, PlanData &pageplan, CString sFixedEditionNames, CString sFixedSectionName, int &nPagesImported, CString &sInfo, CString &sErrorMessage);

CString LookupPPIName(CString sPPIProduct, CString sPPIEdition);
