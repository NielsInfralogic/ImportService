#pragma once

#include "stdafx.h"
#include "Defs.h"
#include "DatabaseManager.h"

int ImportXMLpages(CDatabaseManager *pDB, PLANSTRUCT *plan, BOOL &bIsUpdate, CString &sErrMsg, CString &sInfo, BOOL bKeepExistingSections);
CString GetSectionItem(CString s);
int GetSectionIndex(PLANSTRUCT *plan, int nSectionID);

