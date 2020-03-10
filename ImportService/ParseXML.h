#pragma once
#include "stdafx.h"
#include "Defs.h"

BOOL IsInternalXML(CString sFullInputPath, CString &sErrorMessage);
int ParseXML(CString sFullInputPath, PLANSTRUCT *plan, int &nPagesImported, CString &sErrMsg, CString &sInfo,
										CString &sPressName, BOOL &bIgnorePostCommand, BOOL &bKeepExistingSections);
int GetPagePosition(int nPagesOnPlate, int Xpos, int Ypos);
int GetPageColors(PLANSTRUCT *plan, CString sPageID);
void SetMaxPageColors(PLANSTRUCT *plan);
int  GetPageSectionID(PLANSTRUCT *plan, int ed, CString sPageID);

void LogXMLcontent(CString sFileName, PLANSTRUCT *plan);



