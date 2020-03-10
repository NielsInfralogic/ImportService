#pragma once
#include "stdafx.h"




class CNPPrefs : public CObject
{
public:
	CNPPrefs();
	virtual ~CNPPrefs();
	void  LoadIniFile(CString sIninFile);

	BOOL m_multipleregextests;

	BOOL		m_useregexp;
	CString	m_matchexpressionheader;
	CString	m_formatexpressionheader;
	CString	m_matchexpressiontop;
	CString	m_formatexpressiontop;
	int		m_numberofexpressions;

	CString	m_matchexpressionbody[16];
	CString	m_formatexpressionbody[16];

	CString m_headercontents;
	CString m_subheadertriggerexpr;
	CString m_subheadercontents;
	CString m_datarowcontents;
	CString m_dateformat;
	CString m_timeformat;
	CString m_defaultlocation;
	CString m_matchexpressionfilename;
	CString m_formatexpressionfilename;
	CString m_filenamecontents;
	CString m_defaultzone;
	CString m_defaultedition;

	CString	m_preparsematchexpression;
	CString	m_preparseformatexpression;

	int m_numberofzonelocations;
	CString m_rootfolder;
	ZONELOCATIONSTRUCT m_zonelocations[MAXZONES];

	CString m_filenamefilterexpression;

	 BOOL m_allzonesforced;
	 BOOL	m_rejectdifferentpagecount;
	 BOOL m_rejectidenticalplan;

	 BOOL m_waitforimportcenter;

	 BOOL m_pagenameprefixfromcomment;

	 BOOL m_zeroextend;
	 NPEXTRAEDITIONLIST m_ExtraEditionList;
	 int m_daystokeepdata;
	 CString m_archivefolder;
};


