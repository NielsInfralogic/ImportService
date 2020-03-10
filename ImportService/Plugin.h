#pragma once

typedef int(*PLUGININIT)();
typedef int(*PLUGINEXIT)();
typedef int(*PLUGINSETUP)();
typedef int(*PLUGINJOB)(char *szFilePath, char *szErrorMessage);
typedef int(*PLUGINJOBEX)(int nSpoolNumber, char *szFilePath, char *szErrorMessage);
typedef int(*PLUGINJOBSTATUS)(int nStatus, char *szXMLFile, char *szErrorMessage);
