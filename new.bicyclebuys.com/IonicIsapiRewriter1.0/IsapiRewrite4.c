/*

IsapiRewrite4.c

Ionic's Isapi Rewrite Filter [IIRF]

ISAPI Filter that does  URL-rewriting. 
Inspired by Apache's mod_rewrite .
Implemented in C, does not use MFC. 

==================================================================


License
---------------------------------

Ionic's ISAPI Rewrite Filter is an add-on to IIS that can
rewrite URLs.  IIRF and its documentation is distributed under
the license terms spelled out here, terms derived from the BSD license. 

Written by: Ionic Shade
dpchiesa [AT] hotmail.com 
Copyright (c) 2005  Ionic Shade
All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

    * Redistributions of source code must retain the above copyright notice,
      this list of conditions and the following disclaimer.

    * Redistributions in binary form must reproduce the above copyright
      notice, this list of conditions and the following disclaimer in the
      documentation and/or other materials provided with the distribution.

    * Neither the name of Ionic Shade nor the names of any
      contributors to this project may be used to endorse or
      promote products derived from this software without
      specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE
LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
POSSIBILITY OF SUCH DAMAGE.



==================================================================
related: 

 http://www.phys-iasi.ro/Library/Computing/Using_ISAPI/index14.htm
 http://www.codeproject.com/isapi/isapiredirector.asp
 http://support.zeus.com/doc/examples/isapi/lang.html  AllocMem
 http://www.alphasierrapapa.com/IisDev/Articles/XAspFilter/
 
 http://msdn.microsoft.com/library/en-us/iissdk/iis/redirecting_in_an_isapi_filter_using_sf_notify_send_raw_data.asp
 
 dependencies: 
  AdvApi32.lib - for RegKeyOpenEx, etc
  PCRE - the Perl-compatible Regular Expression library, from Hazel.  this is for pattern matching.
 
 to build:
   (see makefile)
 
 (c) Copyright Ionic Shade, 2005
 Tue, 04 Jan 2005  17:14


*/


#include <stdio.h>
#include <time.h>

#include <WTypes.h>
#include <HttpFilt.h>
#include <WinInet.h>
#include <WinReg.h>  // for registry stuff

#include <pcre.h>

#include "RewriteRule.h"


#define ISAPI_FILTER_VERSION_STRING "Ionic URL Rewriting ISAPI Filter v1.0.1"
#define INI_FILENAME "Isapi-Rewrite.ini"
#define REGY_KEY     "Software\\Ionic\\Isapi Rewriter 1.0"
#define DEFAULT_ITERATION_LIMIT 8
#define DEFAULT_MATCH_COUNT 10



// statics and globals
volatile BOOL     FilterInitialized = FALSE;
CRITICAL_SECTION  g_CS;
static  int    g_LogLevel= 0;
char LogFileName[_MAX_PATH];
static 	RewriteRule * root;
static  char IniFileName[_MAX_PATH];
static int IterationLimit= DEFAULT_ITERATION_LIMIT;
static int MaxMatchCount= DEFAULT_MATCH_COUNT;



// forward decls
void LogMe(int level, char *p); 
void LogMeNoTimestamp(char *p) ;
VOID Initialize() ;
void ReadConfig(char * ConfigFile) ;
boolean ApplyRules(char * subject, int depth, /* out */ char **result);
void ReadRegistry () ;



BOOL WINAPI GetFilterVersion( PHTTP_FILTER_VERSION pVer ) 
     /* 
Purpose:

    Required entry point for ISAPI filters.  This function
    is called when the server initially loads this DLL.

Arguments:

    pVer - Points to the filter version info structure

Returns:

    TRUE on successful initialization
    FALSE on initialization failure

     */

{
    LogMe(1, "GetFilterVersion");

    //    if(!IsBadWritePtr(pVer,sizeof(HTTP_FILTER_VERSION))) 
    {
	pVer->dwFilterVersion = HTTP_FILTER_REVISION;

	// filter priority
	pVer->dwFlags |=   SF_NOTIFY_ORDER_LOW ;

	// security
	pVer->dwFlags |=  SF_NOTIFY_SECURE_PORT | SF_NOTIFY_NONSECURE_PORT ;

	// notification flags
	pVer->dwFlags |= SF_NOTIFY_PREPROC_HEADERS ;

	strncpy(pVer->lpszFilterDesc, ISAPI_FILTER_VERSION_STRING, sizeof(pVer->lpszFilterDesc));
	return TRUE; 
    }
    //    else {
    //	return FALSE; 
    //    }

}


DWORD PreProcessHeaders(PHTTP_FILTER_CONTEXT pFilterContext, PHTTP_FILTER_PREPROC_HEADERS pHeaderInfo) 
{

    char url[INTERNET_MAX_URL_LENGTH];
    DWORD dwSize= sizeof(url);
    if (!pHeaderInfo->GetHeader(pFilterContext, "url", url, &dwSize)) {
        return SF_STATUS_REQ_ERROR;
    } 

    LogMe(1, "PreProcessHeaders");
    
    if (url[0]!='\0') {
	boolean rc;
	char * resultString;

	char buf[INTERNET_MAX_URL_LENGTH + 32];
	sprintf(buf, "Url is: '%s'", url);
	LogMe(1, buf);

	// TODO: Load new initialization for each new vdir ?

	rc= ApplyRules(url, 0, &resultString);
	if (rc) {
	    sprintf(buf, "Rewrite Url to: '%s'", resultString);
	    LogMe(1, buf);
	    pHeaderInfo->SetHeader(pFilterContext, "url", resultString); 
	    free(resultString);   // or not? 
	    //return SF_STATUS_REQ_HANDLED_NOTIFICATION; // need this?? 
	}
	else {
	    LogMe(1, "No Rewrite");
	}
    }
    
    return SF_STATUS_REQ_NEXT_NOTIFICATION;
}



/* extern "C" */

DWORD WINAPI 
HttpFilterProc(
    HTTP_FILTER_CONTEXT *      pfc,
    DWORD                      dwNotificationType,
    VOID *                     pvNotification
    )
{

    DWORD dwRetval = SF_STATUS_REQ_NEXT_NOTIFICATION;

    LogMe(1, "HttpFilterProc");
    if (dwNotificationType == SF_NOTIFY_PREPROC_HEADERS) {
	if ( ! FilterInitialized ) {
	    Initialize();
	}

	dwRetval = PreProcessHeaders(pfc, (PHTTP_FILTER_PREPROC_HEADERS) pvNotification);
    }

    return dwRetval;
}





BOOL WINAPI TerminateFilter(DWORD dwFlags) {
    /* free / unload / unlock any allocated/loaded/locked resources */

    return TRUE;
}



static HINSTANCE g_hInstance = NULL;

HINSTANCE __stdcall AfxGetResourceHandle()
{ 
    LogMe(1, "Terminate");
    return g_hInstance;
}




BOOL WINAPI DllMain(HINSTANCE hInst, ULONG ulReason, LPVOID lpReserved) 
{
    char szLastAd[4];
    char drive[_MAX_DRIVE];
    char dir[_MAX_DIR];
    char fname[_MAX_FNAME];
    char ModuleFileName[_MAX_PATH];

    boolean retVal= FALSE;

    switch( ulReason ) {

    case DLL_PROCESS_ATTACH: 
	{
	    // on process attach we can initialize the state of the filter. 
	    ReadRegistry();
	    InitializeCriticalSection(&g_CS);

	    LogMe(1, "DllMain (PROCESS_ATTACH)");

	    if (GetModuleFileName(hInst, ModuleFileName, sizeof(ModuleFileName))) {
		char msg[_MAX_PATH+32];
		_splitpath(ModuleFileName, drive, dir, fname, NULL);
		_makepath(IniFileName, drive, dir, fname, ".ini");
		sprintf(msg,"target ini file: '%s'", IniFileName);
		LogMe(1, msg);

		retVal= TRUE; 
	    }
	    else 
		LogMe(1, "Cannot get module name");


	    /* giCurrentAd = GetPrivateProfileInt(TEXT("Info"), TEXT("LastAd"), 1, TEXT(INI_FILENAME)); */
	    break;
	}

    case DLL_PROCESS_DETACH:
	{
	    /* WritePrivateProfileString(TEXT("Info"), TEXT("LastAd"), szLastAd, TEXT(INI_FILENAME)); */
	    DeleteCriticalSection(&g_CS);
	    LogMe(1, "DllMain (PROCESS_DETACH)");
	    break;
	}

	// case DLL_THREAD_ATTACH:
	// case DLL_THREAD_DETACH:

    }
    return retVal;
}


// util routines
static boolean FirstLog= TRUE;

void LogMe(int msgLevel, char *s) {

    //   printf("FILE[%s] Lev[%d,%d]  %s\n",LogFileName, g_LogLevel, msgLevel, s);

    if ( ( g_LogLevel >= msgLevel ) && (LogFileName[0]!='\0')) {
	
	FILE *fp=fopen(LogFileName,"a+");
	if (fp==NULL) 
	    fp=fopen(LogFileName,"w");

	if (fp!=NULL) {
	    time_t t;
	    char tmbuf[25] ;
	    time(&t);
	    strncpy(tmbuf,ctime(&t),24);
	    tmbuf[24]= '\0';
	    if (FirstLog) {
		fprintf(fp,"\n--------------------------------------------\n");
		FirstLog=FALSE;
	    }
	    fprintf(fp,"%s - %s\n", tmbuf,s);
	    fclose(fp);
	}
    }
}


#define RE_SIZE 1024

VOID Initialize() 
{
    EnterCriticalSection(&g_CS);
    if ( ! FilterInitialized ) {
	LogMe(1, ISAPI_FILTER_VERSION_STRING);
	LogMe(1, "Initialize");

	// TODO: create thread to watch ini file for changes and re-init
	ReadConfig(IniFileName);
	
	FilterInitialized= TRUE;
    }
    LeaveCriticalSection(&g_CS);
}




boolean IsDuplicate(char * pattern) 
{
    RewriteRule * current= root;
    boolean retVal= FALSE;
    while (current!=NULL) {
	if (strcmp(current->Pattern, pattern)==0) {
	    retVal= TRUE;
	    current=NULL; // break;
	}
	else current= current->next;
    }
    return retVal;
}


// directives:

#define REWRITE_RULE "ReWriteRule"
#define ITERATION_LIMIT "IterationLimit"
#define MAX_MATCH_COUNT "MaxMatchCount"

// Directives still needing implementation: 
#define REWRITE_COND "ReWriteCond"
#define PROXY_PASS   "ProxyPass"
#define PROXY_PASS_REVERSE   "ProxyPassReverse"
#define RECEIVE_BUFFER_SIZE "ProxyReceiveBufferSize"
#define IO_BUF_SIZE  "ProxyIOBufferSize"


#define RE_SIZE 1024

void ReadConfig(char * ConfigFile) {
    char logMsg[RE_SIZE*3];
    const char *error;
    int erroffset;
    BOOL done= FALSE;
    FILE *infile ;
    unsigned char *p1;
    unsigned char *p2;
    pcre * re;
    //RewriteRule * root= NULL;
    RewriteRule * current= NULL;
    RewriteRule * previous;
    int lineNum=0; 
    int nRules=0;
    int nFailed=0;
    char delims[]= " \n\r\t";

    unsigned char * buffer;

    LogMe(1, "ReadConfig");

    buffer = (unsigned char *)malloc(RE_SIZE);
    if (buffer==NULL) return ;

    /* read config file here, slurping in Rewrite rules */ 

    infile= fopen(ConfigFile, "r");
    if (infile==NULL) return;

    while (!done) {
	lineNum++;
	if (fgets((char *)buffer, RE_SIZE, infile) == NULL) break;

	p1 = buffer;
	while (isspace(*p1)) p1++;
	if (*p1 == 0) continue; // nothing
	if (*p1 == '#') continue; // comment
	p2= strtok(p1, " ");
	if (strnicmp(p2,REWRITE_RULE, strlen(REWRITE_RULE))==0) {
	    char *pPattern = strtok (NULL, delims);
	    char *pReplacement = strtok (NULL, delims);
	    char *pModifiers = strtok (NULL, delims);

	    sprintf(logMsg,"ini line %3d: RewriteRule %3d %-46s %-42s %8s", 
		    lineNum, nRules+1, pPattern, pReplacement, pModifiers );
	    LogMe(1, logMsg);

	    if (IsDuplicate(pPattern)) {
		sprintf(logMsg,"ini file line %d: duplicate expression '%s'", lineNum, pPattern);
		LogMe(1, logMsg);
		continue; 
	    }
	    
	    // Want to create a compiled regex here

	    re = pcre_compile(
			      pPattern,              /* the pattern */
			      0,                    /* default options */
			      &error,               /* for error message */
			      &erroffset,           /* for error offset */
			      NULL);                /* use default character tables */

	    nRules++;
	    if (re == NULL) {
		sprintf(logMsg,"compilation of expression '%s' failed at offset %d: %s", pPattern, erroffset, error);
		LogMe(1, logMsg);
		nFailed++;
	    }
	    
	    previous= current;

	    current= (RewriteRule *) malloc(sizeof(RewriteRule)); 
	    if (root==NULL) root= current; 

	    current->RE= re; //  NULL or not
	    current->Pattern= (char*) malloc(strlen(pPattern)+1); 
	    strcpy(current->Pattern, pPattern);
	    current->Replacement= (char*) malloc(strlen(pReplacement)+1); 
	    strcpy(current->Replacement, pReplacement);
	    current->next= NULL;
	    current->previous= previous;  //NULL for first node, not-NULL for successive

	    if (previous!=NULL) 
		previous->next= current; 
	}

	else if (strnicmp(p2,ITERATION_LIMIT, strlen(ITERATION_LIMIT))==0) {
	    char *pLimit = strtok (NULL, delims);
	    if (pLimit!=NULL) IterationLimit= atoi(pLimit);
	    else IterationLimit=DEFAULT_ITERATION_LIMIT;
	    sprintf(logMsg,"setting Iteration Limit to %d", IterationLimit);
	    LogMe(1, logMsg);
	}

	else if (strnicmp(p2,MAX_MATCH_COUNT, strlen(MAX_MATCH_COUNT))==0) {
	    char *pCount = strtok (NULL, delims);
	    if (pCount!=NULL) MaxMatchCount= atoi(pCount);
	    else MaxMatchCount=DEFAULT_MATCH_COUNT;
	    sprintf(logMsg,"setting MaxMatchCount to %d", MaxMatchCount);
	    LogMe(1, logMsg);
	}

	else {
	    sprintf(logMsg,"ini file line %d: Ignoring line: '%s'", lineNum, p2);
	    LogMe(1, logMsg);
	}
    }

    sprintf(logMsg,"Found %d rules (%d failed) on %d lines", nRules, nFailed, lineNum);
    LogMe(1, logMsg);

    fclose(infile);
    free(buffer);
    return ;
}



char * GenerateReplacementString(char *src, char *ReplacePattern, int* vec, int count) 
{
    char logMsg[256];
    char *p1= ReplacePattern; 
    char *outString= malloc(512); 
    char *pOut= outString; 
    boolean done= FALSE;

    while (*p1!='\0') {
	if ((p1[0]=='$') && ( isdigit(p1[1]) )) {
	    int n= atoi(p1+1);
	    //sprintf(logMsg,"replacing substring %d", n);
	    //LogMe(logMsg);
	    if (n < count) {
 		char *substring_start = src + vec[2*n];
 		int substring_length = vec[2*n+1] - vec[2*n]; 
		strncpy(pOut,substring_start, substring_length);
		pOut+= substring_length;
	    }
	    *p1++;
	}
	else {
	    // sprintf(logMsg,"pass-thru: %c", *p1);
	    //LogMe(logMsg);

	    *pOut= *p1;
	    *pOut++;
	}
	*p1++;
    }
    *pOut='\0';
    return outString;
}



/* EXPORT */
boolean ApplyRules( char * subject, int depth, /* out */ char **result) 
{
    RewriteRule * current= root;
    boolean retVal= FALSE;
    char logMsg[512];
    int c=0;
    int rc, i; 
    int *ovector;

    sprintf(logMsg,"ApplyRules (depth=%d)", depth);
    LogMe(3, logMsg);

    // pcre doc says should be a multiple of 3??  why? seems like it ought to be 2. 
    ovector= (int *) malloc((MaxMatchCount*3)*sizeof(int));  

    if (root==NULL) return retVal;

    // TODO: employ a MRU cache to map URLs

    while (current!=NULL) {
	c++;
	rc = pcre_exec(
		       current->RE,          /* the compiled pattern */
		       NULL,                 /* no extra data - we didn't study the pattern */
		       subject,              /* the subject string */
		       strlen(subject),      /* the length of the subject */
		       0,                    /* start at offset 0 in the subject */
		       0,                    /* default options */
		       ovector,              /* output vector for substring information */
		       MaxMatchCount*3);     /* number of elements in the output vector */
	
	if (rc < 0) {
	    if (rc== PCRE_ERROR_NOMATCH) {
		sprintf(logMsg,"Rule %d : %d", c, rc );
		LogMe(3, logMsg);
	    }
	    else {
		sprintf(logMsg,"Rule %d : %d (unknown error)", c, rc);
		LogMe(2, logMsg);
	    }
	}
	else if (rc == 0) {
	    sprintf(logMsg,"Rule %d : %d (The output vector (%d slots) was not large enough)", 
		    c, rc, MaxMatchCount*3);
	    LogMe(2, logMsg);
	}
	else {
	    char * newString;
	    sprintf(logMsg,"Rule %d : %d", c, rc);
	    LogMe(2, logMsg);
	    retVal= TRUE;

	    /* generate and Emit the replacement string */
	    newString= GenerateReplacementString(subject, current->Replacement, ovector, rc);

	    if (sizeof(logMsg)-11> strlen(newString)) { 
		sprintf(logMsg,"Result: %s", newString);
		LogMe(2,logMsg);
	    }
	    else
		LogMe(2,"(Log Buffer overflow)");
	    sprintf(logMsg,"Result length: %d", strlen(newString));
	    LogMe(3,logMsg);

	    *result= newString; 

	    if (depth < IterationLimit) {
		char * t; 
		boolean rc= ApplyRules(newString, depth+1, &t);
		if (rc) { 
		    *result= t;
		    free(newString);
		}
	    }
	    else {
		sprintf(logMsg,"Iteration stopped; reached limit of %d cycles.", IterationLimit);
		LogMe(2,logMsg);
	    }

	    break;  // stop on first match
	}

	current= current->next;
    }
    return retVal;
}



void ReadRegistry () {
    long rc;
    HKEY hkey;
    DWORD type = 0;
    DWORD size;
    LONG lrc;
    char buf1[INTERNET_MAX_URL_LENGTH];
    char buf2[INTERNET_MAX_URL_LENGTH];

    LogMe(1, "ReadRegistry");
    LogFileName[0]='\0';

    rc = RegOpenKeyEx(HKEY_LOCAL_MACHINE,
		      REGY_KEY, (DWORD) 0, KEY_READ, &hkey);
    if (ERROR_SUCCESS != rc) return ; 

    size= sizeof(buf1);
    LogMe(1, "RegistryRead: Reading LogLevel");
    rc = RegQueryValueEx(hkey, "LogLevel", (LPDWORD) 0, &type, (LPBYTE) buf1, &size);
    if ((ERROR_SUCCESS == rc) || (type == REG_SZ)) {
	buf1[size] = '\0';  // want this? 
	g_LogLevel= atoi(buf1);
	sprintf(buf2,"LogLevel= %d", g_LogLevel);
	LogMe(1, buf2);
    }


    LogMe(1, "RegistryRead: Reading LogFile");
    size= sizeof(buf1);
    rc = RegQueryValueEx(hkey, "LogFile", (LPDWORD) 0, &type, (LPBYTE) buf1, &size);
    if ((ERROR_SUCCESS == rc) || (type == REG_SZ)) {
	buf1[size] = '\0';
	strcpy(LogFileName, buf1);
	sprintf(buf2, "Logging at level=%d, file='%s'", g_LogLevel, LogFileName);
	LogMe(1, buf2); // should be first actual log msg !
    }

    RegCloseKey(hkey);
}





/* extern "C" */
void IsapiFilterTestSetup(char * psz_iniFileName) 
{
    // force logging to console
    strcpy(LogFileName, "CON"); 
    Initialize(); 
}

