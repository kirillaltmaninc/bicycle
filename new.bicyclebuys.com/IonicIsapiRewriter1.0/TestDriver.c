/*
 * TestDriver.c
 *
 * Test driver module for the Isapi Rewrite filter.
 * This module links to the ISAPI filter DLL, and drives it.  
 * This permits testing of the RewriteRules against a specific set of URLs, 
 * from the command line, without needing to configure the ISAPI filter to IIS. 
 * Output is sent to the console.  
 *
 * Use this app to verify that the rewrite rules you have authored are working as desired. 
 *
 * (c) Copyright Ionic Shade, 2005
 * Tue, 04 Jan 2005  17:13
 *
 */

#include <stdio.h>
#include <string.h>
#include <stdlib.h>
#include <time.h>

#include <WTypes.h>

#include <pcre.h>

#include "RewriteRule.h"


void ProcessUrls(char * SampleUrlsFile) 
{
    FILE *infile; 
    int lineNum=0; 
    char url[1024];
    //char logMsg[256];
    char *p1;
    char * resultString;
    boolean rc;
    int len;

    printf("Processing URLs...\n\n");

    infile= fopen(SampleUrlsFile, "r");
    if (infile==NULL) {
	printf("Cannot open Urls file\n");
	return ;
    }

    while (TRUE) {
	lineNum++;
	if (fgets((char *)url, sizeof(url), infile) == NULL) break;
	p1 = url;
	while (isspace(*p1)) p1++;
	if (p1[0]=='\0') continue;  // empty line
	len= strlen(p1);
	while ((len>1) && (isspace(p1[len-1]))) {
	    p1[len-1]='\0';
	    len= strlen(p1);
	}

	rc= ApplyRules(p1, 0, &resultString);
	if (rc) {
	    printf("\n'%s' ==> '%s'\n\n", p1, resultString);
	    free(resultString); 
	}
	else 
	    printf("\n'%s' ==> No Rewrite\n\n", p1);

    }
    fclose(infile);
}




int main(int argc, char **argv)
{
    IsapiFilterTestSetup("IsapiRewrite4.ini");

    ProcessUrls("SampleUrls.txt");
}
