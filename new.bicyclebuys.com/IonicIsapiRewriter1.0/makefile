# makefile
#
# for IIRF,  Iconic's Isapi Rewriting Filter.
#
# Thu, 06 Jan 2005  10:10
#


VC7="C:\Program Files\Microsoft Visual Studio 9.0\VC"
PSDK=$(VC7)\PlatformSDK
CC=$(VC7)\bin\cl.exe
LINK=$(VC7)\bin\link.exe

PCRE="D:\Root\new.bicyclebuys.com\IonicIsapiRewriter1.0"

#ADVAPI32=C:\Program Files\Microsoft SDKs\Windows\v6.0A\Lib\advapi32.dll
ADVAPI32="C:\Program Files\Microsoft SDKs\Windows\v6.0A\Lib\AdvAPI32.Lib"

INSTALL_DIR=c:\Windows\system32\inetsrv

ZIPIT=c:\dinoch\bin\backdir.bat 

# =======================================================

IsapiRewrite4.dll: IsapiRewrite4.c RewriteRule.h makefile
	$(CC) -I$(VC7)\Include -I$(PSDK)\Include  -I$(PCRE)   -LDd IsapiRewrite4.c  $(PCRE)\pcre.lib  $(ADVAPI32)  -FeIsapiRewrite4.dll  -link /LIBPATH:$(VC7)\Lib /LIBPATH:$(PSDK)\Lib    /SUBSYSTEM:CONSOLE /EXPORT:GetFilterVersion /EXPORT:HttpFilterProc /EXPORT:IsapiFilterTestSetup /EXPORT:TerminateFilter /EXPORT:ApplyRules /DEBUG /NODEFAULTLIB:libc


TestDriver.exe: TestDriver.c RewriteRule.h IsapiRewrite4.dll
	$(CC) /c -I$(VC7)\Include -I$(PSDK)\Include -I$(PCRE) TestDriver.c 
	$(LINK)  TestDriver.obj IsapiRewrite4.lib $(PCRE)\pcre.lib  /LIBPATH:$(VC7)\Lib /LIBPATH:$(PSDK)\Lib   /SUBSYSTEM:CONSOLE  /DEBUG



install: $(INSTALL_DIR)\IsapiRewrite4.dll $(INSTALL_DIR)\IsapiRewrite4.ini


$(INSTALL_DIR)\IsapiRewrite4.dll: IsapiRewrite4.dll
	copy /y  IsapiRewrite4.dll $(INSTALL_DIR)

$(INSTALL_DIR)\IsapiRewrite4.ini: IsapiRewrite4.ini
	copy /y  IsapiRewrite4.ini $(INSTALL_DIR)\IsapiRewrite4.ini



# Rule for creating ZIP file
# Note: this will not work for you if you do not have the ZIPIT
# tool
zip: IsapiRewrite4.ini IsapiRewrite4.dll Readme.txt makefile IirfLoggingOn.reg IirfLoggingOff.reg TestDriver.exe SampleUrls.txt License.txt filtered.htm
	mkdir IonicIsapiRewriter1.0
	copy makefile IonicIsapiRewriter1.0
	copy Readme.txt IonicIsapiRewriter1.0
	copy License.txt IonicIsapiRewriter1.0
	copy TestDriver.c IonicIsapiRewriter1.0
	copy TestDriver.exe IonicIsapiRewriter1.0
	copy RewriteRule.h IonicIsapiRewriter1.0
	copy IirfLoggingOn.Reg IonicIsapiRewriter1.0
	copy IirfLoggingOff.Reg IonicIsapiRewriter1.0
	copy IsapiRewrite4.ini IonicIsapiRewriter1.0
	copy SampleUrls.txt IonicIsapiRewriter1.0
	copy IsapiRewrite4.dll IonicIsapiRewriter1.0
	copy IsapiRewrite4.c IonicIsapiRewriter1.0
	copy $(PCRE)\pcre.lib IonicIsapiRewriter1.0
	copy $(PCRE)\pcre.h IonicIsapiRewriter1.0
	copy $(PCRE)\Makefile IonicIsapiRewriter1.0\PCRE.mak
	copy $(PCRE)\LICENCE IonicIsapiRewriter1.0\License.pcre.txt

	$(ZIPIT) IonicIsapiRewriter1.0 IonicIsapiRewriter.zip 
	rd /s /q IonicIsapiRewriter1.0


clean: 
	-del IsapiRewrite4.dll
	-del IonicIsapiRewriter.zip 
	-del IsapiRewrite4.obj
	-del IsapiRewrite4.pdb
	-del IsapiRewrite4.lib
	-del IsapiRewrite4.exp
	-del IsapiRewrite4.ilk
	-del TestDriver.exe
	-del TestDriver.obj
	-del TestDriver.ilk
	-del TestDriver.pdb
