# win32 makefile for nmake
#
# for PCRE, the Perl-Compatible regular expression library
# See http://www.pcre.org
#
# Disclaimer: 
# this worked for me, may not work for you! 
# Works with PCRE v5.0 . 
#
# To use this thing:
# - download PCRE v5.0 (or later?) 
# - copy this makefile into the PCRE dir. 
# - modify the settings for VC7 and PSDK
# - run nmake!
# 
# =======================================================
# Ionic Shade
# Thu, 06 Jan 2005  10:10
#


VC7="C:\Program Files\Microsoft Visual Studio 9.0\VC"
PSDK=$(VC7)\PlatformSDK
CC=$(VC7)\bin\cl.exe
LIB=$(VC7)\bin\lib.exe

# =======================================================

all: pcre pcretest

pcre: pcre.dll pcre.lib pcreposix.lib

pcretest: pcretest.exe

dftables.exe: dftables.c
	$(CC) -DSUPPORT_UTF8 dftables.c -link /LIBPATH:$(VC7)\Lib /LIBPATH:$(PSDK)\Lib  

chartables.c: dftables.exe
	dftables.exe chartables.c

#static library
pcre.lib: chartables.c makefile
	$(CC)  -DSTATIC -DSUPPORT_UTF8 -DPOSIX_MALLOC_THRESHOLD=10 /c maketables.c get.c study.c pcre.c
	$(LIB) /OUT:pcre.lib maketables.obj get.obj study.obj pcre.obj /NODEFAULTLIB

pcreposix.lib: chartables.c makefile
	$(CC)  -DSTATIC  -DSUPPORT_UTF8 -DPOSIX_MALLOC_THRESHOLD=10 /c pcreposix.c
	$(LIB) /OUT:pcreposix.lib pcreposix.obj

#dynamic library
pcre.dll: chartables.c makefile
	$(CC) -DSUPPORT_UTF8 -DPOSIX_MALLOC_THRESHOLD=10 /LDd maketables.c get.c study.c pcre.c /Fepcre.dll  /link /implib:pcre-tmp.lib /LIBPATH:$(VC7)\Lib /LIBPATH:$(PSDK)\Lib  /SUBSYSTEM:CONSOLE  /DEBUG  /EXPORT:pcre_compile /EXPORT:pcre_fullinfo  /EXPORT:pcre_exec   /EXPORT:pcre_version    /EXPORT:pcre_config   


pcreposix.dll: chartables.c makefile
	$(CC) -DSUPPORT_UTF8 -DPOSIX_MALLOC_THRESHOLD=10 /LDd pcreposix.c /Fepcreposix.dll  /link /LIBPATH:$(VC7)\Lib /LIBPATH:$(PSDK)\Lib  /SUBSYSTEM:CONSOLE  /DEBUG 


pcretest.exe: pcretest.c pcre.lib pcreposix.lib
	$(CC) -DSUPPORT_UTF8 pcretest.c pcre.lib pcreposix.lib /link /LIBPATH:$(VC7)\Lib /LIBPATH:$(PSDK)\Lib  /SUBSYSTEM:CONSOLE  /DEBUG



# =======================================================

clean:
	-del pcre.dll
	-del pcre.lib
	-del pcreposix.lib
	-del pcreposix.dll
	-del pcretest.exe

