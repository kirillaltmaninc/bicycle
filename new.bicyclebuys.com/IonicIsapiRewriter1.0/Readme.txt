Tue, 04 Jan 2005  15:47

updated Mon, 18 Jul 2005  16:03


Ionic's Isapi Rewrite Filter (IIRF) 

The mod_rewrite for Apache is widely used to rewrite URL
requests before they reach web application engines.  IIS 5.0
(Windows 2000) 5.1 (Windows XP) and v6.0 (Windows 2003) lack a
built-in URL rewriting function.

There are several add-on tools that can add the URL re-writing
function to IIS. All are implemented as ISAPI Filters. 
  - ASP.NET provides some URL mapping capability
  - the IIS v6.0 Resource Kit includes the UrlRemap tool.
  - IISRewrite 
  - ISAPI_Rewrite - 
    includes a "lite" version available for free. 
  - Mod_Rewrite 2

All of them have drawbacks. 
The ASP.NET support requires that all URLs map through ASP.NET;
even static files would have to be served by the ASP.NET
runtime.  This can imply a large performance cost.  

The IIS6 Resource Kit filter is provided by Microsoft, and is
supported.  However, it does not allow specifying rules via
Regular expressions, and as a result it is not as flexible as
some would like.

The ISAPI_Rewrite module is commercially supported, and does
Regular Expressions, but is expensive. Likewise Mod_Rewrite and
IISRewrite.  There is a "Lite" version of ISAPI_Rewrite, but it
does not do Regular Expressions. Doh! 

IIRF adds one more option to the list. IIRF is an open-source
Rewrite filter, available in source or binary form at no
cost. It does Regular-Expression mapping based on the PCRE
library.  The source is available so you can modify and tweak
the filter if you want to, or need to.



Implementation
---------------------------------
This filter is implemented in C, and has been compiled with the
Microsoft VC7 compiler that shipped with Visual Studio 2003.

If you'd like to compile it yourself, you can use the compiler
that ships with Visual Studio 2003, or you can also use the free
version of the Microsoft C++ compiler and tools, called the
Microsoft Visual C++ toolkit.  It is available here:

  http://msdn.microsoft.com/visualc/vctoolkit2003/




Installation 
---------------------------------
Installation is simple, but manual.  You need to install the
ISAPI Filter, as with any ISAPI.  Do it through the IIS MMC
panel.  Add the filter to the entire web site, or to a
particular website.  

You can place the IsapiRewrite4.dll file anywhere it will be
accessible to IIS.  In Windows XP Pro, this might be
c:\windows\system32\inetsrv .   Or it could be in c:\wwwroot
somewhere.  It's up to you. 

On Windows Server 2003, the same thing applies.  If you have
multiple web sites on WS2003, and want multiple installations of
IIRF, then install the DLL and .ini file in separate places, and
configure the ISAPI filter separately for each website.  



Logging Settings
---------------------------------

You can turn on Logging in the filter, by setting a few fields
in the Windows registry.  A .reg file (IirfLoggingOn.reg) is
provided for your convenience.  Just run the .reg file, and it
should update the Registry appropriately.  

Logging settings take effect when the ISAPI is restarted, which
is to say, when IIS is restarted.  

Log levels:
  0 - no logging
  1 - a little logging
  2 - a little more
  3 - all (verbose) logging

The LogFile value should specify a fully-qualified filename.
The file will be created by the ISAPI filter when it handles its
first URL.  If the file path is not valid, or if the ISAPI does
not have permissions to write to the specified file path, then
no logging will be generated.  If you expect to see a logging
file and don't see one, check your paths and permissions.

Logging is relatively expensive.  Each logged statement opens a
file, seeks to the end, writes to the file, and closes the file.
For best performance, Turn logging down as low as possible (to
zero).  For best debugging or monitoring of URL re-writing, turn
logging up to 3.



Directives
---------------------------------

It is possible to specify settings for the ISAPI in an ini
file.  The location of the Ini file is in the same directory as
the DLL, with the same name, but with an extension of .ini.  

The format of the ini file is similar in philosophy to that of
the properties file used by Apache's mod_rewrite.  Because this
ISAPI is simpler in intent and execution than mod_rewrite, the
Ini file is correspondingly simpler. 

There are currently 3 directives supported for the ini file: 

  IterationLimit
  MaxMatchCount  
  RewriteRule

The directive names are case-sensitive.  Whitespace following
the directive name is not significant.

A sample ini file is included in this shipment. 



RewriteRule
-----------

With this directive, administrators can specify how to map and
transform incoming URL requests.  There are 2 arguments for the
directive:  a regular expression pattern, and a replacement
string. 

For example, the following rule will transform incoming URLs
that contain approot1 into URLs that contain vdir2

Example 1: 

  RewriteRule  ^/approot1/(.*).php  /vdir2/$1.aspx

With this rule, an incoming URL like this: 
   /approot1/subdir/page.php

will get mapped to this:
   /vdir2/subdir/page.aspx


The Regular Expression support is from PCRE.  It provides a
fully-powered regular expression library.  For example, with
IIRF, you can use rules to transform path elements into query
string params.

Example 2: 

   RewriteRule  ^/dinoch/album/([^/]+)/([^/]+).(jpg|JPG|PNG)   /chiesa/pics.aspx?d=$1&p=$2.$3


See the provided Ini file for more examples. 
For more information on regular expressions, see
http://www.google.com/search?hl=en&num=30&q=regular+expressions
or try a book, like this: 
http://www.amazon.com/exec/obidos/tg/detail/-/1565922573?v=glance



IterationLimit
---------------------------------

The IterationLimit ini-file directive specifies how many times the
rewrite filter will loop on a single URL.  After a URL has been
transformed successfully, the result will be run through the
Rewrite rules again, to be transformed again.  This continues
for a single URL, as many times as necessary, or until the 
IterationLimit is exceeded, whichever comes first.

The IterationLimit is included as a fail-safe mechanism.
Consider the simplest case: It is possible to create a rewrite
rule that generates an output that matches its own input
pattern.  The effect is a logical infinite loop.  It is also
possible to have more complex loops, for example the output of
one rule matches the input of another rule, and vice versa.  

With an infinite loop in your ini-file Rewrite rules, and
without an IterationLimit, the ISAPI filter would loop
infinitely, and resulting in a stack overflow, and a
denial-of-service in IIS.

The simple solution would be: don't design sets of rules that
induce loops.  But sometimes it can be difficult to determine if
a loop might occur. The IterationLimit removes the potential for
this infinite loop.

The default IterationLimit is 8.  This default applies if you do
not specify an IterationLimit in the ini file.  The default may
or may not be suitable for any particular deployment.  It is not
likely that you will need more iterations, but if you do, change
the limit with the IterationLimit directive.



MaxMatchCount
---------------------------------

The MaxMatchCount directive specifies the maximum number of
matches to collect for each match.  The example regular
expressions in this document have 1 or 2 matches. In more
complex scenarios, an RE might have 8 or 12 matches.  The
MaxMatchCount allows an admin to specify the upper limit for
this number.  The default is 10.



Unrecognized Directives
---------------------------------

Any unrecognized directives in the INI file will be tolerated.
So, be careful of spelling and case.  If you specify 

  REWriterule xyz abc

...the filter won't do anything. 

Any unrecognized directives will be logged, but not everyone
turns on logging, and not everyone who enables logging actually
examines their logs. So do be careful.  



Processing Rules
---------------------------------
Rules are processed in the order they appear in the file.  



Features Not Included
---------------------------------

mod_rewrite includes the ability to specify local or remote
forwarding.  This is specified with an optional, third argument
to the RewriteRule directive in the .properties file (for the
Apache case).  This is not supported by Ionic's ISAPI rewrite filter.

mod_rewrite also has a RewriteCond directive, which allows you
to apply a conditional to the succeeding RewriteRule.  This
feature is not included in the current version of IIRF. 



Building the Filter
---------------------------------

To build the filter, you need these pre-requisites: 

 - the MS VC++ toolkit, or Visual Studio 2003 or better.  
   This includes a make utility, nmake.exe . 

 - a PCRE statically linked library, pcre.lib 
   This is included in the IIRF distribution. 

After you arrange the pre-requisites, modify the makefile for
IIRF to specify the locations of PCRE and the PlatformSDK
libraries.  Run nmake.




Building PCRE
---------------------------------

IIRF includes a version of PCRE, so it is not required that you
build a PCRE library to use IIRF.  However, you may wish to
download and build PCRE yourself.

The build procedure for PCRE on Windows is not well documented,
but it isn't difficult. You can use one of the MS c++ compilers.

You need to create your own PCRE makefile.  
 


Compatibility
---------------------------------

This URL re-writer works with any web application logic,
including but not limited to:  Cold Fusion (CFM), Active Server
Pages (ASP), ASP.NET, JSP, as well as static files like CSS, JPG,
HTM, xml, and so on.  

Compatibility issues may arise with other ISAPI filters.  
Let me know if you run into problems.



Testing
---------------------------------

You can test your sets of rules against sample URLs by using the
included TestDriver program.  Source for that is also included.   The
TestDriver links with the ISAPI filter, and asks it to apply the
Rewrite rules to a given set of sample URLs.  The sample Urls
are obtained from SampleUrls.txt, and the rewrite rules are
obtained from the ini file, as normal.  (The ini file name
derives from the name of the ISAPI DLL; by default it works out
to be IsapiRewrite4.dll ).  

When you run the TestDriver.exe, the ISAPI reads the registry as
normal to determine the Logging level (0,1,2,3).  But, the
logfile is overridden - all logging info is sent to stdout, the
console.



License
---------------------------------

Ionic's ISAPI Rewrite Filter (IIRF) is an add-on to IIS that can
rewrite URLs.  IIRF and its documentation is distributed under
the license terms spelled out here.

Written by: Ionic Shade
dpchiesa [AT] hotmail.com 
Copyright (c) 2004  Ionic Shade
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


IIRF depends upon PCRE, which is licensed independently and
separately.  Consult the License.pcre.txt file for details.



Futures
---------------------------------
Depending on interest I could extend the filter pretty
easily. Top features I would consider: 

 - auto-reload of changed INI file (no IIS restart)
 - RewriteCond directive
 - testing on IIS6

But I welcome feedback on this.  Let me know if you've used it,
and how it went.  Let me know what else you would like to see. 

Another possibility is to start up an IIRF project on
sourceforge.  If there is interest in that, I'd be open to the
idea. 


Bugs
---------------------------------
 - Not tested on clustered IIS
 - Not tested with IIS6, or with application pools, etc
 - Requires restart of IIS in order to pick up changes in INI file
 - Performance has not been measured
 - There is no MRU cache employed for URL re-writing
 - Does not use the Windows Event Log
 - No rollover of text-based Log files
 - The filter silently tolerates syntax errors in the INI file. 
   This behavior should be settable. 
 - The Logging strategy is naive. 
 - Not under source control. 
 - No RewriteCond directive supported yet



Fixed Bugs
---------------------------------

v1.0.1
 - There was a buffer overflow in ApplyRules which occurred when
   the resulting (rewritten) URL was larger than the Log message
   buffer.  I inserted a check to protect against this. 




Related
---------------------------------

IIS 6 Resource Kit
http://www.microsoft.com/downloads/details.aspx?familyid=56fc92ee-a71a-4c73-b628-ade629c89499&displaylang=en

IISRewrite
http://www.qwerksoft.com/products/iisrewrite/ 

ISAPI_Rewrite
www.isapirewrite.com

Mod_Rewrite 2
http://www.iismods.com/



Thanks
-Dino

dpchiesa@hotmail.com
