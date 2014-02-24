'Create an instance of IE
Dim IE
Set IE = CreateObject("InternetExplorer.Application")

'Execute our URL
ie.navigate("http://www.bicyclebuys.com/encrypt/auto_encrypt.asp")

'Clean up...
Set IE = Nothing