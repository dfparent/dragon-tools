A few applications tend not to work well with MSAA callouts, even though that is the mode that will work best for the application. In order to get it working for these applications, you have to do the following:

1. Install the Windows SDK.  You only need the following components:
	a. Windows SDK signing tools for desktop apps
	B. Windows SDK for UWP managed apps
2. Run C:\Program Files (x86)\Windows Kits\10\bin\10.0.17763.0\x86\Inspect.exe
3. Under options, select "MSAA Mode"
4. You might want to have the app that is having problems open while you do this.

There is SOMETHING that this app is doing that click my numbers is not doing. Someday I will figure this out.
