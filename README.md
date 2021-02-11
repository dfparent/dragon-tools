# dragon-tools
Macros and tools for use with Dragon Naturallyspeaking 15 and KnowBrainer 2017

<h1>Overview</h1>
<p>
This is a collection of macros and tools that make it easier to work hands free on a Windows PC using Dragon Naturallyspeaking 15.  To use these tools you also need to purchase and install a Dragon add-on called KnowBrainer Professional 2017.
</p>
<p>
KnowBrainer Professionial comes with its own collection of macros that you can immediately use.  My collection of tools takes it to the "next level".  They crank it up to "11" if you know what I mean.  This collecton consists of a bunch of macros that I have found useful as well as some applications/tools created to work completely with keystrokes or via an API that can be called from macro code.  These tools in conjunction with the macros that work with them provide some very powerful functionality that attempts to "fill in the blanks" left by Dragon and KnowBrainer.
</p>
<p>
The downside:  these tools are free and as such there is no technical support available.  I have tried my best to make them universally applicable, but I have only used them on my personal and work PCs.  If you experience any problems setting things up, please let me know so we can make the install instructions better.
</p>
<p>
Here are some functionality highlights to pique your interest:
<ul>
	<li>Touch Locations</li>
		<P>Define named click location by voice.  Assign a name to a mouse location using voice commands and use the name to click there.  For example, you could position your mouse over the Browser Back button and call it "Back".  Then you could say "touch back" and it will perform a mouse click at that location.  Locations are saved between dictation sessions and you can have different loctions for each application.</p>
	<li>Click By Numbers</li>
		<p>This is functionality similar to other applications you can buy that show a number for each control in an window.  Say "flag" followed by the number and it will click at the center of that control.  You can choose to show flags by default (or not) for a particular application or choose to make them "sticky" or not.</p>
	<li>Mouse Grid</li>
		<p>Superimposes a grid on the current window.  You can say the row number and column number to click in a particular square.  Have a grid cover the whole window, use the default "top", "left", "right", or "bottom" grids, or setup your own custom grid size and location in a macro.</p>
	<li>Memory For Macros</li>
		<p>A utility service that enables many of the macros in my macro set.  Lets you remember values in between macro runs, save values to a file or registry, run a series of "delayed" commands, access saved paths, files, people, snippets, touch locations and other data.</p>
	<li>Show Active Window</li>
		<p>Shows a red box around the foreground window to overcome Window 10's terrible color scheme that makes it impossible sometimes to know what the active window is.  I mean, what were they thinking?</p>
	<li>Kill Dragon</li>
		<p>When Dragon freezes or crashes, (not if, but when) this will kill all remaining Dragon processes so you can start it again without a reboot or logout.</p>
</ul>
</p>

<h1>Installation Instructions</h1>
<p>
<ol>
	<li>Install Dragon Naturallyspeaking 15 using the default folder locations.</li>
	<li>Install Knowbrainer Professional 2017 using the default folder locations.</li>
	<li>Create C:\Users\KnowBrainer\CommonModules.  You will need administrator privileges to do so.  Grant either "Everyone" or your user full control of this folder if your user is not already an administrator.</li>
	<li>Copy the entire contents of CommonModules into C:\Users\KnowBrainer\CommonModules.</li>
	<li>Rename the following files in C:\Users\KnowBrainer\CommonModules\Data by replacing "MyComputer" with your computer name.  Also, update the contents of these files to match what is on your computer:
	<ul>
		<li>app-MyComputer.txt </li>
		<li>paths-MyComputer.txt </li>
	</ul>
	</li>
	<li>Update the contents of these files in C:\Users\KnowBrainer\CommonModules\Data to match what is on your computer:
	<ul>
		<li>paths.txt</li>
	</ul>
	</li>
	<li>Update C:\Users\KnowBrainer\CommonModules\Data\people.txt to contain the names of the people you communicate with the most.  Follow the instructions contained in this file.</li>
	<li>Register the Memory For Macros tools by running the C:\Users\KnowBrainer\CommonModules\MemoryForMacros\RegisterMemoryForMacros.bat file as an administrator from within the same folder as the bat file.  You may need to update the batch file with the correct path to your .NET Framework install.</li>
	<li>If you wish, you can also add the Office Add-ins to your Office installation.  These will run when your apps startup and add commands to the Add-Ins Ribbon.  From this ribbon you can run various helpful commands.  There are some voice macros that make use of the functionality in these Office Add-ins so if you want to use those voice macros, you'll need to install the add-ins.
		<ol>
			<li>Copy MacrosForVoice.ppam and PowerpointUtilities.ppam to C:\Users\&lt;user folder&gt;\AppData\Roaming\Microsoft\AddIns.  Start Powerpoint and add the addins in Powerpoint Options.</li>
			<li>Copy MyUtilities.xlam to C:\Users\&lt;user folder&gt;\AppData\Roaming\Microsoft\AddIns.  Start Excel and add the addins in Excel Options.</li>
			<li>Copy WordUtilities.dotm to C:\Users\&lt;user folder&gt;\AppData\Roaming\Microsoft\Word\STARTUP.  The template file should be automaticallly picked up and run when Words starts.</li>
		</ol>
	</li>
</ol>
</p>
