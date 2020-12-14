Memory For Macros dictionary file tips.

You can comment lines in files using the ; character.
All comment lines found at the top of the file until the first non-commented line will be retained when saving a dictionary.  All other lines in the dictionary file will be rewritten when saving using the voice macros.  Dictionary files are only saved when calling macro code that explicitly saves them.

Blank lines are ignored.

Dictionary name must be included in brackets like this:  [<dictionary name>] e.g. [people]
Entries will be loaded into a dictionary keyed by this name.  Dictionary names across all types of data must be unique.  You may include multiple dictionaries per file.

Format dictionary entry lines like this: <key>=<value1>,<value2>,...
e.g. mj=Michael Jackson,Michael Jordan,Mark Jones

Entry values can include references to other entry values.  For example, given this value in a dictionary file:

	documents=C:\Users\Doug\Documents

you can refer to it in other entries like this:

 	personal=${documents}\personal

If the referred entry is found in a different dictionary, specify the dictionary as follows:

	my personal file=${paths:personal}\my file.txt

If the referred entry includes more than one value, the first value is used.

NOTE: when loading dictionaries, make sure to load referred dictionaries first so that the references are valid. 

Example:

[paths]
documents=C:\Users\Doug\Documents
windows=C:\Windows
personal=${documents}\personal
knowbrainer commands=C:\Users\Doug\AppData\Roaming\KnowBrainer\KnowBrainerCommands 

[files]
mug shot=${paths:personal}\mugshot.jpg
knowbrainer command file=${paths:knowbrainer commands}\MyCommands.xml

