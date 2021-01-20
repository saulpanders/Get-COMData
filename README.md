# Get-COMData
Enumerates all of the COM Objects & their exposed methods on your local machine, to hunt for unusual COM goodies.

Pulls all of the subkeys from HKLM\Software\Classes registry key, then instantiates the associated COM object by CLSID & enumerates available methods

## Usage
To use this script, clone this repo to your local system, then `Import-Module .\Get-COMData

To run:
```
PS C:\>  Get-COMData -ScriptUpdates -CLSIDFileName clsids.txt -OutputFileName COM-objects-methods.txt
```
#### OPSEC Considerations
* Currently, this script drops 2 files to disk
* Will cause some strange behavior (pop-ups etc) on the local machine during execution
### Arguments
There are two mandatory arguments:
* CLSIDFileName
* OutputFileName
And one optional:
* ScriptUpdates 

#### CLSIDFileName
Name of file which stores CLSID list

#### OutputFileName
Name of full output containing CLSID & its exposed methods

#### ScriptUpdates
Acts as a boolean switch for printing verbose command-line status updates


#### Source(s):
https://www.fireeye.com/blog/threat-research/2019/06/hunting-com-objects.html
