# Class 3 of 2026 KMS Activation Service
[简体中文](/zh) | English
Provide KMS Activation Services for Microsoft Windows and Microsoft Office.
## Microsoft Windows Activation
```shell
slmgr /ipk WIN_GVLK # Install GVLK
slmgr /skms cn-cd-txy-1.starryfrp.com:52206 # or `162.14.112.182:52206`
slmgr /ato # Activate now
slmgr /dlv # Display detailed license information
```
Go [here](win-key.html) to view `WIN_GVLK`
### Other usage for Slmgr.vbs
See [learn.microsoft.com](https://learn.microsoft.com/en-us/windows-server/get-started/activation-slmgr-vbs-options)
## Microsoft Office Activation
`cd`to the directory where Office ospp.vbs is located:
(!!! ONLY FOR MICROSOFT OFFICE 2016-2021 !!!)
```shell
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"
```
Activate Office：
（`OSPP.VBS` needs `cscript`）
```shell
cscript ospp.vbs /inpkey:OFFICE_GVLK # Install GVLK
cscript ospp.vbs /sethst:cn-cd-txy-1.starryfrp.com # or `162.14.112.182:52206`
cscript ospp.vbs /setprt:52206 # Specific port
cscript ospp.vbs /act # Activate now
```
Go [here](office-key.html) to view `OFFICE_GVLK`
### Useful Activation Commands

Command	|Description
-|-
cscript ospp.vbs /dstatus	|	Displays license information for installed product keys.
cscript ospp.vbs /dstatusall	|	Displays license information for all installed licenses.
cscript ospp.vbs /unpkey:XXXXX|	Uninstalls an installed product key **with the last five digits of the product key** to uninstall.
cscript ospp.vbs /inpkey:XXXXX-XXXXX-XXXXX-XXXXX-XXXXX	|	Installs a product key (replaces existing key) with a user-provided product key. Value option is required.
cscript ospp.vbs /sethst:XXXXXX	|Sets a KMS host name with a user-provided host name.
cscript ospp.vbs /serprt:XXXXX|Sets a KMS port with a user-provided port number. The default port number is 1688.
cscript ospp.vbs /remhst	|	Removes KMS host name and sets port to default. The default port is 1688.
cscript ospp.vbs /act|		Activates installed Office product keys.

The above command is only used to activate the VL version of Office. If it is currently the Retail version, please convert it to the batch authorization version first.
---