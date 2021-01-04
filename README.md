# epm_Tools_CopyDrivers
CopyDrivers Tool by Jan Buelens (former Landesk employee)

## origin
http://community.landesk.com/support/docs/DOC-2714
as part of the Hardware Indepenedent Imaging Concept that Jan Buelens published.
https://www.yumpu.com/en/document/view/48139178/hardware-independent-imaging-hii-and-landesk-community

## 2020 still in use in our company
Unfortunately, it can no longer be found in the Landesk forum. But I still had it on my hard disk.

## License
All rights reserved by Jan Buelens or LANDESK


## Manual (from linked PDF)
CopyDrivers.exe
CopyDrivers uses the WMI model name to copy machine specific drivers from a source
folder on a server to a destination folder on the local machine.
CopyDrivers reads out the local machine's WMI model name and uses a file called
copydrivers.ini to map the WMI model to a machine specific driver folder.

``` ini
[Config]
DriversSource=\\Core88a\Provisioning\Drivers
DriversTarget=c:\drivers
[Models]
Precision Workstation T3400=Dell Precision T3400
VMware Virtual Platform=VMWare_Workstation
PowerEdge 1950=Dell PowerEdge 1950
9071A1G=ThinkCenter M57p
2007FVG=ThinkPad T60
64575KG=Thinkpad T61p
```
In the [Models] section, each line associates one model name (left of the equal sign) with
a subfolder name. Example: if the WMI model name is 2007FVG, the subfolder name is
ThinkPad T60. CopyDrivers appends this subfolder name to the base path defined by the
DriversSource line in the [Config] section. CopyDrivers then copies this entire folder
(\\Core88a\Provisioning\Drivers\ThinkPad T60 in the example) to a local folder on the
target machine, as defined by the DriversTarget line (c:\drivers in the example).
When matching WMI model names with lines in copydrivers.ini, CopyDrivers performs
wildcard comparsions. See section 3.1.1, Wildcard matching.
When you invoke CopyDrivers without command line parameters, a little GUI programs
opens that can be used to build or edit CopyDrivers.ini.
Below is the usage info displayed when you type copydrivers /?


![Options](/assets/2ec4ba39f54745209e5853779e99bc9a.png)


The /s and /d parameters can be used to override the DriversSource and DriversTarget
lines in copydrivers.ini22.
CopyDrivers also provides support for drivers that require a setup program, as explained
in section 3.4. CopyDrivers looks for 2 files in the machine specific driver folder:
cmdlines.txt and GuiRunonce.ini. These might have been built by the CopyDrivers GUI,
or you might have built them manually.
When you use the CopyDrivers GUI, there is no need to be aware of the format of these
files. Their format is as described in the sysprep documentation. Here is a smaple
cmdlines.txt file:
[cmdlines]
"c:\drivers\setup\driver1\setup.exe"
"c:\drivers\setup\driver2\setup.exe"
And here is a sample GuiRunonce.ini file:
[GuiRunOnce]
Command0="c:\drivers\setup\\driver1\setup.exe"
Command1="c:\drivers\setup\\driver2\setup.exe"
If your image is already using cmdlines and GuiRunOnce, the lines will be added to the
existing cmdlines.txt file or sysprep.inf [GuiRunOnce] section.
If you follow the conventions used in this document, CopyDrivers will create a log file
under c:\drivers (see section 7.5). There is no need to specify the /log parameter unless
you want the log created somewhere else.
