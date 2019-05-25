# KMS_VL_ALL - Smart Activation Script (Version 32)

## Supported Volume Products:  
[see here](https://github.com/lixuy/vlmcsd#valid-apps)
>Server/Windows  
https://docs.microsoft.com/en-us/windows-server/get-started/kmsclientkeys  
office2016 / office 2019  
https://docs.microsoft.com/en-us/DeployOffice/vlactivation/gvlks   
office2013   
https://technet.microsoft.com/en-us/library/dn385360.aspx   
office2010  
https://technet.microsoft.com/en-us/library/ee624355(v=office.14).aspx   

## Auto Renewal:

To install this solution for auto renewal activation, run these scripts respectively:

1-SppExtComObjPatcher.cmd
install/uninstall the Patcher Hook.

2-Activate-Local.cmd
activate installed supported products (you must run it at least once).
you may need to run it again if you installed Office product afterwards.

## Manual:

To only activate without installing and without renewal, run this script only:

KMS_VL_ALL.cmd

you will need to run it again before the activation period expire (6 months by default).

## Windows 10 KMS 2038:

Both KMS_VL_ALL.cmd and Activate-Local.cmd are set by default
to check and skip Windows activation if KMS 2038 detected

However, if you would like to revert or use normal KMS activation:
- edit KMS_VL_ALL.cmd or 2-Activate-Local.cmd with Notepad
- change KMS38=1 to zero 0
- save the script, and run it as administrator

## Online KMS:

You may use Activate-Local.cmd for online activation,
if you have valid/trusted external KMS host server.

- edit Activate-Local.cmd with Notepad
- change KMS_IP=172.16.0.2 to the IP/address of online KMS server
- change Online=0 from zero 0 to 1
- save the script, and run it as administrator

## Setup Preactivate:

- To preactivate the system during installation, copy $oem$ to "sources" folder in the installation media (iso/usb)

- If you already use another setupcomplete.cmd, rename this one to KMS_VL_ALL.cmd or similar name
then add a command to run it in your setupcomplete.cmd, example:
call KMS_VL_ALL.cmd

- Use SppExtComObjPatcher.cmd if you want to uninstall the project afterwards.

- Note: setupcomplete.cmd is disabled if the default installed key for the edition is OEM Channel

## Remarks:

- Some security programs will report infected files, that is false-positive due KMS emulating.
- Remove any other KMS solutions. Temporary turn off AV security protection. Run as administrator.
- If you installed the solution for auto renewal, exclude this file in AV security protection:
C:\Windows\system32\SppExtComObjHook.dll

## KMS Options for advanced users:

You can modify KMS-related options by editing SppExtComObjPatcher.cmd or KMS_VL_ALL.cmd or setupcomplete.cmd

- KMS_Emulation
Enable embedded KMS Emulator functions
never change this option

- KMS_RenewalInterval
Set interval (minutes) for activated clients to auto renew KMS activation
this does not affect the overall KMS period (6 months)
allowed values: from 15 to 43200

- KMS_ActivationInterval
Set interval (minutes) for products to attempt KMS activation, whether unactivated or failed activation renewal
this does not affect the overall KMS period (6 months)
allowed values: from 15 to 43200

- KMS_HWID
Set custom KMS host Hardware ID hash, 0x prefix is mandatory
only affect Windows 8.1/ 10

- Windows, Office2010, Office2013, Office2016, Office2019
Set custom fixed KMS host ePID for each product, instead generating it randomly

## Debug:

If the activation failed, you may run the debug mode to help determining the reason

move SppExtComObjPatcher-kms folder to a short path
with Notepad open/edit KMS_VL_ALL.cmd
change the zero 0 to 1 in set _Debug=0
save the script, and run it as administrator
wait until command prompt window is closed and Debug.log is created
then upload or copy/post the log file

Note: this will auto remove SppExtComObjPatcher if it was installed

* * *
# KMS_VL_ALL - 一个精巧灵活的激活批处理（版本28）

## 安装自动续期的KMS激活
  - 1、把KMS_VL_ALL目录放到合适的位置，确保你不会删除它。  
 先右键运行1-SppExtComObjPatcher.cmd，这是一个带有KMS服务器的HOOK，会劫持系统的KMS组件，请让杀毒软件放行，输入y安装；如果你需要卸载，只需要再次运行他，输入y卸载。
 
  - 2、右键运行2-Activate-Local.cmd，这是一个自动激活本机所有批量产品的批处理，它可以单独使用，如果你想用第三方的KMS服务器而不调用1的劫持服务器的话，可以编辑它，设置Online值为1，并填上IP（或者域名）和port（默认1688）。  
  不管你是用哪种服务器激活，如果你安装了新的产品，你仍至少需要运行一次2来处理产品激活。

## 安装一次性的KMS手动续期激活
如果你不需要自动续期，可以直接运行KMS_VL_ALL.cmd即可。该脚本是1和2的合体，并且在激活处理完成后会自动卸载1。  

## 其他文件
 - Check-Activation是检查激活状态的脚本。  
 - $OEM$ 文件夹是用于系统部署自动激活的脚本。


## Credits:

 - namazso - SppExtComObjHook, IFEO AVrf custom provider.
 - qad - SppExtComObjPatcher, IFEO Debugger.
 - [Mouri_Naruto](https://github.com/MouriNaruto)   - SppExtComObjPatcher-DLL  
 - os51 - SppExtComObjPatcher ported to MinGW GCC, Retail/MAK checks examples.
 - MasterDisaster - Original script, WMI methods.
 - qewpal - KMS-VL-ALL script.
 - Windows_Addict - suggestions, ideas and documentation help.
 - NormieLyfe - GVLK categorize, Office checks help.
 - rpo, presto1234 - scripting suggestions.
 - Nucleus, Enthousiast, s1ave77, l33tisw00t, LostED, Sajjo and MDL Community for interest, feedback and assistance.
 - abbodi1406 - KMS_VL_ALL-SppExtComObjPatcher-kms

>This is a copy from the mydigitallife forum.  
https://forums.mydigitallife.net/threads/kms-activate-windows-8-1-en-pro-and-office-2013.49686/page-76#post-838808
