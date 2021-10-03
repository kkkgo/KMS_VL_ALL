# KMS_VL_ALL - Smart Activation Script (Version 44)

## Supported Volume Products:  
[see here](https://github.com/lixuy/vlmcsd#valid-apps)
>Server/Windows  
https://docs.microsoft.com/en-us/windows-server/get-started/kmsclientkeys  
office2016 / office 2019  /office 2021  
https://docs.microsoft.com/en-us/DeployOffice/vlactivation/gvlks   
office2013   
https://technet.microsoft.com/en-us/library/dn385360.aspx   
office2010  
https://technet.microsoft.com/en-us/library/ee624355(v=office.14).aspx   
*Office Retail must be [converted](https://github.com/kkkgo/office-C2R-to-VOL) to Volume first, before it can be activated with KMS.  
KMS activation on Windows 7 have a limitation related to SLIC 2.1 and Windows marker
## How To Use:
>Remove any other KMS solutions.  
Temporary suspend Antivirus realtime protection, or exclude the downloaded file and extracted folder from scanning to avoid quarantine.  
Extract the downloaded file contents to a simple path without special characters or long spaces.  
Administrator rights are require to run the activation script(s).  

**KMS_VL_ALL offer 3 flavors of activation modes:**
### Manual mode (without leaving any KMS emulator traces in the system.)
make sure that auto renewal solution is not installed, or remove it  
then, just run **Activate.cmd**  

>You will have to run Activate.cmd again before the **KMS activation period expire**(default 180 days).  
You can run Activate.cmd anytime during that period to renew the period to the max interval.  
If Activate.cmd is accidentally terminated before it completes, run the script again to clean any leftovers.  

### Auto Renewal mode (the system itself handle and renew activation per schedule.)
first, run the script AutoRenewal-Setup.cmd, press **Y** to approve the installation  
then, just run **Activate.cmd**  

>If you use Antivirus software, it is best to exclude this file from scanning protection:
**C:\Windows\System32\SppExtComObjHook.dll**

### Online KMS mode:

You may use Activate-Local.cmd for online activation,
if you have valid/trusted external KMS host server.
 - edit Activate.cmd with Notepad (or text editor)
 - change External=0 to 1
 - change KMS_IP=172.16.0.2 to the IP/address of the server
 - save the script, and run it as administrator

## Setup Preactivate:

 - To preactivate the system during installation, copy $oem$ folder to sources folder in the installation media (iso/usb).

 - If you already use another setupcomplete.cmd, rename this one to KMS_VL_ALL.cmd or similar name
then add a command to run it in your setupcomplete.cmd, example:
call KMS_VL_ALL.cmd

 - Use AutoRenewal-Setup.cmd if you want to uninstall the project afterwards.

>Notes:  
The included setupcomplete.cmd support the Additional Options described previously, except Unattended Switches.
Use AutoRenewal-Setup.cmd if you want to uninstall the project afterwards.
In Windows 8 and later, running setupcomplete.cmd is disabled if the default installed key for the edition is OEM Channel.

## More help
see ReadMe.html

* * *
# KMS_VL_ALL - 一个精巧灵活的激活批处理
>准备：把KMS_VL_ALL目录放到合适的位置（无特殊字符的路径），删除或卸载其他相关KMS软件，退出杀毒软件。  
零售版office需要经过[转换](https://github.com/kkkgo/office-C2R-to-VOL)成VL后才能使用KMS激活。  
*注意：由于微软的限制，对于BIOS具有SLIC的品牌机，可能无法使用KMS激活Window7系统。[支持的产品](https://github.com/kkkgo/KMS_VL_ALL#supported-volume-products)
 
## 使用一次性的KMS手动续期激活（系统不会增加任何文件）
如果你不需要自动续期，可以直接右键管理员运行**Activate.cmd**即可。  
你必须在KMS到期（默认是180天）前再次运行一次。 
>Activate.cmd是一个自动激活本机所有批量产品的批处理。
它可以单独使用，如果你想用第三方的KMS服务器的话，可以编辑设置External值为1，并填上IP（或者域名）和port（默认1688）。  

## 安装自动续期的KMS激活（系统会增加计划任务和必要的hook）
  - 1、
 先右键管理员运行脚本**AutoRenewal-Setup.cmd**，这是一个带有KMS服务器的hook，会劫持系统的KMS组件，请让杀毒软件放行，输入y安装；如果你需要卸载，只需要再次运行他，输入y卸载。如果您使用防病毒软件，最好从扫描保护中排除此文件：
**C:\Windows\System32\SppExtComObjHook.dll**
  - 2、运行**Activate.cmd**即可
>不管你是用哪种方式激活，如果你安装了新的产品，你仍至少需要运行一次**Activate.cmd**来处理产品激活。

## 其他文件
 - **Check-Activation-Status-vbs.cmd** 是检查激活状态的脚本（使用VBS）。  
 - **Check-Activation-Status-wmic.cmd** 是检查激活状态的脚本（使用WMI）。  
 - **$OEM$** 是用于封装系统部署自动激活的文件夹。


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
