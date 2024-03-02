# 2026届三班 KMS 激活服务
简体中文 | [English](/en/index.html)
提供Microsoft Windows 和 Microsoft Office 的 KMS 激活服务
## Microsoft Windows 激活
```shell
slmgr /ipk WIN_GVLK # 安装 KMS 密钥
slmgr /skms cn-cd-txy-1.starryfrp.com:52206 # 或 `162.14.112.182:52206`
slmgr /ato # 立即尝试激活
slmgr /dlv # 查看许可证状态
```
在[此处](win-key.html)查看`WIN_GVLK`
### 更多 Slmgr.vbs 的用法
见 [learn.microsoft.com](https://learn.microsoft.com/zh-cn/windows-server/get-started/activation-slmgr-vbs-options)
## Microsoft Office 激活
`cd`到 Office ospp.vbs 所在的目录：
（对于下面命令中的四个`Office14`，2010 版本不需要改动，2013 版本使用`Office15`，2016-2021 版本使用`Office16`）
```shell
if exist "%ProgramFiles%\Microsoft Office\Office14\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office14"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office14\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office14"
```
激活 Office：
（`OSPP.VBS`需要`cscript`运行）
```shell
cscript ospp.vbs /inpkey:OFFICE_GVLK # 安装 KMS 密钥
cscript ospp.vbs /sethst:cn-cd-txy-1.starryfrp.com # 或 `162.14.112.182:52206`
cscript ospp.vbs /setprt:52206 # 设置端口
cscript ospp.vbs /act # 立即尝试激活
```
在[此处](office-key.html)查看`OFFICE_GVLK`
2010和2013版本在[这里](https://kms.343.re/office)
### 常用激活命令

命令	|说明
-|-
cscript ospp.vbs /dstatus	|显示当前已安装产品密钥的许可证信息
cscript ospp.vbs /dstatusall	|显示当前已安装的所有许可证信息
cscript ospp.vbs /unpkey:XXXXX|	卸载已安装的产品密钥（最后5位）
cscript ospp.vbs /inpkey:XXXXX-XXXXX-XXXXX-XXXXX-XXXXX	|安装产品密钥
cscript ospp.vbs /sethst:XXXXXX	|设置 KMS 主机名
cscript ospp.vbs /serprt:XXXXX|设置 KMS 端口（默认是`1688`）
cscript ospp.vbs /remhst	|删除 KMS 主机名
cscript ospp.vbs /act|	激活 Office

以上命令仅用于激活 VL 版本的 Office，如果当前为 Retail 版本，请先转化为批量授权版本。
---