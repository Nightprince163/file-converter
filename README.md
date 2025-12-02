# file-converter
基于node + libreoffice库的文档转换工具，支持WORD/PPT/Excel文件（乃至更多文件格式，需要自行调整代码）转PDF，只要libreoffice库支持的格式均可。

## 技术栈
-- node 22.*
-- libreoffice 25.* （libreoffice 7.5+，但最好用25，否则打印PDF格式可能存在异常，容易错乱）

## 使用方法
需要在电脑上安装libreoffice25。windows可直接从官网下载安装包。linux可以通过apt安装，但默认版本只有7.5，我是通过flatpak安装的libreoffice 25版本。
``` bash
# Ubuntu/Debian
sudo apt update
sudo apt install libreoffice-writer libreoffice-core
# 或者通过flatpak安装：
flatpak install flathub org.libreoffice.LibreOffice
```
注意：flatpak安装的libreoffice，调用方式也要通过flatpak，而apt安装的libreoffice，是通过soffice命令来调用。而我的代码中linux环境下使用的是flatpak方式，如果你的linux服务器用的是soffice，请直接用windows环境的soffice部分代码。

### 开发环境启动：
``` bash
npm start //启动服务
```
调用方式：
``` typescript
const formData = new FormData()
      formData.append('file', file)
const response = await fetch('http://localhost:3001/api/convert-docx-to-pdf', {
          method: 'POST',
          body: formData,
        })
if (!response.ok) {
          throw new Error('文件转换失败，请稍后重试')
}
const arrayBuffer = await response.arrayBuffer()
```
### 生产环境：
无需编译，直接全部部署到服务器，然后用pm2启动。具体方法恕不详述，请自行查阅资料。
