# <center>半自动国税发票查验</center>
<p align="right">by 计算机小白</p>

<font face="仿宋" >之所以说“半自动”因为验证码还得手动填写o(╥﹏╥)o</font>
## 代码步骤
### 1. 发票信息填写
发票信息如发票代码、发票号码、发票时间、税前金额需先行填写在excel中，如[invoice_sample.xlsx](https://github.com/Snowing-ST/invoice_checking/blob/master/invoice_sample.xlsx)所示


|业务编号|	发票代码|	发票号码|	发票日期|	税前金额|
|  ----  | ----  | ----  | ----  | ----  |
|ABC2019FP00004|	3400183130|	02049280|	20181225	|471495.24 |
|ABC2019FP00004|	3400183130|	02049281	|20181225|	937220.75 |

### 2. 自动填写表单
运行[invoice_checking.py](https://github.com/Snowing-ST/invoice_checking/blob/master/invoice_checking.py),提示填写excel表格的路径，填写完毕回车后，selenium库的webdriver将打开谷歌浏览器，自动登陆[国家税务总局全国增值税发票查验平台](https://inv-veri.chinatax.gov.cn/index.html)，自动填写发票信息，**<font color=#FF0000 >验证码需手动输入程序端，不能输在网页上</font>**，如果输错等待浏览器自动跳转，**<font color=#FF0000 >不要点击浏览器任何按钮</font>**，再次在程序端输入验证码。

<font size=2 >注：使用selenium库的webdriver打开谷歌浏览器需到[此处](http://chromedriver.storage.googleapis.com/index.html)下载chromedriver.exe，注意下载版本一定要和电脑中谷歌浏览器版本一致！并将该文件放入excel文件的目录下（可在代码中修改存放路径）</font>

### 3. 自动截图
跳转至发票验证页面后，使用selenium库控制鼠标在页面滚动至最上方，自动截图，以发票号码命名，保存在excel文件的目录下的“截图”文件夹中。

<font size=2 >注：由于网速等问题网页跳转慢，可能出现截错页面的情况</font>

### 4. 自动生成word文档
所有发票验证截图后，将同一个业务项下的发票插入word文档，以业务编码命名，保存在excel文件的目录下的“文档”文件夹中，方便后续打印操作。

## 打包exe文件
为了让不懂程序的人使用该程序，提高工作效率，本代码还可用pyinstaller打包成exe文件，总文件大概700M，打包代码如下：

    pyinstaller -p C:/.../Anaconda2/envs/py3/Lib/site-packages -D invoice_checking.py

-p 指定python第三方包的路径，将代码依赖的包都打包进去

-D 产生一个目录（包含多个文件）作为可执行程序

可执行文件路径：../dist/invoice_checking/invoice_checking.exe


## 关于验证码识别
github上有搜到几篇关于验证码识别的用python写的project，但无奈没有说明文档本计算机小白看不懂，如果哪天github上有详细说明了我再来借鉴借鉴~~



