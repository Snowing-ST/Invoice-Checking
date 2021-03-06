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

- <font size=2 > 注：1. 若使用谷歌浏览器，则必须是77版本及以下，79版本亲测无法获取网站验证码。2. 使用selenium库的webdriver打开谷歌浏览器需到[此处](http://chromedriver.storage.googleapis.com/index.html)下载chromedriver.exe，注意下载版本一定要和电脑中谷歌浏览器版本一致！并将该文件放入excel文件的目录下（可在代码中修改存放路径）</font>

### 3. 自动截图
跳转至发票验证页面后，使用selenium库控制鼠标在页面滚动至最上方，自动截图，以发票号码命名，保存在excel文件的目录下的“截图”文件夹中。

- <font size=2 >注：由于网速等问题网页跳转慢，可能出现截错页面的情况</font>

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



## 2.0版新特点
- 可直接在python的console中输出验证码图片、提示性文字，以及输入验证码，提高效率
- 识别“查验太频繁，请1分钟后再试”的问题，程序将等待一分钟后自动跳转
- 截图时简化模拟鼠标滚动操作，鼠标仅模拟滑动至最上方

## 3.0版新特点
- 用tkinter包构建图形界面，打包成exe程序后方便不懂代码的人使用，如导入文件时可直接打开文件管理器选择文件，简单直观。
- gui编程的设计思路为事件触发型，因此代码逻辑相比于2.0版本有重大修改，主要是：
    - 1 把获取验证码和提交验证码分为两步； 
    - 2 原2.0逐张验证发票使用for循环，3.0无明显循环的关键字，改为截图后重复一次查验过程，直到已查验发票数i与导入的excel文件中样本数相等； 
    - 3 出错时不再需要while循环，遂删除
    - 4 整个代码写成一个类的形式，大大方便了参数调用

![图形界面示例](https://github.com/Snowing-ST/Invoice-Checking/blob/master/gui.png)