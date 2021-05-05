VBA学习（Excel篇）
======

VBA概述
------
VBA代表Visual Basic for Applications，这是一种来自Microsoft的事件驱动编程语言，现在主要与Microsoft Office应用程序(如MSExcel，MS-Word和MS-Access)一起使用。

### 亮点
* 容易上手
    * 基本上有过开发经验的人经过基本的代码扫盲就能够快速进行初期开发。
* 提高效率
    * 在处理一些大量重复的操作时，VBA明显能够更快的提供解决方案。
* 保证准确性
    * 在处理数据时，人难免会有一些小差错，而使用程序去进行操作则可以最大程度上的去避免。
* 开发速度快
    * 有时处理一项事务时给的时间是非常紧急的，相比与其他语言，VBA实现可以至少节省两到三倍的时间。
* 安装成本低
    * Microsoft Office应用程序基本上所有办公电脑上都会有安装，只需要配置一下就可以作为开发平台。

### 开发工具配置
![](https://github.com/goro-hyronjs/wulu-VBA-learn/raw/master/Image/A0001.PNG) 
![](https://github.com/goro-hyronjs/wulu-VBA-learn/raw/master/Image/A0002.PNG) 
![](https://github.com/goro-hyronjs/wulu-VBA-learn/raw/master/Image/A0003.PNG) 
![](https://github.com/goro-hyronjs/wulu-VBA-learn/raw/master/Image/A0004.PNG) 

### 打开开发工具
![](https://github.com/goro-hyronjs/wulu-VBA-learn/raw/master/Image/B0001.PNG) 
![](https://github.com/goro-hyronjs/wulu-VBA-learn/raw/master/Image/B0002.PNG) 
![](https://github.com/goro-hyronjs/wulu-VBA-learn/raw/master/Image/B0003.PNG) 
![](https://github.com/goro-hyronjs/wulu-VBA-learn/raw/master/Image/B0004.PNG) 

### 新建模块
![](https://github.com/goro-hyronjs/wulu-VBA-learn/raw/master/Image/C0001.PNG) 
![](https://github.com/goro-hyronjs/wulu-VBA-learn/raw/master/Image/C0002.PNG) 
![](https://github.com/goro-hyronjs/wulu-VBA-learn/raw/master/Image/C0003.PNG) 
![](https://github.com/goro-hyronjs/wulu-VBA-learn/raw/master/Image/C0004.PNG)

### 录制宏
![](https://github.com/goro-hyronjs/wulu-VBA-learn/raw/master/Image/D0001.PNG) 
![](https://github.com/goro-hyronjs/wulu-VBA-learn/raw/master/Image/D0002.PNG) 
![](https://github.com/goro-hyronjs/wulu-VBA-learn/raw/master/Image/D0003.PNG) 
![](https://github.com/goro-hyronjs/wulu-VBA-learn/raw/master/Image/D0004.PNG)

### 代码快速扫盲
#### 注释
任何以单引号(')或者关键字"REM"开头的语句都被视为注释。以下是注释的一个例子。
```
' This Script is invoked after successful login 
REM This Script is written to Validate the Entered Input 
```

#### 模块声明
```
Sub xxxxx()
XXXXXXXXX
End Sub
```

#### 变量声明
```
Dim Para As Type
```

#### 常量
```
Const MyPara As Type = xxx
```

#### 消息框
```MsgBox(prompt[,buttons][,title][,helpfile,context])```
prompt - 必需的参数。在对话框中显示为消息的字符串。提示的最大长度大约为1024个字符。 如果消息扩展为多行，则可以使用每行之间的回车符(Chr(13))或换行符(Chr(10))来分隔行。
buttons - 可选参数。一个数字表达式，指定要显示的按钮的类型，要使用的图标样式，默认按钮的标识以及消息框的形式。如果留空，则按钮的默认值为0。
title - 可选参数。 显示在对话框的标题栏中的字符串表达式。 如果标题留空，应用程序名称将被放置在标题栏中。
helpfile - 可选参数。一个字符串表达式，标识用于为对话框提供上下文相关帮助的帮助文件。
Context - 可选参数。一个数字表达式，用于标识由帮助作者分配给相应帮助主题的帮助上下文编号。 

#### 输入框
```InputBox(prompt[,title][,default][,xpos][,ypos][,helpfile,context])```
Prompt - 必需的参数。 在对话框中显示为消息的字符串。提示的最大长度大约为1024个字符。 如果消息扩展为多行，则可以使用每行之间的回车符(Chr(13))或换行符(Chr(10))来分隔行。
title - 一个可选参数。显示在对话框的标题栏中的字符串表达式。如果标题留空，应用程序名称将被放置在标题栏中。
default - 一个可选参数。用户希望显示的文本框中的默认文本。
xpos - 一个可选参数。X轴的位置表示水平从屏幕左侧的提示距离。 如果留空，则输入框水平居中。
ypos - 一个可选参数。Y轴的位置表示竖直方向从屏幕左侧的提示距离。如果留空，则输入框垂直居中。
helpfile - 一个可选参数。一个字符串表达式，标识用于为对话框提供上下文相关帮助的帮助文件。 
context - 一个可选参数。一个数字表达式，用于标识由帮助作者分配给相应帮助主题的帮助上下文编号。如果提供上下文，则还必须提供helpfile。

#### 运算符
* 算术运算符(+,-,*,/,%)
* 比较运算符(=,<>,<,>,>=,<=)
* 逻辑(或关系)运算符(AND,OR,NOT,XOR)
* 连接运算符(+,&)

#### 决策
```
if...else
if...elseif...else
switch
```

#### 循环
```
for
for...each
while...wend
do...while
do...until
```

#### 自定义函数
```
Function Functionname(parameter-list)
   statement 1
   statement 2
   statement 3
   .......
   statement n
End Function
```

#### 子程序
```
Sub Area(x As Double, y As Double)
   MsgBox x * y
End Sub
```


