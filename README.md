# actreerule
AC 多模树修改正则

只有两个对比类：

1 PatternMatchTest ：

普通正则匹配模式：使用超长正则匹配文本，耗时约 5000 ms


2 ACTrie 使用拆分后的类：

使用AC多模匹配，将超长正则拆分后结果耗时在约 100ms ~ 200ms


详细解释见代码注释。

AC 多模树创建见 ACTrie  类
