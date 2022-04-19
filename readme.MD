# word replace 文档文字替换

### 文档文字替换
```
1. 替换 word 中 红色字 为空格(Replace red words with spaces in word)
输入目录：D:\upload\a
输出目录：D:\upload\a_new
java -jar word2replace-0.0.1-SNAPSHOT.jar D:\upload\a D:\upload\a_new

2. 替换 word 中 红色字 为空格，匹配到同名的使用pic目录下的图片替换
java -Dfile.encoding=GB2312 -jar word2replace-0.0.1-SNAPSHOT.jar D:\upload\a D:\upload\a_new 测试.png,监测.png
```

