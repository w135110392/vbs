dim str
str=inputbox("输入要读的英文单词或句子","提示")
set s=CreateObject("SAPI.SpVoice")
s.speak str
