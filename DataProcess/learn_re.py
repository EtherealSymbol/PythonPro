import re

# 1.表示匹配连续的多个数值
reg = r"\d+"
# 从"abcd123efg"字符串中搜索连续的数值,得到"123",返回一个匹配对象
m = re.search(reg, "abcd123ef789g")
print(m)
# 获得一个或多个分组匹配的字符串，当要获得整个匹配的子串时，可直接使用 group() 或 group(0)
print(m.group())
# print(m.group(0))
# 获取分组匹配的子串在整个字符串中的起始位置（子串第一个字符的索引），参数默认值为 0
print(m.start())
# 获取分组匹配的子串在整个字符串中的结束位置（子串最后一个字符的索引+1），参数默认值为 0
print(m.end())
print(m.span())

# 2.字符串"\d"匹配0~9之间的一个数值
reg1 = r"\d"
m1 = re.search(reg1, "abc123def456ghi")
print(m1)

# 3.注意:r"b\d+"第一个字符要匹配"b",后面是连续的多个数字,因此"是b1233",不是"a12"
reg2 = r"b\d+"
m2 = re.search(reg2, "a12b3456cd")
print(m2)

# 4."*" 与 "+"类似,但有区别,列如:
# 可见 r"ab+“匹配的是"ab”,但是r"ab “匹配的是"a”,因为表示"b"可以重复零次，但是”+“却要求"b"重复一次以上
reg3 = r"ab+"
m3 = re.search(reg3, "acabc")
print(m3)

reg4 = r"ab*"
m4 = re.search(reg4, "acabc")
print("m4   ", m4)

# 5.字符"?"重复前面一个匹配字符零次或者一次.
reg5 = r"ab?"
m5 = re.search(reg5, "abbcabc")
print(m5)
m6 = re.search(reg5, "addbbcabc")
print(m6)

# 6.字符".“代表任何一个字符,但是没有特别声明时不代表字符”\n"
reg6 = r"a.b"
m7 = re.search(reg6, "xxaxxxxbxxxxx")
print(m7)
m8 = re.search(reg6, "xxaxbxxxxx")
print(m8)

# 7."|"代表把左右分成两个部分, 结果匹配"ab"或者"ba"都可以
reg7 = r"ab|ba"
m9 = re.search(reg7, "xaabababy")
print(m9)

# 8.特殊字符使用反斜杠"“引导,例如”\r"、"\n"、"\t"、"\"分别表示回车、换行、制表符号与反斜线自己本身
reg8 = r"a\nb"
n = re.search(reg8, "ca\nbcaba")
print(n)

# 9.字符"\b"表示单词结尾,单词结尾包括各种空白字符或者字符串结尾.
# 结果匹配"car",因为"car"后面是一个空格.
reg9 = r"car\b"
n1 = re.search(reg9, "this car is black")
print(n1)
n2 = re.search(reg9, "these cars are black")
print(n2)

# 10."[]中的字符是任选择一个,如果字符ASCll码中连续的一组,那么可以使用"-"字符连接,例如[0-9]表示0-9的其中一个数字,
# [A-Z]表示A-Z的其中一个大写字符,[0-9A-z]表示0-9的其中一个数字或者A-z的其中一个大写字符.
reg10 = r"x[0-9]y"
n3 = re.search(reg10, "xy9x19yyx3yyy")
print(n3)

# 11."^"出现在[]的第一个字符位置,就代表取反,例如[ ^ab0-9]表示不是a、b，也不是0-9的数字.
reg11 = r"x[^ab0-9]y"
n4 = re.search(reg11, "xayxbyx2yxcy")
print(n4)

# 12."\s"匹配任何空白字符,等价于"[\r\n\x20\t\f\v]"
reg12 = r"a\sb"
n5 = re.search(reg12, "xa\tba bxy")
print(n5)
n6 = re.search(reg12, "xa ba bxy")
print(n6)

# 13."\w"匹配包括下划线子内的单词字符,等价于"[a-zA-Z0-9]"
reg13 = r"\w+"
n7 = re.search(reg13, "Python is easy.")
print(n7)

# 14."$"字符匹配字符串的结尾位置
# 匹配结果是最后一个"ab",而不是第一个"ab"
reg14 = r"ab$"
n8 = re.search(reg14, "abcab")
print(n8)

# 15.使用括号()可以把()看出一个整体,经常与"+"、"*"、"?"的连续使用,对（）部分进行重复.
# 结果匹配"abab","+“对"ab"进行了重复
reg15 = r"(ab)+"
n9 = re.search(reg15, "abababcab")
print(n9)

# 16.使用正则表达式r"[A-Za-z]+\b"匹配单词,它表示匹配由大小写字母组成的连续多个字符,一般是一个单词,之后"\b"表示单词结尾.
reg16 = r"[A-Za-z]+\b"
str = "I love Java! But now I'm learning python."
n10 = re.search(reg16, str)
# print(n10)
# end = n10.end()
# n11 = re.search(reg16, str[end:])
# print(n11)
temp = 0
while n10 is not None:
    print(n10)
    start = n10.start()
    end = n10.end() + temp
    temp = end
    n10 = re.search(reg16, str[end:])




