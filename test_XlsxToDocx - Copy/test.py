import re
import stream

# Data = stream.ClassStream()
#
# Data[1] = "ok"
# # Data[2] = "ook"
# # for j in Data:
# #     print(j)
#
# ree = re.compile(r"\[(.*)(?=|)[|]?(?<=[|\[])(.*?[A-Z]+.*?)(?=[|\]])[|]?(?<=|)(.*)]")
#
# text = "[G+1]"
#
# print(ree.match(text).groups())
# print("sos"[1:])
#
# ree = re.compile(r"\[(\d*?)([ecnf]?)(\d*?)([yta]?)(\d*)([A-Z]+)(.*?)]")
#
# text = "[10A]"
#
# for ii in {}.keys():
#     print(type({}.keys()))

test = re.findall("[a-z]+", "你好")
print(test)

A = [0, 1, 2]
B = [0, 1, 2]

C = [x-y for x, y in zip(A, B)]
print(C.count(1))
print(A == B)
print(len(A))

print("{value:<{width}}ok".format(value="ok_ok", width=1.1))
print(int(""))

a = "5"
b = ".6"

print(exec("a="+a+b))

print(a)

