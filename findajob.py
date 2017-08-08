#
# # # -*- coding: utf-8 -*-1
# # a = 1
# # def fun(a):
# #     a = 2
# # fun(a)
# # print(a)
# #
# # a=[]
# # def fun(a):
# #     a.append(1)
# # fun(a)
# # print(a)
#
# def foo(x):
#     print("excuting foo(%s)"%(x))
#
# class A(object):
#     def foo(self,x):
#         print("excuting foo(%s,%s)"%(self,x))
#     @classmethod
#     def class_foo(cls,x):
#         print("excuting class_foo(%s,%s)"%(cls,x))
#     @staticmethod
#     def static_foo(x):
#         print ("excuting static_foo(%s)"%(x))
# a=A()
# a.foo(10)
# a.class_foo(20)
# A.static_foo(30)
#
# def print_everything(*args):
#     for count,thing in enumerate(args):
#         print('{0}.{1}'.format(count,thing))
# print_everything('apple','banana','peach','oil')
#
# def print_fruit(**kwargs):
#     for name,value in kwargs.items():
#         print('{0}={1}'.format(name,value))
# print_fruit(apple='fruit',cabagge='vegetable')
#
# def pring_three_things(a,b,c):
#     print('a={0},b={1},c={2}'.format(a,b,c))
# pring_three_things('tom','jack','black')

str = 'Hello World!'

print str  # 输出完整字符串
print str[0]  # 输出字符串中的第一个字符
print str[2:5]  # 输出字符串中第三个至第五个之间的字符串
print str[2:]  # 输出从第三个字符开始的字符串
print str * 2  # 输出字符串两次
print str + "TEST"  # 输出连接的字符串