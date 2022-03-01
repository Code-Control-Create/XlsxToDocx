import sys
import time
import datetime


class ClassError:

    PATH = None
    INFO = None
    START_TIME = None
    TEST = False

    def __init__(self):
        pass

    @staticmethod
    def init(path):
        ClassError.PATH = path
        ClassError.START_TIME = time.time()
        ClassError.INFO = open(path + "输出信息.txt", 'w')
        sys.stderr = ClassError.INFO

    @staticmethod
    def warn(msg: str):
        string = "!warning:{str:^20}.\n".format(str=msg)
        ClassError.write(string)
        print(string)

    @staticmethod
    def error(msg: str):
        string = "!!!error:{str:^20}.\n".format(str=msg)
        ClassError.write(string)
        ClassError.end()
        print(string)
        sys.exit()

    @staticmethod
    def write(string):
        ClassError.INFO.write(str(datetime.datetime.now()) + "    " + string)

    @staticmethod
    def end():
        nowTime = time.time() - ClassError.START_TIME
        string = "完成！共用时(秒)：%f" % nowTime
        ClassError.write(string)
        ClassError.INFO.close()

    @staticmethod
    def get():
        ClassError.end()
        sys.exit()


if __name__ == "__main__":
    ClassError.init("result/")
    ClassError.warn("warn1")

    ClassError.error("error1")
