class ClassDeleteElement:
    pass


class ClassStream:

    def __init__(self):

        self.__DictStream = {}

    def __iter__(self):
        return self

    def __next__(self):
        return self.__DictStream.values()

    def __getitem__(self, item):
        return self.__DictStream[item]

    def __setitem__(self, key, value):
        self.__DictStream[key] = value

    def this(self) -> dict:
        return self.__DictStream

    def values(self):
        return self.__DictStream.values()

    def keys(self):
        return self.__DictStream.keys()

    def items(self):
        return self.__DictStream.items()


class ClassShadowStream:
    """单键双值字典

        单键双值字典，用于映射的处理时储存的数据结构

    """
    def __init__(self):
        self.__ShadowStream = {}
        self.__SunStream = {}
        self.__DelKeys = []

        self.Translated = False  # 是否转换完毕

        self.__Iter = False

    def __iter__(self):
        self.__Iter = True
        return self

    def __next__(self):
        if self.__Iter:
            self.__Iter = False
            return list(self.__ShadowStream.values())
            # return 5
        else:
            raise StopIteration

    def __getitem__(self, item):
        return self.__ShadowStream[item]

    def __setitem__(self, key, value):
        """令双字典同键同序

            __ShadowStream 为待处理序列
            __SunStream    为处理完序列

        :param key:
        :param value:
        :return:
        """

        # 对冷数据的优化
        # if value is not None:
        #    self.__ShadowStream[key] = value
        self.__ShadowStream[key] = value
        try:
            jin = self.__SunStream[key]
        except KeyError:
            self.Translated = False
            self.__SunStream[key] = None

    def set(self, key, value):
        self.__SunStream[key] = value

    def this(self) -> dict:
        return self.__SunStream

    def values_(self):
        return self.__SunStream.values()

    def values(self):
        return self.__ShadowStream.values()

    def value_(self, key):
        return self.__SunStream[key]

    def value(self, key):
        return self.__ShadowStream[key]

    def keys_(self):
        return self.__SunStream.keys()

    def keys(self):
        self.del_()
        return self.__ShadowStream.keys()

    def items(self):
        return self.__SunStream.items()

    def move(self, key):
        """映射处理完转移

        :param key:
        :return:
        """
        self.__SunStream[key] = self.__ShadowStream[key]
        del self.__ShadowStream[key]

        if len(self.__ShadowStream) == 0:
            self.Translated = True

    def take(self, key, value):
        if value is not None:
            self.__SunStream[key] = value
            self.__DelKeys += [key]
        if value is ClassDeleteElement:
            del self.__SunStream[key]

    def del_(self):
        for key in self.__DelKeys:
            del self.__ShadowStream[key]

            if len(self.__ShadowStream) == 0:
                self.Translated = True
            else:
                self.Translated = False

        self.__DelKeys = []

    def change(self, key):
        self.__SunStream[key] = self.__ShadowStream[key]
        self.__ShadowStream[key] = None

    def results(self):
        """结果输出

            对映射序列进行输出，无结果则输出处理中值

        :return:
        """
        for key, value in self.items():
            if value is None:
                yield key, self.__ShadowStream[key]

            else:
                yield key, value

    def translated(self):
        self.del_()
        return self.Translated

    def reversed(self):
        self.__SunStream, self.__ShadowStream = self.__ShadowStream, self.__SunStream

    def test_show(self):
        print(self.__ShadowStream)
        print(self.__SunStream)
        print("")


if __name__ == "__main__":  # test
    Data = ClassShadowStream()

    Data[1] = "ok"
    Data[2] = "ook"

    for j in Data:
        print(j)
