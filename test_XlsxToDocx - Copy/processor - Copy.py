import re
import copy

import stream
import doctemplate
import xlsxmethod
from error import ClassError


g_value_for_exec = [None]


class ClassWord:
    """运行时词单元

        运行中对于填充词法个体的表达
        由ClassWordProcessor初始化

    """

    CON_NEXT_SENTENCE = None

    def __init__(self, prefix: str, suffix: str, pure: str):
        # 词定义
        self.PureValue = pure
        self.Prefix = prefix
        self.Suffix = suffix
        self.Value = None                # 对应的ABC的xlsx列，全由WordProcessor填写
        self.Width = 0
        self.Method = None
        self.ExpressTime = 1
        self.NextSentence = None
        self.FixType = 0                 # 枚举体中结点后数据索引方式，0并集，1交集 wuyong
        self.Enum = False                # 是否枚举遇见的所有值用以区分ct 与 ey

        # 词的值
        self.Meaning = ""
        self.ValueCache = []

    def get_data(self, line_data: xlsxmethod.ClassXlsxData):
        """各个get_data调用，词单元的数据填充

            对词单元的填充
            大致流程：
            1，像当前数据对象请求此刻数据
            2，根据环境与自身属性对数据进行缓存，并向数据对象返回信息
            3，接收完毕输出

        :param line_data:
        :return:
        """
        jian = line_data.ask(self.Value, self.Enum)
        # Enum共享一组数据，并进行抢夺
        # 可以解决eyA,A的情况
        # 对于e2yA,c1tA,c2tA的表述存在歧义，对于A:1,1,2输出为1 2， 1， 1（ey设置结点强制不重复）应有用户选择意义
        if jian is not None:  # 接收到
            self.ExpressTime -= 1
            if self.ExpressTime == 0:    # 接收完毕
                self.ValueCache += [jian]
                if self.Enum:
                    if ClassMainProcessor.XLSX_FINISHED:     # 再无后续数据，不应再接收数据，但对Enum刚好差次结束，吐出
                        self.ValueCache = self.ValueCache[:-1]
                    else:
                        ClassWord.CON_NEXT_SENTENCE = self.NextSentence  # 后续还有
                        line_data.none(self.Value)
                return self.value()
            else:   # 接收未完毕
                if self.Enum:
                    ClassWord.CON_NEXT_SENTENCE = None   # 仅允许最后结点的设置生效(不合理，应考虑队列)在InputBody增加句子时Key异常
                    if ClassMainProcessor.XLSX_FINISHED:
                        return self.value()
                    else:
                        try:    # 枚举体拒绝录入重复数据
                            test = self.ValueCache.index(jian)
                            self.ExpressTime += 1
                            return None
                        except ValueError:
                            self.ValueCache += [jian]
                            line_data.none(self.Value)
                        return None
                elif self.ExpressTime < -1:  # 为冷数据做匹配，由于冷数据分开储存，且只在结束时录入一次，注意jian不能为None
                    self.ValueCache += [line_data[self.Value].get_value(self.ExpressTime + 1)]
                    return self.value()
                elif self.ExpressTime >= 1 and ClassMainProcessor.XLSX_FINISHED:  # 数据全导入后，数量却不够
                    self.ValueCache += [jian]
                    return self.value()
                else:
                    return None
        else:
            if ClassMainProcessor.XLSX_FINISHED:  # 强制返回，之前在ClassSentenceWithData中debug所留
                return self.value()
            return None

    def value(self):
        """由self.get_data调用返回值"""
        jian = []
        for i in self.ValueCache:
            jian += [self.__exec(i)]
        self.ValueCache = jian

        try:
            self.Meaning += self.ValueCache[0]
            for jian in self.ValueCache[1:]:
                self.Meaning += self.Prefix + jian
            if self.Meaning != "":
                self.Meaning += self.Suffix

        except IndexError:
            pass
        finally:
            # 用于未完成的句的判定
            if ClassMainProcessor.XLSX_FINISHED and self.Meaning == "":
                return None
            self.Meaning = "{value:<{width}}".format(value=self.Meaning, width=self.Width)
            return self.Meaning

    def __exec(self, value):
        """self.value调用对于数字元素的简单处理"""
        if self.Method == "":
            return value

        test = re.findall("[a-z\n.,，。]+", self.Method)
        try:
            method = test[0]
            return value

        except IndexError:
            method = self.Method
            test = value
            g_value_for_exec[0] = value
            try:
                exec("g_value_for_exec[0] = " + test + method, {"g_value_for_exec": g_value_for_exec})
                return str(g_value_for_exec[0])
            except:
                return value


class ClassMsg:
    """初始化处理中信使

        用于在初始化的各个处理进程之间传递信息

    """
    def __init__(self):
        self.NewEnumBody = 1  # 新的枚举体？由MainProcessor.pretreatment减，由WordProcessor加
        self.SentenceNumber = 1  # 枚举体内句子计数
        self.WordCount = 0
        self.Address = None  # 由MainProcessor.pretreatment填入
        self.NowEnumBody = False  # 当前输入是否是真枚举体, 用于枚举体转入非枚举体

    def next(self):              # 由ClassMainProcessor.pretreatment 调用
        self.WordCount = 0
        self.SentenceNumber += 1

    def new(self):               # 由MainProcessor.pretreatment减
        self.NewEnumBody -= 1
        self.SentenceNumber = 1
        self.WordCount = 0
        self.NowEnumBody = False

    def new_address(self, address_body):  # 由MainProcessor.pretreatment减
        self.Address = address_body
        self.new()


class ClassSentenceWithData:
    """句子对象和分析产生的词对象列表"""
    def __init__(self, sentence: doctemplate.ClassSentences, word_list: list):
        self.Sentence = sentence
        self.WordList = word_list
        self.Data = stream.ClassShadowStream()

    def get_data(self, line_data: xlsxmethod.ClassXlsxData):
        """InputBody.get_data接收数据并使其遍历Word对象"""
        for word in self.Data.keys():
            if self.Data[word] is None:
                continue
            self.Data.take(word, self.Data[word].get_data(line_data))

        if self.Data.translated():
            text = self.Sentence.Text
            for i in range(len(self.WordList)):
                text = text.replace(self.WordList[i], self.Data.value_(i), 1)
            self.Sentence.Result = text[2:-2]
            return self.Sentence
        else:
            if ClassMainProcessor.XLSX_FINISHED:  # 数据输入结束后强制求值
                text = self.Sentence.Text
                for i in range(len(self.WordList)):
                    try:
                        text = text.replace(self.WordList[i], self.Data.value_(i), 1)
                    except TypeError:
                        try:
                            text = text.replace(self.WordList[i],
                                                ClassMainProcessor.TIMING_COLD_WORD.value_(self.WordList[i]),
                                                1)
                        except KeyError:  # 如无任何数据接收删除该句
                            return stream.ClassDeleteElement

                self.Sentence.Result = text[2:-2]  # 去头尾
                return self.Sentence
            else:
                return None


class ClassInputBody:
    """运行时句子管理单元

        运行时句子管理

    """
    def __init__(self):
        # 接受数据的句子结构流 填充ClassSentence
        self.SentenceStream = stream.ClassShadowStream()  # 由SentenceProcessor填入SentenceWithData
        self.PureSentence = stream.ClassShadowStream()
        self.AddSentence = 1
        self.NextSentence = None

        # # 映射池 填充ClassWord 要具体到句
        # self.MappingPool = stream.ClassShadowStream()  # 由WordProcessor填入,并由SentenceProcessor填入控制符

    def get_data(self, line_data: xlsxmethod.ClassXlsxData):
        """从EnumBody.get_data调用接收对象并遍历SentenceWithData对象"""
        self.__add_sentence()

        for sen in self.SentenceStream.keys():
            if self.SentenceStream[sen] is None:
                continue
            self.SentenceStream.take(sen, self.SentenceStream[sen].get_data(line_data))

        self.__get_sentence()

        if self.SentenceStream.translated() and self.NextSentence is None:

            return self.SentenceStream
        else:
            return None

    def set_sentence(self, count, sentence_body: ClassSentenceWithData):
        """初始化"""
        self.SentenceStream[count] = copy.deepcopy(sentence_body)
        self.PureSentence[count] = copy.deepcopy(sentence_body)

    def __add_sentence(self):
        """新增句子
            读入句子是新句子和最后都要读入，最后不会生成新的
        """
        if (not ClassMainProcessor.XLSX_FINISHED) and self.NextSentence is not None:
            self.SentenceStream["add"+str(self.AddSentence)] = copy.deepcopy(self.PureSentence[self.NextSentence])
            self.NextSentence = None
            self.AddSentence += 1
        elif ClassMainProcessor.XLSX_FINISHED:
            self.NextSentence = None

    def __get_sentence(self):
        """缓存要增加的句子，仅当由有新数据的时候才生成"""
        if ClassWord.CON_NEXT_SENTENCE is not None and ClassWord.CON_NEXT_SENTENCE > 0:
            self.NextSentence = ClassWord.CON_NEXT_SENTENCE
            ClassWord.CON_NEXT_SENTENCE = None
        if ClassMainProcessor.XLSX_FINISHED:
            self.NextSentence = None


class ClassNameFileBody:
    """用于表达输出文件的路径，名称的构成方法

        由由WordProcessor填入负责填写

    """
    NAME_INTERVAL = "_"
    FILE_ROOT_PATH = "./result/"

    def __init__(self):
        """
            NamePool : ClassWord
            FilePool : ClassWord
        """
        self.NamePool = stream.ClassShadowStream()  # 均由WordProcessor调用get填入
        self.FilePool = stream.ClassShadowStream()
        self.NameIndex = []  # 用于确定docx索引

        self.NameOrderDict = {}  # 命名的引索与键的字典
        self.FileOrderDict = {}  # 路径的引索与键的字典

    def get(self, word_name: str, word_body: ClassWord or None, kind: str, order: int):
        """WordProcessor.analysis调用填充stream

        :param word_name:
        :param word_body:
        :param kind:
        :param order:
        :return:
        """
        # 异地判定是很不好的习惯，请避免
        if kind == "normal":
            pass
        elif kind == "Name":
            self.NamePool[word_name] = word_body
            try:
                test = self.NameOrderDict[order]
                ClassError.error("重复定义名称顺序：" + str(order))
            except KeyError:
                self.NameOrderDict[order] = word_name

            if word_body is not None and word_body.ExpressTime == 1:
                self.NameIndex += [word_body.Value]

        elif kind == "File":
            self.FilePool[word_name] = word_body
            try:
                test = self.FileOrderDict[order]
                ClassError.error("重复定义文件夹顺序：" + str(order))
            except KeyError:
                self.FileOrderDict[order] = word_name

    def get_temp_name(self, data: list) -> str:
        """DocxData.get_name调用对新的数据行求归属文件"""
        chars = ""

        for key in self.NameIndex:
            chars += "_0_" + str(data[key])

        return chars

    def name_init(self, line_data: xlsxmethod.ClassXlsxData):
        """设置名称的值,由DocxData.name_init调用
            多次引用确保n2t2
        """
        for key in self.NamePool.keys():
            if self.NamePool[key] is None:
                continue
            self.NamePool.take(key, self.NamePool[key].get_data(line_data))
        for key in self.FilePool.keys():
            if self.FilePool[key] is None:
                continue
            self.FilePool.take(key, self.FilePool[key].get_data(line_data))

    def finish(self):
        """DocxData.finish_data_input最后填满并得出结果"""
        for key in self.NamePool.keys():
            if self.NamePool[key] is None:
                self.NamePool.take(key, ClassMainProcessor.TIMING_COLD_WORD.value_(key))
        for key in self.FilePool.keys():
            if self.FilePool[key] is None:
                self.FilePool.take(key, ClassMainProcessor.TIMING_COLD_WORD.value_(key))

        # test
        if not (self.NamePool.translated() and self.FilePool.translated()):
            print("error name finish error")

        # make name string
        name = ""
        for i in range(len(self.NameOrderDict)):
            try:
                name += ClassNameFileBody.NAME_INTERVAL + self.NamePool.value_(self.NameOrderDict[i+1])
            except KeyError:
                print("error 命名顺序错误")

        name = name.replace(ClassNameFileBody.NAME_INTERVAL, "", 1)
        name += ".docx"

        # make filepath string
        file = ClassNameFileBody.FILE_ROOT_PATH
        for i in range(len(self.FileOrderDict)):
            try:
                file += self.FilePool.value_(self.FileOrderDict[i + 1]) + "/"
            except KeyError:
                print("error 地址顺序错误")

        return name, file


class ClassEnumBody:
    """运行时句子组管理单元

        对于需要枚举的变量的不同句子组的统一管理
            对不同数据组填入从原始接受体的分裂的不同实体中

    """
    def __init__(self):
        # 接受数据的接受体流 填充ClassInputBody
        self.InputStream = stream.ClassShadowStream()
        self.Translated = self.InputStream.Translated

        # 原始接受体
        self.PureInputStream = ClassInputBody()

        # 接受体地址对象
        self.InputAddress = None

    def get_data(self, line_data: xlsxmethod.ClassXlsxData):
        """从DocxData获取数据并遍历InputBody"""
        if ClassMainProcessor.XLSX_FINISHED:
            for key in self.InputStream.keys():
                if self.InputStream[key] is None:
                    continue
                self.InputStream.take(key, self.InputStream[key].get_data(line_data))

            if self.InputStream.translated():
                return self.InputStream
            else:
                ClassError.error("error EnumBody finish error")
                return None

        else:
            index = line_data.index(self.InputAddress)

            try:
                self.InputStream.take(index, self.InputStream[index].get_data(line_data))
            except KeyError:
                self.InputStream[index] = copy.deepcopy(self.PureInputStream)
                self.InputStream.take(index, self.InputStream[index].get_data(line_data))
            finally:
                if self.InputStream.translated():
                    return self.InputStream
                else:
                    return None


class ClassWordProcessor:
    """词单元的初始化处理器

        将来自句法处理单元原始词单元初始化
        对于词单元内容的填入将由词单元方法完成

    """
    # 用于分析词法的正则表达式，及其对应单元编号
    WordFixRegular = re.compile(r"\[(.*)(?=|)[|]?(?<=[|\[])(.*?[A-Z]+.*?)(?=[|\]])[|]?(?<=|)(.*)]")
    WordBodyRegular = re.compile(r"^(\d*?)([ecnf]?)(\d*?)([yta]?)(\d*)([A-Z]+)(\d*)(.*?)$")

    ERROR_SIGN = "Word Processor"

    WordBodyLineSet = 0
    WordBodyExpressTime = 2
    WordBodyNextLine = 4
    WordBodyValue = 5
    WordBodyWidth = 6
    WordBodyMethod = 7

    def __init__(self):
        # 类LR构造表，连同__acc接受函数组
        self.GrammarMap = [{"": 1, "number": 2},                                # state 0
                           {"c": 3, "n": 4, "f": 11, "e": 5, "": self.__acc5},  # state 1
                           {"e": 5},                                            # state 2
                           {"number": 6, "": 6},                                # state 3
                           {"number": 7},                                       # state 4
                           {"number": 9, "": 8},                                # state 5
                           {"t": self.__acc3},                                  # state 6
                           {"a": self.__acc4, "t": 13},                         # state 7
                           {"y": self.__acc1},                                  # state 8
                           {"y": 10},                                           # state 9
                           {"number": self.__acc2},                             # state 10
                           {"number": 12},                                      # state 11
                           {"a": 13},                                           # state 12
                           {"number": self.__acc6}]                             # state 13

        self.MapHelperWordNumberStockZero = 0b1  # 用于辅助解析词结构
        self.MapHelperWordType = 1  # 用于识别词的类型

        self.NameHelperFileType = "normal"
        self.NameHelperFileOrder = 0

        self.InputHelperWordCold = False

        self.MsgFirstLine = False

    # 分析主程序，由SentenceProcessor调用，完成对单个词单元的分析填入与
    def analysis(self, word_template: str, msg: ClassMsg,
                 sentence_data, name_body: ClassNameFileBody(),
                 cold_word: stream.ClassShadowStream):
        """SentenceProcessor.analysis调用"""
        def __str_to_int__(string) -> int:  # 只会匹到大写字符不存在其他
            _int = 0

            for ch in string:
                _int *= 26
                _int += ord(ch) - ord('A') + 1

            return _int - 1

        wordPortList = self.WordFixRegular.match(word_template).groups()
        # error with

        wordMsg = list(self.WordBodyRegular.match(wordPortList[1]).groups())
        # error with
        wordMsg[self.WordBodyValue] = __str_to_int__(wordMsg[self.WordBodyValue])

        wordBody = ClassWord(wordPortList[0][:-1], wordPortList[-1], word_template)
        wordBody.Value = wordMsg[self.WordBodyValue]
        wordBody.Method = wordMsg[self.WordBodyMethod]
        try:
            wordBody.Width = int(wordMsg[self.WordBodyWidth])
        except ValueError:
            wordBody.Width = 0

        self.__map_init()
        nowState = 0
        for element in wordMsg:
            try:
                nowState = self.GrammarMap[nowState][self.__map_key(element)]

            except TypeError:
                nowState(wordBody, wordMsg, msg)
                break

            except KeyError:
                ClassError.error(ClassWordProcessor.ERROR_SIGN + "word error" + word_template)
                break
                pass

        # error with

        # word填入
        if self.InputHelperWordCold:
            cold_word[word_template] = wordBody
            wordBody = None

        sentence_data.Data[msg.WordCount] = wordBody  # 编号为引索

        name_body.get(word_template, wordBody, self.NameHelperFileType, self.NameHelperFileOrder)

    def __acc1(self, word_body: ClassWord, word_msg: list, msg):
        """对ey的处理的方式

            将词放入冷堆

        :param word_body:
        :param word_msg:
        :param msg:
        :return:
        """
        if self.MapHelperWordNumberStockZero != 0b1:
            pass
        # error 防止e0y0

        # self.InputHelperWordCold = True

        word_body.Enum = True
        word_body.ExpressTime = -1

    def __acc2(self, word_body: ClassWord, word_msg: list, msg: ClassMsg):
        """对02e02y2,2e2y2,0e2 y的处理状态

        :param word_body:
        :param word_msg:
        :param msg:
        :return:
        """
        word_body.Enum = True
        msg.NowEnumBody = True
        fixType = 0

        # 处理倒数第三个数
        if self.MapHelperWordNumberStockZero & 0b1000 != 0:
            if self.MapHelperWordNumberStockZero & 0b100 != 0:
                msg.NewEnumBody += 1
                if msg.NewEnumBody >= 3:
                    ClassError.error(ClassWordProcessor.ERROR_SIGN + "error 嵌套ey*" + word_body.PureValue)

            # 是否首次定义句子
            if self.MsgFirstLine:
                msg.SentenceNumber = int(word_msg[self.WordBodyLineSet])
                if int(word_msg[self.WordBodyLineSet]) == 0:
                    ClassError.error(ClassWordProcessor.ERROR_SIGN + "error 句子设置为0" + word_body.PureValue)
                self.MsgFirstLine = False
        # end

        # 处理倒数第二个数
        if self.MapHelperWordNumberStockZero & 0b010 != 0:
            fixType = 1
            if int(word_msg[self.WordBodyExpressTime]) == 0:
                ClassError.error(ClassWordProcessor.ERROR_SIGN + "error 枚举次数为0" + word_body.PureValue)
        word_body.ExpressTime = int(word_msg[self.WordBodyExpressTime])
        # end

        # 处理倒数第一个数
        word_body.NextSentence = int(word_msg[self.WordBodyNextLine])
        if word_body.NextSentence == 0:
            word_body.NextSentence = None

        if word_body.NextSentence == 1 \
                and msg.SentenceNumber == 1\
                and not self.MapHelperWordNumberStockZero & 0b1:
            msg.NewEnumBody += 1
            if msg.NewEnumBody >= 3:
                ClassError.error(ClassWordProcessor.ERROR_SIGN + "error 嵌套ey" + word_body.PureValue)
        # end

        # 处理地址
        if msg.SentenceNumber == 1:
            if fixType == 1:
                xlsxmethod.ClassDataAddress.add(word_msg[self.WordBodyValue], word_body.ExpressTime * -1)
            else:
                xlsxmethod.ClassDataAddress.add(word_msg[self.WordBodyValue], word_body.ExpressTime)

    def __acc3(self, word_body: ClassWord, word_msg: list, msg):
        """对c02t,c2t,ct的处理状态

        :param word_body:
        :param word_msg:
        :param msg:
        :return:
        """
        if self.MapHelperWordNumberStockZero & 0b10:
            word_body.ExpressTime = int(word_msg[self.WordBodyExpressTime])

            if self.MapHelperWordNumberStockZero & 0b11:
                if word_body.ExpressTime == 0:
                    word_body.ExpressTime = -1
                else:
                    word_body.ExpressTime *= 1
                self.InputHelperWordCold = True
        else:
            word_body.ExpressTime = 1

    def __acc4(self, word_body: ClassWord, word_msg: list, msg):
        """对n2a的处理方式

        :param word_body:
        :param word_msg:
        :param msg:
        :return:
        """
        if self.MapHelperWordNumberStockZero != 0b10:
            ClassError.error(ClassWordProcessor.ERROR_SIGN + "error 非n2a" + word_body.PureValue)

        self.NameHelperFileType = "Name"
        self.NameHelperFileOrder = int(word_msg[self.WordBodyExpressTime])

    def __acc5(self, word_body: ClassWord, word_msg: list, msg):
        """对A的接受状态

        :param word_body:
        :param word_msg:
        :param msg:
        :return:
        """
        if (word_msg[2] + word_msg[3] + word_msg[4]) != "":
            ClassError.error(ClassWordProcessor.ERROR_SIGN + "error 非A" + word_body.PureValue)

        word_body.ExpressTime = 1

    def __acc6(self, word_body: ClassWord, word_msg: list, msg):
        """对n2t2,f2a2的接受状态

        :param word_body:
        :param word_msg:
        :param msg:
        :return:
        """
        self.NameHelperFileType = {"n": "Name", "f": "File"}[word_msg[self.MapHelperWordType]]

        # 匹配到两个数字
        word_body.ExpressTime = int(word_msg[self.WordBodyExpressTime])

        if self.MapHelperWordNumberStockZero & 0b10:
            if word_body.ExpressTime == 0:
                word_body.ExpressTime = -1
            else:
                word_body.ExpressTime *= -1
            self.InputHelperWordCold = True
        else:
            word_body.ExpressTime = 1

        self.NameHelperFileOrder = int(word_msg[self.WordBodyNextLine])
        if self.NameHelperFileOrder == 0:
            pass
        # error 顺序为0

    def __map_key(self, string: str) -> str:
        try:
            value = int(string)

            self.MapHelperWordNumberStockZero <<= 1
            if string[0] == "0":
                self.MapHelperWordNumberStockZero += 1

            return "number"

        except ValueError:
            return string

    def __map_init(self):
        self.MapHelperWordNumberStockZero = 0b1
        self.MapHelperWordType = 1

        self.NameHelperFileType = "normal"
        self.NameHelperFileOrder = 0

        self.InputHelperWordCold = False

    def new_sentence(self):
        """由SentenceProcessor调用，防止重复定义句子"""
        self.MsgFirstLine = True


class ClassSentenceProcessor:
    """句子单元的初始化处理器

        将原始句子单元初始化

    """
    WordRegular = re.compile(r"(\[.*?])")
    WordProcessor = ClassWordProcessor()

    def __init__(self):
        pass

    def analysis(self, sentence: doctemplate.ClassSentences,
                 msg: ClassMsg,
                 # input_body: ClassInputBody(),
                 name_body: ClassNameFileBody(),
                 cold_word: stream.ClassShadowStream):
        """MainProcessor.pretreatment调用"""

        wordList = self.WordRegular.findall(sentence.Text)

        sentenceData = ClassSentenceWithData(sentence, wordList)

        # 当前输入是否是真枚举体, 用于枚举体转入非枚举体
        msgNowEnum = bool(msg.NowEnumBody)
        msg.NowEnumBody = False

        self.WordProcessor.new_sentence()
        for wordTemplate in wordList:
            self.WordProcessor.analysis(wordTemplate, msg, sentenceData, name_body, cold_word)
            msg.WordCount += 1

        # input_body.SentenceStream[msg.SentenceNumber] = sentenceData
        if msgNowEnum and not msg.NowEnumBody:
            msg.NewEnumBody += 1

        return sentenceData


class ClassDocxData:
    def __init__(self):
        self.EnumBodyStream = stream.ClassShadowStream()
        self.ColdWordStream = stream.ClassShadowStream()  # 填入冷数据词
        self.NameBody = ClassNameFileBody()  # 用于装名称路径的输入体

    def get_enum(self, enum_body, order):
        """MainProcessor.pretreatment调用填入EnumBody"""
        self.EnumBodyStream[order] = enum_body

    def get_name(self, data: list) -> str:
        """MainProcessor.data_filling调用返回数据归属的文件名"""
        return self.NameBody.get_temp_name(data)

    def name_init(self, line_data: xlsxmethod.ClassXlsxData):
        """ClassDocxData.get_template&get_data调用传递给NameFileBody.name_init"""
        self.NameBody.name_init(line_data)

    def get_data(self, line_data: xlsxmethod.ClassXlsxData):
        """ClassDocxData.get_data调用并遍历传递EnumBodyStream"""
        for key in self.EnumBodyStream.keys():
            if self.EnumBodyStream[key] is None:
                continue
            self.EnumBodyStream.take(key, self.EnumBodyStream[key].get_data(line_data))

    def finish_data_input(self, data: xlsxmethod.ClassXlsxData):
        """ClassDocxData.finish_data_input调用并作处理"""
        for word in self.ColdWordStream.keys():
            # code test
            if self.ColdWordStream[word] is None:
                print("error cold word stream has None")
                continue
            # code test end
            self.ColdWordStream.take(word, self.ColdWordStream[word].get_data(data))
        if self.ColdWordStream.translated():
            pass
        else:
            print("error")

        ClassMainProcessor.TIMING_COLD_WORD = self.ColdWordStream

        for key in self.EnumBodyStream.keys():
            if self.EnumBodyStream[key] is None:
                continue
            self.EnumBodyStream.take(key, self.EnumBodyStream[key].get_data(data))

        # test
        if not self.EnumBodyStream.translated():
            print("error EnumBody finish error")
        # test end

        name, file = self.NameBody.finish()
        return name, file


class ClassMainProcessor:
    """主处理器

        负责启动模板初始化；负责启动数据的填入

    """

    XLSX_FINISHED = False
    TIMING_COLD_WORD = None

    def __init__(self):
        pass

    def pretreatment(self, template):

        dataBody = ClassDocxData()

        # 初始化枚举体流
        countEnumBody = 1   # 枚举体索引
        currentEnumBody = ClassEnumBody()    # 当前枚举体
        currentAddress = xlsxmethod.ClassDataAddress()  # 当前的输入体地址表达
        currentInputBody = ClassInputBody()  # 用于填入的输入体
        # initial end

        # 初始化名称表达与冷数据
        coldWordStream = dataBody.ColdWordStream  # 填入冷数据词
        nameBody = dataBody.NameBody   # 用于装名称路径的输入体
        # initial end

        # 初始化语法分析器
        sentenceProcessor = ClassSentenceProcessor()
        # initial end

        # 初始化信使
        msg = ClassMsg()
        msg.new()
        # initial end

        for sentences in template.TemplateSentenceList:
            sentenceData = sentenceProcessor.analysis(sentences,
                                                      msg,
                                                      # currentInputBody,
                                                      nameBody,
                                                      coldWordStream)

            if msg.NewEnumBody > 0:
                currentEnumBody.PureInputStream = currentInputBody
                currentEnumBody.InputAddress = currentAddress

                dataBody.get_enum(currentEnumBody, countEnumBody)
                countEnumBody += 1

                currentEnumBody = ClassEnumBody()  # 当前枚举体
                currentAddress = xlsxmethod.ClassDataAddress()  # 当前的输入体地址表达
                currentInputBody = ClassInputBody()  # 用于填入的输入体\
                msg.new()

            currentInputBody.set_sentence(msg.SentenceNumber,  sentenceData)
            msg.next()

        currentEnumBody.PureInputStream = currentInputBody
        currentEnumBody.InputAddress = currentAddress

        dataBody.get_enum(currentEnumBody, countEnumBody)
        countEnumBody += 1

        template.DocxTemplateData = dataBody

    def data_filling(self, template: doctemplate.ClassDocxStreamWithTemplate, xlsx_body: xlsxmethod.ClassXlsxSource):
        for line in xlsx_body.read_all_lines():
            indexName = template.DocxTemplateData.get_name(line)

            try:
                docxBody = template.DocxStream[indexName]
            except KeyError:
                docxBody = copy.deepcopy(template.PureTemplate)
                templateData = copy.deepcopy(template.DocxTemplateData)
                docxBody.get_template(templateData, xlsxmethod.ClassXlsxData(line))
                template.DocxStream[indexName] = docxBody
            finally:
                docxBody = template.DocxStream[indexName]
                docxBody.get_data(line)

        ClassWord.CON_NEXT_SENTENCE = -1
        ClassMainProcessor.XLSX_FINISHED = True

        for key in template.DocxStream.keys():
            template.DocxStream[key].finish_data_input()


if __name__ == "__main__":  # test
    docx = doctemplate.ClassDocxStreamWithTemplate("./data/table2.docx")
    xlsx = xlsxmethod.ClassXlsxSource("./data/test_xlsx2.xlsx")

    process = ClassMainProcessor()

    process.pretreatment(docx)

    process.data_filling(docx, xlsx)

    docx.fill_in_docx()
    docx.output([])

    # 输出测试
    # docx.DocxStream["_0_俞振炀_0_2017217193"].DataDocx.EnumBodyStream.test_show()
    # docx.DocxStream["_0_俞振炀_0_2017217193"].DataDocx.EnumBodyStream.value_(1).test_show()
    # docx.DocxStream["_0_俞振炀_0_2017217193"].DataDocx.EnumBodyStream.value_(1).value_("").test_show()
    # print(docx.DocxStream["_0_俞振炀_0_2017217193"].DataDocx.EnumBodyStream.value_(1).value_("").value_(1).Result)
    # print(docx.DocxStream["_0_俞振炀_0_2017217193"].DataDocx.EnumBodyStream.value_(1).value_("").value_(2).Result)
    # print(docx.DocxStream["_0_俞振炀_0_2017217193"].DataDocx.EnumBodyStream.value_(1).value_("").value_(5).Result)
    docx.DocxStream["_0_俞振炀_0_2017217193"].DataDocx.EnumBodyStream.test_show()
    docx.DocxStream["_0_俞振炀_0_2017217193"].DataDocx.EnumBodyStream.value_(2).test_show()
    docx.DocxStream["_0_俞振炀_0_2017217193"].DataDocx.EnumBodyStream.value_(2).value_("1").test_show()
    print(docx.DocxStream["_0_俞振炀_0_2017217193"].DataDocx.EnumBodyStream.value_(2).value_("1").value_(1).Result)
    print(docx.DocxStream["_0_俞振炀_0_2017217193"].DataDocx.EnumBodyStream.value_(2).value_("1").value_(2).Result)
    print(docx.DocxStream["_0_俞振炀_0_2017217193"].DataDocx.EnumBodyStream.value_(2).value_("1").value_('add1').Result)

    # print(docx.DocxStream["_0_俞振炀_0_2017217193"].DataDocx.EnumBodyStream.value_(2).value_("1").value_('add2').Result)
    # docx.DocxStream["_0_俞振炀_0_2017217193"].DataDocx.EnumBodyStream[2].InputStream["3"].SentenceStream.test_show()
    # docx.DocxStream["_0_俞振炀_0_2017217193"].DataDocx.EnumBodyStream.value_(2).SentenceStream.value_(2).Data.test_show()

    # 输入测试
    # print("docx.DocxTemplateData.EnumBodyStream")
    # docx.DocxTemplateData.EnumBodyStream.test_show()
    #
    # print("docx.DocxTemplateData.ColdWordStream")
    # docx.DocxTemplateData.ColdWordStream.test_show()

    # print("docx.DocxTemplateData.NameBody.NameIndex")
    # docx.DocxTemplateData.NameBody.NameIndex.test_show()
    #
    # print("docx.DocxTemplateData.NameBody.NamePool")
    # docx.DocxTemplateData.NameBody.NamePool.test_show()
    #
    # print("docx.DocxTemplateData.NameBody.NamePool")
    # docx.DocxTemplateData.NameBody.NamePool.test_show()
    # print("docx.DocxTemplateData.NameBody.NamePool")
    # docx.DocxStream["_0_姓名_0_学号"].DataDocx.NameBody.NamePool.test_show()
    #
    # docx.DocxTemplateData.EnumBodyStream[2].PureInputStream.SentenceStream.test_show()
    #
    # print(docx.DocxTemplateData.EnumBodyStream[2].PureInputStream.SentenceStream[2].Sentence.Text)
    # docx.DocxTemplateData.EnumBodyStream[2].PureInputStream.SentenceStream[1].Data.test_show()
    #
    # print(docx.DocxTemplateData.EnumBodyStream[2].InputAddress.Path)


