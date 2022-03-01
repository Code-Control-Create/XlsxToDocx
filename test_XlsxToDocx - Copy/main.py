import os
import shutil
import sys
import getopt


import processor
import xlsxmethod
import doctemplate
import error


class ClassConsole:

    HELP_INFO_ZH = "    ！！！XlsxToDocx！！！\n" \
                   "    Version: 0.5.0 Bate (厕逝版)\n" \
                   "    可以将已有的Xlsx数据通过Docx文件模板生成格式化文件\n"\
                   "   -d        --docx=      模板docx的路径(文件名)\n" \
                   "   -x        --xlsx=      数据源xlsx的路径(文件名)\n" \
                   "   -i    --interval=      输出文件名中的间隔\n" \
                   "   _t         --test      模板测试\n" \
                   "    更多帮助请联系Q群：686545133\n"

    @staticmethod
    def _help():
        print(ClassConsole.HELP_INFO_ZH)

    def __init__(self):
        self.DocxPath = None
        self.XlsxPath = None
        self.NameInterval = "_"
        self.NowPath = None
        self.ResultPath = "result/"
        self.ResultPathList = []

    def init(self):
        opts, args = getopt.getopt(sys.argv[1:], "htd:x:i:", ["help", "docx=", "xlsx=", "interval=", "test"])
        for opt_name, opt_value in opts:
            if opt_name in ("-h", "--help"):
                ClassConsole._help()
                sys.exit()
            elif opt_name in ("-d", "--docx"):
                self.DocxPath = opt_value
                continue
            elif opt_name in ("-x", "--xlsx"):
                self.XlsxPath = opt_value
                continue
            elif opt_name in ("-i", "--interval"):
                error.ClassError.TEST = True
                continue
            elif opt_name in ("-t", "--test"):
                self.NameInterval = opt_value
                continue

        if (self.DocxPath is not None) and (self.XlsxPath is not None) and (self.NameInterval is not None):
            self.NowPath = os.getcwd()
            try:
                os.mkdir(self.ResultPath)
            except FileExistsError:
                shutil.rmtree(self.ResultPath)
                os.mkdir(self.ResultPath)
            self.define()

        else:
            ClassConsole._help()
            sys.exit()

    def define(self):
        processor.ClassNameFileBody.NAME_INTERVAL = self.NameInterval
        processor.ClassNameFileBody.FILE_ROOT_PATH = self.ResultPath
        error.ClassError.init(self.ResultPath)

    def start(self):
        docx = doctemplate.ClassDocxStreamWithTemplate(self.DocxPath)
        xlsx = xlsxmethod.ClassXlsxSource(self.XlsxPath)

        process = processor.ClassMainProcessor()

        process.pretreatment(docx)

        process.data_filling(docx, xlsx)

        docx.fill_in_docx()

        docx.output(self.ResultPathList)

        error.ClassError.end()


if __name__ == "__main__":
    main = ClassConsole()

    main.init()

    main.start()
