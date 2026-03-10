import copy
import tempfile
import openpyxl
import os
import re
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill
import sys
import shutil
import uuid

class config_t:
    config_inputFilePath = "\input"
    config_outputFilePath = "\output"
    config_changesInfoFilePath = config_inputFilePath + "\signalInfoInput.xlsx"
    config_dbcFilePath = config_inputFilePath
    config_logFilePath = config_outputFilePath + "\log.xlsx"
    config_debugMode = False
    config_refWindowNum = 15

class miscFuncs_t:
    class queue:
        def __init__(self):
            self.items = []  # 使用列表存储队列元素
            self.unique_set = set()  # 使用集合跟踪已存在的元素

        def enqueue(self, item):
            if item not in self.unique_set:
                self.items.append(item)
                self.unique_set.add(item)
                return True  # 表示成功添加
            return False  # 表示元素已存在，未添加

        def dequeue(self):
            if self.is_empty():
                raise IndexError("队列为空")
            item = self.items.pop(0)
            self.unique_set.remove(item)
            return item

        def front(self):
            if self.is_empty():
                raise IndexError("队列为空")
            return self.items[0]

        def is_empty(self):
            return len(self.items) == 0

        def size(self):
            return len(self.items)

        def __str__(self):
            return f"队列元素: {self.items}"

    class color:
        # 基本颜色
        BLACK = "\033[30m"
        RED = "\033[31m"
        GREEN = "\033[32m"
        YELLOW = "\033[33m"
        BLUE = "\033[34m"
        MAGENTA = "\033[35m"
        CYAN = "\033[36m"
        WHITE = "\033[37m"

        # 亮色（高对比度）
        BRIGHT_BLACK = "\033[90m"
        BRIGHT_RED = "\033[91m"
        BRIGHT_GREEN = "\033[92m"
        BRIGHT_YELLOW = "\033[93m"
        BRIGHT_BLUE = "\033[94m"
        BRIGHT_MAGENTA = "\033[95m"
        BRIGHT_CYAN = "\033[96m"
        BRIGHT_WHITE = "\033[97m"

        # 背景色
        BG_BLACK = "\033[40m"
        BG_RED = "\033[41m"
        BG_GREEN = "\033[42m"
        BG_YELLOW = "\033[43m"
        BG_BLUE = "\033[44m"
        BG_MAGENTA = "\033[45m"
        BG_CYAN = "\033[46m"
        BG_WHITE = "\033[47m"

        # 亮背景色
        BG_BRIGHT_BLACK = "\033[100m"
        BG_BRIGHT_RED = "\033[101m"
        BG_BRIGHT_GREEN = "\033[102m"
        BG_BRIGHT_YELLOW = "\033[103m"
        BG_BRIGHT_BLUE = "\033[104m"
        BG_BRIGHT_MAGENTA = "\033[105m"
        BG_BRIGHT_CYAN = "\033[106m"
        BG_BRIGHT_WHITE = "\033[107m"

        # 样式
        RESET = "\033[0m"
        BOLD = "\033[1m"
        UNDERLINE = "\033[4m"
        REVERSE = "\033[7m"  # 反显（前景背景互换）

        # 配色方案
        INFO_STYLE = BRIGHT_GREEN + BOLD  # 亮绿色加粗
        DEBUG_STYLE = BRIGHT_CYAN + BOLD  # 亮青色加粗
        ERROR_STYLE = BRIGHT_RED + BOLD  # 亮红色加粗
        WARNING_STYLE = BRIGHT_YELLOW + BOLD  # 亮黄色加粗（可选）


    # 使能终端输出颜色字符
    def enableAnsiSupport(self):
        if sys.platform.lower() == "win32":
            os.system("")
            import ctypes
            kernel32 = ctypes.windll.kernel32
            kernel32.SetConsoleMode(kernel32.GetStdHandle(-11), 7)

    def __init__(self):
        self.warningCnt = 0
        self.errorCnt = 0
        self.infoCnt = 0
        self.errorQueue = self.queue()
        self.warningQueue = self.queue()
        self.prograss_length = 50
        self.prograssTotal = 0
        self.prograssCnt = 0


    def notification(self, type, str, errorFunctionName=None, prograss_1=None):
        if type == 'INFO':
            self.infoCnt += 1
            print(f"{self.color.INFO_STYLE}INFO {str}{self.color.RESET}")
        elif type == 'DEBUG':
            print(f"{self.color.DEBUG_STYLE}LOG {self.color.RESET}{str}{self.color.RESET}", end='')
        elif type == 'ERROR':
            rtn = self.errorQueue.enqueue(str)
            if rtn is True:
                self.errorCnt += 1
            print(f"{self.color.ERROR_STYLE}ERROR {str}{self.color.RESET}")
            if errorFunctionName is not None:
                print(f"{self.color.ERROR_STYLE} ErrorFunctionName: {errorFunctionName}{self.color.RESET}")
        elif type == 'FATAL':
            print(f"{self.color.ERROR_STYLE}FATAL! {str}{self.color.RESET}")
            if errorFunctionName is not None:
                print(f"{self.color.ERROR_STYLE} ErrorFunctionName: {errorFunctionName}{self.color.RESET}")
            self.getUserInput("输入任意键结束:")
            sys.exit(2)
        elif type == 'WARNING':
            rtn = self.warningQueue.enqueue(str)
            if rtn is True:
                self.warningCnt += 1
            #print(f"{self.color.WARNING_STYLE}WARNING: {str}{self.color.RESET}")
        elif type == 'TERMINAL_WARNING':
            print(f"{self.color.WARNING_STYLE}{str}{self.color.RESET}")
        elif type == 'TERMINAL_ERROR':
            print(f"{self.color.ERROR_STYLE}{str}{self.color.RESET}")
        elif type == 'LODING':
            percent = float(str)

            bar_length = int(self.prograss_length * percent)
            bar = '█' * bar_length
            spaces = ' ' * (self.prograss_length - bar_length)

            bar_str = f"{self.color.INFO_STYLE}{bar}{spaces}{self.color.RESET}"
            percent_str = f"{self.color.INFO_STYLE}{percent * 100:.1f}%"

            sys.stdout.write(f"\r{self.color.INFO_STYLE}INFO {bar_str} {percent_str} {prograss_1}{self.color.RESET}")
            sys.stdout.flush()

    def printTerminalInfo(self):
        print("")                                                       # 这个print应当写在mainTask中维持UI，藏在这里显得mainTask美观
        if (miscFuncs.warningCnt == 0) and (miscFuncs.errorCnt == 0):
            miscFuncs.notification('INFO', f"WARNING: {miscFuncs.warningCnt}  ERROR: {miscFuncs.errorCnt}")
        else:
            miscFuncs.notification('TERMINAL_WARNING', f"WARNING: {miscFuncs.warningCnt}  ERROR: {miscFuncs.errorCnt}")
            for i in range(0, miscFuncs.errorQueue.size(), 1):
                miscFuncs.notification('TERMINAL_ERROR', f"ERROR: {miscFuncs.errorQueue.dequeue()}")
            for i in range(0, miscFuncs.warningQueue.size(), 1):
                miscFuncs.notification('TERMINAL_WARNING', f"WARNING: {miscFuncs.warningQueue.dequeue()}")

    def waitForKey(str):
        return input(str)

    def getUserInput(self, str):
        print("\n")
        while True:
            rtn = miscFuncs_t.waitForKey(str)
            if rtn is not None:
                break
        return rtn

    # 获取进度条总进度值
    def getProgressTotal(self):
        miscFuncs.prograssTotal = len(inputFuncs.signalDetailInfo) * 12
        for newSignalInfo in inputFuncs.signalDetailInfo:
            if newSignalInfo['changeType'] == "调整信号信息":
                miscFuncs.prograssTotal += 12
            elif newSignalInfo["changeType"] == "新增报文":
                miscFuncs.prograssTotal -= 12

class inputFuncs_t:
    SIGNAL_NUM = "Num"
    SIGNAL_CAN = "总线"
    SIGNAL_MESSAGE_ID = "报文ID"
    SIGNAL_SEND_POINT = "发送节点"
    SIGNAL_RECEIVE_POINT = "接收节点"
    SIGNAL_ENG_NAME = "信号名称"
    SIGNAL_CHA_NAME = "信号名称_2"
    SIGNAL_CHANGE_TYPE = "变更类型"
    SIGNAL_FRONT_CHANGE = "前缀替换"
    SIGNAL_REFERENCE = "参考信号"

    def __init__(self):
        self.columnIndices = {
            inputFuncs_t.SIGNAL_NUM: None,
            inputFuncs_t.SIGNAL_CAN: None,
            inputFuncs_t.SIGNAL_MESSAGE_ID: None,
            inputFuncs_t.SIGNAL_SEND_POINT: None,
            inputFuncs_t.SIGNAL_RECEIVE_POINT: None,
            inputFuncs_t.SIGNAL_ENG_NAME: None,
            inputFuncs_t.SIGNAL_CHA_NAME: None,
            inputFuncs_t.SIGNAL_CHANGE_TYPE: None,
            inputFuncs_t.SIGNAL_FRONT_CHANGE: None,
            inputFuncs_t.SIGNAL_REFERENCE: None,
        }
        self.inputDicts = []
        self.signalDetailInfo = []
        self.signalDetailInfoAll = []

    # 解析excel表头的行坐标
    def __getColumnIndex(self, sheet):
        sheetRows = list(sheet.rows)
        for cell in sheetRows[0]:
            if cell.value in self.columnIndices:
                self.columnIndices[cell.value] = cell.column - 1
                #print(f"{cell.value}found in colum {cell.column}")

    def __getMessageIDandTransPoint(self, value):
        parts = None

        if '_' in value:
            parts = value.split('_')
        elif ' ' in value:
            parts = value.split(' ')
        else:
            print("Input massage format invalid!")

        if parts is not None:

            return parts[len(parts)-1], parts[0]
        else:
            return None, None

    def __getMessageID(self, value):
        return value

    def __getReceivePoint(self, value):
        return value

    def __getMessageCAN(self, value):
        return value

    def __getSendPoint(self, value):
        return value

    def __getSignalEngName(self, value):
        return value

    def __getSignalChaName(self, value):
        return value

    def __getChangeType(self, value):
        return value

    def __getDBCFile(self, pyToolCurrentPath):
        dbcFilePath = pyToolCurrentPath + config.config_dbcFilePath

        dbcFiles = [file for file in os.listdir(dbcFilePath) if file.lower().endswith('.dbc') and
                    os.path.isfile(os.path.join(dbcFilePath, file))]

        if not dbcFiles:
            miscFuncs.notification('ERROR', "当前目录无dbc文件！")
            return

        with tempfile.NamedTemporaryFile(mode='wb', delete=False, suffix='.tmp', dir=dbcFilePath) as tempFile:
            tempFilePath = tempFile.name
            if config.config_debugMode == True:
                miscFuncs.notification('INFO', f"创建临时文件 {tempFilePath}")

            for fileName in dbcFiles:
                filePath = os.path.join(dbcFilePath, fileName)
                try:
                    with open(filePath,'r', encoding='utf-8', errors='replace') as dbcFile:
                        if config.config_debugMode == True:
                            miscFuncs.notification('INFO', f"处理DBC文件 {fileName}")
                        content = dbcFile.read()

                    header = f"\n\n===== 文件: {os.path.basename(fileName)} =====\n\n".encode('utf-8')
                    tempFile.write(header)
                    tempFile.write(content.encode('utf-8'))

                    if content is not None:
                        if config.config_debugMode == True:
                            miscFuncs.notification('INFO', f"已追加文件 {fileName}")
                    else:
                        miscFuncs.notification('ERROR', f"追加文件错误 {fileName}")
                except Exception as e:
                    miscFuncs.notification('ERROR', f"处理{fileName}时出错:{str(e)}")

        return tempFilePath

    # 校验参考信号是否在ARXML文件中存在
    def __checkReferenceSignalExcite(self, pyToolCurrentPath, referenceSignal, signalEnglishName, signalChineseName):
        arxmlFilePath = pyToolCurrentPath + config.config_inputFilePath + "\MCU_SYSTEM.arxml"
        referenceSignalFind = 0
        # 检查参考信号在代码中是否存在
        try:
            with open(arxmlFilePath, 'r', encoding='utf-8', errors='replace') as arxmlFile:
                for line in arxmlFile:
                    if referenceSignal in line:
                        # miscFuncs.notification('INFO',f"参考信号存在")
                        referenceSignalFind = 1
                        break

        except Exception as e:
            miscFuncs.notification('ERROR', f"处理{arxmlFilePath}时出错:{str(e)}")

        if referenceSignalFind == 0:
            miscFuncs.notification('ERROR',
                                   f"{signalEnglishName}:{signalChineseName} 在ARXML中未找到与DBC同名的参考信号，检查log文件中参考信号是否在ARXML中存在！")

        return referenceSignalFind

    # 信号前缀替换
    def __changeSignalFront(self, signalFrontChange, referenceSignal, signalEnglishName):
        newSignal = {"referenceSignal":None, "signalEnglishName":None}

        if signalFrontChange is not None:
            parts = referenceSignal.split('_')
            if len(parts) >= 2:
                new_parts = [signalFrontChange]  # 用新前缀替换第一部分
                # 循环添加剩余部分（从索引1开始）
                for i in range(1, len(parts)):
                    new_parts.append(parts[i])
                newSignal["referenceSignal"] = '_'.join(new_parts)

            parts = signalEnglishName.split('_')
            if len(parts) >= 2:
                new_parts = [signalFrontChange]  # 用新前缀替换第一部分
                # 循环添加剩余部分（从索引1开始）
                for i in range(1, len(parts)):
                    new_parts.append(parts[i])
                newSignal["signalEnglishName"] = '_'.join(new_parts)

        return newSignal

    # 从拼接的DBC文件中检索所需信号的参考信号
    def __getReferenceInfo(self, dbcFilePath, messageID):
        referenceInfo = {"message":None, "signal":None}
        referenceInfo["message"] = messageID
        findMessageFlag = 0
        testCnt = 0

        try:
            with open(dbcFilePath, 'r', encoding='utf-8', errors='replace') as dbcFile:

                for line in dbcFile:
                    if line.startswith("BO_"):
                        parts = line.split()
                        dbc_sendPoint = parts[1]

                        if dbc_sendPoint == str(messageID):
                            findMessageFlag = 1
                            continue

                    if findMessageFlag == 1 and line.startswith(" SG_"):
                        findMessageFlag = 0
                        testCnt += 1
                        parts = line.split()
                        referenceInfo["signal"] = parts[1]
                        break

        except Exception as e:
            miscFuncs.notification('ERROR', f"处理{dbcFilePath}时出错:{str(e)}")

        return referenceInfo

    def __getSignalUpperLimit(self, signalBitLen):
        info = {"signalUpperLimit": None, "signalType":None}
        info["signalUpperLimit"] = str(2**int(signalBitLen) - 1)

        if int(signalBitLen) > 0 and int(signalBitLen) <= 8:
            info["signalType"] = "UINT8"
        elif int(signalBitLen) > 8 and int(signalBitLen) <= 16:
            info["signalType"] = "UINT16"
        elif int(signalBitLen) > 16 and int(signalBitLen) <= 32:
            info["signalType"] = "UINT32"
        elif int(signalBitLen) > 32 and int(signalBitLen) <= 64:
            info["signalType"] = "UINT64"

        return info

    # 从DBC文件读取信号的起始位，长度等信息
    def __getSignalInfo(self, tempFilePath, signalEnglishName):
        signalInfo = {"signalStartBit": None, "signalBitLen": None, "signalUpperLimit":None, "signalType":None}

        try:
            with open(tempFilePath, 'r', encoding='utf-8', errors='replace') as dbcFile:
                for line in dbcFile:
                    #正则表达式全词匹配
                    pattern = r'\b' + re.escape(signalEnglishName) + r'\b'

                    if line.startswith(" SG_") and re.search(pattern, line):
                        parts = line.split()
                        signalInfo["signalStartBit"], signalInfo["signalBitLen"] = parts[3].split('|')
                        signalInfo["signalBitLen"] = signalInfo["signalBitLen"].split('@')[0]
                        signalInfo["signalType"] = self.__getSignalUpperLimit(signalInfo["signalBitLen"])["signalType"]
                        signalInfo["signalUpperLimit"] = self.__getSignalUpperLimit(signalInfo["signalBitLen"])["signalUpperLimit"]

        except Exception as e:
            miscFuncs.notification('ERROR', f"处理{tempFilePath}时出错:{str(e)}")

        return signalInfo

    def __transfer2signalInfo(self, pyToolCurrentPath):
        dbcFilePath = self.__getDBCFile(pyToolCurrentPath)

        for index, dict in enumerate(self.inputDicts, 0):
            messageID = None
            messageID_dec = None
            messageCAN = None
            messageSendPoint = None
            messageReceivePoint = None
            signalEnglishName = None
            signalChineseName = None
            changeType = None
            changeDetail = None
            currentECU = None
            transmitDirection = None
            referenceSignal = None
            referenceSignal_1 = None
            signalStartBit = None
            signalBitLen = None
            signalType = None
            signalUpperLimit = None
            signalFrontChange = None
            changePointID = None
            isChanged = "NULL"

            for key, value in dict.items():
                if key == inputFuncs_t.SIGNAL_MESSAGE_ID and value is not None:
                    messageID = self.__getMessageID(value)
                elif key == inputFuncs_t.SIGNAL_CAN and value is not None:
                    messageCAN = self.__getReceivePoint(value)
                elif key == inputFuncs_t.SIGNAL_RECEIVE_POINT and value is not None:
                     messageReceivePoint = self.__getMessageCAN(value)
                elif key == inputFuncs_t.SIGNAL_SEND_POINT and value is not None:
                    messageSendPoint = self.__getSendPoint(value)
                elif key == inputFuncs_t.SIGNAL_ENG_NAME and value is not None:
                    signalEnglishName = self.__getSignalEngName(value)
                elif key == inputFuncs_t.SIGNAL_CHA_NAME and value is not None:
                    signalChineseName = self.__getSignalChaName(value)
                elif key == inputFuncs_t.SIGNAL_CHANGE_TYPE and value is not None:
                    changeType = self.__getChangeType(value)
                elif key == inputFuncs_t.SIGNAL_FRONT_CHANGE and value is not None:
                    signalFrontChange = value
                elif key == inputFuncs_t.SIGNAL_NUM and value is not None:
                    changePointID = value
                elif key == inputFuncs_t.SIGNAL_REFERENCE and value is not None:
                    referenceSignal = value

            if changeType == "新增信号" or changeType == "调整信号信息":
                # 若referenceSignal输入为空，从DBC中解析参考信号
                if referenceSignal is None:
                    # 以报文id的十进制数做dbc报文的索引
                    messageID_dec = int(messageID, base=16)
                    referenceSignal = self.__getReferenceInfo(dbcFilePath, messageID_dec)["signal"]

                # 检查DBC中参考信号是否存在
                if referenceSignal is None:
                    miscFuncs.notification('FATAL', f"参考信号不存在！检查DBC文件中是否存在{messageID}报文\n")
                    return 0

                if changeType == "调整信号信息":
                    referenceSignal_1 = signalEnglishName
            elif changeType == "删除信号":
                referenceSignal = signalEnglishName
            elif changeType == "调整信号名称":
                # 调整信号名称时，参考信号输入应当是旧信号名，如果未输入，则报错
                if referenceSignal is None:
                    isChanged = "ERROR"
                    miscFuncs.notification('ERROR', f"调整 {signalEnglishName}：{signalChineseName} 信号名称，参考信号应当输入旧信号名称！当前参考输入为空。")
                    continue

            signalInfo = self.__getSignalInfo(dbcFilePath, signalEnglishName)
            #删除信号，调整名称不用解析信号信息
            if changeType == "新增信号" or changeType == "调整信号信息":
                if signalInfo["signalStartBit"] is None:
                    miscFuncs.notification('ERROR', f"解析{signalEnglishName}信号起始位出错！")
                    isChanged = "ERROR"
                else:
                    signalStartBit = signalInfo['signalStartBit']
                if signalInfo["signalBitLen"] is None:
                    miscFuncs.notification('ERROR', f"解析{signalEnglishName}信号长度出错！")
                    isChanged = "ERROR"
                else:
                    signalBitLen = signalInfo['signalBitLen']
                if signalInfo["signalType"] is None:
                    miscFuncs.notification('ERROR', f"解析{signalEnglishName}信号数据类型出错！")
                    isChanged = "ERROR"
                else:
                    signalType = signalInfo['signalType']
                if signalInfo["signalUpperLimit"] is None:
                    miscFuncs.notification('ERROR', f"解析{signalEnglishName}信号上限值出错！")
                    isChanged = "ERROR"
                else:
                    signalUpperLimit = signalInfo['signalUpperLimit']

            # 前缀替换
            rtn = self.__changeSignalFront(signalFrontChange, referenceSignal, signalEnglishName)
            if rtn["referenceSignal"] is not None:
                referenceSignal = rtn["referenceSignal"]
            if rtn["signalEnglishName"] is not None:
                signalEnglishName = rtn["signalEnglishName"]

            # 参考信号校验，校验参考信号是否在ARXML文件中存在，只有支持的功能才会参考信号校验
            if changeType == "新增信号" or changeType == "删除信号" or changeType == "调整信号信息" or changeType == "调整信号名称":
                rtn = self.__checkReferenceSignalExcite(pyToolCurrentPath, referenceSignal, signalEnglishName, signalChineseName)
                # 校验不通过不予变更
                if rtn == 0:
                    isChanged = "ERROR"
            else:
                #不支持的功能则跳过
                isChanged = "SKIP"

            dict_temp = {
                'index': changePointID,
                'messageCAN': messageCAN,
                'messageID': messageID,
                'messageSendPoint': messageSendPoint,
                'messageReceivePoint': messageReceivePoint,
                'signalEnglishName': signalEnglishName,
                'signalChineseName': signalChineseName,
                'changeType': changeType,
                'signalStartBit': signalStartBit,
                'signalBitLen': signalBitLen,
                'signalType': signalType,
                'signalUpperLimit': signalUpperLimit,
                'referenceSignal': referenceSignal,
                'isChanged': isChanged,
                'currentECU': currentECU,
                'transmitDirection': transmitDirection,
                'changeDetail':changeDetail,
                'referenceSignal_1': referenceSignal_1
            }
            self.signalDetailInfoAll.append(dict_temp)

            if changeType == "新增信号" or changeType == "删除信号" or changeType == "调整信号信息" or changeType == "调整信号名称":
                self.signalDetailInfo.append(dict_temp)
            else:
                miscFuncs.notification('INFO', f"非脚本变更项：{signalEnglishName}：{signalChineseName}，{changeType}")

        miscFuncs.notification('INFO', "变更点解析完成")

        # 删除临时拼接DBC文件
        if os.path.exists(dbcFilePath):
            os.remove(dbcFilePath)
            if config.config_debugMode == True:
                miscFuncs.notification('INFO', f"文件 {dbcFilePath} 已被删除")
        else:
            miscFuncs.notification('ERROR', f"文件 {dbcFilePath} 不存在", "__transfer2signalInfo")

        if config.config_debugMode == True:
            miscFuncs.notification('INFO', "变更点Excel解析完成")
            inputFuncs.showChangeDicts("detailInfo")

    # 读取输入EXCEL信息
    def getChangeDicts(self, pyToolCurrentPath):
        miscFuncs.notification('INFO', "解析变更点进行中...")

        excelFilePath = pyToolCurrentPath + config.config_changesInfoFilePath
        # TO DO
        # 添加excelFilePath判空退出的逻辑
        workbook = openpyxl.load_workbook(excelFilePath)
        sheet = workbook["Sheet1"]
        self.__getColumnIndex(sheet)
        sheet_rows = list(sheet.rows)

        for row in sheet_rows[1:]:
            if row[self.columnIndices[inputFuncs_t.SIGNAL_ENG_NAME]].value is not None:
                dictTemp = {
                    inputFuncs_t.SIGNAL_NUM: row[self.columnIndices[inputFuncs_t.SIGNAL_NUM]].value,
                    inputFuncs_t.SIGNAL_CAN: row[self.columnIndices[inputFuncs_t.SIGNAL_CAN]].value,
                    inputFuncs_t.SIGNAL_MESSAGE_ID: row[self.columnIndices[inputFuncs_t.SIGNAL_MESSAGE_ID]].value,
                    inputFuncs_t.SIGNAL_SEND_POINT: row[self.columnIndices[inputFuncs_t.SIGNAL_SEND_POINT]].value,
                    inputFuncs_t.SIGNAL_RECEIVE_POINT: row[self.columnIndices[inputFuncs_t.SIGNAL_RECEIVE_POINT]].value,
                    inputFuncs_t.SIGNAL_ENG_NAME: row[self.columnIndices[inputFuncs_t.SIGNAL_ENG_NAME]].value,
                    inputFuncs_t.SIGNAL_CHA_NAME: row[self.columnIndices[inputFuncs_t.SIGNAL_CHA_NAME]].value,
                    inputFuncs_t.SIGNAL_CHANGE_TYPE: row[self.columnIndices[inputFuncs_t.SIGNAL_CHANGE_TYPE]].value,
                    inputFuncs_t.SIGNAL_FRONT_CHANGE: row[self.columnIndices[inputFuncs_t.SIGNAL_FRONT_CHANGE]].value,
                    inputFuncs_t.SIGNAL_REFERENCE: row[self.columnIndices[inputFuncs_t.SIGNAL_REFERENCE]].value,
                }
                self.inputDicts.append(dictTemp)
        if config.config_debugMode == True:
            inputFuncs.showChangeDicts("inputInfo")

        # 解析excel
        self.__transfer2signalInfo(pyToolCurrentPath)

    def showChangeDicts(self, dictsName):
        if dictsName == "detailInfo":
            if self.signalDetailInfoAll is not None:
                for detailInfo in enumerate(self.signalDetailInfoAll):
                    miscFuncs.notification('DEBUG', f"{detailInfo}\n")
            else:
                miscFuncs.notification('ERROR', f"Signal change dicts is null!")
        else:
            if self.inputDicts is not None:
                for rowDict in enumerate(self.inputDicts):
                    miscFuncs.notification('DEBUG', f"{rowDict}\n")
            else:
                miscFuncs.notification('ERROR', f"Signal change dicts is null!")

class masterFuncs_t:

    # 输出期望字符串滑动窗口
    def getStringContext(self, filePath, windowSize, refStr):
        class refContextStruct:
            def __init__(self, config_windowSize, config_refStr):
                self.config_windowSize = config_windowSize
                self.config_refStr = config_refStr
                self.res_isValid = 0
                self.res_start_index = 0
                self.res_end_index = 0
                self.res_find_index = 0
                self.res_find_flag = 0
                self.res_finish_flag = 0
                self.insertLines = 0
                self.lines = []

        class windowLinesStruct:
            def __init__(self, start_index=0, end_index=0):
                self.start_index = start_index
                self.end_index = end_index
                self.lines = []

            def __copy__(self):
                new_obj = type(self)(self.start_index, self.end_index)
                new_obj.lines = self.lines
                return new_obj

        currentWindow = windowLinesStruct(0, windowSize - 1)
        lastWindow = windowLinesStruct()
        refWindow = [refContextStruct(windowSize, refStr) for _ in range(config.config_refWindowNum)]
        nextLine = None


        with open(filePath, 'r', encoding='utf-8') as file:
            window_cnt = 0
            current_index = 0
            refWindow_cnt = 0

            for line in file:
                currentWindow.lines.append(line)
                current_index += 1

                if current_index % windowSize == 0:
                    window_cnt += 1
                   # print(currentWindow.lines, currentWindow.start_index, currentWindow.end_index)
                    currentWindow.start_index = current_index
                    currentWindow.end_index = current_index + windowSize - 1
                    lastWindow = copy.copy(currentWindow)
                    currentWindow.lines = currentWindow.lines[windowSize:]

                if refStr in line:
                    refWindow[refWindow_cnt].res_find_flag = 1
                    refWindow[refWindow_cnt].res_find_index = current_index
                    refWindow[refWindow_cnt].res_start_index = max(current_index - windowSize // 2, 0)      # 窗口起始行应当是当前行（包含目标str）-一半窗口大小
                    refWindow[refWindow_cnt].res_end_index = current_index + windowSize // 2                # 窗口结束行应当是当前行（包含目标str）+一半窗口大小
                    temp = refWindow[refWindow_cnt].res_start_index % windowSize                            # 判断找到期望str时，期望str的起始行在窗口中的位置
                    if (current_index % windowSize) < temp:
                        refWindow[refWindow_cnt].lines.clear()                                              # 没有这一行会出BUG，这个BUG很难描述
                        refWindow[refWindow_cnt].lines[0:(windowSize - temp + 1)] = lastWindow.lines[temp:]
                    else:
                        refWindow[refWindow_cnt].lines.clear()
                        refWindow[refWindow_cnt].lines[0:windowSize // 2] = currentWindow.lines[temp:current_index]

                if (refWindow[refWindow_cnt].res_find_flag and current_index == refWindow[refWindow_cnt].res_end_index) or nextLine == '':
                    temp = refWindow[refWindow_cnt].res_find_index % windowSize
                    if (refWindow[refWindow_cnt].res_end_index % windowSize) < temp:
                        refWindow[refWindow_cnt].lines.extend(lastWindow.lines[temp:])
                        refWindow[refWindow_cnt].lines.extend(currentWindow.lines[0:])
                    else:
                        refWindow[refWindow_cnt].lines.extend(currentWindow.lines[0:(temp + windowSize // 2)])
                    refWindow[refWindow_cnt].res_find_flag = 0
                    refWindow[refWindow_cnt].res_finish_flag = 1
                    refWindow[refWindow_cnt].res_isValid = 1
                    refWindow_cnt += 1

        return refWindow

    # arxml文件更新函数_删除信号
    def updataArxmlFile_Delete(self, inputFilePath, copyWindowInfo):
        # 把copywindow列表中的元素，按照插入行从小到大排序
        copyWindowInfo.sort(key=lambda x: x["insertLines"])
        outputFilePath = inputFilePath + '.tmp'
        currentLine = 1
        try:
            with open(inputFilePath, 'r', encoding='utf-8') as inputFile,\
                 open(outputFilePath, 'w', encoding='utf-8') as outputFile:
                for copyWindow in copyWindowInfo:
                    # insertLine之前，inputFile直接复制到outputFile
                    while currentLine < copyWindow["insertLines"] + 1:
                        line = inputFile.readline()
                        currentLine += 1
                        if line != "\n" : outputFile.write(line)

                    for i in range(0, copyWindow["windowSize"], 1):
                        line = inputFile.readline()
                        currentLine += 1
                        #删除行校验
                        if line == copyWindow["lines"][i]:
                            if config.config_debugMode == True:
                                miscFuncs.notification('DEBUG', f"删除行：{line}")
                        else:
                            temp = copyWindow["insertLines"]
                            miscFuncs.notification('FATAL', f"路径:{inputFilePath} \n"
                                                            f"原arxml行数:{temp}\n"
                                                            f"删除行计算错误！该变更无法执行，ARXML文件已损坏 ")
                            # TO DO
                            # 在该处删除行校验代码不健全，有待完善
                            if line != "\n" : outputFile.write(line)

                    continue

                while True:
                    line = inputFile.readline()
                    if not line:
                        break
                    if line != "\n" : outputFile.write(line)

            os.replace(outputFilePath, inputFilePath)

        except Exception as e:
            # 发生错误时删除临时文件
            if os.path.exists(outputFilePath):
                os.remove(outputFilePath)
            raise e

    # arxml文件更新函数_新增信号
    def updataArxmlFile_Add(self, inputFilePath, copyWindowInfo):
        # 把copywindow列表中的元素，按照插入行从小到大排序
        copyWindowInfo.sort(key=lambda x: x["insertLines"])
        outputFilePath = inputFilePath + '.tmp'
        currentLine = 1
        try:
            with open(inputFilePath, 'r', encoding='utf-8') as inputFile,\
                 open(outputFilePath, 'w', encoding='utf-8') as outputFile:
                for copyWindow in copyWindowInfo:
                    #insertLine之前，inputFile直接复制到outputFile
                    while currentLine < copyWindow["insertLines"] + 1:
                        line = inputFile.readline()
                        if line != "\n" : outputFile.write(line)
                        currentLine += 1

                    #insertLine时，copywindow复制到outputFile
                    for i in range(len(copyWindow["lines"])):
                        outputFile.write(copyWindow["lines"][i])

                    continue

                while True:
                    line = inputFile.readline()
                    if not line:
                        break
                    if line != "\n" : outputFile.write(line)
            os.replace(outputFilePath, inputFilePath)

        except Exception as e:
            # 发生错误时删除临时文件
            if os.path.exists(outputFilePath):
                os.remove(outputFilePath)
            raise e

    # 计算copyWindow内容，和copyWindow的插入行，该函数用做删除信号
    def __calculateCopyWindowInfo_2(self, refWindow, tag, newSignalInfo):
        copyWindow = {"lines": [], "insertLines": 0, "windowSize": 0, "changeType": None}           # 函数返回值
        finishFlag = 0                                                                              # 结束标志位
        refWindow_start_line = 0                                                                    # 期望字符串在refWindow中的绝对位置。在arxml文件视角下，这个值是相对的。
        windowSize = 0                                                                              # 窗口大小

        copyWindow["changeType"] = "delete"

        for i in range(0, refWindow.config_windowSize-1, 1):
            if refWindow.config_refStr in refWindow.lines[i]:
                refWindow_start_line = i                                                #找到refWindow中期望字符串所在的位置，利用该值计算出copyWindow应当在arxml中的插入行

        for i in range(refWindow_start_line, -1, -1):                                   #i是在refWindow窗口中的绝对位置，向上遍历找到起始tag
            if tag in refWindow.lines[i]:
                for j in range(i, refWindow.config_windowSize - 1, 1):                  #j是在refWindow窗口中的绝对位置，向下遍历找到结尾tag
                    line = refWindow.lines[j]
                    windowSize += 1
                    copyWindow["lines"].append(line)

                    if "</"+tag+">" in line:
                        finishFlag = 1
                        refWindow.insertLines = refWindow.res_find_index + j - refWindow_start_line - windowSize
                        copyWindow["insertLines"] = refWindow.insertLines
                        copyWindow["windowSize"] = windowSize
                        break
                break

        if finishFlag:
            return copyWindow
        else:
            return None


    # 计算copyWindow内容，和copyWindow的插入行，该函数用做新增信号
    def __calculateCopyWindowInfo_1(self, refWindow, tag, newSignalInfo):
        copyWindow = {"lines": [], "insertLines": 0, "windowSize": 0, "changeType": None}    # 函数返回值
        finishFlag = 0                                                                       # 结束标志位
        refWindow_start_line = 0                                                             # 期望字符串在refWindow中的绝对位置。在arxml文件视角下，这个值是相对的。
        windowSize = 0

        copyWindow["changeType"] = "delete"

        flag = [{"tag":"ComSignalType", "state":0},     # "ECUC-CONTAINER-VALUE"配置项专属
                {"tag":"ComBitSize", "state":0},
                {"tag":"ComBitPosition", "state":0}]

        for i in range(0, refWindow.config_windowSize-1, 1):
            if refWindow.config_refStr in refWindow.lines[i]:
                refWindow_start_line = i                                                #找到refWindow中期望字符串所在的位置，利用该值计算出copyWindow应当在arxml中的插入行
                                                                                        #检查：refWindow.lines[refWindow_start_line]和arxml[refWindow.res_find_index]应当指向同一行
        for i in range(refWindow_start_line, -1, -1):                                   #i是在refWindow窗口中的绝对位置，向上遍历找到起始tag
            if tag in refWindow.lines[i]:
                for j in range(i, refWindow.config_windowSize - 1, 1):                  #j是在refWindow窗口中的绝对位置，向下遍历找到结尾tag
                    line = refWindow.lines[j]
                    windowSize += 1

                    if refWindow.config_refStr in line:
                        line = line.replace(refWindow.config_refStr, newSignalInfo["signalEnglishName"])

                    # 下列是arxml文件需要根据DBC配置的配置项
                    if tag == "DATA-CONSTR" and ("<UPPER-LIMIT>" in line):
                        temp = newSignalInfo["signalUpperLimit"] + '.0'
                        pattern = r'(<UPPER-LIMIT>).*?(</UPPER-LIMIT>)'
                        line = re.sub(pattern, rf'\g<1>{temp}\g<2>', line)
                    elif tag == "IMPLEMENTATION-DATA-TYPE" and ("IMPLEMENTATION-DATA-TYPE-REF" in line):
                        pattern = r'(<IMPLEMENTATION-DATA-TYPE-REF[^>]+>.*?/)([^/<]+)(</IMPLEMENTATION-DATA-TYPE-REF>)'
                        temp = newSignalInfo["signalType"].lower()
                        line = re.sub(pattern, rf'\g<1>{temp}\g<3>', line)
                    elif tag == "TRANSMISSION-MODE-CONDITION" and ("<MAX>" in line):
                        pattern = r'(<MAX>).*?(</MAX>)'
                        line = re.sub(pattern, rf'\g<1>{newSignalInfo["signalUpperLimit"]}\g<2>', line)
                    elif tag == "I-SIGNAL-TO-I-PDU-MAPPING" and ("<START-POSITION>" in line):
                        pattern = r'(<START-POSITION>).*?(</START-POSITION>)'
                        line = re.sub(pattern, rf'\g<1>{newSignalInfo["signalStartBit"]}\g<2>', line)
                    elif tag == "I-SIGNAL":
                        if "<LENGTH>" in line:
                            pattern = r'(<LENGTH>).*?(</LENGTH>)'
                            line = re.sub(pattern, rf'\g<1>{newSignalInfo["signalBitLen"]}\g<2>', line)
                        if "BASE-TYPE-REF" in line:
                            pattern = r'(<BASE-TYPE-REF[^>]+>.*?/)([^/<]+)(</BASE-TYPE-REF>)'
                            temp = newSignalInfo["signalType"].lower()
                            line = re.sub(pattern, rf'\g<1>{temp}\g<3>', line)
                    elif tag == "COMPU-METHOD" and ("<UPPER-LIMIT>" in line):
                        temp = newSignalInfo["signalUpperLimit"] + '.0'
                        pattern = r'(<UPPER-LIMIT>).*?(</UPPER-LIMIT>)'
                        line = re.sub(pattern, rf'\g<1>{temp}\g<2>', line)
                    elif tag == "ECUC-CONTAINER-VALUE":
                        if flag[0]["state"] == 1:
                            pattern = r'(<VALUE>).*?(</VALUE>)'
                            if newSignalInfo["signalBitLen"] == '1':
                                temp = "BOOLEAN"
                            else:
                                temp = newSignalInfo["signalType"]
                            line = re.sub(pattern, rf'\g<1>{temp}\g<2>', line)
                            flag[0]["state"] = 0
                        if flag[1]["state"] == 1:
                            pattern = r'(<VALUE>).*?(</VALUE>)'
                            line = re.sub(pattern, rf'\g<1>{newSignalInfo["signalBitLen"]}\g<2>', line)
                            flag[1]["state"] = 0
                        if flag[2]["state"] == 1:
                            pattern = r'(<VALUE>).*?(</VALUE>)'
                            line = re.sub(pattern, rf'\g<1>{newSignalInfo["signalStartBit"]}\g<2>', line)
                            flag[2]["state"] = 0

                        # 当前行检索到标签，变更需要在下一行体现，故flag[0]["state"]的置位要在flag[0]["state"]判断之后，与上面的if位置不可调换
                        if flag[0]["tag"] in line and flag[0]["state"] == 0: flag[0]["state"] = 1
                        if flag[1]["tag"] in line and flag[1]["state"] == 0: flag[1]["state"] = 1
                        if flag[2]["tag"] in line and flag[2]["state"] == 0: flag[2]["state"] = 1

                    copyWindow["lines"].append(line)

                    if "</"+tag+">" in line:
                        finishFlag = 1
                        refWindow.insertLines = refWindow.res_find_index + j - refWindow_start_line       # 忘了为什么要这么算，debug一下即可揭晓答案
                        copyWindow["insertLines"] = refWindow.insertLines
                        copyWindow["windowSize"] = windowSize

                        # print(f'{tag} is find')
                        # line_cnt = 0
                        # with open(filePath, 'r', encoding='utf-8') as file:
                        #     for line in file:
                        #         line_cnt += 1
                        #         if line_cnt == refWindow.res_find_index: print(f'find_index_line:{line}', end = '')
                        # print(f'ref_window_line:{refWindow.lines[refWindow_start_line]}', end = '')
                        # print(f'find_index:{refWindow.res_find_index}')
                        # print(f'up_find_cnt_i:{i}')
                        # print(f'down_find_cnt_j:{j}')
                        # print(f'refWindow_start_line:{refWindow_start_line}')
                        # print(f'insert_line:{refWindow.insertLines}')
                        # for copyWindow_temp in copyWindow:
                        #     print(copyWindow_temp, end='')
                        break
                break

        if finishFlag:
            return copyWindow
        else:
            return None

    # 获得某一信号的替换窗口和替换行
    # eg.   新增信号
    #       copyWindowInfo["copyWindow"]:新增信号的插入窗口
    #       copyWindowInfo["insertLines"]:新增信号的插入行
    #
    # eg.   删除信号
    #       copyWindowInfo["copyWindow"]:要删除信号所在窗口
    #       copyWindowInfo["insertLines"]:删除的起始行
    def getSingalSignalCopyWindowInfo(self, arxmlFileName, refWindow, newSignalInfo):
        tagInfo = [#APPCOM_SWC.arxml （0,5）
                   {"tag": "APPLICATION-PRIMITIVE-DATA-TYPE", "state":0},
                   {"tag": "VARIABLE-ACCESS", "state":0},
                   {"tag": "DATA-CONSTR", "state":0},
                   {"tag": "DATA-TYPE-MAP", "state":0},
                   {"tag": "VARIABLE-DATA-PROTOTYPE", "state":0},
                   {"tag": "IMPLEMENTATION-DATA-TYPE", "state":0},
                   #DBC_SysDesc.arxml （6,13）
                   {"tag": "I-SIGNAL-PORT", "state":0},
                   {"tag": "I-SIGNAL-TRIGGERING", "state":0},
                   {"tag": "I-SIGNAL-TRIGGERING-REF-CONDITIONAL", "state":0},
                   {"tag": "TRANSMISSION-MODE-CONDITION", "state":0},
                   {"tag": "I-SIGNAL-TO-I-PDU-MAPPING", "state":0} ,
                   {"tag": "SYSTEM-SIGNAL", "state":0},
                   {"tag": "I-SIGNAL", "state":0},
                   {"tag": "COMPU-METHOD", "state":0},
                   #MCU_SYSTEM.arxml （14,14）
                   {"tag": "SENDER-RECEIVER-TO-SIGNAL-MAPPING", "state":0},
                   #Project_EcucValues.arxml （15,16）
                   {"tag": "ECUC-REFERENCE-VALUE", "state":0},
                   {"tag": "ECUC-CONTAINER-VALUE", "state":0},]

        copyWindowInfo = []

        for window in (w for w in refWindow[:config.config_refWindowNum] if w.res_isValid):
            for tag in (t for t in tagInfo[0:6] if ((t["state"] == 0 or t["state"] == 2) and arxmlFileName == "APPCOM_SWC.arxml")):
                rtn = None

                if newSignalInfo['changeType'] == "新增信号":
                    rtn = self.__calculateCopyWindowInfo_1(window, tag["tag"], newSignalInfo)
                elif newSignalInfo['changeType'] == "删除信号":
                    rtn = self.__calculateCopyWindowInfo_2(window, tag["tag"], newSignalInfo)
                if rtn is not None:
                    tempWindow = {
                        "lines": rtn["lines"],
                        "insertLines": rtn["insertLines"],
                        "windowSize": rtn["windowSize"],
                        "changeType": rtn["changeType"],
                    }
                    copyWindowInfo.append(tempWindow)
                    tag["state"] = 1
                    break
                else:
                    # tag["state"] = 2 表示该tag未找到
                    tag["state"] = 2
                    temp_0 = tag["tag"]
                    #miscFuncs.notification('WARNING',
                    #                       f"filePath: APPCOM_SWC.arxml, signal: {window.config_refStr}, tag: {temp_0} is not find!")


            for tag in (t for t in tagInfo[6:14] if ((t["state"] == 0 or t["state"] == 2) and arxmlFileName == "DBC_SysDesc.arxml")):
                rtn = None

                if newSignalInfo['changeType'] == "新增信号":
                    rtn = self.__calculateCopyWindowInfo_1(window, tag["tag"], newSignalInfo)
                elif newSignalInfo['changeType'] == "删除信号":
                    rtn = self.__calculateCopyWindowInfo_2(window, tag["tag"], newSignalInfo)
                if rtn is not None:
                    tempWindow = {
                        "lines": rtn["lines"],
                        "insertLines": rtn["insertLines"],
                        "windowSize": rtn["windowSize"],
                        "changeType": rtn["changeType"],
                    }
                    copyWindowInfo.append(tempWindow)
                    tag["state"] = 1
                    break
                else:
                    # tag["state"] = 2 表示该tag未找到
                    tag["state"] = 2
                    temp_0 = tag["tag"]
                    #miscFuncs.notification('WARNING',
                    #                       f"filePath: DBC_SysDesc.arxml, signal: {window.config_refStr}, tag: {temp_0} is not find!")

            for tag in (t for t in tagInfo[14:15] if ((t["state"] == 0 or t["state"] == 2) and arxmlFileName == "MCU_SYSTEM.arxml")):
                rtn = None

                if newSignalInfo['changeType'] == "新增信号":
                    rtn = self.__calculateCopyWindowInfo_1(window, tag["tag"], newSignalInfo)
                elif newSignalInfo['changeType'] == "删除信号":
                    rtn = self.__calculateCopyWindowInfo_2(window, tag["tag"], newSignalInfo)
                if rtn is not None:
                    tempWindow = {
                        "lines": rtn["lines"],
                        "insertLines": rtn["insertLines"],
                        "windowSize": rtn["windowSize"],
                        "changeType": rtn["changeType"],
                    }
                    copyWindowInfo.append(tempWindow)
                    tag["state"] = 1
                    break
                else:
                    # tag["state"] = 2 表示该tag未找到
                    tag["state"] = 2
                    temp_0 = tag["tag"]
                    #miscFuncs.notification('WARNING',
                    #                      f"filePath: MCU_SYSTEM.arxml, signal: {window.config_refStr}, tag: {temp_0} is not find!")

            for tag in (t for t in tagInfo[15:17] if ((t["state"] == 0 or t["state"] == 2) and arxmlFileName == "Project_EcucValues.arxml")):
                rtn = None

                if newSignalInfo['changeType'] == "新增信号":
                    rtn = self.__calculateCopyWindowInfo_1(window, tag["tag"], newSignalInfo)
                elif newSignalInfo['changeType'] == "删除信号":
                    rtn = self.__calculateCopyWindowInfo_2(window, tag["tag"], newSignalInfo)
                if rtn is not None:
                    tempWindow = {
                        "lines": rtn["lines"],
                        "insertLines": rtn["insertLines"],
                        "windowSize": rtn["windowSize"],
                        "changeType": rtn["changeType"],
                    }
                    copyWindowInfo.append(tempWindow)
                    tag["state"] = 1
                    break
                else:
                    # tag["state"] = 2 表示该tag未找到
                    tag["state"] = 2
                    temp_0 = tag["tag"]
                    #miscFuncs.notification('WARNING',
                    #                       f"filePath: Project_EcucValues.arxml, signal: {window.config_refStr}, tag: {temp_0} is not find!")

        # 检查copyWindowInfo列表
        if config.config_debugMode == True:
            for temp in copyWindowInfo:
                miscFuncs.notification('DEBUG', f"copyWindowInfo.insertLines: {temp['insertLines']}\n")
                for temp_2 in temp["copyWindow"]:
                    miscFuncs.notification('DEBUG', f"copyWindowInfo.copyWindow: {temp_2}")

        return copyWindowInfo

        # self.updataArxmlFile(filePath, copyWindowInfo)
        # for i in range(len(copyWindow)):
        #     copyWindow[i] = copyWindow[i].replace(refWindow.config_refStr, newSignalInfo.singleName)

        # for lines in copyWindow:
        #     print(lines.rstrip('\n'))
        # print(insertLines)  #关键参数，要插入的行数

    def logFileInit(self, pyToolCurrentPath, newSignalInfo):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "log"

        headers = list(newSignalInfo.keys())
        ws.append(headers)

        for i, header in enumerate(headers, 1):
            col_width = len(header) + 2
            #获取列字母
            col_letter = get_column_letter(i)
            ws.column_dimensions[col_letter].width = col_width

        header_font = Font(name='等线', size=12, bold=True, color='000000')  # 黑色等线字体
        header_fill = PatternFill(start_color='D3D3D3', end_color='D3D3D3', fill_type='solid')  # 浅灰色背景

        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill

        ws.freeze_panes = 'A2'  # 冻结第一行

        outputFilePath = pyToolCurrentPath + config.config_logFilePath
        wb.save(outputFilePath)

    def logFileIsExists(self, pyToolCurrentPath):
        logFilePath = pyToolCurrentPath + config.config_logFilePath

        if os.path.exists(logFilePath):
            return True

        return False
        # 检查第二行是否为空
        # try:
        #     wb = openpyxl.load_workbook(filename=filePath, read_only=True)
        #     sheet = wb.worksheets[0]
        #
        #     isEmpty = True
        #     value = sheet.cell(row=2, column=1).value
        #     if value is not None and str(value).strip() != "":
        #         isEmpty = False
        #
        #     return True
        #
        # except Exception as e:
        #     print(f"读取文件时出错: {e}")
        #     return False  # 文件存在但读取失败

    # 输出refHeader列下值为refHeaderValue所在行outputHeader列的值
    def logFileGetState(self, pyToolCurrentPath, inputHeader, inputHeaderValue, outputHeader):
        logFilePath = pyToolCurrentPath + config.config_logFilePath
        inputHeaderColum = 0
        outputHeaderColum = 0
        inputHeaderValueRow = 0
        isFindFlag = 0

        try:
            wb = openpyxl.load_workbook(logFilePath, read_only=True)
            sheet = wb.worksheets[0]
            header = sheet[1]

            #找表头所在列
            for cell in header:
                #miscFuncs.notification('INFO', f"header:{cell.value}")
                if cell.value == inputHeader:
                    inputHeaderColum = cell.column
                    # miscFuncs.notification('INFO', f"inputHeaderColum:{inputHeaderColum}")
                if cell.value == outputHeader:
                    outputHeaderColum = cell.column
                    # miscFuncs.notification('INFO', f"outputHeaderColum:{outputHeaderColum}")

            #找匹配值所在行
            for row in sheet.iter_rows(min_row=2, min_col=inputHeaderColum, max_col=inputHeaderColum):
                cell = row[0]
                if cell.value == inputHeaderValue:
                    inputHeaderValueRow = cell.row
                    if isFindFlag == 1:
                        miscFuncs.notification('WARNING', f"信号{inputHeaderValue}重复")
                    isFindFlag = 1
                    # miscFuncs.notification('INFO', f"refHeaderValue:{inputHeaderValue}  "
                                                   # f"refHeaderRow:{inputHeaderValueRow}")

            rtn = sheet.cell(row=inputHeaderValueRow, column=outputHeaderColum).value
            # miscFuncs.notification('INFO', f"rtn:{rtn}")

        except Exception as e:
            miscFuncs.notification('ERROR', f"发生错误: {e}")
            return False

        return rtn

    # 在输出log文件中追加数据
    def logFileAppand(self, pyToolCurrentPath, newSignalInfo):
        logFilePath = pyToolCurrentPath + config.config_logFilePath

        try:
            wb = openpyxl.load_workbook(logFilePath)
            sheet = wb.worksheets[0]

            headerRow = 1
            headerMap = {}

            for columnIndex, cell in enumerate(sheet[headerRow], start=1):
                headerMap[cell.value] = columnIndex     # 格式：{表头文本: 列索引}

            startRow = sheet.max_row + 1

            for key, value in newSignalInfo.items():
                if key in headerMap:
                    columnIndex = headerMap[key]
                    sheet.cell(row=startRow, column=columnIndex, value=value)

            wb.save(logFilePath)

        except Exception as e:
            miscFuncs.notification('ERROR', f"发生错误: {e}")
            return False

    def logFileChangeState(self, pyToolCurrentPath, inputHeader, inputHeaderValue, outputHeader, outputHeaderValue):
        logFilePath = pyToolCurrentPath + config.config_logFilePath
        inputHeaderColum = 0
        outputHeaderColum = 0
        inputHeaderValueRow = 0

        try:
            wb = openpyxl.load_workbook(logFilePath)
            sheet = wb.worksheets[0]
            header = sheet[1]

            # 找表头所在列
            for cell in header:
                # miscFuncs.notification('INFO', f"header:{cell.value}")
                if cell.value == inputHeader:
                    inputHeaderColum = cell.column
                    # miscFuncs.notification('DEBUG', f"inputHeaderColum:{inputHeaderColum}")
                if cell.value == outputHeader:
                    outputHeaderColum = cell.column
                    # miscFuncs.notification('DEBUG', f"outputHeaderColum:{outputHeaderColum}")

            # 找匹配值所在行
            for row in sheet.iter_rows(min_row=2, min_col=inputHeaderColum, max_col=inputHeaderColum):
                cell = row[0]
                if cell.value == inputHeaderValue:
                    inputHeaderValueRow = cell.row
                    # miscFuncs.notification('DEBUG', f"refHeaderValue:{inputHeaderValue} refHeaderRow:{inputHeaderValueRow}")

            sheet.cell(row=inputHeaderValueRow, column=outputHeaderColum).value = outputHeaderValue
            # miscFuncs.notification('INFO', f"rtn:{rtn}")

            wb.save(logFilePath)

        except Exception as e:
            miscFuncs.notification('ERROR', f"发生错误: {e}")
            return False

        return 0

    def logFileCreate(self, arxmlToolPath):
        if config.config_debugMode == True:
            miscFuncs.notification('INFO', f"新建LOG文件")
        masterFuncs.logFileInit(arxmlToolPath, inputFuncs.signalDetailInfoAll[0])
        # TO DO
        # 这里把detail列表加入LOG,策略待修改。应当在变更完毕后，回读arxml，把回读数据传入LOG
        for newSignalInfo in inputFuncs.signalDetailInfoAll:
            masterFuncs.logFileAppand(arxmlToolPath, newSignalInfo)

        print("")                                                    # 这个print应当写在mainTask中维持UI，藏在这里显得mainTask美观

    # 检查新信号名称在ARXML中是否存在，新信号名称存在的话就不予变更，会有bug
    def checkNewSiganlName(self, pyToolCurrentPath, newSiganl):
        filePath = pyToolCurrentPath + config.config_inputFilePath + "\MCU_SYSTEM.arxml"

        temp_0 = newSiganl["signalEnglishName"]
        temp_1 = newSiganl["signalChineseName"]

        try:
            with open(filePath, 'r', encoding='utf-8') as file:
                for line in file:
                    if newSiganl["signalEnglishName"] in line:
                        miscFuncs.notification('ERROR', f"{temp_0}:{temp_1} 信号名称变更无法进行，需手动变更!")
                        rtn = 0
                        break
                rtn = 1
        except Exception as e:
            raise e
        return rtn

    #ARXML文件预处理
    # TO DO
    # 在文件预处理处，应当传入newSignal，保存referenceSignal在每个arxml中出现的次数，用于信号回读校验
    def arxmlFilePreprocessing(self, pyToolCurrentPath, newSiganl):
        arxmlFileInfo = [{"fileName": "\APPCOM_SWC.arxml", "fileLineNum": 0},
                         {"fileName": "\DBC_SysDesc.arxml", "fileLineNum": 0},
                         {"fileName": "\MCU_SYSTEM.arxml", "fileLineNum": 0},
                         {"fileName": "\Project_EcucValues.arxml", "fileLineNum": 0}]

        for arxmlFile in arxmlFileInfo:
            inputFilePath = pyToolCurrentPath + config.config_inputFilePath + str(arxmlFile["fileName"])
            outputFilePath = inputFilePath + ".tmp"
            try:
                with open(inputFilePath, 'r', encoding='utf-8') as inputFile, \
                        open(outputFilePath, 'w', encoding='utf-8') as outputFile:

                    line = inputFile.readline()
                    while line:
                        arxmlFile["fileLineNum"] += 1
                        nextLine = inputFile.readline()                                     #预读下一行，用于判断是否到文件末尾
                        isFinalLine = (nextLine == '')

                        outputFile.write(line)
                        if isFinalLine:                                                     #如果到文件末尾，在末尾添加100个回车符
                            line = "\n"
                            for i in range(0, 100, 1):                                      #这个参数100，取决于滑动窗口的最大值，
                                                                                            #Project_EcucValues文件滑动窗口的大小为100
                                outputFile.write(line)
                            break

                        line = nextLine
            except Exception as e:
                raise e

            # 对Project_EcucValues的特殊处理，解决新增报文情况下，Project_EcucValues无法完整新增配置项的问题
            if arxmlFile["fileName"] == "\Project_EcucValues.arxml":
                lineNum = 0
                startInsertFlag = 0
                tagCnt = 0
                try:
                    with open(inputFilePath, 'r', encoding='utf-8') as inputFile, \
                            open(outputFilePath, 'w', encoding='utf-8') as outputFile:
                        for line in inputFile:
                            lineNum += 1
                            outputFile.write(line)

                            # 遍历Project_EcucValues文件
                            # 当出现参考信号时，参考信号后tagCnt个</ECUC-CONTAINER-VALUE>配置项后会被插入50个换行
                            # 确保每个窗口内只包含一个配置项

                            if newSiganl["referenceSignal"] in line:
                                startInsertFlag = 1

                            if "</ECUC-CONTAINER-VALUE>" in line and startInsertFlag == 1:
                                tagCnt += 1
                                line = "\n"
                                for i in range(0, 50, 1):
                                    outputFile.write(line)

                            if tagCnt > 10:
                                startInsertFlag = 0

                except Exception as e:
                    raise e

            os.replace(outputFilePath, inputFilePath)

        if config.config_debugMode == True:
            miscFuncs.notification('INFO', arxmlFileInfo)

    def copyArxmlFile(self, originalPath, destPath):
        arxmlFileInfo = [{"fileName": "\APPCOM_SWC.arxml"},
                         {"fileName": "\DBC_SysDesc.arxml"},
                         {"fileName": "\MCU_SYSTEM.arxml"},
                         {"fileName": "\Project_EcucValues.arxml"}]

        for arxmlFile in arxmlFileInfo:
            inputFilePath = originalPath + str(arxmlFile["fileName"])
            outputFilePath = destPath + str(arxmlFile["fileName"])

            if not os.path.isfile(inputFilePath):
                temp = str(arxmlFile["fileName"])
                miscFuncs.notification('FATAL', f"{temp}不存在！")

            try:
                shutil.copy2(inputFilePath, outputFilePath)
            except Exception as e:
                miscFuncs.notification('FATAL', f"复制文件时出错: {e}")

    def checkUUIDrepeat(self, pyToolCurrentPath, arxmlName):
        inputFilePath = pyToolCurrentPath + config.config_inputFilePath + "\\" + arxmlName
        outputFilePath = inputFilePath + '.tmp'

        uuidCnt = 0
        repeatCnt = 0
        checkCnt = 0
        uuidQueue = miscFuncs.queue()
        checkQueue = miscFuncs.queue()

        try:
            with open(inputFilePath, 'r', encoding='utf-8') as file:
                for line in file:
                    if "UUID" in line:
                        uuidCnt += 1

        except Exception as e:
            raise e

        try:
            with open(inputFilePath, 'r', encoding='utf-8') as inputFile, \
                    open(outputFilePath, 'w', encoding='utf-8') as outputFile:

                for line in inputFile:
                    if "UUID" in line:
                        uuids = re.findall(
                            r'UUID="([^"]{8}-[^"]{4}-[^"]{4}-[^"]{4}-[^"]{12})"',
                            line
                        )

                        if uuidQueue.enqueue(str(uuids[0])) == False:
                            repeatCnt += 1
                            # 输出重复的UUID
                            # print(f"repeat uuid: {str(uuids[0])}")

                            while True:
                                new_uuid = str(uuid.uuid4())
                                if uuidQueue.enqueue(new_uuid) == True:
                                    # 输出新生成的UUID
                                    # print(f"new uuid: {str(new_uuid)}")
                                    break

                            line = re.sub(
                                r'UUID="([^"]{8}-[^"]{4}-[^"]{4}-[^"]{4}-[^"]{12})"',
                                f'UUID="{new_uuid}"',
                                line,
                                count=1  # 只替换第一个匹配项
                            )
                            # 输出替换后的行
                            # print(f"change line: {line}")

                    outputFile.write(line)

            os.replace(outputFilePath, inputFilePath)
            # 输出uuid数量和重复uuid数量
            # print(f"total uuid num: {uuidCnt}")
            # print(f"repeat uuid num: {repeatCnt}")

        except Exception as e:
            if os.path.exists(outputFilePath):
                os.remove(outputFilePath)
            raise e

        try:
            with open(inputFilePath, 'r', encoding='utf-8') as file:
                for line in file:
                    if "UUID" in line:
                        uuids = re.findall(
                            r'UUID="([^"]{8}-[^"]{4}-[^"]{4}-[^"]{4}-[^"]{12})"',
                            line
                        )
                        if checkQueue.enqueue(str(uuids[0])) == False:
                            checkCnt += 1

            if checkCnt == 0:
                return True
            else:
                return False

        except Exception as e:
            raise e

    def uuidUpdata(self, pyCurrentPath):
        print("")                                                                           # 这个print应当写在mainTask中维持UI，藏在这里显得mainTask美观
        rtn1 = masterFuncs.checkUUIDrepeat(pyCurrentPath, "APPCOM_SWC.arxml")
        rtn2 = masterFuncs.checkUUIDrepeat(pyCurrentPath, "DBC_SysDesc.arxml")
        rtn3 = masterFuncs.checkUUIDrepeat(pyCurrentPath, "MCU_SYSTEM.arxml")
        rtn4 = masterFuncs.checkUUIDrepeat(pyCurrentPath, "Project_EcucValues.arxml")
        if rtn1 and rtn2 and rtn3 and rtn4 == 1:
            miscFuncs.notification('INFO', f"动态UUID已完成")
        else:
            miscFuncs.notification('ERROR', f"仍存在重复UUID")

    def traverseChangePoints(self, pyCurrentPath):
        for index, newSignalInfo in enumerate(inputFuncs.signalDetailInfo):

            # 读取LOG中信号变更状态
            rtn = masterFuncs.logFileGetState(pyCurrentPath, "signalEnglishName",
                                                        newSignalInfo['signalEnglishName'], "isChanged")
            # 该信号仍需变更
            if rtn != "ERROR":
                temp_0 = newSignalInfo["index"]
                temp_1 = newSignalInfo["signalChineseName"]
                temp_2 = newSignalInfo["changeType"]

                miscFuncs.notification('LODING', f"{miscFuncs.prograssCnt / miscFuncs.prograssTotal}", None,
                                       f"变更点{temp_0}进行中... {temp_1}: {temp_2} ")

                if newSignalInfo['changeType'] == "新增信号":
                    rtn = task.addSignalTask(pyCurrentPath, newSignalInfo)
                elif newSignalInfo['changeType'] == "调整信号信息":
                    rtn = task.changeSignalInfoTask(pyCurrentPath, newSignalInfo)
                elif newSignalInfo['changeType'] == "删除信号":
                    # TO DO
                    # 删除信号前，应判断该信号是否存在
                    rtn = task.deleteSignalTask(pyCurrentPath, newSignalInfo)
                elif newSignalInfo['changeType'] == "调整信号名称":
                    rtn = task.changeSignalNameTask(pyCurrentPath, newSignalInfo)

                if rtn is True:
                    newSignalInfo['isChanged'] = "PASS"
                    masterFuncs.logFileChangeState(pyCurrentPath, "signalEnglishName",
                                                   newSignalInfo['signalEnglishName'],"isChanged",
                                                   newSignalInfo['isChanged'])
                else:
                    newSignalInfo['isChanged'] = "ERROR"
                    miscFuncs.notification('ERROR',f"\n变更点{index}，{newSignalInfo['signalChineseName']}："
                                                   f"{newSignalInfo['changeType']} 错误\n")

        miscFuncs.notification('LODING', f"{miscFuncs.prograssCnt / miscFuncs.prograssTotal}", None,
                                       f"变更已完成                                                           ")

    def creatZentaoLog(self, pyCurrentPath):
        zenTaoLogFilePath = pyCurrentPath + config.config_outputFilePath + "\zenTao.txt"
        cnt_pCAN = 1
        cnt_cCAN = 1
        cnt_imuCAN = 1

        contentList_pCAN = []
        contentList_cCAN = []
        contentList_imuCAN = []

        for newSignalInfo in inputFuncs.signalDetailInfoAll:
            temp = [newSignalInfo["messageID"],
                    newSignalInfo["signalEnglishName"],
                    newSignalInfo["signalChineseName"],
                    newSignalInfo["changeType"]]

            if newSignalInfo["messageCAN"] == "P_CANFD":
                content = f"报文{temp[0]}，{temp[1]}：{temp[2]}，{temp[3]}\n"
                contentList_pCAN.append(content)
                cnt_pCAN += 1
            elif newSignalInfo["messageCAN"] == "C_CANFD":
                content = f"报文{temp[0]}，{temp[1]}：{temp[2]}，{temp[3]}\n"
                contentList_cCAN.append(content)
                cnt_cCAN += 1
            elif newSignalInfo["messageCAN"] == "IMU_CANFD":
                content = f"报文{temp[0]}，{temp[1]}：{temp[2]}，{temp[3]}\n"
                contentList_imuCAN.append(content)
                cnt_imuCAN += 1

        try:
            with open(zenTaoLogFilePath, 'w', encoding='utf-8') as inputFile:
                if contentList_pCAN is not None:
                    inputFile.write("P_CANFD:\n")
                    for line in contentList_pCAN:
                        inputFile.write(line)
                    inputFile.write("\n")
                if contentList_cCAN is not None:
                    inputFile.write("C_CANFD:\n")
                    for line in contentList_cCAN:
                        inputFile.write(line)
                    inputFile.write("\n")
                if contentList_imuCAN is not None:
                    inputFile.write("IMU_CANFD:\n")
                    for line in contentList_imuCAN:
                        inputFile.write(line)
                    inputFile.write("\n")

        except Exception as e:
            raise e

class task_t:

    def changeSignalInfoTask(self, pyToolCurrentPath, newSignalInfo):
        # 调整信号信息涉及两个步骤，即删除信号后再新增信号
        # 删除信号时以信号本身为参考信号
        # 新增信号时以报文第一帧信号为参考信号或手动输入参考信号
        # 因此，调整信号信息时储存了两个参考信号referenceSignal（报文参考信号），referenceSignal_1（该信号本身）

        # 预处理ARXML文件
        masterFuncs.arxmlFilePreprocessing(pyToolCurrentPath, newSignalInfo)
        temp = newSignalInfo['referenceSignal']
        newSignalInfo['referenceSignal'] = newSignalInfo['referenceSignal_1']
        newSignalInfo['changeType'] = "删除信号"
        self.deleteSignalTask(pyToolCurrentPath, newSignalInfo)

        # 预处理ARXML文件
        masterFuncs.arxmlFilePreprocessing(pyToolCurrentPath, newSignalInfo)
        newSignalInfo['referenceSignal'] = temp
        newSignalInfo['changeType'] = "新增信号"
        self.addSignalTask(pyToolCurrentPath, newSignalInfo)

        rtn = True
        return rtn

    def addSignalTask(self, pyToolCurrentPath, newSignalInfo):
        temp_0 = newSignalInfo["index"]
        temp_1 = newSignalInfo["signalChineseName"]
        temp_2 = newSignalInfo["changeType"]

        # 预处理ARXML文件
        masterFuncs.arxmlFilePreprocessing(pyToolCurrentPath, newSignalInfo)

        filePath = pyToolCurrentPath + config.config_inputFilePath + "\APPCOM_SWC.arxml"
        if config.config_debugMode == True:
            miscFuncs.notification('INFO', f"检索APPCOM_SWC")
        else:
            miscFuncs.prograssCnt += 1
            miscFuncs.notification('LODING', f"{miscFuncs.prograssCnt / miscFuncs.prograssTotal}", None,
                                       f"变更点{temp_0}进行中... {temp_1}, {temp_2}")
        refWindowInfo = masterFuncs.getStringContext(filePath, 40, newSignalInfo["referenceSignal"])
        if config.config_debugMode == True:
            miscFuncs.notification('INFO', f"计算变更值")
        else:
            miscFuncs.prograssCnt += 1
            miscFuncs.notification('LODING', f"{miscFuncs.prograssCnt / miscFuncs.prograssTotal}", None,
                                       f"变更点{temp_0}进行中... {temp_1}, {temp_2}")
        copyWindowInfo = masterFuncs.getSingalSignalCopyWindowInfo("APPCOM_SWC.arxml", refWindowInfo, newSignalInfo)
        if copyWindowInfo is not None:
            if config.config_debugMode == True:
                miscFuncs.notification('INFO', f"变更APPCOM_SWC")
            else:
                miscFuncs.prograssCnt += 1
                miscFuncs.notification('LODING', f"{miscFuncs.prograssCnt / miscFuncs.prograssTotal}", None,
                                       f"变更点{temp_0}进行中... {temp_1}, {temp_2} ")
            masterFuncs.updataArxmlFile_Add(filePath, copyWindowInfo)
        else:
            miscFuncs.notification('ERROR', f"变更值copyWindow计算为空！")



        filePath = pyToolCurrentPath + config.config_inputFilePath + "\DBC_SysDesc.arxml"
        if config.config_debugMode == True:
            miscFuncs.notification('INFO', f"检索DBC_SysDesc")
        else:
            miscFuncs.prograssCnt += 1
            miscFuncs.notification('LODING', f"{miscFuncs.prograssCnt / miscFuncs.prograssTotal}", None,
                                       f"变更点{temp_0}进行中... {temp_1}, {temp_2}")
        refWindowInfo = masterFuncs.getStringContext(filePath, 44, newSignalInfo["referenceSignal"])
        if config.config_debugMode == True:
            miscFuncs.notification('INFO', f"计算变更值")
        else:
            miscFuncs.prograssCnt += 1
            miscFuncs.notification('LODING', f"{miscFuncs.prograssCnt / miscFuncs.prograssTotal}", None,
                                       f"变更点{temp_0}进行中... {temp_1}, {temp_2}")
        copyWindowInfo = masterFuncs.getSingalSignalCopyWindowInfo("DBC_SysDesc.arxml", refWindowInfo, newSignalInfo)
        if copyWindowInfo is not None:
            if config.config_debugMode == True:
                miscFuncs.notification('INFO', f"变更DBC_SysDesc")
            else:
                miscFuncs.prograssCnt += 1
                miscFuncs.notification('LODING', f"{miscFuncs.prograssCnt / miscFuncs.prograssTotal}", None,
                                       f"变更点{temp_0}进行中... {temp_1}, {temp_2} ")
            masterFuncs.updataArxmlFile_Add(filePath, copyWindowInfo)
        else:
            miscFuncs.notification('ERROR', f"变更值copyWindow计算为空！")



        filePath = pyToolCurrentPath + config.config_inputFilePath + "\MCU_SYSTEM.arxml"
        if config.config_debugMode == True:
            miscFuncs.notification('INFO', f"检索MCU_SYSTEM")
        else:
            miscFuncs.prograssCnt += 1
            miscFuncs.notification('LODING', f"{miscFuncs.prograssCnt / miscFuncs.prograssTotal}", None,
                                       f"变更点{temp_0}进行中... {temp_1}, {temp_2}")
        refWindowInfo = masterFuncs.getStringContext(filePath, 40, newSignalInfo["referenceSignal"])
        if config.config_debugMode == True:
            miscFuncs.notification('INFO', f"计算变更值")
        else:
            miscFuncs.prograssCnt += 1
            miscFuncs.notification('LODING', f"{miscFuncs.prograssCnt / miscFuncs.prograssTotal}", None,
                                       f"变更点{temp_0}进行中... {temp_1}, {temp_2}")
        copyWindowInfo = masterFuncs.getSingalSignalCopyWindowInfo("MCU_SYSTEM.arxml", refWindowInfo, newSignalInfo)
        if copyWindowInfo is not None:
            if config.config_debugMode == True:
                miscFuncs.notification('INFO', f"变更MCU_SYSTEM")
            else:
                miscFuncs.prograssCnt += 1
                miscFuncs.notification('LODING', f"{miscFuncs.prograssCnt / miscFuncs.prograssTotal}", None,
                                       f"变更点{temp_0}进行中... {temp_1}, {temp_2} ")
            masterFuncs.updataArxmlFile_Add(filePath, copyWindowInfo)
        else:
            miscFuncs.notification('ERROR', f"变更值copyWindow计算为空！")



        filePath = pyToolCurrentPath + config.config_inputFilePath + "\Project_EcucValues.arxml"
        if config.config_debugMode == True:
            miscFuncs.notification('INFO', f"检索Project_EcucValues")
        else:
            miscFuncs.prograssCnt += 1
            miscFuncs.notification('LODING', f"{miscFuncs.prograssCnt / miscFuncs.prograssTotal}", None,
                                       f"变更点{temp_0}进行中... {temp_1}, {temp_2}")
        refWindowInfo = masterFuncs.getStringContext(filePath, 100, newSignalInfo["referenceSignal"])
        if config.config_debugMode == True:
            miscFuncs.notification('INFO', f"计算变更值")
        else:
            miscFuncs.prograssCnt += 1
            miscFuncs.notification('LODING', f"{miscFuncs.prograssCnt / miscFuncs.prograssTotal}", None,
                                       f"变更点{temp_0}进行中... {temp_1}, {temp_2}")
        copyWindowInfo = masterFuncs.getSingalSignalCopyWindowInfo("Project_EcucValues.arxml", refWindowInfo, newSignalInfo)
        if copyWindowInfo is not None:
            if config.config_debugMode == True:
                miscFuncs.notification('INFO', f"变更Project_EcucValues")
            else:
                miscFuncs.prograssCnt += 1
                miscFuncs.notification('LODING', f"{miscFuncs.prograssCnt / miscFuncs.prograssTotal}", None,
                                       f"变更点{temp_0}进行中... {temp_1}, {temp_2} ")
            masterFuncs.updataArxmlFile_Add(filePath, copyWindowInfo)
        else:
            miscFuncs.notification('ERROR', f"变更值copyWindow计算为空！.arxml")

        # 观察refWindow
        # i = 0
        # for temp in refWindow:
        #     i += 1
        #     miscFuncs.notification('DEBUG', f"refWindow NUM:{i}\n")
        #     for line in temp.lines:
        #         miscFuncs.notification('DEBUG', f"refWindow:{line}")

        # TO DO
        # 在此处加入回读函数校验
        rtn = True
        return rtn

    def __changeSiganlNameFuc(self, inputFilePath, outputFilePath, newSignalInfo):
        temp_old = newSignalInfo["referenceSignal"]
        temp_new = newSignalInfo["signalEnglishName"]

        try:
            with open(inputFilePath, 'r', encoding='utf-8') as inputFile, \
                    open(outputFilePath, 'w', encoding='utf-8') as outputFile:
                for line in inputFile:
                    if newSignalInfo["referenceSignal"] in line:
                        line = re.sub(f"{temp_old}", f"{temp_new}", line)
                    outputFile.write(line)

        except Exception as e:
            if os.path.exists(outputFilePath):
                os.remove(outputFilePath)
            raise e

        os.replace(outputFilePath, inputFilePath)
        miscFuncs.prograssCnt += 3

    def changeSignalNameTask(self, pyToolCurrentPath, newSignalInfo):
        # 检查新信号名称是否在ARXML中已存在，若已存在则需要手动变更
        rtn = masterFuncs.checkNewSiganlName(pyToolCurrentPath, newSignalInfo)
        # 新信号名称已经在ARXML中已存在，跳过这个变更项
        if rtn == 0: return 0

        inputFilePath = pyToolCurrentPath + config.config_inputFilePath + "\Project_EcucValues.arxml"
        outputFilePath = inputFilePath + '.tmp'
        self.__changeSiganlNameFuc(inputFilePath, outputFilePath, newSignalInfo)

        inputFilePath = pyToolCurrentPath + config.config_inputFilePath + "\APPCOM_SWC.arxml"
        outputFilePath = inputFilePath + '.tmp'
        self.__changeSiganlNameFuc(inputFilePath, outputFilePath, newSignalInfo)

        inputFilePath = pyToolCurrentPath + config.config_inputFilePath + "\DBC_SysDesc.arxml"
        outputFilePath = inputFilePath + '.tmp'
        self.__changeSiganlNameFuc(inputFilePath, outputFilePath, newSignalInfo)

        inputFilePath = pyToolCurrentPath + config.config_inputFilePath + "\MCU_SYSTEM.arxml"
        outputFilePath = inputFilePath + '.tmp'
        self.__changeSiganlNameFuc(inputFilePath, outputFilePath, newSignalInfo)

        return True

    # 删除信号任务
    def deleteSignalTask(self, pyToolCurrentPath, newSignalInfo):
        temp_0 = newSignalInfo["index"]
        temp_1 = newSignalInfo["signalChineseName"]
        temp_2 = newSignalInfo["changeType"]

        # 预处理ARXML文件
        masterFuncs.arxmlFilePreprocessing(pyToolCurrentPath, newSignalInfo)

        filePath = pyToolCurrentPath + config.config_inputFilePath + "\APPCOM_SWC.arxml"
        if config.config_debugMode == True:
            miscFuncs.notification('INFO', f"检索APPCOM_SWC")
        else:
            miscFuncs.prograssCnt += 1
            miscFuncs.notification('LODING', f"{miscFuncs.prograssCnt / miscFuncs.prograssTotal}", None,
                                       f"变更点{temp_0}进行中... {temp_1}, {temp_2}")
        refWindowInfo = masterFuncs.getStringContext(filePath, 40, newSignalInfo["referenceSignal"])
        if config.config_debugMode == True:
            miscFuncs.notification('INFO', f"计算变更值")
        else:
            miscFuncs.prograssCnt += 1
            miscFuncs.notification('LODING', f"{miscFuncs.prograssCnt / miscFuncs.prograssTotal}", None,
                                       f"变更点{temp_0}进行中... {temp_1}, {temp_2}")
        copyWindowInfo = masterFuncs.getSingalSignalCopyWindowInfo("APPCOM_SWC.arxml", refWindowInfo, newSignalInfo)
        if copyWindowInfo is not None:
            if config.config_debugMode == True:
                miscFuncs.notification('INFO', f"变更APPCOM_SWC")
            else:
                miscFuncs.prograssCnt += 1
                miscFuncs.notification('LODING', f"{miscFuncs.prograssCnt / miscFuncs.prograssTotal}", None,
                                       f"变更点{temp_0}进行中... {temp_1}, {temp_2} ")
            masterFuncs.updataArxmlFile_Delete(filePath, copyWindowInfo)
        else:
            miscFuncs.notification('ERROR', f"变更值copyWindow计算为空！")



        filePath = pyToolCurrentPath + config.config_inputFilePath + "\DBC_SysDesc.arxml"
        if config.config_debugMode == True:
            miscFuncs.notification('INFO', f"检索DBC_SysDesc")
        else:
            miscFuncs.prograssCnt += 1
            miscFuncs.notification('LODING', f"{miscFuncs.prograssCnt / miscFuncs.prograssTotal}", None,
                                       f"变更点{temp_0}进行中... {temp_1}, {temp_2}")
        refWindowInfo = masterFuncs.getStringContext(filePath, 44, newSignalInfo["referenceSignal"])
        if config.config_debugMode == True:
            miscFuncs.notification('INFO', f"计算变更值")
        else:
            miscFuncs.prograssCnt += 1
            miscFuncs.notification('LODING', f"{miscFuncs.prograssCnt / miscFuncs.prograssTotal}", None,
                                       f"变更点{temp_0}进行中... {temp_1}, {temp_2}")
        copyWindowInfo = masterFuncs.getSingalSignalCopyWindowInfo("DBC_SysDesc.arxml", refWindowInfo, newSignalInfo)
        if copyWindowInfo is not None:
            if config.config_debugMode == True:
                miscFuncs.notification('INFO', f"变更DBC_SysDesc")
            else:
                miscFuncs.prograssCnt += 1
                miscFuncs.notification('LODING', f"{miscFuncs.prograssCnt / miscFuncs.prograssTotal}", None,
                                       f"变更点{temp_0}进行中... {temp_1}, {temp_2} ")
            masterFuncs.updataArxmlFile_Delete(filePath, copyWindowInfo)
        else:
            miscFuncs.notification('ERROR', f"变更值copyWindow计算为空！")



        filePath = pyToolCurrentPath + config.config_inputFilePath + "\MCU_SYSTEM.arxml"
        if config.config_debugMode == True:
            miscFuncs.notification('INFO', f"检索MCU_SYSTEM")
        else:
            miscFuncs.prograssCnt += 1
            miscFuncs.notification('LODING', f"{miscFuncs.prograssCnt / miscFuncs.prograssTotal}", None,
                                       f"变更点{temp_0}进行中... {temp_1}, {temp_2}")
        refWindowInfo = masterFuncs.getStringContext(filePath, 40, newSignalInfo["referenceSignal"])
        if config.config_debugMode == True:
            miscFuncs.notification('INFO', f"计算变更值")
        else:
            miscFuncs.prograssCnt += 1
            miscFuncs.notification('LODING', f"{miscFuncs.prograssCnt / miscFuncs.prograssTotal}", None,
                                       f"变更点{temp_0}进行中... {temp_1}, {temp_2}")
        copyWindowInfo = masterFuncs.getSingalSignalCopyWindowInfo("MCU_SYSTEM.arxml", refWindowInfo, newSignalInfo)
        if copyWindowInfo is not None:
            if config.config_debugMode == True:
                miscFuncs.notification('INFO', f"变更MCU_SYSTEM")
            else:
                miscFuncs.prograssCnt += 1
                miscFuncs.notification('LODING', f"{miscFuncs.prograssCnt / miscFuncs.prograssTotal}", None,
                                       f"变更点{temp_0}进行中... {temp_1}, {temp_2} ")
            masterFuncs.updataArxmlFile_Delete(filePath, copyWindowInfo)
        else:
            miscFuncs.notification('ERROR', f"变更值copyWindow计算为空！")



        filePath = pyToolCurrentPath + config.config_inputFilePath + "\Project_EcucValues.arxml"
        if config.config_debugMode == True:
            miscFuncs.notification('INFO', f"检索Project_EcucValues")
        else:
            miscFuncs.prograssCnt += 1
            miscFuncs.notification('LODING', f"{miscFuncs.prograssCnt / miscFuncs.prograssTotal}", None,
                                       f"变更点{temp_0}进行中... {temp_1}, {temp_2}")
        refWindowInfo = masterFuncs.getStringContext(filePath, 100, newSignalInfo["referenceSignal"])
        if config.config_debugMode == True:
            miscFuncs.notification('INFO', f"计算变更值")
        else:
            miscFuncs.prograssCnt += 1
            miscFuncs.notification('LODING', f"{miscFuncs.prograssCnt / miscFuncs.prograssTotal}", None,
                                       f"变更点{temp_0}进行中... {temp_1}, {temp_2}")
        copyWindowInfo = masterFuncs.getSingalSignalCopyWindowInfo("Project_EcucValues.arxml", refWindowInfo, newSignalInfo)
        if copyWindowInfo is not None:
            if config.config_debugMode == True:
                miscFuncs.notification('INFO', f"变更Project_EcucValues")
            else:
                miscFuncs.prograssCnt += 1
                miscFuncs.notification('LODING', f"{miscFuncs.prograssCnt / miscFuncs.prograssTotal}", None,
                                       f"变更点{temp_0}进行中... {temp_1}, {temp_2} ")
            masterFuncs.updataArxmlFile_Delete(filePath, copyWindowInfo)
        else:
            miscFuncs.notification('ERROR', f"变更值copyWindow计算为空！.arxml")

        # 观察refWindow
        # i = 0
        # for temp in refWindow:
        #     i += 1
        #     miscFuncs.notification('DEBUG', f"refWindow NUM:{i}\n")
        #     for line in temp.lines:
        #         miscFuncs.notification('DEBUG', f"refWindow:{line}")

        # TO DO
        # 在此处加入回读函数校验
        rtn = True
        return rtn

    def __init__(self):
        miscFuncs.enableAnsiSupport()

    def mainTask(self):
        # 当前py路径
        pyCurrentPath = os.getcwd()

        # copy输入arxml到output
        masterFuncs.copyArxmlFile(pyCurrentPath + config.config_inputFilePath,
                                  pyCurrentPath + config.config_outputFilePath + "\ArxmlFile_Original")

        # 处理输入EXCEL,解析为detail列表
        inputFuncs.getChangeDicts(pyCurrentPath)

        # 获取进度条总进度值
        miscFuncs.getProgressTotal()

        # 新建变更模板
        masterFuncs.creatZentaoLog(pyCurrentPath)

        # 新建log文件
        masterFuncs.logFileCreate(pyCurrentPath)

        # 遍历信号变更细节列表
        masterFuncs.traverseChangePoints(pyCurrentPath)

        # 打印终端报错信息
        miscFuncs.printTerminalInfo()

        # 动态UUID
        # masterFuncs.uuidUpdata(pyCurrentPath)

        # copy变更后arxml到output
        masterFuncs.copyArxmlFile(pyCurrentPath + config.config_inputFilePath,
                                  pyCurrentPath + config.config_outputFilePath + "\ArxmlFile_Changed")
        # copy原arxml到input
        masterFuncs.copyArxmlFile(pyCurrentPath + config.config_outputFilePath + "\ArxmlFile_Original",
                                  pyCurrentPath + config.config_inputFilePath)

        # 结束
        miscFuncs.getUserInput("输入任意键结束:")

        return 0

config = config_t()
inputFuncs = inputFuncs_t()
masterFuncs = masterFuncs_t()
miscFuncs = miscFuncs_t()
task = task_t()

task.mainTask()

