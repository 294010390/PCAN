# 读写2007 excel
import openpyxl

StartPoint = 3

# For IB_MsgSig Sheet
Message = 'A'
CAN_ID = 'B'
Type = 'C'
DiagConnection = 'D'
Signal = 'E'  # Sig full name
Short_Name = 'F'  # Sig short name
Start_Byte = 'G'
Start_Bit = 'H'
Len = 'I'
Data = 'J'
Range = 'K'
Conversion = 'L'  # Sig Value
DLC = 'M'

# For IB_Tx Sheet
CycleMessage = 'C'
CycleCAN_ID = 'D'
CyclePeriodic = 'F'  # CyclePeriodic

Message_Table = {}
AllValue = []

def readIB_MsgSig(path):
    wb = openpyxl.load_workbook(path)
    sheet = wb.get_sheet_by_name('IB_MsgSig')
    # 每一列的标志

    CAN_ID_Name = "CAN_ID"
    Sig_Short_Name = "Sig_Short_Name"
    Sig_Value_Name = "Sig_Value_Name"

    global Message_Table
    global AllValue

    SigValueDict = {}
    RangeValueDict = {}
    LocationandDLCDic = {}
    SigNameList = []
    SigValueList = []

    PreMessageName = None


    # for row in sheet.rows:
    #     for cell in row:
    #         print(cell.value, "\t", end="")
    #     print()

    print(sheet.max_row)

    for i in range(StartPoint, sheet.max_row):

        # global SigNameList
        # global PreMessageName
        # global SigValueList
        # global SigValueDict
        # SigValueDict = {}

        # SigNameList = []

        # PreMessageName = None

        Value = []
        AllValue = []

        MessageName = sheet[Message + str(i)].value
        CAN_ID_Value = sheet[CAN_ID + str(i)].value
        Sig_ShortName_Value = sheet[Short_Name + str(i)].value
        DataValue = sheet[Data + str(i)].value
        RangeValue = sheet[Range + str(i)].value
        SigValue = sheet[Conversion + str(i)].value
        TypeValue = sheet[Type + str(i)].value
        StartByteValue = sheet[Start_Byte + str(i)].value
        StartBitValue = sheet[Start_Bit + str(i)].value
        LenValue = sheet[Len + str(i)].value
        DLCValue = sheet[DLC + str(i)].value

        if PreMessageName == None:
            PreMessageName = MessageName

        if PreMessageName == MessageName or MessageName == None:
            # MessageName = PreMessageName
            SigNameList.append(Sig_ShortName_Value)
            # SigValueList.append(SigValueDict)
        else:
            PreMessageName = MessageName
            SigNameList = []
            # SigValueList = []
            SigValueDict = {}
            RangeValueDict = {}
            LocationandDLCDic = {}
            SigNameList.append(Sig_ShortName_Value)
            # SigValueDict[Sig_ShortName_Value] = Value
            # PreMessageName = None
        LocationandDLCDic[Sig_ShortName_Value] = (StartByteValue, StartBitValue, LenValue, DLCValue)
        # if MessageName != "Message":
        # Message_Table[MessageName] = CAN_ID_Value


        # if PreMessageName == None:
        #     PreMessageName = MessageName

        if DataValue == 'BLN':
            Value = [("False", 0), ("True", 1)]
            SigValueDict[Sig_ShortName_Value] = Value
            RangeValueDict[Sig_ShortName_Value] = ['NA']
        elif DataValue == 'ENM':
            newSigValue = SigValue.split(";")
            for i in newSigValue:
                j = i.split('=')
                k = j[0][-1]
                v = j[1]
                Value.append((k, v))
            SigValueDict[Sig_ShortName_Value] = Value
            RangeValueDict[Sig_ShortName_Value] = ['NA']
        elif DataValue == 'SNM' or DataValue == 'UNM':
            Value = ['NA']
            SigValueDict[Sig_ShortName_Value] = Value
            nRangeValueList = RangeValue.split(' ')
            RangeValueDict[Sig_ShortName_Value] = (nRangeValueList[0], nRangeValueList[2])

        elif DataValue == 'PKT' or DataValue == 'ASC':
            Value = ['NA']
            SigValueDict[Sig_ShortName_Value] = Value
            RangeValueDict[Sig_ShortName_Value] = Value

        # elif DataValue == 'ASC':
        #     Value = ['NA']
        #     SigValueDict[Sig_ShortName_Value] = Value
        #     RangeValueList = Value

        AllValue.append(CAN_ID_Value)
        AllValue.append(TypeValue)
        AllValue.append(SigNameList)
        AllValue.append(RangeValueDict)
        AllValue.append(SigValueDict)
        AllValue.append(LocationandDLCDic)

        Message_Table[MessageName] = AllValue
        # Message_Table[MessageName] = CAN_ID_Value, TypeValue, SigNameList, RangeValueDict, SigValueDict, LocationandDLCDic



    # for k,v in Message_Table.items():
    #     print(k, "---", v)
    # # print(Message_Table)
    # print(len(Message_Table))

def readCycleTime(path):
    wb = openpyxl.load_workbook(path)
    sheet = wb.get_sheet_by_name('IB_Tx')

    global Message_Table
    global AllValue

    for i in range(StartPoint, sheet.max_row):
        AllValue = []

        CycleMessageValue = sheet[CycleMessage + str(i)].value
        CycleCAN_IDValue = sheet[CycleCAN_ID + str(i)].value
        CyclePeriodicValue = sheet[CyclePeriodic + str(i)].value

        if CyclePeriodicValue == '0':
            CyclePeriodicValue = '10.0'

        if CycleMessageValue != None:
            # Message_Table[CycleMessageValue].append(CycleCAN_IDValue)
            Message_Table[CycleMessageValue].append(CyclePeriodicValue)

            print('My AllValue is ', Message_Table[CycleMessageValue])

        # Message_Table[CycleMessageValue] = AllValue
    for k, v in Message_Table.items():
        print(k, "---", v)

if __name__ == '__main__':
    dbcfilepath = r'D:\15.NM\DBC\20.20\CLEA_Family_I-CAN_v20.18.0_ICI2.xlsx'
    readIB_MsgSig(dbcfilepath)
    readCycleTime(dbcfilepath)

