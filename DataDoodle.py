import pyautogui as auto
import openpyxl
from openpyxl import load_workbook
from openpyxl.workbook import Workbook
import sys

def Taking_TextFile_Path():
    try:
        Text_File_Path = auto.prompt("Path Of Text File","Input Here...")
        if Text_File_Path.isspace():
            auto.alert("\tInvalid Input\t ","Try Again..")
            Taking_TextFile_Path()
        elif len(Text_File_Path) < 3 :
            auto.alert("\tInvalid Input\t","Try Again..")
            Taking_TextFile_Path()
        elif Text_File_Path == 'None':
            sys.exit()
        else:
            Text_File_Path = Text_File_Path.strip(''' "',.''')
            Text_File_Path = Text_File_Path.replace("\\", "\\\\")
            TakingExcelSheet_Path(Text_File_Path)
    except:
        sys.exit()

def TakingExcelSheet_Path(Text_File_Path):
    try:
        Excel_Sheet_Path = auto.prompt("Path Of Excel Sheet","Input Here...")
        if Excel_Sheet_Path.isspace():
            auto.alert("\tInvalid Input\t","Try Again..")
            TakingExcelSheet_Path(Text_File_Path)
        elif len(Excel_Sheet_Path) < 3 :
            auto.alert("\tInvalid Input\t","Try Again..")
            TakingExcelSheet_Path(Text_File_Path)
        elif Text_File_Path == 'None':
            sys.exit()
        else:
            Excel_Sheet_Path = Excel_Sheet_Path.strip(''' "',.''')
            Excel_Sheet_Path = Excel_Sheet_Path.replace("\\", "\\\\")
            DeleteAndCreateSheet(Text_File_Path,Excel_Sheet_Path)
    except:
        sys.exit()

def DeleteAndCreateSheet(Text_File_Path, Excel_Sheet_Path):
    try :
        # Open the Excel file
        workbook = load_workbook(Excel_Sheet_Path)

        # Delete all existing sheets
        for sheet_name in workbook.sheetnames:
            workbook.remove(workbook[sheet_name])

        # Create a new sheet with a specific name
        workbook.create_sheet(title='Detailed Data')
        workbook.create_sheet(title='RewardsNumber')

        # Save the modified Excel file
        workbook.save(Excel_Sheet_Path)
    except : 
        pass
    DetailedData(Text_File_Path,Excel_Sheet_Path)

def DetailedData(Text_File_Path,Excel_Sheet_Path):
    try:
        Data_List = []
        with open(Text_File_Path, 'r',errors='ignore',encoding='utf8') as MemberData:
            for line in MemberData:
                try:
                    All_Data = line[line.find("reportOutput"):]
                    Loops_End = All_Data.count("jobId")
                    Required_Data = All_Data
                    for i in range (Loops_End):
                        if '}' in Required_Data:
                            Brace_End = Required_Data.find('}')+1
                        elif "[truncated]" in Required_Data:
                            Brace_End =  Required_Data.find("[truncated]")

                        One_Brace = Required_Data[Required_Data.find("{"):Brace_End]
                        One_Brace.strip("")

                        Actual_Job_Id = Job_ID(One_Brace)
                        Actual_Object = Entity_Type(One_Brace)
                        Actual_Operation = Operations(One_Brace)
                        Actual_Processed_Records = Processed_Records(One_Brace)
                        Actual_Failed_Records = Failed_Records(One_Brace)
                        Actual_Error_Reason = Error_Reason(One_Brace)
                        Data_List.append([Actual_Job_Id,Actual_Object,Actual_Operation,Actual_Processed_Records,Actual_Failed_Records,Actual_Error_Reason])
                        Required_Data = Required_Data[Brace_End:]
                except:
                    pass
            excelFile = openpyxl.load_workbook(Excel_Sheet_Path)
            detailedData = excelFile["Detailed Data"]
            detailedData.append(["Job Id","Object","Operation","Processed Records","Failed Records","Error Reason"])
            for data in Data_List:
                detailedData.append(data)
            excelFile.save(Excel_Sheet_Path)

            RewardNumber(Text_File_Path,Excel_Sheet_Path)   # Filtering RewardsNumber
            auto.alert("\tDone ✅ \t ","Successful..")
    except:
        auto.alert("\tError Occured\t","Failed..")
        Taking_TextFile_Path()
        pass

def Job_ID(One_Brace):
    if "jobId" in One_Brace:
        indexNo = One_Brace.find("jobId")
        try:
            Actual_Job_id = One_Brace[indexNo+5:indexNo+30]
            Actual_Job_id = Actual_Job_id.strip(''',. : '".,''')
            return Actual_Job_id
        except:
            return "Null"
    else:
        return None

def Entity_Type(One_Brace):
    if "entityType" in One_Brace:
        indexNo = One_Brace.find("entityType")
        try:
            Object = One_Brace[indexNo+10:indexNo+33]
            Object = Object.strip(''',. : '".,''')
            if "ContactMemberMaster" in Object:
                Contact__Index = Object.find("ContactMemberMaster")
                Actual_Object = "ContactMemberMaster__c"
            elif "Contact" in Object:
                Contact__Index = Object.find("Contact")
                Actual_Object = Object[Contact__Index:Contact__Index+7]
            return Actual_Object
        except:
            return "Null"
    else:
        return "Null"


def Operations(One_Brace):
    if "operation" in One_Brace:
        indexNo = One_Brace.find("operation")
        try:
            Actual_Operation = One_Brace[indexNo+13:indexNo+24]
            Actual_Operation = Actual_Operation[:Actual_Operation.find('''"''')]
            Actual_Operation = Actual_Operation.strip(''',. : '".,''')
            return Actual_Operation
        except:
            return "Null"
    else:
        return "Null"

def Processed_Records(One_Brace):
    if "processedRecords" in One_Brace:
        indexNo = One_Brace.find("processedRecords")
        try:
            Actual_Processed_Records = One_Brace[indexNo+16:indexNo+28]
            Actual_Processed_Records = Actual_Processed_Records.strip('''!~`@#$%^&*(-_+=[{)}]|\;></?,. : "'abcdefghijklmnopqrstuvwxyx.,ABCDEFGHIJKLMNOPQRSTUVWXYZ''')
            if int(Actual_Processed_Records)>=0:
                return int(Actual_Processed_Records)
            else:
               return 0
        except:
            return "Null"
    else:
        return "Null"
    
def Failed_Records(One_Brace):
    if "failedRecords" in One_Brace:
        indexNo = One_Brace.find("failedRecords")
        try:
            Actual_Failed_Records = One_Brace[indexNo+16:indexNo+20]
            Actual_Failed_Records = Actual_Failed_Records.strip('''!~`@#$%^&*(-_+=[{)}]|\;></?,. : "'abcdefghijklmnopqrstuvwxyx.,ABCDEFGHIJKLMNOPQRSTUVWXYZ''')

            if int(Actual_Failed_Records)>=0:
                return int(Actual_Failed_Records)
            else:
               return 0
        except:
            return "Null"
    else:
        return "Null"

def Error_Reason(One_Brace):
    if "errorsList" in One_Brace:
        indexNo = One_Brace.find("errorsList")
        try:
            Actual_Error_Reason = One_Brace[indexNo+10:One_Brace.find('''"}''')]
            Actual_Error_Reason = Actual_Error_Reason.strip(''',. : '".,''')
    
            if Actual_Error_Reason == "No errors":
                return Actual_Error_Reason
            else:
                if "RewardsNumber" in One_Brace:
                    Reward_Number_Count = One_Brace.count("RewardsNumber")
                    Real_Error_Reason = ""
                    for i in range(Reward_Number_Count):
                        indexNo = One_Brace.find("RewardsNumber")
                        Rewards_Number = One_Brace[indexNo+13:indexNo+28]
                        Rewards_Number = Rewards_Number.strip(",. : .,")
                        One_Brace = One_Brace[One_Brace.find(Rewards_Number):]
    
                        if "--" in One_Brace:
                            Find_Hyfen = One_Brace.find("--")
                            error_Reason = One_Brace[0:Find_Hyfen+2]
                            error_Reason = error_Reason.strip(""" 1234567890,.;:}[]{\|!@#$%^&*(+=)""")
                            if i == Reward_Number_Count - 1 :
                                Real_Error_Reason = Real_Error_Reason + error_Reason
                            else :
                                Real_Error_Reason = Real_Error_Reason + error_Reason + "\n"
                return Real_Error_Reason
        except:
            return "Null"
            
    elif "errorIfAny" in One_Brace :
        indexNo = One_Brace.find("errorIfAny")
        try:
            One_Brace = One_Brace[indexNo:]
            Actual_Error_Reason = One_Brace[10:One_Brace.find(''',"''')]
            Actual_Error_Reason = Actual_Error_Reason.strip(''',. : '".,''')
            return Actual_Error_Reason 
        except:
            return "Null"
    else:
        return "Null"

# Filtering the rewardsNumber with the error reason
def RewardNumber(Text_File_Path,Excel_Sheet_Path):
    try:
        rowNo = int(2)
        excelFile = openpyxl.load_workbook(Excel_Sheet_Path)
        reportOutput = excelFile["RewardsNumber"]
        Excel_Column = reportOutput['B1']
        Excel_Column.value = "Rewards Number"

        with open(Text_File_Path, 'r',errors='ignore') as MemberData:
            for line in MemberData:
                a = line.strip()
                Reward_Number_Count = line.count("RewardsNumber")
                for i in range(Reward_Number_Count):
                    if "RewardsNumber" in a:
                        indexNo = a.find("RewardsNumber")
                        Rewards_Number = a[indexNo+13:indexNo+28]
                        Rewards_Number = Rewards_Number.strip(",. : .,")
                        Excel_Col = reportOutput['B'+str(rowNo)]
                        rowNo += 1
                        Excel_Col.value = Rewards_Number
                        a = a[indexNo+27:]
        excelFile.save(Excel_Sheet_Path)
        ErrorReason(Text_File_Path, Excel_Sheet_Path)
    except:
        pass

def ErrorReason(Text_File_Path,Excel_Sheet_Path):
    try:
        excelFile = openpyxl.load_workbook(Excel_Sheet_Path)
        reportOutput = excelFile["RewardsNumber"]
        Excel_Column = reportOutput['A1']
        Excel_Column.value = "Error Reason"
        Reward_Number = []
        Reward_Number_Count = 0
        rowNo = 2

        for row in range(1, reportOutput.max_row):
            for col in reportOutput.iter_cols(2):
                Number = col[row].value
                Reward_Number.append(Number)

        with open(Text_File_Path, 'r',errors='ignore',encoding='utf8') as MemberData:
            for line in MemberData:
                a = line.strip(" ")
                for i in range(len(Reward_Number)):
                    try:
                        if Reward_Number[Reward_Number_Count] in a:
                            indexNo = a.find(Reward_Number[Reward_Number_Count])    
                            reason = a[indexNo:indexNo+200]
                            if "--" in reason:
                                Find_Hyfen = reason.find("--")

                                Real_Reason = reason[0:Find_Hyfen+2]
                                Real_Reason = Real_Reason.strip(""" 1234567890,.;:[]{}\|!@#$%^&*(+=)""")

                                ErrorReasonUpdate = reportOutput['A'+str(rowNo)]
                                rowNo += 1
                                ErrorReasonUpdate.value = Real_Reason   
                                Reward_Number_Count += 1
                    except:
                        pass
        excelFile.save(Excel_Sheet_Path)
        Process_Record(Text_File_Path, Excel_Sheet_Path)
    except:
        pass

def Process_Record(Text_File_Path,Excel_Sheet_Path):
    try:
        excelFile = openpyxl.load_workbook(Excel_Sheet_Path)
        processedRecords = excelFile["RewardsNumber"]
        Excel_Column = processedRecords['C1']
        Excel_Column.value = "Processed Records"
        processed_Records_Sum = []
        with open(Text_File_Path, 'r',errors='ignore') as MemberMaster:
            for line in MemberMaster:
                a = line.strip()
                ProcessedRecords_Count = line.count('''processedRecords''')
                for i in range(ProcessedRecords_Count):
                    if "processedRecords"  in a:
                        indexNo = a.find("processedRecords")
                        processed_Records = a[indexNo+16:indexNo+28]
                        processed_Records = processed_Records.strip('''!~`@#$%^&*(-_+=[{)}]|\;></?,. : "'abcdefghijklmnopqrstuvwxyx.,ABCDEFGHIJKLMNOPQRSTUVWXYZ''')
                        if len(processed_Records)>0:
                            processed_Records_Sum.append(int(processed_Records))
                        else:
                            processed_Records_Sum.append(0)
                        a = a[indexNo+27:]
        Excel_Col = processedRecords['C2']
        Excel_Col.value = sum(processed_Records_Sum)
        excelFile.save(Excel_Sheet_Path)
    except :
        pass

auto.alert("\tWelcome\t\t","DataDoodle")
Taking_TextFile_Path()