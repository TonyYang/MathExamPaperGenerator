# _*_ coding: utf-8 _*_


from win32com.client import Dispatch

import random, time, os


###############################################################################
## EPG_XlsGen Class                                                          ##
## - Generate testing questions into Excel file.                             ##
###############################################################################
class EPG_GenXls:
    def __init__(self, pFileName):
        self.mFileName = pFileName
        self.mXlsApp = Dispatch("Excel.Application")
        self.mXlsApp.Visible = False
        self.mXlsApp.DisplayAlerts = False
        
        if os.path.exists(self.mFileName):
            os.remove(self.mFileName)
        
        self.mXlsBook = self.mXlsApp.Workbooks.Add()
        self.mXlsSheet = self.mXlsBook.Worksheets.Add()
        
    def writeQuestion(self, pIdx, pContent):
        tRow = int(pIdx) / 5 + 2
        tCol = int(pIdx) % 5 + 1
        self.mXlsSheet.Cells(tRow, tCol).Value = pContent
        
    def writeComment(self, pExamCalParamNumber, pExamQuestionNumber):
        tTime = pExamQuestionNumber * 5 / 60 * (pExamCalParamNumber / 2)
        tComments = "在" + str(tTime) + "分钟内完成!!"
        self.mXlsSheet.Cells(1, 1).Value = tComments.decode("utf8").encode("gbk")
        
    def formatContent(self, pExamQuestionNumber):
        # Format column width
        self.mXlsSheet.Range(self.mXlsSheet.Cells(1, 1), self.mXlsSheet.Cells(1, 5)).ColumnWidth = 25
        # Format font
        self.mXlsSheet.Range(self.mXlsSheet.Cells(1, 1), self.mXlsSheet.Cells(pExamQuestionNumber / 5 + 1, 5)).Font.Bold = True
        self.mXlsSheet.Range(self.mXlsSheet.Cells(1, 1), self.mXlsSheet.Cells(pExamQuestionNumber / 5 + 1, 5)).Font.Name = "Calibri"
        self.mXlsSheet.Range(self.mXlsSheet.Cells(1, 1), self.mXlsSheet.Cells(pExamQuestionNumber / 5 + 1, 5)).Font.Size = 16
        # Format border
        for iRow in range(1, pExamQuestionNumber / 5 + 2):
            for iCol in range(1, 6):
                for iBorder in range(1, 5):
                    self.mXlsSheet.Cells(iRow, iCol).Borders(iBorder).LineStyle = 1
        
    def saveFile(self):
        self.mXlsBook.SaveAs(self.mFileName)
        self.mXlsBook.Close(SaveChanges = 0)
        del self.mXlsApp


###############################################################################
## 'EPG_ExamAdd' Class                                                       ##
## - Generate addition calculation examination paper.                        ##
###############################################################################
class EPG_ExamAdd:
    def __init__(self, pExamCalParamNumber, pExamCalRange, pExamQuestionNumber):
        self.mExamCalParamNumber = pExamCalParamNumber
        self.mExamCalRange       = pExamCalRange
        self.mExamQuestionNumber = pExamQuestionNumber
        
        self.mFileName = "Addition_" + str(self.mExamCalParamNumber) + "_" + str(self.mExamCalRange) + "_" + str(self.mExamQuestionNumber) + "_" + time.strftime("%Y%m%d") + "_" + time.strftime("%H%M%S")
        self.mFileName = os.getcwd() + "\\" + self.mFileName
        
        print "[EPG_ExamAdd] Excel file name is %s!! \n" % self.mFileName        
        
        self.mGenXls = EPG_GenXls(self.mFileName)
        
    def genExamPaper(self):
        for tQuestionIdx in range(self.mExamQuestionNumber):
            print "[EPG_ExamAdd] Generate No.%d question!! \n" % tQuestionIdx
            tCalParam  = 0
            tQuestion  = ""
            tCalResult = 0
            tSuccess   = False
            
            while not tSuccess:
                for tCalParamIdx in range(self.mExamCalParamNumber):
                    print "[EPG_ExamAdd] Generate No.%d parameter!! \n" % tCalParamIdx
                    
                    if tCalParamIdx == 0:
                        tCalParam  = random.randint(1, self.mExamCalRange -1)
                        tQuestion  = tQuestion + str(tCalParam)
                        tCalResult = tCalParam
                        tSuccess   = True
                    elif (self.mExamCalRange - tCalResult) > 1:
                        tCalParam  = random.randint(1, self.mExamCalRange - tCalResult - 1)
                        tQuestion  = tQuestion + " + " + str(tCalParam)
                        tCalResult = tCalResult + tCalParam
                        tSuccess   = True
                    else:
                        print "[EPG_ExamAdd] Generating question failed!! Re-trying!! \n"
                        tCalParam  = 0
                        tQuestion  = ""
                        tCalResult = 0
                        tSuccess   = False
                        break
            
            tQuestion = tQuestion + " = "
            self.mGenXls.writeQuestion(tQuestionIdx, tQuestion)
            
        self.mGenXls.writeComment(self.mExamCalParamNumber, self.mExamQuestionNumber)
        self.mGenXls.formatContent(self.mExamQuestionNumber)
        self.mGenXls.saveFile()
        
        print "[EPG_ExamAdd] Generating examination pager completed!! \n"


###############################################################################
## 'EPG_ExamSub' Class                                                       ##
## - Generate subtraction calculation questions.                             ##
###############################################################################
class EPG_ExamSub:
    def __init__(self, pExamCalParamNumber, pExamCalRange, pExamQuestionNumber):
        self.mExamCalParamNumber = pExamCalParamNumber
        self.mExamCalRange       = pExamCalRange
        self.mExamQuestionNumber = pExamQuestionNumber
        
        self.mFileName = "Subtraction_" + str(self.mExamCalParamNumber) + "_" + str(self.mExamCalRange) + "_" + str(self.mExamQuestionNumber) + "_" + time.strftime("%Y%m%d") + "_" + time.strftime("%H%M%S")
        self.mFileName = os.getcwd() + "\\" + self.mFileName
        
        print "[EPG_ExamSub] Excel file name is %s!!" % self.mFileName
        
        self.mGenXls = EPG_GenXls(self.mFileName)
    
    def genExamPaper(self):
        for tQuestionIdx in range(self.mExamQuestionNumber):
            print "[EPG_ExamSub] Generate No.%d question!! \n" % tQuestionIdx
            tCalParam  = 0
            tQuestion  = ""
            tCalResult = 0
            tSuccess   = False
            
            while not tSuccess:
                for tCalParamIdx in range(self.mExamCalParamNumber):
                    print "[EPG_ExamSub] Generate No.%d parameter!! \n" % tCalParamIdx
                    
                    if tCalParamIdx == 0:
                        tCalParam  = random.randint(1, self.mExamCalRange - 1)
                        tQuestion  = tQuestion + str(tCalParam)
                        tCalResult = tCalParam
                        tSuccess   = True
                    elif tCalResult > 1:
                        tCalParam  = random.randint(1, tCalResult - 1)
                        tQuestion  = tQuestion + " - " + str(tCalParam)
                        tCalResult = tCalResult - tCalParam
                        tSuccess   = True
                    else:
                        print "[EPG_ExamSub] Generating question failed!! Re-trying!! \n"
                        tCalParam  = 0
                        tQuestion  = ""
                        tCalResult = 0
                        tSuccess   = False
                        break
                        
            tQuestion = tQuestion + " = "
            self.mGenXls.writeQuestion(tQuestionIdx, tQuestion)

        self.mGenXls.writeComment(self.mExamCalParamNumber, self.mExamQuestionNumber)
        self.mGenXls.formatContent(self.mExamQuestionNumber)
        self.mGenXls.saveFile()
        
        print "[EPG_ExamSub] Generating examination paper completed!! \n"


###############################################################################
## 'EPG_ExamAddSub' Class                                                    ##
## - Generate addition calculation questions randomly.                       ##
## - Generate subtraction calculation questions randomly.                    ##
###############################################################################
class EPG_ExamAddSub:
    def __init__(self, pExamCalParamNumber, pExamCalRange, pExamQuestionNumber):
        self.mExamCalParamNumber = pExamCalParamNumber
        self.mExamCalRange       = pExamCalRange
        self.mExamQuestionNumber = pExamQuestionNumber
        
        self.mFileName = "AddSub_" + str(self.mExamCalParamNumber) + "_" + str(self.mExamCalRange) + "_" + str(self.mExamQuestionNumber) + "_" + time.strftime("%Y%m%d") + "_" + time.strftime("%H%M%S")
        self.mFileName = os.getcwd() + "\\" + self.mFileName
        
        print "[EPG_ExamAddSub] Excel file name is %s!!" % self.mFileName        
        
        self.mGenXls = EPG_GenXls(self.mFileName)
        
    def genExamPaper(self):
        for tQuestionIdx in range(self.mExamQuestionNumber) :
            print "[EPG_ExamAddSub] Generate No.%d question!! \n" % tQuestionIdx
            
            tCalResult = 0
            tCalType   = 0
            tCalParam  = 0
            tQuestion  = ""
            tSuccess   = False
            
            while not tSuccess:
                for tParamIdx in range(self.mExamCalParamNumber):
                    print "[EPG_ExamAddSub] Generate No.%d parameter!! \n" % tParamIdx
                    
                    if tParamIdx == 0:
                        tCalParam  = random.randint(1, self.mExamCalRange - 1)
                        tQuestion  = tQuestion + str(tCalParam)
                        tCalResult = tCalParam
                        tSuccess   = True
                    else:
                        tCalType = random.randint(1, 2)
                        
                        if tCalType == 1 and (self.mExamCalRange - tCalResult ) > 1:
                            tCalParam  = random.randint(1, self.mExamCalRange - tCalResult - 1)
                            tQuestion  = tQuestion + " + " + str(tCalParam)
                            tCalResult = tCalResult + tCalParam
                            tSuccess   = True
                        elif tCalType == 2 and tCalResult > 1:
                            tCalParam  = random.randint(1, tCalResult - 1)
                            tQuestion  = tQuestion + " - " + str(tCalParam)
                            tCalResult = tCalResult - tCalParam
                            tSuccess   = True
                        else:
                            tCalResult = 0
                            tCalType   = 0
                            tCalParam  = 0
                            tQuestion  = ""
                            tSuccess   = False
                            break
            
            tQuestion = tQuestion + " = "
            self.mGenXls.writeQuestion(tQuestionIdx, tQuestion)

        self.mGenXls.writeComment(self.mExamCalParamNumber, self.mExamQuestionNumber)
        self.mGenXls.formatContent(self.mExamQuestionNumber)
        self.mGenXls.saveFile()
        
        print "[EPG_ExamAddSub] Generate examination paper completed!! \n"


###############################################################################
## 'EPG_ExamMul' Class                                                       ##
## - Generate multiplication calculation questions randomly.                 ##
###############################################################################
class EPG_ExamMul:
    def __init__(self, pExamCalParamNum, pExamCalRange, pExamQuestionNum):
        self.mExamCalParamNum = pExamCalParamNum
        self.mExamCalRange    = pExamCalRange
        self.mExamQuestionNum = pExamQuestionNum
        
        self.mFileName = "Mul_" + str(self.mExamCalParamNum) + "_" + str(self.mExamCalRange) + "_" + str(self.mExamQuestionNum) + "_" + time.strftime("%Y%m%d") + "_" + time.strftime("%H%M%S")
        self.mFileName = os.getcwd() + "\\" + self.mFileName
        
        print "[EPG_ExamMul] Excel file name is %s!!" % self.mFileName
        
        self.mGenXls = EPG_GenXls(self.mFileName)
        
    def genExamPaper(self):
        for tQuestionIdx in range(self.mExamQuestionNum):
            print "[EPG_ExamMul] Generate No.%d question!! \n" % tQuestionIdx
            tCalParam  = 0
            tQuestion  = ""

            print "[EPG_ExamMul] Generate No.1 parameter!! \n"
            tCalParam = random.randint(1, self.mExamCalRange)
            tQuestion = tQuestion + str(tCalParam)

            print "[EPG_ExamMul] Generate No.2 parameter!! \n"            
            tCalParam = random.randint(1, self.mExamCalRange)
            tQuestion = tQuestion + " x " + str(tCalParam)
                        
            tQuestion = tQuestion + " = "
            self.mGenXls.writeQuestion(tQuestionIdx, tQuestion)

        self.mGenXls.writeComment(self.mExamCalParamNum, self.mExamQuestionNum)
        self.mGenXls.formatContent(self.mExamQuestionNum)
        self.mGenXls.saveFile()
        
        print "[EPG_ExamMul] Generating examination paper completed!! \n"


###############################################################################
## 'EPG_ExamAddSubMul' Class                                                 ##
## - Generate Addition calculation questions randomly.                       ##
## - Generate Subtraction calculation questions randomly.                    ##
## - Generate multiplication calculation questions randomly.                 ##
###############################################################################
class EPG_ExamAddSubMul:
    def __init__(self, pExamAddSubMulCalParamNum, pExamAddSubCalRange, pExamMulCalRange, pExamQuestionNum):
        self.mExamAddSubMulCalParamNum = pExamAddSubMulCalParamNum
        self.mExamAddSubCalRange       = pExamAddSubCalRange
        self.mExamMulCalRange          = pExamMulCalRange
        self.mExamQuestionNum          = pExamQuestionNum
        
        self.mFileName = "AddSubMul_" + str(self.mExamAddSubMulCalParamNum) + "_" + str(self.mExamAddSubCalRange) + "_" + "_" + str(self.mExamMulCalRange) + "_" + str(self.mExamQuestionNum) + "_" + time.strftime("%Y%m%d") + "_" + time.strftime("%H%M%S")
        self.mFileName = os.getcwd() + "\\" + self.mFileName
        
        print "[EPG_ExamAddSubMul] Excel file name is %s!!" % self.mFileName
        
        self.mGenXls = EPG_GenXls(self.mFileName)

    def genExamPaper(self):
        for tQuestionIdx in range(self.mExamQuestionNum):
            print "[EPG_ExamAddSubMul] Generate No.%d question!! \n" % tQuestionIdx
            tCalParam1 = 0
            tCalParam2 = 0
            tCalType   = 0
            tCalSymbol = ""
            tQuestion  = ""
            
            tCalType = random.randint(1, 3)
            
            ## Addition calculation.
            if tCalType == 1:
                tCalParam1 = random.randint(1, self.mExamAddSubCalRange - 1)
                tCalParam2 = random.randint(1, self.mExamAddSubCalRange - tCalParam1)
                tCalSymbol = " + "

            ## Subtraction calculation    
            elif tCalType == 2:
                tCalParam1 = random.randint(2, self.mExamAddSubCalRange - 1)
                tCalParam2 = random.randint(1, tCalParam1 - 1)
                tCalSymbol = " - "

            ## Multiplication calculation
            else:
                tCalParam1 = random.randint(1, self.mExamMulCalRange)
                tCalParam2 = random.randint(1, self.mExamMulCalRange)
                tCalSymbol = " x "
                
            tQuestion = tQuestion + str(tCalParam1) + tCalSymbol + str(tCalParam2) + " = "
            self.mGenXls.writeQuestion(tQuestionIdx, tQuestion)
            
        self.mGenXls.writeComment(self.mExamAddSubMulCalParamNum, self.mExamQuestionNum)
        self.mGenXls.formatContent(self.mExamQuestionNum)
        self.mGenXls.saveFile()
        
        print "[EPG_ExamAddSubMul] Generating examination paper completed!! \n"


###############################################################################
## Program Entry                                                             ##
###############################################################################
if __name__ == "__main__":
    print '''
\n
Please choose calculation type:
1 - Addition
2 - Subtraction
3 - Addition-Subtraction mixed
4 - Multiplication
5 - Addition-Subtraction-Multiplication mixed
6 - Division [N/A]
7 - Multiplication-Division mixed [N/A]
8 - Addition-Subtraction-Multiplication-Division mixed [N/A]
          '''
    
    ## Ask user to select calculation type.
    tExamType = int(raw_input("Please input calculation type index (1, 2, 3, ...): "))
    
    ## Ask user to input how many questions will be generated in one paper.
    print "\n"
    tExamQuestionNum = int(raw_input("Please input the number of calculation questions (50, 100, ...): "))
        
    ## Ask user to input how many examination papers will be generated.
    print "\n"
    tExamPaperNum = int(raw_input("Please input the number of examination papers (20, 40, ...): "))
    
    ## Ask user to input calculation parameter and range.
    if tExamType == 1 or tExamType == 2 or tExamType == 3:
        print "\n"
        tExamAddSubCalParamNum = int(raw_input("Please input the number of add-sub calculation parameters (2, 3, 4, ...): "))
        
        print "\n"
        tExamAddSubCalRange = int(raw_input("Please input the range of add-sub calculation (within 10, within 20, within 30, ...): "))

    elif tExamType == 4:
        tExamMulCalParamNum = 2
        
        print "\n"
        tExamMulCalRange = int(raw_input("Please input the range of multiplication calculation (within 5, 9, 12, ...): "))

    elif tExamType == 5:
        tExamAddSubMulCalParamNum = 2
        
        print "\n"
        tExamAddSubCalRange = int(raw_input("Please input the range of add-sub calculation (within 10, within 20, within 30, ...): "))
        
        print "\n"
        tExamMulCalRange = int(raw_input("Please input the range of multiplication calculation (within 5, 9, 12, ...): "))

    else:
        print "\n"
        print "The calculation type is NOT supported!!"

    ## Generate examination papers.
    for tExamPapgerIdx in range(tExamPaperNum):
        if tExamType == 1:
            tAdd = EPG_ExamAdd(tExamAddSubCalParamNum, tExamAddSubCalRange, tExamQuestionNum)
            tAdd.genExamPaper()
        
        elif tExamType == 2:
            tSub = EPG_ExamSub(tExamAddSubCalParamNum, tExamAddSubCalRange, tExamQuestionNum)
            tSub.genExamPaper()
        
        elif tExamType == 3:
            tAddSub = EPG_ExamAddSub(tExamAddSubCalParamNum, tExamAddSubCalRange, tExamQuestionNum)
            tAddSub.genExamPaper()
        
        elif tExamType == 4:
            tMul = EPG_ExamMul(tExamMulCalParamNum, tExamMulCalRange, tExamQuestionNum)
            tMul.genExamPaper()
        
        elif tExamType == 5:
            tAddSubMul = EPG_ExamAddSubMul(tExamAddSubMulCalParamNum, tExamAddSubCalRange, tExamMulCalRange, tExamQuestionNum)
            tAddSubMul.genExamPaper()
            
        else:
            print "\n\n"
            print "The calculation type is NOT supported!!"
