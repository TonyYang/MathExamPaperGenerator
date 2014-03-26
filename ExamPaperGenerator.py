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
## 'EPG_ExamAddition' Class                                                  ##
## - Generate addition calculation examination paper.                        ##
###############################################################################
class EPG_ExamAddition:
    def __init__(self, pExamCalParamNumber, pExamCalRange, pExamQuestionNumber):
        self.mExamCalParamNumber = pExamCalParamNumber
        self.mExamCalRange       = pExamCalRange
        self.mExamQuestionNumber = pExamQuestionNumber
        
        self.mFileName = "Addition_" + str(self.mExamCalParamNumber) + "_" + str(self.mExamCalRange) + "_" + str(self.mExamQuestionNumber) + "_" + time.strftime("%Y%m%d") + "_" + time.strftime("%H%M%S")
        self.mFileName = os.getcwd() + "\\" + self.mFileName
        
        print "EPG_ExamAddition XLS file name is %s!! \n" % self.mFileName        
        
        self.mGenXls = EPG_GenXls(self.mFileName)
        
    def genExamPaper(self):
        for tQuestionIdx in range(self.mExamQuestionNumber):
            print "EPG_ExamAddition generate No.%d question!! \n" % tQuestionIdx
            tCalParam  = 0
            tQuestion  = ""
            tCalResult = 0
            tSuccess   = False
            
            while not tSuccess:
                for tCalParamIdx in range(self.mExamCalParamNumber):
                    print "EPG_ExamAddition generate No.%d parameter!! \n" % tCalParamIdx
                    
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
                        print "EPG_ExamAddition generate question failed!! Re-trying!! \n"
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
        
        print "EPG_ExamAddition generate examination pager completed!! \n"


###############################################################################
## 'EPG_ExamSubtractio' Class                                                ##
## - Generate subtraction calculation questions.                             ##
###############################################################################
class EPG_ExamSubtraction:
    def __init__(self, pExamCalParamNumber, pExamCalRange, pExamQuestionNumber):
        self.mExamCalParamNumber = pExamCalParamNumber
        self.mExamCalRange       = pExamCalRange
        self.mExamQuestionNumber = pExamQuestionNumber
        
        self.mFileName = "Subtraction_" + str(self.mExamCalParamNumber) + "_" + str(self.mExamCalRange) + "_" + str(self.mExamQuestionNumber) + "_" + time.strftime("%Y%m%d") + "_" + time.strftime("%H%M%S")
        self.mFileName = os.getcwd() + "\\" + self.mFileName
        
        print "EPG_ExamSubtraction XLS file name is %s!!" % self.mFileName
        
        self.mGenXls = EPG_GenXls(self.mFileName)
    
    def genExamPaper(self):
        for tQuestionIdx in range(self.mExamQuestionNumber):
            print "EPG_ExamSubtraction generate No.%d question!! \n" % tQuestionIdx
            tCalParam  = 0
            tQuestion  = ""
            tCalResult = 0
            tSuccess   = False
            
            while not tSuccess:
                for tCalParamIdx in range(self.mExamCalParamNumber):
                    print "EPG_ExamSubtraction generate No.%d parameter!! \n" % tCalParamIdx
                    
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
                        print "EPG_ExamSubtraction generate question failed!! Re-trying!! \n"
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
        
        print "EPG_ExamSubtraction generate examination paper completed!! \n"


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
        
        print "EPG_ExamAddSub XLS file name is %s!!" % self.mFileName        
        
        self.mGenXls = EPG_GenXls(self.mFileName)
        
    def genExamPaper(self):
        for tQuestionIdx in range(self.mExamQuestionNumber) :
            print "EPG_ExamAddSub generate No.%d question!! \n" % tQuestionIdx
            
            tCalResult = 0
            tCalType   = 0
            tCalParam  = 0
            tQuestion  = ""
            tSuccess   = False
            
            while not tSuccess:
                for tParamIdx in range(self.mExamCalParamNumber):
                    print "EPG_ExamAddSub generate No.%d parameter!! \n" % tParamIdx
                    
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
        
        print "EPG_ExamAddSub generate examination paper completed!! \n"


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
4 - Multiplication [N/A]
5 - Division [N/A]
6 - Multiplication-Division mixed [N/A]
7 - Addition-Subtraction-Multiplication-Division mixed [N/A]
          '''
    tExamCalType = int(raw_input("Please input calculation type index (1, 2, 3, ...): "))
    
    print "\n"
    tExamCalParamNumber = int(raw_input("Please input the number of calculation parameters (2, 3, 4, ...): "))
    
    print "\n"
    tExamCalRange = int(raw_input("Please input the range of calculation (within 10, within 20, within 30, ...): "))
    
    print "\n"
    tExamQuestionNumber = int(raw_input("Please input the number of calculation questions (50, 100, ...): "))
    
    print "\n"
    tExamPaperNumber = int(raw_input("Please input the number of examination papers (20, 40, ...): "))
    
    for tExamPapgerIdx in range(tExamPaperNumber):
        if tExamCalType == 1:
            tAddition = EPG_ExamAddition(tExamCalParamNumber, tExamCalRange, tExamQuestionNumber)
            tAddition.genExamPaper()
        elif tExamCalType == 2:
            tSubtraction = EPG_ExamSubtraction(tExamCalParamNumber, tExamCalRange, tExamQuestionNumber)
            tSubtraction.genExamPaper()
        elif tExamCalType == 3:
            tAddSub = EPG_ExamAddSub(tExamCalParamNumber, tExamCalRange, tExamQuestionNumber)
            tAddSub.genExamPaper()
        else:
            print "\nThe function is in development, please wait...!!"
    
    