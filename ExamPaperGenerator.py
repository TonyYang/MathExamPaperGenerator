# _*_ coding: utf-8 _*_


from win32com.client import Dispatch


import random, time, os


###############################################################################
## Global variable definition                                                ##
##                                                                           ##
###############################################################################
CAL_OP_ADD = 1
CAL_OP_SUB = 2
CAL_OP_MUL = 3
CAL_OP_DIV = 4

CAL_OP_ADD_STR = " + "
CAL_OP_SUB_STR = " - "
CAL_OP_MUL_STR = " x "
CAL_OP_DIV_STR = " / "


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

        
    def writeQuestion2(self, pIdx, pContent):
        tRow = int(pIdx) / 3 + 2
        tCol = int(pIdx) % 3 + 1
        self.mXlsSheet.Cells(tRow, tCol).Value = pContent


    def writeComment(self, pExamCalParamNumber, pExamQuestionNumber):
        tTime = pExamQuestionNumber * (pExamCalParamNumber / 3)
        tComments = "在" + str(tTime) + "分钟内完成!!"
        self.mXlsSheet.Cells(1, 1).Value = tComments.decode("utf8").encode("gbk")


    def writeComment2(self, pExamCalParamNumber, pExamQuestionNumber):
        tTime = pExamQuestionNumber * (pExamCalParamNumber / 2)
        tComments = "在" + str(tTime) + "分钟内完成!!"
        self.mXlsSheet.Cells(1, 1).Value = tComments.decode("utf8").encode("gbk")


    def formatContent(self, pExamQuestionNumber):
        # Format column width
        self.mXlsSheet.Range(self.mXlsSheet.Cells(1, 1), self.mXlsSheet.Cells(1, 5)).ColumnWidth = 25
        
        # Format font
        if (pExamQuestionNumber % 5) != 0:
            rowPlus = 2
        else:
            rowPlus = 1
        #
        self.mXlsSheet.Range(self.mXlsSheet.Cells(1, 1), self.mXlsSheet.Cells(pExamQuestionNumber / 5 + rowPlus, 5)).Font.Bold = True
        self.mXlsSheet.Range(self.mXlsSheet.Cells(1, 1), self.mXlsSheet.Cells(pExamQuestionNumber / 5 + rowPlus, 5)).Font.Name = "Calibri"
        self.mXlsSheet.Range(self.mXlsSheet.Cells(1, 1), self.mXlsSheet.Cells(pExamQuestionNumber / 5 + rowPlus, 5)).Font.Size = 16
        
        # Format border
        for iRow in range(1, pExamQuestionNumber / 5 + 2):
            for iCol in range(1, 6):
                for iBorder in range(1, 5):
                    self.mXlsSheet.Cells(iRow, iCol).Borders(iBorder).LineStyle = 1


    def formatContent2(self, pExamQuestionNumber):
        # Define column width
        self.mXlsSheet.Range(self.mXlsSheet.Cells(1, 1), self.mXlsSheet.Cells(1, 3)).ColumnWidth = 45
        # Define font
        if (pExamQuestionNumber % 3) != 0:
            rowPlus = 2
        else:
            rowPlus = 1
        #    
        self.mXlsSheet.Range(self.mXlsSheet.Cells(1, 1), self.mXlsSheet.Cells(pExamQuestionNumber / 3 + rowPlus, 3)).Font.Bold = True
        self.mXlsSheet.Range(self.mXlsSheet.Cells(1, 1), self.mXlsSheet.Cells(pExamQuestionNumber / 3 + rowPlus, 3)).Font.Name = "Calibri"
        self.mXlsSheet.Range(self.mXlsSheet.Cells(1, 1), self.mXlsSheet.Cells(pExamQuestionNumber / 3 + rowPlus, 3)).Font.Size = 16
        # Define row's height
        self.mXlsSheet.Range(self.mXlsSheet.Cells(2, 1), self.mXlsSheet.Cells(pExamQuestionNumber / 3 + rowPlus, 3)).RowHeight = 120
        
        # Format border
        for iRow in range(1, pExamQuestionNumber / 3 + rowPlus + 1):
            for iCol in range(1, 4):
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
## 'EPG_ExamAddSubMixed' Class                                               ##
## - Generate addition calculation questions randomly.                       ##
## - Generate subtraction calculation questions randomly.                    ##
###############################################################################
class EPG_ExamAddSubMixed:
    def __init__(self, pExamCalParamNumber, pExamCalMinRange, pExamCalMaxRange, pExamQuestionNumber):
        self.mExamCalParamNumber = pExamCalParamNumber
        self.mExamCalMinRange    = pExamCalMinRange
        self.mExamCalMaxRange    = pExamCalMaxRange
        self.mExamQuestionNumber = pExamQuestionNumber
        
        self.mFileName = "AddSubMixed_" + str(self.mExamCalParamNumber) + "_" + str(self.mExamCalMinRange) + "_" + str(self.mExamCalMaxRange) + "_" + str(self.mExamQuestionNumber) + "_" + time.strftime("%Y%m%d") + "_" + time.strftime("%H%M%S")
        self.mFileName = os.getcwd() + "\\" + self.mFileName
        
        print "[EPG_ExamAddSubMixed] Excel file name is %s!!" % self.mFileName        
        
        self.mGenXls = EPG_GenXls(self.mFileName)
        
    def genExpression(self):
        ## Variables definition and initialization
        tCalResult = 0
        tCalType   = 0
        tCalParam  = 0
        tQuestion  = ""
        tSuccess   = False
        
        while not tSuccess:
            for tParamIdx in range(self.mExamCalParamNumber):
                print "[EPG_ExamAddSubMixed] Generate No.%d parameter!! \n" % tParamIdx
                
                if tParamIdx == 0:
                    tCalParam  = random.randint(self.mExamCalMinRange, self.mExamCalMaxRange - 1)
                    tQuestion  = tQuestion + str(tCalParam)
                    tCalResult = tCalParam
                    tSuccess   = True
                else:
                    tCalType = random.randint(1, 2)
                    
                    if tCalType == 1 and (self.mExamCalMaxRange - tCalResult ) > self.mExamCalMinRange:
                        tCalParam  = random.randint(self.mExamCalMinRange, self.mExamCalMaxRange - tCalResult - 1)
                        tQuestion  = tQuestion + " + " + str(tCalParam)
                        tCalResult = tCalResult + tCalParam
                        tSuccess   = True
                    elif tCalType == 2 and tCalResult > self.mExamCalMinRange:
                        tCalParam  = random.randint(self.mExamCalMinRange, tCalResult - 1)
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
           
        return tQuestion
    
    def genExamPaper(self):
        for tQuestionIdx in range(self.mExamQuestionNumber) :
            print "[EPG_ExamAddSubMixed] Generate No.%d question!! \n" % tQuestionIdx
            
            mExpression = self.genExpression()
            
            self.mGenXls.writeQuestion2(tQuestionIdx, mExpression)

        self.mGenXls.formatContent2(self.mExamQuestionNumber)
        self.mGenXls.writeComment2(self.mExamCalParamNumber, self.mExamQuestionNumber)
        self.mGenXls.saveFile()
        
        print "[EPG_ExamAddSubMixed] Generate examination paper completed!! \n"


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
## 'EPG_ExamAddSubMulRandom' Class                                           ##
## - Generate Addition calculation questions randomly.                       ##
## - Generate Subtraction calculation questions randomly.                    ##
## - Generate multiplication calculation questions randomly.                 ##
###############################################################################
class EPG_ExamAddSubMulRandom:
    def __init__(self, pOperandNum, pAddSubCalRange, pMulCalRange, pExpressionNum):
        self.mOperandNum     = pOperandNum
        self.mAddSubCalRange = pAddSubCalRange
        self.mMulCalRange    = pMulCalRange
        self.mExpressionNum  = pExpressionNum
        
        self.mFileName = "AddSubMulRandom_" + str(self.mOperandNum) + "_" + str(self.mAddSubCalRange) + "_" + str(self.mMulCalRange) + "_" + str(self.mExpressionNum) + "_" + time.strftime("%Y%m%d") + "_" + time.strftime("%H%M%S")
        self.mFileName = os.getcwd() + "\\" + self.mFileName
        
        print "[EPG_ExamAddSubMulRandom] Excel file name is %s!!" % self.mFileName
        
        self.mGenXls = EPG_GenXls(self.mFileName)

    def genExpression(self):
        mOperand1 = 0
        mOperand2 = 0
        mOperator = 0
        
        mOperatorStr = ""
        mExpression  = ""
        
        mOperator = random.randint(CAL_OP_ADD, CAL_OP_MUL)
        
        ## Addition calculation.
        if mOperator == 1:
            mOperatorStr = " + "
            mOperand1 = random.randint(1, self.mAddSubCalRange - 1)
            mOperand2 = random.randint(1, self.mAddSubCalRange - mOperand1)

        ## Subtraction calculation    
        elif mOperator == 2:
            mOperatorStr = " - "
            mOperand1 = random.randint(1, self.mAddSubCalRange)
            mOperand2 = random.randint(1, mOperand1)

        ## Multiplication calculation
        else:
            mOperatorStr = " x "
            mOperand1 = random.randint(1, self.mMulCalRange)
            mOperand2 = random.randint(1, self.mMulCalRange)
            
        mExpression = str(mOperand1) + mOperatorStr + str(mOperand2) + " = "
        return mExpression
    
    def genExamPaper(self):
        for mExpressionIdx in range(self.mExpressionNum):
            print "[EPG_ExamAddSubMulRandom] Generate No.%d question!! \n" % mExpressionIdx
            
            mExpression = self.genExpression()
            self.mGenXls.writeQuestion(mExpressionIdx, mExpression)
            
        self.mGenXls.writeComment(self.mOperandNum, self.mExpressionNum)
        self.mGenXls.formatContent(self.mExpressionNum)
        self.mGenXls.saveFile()
        
        print "[EPG_ExamAddSubMulRandom] Generating examination paper completed!! \n"


###############################################################################
## 'EPG_ExamAddSubMulMixed' Class                                            ##
## - Generate Addition calculation questions randomly.                       ##
## - Generate Subtraction calculation questions randomly.                    ##
## - Generate multiplication calculation questions randomly.                 ##
###############################################################################
class EPG_ExamAddSubMulMixed:
    def __init__(self, pExamParamNum, pExamAddSubCalRange, pExamMulCalRange, pExamQuestionNum):
        self.mExamParamNum = pExamParamNum
        self.mExamAddSubCalRange       = pExamAddSubCalRange
        self.mExamMulCalRange          = pExamMulCalRange
        self.mExamQuestionNum          = pExamQuestionNum
        
        self.mFileName = "AddSubMulMixed_" + str(self.mExamParamNum) + "_" + str(self.mExamAddSubCalRange) + "_" + str(self.mExamMulCalRange) + "_" + str(self.mExamQuestionNum) + "_" + time.strftime("%Y%m%d") + "_" + time.strftime("%H%M%S")
        self.mFileName = os.getcwd() + "\\" + self.mFileName
        
        print "[EPG_ExamAddSubMulMixed] Excel file name is %s!!" % self.mFileName
        
        self.mGenXls = EPG_GenXls(self.mFileName)
        
    def genExpression(self):
        ## Variables definition and initialization
        mMulDivFirst = 0
        mParam1      = 0
        mParam2      = 0
        mParam3      = 0
        mOp1         = 0
        mOp2         = 0
        mOpStr1      = ""
        mOpStr2      = ""
        mExpression  = ""
        
        mMulDivFirst = random.randint(0, 1)
        
        ## Add/Sub first, Mul/Div last
        if mMulDivFirst == 0:
            mParam2 = random.randint(1, self.mExamMulCalRange)
            mParam3 = random.randint(1, self.mExamMulCalRange)
            mOp2    = CAL_OP_MUL
            mOpStr2 = CAL_OP_MUL_STR
            
            mOp1 = random.randint(CAL_OP_ADD, CAL_OP_SUB)
            
            if mOp1 == CAL_OP_ADD:
                mOpStr1 = CAL_OP_ADD_STR
                mParam1 = random.randint(1, self.mExamAddSubCalRange - mParam2 * mParam3)
            else:
                mOpStr1 = CAL_OP_SUB_STR
                mParam1 = random.randint(mParam2 * mParam3 + 1, self.mExamAddSubCalRange)
        
        ## Mul/Div first, Add/Sub last    
        else:
            mParam1 = random.randint(1, self.mExamMulCalRange)
            mParam2 = random.randint(1, self.mExamMulCalRange)
            mOp1 = CAL_OP_MUL
            mOpStr1 = CAL_OP_MUL_STR
            
            mOp2 = random.randint(CAL_OP_ADD, CAL_OP_SUB)
            
            if mOp2 == CAL_OP_ADD:
                mOpStr2 = CAL_OP_ADD_STR
                mParam3 = random.randint(1, self.mExamAddSubCalRange - mParam1 * mParam2)
            else:
                mOpStr2 = CAL_OP_SUB_STR
                mParam3 = random.randint(1, mParam1 * mParam2)
            
        mExpression = str(mParam1) + mOpStr1 + str(mParam2) + mOpStr2 + str(mParam3) + " = "
        
        return mExpression
    
    def genExamPaper(self):
        for mExpressionIdx in range(self.mExamQuestionNum):
            print "[EPG_ExamAddSubMulMixed] Generate No.%d question!! \n" % mExpressionIdx
            
            mExpression = self.genExpression()
            self.mGenXls.writeQuestion(mExpressionIdx, mExpression)
            
        self.mGenXls.writeComment(self.mExamParamNum, self.mExamQuestionNum)
        self.mGenXls.formatContent(self.mExamQuestionNum)
        self.mGenXls.saveFile()
        
        print "[EPG_ExamAddSubMulMixed] Generating examination paper completed!! \n"


###############################################################################
## 'EPG_ExamAddSubMulRandomMixed' Class                                      ##
## - Generate Addition calculation questions randomly.                       ##
## - Generate Subtraction calculation questions randomly.                    ##
## - Generate multiplication calculation questions randomly.                 ##
###############################################################################
class EPG_ExamAddSubMulRandomMixed:
    def __init__(self, pExpressionNum):
        self.mExpressionNum = pExpressionNum
        
        self.mAddSubMulRandom = EPG_ExamAddSubMulRandom(2, 100, 9, self.mExpressionNum)
        self.mAddSubMulMixed  = EPG_ExamAddSubMulMixed(3, 100, 9, self.mExpressionNum)
        
        self.mFileName = "AddSubMulRandomMixed_" + "2&3" + "_" + str(100) + "_" + str(9) + "_" + str(self.mExpressionNum) + "_" + time.strftime("%Y%m%d") + "_" + time.strftime("%H%M%S")
        self.mFileName = os.getcwd() + "\\" + self.mFileName
        
        print "[EPG_ExamAddSubMulRandom] Excel file name is %s!!" % self.mFileName
        
        self.mGenXls = EPG_GenXls(self.mFileName)
        
    def genExpression(self):
        mExpression    = ""
        mRandomOrMixed = 0
        
        mRandomOrMixed = random.randint(1, 2)
        
        if mRandomOrMixed == 1:
            mExpression = self.mAddSubMulRandom.genExpression()
        else:
            mExpression = self.mAddSubMulMixed.genExpression()
            
        return mExpression
    
    def genExamPaper(self):
        for mExpressionIdx in range(self.mExpressionNum):
            print "[EPG_ExamAddSubMulRandomMixed] Generate No.%d question!! \n" % (mExpressionIdx + 1)
            
            mExpression = self.genExpression()
            self.mGenXls.writeQuestion(mExpressionIdx, mExpression)
            
        self.mGenXls.writeComment(3, self.mExpressionNum)
        self.mGenXls.formatContent(self.mExpressionNum)
        self.mGenXls.saveFile()
        
        print "[EPG_ExamAddSubMulRandomMixed] Generating examination paper completed!! \n"


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
5 - Addition-Subtraction-Multiplication random
6 - Addition-Subtraction-Multiplication mixed
7 - Addition-Subtraction-Multiplication random & mixed
8 - Addition-Subtraction mixed (3 digits)
============================================================
# - Division [N/A]
# - Multiplication-Division mixed [N/A]
# - Addition-Subtraction-Multiplication-Division mixed [N/A]
          '''
    
    ## Ask user to select calculation type.
    mCalType = int(raw_input("Please choose calculation type index (1, 2, 3, ...): "))
    
    ## Ask user to input how many questions will be generated in one paper.
    print "\n"
    mExpressionNum = int(raw_input("Please input the number of expression (20, 50, 100, ...): "))
        
    ## Ask user to input how many examination papers will be generated.
    print "\n"
    mExamPaperNum = int(raw_input("Please input the number of examination papers (20, 40, ...): "))
    
    ## Addition and Subtraction
    if mCalType == 1 or mCalType == 2:
        mOperandNum = 2
        mCalRange = 100
    ## Addition-Subtraction mixed
    elif mCalType == 3:
        mOperandNum = 3
        mCalMinRange = 1
        mCalMaxRange = 100
    ## Multiplication
    elif mCalType == 4:
        mOperandNum = 2
        mCalRange = 9
    ## Addition-Subtraction-Multiplication random
    elif mCalType == 5:
        mOperandNum = 2
        mAddSubCalRange = 100
        mMulCalRange = 9
    ## Addition-Subtraction-Multiplication mixed
    elif mCalType == 6:
        mOperandNum = 3
        mAddSubCalRange = 100
        mMulCalRange = 9
    ## Addition-Subtraction-Multiplication random & mixed
    elif mCalType == 7:
        mOperandNum = 3
        mAddSubCalRange = 100
        mMulCalRange = 9
    ## Addition-Subtraction mixed (3 digits)
    elif mCalType == 8:
        mOperandNum = 2
        mCalMinRange = 100
        mCalMaxRange = 1000
    ## Not supported
    else:
        print "\n"
        print "The calculation type is NOT supported!!"

    ## Generate examination papers.
    for mExamPapgerIdx in range(mExamPaperNum):
        if mCalType == 1:
            mAdd = EPG_ExamAdd(mOperandNum, mCalRange, mExpressionNum)
            mAdd.genExamPaper()
        
        elif mCalType == 2:
            mSub = EPG_ExamSub(mOperandNum, mCalRange, mExpressionNum)
            mSub.genExamPaper()
        
        elif mCalType == 3:
            mAddSubRandom = EPG_ExamAddSubMixed(mOperandNum, mCalMinRange, mCalMaxRange, mExpressionNum)
            mAddSubRandom.genExamPaper()
        
        elif mCalType == 4:
            mMul = EPG_ExamMul(mOperandNum, mCalRange, mExpressionNum)
            mMul.genExamPaper()
        
        elif mCalType == 5:
            mAddSubMulRandom = EPG_ExamAddSubMulRandom(mOperandNum, mAddSubCalRange, mMulCalRange, mExpressionNum)
            mAddSubMulRandom.genExamPaper()
            
        elif mCalType == 6:
            mAddSubMulMixed = EPG_ExamAddSubMulMixed(mOperandNum, mAddSubCalRange, mMulCalRange, mExpressionNum)
            mAddSubMulMixed.genExamPaper()
        
        elif mCalType == 7:
            mAddSubMulRandomMixed = EPG_ExamAddSubMulRandomMixed(mExpressionNum)
            mAddSubMulRandomMixed.genExamPaper()
            
        elif mCalType == 8:
            mAddSubMixed = EPG_ExamAddSubMixed(mOperandNum, mCalMinRange, mCalMaxRange, mExpressionNum)
            mAddSubMixed.genExamPaper()
        
        else:
            print "\n\n"
            print "The calculation type is NOT supported!!"
