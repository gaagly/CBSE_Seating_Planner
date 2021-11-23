Option Explicit

Sub SEATING()
Dim StuRoll() As Long
Dim TotalClasses As Integer, TotalRooms As Integer
Dim SeatLimit As Integer, LimitCount As Integer
Dim SchoolName As String, ExamName As String, SubjectNameCode As String, HeadingTitle As String
Dim ExamDate As String
Dim i As Integer, BlockHeight As Integer
Dim j As Integer

Dim TotalStudents As Integer
Dim FirstRoll As Long, LastRoll As Long, CentreNumber As Long
Dim StudentRoll As Long, StuIndex As Integer
Dim TotalBoardStudents As Long
Dim Label As Integer
Dim k As Integer, l As Integer
''' Rows and Column value provided by the teacher'''
Dim row As Integer, col As Integer
row = Worksheets("Sheet1").Range("D3").Value
col = Worksheets("Sheet1").Range("D4").Value

TotalStudents = Worksheets("Sheet1").Range("D8").Value
ReDim StuRoll(TotalStudents) As Long
TotalClasses = Worksheets("Sheet1").Range("D7").Value


TotalRooms = Worksheets("Sheet1").Range("A2").Value

'''Combining All Students Roll no in One Array '''
''' THIS NEEDS TO BE CHANGED '''
StuIndex = 0
For i = 1 To TotalClasses
       FirstRoll = Worksheets("Sheet1").Cells(3, 7).Offset(0, i).Value
       ''' LastRoll = Worksheets("Sheet1").Cells(3, 7).Offset(1, i).Value '''
       
       ''' count total number of students in one list '''
       TotalBoardStudents = Worksheets("Sheet1").Cells(2, 7).Offset(0, i).Value
       
       '''For StudentRoll = FirstRoll To LastRoll'''
            '''StuRoll(StuIndex) = StudentRoll'''
            '''StuIndex = StuIndex + 1'''
       '''Next StudentRoll'''
       ''' CBSE The code given below copies the pasted roll number '''
       For StudentRoll = 0 To TotalBoardStudents
            StuRoll(StuIndex) = Worksheets("Sheet1").Cells(3, 7).Offset(StudentRoll, i).Value
            StuIndex = StuIndex + 1
       Next StudentRoll
   Next i

Debug.Print "THE TOTAL STUDENTS ARE: " & StuIndex
HeadingTitle = Worksheets("Sheet1").Range("D14").Value
SchoolName = Worksheets("Sheet1").Range("D12").Value
ExamName = Worksheets("Sheet1").Range("D13").Value
SubjectNameCode = Worksheets("Sheet1").Range("D17").Value
ExamDate = Worksheets("Sheet1").Range("D16").Value
CentreNumber = Worksheets("Sheet1").Range("D15").Value
''' LABEL is width caused by '''
'''                    HEADING                     '''
''' 1. School Name '''          '''4. CENTRE NUMBER'''
''' 2. Exam Name '''            '''5. SUBJECT      '''
''' 3. Room NAME '''            '''6. DATE         '''
'''                     ROOM NAME                  '''
''' ROLL NUMBER, QP CODE, ROLL NUMBER, QP CODE     '''
Label = 9
''' SCHOOL NAME till Students seats ''''
''' EXAM NAME '''
''' STUDENTS NAME '''
''' TOTAL STUDENTS , PRESENT ,ABSENT '''
''' NOW STATIC TEXT WILL FOLLOW'''
''' NAME AND SIGNATURE OF ASSISTANT SUPRITENDENT '''
''' SIGNATURE OF CENTRE SUPRINTENDENT '''
''' BLOCKHEIGHT is parchi ki height'''
BlockHeight = Label + 1 + Worksheets("Sheet1").Range("D3").Value + 8 ''' THE PLUS 8 IN THE END IS FOR STATIC TEXT LIKE NAME AND SIGNARURE OF ASST. SUPDT. CENTRE SUPDT.'''

SeatLimit = Worksheets("Sheet1").Range("F5").Value
Debug.Print "There is a limit of " & SeatLimit & " students."
StuIndex = 0
For i = 1 To TotalRooms
    j = 1 + (i - 1) * BlockHeight
    
    '''        HEADINGS                 '''
    Cells(j, 1).Offset(0, 0).Value = HeadingTitle
    
    '''         Centre Name             '''
    Cells(j, 1).Offset(1, 0).Value = "CENTRE NAME"
    Cells(j, 1).Offset(1, 1).Value = SchoolName
    
    '''        Centre Number            '''
    Cells(j, 1).Offset(2, 0).Value = "CENTRE NO."
    Cells(j, 1).Offset(2, 1).Value = CentreNumber
    
    ''' Name of the Examination         '''
    Cells(j, 1).Offset(3, 0).Value = "EXAM NAME"
    Cells(j, 1).Offset(3, 1).Value = ExamName
    
    '''              Subject Code & Name       '''
    Cells(j, 1).Offset(4, 0).Value = "SUBJECT"
    Cells(j, 1).Offset(4, 1).Value = SubjectNameCode
    
    '''             Exam Date           '''
    Cells(j, 1).Offset(5, 0).Value = "DATE"
    Cells(j, 1).Offset(5, 1).Value = Format(ExamDate, "Long Date")
    
    '''     Room Name                   '''
    Cells(j, 1).Offset(7, 0).Value = Worksheets("Sheet1").Cells(i + 2, 1).Value
    
    '''             ROLL NO AND QP HEADINGS         '''
    For l = 0 To (row / 2)
        Cells(j, 1).Offset(8, 2 * l).Value = "Roll No"
        Cells(j, 1).Offset(8, 2 * l + 1).Value = "QP. Code"
    Next l
    
    ''' NAME & Signature of Asstt. Superintendent''''
        Cells(j, 1).Offset(Label + row + 1, 0).Value = "Name & Signature of Asstt. Superintendent"
        Cells(j, 1).Offset(Label + row + 1, 0).WrapText = True
            
        ''' Total NO. registered '''
        Cells(j, 1).Offset(Label + row + 1, 2 * (col - 1)).Value = "Total No. Students"
        Cells(j, 1).Offset(Label + row + 1, 2 * (col)).Value = row * col
        '''   Present            '''
        Cells(j, 1).Offset(Label + row + 2, 2 * (col - 1)).Value = "Present"
        '''     Absent           '''
        Cells(j, 1).Offset(Label + row + 3, 2 * (col - 1)).Value = "Absent"
        
    '''       Signature of Centre Suprintendent            '''
        Cells(j, 1).Offset(Label + row + 5, 2 * (col - 2)).Value = "Signature of Centre Suprintendent"
        Cells(j, 1).Offset(Label + row + 5, 2 * (col - 2)).WrapText = True
    
    '''''''''''''''''''roll number''''''''''''''''''''''''''''''''''''
    LimitCount = 0
        For l = 1 To col
            For k = 1 To row
                Cells(j, 1).Offset(Label + k - 1, 2 * (l - 1)).Value = StuRoll(StuIndex)
                StuIndex = StuIndex + 1
                If StuIndex = TotalStudents Then
                    Exit For
                End If
                If SeatLimit > 0 Then
                    LimitCount = LimitCount + 1
                    Debug.Print LimitCount
                    If LimitCount >= SeatLimit Then
                        Exit For
                    End If
                End If
                
            Next k
            
           
            If StuIndex = TotalStudents Then
                    Exit For
            End If
            If SeatLimit > 0 Then
                
                If LimitCount >= SeatLimit Then
                    Exit For
                End If
            End If
          
            
        Next l
    If StuIndex = TotalStudents Then
                    Exit For
    End If


Next i
    
Debug.Print "-------------------------------------------"
End Sub





