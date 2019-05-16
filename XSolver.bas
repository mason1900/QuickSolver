Attribute VB_Name = "XSolver1"
Option Explicit

Private Function CheckConstraints(darr() As Double, dExclusiveBit() As Double, dInvestment() As Double, dInvestmentBeta() As Double, _
    dMinBeta As Double, dMaxBeta As Double, dMinCap As Double, dMaxCap As Double) As Boolean

'This function return True if constraints are not satisfied.

Dim itemp3 As Long
Dim ExclusiveBit As Long
Dim dSumBits As Double
Dim dBetaNumerater As Double
Dim dBetaDenominator As Double
Dim dOverallBeta As Double

If vbErrHandler Then On Error GoTo Handler

dSumBits = 0
dBetaNumerater = 0
dBetaDenominator = 0
dOverallBeta = 0

For itemp3 = 0 To UBound(darr) - 1
    'Note: the iteration outside of this function starts from all zero to all filled so no problem with "UBound(darr) - 1"
    ExclusiveBit = dExclusiveBit(itemp3)
    dSumBits = dSumBits + darr(itemp3) * ExclusiveBit
    dBetaNumerater = dBetaNumerater + darr(itemp3) * dInvestmentBeta(itemp3)
    dBetaDenominator = dBetaDenominator + darr(itemp3) * dInvestment(itemp3)
    
Next

'All projects rejected
If dBetaDenominator = 0 Then
    CheckConstraints = True
    Exit Function
End If

dOverallBeta = dBetaNumerater / dBetaDenominator

'Debug.Assert dSumBits > 1

If dSumBits > 1 Then
    CheckConstraints = True
ElseIf dOverallBeta < dMinBeta Or dOverallBeta > dMaxBeta Then
    CheckConstraints = True
ElseIf dBetaDenominator < dMinCap Or dBetaDenominator > dMaxCap Then
    CheckConstraints = True
End If

'Debug.Print "Got here"

Exit Function
Handler:
    Debug.Print "CheckConstraints Failed!"
    Debug.Print "itemp3 = " & itemp3 & "ExclusiveBit = " & ExclusiveBit & "dSumBits = " & dSumBits & "dBetaNumerater = " & dBetaNumerater
    Debug.Print "dBetaDenominator = " & dBetaDenominator & "dOverallBeta = " & dOverallBeta
    Exit Function
    
End Function


Sub XSolver()

Dim darr() As Double
Dim dArrSum() As Double
Dim dArrOutput() As Double
Dim dBits() As Double
Dim dExclusiveBit() As Double
Dim dInvestment() As Double
Dim dInvestmentBeta() As Double

Dim index As Long
Dim itemp, itemp2 As Long
Dim dSum As Double
Dim bFlag As Boolean
Dim bisOptimalExist As Boolean
Dim dMinusInfinity As Double
    
Dim n As Long
Dim dMaxSum  As Double
Dim iMaxSumIndex As Long

Dim dMinCap  As Double
Dim dMaxCap As Double
Dim dMinBeta As Double
Dim dMaxBeta As Double

Dim dBetaNumerater As Double
Dim dBetaDenominator As Double

Dim t As Double
Dim t2 As Double

If vbErrHandler Then On Error GoTo Handler

t = Timer

n = ThisWorkbook.Sheets("Summary").Range("bAcceptDecision").Count
bisOptimalExist = False
dMinusInfinity = -4.2E+20

'The algorithm requires a lot of memory, (2 ^ n - 1). It cause the Excel to crash when n is too large on some computers.
'In addition, execution time increase significantly when n>=24 (more than 1 minute on my computer)
If n > 21 Then
    MsgBox "Will run Excel solver Add-in instead.", vbInformation + vbOKOnly, "Info"
    Call RunSolver
    Exit Sub
End If

ReDim darr(n), dArrOutput(n) As Double
ReDim dArrSum(2 ^ n - 1) As Double
ReDim dBits(n) As Double
ReDim dExclusiveBit(n) As Double
ReDim dInvestment(n) As Double
ReDim dInvestmentBeta(n) As Double
'Please note that All of the arrays above are 1 size bigger than needed

'dBits = Array(5, -7, -9) won't work for VBA
'dBits(0) = 5
'dBits(1) = -7
'dBits(2) = -9


'Read data and constraints
With ThisWorkbook.Sheets("Summary")
    For itemp = 0 To n - 1
        'Debug.Print ThisWorkbook.Sheets("Summary").Range("bAcceptDecision").Cells(itemp + 1, 1)
        dBits(itemp) = .Range("NPVvalues").Cells(itemp + 1, 1)
        dExclusiveBit(itemp) = .Range("bExclusiveBits").Cells(itemp + 1, 1)
        dInvestment(itemp) = .Range("Investment").Cells(itemp + 1, 1)
        dInvestmentBeta(itemp) = .Range("InvestmentBeta").Cells(itemp + 1, 1)
    Next
    dMinCap = .Range("MinCap").Value2
    dMaxCap = .Range("MaxCap").Value2
    dMinBeta = .Range("MinBeta").Value2
    dMaxBeta = .Range("MaxBeta").Value2
    
End With

If dMinCap > dMaxCap Then
    MsgBox "Please check Capital Constraints input. Minimun capital must be smaller than Maximum capital.", _
        vbCritical, "Error"
    Exit Sub
ElseIf dMinBeta > dMaxBeta Then
    MsgBox "Please check Beta Constraints input. Minimun beta must be smaller than Maximum beta.", _
        vbCritical, "Error"
    Exit Sub
End If

t2 = Timer

For itemp = 0 To (2 ^ n - 1)
    
    'Starting from 0 is a trick. Convert index of dArrSum to binary number will directly become the
    'corresponding decision bits.
    index = 0
    itemp2 = itemp
    While itemp2 <> 0
        darr(index) = itemp2 Mod 2
        'Note: Use \ instead of /
        itemp2 = itemp2 \ 2
        index = index + 1
    Wend
    
    
    dSum = 0
    'Hard code overall beta numerator here for efficiency
    'Not implemented yet
    'dBetaNumerater = 0
    'dBetaDenominator = 0
    
    For itemp2 = 0 To index - 1
        'MsgBox dArr(itemp)
        dSum = dSum + darr(itemp2) * dBits(itemp2)
        'Hard code overall beta numerator and denominator here for efficiency
        'Not implemented yet
        'dBetaNumerater = dBetaNumerater + dInvestmentBeta(itemp2) * darr(itemp2)
        'dBetaDenominator = dBetaDenominator + dInvestment(itemp2) * darr(itemp2)
      
    Next
    
    bFlag = False
    
    'Hard Code capital constraints here for efficiency
    'Not implemented yet
    'Reject NPV <= 0
    If dSum <= 0 Then
    'Note: itemp = 0 already rejected here.
        bFlag = True
    Else
        bFlag = CheckConstraints(darr, dExclusiveBit, dInvestment, dInvestmentBeta, dMinBeta, dMaxBeta, dMinCap, dMaxCap)
    End If
    
    If bFlag Then dSum = dMinusInfinity
    dArrSum(itemp) = dSum

    ' If no feasible projects, this var will always be FALSE
    bisOptimalExist = bisOptimalExist Or (Not bFlag)
Next

If Not bisOptimalExist Then
    t2 = Timer - t2
    Debug.Print t2
    MsgBox "No possible optimal set of projects. Please check your input.", vbExclamation, "Error"
    Exit Sub
End If


dMaxSum = -4.2E+20
iMaxSumIndex = -1
For itemp = 0 To (2 ^ n - 1)
    If dArrSum(itemp) > dMaxSum Then
        dMaxSum = dArrSum(itemp)
        iMaxSumIndex = itemp
    End If
Next
    
index = 0
itemp2 = iMaxSumIndex
While itemp2 <> 0
    dArrOutput(index) = itemp2 Mod 2
    'Note: Use \ instead of /
    itemp2 = itemp2 \ 2
    index = index + 1
Wend

t2 = Timer - t2
Debug.Print t2
    
' Output!
Application.ScreenUpdating = False
For itemp2 = 0 To n - 1
    'Debug.Print dArr(itemp)
    'Debug.Print dArrOutput(itemp2)
    ThisWorkbook.Sheets("Summary").Range("bAcceptDecision").Cells(itemp2 + 1, 1).Value2 = dArrOutput(itemp2)
Next
Application.ScreenUpdating = True

    
    t = Timer - t
    Debug.Print t

    MsgBox "Done!"
    
Exit Sub
Handler:
    MsgBox "Solver failed to run. Will run Excel solver Add-in instead.", vbCritical + vbOKOnly, "Error"
    Call RunSolver
    Exit Sub
    
End Sub

