VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsLoanOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Constants for error codes
Private Const ERR_NEGATIVE_PRINCIPAL As Long = vbObjectError + 1
Private Const ERR_INVALID_TENURE As Long = vbObjectError + 2
Private Const ERR_INVALID_BALLOON As Long = vbObjectError + 3
Private Const ERR_NEGATIVE_INSTALLMENT As Long = vbObjectError + 4
Private Const ERR_INVALID_INTEREST As Long = vbObjectError + 5
Private Const ERR_INVALID_DAYS_OPTION As Long = vbObjectError + 6
Private Const ERR_INVALID_START_DATE As Long = vbObjectError + 7
Private Const ERR_INVALID_INST_START_DATE As Long = vbObjectError + 8
Private Const ERR_INVALID_YEAR_END_MONTH As Long = vbObjectError + 9

' Private variables
Private mPrincipal As Double
Private mTenure As Double
Private mBalloon As Double
Private mInstallment As Double
Private mInterest As Double
Private mDaysOption As Integer
Private mInterestRounding As Boolean
Private mStart As Date
Private mInst_Start As Date
Private mYearEndMonth As Integer

' Class Initialize
Private Sub Class_Initialize()
    mPrincipal = 0
    mTenure = 0
    mBalloon = 0
    mInstallment = 0
    mInterest = 0
    mDaysOption = 0
    mInterestRounding = False
    mStart = Now()
    mInst_Start = DateAdd("m", 1, Now())
    mYearEndMonth = 2
End Sub

' Class Terminate
Private Sub Class_Terminate()
    ' Optional cleanup logic
End Sub

' Helper to validate positive numbers
Private Sub ValidatePositiveNumber(Value As Double, PropertyName As String)
    If Value < 0 Then
        Err.Raise vbObjectError + 10, PropertyName, PropertyName & " must be a positive number."
    End If
End Sub

' Property: Principal
Public Property Get Principal() As Double
    Principal = mPrincipal
End Property

Public Property Let Principal(Value As Double)
    ValidatePositiveNumber Value, "Principal"
    mPrincipal = Value
End Property

' Property: Tenure
Public Property Get Tenure() As Double
    Tenure = mTenure
End Property

Public Property Let Tenure(Value As Double)
    If Value <= 0 Then
        Err.Raise ERR_INVALID_TENURE, "Tenure Property", "Tenure must be greater than 0."
    End If
    mTenure = Value
End Property

' Property: Balloon
Public Property Get Balloon() As Double
    Balloon = mBalloon
End Property

Public Property Let Balloon(Value As Double)
    If Value < 0 Or Value > mPrincipal Then
        Err.Raise ERR_INVALID_BALLOON, "Balloon Property", "Balloon amount cannot be negative or exceed the principal."
    End If
    mBalloon = Value
End Property

' Property: Installment
Public Property Get Installment() As Double
    Installment = mInstallment
End Property

Public Property Let Installment(Value As Double)
    ValidatePositiveNumber Value, "Installment"
    mInstallment = Value
End Property

' Property: Interest
Public Property Get Interest() As Double
    Interest = mInterest
End Property

Public Property Let Interest(Value As Double)
    If Value < 0 Or Value > 100 Then
        Err.Raise ERR_INVALID_INTEREST, "Interest Property", "Interest rate must be between 0 and 100."
    End If
    mInterest = Value
End Property

' Property: DaysOption
Public Property Get DaysOption() As Integer
    DaysOption = mDaysOption
End Property

Public Property Let DaysOption(Value As Integer)
    If Value < 0 Or Value > 2 Then
        Err.Raise ERR_INVALID_DAYS_OPTION, "DaysOption Property", "DaysOption must be 0, 1, or 2."
    End If
    mDaysOption = Value
End Property

' Property: InterestRounding
Public Property Get InterestRounding() As Boolean
    InterestRounding = mInterestRounding
End Property

Public Property Let InterestRounding(Value As Boolean)
    mInterestRounding = Value
End Property

' Property: Start Date
Public Property Get StartDate() As Date
    StartDate = mStart
End Property

Public Property Let StartDate(Value As Date)
    If Not IsDate(Value) Then
        Err.Raise ERR_INVALID_START_DATE, "StartDate Property", "Invalid Start date."
    End If
    mStart = Value
End Property

' Property: Installment Start Date
Public Property Get InstallmentStartDate() As Date
    InstallmentStartDate = mInst_Start
End Property

Public Property Let InstallmentStartDate(Value As Date)
    If Value < mStart Then
        Err.Raise ERR_INVALID_INST_START_DATE, "InstallmentStartDate Property", "Installment start date cannot be before the loan start date."
    End If
    mInst_Start = Value
End Property

' Property: YearEndMonth
Public Property Get YearEndMonth() As Integer
    YearEndMonth = mYearEndMonth
End Property

Public Property Let YearEndMonth(Value As Integer)
    If Value < 1 Or Value > 12 Then
        Err.Raise ERR_INVALID_YEAR_END_MONTH, "YearEndMonth Property", "Year end month must be between 1 and 12."
    End If
    mYearEndMonth = Value
End Property

