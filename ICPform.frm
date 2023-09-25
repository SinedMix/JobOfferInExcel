VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ICPform 
   Caption         =   "Заполнение индивидуального компенсационного пакета (ИКП)"
   ClientHeight    =   5856
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   9825.001
   OleObjectBlob   =   "ICPform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ICPform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()
'Обработка кнопки "Отмена"
    ICPform.Hide
    TypeContract_value.Clear
    ProbPeriod_value.Clear
    TypeEmployment_value.Clear
    WorkSchedule_value.Clear
End Sub

Private Sub Clear_Click()
' Обработка кнопки "Очистить поля"
    FullName_value.Text = ""
    Company_value.Text = ""
    Position_value.Text = ""
    MonthlyBonus_value.Text = ""
    Salary_value.Text = ""
    ManagmentPosOfStructuralUnit_value.Text = ""
    StructuralUnit_value.Text = ""
    PlaceOfWork_value.Text = ""
    TypeContract_value.Value = ""
    ProbPeriod_value.Value = ""
    WorkSchedule_value.Value = ""
    TypeEmployment_value.Value = ""
    SN_value.Value = False
    DMS_value.Value = False
End Sub

Private Sub Fill_Click()
' Обработка кнопки "Заполнить"
    With ICP.ListObjects("ИКП").ListColumns(2)
        .DataBodyRange(1).Value = FullName_value.Text
        .DataBodyRange(2).Value = Company_value.Text
        .DataBodyRange(3).Value = Position_value.Text
        .DataBodyRange(4).Value = StructuralUnit_value.Text
        .DataBodyRange(5).Value = ManagmentPosOfStructuralUnit_value.Text
        .DataBodyRange(6).Value = PlaceOfWork_value.Text
        .DataBodyRange(7).Value = TypeEmployment_value.Text
        .DataBodyRange(8).Value = WorkSchedule_value.Text
        .DataBodyRange(9).Value = Salary_value.Text
        .DataBodyRange(10).Value = (MonthlyBonus_value.Text / 100)
        If SN_value.Value = True Then .DataBodyRange(12).Value = 1
        If DMS_value.Value = True Then .DataBodyRange(11).Value = 1
        .DataBodyRange(13).Value = ProbPeriod_value.Text
        .DataBodyRange(14).Value = TypeContract_value.Text
    End With
    
End Sub
