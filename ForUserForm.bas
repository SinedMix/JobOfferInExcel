Attribute VB_Name = "ForUserForm"
Sub CallUserForm()

    With ICPform.WorkSchedule_value
        .AddItem "������ ����"
        .AddItem "������� ������"
    End With
    
    With ICPform.TypeEmployment_value
        .AddItem "������"
        .AddItem "��������"
        .AddItem "������ ����"
    End With
    
    With ICPform.TypeContract_value
        .AddItem "�������"
        .AddItem "����������"
    End With
    
    With ICPform.ProbPeriod_value
        .AddItem "1 �����"
        .AddItem "2 ������"
        .AddItem "3 ������"
    End With
    
    ICPform.Show
End Sub
