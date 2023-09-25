Attribute VB_Name = "ForUserForm"
Sub CallUserForm()

    With ICPform.WorkSchedule_value
        .AddItem "полный день"
        .AddItem "сменный график"
    End With
    
    With ICPform.TypeEmployment_value
        .AddItem "гибрид"
        .AddItem "удаленка"
        .AddItem "полный офис"
    End With
    
    With ICPform.TypeContract_value
        .AddItem "срочный"
        .AddItem "бессрочный"
    End With
    
    With ICPform.ProbPeriod_value
        .AddItem "1 мес€ц"
        .AddItem "2 мес€ца"
        .AddItem "3 мес€ца"
    End With
    
    ICPform.Show
End Sub
