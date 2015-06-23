Function load_data()
Set rs = Nothing
ListView1.ListItems.clear
InPconn "select * from table1"

With rs
    If .RecordCount Then
        For Counter = 1 To .RecordCount
            ListView1.ListItems.Add , , rs(0).Value
            For X = 1 To rs.Fields.Count - 1
                ListView1.ListItems(Counter).ListSubItems.Add , , IIf(IsNull(rs(X).Value), "", rs(X).Value)
            Next X
            .MoveNext
        Next Counter
    Else
    End If
End With
End Function
