Function load_data()
	rsconn "select * from tablename"
	Do Until rs.EOF
	With ListView1
	Set li = .ListItems.Add(, , rs.Fields!fieldname)
    		li.SubItems(1) = rs.Fields!fieldname
	End With
	rs.MoveNext
	Loop
End Function
