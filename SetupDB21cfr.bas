Attribute VB_Name = "SetupDB21cfr"
Sub setuup()
CreateDB
errorhandeling21cfrtableconstruct
roletable21cfrconstruct
usertable21cfrconstruct
End Sub
Sub uninstall()
droptable "errorhandeling"
End Sub


Sub errorhandeling21cfrtableconstruct()
createtable "errorhandeling"
Addcolumn "errorhandeling", "Datecol", "DATETIME DEFAULT =NOW()"
Addcolumn "errorhandeling", "Errornumber", "LONG NOT NULL"
Addcolumn "errorhandeling", "Error_des", "MEMO"
Addcolumn "errorhandeling", "Error_source", "TEXT"
Addcolumn "errorhandeling", "Error_fct", "TEXT"
Addcolumn "errorhandeling", "Who", "TEXT"
End Sub
Sub roletable21cfrconstruct()
Dim role(), rolename(0) As Variant
createtable "role"
Addcolumn "role", "role_name", "TEXT"
dbExecute "INSERT INTO role (role_name) VALUES ('User') ;"
dbExecute "INSERT INTO role (role_name) VALUES ('Administrator') ;"

End Sub
Sub usertable21cfrconstruct()
createtable "usertable"
Addcolumn "usertable", "user_ID", "TEXT"
Addcolumn "usertable", "user_firtstname", "TEXT"
Addcolumn "usertable", "user_lastname", "TEXT"
Addcolumn "usertable", "role_ID", "INTEGER DEFAULT =1 "
End Sub
