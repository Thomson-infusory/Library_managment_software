Attribute VB_Name = "Module1"
Public con As New ADODB.Connection
Public rs As New ADODB.Recordset

Sub main()
    con.Open "jegathConnect"
    frmSplashScreen.Show
End Sub
