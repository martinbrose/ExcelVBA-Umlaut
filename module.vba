Function umlaut(text As String, Optional replaceEMPTYby As String = "")

Dim umlaut1 As String, rplString As String
Dim i As Long, j As Long
Dim MyArray

    '~~> One time slogging
rplString = "EUR,,,f,,,,,,,S,,OE,,Z,,,,,,,,,,,(TM),s,,oe,,z,Y,,i,c,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,A,A,A,A,Ae,A,A,C,E,E,E,E,I,I,I,I,G,N,O,O,O,O,Oe,x,0,U,U,U,Ue,Y,b,ss,a,a,a,a,ae,a,ae,c,e,e,e,e,i,i,i,i,o,n,o,o,o,o,o,-,o,u,u,u,u,y,b,y" '<~~ and so on.
    '~~> The first one before the comma is empty since we do
    '~~> not have any replacement for character represented by 128.
    '~~> The next one is for 129 and then 130 and so on so forth.
    '~~> The characters for which you do not have the replacement,
    '~~> leave them empty
    'how to find out your own signs: in Excel in Cell A128 type formula =CHAR(ROW())
    'copy that down to 255. replace characters not wanted by the charcater wanted.
    'in B128 formula: =A128
    'in all cells from B129 down to 255 type/copy formula: =CONCATENATE(R[-1]C,"","",RC[-1])
    'paste the value from B255 in "rplstring" above!

If replaceEMPTYby <> "" Then
    rplString = Replace(rplString, ",,", "," & replaceEMPTYby & ",")
    rplString = Replace(rplString, ",,", "," & replaceEMPTYby & ",")
    rplString = Replace(rplString, ",,", "," & replaceEMPTYby & ",")
    If Mid(rplString, 1, 1) = "," Then rplString = replaceEMPTYby & rplString
    If Mid(rplString, Len(rplString), 1) = "," Then rplString = rplString & replaceEMPTYby
    Debug.Print rplString
End If

    MyArray = Split(rplString, ",")
    umlaut1 = text: j = 0

    For i = 128 To 255
        umlaut1 = Replace(umlaut1, Chr(i), MyArray(j))
        j = j + 1
    Next
    umlaut = umlaut1
End Function
