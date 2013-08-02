VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmFormulieren"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnCancel_Click()
    DoCmd.Close
End Sub

Private Sub btnOk_Click()
    Dim objWord As Word.Application
    Dim dbs As DAO.Database
    Dim rstPosten As DAO.Recordset
    
    Dim intCounter As Integer
    Dim strWerkblad As String
        
    On Error GoTo MergeButton_Err
    
    intCounter = 1
    
    Set dbs = CurrentDb
    Set rstPosten = dbs.OpenRecordset("tblPosten", dbOpenTable)
    rstPosten.MoveFirst
        
    Set objWord = CreateObject("Word.Application")
    objWord.Visible = True
    
    Me.edtCounter.SetFocus
    
    Do While Not rstPosten.EOF
        Me.edtCounter.Text = CStr(intCounter)
        Me.Refresh
        DoEvents
        
        Select Case rstPosten.Fields("Werkblad").Value
            Case "A"
                strWerkblad = "A - Aanstelling.doc"
            Case "B"
                strWerkblad = "B - Ontslag.doc"
            Case "C"
                strWerkblad = "C - Wijziging arbeidstijd.doc"
            Case "D"
                strWerkblad = "D - Bevordering.doc"
            Case "E"
                strWerkblad = "E - Vaste toelagen.doc"
            Case Else
                MsgBox "Onbekende werkblad code: " & rstPosten.Fields("Werkblad").Value
                Exit Sub
        End Select
                
        With objWord
            ' Open the document.
            .Documents.Open ("H:\aaa\" & strWerkblad)
            
            ' Move to each bookmark and insert text from the form.
            .ActiveDocument.Bookmarks("bmDienst").Select
            .Selection.Text = rstPosten.Fields("Dienst").Value
                                    
            .ActiveDocument.Bookmarks("bmVolgNr").Select
            .Selection.Text = rstPosten.Fields("Vlg_nr").Value
            
            .ActiveDocument.Bookmarks("bmNaam").Select
            .Selection.Text = rstPosten.Fields("Naam").Value
            
            .ActiveDocument.Bookmarks("bmSapNr").Select
            .Selection.Text = rstPosten.Fields("Sap_nr").Value
            
            .ActiveDocument.Bookmarks("bmMaand").Select
            .Selection.Text = rstPosten.Fields("Maand").Value
           
            .ActiveDocument.Bookmarks("bmLC").Select
            .Selection.Text = rstPosten.Fields("LC_code").Value
        
            .ActiveDocument.Bookmarks("bmLcOmschrijving").Select
            .Selection.Text = rstPosten.Fields("LC_omschrijving").Value
        
            .ActiveDocument.Bookmarks("bmAardMutatie").Select
            .Selection.Text = rstPosten.Fields("Aard_mutatie").Value
                
        End With
        
        ' Print the document in the foreground so Microsoft Word 97
        ' will not close until the document finishes printing.
        objWord.ActiveDocument.SaveAs "H:\aaa\" & rstPosten.Fields("Vlg_nr").Value & ".doc"
        
        ' Close the document without saving changes.
        objWord.ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
        
        intCounter = intCounter + 1
        rstPosten.MoveNext
    Loop
    
    ' Quit Microsoft Word 97 and release the object variable.
    objWord.Quit
    Set objWord = Nothing
    
    rstPosten.Close
    Set rstPosten = Nothing

    MsgBox "Overzichten gereed."
    DoCmd.Close

    Exit Sub
    
MergeButton_Err:
    If Err.Number = 94 Then
        objWord.Selection.Text = ""
        Resume Next
    Else
        MsgBox Err.Number & vbCr & Err.Description
    End If
End Sub

