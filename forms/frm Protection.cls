Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnAnnuler_Click()
DoCmd.Close
End Sub

Private Sub btnParcourir_Click()
Dim fd As FileDialog

On Error Resume Next
Set fd = FileDialog(msoFileDialogOpen)
fd.Title = "Sélectionnez une base de données..."
fd.AllowMultiSelect = False
With fd.Filters
    .Add "Bases de données", "*.mdb"
    .Add "Fichiers MDE", "*.mde"
    .Add "Tous les fichiers", "*.*"
End With
If fd.Show = True Then
    Me.txtBase = fd.SelectedItems(1)
End If
Set fd = Nothing
End Sub

Private Sub btnProtect_Click()
If IsNull(Me.txtBase) Then
    MsgBox "Sélectionnez une base de données au préalable !", vbInformation
    Exit Sub
End If
If Dir(Me.txtBase) = "" Then
    MsgBox "Fichier introuvable !", vbExclamation
    Exit Sub
End If
DefinirToucheMAJ Me.txtBase, (Me.optProtection = 1)
If Me.optProtection = 1 Then
    MsgBox "Touche Majuscule activée", vbInformation
Else
    MsgBox "Touche Majuscule désactivée", vbInformation
End If
End Sub
