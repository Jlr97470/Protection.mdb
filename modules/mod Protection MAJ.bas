Attribute VB_Name = "mod Protection MAJ"
Option Compare Database
Option Explicit
' ---
' ACTIVATION / DESACTIVATION DE LA TOUCHE MAJ
' ---
'
Function DefinirToucheMAJ(strBaseDeDonnees As String, blnEtat As Boolean)
Dim db As DAO.Database
Dim prp As DAO.Property
Dim blnResultat As Boolean

Const PROPRIETE_NON_TROUVEE = 3270

' Ouvrir la base de données
If Dir(strBaseDeDonnees) = "" Then
    MsgBox "Base de données introuvable.", vbExclamation
    Exit Function
End If
Set db = OpenDatabase(strBaseDeDonnees)

' Définir la propriété
On Error Resume Next
db.Properties("AllowBypassKey") = blnEtat
If Err.Number = PROPRIETE_NON_TROUVEE Then
    Set prp = db.CreateProperty("AllowBypassKey", dbBoolean, blnEtat)
    db.Properties.Append prp
End If

' Libérer les objets
db.Close
Set db = Nothing
End Function
