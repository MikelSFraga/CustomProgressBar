VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------
' Classe          : cProgressBar Class
' Data            :
' Developer       : Mikel S Fraga
' Contato/GitHub  : https://github.com/MikelSFraga/CustomProgressBar
' Descri��o       : Cria��o de um Progress Bar Personalizada, atrav�s da cria��o de um Userform
'                   dinamicamente exobindo a evolu��o de um processo envocado simultaneamente.
'-----------------------------------------------------------------------------------------
Option Explicit
'=========================================================================================
' DECLARA��O DE VARI�VEIS PERSONALIZADAS DA CLASSE.
Enum BarType
  Progress = 0
  DisplayText = 1
End Enum
Enum CountType
  Percentage = 0
  Quantity = 1
End Enum
'=========================================================================================

'=========================================================================================
' DECLARA��O DE VARI�VEIS CONSTANTES DA CLASSE.
  Private Const aTagForm                As String = "frmCPBar"
  Private Const aNameFrame              As String = "fraCPBar"
  Private Const aNameBar                As String = "ctlCPBar"
  Private Const aMaxWidthBar            As Integer = 300
  Private Const aClass                  As String = "ApiFunction"
'=========================================================================================

'=========================================================================================
' DECLARA��O DE VARI�VEIS/PROPRIEDADES PRIVADAS DA CLASSE.
'Declare Controls
  Private aControl                      As Object 'VBIDE.VBComponent
  Private aForm                         As Object 'MSForms.UserForm
  Private aFrame                        As Object 'MSForms.Frame
  Private aProgressBar                  As Object 'MSForms.Control
'Declare UDV (User Definition Variable)
  Private aTBar                         As BarType
  Private aTCount                       As CountType
'Declare Validate
  Private aApiClass                     As Boolean
  Private aRecordMax                    As Long
  Private aRecordNow                    As Long
  Private aPercentil                    As Double
'=========================================================================================

'=========================================================================================
' DECLARA��O DE VARI�VEIS/PROPRIEDADES PUBLICAS DA CLASSE.
'Propriedade para o Tipo de Barra do ProgressBar.
Public Property Get BarType() As BarType: BarType = aTBar: End Property
Public Property Let BarType(pTBar As BarType): aTBar = pTBar: End Property
'Propriedade para o Tipo de Contador do ProgressBar.
Public Property Get CountType() As CountType: CountType = aTCount: End Property
Public Property Let CountType(pTCount As CountType): aTCount = pTCount: End Property
'Propriedade para definir o m�ximo de registros para o ProgressBar.
Public Property Get RecordMax() As Long: RecordMax = aRecordMax: End Property
Public Property Let RecordMax(pRecordMax As Long): aRecordMax = pRecordMax: End Property
'=========================================================================================

'=========================================================================================
' DECLARA��O DOS M�TODOS PUBLICOS DA CLASSE.
Public Sub Initialize()
  'This is to stop screen flashing while creating form
  Application.VBE.MainWindow.Visible = False
  'Cria novo formul�rio.
  Call CreateNewForm
  'Cria caixa de grupo.
  Call CreateNewFrame
  'Cria barra com r�tulo.
  Call CreateNewBar
  'Abre o formul�rio que foi criado.,
  Set aForm = UserForms(UserForms.Count - 1)
  aForm.Show vbModeless
End Sub
Public Sub Update(ByVal pRecordNow As Long, Optional ByVal pRecordMax As Long, _
                  Optional ByVal pTextBar As String = "Carregando Registros...")
  'Confronto entre atributo e par�metro RecordMax.
  If pRecordMax > aRecordMax Then aRecordMax = pRecordMax
  'Define o percentual de pRecordNow sobre aRecordMax.
  aPercentil = pRecordNow / aRecordMax
  'Define valor retorno contador do ProgressBar.
  aFrame.Caption = VBA.IIf(aTCount = Percentage, _
          VBA.Format(aPercentil * 100, "0.00") & "% Conclu�do.", _
          pRecordNow & " de " & aRecordMax & " Conclu�do.")
  'Define valor retorno para a barra do ProgressBa.
  If aTBar = Progress Then
    aProgressBar.Width = aMaxWidthBar * aPercentil
  Else
    aProgressBar.Text = pTextBar
  End If
  'Atualiza a barra em tempo de execu��o.
  VBA.DoEvents
  'Oculta Userform se for o �ltimo registro.
  If pRecordNow = aRecordMax Then aForm.Hide
End Sub
'=========================================================================================

'=========================================================================================
' DECLARA��O DOS M�TODOS PRIVADOS DA CLASSE.
Private Function IsClassPresent(ByVal pClass As String) As Boolean
  'Cria um objeto local do tipo VBComponent.
  Dim aVbComp                           As VBComponent
  'Define um valor padr�o para retorno da classe.
  IsClassPresent = False
  'Realiza um lan�o entre todos os componentes do projeto.
  For Each aVbComp In ThisWorkbook.VBProject.VBComponents
    'Testa todos os componentes, verificando se o ApiFunction esta presente.
    If aVbComp.Name = pClass Then IsClassPresent = True: Exit Function
  Next aVbComp
End Function
'=========================================================================================

'=========================================================================================
' DECLARA��O DOS EVENTOS DA CLASSE.
Private Sub Class_Initialize()
  'Define alguns atributos padr�es.
  aTBar = Progress: aTCount = Percentage
End Sub
Private Sub Class_Terminate()
  'Deleta o formul�rio criado.
  Set aControl = ThisWorkbook.VBProject.VBComponents(aForm.Name)
  ThisWorkbook.VBProject.VBComponents.Remove aControl
  'Clear the variables.
  Set aProgressBar = Nothing
  Set aFrame = Nothing
  Set aForm = Nothing
  Set aControl = Nothing
End Sub
'=========================================================================================

'=========================================================================================
' DECLARA��O DOS EVENTOS DA CLASSE.
Private Sub CreateNewForm()
  'Atribui o novo VbComponent ao atributo aForm.
  Set aControl = ThisWorkbook.VBProject.VBComponents.Add(3)
  'Configura a estrutura do novo Userform.
  With aControl
    .Properties("Tag") = aTagForm
    .Properties("Caption") = ""
    .Properties("Width") = 344
    .Properties("Height") = VBA.IIf(aTBar = Progress, 75, 85)
  End With
  'Adiciona o Userform criado a lista de Userforms.
  VBA.UserForms.Add aControl.Name
  'Atribui o novo Userform ao atributo aForm.
  Set aForm = VBA.UserForms(UserForms.Count - 1)
End Sub
Private Sub CreateNewFrame()
  'Atribui o novo Frame ao atributo aFrame.
  Set aFrame = aForm.Controls.Add("Forms.frame.1", aNameFrame, True)
  'Configura a estrutra do novo Frame.
  With aFrame
    .Height = VBA.IIf(aTBar = Progress, 34, 43)
    .Left = 6
    .Top = 6
    .Width = 320
    .Caption = ""
    .SpecialEffect = fmSpecialEffectSunken
  End With
End Sub
Private Sub CreateNewBar()
  'Atribui a no Control ao atributo aProgressBar.
  Set aProgressBar = aFrame.Controls.Add(VBA.IIf( _
                     aTBar = Progress, "Forms.Label.1", "Forms.TextBox.1"), _
                     aNameBar, True)
  'Configura a estrutura do novo Control.
  With aProgressBar
    .Height = VBA.IIf(aTBar = Progress, 12, 18)
    .Left = VBA.IIf(aTBar = Progress, 9, 12)
    .Top = VBA.IIf(aTBar = Progress, 8, 12)
    .Width = VBA.IIf(aTBar = Progress, 0, aMaxWidthBar)
    .BackColor = VBA.IIf(aTBar = Progress, vbBlue, &H8000000F)
    If aTBar <> Progress Then .Locked = True
  End With
End Sub
'=========================================================================================
