VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=========================================================================================
' DECLARAÇÃO DE VARIÁVEIS PERSONALIZADAS DA CLASSE.
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
' DECLARAÇÃO DE VARIÁVEIS CONSTANTES DA CLASSE.
  Private Const aTagForm                As String = "frmCPBar"
  Private Const aNameFrame              As String = "fraCPBar"
  Private Const aNameBar                As String = "ctlCPBar"
  Private Const aMaxWidthBar            As Integer = 300
  Private Const aClass                  As String = "ApiFunction"
  Private Const aOpEx                   As String = "Option Explicit"
'=========================================================================================

'=========================================================================================
' DECLARAÇÃO DE VARIÁVEIS/PROPRIEDADES PRIVADAS DA CLASSE.
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
' DECLARAÇÃO DE VARIÁVEIS/PROPRIEDADES PUBLICAS DA CLASSE.
'Propriedade para o Tipo de Barra do ProgressBar.
Public Property Get BarType() As BarType: BarType = aTBar: End Property
Public Property Let BarType(pTBar As BarType): aTBar = pTBar: End Property
'Propriedade para o Tipo de Contador do ProgressBar.
Public Property Get CountType() As CountType: CountType = aTCount: End Property
Public Property Let CountType(pTCount As CountType): aTCount = pTCount: End Property
'Propriedade para definir o máximo de registros para o ProgressBar.
Public Property Get RecordMax() As Long: RecordMax = aRecordMax: End Property
Public Property Let RecordMax(pRecordMax As Long): aRecordMax = pRecordMax: End Property
'=========================================================================================

'=========================================================================================
' DECLARAÇÃO DOS MÉTODOS PUBLICOS DA CLASSE.
Public Sub Initialize()
  'Cria novo formulário.
  Call CreateNewForm
  'Cria caixa de grupo.
  Call CreateNewFrame
  'Cria barra com rótulo.
  Call CreateNewBar
  'Abre o formulário que foi criado.
  aForm.Show vbModeless
End Sub
Public Sub Update(ByVal pRecordNow As Long, Optional ByVal pRecordMax As Long, _
                  Optional ByVal pTextBar As String = "Carregando Registros...")
  'Confronto entre atributo e parâmetro RecordMax.
  If pRecordMax > aRecordMax Then aRecordMax = pRecordMax
  'Define o percentual de pRecordNow sobre aRecordMax.
  aPercentil = pRecordNow / aRecordMax
  'Define valor retorno contador do ProgressBar.
  aFrame.Caption = VBA.IIf(aTCount = Percentage, _
          VBA.Format(aPercentil * 100, "0.00") & "% Concluído.", _
          pRecordNow & " de " & aRecordMax & " Concluído.")
  'aFrame.Caption = aPercentil * 100
  'Define valor retorno para a barra do ProgressBa.
  If aTBar = Progress Then
    aProgressBar.Width = aMaxWidthBar * aPercentil
  Else
    aProgressBar.Text = pTextBar
  End If
  'Atualiza a barra em tempo de execução.
  VBA.DoEvents
End Sub
'=========================================================================================

'=========================================================================================
' DECLARAÇÃO DOS MÉTODOS PRIVADOS DA CLASSE.
'=========================================================================================

'=========================================================================================
' DECLARAÇÃO DOS EVENTOS DA CLASSE.
Private Sub Class_Initialize()
  'Define alguns atributos padrões.
  aTBar = Progress: aTCount = Percentage
End Sub
Private Sub Class_Terminate()
  'Deleta o formulário criado.
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
' DECLARAÇÃO DOS EVENTOS DA CLASSE.
Private Sub CreateNewForm()
  'Atribui o novo VbComponent ao atributo aForm.
  Set aControl = ThisWorkbook.VBProject.VBComponents.Add(3)
  'Adiciona o Userform criado a lista de Userforms.
  VBA.UserForms.Add aControl.Name
  'Atribui o novo Userform ao atributo aForm.
  Set aForm = UserForms(UserForms.Count - 1)
  'Chama rotina para configurar formulário.
  Call FormConfig
End Sub
Private Sub FormConfig()
  'Configura a estrutura do novo Userform.
  With aForm
    .Tag = aTagForm
    .Caption = ""
    .Width = 344
    .Height = VBA.IIf(aTBar = Progress, 85, 90)
  End With
End Sub
Private Sub CreateNewFrame()
  'Atribui o novo Frame ao atributo aFrame.
  Set aFrame = aForm.Controls.Add("Forms.frame.1", aNameFrame, True)
  'Chama rotina para configurar quadro.
  Call FrameConfig
End Sub
Private Sub FrameConfig()
  'Configura a estrutra do novo Frame.
  With aFrame
    .Height = VBA.IIf(aTBar = Progress, 42, 48)
    .Left = 6
    .Top = 6
    .Width = 320
    .Caption = ""
  End With
End Sub
Private Sub CreateNewBar()
  'Verifica se o controle já esta atribuido ou existe no projeto.
  For Each aControl In aForm.Controls
    'Testa o elemento com o nome utilizado no projeto e o deleta.
    If aControl.Name = aNameBar Then aFrame.Controls.Remove aControl
  Next aControl
  'Atribui a no Control ao atributo aProgressBar.
  Set aProgressBar = aFrame.Controls.Add(VBA.IIf( _
                     aTBar = Progress, "Forms.Label.1", "Forms.TextBox.1"), _
                     aNameBar, True)
  'Chama rotina para configurar barra.
  Call BarConfig
End Sub
Private Sub BarConfig()
  'Configura a estrutura do novo Control.
  With aProgressBar
    .Height = VBA.IIf(aTBar = Progress, 12, 18)
    .Left = VBA.IIf(aTBar = Progress, 9, 12)
    .Top = 12
    .Width = VBA.IIf(aTBar = Progress, 0, aMaxWidthBar)
    .BackColor = VBA.IIf(aTBar = Progress, vbBlue, &H8000000F)
    If aTBar <> Progress Then .Locked = True
  End With
End Sub
'=========================================================================================
