VERSION 5.00
Begin VB.Form frmLerChaveOffice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ler Chave Office"
   ClientHeight    =   3930
   ClientLeft      =   1155
   ClientTop       =   2070
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   7095
   Begin VB.TextBox txtChaveOneNoteVersao 
      Height          =   315
      Left            =   2400
      TabIndex        =   17
      Top             =   900
      Width           =   4575
   End
   Begin VB.TextBox txtChaveMSProjectVersao 
      Height          =   315
      Left            =   2400
      TabIndex        =   15
      Top             =   1980
      Width           =   4575
   End
   Begin VB.TextBox txtChaveVisioVersao 
      Height          =   315
      Left            =   2400
      TabIndex        =   13
      Top             =   2700
      Width           =   4575
   End
   Begin VB.TextBox txtChavePublisherVersao 
      Height          =   315
      Left            =   2400
      TabIndex        =   11
      Top             =   2340
      Width           =   4575
   End
   Begin VB.TextBox txtChavePowerPointVersao 
      Height          =   315
      Left            =   2400
      TabIndex        =   9
      Top             =   1620
      Width           =   4575
   End
   Begin VB.TextBox txtChaveOutlookVersao 
      Height          =   315
      Left            =   2400
      TabIndex        =   7
      Top             =   1260
      Width           =   4575
   End
   Begin VB.TextBox txtChaveAccessVersao 
      Height          =   315
      Left            =   2400
      TabIndex        =   5
      Top             =   180
      Width           =   4575
   End
   Begin VB.TextBox txtChaveWordVersao 
      Height          =   315
      Left            =   2400
      TabIndex        =   3
      Top             =   3060
      Width           =   4575
   End
   Begin VB.TextBox txtChaveExcelVersao 
      Height          =   315
      Left            =   2400
      TabIndex        =   2
      Top             =   540
      Width           =   4575
   End
   Begin VB.CommandButton cmdLerChaves 
      Caption         =   "Ler &Chaves"
      Height          =   375
      Left            =   5760
      TabIndex        =   0
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lblChaveOneNoteVersao 
      Caption         =   "Chave do OneNote/Versão:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label lblChaveMSProjectVersao 
      Caption         =   "Chave do MSProject/Versão:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label lblChaveVisioVersao 
      Caption         =   "Chave do Visio/Versão:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label lblChavePublisherVersao 
      Caption         =   "Chave do Publisher/Versão:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label lblChavePowerPointVersao 
      Caption         =   "Chave do PowerPoint/Versão:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label lblChaveOutlookVersao 
      Caption         =   "Chave do Outlook/Versão:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label lblChaveAccessVersao 
      Caption         =   "Chave do Access/Versão:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label lblChaveWordVersao 
      Caption         =   "Chave do Word/Versão:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label lblChaveExcelVersao 
      Caption         =   "Chave do Excel/Versão:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
End
Attribute VB_Name = "frmLerChaveOffice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------
'Links Interessantes:
'---------------------------------------------------------------
'CreateObject função (Visual Basic) - Visual Studio 2008
'http://msdn.microsoft.com/pt-br/library/7t9k08y5(v=vs.90).aspx
'---------------------------------------------------------------
'INFORMAÇÕES: Solucionando o erro 429 ao automatizar aplicativos do Office
'http://support.microsoft.com/kb/244264/pt-br
'
'Para verificar a chave do caminho que está armazenada para o servidor, inicie _
 o Editor do Registro do Windows, digitando regedit no menu Iniciar e depois na _
 caixa de diálogo Executar. Navegue até a chave HKEY_CLASSES_ROOT\Clsid. _
 Nessa chave você encontrará os CLSIDs para os servidores de automação _
 registrados no sistema. Depois, usando os valores, encontre a chave que _
 representa o aplicativo do Office que você quiser automatizar e verifique o _
 caminho da chave LocalServer32 do mesmo.
'
'   +========================+=========================================+
'   | Servidor do Office     | Chave CLSID                             |
'   +========================+=========================================+
'   | Access.Application     | {73A4C9C1-D68D-11D0-98BF-00A0C90DC8D9}  |
'   +------------------------+-----------------------------------------+
'   | Excel.Application      | {00024500-0000-0000-C000-000000000046}  |
'   +------------------------+-----------------------------------------+
'   | FrontPage.Application  | {04DF1015-7007-11D1-83BC-006097ABE675}  |
'   +------------------------+-----------------------------------------+
'   | Outlook.Application    | {0006F03A-0000-0000-C000-000000000046}  |
'   +------------------------+-----------------------------------------+
'   | PowerPoint.Application | {91493441-5A91-11CF-8700-00AA0060263B}  |
'   +------------------------+-----------------------------------------+
'   | Word.Application       | {000209FF-0000-0000-C000-000000000046}  |
'   +------------------------+-----------------------------------------+
'
'---------------------------------------------------------------
'Microsoft Office Compatibility Pack for Word, Excel, and PowerPoint File Formats
'http://www.microsoft.com/pt-BR/download/details.aspx?id=3
'---------------------------------------------------------------
'WINREG32.BAS (Third Version)
'http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=5354&lngWId=1
'---------------------------------------------------------------

Option Explicit

Private Sub cmdLerChaves_Click()

    '---------------------------------------------------------------
    ' Ler chave do Access/Versão
    txtChaveAccessVersao.Text = objBuscaChaves.BuscaVersao("Access")
    '---------------------------------------------------------------
    ' Ler chave do Excel/Versão
    txtChaveExcelVersao.Text = objBuscaChaves.BuscaVersao("Excel")
    '---------------------------------------------------------------
    ' Ler chave do OneNote/Versão
    txtChaveOneNoteVersao.Text = objBuscaChaves.BuscaVersao("OneNote")
    '---------------------------------------------------------------
    ' Ler chave do Outlook/Versão
    txtChaveOutlookVersao.Text = objBuscaChaves.BuscaVersao("Outlook")
    '---------------------------------------------------------------
    ' Ler chave do PowerPoint/Versão
    txtChavePowerPointVersao.Text = objBuscaChaves.BuscaVersao("PowerPoint")
    '---------------------------------------------------------------
    ' Ler chave do MSProject/Versão
    txtChaveMSProjectVersao.Text = objBuscaChaves.BuscaVersao("MSProject")
    '---------------------------------------------------------------
    ' Ler chave do Publisher/Versão
    txtChavePublisherVersao.Text = objBuscaChaves.BuscaVersao("Publisher")
    '---------------------------------------------------------------
    ' Ler chave do Visio/Versão
    txtChaveVisioVersao.Text = objBuscaChaves.BuscaVersao("Visio")
    '---------------------------------------------------------------
    ' Ler chave do Word/Versão
    txtChaveWordVersao.Text = objBuscaChaves.BuscaVersao("Word")
    '---------------------------------------------------------------

End Sub

Private Sub Form_Load()

End Sub
