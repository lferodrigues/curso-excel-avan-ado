VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
Inseri_dados

End Sub

Private Sub CommandButton2_Click()
limpeza
End Sub

Private Sub CommandButton3_Click()
busca_data
End Sub

Private Sub ComboBox1_Change()

End Sub

Private Sub ComboBox2_Change()

End Sub

Private Sub CommandButton4_Click()
procura_vertical
monta_tabela
gera_filtro
Gera_tabela

End Sub
