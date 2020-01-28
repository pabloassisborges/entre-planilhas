Sub Copiar dados de outro arquivo ()

'Passo 1: Declarações.
Dim wsOrigem As Worksheet
Dim wsDestino As Worksheet

'Passo 2: Especifica o intervalo que deseja excluir os dados na planilha de destino (Porque se tiver algo lá antes, tu tira né?).
Range("A1:CU100000").ClearContents

'Passo 3: Especifica o caminho do arquivo de origem.
Workbooks.Open Filename:="C:\Exemplo\Exemplo\Exemplo.xlsb"

'Passo 4: Especifica o nome e a aba do arquivo de origem, que deseja copiar os dados.
Set wsOrigem = Workbooks("NOME DO ARQUIVO.xlsb").Worksheets("NOME DA ABA")

' Passo 5: Especifica a aba no arquivo de destino, que deseja colar os dados.
Set wsDestino = ThisWorkbook.Sheets("NOME DA ABA")

'Passo 6: Realiza o procedimento de copiar e colar os dados, no intervalo que desejar. Neste caso está sendo copiado todos os dados da planilha, exceto a primeira linha.
With wsOrigem
Range("A2:CU100000").Copy Destination:=wsDestino.Range("A2:CU100000")
End With

'Passo 7: Especifica o nome da planilha de origem, para salvar e fechar.
Workbooks("NOME DO ARQUIVO.xlsb").Close SaveChanges:=True

End Sub