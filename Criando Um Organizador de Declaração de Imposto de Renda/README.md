## Menu

- Construir um menu na lateral direita:
    - Selecionar a Coluna A
    - Aumentar o tamanho
    - Pintar com o código #0E1317
    - Criar um logo com aparência de ícone
    - Inserir o nome "LionForme"
        - Deixar transparente
        - Alinhar centralizado
        - Cor da fonte gradiente #EE37BF e #6124E8
    - Três botões com formas:
        - Titular
        - Fonte Verdana e Segoe UI Light
        - Cor gradiente entre #EE37BF e #6124E8
        - Deixar bordas arredondadas

## Itens do Menu

```VBA
Sub MoverIconeParaPosicao()
    Dim shp As Shape
    Dim ws As Worksheet
    Dim nomeIconeProcurado As String
    Dim novaPosicaoX As Double
    Dim novaPosicaoY As Double
    
    ' Defina a planilha atual
    Set ws = ActiveSheet
    
    ' Defina o nome do Ã­cone que vocÃª quer mover (exato, como aparece no Excel)
    nomeIconeProcurado = "Ãcone 1" ' <-- Troque aqui pelo nome do seu Ã­cone
    
    ' Defina a posiÃ§Ã£o desejada
    novaPosicaoX = 100 ' PosiÃ§Ã£o X em pontos
    novaPosicaoY = 50  ' PosiÃ§Ã£o Y em pontos
    
    ' Procura pelo Ã­cone na planilha
    For Each shp In ws.Shapes
        If shp.Name = nomeIconeProcurado Then
            ' Move o Ã­cone para a nova posiÃ§Ã£o
            shp.Left = novaPosicaoX
            shp.Top = novaPosicaoY
            MsgBox "Ãcone '" & nomeIconeProcurado & "' movido com sucesso!", vbInformation
            Exit Sub
        End If
    Next shp
    
    ' Se nÃ£o encontrar
    MsgBox "Ãcone '" & nomeIconeProcurado & "' nÃ£o encontrado.", vbExclamation
End Sub
```
### Criando funções para os itens

- Criar uma caixa para colocar o texto "System by {nome}"
- Colocar uma linha entre o botão e o texto
- Duplicar as abas em 3: "TITULAR", "INFORME" e "NOTAS"
- Deixar apenas a aba selecionada colorida
- Vincular o botão às abas

## Criando uma função no Excel

- Programar em VBA para ajustar os ícones
- Abrir o editor de macros
    - Inserir macro e adicionar o código
    - Adicionar a posição e o nome do ícone no script
    - Remover a macro

## Formulário do Titular

- Inserir:
    - Nome
    - CPF
    - Título
    - Cônjuge
    - Rua
    - Rua abreviada
    - CEP
    - Telefone
    - Celular
    - E-mail
    - Houve alterações
    - Dependente cônjuge
    - Residente no exterior

- Fonte: Segoe UI Light
- Ajustar o tamanho das células
- Inserir contexto "Título" - Dados do titular
- Inserir subtítulo:
    - "Preencha os dados da pessoa física abaixo"
- Inserir validação de dados para os 3 últimos campos
- Inserir um botão "Próximo"

## Formatações personalizadas

- Na formatação personalizada do Excel, dígito é representado por zero

## Telas de Informes

- Copiar o formulário para gerar uniformidade visual em todos os itens
- Título: Informes de rendimentos bancários
- Subtítulo: Preencha os dados atuais de cada banco
- Criar os campos:
    - Banco
    - Valor atual
    - Anexo
    - Usar a planilha com o código de todos os bancos
    - Validar o dado com o banco
    - Total: soma de todos os bancos
    - Adicionar o botão para o próximo campo

## Tela de Notas

- Título: Notas bancárias ou extrato/holerites
- Subtítulo: Insira todos os dados de receita
- Colunas:
    - Data
    - Categoria
        - Validar com:
            - CNPJ
            - Holerite
            - Freelance
    - Valor

## Toques Finais

- Remover barra de fórmulas e barra de título
- Desbloquear células necessárias
- Bloquear a planilha

## Links Úteis

- [Bancos](https://docs.google.com/spreadsheets/d/1wl4pcGdPb1HQHWTsneyI5MkrBYBceKP_/edit?usp=drive_link&ouid=101647331545466450705&rtpof=true&sd=true)
- [Projeto Exemplo](https://docs.google.com/spreadsheets/d/1mRuj0vnmiX3CUM3cyMVRB9-HQRcUY6DZ/edit?usp=drive_link&ouid=101647331545466450705&rtpof=true&sd=true)
- [Meu projeto](https://docs.google.com/spreadsheets/d/1DMM4MQJOggq-UOVtBpR_bNyXANoSFB7srNfs0EKqlSk/edit?usp=drive_link)
