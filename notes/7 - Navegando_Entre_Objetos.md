# Introdução 

Este documento apresenta, em linguagem científica e estruturada, o funcionamento da hierarquia de objetos do Excel no VBA, incluindo navegação, reutilização e relação entre classes dentro do modelo de objetos. O objetivo é fornecer uma base sólida para compreender como o Excel organiza seus elementos internos, permitindo ao programador manipular dados e automatizar tarefas de forma precisa e avançada.

## Hierarquia Completa do Modelo de Objetos do Excel

```bash
Application (O próprio Excel em si)
    ├── Workbooks (coleção) (Pastas de trabalhos)
    │   └── Workbook
    │       ├── Sheets (coleção) (Planilhas levando em consideração os gráficos)
    │       │   ├── Worksheet (Planilhas sem lavar em consideração os gráficos )
    │       │   │ ├── Range (Intervalo de células)
    │       │   │ │     ├── Cells (coleção) (células do Excel)
    │       │   │ │     ├── Rows (coleção) (Linha do Excel)
    │       │   │ │     └── Columns (coleção) (Colunas do Excel)
    │       │   │ ├── ListObjects (Tabelas) 
    │       │   │ ├── Shapes (objetos gráficos)
    │       │   │ └── Names (intervalos nomeados)
    │       │   └── Chart
    │       ├── Names (nomes do workbook)
    │       └── Connections / PivotCaches / VBProject
    └── Global Objects
        ├── ActiveWorkbook
        ├── ActiveSheet
        ├── ActiveCell
        ├── ActiveWindow
        ├── Selection
        └── Events (Application-Level Events)
```

## 1 - Application

Application é o objeto raiz da hierarquia. Ele representa a instância inteira do Excel e fornece acesso global a:

- ambiente de execução,

- configurações gerais,

- coleções principais (Workbooks, AddIns, CommandBars),

- propriedades de controle (ScreenUpdating, DisplayAlerts, Calculation).

Exemplo:

```VB
    Application.ScreenUpdating = False
```

## 2 - Workbooks

Workbooks é uma coleção que representa todas as pastas de trabalho abertas na instância do Excel.

Acesso a um workbook específico:

- Workbooks("Relatorio.xlsx")

- Enumerar todos os workbooks abertos:

```Vb
Dim wb As Workbook
For Each wb In Application.Workbooks
    Debug.Print wb.Name
Next wb
```

## 3 - Workbook

Um Workbook representa um arquivo Excel (*.xls, *.xlsx, .xlsm). Ele contém:

- Coleção Sheets

- Coleção Names

- Propriedades como Path, FullName, Saved

Exemplo de acesso ao caminho:

```VB
MsgBox ActiveWorkbook.Path
```

### Exemplo de como navegar e acessar objetos com VBA

```VB

'  Se uma pasta de trabalho NÃO ESTÁ ATIVA, você pode acessar os objetos dessa pasta de trabalho da seguinte forma:
Workbooks("Pasta2.xlsm").Sheets("Planilha1").Range("A1").value = 1

' Ou seja, bem a grosso modo, acessanos a Pasta de Trabalho Planilha2.xlsx, acessamos a Planilha1 e na célula A1 atribuimos o valor 1 nela. 

' Entretanto, se a pasta de trabalho estiver Ativa, você pode omitir a referência ao objeto Workbook:
Sheets("Planilha1").Range("A1").value = 1

' O que isso quer dizer? Se estamos trabalhando justamente nesta pasta de trabalho...  Então não é necessário ativa-la, pois ela já está ativa e aberta
' Logo acessamos somente  as planilhas, celulas e intervalos de células

' E se você quiser interagir com a planilha ativa da pasta de trabalho, você pode também omitir o objeto Sheets (Planilhas):

Range("A1").value = 1


' Mesma Lógica acima, só que para  planilhas dentro desta pasta de trabalho.
```

## 4 - Sheets

Sheets é a coleção que engloba todos os tipos de folhas existentes dentro de um Workbook:

- Worksheets (planilhas tradicionais)

- Charts (planilhas de gráfico)

Acesso por índice ou nome:

```VB
Set sh = ActiveWorkbook.Sheets(1)
Set sh = ActiveWorkbook.Sheets("Planilha1")
```

## 5 - Worksheet

Worksheet representa uma planilha individual. É o objeto mais manipulado em VBA.

Exemplo:

```VB
Worksheets("Dados").Activate
```

Principais elementos contidos em um Worksheet:

- Range

- Cells

- Rows

- Columns

- ListObjects

- Shapes

- Names

## 6 - Range

Range é o objeto central de manipulação de dados. Ele representa uma célula ou conjunto de células.

Exemplos:

```VB
Range("A1")
Range("A1:B10")
Range("A:A")
```

Range é um contêiner para objetos Cells, Rows e Columns.

## 7 - Cells

Cells é uma coleção que representa todas as células de um Worksheet.

Acesso posicional:

```VB
Cells(1, 1)  ' Equivalente a Range("A1")
```

A propriedade Cells está sempre subordinada a um Range ou Worksheet.

## 8 - Rows

Rows representa todas as linhas de um Worksheet ou de um Range.

Exemplos:

```VB
Rows(1).Select
Range("A1:D10").Rows(3).Interior.Color = vbGreen
Debug.Print Rows.Count
```

Rows é conveniente para manipulação horizontal de dados.

## 9 - Columns

Columns é uma coleção que representa todas as colunas em um Worksheet ou dentro de um Range.


Hierarquia:

```bash 
Worksheet
 └── Range
       └── Columns
```

Exemplos:

```VB
Columns("C").Select
Range("A1:D10").Columns(2).Interior.Color = vbYellow
Debug.Print Range("A1:F10").Columns.Count
Columns("A").ColumnWidth = 20
```

## 10 - Names (Nomes definidos)

O objeto Names representa intervalos nomeados, tanto no Workbook quanto na Worksheet.

Exemplos:

```VB
ThisWorkbook.Names.Add Name:="TotalVendas", RefersTo:=Range("B2:B100")
MsgBox ThisWorkbook.Names("TotalVendas").RefersTo
```

Names é essencial para modelos automatizados e relatórios estruturados.

## 11 - Shapes

Shapes é a coleção de todos os objetos gráficos:

- caixas de texto,

- botões,

- imagens,

- linhas,

- objetos de desenho.

Exemplos:

```VB
    ActiveSheet.Shapes("Imagem1").Delete
    ActiveSheet.Shapes.AddShape msoShapeRectangle, 20, 20, 100, 60
```

Shapes permite criar interfaces gráficas dentro do Excel.

## 12 - ListObjects (Tabelas Estruturadas)

ListObjects representa as tabelas do Excel criadas via Inserir → Tabela.

Exemplos:

```Vb
Dim tbl As ListObject
Set tbl = Worksheets("Dados").ListObjects("Tabela1")
Debug.Print tbl.DataBodyRange.Rows.Count
```

ListObjects é fundamental em automações modernas de BI e manipulação tabular.

## 13 - Eventos (Application, Workbook, Worksheet)

Eventos permitem reagir a ações do usuário ou do próprio Excel.

Exemplos de eventos:

- Application.OnTime

- Workbook.Open

- Worksheet.Change

- Worksheet.SelectionChange

Exemplo em VBA:

```VB
Private Sub Worksheet_Change(ByVal Target As Range)
    MsgBox "Célula alterada: " & Target.Address
End Sub
```

Eventos são o núcleo do desenvolvimento orientado a eventos no Excel.