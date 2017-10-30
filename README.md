# cProgressBar V1.0.0

O __Custom Progress Bar__, é uma classe desenvolvida para simular uma __Barra de Progresso Personalizada__. Essa classe pode ser utilizada em aplicações desenvolvidas com __Microsoft Excel__ e __Visual Basic for Application__. A classe consiste na construção de um ___Useform___, ___Em Tempo de Execução___, que simula uma __ProgressBar__, semelhante ao da biblioteca __Microsoft Common Control (ListView, TreeView, ProgressBar, StatusBar, etc)__.

A mesma é considerada personalizada, pois é possível realizar alterações na estrutura da mesma, da seguinte maneira:
- __Contador__: pode assumir um valor Percentual, ou Quantitativo, da evolução dos processos;
- __Barra__: pode assumir uma forma de Barra Crescente, ou Display de Informações, com os valores a serem processados.

Veremos o detalhamento dos __Métodos__ e __Propriedades__ disponibilizados na ___cProgressBar___.

***

#### Métodos e Propriedades Classe

Através da criação destas ___Propriedades___ e ___Métodos___, tivemos o intuito de que a mesma funcionasse como um ___Framework de Projetos VBA___, onde a mesma possa ser chamada em qualquer momento do projeto, seja em um _Módulo_ ou um _Formulário de Pesquisa_. 
Ressaltando que os nomes das ___Propriedades___ e ___Métodos___ são em Inglês, para seguir o padrão da língua estrangeira, utilizada como base para a __Programação VBA__.
***
#### Propriedades

##### RecordMax

Essa propriedade nos permite definir/informar, o Número Máximo de Registros, que a cProgressBar vai contabilizar e exibir para os usuários do sistema/aplicação. 

__Exemplo:__
```vb
Option Explicit

Private Sub RecordsLoad()
  ' Declaração do objeto da classe.
  Dim oProgressBar As New cProgressBar
  ' Instancia o objeto para uso no ambiente.
  Set oProgressBar = New cProgressBar
  ' Define o valor máximo de registros de uma pesquisa.
  oProgressBar.RecordMax = ws.UsedRange.Rows.Count ' Total de linhas usadas em uma Sheet, por exemplo.
End Sub
```

##### CountType

Essa propriedade nos permite definir o __Tipo do Contador__ que será utilizado na __Barra de Progresso__.

__Exemplo 1:__ Valor definido como ___Percentage___ (_default_)
```vb
Option Explicit

Private Sub RecordsLoad()
  ' Declaração do objeto da classe.
  Dim oProgressBar As New cProgressBar
  ' Instancia o objeto para uso no ambiente.
  Set oProgressBar = New cProgressBar
  ' Define o tipo de contador como Percentual.
  oProgressBar.CountType = Percentage
End Sub
```
<p align="center">
  <img src="https://github.com/MikelSFraga/CustomProgressBar/blob/master/img/CountType_Percentage.png">
</p>

__Exemplo 2:__ Valor definido como ___Quantitative___
```vb
Option Explicit

Private Sub RecordsLoad()
  ' Declaração do objeto da classe.
  Dim oProgressBar As New cProgressBar
  ' Instancia o objeto para uso no ambiente.
  Set oProgressBar = New cProgressBar
  ' Define o tipo de contador como Quantitativo.
  oProgressBar.CountType = Quantitative
End Sub
```
<p align="center">
  <img src="https://github.com/MikelSFraga/CustomProgressBar/blob/master/img/CountType_Quantity.png">
</p>

##### BarType

Essa propriedade nos permite definir o __Tipo de Barra__ que será utilizado na __Barra de Progresso__.

__Exemplo 1:__ Valor definido como ___Progress___ (_default_)
```vb
Option Explicit

Private Sub RecordsLoad()
  ' Declaração do objeto da classe.
  Dim oProgressBar As New cProgressBar
  ' Instancia o objeto para uso no ambiente.
  Set oProgressBar = New cProgressBar
  ' Define o tipo de barra como Progresso.
  oProgressBar.BarType = Progress
End Sub
```
<p align="center">
  <img src="https://github.com/MikelSFraga/CustomProgressBar/blob/master/img/BatType_Progress.png">
</p>

__Exemplo 2:__ Valor definido como ___DisplayText___
```vb
Option Explicit

Private Sub RecordsLoad()
  ' Declaração do objeto da classe.
  Dim oProgressBar As New cProgressBar
  ' Instancia o objeto para uso no ambiente.
  Set oProgressBar = New cProgressBar
  ' Define o tipo de barra como Exibição de Texto.
  oProgressBar.CountType = DisplayText
End Sub
```
<p align="center">
  <img src="https://github.com/MikelSFraga/CustomProgressBar/blob/master/img/BarType_DisplayText.png">
</p>

***
#### Métodos

##### Initialize

Esse método, quando chamado, inicia a construção do ___Userform___, que exibirá a evolução dos processos definidos no sistema/aplicação vinculado, a partir dos valores definidos para as propriedades __CountType__ e __BarType__ (assumindo os valores padrões, caso o desenvolvedor não os tenha definido).

Este método não possui nenhum parâmetro a ser passado/informado.

__Exemplo:__
```vb
Option Explicit

Private Sub RecordsLoad()
  ' Declaração do objeto da classe.
  Dim oProgressBar As New cProgressBar
  ' Instancia o objeto para uso no ambiente.
  Set oProgressBar = New cProgressBar
  ' Chama método de construção do Userform, instanciado pelo cProgressBar.
  oProgressBar.Initialize
End Sub
```

##### Update

Através deste método, que as informações de leitura dos registros, serão enviados para o Userfore, através da classe cProgressBar. A cada laço do processo em análise na rotina VBA, faz-se uma chamada para esse método, atualizando o registro atual a ser analisado pela rotina.

Este método possui três parâmetros a ser passados/informados, conforme abaixo:
- __pRecordNow__: _obrigatório_, esse parâmetro informa a classe, o número (valor) do registro atual e envia para o ___Engine da Classe___, que irá realizar os cálculos necessários e exibir a evolução do processo (tanto pelo __Contador__, como pela __Barra__);
- __pRecordMax__: _opicional_, através desse parâmetro, é possível informar para a classe o _Valor Máximo de Registros_ existentes no processo atual. Este parâmetro pode substituir a propriedade ___RecordMax___, ou informar a classe um _Novo Valor Máximo de Registros_, caso necessário; 
- __pTextBar__: _opicional_, quando o valor da propriedade __BarType__ for definido como ___DisplayText___, deve-se utilizar esse parâmetro para definir o tipo de texto/informação que será exibida na __Barra de Progresso__. Por exemplo, se esta gerando um relatório de pedidos, pode adicionar ao texto o número do pedido do registro atual que esta sendo analisado.

__Exemplo:__ 
```vb
Option Explicit

Private Sub RecordsLoad()
  ' Declara variáveis.
  Dim RegistroAtual, TotalRegistros As Long
  TotalRegistros = Sheets("Plan1").UsedRange.Rows.Count
  ' Declaração do objeto da classe.
  Dim oProgressBar As New cProgressBar
  ' Instancia o objeto para uso no ambiente.
  Set oProgressBar = New cProgressBar
  With oProgressBar
    ' Chama método de construção do Userform, instanciado pelo cProgressBar.
    .Initialize
    For RegistroAtual = 1 To TotalRegistros
      ' Chama método que irá atualizar a evolução do processo.
      .Update RegistroAtual, TotalRegistros, "Registro número " & RegistroAtual
    Next RegistroAtual
  End With
  'Limpa objeto na memória.
  Set oProgressBar = Nothing
End Sub
```

***

##### OBSERVAÇÃO
Quando um objeto é destruído, ou seja, é instanciado um valor ___Nothing___ ao objeto, o Userform que foi criado _Em Tempo de Execução_, será eliminado do __Projeto VBA__.
