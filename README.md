# cProgressBar V1.0.0

O __Custom Progress Bar__, é uma classe desenvolvida para simular uma __Barra de Progresso Personalizada__. Essa classe pode ser utilizada em aplicações desenvolvidas com __Microsoft Excel__ e __Visual Basic for Application__ e consiste na construção de um ___Useform___, ___Em Tempo de Execução___, que simula uma __ProgressBar__ da biblioteca __Microsoft Common Control (ListView, TreeView, ProgressBar, StatusBar, etc)__.

A mesma é considerada personalizada, pois é possível realizar alterações na estrutura da mesma, da seguinte maneira:
- __Contador__: pode assumir um valor Percentual, ou Quantitativo, da evolução dos processos;
- __Barra__: pode assumir uma forma de Barra Crescente, ou Display de Informações, com os valores a serem processados.

Veremos o detalhamento dos __Métodos__ e __Propriedades__ disponibilizados na ___cProgressBar___.

### Métodos e Propriedades Classe

Através da criação destas ___Propriedades___ e ___Métodos___, tivemos o intuito de que a mesma funcionasse como um ___Framework de Projetos VBA___, onde a mesma possa ser chamada em qualquer momento do projeto, seja em um _Módulo_ ou um _Formulário de Pesquisa_. 
Ressaltando que os nomes das ___Propriedades___ e ___Métodos___ são em Inglês, para seguir o padrão da língua estrangeira, utilizada como base para a __Programação VBA__.

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

![alt text](https://raw.githubusercontent.com/username/projectname/branch/path/to/img.png)
