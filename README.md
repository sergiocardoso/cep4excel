cep4excel
=========

Eu não sei se você já passou por esse problema, mas aqui no escritório por trabalharmos com muitas planilhas do Excel estamos sempre precisando limpar elas, ou seja, a maioria vem apenas com o campo CEP, sem o endereço, bairro, cidade (e muitas vezes até sem o número da residência). Por causa disso eu desenvolvi um macro em Excel que permite que informando apenas o cep o Excel irá retornar todas as informações necessárias.

Veja como fazer:

Faça o download da macro clicando neste link e descompacte.<br/>
Abra o excel e clique na aba ***“Desenvolvedor”***.<br/>
Depois clique no botão “Visual Basic”.<br/>
![](http://www.sergiocardoso.org/wp-content/uploads/2012/11/excel2.gif)

No menu do Visual Basic, clique em Arquivo e depois em Importar.
Selecione o arquivo “cep.bas” e clique em “Abrir”.
Agora um detalhe importante, é necessário que o complemento que lê arquivos XML esteja habilitado, para isso clique em Ferramentas e depois em Referências.
Depois confira se o complemento “Microsoft XML” está habilitado, se não, desça a janela e habilite-o. (Eu estou usando a versão 3, mas acredito que também funciona com versões superiores).
Pronto, o arquivo já esta configurado. Feche a janela do VBA e volte ao Excel. (este arquivo habilitará a formula “CEP” na sua biblioteca de formulas do Excel).
Para utilizar é bem simples, crie uma nova formula e digite:

Pronto, repare que na célula que você iniciou a formula neste exemplo, o macro irá retornar o endereço da Av. Paulista que obviamente foi o cep que eu utilizei.

Então reforçando, a sintax de uso da formula é CEP(“CEP-DESEJADO”;”TIPO DE CAMPO A RETORNAR”) aonde o tipo do campo pode ser logradouro, bairro, cidade ou uf.

Espero que esse macro possa ajudas as pessoas que as vezes se veem perdida em tantos campos no Excel, principalmente quando várias pessoas fazer o cadastramento dessas planilhas de formas diferentes (umas digitam o nome Rua, outras apenas R.), com essa formula você poderá organizar melhor e de uma forma uniforme todos os seus cadastros.

IS US!


