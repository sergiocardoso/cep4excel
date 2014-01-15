cep4excel
=========

Esta macro em Excel permite que informando o cep retorne todas as informações necessárias *(rua, bairro, logradouro, cidade, estado)*.

Veja como fazer:

Faça o download da macro.<br/><br/>
Abra o excel e clique na aba **“Desenvolvedor”**.<br/><br/>
Depois clique no botão **“Visual Basic”**.<br/><br/>
![](http://www.sergiocardoso.org/wp-content/uploads/2012/11/excel2.gif)<br/><br/>

No menu do Visual Basic, clique em **Arquivo** e depois em **Importar**.<br/>
![](http://www.sergiocardoso.org/wp-content/uploads/2012/11/excel3.gif)<br/>

Selecione o arquivo “cep.bas” e clique em “Abrir”.<br/>
Agora um detalhe importante, é necessário que o **complemento que lê arquivos XML esteja habilitado**, para isso clique em Ferramentas e depois em Referências.<br/>
![](http://www.sergiocardoso.org/wp-content/uploads/2012/11/excel4.gif)<br/>
Depois confira se o complemento “Microsoft XML” está habilitado, se não, desça a janela e habilite-o. (Eu estou usando a versão 3, mas acredito que também funciona com versões superiores).
![](http://www.sergiocardoso.org/wp-content/uploads/2012/11/excel5.gif)<br/>
Pronto, o arquivo já esta configurado. Feche a janela do VBA e volte ao Excel. *(este arquivo habilitará a formula “CEP” na sua biblioteca de formulas do Excel)*.
Para utilizar é bem simples, crie uma nova formula e digite:
![](http://www.sergiocardoso.org/wp-content/uploads/2012/11/excel6.gif)<br/>

Pronto, repare que na célula que você iniciou a formula neste exemplo, o macro irá retornar o endereço da Av. Paulista que obviamente foi o cep que eu utilizei.<br/>

Então reforçando, a sintax de uso da formula é **CEP(“CEP-DESEJADO”;”TIPO DE CAMPO A RETORNAR”) **aonde o tipo do campo pode ser *logradouro, bairro, cidade ou uf*.<br/>

Espero que esse macro possa ajudas as pessoas que as vezes se veem perdida em tantos campos no Excel, principalmente quando várias pessoas fazer o cadastramento dessas planilhas de formas diferentes (umas digitam o nome Rua, outras apenas R.), com essa formula você poderá organizar melhor e de uma forma uniforme todos os seus cadastros.
<br/>
IS US!<br/>


