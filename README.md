# QlikSend
## Aplicação de Envio de Relatórios por Email Qlikview

Essa aplicação tem o intuito de permirtir gerar em pdf todos os relatórios que estão dentro de um arquivo **.qvw**. 
Na aplicação Web, o usuário poderá os emails que irão receber cada relatório, além período que deseja receber o email:
  - D: Diariamente
  - S: Semanalmente
  - Q: Quinzenalmente
  - M: Mensalmente

Essa Aplicação foi desenvolvida utilizando:
  - HTML, CSS, JS (Front End)
  - PHP (Back End)
  - Visual Basic (Back End. Linguagem usada para manipular o Qlikview)
## Para funcionar
  - Para o funcionamento dessa aplicação é necessário ter um XAMPP instalado.
  - Ter o Qlikview instalado.
## Configurar
  - Cada aplicação sua que irá gerar os relatórios em pdf deverá ter um arquivo **.vbs** com o mesmo nome:
    ###### Exemplo: "File.qvw" ----> "File.vbs"
  - Em cada "File.vbs" configure os diretórios descritos nas linhas 5 a 8.
    ###### Exemplo: Veja o arquivo "Visualization.vbs"
  - No arquivo "fsendEmail.vbs", das linhas 78 a 91, é necessário configurar o servidor de email da sua empresa.
    ###### Informações como Ip do Servido SMNP, Porta, Email e Senha do email que será usado para ser o remetente.
  - Exit um arquivo chamado "execute.bat", crie uma tarefa em algum Task Scheduler de sua preferencia, e peça para executar o **.bat** uma vez ao dia. Esse bat vai executar cada arquivo **.vbs** automaticamente.
  
## Detalhe
Exite em alguns casos a necessidade de que cada Relatório do Qlikview seja gerado com um filtro específico.
Caso haja essa necessidade coloque os filtros no campo "Comentários" do Relatórios do Qlik. 
O Script **.vbs** tem uma tarefa que ler esse campo e transforma a **String** em um texto executável, como uma **Macro**. O padrão para seguir é esse:
  ###### document.ClearAll true 'Limpa todos filtros
  ###### document.Fields("Ano").Select document.Evaluate("(Year(Today()))") 'Exemplo filtrando o campo Ano
Isso faz com que cada relatório possa tem um filtro diferente na hora de gerar.
  - Note que é como uma Macro no Qlikview, porém o **ActiveDocument** que é alterado para **document**. Isso acontece porque no código em VB, o _ActiveDocument_ vira um objeto chamado _document_. Então se desejar fazer filtros mais complexo, a sintaxe é a mesma da Macro.

