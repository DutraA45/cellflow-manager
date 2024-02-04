function FormPedido(Id) {

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaCliente = planilha.getSheetByName("Clientes");
  var guiaProduto = planilha.getSheetByName("Inventario");

  var ultimaLinha = guiaCliente.getLastRow();

  if(ultimaLinha == 0){
    var ultimaLinha = 1;
  }

  var list = guiaCliente.getRange(2, 2, ultimaLinha, 1).getValues();

  list.sort();

  // Removendo código desnecessário para lista 3
  /*
  var ultimaLinha = guiaProduto.getLastRow() - 1;

  if(ultimaLinha == 0){
    var ultimaLinha = 1;
  }

  var dadosLinhas = guiaProduto.getRange(2,2, ultimaLinha, 1).getValues();

  var listaUnica = [...new Set(dadosLinhas.flat())];

  var listaLinhas = [];

  for(var i = 0; i < listaUnica.length; i++){
    listaLinhas.push([listaUnica[i]]);
  }  

  var list3 = listaLinhas.sort();
  */

  // Lista 3 fixa como "Reparo" e "Venda"
  var list3 = [["Reparo"], ["Venda"]];

  // Removendo código desnecessário para lista 4
  /*
  var ultimaLinha = guiaProduto.getLastRow() - 1;

  if(ultimaLinha == 0){
    var ultimaLinha = 1;
  }

  var dadosLinhas = guiaProduto.getRange(2,3, ultimaLinha, 1).getValues();

  var listaUnica = [...new Set(dadosLinhas.flat())];

  var listaLinhas = [];

  for(var i = 0; i < listaUnica.length; i++){
    listaLinhas.push([listaUnica[i]]);
  }  

  var list4 = listaLinhas.sort();
  */

  // Lista 4 fixa como "Computador", "Celular" e "Outros"
  var list4 = [["Computador"], ["Celular"], ["Outros"]];

  var Form = HtmlService.createTemplateFromFile("FormPedido");

  Form.list = list.map(function(r){
    return r[0];
  });

  Form.list3 = list3.map(function(r){
    return r[0];
  });

  Form.list4 = list4.map(function(r){
    return r[0];
  });

  Form.Id = Id;

  var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);

  MostrarForm.setTitle("Cadastro de Serviços").setHeight(550).setWidth(800);

  SpreadsheetApp.getUi().showModalDialog(MostrarForm, "Cadastro de Serviços");
  
}


function buscaListas(){

  var planilha = SpreadsheetApp.getActiveSpreadsheet();

  var guiaClientes = planilha.getSheetByName("Clientes");

  var ultimaLinha = guiaClientes.getLastRow();

  var dadosClientes = guiaClientes.getRange(2, 2, ultimaLinha, 1).getValues();

  var guiaProdutos = planilha.getSheetByName("Inventario");

  var ultimaLinha = guiaProdutos.getLastRow();

  var dadosProdutos = guiaProdutos.getRange(2, 1, ultimaLinha, 10).getValues();

  var arrays = {
    dadosClientes: dadosClientes,
    dadosProdutos: dadosProdutos,
  }
  
   return arrays;

}

function buscaPedidoId(){

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaPedido = planilha.getSheetByName("Cadastro_Servico");

  var novoPedido = Math.max.apply(null, guiaPedido.getRange("X2:X").getValues());
  var novoPedido = novoPedido + 1;

  var novoId = Math.max.apply(null, guiaPedido.getRange("A2:A").getValues());
  var novoId = novoId + 1;

  var dados = {
    novoPedido: novoPedido,
    novoId: novoId,
  }

  return dados;

}

function SalvarPedido(Dados){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){

    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var guiaPedido = planilha.getSheetByName("Cadastro_Servico");
    var guiaCliente = planilha.getSheetByName("Clientes");

    var linha = guiaPedido.getLastRow() + 1;


    // PESQUISA DO TELEFONE

    guiaPedido.getRange(linha, 3).setValue(Dados.Cliente); // Defina o cliente na coluna C

    // Pesquise o cliente na planilha "Clientes"
    var clienteRange = guiaCliente.getRange("B:B");
    var clienteValues = clienteRange.getValues();
    var clienteIndex = clienteValues.findIndex(function (row) {
      return row[0] === Dados.Cliente;
    });

    if (clienteIndex !== -1) {
      // Se o cliente for encontrado, obtenha o valor da coluna D
      var valorColunaD = guiaCliente.getRange(clienteIndex + 1, 4).getValue();
      // Defina o valor encontrado na coluna E do "Cadastro_Servico"
      guiaPedido.getRange(linha, 5).setValue(valorColunaD);
    } else {
      // Caso o cliente não seja encontrado, você pode tratar isso como achar adequado
      guiaPedido.getRange(linha, 5).setValue("Cliente não encontrado");
    }




    var dataQuebrada = Dados.Data.split("/");

    var Ano = dataQuebrada[0];
    var Mes = dataQuebrada[1];
    var Dia = dataQuebrada[2];

    var Data = Dia + "/" + Mes + "/" + Ano;

    guiaPedido.getRange(linha,2).setValue(Data);

    var data = new Date(Dados.Data);
    var m = data.getMonth();

    var meses = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"];

    var Mes = meses[m];

    guiaPedido.getRange(linha,1).setValue(Dados.Id);
    guiaPedido.getRange(linha,24).setValue(Dados.Pedido);
    //guiaPedido.getRange(linha,4).setValue(Mes);
    //guiaPedido.getRange(linha,5).setValue(Ano);
    guiaPedido.getRange(linha,11).setValue(Dados.Linha);
    guiaPedido.getRange(linha,7).setValue(Dados.Produto);
    guiaPedido.getRange(linha,14).setValue(Dados.Qtd);
    //guiaPedido.getRange(linha,15).setValue(Dados.Preco);
    guiaPedido.getRange(linha,19).setValue(Dados.Total);
    guiaPedido.getRange(linha,3).setValue(Dados.Cliente);
    guiaPedido.getRange(linha,16).setValue(Dados.Status);
    guiaPedido.getRange(linha,22).setValue(Dados.Obs);
    guiaPedido.getRange(linha,12).setValue(Dados.Descricao);
    guiaPedido.getRange(linha,15).setValue(Dados.Custo);

    return "REGISTRADO COM SUCESSO!";

  }

}

function PesquisarPedido(Id){

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaPedido = planilha.getSheetByName("Cadastro_Servico");

  var ultimaLinha = guiaPedido.getLastRow();

  var dados = guiaPedido.getRange(2, 1, ultimaLinha, 25).getValues();

  for(var i = 0; i < dados.length; i++){

    if(dados[i][0] == Id){

      var Pedido = dados[i][23];

      var Data = Utilities.formatDate(new Date(dados[i][1]), planilha.getSpreadsheetTimeZone(), "yyyy-MM-dd");

      var Linha = dados[i][10];
      var Produto = dados[i][6];
      var Qtd = dados[i][13];      


      var T = dados[i][18].toLocaleString({style:'decimal', decimal: 'pt-BR'});
      var Total = T.replace(/\./g,"");

      var C = dados[i][14].toLocaleString({style:'decimal', decimal: 'pt-BR'});
      var Custo = C.replace(/\./g,"");

      var Cliente = dados[i][2];
      var Status = dados[i][15];
      var Obs = dados[i][21];
      var Descricao = dados[i][11];

      // Calcular o preço (Total / Qtd)
      var Preco = (parseFloat(Total) / parseFloat(Qtd)).toFixed(2);

      dados.length = 0;

      return ([Pedido,Data,Linha,Produto,Qtd,Total,Custo ,Cliente,Status,Obs, Descricao, Preco]);

    }

  }
  
  dados.length = 0;
  return "ID NÃO ENCONTRADO!";

}

function EditarPedido(Dados){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaPedido = planilha.getSheetByName("Cadastro_Servico");

  var ultimaLinha = guiaPedido.getLastRow();

  var dados = guiaPedido.getRange(2, 1, ultimaLinha, 1).getValues();

  for(var i = 0; i < dados.length; i++){

    if(dados[i][0] == Dados.Id){

      var linha = i + 2;

      var dataQuebrada = Dados.Data.split("/");

      var Ano = dataQuebrada[0];
      var Mes = dataQuebrada[1];
      var Dia = dataQuebrada[2];

      var Data = Dia + "/" + Mes + "/" + Ano;

      guiaPedido.getRange(linha,2).setValue(Data);

      var data = new Date(Dados.Data);
      var m = data.getMonth();

      var meses = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"];

      var Mes = meses[m];

      guiaPedido.getRange(linha,1).setValue(Dados.Id);
      guiaPedido.getRange(linha,24).setValue(Dados.Pedido);
      //guiaPedido.getRange(linha,4).setValue(Mes);
      //guiaPedido.getRange(linha,5).setValue(Ano);
      guiaPedido.getRange(linha,11).setValue(Dados.Linha);
      guiaPedido.getRange(linha,7).setValue(Dados.Produto);
      guiaPedido.getRange(linha,14).setValue(Dados.Qtd);
      //guiaPedido.getRange(linha,15).setValue(Dados.Preco);
      guiaPedido.getRange(linha,19).setValue(Dados.Total);
      guiaPedido.getRange(linha,3).setValue(Dados.Cliente);
      guiaPedido.getRange(linha,16).setValue(Dados.Status);
      guiaPedido.getRange(linha,22).setValue(Dados.Obs);
      guiaPedido.getRange(linha,12).setValue(Dados.Descricao);
      guiaPedido.getRange(linha,15).setValue(Dados.Custo);

      dados.length = 0;
      return "EDITADO COM SUCESSO!";

    }

   }

   dados.length = 0;
   return "ID NÃO ENCONTRADO!";

  }

}

function ExcluirPedido(Id){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){

    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var guiaPedido = planilha.getSheetByName("Cadastro_Servico");

    var ultimaLinha = guiaPedido.getLastRow();

    var dados = guiaPedido.getRange(2, 1, ultimaLinha, 1).getValues();

    for(var i = 0; i < dados.length; i++){

      if(dados[i][0] == Id){

        var linha = i + 2;
        guiaPedido.deleteRow(linha);

        dados.length = 0;
        return "EXCLUÍDO COM SUCESSO!";

      }

    }

    dados.length = 0;
    return "ID NÃO ENCONTRADO!";

  }  

}
