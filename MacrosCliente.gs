function FormCliente(Cliente) {
  
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaCliente = planilha.getSheetByName("Clientes");
  var guiaEstado = planilha.getSheetByName("Estados");
  
  var ultimaLinha = guiaCliente.getLastRow() - 1;

  if(ultimaLinha == 0){
    ultimaLinha = 1;
  }

  var list = guiaCliente.getRange(2,2,ultimaLinha,1).getValues();

  list.sort();

  var ultimaLinha = guiaEstado.getLastRow() - 1;

  if(ultimaLinha == 0){
    ultimaLinha = 1;
  }

  var list2 = guiaEstado.getRange(2, 1, ultimaLinha, 1).getValues();

  var Form = HtmlService.createTemplateFromFile("FormCliente");

  Form.list = list.map(function(r){return r[0];});
  Form.list2 = list2.map(function(r){return r[0];}); 
  
  Form.Cliente = Cliente;

  var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);

  MostrarForm.setTitle("Cadastro de Clientes").setHeight(470).setWidth(700);

  SpreadsheetApp.getUi().showModalDialog(MostrarForm,"Cadastro de Clientes");

}


function SalvarCliente(Dados){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){

    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var guiaCliente = planilha.getSheetByName("Clientes");

    var ultimaLinha = guiaCliente.getLastRow();

    var dadosCliente = guiaCliente.getRange(2, 2, ultimaLinha, 1).getValues();

    for(var i = 0; i < dadosCliente.length; i++){

      if(dadosCliente[i][0] == Dados.Cliente){
          return "CLIENTE JÁ CADASTRADO!";
      }

    }

    dadosCliente.length = 0;

    var linha = ultimaLinha + 1;
    var data = new Date();

    guiaCliente.getRange(linha,1).setValue(data);
    guiaCliente.getRange(linha,2).setValue(Dados.Cliente);
    guiaCliente.getRange(linha,3).setValue(Dados.Cpf);
    guiaCliente.getRange(linha,4).setValue(Dados.Contato);
    guiaCliente.getRange(linha,5).setValue(Dados.Rua);
    guiaCliente.getRange(linha,6).setValue(Dados.Bairro);
    guiaCliente.getRange(linha,7).setValue(Dados.Cidade);
    guiaCliente.getRange(linha,8).setValue(Dados.Estado);
    guiaCliente.getRange(linha,9).setValue(Dados.Obs);

    guiaCliente.getRange("A:A").setNumberFormat("dd/MM/yyyy");

    var guiaCidade = planilha.getSheetByName("Estados/Cidades");

    var ultimaLinha = guiaCidade.getLastRow();

    var dadosCidade = guiaCidade.getRange(2, 1, ultimaLinha, 2).getValues();

    for(var i = 0; i < dadosCidade.length; i++){

      if(dadosCidade[i][0] == Dados.Estado && dadosCidade[i][1] == Dados.Cidade){
          var existeCidade = "SIM";
      }

    }

    if(existeCidade != "SIM"){
      var linha = ultimaLinha + 1;
      guiaCidade.getRange(linha,1).setValue(Dados.Estado);
      guiaCidade.getRange(linha,2).setValue(Dados.Cidade);
    }

    dadosCidade.length = 0;
    return "REGISTRADO COM SUCESSO!";

  }

}


function AtualizarListaClientes(){

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaCliente = planilha.getSheetByName("Clientes");

  var ultimaLinha = guiaCliente.getLastRow() - 1;

  var list = guiaCliente.getRange(2,2,ultimaLinha,1).getValues();

  return list.sort();

}


function PesquisarCliente(nomeCliente){

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaCliente = planilha.getSheetByName("Clientes");

  var ultimaLinha = guiaCliente.getLastRow();

  var dadosCliente = guiaCliente.getRange(2,2,ultimaLinha,8).getValues();

  for(var i = 0; i<dadosCliente.length; i++){

      if(dadosCliente[i][0] == nomeCliente){

        var Cliente = dadosCliente[i][0];
        var Cpf = dadosCliente[i][1];
        var Contato = dadosCliente[i][2];
        var Rua = dadosCliente[i][3];
        var Bairro = dadosCliente[i][4];
        var Cidade = dadosCliente[i][5];
        var Estado = dadosCliente[i][6];
        var Obs = dadosCliente[i][7];

        dadosCliente.length = 0;

        return ([Cliente,Cpf,Contato,Rua,Bairro,Cidade,Estado,Obs]);

      }

  }

  dadosCliente.length = 0;

  return "CLIENTE NÃO ENCONTRADO!";

}


function VerificarCliente(nomeCliente){

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaCliente = planilha.getSheetByName("Clientes");

  var ultimaLinha = guiaCliente.getLastRow();

  var dadosCliente = guiaCliente.getRange(2,2,ultimaLinha,1).getValues();

  for(var i = 0; i < dadosCliente.length; i++){

    if(dadosCliente[i][0] == nomeCliente){
        dadosCliente.length = 0;
        return "CLIENTE JÁ CADASTRADO!";
    }

  }

  dadosCliente.length = 0;
  return "CLIENTE NÃO CADASTRADO!";

}


function EditarCliente(Dados){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){

    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var guiaCliente = planilha.getSheetByName("Clientes");
    //var guiaPedido = planilha.getSheetByName("Pedidos");

    var ultimaLinha = guiaCliente.getLastRow();

    var dadosCliente = guiaCliente.getRange(2,2,ultimaLinha,1).getValues();

    //var ultimaLinha = guiaPedido.getLastRow();

    //var dadosPedidos = guiaPedido.getRange(2,11,ultimaLinha,1).getValues();

    // var ver = dadosPedidos.filter(function(value,i,arr){
    //   return Dados.nomeCliente == arr[i][0];
    // });

    for(var i = 0; i < dadosCliente.length; i++){

      if(dadosCliente[i][0] == Dados.nomeCliente){

        var linha = i + 2;

        guiaCliente.getRange(linha,3).setValue(Dados.Cpf);
        guiaCliente.getRange(linha,4).setValue(Dados.Contato);
        guiaCliente.getRange(linha,5).setValue(Dados.Rua);
        guiaCliente.getRange(linha,6).setValue(Dados.Bairro);
        guiaCliente.getRange(linha,7).setValue(Dados.Cidade);
        guiaCliente.getRange(linha,8).setValue(Dados.Estado);
        guiaCliente.getRange(linha,9).setValue(Dados.Obs);

        dadosCliente.length = 0;
        //dadosPedidos.length = 0;

        // if(ver.length > 0){
        //     return "EDITADO COM SUCESSO, EXCETO NOME DO CLIENTE. JÁ TEM PEDIDO!";
        //     }else{
        //     guiaCliente.getRange(linha,2).setValue(Dados.Cliente);
        //  }

        return "CLIENTE EDITADO COM SUCESSO!";

      }

    }

    dadosCliente.length = 0;
    //dadosPedidos.length = 0;

    return "CLIENTE NÃO ENCONTRADO!";

  }  

}

function ExcluirCliente(nomeCliente){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){

    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var guiaCliente = planilha.getSheetByName("Clientes");
    //var guiaPedido = planilha.getSheetByName("Pedidos");

    var ultimaLinha = guiaCliente.getLastRow();

    var dadosClientes = guiaCliente.getRange(2,2,ultimaLinha,1).getValues();

    //var ultimaLinha = guiaPedido.getLastRow();

    //var dadosPedidos = guiaPedido.getRange(2,11,ultimaLinha,1).getValues();

    // var ver = dadosPedidos.filter(function(value, i, arr){
    //   return nomeCliente == arr[i][0];
    // });

    // if(ver.length > 0){
    //   dadosClientes.length = 0;
    //   dadosPedidos.length = 0;
    //   ver.length = 0;
    //   return "CLIENTE NÃO PODE SER EXCLUÍDO. JÁ TEM LANÇAMENTO DE PEDIDO!";
    // }

    for(var i = 0; i < dadosClientes.length; i++){

      if(dadosClientes[i][0] == nomeCliente){

        var linha = i + 2;
        guiaCliente.deleteRow(linha);

        dadosClientes.length = 0;
        //dadosPedidos.length = 0;
        //ver.length = 0;

        return "EXCLUÍDO COM SUCESSO!";

      }

    }

  dadosClientes.length = 0;
  //dadosPedidos.length = 0;
  //ver.length = 0;

  return "CLIENTE NÃO ENCONTRADO!";

  } 

}
