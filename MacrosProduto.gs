function FormProdutos() {

var planilha = SpreadsheetApp.getActiveSpreadsheet();
var guiaProduto = planilha.getSheetByName("Inventario");
var ultimaLinha = guiaProduto.getLastRow();

var dadosLinhas = guiaProduto.getRange(2,2,ultimaLinha,1).getValues();

var b = {};

for(var i = 0; i < dadosLinhas.length; i++){
  b[dadosLinhas[i][0]] = dadosLinhas[i][0];
}

var listaUnica = [];

for(var key in b){
  listaUnica.push([key]);
}

dadosLinhas.length = 0;

var list = listaUnica;

list.sort();

var Form = HtmlService.createTemplateFromFile("FormProduto");

Form.list = list.map(function(r){return r[0];});

var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);

MostrarForm.setTitle("Cadastro de Produtos").setHeight(490).setWidth(510);

SpreadsheetApp.getUi().showModalDialog(MostrarForm,"Cadastro de Produtos");

}

function Chamar(Arquivo) {

  return HtmlService.createHtmlOutputFromFile(Arquivo).getContent();
  
}

function ListaProdutos(Linha){

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaProduto = planilha.getSheetByName("Inventario");

  var ultimaLinha = guiaProduto.getLastRow();

  var dadosProdutos = guiaProduto.getRange(2,2,ultimaLinha,2).getValues();

  var produtos = [];
   var produtosSet = new Set();  // Usando um conjunto para garantir valores únicos

   for (var i = 0; i < dadosProdutos.length; i++) {
      if (dadosProdutos[i][0] == Linha) {
         var produto = dadosProdutos[i][1];

         // Verifica se o produto já está no conjunto antes de adicionar
         if (!produtosSet.has(produto)) {
            produtos.push([produto]);
            produtosSet.add(produto);
         }
      }
   }

   dadosProdutos.length = 0;
   return produtos;

}


function SalvarProduto(Dados){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){

    var Tipo = Dados.Tipo;
    var Produto = Dados.Produto;
    var Preco = Dados.Preco;
    var Quantidade = Dados.Quantidade;
    var Descricao = Dados.Descricao;

    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var guiaProduto = planilha.getSheetByName("Inventario");

    var ultimaLinha = guiaProduto.getLastRow();

    var dadosProdutos = guiaProduto.getRange(2,3,ultimaLinha,2).getValues();

    for(var i = 0; i<dadosProdutos.length; i++){

      if(dadosProdutos[i][1] == Produto){
          return "PRODUTO JÁ CADASTRADO!";
      }

    }

    var linha = guiaProduto.getLastRow() + 1;

    guiaProduto.getRange(linha,2).setValue(Tipo);
    guiaProduto.getRange(linha,3).setValue(Produto);
    guiaProduto.getRange(linha,6).setValue(Descricao);
    guiaProduto.getRange(linha,9).setValue(Preco);
    guiaProduto.getRange(linha,11).setValue(Quantidade);

     return "REGISTRADO COM SUCESSO!";

  }

}

function AtualizarListaLinhas(){

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaProduto = planilha.getSheetByName("Inventario");

  var ultimaLinha = guiaProduto.getLastRow();

  var dadosLinhas = guiaProduto.getRange(2,2,ultimaLinha,1).getValues();

  var b = {};

  for(var i = 0; i < dadosLinhas.length; i++){
    b[dadosLinhas[i][0]] = dadosLinhas[i][0];
  }

  var listaUnica = [];

  for(var key in b){
    listaUnica.push([key]);
  }

  dadosLinhas.length = 0;

  return listaUnica.sort();

}

function PesquisarProduto(Dados){

    var Linha = Dados.Linha;
    var Produto = Dados.Produto;

    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var guiaProduto = planilha.getSheetByName("Inventario");

    var ultimaLinha = guiaProduto.getLastRow();

    var dadosProdutos = guiaProduto.getRange(2,2,ultimaLinha,10).getValues();

    for(var i = 0; i <dadosProdutos.length; i++){

      if(dadosProdutos[i][0] == Linha && dadosProdutos[i][1] == Produto){

        var Descricao = dadosProdutos[i][4].toLocaleString({style: 'decimal',decimal: 'pt-BR'});
        var Descricao = Descricao.replace(/\./g,"");

        var Preco = dadosProdutos[i][7].toLocaleString({style: 'decimal',decimal: 'pt-BR'});
        var Preco = Preco.replace(/\./g,"");

        var Quantidade = dadosProdutos[i][9].toLocaleString({style: 'decimal',decimal: 'pt-BR'});
        var Quantidade = Quantidade.replace(/\./g,"");        
        
        dadosProdutos.length = 0;

        
        return {
                Descricao: Descricao,
                Preco: Preco,
                Quantidade: Quantidade
            };

      }

    }
   
    dadosProdutos.length = 0;    
    return "NÃO ENCONTRADO!";

}

function EditarProduto(Dados){

const user = LockService.getScriptLock();
user.tryLock(10000);

if(user.hasLock()){

  var LinhaLista = Dados.LinhaLista;
  var ProdutoLista = Dados.ProdutoLista;

  var Tipo = Dados.Tipo;
  var Produto = Dados.Produto;
  var Preco = Dados.Preco;
  var Quantidade = Dados.Quantidade;
  var Descricao = Dados.Descricao;

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaProduto = planilha.getSheetByName("Inventario");

  var ultimaLinha = guiaProduto.getLastRow();

  var dadosProdutos = guiaProduto.getRange(2,2,ultimaLinha,2).getValues();

  for(var i = 0; i <dadosProdutos.length; i++){

    if(dadosProdutos[i][0] == LinhaLista && dadosProdutos[i][1] == ProdutoLista ){

        var linha = i + 2;

          guiaProduto.getRange(linha,2).setValue(Tipo);
          guiaProduto.getRange(linha,3).setValue(Produto);
          guiaProduto.getRange(linha,6).setValue(Descricao);
          guiaProduto.getRange(linha,9).setValue(Preco);
          guiaProduto.getRange(linha,11).setValue(Quantidade);

          dadosProdutos.length = 0;

          return "EDITADO COM SUCESSO!";

    }

  }

dadosProdutos.length = 0;

return "PRODUTO NÃO ENCONTRADO!";

}

}


function ExcluirProduto(Dados){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){

    var Linha = Dados.LinhaLista;
    var Produto = Dados.ProdutoLista;

    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var guiaProduto = planilha.getSheetByName("Inventario");

    var ultimaLinha = guiaProduto.getLastRow();

    var dadosProdutos = guiaProduto.getRange(2,1,ultimaLinha,9).getValues();
   

    for(var i = 0; i < dadosProdutos.length; i++){

      if(dadosProdutos[i][1] == Linha && dadosProdutos[i][2] == Produto){

        var linha = i + 2;
        guiaProduto.deleteRow(linha);

        dadosProdutos.length = 0;

        return "EXCLUÍDO COM SUCESSO!";

      }

    }

    dadosProdutos.length = 0;

    return "PRODUTO NÃO ENCONTRADO!";

  }

}
