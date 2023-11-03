function CallFunction(e) {
  ProcFunction(e);
  useDataRange(e);
}

function ProcFunction(e) {
  
    var Timestamp = e.values[0];
    var Nome = e.values[1];
    var Nacionalidade = e.values[2]
    var EstadoCivil = e.values[3];
    var Cpf = e.values[4];
    var Rg = e.values[5]
    var Profissao = e.values[6];
    var Rua = e.values[7];
    var Numero = e.values[8];
    var Bairro = e.values[9];
    var Cidade = e.values[10];
    var Cep = e.values[11];
    var Telefone = e.values[12];
    var Email = e.values[13];
    var Data = e.values[14];
    var TipoAcao = e.values[16];
    var Peticionante = e.values[18];
   
    var NeydianneKit = 'Dra....';

    var LaisKit = 'Dra...';

    var BrendaKit = 'Dra...';
    
    
    if (TipoAcao=='APOSENTADORIA POR INVALIDEZ' || TipoAcao=='AUX.DOENÇA' || TipoAcao=='AUX.ACIDENTE') {
      var urlTemplate = ("152EmgdmyTXtBSTaw5I2_5llLQewxrZhurY-cWrfKJ2w");
    }else if (TipoAcao=='CONTAGEM DE TEMPO' ||TipoAcao=='REVISÃO BENEFICIO' || TipoAcao=='CONCESSÃO     APOSENTADORIA') {
      var urlTemplate = ("1N92jNrB-2c3LTQk8");
    }else {
      var urlTemplate = ("1NKotqS9grYVTjLfqvIvxc");
    }Ws
    
    var TemplateDocu = DriveApp.getFileById(urlTemplate);
    var PastaDest = DriveApp.getFolderById("1ymBLodt1OMr3YmsgCtgFF");

    var Copi = TemplateDocu.makeCopy(Nome, PastaDest); 

    var Procuracao = DocumentApp.openById(Copi.getId());

    var body = Procuracao.getBody();

    var foooter = Procuracao.getFooter();

    body.replaceText("{{nome}}",Nome);
    body.replaceText("{{nac}}",Nacionalidade);
    body.replaceText("{{estadocivil}}",EstadoCivil);
    body.replaceText("{{cpf}}",Cpf);
    body.replaceText("{{rg}}",Rg);
    body.replaceText("{{profissao}}",Profissao);
    body.replaceText("{{rua}}",Rua);
    body.replaceText("{{numero}}",Numero);
    body.replaceText("{{bairro}}",Bairro);
    body.replaceText("{{cidade}}",Cidade);
    body.replaceText("{{cep}}",Cep);
    body.replaceText("{{telefone}}",Telefone);
    body.replaceText("{{email}}",Email);
    body.replaceText("{{data}}",Data);
    body.replaceText("{{tipoação}}", TipoAcao);


  if (Peticionante== 'NEYDIANNE') {
    body.replaceText("{{peticionante}}",NeydianneKit );
    } else if (Peticionante=='LAIS') {
      body.replaceText("{{peticionante}}",LaisKit );
    } else if (Peticionante=='BRENDA') {
      body.replaceText("{{peticionante}}",BrendaKit );
  }

Procuracao.saveAndClose();      
}

var url = "https://docs.google.com/spreadsheets"

function useDataRange (e) {

  var planilhabase = SpreadsheetApp.openByUrl(url);
  var guiadados = planilhabase.getSheetByName('DADOS')

  var planilhamenu =  SpreadsheetApp.getActiveSpreadsheet();
  var guiamenu = planilhamenu.getSheetByName('MENU');

  var ultimalinha = guiamenu.getLastRow();

  var area = "P:Q" + ultimalinha;
  var dados = guiamenu.getRange(area).getValues();
    
  var linhadados = guiadados.getLastRow() + 1;

  var areaData = "A" + ultimalinha
  var data = guiamenu.getRange(areaData).getValues();

  guiadados.getRange(linhadados, 1,dados.length, dados[0].length).setValues(dados);
  guiadados.getRange(linhadados, 3, data.length, data[0].length).setValues(data);

  guiamenu.getRange(area);
  guiamenu.getRange("A2").activate();

  guiamenu.getRange(areaData);
  guiamenu.getRange("A2").activate();

    Logger.log(areaData)

  Browser.msgBox('salvo com sucesso')

}

