function ContratcFunction(e) {
  
    var Timestamp = e.values[0];
    var CabecalhoUm = e.values[1];
    var Nome = e.values[3]
    var Nacionalidade = e.values[3];
    var Profissao = e.values[4];
    var Rg = e.values[5]
    var Cpf = e.values[6];
    var Email = e.values[7];
    var EstadoCivil = e.values[8];
    var NomeConj = e.values[10];
    var Nacionalidadeconj = e.values[11];
    var CabecalhoDois = e.values[12];
    var ProfissaoConj = e.values[13];
    var RGConj = e.values[14];
    var CPFConj = e.values[15];
    var EmailConj = e.values[16];
    var Rua = e.values[17];
    var Bairro = e.values[18];
    var Cidade = e.values[19];
    var UF = e.values[20];
    var CEP = e.values[21];
    var ComiC1 = e.values[22];
    var ComiC1Text = e.values[23];
    var ComiC2 = e.values[24];
    var ComiC2Text = e.values[25];
    var ComiC3 = e.values[27];
    var ComiC3Text = e.values[28];
    var DataPagC1 = e.values[30];
    var ContaC1 = e.values[31];
    var ContaC3 = e.values[34];
    var DataAssContrato = e.values[35];
    var AssinaturaUm = e.values[36];
    var AssinaturaDois = e.values[37];
    var AssinaturaTres = e.values[38];
  
    var TemplateDocu = DriveApp.getFileById("2nk6RPxv...7b6J329gWedM");
    var PastaContratos = DriveApp.getFolderById("1m1S5V...qQ42h7Ow");
  
    var Copi = TemplateDocu.makeCopy(Nome, PastaContratos); 
        var Documento = DocumentApp.openById(Copi.getId());
        var body = Documento.getBody();
        var foooter = Documento.getFooter();
  
  
      body.replaceText("{{cabecalhocorretorvendedor1}}",CabecalhoUm);
      body.replaceText("{{cabecalhocorretorvendedor2}}",CabecalhoDois);
      body.replaceText("{{nomeclientevendedor}}",Nome);
      body.replaceText("{{naccliente}}",Nacionalidade);
      body.replaceText("{{profissaocliente}}",Profissao);
      body.replaceText("{{rgcliente}}",Rg);
      body.replaceText("{{cpfvendedor}}",Cpf);
      body.replaceText("{{emailcliente}}",Email);
      body.replaceText("{{estadocivilcliente}}",EstadoCivil);
      body.replaceText("{{nomeconje}}",NomeConj );
      body.replaceText("{{nacconje}}",Nacionalidadeconj);
      body.replaceText("{{profissaoconje}}",ProfissaoConj);
      body.replaceText("{{rgconje}}",RGConj);
      body.replaceText("{{cpfconje}}",CPFConj);
      body.replaceText("{{emailconje}}",EmailConj);
      body.replaceText("{{ruacliente}}",Rua);
      body.replaceText("{{bairrocliente}}",Bairro);
      body.replaceText("{{cidadecliente}}", Cidade);
      body.replaceText("{{ufcliente}}",UF);
      body.replaceText("{{cepcliente}}", CEP);
      body.replaceText("{{comiimobi}}",ComiC1);
      body.replaceText("{{comiimobitext}}",ComiC1Text);
      body.replaceText("{{datapgtocomi}}",DataPagC1);
      body.replaceText("{{comicorretor1}}",ComiC2);
      body.replaceText("{{comicorretor1text}}",ComiC2Text);
      body.replaceText("{{datapgtocomi}}",DataPagC1);
      body.replaceText("{{ctabancáriacorretor1}}",ContaC1);
      body.replaceText("{{comicorretor2}}",ComiC3);
      body.replaceText("{{comicorretor2text}}",ComiC3Text);
      body.replaceText("{{datapgtocomi}}",DataPagC1);
      body.replaceText("{{ctabancáriacorretor2}}",ContaC3);
      body.replaceText("{{datacontrato}}",DataAssContrato);
      body.replaceText("{{asscorretor1}}",AssinaturaUm);
      body.replaceText("{{asscorretor2}}",AssinaturaDois);
      body.replaceText("{{asscorretor3}}",AssinaturaTres);
  
      Documento.saveAndClose();
  }