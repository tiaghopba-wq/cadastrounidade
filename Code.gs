function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.openById('1VPh5iIM4nvT8fMDNelnlRPRS_maVxg6uuYsMrbAJXVg');

    if (data.tipo_cadastro === 'fonte_pagadora') {
      salvarFontePagadora(ss, data);
    } else if (data.tipo_cadastro === 'usuario') {
      salvarUsuario(ss, data);
    } else {
      salvarUnidadeColeta(ss, data);
    }
    return ContentService.createTextOutput(JSON.stringify({ result: 'success' })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ result: 'error', error: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) sheet.appendRow(headers);
  return sheet;
}

function multi(d, name) {
  var v = d[name];
  if (!v) return '';
  return Array.isArray(v) ? v.join(', ') : v;
}

function salvarArquivo(fileObj, pastaNome) {
  if (!fileObj || !fileObj.base64) return '';
  var pastas = DriveApp.getFoldersByName(pastaNome);
  var pasta = pastas.hasNext() ? pastas.next() : DriveApp.createFolder(pastaNome);
  var bytes = Utilities.base64Decode(fileObj.base64);
  var blob = Utilities.newBlob(bytes, fileObj.mimeType, fileObj.filename);
  var arquivo = pasta.createFile(blob);
  // Por padrão o arquivo fica privado (só você acessa). Se quiser um link público
  // para usar em laudos, descomente a linha abaixo (isso torna o arquivo acessível
  // por qualquer pessoa com o link):
  // arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return arquivo.getUrl();
}

function salvarUnidadeColeta(ss, d) {
  var headers = [
    'Timestamp','Grupo/Empresa','Descrição (Nome da Unidade)','Regional',
    'C.N.E.S.','Responsável técnico','Obrigatoriedade',
    'CEP','Logradouro','Complemento','Número','Bairro','Cidade','Estado','Telefone',
    'Segmento','Categoria','Atendimento',
    'Horário Segunda (Abre)','Horário Segunda (Intervalo Início)','Horário Segunda (Intervalo Fim)','Horário Segunda (Fecha)',
    'Horário Terça (Abre)','Horário Terça (Intervalo Início)','Horário Terça (Intervalo Fim)','Horário Terça (Fecha)',
    'Horário Quarta (Abre)','Horário Quarta (Intervalo Início)','Horário Quarta (Intervalo Fim)','Horário Quarta (Fecha)',
    'Horário Quinta (Abre)','Horário Quinta (Intervalo Início)','Horário Quinta (Intervalo Fim)','Horário Quinta (Fecha)',
    'Horário Sexta (Abre)','Horário Sexta (Intervalo Início)','Horário Sexta (Intervalo Fim)','Horário Sexta (Fecha)',
    'Horário Sábado (Abre)','Horário Sábado (Intervalo Início)','Horário Sábado (Intervalo Fim)','Horário Sábado (Fecha)',
    'Horário Domingo (Abre)','Horário Domingo (Intervalo Início)','Horário Domingo (Intervalo Fim)','Horário Domingo (Fecha)',
    'Dias de produção dos exames','Dias de entrega dos resultados',
    'Empresa prestadora','Percentual máximo desconto','Imprime resultado em débito',
    'Tempo até a unidade de processamento (min)',
    'Usará template com logomarca','Logomarca do template (link)'
  ];
  var sheet = getOrCreateSheet(ss, 'UnidadeColeta', headers);

  var linkTemplateLogo = salvarArquivo(d.template_logo, 'Templates Unidade de Coleta');

  var row = [
    new Date(), d.grupo_empresa || '', d.descricao || '', d.regional || '',
    d.cnes || '', d.responsavel_tecnico || '', multi(d, 'obrigatoriedade'),
    d.cep || '', d.logradouro || '', d.complemento || '', d.numero_endereco || '', d.bairro || '', d.cidade || '', d.estado || '', d.telefone || '',
    multi(d, 'segmento'), multi(d, 'categoria'), d.atendimento || '',
    d.horario_seg_inicio || '', d.horario_seg_intervalo_inicio || '', d.horario_seg_intervalo_fim || '', d.horario_seg_fim || '',
    d.horario_ter_inicio || '', d.horario_ter_intervalo_inicio || '', d.horario_ter_intervalo_fim || '', d.horario_ter_fim || '',
    d.horario_qua_inicio || '', d.horario_qua_intervalo_inicio || '', d.horario_qua_intervalo_fim || '', d.horario_qua_fim || '',
    d.horario_qui_inicio || '', d.horario_qui_intervalo_inicio || '', d.horario_qui_intervalo_fim || '', d.horario_qui_fim || '',
    d.horario_sex_inicio || '', d.horario_sex_intervalo_inicio || '', d.horario_sex_intervalo_fim || '', d.horario_sex_fim || '',
    d.horario_sab_inicio || '', d.horario_sab_intervalo_inicio || '', d.horario_sab_intervalo_fim || '', d.horario_sab_fim || '',
    d.horario_dom_inicio || '', d.horario_dom_intervalo_inicio || '', d.horario_dom_intervalo_fim || '', d.horario_dom_fim || '',
    multi(d, 'dias_producao'), multi(d, 'dias_entrega'),
    d.empresa_prestadora || '', d.percentual_max_desconto || '', d.imprime_resultado_debito ? 'Sim' : 'Não',
    d.tempo_transporte || '',
    d.usa_template_logomarca || '', linkTemplateLogo
  ];
  sheet.appendRow(row);
}

function salvarFontePagadora(ss, d) {
  var headers = [
    'Timestamp','Descrição','Razão social','Nome fantasia','Tipo','CNPJ','Inscrição estadual','Inscrição municipal','E-mail',
    'CEP','Logradouro','Complemento','Número','Bairro','Cidade','Estado','Telefone',
    'Contato responsável faturamento','Dia vencimento faturamento',
    'Logomarca cabeçalho (link)','Logomarca rodapé (link)','Logomarca assinatura (link)'
  ];
  var sheet = getOrCreateSheet(ss, 'FontePagadora', headers);

  var linkCabecalho = salvarArquivo(d.logo_cabecalho, 'Logomarcas Fonte Pagadora');
  var linkRodape = salvarArquivo(d.logo_rodape, 'Logomarcas Fonte Pagadora');
  var linkAssinatura = salvarArquivo(d.logo_assinatura, 'Logomarcas Fonte Pagadora');

  var row = [
    new Date(), d.descricao_fp || '', d.razao_social || '', d.nome_fantasia || '', d.tipo_fp || '',
    d.cnpj || '', d.inscricao_estadual || '', d.inscricao_municipal || '', d.email_fp || '',
    d.cep_fp || '', d.logradouro_fp || '', d.complemento_fp || '', d.numero_fp || '', d.bairro_fp || '', d.cidade_fp || '', d.estado_fp || '', d.telefone_fp || '',
    d.contato_responsavel_faturamento || '', d.dia_vencimento_faturamento || '',
    linkCabecalho, linkRodape, linkAssinatura
  ];
  sheet.appendRow(row);
}

function salvarUsuario(ss, d) {
  var headers = [
    'Timestamp','Nome do usuário','RG','CPF','Data nascimento','Sexo','E-mail',
    'CEP','Tipo endereço','Endereço','Número','Bairro','Cidade','Estado','País',
    'DDD','Telefone celular','Nro. conselho + estado','Local de trabalho','Cargo','Área de trabalho','Área de trabalho (outra)'
  ];
  var sheet = getOrCreateSheet(ss, 'Usuarios', headers);

  var row = [
    new Date(), d.nome_usuario || '', d.rg || '', d.cpf || '', d.data_nascimento || '', d.sexo || '', d.email_usuario || '',
    d.cep_usuario || '', d.tipo_endereco || '', d.endereco_usuario || '', d.numero_usuario || '', d.bairro_usuario || '', d.cidade_usuario || '', d.estado_usuario || '', d.pais || '',
    d.ddd || '', d.telefone_celular || '', d.conselho_estado || '', d.local_trabalho || '', d.cargo || '', multi(d, 'area_trabalho'), d.area_trabalho_outra || ''
  ];
  sheet.appendRow(row);
}
