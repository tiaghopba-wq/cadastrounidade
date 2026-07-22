/**
 * Apps Script (Web App) que recebe os 3 formulários de index.html e grava
 * cada um em sua própria aba na planilha de destino, com base em
 * payload.tipo_cadastro: "usuario" | "fonte_pagadora" | "unidade_coleta".
 *
 * Como usar:
 * 1. Abra a planilha Google que já recebe os dados (a apontada por SCRIPT_URL).
 * 2. Extensões > Apps Script.
 * 3. Substitua o conteúdo de Code.gs por este arquivo (ajuste FOLDER_ID se
 *    quiser guardar os uploads de logomarca no Drive).
 * 4. Implantar > Gerenciar implantações > editar a implantação existente
 *    (mantém a mesma URL que já está em index.html).
 */

// ID de uma pasta do Drive para salvar os uploads de logomarca (Fonte Pagadora).
// Deixe '' para pular o upload e gravar apenas o nome do arquivo.
var FOLDER_ID = '';

var CONFIG = {
  usuario: {
    aba: 'Usuários',
    campos: [
      'nome_usuario', 'rg', 'cpf', 'data_nascimento', 'sexo', 'email_usuario',
      'cep_usuario', 'tipo_endereco', 'endereco_usuario', 'numero_usuario',
      'bairro_usuario', 'cidade_usuario', 'estado_usuario', 'pais', 'ddd',
      'telefone_celular', 'conselho_estado', 'local_trabalho', 'cargo', 'area_trabalho'
    ],
    cabecalho: [
      'Nome do usuário', 'RG', 'CPF', 'Data nascimento', 'Sexo', 'E-mail',
      'CEP', 'Tipo endereço', 'Endereço', 'Número', 'Bairro', 'Cidade',
      'Estado', 'País', 'DDD', 'Telefone celular', 'Nro. conselho + estado',
      'Local de trabalho', 'Cargo', 'Área de trabalho'
    ]
  },
  fonte_pagadora: {
    aba: 'Fonte Pagadora',
    campos: [
      'descricao_fp', 'razao_social', 'nome_fantasia', 'tipo_fp', 'cnpj',
      'inscricao_estadual', 'inscricao_municipal', 'email_fp', 'cep_fp',
      'logradouro_fp', 'complemento_fp', 'numero_fp', 'bairro_fp', 'cidade_fp',
      'estado_fp', 'telefone_fp', 'contato_responsavel_faturamento',
      'dia_vencimento_faturamento', 'logo_cabecalho', 'logo_rodape', 'logo_assinatura'
    ],
    cabecalho: [
      'Descrição', 'Razão social', 'Nome fantasia', 'Tipo', 'CNPJ',
      'Inscrição estadual', 'Inscrição municipal', 'E-mail', 'CEP',
      'Logradouro', 'Complemento', 'Número', 'Bairro', 'Cidade', 'Estado',
      'Telefone', 'Contato faturamento', 'Dia vencimento faturamento',
      'Logo cabeçalho', 'Logo rodapé', 'Logo assinatura'
    ],
    arquivos: ['logo_cabecalho', 'logo_rodape', 'logo_assinatura']
  },
  unidade_coleta: {
    aba: 'Unidade de Coleta',
    campos: [
      'grupo_empresa', 'descricao', 'cartao_internet', 'regional', 'unidade',
      'maioridade_excecao', 'cnes', 'responsavel_tecnico', 'obrigatoriedade',
      'cep', 'logradouro', 'complemento', 'numero_endereco', 'bairro', 'cidade',
      'estado', 'telefone', 'segmento', 'categoria', 'atendimento',
      'inf_atendimento', 'permite_emergencia', 'permite_prazo_minimo',
      'endereco_hospitalar', 'pedido_externo', 'prontuario',
      'autorizacao_paciente_os', 'status_inicial_procedimento',
      'dia_producao_padrao', 'dia_entrega_padrao', 'empresa_prestadora',
      'percentual_max_desconto', 'imprime_resultado_debito',
      'transportar_amostras_em', 'fechar_malote', 'malote_agrupado_por',
      'unidade_destino_por', 'tempo_transporte', 'template_laudo',
      'template_elis', 'regras_impressao_acao'
    ],
    cabecalho: [
      'Grupo/Empresa', 'Descrição', 'Cartão internet', 'Regional', 'Unidade',
      'Maioridade exceção', 'C.N.E.S.', 'Responsável técnico', 'Obrigatoriedade',
      'CEP', 'Logradouro', 'Complemento', 'Número', 'Bairro', 'Cidade', 'Estado',
      'Telefone', 'Segmento', 'Categoria', 'Atendimento', 'Inf. atendimento',
      'Permite emergência', 'Permite prazo mínimo', 'Endereço hospitalar',
      'Pedido externo', 'Prontuário', 'Autorização paciente O.S.',
      'Status inicial procedimento', 'Dia produção padrão', 'Dia entrega padrão',
      'Empresa prestadora', '% máx desconto', 'Imprime resultado débito',
      'Transportar amostras em', 'Fechar malote', 'Malote agrupado por',
      'Unidade destino por', 'Tempo transporte', 'Template laudo',
      'Template e-LIS', 'Regras impressão'
    ]
  }
};

function doPost(e) {
  var payload = JSON.parse(e.postData.contents);
  var tipo = payload.tipo_cadastro;
  var config = CONFIG[tipo];
  if (!config) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, erro: 'tipo_cadastro inválido' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var aba = planilha.getSheetByName(config.aba);
  if (!aba) {
    aba = planilha.insertSheet(config.aba);
    aba.appendRow(['Data/Hora envio'].concat(config.cabecalho));
  }

  var linha = [new Date()];
  config.campos.forEach(function (campo) {
    var valor = payload[campo];

    if (config.arquivos && config.arquivos.indexOf(campo) !== -1) {
      linha.push(salvarArquivo(valor));
      return;
    }
    if (Array.isArray(valor)) {
      linha.push(valor.join(', '));
      return;
    }
    linha.push(valor === undefined || valor === null ? '' : valor);
  });

  aba.appendRow(linha);

  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function salvarArquivo(arquivo) {
  if (!arquivo || !arquivo.base64) return '';
  if (!FOLDER_ID) return arquivo.filename || '';

  var bytes = Utilities.base64Decode(arquivo.base64);
  var blob = Utilities.newBlob(bytes, arquivo.mimeType, arquivo.filename);
  var pasta = DriveApp.getFolderById(FOLDER_ID);
  var arquivoSalvo = pasta.createFile(blob);
  return arquivoSalvo.getUrl();
}
