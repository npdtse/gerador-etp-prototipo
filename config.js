/**
 * ===================================================================================
 * CONFIG.JS - ARQUIVO DE CONFIGURAÇÃO E CONSTANTES GLOBAIS
 * ===================================================================================
 * 
 * Este arquivo centraliza todas as variáveis de configuração, constantes, chaves de API
 * e mapeamentos de dados estáticos da aplicação Gerador de ETP.
 * 
 * Separar a configuração da lógica (script.js) torna o código mais limpo,
 * mais fácil de manter e permite alterar parâmetros importantes sem tocar
 * no fluxo principal da aplicação.
 * 
 */

// --- CHAVES DE API E ENDPOINTS ---

// Chave para armazenamento de dados no Local Storage do navegador.
// Alterar esta chave invalida os dados salvos anteriormente (útil para atualizações que quebram a compatibilidade).
const ETP_DATA_KEY = 'etpAppData_v100';

// Chave de API para o serviço de IA da Groq.
// ATENÇÃO: Esta é uma chave padrão para o funcionamento inicial. O usuário poderá substituir pela sua.
const GROQ_API_KEY = '';

// URL do endpoint da API de chat da Groq.
const GROQ_API_URL = 'https://api.groq.com/openai/v1/chat/completions';

// Lista de modelos de LLM padrão, recomendados pelo aplicativo.
// Esta lista será usada quando o usuário optar pela configuração "padrão" no modal de configurações de IA.
// A ordem representa a prioridade de uso (fallback).
const LLM_MODELS_PRIORITY = ["llama-3.3-70b-versatile", "llama-3.1-8b-instant"];


// --- MAPEAMENTO DE CAMPOS CONDICIONAIS ---

// Objeto que mapeia o ID de um campo de controle (como um <select> ou radio button)
// para um array de IDs dos elementos HTML que devem ser exibidos ou ocultados
// com base na seleção do usuário.
const conditionalFieldIds = {
    // Campos da Identificação
    etp_tipo: ['conditional_etp_simplificado_auth'],
    etp_auth: ['conditional_etp_auth_sei'],
    // Campos do Capítulo 1
    alinhamento_estrategico: ['conditional_c1_7_sim_alinhamento'],
    // Campos do Capítulo 3
    natureza_continua: ['conditional_c3_7_natureza_continua'],
    beneficios_vigencia: ['conditional_c3_8_1_beneficios_vigencia'],
    garantia: ['conditional_c3_9_1_prazo_garantia', 'conditional_c3_9_2_garantia_mercado'],
    justificativa_garantia: ['conditional_c3_9_2_1_justificativa_garantia'],
    assistencia: ['conditional_c3_10_1_prazo_assistencia', 'conditional_c3_10_2_assistencia_mercado'],
    justificativa_assistencia: ['conditional_c3_10_2_1_justificativa_assistencia'],
    transferencia: ['conditional_c3_11_transferencia'],
    capacitacao: ['conditional_c3_12_capacitacao'],
    contratacao_adicional: ['conditional_c3_13_contratacao_adicional'],
    ajuste_contratacoes: ['conditional_c3_14_ajuste_contratacoes'],
    acessibilidade: ['conditional_c3_15_acessibilidade'],
    // Campos do Capítulo 7
    parcelamento_grupo_unico: ['conditional_c7_grupo_unico'],
    parcelamento_grupos_separados: ['conditional_c7_grupos_separados'],
    // Campos do Capítulo 8
    contratacao_dispensa: ['conditional_c8_1_dispensa'],
    contratacao_inexigibilidade: ['conditional_c8_1_inexigibilidade'],
    dispensa_outra_hipotese: ['conditional_c8_1_1_outra_hipotese'],
    qualificacao_tecnica: ['conditional_c8_4_qualificacao_tecnica'],
    amostras: ['conditional_c8_5_amostras'],
    vistoria: ['conditional_c8_6_vistoria'],
    confidencialidade: ['conditional_c8_7_confidencialidade'],
    subcontratacao: ['conditional_c8_8_subcontratacao'],
    consorcio_proibicao: ['conditional_c8_9_proibicao_consorcio'],
    consorcio_permissao: ['conditional_c8_9_permisao_consorcio'],
    limite_consorcio: ['conditional_c8_9_2_detalhes_limite'],
    cooperativas_proibicao: ['conditional_c8_10_proibicao_cooperativas'],
    estrangeiras_proibicao: ['conditional_c8_11_proibicao_estrangeiras'],
    estrangeiras_permissao: ['conditional_c8_11_permisao_estrangeiras'],
    margem_preferencia: ['conditional_c8_11_2_detalhes_margem'],
    pessoa_fisica_proibicao: ['conditional_c8_12_proibicao_pessoa_fisica'],
    // Campos do Capítulo 9
    quantitativo_inferior: ['conditional_c9_1_quantitativo_inferior'],
    precos_diferentes: ['conditional_c9_2_precos_diferentes'],
    mais_de_um_fornecedor: ['conditional_c9_3_mais_de_um_fornecedor'],
    adesao_futura: ['conditional_c9_4_adesao_futura'],
    prorrogacao_ata: ['conditional_c9_5_prorrogacao_detalhes'],
    renovacao_quantidades_ata: ['conditional_c9_5_1_1_justificativa_renovacao'],
};

// --- CONFIGURAÇÃO DE ITENS DINÂMICOS ---

// Objeto que define os templates e comportamentos para cada tipo de item dinâmico
// que pode ser adicionado pelo usuário (Soluções, Riscos, Contratações, Anexos).
const dynamicItemConfigs = {
    solucao: {
        containerId: 'solucoes_mercado_container',
        itemClass: 'solucao-item',
        indexAttribute: 'data-solucao-index',
        titleSuffix: 'ª Solução',
        removeConfirmText: 'Tem certeza que deseja remover esta solução?',
        removeClass: 'remove-solucao-button',
        removeButtonText: 'Remover Solução',
        removeAction: 'remove-solucao',
        template: (i, showRemove) => {
            const removeBtn = showRemove ? `<button class="remove-solucao-button" data-action="remove-solucao"><i class="fas fa-trash-alt"></i> Remover Solução</button>` : '';
            return `<h3>${i}ª Solução ${removeBtn}</h3>` + [
                { id: 'a_descricao_sucinta', label: 'Título da solução', rows: 2 },
                { id: 'b_indicacao_servicos_materiais', label: 'Descrição resumida dos serviços e materiais', rows: 3 },
                { id: 'c_potenciais_fornecedores', label: 'Potenciais fornecedores ou fabricantes', rows: 2 },
                { id: 'd_orgaos_similares', label: 'Órgãos públicos ou entidades que tenham adotado solução similar', rows: 2 },
                { id: 'e_analise_contratos', label: 'Análise dos contratos identificados', rows: 3 },
                { id: 'f_servicos_materiais_complementares', label: 'Serviços e materiais não contemplados na solução', rows: 2 },
                { id: 'g_custos_estimados', label: 'Estimativa preliminar de custos', rows: 2 },
                { id: 'h_vantagens', label: 'Vantagens', rows: 3 },
                { id: 'i_desvantagens', label: 'Desvantagens', rows: 3 }
            ].map((f, idx) => `<div class="form-group"><label for="c2_sol${i}_${f.id}">2.${i}.${idx + 1} ${f.label}</label><div class="input-with-ai"><textarea id="c2_sol${i}_${f.id}" rows="${f.rows}"></textarea><button class="ai-help-button" data-action="get-ai-help" data-target-field="c2_sol${i}_${f.id}"><i class="fas fa-book-open"></i></button></div></div>`).join('');
        },
        renumberFields: (item, newIndex) => {
            const fieldSuffixes = ['a_descricao_sucinta', 'b_indicacao_servicos_materiais', 'c_potenciais_fornecedores', 'd_orgaos_similares', 'e_analise_contratos', 'f_servicos_materiais_complementares', 'g_custos_estimados', 'h_vantagens', 'i_desvantagens'];
            const labelsMap = [
                "Título da solução", "Descrição resumida dos serviços e materiais", "Potenciais fornecedores ou fabricantes",
                "Órgãos públicos ou entidades que tenham adotado solução similar", "Análise dos contratos identificados",
                "Serviços e materiais não contemplados na solução",
                "Estimativa preliminar de custos", "Vantagens", "Desvantagens"
            ];
            item.querySelectorAll('.form-group').forEach((group, index) => {
                const newId = `c2_sol${newIndex}_${fieldSuffixes[index]}`;
                group.querySelector('label').htmlFor = newId;
                group.querySelector('label').textContent = `2.${newIndex}.${index + 1} ${labelsMap[index]}`;
                group.querySelector('textarea').id = newId;
                // Atualiza o botão da IA para apontar para o novo ID específico
                group.querySelector('.ai-help-button').dataset.targetField = newId;
            });
        }
    },
    risco: {
        containerId: 'riscos_container',
        itemClass: 'risco-item',
        indexAttribute: 'data-risco-index',
        titleSuffix: 'º Risco',
        removeConfirmText: 'Tem certeza que deseja remover este risco?',
        removeClass: 'remove-risco-button',
        removeButtonText: 'Remover Risco',
        removeAction: 'remove-risco',
        template: (i, showRemove) => {
            const removeBtn = showRemove ? `<button class="remove-risco-button" data-action="remove-risco"><i class="fas fa-trash-alt"></i> Remover Risco</button>` : '';
            return `<h3>${i}º Risco ${removeBtn}</h3>
                <div class="form-group"><label for="c6_risco${i}_identificacao">6.${i}.1 Identificação do Risco</label><div class="input-with-ai"><textarea id="c6_risco${i}_identificacao" rows="3" placeholder="Descreva o risco. Ex: Atraso na entrega dos equipamentos."></textarea><button class="ai-help-button" data-action="get-ai-help" data-target-field="c6_risco${i}_identificacao"><i class="fas fa-book-open"></i></button></div></div>
                <div class="form-group"><label for="c6_risco${i}_danos_possiveis">6.${i}.2 Danos possíveis</label><div class="input-with-ai"><textarea id="c6_risco${i}_danos_possiveis" rows="3" placeholder="Descreva os danos caso o risco se concretize. Ex: Paralisação de atividades críticas da unidade."></textarea><button class="ai-help-button" data-action="get-ai-help" data-target-field="c6_risco${i}_danos_possiveis"><i class="fas fa-book-open"></i></button></div></div>
                <div class="form-group"><label for="c6_risco${i}_probabilidade">6.${i}.3 Probabilidade do risco</label><div class="input-with-ai"><select id="c6_risco${i}_probabilidade"><option value="">Selecione...</option><option value="Baixa">Baixa</option><option value="Média">Média</option><option value="Alta">Alta</option></select><button class="ai-help-button" data-action="get-ai-help" data-target-field="c6_risco${i}_probabilidade"><i class="fas fa-book-open"></i></button></div></div>
                <div class="form-group"><label for="c6_risco${i}_impacto">6.${i}.4 Impacto do risco</label><div class="input-with-ai"><select id="c6_risco${i}_impacto"><option value="">Selecione...</option><option value="Baixo">Baixo</option><option value="Médio">Médio</option><option value="Alto">Alto</option></select><button class="ai-help-button" data-action="get-ai-help" data-target-field="c6_risco${i}_impacto"><i class="fas fa-book-open"></i></button></div></div>
                <div class="form-group"><label for="c6_risco${i}_acoes_preventivas">6.${i}.5 Ações preventivas</label><div class="input-with-ai"><textarea id="c6_risco${i}_acoes_preventivas" rows="3" placeholder="Ações para evitar que o risco ocorra. Ex: Definir cronograma de entregas com margem de segurança."></textarea><button class="ai-help-button" data-action="get-ai-help" data-target-field="c6_risco${i}_acoes_preventivas"><i class="fas fa-book-open"></i></button></div></div>
                <div class="form-group"><label for="c6_risco${i}_acoes_contingencia">6.${i}.6 Ações de contingência</label><div class="input-with-ai"><textarea id="c6_risco${i}_acoes_contingencia" rows="3" placeholder="Ações a serem tomadas caso o risco se materialize. Ex: Notificar a empresa e aplicar sanções contratuais."></textarea><button class="ai-help-button" data-action="get-ai-help" data-target-field="c6_risco${i}_acoes_contingencia"><i class="fas fa-book-open"></i></button></div></div>`;
        },
        renumberFields: (item, newIndex) => {
            const fieldSuffixes = ['identificacao', 'danos_possiveis', 'probabilidade', 'impacto', 'acoes_preventivas', 'acoes_contingencia'];
            const labelsMap = ["Identificação do Risco", "Danos possíveis", "Probabilidade do risco", "Impacto do risco", "Ações preventivas", "Ações de contingência"];
            item.querySelectorAll('.form-group').forEach((group, index) => {
                const newId = `c6_risco${newIndex}_${fieldSuffixes[index]}`;
                group.querySelector('label').htmlFor = newId;
                group.querySelector('label').textContent = `6.${newIndex}.${index + 1} ${labelsMap[index]}`;
                group.querySelector('textarea, select').id = newId;
                const aiButton = group.querySelector('.ai-help-button');
                if (aiButton) aiButton.dataset.targetField = newId;
            });
        }
    },
    contratacao: {
        containerId: 'contratacoes_anteriores_container',
        itemClass: 'contratacao-item',
        indexAttribute: 'data-contratacao-index',
        titleSuffix: 'ª Contratação Anterior',
        removeConfirmText: 'Tem certeza que deseja remover esta contratação anterior?',
        removeClass: 'remove-contratacao-button',
        removeButtonText: 'Remover',
        removeAction: 'remove-contratacao',
        template: (i, showRemove) => {
            const removeBtn = showRemove ? `<button class="remove-contratacao-button" data-action="remove-contratacao"><i class="fas fa-trash-alt"></i> Remover</button>` : '';
            return `<h3>${i}ª Contratação Anterior ${removeBtn}</h3>
                <div class="form-group"><label for="c5_contratacao${i}_contrato_ne">5.${i}.1 Contrato ou nota de empenho anterior</label><div class="input-with-ai"><textarea id="c5_contratacao${i}_contrato_ne" rows="2" placeholder="Nº do Contrato, Nota de Empenho, etc."></textarea><button class="ai-help-button" data-action="get-ai-help" data-target-field="c5_contratacao${i}_contrato_ne"><i class="fas fa-book-open"></i></button></div></div>
                <div class="form-group"><label for="c5_contratacao${i}_processo_sei">5.${i}.2 Número do Processo SEI</label><div class="input-with-ai"><input type="text" id="c5_contratacao${i}_processo_sei" placeholder="0000.00.000000000-0" maxlength="20"><button class="ai-help-button" data-action="get-ai-help" data-target-field="c5_contratacao${i}_processo_sei"><i class="fas fa-book-open"></i></button></div></div>
                <div class="form-group"><label for="c5_contratacao${i}_fase_planejamento">5.${i}.3 Lições aprendidas na fase de planejamento da licitação</label><div class="input-with-ai"><textarea id="c5_contratacao${i}_fase_planejamento" rows="3" placeholder="Análise das recomendações da SAD e Assessoria Jurídica..."></textarea><button class="ai-help-button" data-action="get-ai-help" data-target-field="c5_contratacao${i}_fase_planejamento"><i class="fas fa-book-open"></i></button></div></div>
                <div class="form-group"><label for="c5_contratacao${i}_fase_externa">5.${i}.4 Lições aprendidas na fase externa da licitação</label><div class="input-with-ai"><textarea id="c5_contratacao${i}_fase_externa" rows="3" placeholder="Questionamentos, impugnações, diligências, inabilitações, recursos..."></textarea><button class="ai-help-button" data-action="get-ai-help" data-target-field="c5_contratacao${i}_fase_externa"><i class="fas fa-book-open"></i></button></div></div>
                <div class="form-group"><label for="c5_contratacao${i}_dificuldades_execucao">5.${i}.5 Dificuldades da execução contratual</label><div class="input-with-ai"><textarea id="c5_contratacao${i}_dificuldades_execucao" rows="3" placeholder="Apontar dificuldades e problemas durante a execução..."></textarea><button class="ai-help-button" data-action="get-ai-help" data-target-field="c5_contratacao${i}_dificuldades_execucao"><i class="fas fa-book-open"></i></button></div></div>`;
        },
        onAdd: (newItemDiv) => {
            newItemDiv.querySelector('input[id*="_processo_sei"]').addEventListener('input', (e) => formatSEIProcess(e.target));
        },
        renumberFields: (item, newIndex) => {
            const fieldSuffixes = ['contrato_ne', 'processo_sei', 'fase_planejamento', 'fase_externa', 'dificuldades_execucao'];
            const labelsMap = ["Contrato ou nota de empenho anterior", "Número do Processo SEI", "Lições aprendidas na fase de planejamento da licitação", "Lições aprendidas na fase externa da licitação", "Dificuldades da execução contratual"];
            item.querySelectorAll('.form-group').forEach((group, index) => {
                const newId = `c5_contratacao${newIndex}_${fieldSuffixes[index]}`;
                group.querySelector('label').htmlFor = newId;
                group.querySelector('label').textContent = `5.${newIndex}.${index + 1} ${labelsMap[index]}`;
                group.querySelector('textarea, input').id = newId;
                const aiButton = group.querySelector('.ai-help-button');
                if (aiButton) aiButton.dataset.targetField = newId;
            });
        }
    },
    anexo: {
        containerId: 'anexos_container',
        itemClass: 'anexo-item',
        indexAttribute: 'data-anexo-index',
        titleSuffix: 'º Anexo',
        removeConfirmText: 'Tem certeza que deseja remover este anexo?',
        removeClass: 'remove-anexo-button',
        removeButtonText: 'Remover Anexo',
        removeAction: 'remove-anexo',
        template: (i, showRemove) => {
            const removeBtn = showRemove ? `<button class="remove-anexo-button" data-action="remove-anexo"><i class="fas fa-trash-alt"></i> Remover Anexo</button>` : '';
            return `<h3>${i}º Anexo ${removeBtn}</h3>
                <div class="form-group"><label for="anexo${i}_titulo">A.${i}.1 Título do Anexo</label><div class="input-with-ai"><input type="text" id="anexo${i}_titulo"><button class="ai-help-button" data-action="get-ai-help" data-target-field="anexo${i}_titulo"><i class="fas fa-book-open"></i></button></div></div>
                <div class="form-group"><label for="anexo${i}_capitulo_relacionado">A.${i}.2 Capítulo relacionado com o anexo</label><div class="input-with-ai"><select id="anexo${i}_capitulo_relacionado" data-action="populate-anexo-items" data-anexo-index="${i}"><option value="">Selecione um capítulo...</option><option value="cap1">Capítulo 1: Necessidade</option><option value="cap2">Capítulo 2: Soluções de Mercado</option><option value="cap3">Capítulo 3: Solução Escolhida</option><option value="cap4">Capítulo 4: Valor Estimado</option><option value="cap5">Capítulo 5: Contratações Anteriores</option><option value="cap6">Capítulo 6: Análise dos Riscos</option><option value="cap7">Capítulo 7: Parcelamento</option><option value="cap8">Capítulo 8: Aspectos Administrativos</option><option value="cap9">Capítulo 9: Registro de Preços</option></select><button class="ai-help-button" data-action="get-ai-help" data-target-field="anexo${i}_capitulo_relacionado"><i class="fas fa-book-open"></i></button></div></div>
                <div class="form-group" style="width:100%"><label for="anexo${i}_itens_relacionados_container">A.${i}.3 Item relacionado com o anexo</label><div class="input-with-ai"><div class="checkbox-group scrollable-checkbox-group" id="anexo${i}_itens_relacionados_container"><small><i>Selecione um capítulo acima para ver os itens.</i></small></div><button class="ai-help-button" data-action="get-ai-help" data-target-field="anexo${i}_itens_relacionados_container"><i class="fas fa-book-open"></i></button></div></div>
                <div class="form-group"><label for="anexo${i}_numero_sei">A.${i}.4 Número SEI do Anexo</label><div class="input-with-ai"><input type="text" id="anexo${i}_numero_sei" maxlength="7" pattern="\\d{7}" oninput="this.value = this.value.replace(/\\D/g, '')"><button class="ai-help-button" data-action="get-ai-help" data-target-field="anexo${i}_numero_sei"><i class="fas fa-book-open"></i></button></div></div>`;
        },
        renumberFields: (item, newIndex) => {
            const fieldSuffixes = ['titulo', 'capitulo_relacionado', 'itens_relacionados_container', 'numero_sei'];
            const labelsMap = ["Título do Anexo", "Capítulo relacionado com o anexo", "Item relacionado com o anexo", "Número SEI do Anexo"];
            item.querySelectorAll('.form-group').forEach((group, index) => {
                const newId = `anexo${newIndex}_${fieldSuffixes[index]}`;
                group.querySelector('label').htmlFor = newId;
                group.querySelector('label').textContent = `A.${newIndex}.${index + 1} ${labelsMap[index]}`;
                const input = group.querySelector('input, select, div.checkbox-group');
                input.id = newId;
                if (input.tagName === 'SELECT') {
                    input.dataset.anexoIndex = newIndex;
                }
                const aiButton = group.querySelector('.ai-help-button');
                if (aiButton) aiButton.dataset.targetField = newId;
            });
        }
    }
};