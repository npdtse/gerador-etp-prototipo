/**
 * ===================================================================================
 * MODELO_ETP.JS - GERADOR DE MODELO EM BRANCO (MODO ADMINISTRATIVO)
 * ===================================================================================
 * 
 * Este script adiciona uma funcionalidade oculta para gerar um documento DOCX em branco
 * contendo todos os campos do ETP, formatado de acordo com o padrão oficial.
 * 
 * COMO ACESSAR:
 * Adicione "?modo=admin" ao final da URL do aplicativo (ex: index.html?modo=admin).
 * 
 */

document.addEventListener('DOMContentLoaded', () => {
    const urlParams = new URLSearchParams(window.location.search);
    if (urlParams.get('modo') === 'admin') {
        initAdminOverlay();
    }
});

/**
 * Cria a interface administrativa sobreposta ao app.
 */
function initAdminOverlay() {
    const elementsToHide = ['.tab-container', '.etp-toolbar', 'header', '#beta-alert-banner'];
    elementsToHide.forEach(selector => {
        const el = document.querySelector(selector);
        if (el) el.style.display = 'none';
    });

    const overlay = document.createElement('div');
    Object.assign(overlay.style, {
        position: 'fixed', top: '0', left: '0', width: '100%', height: '100%',
        backgroundColor: '#f4f6f9', display: 'flex', justifyContent: 'center', alignItems: 'center',
        zIndex: '9999', fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif'
    });

    overlay.innerHTML = `
        <div style="background: white; padding: 50px; border-radius: 8px; box-shadow: 0 4px 20px rgba(0,0,0,0.08); text-align: center; width: 100%; max-width: 500px; border: 1px solid #dee2e6;">
            <div style="font-size: 52px; color: #004080; margin-bottom: 25px;"><i class="fas fa-file-word"></i></div>
            <h1 style="color: #333; font-size: 24px; margin-bottom: 15px; font-weight: 600;">Gerador de Modelos de ETP</h1>
            <p style="color: #666; margin-bottom: 30px; line-height: 1.6;">
                Selecione qual estrutura de modelo em branco você deseja baixar.
            </p>
            
            <div style="display: flex; flex-direction: column; gap: 15px;">
                <button id="btnGenCompleto" style="padding: 14px; font-size: 16px; background-color: #0056b3; color: white; border: none; border-radius: 5px; cursor: pointer; font-weight: 500; transition: 0.2s;">
                    <i class="fas fa-download" style="margin-right: 8px;"></i> Modelo ETP Completo
                </button>
                
                <button id="btnGenSimplificado" style="padding: 14px; font-size: 16px; background-color: #28a745; color: white; border: none; border-radius: 5px; cursor: pointer; font-weight: 500; transition: 0.2s;">
                    <i class="fas fa-download" style="margin-right: 8px;"></i> Modelo ETP Simplificado
                </button>
            </div>

            <div style="margin-top: 30px; border-top: 1px solid #eee; padding-top: 20px;">
                <a href="${window.location.pathname}" style="color: #6c757d; text-decoration: none; font-size: 14px; display: inline-flex; align-items: center; gap: 5px;">
                    <i class="fas fa-arrow-left"></i> Voltar ao Aplicativo
                </a>
            </div>
        </div>
    `;

    document.body.appendChild(overlay);

    // Listeners para os botões
    document.getElementById('btnGenCompleto').addEventListener('click', () => generateBlankTemplate(false));
    document.getElementById('btnGenSimplificado').addEventListener('click', () => generateBlankTemplate(true));
}

// --- DICIONÁRIO DE NOMES AMIGÁVEIS ---
const friendlyConditionNames = {
    "etp_tipo": "tipo de ETP",
    "etp_auth": "autorização de ETP simplificado",
    "alinhamento_estrategico": "alinhamento estratégico",
    "natureza_continua": "natureza contínua",
    "beneficios_vigencia": "benefícios da vigência",
    "garantia": "garantia",
    "justificativa_garantia": "justificativa de garantia fora do padrão",
    "assistencia": "assistência técnica",
    "justificativa_assistencia": "justificativa de assistência fora do padrão",
    "transferencia": "transferência de conhecimento",
    "capacitacao": "capacitação",
    "contratacao_adicional": "contratação adicional",
    "ajuste_contratacoes": "ajuste em contratações",
    "acessibilidade": "acessibilidade",
    "parcelamento": "forma de parcelamento",
    "parcelamento_grupo_unico": "agrupamento único",
    "parcelamento_grupos_separados": "grupos separados",
    "modalidade_contratacao": "modalidade da contratação",
    "contratacao_dispensa": "dispensa de licitação",
    "contratacao_inexigibilidade": "inexigibilidade de licitação",
    "dispensa_outra_hipotese": "outra hipótese de dispensa",
    "qualificacao_tecnica": "qualificação técnica",
    "amostras": "amostras ou prova de conceito",
    "vistoria": "vistoria",
    "confidencialidade": "confidencialidade",
    "subcontratacao": "subcontratação",
    "consorcio": "formação de consórcio",
    "consorcio_proibicao": "proibição de consórcio",
    "consorcio_permissao": "permissão de consórcio",
    "limite_consorcio": "limite de consorciadas",
    "cooperativas": "participação de cooperativas",
    "cooperativas_proibicao": "proibição de cooperativas",
    "estrangeiras": "empresas estrangeiras",
    "estrangeiras_proibicao": "proibição de estrangeiras",
    "estrangeiras_permissao": "permissão de estrangeiras",
    "margem_preferencia": "margem de preferência",
    "pessoa_fisica": "participação de pessoa física",
    "pessoa_fisica_proibicao": "proibição de pessoa física",
    "quantitativo_inferior": "quantitativo inferior",
    "precos_diferentes": "preços diferentes",
    "mais_de_um_fornecedor": "mais de um fornecedor",
    "adesao_futura": "adesão futura",
    "prorrogacao_ata": "prorrogação da ata",
    "renovacao_quantidades_ata": "renovação de quantidades"
};

function getFriendlyTriggerText(targetFieldId) {
    if (typeof conditionalFieldIds === 'undefined') return "opção anterior";
    for (const [triggerKey, targets] of Object.entries(conditionalFieldIds)) {
        if (targets.includes(targetFieldId)) {
            return friendlyConditionNames[triggerKey] || triggerKey;
        }
    }
    return "opção anterior";
}

/**
 * Função Principal de Geração do DOCX
 */
function generateBlankTemplate(isSimplificado) {
    const { Document, Packer, Paragraph, TextRun, HeadingLevel, BorderStyle, AlignmentType } = window.docx;

    const FONT_FAMILY = "Calibri";
    const COLOR_BLACK = "000000";
    const COLOR_RED = "FF0000";
    const docTitleText = isSimplificado ? "ESTUDO TÉCNICO PRELIMINAR SIMPLIFICADO" : "ESTUDO TÉCNICO PRELIMINAR COMPLETO";

    const styles = {
        docTitle: { size: 32, bold: true, font: FONT_FAMILY, color: COLOR_BLACK, allCaps: true },
        chapterTitle: { size: 28, bold: true, font: FONT_FAMILY, color: COLOR_BLACK, spacing: { before: 1200, after: 200 } },
        dynamicSectionTitle: { size: 24, bold: true, font: FONT_FAMILY, color: "333333", italics: true, spacing: { before: 400, after: 100 } },
        fieldLabel: { size: 22, bold: true, font: FONT_FAMILY, color: COLOR_BLACK, spacing: { before: 600, after: 100 } },
        instruction: { size: 20, italics: true, color: "666666", font: FONT_FAMILY, spacing: { after: 100 } },
        redNote: { size: 22, italics: true, color: COLOR_RED, font: FONT_FAMILY, spacing: { after: 200 } },
        placeholder: { size: 22, color: "555555", italics: true, font: FONT_FAMILY },
        optionText: { size: 22, color: COLOR_BLACK, font: FONT_FAMILY }
    };

    const docChildren = [];
    const tabOrder = ['identificacao', 'cap1', 'cap2', 'cap3', 'cap4', 'cap5', 'cap6', 'cap7', 'cap8', 'cap9', 'anexos'];
    const processedFieldIds = new Set();

    // Título Principal
    docChildren.push(new Paragraph({
        children: [new TextRun({ text: docTitleText, ...styles.docTitle })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 600 }
    }));

    // Flag para controlar se já inserimos a mensagem de exclusão do range 3.9-3.15
    let hasInsertedCap3Exclusion = false;

    // === FUNÇÃO HELPER: Renderizar Campo ===
    function renderField(groupElement, forcedLabel = null) {
        if (groupElement.id.includes('table_container') || groupElement.id.includes('resumo_')) return;

        const label = groupElement.querySelector('label') || groupElement.querySelector('.label-with-help label');
        const inputElement = groupElement.querySelector('input:not([type="hidden"]), select, textarea');
        
        if (!label) return;
        if (groupElement.id && processedFieldIds.has(groupElement.id)) return;
        if (groupElement.id) processedFieldIds.add(groupElement.id);
        if (inputElement && inputElement.id) {
            if (processedFieldIds.has(inputElement.id)) return;
            processedFieldIds.add(inputElement.id);
        }

        // Título do Campo
        let labelText = forcedLabel || label.textContent.trim();
        docChildren.push(new Paragraph({
            children: [new TextRun({ text: labelText, ...styles.fieldLabel })],
            keepNext: true
        }));

        // Nota Condicional
        if (groupElement.classList.contains('conditional-field')) {
            const friendlyName = getFriendlyTriggerText(groupElement.id);
            docChildren.push(new Paragraph({
                children: [new TextRun({ text: `Nota: Preencher somente se necessário conforme a opção "${friendlyName}".`, ...styles.instruction })]
            }));
        }

        // Opções ou Placeholder
        const checkOptions = groupElement.querySelectorAll('.checkbox-group label');
        
        if (checkOptions.length > 0) {
            checkOptions.forEach(opt => {
                let marker = "(   )";
                const optText = opt.textContent.trim();
                
                // Lógica de Pré-Seleção (X)
                if (isSimplificado) {
                    if (labelText.includes("Tipo de ETP") && optText.includes("Simplificado")) marker = "( X )";
                    if (labelText.includes("autorização formal") && optText === "Sim") marker = "( X )";
                } else {
                    if (labelText.includes("Tipo de ETP") && optText.includes("Completo")) marker = "( X )";
                }

                docChildren.push(new Paragraph({
                    children: [
                        new TextRun({ text: `${marker} `, font: "Courier New", size: 22 }),
                        new TextRun({ text: optText, ...styles.optionText })
                    ],
                    indent: { left: 400 }
                }));
            });
        } else if (inputElement && inputElement.tagName === 'SELECT') {
            const options = Array.from(inputElement.options);
            let hasValidOptions = false;
            options.forEach(opt => {
                if (opt.value && opt.value.trim() !== "") {
                    hasValidOptions = true;
                    docChildren.push(new Paragraph({
                        children: [
                            new TextRun({ text: "(   ) ", font: "Courier New", size: 22 }),
                            new TextRun({ text: opt.text.trim(), ...styles.optionText })
                        ],
                        indent: { left: 400 }
                    }));
                }
            });
            if (!hasValidOptions) {
                docChildren.push(new Paragraph({ children: [new TextRun({ text: "[Insira a resposta aqui]", ...styles.placeholder })], spacing: { after: 200 } }));
            }
        } else {
            docChildren.push(new Paragraph({ children: [new TextRun({ text: "[Insira a resposta aqui]", ...styles.placeholder })], spacing: { after: 200 } }));
        }
    }

    tabOrder.forEach(tabId => {
        const tabEl = document.getElementById(tabId);
        if (!tabEl) return;

        // Título do Capítulo
        const h2 = tabEl.querySelector('h2');
        if (h2) {
            let titleText = h2.innerText.trim().replace(/^\W+/, '');
            docChildren.push(new Paragraph({
                children: [new TextRun({ text: titleText, ...styles.chapterTitle })],
                heading: HeadingLevel.HEADING_1,
                border: { bottom: { color: "BFBFBF", space: 4, value: "single", size: 6 } }
            }));
        }

        // LÓGICA DE EXCLUSÃO DE CAPÍTULOS INTEIROS (SIMPLIFICADO)
        if (isSimplificado && (tabId === 'cap2' || tabId === 'cap5')) {
            docChildren.push(new Paragraph({
                children: [new TextRun({ text: "Nota: Este capítulo não se aplica ao ETP Simplificado.", ...styles.redNote })],
                spacing: { after: 400 }
            }));
            return;
        }

        // LÓGICA CAP 5 (COMPLETO) - Nota inicial
        if (tabId === 'cap5' && !isSimplificado) {
            const naoSeAplicaGroup = document.getElementById('c5_nao_se_aplica')?.closest('.form-group');
            if (naoSeAplicaGroup) {
                docChildren.push(new Paragraph({
                    children: [new TextRun({ text: "5.0 Não se aplica / Primeira Contratação", ...styles.fieldLabel })],
                    keepNext: true
                }));
                docChildren.push(new Paragraph({
                    children: [new TextRun({ text: "Nota: Caso este item seja marcado, os itens do bloco dinâmico abaixo não devem ser preenchidos.", ...styles.instruction })]
                }));
                docChildren.push(new Paragraph({
                    children: [
                        new TextRun({ text: "(   ) ", font: "Courier New", size: 22 }),
                        new TextRun({ text: "Sim", ...styles.optionText })
                    ],
                    indent: { left: 400 }
                }));
                processedFieldIds.add('c5_nao_se_aplica');
                if (naoSeAplicaGroup.id) processedFieldIds.add(naoSeAplicaGroup.id);
            }
        }

        // BLOCOS DINÂMICOS
        const dynamicContainers = {
            'cap2': { id: 'solucoes_mercado_container', config: dynamicItemConfigs.solucao, label: "Solução de Mercado", prefix: "2" },
            'cap5': { id: 'contratacoes_anteriores_container', config: dynamicItemConfigs.contratacao, label: "Contratação Anterior", prefix: "5" },
            'cap6': { id: 'riscos_container', config: dynamicItemConfigs.risco, label: "Risco Identificado", prefix: "6" },
            'anexos': { id: 'anexos_container', config: dynamicItemConfigs.anexo, label: "Anexo", prefix: "A" }
        };

        if (dynamicContainers[tabId]) {
            const dynData = dynamicContainers[tabId];
            docChildren.push(new Paragraph({
                children: [new TextRun({ text: `Exemplo de Bloco Dinâmico: ${dynData.label}`, ...styles.dynamicSectionTitle })],
                spacing: { before: 300 }
            }));
            
            const tempDiv = document.createElement('div');
            tempDiv.innerHTML = dynData.config.template(1, false);
            
            tempDiv.querySelectorAll('.form-group').forEach(group => {
                const lbl = group.querySelector('label');
                if(lbl) {
                    let labelText = lbl.textContent.trim();
                    if (dynData.prefix !== "A") {
                        labelText = labelText.replace(/^(\d+)\.\d+\.(\d+)/, '$1.X.$2'); 
                    } else {
                        labelText = labelText.replace(/^A\.\d+\.(\d+)/, 'A.X.$1');
                    }
                    const cloneGroup = group.cloneNode(true);
                    renderField(cloneGroup, labelText);
                }
            });
        }

        // CAMPOS ESTÁTICOS
        tabEl.querySelectorAll('.form-group').forEach(group => {
            if (group.closest('.solucao-item, .risco-item, .contratacao-item, .anexo-item')) return;
            if (group.id.includes('resumo_') || group.querySelector('.summary-table-container')) return;

            // Filtros de visibilidade do formulário
            if (isSimplificado && group.classList.contains('simplificado-hide')) return;

            // Lógica Item 1.4 (Completo)
            if (!isSimplificado && group.classList.contains('completo-hide')) {
                const label = group.querySelector('label');
                if (label && label.textContent.includes('1.4')) {
                    docChildren.push(new Paragraph({
                        children: [new TextRun({ text: label.textContent.trim(), ...styles.fieldLabel })],
                        spacing: { before: 600 }
                    }));
                    docChildren.push(new Paragraph({
                        children: [new TextRun({ text: "Nota: Este item não se aplica ao ETP Completo.", ...styles.instruction })],
                        spacing: { after: 200 }
                    }));
                    processedFieldIds.add(group.id);
                    return;
                }
            }

            // Lógica Capítulo 3 (Simplificado - Range 3.9 a 3.15)
            if (isSimplificado && tabId === 'cap3') {
                const id = group.id || "";
                // Verifica se é um dos itens do range
                if (id.startsWith('c3_9') || id.startsWith('c3_10') || id.startsWith('c3_11') || 
                    id.startsWith('c3_12') || id.startsWith('c3_13') || id.startsWith('c3_14') || id.startsWith('c3_15') ||
                    group.querySelector('label')?.getAttribute('for')?.startsWith('c3_9')) {
                    
                    if (!hasInsertedCap3Exclusion) {
                        docChildren.push(new Paragraph({
                            children: [new TextRun({ text: "3.9 a 3.15 Os itens 3.9 a 3.15 não se aplicam ao ETP Simplificado.", ...styles.redNote })],
                            spacing: { before: 600, after: 400 }
                        }));
                        hasInsertedCap3Exclusion = true;
                    }
                    processedFieldIds.add(group.id);
                    return; // Pula a renderização individual
                }
            }

            // Lógica Item 8.7 (Simplificado)
            if (isSimplificado && tabId === 'cap8' && group.id.includes('c8_7_')) {
                const label = group.querySelector('label');
                if (label && label.textContent.includes('8.7')) {
                    docChildren.push(new Paragraph({
                        children: [new TextRun({ text: "8.7 O item 8.7 não se aplica ao ETP Simplificado.", ...styles.redNote })],
                        spacing: { before: 600 }
                    }));
                    processedFieldIds.add(group.id);
                    return;
                }
                // Se for o subitem (justificativa), apenas ignora
                if (label && label.textContent.includes('8.7.1')) return;
            }

            // Tratamento 3.1
            if (tabId === 'cap3') {
                const isCompletoWrapper = group.id === 'c3_1_completo_wrapper';
                const isSimplificadoWrapper = group.id === 'c3_1_simplificado_wrapper';
                if (isSimplificado && isCompletoWrapper) return;
                if (!isSimplificado && isSimplificadoWrapper) return;
            }

            // Renumeração Inexigibilidade (Cap 8)
            let forcedLabel = null;
            const input = group.querySelector('select, textarea, input');
            if (input) {
                if (input.id === 'c8_1_1_hipotese_inexigibilidade') {
                    forcedLabel = "8.1.2 Hipótese de Inexigibilidade";
                } else if (input.id === 'c8_1_2_justificativa_inexigibilidade') {
                    forcedLabel = "8.1.2.1 Justificativa da inviabilidade de competição";
                }
            }

            renderField(group, forcedLabel);
        });
        
        // Reseta flag para o próximo capítulo (segurança)
        hasInsertedCap3Exclusion = false;
    });

    const filename = isSimplificado ? "Modelo_ETP_Simplificado.docx" : "Modelo_ETP_Completo.docx";

    Packer.toBlob(new Document({
        sections: [{
            properties: { page: { margin: { top: 1440, bottom: 1440, left: 1134, right: 1134 } } },
            children: docChildren,
        }],
    })).then(blob => saveAs(blob, filename));
}