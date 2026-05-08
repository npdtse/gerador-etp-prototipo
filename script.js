document.addEventListener('DOMContentLoaded', () => {
    // --- SELETORES DE ELEMENTOS ---
    // As constantes de configuração e os objetos de mapeamento estão em config.js

    // Modal de Ajuda da IA
    const aiModal = document.getElementById('aiResponseModal');
    const aiComplianceAnalysisEl = document.getElementById('aiComplianceAnalysis');
    const aiModelNameEl = document.getElementById('aiModelName');

    // Barras de Progresso
    const mainProgressBarEl = document.getElementById('etpProgressBar');
    const mainProgressTextEl = document.getElementById('etpProgressText');
    const tabProgressBarEl = document.getElementById('tabProgressBar');
    const tabProgressTextEl = document.getElementById('tabProgressText');
    const tabProgressWrapper = document.getElementById('tabProgressWrapper');

    // Seletores do Modal de Configurações de IA ---
    const aiSettingsModal = document.getElementById('aiSettingsModal');
    const useAiRadioYes = document.getElementById('useAiRadioYes');
    const useAiRadioNo = document.getElementById('useAiRadioNo');
    const aiSettingsContainer = document.getElementById('aiSettingsContainer');
    const apiKeyInput = document.getElementById('apiKeyInput');
    const modelModeDefault = document.getElementById('modelModeDefault');
    const modelModeCustom = document.getElementById('modelModeCustom');
    const customModelsSection = document.getElementById('customModelsSection');
    const customModelsList = document.getElementById('customModelsList');
    const addModelButton = document.getElementById('addModelButton');
    const testApiConnectionButton = document.getElementById('testApiConnectionButton');
    const testConnectionResult = document.getElementById('testConnectionResult');

    // Recursos
    const btnGuiaSobre = document.getElementById('btnAbrirGuiaRapido');
    if (btnGuiaSobre) btnGuiaSobre.addEventListener('click', () => window.open('guia_rapido.html', '_blank'));

    // Bibliotecas externas
    const { Packer, Document, Paragraph, TextRun, Table, TableCell, TableRow, HeadingLevel, WidthType, BorderStyle, VerticalAlign, HeightRule } = window.docx;

    // Variáveis Globais ---
    let aiConfig = {}; // Será preenchida pela função loadAiSettings
    let currentEtpFilename = null; // Armazena o nome do arquivo atual
    let isAppInitializing = false;


    // --- FUNÇÕES AUXILIARES ---

    // --- FUNÇÕES DE UX ---

    function updateHomeTab() {
        const btnContinue = document.getElementById('btnContinueEtpHome');
        const etpDataString = localStorage.getItem(ETP_DATA_KEY);

        if (btnContinue) {
            if (etpDataString) {
                // Se tem dados salvos, habilita o botão de continuar
                btnContinue.disabled = false;
            } else {
                // Se não tem dados, desabilita (fica cinza)
                btnContinue.disabled = true;
            }
        }
    }

    function formatSEIProcess(input) {
        let value = input.value.replace(/\D/g, '');
        let formattedValue = '';
        if (value.length > 0) formattedValue = value.substring(0, 4);
        if (value.length > 4) formattedValue += '.' + value.substring(4, 6);
        if (value.length > 6) formattedValue += '.' + value.substring(6, 15);
        if (value.length > 15) formattedValue += '-' + value.substring(15, 16);
        input.value = formattedValue;
    }

    function stripHtml(html) {
        if (!html) return "";
        const tmp = document.createElement("DIV");
        tmp.innerHTML = html;
        let text = tmp.textContent || tmp.innerText || "";
        // Remove espaços extras e quebras de linha excessivas
        return text.replace(/\s+/g, " ").trim();
    }

    function getEtpContext() {
        const getVal = (id) => {
            const el = document.getElementById(id);
            return (el && el.value.trim()) ? el.value.trim() : "Não informado/Não preenchido";
        };

        const unidade = getVal('etp_unidade_demandante');
        const titulo = getVal('etp_titulo');
        const necessidade = getVal('c1_1_indicacao_necessidade');
        const problemas = getVal('c1_2_situacoes_problemas');

        // Lógica específica para capturar a Solução Escolhida
        let solucao = "Não definida";
        const isSimplificado = document.body.classList.contains('etp-simplificado-mode');

        if (isSimplificado) {
            solucao = getVal('c3_1_solucao_escolhida_simplificada');
        } else {
            const selectEl = document.getElementById('c3_1_solucao_escolhida');
            if (selectEl && selectEl.selectedIndex > 0) {
                // Pega o texto da opção selecionada (ex: "Solução 1: Compra de Notebooks...")
                solucao = selectEl.options[selectEl.selectedIndex].text;
            } else {
                solucao = "Não selecionada na lista";
            }
        }

        return `
            - Unidade Demandante: ${unidade}
            - Título do ETP: ${titulo}
            - Necessidade a atender: ${necessidade}
            - Problemas enfrentados: ${problemas}
            - Solução Escolhida: ${solucao}`;
    }

    function smartSubstring(text, maxLength) {
        if (!text || text.length <= maxLength) return text;
        return text.substring(0, maxLength - 3) + "...";
    }

    function clearField(fieldId) {
        if (!fieldId) return;
        const field = document.getElementById(fieldId);
        if (field) {
            if (field.tagName === 'SELECT') field.value = "";
            else if (field.type === 'checkbox' || field.type === 'radio') field.checked = false;
            else field.value = "";
        }
    }

    function initializeConditionalFields() {
        document.querySelectorAll('[data-action="toggle-conditional"]').forEach(el => {
            const event = new Event('change', { 'bubbles': true });
            if (el.type === 'radio') {
                const checkedRadio = document.querySelector(`input[name="${el.name}"]:checked`);
                if (checkedRadio) checkedRadio.dispatchEvent(event);
                else el.dispatchEvent(event);
            } else {
                el.dispatchEvent(event);
            }
        });

        const cap5Checkbox = document.getElementById('c5_nao_se_aplica');
        if (cap5Checkbox) {
            cap5Checkbox.dispatchEvent(new Event('change', { 'bubbles': true }));
        }

        const vigenciaInput = document.getElementById('c3_8_vigencia_meses');
        if (vigenciaInput) {
            toggleVigenciaBenefitsField(vigenciaInput);
        }
    }

    function resetAuthChoice() {
        const authSelect = document.getElementById('etp_auth');
        if (authSelect) {
            authSelect.value = '';
            authSelect.dispatchEvent(new Event('change', { 'bubbles': true }));
        }
    }

    // --- LÓGICA DE NAVEGAÇÃO E UI ---

    function openTab(tabName) {
        const tabButton = document.querySelector(`.tab-button[data-tab-target="${tabName}"]`);

        // NOVO: Impede a abertura se o botão não existir, estiver invisível ou estiver desativado (inactive)
        if (!tabButton || window.getComputedStyle(tabButton).display === 'none' || tabButton.classList.contains('inactive')) {
            return;
        }

        document.querySelectorAll(".tab-content").forEach(tc => {
            tc.style.display = "none";
            tc.classList.remove("active");
        });

        document.querySelectorAll(".tab-button").forEach(tb => {
            tb.classList.remove("active");
        });

        const activeTab = document.getElementById(tabName);
        if (activeTab) {
            activeTab.style.display = "block";
            activeTab.classList.add("active");
        }

        tabButton.classList.add("active");

        // --- INÍCIO DA CORREÇÃO ---
        if (tabName === 'consolidacao') {
            updateConsolidatedETP(); // Isso só deve rodar para a consolidação
            tabProgressWrapper.style.display = 'none';
        } else if (tabName === 'inicio') {
            tabProgressWrapper.style.display = 'none'; // Para o início, apenas escondemos a barra
        } else {
            tabProgressWrapper.style.display = 'flex';
        }
        // --- FIM DA CORREÇÃO ---

        updateAllProgressBars();
    }

    function navigateToChapter(direction) {
        const tabOrder = ['identificacao', 'cap1', 'cap2', 'cap3', 'cap4', 'cap5', 'cap6', 'cap7', 'cap8', 'cap9', 'anexos', 'consolidacao'];
        const currentTab = document.querySelector('.tab-content.active');
        if (!currentTab) return;

        let currentIndex = tabOrder.indexOf(currentTab.id);
        if (currentIndex === -1) return;

        let step = direction === 'next' ? 1 : -1;
        let targetIndex = currentIndex + step;

        // Loop para pular capítulos ocultos/inativos
        while (targetIndex >= 0 && targetIndex < tabOrder.length) {
            const targetId = tabOrder[targetIndex];
            const navButton = document.querySelector(`.tab-button[data-tab-target="${targetId}"]`);

            const isInactive = navButton && navButton.classList.contains('inactive');
            const isHidden = navButton && window.getComputedStyle(navButton).display === 'none';

            if (!isInactive && !isHidden) {
                openTab(targetId);

                // CORREÇÃO DEFINITIVA: Rola a janela inteira (window) para o topo.
                setTimeout(() => {
                    window.scrollTo({ top: 0, behavior: 'smooth' });
                }, 0);

                break;
            }
            targetIndex += step;
        }
    }

    function applyEtpMode(mode) {
        const isSimplificado = mode === 'simplificado';
        document.body.classList.toggle('etp-simplificado-mode', isSimplificado);

        // 1. Gerencia itens que devem ser desativados no modo SIMPLIFICADO
        document.querySelectorAll('.form-group.simplificado-hide').forEach(container => {
            container.querySelectorAll('input, textarea, select, button').forEach(input => {
                input.disabled = isSimplificado;
            });
        });

        // 2. Gerencia itens que devem ser desativados no modo COMPLETO (ex: item 1.4)
        document.querySelectorAll('.form-group.completo-hide').forEach(container => {
            container.querySelectorAll('input, textarea, select, button').forEach(input => {
                // Se NÃO for simplificado (ou seja, modo Completo), o campo deve ser desabilitado
                input.disabled = !isSimplificado;
            });
        });

        // 3. Gerencia a troca de inputs do item 3.1
        const completoWrapper = document.getElementById('c3_1_completo_wrapper');
        const simplificadoWrapper = document.getElementById('c3_1_simplificado_wrapper');

        if (completoWrapper && simplificadoWrapper) {
            if (isSimplificado) {
                completoWrapper.style.display = 'none';
                simplificadoWrapper.style.display = 'block';
            } else {
                completoWrapper.style.display = 'block';
                simplificadoWrapper.style.display = 'none';
            }
        }

        // 4. Se o usuário estiver em uma aba que se tornou inativa, volta para identificação
        const activeNavButton = document.querySelector('.tab-button.active');
        if (activeNavButton && activeNavButton.classList.contains('inactive')) {
            openTab('identificacao');
        }

        updateAllProgressBars();
        updateAllSummariesAndDropDowns();
    }

    function updateChapterAccess() {
        const tipoEtpSelect = document.getElementById('etp_tipo');
        const authSelect = document.getElementById('etp_auth');
        const etpDataString = localStorage.getItem(ETP_DATA_KEY);
        const activeTab = document.querySelector('.tab-content.active');

        // --- BLOQUEIO DA ABA IDENTIFICAÇÃO ---
        const btnIdentificacao = document.querySelector('.tab-button[data-tab-target="identificacao"]');
        if (btnIdentificacao) {
            // A aba Identificação só fica ativa se houver dados salvos OU se o usuário já estiver nela (onboarding)
            const isCurrentlyOnIdentificacao = activeTab && activeTab.id === 'identificacao';
            btnIdentificacao.classList.toggle('inactive', !etpDataString && !isCurrentlyOnIdentificacao);
        }

        // --- LÓGICA DE UX (PULSO VISUAL) ---
        if (tipoEtpSelect) {
            if (tipoEtpSelect.value === '') {
                tipoEtpSelect.classList.add('highlight-pulse');
            } else {
                tipoEtpSelect.classList.remove('highlight-pulse');
            }
        }

        const isCompleto = tipoEtpSelect?.value === 'completo';
        const isSimplificado = tipoEtpSelect?.value === 'simplificado';
        const hasAuth = authSelect?.value === 'sim';
        const seiNumber = document.getElementById('etp_auth_sei_number')?.value.trim();

        // Verifica se SRP foi selecionado no Capítulo 8 (item 8.2)
        const isSrpSelected = document.getElementById('c8_2_procedimentos_auxiliares')?.value === 'srp';

        const enableCompleto = isCompleto;
        const enableSimplificado = isSimplificado && hasAuth && seiNumber !== '';

        // Seleciona apenas os botões de CAPÍTULOS (exclui Início e Identificação da regra geral)
        const chapterNavButtons = document.querySelectorAll('.tab-nav .tab-button:not([data-tab-target="identificacao"]):not([data-tab-target="inicio"])');

        chapterNavButtons.forEach(button => {
            const tabTarget = button.dataset.tabTarget;
            const isSimplificadoHidden = button.classList.contains('simplificado-hide');
            let shouldBeActive = false;
            let tooltipText = "";

            if (enableCompleto) {
                shouldBeActive = true;
            } else if (enableSimplificado) {
                if (!isSimplificadoHidden) {
                    shouldBeActive = true;
                } else {
                    tooltipText = "Este capítulo não se aplica ao ETP Simplificado.";
                }
            } else {
                tooltipText = "Preencha a aba Identificação e o Tipo de ETP para liberar este capítulo.";
            }

            // Lógica específica para o Capítulo 9 (Registro de Preços)
            if (tabTarget === 'cap9') {
                // Para o tooltip funcionar, o botão deve estar visível na tela
                button.style.display = 'flex'; 
                
                if ((enableCompleto || enableSimplificado) && isSrpSelected) {
                    shouldBeActive = true;
                    tooltipText = ""; // Limpa a tooltip pois a aba está ativa
                } else if (enableCompleto || enableSimplificado) {
                    shouldBeActive = false;
                    tooltipText = "Requer a seleção de Sistema de Registro de Preços (SRP) no item 8.2.";
                }
            } else {
                button.style.display = 'flex';
            }

            button.classList.toggle('inactive', !shouldBeActive);
            
            // Aplica a tooltip (se vazio, remove o atributo title)
            if (tooltipText) {
                button.setAttribute('title', tooltipText);
            } else {
                button.removeAttribute('title');
            }
        });

        // Se a aba atual for um capítulo bloqueado, volta para identificação (mas nunca expulsa da aba Início)
        const activeNavButton = document.querySelector('.tab-button.active');
        if (activeNavButton && activeNavButton.classList.contains('inactive')) {
            if (activeNavButton.dataset.tabTarget !== 'inicio') {
                openTab('identificacao');
            }
        }

        if (enableSimplificado) {
            applyEtpMode('simplificado');
        } else {
            applyEtpMode('completo');
        }

        updateAllProgressBars();

        const btnNextIdentificacao = document.getElementById('btnNext_identificacao');
        if (btnNextIdentificacao) {
            btnNextIdentificacao.disabled = !(enableCompleto || enableSimplificado);
        }
    }

    function toggleVigenciaBenefitsField(inputElement) {
        const benefitsContainer = document.getElementById('conditional_c3_8_1_beneficios_vigencia');
        if (!benefitsContainer || !inputElement) return;

        const vigenciaValue = parseInt(inputElement.value, 10);
        const showBenefits = vigenciaValue > 12;

        benefitsContainer.style.display = showBenefits ? 'block' : 'none';

        if (!showBenefits) {
            const fieldToClear = document.getElementById('c3_8_1_beneficios_vigencia_superior');
            if (fieldToClear) clearField(fieldToClear.id);
        }
        updateAllProgressBars();
    }

    function toggleConditionalFields(groupKey, selectedValue) {
        const toggle = (ids, condition) => {
            if (!ids) return;
            ids.forEach(id => {
                const el = document.getElementById(id);
                if (el) {
                    const show = condition;
                    el.style.display = show ? 'block' : 'none';
                    if (!show) {
                        el.querySelectorAll('input, textarea, select').forEach(input => {
                            if (input.tagName === 'SELECT' && input.dataset.action === 'toggle-conditional') {
                                if (input.value !== '') {
                                    input.value = '';
                                    input.dispatchEvent(new Event('change', { 'bubbles': true }));
                                }
                            } else {
                                clearField(input.id);
                            }
                        });
                        if (id === 'conditional_c3_7_natureza_continua') {
                            toggleVigenciaBenefitsField(document.getElementById('c3_8_vigencia_meses'));
                        }
                    }
                }
            });
        };

        if (groupKey === 'etp_tipo') {
            toggle(conditionalFieldIds.etp_tipo, selectedValue === 'simplificado');
            if (selectedValue === 'completo') {
                resetAuthChoice();
            }
        } else if (groupKey === 'etp_auth') {
            if (selectedValue === 'nao') {
                alert("Para prosseguir com um ETP Simplificado, é obrigatório possuir a autorização prévia da Secretaria de Administração ou da Diretoria-Geral. O formulário foi redefinido para 'ETP Completo'.");
                const tipoEtpSelect = document.getElementById('etp_tipo');
                if (tipoEtpSelect) {
                    tipoEtpSelect.value = 'completo';
                    tipoEtpSelect.dispatchEvent(new Event('change', { 'bubbles': true }));
                }
                return;
            }

            const showSeiField = selectedValue === 'sim';
            toggle(conditionalFieldIds.etp_auth, showSeiField);
            if (!showSeiField) {
                clearField('etp_auth_sei_number');
            }
        } else if (groupKey === 'parcelamento') {
            toggle(conditionalFieldIds.parcelamento_grupo_unico, selectedValue === 'grupo_unico');
            toggle(conditionalFieldIds.parcelamento_grupos_separados, selectedValue === 'grupos_separados');
        } else if (groupKey === 'alinhamento_estrategico') {
            toggle(conditionalFieldIds.alinhamento_estrategico, selectedValue === 'sim');
        } else if (groupKey === 'modalidade_contratacao') {
            toggle(conditionalFieldIds.contratacao_dispensa, selectedValue === 'dispensa');
            toggle(conditionalFieldIds.contratacao_inexigibilidade, selectedValue === 'inexigibilidade');
        } else if (groupKey === 'dispensa_outra_hipotese') {
            toggle(conditionalFieldIds.dispensa_outra_hipotese, selectedValue === 'outra');
        } else if (groupKey === 'consorcio') {
            toggle(conditionalFieldIds.consorcio_permissao, selectedValue === 'sim');
            toggle(conditionalFieldIds.consorcio_proibicao, selectedValue === 'nao');
        } else if (groupKey === 'cooperativas') {
            toggle(conditionalFieldIds.cooperativas_proibicao, selectedValue === 'nao');
        } else if (groupKey === 'estrangeiras') {
            toggle(conditionalFieldIds.estrangeiras_permissao, selectedValue === 'sim');
            toggle(conditionalFieldIds.estrangeiras_proibicao, selectedValue === 'nao');
        } else if (groupKey === 'pessoa_fisica') {
            toggle(conditionalFieldIds.pessoa_fisica_proibicao, selectedValue === 'nao');
        } else {
            const condition = (groupKey.includes('justificativa')) ? selectedValue === 'nao'
                : selectedValue === 'sim';
            toggle(conditionalFieldIds[groupKey], condition);
        }
        updateAllProgressBars();
    }

    function toggleCapNFields(fieldsContainerId, isChecked) {
        const fieldsContainer = document.getElementById(fieldsContainerId);
        if (!fieldsContainer) return;

        fieldsContainer.style.display = isChecked ? 'none' : 'block';
        if (isChecked) {
            fieldsContainer.querySelectorAll('textarea, input[type="text"], input[type="number"], select').forEach(input => {
                clearField(input.id);
            });
        }
        updateAllProgressBars();
        if (fieldsContainerId === 'cap5_fields_container') {
            updateContratacaoSummary();
        }
    }

    // --- LÓGICA DE PROGRESSO ---

    function isFieldLogicallyActive(input) {
        const tabContent = input.closest('.tab-content');
        if (tabContent) {
            const navButton = document.querySelector(`.tab-button[data-tab-target="${tabContent.id}"]`);
            if (navButton && navButton.classList.contains('inactive')) {
                return false;
            }
        }

        const isSimplificadoMode = document.body.classList.contains('etp-simplificado-mode');

        if (isSimplificadoMode && input.closest('.simplificado-hide')) {
            return false;
        }

        if (!isSimplificadoMode && input.closest('.completo-hide')) {
            return false;
        }

        let current = input;
        while (current && !current.classList.contains('tab-content')) {
            if (current.style.display === 'none') {
                return false;
            }
            current = current.parentElement;
        }

        return true;
    }

    function calculateProgress(scopeElement) {
        if (!scopeElement) return { filled: 0, total: 0, percentage: 0 };

        const inputs = scopeElement.querySelectorAll('input[type="text"]:not([readonly]), input[type="number"]:not([readonly]), textarea:not([readonly]), select, input[type="checkbox"], input[type="radio"]');
        let total = 0;
        let filled = 0;
        const processedRadioGroups = new Set();

        inputs.forEach(input => {
            if (input.id === 'c5_nao_se_aplica') {
                return;
            }

            if (isFieldLogicallyActive(input)) {
                if (input.type === 'radio') {
                    const groupName = input.name;
                    if (processedRadioGroups.has(groupName)) return;
                    processedRadioGroups.add(groupName);
                    total++;
                    if (scopeElement.querySelector(`input[name="${groupName}"]:checked`)) {
                        filled++;
                    }
                } else {
                    total++;
                    let isFilled = (input.type === 'checkbox') ? input.checked : (input.value && input.value.trim() !== '');

                    if (isFilled) {
                        if (input.id === 'etp_titulo' && input.value.startsWith("Será gerado")) return;
                        filled++;
                    }
                }
            }
        });

        const percentage = total > 0 ? (filled / total) * 100 : 0;
        return { filled, total, percentage };
    }

    function updateMainProgressBar() {
        let totalFields = 0;
        let filledFields = 0;
        const processedRadioGroups = new Set();

        document.querySelectorAll('.tab-content:not(#consolidacao)').forEach(tab => {
            const navButton = document.querySelector(`.tab-button[data-tab-target="${tab.id}"]`);
            if (navButton && navButton.classList.contains('inactive')) {
                return;
            }

            if (tab.id === 'cap5' && document.getElementById('c5_nao_se_aplica')?.checked) return;

            const inputs = tab.querySelectorAll('input[type="text"]:not([readonly]), input[type="number"]:not([readonly]), textarea:not([readonly]), select, input[type="checkbox"], input[type="radio"]');
            inputs.forEach(input => {
                if (input.id === 'c5_nao_se_aplica') return;

                if (isFieldLogicallyActive(input)) {
                    if (input.type === 'radio') {
                        const groupName = input.name;
                        if (processedRadioGroups.has(groupName)) return;
                        processedRadioGroups.add(groupName);
                        totalFields++;
                        if (document.querySelector(`input[name="${groupName}"]:checked`)) {
                            filledFields++;
                        }
                    } else {
                        totalFields++;
                        let isFilled = (input.type === 'checkbox') ? input.checked : (input.value && input.value.trim() !== '');
                        if (isFilled) {
                            if (input.id === 'etp_titulo' && input.value.startsWith("Será gerado")) return;
                            filledFields++;
                        }
                    }
                }
            });
        });

        const percentage = totalFields > 0 ? (filledFields / totalFields) * 100 : 0;
        mainProgressBarEl.style.width = percentage + '%';
        mainProgressTextEl.textContent = Math.round(percentage) + '%';
    }

    function updateTabProgressBar() {
        const activeTab = document.querySelector('.tab-content.active');
        if (!activeTab || activeTab.id === 'consolidacao') {
            tabProgressWrapper.style.display = 'none';
            return;
        }

        tabProgressWrapper.style.display = 'flex';
        const progress = calculateProgress(activeTab);

        tabProgressBarEl.style.width = '0%';
        tabProgressTextEl.textContent = '0%';

        setTimeout(() => {
            tabProgressBarEl.style.width = progress.percentage + '%';
            tabProgressTextEl.textContent = Math.round(progress.percentage) + '%';
        }, 50);
    }

    function updateAllProgressBars() {
        updateMainProgressBar();
        updateTabProgressBar();
        updateAllChapterCompletionStatus();
    }

    function checkChapterCompletion(tabId) {
        const tabElement = document.getElementById(tabId);
        if (!tabElement || tabId === 'consolidacao') return false;

        const navButton = document.querySelector(`.tab-button[data-tab-target="${tabId}"]`);
        if (navButton && navButton.classList.contains('inactive')) return false;

        if (tabId === 'cap5' && document.getElementById('c5_nao_se_aplica')?.checked) return true;

        const progress = calculateProgress(tabElement);
        return progress.total > 0 && progress.percentage === 100;
    }

    function updateAllChapterCompletionStatus() {
        const tabIds = ['identificacao', 'cap1', 'cap2', 'cap3', 'cap4', 'cap5', 'cap6', 'cap7', 'cap8', 'cap9', 'anexos'];
        tabIds.forEach(tabId => {
            const isComplete = checkChapterCompletion(tabId);
            const navButton = document.querySelector(`.tab-button[data-tab-target="${tabId}"]`);
            if (navButton) {
                navButton.classList.toggle('completed', isComplete);
            }
        });
    }

    // --- LÓGICA DE IA ---

    async function makeAIRequest(prompt, temperature, max_tokens, attempt = 0) {
        const currentApiKey = aiConfig.apiKey || GROQ_API_KEY;
        if (!currentApiKey) {
            throw new Error("Chave de API não configurada. Por favor, configure sua chave no menu de Configurações de IA.");
        }
        const modelsToTry = aiConfig.modelSelectionMode === 'custom' && aiConfig.customModels.length > 0
            ? aiConfig.customModels
            : LLM_MODELS_PRIORITY;

        if (attempt >= modelsToTry.length) {
            throw new Error("Todos os modelos LLM configurados falharam ou não estão disponíveis.");
        }
        const model = modelsToTry[attempt];
        console.log(`Tentando requisição com modelo: ${model} (Tentativa ${attempt + 1})`);

        try {
            const response = await fetch(GROQ_API_URL, {
                method: 'POST',
                headers: { 'Authorization': `Bearer ${currentApiKey}`, 'Content-Type': 'application/json' },
                body: JSON.stringify({ model: model, messages: [{ role: "user", content: prompt }], temperature: temperature, max_tokens: max_tokens })
            });

            if (!response.ok) {
                const errorBody = await response.text();
                console.error(`Erro com modelo ${model}: ${response.status} ${response.statusText}`, errorBody);
                if (response.status === 429 || response.status >= 500) {
                    return makeAIRequest(prompt, temperature, max_tokens, attempt + 1);
                }
                throw new Error(`Erro ${response.status} com ${model}: ${response.statusText} - ${errorBody}`);
            }
            const data = await response.json();
            return { data, modelUsed: model };
        } catch (error) {
            console.error(`Exceção ao tentar modelo ${model}:`, error);
            if (attempt + 1 < modelsToTry.length) {
                return makeAIRequest(prompt, temperature, max_tokens, attempt + 1);
            }
            throw error;
        }
    }

    async function generateTitleWithAI() {
        const cap1_1 = document.getElementById('c1_1_indicacao_necessidade')?.value.trim();
        const cap1_2 = document.getElementById('c1_2_situacoes_problemas')?.value.trim();
        const cap1_5 = document.getElementById('c1_5_publico_alvo')?.value.trim();
        const etpUnidade = document.getElementById('etp_unidade_demandante')?.value.trim();

        if (!cap1_1 && !cap1_2) {
            alert("Preencha ao menos a 'Indicação da necessidade' ou 'Situações/problemas' no Capítulo 1 para gerar o título.");
            return;
        }

        const titleField = document.getElementById('etp_titulo');
        const generateButton = document.querySelector('button[data-action="generate-title-ai"]');
        if (generateButton) generateButton.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Gerando...';
        if (titleField) titleField.value = "Gerando título com IA...";

        // Monta o prompt usando os dados do usuário + o sufixo definido em prompts.js
        const userContext = `Unidade Demandante: "${etpUnidade || 'Não informada'}"\nIndicação da Necessidade: "${cap1_1 || 'Não preenchido'}"\nSituações/Problemas enfrentados: "${cap1_2 || 'Não preenchido'}"\nPublico Alvo: "${cap1_5 || 'Não preenchido'}"\n`;

        const titlePrompt = userContext + TITLE_GENERATION_PROMPT_SUFFIX;

        try {
            const result = await makeAIRequest(titlePrompt, 0.5, 60);
            const suggestedTitle = result.data.choices[0]?.message?.content.trim().replace(/^"|"$/g, '');
            if (titleField) {
                titleField.value = suggestedTitle || "Não foi possível gerar o título.";
            }
        } catch (error) {
            console.error("Erro ao gerar título com IA:", error);
            if (titleField) titleField.value = "Erro ao gerar título.";
            alert(`Erro ao gerar título: ${error.message}`);
        } finally {
            if (generateButton) generateButton.innerHTML = `<i class="fas fa-wand-magic-sparkles"></i> Gerar Título`;
            debouncedSave();
        }
    }

    async function getAIHelp(fieldId) {
        const field = document.getElementById(fieldId);
        const currentButton = document.querySelector(`button[data-target-field="${fieldId}"]`);
        const loadingText = '<p class="placeholder-text">Carregando...</p>';

        // Seletores dos elementos do modal
        const manualItemTitleEl = document.getElementById('manualItemTitle');
        const manualConceitoEl = document.getElementById('manualConceitoContent');
        const manualComoPreencherEl = document.getElementById('manualComoPreencherContent');
        const manualFundamentacaoEl = document.getElementById('manualFundamentacaoContent');
        const aiSectionContainer = document.getElementById('aiHelpSectionContainer');

        if (currentButton) currentButton.innerHTML = '<i class="fas fa-spinner fa-spin"></i>';

        // Reset e abertura do modal
        manualItemTitleEl.textContent = 'Carregando...';
        manualConceitoEl.innerHTML = loadingText;
        manualComoPreencherEl.innerHTML = loadingText;
        if (manualFundamentacaoEl) manualFundamentacaoEl.innerHTML = loadingText;
        aiModal.style.display = "block";

        // Identifica o título do campo
        const labelElement = document.querySelector(`label[for="${fieldId}"]`) || currentButton.closest('.form-group, .label-with-help').querySelector('label');
        const fieldLabelText = labelElement ? labelElement.textContent.trim() : 'Este campo';
        manualItemTitleEl.textContent = fieldLabelText;

        // --- 1. NORMALIZAÇÃO DE ID E CARREGAMENTO DO MANUAL ---

        // Lógica para converter IDs específicos (ex: c6_risco1_impacto) em IDs genéricos do manual (ex: c6_risco_impacto)
        let normalizedFieldId = fieldId;

        // Remove números de IDs dinâmicos para encontrar a chave no manual_content.js
        if (fieldId.includes('sol') && fieldId.includes('_')) normalizedFieldId = fieldId.replace(/sol\d+_/, 'sol_');
        else if (fieldId.includes('risco') && fieldId.includes('_')) normalizedFieldId = fieldId.replace(/risco\d+_/, 'risco_');
        else if (fieldId.includes('contratacao') && fieldId.includes('_')) normalizedFieldId = fieldId.replace(/contratacao\d+_/, 'contratacao_');
        else if (fieldId.startsWith('anexo') && fieldId.includes('_')) normalizedFieldId = fieldId.replace(/anexo\d+_/, 'anexo_');

        // Acessa o objeto global manualContent usando o ID normalizado
        const manualEntry = window.manualContent && window.manualContent[normalizedFieldId];

        let conceitoClean = "Não disponível.";
        let comoPreencherClean = "Não disponível.";
        let fundamentacaoClean = "Não disponível.";

        if (manualEntry) {
            manualConceitoEl.innerHTML = manualEntry.conceito || '<p>Conteúdo não disponível.</p>';
            manualComoPreencherEl.innerHTML = manualEntry.comoPreencher || '<p>Conteúdo não disponível.</p>';
            if (manualFundamentacaoEl) manualFundamentacaoEl.innerHTML = manualEntry.fundamentacao || '<p>Conteúdo não disponível.</p>';

            // Prepara textos para a IA (remove HTML)
            conceitoClean = stripHtml(manualEntry.conceito);
            comoPreencherClean = stripHtml(manualEntry.comoPreencher);
            fundamentacaoClean = stripHtml(manualEntry.fundamentacao);
        } else {
            const defaultText = '<p>Não foi encontrada uma orientação específica para este campo no manual.</p>';
            manualConceitoEl.innerHTML = defaultText;
            manualComoPreencherEl.innerHTML = defaultText;
            if (manualFundamentacaoEl) manualFundamentacaoEl.innerHTML = defaultText;
        }

        // --- 2. CONSULTAR A IA (MÁQUINA) ---
        if (aiConfig.useAI === 'yes') {
            aiSectionContainer.style.display = 'block';
            aiModelNameEl.textContent = "Carregando...";

            // Verifica se o campo é apropriado para análise de texto
            const isTextField = field && (field.tagName === 'TEXTAREA' || field.type === 'text' || field.tagName === 'SELECT');
            const userText = isTextField ? field.value.trim() : "";

            if (isTextField && manualEntry) {
                // Se o usuário ainda não escreveu nada, pede para preencher
                if (!userText && !field.readOnly) {
                    aiComplianceAnalysisEl.textContent = "Preencha o campo primeiro para receber uma avaliação de conformidade da IA.";
                    aiModelNameEl.textContent = "Aguardando";
                } else {
                    aiComplianceAnalysisEl.textContent = "Aguardando avaliação...";

                    // --- MONTAGEM DO PROMPT ENRIQUECIDO ---

                    // 1. Captura o contexto atual do ETP (Necessidade, Solução, etc.)
                    const etpContext = getEtpContext();

                    // 2. Injeta Manual + Contexto + Texto do Usuário no template
                    let prompt = FIELD_EVALUATION_TEMPLATE
                        .replace('{{ETP_CONTEXT}}', etpContext)
                        .replace('{{LABEL}}', fieldLabelText)
                        .replace('{{CONCEITO}}', conceitoClean)
                        .replace('{{COMO_PREENCHER}}', comoPreencherClean)
                        .replace('{{FUNDAMENTACAO}}', fundamentacaoClean)
                        .replace('{{USER_TEXT}}', userText || "[Campo vazio]");

                    try {
                        const result = await makeAIRequest(prompt, 0.3, 1500);
                        aiComplianceAnalysisEl.textContent = result.data.choices[0]?.message?.content.trim() || "Não foi possível realizar a avaliação.";
                        aiModelNameEl.textContent = result.modelUsed;
                    } catch (error) {
                        console.error("Erro na avaliação de conformidade:", error);
                        aiComplianceAnalysisEl.textContent = `Ocorreu um erro ao contatar a IA: ${error.message}`;
                        aiModelNameEl.textContent = "Falha";
                    }
                }
            } else if (!manualEntry) {
                aiComplianceAnalysisEl.textContent = "Não há dados do manual vinculados a este campo para embasar a análise da IA.";
                aiModelNameEl.textContent = "N/A";
            } else {
                aiComplianceAnalysisEl.textContent = "A avaliação de conformidade não se aplica a este tipo de campo.";
                aiModelNameEl.textContent = "N/A";
            }
        } else {
            aiSectionContainer.style.display = 'none';
        }

        if (currentButton) currentButton.innerHTML = '<i class="fas fa-book-open"></i>';
    }

    function closeAIModal() {
        if (aiModal) aiModal.style.display = "none";
    }

    function closeAboutModal() {
        const aboutModal = document.getElementById('aboutModal');
        if (aboutModal) aboutModal.style.display = "none";
    }

    function openVideoTutorialModal() {
        const modal = document.getElementById('videoTutorialModal');
        if (modal) modal.style.display = 'block';
    }

    function closeVideoTutorialModal() {
        const modal = document.getElementById('videoTutorialModal');
        if (modal) modal.style.display = 'none';
    }

    // --- GERENCIADOR DE ITENS DINÂMICOS ---
    class DynamicItemManager {
        constructor(config) {
            this.config = config;
            this.container = document.getElementById(config.containerId);
            this.count = 0;
            this.init();
        }

        init() {
            if (!this.container) {
                console.error(`Container not found for dynamic item: ${this.config.containerId}`);
            }
        }

        add() {
            this.count++;
            const newItemDiv = document.createElement('div');
            newItemDiv.className = this.config.itemClass;
            newItemDiv.setAttribute(this.config.indexAttribute, this.count);
            newItemDiv.innerHTML = this.config.template(this.count, this.container.children.length > 0 || this.count > 1);
            this.container.appendChild(newItemDiv);

            if (this.config.onAdd) {
                this.config.onAdd(newItemDiv);
            }
            updateAllProgressBars();
            applyAiFeatureState(aiConfig.useAI === 'yes');
        }

        remove(element) {
            if (confirm(this.config.removeConfirmText)) {
                element.remove();
                this.renumber();
                updateAllProgressBars();
            }
        }

        renumber() {
            const remainingItems = this.container.querySelectorAll(`.${this.config.itemClass}`);
            this.count = 0;
            remainingItems.forEach(item => {
                this.count++;
                item.setAttribute(this.config.indexAttribute, this.count);

                const h3 = item.querySelector('h3');
                let removeBtn = h3.querySelector(`.${this.config.removeClass}`);
                h3.firstChild.nodeValue = `${this.count}${this.config.titleSuffix} `;

                if (remainingItems.length > 1) {
                    if (!removeBtn) {
                        removeBtn = document.createElement('button');
                        removeBtn.className = this.config.removeClass;
                        removeBtn.innerHTML = `<i class="fas fa-trash-alt"></i> ${this.config.removeButtonText}`;
                        removeBtn.dataset.action = this.config.removeAction;
                        h3.appendChild(removeBtn);
                    }
                } else if (removeBtn) {
                    removeBtn.remove();
                }

                this.config.renumberFields(item, this.count);
            });

            if (this.count === 0) {
                this.add();
            }
            updateAllSummariesAndDropDowns();
        }

        reset() {
            this.container.innerHTML = '';
            this.count = 0;
        }
    }

    const solucaoManager = new DynamicItemManager(dynamicItemConfigs.solucao);
    const riscoManager = new DynamicItemManager(dynamicItemConfigs.risco);
    const contratacaoManager = new DynamicItemManager(dynamicItemConfigs.contratacao);
    const anexoManager = new DynamicItemManager(dynamicItemConfigs.anexo);

    // --- FUNÇÕES DE RESUMO E ATUALIZAÇÃO DE UI ---
    function createSummaryTable(headers, rows) {
        const table = document.createElement('table');
        table.className = 'summary-table';
        const thead = table.createTHead();
        const tbody = table.createTBody();
        const headerRow = thead.insertRow();
        headers.forEach(text => {
            const th = document.createElement('th');
            th.textContent = text;
            headerRow.appendChild(th);
        });
        rows.forEach(rowData => {
            const row = tbody.insertRow();
            rowData.forEach(cellData => {
                row.insertCell().textContent = cellData;
            });
        });
        return table.outerHTML;
    }

    function updateSolutionSummary() {
        const tableContainer = document.getElementById('c2_X_resumo_comparativo_table_container');
        if (!tableContainer) return;
        let rowsData = [];
        document.querySelectorAll('.solucao-item').forEach(item => {
            const descricaoEl = item.querySelector(`textarea[id$="_a_descricao_sucinta"]`);
            if (descricaoEl && descricaoEl.value.trim()) {
                rowsData.push([
                    item.getAttribute('data-solucao-index') + "ª",
                    descricaoEl.value.trim(),
                    item.querySelector(`textarea[id$="_g_custos_estimados"]`)?.value.trim() || 'N/P',
                    item.querySelector(`textarea[id$="_h_vantagens"]`)?.value.trim() || 'N/P',
                    item.querySelector(`textarea[id$="_i_desvantagens"]`)?.value.trim() || 'N/P',
                ]);
            }
        });
        if (rowsData.length === 0) {
            tableContainer.innerHTML = '<p class="placeholder-text">Nenhuma solução com título preenchido.</p>';
        } else {
            tableContainer.innerHTML = createSummaryTable(["Nº Solução", "Título da Solução", "Custos Estimados", "Vantagens", "Desvantagens"], rowsData);
        }
    }

    function populateSolucaoEscolhidaDropdown() {
        const dropdown = document.getElementById('c3_1_solucao_escolhida');
        if (!dropdown) return;
        const savedValue = dropdown.value;
        dropdown.innerHTML = '';
        let options = [];
        document.querySelectorAll('.solucao-item').forEach((sol) => {
            const descricaoEl = sol.querySelector(`textarea[id$="_a_descricao_sucinta"]`);
            if (descricaoEl && descricaoEl.value.trim()) {
                const solucaoIndex = sol.getAttribute('data-solucao-index');
                const optionText = `Solução ${solucaoIndex}: ${smartSubstring(descricaoEl.value.trim(), 70)}`;
                options.push({ value: optionText, text: optionText });
            }
        });
        if (options.length === 0) {
            dropdown.innerHTML = '<option value="">Nenhuma solução cadastrada no Cap. 2</option>';
            dropdown.disabled = true;
        } else {
            dropdown.disabled = false;
            dropdown.innerHTML = '<option value="">Selecione uma solução...</option>';
            options.forEach(opt => {
                dropdown.innerHTML += `<option value="${opt.value}">${opt.text}</option>`;
            });
        }
        if (Array.from(dropdown.options).some(opt => opt.value === savedValue)) {
            dropdown.value = savedValue;
        }
    }

    function updateRiskSummary() {
        const tableContainer = document.getElementById('c6_X_resumo_riscos_table_container');
        if (!tableContainer) return;
        let rowsData = [];
        document.querySelectorAll('.risco-item').forEach(item => {
            const identificacaoEl = item.querySelector(`textarea[id$="_identificacao"]`);
            if (identificacaoEl && identificacaoEl.value.trim()) {
                rowsData.push([
                    item.getAttribute('data-risco-index') + "º",
                    identificacaoEl.value.trim(),
                    item.querySelector(`textarea[id$="_danos_possiveis"]`)?.value.trim() || 'N/P',
                    item.querySelector(`select[id$="_probabilidade"]`)?.value || 'N/P',
                    item.querySelector(`select[id$="_impacto"]`)?.value || 'N/P',
                    item.querySelector(`textarea[id$="_acoes_preventivas"]`)?.value.trim() || 'N/P',
                    item.querySelector(`textarea[id$="_acoes_contingencia"]`)?.value.trim() || 'N/P'
                ]);
            }
        });
        if (rowsData.length === 0) {
            tableContainer.innerHTML = '<p class="placeholder-text">Nenhum risco com identificação preenchida.</p>';
        } else {
            tableContainer.innerHTML = createSummaryTable(["Nº", "Identificação", "Danos Possíveis", "Probabilidade", "Impacto", "Ações Preventivas", "Ações de Contingência"], rowsData);
        }
    }

    function updateContratacaoSummary() {
        const tableContainer = document.getElementById('c5_Z_resumo_contratacoes_table_container');
        if (!tableContainer) return;
        if (document.getElementById('c5_nao_se_aplica')?.checked) {
            tableContainer.innerHTML = '<p class="placeholder-text">Não se aplica / Primeira Contratação.</p>';
            return;
        }
        let rowsData = [];
        document.querySelectorAll('.contratacao-item').forEach(item => {
            const contratoNeEl = item.querySelector(`[id$="_contrato_ne"]`);
            if (contratoNeEl && contratoNeEl.value.trim()) {
                rowsData.push([
                    item.getAttribute('data-contratacao-index') + "ª",
                    contratoNeEl.value.trim(),
                    item.querySelector(`[id$="_processo_sei"]`)?.value.trim() || 'N/P',
                    item.querySelector(`[id$="_fase_planejamento"]`)?.value.trim() || 'N/P',
                    item.querySelector(`[id$="_fase_externa"]`)?.value.trim() || 'N/P',
                    item.querySelector(`[id$="_dificuldades_execucao"]`)?.value.trim() || 'N/P'
                ]);
            }
        });
        if (rowsData.length === 0) {
            tableContainer.innerHTML = '<p class="placeholder-text">Nenhuma contratação anterior com Contrato/NE preenchido.</p>';
        } else {
            tableContainer.innerHTML = createSummaryTable(["Nº", "Contrato/NE", "Proc. SEI", "Lições - Planejamento", "Lições - Fase Externa", "Dificuldades Execução"], rowsData);
        }
    }

    function populateAnexoItems(anexoIndex, chapterId) {
        const container = document.getElementById(`anexo${anexoIndex}_itens_relacionados_container`);
        if (!container) return;
        container.innerHTML = '';
        if (!chapterId) {
            container.innerHTML = '<small><i>Selecione um capítulo acima para ver os itens.</i></small>';
            return;
        }

        let itemsGenerated = false;

        const dynamicChapters = {
            'cap2': { itemSelector: '.solucao-item', idSelector: '[id*="_a_descricao_sucinta"]', prefix: 'Solução', indexAttr: 'data-solucao-index', emptyMsg: 'Nenhuma solução com título preenchido no Capítulo 2.' },
            'cap5': { itemSelector: '.contratacao-item', idSelector: '[id*="_contrato_ne"]', prefix: 'Contratação', indexAttr: 'data-contratacao-index', emptyMsg: 'Nenhuma contratação anterior preenchida no Capítulo 5.' },
            'cap6': { itemSelector: '.risco-item', idSelector: '[id*="_identificacao"]', prefix: 'Risco', indexAttr: 'data-risco-index', emptyMsg: 'Nenhum risco preenchido no Capítulo 6.' }
        };

        if (dynamicChapters[chapterId]) {
            const config = dynamicChapters[chapterId];
            const dynamicItems = document.querySelectorAll(config.itemSelector);
            let hasContent = false;

            dynamicItems.forEach(item => {
                const titleField = item.querySelector(config.idSelector);
                const titleText = titleField ? titleField.value.trim() : '';
                if (titleText) {
                    hasContent = true;
                    const itemIndex = item.getAttribute(config.indexAttr);
                    const checkboxId = `anexo${anexoIndex}_item_${titleField.id}`;
                    const labelText = `${config.prefix} ${itemIndex}: ${smartSubstring(titleText, 60)}`;

                    const span = document.createElement('span');
                    span.innerHTML = `<input type="checkbox" id="${checkboxId}" value="${titleField.id}"><label for="${checkboxId}">${labelText}</label>`;
                    container.appendChild(span);
                    itemsGenerated = true;
                }
            });

            if (!hasContent) {
                container.innerHTML = `<small><i>${config.emptyMsg}</i></small>`;
                return;
            }
        } else {
            const chapterElement = document.getElementById(chapterId);
            if (!chapterElement) return;

            const labels = Array.from(chapterElement.querySelectorAll('label[for]'));
            labels.forEach(label => {
                const forId = label.getAttribute('for');
                const match = label.textContent.trim().match(/^(\d+\.[\d\.]*|\w\.\d)\s/);
                if (forId && match && !label.closest('.solucao-item, .risco-item, .contratacao-item, .anexo-item')) {
                    const checkboxId = `anexo${anexoIndex}_item_${forId}`;
                    const span = document.createElement('span');
                    span.innerHTML = `<input type="checkbox" id="${checkboxId}" value="${forId}"><label for="${checkboxId}">${label.textContent.trim()}</label>`;
                    container.appendChild(span);
                    itemsGenerated = true;
                }
            });
        }

        if (!itemsGenerated) {
            container.innerHTML = '<small><i>Nenhum item numerado ou dinâmico encontrado neste capítulo.</i></small>';
        }
    }

    function updateAnexosLista() {
        const tableContainer = document.getElementById('anexos_lista_texto_table_container');
        if (!tableContainer) return;
        let rowsData = [];
        document.querySelectorAll('.anexo-item').forEach(item => {
            const tituloEl = item.querySelector(`input[id$="_titulo"]`);
            if (tituloEl && tituloEl.value.trim()) {
                const capSelectEl = item.querySelector(`select[id$="_capitulo_relacionado"]`);
                const itensSelecionados = Array.from(item.querySelectorAll('.checkbox-group input:checked')).map(cb => {
                    const label = item.querySelector(`label[for="${cb.id}"]`);
                    return label ? label.textContent.split(' ')[0] : '';
                }).join(', ');
                rowsData.push([
                    item.getAttribute('data-anexo-index') + "º",
                    tituloEl.value.trim(),
                    item.querySelector(`input[id$="_numero_sei"]`)?.value.trim() || 'N/P',
                    capSelectEl && capSelectEl.value ? capSelectEl.options[capSelectEl.selectedIndex].text : "N/P",
                    itensSelecionados || 'Nenhum'
                ]);
            }
        });
        if (rowsData.length === 0) {
            tableContainer.innerHTML = '<p class="placeholder-text">Nenhum anexo com título preenchido.</p>';
        } else {
            tableContainer.innerHTML = createSummaryTable(["Anexo Nº", "Título do Anexo", "Nº SEI", "Capítulo Relacionado", "Itens Relacionados"], rowsData);
        }
    }

    function updateAllSummariesAndDropDowns() {
        updateSolutionSummary();
        updateRiskSummary();
        updateContratacaoSummary();
        updateAnexosLista();
        populateSolucaoEscolhidaDropdown();
    }

    // --- FUNÇÕES DE PERSISTÊNCIA ---
    function saveAllDataToLocalStorage() {
        console.log("Salvamento automático acionado...");
        const etpData = {};
        const tabIds = ['identificacao', 'cap1', 'cap2', 'cap3', 'cap4', 'cap5', 'cap6', 'cap7', 'cap8', 'cap9', 'anexos'];

        tabIds.forEach(tabId => {
            etpData[tabId] = {};
            const tabElement = document.getElementById(tabId);
            if (!tabElement) return;

            tabElement.querySelectorAll('textarea, input[type="text"], input[type="number"], input[type="checkbox"], input[type="radio"], select').forEach(input => {
                if (!input.closest('.solucao-item, .risco-item, .anexo-item, .contratacao-item')) {
                    if (input.type === 'checkbox' || input.type === 'radio') {
                        etpData[tabId][input.id] = input.checked;
                    } else {
                        etpData[tabId][input.id] = input.value;
                    }
                }
            });
        });

        etpData.cap2.solucoesMercado = Array.from(document.querySelectorAll('.solucao-item')).map(item => Object.fromEntries(Array.from(item.querySelectorAll('textarea')).map(input => [input.id, input.value])));
        etpData.cap6.riscos = Array.from(document.querySelectorAll('.risco-item')).map(item => Object.fromEntries(Array.from(item.querySelectorAll('textarea, select')).map(input => [input.id, input.value])));
        etpData.cap5.contratacoesAnteriores = Array.from(document.querySelectorAll('.contratacao-item')).map(item => Object.fromEntries(Array.from(item.querySelectorAll('textarea, input')).map(input => [input.id, input.value])));
        etpData.anexos.items = Array.from(document.querySelectorAll('.anexo-item')).map(item => ({
            titulo: item.querySelector(`input[id$="_titulo"]`)?.value,
            capitulo: item.querySelector(`select[id$="_capitulo_relacionado"]`)?.value,
            numero_sei: item.querySelector(`input[id$="_numero_sei"]`)?.value,
            itens: Array.from(item.querySelectorAll('.checkbox-group input:checked')).map(cb => cb.value)
        }));

        localStorage.setItem(ETP_DATA_KEY, JSON.stringify(etpData));
        updateAllProgressBars();
        updateAllSummariesAndDropDowns();
        updateHomeTab(); // <--- NOVA LINHA ADICIONADA AQUI (Atualiza os botões do Início)
    }

    const debouncedSave = debounce(saveAllDataToLocalStorage, 750);

    function debounce(func, delay) {
        let timeout;
        return function (...args) {
            if (isAppInitializing) return;
            clearTimeout(timeout);
            timeout = setTimeout(() => func.apply(this, args), delay);
        };
    }

    function loadAllData() {
        isAppInitializing = true; // Bloqueia o salvamento automático durante a inicialização

        const etpDataString = localStorage.getItem(ETP_DATA_KEY);

        solucaoManager.reset();
        riscoManager.reset();
        contratacaoManager.reset();
        anexoManager.reset();

        if (!etpDataString) {
            document.querySelectorAll('.tab-content').forEach(tab => {
                tab.querySelectorAll('textarea, input[type="text"], input[type="number"], select').forEach(input => {
                    input.value = '';
                });
                tab.querySelectorAll('input[type="checkbox"], input[type="radio"]').forEach(input => {
                    input.checked = false;
                });
            });

            applyEtpMode('completo');
            solucaoManager.add();
            riscoManager.add();
            contratacaoManager.add();
            anexoManager.add();

            initializeConditionalFields();

            updateAllSummariesAndDropDowns();
            updateAllProgressBars();
            
            isAppInitializing = false; // Libera o salvamento automático antes do return
            return;
        }

        const etpData = JSON.parse(etpDataString);
        
        // 1. Carrega os dados das abas (Isso reconstrói os itens dinâmicos do Cap 2)
        Object.keys(etpData).forEach(tabId => {
            if (tabId !== 'consolidacao') loadTabData(tabId, etpData);
        });

        if (document.getElementById('solucoes_mercado_container').children.length === 0) solucaoManager.add();
        if (document.getElementById('riscos_container').children.length === 0) riscoManager.add();
        if (document.getElementById('contratacoes_anteriores_container').children.length === 0) contratacaoManager.add();
        if (document.getElementById('anexos_container').children.length === 0) anexoManager.add();

        // 2. Atualiza os componentes (Isso preenche as opções do Select da Solução Escolhida)
        updateAllSummariesAndDropDowns();

        // 3. --- CORREÇÃO: REAPLICAÇÃO DO VALOR DA SOLUÇÃO ESCOLHIDA ---
        // Agora que o Select já tem opções, forçamos a aplicação do valor salvo
        const savedSolucao = etpData.cap3?.c3_1_solucao_escolhida;
        if (savedSolucao) {
            const dropdown = document.getElementById('c3_1_solucao_escolhida');
            if (dropdown) dropdown.value = savedSolucao;
        }
        // -------------------------------------------------------------

        initializeConditionalFields();
        handleModalidadeChange();
        updateChapterAccess();
        updateAllProgressBars();

        isAppInitializing = false; // Libera o salvamento automático no final
    }

    function resetETPToIdentificacao() {
        // 1. Limpa o armazenamento local
        localStorage.removeItem(ETP_DATA_KEY);
        // localStorage.removeItem('etpAiConfig'); // Opcional: descomente se quiser limpar a chave de IA também
        currentEtpFilename = null;

        // 2. Reinicializa todo o estado da aplicação (DOM e variáveis)
        // Chamamos loadAllData pois ela já possui a lógica de resetar os Managers 
        // e limpar os campos caso não encontre dados no LocalStorage.
        loadAllData();

        // 3. Destrava manualmente a aba Identificação para permitir a entrada via openTab
        const btnIdentificacao = document.querySelector('.tab-button[data-tab-target="identificacao"]');
        if (btnIdentificacao) btnIdentificacao.classList.remove('inactive');

        // 4. Navega diretamente para a Identificação, pulando a aba Início
        openTab('identificacao');
        
        // 5. Atualiza os botões da aba Início (para quando o usuário voltar lá)
        updateHomeTab();
    }

    function loadTabData(tabId, allData) {
        const tabSavedData = allData[tabId];
        if (!tabSavedData) return;

        Object.keys(tabSavedData).forEach(fieldId => {
            if (!['solucoesMercado', 'riscos', 'anexos', 'contratacoesAnteriores', 'items'].includes(fieldId)) {
                const input = document.getElementById(fieldId);
                if (input) {
                    if (input.type === 'checkbox' || input.type === 'radio') {
                        input.checked = tabSavedData[fieldId];
                    } else {
                        input.value = tabSavedData[fieldId] || '';
                    }
                }
            }
        });

        if (tabId === 'cap2' && tabSavedData.solucoesMercado) {
            tabSavedData.solucoesMercado.forEach(solucaoData => {
                solucaoManager.add();
                const newItem = document.querySelector(`.solucao-item[data-solucao-index="${solucaoManager.count}"]`);
                if (newItem) Object.keys(solucaoData).forEach(solFieldId => {
                    const input = newItem.querySelector(`#${solFieldId}`);
                    if (input) input.value = solucaoData[solFieldId];
                });
            });
        }
        if (tabId === 'cap6' && tabSavedData.riscos) {
            tabSavedData.riscos.forEach(riscoData => {
                riscoManager.add();
                const newItem = document.querySelector(`.risco-item[data-risco-index="${riscoManager.count}"]`);
                if (newItem) Object.keys(riscoData).forEach(riscoFieldId => {
                    const input = newItem.querySelector(`#${riscoFieldId}`);
                    if (input) input.value = riscoData[riscoFieldId];
                });
            });
        }
        if (tabId === 'cap5' && tabSavedData.contratacoesAnteriores) {
            tabSavedData.contratacoesAnteriores.forEach(contratacaoData => {
                contratacaoManager.add();
                const newItem = document.querySelector(`.contratacao-item[data-contratacao-index="${contratacaoManager.count}"]`);
                if (newItem) Object.keys(contratacaoData).forEach(contFieldId => {
                    const input = newItem.querySelector(`#${contFieldId}`);
                    if (input) input.value = contratacaoData[contFieldId];
                });
            });
        }
        if (tabId === 'anexos' && tabSavedData.items) {
            tabSavedData.items.forEach(anexoData => {
                if (!anexoData) return;
                anexoManager.add();
                const newItem = document.querySelector(`[data-anexo-index="${anexoManager.count}"]`);
                if (newItem) {
                    newItem.querySelector(`[id$="_titulo"]`).value = anexoData.titulo || '';
                    const capSelect = newItem.querySelector(`[id$="_capitulo_relacionado"]`);
                    capSelect.value = anexoData.capitulo || '';
                    newItem.querySelector(`[id$="_numero_sei"]`).value = anexoData.numero_sei || '';

                    if (capSelect.value) {
                        populateAnexoItems(anexoManager.count, capSelect.value);
                        if (anexoData.itens) {
                            anexoData.itens.forEach(itemId => {
                                const cb = newItem.querySelector(`input[value="${itemId}"]`);
                                if (cb) cb.checked = true;
                            });
                        }
                    }
                }
            });
        }
    }

    // --- FUNÇÕES DE CONSOLIDAÇÃO E EXPORTAÇÃO ---

    function getSelectDisplayValue(selectElement) {
        if (selectElement && selectElement.tagName === 'SELECT' && selectElement.value) {
            const option = Array.from(selectElement.options).find(opt => opt.value === selectElement.value);
            return option ? option.text : selectElement.value;
        }
        return selectElement ? (selectElement.value || '') : '';
    };

    function getFieldData(input) {
        if (!input) { return { value: '[Campo não encontrado]', status: 'inactive' }; }

        if (!isFieldLogicallyActive(input)) {
            return { value: '[Campo inativo]', status: 'inactive' };
        }

        let value = '';
        let isFilled = false;

        if (input.type === 'checkbox') {
            value = input.checked ? 'Sim' : 'Não';
            isFilled = true;
        } else if (input.type === 'radio') {
            const radioGroup = document.querySelectorAll(`input[name="${input.name}"]:checked`);
            if (radioGroup.length > 0) {
                const label = document.querySelector(`label[for="${radioGroup[0].id}"]`);
                value = label ? label.textContent.trim() : 'Sim';
                isFilled = true;
            } else {
                isFilled = false;
            }
        } else if (input.tagName === 'SELECT') {
            value = getSelectDisplayValue(input);
            isFilled = !!input.value;
        } else {
            value = input.value;
            isFilled = !!(input.value && input.value.trim());
        }

        return {
            value: isFilled ? value : '[Sem resposta do usuário]',
            status: isFilled ? 'filled' : 'unfilled'
        };
    }

    function updateConsolidatedETP() {
        const container = document.getElementById('consolidation_container');
        if (!container) return;
        container.innerHTML = '<p class="placeholder-text">Gerando visualização consolidada...</p>';
        container.innerHTML = '';

        const isSimplificado = document.body.classList.contains('etp-simplificado-mode');
        const tabOrder = ['identificacao', 'cap1', 'cap2', 'cap3', 'cap4', 'cap5', 'cap6', 'cap7', 'cap8', 'cap9', 'anexos'];

        tabOrder.forEach(tabId => {
            const chapterElement = document.getElementById(tabId);
            if (!chapterElement) return;

            const navButton = document.querySelector(`.tab-button[data-tab-target="${tabId}"]`);

            // --- INÍCIO DA LÓGICA MODIFICADA ---
            const isInactive = navButton && navButton.classList.contains('inactive');

            if (isInactive && !isSimplificado) {
                return; // Comportamento original: Pula se inativo e NÃO for simplificado.
            }
            // --- FIM DA LÓGICA MODIFICADA ---

            const details = document.createElement('details');
            details.className = 'consolidation-chapter';
            details.open = true;

            const summary = document.createElement('summary');
            summary.className = 'consolidation-title';
            const h2 = chapterElement.querySelector('h2');
            const h2Clone = h2.cloneNode(true);
            h2Clone.querySelector('.chapter-icon')?.remove();
            summary.textContent = h2Clone.textContent.trim();
            details.appendChild(summary);

            const contentDiv = document.createElement('div');
            contentDiv.className = 'consolidation-content';

            // --- INÍCIO DA LÓGICA MODIFICADA ---
            if (isInactive && isSimplificado) {
                const p = document.createElement('p');
                p.innerHTML = '<i>Não se aplica ao ETP Simplificado.</i>';
                contentDiv.appendChild(p);
                details.appendChild(contentDiv);
                container.appendChild(details);
                return; // Pula para o próximo capítulo do loop
            }
            // --- FIM DA LÓGICA MODIFICADA ---

            const table = document.createElement('table');
            table.className = 'consolidation-table';
            const tbody = document.createElement('tbody');

            if (tabId === 'cap3') {
                const tr = document.createElement('tr');
                const labelText = "3.1 Solução escolhida";
                const input = isSimplificado
                    ? document.getElementById('c3_1_solucao_escolhida_simplificada')
                    : document.getElementById('c3_1_solucao_escolhida');
                const fieldData = getFieldData(input);
                tr.innerHTML = `<td>${labelText}</td><td><span class="consolidation-value-${fieldData.status}">${fieldData.value}</span></td>`;
                tbody.appendChild(tr);
            }

            if (tabId === 'cap5' && document.getElementById('c5_nao_se_aplica')?.checked) {
                const tr = document.createElement('tr');
                tr.innerHTML = `<td>Análise de Contratações Anteriores</td><td>Não se aplica / Primeira Contratação</td>`;
                tbody.appendChild(tr);
            } else {
                const formGroups = chapterElement.querySelectorAll('.form-group');
                const processedRadioGroups = new Set();

                formGroups.forEach(group => {
                    if (!isFieldLogicallyActive(group)) return;

                    const labelElement = group.querySelector('label');
                    if (!labelElement) return;

                    const firstInput = group.querySelector('input, select, textarea');
                    if (!firstInput) return;

                    let fieldId;
                    let labelText = labelElement.textContent.trim();

                    if (firstInput.type === 'radio') {
                        const groupName = firstInput.name;
                        if (processedRadioGroups.has(groupName)) return;
                        processedRadioGroups.add(groupName);
                        fieldId = document.querySelector(`input[name="${groupName}"]`)?.id;
                    } else {
                        fieldId = labelElement.getAttribute('for');
                    }

                    if (!fieldId) return;

                    if (fieldId === 'c3_1_solucao_escolhida' || fieldId === 'c3_1_solucao_escolhida_simplificada') {
                        return;
                    }

                    const input = document.getElementById(fieldId);
                    if (!input || input.closest('.solucao-item, .risco-item, .contratacao-item, .anexo-item')) return;

                    const tr = document.createElement('tr');
                    const fieldData = getFieldData(input);
                    tr.innerHTML = `<td>${labelText}</td><td><span class="consolidation-value-${fieldData.status}">${fieldData.value}</span></td>`;
                    tbody.appendChild(tr);
                });

                const dynamicConfigs = {
                    'cap2': { selector: '.solucao-item', titlePrefix: 'ª Solução de Mercado' },
                    'cap5': { selector: '.contratacao-item', titlePrefix: 'ª Contratação Anterior' },
                    'cap6': { selector: '.risco-item', titlePrefix: 'º Risco' },
                    'anexos': { selector: '.anexo-item', titlePrefix: 'º Anexo' }
                };

                if (dynamicConfigs[tabId]) {
                    const config = dynamicConfigs[tabId];
                    chapterElement.querySelectorAll(config.selector).forEach((item, index) => {
                        if (!isFieldLogicallyActive(item)) return;
                        const tr = document.createElement('tr');
                        const tdLabel = document.createElement('td');
                        tdLabel.textContent = `${index + 1}${config.titlePrefix}`;

                        const tdValue = document.createElement('td');
                        const subTable = document.createElement('table');
                        subTable.className = 'consolidation-sub-table';
                        const subTbody = document.createElement('tbody');

                        item.querySelectorAll('label[for]').forEach(subLabel => {
                            if (subLabel.closest('.scrollable-checkbox-group')) return;

                            const subFieldId = subLabel.getAttribute('for');
                            const subInput = document.getElementById(subFieldId);
                            if (subInput) {
                                const subTr = document.createElement('tr');
                                const subTdLabel = document.createElement('td');
                                subTdLabel.textContent = subLabel.textContent.trim().replace(/^(\d+\.[\d\.]*|\w\.[\d\.]+)\s/, '');

                                const subTdValue = document.createElement('td');
                                if (subInput.id.includes('itens_relacionados')) {
                                    const checkedItems = Array.from(subInput.querySelectorAll('input:checked'))
                                        .map(cb => subInput.querySelector(`label[for="${cb.id}"]`).textContent.trim())
                                        .join('; ');
                                    subTdValue.textContent = checkedItems || '[Nenhum item selecionado]';
                                } else {
                                    const fieldData = getFieldData(subInput);
                                    subTdValue.innerHTML = `<span class="consolidation-value-${fieldData.status}">${fieldData.value}</span>`;
                                }
                                subTr.appendChild(subTdLabel);
                                subTr.appendChild(subTdValue);
                                subTbody.appendChild(subTr);
                            }
                        });

                        subTable.appendChild(subTbody);
                        tdValue.appendChild(subTable);
                        tr.appendChild(tdLabel);
                        tr.appendChild(tdValue);
                        tbody.appendChild(tr);
                    });
                }
            }

            if (tbody.children.length > 0) {
                table.appendChild(tbody);
                contentDiv.appendChild(table);
                details.appendChild(contentDiv);
                container.appendChild(details);
            }
        });
    }

    function getConsolidatedDataForExport() {
        const data = [];
        const tabOrder = ['identificacao', 'cap1', 'cap2', 'cap3', 'cap4', 'cap5', 'cap6', 'cap7', 'cap8', 'cap9', 'anexos'];
        const isSimplificado = document.body.classList.contains('etp-simplificado-mode');

        const summaryTableConfigs = {
            'cap2': { id: 'c2_X_resumo_comparativo_table_container', title: 'Resumo Comparativo das Soluções' },
            'cap5': { id: 'c5_Z_resumo_contratacoes_table_container', title: 'Resumo das Contratações Anteriores' },
            'cap6': { id: 'c6_X_resumo_riscos_table_container', title: 'Resumo dos Riscos' },
            'anexos': { id: 'anexos_lista_texto_table_container', title: 'Lista de Anexos' }
        };

        tabOrder.forEach(tabId => {
            const chapterElement = document.getElementById(tabId);
            if (!chapterElement) return;

            const navButton = document.querySelector(`.tab-button[data-tab-target="${tabId}"]`);
            const isInactive = navButton && navButton.classList.contains('inactive');

            // --- INÍCIO DA LÓGICA MODIFICADA ---
            if (isInactive && !isSimplificado) {
                return;
            }
            // --- FIM DA LÓGICA MODIFICADA ---

            const h2 = chapterElement.querySelector('h2');
            const h2Clone = h2.cloneNode(true);
            h2Clone.querySelector('.chapter-icon')?.remove();
            const chapterTitle = h2Clone.textContent.trim();
            const chapterData = { title: chapterTitle, fields: [], summaryTable: null };

            // --- INÍCIO DA LÓGICA MODIFICADA ---
            if (isInactive && isSimplificado) {
                chapterData.isSkipped = true; // Adiciona um marcador
                data.push(chapterData);
                return; // Pula para o próximo capítulo do loop
            }
            // --- FIM DA LÓGICA MODIFICADA ---

            if (tabId === 'cap3') {
                const labelText = "3.1 Solução escolhida";
                const input = isSimplificado
                    ? document.getElementById('c3_1_solucao_escolhida_simplificada')
                    : document.getElementById('c3_1_solucao_escolhida');
                chapterData.fields.push({ label: labelText, ...getFieldData(input) });
            }

            if (summaryTableConfigs[tabId]) {
                const config = summaryTableConfigs[tabId];
                const tableContainer = document.getElementById(config.id);
                if (isFieldLogicallyActive(tableContainer)) {
                    const table = tableContainer?.querySelector('table');
                    if (table) {
                        const headers = Array.from(table.querySelectorAll('thead th')).map(th => th.textContent.trim());
                        const rows = Array.from(table.querySelectorAll('tbody tr')).map(tr =>
                            Array.from(tr.querySelectorAll('td')).map(td => td.textContent.trim())
                        );
                        chapterData.summaryTable = { title: config.title, headers, rows };
                    }
                }
            }

            if (tabId === 'cap5' && document.getElementById('c5_nao_se_aplica')?.checked) {
                chapterData.fields.push({ label: "Análise de Contratações Anteriores", value: "Não se aplica / Primeira Contratação" });
            } else {
                const formGroups = chapterElement.querySelectorAll('.form-group');
                const processedRadioGroups = new Set();

                formGroups.forEach(group => {
                    if (group.closest('.solucao-item, .risco-item, .contratacao-item, .anexo-item')) {
                        return;
                    }

                    if (!isFieldLogicallyActive(group)) return;

                    const labelElement = group.querySelector('label');
                    if (!labelElement) return;

                    const firstInput = group.querySelector('input, select, textarea');
                    if (!firstInput) return;

                    let fieldId;
                    let labelText = labelElement.textContent.trim();

                    if (firstInput.type === 'radio') {
                        fieldId = firstInput.name;
                        if (processedRadioGroups.has(fieldId)) return;
                        processedRadioGroups.add(fieldId);
                        const input = document.querySelector(`input[name="${fieldId}"]`);
                        chapterData.fields.push({ label: labelText, ...getFieldData(input) });
                    } else {
                        fieldId = labelElement.getAttribute('for');
                        if (fieldId === 'c3_1_solucao_escolhida' || fieldId === 'c3_1_solucao_escolhida_simplificada') {
                            return;
                        }
                        const input = document.getElementById(fieldId);
                        chapterData.fields.push({ label: labelText, ...getFieldData(input) });
                    }
                });

                const dynamicConfigs = {
                    'cap2': { selector: '.solucao-item', titlePrefix: 'ª Solução de Mercado' },
                    'cap5': { selector: '.contratacao-item', titlePrefix: 'ª Contratação Anterior' },
                    'cap6': { selector: '.risco-item', titlePrefix: 'º Risco' },
                    'anexos': { selector: '.anexo-item', titlePrefix: 'º Anexo' }
                };

                if (dynamicConfigs[tabId]) {
                    const config = dynamicConfigs[tabId];
                    chapterElement.querySelectorAll(config.selector).forEach((item, index) => {
                        if (!isFieldLogicallyActive(item)) return;
                        const subFields = [];
                        item.querySelectorAll('label[for]').forEach(subLabel => {
                            if (subLabel.closest('.scrollable-checkbox-group')) return;

                            const subInput = document.getElementById(subLabel.getAttribute('for'));
                            if (subInput) {
                                const labelText = subLabel.textContent.trim().replace(/^(\d+\.[\d\.]*|\w\.[\d\.]+)\s/, '');
                                let value;
                                if (subInput.id.includes('itens_relacionados')) {
                                    const checkedItems = Array.from(subInput.querySelectorAll('input:checked')).map(cb => subInput.querySelector(`label[for="${cb.id}"]`).textContent.trim()).join('; ');
                                    value = checkedItems || '[Nenhum item selecionado]';
                                } else {
                                    value = getFieldData(subInput).value;
                                }
                                subFields.push({ label: labelText, value });
                            }
                        });
                        chapterData.fields.push({
                            label: `${index + 1}${config.titlePrefix}`,
                            isDynamic: true,
                            subFields: subFields
                        });
                    });
                }
            }
            if (chapterData.fields.length > 0 || chapterData.summaryTable) {
                data.push(chapterData);
            }
        });
        return data;
    }

    function generateDocx(data, etpTitle, filename) {
        // --- INÍCIO DA LÓGICA MODIFICADA ---
        const isSimplificado = document.body.classList.contains('etp-simplificado-mode');
        const mainDocTitle = isSimplificado ? "Estudo Técnico Preliminar Simplificado" : "Estudo Técnico Preliminar";
        // --- FIM DA LÓGICA MODIFICADA ---

        const FONT_STYLE = { font: "Helvetica", color: "000000" };
        const CELL_MARGINS = { left: 120, right: 120, top: 100, bottom: 100 };
        const ROW_HEIGHT = { value: 450, rule: HeightRule.AT_LEAST };
        const BORDERS = {
            top: { style: BorderStyle.SINGLE, size: 1, color: "BFBFBF" },
            bottom: { style: BorderStyle.SINGLE, size: 1, color: "BFBFBF" },
            left: { style: BorderStyle.SINGLE, size: 1, color: "BFBFBF" },
            right: { style: BorderStyle.SINGLE, size: 1, color: "BFBFBF" },
            insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "BFBFBF" },
            insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "BFBFBF" },
        };

        const docChildren = [];

        docChildren.push(new Paragraph({
            alignment: "center", spacing: { after: 100 },
            children: [new TextRun({ text: mainDocTitle, size: 44, bold: true, ...FONT_STYLE })]
        }));
        docChildren.push(new Paragraph({
            alignment: "center", spacing: { after: 400 },
            children: [new TextRun({ text: etpTitle, size: 32, ...FONT_STYLE })]
        }));

        data.forEach(chapter => {
            docChildren.push(new Paragraph({
                spacing: { before: 400, after: 150 },
                children: [new TextRun({ text: chapter.title, size: 28, bold: true, ...FONT_STYLE })]
            }));

            // --- INÍCIO DA LÓGICA MODIFICADA ---
            if (chapter.isSkipped) {
                docChildren.push(new Paragraph({
                    spacing: { after: 200 },
                    children: [new TextRun({ text: "Não se aplica ao ETP Simplificado.", size: 22, italics: true, ...FONT_STYLE })]
                }));
                return; // Equivalente a 'continue' no forEach
            }
            // --- FIM DA LÓGICA MODIFICADA ---

            if (chapter.summaryTable) {
                docChildren.push(new Paragraph({
                    spacing: { after: 100 },
                    children: [new TextRun({ text: chapter.summaryTable.title, size: 24, bold: true, italics: true, ...FONT_STYLE })]
                }));

                const summaryHeader = new TableRow({
                    tableHeader: true, height: ROW_HEIGHT,
                    children: chapter.summaryTable.headers.map(headerText => new TableCell({
                        children: [new Paragraph({
                            children: [new TextRun({ text: headerText, bold: true, size: 20, ...FONT_STYLE })],
                            keepLines: true,
                        })],
                        shading: { type: "clear", fill: "EAEAEA" }, verticalAlign: VerticalAlign.CENTER, margins: CELL_MARGINS,
                    })),
                });

                const summaryRows = chapter.summaryTable.rows.map((row, index) => new TableRow({
                    height: ROW_HEIGHT,
                    shading: index % 2 !== 0 ? { type: "clear", fill: "F5F5F5" } : undefined,
                    children: row.map(cellText => new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: cellText, size: 22, ...FONT_STYLE })] })],
                        verticalAlign: VerticalAlign.CENTER, margins: CELL_MARGINS,
                    })),
                }));

                const summaryTable = new Table({ rows: [summaryHeader, ...summaryRows], width: { size: 100, type: WidthType.PERCENTAGE }, borders: BORDERS });
                docChildren.push(summaryTable);
                docChildren.push(new Paragraph({ spacing: { after: 200 } }));
            }

            const headerRow = new TableRow({
                tableHeader: true, height: ROW_HEIGHT,
                children: [
                    new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: "Campo", bold: true, size: 22, ...FONT_STYLE })] })],
                        shading: { type: "clear", fill: "EAEAEA" }, width: { size: 35, type: WidthType.PERCENTAGE }, verticalAlign: VerticalAlign.CENTER, margins: CELL_MARGINS,
                    }),
                    new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: "Valor", bold: true, size: 22, ...FONT_STYLE })] })],
                        shading: { type: "clear", fill: "EAEAEA" }, width: { size: 65, type: WidthType.PERCENTAGE }, verticalAlign: VerticalAlign.CENTER, margins: CELL_MARGINS,
                    }),
                ],
            });

            const dataRows = chapter.fields.filter(field => field.status !== 'inactive').map((field, index) => {
                const rowProperties = { height: ROW_HEIGHT, shading: index % 2 !== 0 ? { type: "clear", fill: "F5F5F5" } : undefined };
                const labelCell = new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: field.label, bold: true, size: 22, ...FONT_STYLE })] })],
                    width: { size: 35, type: WidthType.PERCENTAGE }, verticalAlign: VerticalAlign.CENTER, margins: CELL_MARGINS,
                });

                let valueCell;
                if (field.isDynamic) {
                    const subTableRows = field.subFields.map(subField => new TableRow({
                        height: ROW_HEIGHT,
                        children: [
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: subField.label, bold: true, size: 22, ...FONT_STYLE })] })],
                                width: { size: 30, type: WidthType.PERCENTAGE }, verticalAlign: VerticalAlign.CENTER, margins: CELL_MARGINS,
                            }),
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: subField.value, size: 22, ...FONT_STYLE })] })],
                                width: { size: 70, type: WidthType.PERCENTAGE }, verticalAlign: VerticalAlign.CENTER, margins: CELL_MARGINS,
                            })
                        ]
                    }));
                    const subTable = new Table({
                        rows: subTableRows, width: { size: 100, type: WidthType.PERCENTAGE },
                        borders: {
                            top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                            left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                            insideVertical: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                            insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "BFBFBF" },
                        }
                    });
                    valueCell = new TableCell({ children: [subTable], width: { size: 65, type: WidthType.PERCENTAGE }, verticalAlign: VerticalAlign.CENTER });
                } else {
                    valueCell = new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: field.value, size: 22, ...FONT_STYLE })] })],
                        width: { size: 65, type: WidthType.PERCENTAGE }, verticalAlign: VerticalAlign.CENTER, margins: CELL_MARGINS,
                    });
                }
                return new TableRow({ ...rowProperties, children: [labelCell, valueCell] });
            });

            if (dataRows.length > 0) {
                const table = new Table({ rows: [headerRow, ...dataRows], width: { size: 100, type: WidthType.PERCENTAGE }, borders: BORDERS });
                docChildren.push(table);
            }
        });

        const doc = new Document({
            sections: [{
                properties: {
                    page: {
                        margin: {
                            top: 1440,
                            bottom: 1440,
                            left: 720,
                            right: 720,
                        },
                    },
                },
                children: docChildren,
            }],
        });

        Packer.toBlob(doc).then(blob => {
            saveAs(blob, `${filename}.docx`);
        });
    }

    function generateDocxText(data, etpTitle, filename) {
        const isSimplificado = document.body.classList.contains('etp-simplificado-mode');
        const mainDocTitle = isSimplificado ? "Estudo Técnico Preliminar Simplificado" : "Estudo Técnico Preliminar";

        const FONT_STYLE = { font: "Helvetica", color: "000000" };
        const INDENT_LEVEL_1 = 400; // Recuo para respostas normais
        const INDENT_LEVEL_2 = 800; // Recuo para subitens dinâmicos

        const docChildren = [];

        // 1. Cabeçalho Principal
        docChildren.push(new Paragraph({
            alignment: "center",
            spacing: { after: 100 },
            children: [new TextRun({ text: mainDocTitle, size: 44, bold: true, ...FONT_STYLE })]
        }));
        docChildren.push(new Paragraph({
            alignment: "center",
            spacing: { after: 600 },
            children: [new TextRun({ text: etpTitle, size: 32, ...FONT_STYLE })]
        }));

        data.forEach(chapter => {
            // 2. Título do Capítulo (H1)
            docChildren.push(new Paragraph({
                spacing: { before: 400, after: 200 },
                heading: HeadingLevel.HEADING_1,
                children: [new TextRun({ text: chapter.title, size: 28, bold: true, ...FONT_STYLE })]
            }));

            // 3. Caso o capítulo não se aplique (Simplificado)
            if (chapter.isSkipped) {
                docChildren.push(new Paragraph({
                    spacing: { after: 200 },
                    children: [new TextRun({ text: "Não se aplica ao ETP Simplificado.", size: 22, italics: true, ...FONT_STYLE })]
                }));
                return;
            }

            // 4. Tabelas de Resumo (Mantidas como tabela conforme solicitado)
            if (chapter.summaryTable) {
                docChildren.push(new Paragraph({
                    spacing: { before: 200, after: 100 },
                    children: [new TextRun({ text: chapter.summaryTable.title, size: 22, bold: true, italics: true, ...FONT_STYLE })]
                }));

                const summaryHeader = new TableRow({
                    tableHeader: true,
                    children: chapter.summaryTable.headers.map(headerText => new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: headerText, bold: true, size: 18, ...FONT_STYLE })] })],
                        shading: { fill: "EAEAEA" }
                    })),
                });

                const summaryRows = chapter.summaryTable.rows.map(row => new TableRow({
                    children: row.map(cellText => new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: cellText, size: 18, ...FONT_STYLE })] })]
                    })),
                }));

                docChildren.push(new Table({
                    rows: [summaryHeader, ...summaryRows],
                    width: { size: 100, type: WidthType.PERCENTAGE }
                }));
            }

            // 5. Campos de Texto
            chapter.fields.filter(field => field.status !== 'inactive').forEach(field => {
                if (field.isDynamic) {
                    // Título do Item Dinâmico (Ex: 1ª Solução)
                    docChildren.push(new Paragraph({
                        spacing: { before: 300, after: 100 },
                        children: [new TextRun({ text: field.label, size: 24, bold: true, italics: true, ...FONT_STYLE })]
                    }));

                    // Subcampos do item dinâmico (Ex: Descrição, Custos...)
                    field.subFields.forEach(sub => {
                        docChildren.push(new Paragraph({
                            spacing: { before: 150 },
                            indent: { left: INDENT_LEVEL_1 },
                            children: [new TextRun({ text: sub.label, bold: true, size: 22, ...FONT_STYLE })]
                        }));
                        docChildren.push(new Paragraph({
                            spacing: { after: 150 },
                            indent: { left: INDENT_LEVEL_2 },
                            children: [new TextRun({ text: sub.value, size: 22, ...FONT_STYLE })]
                        }));
                    });
                } else {
                    // Campo Normal (Label acima, valor abaixo)
                    docChildren.push(new Paragraph({
                        spacing: { before: 250 },
                        children: [new TextRun({ text: field.label, bold: true, size: 22, ...FONT_STYLE })]
                    }));
                    docChildren.push(new Paragraph({
                        spacing: { after: 100 },
                        indent: { left: INDENT_LEVEL_1 },
                        children: [new TextRun({ text: field.value, size: 22, ...FONT_STYLE })]
                    }));
                }
            });
        });

        const doc = new Document({
            sections: [{
                properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
                children: docChildren,
            }],
        });

        Packer.toBlob(doc).then(blob => {
            saveAs(blob, `${filename}.docx`);
        });
    }

    function exportETP(format) {
        const data = getConsolidatedDataForExport();
        if (data.length === 0) {
            alert("Não há dados preenchidos para exportar.");
            return;
        }

        const etpTitle = document.getElementById('etp_titulo')?.value || "ETP sem título";
        const filename = etpTitle.replace(/[^\w\s.-]/gi, '').replace(/\s+/g, '_').substring(0, 50);

        if (format === 'docx') {
            // Chama o gerador original (Tabela)
            generateDocx(data, etpTitle, filename);
        } else if (format === 'docx-text') {
            // Chama o novo gerador (Texto)
            generateDocxText(data, etpTitle, filename);
        }
    }

    function downloadFile(filename, content, mimeType) {
        const element = document.createElement('a');
        element.setAttribute('href', `data:${mimeType};charset=utf-8,` + encodeURIComponent(content));
        element.setAttribute('download', filename);
        document.body.appendChild(element);
        element.click();
        document.body.removeChild(element);
    }

    // --- FUNÇÕES DE EXPORTAÇÃO DIRETA PARA A ÁREA DE TRANSFERÊNCIA (SEI) ---

    function generateHtmlForSei(data, etpTitle) {
        const isSimplificado = document.body.classList.contains('etp-simplificado-mode');
        const mainDocTitle = isSimplificado ? "Estudo Técnico Preliminar Simplificado" : "Estudo Técnico Preliminar";

        // Estilos embutidos mínimos e seguros para o SEI
        let html = `<div style="font-family: Arial, sans-serif; font-size: 11pt; color: #000000; line-height: 1.5;">`;
        
        // Títulos Principais: Usamos <p> para forçar a quebra de linha absoluta no SEI
        html += `<p style="text-align: center; margin-bottom: 5px; display: block;"><span style="font-size: 15pt; font-weight: bold; text-transform: uppercase;">${mainDocTitle}</span></p>`;
        html += `<p style="text-align: center; margin-top: 0; margin-bottom: 30px; display: block;"><span style="font-size: 13pt; font-weight: bold;">${etpTitle}</span></p>`;

        data.forEach(chapter => {
            // Títulos dos Capítulos: Usamos <p> com borda inferior
            html += `<p style="margin-top: 30px; margin-bottom: 15px; border-bottom: 2px solid #000000; padding-bottom: 5px; display: block;"><span style="font-size: 12pt; font-weight: bold; text-transform: uppercase;">${chapter.title}</span></p>`;

            if (chapter.isSkipped) {
                html += `<p style="display: block;"><i>Não se aplica ao ETP Simplificado.</i></p>`;
                return;
            }

            // Geração de Tabelas de Resumo
            if (chapter.summaryTable) {
                html += `<p style="display: block;"><b>${chapter.summaryTable.title}</b></p>`;
                html += `<table border="1" cellpadding="5" cellspacing="0" style="width: 100%; border-collapse: collapse; margin-bottom: 15px;"><thead><tr>`;
                chapter.summaryTable.headers.forEach(h => html += `<th style="background-color: #f2f2f2; text-align: left;">${h}</th>`);
                html += `</tr></thead><tbody>`;
                chapter.summaryTable.rows.forEach(row => {
                    html += `<tr>`;
                    row.forEach(cell => html += `<td>${cell}</td>`);
                    html += `</tr>`;
                });
                html += `</tbody></table>`;
            }

            // Geração dos Campos
            chapter.fields.filter(field => field.status !== 'inactive').forEach(field => {
                if (field.isDynamic) {
                    html += `<p style="margin-top: 15px; margin-bottom: 5px; display: block;"><b><i>${field.label}</i></b></p>`;
                    html += `<ul style="list-style-type: none; padding-left: 20px; margin-top: 0;">`;
                    field.subFields.forEach(sub => {
                        // Substitui quebras de linha de textarea por <br> para manter no HTML
                        const safeValue = sub.value ? sub.value.replace(/\n/g, '<br>') : '';
                        html += `<li style="margin-bottom: 10px;"><b>${sub.label}</b><br>${safeValue}</li>`;
                    });
                    html += `</ul>`;
                } else {
                    const safeValue = field.value ? field.value.replace(/\n/g, '<br>') : '';
                    html += `<p style="margin-bottom: 15px; display: block;"><b>${field.label}</b><br>${safeValue}</p>`;
                }
            });
        });

        html += `</div>`;
        return html;
    }

    async function copyETPtoClipboardSEI() {
        const data = getConsolidatedDataForExport();
        if (data.length === 0) {
            alert("Não há dados preenchidos para exportar.");
            return;
        }

        const etpTitle = document.getElementById('etp_titulo')?.value || "ETP sem título";
        const htmlContent = generateHtmlForSei(data, etpTitle);

        try {
            // A API Clipboard moderna permite gravar tipos MIME específicos.
            // Gravando como text/html, o SEI entende a formatação nativamente.
            const blobHtml = new Blob([htmlContent], { type: "text/html" });
            const blobText = new Blob([htmlContent.replace(/<[^>]*>?/gm, '')], { type: "text/plain" }); // Fallback sem formatação
            
            const clipboardItem = new ClipboardItem({
                "text/html": blobHtml,
                "text/plain": blobText,
            });

            await navigator.clipboard.write([clipboardItem]);
            
            // Exibe a mensagem de sucesso
            const toast = document.getElementById("toastNotification");
            toast.className = "toast-notification show";
            setTimeout(() => { toast.className = toast.className.replace("show", ""); }, 3900);

        } catch (err) {
            console.error('Falha ao copiar para a área de transferência:', err);
            alert("Não foi possível copiar automaticamente. Seu navegador pode estar bloqueando o acesso à área de transferência.");
        }
    }

    // --- LÓGICA DINÂMICA DO CAPÍTULO 8 ---
    function handleModalidadeChange() {
        const modalidadeSelect = document.getElementById('c8_1_modalidade_contratacao');
        const procAuxiliaresSelect = document.getElementById('c8_2_procedimentos_auxiliares');
        const criterioJulgamentoSelect = document.getElementById('c8_3_criterio_julgamento');

        if (!modalidadeSelect || !procAuxiliaresSelect || !criterioJulgamentoSelect) {
            return;
        }

        const selectedModalidade = modalidadeSelect.value;

        // Mapeamento de opções válidas para "Procedimentos Auxiliares" (8.2) - ATUALIZADO
        const procAuxiliaresOptions = {
            'licitacao': ['pre_qualificacao', 'pmi', 'registro_cadastral', 'srp', 'nao_se_aplica'],
            'dispensa': ['registro_cadastral', 'srp', 'nao_se_aplica'],
            'inexigibilidade': ['credenciamento', 'registro_cadastral', 'srp', 'nao_se_aplica'],
            'leilao': ['registro_cadastral', 'nao_se_aplica'],
            'default': ['credenciamento', 'pre_qualificacao', 'pmi', 'srp', 'registro_cadastral', 'nao_se_aplica']
        };

        // Mapeamento de opções válidas para "Critério de Julgamento" (8.3) - ATUALIZADO
        const criterioJulgamentoOptions = {
            'licitacao': ['menor_preco', 'maior_desconto', 'melhor_tecnica_ou_conteudo_artistico', 'tecnica_e_preco', 'maior_retorno_economico'],
            'dispensa': ['menor_preco', 'maior_desconto'],
            'default': ['menor_preco', 'maior_desconto', 'melhor_tecnica_ou_conteudo_artistico', 'tecnica_e_preco', 'maior_lance', 'maior_retorno_economico', 'nao_se_aplica']
        };

        // Filtrar opções para Procedimentos Auxiliares (8.2)
        const validProcAux = procAuxiliaresOptions[selectedModalidade] || procAuxiliaresOptions['default'];
        Array.from(procAuxiliaresSelect.options).forEach(option => {
            if (option.value) { // Não ocultar o placeholder "Selecione..."
                option.style.display = validProcAux.includes(option.value) ? '' : 'none';
            }
        });
        // Se a opção selecionada não for mais válida, limpa o campo e dispara o evento de mudança
        if (procAuxiliaresSelect.value && !validProcAux.includes(procAuxiliaresSelect.value)) {
            procAuxiliaresSelect.value = '';
            procAuxiliaresSelect.dispatchEvent(new Event('change', { 'bubbles': true }));
        }

        // Lógica para Critério de Julgamento (8.3)
        criterioJulgamentoSelect.disabled = false;
        if (selectedModalidade === 'inexigibilidade') {
            criterioJulgamentoSelect.value = 'nao_se_aplica';
            criterioJulgamentoSelect.disabled = true;
        } else if (selectedModalidade === 'leilao') {
            criterioJulgamentoSelect.value = 'maior_lance';
            criterioJulgamentoSelect.disabled = true;
        } else {
            const validCriterios = criterioJulgamentoOptions[selectedModalidade] || criterioJulgamentoOptions['default'];
            Array.from(criterioJulgamentoSelect.options).forEach(option => {
                if (option.value) {
                    option.style.display = validCriterios.includes(option.value) ? '' : 'none';
                }
            });
            // Se a opção selecionada não for mais válida, limpa o campo
            if (criterioJulgamentoSelect.value && !validCriterios.includes(criterioJulgamentoSelect.value)) {
                criterioJulgamentoSelect.value = '';
            }
        }
    }

    // --- LÓGICA DO MODAL DE CONFIGURAÇÕES DE IA ---

    function openAiSettingsModal() {
        aiSettingsModal.style.display = 'block';
    }

    function closeAiSettingsModal() {
        aiSettingsModal.style.display = 'none';
        testConnectionResult.style.display = 'none';
    }

    function addCustomModelInput(modelName = '') {
        const div = document.createElement('div');
        div.className = 'dynamic-model-input';

        const input = document.createElement('input');
        input.type = 'text';
        input.className = 'full-width-input';
        input.placeholder = 'Ex: llama3-70b-8192';
        input.value = modelName;

        const removeBtn = document.createElement('button');
        removeBtn.className = 'remove-model-button';
        removeBtn.innerHTML = '<i class="fas fa-times"></i>';
        removeBtn.type = 'button';
        removeBtn.addEventListener('click', () => div.remove());

        div.appendChild(input);
        div.appendChild(removeBtn);
        customModelsList.appendChild(div);
    }

    function updateAiSettingsModalUI() {
        useAiRadioYes.checked = aiConfig.useAI === 'yes';
        useAiRadioNo.checked = aiConfig.useAI === 'no';
        aiSettingsContainer.style.display = aiConfig.useAI === 'yes' ? 'block' : 'none';

        apiKeyInput.value = aiConfig.apiKey || '';

        modelModeDefault.checked = aiConfig.modelSelectionMode === 'default';
        modelModeCustom.checked = aiConfig.modelSelectionMode === 'custom';
        customModelsSection.style.display = aiConfig.modelSelectionMode === 'custom' ? 'block' : 'none';

        customModelsList.innerHTML = '';
        if (aiConfig.customModels && aiConfig.customModels.length > 0) {
            aiConfig.customModels.forEach(model => addCustomModelInput(model));
        } else {
            addCustomModelInput();
        }
    }

    function saveAiSettings() {
        const customModels = Array.from(customModelsList.querySelectorAll('input'))
            .map(input => input.value.trim())
            .filter(value => value);

        aiConfig = {
            useAI: useAiRadioYes.checked ? 'yes' : 'no',
            apiKey: apiKeyInput.value.trim(),
            modelSelectionMode: modelModeDefault.checked ? 'default' : 'custom',
            customModels: customModels
        };

        localStorage.setItem('etpAiConfig', JSON.stringify(aiConfig));
        applyAiFeatureState(aiConfig.useAI === 'yes');
        closeAiSettingsModal();
        alert("Configurações de IA salvas com sucesso!");
    }

    function loadAiSettings() {
        const savedConfig = localStorage.getItem('etpAiConfig');
        if (savedConfig) {
            aiConfig = JSON.parse(savedConfig);
        } else {
            aiConfig = {
                useAI: 'no',
                apiKey: '',
                modelSelectionMode: 'default',
                customModels: []
            };
        }
        updateAiSettingsModalUI();
        applyAiFeatureState(aiConfig.useAI === 'yes');
    }

    function applyAiFeatureState(isEnabled) {
        // Adiciona ou remove a classe no body para controlar a cor dos botões
        document.body.classList.toggle('ai-features-disabled', !isEnabled);

        const generateTitleButton = document.querySelector('[data-action="generate-title-ai"]');
        const titleField = document.getElementById('etp_titulo');
        const titleHelperText = document.getElementById('etp_titulo_ai_helper_text');

        if (generateTitleButton) {
            generateTitleButton.style.display = isEnabled ? '' : 'none';
        }

        if (titleField) {
            titleField.placeholder = isEnabled
                ? "Digite ou clique na varinha mágica para uma sugestão da IA"
                : "Digite o título identificador do ETP";
        }

        if (titleHelperText) {
            titleHelperText.style.display = isEnabled ? 'block' : 'none';
        }
    }

    async function testApiConnection() {
        const apiKey = apiKeyInput.value.trim();
        if (!apiKey) {
            alert("Por favor, insira uma chave de API para testar.");
            return;
        }

        const modelsToTest = modelModeCustom.checked
            ? Array.from(customModelsList.querySelectorAll('input')).map(input => input.value.trim()).filter(Boolean)
            : LLM_MODELS_PRIORITY;

        if (modelsToTest.length === 0) {
            alert("Nenhum modelo para testar. Adicione um modelo personalizado ou selecione a opção padrão.");
            return;
        }

        testConnectionResult.className = 'info-text testing';
        testConnectionResult.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Testando conexão...';
        testConnectionResult.style.display = 'block';
        testApiConnectionButton.disabled = true;

        const testPromises = modelsToTest.map(model =>
            fetch(GROQ_API_URL, {
                method: 'POST',
                headers: { 'Authorization': `Bearer ${apiKey}`, 'Content-Type': 'application/json' },
                body: JSON.stringify({ model: model, messages: [{ role: "user", content: "Olá" }], max_tokens: 2 })
            }).then(response => ({ model, success: response.ok, status: response.status }))
                .catch(error => ({ model, success: false, status: 'Network Error' }))
        );

        const results = await Promise.allSettled(testPromises);

        let successCount = 0;
        let resultHtml = '<strong>Resultado do Teste:</strong><ul>';

        results.forEach(result => {
            if (result.status === 'fulfilled' && result.value.success) {
                successCount++;
                resultHtml += `<li><i class="fas fa-check-circle" style="color: green;"></i> <strong>${result.value.model}:</strong> Conectado com sucesso!</li>`;
            } else {
                const status = result.status === 'fulfilled' ? result.value.status : 'Erro de Rede';
                resultHtml += `<li><i class="fas fa-times-circle" style="color: red;"></i> <strong>${result.value.model}:</strong> Falhou (Status: ${status})</li>`;
            }
        });
        resultHtml += '</ul>';

        testConnectionResult.innerHTML = resultHtml;
        testConnectionResult.className = successCount > 0 ? 'info-text success' : 'info-text error';
        testApiConnectionButton.disabled = false;
    }

    // --- INICIALIZAÇÃO E MANIPULADORES DE EVENTOS ---

    function initialize() {
        loadAiSettings();
        loadAllData();
        openTab('inicio');
        updateHomeTab();

        // Listener de cliques principal
        document.body.addEventListener('click', (e) => {
            const tabButton = e.target.closest('.tab-button[data-tab-target]');
            if (tabButton) {
                openTab(tabButton.dataset.tabTarget);
                return;
            }

            const actionTarget = e.target.closest('[data-action]');
            if (!actionTarget) return;

            const action = actionTarget.dataset.action;

            // --- Ações para os botões da aba Início ---
            
            // Botão: Continuar ETP Salvo
            if (action === 'continue-etp-home') {
                openTab('identificacao');
                return;
            }

            // Botão: Começar Novo ETP
            if (action === 'start-new-etp-home') {
                const etpDataString = localStorage.getItem(ETP_DATA_KEY);
                if (etpDataString) {
                    // Se já existir dado, usa a função de reset inteligente
                    if (confirm("Você já tem um ETP em andamento. Tem certeza que deseja iniciar um novo? Todo o progresso atual será perdido.")) {
                        resetETPToIdentificacao();
                    }
                } else {
                    // Se o sistema já estiver vazio, apenas destrava a aba e navega
                    const btnIdentificacao = document.querySelector('.tab-button[data-tab-target="identificacao"]');
                    if (btnIdentificacao) btnIdentificacao.classList.remove('inactive');
                    openTab('identificacao');
                }
                return;
            }

            // Ações para os botões secundários da aba Início
            if (actionTarget.id === 'btnAbrirTutorial_home') {
                openVideoTutorialModal();
                return;
            }
            if (actionTarget.id === 'btnAbrirManualEtp_home') {
                window.open('manual_etp.html', '_blank');
                return;
            }
            if (actionTarget.id === 'btnAbrirSobre_home') {
                const aboutModal = document.getElementById('aboutModal');
                if (aboutModal) aboutModal.style.display = 'block';
                return;
            }

            switch (action) {
                case 'generate-title-ai': generateTitleWithAI(); break;
                case 'get-ai-help': getAIHelp(actionTarget.dataset.targetField); break;
                case 'close-ai-modal': closeAIModal(); break;
                case 'add-solucao': solucaoManager.add(); break;
                case 'remove-solucao': solucaoManager.remove(actionTarget.closest('.solucao-item')); break;
                case 'add-risco': riscoManager.add(); break;
                case 'remove-risco': riscoManager.remove(actionTarget.closest('.risco-item')); break;
                case 'add-contratacao': contratacaoManager.add(); break;
                case 'remove-contratacao': contratacaoManager.remove(actionTarget.closest('.contratacao-item')); break;
                case 'add-anexo': anexoManager.add(); break;
                case 'remove-anexo': anexoManager.remove(actionTarget.closest('.anexo-item')); break;
                case 'close-ai-settings-modal': closeAiSettingsModal(); break;
                case 'close-about-modal': closeAboutModal(); break;
                case 'close-video-modal': closeVideoTutorialModal(); break;
                case 'add-anexo': anexoManager.add(); break;
                case 'remove-anexo': anexoManager.remove(actionTarget.closest('.anexo-item')); break;
                case 'copy-to-sei': copyETPtoClipboardSEI(); break;
                case 'export-etp': exportETP(actionTarget.dataset.format); break;
                case 'next-chapter': navigateToChapter('next'); break;
                case 'prev-chapter': navigateToChapter('prev'); break;
            }
        });

        // Listener de 'change' principal
        document.body.addEventListener('change', (e) => {
            const target = e.target;
            const action = target.dataset.action;

            if (action === 'toggle-conditional') {
                toggleConditionalFields(target.dataset.toggleGroup, target.value);
            } else if (action === 'toggle-cap5') {
                toggleCapNFields('cap5_fields_container', target.checked);
            } else if (action === 'populate-anexo-items') {
                populateAnexoItems(target.dataset.anexoIndex, target.value);
            } else if (target.id === 'c8_2_procedimentos_auxiliares') {
                updateChapterAccess();
            }
        });

        // Listeners de 'input' e 'change' para salvamento automático
        const tabContainer = document.querySelector('.tab-container');
        tabContainer.addEventListener('input', (e) => {
            debouncedSave();
            if (e.target.matches('[data-action="toggle-vigencia-benefits"]')) {
                toggleVigenciaBenefitsField(e.target);
            }
        });
        tabContainer.addEventListener('change', debouncedSave);
        tabContainer.addEventListener('click', (e) => {
            if (e.target.closest('.remove-solucao-button, .remove-risco-button, .remove-anexo-button, .remove-contratacao-button, .add-button')) {
                debouncedSave();
            }
        });

        // Listeners da barra de ferramentas
        document.getElementById('btnNovoETP').addEventListener('click', () => {
            if (confirm("Tem certeza que deseja iniciar um novo ETP? Todo o progresso atual será perdido.")) {
                resetETPToIdentificacao();
            }
        });
        const etpFileInput = document.getElementById('etpFileInput');
        document.getElementById('btnAbrirETP').addEventListener('click', () => {
            if (localStorage.getItem(ETP_DATA_KEY) && !confirm("Abrir um arquivo irá sobrescrever o progresso atual não salvo. Deseja continuar?")) return;
            etpFileInput.click();
        });
        etpFileInput.addEventListener('change', (event) => {
            const file = event.target.files[0];
            if (file) {
                currentEtpFilename = file.name;
                const reader = new FileReader();
                reader.onload = (e) => {
                    try {
                        const importedData = JSON.parse(e.target.result);
                        if (typeof importedData === 'object' && importedData !== null && Object.keys(importedData).length > 0) {
                            localStorage.setItem(ETP_DATA_KEY, e.target.result);
                            loadAllData();
                            alert("ETP carregado com sucesso!");
                        } else { throw new Error("Arquivo JSON inválido ou vazio."); }
                    } catch (error) { alert(`Erro ao carregar o arquivo: ${error.message}`); }
                };
                reader.readAsText(file);
            }
        });
        document.getElementById('btnSalvarETPComputador').addEventListener('click', () => {
            const etpDataString = localStorage.getItem(ETP_DATA_KEY);
            if (!etpDataString) {
                alert("Nenhum dado salvo para exportar.");
                return;
            }
            let filename;
            if (currentEtpFilename) {
                filename = currentEtpFilename;
            } else {
                const title = document.getElementById('etp_titulo')?.value || "ETP_sem_titulo";
                const sanitizedTitle = title.replace(/[^\w\s.-]/gi, '').replace(/\s+/g, '_').substring(0, 50);
                filename = `${sanitizedTitle}.etp.json`;
            }
            downloadFile(filename, etpDataString, "application/json");
            currentEtpFilename = filename;
        });

        // Listener para fechar o banner de alerta beta
        const betaBanner = document.getElementById('beta-alert-banner');
        const closeBetaBannerBtn = document.getElementById('close-beta-alert');

        if (betaBanner && closeBetaBannerBtn) {
            closeBetaBannerBtn.addEventListener('click', () => {
                betaBanner.style.display = 'none';
            });
        }

        // Listener para formatação do campo SEI
        document.getElementById('etp_numero_processo_sei').addEventListener('input', (e) => formatSEIProcess(e.target));

        // Listener para fechar modais
        window.onclick = function (event) {
            const aboutModal = document.getElementById('aboutModal');
            const videoModal = document.getElementById('videoTutorialModal');
            if (event.target == aiModal) {
                closeAIModal();
            }
            if (event.target == aiSettingsModal) {
                closeAiSettingsModal();
            }
            if (event.target == aboutModal) {
                closeAboutModal();
            }
            if (event.target == videoModal) {
                closeVideoTutorialModal();
            }
        }

        // --- INÍCIO: Listeners do Modal de Configurações de IA ---
        document.getElementById('btnAbrirConfigIA').addEventListener('click', openAiSettingsModal);
        document.getElementById('saveAiSettingsButton').addEventListener('click', saveAiSettings);

        useAiRadioYes.addEventListener('change', () => {
            aiSettingsContainer.style.display = 'block';
        });
        useAiRadioNo.addEventListener('change', () => {
            aiSettingsContainer.style.display = 'none';
        });

        modelModeDefault.addEventListener('change', () => {
            customModelsSection.style.display = 'none';
        });
        modelModeCustom.addEventListener('change', () => {
            customModelsSection.style.display = 'block';
        });

        addModelButton.addEventListener('click', () => addCustomModelInput());
        testApiConnectionButton.addEventListener('click', testApiConnection);

        const toggleButton = document.getElementById('toggleApiKeyVisibility');
        if (toggleButton) {
            toggleButton.addEventListener('click', () => {
                const apiKeyInput = document.getElementById('apiKeyInput');
                const isPassword = apiKeyInput.type === 'password';
                apiKeyInput.type = isPassword ? 'text' : 'password';
                toggleButton.classList.toggle('fa-eye', !isPassword);
                toggleButton.classList.toggle('fa-eye-slash', isPassword);
            });
        }
        // --- FIM: Listeners do Modal de Configurações de IA ---

        // --- INÍCIO: Listeners de Recursos e Ajuda ---
        document.getElementById('btnAbrirSobre').addEventListener('click', () => {
            const aboutModal = document.getElementById('aboutModal');
            if (aboutModal) aboutModal.style.display = 'block';
        });

        // Botão do Guia Rápido (agora dentro do modal de tutorial)
        document.getElementById('btnGuiaRapidoModal').addEventListener('click', () => {
            window.open('guia_rapido.html', '_blank');
        });

        // Botão do Manual (na barra de ferramentas principal)
        document.getElementById('btnAbrirManualEtp').addEventListener('click', () => {
            window.open('manual_etp.html', '_blank');
        });

        // Botão do Tutorial (na barra de ferramentas principal)
        document.getElementById('btnAbrirTutorial').addEventListener('click', openVideoTutorialModal);

        // Botão de documentação técnica (sem elemento visual na UI principal, mas mantido para o modal "Sobre")
        const btnDocTecnica = document.getElementById('btnAbrirDocTecnica');
        if (btnDocTecnica) {
            btnDocTecnica.addEventListener('click', () => {
                window.open('documentacao.html', '_blank');
            });
        }
        // --- FIM: Listeners de Recursos e Ajuda ---

        // --- INÍCIO: Listeners para controle de acesso aos capítulos ---
        document.getElementById('etp_tipo')?.addEventListener('change', updateChapterAccess);
        document.getElementById('etp_auth')?.addEventListener('change', updateChapterAccess);
        document.getElementById('etp_auth_sei_number')?.addEventListener('input', updateChapterAccess);
        // --- FIM: Listeners para controle de acesso aos capítulos ---

        // --- INÍCIO: Listener para lógica dinâmica do Capítulo 8 ---
        document.getElementById('c8_1_modalidade_contratacao').addEventListener('change', handleModalidadeChange);
        // --- FIM: Listener para lógica dinâmica do Capítulo 8 ---

        updateChapterAccess();
    }

    initialize();
});