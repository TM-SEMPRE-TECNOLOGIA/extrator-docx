# Planejamento de Redesign Visual e UX (Equipe Design v3)

**Orquestrador:** Antigravity
**Objetivo:** Simplificar a aplicação Extrator DOCX para parecer um "assistente simples", focando em hierarquia, copy e organização, conforme diretrizes da equipe Design v3.

---

## 1. Análise UX Strategist
**Diagnóstico:** A tela atual expõe todo o processo de uma vez ("1) Enviar", "2) Resumo", "3) Somar", "4) Prévia"). Isso cria carga cognitiva desnecessária. O usuário só quer processar o arquivo. O resto é consequência.

**Proposta de Fluxo:**
1.  **Estado Inicial:** Apenas o Card de Upload (Intro + Dropzone). Nada mais. "Detalhes técnicos" deve sumir ou ficar muito discreto no rodapé.
2.  **Estado Processando:** Feedback visual claro no card central.
3.  **Estado Concluído (Sucesso):** O Card de Upload se transforma (ou dá lugar) ao Card de Resultados.
    *   Mostra sucesso.
    *   Mostra botões de ação principais (Baixar Excel, Baixar Log).
    *   Mostra a opção de "Agregação" como um passo seguinte natural, não uma "seção 3" desconectada.
    *   "Prévia dos itens" e "Estatísticas" ficam abaixo, como informações secundárias.

---

## 2. Layout Architect
**Diagnóstico:** Atualmente há muitas caixas (`tm-panel`) competindo por atenção. A numeração "1)", "2)" reforça a sensação de "formulário burocrático".

**Nova Estrutura:**
*   **Header:** Mantém (Logo + Toggle Tema + Status). Limpo.
*   **Área Principal (Container Centralizado):**
    *   **Hero Section (Texto):** Título grande e convidativo ("Extraia itens do seu DOCX em segundos"). Subtítulo explicativo curto.
    *   **Action Card (Upload):** Área de dropzone grande, clean.
*   **Área de Resultados (Só aparece após processamento):**
    *   **Card de Sucesso:** "8 itens encontrados". Destaque para o número.
    *   **Ações Primárias:** Botões grandes de download lado a lado.
    *   **Ferramentas Extras (Agregação):** Um bloco separado, visualmente distinto (ex: fundo cinza claro ou outline), para não confundir com o download principal.
    *   **Preview:** Tabela simples, sem ser um "accordion" complexo se não for necessário. Talvez apenas os 5 primeiros itens.

**Diretriz:** Remover numeração ("1)", "2)", etc.) dos títulos. Usar tamanho de fonte e espaçamento para hierarquia.

---

## 3. UX Writer
**Diagnóstico:** Textos muito técnicos ("Enviar e processar", "Excel bruto", "Excel consolidado", "Quantidade_raw").

**Tabela de Mudanças (De -> Para):**

| Onde | Antes | Depois |
| :--- | :--- | :--- |
| **Título Principal** | "1) Enviar e processar" | "Upload do Arquivo" (ou removido, direto no Hero) |
| **Dropzone** | "Arraste o arquivo aqui ou clique..." | "Arraste seu DOCX aqui" |
| **Botão Processar** | "Processar documento" | "Extrair Itens" |
| **Botão Selecionar** | "Selecionar outro" | "Novo Arquivo" |
| **Seção 2** | "2) Resumo" | "Resumo da Extração" |
| **Seção 3** | "3) Somar itens iguais" | "Consolidar Itens Repetidos" |
| **Botão Soma** | "Gerar planilha somada" | "Agrupar e Baixar" |
| **Seção 4** | "4) Prévia dos itens extraidos" | "Prévia dos Resultados" |
| **Footer** | "Detalhes tecnicos" | "Mais informações" |

---

## 4. UI Designer (Refinamento Ocean Breeze)
**Diagnóstico:** O Ocean Breeze já está aplicado, mas podemos limpar o ruído.

**Ajustes:**
*   **Gradientes:** Manter o gradiente de fundo, mas garantir que o conteúdo principal (Cards) tenha fundo sólido (Branco/Cinza Escuro) para leitura máxima.
*   **Sombras:** Usar sombras mais difusas (`shadow-lg`) para o card principal flutuar, e `shadow-sm` para elementos secundários.
*   **Espaçamento:** Aumentar `gap` entre o título Hero e o Card de Upload.
*   **Ícones:** Usar ícones maiores e com `stroke-width` mais fino para elegância.

---

## 5. Plano de Ação (Para o Developer)

1.  **Refatorar `App.jsx`:**
    *   Remover títulos numerados ("1)", "2)").
    *   Criar um componente `Hero` simples para o título da página.
    *   Condicionar a renderização: Se `phase === 'idle'`, mostrar APENAS Hero + Upload.
    *   Se `phase === 'ok'`, mostrar Hero (menor) + Upload (contraído/botão "Novo") + Resultados.
    *   Reescrever os textos conforme UX Writer.
    *   Agrupar "Resumo", "Preview" e "Agregação" de forma mais inteligente visualmente.

2.  **Limpeza Visual:**
    *   Remover o `<details>` de "Detalhes técnicos" ou movê-lo para um footer link discreto.
    *   Simplificar a "Regras e privacidade" para um tooltip ou link "Como funciona".

3.  **Execução:**
    *   Editar `App.jsx` aplicando a nova estrutura.
    *   Verificar se o CSS `tm-` precisa de ajustes de espaçamento.

---
*Aprovado pela equipe Design v3.*
