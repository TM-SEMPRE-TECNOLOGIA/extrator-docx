# Design System TM - Ocean Breeze v2.0

Sistema de design completo e moderno baseado em páginas, com navegação no header e tema claro/escuro.

## 📁 Estrutura do Projeto

```
ocean-breeze-pages/
├── pages/
│   ├── index.html           # Home / Visão geral
│   ├── tokens.html          # Cores, tipografia, sombras, radius, spacing
│   ├── components.html      # Botões, cards, inputs, alerts, etc.
│   ├── dashboards.html      # Dashboards e gráficos (14 métricas + 8 gráficos)
│   ├── patterns.html        # Padrões de telas (sidebar, chat, blog)
│   ├── ux.html              # UX guidelines
│   ├── accessibility.html   # Acessibilidade (WCAG 2.1)
│   ├── content.html         # Conteúdo & microcopy
│   └── mobile.html          # Mobile Preview (light/dark)
└── assets/
    ├── styles.css           # CSS consolidado com todos os tokens
    └── app.js               # JS: theme toggle + navegação
```

## ✨ Características Principais

### 🎨 Design System Completo
- **Tokens Centralizados**: Todas as variáveis CSS em um único lugar
- **Modo Escuro Nativo**: Suporte completo com persistência de preferência
- **Responsivo**: Layouts que se adaptam a qualquer tamanho de tela
- **CSS Puro**: Sem dependências externas

### 📊 Dashboard Enriquecido
A página `dashboards.html` apresenta:
- **14+ Métricas**: Receita, Conversão, Tickets, SLA, Backlog, Tempo médio, NPS, Churn, CAC, LTV, Custo operacional, Usuários ativos, ROI, Throughput
- **8 Variações de Gráficos**:
  1. Line Chart (tendência)
  2. Area Chart (crescimento)
  3. Bar Chart (volume)
  4. Stacked Bar (distribuição)
  5. Donut Chart (mix de canais)
  6. Gauge (medidor de SLA)
  7. Heatmap (atividade semanal)
  8. Tabela de KPIs (performance por canal)
- **3 Layouts Completos**:
  - Executive Overview
  - Operations Dashboard
  - Marketing Dashboard
- **Interações**: Tabs para períodos (7d/30d/90d) e toggle de densidade

### 📱 Mobile Preview
A página `mobile.html` demonstra:
- Mock de smartphone com moldura
- Tema claro e escuro lado a lado
- Header, stats, gráficos, botões, inputs, alerts e listas
- Navegação bottom (inferior)
- Layout responsivo real

## 🚀 Como Usar

### 1. Abrir o Projeto
Abra qualquer arquivo HTML da pasta `pages/` em um navegador moderno.

### 2. Navegação
Use o header fixo para navegar entre as páginas:
- **Home**: Visão geral do sistema
- **Tokens**: Cores, tipografia, sombras, etc.
- **Componentes**: Biblioteca de componentes
- **Dashboards**: Métricas e gráficos
- **Padrões**: Layouts prontos
- **UX**: Diretrizes de experiência
- **Acessibilidade**: WCAG 2.1
- **Conteúdo**: Guia de escrita
- **Mobile**: Preview mobile

### 3. Alternar Tema
Clique no botão "🌓 Alternar Tema" no header. A preferência é salva automaticamente.

## 🎨 Tokens de Design

### Cores
```css
--TM-primary: #22c55e (verde)
--TM-secondary: #e0f2fe (azul claro)
--TM-accent: #d1fae5 (verde claro)
--TM-destructive: #ef4444 (vermelho)
```

### Tipografia
```css
--TM-font-sans: "DM Sans", sans-serif
--TM-font-serif: "Lora", serif
--TM-font-mono: "IBM Plex Mono", monospace
```

### Sombras
```css
--TM-shadow-sm: 0px 4px 8px -1px rgba(0,0,0,0.10)...
--TM-shadow-md: ...
--TM-shadow-lg: ...
```

### Raios de Borda
```css
--TM-radius-sm: calc(0.5rem - 4px)
--TM-radius-md: calc(0.5rem - 2px)
--TM-radius-lg: 0.5rem
--TM-radius-xl: calc(0.5rem + 4px)
```

## 📋 Checklist de Implementação

- [x] Não existe mais "página única" (múltiplos HTMLs)
- [x] Header fixo com navegação e item ativo
- [x] Theme toggle funciona e persiste
- [x] dashboards.html possui 14+ métricas e 8+ variações de gráficos
- [x] mobile.html apresenta mock mobile completo (light/dark)
- [x] CSS centralizado e reutilizado
- [x] Textos em português correto com acentuação

## 🎯 Páginas Principais

### Dashboard (dashboards.html)
A página mais rica do sistema, com:
- 14 métricas distintas com deltas e sparklines
- 8 tipos de gráficos diferentes
- 3 seções completas (Executive, Operations, Marketing)
- Tabs para filtros de período
- Toggle de densidade de layout

### Mobile (mobile.html)
Demonstração completa em dispositivos móveis:
- Mockup de smartphone com tema claro
- Mockup de smartphone com tema escuro
- Todos os componentes principais
- Navegação bottom
- Diretrizes mobile

## 🌐 Compatibilidade

- ✅ Chrome/Edge (últimas versões)
- ✅ Firefox (últimas versões)
- ✅ Safari (últimas versões)
- ✅ Dispositivos móveis (iOS/Android)

## 📝 Notas Técnicas

- **Sem Backend**: Projeto 100% frontend
- **Sem Dependências**: CSS e JS vanilla
- **Persistência**: localStorage para tema
- **Acessibilidade**: WCAG 2.1 AA
- **SEO**: Meta tags e estrutura semântica

## 🎨 Paleta Ocean Breeze

O tema "Ocean Breeze" é inspirado em tons de oceano e natureza:
- Verde vibrante como cor primária
- Azuis suaves para backgrounds secundários
- Tons neutros para texto e cards
- Modo escuro com azuis profundos

## 📄 Licença

Design System TM - Ocean Breeze v2.0
Todos os direitos reservados.

---

**Versão**: 2.0  
**Data**: Janeiro 2026  
**Autor**: Design System TM
