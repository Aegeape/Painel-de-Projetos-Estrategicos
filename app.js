/**
 * PAINEL EXECUTIVO DE PROJETOS V2 — APE
 * Motor Executivo com Drill-Through, Cards Visuais e Suporte à Planilha em Aba Única (Projetos)
 */

(function () {
  'use strict';

  // State Management
  const state = {
    data: window.PANEL_DATA || {
      metadata: {},
      projetos: [],
      atualizacoes: [],
      entregas: [],
      pontosAtencao: [],
      impedimentos: [],
      proximosPassos: [],
      riscos: [],
      decisoes: [],
      listas: []
    },
    filters: {
      area: 'ALL',
      search: '',
      quickFilter: 'andamento'
    },
    currentView: 'overview', // 'overview' | 'detail'
    selectedProjectId: null
  };

  // ==========================================================================
  // FIELD ACCESSOR HELPER (Insensível a maiúsculas/minúsculas e acentuação)
  // ==========================================================================
  function normalizeKey(str) {
    return (str || '')
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-z0-9]/g, '');
  }

  function getField(obj, ...possibleKeys) {
    if (!obj || typeof obj !== 'object') return '';

    // Direct key matches first
    for (const k of possibleKeys) {
      if (obj[k] !== undefined && obj[k] !== null && String(obj[k]).trim() !== '') {
        return obj[k];
      }
    }

    // Normalized matches
    const normTargets = possibleKeys.map(k => normalizeKey(k));
    for (const prop in obj) {
      if (Object.prototype.hasOwnProperty.call(obj, prop)) {
        const normProp = normalizeKey(prop);
        if (normTargets.includes(normProp)) {
          if (obj[prop] !== undefined && obj[prop] !== null && String(obj[prop]).trim() !== '') {
            return obj[prop];
          }
        }
      }
    }

    return '';
  }

  // ==========================================================================
  // DATE PARSING & FORMATTING UTILITIES (Padrão BR: dd/mm/aaaa)
  // ==========================================================================
  function parseDate(val) {
    if (!val) return null;
    if (val instanceof Date) {
      return isNaN(val.getTime()) ? null : val;
    }
    if (typeof val === 'number') {
      if (val > 20000 && val < 80000) {
        const date = new Date((val - 25569) * 86400 * 1000);
        return isNaN(date.getTime()) ? null : new Date(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate());
      }
      return null;
    }
    if (typeof val !== 'string') return null;
    const trimmed = val.trim();
    if (!trimmed || trimmed === '—' || trimmed === '-' || trimmed === 'undefined' || trimmed === 'null') return null;

    // Excel serial number as string of digits
    if (/^\d{5}(\.\d+)?$/.test(trimmed)) {
      const num = parseFloat(trimmed);
      const date = new Date((num - 25569) * 86400 * 1000);
      return isNaN(date.getTime()) ? null : new Date(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate());
    }

    // Format DD/MM/YYYY or DD-MM-YYYY
    const brMatch = trimmed.match(/^(\d{1,2})[\/\.-](\d{1,2})[\/\.-](\d{2,4})/);
    if (brMatch) {
      const day = parseInt(brMatch[1], 10);
      const month = parseInt(brMatch[2], 10) - 1;
      let year = parseInt(brMatch[3], 10);
      if (year < 100) year += 2000;
      const d = new Date(year, month, day);
      return isNaN(d.getTime()) ? null : d;
    }

    // Format YYYY-MM-DD (ISO)
    const isoMatch = trimmed.match(/^(\d{4})[\/\.-](\d{1,2})[\/\.-](\d{1,2})/);
    if (isoMatch) {
      const year = parseInt(isoMatch[1], 10);
      const month = parseInt(isoMatch[2], 10) - 1;
      const day = parseInt(isoMatch[3], 10);
      const d = new Date(year, month, day);
      return isNaN(d.getTime()) ? null : d;
    }

    const nativeDate = new Date(trimmed);
    return isNaN(nativeDate.getTime()) ? null : nativeDate;
  }

  function formatDateBr(val) {
    if (!val) return '—';
    const date = parseDate(val);
    if (!date || isNaN(date.getTime())) {
      return typeof val === 'string' && val.trim() && val.trim() !== '-' ? val.trim() : '—';
    }
    const d = String(date.getDate()).padStart(2, '0');
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const y = date.getFullYear();
    return `${d}/${m}/${y}`;
  }

  function parsePercentage(val) {
    if (val === null || val === undefined || val === '') return 0;
    if (typeof val === 'number') {
      if (val > 0 && val <= 1) return Math.round(val * 100);
      return Math.round(val);
    }
    const str = String(val).replace('%', '').trim();
    const num = parseFloat(str.replace(',', '.'));
    if (isNaN(num)) return 0;
    if (num > 0 && num <= 1 && str.includes('.')) return Math.round(num * 100);
    return Math.round(num);
  }

  function calculateProjectSchedule(project) {
    const dataInicio = parseDate(getField(project, 'Data Início', 'Data Inicio', 'Inicio'));
    const dataPlan = parseDate(getField(project, 'Data Fim Planejada', 'Data Fim Planejado', 'Fim Planejado'));
    const dataAtual = parseDate(getField(project, 'Data Fim Atual', 'Fim Atual', 'Previsão', 'Data Fim Prevista', 'Fim Previsto'));

    let delayDays = 0;
    let delayPct = null;
    let isDelayed = false;
    let delayLabel = '0 dias (0%)';
    let delayClass = 'delay-ontime';

    // Consumo do Prazo (%) = da tabela ou (Data do dia - Data Início) / (Fim Planejado - Data Início) * 100
    let consumoPct = null;
    let consumoLabel = '—';
    const rawConsumo = getField(project, '% Consumo Prazo', '% Consumo do Prazo', 'Consumo do Prazo', 'Consumo Prazo', '%Consumo Prazo', '%Consumo \nPrazo', '%Consumo \r\nPrazo', 'Consumo');

    if (rawConsumo !== '' && rawConsumo !== null && rawConsumo !== undefined) {
      consumoPct = parsePercentage(rawConsumo);
      consumoLabel = `${consumoPct}%`;
    } else {
      const refDate = (state.data.metadata && state.data.metadata.referenceDate)
        ? parseDate(state.data.metadata.referenceDate)
        : new Date();

      if (dataInicio && dataPlan) {
        const startMs = new Date(dataInicio.getFullYear(), dataInicio.getMonth(), dataInicio.getDate()).getTime();
        const endMs = new Date(dataPlan.getFullYear(), dataPlan.getMonth(), dataPlan.getDate()).getTime();
        const today = refDate ? new Date(refDate.getFullYear(), refDate.getMonth(), refDate.getDate()) : new Date();
        today.setHours(0, 0, 0, 0);
        const todayMs = today.getTime();

        const totalMs = endMs - startMs;
        const elapsedMs = todayMs - startMs;

        if (totalMs > 0) {
          if (elapsedMs <= 0) {
            consumoPct = 0;
          } else {
            consumoPct = Math.min(Math.round((elapsedMs / totalMs) * 100), 100);
          }
          consumoLabel = `${consumoPct}%`;
        } else if (totalMs === 0) {
          consumoPct = todayMs >= endMs ? 100 : 0;
          consumoLabel = `${consumoPct}%`;
        }
      }
    }

    if (dataPlan && dataAtual) {
      // O desvio do cronograma (em dias) é a diferença entre Fim Atual e Fim Planejado
      const diffMs = dataAtual.getTime() - dataPlan.getTime();
      delayDays = Math.round(diffMs / (1000 * 60 * 60 * 24));

      // Percentual de Desvio = (Fim Atual - Fim Planejado) / (Fim Planejado - Data de Início) * 100
      let pctStr = '';
      if (dataInicio && dataPlan.getTime() > dataInicio.getTime()) {
        const plannedDurationMs = dataPlan.getTime() - dataInicio.getTime();
        const rawPct = (diffMs / plannedDurationMs) * 100;
        const roundedPct = Math.round(rawPct);
        const sign = roundedPct > 0 ? '+' : '';
        pctStr = ` (${sign}${roundedPct}%)`;
        delayPct = roundedPct;
      }

      if (delayDays > 0) {
        isDelayed = true;
        delayLabel = `+${delayDays} dias${pctStr}`;
        delayClass = 'delay-problem';
      } else if (delayDays < 0) {
        delayLabel = `${delayDays} dias${pctStr}`;
        delayClass = 'delay-ahead';
      } else {
        delayLabel = `0 dias${pctStr || ' (0%)'}`;
        delayClass = 'delay-ontime';
      }
    } else {
      delayLabel = '—';
      delayClass = 'delay-ontime';
    }

    return {
      delayDays,
      delayPct,
      isDelayed,
      delayLabel,
      delayClass,
      consumoPct,
      consumoLabel
    };
  }

  // ==========================================================================
  // TEXT & BULLET POINT FORMATTING UTILITY
  // ==========================================================================
  function formatMultiLineText(rawText, emptyFallbackMsg) {
    if (!rawText || !String(rawText).trim()) {
      return `<div style="font-size: 12.5px; color: var(--text-muted); padding: 6px 0;">${emptyFallbackMsg || 'Nenhum registro informado.'}</div>`;
    }

    const lines = String(rawText)
      .split(/\r?\n/)
      .map(l => l.trim())
      .filter(Boolean);

    if (lines.length === 0) {
      return `<div style="font-size: 12.5px; color: var(--text-muted); padding: 6px 0;">${emptyFallbackMsg || 'Nenhum registro informado.'}</div>`;
    }

    const isBulletList = lines.some(l => /^[-•*]|\d+[\.\-\)]\s/.test(l));

    if (isBulletList) {
      const itemsHtml = lines.map(l => {
        const cleanLine = l.replace(/^[-•*]\s*|\d+[\.\-\)]\s*/, '');
        return `<li style="margin-bottom: 4px;">${cleanLine}</li>`;
      }).join('');
      return `<ul style="padding-left: 20px; font-size: 13px; line-height: 1.6; margin: 0;">${itemsHtml}</ul>`;
    }

    return lines.map(l => `<p style="margin-bottom: 6px; font-size: 13px; line-height: 1.6;">${l}</p>`).join('');
  }

  // ==========================================================================
  // FILTERING & SORTING LOGIC
  // ==========================================================================
  function getFilteredProjects() {
    const list = state.data.projetos || [];

    return list.filter(p => {
      // Área
      if (state.filters.area !== 'ALL') {
        const area = (getField(p, 'Área', 'Area') || '').trim();
        if (area !== state.filters.area) return false;
      }

      // Quick KPI Filter
      if (state.filters.quickFilter !== 'ALL') {
        const situacao = (getField(p, 'Situação', 'Situacao') || '').trim().toLowerCase();
        const status = (getField(p, 'Status') || '').trim().toLowerCase();

        if (state.filters.quickFilter === 'andamento' && !status.includes('andamento') && (status !== '' || situacao === '')) {
          if (!status.includes('andamento')) return false;
        }
        if (state.filters.quickFilter === 'saudavel' && !situacao.includes('saud')) return false;
        if (state.filters.quickFilter === 'atencao' && !situacao.includes('aten')) return false;
        if (state.filters.quickFilter === 'problema' && !situacao.includes('problem')) return false;
        if (state.filters.quickFilter === 'concluido' && !status.includes('conclu')) return false;
        if (state.filters.quickFilter === 'suspenso' && !status.includes('suspenso')) return false;
        if (state.filters.quickFilter === 'cancelado' && !status.includes('cancelad')) return false;
      }

      // Busca por nome ou ID do projeto
      if (state.filters.search) {
        const q = state.filters.search.toLowerCase();
        const name = (getField(p, 'Projeto', 'Nome') || '').toLowerCase();
        const id = (getField(p, 'ID Projeto', 'ID') || '').toLowerCase();
        const gp = (getField(p, 'GP', 'Gestor') || '').toLowerCase();
        if (!name.includes(q) && !id.includes(q) && !gp.includes(q)) {
          return false;
        }
      }

      return true;
    });
  }

  function getSortedProjects(projects) {
    const list = [...projects];

    return list.sort((a, b) => {
      // Ordenar por Criticidade: Problema -> Atenção -> Saudável -> Outros
      const sitA = (getField(a, 'Situação', 'Situacao') || '').toLowerCase();
      const sitB = (getField(b, 'Situação', 'Situacao') || '').toLowerCase();

      const getWeight = (s) => {
        if (s.includes('problem')) return 4;
        if (s.includes('aten')) return 3;
        if (s.includes('saud')) return 2;
        return 1;
      };

      const wA = getWeight(sitA);
      const wB = getWeight(sitB);
      if (wA !== wB) return wB - wA;

      // Empate: maior atraso / desvio
      const schedA = calculateProjectSchedule(a);
      const schedB = calculateProjectSchedule(b);
      if (schedB.delayDays !== schedA.delayDays) {
        return schedB.delayDays - schedA.delayDays;
      }

      // Empate secundário: nome do projeto
      const nameA = (getField(a, 'Projeto', 'Nome') || '').trim();
      const nameB = (getField(b, 'Projeto', 'Nome') || '').trim();
      return nameA.localeCompare(nameB);
    });
  }

  // ==========================================================================
  // POPULATE SELECTS
  // ==========================================================================
  function populateFilterOptions() {
    const projetos = state.data.projetos || [];

    // Áreas
    const areas = Array.from(new Set(projetos.map(p => (getField(p, 'Área', 'Area') || '').trim()).filter(Boolean))).sort();
    const areaSelect = document.getElementById('filterArea');
    if (areaSelect) {
      const current = areaSelect.value;
      areaSelect.innerHTML = '<option value="ALL">Todas as Áreas</option>' + 
        areas.map(a => `<option value="${a}">${a}</option>`).join('');
      areaSelect.value = areas.includes(current) ? current : 'ALL';
    }
  }

  // ==========================================================================
  // RENDER KPI SUMMARY CARDS
  // ==========================================================================
  function renderKPIs(allProjects) {
    const total = allProjects.length;
    let ongoingCount = 0;
    let healthyCount = 0;
    let attentionCount = 0;
    let problemCount = 0;
    let completedCount = 0;
    let suspendedCount = 0;
    let cancelledCount = 0;

    let latestRapDate = null;

    allProjects.forEach(p => {
      const status = (getField(p, 'Status') || '').trim().toLowerCase();
      const sit = (getField(p, 'Situação', 'Situacao') || '').trim().toLowerCase();

      if (status.includes('conclu')) completedCount++;
      else if (status.includes('suspenso')) suspendedCount++;
      else if (status.includes('cancelad')) cancelledCount++;
      else ongoingCount++; // Default to ongoing

      if (sit.includes('saud')) healthyCount++;
      else if (sit.includes('aten')) attentionCount++;
      else if (sit.includes('problem')) problemCount++;

      const rap = parseDate(getField(p, 'Data Última RAP', 'Data Ultima RAP', 'RAP'));
      if (rap && (!latestRapDate || rap > latestRapDate)) {
        latestRapDate = rap;
      }
    });

    const setCard = (valId, barId, count) => {
      const valEl = document.getElementById(valId);
      if (valEl) valEl.textContent = count;
      const barEl = document.getElementById(barId);
      if (barEl) {
        const pct = total > 0 ? Math.round((count / total) * 100) : 0;
        barEl.style.width = `${pct}%`;
      }
    };

    setCard('kpiTotal', 'kpiTotalBar', total);
    setCard('kpiAndamento', 'kpiAndamentoBar', ongoingCount);
    setCard('kpiSaudavel', 'kpiSaudavelBar', healthyCount);
    setCard('kpiAtencao', 'kpiAtencaoBar', attentionCount);
    setCard('kpiProblema', 'kpiProblemaBar', problemCount);
    setCard('kpiConcluido', 'kpiConcluidoBar', completedCount);
    setCard('kpiSuspenso', 'kpiSuspensoBar', suspendedCount);
    setCard('kpiCancelado', 'kpiCanceladoBar', cancelledCount);

    // Header updates
    const countEl = document.getElementById('headerProjectCount');
    if (countEl) countEl.innerHTML = `<strong>${total}</strong> projetos`;

    const refDateEl = document.getElementById('headerReferenceDate');
    if (refDateEl) {
      refDateEl.textContent = state.data.metadata.referenceDate || formatDateBr(new Date());
    }
  }

  // ==========================================================================
  // RENDER ACTIVE FILTER CHIPS
  // ==========================================================================
  function renderActiveFilterChips() {
    const bar = document.getElementById('activeFiltersBar');
    const container = document.getElementById('activeFiltersChips');
    if (!bar || !container) return;

    const chips = [];

    if (state.filters.quickFilter !== 'ALL') {
      const map = {
        'andamento': 'Filtro: Em Andamento',
        'saudavel': 'Filtro: Saudáveis',
        'atencao': 'Filtro: Em Atenção',
        'problema': 'Filtro: Com Problema',
        'concluido': 'Filtro: Concluídos',
        'suspenso': 'Filtro: Suspensos',
        'cancelado': 'Filtro: Cancelados'
      };
      chips.push({ key: 'quickFilter', label: map[state.filters.quickFilter] || state.filters.quickFilter });
    }

    if (state.filters.area !== 'ALL') {
      chips.push({ key: 'area', label: `Área: ${state.filters.area}` });
    }

    if (state.filters.search) {
      chips.push({ key: 'search', label: `Busca: "${state.filters.search}"` });
    }

    if (chips.length === 0) {
      bar.style.display = 'none';
      container.innerHTML = '';
      return;
    }

    bar.style.display = 'flex';
    container.innerHTML = chips.map(c => `
      <div class="filter-chip">
        <span>${c.label}</span>
        <button data-chip-key="${c.key}" title="Remover filtro">✕</button>
      </div>
    `).join('');

    container.querySelectorAll('button[data-chip-key]').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const key = btn.getAttribute('data-chip-key');
        if (key === 'search') {
          state.filters.search = '';
          const searchInput = document.getElementById('searchInput');
          if (searchInput) searchInput.value = '';
          const clearBtn = document.getElementById('btnSearchClear');
          if (clearBtn) clearBtn.style.display = 'none';
        } else if (key === 'quickFilter') {
          state.filters.quickFilter = 'ALL';
          document.querySelectorAll('.kpi-card').forEach(k => k.classList.remove('active-kpi'));
        } else if (key === 'area') {
          state.filters.area = 'ALL';
          const el = document.getElementById('filterArea');
          if (el) el.value = 'ALL';
        }
        updateOverview();
      });
    });
  }

  // ==========================================================================
  // RENDER PROJECT VISUAL CARD (SINGLE)
  // ==========================================================================
  function createProjectCardHtml(p) {
    const id = getField(p, 'ID Projeto', 'ID', 'Codigo') || '';
    const name = getField(p, 'Projeto', 'Nome', 'Nome do Projeto') || 'Projeto sem nome';
    const area = getField(p, 'Área', 'Area') || '—';
    const rawSituacao = getField(p, 'Situação', 'Situacao');
    const situacao = rawSituacao || 'Não informada';
    const status = getField(p, 'Status') || '';
    const pctExec = parsePercentage(getField(p, '% Execução', '% Execucao', 'Progresso', 'Execucao'));

    // Classes de criticidade
    let sitClass = 'sit-indefinido';
    let sitBadgeClass = 'status-indefinido';
    let badgeText = situacao;
    const sitLower = situacao.toLowerCase();
    const stLower = status.toLowerCase();

    if (stLower.includes('conclu') || sitLower.includes('conclu')) {
      sitClass = 'sit-concluido';
      sitBadgeClass = 'status-concluido';
      badgeText = 'Concluído';
    } else if (sitLower.includes('problem')) {
      sitClass = 'sit-problema';
      sitBadgeClass = 'status-problema';
    } else if (sitLower.includes('aten')) {
      sitClass = 'sit-atencao';
      sitBadgeClass = 'status-atencao';
    } else if (sitLower.includes('saud')) {
      sitClass = 'sit-saudavel';
      sitBadgeClass = 'status-saudavel';
    }

    const isOngoing = stLower.includes('andamento') || (!stLower.includes('conclu') && !stLower.includes('suspenso') && !stLower.includes('cancelad'));

    // Consumo do prazo para projetos em andamento
    let consumoSectionHtml = '';
    if (isOngoing) {
      const rawConsumo = getField(p, '% Consumo Prazo', '% Consumo do Prazo', 'Consumo do Prazo', 'Consumo Prazo', '%Consumo Prazo', '%Consumo \nPrazo', '%Consumo \r\nPrazo', 'Consumo');
      let pctConsumo = null;
      if (rawConsumo !== '' && rawConsumo !== null && rawConsumo !== undefined) {
        pctConsumo = parsePercentage(rawConsumo);
      } else {
        const sched = calculateProjectSchedule(p);
        pctConsumo = sched.consumoPct;
      }

      if (pctConsumo !== null && pctConsumo !== undefined) {
        consumoSectionHtml = `
          <!-- Progress Bar (Consumo do Prazo) -->
          <div class="card-progress-section">
            <div class="card-progress-header">
              <span class="card-progress-label">Consumo do Prazo</span>
              <span class="card-progress-pct">${pctConsumo}%</span>
            </div>
            <div class="card-progress-bar">
              <div class="card-progress-fill" style="width: ${Math.min(pctConsumo, 100)}%;"></div>
            </div>
          </div>
        `;
      }
    }

    return `
      <article class="project-card ${sitClass}" data-project-id="${id}" tabindex="0" role="button" aria-label="Ver detalhes de ${name}">
        <div>
          <!-- Top Meta -->
          <div class="card-meta-top">
            <span class="card-area-badge">${area}</span>
          </div>

          <!-- Title -->
          <div class="card-main-info">
            <h3 class="card-project-title" title="${name}">${name}</h3>
          </div>

          <!-- Badges Row (Situação) -->
          <div class="card-badges-row">
            <span class="badge-pill ${sitBadgeClass}">
              ${badgeText}
            </span>
          </div>

          <!-- Progress Bar (Execução Física) -->
          <div class="card-progress-section">
            <div class="card-progress-header">
              <span class="card-progress-label">Execução Física</span>
              <span class="card-progress-pct">${pctExec}%</span>
            </div>
            <div class="card-progress-bar">
              <div class="card-progress-fill" style="width: ${pctExec}%;"></div>
            </div>
          </div>

          ${consumoSectionHtml}
        </div>
      </article>
    `;
  }

  // ==========================================================================
  // RENDER OVERVIEW (GRID)
  // ==========================================================================
  function renderProjectsShowcase(filteredProjects) {
    const container = document.getElementById('projectsContainer');
    const badge = document.getElementById('resultsCountBadge');
    if (!container) return;

    if (badge) {
      badge.innerHTML = `Exibindo <strong>${filteredProjects.length}</strong> projetos`;
    }

    if (filteredProjects.length === 0) {
      container.className = 'projects-grid';
      container.innerHTML = `
        <div class="empty-state-box">
          <div class="empty-state-title">Nenhum projeto encontrado</div>
          <p>Tente ajustar os filtros ou termo de busca para visualizar os projetos correspondentes.</p>
        </div>
      `;
      return;
    }

    const sorted = getSortedProjects(filteredProjects);
    container.className = 'projects-grid';
    container.innerHTML = sorted.map(p => createProjectCardHtml(p)).join('');

    // Attach Drill-Through Event Listeners to all project cards
    container.querySelectorAll('.project-card[data-project-id]').forEach(card => {
      const pid = card.getAttribute('data-project-id');
      card.addEventListener('click', () => {
        navigateToDetail(pid);
      });
      card.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault();
          navigateToDetail(pid);
        }
      });
    });
  }

  // ==========================================================================
  // DRILL-THROUGH: DETALHAMENTO DO PROJETO (VISÃO 360°)
  // ==========================================================================
  function renderProjectDetail(projectId) {
    const project = (state.data.projetos || []).find(p => getField(p, 'ID Projeto', 'ID', 'Codigo') === projectId);
    if (!project) return;

    // Header Hero Elements
    const titleEl = document.getElementById('detailProjectTitle');
    const gpEl = document.getElementById('detailProjectGp');
    const areaEl = document.getElementById('detailProjectArea');
    const rapEl = document.getElementById('detailProjectRap');
    const portariaEl = document.getElementById('detailProjectPortaria');

    if (titleEl) titleEl.textContent = getField(project, 'Projeto', 'Nome', 'Nome do Projeto') || 'Projeto sem nome';
    if (areaEl) areaEl.textContent = getField(project, 'Área', 'Area') || '—';
    if (gpEl) gpEl.textContent = getField(project, 'GP', 'Gestor') || '—';
    if (rapEl) rapEl.textContent = formatDateBr(getField(project, 'Data Última RAP', 'Data Ultima RAP', 'RAP'));
    if (portariaEl) portariaEl.textContent = getField(project, 'Portaria', 'Nº Portaria', 'No Portaria', 'Numero Portaria') || '—';

    // Status Badges in Hero (Apenas Status)
    const rawSituacao = getField(project, 'Situação', 'Situacao');
    const situacao = rawSituacao || 'Não informada';
    const status = getField(project, 'Status') || 'Em andamento';

    const badgesContainer = document.getElementById('detailHeroBadges');
    if (badgesContainer) {
      let statusBadgeClass = 'status-andamento';
      const stLower = status.toLowerCase();
      if (stLower.includes('suspenso')) statusBadgeClass = 'status-suspenso';
      else if (stLower.includes('cancelad')) statusBadgeClass = 'status-cancelado';
      else if (stLower.includes('conclu')) statusBadgeClass = 'status-concluido';

      badgesContainer.innerHTML = `
        <span class="detail-badge ${statusBadgeClass}">${status}</span>
      `;
    }

    // Card de Situação & Justificativa da Situação (abaixo do cabeçalho)
    let sitBadgeClass = 'status-indefinido';
    const sitLower = situacao.toLowerCase();
    if (sitLower.includes('conclu')) sitBadgeClass = 'status-concluido';
    else if (sitLower.includes('problem')) sitBadgeClass = 'status-problema';
    else if (sitLower.includes('aten')) sitBadgeClass = 'status-atencao';
    else if (sitLower.includes('saud')) sitBadgeClass = 'status-saudavel';

    const sitBadgeContainer = document.getElementById('detailSituacaoBadgeContainer');
    if (sitBadgeContainer) {
      sitBadgeContainer.innerHTML = `<span class="detail-badge ${sitBadgeClass}">${situacao}</span>`;
    }

    const justificativaContainer = document.getElementById('detailSituacaoJustificativa');
    const justificativaText = getField(project, 'Justificativa da Situação', 'Justificativa da Situacao', 'Justificativa Situação', 'Justificativa Situacao', 'Justificativa', 'Motivo da Situação', 'Motivo Situacao');
    if (justificativaContainer) {
      justificativaContainer.innerHTML = `<div class="rap-report-text">${formatMultiLineText(justificativaText, 'Nenhuma justificativa informada para a situação atual.')}</div>`;
    }

    // Card: O que foi feito na última semana
    const ultimaSemanaPeriodo = document.getElementById('detailUltimaSemanaPeriodo');
    const ultimaSemanaConteudo = document.getElementById('detailUltimaSemanaConteudo');
    const feitoText = getField(project, 'O que foi feito na última semana', 'O que foi feito na ultima semana', 'O que foi feito', 'Feito na última semana', 'Feito na ultima semana', 'Realizações da semana', 'Realizacoes da semana');
    const periodoText = getField(project, 'Período da última semana', 'Periodo da última semana', 'Período da ultima semana', 'Periodo da ultima semana', 'Período', 'Periodo');

    if (ultimaSemanaPeriodo) {
      if (periodoText) {
        ultimaSemanaPeriodo.textContent = `Período: ${periodoText}`;
        ultimaSemanaPeriodo.style.display = 'inline-block';
      } else {
        ultimaSemanaPeriodo.textContent = '';
        ultimaSemanaPeriodo.style.display = 'none';
      }
    }

    if (ultimaSemanaConteudo) {
      ultimaSemanaConteudo.innerHTML = `<div class="rap-report-text">${formatMultiLineText(feitoText, 'Nenhuma realização relevante registrada no período.')}</div>`;
    }

    // Controle de exibição dos cards conforme status:
    // Para projetos com status "Concluído", a ficha exibe apenas o cabeçalho e os cards "Escopo" e "Objetivos Estratégicos".
    // Para projetos com status "Em andamento" (e outros), mantém a ficha completa.
    // A seção "O que foi feito na última semana" é exibida exclusivamente para projetos com status "Em andamento".
    const isConcluido = status.toLowerCase().includes('conclu');
    const isOngoing = status.toLowerCase().includes('andamento');

    const ongoingOnlyCards = [
      'detailSituacaoCard',
      'detailUltimaSemanaCard',
      'detailCardExecucao',
      'detailCardSituacaoAtual',
      'detailCardRiscos',
      'detailCardAtencao',
      'detailCardProximosPassos',
      'detailCardPendencias'
    ];

    ongoingOnlyCards.forEach(cardId => {
      const el = document.getElementById(cardId);
      if (el) {
        if (cardId === 'detailUltimaSemanaCard') {
          el.style.display = isOngoing ? '' : 'none';
        } else {
          el.style.display = isConcluido ? 'none' : '';
        }
      }
    });

    // BLOCO 1: Execução & Prazos
    const pctExec = parsePercentage(getField(project, '% Execução', '% Execucao', 'Progresso', 'Execucao'));
    const pctEl = document.getElementById('detailPctExec');
    const pctBar = document.getElementById('detailPctExecBar');
    if (pctEl) pctEl.textContent = `${pctExec}%`;
    if (pctBar) pctBar.style.width = `${pctExec}%`;

    const sched = calculateProjectSchedule(project);
    const dataInicioEl = document.getElementById('detailDataInicio');
    const dataPlanEl = document.getElementById('detailDataPlan');
    const dataAtualEl = document.getElementById('detailDataAtual');
    const consumoPrazoEl = document.getElementById('detailConsumoPrazo');
    const desvioValueEl = document.getElementById('detailDesvioValue');
    const desvioBox = document.getElementById('detailDesvioBox');
    if (dataInicioEl) dataInicioEl.textContent = formatDateBr(getField(project, 'Data Início', 'Data Inicio', 'Inicio'));
    if (dataPlanEl) dataPlanEl.textContent = formatDateBr(getField(project, 'Data Fim Planejada', 'Data Fim Planejado', 'Fim Planejado'));
    if (dataAtualEl) dataAtualEl.textContent = formatDateBr(getField(project, 'Data Fim Atual', 'Fim Atual', 'Previsão'));
    if (consumoPrazoEl) consumoPrazoEl.textContent = sched.consumoLabel;
    if (desvioValueEl) desvioValueEl.textContent = sched.delayLabel;

    if (desvioBox) {
      if (sched.delayDays > 0) {
        desvioBox.style.background = '#fef2f2';
        desvioBox.style.borderColor = '#fecaca';
        if (desvioValueEl) desvioValueEl.style.color = '#b91c1c';
      } else if (sched.delayDays < 0) {
        desvioBox.style.background = '#f0fdf4';
        desvioBox.style.borderColor = '#bbf7d0';
        if (desvioValueEl) desvioValueEl.style.color = '#15803d';
      } else {
        desvioBox.style.background = '#f8fafc';
        desvioBox.style.borderColor = '#e2e8f0';
        if (desvioValueEl) desvioValueEl.style.color = '#1e293b';
      }
    }

    // BLOCO 2: Situação Atual (Últimas Atualizações)
    const rapReportContainer = document.getElementById('detailRapReport');
    const singleSheetUpdate = getField(project, 'Últimas Atualizações', 'Ultimas Atualizacoes', 'Situação Atual', 'Situacao Atual', 'Atualizações', 'Atualizacoes');
    const legacyUpdate = (state.data.atualizacoes || []).find(u => getField(u, 'ID Projeto', 'ID') === projectId);
    const updateText = singleSheetUpdate || (legacyUpdate ? getField(legacyUpdate, 'Situação Atual', 'Situacao Atual', 'Atualização', 'Atualizacao') : '');

    if (rapReportContainer) {
      rapReportContainer.innerHTML = `<div class="rap-report-text">${formatMultiLineText(updateText, 'Nenhum relato adicional registrado na última RAP deste projeto.')}</div>`;
    }

    // BLOCO 3: Escopo
    const escopoContainer = document.getElementById('detailEscopoContainer');
    const escopoText = getField(project, 'Escopo', 'Escopo do Projeto');
    if (escopoContainer) {
      if (escopoText) {
        escopoContainer.innerHTML = `
          <div class="rap-report-box" style="background: #f8fafc; border-left: 4px solid var(--primary-accent); padding: 12px 14px;">
            <div class="rap-report-text" style="font-size: 13px; line-height: 1.6; color: var(--text-primary);">${formatMultiLineText(escopoText, '')}</div>
          </div>
        `;
      } else {
        escopoContainer.innerHTML = '<div style="font-size: 12.5px; color: var(--text-muted); padding: 8px 0;">Nenhum escopo informado para este projeto.</div>';
      }
    }

    // BLOCO 4: Objetivos Estratégicos
    const objetivosContainer = document.getElementById('detailObjetivosContainer');
    const objetivosText = getField(project, 'Objetivos Estratégicos', 'Objetivos Estrategicos', 'Objetivo Estratégico', 'Objetivo Estrategico', 'Alinhamento Estratégico', 'Alinhamento Estrategico');
    if (objetivosContainer) {
      if (objetivosText) {
        const lines = String(objetivosText).split(/\r?\n/).map(l => l.trim()).filter(Boolean);
        if (lines.length > 1) {
          objetivosContainer.innerHTML = lines.map(line => `
            <div class="detail-item-card" style="border-left: 3px solid #0284c7; margin-bottom: 6px; padding: 8px 12px;">
              <div class="detail-item-desc" style="font-size: 12px; color: var(--text-primary); line-height: 1.45;">
                <strong style="color: var(--primary-accent); margin-right: 4px;">•</strong> ${line.replace(/^[-•*]\s*/, '')}
              </div>
            </div>
          `).join('');
        } else {
          objetivosContainer.innerHTML = `
            <div class="detail-item-card" style="border-left: 3px solid #0284c7;">
              <div class="detail-item-desc" style="font-size: 12.5px; line-height: 1.5;">${formatMultiLineText(objetivosText, '')}</div>
            </div>
          `;
        }
      } else {
        objetivosContainer.innerHTML = '<div style="font-size: 12.5px; color: var(--text-muted); padding: 8px 0;">Nenhum objetivo estratégico vinculado a este projeto.</div>';
      }
    }

    // BLOCO 5: Gestão de Riscos
    const risksContainer = document.getElementById('detailRisksContainer');
    const singleSheetRisk = getField(project, 'Riscos', 'Risco', 'Gestão de Riscos', 'Gestao de Riscos');
    const legacyRisks = (state.data.riscos || []).filter(r => getField(r, 'ID Projeto', 'ID') === projectId && (getField(r, 'Risco / Impacto', 'Risco') || '').trim() !== '');

    if (risksContainer) {
      if (singleSheetRisk) {
        risksContainer.innerHTML = `
          <div class="detail-item-card" style="border-left: 3px solid var(--status-problem);">
            <div class="detail-item-header">
              <span class="detail-item-title" style="color: var(--status-problem-text);">Risco Identificado</span>
              <span class="badge-pill status-problema" style="font-size: 10px;">Monitoramento APE</span>
            </div>
            <div class="detail-item-desc">${formatMultiLineText(singleSheetRisk, '')}</div>
          </div>
        `;
      } else if (legacyRisks.length > 0) {
        risksContainer.innerHTML = legacyRisks.map(r => `
          <div class="detail-item-card" style="border-left: 3px solid var(--status-problem);">
            <div class="detail-item-header">
              <span class="detail-item-title">Risco / Impacto</span>
              <span class="badge-pill status-problema" style="font-size: 10px;">
                Prob: ${getField(r, 'Probabilidade') || '—'} | Imp: ${getField(r, 'Impacto') || '—'}
              </span>
            </div>
            <div class="detail-item-desc">${getField(r, 'Risco / Impacto', 'Risco')}</div>
            ${getField(r, 'Resposta / Mitigação') ? `
              <div class="detail-item-meta">
                <strong>Mitigação:</strong> ${getField(r, 'Resposta / Mitigação')}
              </div>
            ` : ''}
          </div>
        `).join('');
      } else {
        risksContainer.innerHTML = '<div style="font-size: 12.5px; color: var(--text-muted); padding: 8px 0;">Nenhum risco crítico registrado para este projeto.</div>';
      }
    }

    // BLOCO 4: Pontos de Atenção & Impedimentos
    const attentionContainer = document.getElementById('detailAttentionContainer');
    const singleSheetAttention = getField(project, 'Pontos de Atenção', 'Pontos de Atencao', 'Ponto de Atenção', 'Ponto de Atencao');
    const singleSheetImpediment = getField(project, 'Impedimentos', 'Impedimento');
    const legacyPontos = (state.data.pontosAtencao || []).filter(pt => getField(pt, 'ID Projeto', 'ID') === projectId && (getField(pt, 'Ponto de Atenção', 'Ponto') || '').trim() !== '');
    const legacyImpedimentos = (state.data.impedimentos || []).filter(i => getField(i, 'ID Projeto', 'ID') === projectId && (getField(i, 'Impedimento') || '').trim() !== '');

    if (attentionContainer) {
      let html = '';
      let hasContent = false;

      if (singleSheetAttention) {
        hasContent = true;
        html += `
          <div class="detail-item-card" style="border-left: 3px solid var(--status-attention);">
            <div class="detail-item-header">
              <span class="detail-item-title" style="color: var(--status-attention-text);">Ponto de Atenção</span>
            </div>
            <div class="detail-item-desc">${formatMultiLineText(singleSheetAttention, '')}</div>
          </div>
        `;
      } else if (legacyPontos.length > 0) {
        hasContent = true;
        legacyPontos.forEach(pt => {
          html += `
            <div class="detail-item-card" style="border-left: 3px solid var(--status-attention);">
              <div class="detail-item-header">
                <span class="detail-item-title" style="color: var(--status-attention-text);">Ponto de Atenção</span>
              </div>
              <div class="detail-item-desc">${getField(pt, 'Ponto de Atenção')}</div>
              ${getField(pt, 'Impacto / Observação') ? `<div class="detail-item-meta">${getField(pt, 'Impacto / Observação')}</div>` : ''}
            </div>
          `;
        });
      }

      if (singleSheetImpediment) {
        hasContent = true;
        html += `
          <div class="detail-item-card" style="border-left: 3px solid var(--status-problem);">
            <div class="detail-item-header">
              <span class="detail-item-title" style="color: var(--status-problem-text);">Impedimento</span>
            </div>
            <div class="detail-item-desc">${formatMultiLineText(singleSheetImpediment, '')}</div>
          </div>
        `;
      } else if (legacyImpedimentos.length > 0) {
        hasContent = true;
        legacyImpedimentos.forEach(i => {
          html += `
            <div class="detail-item-card" style="border-left: 3px solid var(--status-problem);">
              <div class="detail-item-header">
                <span class="detail-item-title" style="color: var(--status-problem-text);">Impedimento</span>
                <span class="badge-pill status-problema" style="font-size: 10px;">Impacto: ${getField(i, 'Impacto') || '—'}</span>
              </div>
              <div class="detail-item-desc">${getField(i, 'Impedimento')}</div>
              ${getField(i, 'Situação / Ação') ? `<div class="detail-item-meta"><strong>Ação:</strong> ${getField(i, 'Situação / Ação')}</div>` : ''}
            </div>
          `;
        });
      }

      if (!hasContent) {
        html = '<div style="font-size: 12.5px; color: var(--text-muted); padding: 8px 0;">Nenhum ponto de atenção ou impedimento crítico registrado.</div>';
      }

      attentionContainer.innerHTML = html;
    }

    // BLOCO 5: Próximos Passos
    const nextStepsContainer = document.getElementById('detailNextStepsContainer');
    const singleSheetNextSteps = getField(project, 'Próximos Passos', 'Proximos Passos', 'Próximo Passo', 'Proximo Passo');
    const legacyPassos = (state.data.proximosPassos || []).filter(p => getField(p, 'ID Projeto', 'ID') === projectId && (getField(p, 'Próximo Passo') || '').trim() !== '');
    const legacyEntregas = (state.data.entregas || []).filter(e => getField(e, 'ID Projeto', 'ID') === projectId && (getField(e, 'Entrega') || '').trim() !== '');

    if (nextStepsContainer) {
      if (singleSheetNextSteps) {
        nextStepsContainer.innerHTML = `
          <div class="detail-item-card" style="border-left: 3px solid var(--primary-accent);">
            <div class="detail-item-header">
              <span class="detail-item-title" style="color: var(--primary-dark);">Ações & Próximos Passos</span>
              <span class="badge-pill status-atencao" style="font-size: 10px;">Em andamento</span>
            </div>
            <div class="detail-item-desc">${formatMultiLineText(singleSheetNextSteps, '')}</div>
          </div>
        `;
      } else if (legacyPassos.length > 0 || legacyEntregas.length > 0) {
        let html = '';
        legacyEntregas.forEach(e => {
          html += `
            <div class="detail-item-card">
              <div class="detail-item-header">
                <span class="detail-item-title">Entrega</span>
                <span class="badge-pill status-andamento" style="font-size: 10px;">${getField(e, 'Status') || 'Pendente'}</span>
              </div>
              <div class="detail-item-desc">${getField(e, 'Entrega')}</div>
              ${getField(e, 'Observação') ? `<div class="detail-item-meta">${getField(e, 'Observação')}</div>` : ''}
            </div>
          `;
        });
        legacyPassos.forEach(p => {
          html += `
            <div class="detail-item-card">
              <div class="detail-item-header">
                <span class="detail-item-title">Próximo Passo</span>
                <span class="badge-pill status-atencao" style="font-size: 10px;">${getField(p, 'Status') || 'Em andamento'}</span>
              </div>
              <div class="detail-item-desc">${getField(p, 'Próximo Passo')}</div>
              <div class="detail-item-meta">
                <span>Responsável: <strong>${getField(p, 'Responsável') || '—'}</strong></span> • 
                <span>Prazo: <strong>${formatDateBr(getField(p, 'Prazo'))}</strong></span>
              </div>
            </div>
          `;
        });
        nextStepsContainer.innerHTML = html;
      } else {
        nextStepsContainer.innerHTML = '<div style="font-size: 12.5px; color: var(--text-muted); padding: 8px 0;">Nenhum próximo passo específico registrado no momento.</div>';
      }
    }

    // BLOCO 6: Pendências / Decisões
    const decisionsContainer = document.getElementById('detailDecisionsContainer');
    const singleSheetDecisions = getField(project, 'Decisões', 'Decisoes', 'Decisão / Escalonamento', 'Decisao / Escalonamento', 'Pendências', 'Pendencias');
    const legacyDecisoes = (state.data.decisoes || []).filter(d => getField(d, 'ID Projeto', 'ID') === projectId && (getField(d, 'Decisão / Escalonamento', 'Decisão') || '').trim() !== '');

    if (decisionsContainer) {
      if (singleSheetDecisions) {
        decisionsContainer.innerHTML = `
          <div class="detail-item-card" style="border-left: 3px solid var(--primary-accent);">
            <div class="detail-item-header">
              <span class="detail-item-title">Pendência / Decisão</span>
              <span class="badge-pill status-atencao" style="font-size: 10px;">Pendente</span>
            </div>
            <div class="detail-item-desc">${formatMultiLineText(singleSheetDecisions, '')}</div>
          </div>
        `;
      } else if (legacyDecisoes.length > 0) {
        decisionsContainer.innerHTML = legacyDecisoes.map(d => `
          <div class="detail-item-card" style="border-left: 3px solid var(--primary-accent);">
            <div class="detail-item-header">
              <span class="detail-item-title">Decisão / Escalonamento</span>
              <span class="badge-pill status-atencao" style="font-size: 10px;">${getField(d, 'Status') || 'Pendente'}</span>
            </div>
            <div class="detail-item-desc">${getField(d, 'Decisão / Escalonamento')}</div>
            <div class="detail-item-meta">
              <span>Responsável: <strong>${getField(d, 'Responsável') || '—'}</strong></span> • 
              <span>Data: <strong>${formatDateBr(getField(d, 'Data'))}</strong></span>
            </div>
          </div>
        `).join('');
      } else {
        decisionsContainer.innerHTML = '<div style="font-size: 12.5px; color: var(--text-muted); padding: 8px 0;">Nenhuma decisão ou pendência crítica registrada para este projeto.</div>';
      }
    }

    // Update Stepper Index
    const filteredList = getSortedProjects(getFilteredProjects());
    const currentIndex = filteredList.findIndex(p => getField(p, 'ID Projeto', 'ID', 'Codigo') === projectId);
    const indicator = document.getElementById('detailProjectIndex');
    if (indicator) {
      indicator.textContent = `${currentIndex >= 0 ? currentIndex + 1 : 1} de ${filteredList.length}`;
    }
  }

  // ==========================================================================
  // NAVIGATION (DRILL-THROUGH ROUTER)
  // ==========================================================================
  function navigateToDetail(projectId) {
    state.currentView = 'detail';
    state.selectedProjectId = projectId;

    const overviewEl = document.getElementById('viewOverview');
    const detailEl = document.getElementById('viewDetail');

    if (overviewEl) overviewEl.classList.remove('active');
    if (detailEl) detailEl.classList.add('active');

    window.scrollTo({ top: 0, behavior: 'smooth' });
    renderProjectDetail(projectId);

    // Update history / hash
    window.location.hash = `detail/${projectId}`;
  }

  function navigateToOverview() {
    state.currentView = 'overview';
    state.selectedProjectId = null;

    const overviewEl = document.getElementById('viewOverview');
    const detailEl = document.getElementById('viewDetail');

    if (detailEl) detailEl.classList.remove('active');
    if (overviewEl) overviewEl.classList.add('active');

    window.scrollTo({ top: 0, behavior: 'smooth' });
    window.location.hash = 'overview';
  }

  function stepProject(direction) {
    const filteredList = getSortedProjects(getFilteredProjects());
    if (filteredList.length === 0) return;

    let currentIndex = filteredList.findIndex(p => getField(p, 'ID Projeto', 'ID', 'Codigo') === state.selectedProjectId);
    if (currentIndex === -1) currentIndex = 0;

    let nextIndex = currentIndex + direction;
    if (nextIndex < 0) nextIndex = filteredList.length - 1;
    if (nextIndex >= filteredList.length) nextIndex = 0;

    const nextProject = filteredList[nextIndex];
    if (nextProject) {
      navigateToDetail(getField(nextProject, 'ID Projeto', 'ID', 'Codigo'));
    }
  }

  // ==========================================================================
  // MASTER UPDATE DASHBOARD
  // ==========================================================================
  function updateKPIActiveState() {
    document.querySelectorAll('.kpi-card[data-filter]').forEach(card => {
      const f = card.getAttribute('data-filter');
      if (state.filters.quickFilter !== 'ALL' && f === state.filters.quickFilter) {
        card.classList.add('active-kpi');
      } else {
        card.classList.remove('active-kpi');
      }
    });
  }

  function updateOverview() {
    renderKPIs(state.data.projetos || []);
    renderProjectsShowcase(getFilteredProjects());
    updateKPIActiveState();
    renderActiveFilterChips();

    // Exibir legenda exclusivamente na visualização de projetos em andamento
    const legendEl = document.getElementById('badgesLegendSection') || document.querySelector('.badges-legend-section');
    if (legendEl) {
      const isOngoingView = state.filters.quickFilter === 'andamento' || ['saudavel', 'atencao', 'problema'].includes(state.filters.quickFilter);
      legendEl.style.display = isOngoingView ? '' : 'none';
    }
  }

  // ==========================================================================
  // EVENT LISTENERS & INITIALIZATION
  // ==========================================================================
  function initEvents() {
    // KPI Cards Filter Click
    const kpiCards = document.querySelectorAll('.kpi-card[data-filter]');
    kpiCards.forEach(card => {
      card.addEventListener('click', () => {
        const filterVal = card.getAttribute('data-filter');
        if (state.filters.quickFilter === filterVal && filterVal !== 'ALL') {
          state.filters.quickFilter = 'ALL';
        } else {
          state.filters.quickFilter = filterVal;
        }
        updateOverview();
      });

      card.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault();
          card.click();
        }
      });
    });

    // Drill-Through Navigation Buttons
    const btnBack = document.getElementById('btnBackToOverview');
    const btnBackBottom = document.getElementById('btnBackToOverviewBottom');
    if (btnBack) btnBack.addEventListener('click', navigateToOverview);
    if (btnBackBottom) btnBackBottom.addEventListener('click', navigateToOverview);

    const btnPrev = document.getElementById('btnPrevProject');
    const btnNext = document.getElementById('btnNextProject');
    if (btnPrev) btnPrev.addEventListener('click', () => stepProject(-1));
    if (btnNext) btnNext.addEventListener('click', () => stepProject(1));

    // Keyboard navigation (Esc to return, Arrow keys to step)
    window.addEventListener('keydown', (e) => {
      if (state.currentView === 'detail') {
        if (e.key === 'Escape') {
          navigateToOverview();
        } else if (e.key === 'ArrowLeft') {
          stepProject(-1);
        } else if (e.key === 'ArrowRight') {
          stepProject(1);
        }
      }
    });

    // Print buttons
    const btnPrint = document.getElementById('btnPrintReport');
    if (btnPrint) {
      btnPrint.addEventListener('click', () => window.print());
    }

    const btnPrintProject = document.getElementById('btnPrintProjectDetail');
    if (btnPrintProject) {
      btnPrintProject.addEventListener('click', () => window.print());
    }

    // Handle initial URL hash
    window.addEventListener('hashchange', checkHashRoute);
  }

  function checkHashRoute() {
    const hash = window.location.hash;
    if (hash.startsWith('#detail/')) {
      const pid = hash.replace('#detail/', '').trim();
      if (pid) {
        navigateToDetail(pid);
        return;
      }
    }
    if (state.currentView === 'detail' && (!hash || hash === '#overview')) {
      navigateToOverview();
    }
  }

  // ==========================================================================
  // SHEETJS DYNAMIC PARSER & AUTO-LOADER (Suporte Aba Única & Multi-Aba)
  // ==========================================================================
  function extractSpreadsheetDate(workbook, file, httpLastModified) {
    // 1. Metadados do OpenXML do Workbook (Propriedade ModifiedDate do Excel)
    if (workbook && workbook.Props) {
      if (workbook.Props.ModifiedDate) {
        const d = parseDate(workbook.Props.ModifiedDate);
        if (d) return d;
      }
      if (workbook.Props.modified) {
        const d = parseDate(workbook.Props.modified);
        if (d) return d;
      }
    }

    // 2. Propriedade do arquivo no computador (File Input ou Drag & Drop)
    if (file && file.lastModified) {
      const d = new Date(file.lastModified);
      if (!isNaN(d.getTime())) return d;
    }

    // 3. Cabeçalho de rede Last-Modified (se carregado via fetch/servidor)
    if (httpLastModified) {
      const d = new Date(httpLastModified);
      if (!isNaN(d.getTime())) return d;
    }

    // 4. Fallback: metadata inicial do arquivo ou data atual
    if (state.data.metadata && state.data.metadata.referenceDate) {
      const d = parseDate(state.data.metadata.referenceDate);
      if (d) return d;
    }

    return new Date();
  }

  function processWorkbook(workbook, fileName, file, httpLastModified) {
    const parseSheet = (sheetName) => {
      const sheet = workbook.Sheets[sheetName];
      if (!sheet) return [];
      return window.XLSX.utils.sheet_to_json(sheet, { defval: '' });
    };

    // Identificar a aba principal de projetos (preferência por "Projetos" ou primeira aba não auxiliar)
    let mainSheetName = 'Projetos';
    if (!workbook.Sheets[mainSheetName]) {
      const nonListSheet = workbook.SheetNames.find(n => {
        const lower = n.toLowerCase();
        return lower !== 'listas' && lower !== 'instruções' && lower !== 'instrucoes';
      });
      mainSheetName = nonListSheet || workbook.SheetNames[0];
    }

    const rawProjects = parseSheet(mainSheetName);

    // Filtrar apenas linhas com ID ou Nome de projeto válidos
    state.data.projetos = rawProjects.filter(p => {
      const id = String(getField(p, 'ID Projeto', 'ID', 'Codigo') || '').trim();
      const name = String(getField(p, 'Projeto', 'Nome', 'Nome do Projeto') || '').trim();
      return id !== '' || name !== '';
    });

    // Suporte a abas legadas/secundárias caso existam no arquivo
    state.data.atualizacoes = parseSheet('Atualizações');
    state.data.entregas = parseSheet('Entregas');
    state.data.pontosAtencao = parseSheet('Pontos de Atenção');
    state.data.impedimentos = parseSheet('Impedimentos');
    state.data.proximosPassos = parseSheet('Próximos Passos');
    state.data.riscos = parseSheet('Riscos');
    state.data.decisoes = parseSheet('Decisões');
    state.data.listas = parseSheet('Listas');

    const modDate = extractSpreadsheetDate(workbook, file, httpLastModified);
    state.data.metadata.referenceDate = formatDateBr(modDate);
    state.data.metadata.sourceFile = fileName;
    state.data.metadata.loadedAt = new Date().toISOString();

    const sourceFileEl = document.getElementById('headerSourceFile');
    if (sourceFileEl) sourceFileEl.textContent = fileName;

    const refDateEl = document.getElementById('headerReferenceDate');
    if (refDateEl) refDateEl.textContent = state.data.metadata.referenceDate;

    populateFilterOptions();
    updateOverview();

    if (state.currentView === 'detail' && state.selectedProjectId) {
      renderProjectDetail(state.selectedProjectId);
    }
  }

  function handleExcelFile(file) {
    if (!window.XLSX) {
      alert('Biblioteca XLSX não carregada no navegador. Utilize a base já embutida.');
      return;
    }

    const reader = new FileReader();
    reader.onload = function (e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = window.XLSX.read(data, { type: 'array' });
        processWorkbook(workbook, file.name, file);
        alert(`Planilha "${file.name}" carregada com sucesso! (${state.data.projetos.length} projetos processados).`);
      } catch (err) {
        console.error('Erro ao processar planilha Excel:', err);
        alert('Erro ao processar o arquivo Excel.');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  async function loadAutoExcel() {
    if (!window.XLSX) return false;
    try {
      const res = await fetch('./Painel-de-Projetos.xlsx?t=' + Date.now());
      if (!res.ok) return false;
      const httpLastModified = res.headers.get('Last-Modified');
      const arrayBuffer = await res.arrayBuffer();
      const data = new Uint8Array(arrayBuffer);
      const workbook = window.XLSX.read(data, { type: 'array' });
      processWorkbook(workbook, 'Painel-de-Projetos.xlsx', null, httpLastModified);
      return true;
    } catch (err) {
      console.warn('Auto-carregamento de Painel-de-Projetos.xlsx indisponível via HTTP/fetch, mantendo base local:', err);
      return false;
    }
  }

  // ==========================================================================
  // INITIALIZATION
  // ==========================================================================
  async function init() {
    populateFilterOptions();
    initEvents();
    updateOverview();
    checkHashRoute();

    // Carregamento automático transparente da planilha XLSX se servido por servidor web
    await loadAutoExcel();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }

})();
