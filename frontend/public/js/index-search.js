// frontend/public/js/index-search.js
// This script powers the public lessons-learned search on the index page.
//
// It assumes a Postgres VIEW named `public_lessons_search` exists in your Supabase
// database with (at least) the following columns:
//
//   metadata_text                  -- text version of the metadata jsonb column
//   lesson_title                   -- lessons_learned.title (Issue/Success text)
//   lesson_category                -- lessons_learned.category ("Issue" / "Success")
//   lessons_learned_id             -- lessons_learned.id
//   fpc_id                         -- future_project_considerations.id
//   future_project_consideration   -- future_project_considerations.fpc
//   lessons_learned_cause_id       -- future_project_considerations.lessons_learned_cause_id
//   lessons_learned_impact_id      -- future_project_considerations.lessons_learned_impact_id
//   project_name                   -- projects.project_name (nullable)
//   project_type                   -- project_type.project_type (nullable)
//   industry                       -- project_type.industry (nullable)
//
// Example SQL you can run in Supabase to create the view (adjust table/column names if needed):
//
//   create or replace view public_lessons_search as
//   select
//     llm.id as metadata_id,
//     llm.lessons_learned_id,
//     llm.metadata::text as metadata_text,
//     ll.title as lesson_title,
//     ll.category as lesson_category,
//     ll.id as lessons_learned_id,
//     fpc.id as fpc_id,
//     fpc.fpc as future_project_consideration,
//     fpc.lessons_learned_cause_id,
//     fpc.lessons_learned_impact_id,
//     p.project_name,
//     pt.project_type,
//     pt.industry
//   from lessons_learned_metadata llm
//   join lessons_learned ll
//     on llm.lessons_learned_id = ll.id
//   left join future_project_considerations fpc
//     on fpc.lessons_learned_id = ll.id
//   left join projects p
//     on ll.project_id = p.project_id
//   left join project_type pt
//     on p.project_type_id = pt.id;
//
// The frontend uses Supabase to query this view with an ILIKE filter on `metadata_text`.

import { supabase } from './supabase-client.js';

const SEARCH_VIEW = 'public_lessons_search';
let lastResultsCache = null;
let lastTermCache = '';

document.addEventListener('DOMContentLoaded', () => {
  const form = document.getElementById('lessonsSearchForm');
  const input = document.getElementById('lessonsSearchInput');
  const projectsButton = document.getElementById('searchProjectsButton');
  const searchButton = form ? form.querySelector('button[type="submit"]') : null;

  if (!form || !input) {
    // If elements aren't present, quietly exit.
    return;
  }

  // #region agent log
  fetch('http://127.0.0.1:7242/ingest/3f684587-b61e-4851-8662-761311dbc082',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({runId:'pre-fix',hypothesisId:'H1',location:'index-search.js:49',message:'DOMContentLoaded wired',data:{hasForm:!!form,hasInput:!!input},timestamp:Date.now()})}).catch(()=>{});
  // #endregion

  form.addEventListener('submit', async (event) => {
    // #region agent log
    fetch('http://127.0.0.1:7242/ingest/3f684587-b61e-4851-8662-761311dbc082',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({runId:'pre-fix',hypothesisId:'H4',location:'index-search.js:63',message:'form submit fired',data:{submitterId:event?.submitter?.id||null,submitterText:event?.submitter?.textContent||null},timestamp:Date.now()})}).catch(()=>{});
    // #endregion
    event.preventDefault();
    const rawTerm = input.value || '';
    const term = rawTerm.trim();

    if (!term) {
      clearResults();
      setStatus('Please enter a keyword to search.');
      return;
    }

    setStatus('Searching…');
    try {
      // #region agent log
      fetch('http://127.0.0.1:7242/ingest/3f684587-b61e-4851-8662-761311dbc082',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({runId:'pre-fix',hypothesisId:'H2',location:'index-search.js:72',message:'searchLessons start',data:{termLength:term.length},timestamp:Date.now()})}).catch(()=>{});
      // #endregion
      const results = await searchLessons(term);
      // #region agent log
      fetch('http://127.0.0.1:7242/ingest/3f684587-b61e-4851-8662-761311dbc082',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({runId:'pre-fix',hypothesisId:'H2',location:'index-search.js:74',message:'searchLessons results',data:{count:Array.isArray(results)?results.length:null},timestamp:Date.now()})}).catch(()=>{});
      // #endregion
      renderResults(results, term);
    } catch (error) {
      // eslint-disable-next-line no-console
      console.error('Error searching lessons:', error);
      // #region agent log
      fetch('http://127.0.0.1:7242/ingest/3f684587-b61e-4851-8662-761311dbc082',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({runId:'pre-fix',hypothesisId:'H2',location:'index-search.js:80',message:'searchLessons error',data:{errorMessage:error?.message||'unknown'},timestamp:Date.now()})}).catch(()=>{});
      // #endregion
      setStatus('Something went wrong while searching. Please try again.');
      clearResults();
    }
  });

  if (projectsButton) {
    projectsButton.addEventListener('click', async () => {
      // #region agent log
      fetch('http://127.0.0.1:7242/ingest/3f684587-b61e-4851-8662-761311dbc082',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({runId:'pre-fix',hypothesisId:'H4',location:'index-search.js:88',message:'projects button clicked',data:{},timestamp:Date.now()})}).catch(()=>{});
      // #endregion
      const rawTerm = input.value || '';
      const term = rawTerm.trim();

      if (!term) {
        clearResults();
        setStatus('Please enter a keyword to search projects.');
        return;
      }

      setStatus('Searching projects…');
      try {
        const results = await searchProjects(term);
        renderProjectResults(results, term);
      } catch (error) {
        // eslint-disable-next-line no-console
        console.error('Error searching projects:', error);
        setStatus('Something went wrong while searching projects. Please try again.');
        clearResults();
      }
    });
  }

  if (searchButton) {
    searchButton.addEventListener('click', () => {
      // #region agent log
      fetch('http://127.0.0.1:7242/ingest/3f684587-b61e-4851-8662-761311dbc082',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({runId:'pre-fix',hypothesisId:'H4',location:'index-search.js:109',message:'search lessons button clicked',data:{},timestamp:Date.now()})}).catch(()=>{});
      // #endregion
    });
  }
});

/**
 * Calls Supabase to search the lessons view by metadata keyword.
 * Returns an array of normalized result objects ready for rendering.
 * Each object represents one future project consideration.
 */
async function searchLessons(term) {
  const extendedSelect = `
        metadata_text,
        lesson_title,
        lesson_category,
        lessons_learned_id,
        fpc_id,
        future_project_consideration,
        lessons_learned_cause_id,
        lessons_learned_impact_id,
        project_name,
        project_type,
        industry
      `;
  const legacySelect = `
        metadata_text,
        lesson_title,
        lesson_category,
        future_project_consideration,
        project_name,
        project_type,
        industry
      `;

  let { data, error } = await supabase
    .from(SEARCH_VIEW)
    .select(extendedSelect)
    .ilike('metadata_text', `%${term}%`);

  // #region agent log
  fetch('http://127.0.0.1:7242/ingest/3f684587-b61e-4851-8662-761311dbc082',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({runId:'pre-fix',hypothesisId:'H3',location:'index-search.js:132',message:'Supabase view response',data:{hasData:!!data,rows:Array.isArray(data)?data.length:null,hasError:!!error,errorMessage:error?.message||null},timestamp:Date.now()})}).catch(()=>{});
  // #endregion

  if (error) {
    const message = String(error?.message || '');
    const isMissingColumn =
      message.includes('does not exist') &&
      (message.includes('fpc_id') ||
        message.includes('lessons_learned_cause_id') ||
        message.includes('lessons_learned_impact_id') ||
        message.includes('lessons_learned_id'));

    if (isMissingColumn) {
      const fallback = await supabase
        .from(SEARCH_VIEW)
        .select(legacySelect)
        .ilike('metadata_text', `%${term}%`);
      data = fallback.data;
      error = fallback.error;
    }
  }

  if (error) throw error;

  if (!data) {
    return [];
  }

  return data.map((row) => normalizeResultRow(row));
}

/**
 * Calls backend /api/search-projects to perform semantic project search.
 */
async function searchProjects(term) {
  const response = await fetch('/api/search-projects', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ queryText: term }),
  });

  if (!response.ok) {
    const text = await response.text().catch(() => '');
    throw new Error(`Project search failed with ${response.status}: ${text}`);
  }

  const payload = await response.json();
  return Array.isArray(payload.results) ? payload.results : [];
}

/**
 * Normalizes a raw row from Supabase into the shape needed for the UI.
 */
function normalizeResultRow(row) {
  const categoryRaw = (row.lesson_category || '').toString().toLowerCase();
  const isSuccess = categoryRaw === 'success';
  const isIssue = categoryRaw === 'issue';

  const categoryLabel = isSuccess
    ? 'Success'
    : isIssue
    ? 'Issue'
    : (row.lesson_category || 'Lesson');

  return {
    futureProjectConsideration: row.future_project_consideration || '',
    issueOrSuccessText: row.lesson_title || '',
    categoryLabel,
    categoryType: isSuccess ? 'success' : isIssue ? 'issue' : 'other',
    projectName: row.project_name || null,
    projectType: row.project_type || null,
    industry: row.industry || null,
    lessonsLearnedId: row.lessons_learned_id ?? null,
    fpcId: row.fpc_id ?? null,
    fpcCauseId: row.lessons_learned_cause_id ?? null,
    fpcImpactId: row.lessons_learned_impact_id ?? null,
  };
}

/**
 * Renders project search results as "<Project Name>: <Project Description>" lines.
 */
function renderProjectResults(results, term) {
  const container = document.getElementById('lessonsResultsContainer');
  if (!container) return;

  container.innerHTML = '';

  if (!results.length) {
    setStatus(`No projects found for “${term}”.`);
    return;
  }

  setStatus(`Showing ${results.length} project result${results.length === 1 ? '' : 's'} for “${term}”.`);

  results.forEach((item) => {
    const div = document.createElement('div');
    div.className = 'project-result-line';

    const name = item.project_name || '(Unnamed Project)';
    const desc = item.project_description || '';

    div.textContent = `${name}: ${desc}`;
    container.appendChild(div);
  });
}

/**
 * Renders all search results as lesson cards on the page.
 */
function renderResults(results, term) {
  const container = document.getElementById('lessonsResultsContainer');
  if (!container) return;

  container.innerHTML = '';

  if (!results.length) {
    setStatus(`No lessons found for “${term}”.`);
    return;
  }

  setStatus(`Showing ${results.length} result${results.length === 1 ? '' : 's'} for “${term}”.`);
  lastResultsCache = results;
  lastTermCache = term;

  // #region agent log
  const resultProjectNameCount = results.filter((item) => !!item?.projectName).length;
  const resultProjectTypeCount = results.filter((item) => !!item?.projectType).length;
  fetch('http://127.0.0.1:7242/ingest/3f684587-b61e-4851-8662-761311dbc082',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({runId:'pre-fix',hypothesisId:'H10',location:'index-search.js:227',message:'renderResults',data:{count:results.length,projectNameCount:resultProjectNameCount,projectTypeCount:resultProjectTypeCount},timestamp:Date.now()})}).catch(()=>{});
  // #endregion

  results.forEach((item) => {
    const card = document.createElement('article');
    card.className = 'lesson-card lesson-card-clickable';
    card.setAttribute('role', 'button');
    card.setAttribute('tabindex', '0');

    // Future Project Consideration line
    const fpc = document.createElement('div');
    fpc.className = 'lesson-card-fpc';
    fpc.innerHTML = `<span class="lesson-card-fpc-label">Future Project Consideration</span>: ${escapeHtml(
      item.futureProjectConsideration
    )}`;

    // Issue / Success inner box
    const issueBox = document.createElement('div');
    issueBox.className =
      item.categoryType === 'success'
        ? 'lesson-card-inner lesson-card-success-box'
        : item.categoryType === 'issue'
        ? 'lesson-card-inner lesson-card-issue-box'
        : 'lesson-card-inner lesson-card-generic-box';

    issueBox.innerHTML = `<strong>${escapeHtml(
      item.categoryLabel
    )}:</strong> ${escapeHtml(item.issueOrSuccessText)}`;

    // Project metadata line (conditionally show project name)
    const meta = document.createElement('div');
    meta.className = 'lesson-card-meta';

    const metaParts = [];

    if (item.projectName) {
      metaParts.push(`<span>Project: ${escapeHtml(item.projectName)}</span>`);
    }
    if (item.projectType) {
      metaParts.push(`<span>Project Type: ${escapeHtml(item.projectType)}</span>`);
    }
    if (item.industry) {
      metaParts.push(`<span>Industry: ${escapeHtml(item.industry)}</span>`);
    }

    meta.innerHTML = metaParts.join('   ');

    card.appendChild(fpc);
    card.appendChild(issueBox);
    if (metaParts.length) {
      card.appendChild(meta);
    }

    const handleClick = () => {
      // #region agent log
      fetch('http://127.0.0.1:7242/ingest/3f684587-b61e-4851-8662-761311dbc082',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({runId:'pre-fix',hypothesisId:'H6',location:'index-search.js:257',message:'result clicked',data:{lessonsLearnedId:item?.lessonsLearnedId||null,fpcId:item?.fpcId||null,fpcCauseId:item?.fpcCauseId||null,fpcImpactId:item?.fpcImpactId||null},timestamp:Date.now()})}).catch(()=>{});
      // #endregion
      showLessonDetailFromResult(item);
    };

    card.addEventListener('click', handleClick);
    card.addEventListener('keydown', (event) => {
      if (event.key === 'Enter' || event.key === ' ') {
        event.preventDefault();
        handleClick();
      }
    });

    container.appendChild(card);
  });
}

async function showLessonDetailFromResult(result) {
  if (!result || !result.lessonsLearnedId) {
    // #region agent log
    fetch('http://127.0.0.1:7242/ingest/3f684587-b61e-4851-8662-761311dbc082',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({runId:'pre-fix',hypothesisId:'H6',location:'index-search.js:277',message:'missing lesson id',data:{resultKeys:result?Object.keys(result):null,lessonsLearnedId:result?.lessonsLearnedId||null},timestamp:Date.now()})}).catch(()=>{});
    // #endregion
    setStatus('Resolving lesson details…');
    try {
      const resolved = await resolveLegacyLessonIds(result);
      if (resolved && resolved.lessonsLearnedId) {
        result.lessonsLearnedId = resolved.lessonsLearnedId;
        result.fpcId = resolved.fpcId;
        result.fpcCauseId = resolved.fpcCauseId;
        result.fpcImpactId = resolved.fpcImpactId;
      } else {
        setStatus('Missing lesson id for this result.');
        return;
      }
    } catch (error) {
      // eslint-disable-next-line no-console
      console.error('Error resolving lesson ids:', error);
      setStatus('Missing lesson id for this result.');
      return;
    }
  }

  setStatus('Loading lesson details…');
  try {
    // #region agent log
    fetch('http://127.0.0.1:7242/ingest/3f684587-b61e-4851-8662-761311dbc082',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({runId:'pre-fix',hypothesisId:'H5',location:'index-search.js:290',message:'fetchLessonDetail start',data:{lessonId:result.lessonsLearnedId,fpcId:result.fpcId},timestamp:Date.now()})}).catch(()=>{});
    // #endregion
    const detail = await fetchLessonDetail(result.lessonsLearnedId);
    renderLessonDetail(detail, result);
    setStatus('Lesson details loaded.');
  } catch (error) {
    // eslint-disable-next-line no-console
    console.error('Error loading lesson details:', error);
    // #region agent log
    fetch('http://127.0.0.1:7242/ingest/3f684587-b61e-4851-8662-761311dbc082',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({runId:'pre-fix',hypothesisId:'H5',location:'index-search.js:297',message:'fetchLessonDetail error',data:{errorMessage:error?.message||'unknown'},timestamp:Date.now()})}).catch(()=>{});
    // #endregion
    setStatus('Unable to load lesson details. Please try again.');
  }
}

async function resolveLegacyLessonIds(result) {
  const fpcText = result?.futureProjectConsideration || '';
  const titleText = result?.issueOrSuccessText || '';
  const categoryLabel = String(result?.categoryLabel || '').toLowerCase();

  // #region agent log
  fetch('http://127.0.0.1:7242/ingest/3f684587-b61e-4851-8662-761311dbc082',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({runId:'pre-fix',hypothesisId:'H7',location:'index-search.js:320',message:'resolveLegacyLessonIds start',data:{hasFpc:!!fpcText,hasTitle:!!titleText,categoryLabel},timestamp:Date.now()})}).catch(()=>{});
  // #endregion

  if (!fpcText) {
    return null;
  }

  const fpcResp = await supabase
    .from('future_project_considerations')
    .select('id, fpc, lessons_learned_id, lessons_learned_cause_id, lessons_learned_impact_id')
    .eq('fpc', fpcText);

  if (fpcResp.error) {
    throw fpcResp.error;
  }

  const fpcRows = Array.isArray(fpcResp.data) ? fpcResp.data : [];

  // #region agent log
  fetch('http://127.0.0.1:7242/ingest/3f684587-b61e-4851-8662-761311dbc082',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({runId:'pre-fix',hypothesisId:'H7',location:'index-search.js:337',message:'resolveLegacyLessonIds fpc rows',data:{count:fpcRows.length},timestamp:Date.now()})}).catch(()=>{});
  // #endregion

  if (fpcRows.length === 1) {
    const row = fpcRows[0];
    return {
      lessonsLearnedId: row.lessons_learned_id ?? null,
      fpcId: row.id ?? null,
      fpcCauseId: row.lessons_learned_cause_id ?? null,
      fpcImpactId: row.lessons_learned_impact_id ?? null,
    };
  }

  if (fpcRows.length > 1 && titleText) {
    const lessonIds = fpcRows
      .map((row) => row.lessons_learned_id)
      .filter((id) => id != null);
    if (lessonIds.length) {
      const lessonsResp = await supabase
        .from('lessons_learned')
        .select('id, title, category')
        .in('id', lessonIds);
      if (lessonsResp.error) {
        throw lessonsResp.error;
      }
      const lessons = Array.isArray(lessonsResp.data) ? lessonsResp.data : [];
      const match = lessons.find((lesson) => {
        const sameTitle = String(lesson.title || '').trim().toLowerCase() === titleText.trim().toLowerCase();
        const sameCategory =
          !categoryLabel ||
          String(lesson.category || '').trim().toLowerCase() === categoryLabel;
        return sameTitle && sameCategory;
      });
      if (match) {
        const row = fpcRows.find((fpc) => fpc.lessons_learned_id === match.id) || fpcRows[0];
        return {
          lessonsLearnedId: row.lessons_learned_id ?? null,
          fpcId: row.id ?? null,
          fpcCauseId: row.lessons_learned_cause_id ?? null,
          fpcImpactId: row.lessons_learned_impact_id ?? null,
        };
      }
    }
  }

  if (fpcRows.length > 0) {
    const row = fpcRows[0];
    return {
      lessonsLearnedId: row.lessons_learned_id ?? null,
      fpcId: row.id ?? null,
      fpcCauseId: row.lessons_learned_cause_id ?? null,
      fpcImpactId: row.lessons_learned_impact_id ?? null,
    };
  }

  if (titleText) {
    const lessonsResp = await supabase
      .from('lessons_learned')
      .select('id, title, category')
      .eq('title', titleText);
    if (lessonsResp.error) {
      throw lessonsResp.error;
    }
    const lessons = Array.isArray(lessonsResp.data) ? lessonsResp.data : [];
    const match = lessons.find((lesson) => {
      if (!categoryLabel) return true;
      return String(lesson.category || '').trim().toLowerCase() === categoryLabel;
    });
    if (match) {
      return { lessonsLearnedId: match.id ?? null, fpcId: null, fpcCauseId: null, fpcImpactId: null };
    }
  }

  return null;
}

async function fetchLessonDetail(lessonId) {
  const lessonResp = await supabase
    .from('lessons_learned')
    .select('id, title, category, project_id, project_type_id')
    .eq('id', lessonId)
    .single();

  if (lessonResp.error) throw lessonResp.error;

  const lesson = lessonResp.data || {};
  const projectId = lesson.project_id ?? null;
  let projectTypeId = lesson.project_type_id ?? null;

  const [causesResp, impactsResp, fpcResp, metadataResp, projectResp] = await Promise.all([
    supabase
      .from('lessons_learned_causes')
      .select('id, cause')
      .eq('lessons_learned_id', lessonId),
    supabase
      .from('lessons_learned_impacts')
      .select('id, impact')
      .eq('lessons_learned_id', lessonId),
    supabase
      .from('future_project_considerations')
      .select('id, fpc, lessons_learned_cause_id, lessons_learned_impact_id')
      .eq('lessons_learned_id', lessonId),
    supabase
      .from('lessons_learned_metadata')
      .select('metadata')
      .eq('lessons_learned_id', lessonId),
    projectId
      ? supabase
          .from('projects')
          .select('project_id, project_name, project_type_id')
          .eq('project_id', projectId)
          .single()
      : Promise.resolve({ data: null, error: null }),
  ]);

  if (causesResp.error) throw causesResp.error;
  if (impactsResp.error) throw impactsResp.error;
  if (fpcResp.error) throw fpcResp.error;
  if (metadataResp.error) throw metadataResp.error;
  if (projectResp.error) throw projectResp.error;

  const project = projectResp.data || null;
  if (!projectTypeId && project?.project_type_id) {
    projectTypeId = project.project_type_id;
  }

  let projectType = null;
  if (projectTypeId) {
    const projectTypeResp = await supabase
      .from('project_type')
      .select('id, project_type')
      .eq('id', projectTypeId)
      .single();
    if (projectTypeResp.error) throw projectTypeResp.error;
    projectType = projectTypeResp.data || null;
  }

  return {
    lesson,
    causes: Array.isArray(causesResp.data) ? causesResp.data : [],
    impacts: Array.isArray(impactsResp.data) ? impactsResp.data : [],
    fpcs: Array.isArray(fpcResp.data) ? fpcResp.data : [],
    metadata: Array.isArray(metadataResp.data) ? metadataResp.data : [],
    project,
    projectType,
  };
}

function renderLessonDetail(detail, resultMeta) {
  const container = document.getElementById('lessonsResultsContainer');
  if (!container) return;

  container.innerHTML = '';

  const lesson = detail.lesson || {};
  // #region agent log
  fetch('http://127.0.0.1:7242/ingest/3f684587-b61e-4851-8662-761311dbc082',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({runId:'pre-fix',hypothesisId:'H9',location:'index-search.js:559',message:'renderLessonDetail data',data:{resultProjectName:resultMeta?.projectName||null,resultProjectType:resultMeta?.projectType||null,lessonProjectId:lesson.project_id??null,projectName:detail?.project?.project_name||null,projectType:detail?.projectType?.project_type||null,metadataCount:Array.isArray(detail?.metadata)?detail.metadata.length:null},timestamp:Date.now()})}).catch(()=>{});
  // #endregion
  const categoryRaw = String(lesson.category || '').toLowerCase();
  const categoryLabel =
    categoryRaw === 'success' ? 'Success' : categoryRaw === 'issue' ? 'Issue' : 'Lesson';

  const expandCauseId = resultMeta?.fpcCauseId ?? null;
  const expandImpactId = resultMeta?.fpcImpactId ?? null;
  const clickedFpcId = resultMeta?.fpcId ?? null;

  const actions = document.createElement('div');
  actions.className = 'lesson-detail-actions';

  const backButton = document.createElement('button');
  backButton.type = 'button';
  backButton.className = 'lesson-detail-back-button';
  backButton.textContent = 'Back to Results';
  backButton.addEventListener('click', () => {
    if (Array.isArray(lastResultsCache) && lastResultsCache.length) {
      renderResults(lastResultsCache, lastTermCache || '');
      return;
    }
    clearResults();
    setStatus('Please run a search to see results.');
  });

  actions.appendChild(backButton);

  const card = document.createElement('article');
  card.className = 'lesson-detail-card';

  const title = document.createElement('div');
  title.className = 'lesson-detail-title';
  title.innerHTML = `<strong>${escapeHtml(categoryLabel)}:</strong> ${escapeHtml(
    lesson.title || ''
  )}`;

  const causesSection = buildExpandableSection({
    label: 'Cause(s):',
    items: detail.causes,
    itemKey: 'cause',
    itemIdKey: 'id',
    fpcs: detail.fpcs,
    fpcLinkKey: 'lessons_learned_cause_id',
    defaultExpandedId: expandCauseId,
    clickedFpcId,
    sectionClass: 'lesson-detail-cause-card',
  });

  const impactsSection = buildExpandableSection({
    label: 'Impact(s):',
    items: detail.impacts,
    itemKey: 'impact',
    itemIdKey: 'id',
    fpcs: detail.fpcs,
    fpcLinkKey: 'lessons_learned_impact_id',
    defaultExpandedId: expandImpactId,
    clickedFpcId,
    sectionClass: 'lesson-detail-impact-card',
  });

  const projectMetaBlock = buildProjectMetaBlock(detail);

  card.appendChild(title);
  card.appendChild(causesSection);
  card.appendChild(impactsSection);
  if (projectMetaBlock) {
    card.appendChild(projectMetaBlock);
  }

  container.appendChild(actions);
  container.appendChild(card);
}

function buildProjectMetaBlock(detail) {
  const projectName = detail?.project?.project_name || '';
  const projectType = detail?.projectType?.project_type || '';
  const metadataValues = (Array.isArray(detail?.metadata) ? detail.metadata : [])
    .map((row) => normalizeMetadataValue(row?.metadata))
    .filter((value) => value);

  const hasProjectType = !!projectType;
  const hasMetadata = metadataValues.length > 0;

  // #region agent log
  fetch('http://127.0.0.1:7242/ingest/3f684587-b61e-4851-8662-761311dbc082',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({runId:'pre-fix',hypothesisId:'H8',location:'index-search.js:664',message:'buildProjectMetaBlock data',data:{hasProjectName:!!projectName,projectName:projectName||null,projectType:projectType||null,metadataCount:metadataValues.length,projectId:detail?.project?.project_id||null,lessonProjectId:detail?.lesson?.project_id||null},timestamp:Date.now()})}).catch(()=>{});
  // #endregion

  if (!projectName && !hasProjectType && !hasMetadata) {
    return null;
  }

  const wrapper = document.createElement('div');
  wrapper.className = 'lesson-detail-project-meta';

  if (projectName) {
    const projectRow = document.createElement('div');
    projectRow.className = 'lesson-detail-project-row';
    projectRow.innerHTML = `<span class="lesson-detail-project-label">Project:</span> ${escapeHtml(
      projectName
    )}`;
    wrapper.appendChild(projectRow);
  }

  if (hasProjectType) {
    const typeRow = document.createElement('div');
    typeRow.className = 'lesson-detail-project-row';
    typeRow.innerHTML = `<span class="lesson-detail-project-label">Project Type:</span> ${escapeHtml(
      projectType
    )}`;
    wrapper.appendChild(typeRow);
  }

  if (hasMetadata) {
    const metadataRow = document.createElement('div');
    metadataRow.className = 'lesson-detail-project-row';
    metadataRow.innerHTML = `<span class="lesson-detail-project-label">Metadata:</span> ${escapeHtml(
      metadataValues.join(', ')
    )}`;
    wrapper.appendChild(metadataRow);
  }

  return wrapper;
}

function normalizeMetadataValue(value) {
  if (value == null) return '';
  if (typeof value === 'string') return value.trim();
  if (typeof value === 'number' || typeof value === 'boolean') return String(value);
  if (Array.isArray(value)) {
    return value.map((item) => normalizeMetadataValue(item)).filter(Boolean).join(', ');
  }
  if (typeof value === 'object') {
    if (typeof value.metadata === 'string') return value.metadata.trim();
    if (typeof value.value === 'string') return value.value.trim();
    const primitiveValues = Object.values(value)
      .map((item) => normalizeMetadataValue(item))
      .filter(Boolean);
    if (primitiveValues.length) return primitiveValues.join(', ');
    try {
      return JSON.stringify(value);
    } catch (err) {
      return '';
    }
  }
  return '';
}

function buildExpandableSection({
  label,
  items,
  itemKey,
  itemIdKey,
  fpcs,
  fpcLinkKey,
  defaultExpandedId,
  clickedFpcId,
  sectionClass,
}) {
  const wrapper = document.createElement('section');
  wrapper.className = `lesson-detail-section ${sectionClass}`;

  const header = document.createElement('div');
  header.className = 'lesson-detail-section-header';
  header.textContent = label;
  wrapper.appendChild(header);

  const safeItems = Array.isArray(items) ? items : [];
  const safeFpcs = Array.isArray(fpcs) ? fpcs : [];

  safeItems.forEach((item) => {
    const itemId = item?.[itemIdKey];
    const row = document.createElement('div');
    row.className = 'lesson-detail-item';

    const rowHeader = document.createElement('button');
    rowHeader.type = 'button';
    rowHeader.className = 'lesson-detail-item-header';

    const labelSpan = document.createElement('span');
    labelSpan.className = 'lesson-detail-item-label';
    labelSpan.textContent = item?.[itemKey] || '';

    const toggle = document.createElement('span');
    toggle.className = 'lesson-detail-toggle';

    rowHeader.appendChild(labelSpan);
    rowHeader.appendChild(toggle);

    const body = document.createElement('div');
    body.className = 'lesson-detail-item-body';

    const linkedFpcs = safeFpcs.filter((fpc) => fpc?.[fpcLinkKey] === itemId);
    if (linkedFpcs.length) {
      const list = document.createElement('ul');
      list.className = 'lesson-detail-fpc-list';
      linkedFpcs.forEach((fpc) => {
        const li = document.createElement('li');
        const isClicked = clickedFpcId != null && fpc?.id === clickedFpcId;
        li.className = `lesson-detail-fpc-item ${
          isClicked ? 'lesson-fpc-highlight' : 'lesson-fpc-default'
        }`;
        li.textContent = fpc?.fpc || '';
        list.appendChild(li);
      });
      body.appendChild(list);
    }

    const isExpanded = defaultExpandedId != null && itemId === defaultExpandedId;
    applyExpandState(rowHeader, body, toggle, isExpanded);

    rowHeader.addEventListener('click', () => {
      const willExpand = body.style.display !== 'block';
      applyExpandState(rowHeader, body, toggle, willExpand);
    });

    row.appendChild(rowHeader);
    row.appendChild(body);
    wrapper.appendChild(row);
  });

  return wrapper;
}

function applyExpandState(header, body, toggle, isExpanded) {
  body.style.display = isExpanded ? 'block' : 'none';
  toggle.innerHTML = isExpanded ? '&#9650;' : '&#9660;';
  if (isExpanded) {
    header.classList.add('is-expanded');
  } else {
    header.classList.remove('is-expanded');
  }
}

function clearResults() {
  const container = document.getElementById('lessonsResultsContainer');
  if (container) {
    container.innerHTML = '';
  }
}

function setStatus(message) {
  const statusEl = document.getElementById('lessonsSearchStatus');
  if (statusEl) {
    statusEl.textContent = message;
  }
}

// Basic HTML escaping to avoid accidentally injecting HTML into the page.
function escapeHtml(value) {
  if (value == null) return '';
  return value
    .toString()
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
