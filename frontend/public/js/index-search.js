// frontend/public/js/index-search.js
// This script powers the public lessons-learned search on the index page.
//
// It assumes a Postgres VIEW named `public_lessons_search` exists in your Supabase
// database with (at least) the following columns:
//
//   metadata_text                  -- text version of the metadata jsonb column
//   lesson_title                   -- lessons_learned.title (Issue/Success text)
//   lesson_category                -- lessons_learned.category ("Issue" / "Success")
//   future_project_consideration   -- future_project_considerations.fpc
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
//     fpc.fpc as future_project_consideration,
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

document.addEventListener('DOMContentLoaded', () => {
  const form = document.getElementById('lessonsSearchForm');
  const input = document.getElementById('lessonsSearchInput');
  const projectsButton = document.getElementById('searchProjectsButton');

  if (!form || !input) {
    // If elements aren't present, quietly exit.
    return;
  }

  form.addEventListener('submit', async (event) => {
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
      const results = await searchLessons(term);
      renderResults(results, term);
    } catch (error) {
      // eslint-disable-next-line no-console
      console.error('Error searching lessons:', error);
      setStatus('Something went wrong while searching. Please try again.');
      clearResults();
    }
  });

  if (projectsButton) {
    projectsButton.addEventListener('click', async () => {
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
});

/**
 * Calls Supabase to search the lessons view by metadata keyword.
 * Returns an array of normalized result objects ready for rendering.
 * Each object represents one future project consideration.
 */
async function searchLessons(term) {
  const { data, error } = await supabase
    .from(SEARCH_VIEW)
    .select(
      `
        metadata_text,
        lesson_title,
        lesson_category,
        future_project_consideration,
        project_name,
        project_type,
        industry
      `
    )
    .ilike('metadata_text', `%${term}%`);

  if (error) {
    throw error;
  }

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

  results.forEach((item) => {
    const card = document.createElement('article');
    card.className = 'lesson-card';

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

    container.appendChild(card);
  });
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
