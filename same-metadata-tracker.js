const stemmerModule = require('stemmer');

const stemWord =
  typeof stemmerModule === 'function'
    ? stemmerModule
    : typeof stemmerModule?.stemmer === 'function'
    ? stemmerModule.stemmer
    : (word) => word;

function normalizeWhitespace(value) {
  return String(value || '')
    .trim()
    .replace(/\s+/g, ' ');
}

function normalizeTerm(value) {
  const compact = normalizeWhitespace(value).toLowerCase();
  if (!compact) return '';
  return compact
    .replace(/[_-]+/g, ' ')
    .replace(/[^\w\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function tokenizeForStem(term) {
  const normalized = normalizeTerm(term);
  if (!normalized) return [];
  const tokens = normalized.match(/[a-z0-9]+/g) || [];
  return tokens.filter((token) => token.length > 1);
}

function toStemSet(term) {
  const stems = tokenizeForStem(term).map((token) => stemWord(token));
  return new Set(stems.filter(Boolean));
}

function tokenRootsMatch(a, b) {
  if (!a || !b) return false;
  if (a === b) return true;
  if (a.length < 4 || b.length < 4) return false;
  return a.startsWith(b) || b.startsWith(a);
}

function getStemOverlap(aSet, bSet) {
  const aTokens = Array.from(aSet || []);
  const bTokens = Array.from(bSet || []);
  if (!aTokens.length || !bTokens.length) {
    return { intersection: 0, union: aTokens.length + bTokens.length };
  }

  let intersection = 0;
  const usedB = new Set();
  for (let i = 0; i < aTokens.length; i += 1) {
    const aToken = aTokens[i];
    for (let j = 0; j < bTokens.length; j += 1) {
      if (usedB.has(j)) continue;
      if (tokenRootsMatch(aToken, bTokens[j])) {
        usedB.add(j);
        intersection += 1;
        break;
      }
    }
  }

  const union = aTokens.length + bTokens.length - intersection;
  return { intersection, union };
}

function stemSetSimilarity(aSet, bSet) {
  if (!aSet.size || !bSet.size) return 0;
  const { intersection, union } = getStemOverlap(aSet, bSet);
  if (!intersection) return 0;
  return intersection / union;
}

function isCloseMatch(aSet, bSet) {
  const similarity = stemSetSimilarity(aSet, bSet);
  if (similarity >= 0.5) return true;

  if (!aSet.size || !bSet.size) return false;
  const { intersection } = getStemOverlap(aSet, bSet);
  if (!intersection) return false;

  const minSize = Math.min(aSet.size, bSet.size);
  const maxSize = Math.max(aSet.size, bSet.size);
  return intersection === minSize && maxSize <= minSize + 1;
}

function extractMetadataTerms(value, output = []) {
  if (value == null) {
    return output;
  }

  if (typeof value === 'string') {
    const term = normalizeTerm(value);
    if (term) output.push(term);
    return output;
  }

  if (typeof value === 'number' || typeof value === 'boolean') {
    const term = normalizeTerm(String(value));
    if (term) output.push(term);
    return output;
  }

  if (Array.isArray(value)) {
    value.forEach((item) => extractMetadataTerms(item, output));
    return output;
  }

  if (typeof value === 'object') {
    Object.values(value).forEach((item) => extractMetadataTerms(item, output));
  }

  return output;
}

function rankMetadataTerms(terms, limit = 10) {
  const exactCounts = new Map();
  for (const term of terms) {
    if (!term) continue;
    exactCounts.set(term, (exactCounts.get(term) || 0) + 1);
  }

  const uniqueTerms = Array.from(exactCounts.keys());
  const stemSets = new Map(uniqueTerms.map((term) => [term, toStemSet(term)]));

  const ranked = uniqueTerms.map((term) => {
    const exactCount = exactCounts.get(term) || 0;
    let closeCount = 0;
    const aSet = stemSets.get(term) || new Set();
    for (const otherTerm of uniqueTerms) {
      if (otherTerm === term) continue;
      const bSet = stemSets.get(otherTerm) || new Set();
      if (isCloseMatch(aSet, bSet)) {
        closeCount += exactCounts.get(otherTerm) || 0;
      }
    }
    return {
      term,
      exactCount,
      closeCount,
      totalCount: exactCount + closeCount,
    };
  });

  ranked.sort((a, b) => {
    if (b.exactCount !== a.exactCount) return b.exactCount - a.exactCount;
    if (b.closeCount !== a.closeCount) return b.closeCount - a.closeCount;
    return a.term.localeCompare(b.term);
  });

  return ranked.slice(0, limit);
}

async function getTopMetadataTermsForProjectType({ supabase, projectTypeId, limit = 10 }) {
  const normalizedProjectTypeId = normalizeWhitespace(projectTypeId);
  if (!normalizedProjectTypeId) {
    return [];
  }

  const { data: lessonRows, error: lessonError } = await supabase
    .from('lessons_learned')
    .select('id')
    .eq('project_type_id', normalizedProjectTypeId)
    .is('organization_id', null);

  if (lessonError) {
    throw lessonError;
  }

  const lessonIds = (lessonRows || []).map((row) => row?.id).filter((id) => id != null);
  // #region agent log
  if (typeof fetch === 'function') {
    fetch('http://127.0.0.1:7242/ingest/3f684587-b61e-4851-8662-761311dbc082',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({runId:'debug-run-1',hypothesisId:'H15',location:'same-metadata-tracker.js:getTopMetadataTermsForProjectType:lessonIds',message:'metadata tracker loaded lesson ids',data:{projectTypeId:normalizedProjectTypeId,lessonCount:lessonIds.length},timestamp:Date.now()})}).catch(()=>{});
  }
  // #endregion
  if (!lessonIds.length) {
    return [];
  }

  const { data: metadataRows, error: metadataError } = await supabase
    .from('lessons_learned_metadata')
    .select('metadata')
    .in('lessons_learned_id', lessonIds);

  if (metadataError) {
    throw metadataError;
  }

  const allTerms = [];
  (metadataRows || []).forEach((row) => {
    const terms = extractMetadataTerms(row?.metadata);
    allTerms.push(...terms);
  });
  // #region agent log
  if (typeof fetch === 'function') {
    fetch('http://127.0.0.1:7242/ingest/3f684587-b61e-4851-8662-761311dbc082',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({runId:'debug-run-1',hypothesisId:'H15',location:'same-metadata-tracker.js:getTopMetadataTermsForProjectType:terms',message:'metadata tracker extracted terms',data:{projectTypeId:normalizedProjectTypeId,metadataRowsCount:Array.isArray(metadataRows)?metadataRows.length:null,termCount:allTerms.length},timestamp:Date.now()})}).catch(()=>{});
  }
  // #endregion

  return rankMetadataTerms(allTerms, limit);
}

module.exports = {
  getTopMetadataTermsForProjectType,
  _internal: {
    normalizeTerm,
    extractMetadataTerms,
    rankMetadataTerms,
    toStemSet,
    stemSetSimilarity,
    isCloseMatch,
  },
};
