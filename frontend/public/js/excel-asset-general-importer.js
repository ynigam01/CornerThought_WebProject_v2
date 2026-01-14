// excel-asset-general-importer.js
// Importer for the "Excel - General" file type in Organization Settings > Manage Assets > Asset Details.

import * as XLSX from 'https://cdn.sheetjs.com/xlsx-0.20.1/package/xlsx.mjs';

function toTrimmedStringOrEmpty(value) {
    return value != null ? String(value).trim() : '';
}

function toNullIfEmpty(value) {
    const s = toTrimmedStringOrEmpty(value);
    return s ? s : null;
}

function chunkArray(arr, size) {
    if (!Array.isArray(arr) || arr.length === 0) return [];
    const n = Math.max(1, Number(size) || 1);
    const chunks = [];
    for (let i = 0; i < arr.length; i += n) {
        chunks.push(arr.slice(i, i + n));
    }
    return chunks;
}

async function insertChunked({ supabase, table, rows, chunkSize, onProgress, label }) {
    if (!rows || rows.length === 0) return [];
    const chunks = chunkArray(rows, chunkSize);
    const inserted = [];

    for (let i = 0; i < chunks.length; i++) {
        const chunk = chunks[i];
        if (typeof onProgress === 'function') {
            onProgress(`${label} (${i + 1}/${chunks.length})...`);
        }

        const { data, error } = await supabase
            .from(table)
            .insert(chunk)
            .select();

        if (error) {
            throw new Error(`Failed inserting into ${table}: ${error.message || String(error)}`);
        }
        if (data && data.length) inserted.push(...data);
    }

    return inserted;
}

function normalizeHeader(value) {
    return String(value || '')
        .trim()
        .toLowerCase()
        .replace(/\s+/g, ' ');
}

function sheetRows(workbook, sheetName) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) return [];
    return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
}

function detectSheetsByHeaders(workbook) {
    const names = workbook.SheetNames || [];
    let componentsSheetName = null;
    let depsSheetName = null;

    for (const name of names) {
        const rows = sheetRows(workbook, name);
        const header = rows && rows[0] ? rows[0] : [];
        const keys = new Set(header.map(normalizeHeader));

        const hasUid = keys.has('uid');
        const hasComponentId = keys.has('component id');
        const hasUpstream = keys.has('upstream dependency');

        if (!componentsSheetName && hasUid && hasComponentId) {
            componentsSheetName = name;
        }
        if (!depsSheetName && hasComponentId && hasUpstream) {
            depsSheetName = name;
        }
    }

    if (!componentsSheetName) {
        throw new Error('Could not find the Asset Components sheet. Expected headers include "UID" and "Component ID".');
    }
    if (!depsSheetName) {
        throw new Error('Could not find the Dependencies sheet. Expected headers include "Component ID" and "Upstream Dependency".');
    }

    return { componentsSheetName, depsSheetName };
}

function parseAssetComponentsSheetRows(rows) {
    if (!Array.isArray(rows) || rows.length === 0) {
        throw new Error('The Asset Components sheet is empty.');
    }

    const headerRow = rows[0] || [];
    let idxUid = -1;
    let idxSystem = -1;
    let idxComponentId = -1;
    let idxComponentName = -1;
    let idxFunction = -1;
    let idxCriticality = -1;

    headerRow.forEach((cell, idx) => {
        const key = normalizeHeader(cell);
        if (key === 'uid') idxUid = idx;
        if (key === 'system') idxSystem = idx;
        if (key === 'component id') idxComponentId = idx;
        if (key === 'component name') idxComponentName = idx;
        if (key === 'function') idxFunction = idx;
        if (key === 'criticality') idxCriticality = idx;
    });

    if (idxUid === -1 || idxSystem === -1 || idxComponentId === -1 || idxComponentName === -1 || idxFunction === -1 || idxCriticality === -1) {
        throw new Error('Asset Components sheet is missing required headers: UID, System, Component ID, Component Name, Function, Criticality.');
    }

    const parsed = [];
    const errors = [];

    for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row) continue;

        const uid = toTrimmedStringOrEmpty(row[idxUid]);
        const system = toNullIfEmpty(row[idxSystem]);
        const component_id = toTrimmedStringOrEmpty(row[idxComponentId]);
        const component_name = toNullIfEmpty(row[idxComponentName]);
        const fn = toNullIfEmpty(row[idxFunction]);
        const criticality = toNullIfEmpty(row[idxCriticality]);

        const isBlankRow = !uid && !system && !component_id && !component_name && !fn && !criticality;
        if (isBlankRow) continue;

        if (!uid) errors.push(`Row ${i + 1}: "UID" is required.`);
        if (!component_id) errors.push(`Row ${i + 1}: "Component ID" is required.`);

        parsed.push({
            uid,
            system,
            component_id,
            component_name,
            function: fn,
            criticality
        });
    }

    if (parsed.length === 0) {
        throw new Error('No Asset Components data rows were found under the required headers.');
    }

    if (errors.length) {
        const first = errors.slice(0, 5).join(' ');
        const more = errors.length > 5 ? ` (and ${errors.length - 5} more)` : '';
        throw new Error(`Some rows are missing required fields. ${first}${more}`);
    }

    return parsed;
}

function parseDependenciesSheetRows(rows) {
    if (!Array.isArray(rows) || rows.length === 0) {
        throw new Error('The Dependencies sheet is empty.');
    }

    const headerRow = rows[0] || [];
    let idxComponentId = -1;
    let idxUpstream = -1;

    headerRow.forEach((cell, idx) => {
        const key = normalizeHeader(cell);
        if (key === 'component id') idxComponentId = idx;
        if (key === 'upstream dependency') idxUpstream = idx;
    });

    if (idxComponentId === -1 || idxUpstream === -1) {
        throw new Error('Dependencies sheet is missing required headers: Component ID, Upstream Dependency.');
    }

    const parsed = [];
    const errors = [];

    for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row) continue;

        const component_id = toTrimmedStringOrEmpty(row[idxComponentId]);
        const upstream_id = toNullIfEmpty(row[idxUpstream]); // may be null

        const isBlankRow = !component_id && !upstream_id;
        if (isBlankRow) continue;

        if (!component_id) errors.push(`Row ${i + 1}: "Component ID" is required.`);

        parsed.push({
            component_id,
            upstream_id
        });
    }

    if (parsed.length === 0) {
        // It is valid for the dependencies sheet to have no data rows.
        return [];
    }

    if (errors.length) {
        const first = errors.slice(0, 5).join(' ');
        const more = errors.length > 5 ? ` (and ${errors.length - 5} more)` : '';
        throw new Error(`Some dependency rows are missing required fields. ${first}${more}`);
    }

    return parsed;
}

/**
 * Import an Excel - General Asset workbook into Supabase.
 *
 * Tables written:
 * - asset_components
 * - asset_component_dependencies
 *
 * @param {object} params
 * @param {any} params.supabase - Supabase client (from supabase-js)
 * @param {File} params.file - Uploaded Excel file (.xlsx or .xls)
 * @param {object} params.context - Required context for inserts
 * @param {number} params.context.organization_id
 * @param {number} params.context.created_by
 * @param {number} params.context.asset_id
 * @param {string} params.context.asset_name
 * @param {number|null} params.context.project_id
 * @param {string|null} params.context.project
 * @param {(msg: string) => void} [params.onProgress]
 * @param {number} [params.chunkSize]
 * @returns {Promise<object>} Summary counts
 */
export async function importAssetGeneralExcelToSupabase({
    supabase,
    file,
    context,
    onProgress,
    chunkSize = 250
}) {
    if (!supabase) throw new Error('Supabase client is required.');
    if (!file) throw new Error('File is required.');
    if (!context) throw new Error('Context is required.');

    const organization_id = context.organization_id;
    const created_by = context.created_by;
    const asset_id = context.asset_id;
    const asset_name = context.asset_name;
    const project_id = context.project_id == null ? null : context.project_id;
    const project = context.project == null ? null : context.project;

    if (organization_id == null || created_by == null || asset_id == null || !asset_name) {
        throw new Error('Missing required context: organization_id, created_by, asset_id, asset_name.');
    }

    if (typeof onProgress === 'function') {
        onProgress('Reading Excel file...');
    }

    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });

    const { componentsSheetName, depsSheetName } = detectSheetsByHeaders(workbook);

    const componentsRows = sheetRows(workbook, componentsSheetName);
    const depsRows = sheetRows(workbook, depsSheetName);

    const parsedComponents = parseAssetComponentsSheetRows(componentsRows);
    const parsedDeps = parseDependenciesSheetRows(depsRows);

    const componentInserts = parsedComponents.map(r => ({
        organization_id,
        asset_id,
        created_by,
        uid: r.uid,
        asset_name,
        system: r.system,
        component_id: r.component_id,
        component_name: r.component_name,
        function: r.function,
        criticality: r.criticality,
        original: true,
        project: project,
        project_id: project_id
    }));

    const insertedComponents = await insertChunked({
        supabase,
        table: 'asset_components',
        rows: componentInserts,
        chunkSize,
        onProgress,
        label: 'Saving asset components'
    });

    const depsInserts = parsedDeps.map(r => ({
        organization_id,
        asset_id,
        created_by,
        component_id: r.component_id,
        upstream_id: r.upstream_id
    }));

    const insertedDeps = depsInserts.length
        ? await insertChunked({
            supabase,
            table: 'asset_component_dependencies',
            rows: depsInserts,
            chunkSize,
            onProgress,
            label: 'Saving asset dependencies'
        })
        : [];

    return {
        inserted: {
            asset_components: insertedComponents.length,
            asset_component_dependencies: insertedDeps.length
        },
        sheets: {
            components: componentsSheetName,
            dependencies: depsSheetName
        }
    };
}

