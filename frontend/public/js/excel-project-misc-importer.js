// excel-project-misc-importer.js
// Importer for the "Excel Project Miscellaneous List - New" file type used in the
// Lessons Learned Metadata sub-module of the user portal.

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

async function insertChunked({ supabase, table, rows, select, chunkSize, onProgress, label }) {
    if (!rows || rows.length === 0) return [];
    const chunks = chunkArray(rows, chunkSize);
    const inserted = [];

    for (let i = 0; i < chunks.length; i++) {
        const chunk = chunks[i];
        if (typeof onProgress === 'function') {
            onProgress(`${label} (${i + 1}/${chunks.length})...`);
        }

        const query = supabase
            .from(table)
            .insert(chunk);

        const { data, error } = select
            ? await query.select(select)
            : await query.select();

        if (error) {
            throw new Error(`Failed inserting into ${table}: ${error.message || String(error)}`);
        }
        if (data && data.length) inserted.push(...data);
    }

    return inserted;
}

function parseProjectMiscSheetRows(rows) {
    if (!Array.isArray(rows) || rows.length === 0) {
        throw new Error('The Excel file is empty.');
    }

    const headerRow = rows[0] || [];
    const normalize = (v) => String(v || '').trim().toLowerCase();

    let idxItemId = -1;
    let idxItem = -1;
    let idxDescription = -1;
    let idxParentItemId = -1;

    headerRow.forEach((cell, idx) => {
        const key = normalize(cell);
        if (key === 'item id') idxItemId = idx;
        if (key === 'item') idxItem = idx;
        if (key === 'description') idxDescription = idx;
        if (key === 'parent item id') idxParentItemId = idx;
    });

    if (idxItemId === -1 || idxItem === -1 || idxDescription === -1 || idxParentItemId === -1) {
        throw new Error('Could not find required headers "Item ID", "Item", "Description", and "Parent Item ID".');
    }

    const parsed = [];
    const errors = [];

    for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row) continue;

        const itemId = toTrimmedStringOrEmpty(row[idxItemId]);
        const item = toTrimmedStringOrEmpty(row[idxItem]);
        const description = toNullIfEmpty(row[idxDescription]);
        const parentItemId = toNullIfEmpty(row[idxParentItemId]);

        const isBlankRow = !itemId && !item && !description && !parentItemId;
        if (isBlankRow) continue;

        if (!itemId) errors.push(`Row ${i + 1}: "Item ID" is required.`);
        if (!item) errors.push(`Row ${i + 1}: "Item" is required.`);

        parsed.push({
            itemId,
            item,
            description,
            parentItemId
        });
    }

    if (parsed.length === 0) {
        throw new Error('No data rows were found under the required headers.');
    }

    if (errors.length) {
        const first = errors.slice(0, 5).join(' ');
        const more = errors.length > 5 ? ` (and ${errors.length - 5} more)` : '';
        throw new Error(`Some rows are missing required fields. ${first}${more}`);
    }

    return parsed;
}

/**
 * Import an Excel Project Miscellaneous List into Supabase.
 *
 * Tables written:
 * - lessons_learned_metadata_list (one row per item; metadata = Item; metadata_type = "miscellaneous")
 * - project_miscellaneous_excel (one row per item; links back to lessons_learned_metadata_list_id)
 *
 * @param {object} params
 * @param {any} params.supabase - Supabase client (from supabase-js)
 * @param {File} params.file - Uploaded Excel file (.xlsx or .xls)
 * @param {object} params.context - Required context for inserts
 * @param {number} params.context.organization_id
 * @param {number} params.context.created_by
 * @param {number|string} params.context.project_id
 * @param {number|string} params.context.project_type_id
 * @param {(msg: string) => void} [params.onProgress]
 * @param {number} [params.chunkSize]
 * @returns {Promise<object>} Summary counts
 */
export async function importProjectMiscListExcelToSupabase({
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
    const project_id = context.project_id;
    const project_type_id = context.project_type_id;

    if (organization_id == null || created_by == null || project_id == null || project_type_id == null) {
        throw new Error('Missing required context: organization_id, created_by, project_id, project_type_id.');
    }

    if (typeof onProgress === 'function') {
        onProgress('Reading Excel file...');
    }

    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });

    const firstSheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[firstSheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });

    const parsedItems = parseProjectMiscSheetRows(rows);

    // 1) Insert lessons_learned_metadata_list rows (one per item)
    const metadataRows = parsedItems.map(it => ({
        metadata_source: 'excel project miscellaneous list',
        metadata: it.item,
        metadata_type: 'miscellaneous',
        created_by,
        organization_id,
        project_id,
        project_type_id
    }));

    const insertedMetadata = await insertChunked({
        supabase,
        table: 'lessons_learned_metadata_list',
        rows: metadataRows,
        select: 'id',
        chunkSize,
        onProgress,
        label: 'Saving miscellaneous metadata'
    });

    if (insertedMetadata.length !== metadataRows.length) {
        throw new Error('Unexpected mismatch inserting lessons_learned_metadata_list rows for miscellaneous list.');
    }

    const lessonsIds = insertedMetadata.map(r => r.id);

    // 2) Insert project_miscellaneous_excel rows (one per item; linked by lessons id)
    const miscRows = parsedItems.map((it, idx) => ({
        lessons_learned_metadata_list_id: lessonsIds[idx],
        organization_id,
        project_id,
        project_type_id,
        created_by,
        name: it.item,
        description: it.description,
        item_id: it.itemId,
        parent_id: it.parentItemId
    }));

    const insertedMisc = await insertChunked({
        supabase,
        table: 'project_miscellaneous_excel',
        rows: miscRows,
        select: 'id',
        chunkSize,
        onProgress,
        label: 'Saving miscellaneous rows'
    });

    if (insertedMisc.length !== miscRows.length) {
        throw new Error('Unexpected mismatch inserting project_miscellaneous_excel rows.');
    }

    return {
        inserted: {
            lessons_learned_metadata_list: insertedMetadata.length,
            project_miscellaneous_excel: insertedMisc.length
        }
    };
}


