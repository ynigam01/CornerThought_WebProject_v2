import type { SupabaseClient } from '@supabase/supabase-js';
import type {
  AppliedMetadataRow,
  ApplyMetadataRequest,
  AssignmentRequest,
  CreateAttachmentRequest,
  CreateSubItemRequest,
  DeleteAttachmentRequest,
  DeleteMetadataRequest,
  DeleteSubItemRequest,
  Id,
  SubItemKind,
  UpdateSubItemRequest,
  UpdateTitleRequest,
} from './types';
import { applyOwnerConstraint } from './ownership';

const MAX_ATTACHMENT_BYTES = 10 * 1024 * 1024; // 10 MB per file

function bufferToPgBytea(buffer: Buffer): string {
  return `\\x${buffer.toString('hex')}`;
}

// Maps a sub-item kind to its backing table and text column.
const SUB_ITEM_TABLE: Record<SubItemKind, { table: string; column: string }> = {
  cause: { table: 'lessons_learned_causes', column: 'cause' },
  impact: { table: 'lessons_learned_impacts', column: 'impact' },
  action: { table: 'action_items', column: 'action_item' },
  fpc: { table: 'future_project_considerations', column: 'fpc' },
  note: { table: 'lessons_learned_notes', column: 'notes' },
};

function resolveSubItem(kind: SubItemKind): { table: string; column: string } {
  const entry = SUB_ITEM_TABLE[kind];
  if (!entry) throw new Error(`Unknown sub-item kind: ${kind}`);
  return entry;
}

type OkResult = { ok: true };

async function unassignCauseDependencies(
  supabase: SupabaseClient,
  orgId: Id,
  pid: Id,
  causeId: Id,
): Promise<void> {
  await supabase
    .from('action_items')
    .update({ lessons_learned_cause_id: null })
    .eq('organization_id', orgId)
    .eq('project_id', pid)
    .eq('lessons_learned_cause_id', causeId);
  await supabase
    .from('future_project_considerations')
    .update({ lessons_learned_cause_id: null })
    .eq('organization_id', orgId)
    .eq('project_id', pid)
    .eq('lessons_learned_cause_id', causeId);
}

async function unassignImpactDependencies(
  supabase: SupabaseClient,
  orgId: Id,
  pid: Id,
  impactId: Id,
): Promise<void> {
  await supabase
    .from('action_items')
    .update({ lessons_learned_impact_id: null })
    .eq('organization_id', orgId)
    .eq('project_id', pid)
    .eq('lessons_learned_impact_id', impactId);
  await supabase
    .from('future_project_considerations')
    .update({ lessons_learned_impact_id: null })
    .eq('organization_id', orgId)
    .eq('project_id', pid)
    .eq('lessons_learned_impact_id', impactId);
}

export async function createSubItem(
  supabase: SupabaseClient,
  req: CreateSubItemRequest,
): Promise<OkResult> {
  const { lessonId, kind, value, organizationId, projectId, userId } = req;
  if (!value) throw new Error('Value is required.');
  const { table, column } = resolveSubItem(kind);

  const row: Record<string, unknown> = {
    lessons_learned_id: lessonId,
    [column]: value,
    created_by: userId,
    organization_id: organizationId,
    project_id: projectId,
  };
  // Action items and future considerations start unassigned.
  if (kind === 'action' || kind === 'fpc') {
    row.lessons_learned_cause_id = null;
    row.lessons_learned_impact_id = null;
  }

  const { error } = await supabase.from(table).insert(row);
  if (error) throw new Error(error.message || 'Save failed.');
  return { ok: true };
}

export async function updateSubItem(
  supabase: SupabaseClient,
  req: UpdateSubItemRequest,
): Promise<OkResult> {
  const { kind, id, value, organizationId, userId, forReviewCollaborative, rowCreatedBy } = req;
  if (!value) throw new Error('Value is required.');
  const { table, column } = resolveSubItem(kind);

  let query = supabase
    .from(table)
    .update({ [column]: value })
    .eq('id', id)
    .eq('organization_id', organizationId);
  query = applyOwnerConstraint(query, { forReviewCollaborative, rowCreatedBy, userId });

  const { error } = await query;
  if (error) throw new Error(error.message || 'Update failed.');
  return { ok: true };
}

export async function deleteSubItem(
  supabase: SupabaseClient,
  req: DeleteSubItemRequest,
): Promise<OkResult> {
  const { kind, id, organizationId, projectId, userId, forReviewCollaborative, rowCreatedBy } = req;
  const { table } = resolveSubItem(kind);

  if (kind === 'cause') {
    await unassignCauseDependencies(supabase, organizationId, projectId, id);
  } else if (kind === 'impact') {
    await unassignImpactDependencies(supabase, organizationId, projectId, id);
  }

  let query = supabase.from(table).delete().eq('id', id).eq('organization_id', organizationId);
  query = applyOwnerConstraint(query, { forReviewCollaborative, rowCreatedBy, userId });

  const { error } = await query;
  if (error) throw new Error(error.message || 'Delete failed.');
  return { ok: true };
}

export async function reassignItem(
  supabase: SupabaseClient,
  req: AssignmentRequest,
): Promise<OkResult> {
  const { kind, id, causeId, impactId, organizationId, projectId, userId, forReviewCollaborative, rowCreatedBy } = req;
  const table = kind === 'action' ? 'action_items' : 'future_project_considerations';

  let query = supabase
    .from(table)
    .update({
      lessons_learned_cause_id: causeId,
      lessons_learned_impact_id: impactId,
    })
    .eq('id', id)
    .eq('organization_id', organizationId)
    .eq('project_id', projectId);
  query = applyOwnerConstraint(query, { forReviewCollaborative, rowCreatedBy, userId });

  const { error } = await query;
  if (error) throw new Error(error.message || 'Failed to save assignment.');
  return { ok: true };
}

export async function updateTitle(
  supabase: SupabaseClient,
  req: UpdateTitleRequest,
): Promise<OkResult> {
  const { lessonId, title, organizationId, userId, forReviewCollaborative } = req;
  if (!title) throw new Error('Title is required.');

  let query = supabase
    .from('lessons_learned')
    .update({ title })
    .eq('id', lessonId)
    .eq('organization_id', organizationId);
  if (forReviewCollaborative) {
    query = query.eq('created_by', userId);
  }

  const { error } = await query;
  if (error) throw new Error(error.message || 'Update failed.');
  return { ok: true };
}

export async function createAttachment(
  supabase: SupabaseClient,
  req: CreateAttachmentRequest,
): Promise<OkResult> {
  const { lessonId, organizationId, projectId, userId, fileName, contentType, base64 } = req;
  if (!base64) throw new Error('No file data provided.');

  const fileBuffer = Buffer.from(base64, 'base64');
  if (!fileBuffer.length) throw new Error('No file data provided.');
  if (fileBuffer.length > MAX_ATTACHMENT_BYTES) {
    throw new Error(`"${fileName || 'attachment'}" exceeds 10 MB.`);
  }

  const { error } = await supabase.from('lessons_learned_attachments').insert({
    lessons_learned_id: lessonId,
    project_id: projectId,
    organization_id: organizationId,
    created_by: userId,
    file_data: bufferToPgBytea(fileBuffer),
    file_name: fileName || 'attachment',
    content_type: contentType || 'application/octet-stream',
  });
  if (error) throw new Error(error.message || 'Attachment upload failed.');
  return { ok: true };
}

export async function deleteAttachment(
  supabase: SupabaseClient,
  req: DeleteAttachmentRequest,
): Promise<OkResult> {
  const { id, organizationId, userId, forReviewCollaborative, rowCreatedBy } = req;

  let query = supabase
    .from('lessons_learned_attachments')
    .delete()
    .eq('id', id)
    .eq('organization_id', organizationId);
  query = applyOwnerConstraint(query, { forReviewCollaborative, rowCreatedBy, userId });

  const { error } = await query;
  if (error) throw new Error(error.message || 'Remove failed.');
  return { ok: true };
}

export async function applyMetadata(
  supabase: SupabaseClient,
  req: ApplyMetadataRequest,
): Promise<{ rows: AppliedMetadataRow[] }> {
  const { lessonId, organizationId, projectId, userId, projectTypeId, listIds } = req;
  const ids = Array.isArray(listIds) ? listIds.filter((v) => v != null) : [];
  if (!ids.length) return { rows: [] };

  // Look up the metadata definitions so we can copy their type/value onto the link row.
  const { data: listRows, error: listErr } = await supabase
    .from('lessons_learned_metadata_list')
    .select('id, metadata_type, metadata')
    .eq('organization_id', organizationId)
    .eq('project_id', projectId)
    .in('id', ids);
  if (listErr) throw new Error(listErr.message || 'Failed to load metadata.');

  const byId = new Map<string, { id: Id; metadata_type: string | null; metadata: unknown }>();
  (listRows || []).forEach((r: any) => byId.set(String(r.id), r));

  const inserted: AppliedMetadataRow[] = [];
  for (const listId of ids) {
    const row = byId.get(String(listId));
    if (!row) continue;
    const { data: ins, error } = await supabase
      .from('lessons_learned_metadata')
      .insert({
        lessons_learned_id: lessonId,
        metadata_type: row.metadata_type || null,
        metadata: row.metadata,
        lessons_learned_metadata_list_id: row.id,
        created_by: userId,
        organization_id: organizationId,
        project_id: projectId,
        project_type_id: projectTypeId,
      })
      .select('id, metadata, metadata_type, lessons_learned_metadata_list_id')
      .single();
    if (error) throw new Error(error.message || 'Apply failed.');
    if (ins) inserted.push(ins as AppliedMetadataRow);
  }

  return { rows: inserted };
}

export async function deleteMetadata(
  supabase: SupabaseClient,
  req: DeleteMetadataRequest,
): Promise<OkResult> {
  const { id, organizationId, userId, forReviewCollaborative, rowCreatedBy } = req;

  let query = supabase
    .from('lessons_learned_metadata')
    .delete()
    .eq('id', id)
    .eq('organization_id', organizationId);
  query = applyOwnerConstraint(query, { forReviewCollaborative, rowCreatedBy, userId });

  const { error } = await query;
  if (error) throw new Error(error.message || 'Remove failed.');
  return { ok: true };
}
