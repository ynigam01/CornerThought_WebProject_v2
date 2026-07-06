import type { SupabaseClient } from '@supabase/supabase-js';
import type {
  AttachmentInput,
  LessonEntryInput,
  SaveLessonsRequest,
  SaveLessonsResult,
} from './types';

const MAX_ATTACHMENT_FILE_BYTES = 10 * 1024 * 1024; // 10 MB per file

// Convert a decoded file buffer into the Postgres bytea hex format (\x...),
// matching the original browser-side arrayBufferToPgBytea helper.
function bufferToPgBytea(buffer: Buffer): string {
  return `\\x${buffer.toString('hex')}`;
}

function normalizeCategory(category: string): 'success' | 'issue' {
  return String(category || '').toLowerCase() === 'success' ? 'success' : 'issue';
}

function splitMetadataLabel(label: string): { metadataType: string | null; metadata: string } {
  if (label.includes(':')) {
    return {
      metadataType: label.split(':')[0].trim(),
      metadata: label.split(':').slice(1).join(':').trim(),
    };
  }
  return { metadataType: null, metadata: label };
}

// Persists one Add Data submission (one or more lesson entries) to Supabase.
// Ported from saveLessonsLearned() in frontend/public/js/user-portal.js.
export async function saveLessons(
  supabase: SupabaseClient,
  req: SaveLessonsRequest,
): Promise<SaveLessonsResult> {
  const { userId, organizationId, projectId, projectTypeId, review, entries } = req;

  if (userId == null || organizationId == null) {
    throw new Error('Missing user information.');
  }
  if (projectId == null) {
    throw new Error('Select a project before saving.');
  }
  if (!Array.isArray(entries) || entries.length === 0) {
    throw new Error('No entries to save.');
  }

  const reviewForDb = review === 'draft' ? 'draft' : 'for review';
  let savedCount = 0;
  const lessonIds: Array<string | number> = [];

  for (const entry of entries) {
    const title = entry.title || '';
    const category = normalizeCategory(entry.category);

    const { data: lessonRows, error: lessonErr } = await supabase
      .from('lessons_learned')
      .insert({
        title,
        category,
        review: reviewForDb,
        share: '',
        created_by: userId,
        organization_id: organizationId,
        project_id: projectId,
        project_type_id: projectTypeId,
      })
      .select('id')
      .single();

    if (lessonErr || !lessonRows) {
      throw new Error(lessonErr?.message || 'Failed to save lesson.');
    }

    const lessonId = lessonRows.id;
    savedCount += 1;
    lessonIds.push(lessonId);

    const causes = Array.isArray(entry.causes) ? entry.causes : [];
    if (causes.length) {
      const { error } = await supabase.from('lessons_learned_causes').insert(
        causes.map((cause) => ({
          lessons_learned_id: lessonId,
          cause,
          created_by: userId,
          organization_id: organizationId,
          project_id: projectId,
        })),
      );
      if (error) throw new Error(error.message || 'Failed to save causes.');
    }

    const impacts = Array.isArray(entry.impacts) ? entry.impacts : [];
    if (impacts.length) {
      const { error } = await supabase.from('lessons_learned_impacts').insert(
        impacts.map((impact) => ({
          lessons_learned_id: lessonId,
          impact,
          created_by: userId,
          organization_id: organizationId,
          project_id: projectId,
        })),
      );
      if (error) throw new Error(error.message || 'Failed to save impacts.');
    }

    const actions = Array.isArray(entry.actions) ? entry.actions : [];
    if (actions.length) {
      const { error } = await supabase.from('action_items').insert(
        actions.map((action_item) => ({
          lessons_learned_id: lessonId,
          action_item,
          lessons_learned_impact_id: null,
          lessons_learned_cause_id: null,
          created_by: userId,
          organization_id: organizationId,
          project_id: projectId,
        })),
      );
      if (error) throw new Error(error.message || 'Failed to save action items.');
    }

    const lessons = Array.isArray(entry.lessons) ? entry.lessons : [];
    if (lessons.length) {
      const { error } = await supabase.from('future_project_considerations').insert(
        lessons.map((fpc) => ({
          lessons_learned_id: lessonId,
          fpc,
          lessons_learned_impact_id: null,
          lessons_learned_cause_id: null,
          created_by: userId,
          organization_id: organizationId,
          project_id: projectId,
        })),
      );
      if (error) throw new Error(error.message || 'Failed to save lessons learned items.');
    }

    const notes = Array.isArray(entry.notes) ? entry.notes : [];
    if (notes.length) {
      const { error } = await supabase.from('lessons_learned_notes').insert(
        notes.map((notesText) => ({
          lessons_learned_id: lessonId,
          notes: notesText,
          created_by: userId,
          organization_id: organizationId,
          project_id: projectId,
        })),
      );
      if (error) throw new Error(error.message || 'Failed to save notes.');
    }

    const metadataItems = Array.isArray(entry.metadataItems) ? entry.metadataItems : [];
    if (metadataItems.length) {
      const { error } = await supabase.from('lessons_learned_metadata').insert(
        metadataItems.map((item) => {
          const { metadataType, metadata } = splitMetadataLabel(item.label);
          return {
            lessons_learned_id: lessonId,
            metadata_type: metadataType,
            metadata,
            lessons_learned_metadata_list_id: item.id,
            created_by: userId,
            organization_id: organizationId,
            project_id: projectId,
            project_type_id: projectTypeId,
          };
        }),
      );
      if (error) throw new Error(error.message || 'Failed to save metadata.');
    }

    const attachments = Array.isArray(entry.attachments) ? entry.attachments : [];
    if (attachments.length) {
      for (const file of attachments) {
        const savedAttachmentError = await saveAttachment(supabase, {
          file,
          lessonId,
          projectId,
          organizationId,
          userId,
        });
        if (savedAttachmentError) {
          throw new Error(savedAttachmentError);
        }
      }
    }
  }

  return { savedCount, lessonIds };
}

interface SaveAttachmentArgs {
  file: AttachmentInput;
  lessonId: string | number;
  projectId: string | number;
  organizationId: string | number;
  userId: string | number;
}

async function saveAttachment(
  supabase: SupabaseClient,
  { file, lessonId, projectId, organizationId, userId }: SaveAttachmentArgs,
): Promise<string | null> {
  const fileName = file && file.fileName ? file.fileName : 'attachment';
  const base64 = file && file.base64 ? file.base64 : '';
  if (!base64) {
    return null;
  }

  const fileBuffer = Buffer.from(base64, 'base64');
  const fileSize = fileBuffer.length;
  if (!fileSize) {
    return null;
  }
  if (fileSize > MAX_ATTACHMENT_FILE_BYTES) {
    return `Attachment "${fileName}" exceeds the 10 MB limit.`;
  }

  const { error } = await supabase.from('lessons_learned_attachments').insert({
    lessons_learned_id: lessonId,
    project_id: projectId,
    organization_id: organizationId,
    created_by: userId,
    file_data: bufferToPgBytea(fileBuffer),
    file_name: fileName,
    content_type: file.contentType || 'application/octet-stream',
  });

  if (error) {
    return error.message || `Failed to save attachment "${fileName}".`;
  }
  return null;
}
