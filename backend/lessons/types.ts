// Input types for the Add Data (lessons learned) save flow.
// These describe the JSON payload the browser sends to the backend save endpoints.

export type LessonReview = 'draft' | 'for review';

export interface MetadataItemInput {
  // Display label as shown in the UI, e.g. "Phase: Demolition" or a bare value.
  label: string;
  // The lessons_learned_metadata_list id this label was selected from (nullable).
  id: string | number | null;
}

export interface AttachmentInput {
  fileName: string;
  contentType: string;
  // Base64-encoded file contents (no data: URI prefix).
  base64: string;
}

export interface LessonEntryInput {
  title: string;
  // 'success' or 'issue'. Anything not 'success' is treated as 'issue'.
  category: string;
  causes: string[];
  impacts: string[];
  actions: string[];
  lessons: string[];
  notes: string[];
  metadataItems: MetadataItemInput[];
  attachments: AttachmentInput[];
}

export interface SaveLessonsRequest {
  userId: string | number;
  organizationId: string | number;
  projectId: string | number;
  projectTypeId: string | number | null;
  review: LessonReview;
  entries: LessonEntryInput[];
}

export interface SaveLessonsResult {
  savedCount: number;
  // Ids of the lessons_learned rows created, in submission order.
  // Used by the client to create review notifications after a submit.
  lessonIds: Array<string | number>;
}
