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

// ---------------------------------------------------------------------------
// Draft lesson editor types
// ---------------------------------------------------------------------------

export type Id = string | number;

// Sub-item kinds map to their backing tables in the draft editor.
export type SubItemKind = 'cause' | 'impact' | 'action' | 'fpc' | 'note';

// Only action items and future considerations can be assigned to a cause/impact.
export type AssignmentKind = 'action' | 'fpc';

// Collaborative-review ownership context passed from the client on each write.
// The backend uses the service-role key (RLS bypassed), so it must re-enforce
// the same ownership rules the client used to gate the UI.
export interface OwnershipContext {
  userId: Id;
  organizationId: Id;
  forReviewCollaborative?: boolean;
  lessonCreatorId?: Id | null;
  // The target row's created_by, when known (for update/delete/assignment).
  rowCreatedBy?: Id | null;
}

export interface UpdateLessonReviewRequest {
  lessonId: Id;
  organizationId: Id;
  userId: Id;
  review: LessonReview;
  forReviewCollaborative?: boolean;
}

export interface CreateSubItemRequest {
  lessonId: Id;
  kind: SubItemKind;
  value: string;
  organizationId: Id;
  projectId: Id;
  userId: Id;
}

export interface UpdateSubItemRequest {
  kind: SubItemKind;
  id: Id;
  value: string;
  organizationId: Id;
  userId: Id;
  forReviewCollaborative?: boolean;
  rowCreatedBy?: Id | null;
}

export interface DeleteSubItemRequest {
  kind: SubItemKind;
  id: Id;
  organizationId: Id;
  projectId: Id;
  userId: Id;
  forReviewCollaborative?: boolean;
  rowCreatedBy?: Id | null;
}

export interface AssignmentRequest {
  kind: AssignmentKind;
  id: Id;
  causeId: Id | null;
  impactId: Id | null;
  organizationId: Id;
  projectId: Id;
  userId: Id;
  forReviewCollaborative?: boolean;
  rowCreatedBy?: Id | null;
}

export interface UpdateTitleRequest {
  lessonId: Id;
  title: string;
  organizationId: Id;
  userId: Id;
  forReviewCollaborative?: boolean;
}

export interface CreateAttachmentRequest {
  lessonId: Id;
  organizationId: Id;
  projectId: Id;
  userId: Id;
  fileName: string;
  contentType: string;
  // Base64-encoded file contents (no data: URI prefix).
  base64: string;
}

export interface DeleteAttachmentRequest {
  id: Id;
  organizationId: Id;
  userId: Id;
  forReviewCollaborative?: boolean;
  rowCreatedBy?: Id | null;
}

export interface ApplyMetadataRequest {
  lessonId: Id;
  organizationId: Id;
  projectId: Id;
  userId: Id;
  projectTypeId: Id | null;
  listIds: Id[];
}

export interface DeleteMetadataRequest {
  id: Id;
  organizationId: Id;
  userId: Id;
  forReviewCollaborative?: boolean;
  rowCreatedBy?: Id | null;
}

// A metadata row returned to the client after apply, so it can update its
// in-memory currentLinks list without a full refetch.
export interface AppliedMetadataRow {
  id: Id;
  metadata: unknown;
  metadata_type: string | null;
  lessons_learned_metadata_list_id: Id | null;
}
