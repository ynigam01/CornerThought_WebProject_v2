export type LessonReview = 'draft' | 'for review';
export interface MetadataItemInput {
    label: string;
    id: string | number | null;
}
export interface AttachmentInput {
    fileName: string;
    contentType: string;
    base64: string;
}
export interface ActionItemInput {
    text: string;
    status: 'completed' | 'recommended';
}
export interface CauseImpactItemInput {
    text: string;
    actions: ActionItemInput[];
    lessons: string[];
}
export interface LessonEntryInput {
    highLevelTitle?: string;
    title: string;
    category: string;
    causes: Array<CauseImpactItemInput | string>;
    impacts: Array<CauseImpactItemInput | string>;
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
    lessonIds: Array<string | number>;
}
export type Id = string | number;
export type SubItemKind = 'cause' | 'impact' | 'action' | 'fpc' | 'note';
export type AssignmentKind = 'action' | 'fpc';
export interface OwnershipContext {
    userId: Id;
    organizationId: Id;
    forReviewCollaborative?: boolean;
    lessonCreatorId?: Id | null;
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
export interface AppliedMetadataRow {
    id: Id;
    metadata: unknown;
    metadata_type: string | null;
    lessons_learned_metadata_list_id: Id | null;
}
//# sourceMappingURL=types.d.ts.map