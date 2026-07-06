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
export interface LessonEntryInput {
    title: string;
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
    lessonIds: Array<string | number>;
}
//# sourceMappingURL=types.d.ts.map