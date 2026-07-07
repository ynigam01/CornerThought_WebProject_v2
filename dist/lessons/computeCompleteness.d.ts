export type CompletenessQuality = 'minimum' | 'okay' | 'good' | 'great';
export interface LessonStructure {
    causeIds: (string | number)[];
    impactIds: (string | number)[];
    linkedItems: Array<{
        causeId: string | number | null;
        impactId: string | number | null;
    }>;
}
export declare function computeCompleteness(s: LessonStructure): CompletenessQuality | null;
//# sourceMappingURL=computeCompleteness.d.ts.map