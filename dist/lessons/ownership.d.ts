import type { Id, OwnershipContext } from './types';
export declare function effectiveRowCreatedBy(createdBy: Id | null | undefined, lessonCreatorId: Id | null | undefined): string;
export declare function canMutateSubRow(createdBy: Id | null | undefined, lessonCreatorId: Id | null | undefined, userId: Id | null | undefined): boolean;
interface Filterable {
    eq(column: string, value: unknown): this;
}
export declare function applyOwnerConstraint<T extends Filterable>(query: T, ctx: Pick<OwnershipContext, 'forReviewCollaborative' | 'rowCreatedBy' | 'userId'>): T;
export {};
//# sourceMappingURL=ownership.d.ts.map