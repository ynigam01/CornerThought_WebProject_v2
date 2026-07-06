import type { Id, OwnershipContext } from './types';

// Ported from my-projects-lesson-draft-editor.js so the backend can enforce the
// exact same ownership rules the client uses to gate the editor UI.

export function effectiveRowCreatedBy(createdBy: Id | null | undefined, lessonCreatorId: Id | null | undefined): string {
  if (createdBy != null && String(createdBy).trim() !== '') return String(createdBy);
  if (lessonCreatorId != null) return String(lessonCreatorId);
  return '';
}

export function canMutateSubRow(
  createdBy: Id | null | undefined,
  lessonCreatorId: Id | null | undefined,
  userId: Id | null | undefined,
): boolean {
  if (userId == null) return false;
  return effectiveRowCreatedBy(createdBy, lessonCreatorId) === String(userId);
}

// Minimal shape of a Supabase/PostgREST filter builder for chaining .eq().
interface Filterable {
  eq(column: string, value: unknown): this;
}

// Mirrors applyReviewOwnerConstraint from the editor: in collaborative "for review"
// mode, scope a mutation to rows the requesting user created.
export function applyOwnerConstraint<T extends Filterable>(
  query: T,
  ctx: Pick<OwnershipContext, 'forReviewCollaborative' | 'rowCreatedBy' | 'userId'>,
): T {
  if (ctx.forReviewCollaborative && ctx.rowCreatedBy != null) {
    return query.eq('created_by', ctx.userId);
  }
  return query;
}
