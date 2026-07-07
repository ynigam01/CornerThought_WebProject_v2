import type { SupabaseClient } from '@supabase/supabase-js';
import type { Id } from './types';
import { computeCompleteness } from './computeCompleteness';

// Queries a lesson's causes, impacts, and linked action items / future
// project considerations, computes the highest completeness tier, and
// writes it back to lessons_learned.completeness_quality.
export async function updateCompleteness(
  supabase: SupabaseClient,
  lessonId: Id,
  organizationId: Id,
): Promise<void> {
  const scope = {
    lessons_learned_id: lessonId,
    organization_id: organizationId,
  };

  const [causesRes, impactsRes, actionsRes, fpcsRes] = await Promise.all([
    supabase
      .from('lessons_learned_causes')
      .select('id')
      .eq('lessons_learned_id', scope.lessons_learned_id)
      .eq('organization_id', scope.organization_id),

    supabase
      .from('lessons_learned_impacts')
      .select('id')
      .eq('lessons_learned_id', scope.lessons_learned_id)
      .eq('organization_id', scope.organization_id),

    supabase
      .from('action_items')
      .select('lessons_learned_cause_id, lessons_learned_impact_id')
      .eq('lessons_learned_id', scope.lessons_learned_id)
      .eq('organization_id', scope.organization_id),

    supabase
      .from('future_project_considerations')
      .select('lessons_learned_cause_id, lessons_learned_impact_id')
      .eq('lessons_learned_id', scope.lessons_learned_id)
      .eq('organization_id', scope.organization_id),
  ]);

  const causeIds = (causesRes.data ?? []).map((r: any) => r.id);
  const impactIds = (impactsRes.data ?? []).map((r: any) => r.id);
  const linkedItems = [
    ...(actionsRes.data ?? []).map((r: any) => ({
      causeId: r.lessons_learned_cause_id ?? null,
      impactId: r.lessons_learned_impact_id ?? null,
    })),
    ...(fpcsRes.data ?? []).map((r: any) => ({
      causeId: r.lessons_learned_cause_id ?? null,
      impactId: r.lessons_learned_impact_id ?? null,
    })),
  ];

  const quality = computeCompleteness({ causeIds, impactIds, linkedItems });

  const { error } = await supabase
    .from('lessons_learned')
    .update({ completeness_quality: quality })
    .eq('id', lessonId)
    .eq('organization_id', organizationId);

  if (error) {
    console.error('Failed to update completeness_quality:', error.message);
  }
}
