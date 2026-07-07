import type { SupabaseClient } from '@supabase/supabase-js';
import type { Id } from './types';
import { type CompletenessQuality } from './computeCompleteness';
export declare function updateCompleteness(supabase: SupabaseClient, lessonId: Id, organizationId: Id): Promise<CompletenessQuality | null>;
//# sourceMappingURL=updateCompleteness.d.ts.map