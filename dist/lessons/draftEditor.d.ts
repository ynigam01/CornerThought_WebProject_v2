import type { SupabaseClient } from '@supabase/supabase-js';
import type { AppliedMetadataRow, ApplyMetadataRequest, AssignmentRequest, CreateAttachmentRequest, CreateSubItemRequest, DeleteAttachmentRequest, DeleteMetadataRequest, DeleteSubItemRequest, UpdateSubItemRequest, UpdateTitleRequest } from './types';
type OkResult = {
    ok: true;
};
export declare function createSubItem(supabase: SupabaseClient, req: CreateSubItemRequest): Promise<OkResult>;
export declare function updateSubItem(supabase: SupabaseClient, req: UpdateSubItemRequest): Promise<OkResult>;
export declare function deleteSubItem(supabase: SupabaseClient, req: DeleteSubItemRequest): Promise<OkResult>;
export declare function reassignItem(supabase: SupabaseClient, req: AssignmentRequest): Promise<OkResult>;
export declare function updateTitle(supabase: SupabaseClient, req: UpdateTitleRequest): Promise<OkResult>;
export declare function createAttachment(supabase: SupabaseClient, req: CreateAttachmentRequest): Promise<OkResult>;
export declare function deleteAttachment(supabase: SupabaseClient, req: DeleteAttachmentRequest): Promise<OkResult>;
export declare function applyMetadata(supabase: SupabaseClient, req: ApplyMetadataRequest): Promise<{
    rows: AppliedMetadataRow[];
}>;
export declare function deleteMetadata(supabase: SupabaseClient, req: DeleteMetadataRequest): Promise<OkResult>;
export {};
//# sourceMappingURL=draftEditor.d.ts.map