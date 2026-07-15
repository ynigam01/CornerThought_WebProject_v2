import type { SupabaseClient } from '@supabase/supabase-js';
interface ApiRequest {
    body?: any;
    params?: Record<string, string>;
}
interface ApiResponse {
    status(code: number): ApiResponse;
    json(payload: unknown): ApiResponse;
    setHeader?(name: string, value: string): void;
    write?(chunk: string): boolean;
    end?(chunk?: string): void;
    flushHeaders?(): void;
    headersSent?: boolean;
}
type Handler = (req: ApiRequest, res: ApiResponse) => void | Promise<void>;
interface RouteApp {
    post(path: string, handler: Handler): void;
    patch(path: string, handler: Handler): void;
    delete(path: string, handler: Handler): void;
}
export declare function registerDraftLessonRoutes(app: RouteApp, supabase: SupabaseClient): void;
export {};
//# sourceMappingURL=draftRoutes.d.ts.map