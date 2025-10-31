// frontend/public/js/supabase-client.js
import { createClient } from 'https://esm.sh/@supabase/supabase-js';

const SUPABASE_URL = 'https://tybkbxdzfywnoiqvcnzr.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InR5YmtieGR6Znl3bm9pcXZjbnpyIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjE4MzE3MDAsImV4cCI6MjA3NzQwNzcwMH0.RgEizFZyAGtbDyqLLQCryMZ0JuikQWjdGn7jkNK5RNM';

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);