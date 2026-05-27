-- Enable Row Level Security on all public Supabase tables.
--
-- This is safe for this project because the Flask app connects server-side via
-- DATABASE_URL. The frontend does not use a Supabase anon client and does not
-- talk directly to Supabase.
--
-- No policies are needed as long as only the server accesses the database.

DO $$
DECLARE r RECORD;
BEGIN
    FOR r IN
        SELECT tablename
        FROM pg_tables
        WHERE schemaname = 'public'
    LOOP
        EXECUTE format(
            'ALTER TABLE public.%I ENABLE ROW LEVEL SECURITY',
            r.tablename
        );
    END LOOP;
END $$;
