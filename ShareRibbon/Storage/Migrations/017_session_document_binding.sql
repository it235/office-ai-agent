-- Bind a chat session to the document it was created for, and let the user rename it.
-- session_summary used to be a legacy snapshot table that nothing read; it is now the
-- per-session profile row (one row per session) joined by the history list.
--
-- The unique index below requires one row per session. The dedupe therefore has to live in
-- this same migration: if a pre-existing duplicate makes the index fail, RunVersionedMigrations
-- rolls the whole transaction back, and the next startup gets another chance to converge
-- instead of being stuck on a half-applied version forever.

DELETE FROM session_summary
WHERE IFNULL(session_id, '') <> ''
  AND id NOT IN (
    SELECT MAX(s.id) FROM session_summary s WHERE s.session_id = session_summary.session_id
);

-- Legacy rows keyed by an empty session id cannot ever be joined, so they are not a profile.
DELETE FROM session_summary WHERE IFNULL(session_id, '') = '';

ALTER TABLE session_summary ADD COLUMN app_type TEXT NOT NULL DEFAULT '';
ALTER TABLE session_summary ADD COLUMN document_id TEXT NOT NULL DEFAULT '';
ALTER TABLE session_summary ADD COLUMN document_name TEXT NOT NULL DEFAULT '';
ALTER TABLE session_summary ADD COLUMN custom_title TEXT;
ALTER TABLE session_summary ADD COLUMN updated_at TEXT;

CREATE UNIQUE INDEX IF NOT EXISTS idx_session_summary_session_unique ON session_summary(session_id);
CREATE INDEX IF NOT EXISTS idx_session_summary_document ON session_summary(app_type, document_id);

UPDATE schema_version SET version = 17;
