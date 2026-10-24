-- Track what actually remains in the document after a run so "undo this run" can be
-- gated on host-provided undo units and verified with before/after document fingerprints.
ALTER TABLE agent_run_change ADD COLUMN rolled_back INTEGER NOT NULL DEFAULT 0;
ALTER TABLE agent_run_change ADD COLUMN undo_units INTEGER NOT NULL DEFAULT 0;
ALTER TABLE agent_run_change ADD COLUMN doc_hash_before TEXT;
ALTER TABLE agent_run_change ADD COLUMN doc_hash_after TEXT;

UPDATE schema_version SET version = 16;
