-- Hosts without a whole-stack native undo (Excel, PowerPoint) can still offer "undo this run"
-- when every remaining write carries a replayable snapshot of the target region before the change.
-- The snapshot is host-authored, capped by AgentRunChangeLog.MaxRestoreChars, and stays in this
-- local table only: never in logs, RunTrace, chat history or the panel payload.
ALTER TABLE agent_run_change ADD COLUMN restore_json TEXT;

UPDATE schema_version SET version = 18;
