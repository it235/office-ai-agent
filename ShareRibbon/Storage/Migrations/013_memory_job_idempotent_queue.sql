-- 队列幂等：memory_job 曾允许同一 job_type 加 target_id 重复入队，
-- 同一批会话事件因此可能被抽取多次。先清掉历史重复的在队任务，只保留最早一条，
-- 再用部分唯一索引把同一目标只允许一个未终结任务固化到数据库层。
DELETE FROM memory_job WHERE job_id IN (
  SELECT j.job_id FROM memory_job j
  WHERE j.status IN ('pending', 'processing') AND IFNULL(j.target_id, '') <> ''
    AND EXISTS (
      SELECT 1 FROM memory_job k
      WHERE k.job_type = j.job_type AND k.target_id = j.target_id
        AND k.status IN ('pending', 'processing')
        AND (k.created_at < j.created_at OR (k.created_at = j.created_at AND k.job_id < j.job_id))
    ));

CREATE UNIQUE INDEX IF NOT EXISTS idx_memory_job_active_unique
    ON memory_job(job_type, target_id)
    WHERE status IN ('pending', 'processing');

UPDATE schema_version SET version = 13;
