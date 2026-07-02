-- ============================================================================
-- Migration: add dedicated `source_tag` column to `sales_data`
-- ============================================================================
-- WHY: the existing `origin` column is already used for two DIFFERENT things:
--   1. On Sales rows: the real business field 출하지 (shipping/factory origin,
--      e.g. "Vietnam" / "Korea") — one real value per row.
--   2. On Manpower / TargetActual rows: a record-type tag ('Manpower' /
--      'TargetActual') used to partition the shared table.
-- Reusing `origin` to also tag Sales rows (origin: 'Sales') would silently
-- overwrite the real per-row shipping-origin value. So we add a separate
-- column purely for record-type partitioning, and leave `origin` alone.
-- ============================================================================

alter table sales_data
  add column if not exists source_tag text;

-- Backfill source_tag for rows that were already tagged the old way
-- (Manpower / TargetActual only — Sales rows never used this convention,
-- so nothing to backfill for them; new Sales inserts will set source_tag
-- going forward).
update sales_data
  set source_tag = origin
  where origin in ('Manpower', 'TargetActual')
    and source_tag is null;

-- Optional but recommended: index for the delete/filter queries each
-- dashboard runs on save.
create index if not exists idx_sales_data_source_tag on sales_data (source_tag);
