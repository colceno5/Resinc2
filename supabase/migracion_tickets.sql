-- =====================================================================
--  MIGRACIÓN: permite guardar los tickets en Supabase
--  Ejecuta este archivo UNA vez en: Supabase → SQL Editor → New query → Run
--  (solo si ya habías ejecutado schema.sql antes)
-- =====================================================================
alter table public.app_datos drop constraint if exists colecciones_validas;
alter table public.app_datos
  add constraint colecciones_validas
  check (coleccion in ('planes','justificaciones','cortes','tickets'));
