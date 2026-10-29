-- =====================================================================
--  Rendimiento de atenciones · Aranda  —  base de datos (Supabase)
--  Pega y ejecuta TODO este archivo en: Supabase → SQL Editor → New query
-- =====================================================================

-- Una sola tabla guarda las tres colecciones de la aplicación:
--   planes           → registro de planes de acción
--   justificaciones  → sustentos de la auditoría de vencidos (por ticket)
--   cortes           → cortes semanales guardados del informe
create table if not exists public.app_datos (
  coleccion  text        not null,
  id         text        not null,
  data       jsonb       not null,
  updated_at timestamptz not null default now(),
  updated_by uuid        default auth.uid(),
  primary key (coleccion, id),
  constraint colecciones_validas check (coleccion in ('planes','justificaciones','cortes'))
);

-- Quién y cuándo modificó cada registro (lo llena la propia base de datos)
create or replace function public.app_datos_audit()
returns trigger language plpgsql as $$
begin
  new.updated_at := now();
  new.updated_by := auth.uid();
  return new;
end $$;

drop trigger if exists app_datos_audit on public.app_datos;
create trigger app_datos_audit
  before insert or update on public.app_datos
  for each row execute function public.app_datos_audit();

-- Seguridad: solo usuarios con sesión iniciada pueden leer y escribir
alter table public.app_datos enable row level security;

drop policy if exists "equipo lee"      on public.app_datos;
drop policy if exists "equipo inserta"  on public.app_datos;
drop policy if exists "equipo edita"    on public.app_datos;
drop policy if exists "equipo elimina"  on public.app_datos;

create policy "equipo lee"     on public.app_datos for select to authenticated using (true);
create policy "equipo inserta" on public.app_datos for insert to authenticated with check (true);
create policy "equipo edita"   on public.app_datos for update to authenticated using (true) with check (true);
create policy "equipo elimina" on public.app_datos for delete to authenticated using (true);

-- Permisos de la API: nada para visitantes anónimos, todo para usuarios con sesión
revoke all on public.app_datos from anon;
grant select, insert, update, delete on public.app_datos to authenticated;
