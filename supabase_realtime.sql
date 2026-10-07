-- Ejecutar en Supabase > SQL Editor (se puede repetir sin problema)
-- Tiempo real: la app se entera sola de los cambios que hacen otras personas
-- (registros, liquidaciones y agencias), sin recargar la página.
-- ATENCIÓN: sin login, acceso abierto. Al activar el login, las políticas de cada tabla ya controlan quién recibe los cambios.

do $$
declare t text;
begin
  foreach t in array array['envios', 'liquidaciones', 'agencias'] loop
    if not exists (select 1 from pg_publication_tables where pubname = 'supabase_realtime' and schemaname = 'public' and tablename = t) then
      execute format('alter publication supabase_realtime add table public.%I', t);
    end if;
  end loop;
end $$;
