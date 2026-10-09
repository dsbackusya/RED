-- Ejecutar en Supabase > SQL Editor, UNA SOLA VEZ y solo despues de comprobar que el ingreso con usuario y contraseña funciona.
-- Cierra el acceso a los datos: solo pueden leer y escribir las personas con sesión iniciada.
-- Si algo falla, ejecutar supabase_seguridad_revertir.sql para volver al acceso abierto.

-- 1) Tablas: se eliminan las políticas abiertas y se crea una sola política para usuarios autenticados.
do $$
declare
  t text;
  p record;
begin
  foreach t in array array['agencias', 'lista_agencias', 'envios', 'metas', 'hrn_detalle', 'liquidaciones', 'parametros', 'recojo_detalle'] loop
    if to_regclass('public.' || t) is not null then
      execute format('alter table public.%I enable row level security', t);
      for p in select policyname from pg_policies where schemaname = 'public' and tablename = t loop
        execute format('drop policy if exists %I on public.%I', p.policyname, t);
      end loop;
      execute format('create policy %I on public.%I for all to authenticated using (true) with check (true)', 'solo autenticados ' || t, t);
      execute format('revoke all on public.%I from anon', t);
    end if;
  end loop;
end $$;

-- 2) Vista del detalle agrupado (hrn_grupos): solo usuarios autenticados.
do $$
begin
  if to_regclass('public.hrn_grupos') is not null then
    revoke all on public.hrn_grupos from anon;
    grant select on public.hrn_grupos to authenticated;
  end if;
end $$;

-- 3) Archivos de facturas (bucket "Facturas"): solo usuarios autenticados.
drop policy if exists "facturas ver" on storage.objects;
drop policy if exists "facturas subir" on storage.objects;
drop policy if exists "facturas cambiar" on storage.objects;
drop policy if exists "facturas borrar" on storage.objects;
create policy "facturas ver" on storage.objects for select to authenticated using (bucket_id = 'Facturas');
create policy "facturas subir" on storage.objects for insert to authenticated with check (bucket_id = 'Facturas');
create policy "facturas cambiar" on storage.objects for update to authenticated using (bucket_id = 'Facturas') with check (bucket_id = 'Facturas');
create policy "facturas borrar" on storage.objects for delete to authenticated using (bucket_id = 'Facturas');

-- 4) Verificación: cada tabla debe mostrar únicamente el rol "authenticated".
select tablename, policyname, roles
from pg_policies
where schemaname = 'public' and tablename in ('agencias', 'lista_agencias', 'envios', 'metas', 'hrn_detalle', 'liquidaciones', 'parametros', 'recojo_detalle')
order by tablename;
