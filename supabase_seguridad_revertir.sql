-- Ejecutar en Supabase > SQL Editor SOLO si hace falta volver al acceso abierto (sin sesión) después de supabase_seguridad.sql.
-- Restablece las políticas originales: cualquier persona con la dirección de la base de datos podrá leer y escribir.

do $$
declare
  t text;
  p record;
begin
  foreach t in array array['agencias', 'lista_agencias', 'envios', 'metas', 'hrn_detalle', 'liquidaciones', 'parametros', 'recojo_detalle'] loop
    if to_regclass('public.' || t) is not null then
      for p in select policyname from pg_policies where schemaname = 'public' and tablename = t loop
        execute format('drop policy if exists %I on public.%I', p.policyname, t);
      end loop;
      execute format('create policy %I on public.%I for all to anon, authenticated using (true) with check (true)', 'auth ' || t, t);
      execute format('grant all on public.%I to anon', t);
    end if;
  end loop;
end $$;

do $$
begin
  if to_regclass('public.hrn_grupos') is not null then
    grant select on public.hrn_grupos to anon, authenticated;
  end if;
end $$;

drop policy if exists "facturas ver" on storage.objects;
drop policy if exists "facturas subir" on storage.objects;
drop policy if exists "facturas cambiar" on storage.objects;
drop policy if exists "facturas borrar" on storage.objects;
create policy "facturas ver" on storage.objects for select to anon, authenticated using (bucket_id = 'Facturas');
create policy "facturas subir" on storage.objects for insert to anon, authenticated with check (bucket_id = 'Facturas');
create policy "facturas cambiar" on storage.objects for update to anon, authenticated using (bucket_id = 'Facturas') with check (bucket_id = 'Facturas');
create policy "facturas borrar" on storage.objects for delete to anon, authenticated using (bucket_id = 'Facturas');
