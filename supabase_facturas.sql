-- Ejecutar en Supabase > SQL Editor (se puede repetir sin problema)
-- Adjuntos de facturas: columna en envios + permisos sobre el bucket "Facturas".
-- ATENCIÓN: sin login, acceso abierto. Al activar el login, cambiar "to anon, authenticated" por "to authenticated".
alter table envios add column if not exists factura_archivo text;
alter table envios add column if not exists factura_fecha date;
alter table envios add column if not exists factura_ruc text;
alter table envios add column if not exists factura_guia text;
alter table envios add column if not exists factura_destino text;
alter table envios add column if not exists factura_bultos int;
alter table envios add column if not exists factura_peso numeric;

drop policy if exists "facturas ver" on storage.objects;
drop policy if exists "facturas subir" on storage.objects;
drop policy if exists "facturas cambiar" on storage.objects;
drop policy if exists "facturas borrar" on storage.objects;

create policy "facturas ver" on storage.objects for select to anon, authenticated using (bucket_id = 'Facturas');
create policy "facturas subir" on storage.objects for insert to anon, authenticated with check (bucket_id = 'Facturas');
create policy "facturas cambiar" on storage.objects for update to anon, authenticated using (bucket_id = 'Facturas') with check (bucket_id = 'Facturas');
create policy "facturas borrar" on storage.objects for delete to anon, authenticated using (bucket_id = 'Facturas');
