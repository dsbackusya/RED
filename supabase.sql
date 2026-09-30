-- Ejecutar en Supabase > SQL Editor
-- ATENCIÓN: sin login. Las políticas permiten acceso a cualquiera que tenga la URL de la app.
--   Al activar el login, cambiar "to anon, authenticated" por "to authenticated".
-- Hoja "Agencias": mismas columnas que AEGENCIAS del Excel
create table if not exists agencias (
  destino text primary key,
  agencia text,
  transporte text,
  zona text,
  observacion text check (observacion in ('AL CONTADO','CREDITO')),
  concatenar text
);
create table if not exists lista_agencias (
  agencia text primary key
);
create table if not exists envios (
  id uuid primary key default gen_random_uuid(),
  fecha date not null,
  origen text not null default 'LIMA',
  transporte text,
  nodo text not null,
  bultos int,
  pedidos int,
  cajas int,
  placa text,
  agencia text,
  detalle_gasto text,
  modo_pago text check (modo_pago in ('CONTADO','CREDITO')),
  importe numeric,
  factura text,
  created_at timestamptz not null default now(),
  unique (fecha, nodo)
);
alter table agencias enable row level security;
alter table lista_agencias enable row level security;
alter table envios enable row level security;
drop policy if exists "auth agencias" on agencias;
create policy "auth agencias" on agencias for all to anon, authenticated using (true) with check (true);
drop policy if exists "auth lista_agencias" on lista_agencias;
create policy "auth lista_agencias" on lista_agencias for all to anon, authenticated using (true) with check (true);
drop policy if exists "auth envios" on envios;
create policy "auth envios" on envios for all to anon, authenticated using (true) with check (true);

insert into agencias (destino, agencia, transporte, zona, observacion, concatenar) values
('HUANCAVELICA', 'LOGISTICA ANTEZANA SAC', 'RELUFI', 'CENTRO', 'CREDITO', 'ENVIO DE PEDIDOS HUANCAVELICA'),
('AYACUCHO', 'LOGISTICA ANTEZANA SAC', 'RELUFI', 'SUR', 'CREDITO', 'ENVIO DE PEDIDOS  AYACUCHO'),
('HUANCAYO', 'LOGISTICA ANTEZANA SAC', 'RELUFI', 'CENTRO', 'CREDITO', 'ENVIO DE PEDIDOS HUANCAYO'),
('BARRANCA', 'COMERCIAL ELIAS A & G S.A.C', 'ELIAS', 'NORTE CHICO', 'CREDITO', 'ENVIO DE PEDIDOS  BARRANCA'),
('CHIMBOTE', 'COMERCIAL ELIAS A & G S.A.C', 'ELIAS', 'NORTE CHICO', 'CREDITO', 'ENVIO DE PEDIDOS CHIMBOTE'),
('HUACHO', 'COMERCIAL ELIAS A & G S.A.C', 'ELIAS', 'NORTE CHICO', 'CREDITO', 'ENVIO DE PEDIDOS HUACHO'),
('HUARAL', 'COMERCIAL ELIAS A & G S.A.C', 'ELIAS', 'NORTE CHICO', 'CREDITO', 'ENVIO DE PEDIDOS HUARAL'),
('PISCO', 'COMERCIAL ELIAS A & G S.A.C', 'ELIAS', 'SUR', 'CREDITO', 'ENVIO DE PEDIDOS PISCO'),
('ICA', 'COMERCIAL ELIAS A & G S.A.C', 'ELIAS', 'SUR', 'CREDITO', 'ENVIO DE PEDIDOS ICA'),
('NAZCA', 'COMERCIAL ELIAS A & G S.A.C', 'ELIAS', 'SUR', 'CREDITO', 'ENVIO DE PEDIDOS NAZCA'),
('CAÑETE', 'COMERCIAL ELIAS A & G S.A.C', 'ELIAS', 'SUR', 'CREDITO', 'ENVIO DE PEDIDOS CAÑETE'),
('CHINCHA', 'COMERCIAL ELIAS A & G S.A.C', 'ELIAS', 'SUR', 'CREDITO', 'ENVIO DE PEDIDOS CHINCHA'),
('HUARAZ', 'CAMIBEL LOGISTIC SAC', 'RELUFI', 'OTROS', 'AL CONTADO', 'ENVIO DE PEDIDOS HUARAZ'),
('CUSCO', 'JATSA CARGO', 'RELUFI', 'SUR', 'CREDITO', 'ENVIO DE PEDIDOS CUSCO'),
('ILO', 'JATSA CARGO', 'RELUFI', 'SUR', 'CREDITO', 'ENVIO DE PEDIDOS ILO'),
('MOQUEGUA', 'JATSA CARGO', 'RELUFI', 'SUR', 'CREDITO', 'ENVIO DE PEDIDOS MOQUEGUA'),
('TACNA', 'JATSA CARGO', 'RELUFI', 'SUR', 'CREDITO', 'ENVIO DE PEDIDOS TACNA'),
('PIURA', 'J & J ALIADOS EIRL', 'RELUFI', 'NORTE', 'AL CONTADO', 'ENVIO DE PEDIDOS PIURA'),
('CHICLAYO', 'LOGISTIC EMPRESARIAL SAC', 'RELUFI', 'NORTE', 'AL CONTADO', 'ENVIO DE PEDIDOS CHICLAYO'),
('TRUJILLO', 'LOGISTIC EMPRESARIAL SAC', 'ELIAS', 'NORTE', 'AL CONTADO', 'ENVIO DE PEDIDOS TRUJILLO'),
('TALARA', 'MARYFER CARGO SAC', 'RELUFI', 'NORTE', 'AL CONTADO', 'ENVIO DE PEDIDOS TALARA'),
('PUNO', 'ROMELISA', 'RELUFI', 'CENTRO', 'AL CONTADO', 'ENVIO DE PEDIDOS PUNO'),
('JULIACA', 'ROMELISA', 'RELUFI', 'CENTRO', 'AL CONTADO', 'ENVIO DE PEDIDOS JULIACA'),
('PUCALLPA', 'SAN PEDRO CARGO SAC', 'RELUFI', 'ESTE', 'AL CONTADO', 'ENVIO DE PEDIDOS PUCALLPA'),
('HUANUCO', 'SAN PEDRO CARGO SAC', 'RELUFI', 'SUR', 'AL CONTADO', 'ENVIO DE PEDIDOS HUANUCO'),
('TINGO MARIA', 'SAN PEDRO CARGO SAC', 'RELUFI', 'ESTE', 'AL CONTADO', 'ENVIO DE PEDIDOS TINGO MARIA'),
('TARMA', 'SAN PEDRO CARGO SAC', 'RELUFI', 'CENTRO', 'AL CONTADO', 'ENVIO DE PEDIDOS TARMA'),
('LA MERCED', 'SAN PEDRO CARGO SAC', 'RELUFI', 'CENTRO', 'AL CONTADO', 'ENVIO DE PEDIDOS LA MERCED'),
('PASCO', 'SAN PEDRO CARGO SAC', 'RELUFI', 'CENTRO', 'AL CONTADO', 'ENVIO DE PEDIDOS PASCO'),
('PUERTO MALDONADO', 'JATSA CARGO', 'RELUFI', 'CENTRO', 'CREDITO', 'ENVIO DE PEDIDOS PUERTO MALDONADO'),
('ANDAHUAYLAS', 'TURISMO INTERNACIONAL PALOMINO SAC', 'RELUFI', 'SUR', 'AL CONTADO', 'ENVIO DE PEDIDOS  ANDAHUAYLAS'),
('ABANCAY', 'TURISMO INTERNACIONAL PALOMINO SAC', 'RELUFI', 'SUR', 'AL CONTADO', 'ENVIO DE PEDIDOS  ABANCAY'),
('AREQUIPA TERRESTRE', 'JATSA CARGO', 'RELUFI', 'SUR', 'CREDITO', 'ENVIO DE PEDIDOS AREQUIPA TERRESTRE'),
('AREQUIPA AEREO', 'KARSIL', 'ZANGHO', 'SUR', 'CREDITO', 'ENVIO DE PEDIDOS  AREQUIPA  AEREO'),
('TARAPOTO', 'SAN MARTIN CARGA EMPRESARIAL SAC', 'FEROZ', 'ESTE', 'CREDITO', 'ENVIO DE PEDIDOS TARAPOTO'),
('AREQUIPA', 'JATSA CARGO', 'RELUFI', 'SUR', 'CREDITO', 'ENVIO DE PEDIDOS AREQUIPA')
on conflict (destino) do nothing;

insert into lista_agencias (agencia) values
('CAMIBEL LOGISTIC SAC'),
('COMERCIAL ELIAS A & G S.A.C'),
('J & J ALIADOS EIRL'),
('JATSA CARGO'),
('LOGISTIC EMPRESARIAL SAC'),
('LOGISTICA ANTEZANA SAC'),
('MARYFER CARGO SAC'),
('MILENIUM'),
('ROMELISA'),
('SAN MARTIN CARGA EMPRESARIAL SAC'),
('SAN PEDRO CARGO SAC'),
('TURISMO INTERNACIONAL PALOMINO SAC')
on conflict (agencia) do nothing;
