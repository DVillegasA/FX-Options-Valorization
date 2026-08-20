from datetime import date
import openpyxl
import pytest
import daily_process as dp
from fx_options_valuation import liquidar_vencimiento


def op(folio, strike, lado, prima_usd=None, margen_clp=None, margen_usd=None,
       prima_clp=None, prima_clp_usd=None, prima_pct=None,
       inicio=date(2026, 4, 29), venc=date(2026, 5, 29), tipo="CALL",
       nominal=100000.0):
    return dict(folio=folio, inicio=inicio, venc=venc, tipo=tipo, lado=lado,
                strike=strike, nominal=nominal, contraparte="CP", modalidad="EUROPEA",
                entrega="Compensación", moneda="USD", par="USD/CLP",
                prima_clp=prima_clp, prima_clp_usd=prima_clp_usd,
                prima_pct=prima_pct, prima_usd=prima_usd,
                margen_clp=margen_clp, margen_usd=margen_usd, usdobs=900.0,
                itm_src=None)


# ---------- clasificacion de cartera ----------
def test_clasificar():
    D = date(2026, 5, 29)
    ops = [op("VIG", 900, "Venta", inicio=date(2026, 5, 1), venc=date(2026, 6, 30)),
           op("ALTA", 900, "Venta", inicio=D, venc=date(2026, 6, 30)),
           op("VENCE", 900, "Venta", inicio=date(2026, 4, 1), venc=D),
           op("VIEJA", 900, "Venta", inicio=date(2026, 1, 1), venc=date(2026, 2, 1))]
    vig, ven, altas = dp.clasificar(ops, D)
    assert {o["folio"] for o in vig} == {"VIG", "ALTA"}
    assert [o["folio"] for o in ven] == ["VENCE"]
    assert [o["folio"] for o in altas] == ["ALTA"]


# ---------- deteccion de prima % ----------
def test_es_prima_pct():
    assert dp._es_prima_pct(op("A", 900, "Venta", prima_usd=1500))
    assert not dp._es_prima_pct(op("B", 900, "Venta", prima_clp=1000000))


# ---------- fx por calce ----------
def test_fx_por_calce_1a1():
    ops = [op("V", 920, "Venta", prima_usd=1500, margen_clp=0, margen_usd=0),
           op("C", 920, "Compra", prima_usd=-601.6,
              margen_clp=804993.352, margen_usd=898.4)]
    fx = dp._fx_por_calce(ops)
    k = dp._clave_calce(ops[0])
    assert fx[k] == pytest.approx(896.03, abs=0.01)   # 804993.352/898.4

def test_fx_por_calce_grupos_independientes():
    ops = [op("V1", 920, "Venta", prima_usd=1500, margen_clp=0, margen_usd=0),
           op("C1", 920, "Compra", prima_usd=-601.6, margen_clp=804993.352, margen_usd=898.4),
           op("V2", 930, "Venta", prima_usd=700, margen_clp=0, margen_usd=0),
           op("C2", 930, "Compra", prima_usd=-21.59, margen_clp=607875.7123, margen_usd=678.41)]
    fx = dp._fx_por_calce(ops)
    assert len(fx) == 2                                # un fx por calce
    k1 = dp._clave_calce(ops[0]); k2 = dp._clave_calce(ops[2])
    assert fx[k1] == pytest.approx(804993.352 / 898.4)
    assert fx[k2] == pytest.approx(607875.7123 / 678.41)

def test_fx_por_calce_n_a_m():
    # 2 compras + 1 venta calzadas: agrega margenes y primas del grupo completo
    ops = [op("V", 920, "Venta", prima_usd=2000, margen_clp=0, margen_usd=0),
           op("C1", 920, "Compra", prima_usd=-500, margen_clp=450000, margen_usd=500),
           op("C2", 920, "Compra", prima_usd=-600, margen_clp=360000, margen_usd=400)]
    fx = dp._fx_por_calce(ops)
    k = dp._clave_calce(ops[0])
    assert fx[k] == pytest.approx((450000 + 360000) / (2000 - 500 - 600))


# ---------- P&L segun tipo de prima ----------
def test_pl_prima_clp():
    o = op("N", 900, "Venta", prima_clp=1000000)
    res = [{"mtm": -300000.0, "spot": 900.0}]
    dp._enriquecer_pl([o], res)
    assert res[0]["pl_clp"] == pytest.approx(700000.0)
    assert res[0]["pl_usd"] is None

def test_pl_prima_pct_usa_fx_del_calce():
    ops = [op("V", 920, "Venta", prima_usd=1500, margen_clp=0, margen_usd=0),
           op("C", 920, "Compra", prima_usd=-601.6,
              margen_clp=804993.352, margen_usd=898.4)]
    res = [{"mtm": -90000.0, "spot": 900.0}, {"mtm": 90000.0, "spot": 900.0}]
    dp._enriquecer_pl(ops, res)
    fx = 804993.352 / 898.4
    assert res[0]["pl_usd"] == pytest.approx(1500 - 100)          # V + mtm/spot
    assert res[0]["pl_clp"] == pytest.approx((1500 - 100) * fx)
    assert res[0]["fx_prima"] == pytest.approx(fx)


# ---------- salida de vencimientos: 1 o 2 tablas ----------
def _liq(o):
    return liquidar_vencimiento(dp._op_obj(o), o["usdobs"])

def test_vencimientos_una_tabla(tmp_path):
    ops = [op("N1", 900, "Venta", prima_clp=1000000, prima_clp_usd=10.0)]
    f = tmp_path / "v.xlsx"
    dp.escribir_vencimientos(ops, [_liq(o) for o in ops], str(f), date(2026, 5, 29))
    ws = openpyxl.load_workbook(f).active
    headers = [r for r in range(1, ws.max_row + 1)
               if ws.cell(r, dp.TABLE_COL0).value == "Folio"]
    assert len(headers) == 1

def test_vencimientos_dos_tablas_y_nota(tmp_path):
    ops = [op("N1", 900, "Venta", prima_clp=1000000, prima_clp_usd=10.0),
           op("P1", 920, "Venta", prima_usd=1500, margen_clp=0, margen_usd=0),
           op("P2", 920, "Compra", prima_usd=-601.6,
              margen_clp=804993.352, margen_usd=898.4)]
    f = tmp_path / "v.xlsx"
    dp.escribir_vencimientos(ops, [_liq(o) for o in ops], str(f), date(2026, 5, 29))
    ws = openpyxl.load_workbook(f).active
    headers = [r for r in range(1, ws.max_row + 1)
               if ws.cell(r, dp.TABLE_COL0).value == "Folio"]
    assert len(headers) == 2                       # dos tablas
    col_pl = dp.TABLE_COL0 + 19 - 1                 # ultima columna (P&L Nevasa)
    notas = [ws.cell(r, col_pl).comment for r in range(1, ws.max_row + 1)
             if ws.cell(r, col_pl).comment]
    assert len(notas) == 2                         # nota en ambas patas %
    assert "896.03" in notas[0].text


# ---------- seguimiento de cambios en operaciones ya vigentes ----------
def test_foto_cartera_serializa_fecha():
    o = op("N1", 900, "Venta")
    foto = dp._foto_cartera([o])
    assert set(foto.keys()) == {"N1"}
    assert foto["N1"]["venc"] == "2026-05-29"       # date -> ISO string
    assert foto["N1"]["contraparte"] == "CP"
    assert foto["N1"]["nominal"] == 100000.0

def test_detectar_cambios_sin_estado_anterior():
    # primera ejecucion (no hay foto previa aun): no debe generar ningun aviso
    avisos = dp._detectar_cambios(None, [op("N1", 900, "Venta")],
                                   [op("N1", 900, "Venta")], date(2026, 5, 29))
    assert avisos == []

def test_detectar_cambios_sin_diferencias():
    ops = [op("N1", 900, "Venta"), op("N2", 920, "Compra")]
    estado = {"operaciones": dp._foto_cartera(ops)}
    avisos = dp._detectar_cambios(estado, ops, ops, date(2026, 5, 1))
    assert avisos == []

def test_detectar_cambios_campo_modificado():
    ops_ayer = [op("N1", 900, "Venta")]
    estado = {"operaciones": dp._foto_cartera(ops_ayer)}
    o_hoy = op("N1", 900, "Venta")
    o_hoy["contraparte"] = "BANCO NUEVO"
    avisos = dp._detectar_cambios(estado, [o_hoy], [o_hoy], date(2026, 5, 1))
    assert len(avisos) == 1
    assert "N1" in avisos[0] and "Contraparte" in avisos[0]
    assert "CP" in avisos[0] and "BANCO NUEVO" in avisos[0]

def test_detectar_cambios_varios_campos_a_la_vez():
    ops_ayer = [op("N1", 900, "Venta", nominal=100000.0)]
    estado = {"operaciones": dp._foto_cartera(ops_ayer)}
    o_hoy = op("N1", 950, "Venta", nominal=150000.0)      # cambia strike y nominal
    avisos = dp._detectar_cambios(estado, [o_hoy], [o_hoy], date(2026, 5, 1))
    campos = {"Strike" in a for a in avisos} | {"Monto Principal" in a for a in avisos}
    assert len(avisos) == 2
    assert any("Strike" in a for a in avisos)
    assert any("Monto Principal" in a for a in avisos)

def test_detectar_cambios_liquidacion_anticipada():
    o_ayer = op("N1", 900, "Venta", inicio=date(2026, 4, 1), venc=date(2026, 8, 21))
    estado = {"operaciones": dp._foto_cartera([o_ayer])}
    # N1 ya no aparece en absoluto en la cartera de hoy (ni vigente ni vencida)
    avisos = dp._detectar_cambios(estado, [], [], date(2026, 5, 1))
    assert len(avisos) == 1
    assert "N1" in avisos[0] and "anticipadamente" in avisos[0]
    assert "2026-08-21" in avisos[0]

def test_detectar_cambios_vencimiento_normal_no_alerta():
    # N1 vence justo hoy: sigue en la lista completa de operaciones (como
    # "vencen"), solo deja de estar en vigentes -> no debe generar alerta.
    o_ayer = op("N1", 900, "Venta", inicio=date(2026, 4, 1), venc=date(2026, 5, 1))
    estado = {"operaciones": dp._foto_cartera([o_ayer])}
    avisos = dp._detectar_cambios(estado, [o_ayer], [], date(2026, 5, 1))
    assert avisos == []

def test_estado_cartera_guardar_y_cargar(tmp_path):
    d = str(tmp_path)
    assert dp._cargar_estado_anterior(d, date(2026, 5, 29)) is None   # no existe aun
    foto = dp._foto_cartera([op("N1", 900, "Venta")])
    dp._guardar_estado_cartera(d, foto, date(2026, 5, 29))
    cargado = dp._cargar_estado_anterior(d, date(2026, 5, 30))   # fecha posterior
    assert cargado["fecha_proceso"] == "2026-05-29"
    assert cargado["operaciones"] == foto

def test_estado_cartera_ignora_fechas_no_anteriores(tmp_path):
    d = str(tmp_path)
    dp._guardar_estado_cartera(d, dp._foto_cartera([op("VIEJO", 900, "Venta")]),
                                date(2026, 5, 27))
    dp._guardar_estado_cartera(d, dp._foto_cartera([op("RECIENTE", 900, "Venta")]),
                                date(2026, 5, 29))
    # al procesar 2026-05-30, debe traer la foto del 29 (la mas reciente < 30)
    cargado = dp._cargar_estado_anterior(d, date(2026, 5, 30))
    assert cargado["fecha_proceso"] == "2026-05-29"
    assert set(cargado["operaciones"]) == {"RECIENTE"}

def test_reprocesar_misma_fecha_no_se_compara_consigo_misma(tmp_path):
    d = str(tmp_path)
    # dia anterior (27): base real de comparacion
    dp._guardar_estado_cartera(d, dp._foto_cartera([op("N1", 900, "Venta")]),
                                date(2026, 5, 27))
    # primera corrida del 29 (con un dato erroneo: contraparte mal digitada)
    o_29_v1 = op("N1", 900, "Venta"); o_29_v1["contraparte"] = "BANCO MAL ESCRITO"
    estado_29 = dp._cargar_estado_anterior(d, date(2026, 5, 29))
    avisos_v1 = dp._detectar_cambios(estado_29, [o_29_v1], [o_29_v1], date(2026, 5, 29))
    dp._guardar_estado_cartera(d, dp._foto_cartera([o_29_v1]), date(2026, 5, 29))
    assert any("Contraparte" in a for a in avisos_v1)     # detecta el error vs. el 27

    # se corrige el input y se reprocesa el 29: debe volver a comparar contra
    # el 27 (no contra la foto erronea que dejo la primera corrida del 29)
    o_29_v2 = op("N1", 900, "Venta")                       # contraparte correcta: "CP"
    estado_29_reproceso = dp._cargar_estado_anterior(d, date(2026, 5, 29))
    assert estado_29_reproceso["fecha_proceso"] == "2026-05-27"   # no "2026-05-29"
    avisos_v2 = dp._detectar_cambios(estado_29_reproceso, [o_29_v2], [o_29_v2],
                                      date(2026, 5, 29))
    assert avisos_v2 == []       # "CP" (hoy) == "CP" (base del 27): sin cambios reales