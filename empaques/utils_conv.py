# embarques/utils_conv.py
from decimal import Decimal, ROUND_HALF_UP

Q5 = Decimal("0.00001")

def q5(x):
    if x is None:
        x = Decimal("0")
    if not isinstance(x, Decimal):
        x = Decimal(str(x))
    return x.quantize(Q5, rounding=ROUND_HALF_UP)

# Pesos por clamshell (aprox, ajusta si tú usas otros):
KG_6OZ   = Decimal("0.170097")  # 6 oz
KG_9_8OZ = Decimal("0.277826")  # 9.8 oz
KG_18OZ  = Decimal("0.510291")  # 18 oz

def clamshells_y_kg_de_presentacion(pres, cajas):
    cs6  = pres.cs_6oz_por_caja   * int(cajas)
    cs98 = pres.cs_9_8oz_por_caja * int(cajas)
    cs18 = pres.cs_18oz_por_caja  * int(cajas)
    kg   = q5(Decimal(cs6)*KG_6OZ + Decimal(cs98)*KG_9_8OZ + Decimal(cs18)*KG_18OZ)
    return cs6, cs98, cs18, kg