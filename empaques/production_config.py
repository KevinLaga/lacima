# empaques/production_config.py

# OJO: Los textos de "presentación" deben coincidir con tus Presentation.name en BD.
# Los "size" deben coincidir con los tamaños que capturas en ShipmentItem.size.
# Si el nombre en tu BD difiere un poquito, cámbialo aquí.

ALLOWED_COMBOS = [
    # 11 LBS (todas excepto TIPS)
    ("11 LBS", "JUMBO"),
    ("11 LBS", "XLarge"),
    ("11 LBS", "LARGE"),
    ("11 LBS", "STANDARD"),
    ("11 LBS", "SMALL"),

    # TIPS (15 lb) — un solo tamaño
    ("TIPS (15 LB)", "Tips"),

    # 2 PZ Walmart (de 11 lbs), un solo tamaño
    ("Walmart 11 lbs 2pz", "STANDARD"),

    # Costco / Sams pesadas (asumo size fijo)
    ("36 LBS COSTCO", "LGE"),
    ("32 LBS SAMS", "STANDARD"),

    # EAT ME BAG 11x450g (varios tamaños)
    ("11x450G EAT ME BAG", "JUMBO"),
    ("11x450G EAT ME BAG", "X-LARGE"),
    ("11x450G EAT ME BAG", "LARGE"),

    # EAT ME 11x450g (sin 'BAG'), varios tamaños
    ("11x450G EAT ME", "JUMBO"),
    ("11x450G EAT ME", "X-LARGE"),
    ("11x450G EAT ME", "LARGE"),

    # EAT ME 20x450g (tamaño LGE)
    ("20x450G EAT ME", "LGE"),

    # GOURMET 10x250g (asumo STANDAR)
    ("10x250G GOURMET", "STANDAR"),

    # EAT ME 10x250g (LGE)
    ("10x250G EAT ME", "LGE"),

    # BUD HOLLAND 11x450g (varios tamaños)
    ("11x450G BUD HOLLAND", "JUMBO"),
    ("11x450G BUD HOLLAND", "X-LARGE"),
    ("11x450G BUD HOLLAND", "LGE"),

    # SPECIAL FRUIT 11x450g (XL)
    ("11x450G SPECIAL FRUIT", "X-LARGE"),

    # 35x150g (LGE)
    ("35x150G", "LGE"),

    # 13.5 LBS 6x1KG (JUMBO y X-LARGE)
    ("13.5 LBS 6x1KG", "JUMBO"),
    ("13.5 LBS 6x1KG", "X-LARGE"),

    # 20x12oz (SMALL)
    ("20x12OZ", "SMALL"),

    # 28 LBS (STANDAR)
    ("28 LBS", "STANDAR"),

    # 2P 11x450g (LARGE)
    ("2P 11x450G", "LARGE"),

    # 11x1LB BAG (STANDAR)
    ("11x1LB BAG", "STANDAR"),
]

# Orden fijo (el orden en la tabla/Excel seguirá este índice)
ORDERED_ALIASES = {combo: idx for idx, combo in enumerate(ALLOWED_COMBOS)}