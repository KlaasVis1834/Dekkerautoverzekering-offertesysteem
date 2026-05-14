# bonus.py

from __future__ import annotations


# ============================================================
# BONUSKORTING TABELLEN
# ============================================================

# Particulier personenauto regio 1 t/m 3
BONUS_TABEL_PARTICULIER_REGIO_1_3: dict[int, float] = {
    0: 45.0,
    1: 50.0,
    2: 55.0,
    3: 60.0,
    4: 65.0,
    5: 67.5,
    6: 70.0,
    7: 72.5,
    8: 75.0,
}

# Regio 4 en zakelijke personenauto
BONUS_TABEL_REGIO4_OF_ZAKELIJK_PERSONEN: dict[int, float] = {
    0: 40.0,
    1: 45.0,
    2: 50.0,
    3: 55.0,
    4: 60.0,
    5: 65.0,
    6: 67.5,
    7: 70.0,
    8: 72.5,
    9: 75.0,
}

# Bestelauto (particulier en zakelijk)
BONUS_TABEL_BESTELAUTO: dict[int, float] = {
    0: 25.0,
    1: 35.0,
    2: 40.0,
    3: 45.0,
    4: 50.0,
    5: 55.0,
    6: 60.0,
    7: 65.0,
    8: 67.5,
    9: 70.0,
    10: 72.5,
    11: 75.0,
}


# ============================================================
# HULPFUNCTIES
# ============================================================

def get_bonus_tabel(
    klant_type: str,
    voertuig_type: str,
    regio,
) -> dict[int, float]:
    """
    Bepaalt welke bonus-malustabel gebruikt moet worden.
    """
    klant_type = (klant_type or "").strip().lower()
    voertuig_type = (voertuig_type or "").strip().lower()

    try:
        regio_int = int(regio)
    except Exception:
        regio_int = 0

    # Bestelauto: altijd aparte tabel
    if voertuig_type == "bestelauto":
        return BONUS_TABEL_BESTELAUTO

    # Zakelijke personenauto
    if klant_type == "zakelijk" and voertuig_type == "personenauto":
        return BONUS_TABEL_REGIO4_OF_ZAKELIJK_PERSONEN

    # Particulier regio 4
    if voertuig_type == "personenauto" and regio_int == 4:
        return BONUS_TABEL_REGIO4_OF_ZAKELIJK_PERSONEN

    # Particulier regio 1-3
    return BONUS_TABEL_PARTICULIER_REGIO_1_3


def bonuskorting_percentage(
    klant_type: str,
    voertuig_type: str,
    regio,
    schadevrije_jaren,
) -> float:
    """
    Geeft het kortingspercentage terug op basis van SVJ.
    """
    tabel = get_bonus_tabel(
        klant_type=klant_type,
        voertuig_type=voertuig_type,
        regio=regio,
    )

    try:
        svj = int(schadevrije_jaren)
    except Exception:
        svj = 0

    if svj < 0:
        svj = 0

    max_svj = max(tabel.keys())

    if svj > max_svj:
        svj = max_svj

    return float(tabel.get(svj, 0.0))


def herbereken_premie_op_svj(
    premie_bij_75_incl: float,
    schadevrije_jaren,
    klant_type: str,
    voertuig_type: str,
    regio,
    svi_bedrag: float = 2.50,
    assurantiebelasting: float = 0.21,
    max_korting: float = 75.0,
) -> dict:
    """
    Rekent een premie terug van 75% korting naar bruto premie,
    en berekent vervolgens de nieuwe premie op basis van het
    werkelijke aantal schadevrije jaren.

    premie_bij_75_incl:
        De premie inclusief assurantiebelasting en inclusief SVI,
        gebaseerd op maximale bonuskorting (75%).

    Retourneert:
        {
            "korting_pct": 45.0,
            "bruto_premie": 155.28,
            "premie_excl": 85.40,
            "premie_incl": 106.36,
        }
    """

    if premie_bij_75_incl is None:
        raise ValueError("premie_bij_75_incl is verplicht")

    premie_bij_75_incl = float(premie_bij_75_incl)

    # Juiste bonuskorting bepalen
    korting_pct = bonuskorting_percentage(
        klant_type=klant_type,
        voertuig_type=voertuig_type,
        regio=regio,
        schadevrije_jaren=schadevrije_jaren,
    )

    # Stap 1: premie exclusief assurantiebelasting
    premie_excl_belasting = premie_bij_75_incl / (1 + assurantiebelasting)

    # Stap 2: SVI eruit halen
    premie_zonder_svi = premie_excl_belasting - svi_bedrag

    if premie_zonder_svi < 0:
        premie_zonder_svi = 0.0

    # Stap 3: terugrekenen naar bruto premie (0% korting)
    factor_max = 1 - (max_korting / 100.0)  # 0.25

    if factor_max <= 0:
        raise ValueError("max_korting resulteert in ongeldige factor")

    bruto_premie = premie_zonder_svi / factor_max

    # Stap 4: nieuwe korting toepassen
    factor_nieuw = 1 - (korting_pct / 100.0)
    premie_excl = bruto_premie * factor_nieuw

    # Stap 5: SVI + assurantiebelasting toevoegen
    premie_incl = (premie_excl + svi_bedrag) * (1 + assurantiebelasting)

    return {
        "korting_pct": round(korting_pct, 2),
        "bruto_premie": round(bruto_premie, 2),
        "premie_excl": round(premie_excl, 2),
        "premie_incl": round(premie_incl, 2),
    }
