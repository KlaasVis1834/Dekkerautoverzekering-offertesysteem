def get_bonus_tabel(klant_type: str, voertuig_type: str, regio) -> dict[int, float]:
    klant_type = (klant_type or "").strip().lower()
    voertuig_type = (voertuig_type or "").strip().lower()

    try:
        regio_int = int(regio)
    except Exception:
        regio_int = 0

    if voertuig_type == "bestelauto":
        return BONUS_TABEL_BESTELAUTO

    if klant_type == "zakelijk" and voertuig_type == "personenauto":
        return BONUS_TABEL_REGIO4_OF_ZAKELIJK_PERSONEN

    if voertuig_type == "personenauto" and regio_int == 4:
        return BONUS_TABEL_REGIO4_OF_ZAKELIJK_PERSONEN

    return BONUS_TABEL_PARTICULIER_REGIO_1_3


def bonuskorting_percentage(klant_type: str, voertuig_type: str, regio, schadevrije_jaren) -> float:
    tabel = get_bonus_tabel(klant_type, voertuig_type, regio)

    try:
        svj = int(schadevrije_jaren)
    except Exception:
        svj = 0

    if svj < 0:
        svj = 0

    max_svj = max(tabel.keys())
    if svj > max_svj:
        svj = max_svj

    return float(tabel.get(svj, 0))


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
    korting_pct = bonuskorting_percentage(
        klant_type=klant_type,
        voertuig_type=voertuig_type,
        regio=regio,
        schadevrije_jaren=schadevrije_jaren,
    )

    premie_excl_belasting = premie_bij_75_incl / (1 + assurantiebelasting)
    premie_zonder_svi = premie_excl_belasting - svi_bedrag

    bruto_premie = premie_zonder_svi / (1 - (max_korting / 100))
    nieuwe_netto = bruto_premie * (1 - (korting_pct / 100))
    nieuwe_incl = (nieuwe_netto + svi_bedrag) * (1 + assurantiebelasting)

    return {
        "korting_pct": korting_pct,
        "bruto_premie": round(bruto_premie, 2),
        "premie_incl": round(nieuwe_incl, 2),
    }
