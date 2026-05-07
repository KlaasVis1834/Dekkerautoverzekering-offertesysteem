def _postcode4(postcode: str | None) -> int | None:
    if not postcode:
        return None
    s = str(postcode).strip().upper()
    digits = "".join(ch for ch in s if ch.isdigit())
    if len(digits) < 4:
        return None
    return int(digits[:4])


REGIO_1_RANGES = [
    (1600, 1799),
    (4300, 4699),
    (7700, 7999),
    (8150, 8159),
    (8300, 9999),
]

REGIO_2_RANGES = [
    (2820, 2839),
    (2850, 2879),
    (2930, 2949),
    (2957, 2979),
    (3180, 3189),
    (3200, 3299),
    (3366, 3366),
    (3380, 3389),
    (3750, 3759),
    (3770, 3799),
    (3828, 3828),
    (3829, 3829),
    (3840, 3909),
    (3925, 3939),
    (3989, 3989),
    (4000, 4289),
    (4756, 4799),
    (5300, 5339),
    (5360, 5364),
    (5430, 5459),
    (5820, 5849),
    (6550, 6799),
    (6850, 7699),
    (8000, 8149),
    (8160, 8199),
    (8250, 8299),
]

REGIO_4_RANGES = [
    (1000, 1109),
    (2500, 2599),
    (3000, 3099),
    (3500, 3599),
]


def _in_ranges(x: int, ranges) -> bool:
    return any(a <= x <= b for a, b in ranges)


def bepaal_regio(postcode: str | None) -> int | None:
    p4 = _postcode4(postcode)
    if p4 is None:
        return None

    if _in_ranges(p4, REGIO_4_RANGES):
        return 4
    if _in_ranges(p4, REGIO_1_RANGES):
        return 1
    if _in_ranges(p4, REGIO_2_RANGES):
        return 2
    return 3


def bepaal_dekking(bouwjaar: int | None) -> str:
    """
    Jouw logica:
    - bouwjaar >= 2017: WA / Casco Compleet / SVI (Allrisk)
    - bouwjaar 2010 t/m 2016: WA / Beperkt Casco / SVI
    - bouwjaar < 2010: WA / SVI
    - bouwjaar onbekend: WA / Casco Compleet / SVI (Allrisk)
    """
    if not bouwjaar:
        return "WA / Casco Compleet (Allrisk) / Schadeverzekering Inzittenden"

    bj = int(bouwjaar)

    if bj >= 2017:
        return "WA / Casco Compleet (Allrisk) / Schadeverzekering Inzittenden"
    if 2010 <= bj <= 2016:
        return "WA / Beperkt Casco / Schadeverzekering Inzittenden"
    return "WA / Schadeverzekering Inzittenden"


def benodigde_svj(klant_type: str, voertuig_type: str, regio: int | None) -> int | None:
    if not voertuig_type:
        return None

    vt = voertuig_type.lower().strip()
    kt = (klant_type or "").lower().strip()

    if vt == "bestelauto":
        return 11

    if vt == "personenauto" and kt == "zakelijk":
        return 9

    if regio == 4:
        return 9
    return 8
