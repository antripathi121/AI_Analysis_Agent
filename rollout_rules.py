import pandas as pd


PHASE_1_STORES = ["Hyrum"]

PHASE_2_STORES = [
    "Hyrum",
    "Fillmore",
    "Nephi",
    "Gunnison",
    "Mt Pleasant",
]


CURRENT_PHASES = [
    {
        "name": "Current Phase 1 - Hyrum Only",
        "start": pd.Timestamp("2026-01-01"),
        "end": pd.Timestamp("2026-04-01"),
        "stores": PHASE_1_STORES,
    },
    {
        "name": "Current Phase 2 - Five Store Rollout",
        "start": pd.Timestamp("2026-04-02"),
        "end": pd.Timestamp("2026-06-15"),
        "stores": PHASE_2_STORES,
    },
    {
        "name": "Current Phase 3 - All Stores Live",
        "start": pd.Timestamp("2026-06-16"),
        "end": pd.Timestamp("2099-12-31"),
        "stores": "ALL",
    },
]


P1_PHASES = [
    {
        "name": "P1 Historical - Hyrum Only",
        "start": pd.Timestamp("2023-01-01"),
        "end": pd.Timestamp("2026-04-01"),
        "stores": PHASE_1_STORES,
    },
    {
        "name": "P1 Phase 2 - Five Store Rollout",
        "start": pd.Timestamp("2026-04-02"),
        "end": pd.Timestamp("2026-06-14"),
        "stores": PHASE_2_STORES,
    },
    {
        "name": "P1 Phase 3 - All Stores Live",
        "start": pd.Timestamp("2026-06-15"),
        "end": pd.Timestamp("2099-12-31"),
        "stores": "ALL",
    },
]


def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_store_text(value):
    text = clean_text(value).lower()

    for char in [",", ".", "-", "_", "(", ")", "&"]:
        text = text.replace(char, " ")

    return " ".join(text.split())


def store_is_allowed(store_value, active_stores):
    if active_stores == "ALL":
        return True

    store_text = normalize_store_text(store_value)

    for store in active_stores:
        store_name = normalize_store_text(store)

        if store_name in store_text:
            return True

    return False


def build_segments_from_timeline(period_start, period_end, timeline, label):
    period_start = pd.to_datetime(period_start).normalize()
    period_end = pd.to_datetime(period_end).normalize()

    segments = []

    for phase in timeline:
        overlap_start = max(period_start, phase["start"])
        overlap_end = min(period_end, phase["end"])

        if overlap_start <= overlap_end:
            segments.append({
                "period": label,
                "phase": phase["name"],
                "start": overlap_start,
                "end": overlap_end,
                "stores": phase["stores"],
            })

    if not segments:
        raise Exception(f"{label} does not overlap with configured rollout rules.")

    return segments


def build_current_coverage_segments(current_start, current_end):
    return build_segments_from_timeline(
        current_start,
        current_end,
        CURRENT_PHASES,
        "Current"
    )


def build_p1_coverage_segments(p1_start, p1_end):
    return build_segments_from_timeline(
        p1_start,
        p1_end,
        P1_PHASES,
        "P1"
    )


def build_p2_coverage_segments(p2_start, p2_end):
    p2_start = pd.to_datetime(p2_start).normalize()
    p2_end = pd.to_datetime(p2_end).normalize()

    return [{
        "period": "P2",
        "phase": "P2 Historical - Hyrum Only",
        "start": p2_start,
        "end": p2_end,
        "stores": PHASE_1_STORES,
    }]