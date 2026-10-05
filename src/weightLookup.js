// ---------------------------------------------------------------------------
// Weight-per-skid lookup, built from 66 historical BOLs — both directions
// (offsite-to-Main return loads and Main-to-offsite restocking loads across
// WSI, WS2/Henderson, and EAB, plus one box truck), May–August 2026. A
// skid's weight doesn't depend on which way it's traveling, so both
// directions are combined into one average per part.
//
// Values are the weighted average lbs/skid across every real BOL that
// included the part: total weight across all loads containing that part,
// divided by total pallets — EXCEPT the items in CONFIRMED_WEIGHTS below,
// which are known single-skid weights given directly by the user and
// override whatever the BOL history would otherwise compute. When merging
// future BOL batches, don't let recalculated averages silently overwrite
// these — they're deliberately fixed until corrected again by hand.
//
// Parts where weight varied by more than ~15% between loads are marked
// "varies" below and listed in UNRELIABLE_WEIGHT_ITEMS — likely
// double-stacking or a partial pallet being recorded differently load to
// load. Parts seen on only one BOL so far are marked "only seen once" —
// still a reasonable starting estimate, just thinner evidence than the rest.
// ---------------------------------------------------------------------------

// Confirmed exact single-skid weights (lbs), given directly by the user —
// these replace the historical-average figure for these parts. All are
// drop cable parts. Note 112153 is double-stackable: a stacked pair is
// simply 2 x 450 = 900, which estimateWeight() already produces correctly
// via skids x per-skid weight — no separate "stacked" case needed.
const CONFIRMED_WEIGHTS = {
  "102092": 970.0,
  "102109": 770.0,
  "102779": 1155.0,
  "112153": 450.0,
};

const HISTORICAL_AVERAGES = {
  "100084": 1175.5, // varies 800-1250 lbs/skid across loads
  "100085": 900.0, // varies 800-1000 lbs/skid across loads
  "100087": 1701.0,
  "102005": 2638.6,
  "102027": 431.0,
  "102037": 1560.0, // only seen once
  "102071": 1033.6,
  "102078": 858.2, // varies 697-1100 lbs/skid across loads
  "102093": 795.3, // varies 450-1455 lbs/skid across loads
  "102101": 2200.0, // only seen once
  "102114": 2401.7,
  "102182": 400.0, // only seen once
  "102838": 910.0,
  "103109": 1923.1, // varies 1500-2000 lbs/skid across loads
  "103324": 912.5, // varies 800-1025 lbs/skid across loads
  "103604": 1370.2, // varies 1200-1525 lbs/skid across loads
  "103605": 1200.0, // only seen once
  "104115": 480.0,
  "104432": 600.0, // only seen once
  "104529": 673.8, // varies 600-895 lbs/skid across loads
  "104910": 1224.5,
  "105695": 1225.0,
  "105704": 2158.3, // varies 1200-2350 lbs/skid across loads
  "106371": 332.4, // only seen once
  "107092": 479.6,
  "108859": 585.0, // only seen once
  "109440": 1300.0, // only seen once
  "109441": 1625.0, // only seen once
  "109635": 1055.8,
  "111844": 518.6, // varies 400-650 lbs/skid across loads
  "112152": 350.0, // only seen once
  "114707": 850.0, // only seen once
  "114734": 450.0, // only seen once
  "114794": 1200.0, // only seen once
  "114795": 885.0, // only seen once
  "114800": 1200.0,
  "115692": 915.4,
  "115695": 450.0, // only seen once
  "115820": 1440.0, // varies 800-1600 lbs/skid across loads
  "115858": 560.0, // varies 400-640 lbs/skid across loads
  "116145": 1811.8,
  "116278": 533.3, // varies 400-640 lbs/skid across loads
  "116417": 450.0, // only seen once
  "116690": 830.8, // varies 800-1000 lbs/skid across loads
  "116786": 640.0,
  "117041": 750.3, // varies 551-850 lbs/skid across loads
  "117478": 450.0, // only seen once
  "117479": 652.5, // varies 450-855 lbs/skid across loads
  "117501": 1345.0, // only seen once
  "117627": 642.5,
  "117768": 1875.0, // only seen once
  "117831": 528.3, // varies 400-552 lbs/skid across loads
};

export const WEIGHT_PER_SKID = { ...HISTORICAL_AVERAGES, ...CONFIRMED_WEIGHTS };

// Parts where the per-load weight varied by more than ~15% — flagged in the
// UI so the estimate gets a second look before the BOL goes out. Confirmed
// weights are never flagged, even if past BOLs disagreed with them.
export const UNRELIABLE_WEIGHT_ITEMS = new Set(
  [
    "100084", "100085", "102078", "102093",
    "103109", "103324", "103604", "104529", "105704",
    "111844", "115820", "115858", "116278", "116690",
    "117041", "117479", "117831",
  ].filter((item) => !(item in CONFIRMED_WEIGHTS))
);

/**
 * Estimates total weight for a line item (skids * lbs/skid for that part),
 * rounded to the nearest 10 lbs. Confirmed weights (see CONFIRMED_WEIGHTS
 * above) take priority over the historical average. Returns null if the
 * part has no weight data at all yet — those lines are left blank for
 * manual entry, same as before.
 */
export function estimateWeight(itemNumber, skids) {
  const perSkid = WEIGHT_PER_SKID[String(itemNumber)];
  if (perSkid == null || !skids) return null;
  return Math.round((perSkid * skids) / 10) * 10;
}
