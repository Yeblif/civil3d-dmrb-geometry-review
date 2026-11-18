# Test data

Synthetic Civil 3D exports for trying the workbook and checking changes. They represent no real scheme, location or person.

| File | Contents |
|---|---|
| `demo_mainline.xml` | LandXML 1.2 in Civil 3D's format: one 690 m alignment `DEMO-MAINLINE`, a design profile, an existing-ground profile, a superelevation record and a 70 kph speed station. |
| `demo_mainline_superelevation.csv` | Superelevation critical stations in the same format as `DMRBEXPORTSUPER` output. |
| `generate_test_data.py` | Recreates both files (`python3 generate_test_data.py`). |

## Geometry

| Element | Chainage (m) | Detail |
|---|---|---|
| Line | 0 – 200 | |
| Clothoid | 200 – 260 | R ∞ → 255, right-hand |
| Arc | 260 – 380 | R 255 m, right-hand |
| Clothoid | 380 – 440 | R 255 → ∞ |
| Line | 440 – 690 | |
| Vertical | 0 – 300 | +2.0% |
| Crest | PVI 300, L 68 m (266 – 334) | A = 4.0%, **K = 17** |
| Vertical | 300 – 520 | −2.0% |
| Sag | PVI 520, L 60 m (490 – 550) | A = 3.0%, **K = 20** |
| Vertical | 520 – 690 | +1.0% |
| Superelevation | 260 – 380 | 5.0% full; 2.5% per 30 m transition |

At 70 kph, CD 109's desirable values are R 360 m, crest K 30 and sag K 20. That makes the R255 arc and the K17 crest each one design step below desirable. They overlap from chainage 266 to 334, which should show up in the Coincidence Review. The sag is exactly at desirable. Confirm against the workbook's own reference table rather than relying on these notes.

## Suggested setup in the workbook

- Review profile: `DEMO-MAINLINE - Design FG`
- Road type: all-purpose single carriageway
- Design speed band: A
- Superelevation environment: rural
