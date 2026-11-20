# Dubai Money Journey 2025–2026 (private)

Rajesh's personal income / spending / debt analysis, Aug 2025 – 19 Nov 2026. All amounts in AED.

## Latest workbook
`output/Dubai_Money_Journey_2025-2026_v6.xlsx`

| Page | What it shows |
|---|---|
| 1. Income | Every salary payment (Thiran, Envision, Envision via Aram in KSA, GenMahi) + Envision travel allowance |
| 2. Money out | Where the money went: buckets, income vs money out, work travel vs Envision allowance, categories, month by month, every outgoing payment |
| 3. India transfers | All 38 transfers to Sornapriya (AED + INR), per month |
| 4. Spending by category | Every payment grouped by category with totals and Need / Cut flags |

## Key numbers (locked)
| | AED |
|---|---|
| Total income | 333,598.83 (salary 310,700.47 + Envision travel allowance 22,898.36) |
| Sent to India | 162,626.46 |
| Total money out | 495,695.74 (FAB & RAKBANK card purchases now itemised, card bill payments no longer counted) |
| Work travel (Envision allowance) | 44,647.59 vs allowance 22,898.36 |
| Debt on 19 Nov 2026 | 117,660 (ENBD 46,867 · Emirates Islamic 44,614 · Tabby 10,000 · FAB 9,890 · RAK 6,289) **+ RAKBANK card-cheque loan 8,274.03** (12 months from 4 Nov 2026, 779.07/month) – to confirm if the 6,289 already includes it |

## Rules agreed with Rajesh
- Only credits labelled *Salary* from Envision count as Envision income (plus the 16 Apr 5,000 advance).
- RAK credits after 30 Mar 2026 are his own transfers, not income.
- Cash deposits during the Thiran period are Thiran salary re-deposited (own money, not income).
- Travel & hotels, Saudi purchases and Saudi ATM cash are work travel, covered by the Envision travel allowance.
- Money moved between his own accounts is never counted.
- Always ask Rajesh for passwords of locked statements (never stored in this repo).

## Still to do
1. ~~FAB & RAKBANK credit card statements~~ – done (v6).
2. Confirm RAK debt: does 6,288.66 include the 8,274.03 card-cheque instalment?
3. Page 5: Debts & payoff plan (10-month goal).
4. Open questions: Rajeswaran Subramanian 6,763.03 (May 2026); Envision-period cash deposits 5,700 / 3,250 / 19,500.

## Rebuild
Original statements are in `statements/` (most are password-protected; passwords are never stored here). The cleaned transaction data is in `work/*.json`; card parsers are in `work/py/` (`rakcc.py`, `fab.py`).

```
cd work/xl
python3 page34.py ../../output/Dubai_Money_Journey.xlsx   # needs openpyxl
```
