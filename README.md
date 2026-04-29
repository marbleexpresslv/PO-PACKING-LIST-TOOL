# Packing List → SPS Slab Info

Browser tool for Marble Express LV. Reads a supplier's packing slip PDF and
generates copy-paste blocks for the SPS **Packinglist Details** popup
(`vSupplierInvoice.aspx` → "Update Slab Info" → blue icon).

## How to use

1. Open `index.html` in a browser (or visit the published URL).
2. Pick the supplier from the dropdown.
3. Drop the packing-slip PDF in the drop zone.
4. The tool shows each bundle with a **Copy TSV** button.
5. In SPS, open the popup for the matching product line.
6. **Mark the unit (m / cm / inches) that matches the supplier's PDF** before
   pasting — otherwise sizes will be wrong.
7. Click the first cell of the **Length** column, paste — the grid fills.
8. Click **Add Slabs**.

## Supported suppliers

- **Milanezi Granitos** — unit: meters · block "—" replaced with PO number

More suppliers are added by writing a parser in the `PARSERS` object inside
`index.html`.

## Field mapping

| SPS popup column | Source in packing slip |
|---|---|
| Length | `L` column |
| Width | `H` column |
| Quantity | always `1` (one row per slab) |
| Block | `BLOCK` column (or PO number when `—`) |
| Bundle | `BUNDLE` header above the table |
| Supplier Reference | `BLOCK-SLAB` (e.g. `12499-02`) |

## Privacy

Fully client-side. The PDF never leaves the browser — no upload, no server.
