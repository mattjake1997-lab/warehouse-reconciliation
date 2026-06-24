# Warehouse Reconciliation Dashboard

Reconciles offsite warehouse inventory reports (EAB, WSI, WS2) against main warehouse shipment records.

## Deploy to Vercel (one-time setup)

1. Push this folder to a GitHub repo
2. Go to [vercel.com](https://vercel.com) → New Project → Import your repo
3. Vercel auto-detects Vite — just click **Deploy**
4. Done. Share the URL with anyone who needs access.

## Local development

```bash
npm install
npm run dev
```

## How to use

1. Open the app in any browser
2. Go to the **Upload** tab
3. Upload all four files:
   - **Main warehouse shipments** — your internal records of pallets sent
   - **EAB warehouse report** — monthly inventory report from EAB
   - **WSI warehouse report** — monthly inventory report from WSI
   - **WS2 warehouse report** — monthly inventory report from WS2
4. The dashboard auto-maps columns and runs reconciliation instantly

## Column mapping

The app auto-detects columns by matching header names. Your files don't need exact headers — it looks for close matches.

| File | Expected fields |
|------|----------------|
| Main warehouse | Part Number, Part Description, Destination Warehouse, Quantity of Pallets Sent, Shipment Date |
| EAB | Part Number, Product Name, Pallets Per Space, Date Received, Date Shipped |
| WSI / WS2 | SKU, SKU Description, Units, LPN, Location |

## Status codes

| Status | Meaning |
|--------|---------|
| Match | Shipped qty equals reported qty |
| Missing | Pallet was shipped but not reported by warehouse |
| Short | Warehouse reports fewer pallets than shipped |
| Over | Warehouse reports more pallets than shipped |
| Not shipped | Warehouse reports a pallet with no matching shipment record |

## Privacy

All file processing happens in the browser. No data is sent to any server.
