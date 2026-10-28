# Freight Bilty-to-Tally Agent — Online Demo

**SN Softech Solutions** · Smart Software Solutions for Growing Businesses · 📞 9993199719

This is the **online demo** of *Freight Bilty-to-Tally Agent*, which turns a transporter's
freight / bilty register (Excel) into Tally entries.

## What the software does
- Reads the bilty register Excel (several files and sheets at once) and finds the right column layout by itself
- Makes the Tally entries for every bilty, like a Tally Day Book:

  | Voucher | Debit | Credit |
  |---|---|---|
  | Payment | Other Expenses | Cash (cash from office) |
  | Journal | TDS | Party |
  | Journal | Diesel Expenses + Other Expenses | Petrol pump |
  | Journal | Transporting Expenses | Truck owner |
  | Sales (item invoice) | Party | Freight — Item, Qty (weight) × Rate |

- Check and correct any value on screen before posting
- Bilty number as voucher number, choose which entries to post
- Full version: sends entries **straight into TallyPrime**, with Replace, Undo and a Pending list for when Tally is closed

## Using the demo
1. Open **Import & Generate**
2. Press **📄 Use the demo sample Excel** (or upload your own register)
3. Check the table and the Day Book-style preview
4. Press **Generate Tally XML file** to download the Tally data
5. **↺ Reset demo** (left side) restores the demo settings

All names and amounts in the sample Excel are invented. Demo data is reset from time to time.

## Run this demo on Streamlit Community Cloud
Repository: this repo · Branch: `main` · Main file path: `app.py` (details in `DEPLOY_DEMO.txt`).

## Buy the full version
SN Softech Solutions — WhatsApp / call **9993199719**
