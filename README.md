# Salman's Pipe & Fitting Order Quotation App

A Python program that automates the order-quotation workflow built up in
our chat: matching items to your master rate list, applying the correct
discount and quantity rules, and producing a formatted Order Quotation
Excel file. It also extracts PVC/APVC/CPVC items from a sales bill PDF,
and can update a price in the master rate list.

## Setup (one-time)

1. Install Python 3 if you don't have it: https://www.python.org/downloads/
2. Install the two packages this app needs:
   ```
   pip install openpyxl pdfplumber
   ```
3. Put your rate list file in the same folder as `app.py`, named
   `RATE_LIST_PIP___FITTING.xlsx` (or just tell the app the path when it
   asks - it'll offer to find it for you).

## Running it

### Option A - as a web app (Streamlit)

This is the one to deploy on **Streamlit Cloud**:

1. Push this whole folder to a GitHub repo.
2. On https://share.streamlit.io, create a new app pointing at that repo,
   with **Main file path** set to `streamlit_app.py` (not `app.py`).
   Streamlit Cloud reads `requirements.txt` automatically and installs
   `streamlit`, `openpyxl`, `pdfplumber`, and `pandas` for you.
3. Once it's live, upload your rate list `.xlsx` in the app itself (top
   of the page) - it isn't bundled into the deployment, so you upload it
   each time you open the app (or keep the browser tab open, since it's
   cached for that session).
4. Three tabs: Generate Order Quotation, Extract items from a Sales Bill
   PDF, and Update a rate list price - each mirrors the command-line
   version described below.

To run the web version locally instead of deploying it, from inside this
folder:
```
pip install -r requirements.txt
streamlit run streamlit_app.py
```

### Option B - from a terminal (command line)

```
python3 app.py
```

You'll see a menu:

```
1) Generate an Order Quotation (priced + discounted)
2) Extract items from a Sales Bill PDF and fill rates
3) Update a rate in the master rate list
4) Reload the rate list (after you've updated prices)
0) Exit
```

### 1) Generate an Order Quotation

Give it a list of items + quantities - either as a CSV file (one item per
row: `Item Name, Qty`) or typed in one at a time. It will:

- match every item to your rate list (PVC/APVC/CPVC, pipe vs fitting)
- apply the right discount % for its category
- apply the AST-brand-only rule for pipe discounts
- divide pipe quantity by 20 (PVC) or 10 (APVC/CPVC) to get billing qty
- handle the special no-discount items (Solution, Long Plug, Khilli,
  Conceal Valve, Teflon tape, 10KG pipe) and the 10KG rate x20 rule
- use the extra manually-supplied rates (CPVC solution, Jasdi Khila,
  PVC reducer 3x2II, PVC bend 1II, PVC multi floor gally tap 4'')
- write a formatted `.xlsx` quotation with a NOTE column flagging
  anything it wasn't fully sure about

**Always check the NOTE column** before sending a quotation out - it
flags low-confidence matches, items it couldn't find at all, and any
special rule it applied, so you can eyeball anything unusual.

Example CSV (`order.csv`):
```
PVC PIPE 1I X 6KG ISI AST,100
PVC ELBOW 1I,10
PVC TEE 1I,5
```

### 2) Extract items from a Sales Bill PDF

Point it at a PDF sales bill. It reads the line items, keeps only the
ones in the product families you choose (PVC/APVC/CPVC by default),
ignores the bill's own printed rate, and - if you say yes - runs them
straight through the same pricing/discount logic as option 1.

This is a simple text-pattern reader built for the kind of bill layout
you've been sending (Sr | Product Name | Qty | Rate | Amount) - always
double-check the extracted list against the PDF, especially for
multi-line product names.

### 3) Update a rate in the master rate list

For when a supplier price changes. Tell it the sheet, the text to search
for in the product name, which column (`RATE` or `PER/RATE`), and the
new price - it updates every matching row in place.

## Editing the business rules

Everything you might need to change over time is collected at the top of
`app.py`, in the section marked `CONFIG`:

- discount percentages per category
- the pipe quantity divisors (÷20 PVC, ÷10 APVC/CPVC)
- the no-discount keyword list (Solution, Long Plug, Khilli, etc.)
- the extra manually-supplied rates
- known abbreviations (MTA→MALE, CUPLINE→COUPLER, etc.)

You don't need to touch anything below the `# ===== END CONFIG =====`
line for normal day-to-day use.

## A note on matching accuracy

Item names on bills are often written differently from how they appear
in the rate list (different word order, abbreviations, missing brand
names). The app does its best with word-overlap matching plus a list of
known abbreviations, but it isn't perfect - that's exactly what the NOTE
column is for. Anything marked "LOW CONFIDENCE" or "NOT FOUND" needs a
quick manual look before the quotation goes out.

## Files in this folder

- `app.py` - all the business logic (rate matching, discount rules,
  Excel generation, PDF extraction, rate updates) + a command-line menu
- `streamlit_app.py` - the web app version for Streamlit Cloud; imports
  everything from `app.py` and wraps it in a browser UI
- `requirements.txt` - packages needed to deploy `streamlit_app.py`
- `sample_order.csv` - a small example item list you can test option 1 with
- `README.md` - this file
