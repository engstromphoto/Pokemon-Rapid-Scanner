# Pokémon Rapid Scanner

A phone-friendly Pokémon card scanner hosted with GitHub Pages. It uses the rear camera and browser OCR, looks cards up through the Pokémon TCG API, counts duplicates, and sends confirmed cards to a Google Sheet through Google Apps Script.

## 1. Install the Apps Script backend

1. Open your Pokémon Collection spreadsheet.
2. Choose **Extensions → Apps Script**.
3. Replace `Code.gs` with the contents of [`apps-script/Code.gs`](apps-script/Code.gs).
4. Run `setupSheets` once and approve access.
5. Choose **Deploy → New deployment → Web app**.
6. Set **Execute as: Me** and **Who has access: Anyone**.
7. Deploy and copy the URL ending in `/exec`.

## 2. Enable GitHub Pages

Open this repository’s **Settings → Pages** and choose:

- Source: **Deploy from a branch**
- Branch: **main**
- Folder: **/(root)**

The scanner will be available at:

`https://engstromphoto.github.io/Pokemon-Rapid-Scanner/`

## 3. Connect the scanner

Open the scanner, tap **Settings**, paste the Apps Script `/exec` URL, and save. Then tap **Start Camera**.

The scanner page stores the backend URL only on your device.