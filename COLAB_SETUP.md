# Google Colab Setup Guide — step-by-step

## TL;DR (2 minutes)
1. Zip this folder
2. Upload `cold_hot_tumour_project.zip` to your Google Drive root
3. Open `ColabRunner.ipynb` in Colab (https://colab.research.google.com)
4. Run cells top-to-bottom
5. Share the Drive folder `cold_hot_tumour_outputs/` with your team

---

## Step 1 — Zip and upload

From this Linux terminal (already done for you — see the zip created in `~/`):
```bash
cd /home/aryan
zip -r cold_hot_tumour_project.zip cold_hot_tumour_project \
    -x "cold_hot_tumour_project/data/*" \
    -x "cold_hot_tumour_project/outputs/*" \
    -x "cold_hot_tumour_project/.venv/*"
```

Then:
- Open https://drive.google.com
- Upload `cold_hot_tumour_project.zip` to **My Drive** (root, not inside a folder)

## Step 2 — Open the notebook in Colab

Two easy ways:

**Option A (upload notebook to Drive):**
1. Upload `ColabRunner.ipynb` to My Drive as well
2. Double-click it in Drive → "Open with Google Colaboratory"

**Option B (open from GitHub):**
1. Create a GitHub repo for this project, push everything
2. In Colab: File → Open notebook → GitHub tab → paste the repo URL

## Step 3 — Run the pipeline

In the notebook:
1. **Runtime → Change runtime type → Python 3** (no GPU needed, but feel free to enable it for speed)
2. Run cell §1.1 — it will ask you to sign in to Google and allow Drive access
3. Run cells §1.2 – §1.4 to install dependencies
4. Run §2 (data acquisition) — **~30-60 minutes, needs ~3 GB**
5. Continue top-to-bottom through §3 – §11
6. All outputs land in `Google Drive → cold_hot_tumour_outputs/`

## Step 4 — Share with your team

1. Right-click `cold_hot_tumour_outputs/` in Drive
2. Share → add your teammates' emails
3. Give them **Editor** access so they can add their own notebook cells

Each teammate opens `ColabRunner.ipynb` themselves, runs §1 (setup), then only their assigned section. Because `outputs/` lives in shared Drive, everyone sees everyone else's results.

---

## Common gotchas

### "My Colab disconnected mid-download"
No problem — `01_data_acquisition.py` is idempotent. Just re-run the cell; it skips files already on disk.

### "It says `ZIP_PATH does not exist`"
Make sure the zip is at the Drive root, i.e. path ends with `/content/drive/MyDrive/cold_hot_tumour_project.zip`, not inside a sub-folder.

### "Colab free tier timed out after 12 hours"
Same fix — just reconnect and re-run. Everything is cached to Drive.

### "Running locally instead of Colab"
```bash
cd /home/aryan/cold_hot_tumour_project
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
python 01_data_acquisition.py   # ≥ 10 GB free disk required
# ...etc
```

---

## Colab quota tips
- Free tier = 12h max runtime, 12 GB RAM, ~100 GB temporary disk
- Colab Pro ($10/mo) = 24h, more RAM — **NOT needed** for this project
- GDC download uses ~3 GB, processed files ~1 GB, so you're well inside free limits
