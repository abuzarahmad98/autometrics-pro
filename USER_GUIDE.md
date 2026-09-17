# 📘 Complete User Guide — Mahindra Car Metrics PPT Generator

> **For everyone** — This guide is written for users of all backgrounds. No coding experience needed. Just follow the steps!

---

## Table of Contents

1. [One-Time Setup (First Time Only)](#1-one-time-setup-first-time-only)
2. [Understanding the Project Folder](#2-understanding-the-project-folder)
3. [How to Select Cars for Comparison](#3-how-to-select-cars-for-comparison)
4. [How to Generate the PPT](#4-how-to-generate-the-ppt)
5. [Understanding the Generated PPT](#5-understanding-the-generated-ppt)
6. [How to Use the Arrow Editor](#6-how-to-use-the-arrow-editor)
7. [How to Add New Cars](#7-how-to-add-new-cars)
8. [How to Add New Metrics](#8-how-to-add-new-metrics)
9. [How to Add or Change Images](#9-how-to-add-or-change-images)
10. [How to Change Arrow Labels and Colors](#10-how-to-change-arrow-labels-and-colors)
11. [Quick Reference (Cheat Sheet)](#11-quick-reference-cheat-sheet)
12. [Troubleshooting](#12-troubleshooting)

---

## 1. One-Time Setup (First Time Only)

You only need to do this **once** on your computer. After this, everything works with simple clicks.

### Step 1.1 — Open VS Code

1. Open **Visual Studio Code** (VS Code) on your computer
2. Click **File** → **Open Folder**
3. Navigate to `C:\mahendra` and click **Select Folder**
4. You should see the project files in the left sidebar

### Step 1.2 — Open the Terminal in VS Code

1. Click **Terminal** in the top menu bar
2. Click **New Terminal**
3. A terminal panel will appear at the bottom of VS Code
4. Make sure it says `PS C:\mahendra>` (PowerShell prompt)

> [!TIP]
> **Shortcut**: Press `` Ctrl + ` `` (the key above Tab) to quickly open/close the terminal.

### Step 1.3 — Install Required Software (One Time)

Type this command in the terminal and press **Enter**:

```
pip install -r requirements.txt
```

Wait for it to finish (you'll see "Successfully installed..." messages). This installs the tools the program needs.

> [!NOTE]
> If you see "Requirement already satisfied" for everything, that means it's already installed. You're good to go!

### Step 1.4 — Verify Everything Works

Type this in the terminal and press **Enter**:

```
python generate_ppt.py
```

If it says **"PPT generated successfully!"** — you're all set! Press **Enter** to close. The test PPT file will be in the `output` folder.

![VS Code showing the terminal with generate_ppt.py running successfully](C:\Users\capta\.gemini\antigravity-ide\brain\3eb3308a-dbd3-49c0-a932-627f50ebb784\vscode_terminal_1788895994167.jpg)

---

## 2. Understanding the Project Folder

Here's what each folder and file does:

```
C:\mahendra\
│
├── 📂 data\
│   ├── 📊 car_metrics.xlsx      ← YOUR MAIN FILE (car data + selections)
│   └── 📂 images\
│       └── 🖼️ side_view.png      ← 2D car model image(s)
│
├── 📂 config\
│   └── ⚙️ arrow_config.json     ← Arrow positions (edited via Arrow Editor)
│
├── 📂 output\                    ← Generated PPTs appear here
│
├── 📂 arrow_editor\             ← Arrow Editor tool (don't modify)
│
├── 🐍 generate_ppt.py           ← The main PPT generator script
├── 🔵 run_generator.bat          ← Double-click to generate PPT
├── 🔵 run_arrow_editor.bat       ← Double-click to open Arrow Editor
└── 📄 requirements.txt           ← Software dependencies list
```

> [!IMPORTANT]
> **The only file you'll regularly edit is** [`car_metrics.xlsx`](file:///c:/mahendra/data/car_metrics.xlsx). Everything else is automated.

---

## 3. How to Select Cars for Comparison

This is what you'll do most often — the president asks "Compare XUV700, Thar, and Bolero", and you generate a PPT.

### Step 3.1 — Open the Excel File

1. Go to `C:\mahendra\data\`
2. Double-click **`car_metrics.xlsx`** to open it in Excel
3. Make sure you're on the **"Car Data"** sheet (tab at the bottom)

### Step 3.2 — Mark Cars with "Y"

1. Look at **Column A** — it says **"Select"**
2. For each car the president wants to compare, type **`Y`** in that car's Select cell
3. For cars you **don't** want, leave the Select cell **empty** (or delete the Y)

![Excel showing the Select column with Y marked for XUV700 and Scorpio N](C:\Users\capta\.gemini\antigravity-ide\brain\3eb3308a-dbd3-49c0-a932-627f50ebb784\excel_screenshot_1788895932738.jpg)

**Example**: To compare XUV700, Thar, and Bolero:
| Select | Car Name |
|--------|----------|
| **Y** | XUV700 |
| | Scorpio N |
| **Y** | Thar |
| | XUV 3XO |
| **Y** | Bolero |

### Step 3.3 — Save the Excel File

Press **Ctrl + S** to save. Close Excel if you want (not mandatory).

> [!WARNING]
> **You must save the Excel file before generating the PPT!** If you don't save, the generator will use the old selections.

---

## 4. How to Generate the PPT

You have **two ways** to generate the PPT:

### Option A: Double-Click Method (Easiest)

1. Go to `C:\mahendra\` in File Explorer
2. Double-click **`run_generator.bat`**
3. A black command window will open, showing progress
4. When it says **"PPT generated successfully!"**, press **Enter** to close
5. Your PPT is in the **`output\`** folder

### Option B: VS Code Terminal Method

1. Open VS Code with the `C:\mahendra` folder
2. Open the terminal (`` Ctrl + ` ``)
3. Type and press Enter:

```
python generate_ppt.py
```

4. Watch the progress:
```
[1/5] Reading Excel data...
  Selected cars (3): XUV700, Thar, Bolero
  Metrics found: 10

[2/5] Loading arrow configuration...
  Arrows configured: 10/10 metrics

[3/5] Assigning car colors...
  XUV700: #E31937
  Thar: #ED7D31
  Bolero: #70AD47

[4/5] Generating presentation...
  Creating title slide...
  Creating slide 1/10: Overall Length (mm)
  Creating slide 2/10: Overall Width (mm)
  ... (continues for all metrics)

[5/5] Saving presentation...
  Saved to: C:\mahendra\output\Car_Comparison_2026-09-09_010500.pptx

PPT generated successfully!
```

5. Press **Enter** to exit
6. Open the PPT from the `output\` folder

### Finding Your Generated PPT

1. Go to `C:\mahendra\output\`
2. Your file will be named like: `Car_Comparison_2026-09-09_010500.pptx`
3. The date and time are in the filename, so each generation creates a new file
4. Double-click to open in PowerPoint

---

## 5. Understanding the Generated PPT

### Slide 1: Title / Cover Slide

The first slide shows:
- **Title**: "Car Metrics Comparison"
- **Date**: When the PPT was generated
- **Car List**: All the cars being compared

### Slides 2 onwards: One Slide Per Metric

Each metric gets its own slide with three key elements:

![Example of a generated metric comparison slide](C:\Users\capta\.gemini\antigravity-ide\brain\3eb3308a-dbd3-49c0-a932-627f50ebb784\ppt_slide_example_1788896016934.jpg)

| Element | Location | What It Shows |
|---------|----------|---------------|
| **Title Bar** | Top of slide | Metric name (e.g., "Overall Length") and unit (e.g., "mm") |
| **2D Car Image with Arrow** | Top-left area | A car diagram with a red arrow showing WHERE on the car this metric is measured |
| **Bar Chart** | Bottom-left area | Colored bars comparing all selected cars, sorted highest to lowest, with actual values on top |
| **Difference Table** | Right side | Cars ranked from lowest to highest, showing how much each car differs from the one below it |

> [!TIP]
> **Each car always has the same color** across all slides. For example, if XUV700 is red on the "Overall Length" slide, it will be red on every slide. This makes it easy to follow one car across metrics.

---

## 6. How to Use the Arrow Editor

The Arrow Editor is a visual tool where you **click on a car image** to place arrows that show where each metric is measured.

### Step 6.1 — Open the Arrow Editor

**Option A**: Double-click **`run_arrow_editor.bat`** in `C:\mahendra\`

**Option B**: In VS Code terminal, type:
```
python arrow_editor\app.py
```

A browser window will automatically open at `http://localhost:5000`

![The Arrow Editor interface in your browser](C:\Users\capta\.gemini\antigravity-ide\brain\3eb3308a-dbd3-49c0-a932-627f50ebb784\arrow_editor_ui_1788895971899.jpg)

### Step 6.2 — Understand the Interface

| Area | What It Is |
|------|-----------|
| **Left Sidebar** | List of all metrics. 🟢 Green dot = arrow configured. 🔴 Red dot = not yet configured |
| **Search Box** | Type to filter/search for a specific metric |
| **Main Canvas** | The car image where you place arrows |
| **Controls Panel** | Label, color, image selection, and Save/Reset/Delete buttons |
| **Configured Counter** | Top-right shows how many metrics have arrows set up (e.g., "Configured: 7/10") |

### Step 6.3 — Place an Arrow (Step by Step)

Let's say you want to configure the arrow for **"Overall Length"**:

1. **Click "Overall Length"** in the left sidebar
2. The car image loads in the main area
3. You'll see a message: **"Click to place START point (green)"**
4. **Click on the image** where the length measurement starts (e.g., the front bumper of the car)
   - A **green dot** appears where you clicked
5. The message changes to: **"Click to place END point (red)"**
6. **Click on the image** where the length measurement ends (e.g., the rear bumper)
   - A **red dot** appears, and a **red double-headed arrow** is drawn between the two points
7. The arrow label automatically says "Overall Length" — you can change it if you want
8. **Click the teal "Save" button**
9. Done! The green dot 🟢 appears next to "Overall Length" in the sidebar

### Step 6.4 — Adjust the Arrow

**Want to redo it?** Just click on the image again:
- First click → new start point
- Second click → new end point
- Click **Save** again

**Want to start over?** Click the **"Reset"** button to clear both points.

**Want to delete the arrow completely?** Click the **"Delete"** button.

### Step 6.5 — Change the Label

1. In the **"Arrow Label"** text field (below the image), type whatever you want
2. For example, change "Overall Length" to "L99-1" or "Body Length"
3. This label appears on the arrow in the PPT slide
4. Click **Save**

### Step 6.6 — Change the Arrow Color

1. Click the **color box** next to "Arrow Color"
2. Pick any color you want from the color picker
3. Click **Save**

### Step 6.7 — Select a Different Image

1. Use the **"Image"** dropdown to select a different image
2. The canvas will reload with the new image
3. Place the arrow on the new image
4. Click **Save**

### Step 6.8 — Close the Arrow Editor

1. Go back to the terminal/command window where the editor is running
2. Press **Ctrl + C** to stop the server
3. Close the browser tab

> [!IMPORTANT]
> **Always click "Save" after placing an arrow!** If you close the browser without saving, your changes will be lost.

---

## 7. How to Add New Cars

### Step 7.1 — Open the Excel File

1. Open `C:\mahendra\data\car_metrics.xlsx`
2. Go to the **"Car Data"** sheet

### Step 7.2 — Add a New Row

1. Go to the **next empty row** below the existing cars
2. Fill in the data:

| Column | What to Type |
|--------|-------------|
| **A (Select)** | Leave blank (or type `Y` if you want it selected right away) |
| **B (Car Name)** | The car's name (e.g., "XUV 400") |
| **C onwards** | The metric values for this car (e.g., 4200 for Overall Length) |

**Example** — Adding "XUV 400":
| Select | Car Name | Overall Length (mm) | Overall Width (mm) | ... |
|--------|----------|---------------------|---------------------|-----|
| | XUV 400 | 4200 | 1822 | ... |

### Step 7.3 — Save

Press **Ctrl + S**. That's it! The new car is now available for comparison.

> [!TIP]
> You can add as many cars as you want. Just keep adding rows. The tool supports up to 70+ cars.

---

## 8. How to Add New Metrics

Adding a new metric requires **two steps** in Excel and **one step** in the Arrow Editor.

### Step 8.1 — Add the Metric Column in "Car Data" Sheet

1. Open `car_metrics.xlsx` → **"Car Data"** sheet
2. Go to the **next empty column** after the last metric
3. Type the **metric name** in the header row (Row 1)
   - Example: `Front Overhang (mm)`
4. Fill in the **values for every car** in that column

**Example**:
| ... | Turning Radius (meters) | Front Overhang (mm) |
|-----|------------------------|---------------------|
| ... | 5.9 | 895 |
| ... | 5.7 | 912 |
| ... | 5.7 | 810 |
| ... | 5.2 | 720 |
| ... | 5.6 | 780 |

### Step 8.2 — Add the Metric Config in "Metric Config" Sheet

1. Click on the **"Metric Config"** sheet tab at the bottom of Excel
2. Add a new row with:

| Column | What to Type | Example |
|--------|-------------|---------|
| **Metric Name** | Exact same name as the column header you added | `Front Overhang (mm)` |
| **Display Name (Alias)** | What you want to show on the slide title | `Front Overhang` |
| **Unit** | The unit of measurement | `mm` |
| **Image File** | Which image to use for this metric's slide | `side_view.png` |

> [!WARNING]
> **The "Metric Name" must exactly match the column header** in the "Car Data" sheet. Even a small typo (extra space, different capitalization) will cause the metric to not work properly.

### Step 8.3 — Configure the Arrow

1. Save the Excel file
2. Open the **Arrow Editor** (`run_arrow_editor.bat` or `python arrow_editor\app.py`)
3. Find your new metric in the sidebar (it will have a 🔴 red dot)
4. Click on it, place the arrow on the image, and click **Save**
5. Close the Arrow Editor

### Step 8.4 — Generate the PPT

Run the generator again — the new metric will appear as a new slide!

---

## 9. How to Add or Change Images

By default, all metrics use the same `side_view.png` image. But you might want different views for different metrics:
- **Side view** for: length, wheelbase, height, ground clearance
- **Front view** for: front width, headlight height
- **Rear view** for: tail dimensions
- **Top view** for: overall width

### Step 9.1 — Get Your Image Ready

1. The image should be a **clear 2D line drawing** of the car (PNG format preferred)
2. **White or light background** works best
3. No text or labels on the image (the tool adds the arrow and label)

### Step 9.2 — Place the Image in the Right Folder

1. Copy your image file to: **`C:\mahendra\data\images\`**
2. Give it a clear name like `front_view.png` or `top_view.png`

### Step 9.3 — Tell the System Which Metrics Use This Image

1. Open `car_metrics.xlsx` → **"Metric Config"** sheet
2. Find the metric(s) that should use the new image
3. Change the **"Image File"** column to your new image filename

**Example** — Making "Overall Width" use a front view:
| Metric Name | Display Name | Unit | Image File |
|---|---|---|---|
| Overall Width (mm) | Overall Width | mm | **front_view.png** |

### Step 9.4 — Configure Arrows on the New Image

1. Open the **Arrow Editor**
2. Click on the metric that uses the new image
3. The new image will load automatically
4. Place the arrow on the new image
5. Click **Save**

> [!NOTE]
> You can have as many images as you want in the `data\images\` folder. Just reference them by filename in the "Metric Config" sheet.

---

## 10. How to Change Arrow Labels and Colors

### Change via the Arrow Editor (Recommended)

1. Open the Arrow Editor
2. Click on the metric you want to change
3. Modify the **"Arrow Label"** text field
4. Click the **color picker** to change the arrow color
5. Click **Save**

### Change Multiple Labels Quickly (Advanced)

If you want to change many labels at once without opening the Arrow Editor:

1. Open `C:\mahendra\config\arrow_config.json` in any text editor (or VS Code)
2. Find the metric you want to change
3. Edit the `"label"` or `"color"` field:

```json
"Overall Length (mm)": {
    "start": [0.05, 0.85],
    "end": [0.95, 0.85],
    "label": "L99-1",           ← Change this to your desired label
    "color": "#0000FF",         ← Change this to your desired color (hex code)
    "image": "side_view.png"
}
```

4. Save the file

> [!TIP]
> **Common color codes**: Red = `#E31937`, Blue = `#2F5496`, Green = `#27ae60`, Orange = `#ED7D31`, Black = `#000000`

---

## 11. Quick Reference (Cheat Sheet)

### Daily Workflow (When the President Asks for a Comparison)

```
1.  Open car_metrics.xlsx
2.  Put 'Y' next to the cars to compare
3.  Save (Ctrl + S)
4.  Double-click run_generator.bat
5.  Open PPT from the output\ folder
6.  Present to the president 🎤
```

### All Tasks at a Glance

| I Want To... | Steps |
|---|---|
| **Compare specific cars** | Excel → put `Y` in Select column → Save → Run generator |
| **Add a new car** | Excel → "Car Data" sheet → add a new row → Save |
| **Add a new metric** | Excel → add column in "Car Data" + row in "Metric Config" → Arrow Editor → Save |
| **Change which image appears** | Copy image to `data\images\` → Update "Image File" in "Metric Config" sheet → Configure arrow in Arrow Editor |
| **Move/adjust an arrow** | Arrow Editor → click metric → click new start/end points → Save |
| **Change arrow label** | Arrow Editor → edit "Arrow Label" field → Save |
| **Change arrow color** | Arrow Editor → click color picker → Save |
| **Find the generated PPT** | Look in `C:\mahendra\output\` — newest file has latest timestamp |

### VS Code Terminal Commands

| Command | What It Does |
|---------|-------------|
| `python generate_ppt.py` | Generate the PPT |
| `python arrow_editor\app.py` | Start the Arrow Editor (opens in browser) |
| `pip install -r requirements.txt` | Install/update required software |
| `Ctrl + C` | Stop the Arrow Editor server |

---

## 12. Troubleshooting

### "No cars selected!" Error

**Problem**: You ran the generator but no cars have `Y` in the Select column.

**Fix**: Open `car_metrics.xlsx` → put `Y` in the Select column for at least 2 cars → Save → Try again.

---

### "Excel file not found" Error

**Problem**: The Excel file is missing or renamed.

**Fix**: Make sure the file is at exactly `C:\mahendra\data\car_metrics.xlsx`. The name must be exactly `car_metrics.xlsx`.

---

### Arrow Not Showing on a Slide

**Problem**: A slide shows the image but no arrow.

**Fix**: Open the Arrow Editor → check if that metric has a 🟢 green dot. If it has a 🔴 red dot, it means the arrow hasn't been configured yet. Click on it and set up the arrow.

---

### Arrow Editor Won't Open

**Problem**: Double-clicking `run_arrow_editor.bat` opens a black window that closes immediately.

**Fix**: Open VS Code terminal and type `python arrow_editor\app.py`. Look at the error message. Most likely Flask isn't installed — run `pip install flask`.

---

### "Port 5000 already in use" Error

**Problem**: The Arrow Editor is already running in another window.

**Fix**: Close all command prompt / terminal windows and try again. Or open your browser and go to `http://localhost:5000` — it might already be working.

---

### PPT Looks Wrong / Missing Elements

**Problem**: Images or charts appear misaligned or missing.

**Fix**:
1. Make sure all images exist in `data\images\`
2. Make sure `arrow_config.json` exists in `config\`
3. Try generating again

---

### New Metric Doesn't Show Up

**Problem**: You added a column in Excel but no new slide appears.

**Fix**: Check these things:
1. Did you add the column in the **"Car Data"** sheet? (not some other sheet)
2. Did you put the column header in **Row 1**?
3. Did you add a matching row in the **"Metric Config"** sheet?
4. Does the **Metric Name in "Metric Config" exactly match** the column header?
5. Did you **save** the Excel file?

---

> [!TIP]
> **Still stuck?** Open VS Code, run the command in the terminal, and share the error message with your IT team. The error messages are designed to be descriptive and will tell you exactly what went wrong.
