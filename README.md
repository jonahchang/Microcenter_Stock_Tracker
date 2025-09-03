# MicroCenter Stock Tracker  

A Python-based automation tool that tracks **real-time inventory** of products across all MicroCenter stores in the U.S., compiles weekly stock reports into **Excel spreadsheets** for trend analysis.  

This project combines **web automation, data engineering, and visualization** into a complete pipeline for market research and business intelligence.  

---

## Features  

- **Automated Stock Collection**  
  - Uses Selenium with stealth configurations to scrape MicroCenter inventory across all stores.  
  - Supports **per-store tracking** for accurate regional demand analysis.

- **Excel Integration**  
  - Generates weekly stock reports in a single `.xlsx` file.  
  - Automatically formats data tables with **category totals, product-level trends, and store totals**.  
  - Color-coded stock changes week-to-week:  
    - 🟩 **Green** → Stock increase  
    - 🟨 **Yellow** → Stock decrease  
    - 🟦 **Blue** → Availability change  

- **Charts & Visualization**  
  - Multi-line charts showing **stock trends of all products per category**.  
  - **Stacked area charts** displaying total stock trends by category.  
  - Store-level **heatmaps** highlighting product availability across all MicroCenter locations.  

- **Product Management UI**  
  - Tkinter-based interface for adding/removing tracked products.  
  - Supports **reset to original dataset** for consistency.  
  - Product lists stored in JSON for easy editing and persistence.  

- **Performance Metrics**  
  - Tracks total runtime per session (hours/minutes/seconds).  
  - Saves progress mid-run to prevent data loss.  

---

## Tech Stack

- **Python 3.11+**  
- **Selenium + undetected_chromedriver** → Browser automation & scraping  
- **OpenPyXL** → Excel automation, chart/table creation  
- **Tkinter** → GUI for managing products  
- **Regex + JSON** → Data parsing & storage  

---

## Example Output  

- Weekly Excel sheet (`Microcenter Stock Tracker.xlsx`) with:  
  - **Category totals** (Power Supplies, Coolers, Chassis, etc.)  
  - **Product-level breakdowns** (each model tracked weekly)  
  - **Store totals** (aggregated across all products)  

- Charts automatically update as new weeks are added, enabling **longitudinal stock analysis**.  

---

## Business Value  

This tool simulates the work of an **analyst team** by automatically collecting and structuring stock data that would otherwise take **dozens of hours weekly**.  

Potential applications include:  
- Competitive market research  
- Regional demand forecasting  
- Supply chain monitoring  
- Product lifecycle analysis  

---

## How It Works  

1. **Start the script**  
   - Choose whether to create a new Excel file or update an existing one.  
   - Enter the **week number** for tracking.  

2. **Stock Tracker runs**  
   - Opens every product URL for every store ID.  
   - Extracts stock values (`25+`, `0`, or exact count).  
   - Writes results into Excel in the correct product/category rows.  

3. **Data Analysis**  
   - Highlights stock changes vs. the previous week.  
   - Prepares tables for charting.  

4. **Output**  
   - Saves updated Excel workbook with all stock data and visualizations.  

---

## Project Structure  

- MCSA.py - Main stock tracker script
- MC_products.json - Editable product list
- MC_original_products.json - Original baseline product list
- Microcenter Stock Tracker.xlsx - Output file

## Quick Start  

```bash
# Clone repo
git clone https://github.com/jonahchang/Microcenter_Stock_Tracker.git
cd Microcenter_Stock_Tracker

# Install dependencies
pip install -r requirements.txt

# Run tracker
python MCSA.py
```


## Quick Start  

```bash
# Clone repo
git clone https://github.com/jonahchang/Microcenter_Stock_Tracker.git
cd Microcenter_Stock_Tracker

# Install dependencies
pip install -r requirements.txt

# Run tracker
python MCSA.py
```

## Author

Developed by Jonah Chang - Computer Science major with a focus on software engineering, automation, and data analytics.