# GUI Plotter App

This is a Python desktop application to visualize data from Excel (.xlsx), CSV, or TSV files using a graphical user interface. It's built using PyQt5 for the GUI, along with Matplotlib, Seaborn, and Plotly for plotting.

---

## Features

- Load `.xlsx`, `.csv`, or `.tsv` files
- Select up to 10 columns each for X and Y axes
- Choose from multiple plot types: Boxplot, Violin, Scatter, Line, Histogram, Bar
- Assign a custom color to each Y-axis column
- Customize the plot title, axis labels, and enable/disable grid
- Export static plots as PNG
- Export interactive plots as HTML (via Plotly)
- View basic summary statistics (mean, median, standard deviation)

---

## Installation

You’ll need Python 3.7 or higher. Install dependencies using:

```bash
pip install pyqt5 pandas matplotlib seaborn plotly openpyxl
```

---

## How to Run

Clone the repository and run:

```bash
python GUI_plotter_app.py
```
![image](https://github.com/user-attachments/assets/329e3e8b-1c58-48a8-8f08-487a0918681b)

![image](https://github.com/user-attachments/assets/a97edd83-0eed-4faa-afae-aa0ff1658eed)

---

## Coming Soon

- Debian installer (`.deb`) for Linux
- Drag and drop file loading
- Export summary stats to `.csv`

---

## About

This app was built to simplify the process of plotting and exploring tabular data. Instead of writing plotting scripts each time, just load your file and generate plots with a few clicks.

If you have suggestions or run into any issues, feel free to open an issue or pull request.
```
