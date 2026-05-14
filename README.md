# Chemical Formulation Cost Optimizer

An automated tool that optimizes chemical recipe costs using **Python (PuLP)** and **Excel (VBA)**.

## 🧪 Project Overview
This project solves a common problem in chemical engineering: finding the most cost-effective mixture of raw materials while maintaining target quality standards (Active Matter Content) and respecting technical limits (Stock & Usage Limits).

## 🚀 Key Features
- **Linear Programming:** Uses the PuLP library to minimize costs.
- **Excel Integration:** Seamless data exchange between Excel and Python via VBA.
- **Constraints Management:** Handles raw material stocks, active matter targets, and maximum usage percentages.
- **Dynamic Reporting:** Automatically writes optimized results back to the Excel "Results" sheet.

## 🛠️ Tech Stack
- **Python 3.x** (Pandas, PuLP, Win32COM)
- **Excel VBA**
- **Linear Solver:** CBC (Coin-or branch and cut)

## 📖 How to Use
1. Clone this repository.
2. Install dependencies: `pip install -r requirements.txt`.
3. Open `optimization_input.xlsm` and enter your data in the *Materials* and *Constraints* sheets.
4. Click the **"Optimize"** button.

You can watch the demo video <a href="https://vimeo.com/1192242165?share=copy&fl=sv&fe=ci">here</a>
