# MES COM7020 — Cloud-Enabled Big Data Ecosystem

## Project overview

This repository contains the technical report and proof-of-concept (PoC) implementation for the MetroEnergy Solutions (MES) cloud-enabled big data ecosystem.

The PoC demonstrates a **batch analytics layer for short-term electricity demand forecasting** using synthetic half-hourly smart-meter data. It covers data acquisition, feature engineering, model training, evaluation and visualisation.

## Repository contents

```text
MES_COM7020/
 README.md
requirements.txt
PoC Data_105602.xlsx
 mes_poc_generation_and_forecasting.ipynb
generate_mes_submission.py
 metrics.csv
 fig1_shared_responsibility.png
fig2_architecture.png
fig3_dataflow.png
fig4_forecast_vs_actual.png
fig5_feature_importance.png
 fig6_daily_profile.png
```

## Proof of concept

The notebook demonstrates:

1. Synthetic half-hourly smart-meter data acquisition
2. Feature engineering using temporal and lag-based features
3. Gradient Boosting Regression for short-term demand forecasting
4. Model evaluation using MAE, MAPE and R²
5. Forecast-versus-actual visualisation
6. Feature-importance analysis
7. Weekday versus weekend demand-profile analysis

The dataset is synthetic and is included to support reproducibility of the PoC.

## How to run

### 1. Create and activate a virtual environment

```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Run the notebook

```bash
jupyter notebook mes_poc_generation_and_forecasting.ipynb
```

The notebook can also be opened in JupyterLab or imported into Google Colab. If using Colab, upload the repository files first and ensure the required dependencies are available.

## Technologies

- Python
- Jupyter Notebook
- Pandas
- NumPy
- Matplotlib
- Scikit-learn
- OpenPyXL

## Academic submission

The Word document contains the technical report. The Jupyter notebook, dataset and supporting outputs provide the implementation evidence for the proof of concept.

The repository is intended to make the implementation and supporting evidence easy to inspect and reproduce.
