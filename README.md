# AAI-530-IoT-Application-Design-Final-Project

Applied Artificial Intelligence for IoT Agriculture 2024. Data origination from College of Computer Science and Mathematics - Tikrit University, Iraq

---

# **IoT Agriculture Project: Smart Greenhouse Monitoring System**

**Team 4: Dylan Scott-Dawkins, Francisco Monarrez Felix, Jeffery Smith**
**Course: AAI-530 IoT Application Design — Final Team Project**
**University of San Diego — M.S. Applied Artificial Intelligence**

## Project Overview

This project focuses on designing and implementing a machine learning-based IoT system for smart agriculture. Using a dataset collected from a greenhouse in Tikrit University, Iraq, we analyze environmental and soil sensor data to provide actionable insights for farmers. The system is tailored for **Industrial IoT** applications, aiming to optimize crop growth in inhospitable environments.

Three machine learning models were developed and compared:
- **Baseline Models** — Linear Regression and Random Forest for initial benchmarking
- **XGBoost Regression** — Gradient-boosted trees for cross-sensor humidity prediction (R² = 0.90, RMSE = 7.15%)
- **LSTM / RNN** — Deep learning models for time-series forecasting of sensor readings

## Dataset Description

**Source**: [IoT Agriculture 2024 — Kaggle](https://www.kaggle.com/datasets/wisam1985/iot-agriculture-2024/data)
**Description**:

- **Observations**: 37,922 (with 1 additional observation not included in the file)
- **Features**:
  - **Environmental Metrics**: Temperature (°C), Humidity (%), Water Level (%)
  - **Soil Nutrients**: Nitrogen (N), Phosphorus (P), Potassium (K)
  - **Actuator Status**: Fan, Watering Pump, and Water Pump operational states (0/1)
- **Data Collection**:
  - Collected from a smart greenhouse at Tikrit University, monitored by Assistant Professor Wissam Dawood Abdullah.
  - Data fed into Google Sheets for real-time monitoring and control.
- **Data Cleaning**:
  - Duplicate rows removed.
  - Missing values handled.
  - Categorical columns encoded with One-Hot Encoding.

## Data Structure

| Column Name               | Type         | Description                           |
| ------------------------- | ------------ | ------------------------------------- |
| `date`                    | `datetime64` | Timestamp of data collection          |
| `temperature`             | `int64`      | Temperature in Celsius                |
| `humidity`                | `int64`      | Humidity percentage                   |
| `water_level`             | `int64`      | Water level as a percentage           |
| `N`                       | `int64`      | Nitrogen level in soil (0-255)        |
| `P`                       | `int64`      | Phosphorus level in soil (0-255)      |
| `K`                       | `int64`      | Potassium level in soil (0-255)       |
| `Fan_actuator_OFF`        | `float64`    | Fan actuator state (0 = off, 1 = on)  |
| `Fan_actuator_ON`         | `float64`    | Fan actuator state (0 = off, 1 = on)  |
| `Watering_plant_pump_OFF` | `float64`    | Watering pump state (0 = off, 1 = on) |
| `Watering_plant_pump_ON`  | `float64`    | Watering pump state (0 = off, 1 = on) |
| `Water_pump_actuator_OFF` | `float64`    | Water pump state (0 = off, 1 = on)    |
| `Water_pump_actuator_ON`  | `float64`    | Water pump state (0 = off, 1 = on)    |

## Repository Structure

### Notebooks

| File | Author | Description |
|------|--------|-------------|
| `AAI_530_Final_Notebook.ipynb` | Dylan Scott-Dawkins | Main EDA notebook with data exploration, visualization, and baseline models (Linear Regression, Random Forest) |
| `XGBoost_Regression_IoTAgriculture.ipynb` | Francisco Monarrez Felix | XGBoost regression model for humidity prediction from sensor/actuator data. Includes EDA, model training, evaluation (R²=0.90, RMSE=7.15%), feature importance analysis, and Tableau export |
| `RNN Model in progress Jeff.ipynb` | Jeffery Smith | RNN/LSTM deep learning model for time-series prediction of greenhouse sensor readings |

### Data

| File | Description |
|------|-------------|
| `IoTProcessed_Data.csv` | Cleaned and preprocessed IoT Agriculture 2024 dataset (37,922 observations, 13 columns) |

### XGBoost Model Artifacts

| File | Description |
|------|-------------|
| `xgboost_humidity_model.json` | Trained XGBoost model saved in JSON format |
| `xgboost_humidity_model.pkl` | Trained XGBoost model saved as pickle for sklearn compatibility |
| `scaler_humidity_model.pkl` | StandardScaler fitted on training features (needed for inference on new data) |
| `xgboost_metrics_summary.csv` | Model performance metrics (RMSE, MAE, R², etc.) for Tableau dashboard integration |
| `xgboost_predictions.csv` | Test set predictions with actual values and residuals (7,584 rows) |
| `xgboost_feature_importance.csv` | Feature importance rankings (gain-based and normalized) |

### Visualizations

| File | Description |
|------|-------------|
| `humidity_distribution.png` | Target variable distribution (histogram + box plot) |
| `sensor_distributions.png` | Distribution plots for all 6 sensor features |
| `correlation_heatmap.png` | Correlation matrix heatmap of all numeric features |
| `features_vs_humidity.png` | Scatter plots of each feature vs humidity |
| `humidity_by_actuator.png` | Humidity box plots grouped by actuator ON/OFF state |
| `xgboost_learning_curves.png` | Training vs test RMSE across boosting rounds |
| `predicted_vs_actual.png` | Predicted vs actual humidity scatter plot |
| `residual_analysis.png` | Residual analysis (residuals vs predicted, distribution, Q-Q plot) |
| `error_by_range.png` | Mean absolute error broken down by humidity range |
| `feature_importance.png` | XGBoost feature importance bar chart (gain-based) |
| `feature_importance_multi.png` | Feature importance from three perspectives (weight, gain, cover) |
| `xgboost_dashboard.png` | Combined 4-panel summary dashboard |

### Report

| File | Description |
|------|-------------|
| `IoT_Agriculture_Final_Report.md` | Final project report in Markdown source format |
| `IoT_Agriculture_Final_Report.docx` | Final project report converted to Word format |
| `IoT_Agriculture_Final_Report_APA7_Sun22.docx` | Final project report with APA 7th edition formatting applied |

### Scripts

| File | Description |
|------|-------------|
| `gen_report.sh` | Shell script to convert Markdown report to PDF using Pandoc |
| `gen_report_docx.sh` | Shell script to convert Markdown report to DOCX using Pandoc |
| `apa7_format.py` | Python script for applying APA 7 formatting to the DOCX report |
| `apply_apa7.py` | Python script for additional APA 7 formatting and reference styling |
| `create_apa7_reference.py` | Python script for generating APA 7 reference list formatting |
| `DOCX_CONVERSION_GUIDE.md` | Guide for the Markdown-to-DOCX conversion process |

## Models Summary

| Model | Type | Target | Key Result | Author |
|-------|------|--------|------------|--------|
| Linear Regression | Baseline | Temperature | Baseline comparison | Dylan |
| Random Forest | Baseline Classification | Fan Actuator State | Baseline comparison | Dylan |
| XGBoost Regression | Gradient-Boosted Trees | Humidity (%) | R²=0.90, RMSE=7.15% | Francisco |
| RNN / LSTM | Deep Learning | Time-series sensor values | Sequential prediction | Jeffery |

## Tableau Dashboard

The interactive Tableau Public dashboard is available at:
[IoT Smart Farm Dashboard](https://public.tableau.com/views/530IoTSmartFarm/Dashboard1)

Visualizations include average temperature, average humidity, water pump utilization rates, time-series sensor trends, and temperature heat gradients.

## How to Use

1. **Clone**: `git clone https://github.com/WiggleCode/AAI-530-IoT-Application-Design-Final-Project.git`
2. **Dataset**: The processed dataset (`IoTProcessed_Data.csv`) is included in the repo. The original raw data is available from [Kaggle](https://www.kaggle.com/datasets/wisam1985/iot-agriculture-2024/data).
3. **Run Notebooks**: Open any `.ipynb` file in Jupyter Notebook or Google Colab. Each notebook is self-contained with its own imports and data loading.
4. **Dependencies**: Python 3.8+, pandas, numpy, matplotlib, seaborn, scikit-learn, xgboost, tensorflow/keras (for RNN/LSTM notebooks)
5. **Attribution**: Credit the dataset to the original source (Tikrit University, 2023-2024).

## License

This dataset is **licensed under CC BY-ND** (Creative Commons Attribution-NonCommercial-ShareAlike). Proper attribution is required for any publication or use.

## Team Members

- **Dylan Scott-Dawkins** — Dataset EDA, Baseline Models, Report Writing
- **Francisco Monarrez Felix** — XGBoost Regression Model, Feature Engineering, Data Export for Tableau
- **Jeffery Smith** — GitHub Repository Setup, RNN/LSTM Model, Tableau Dashboard, APA 7 Formatting
