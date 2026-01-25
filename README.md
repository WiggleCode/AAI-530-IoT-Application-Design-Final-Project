# AAI-530-IoT-Application-Design-Final-Project

Applied Artificial Intelligence for IoT Agriculture 2024. Data origination from College of Computer Science and Mathematics - Tikrit University, Iraq

---

# **IoT Agriculture Project: Smart Greenhouse Monitoring System**

**Team 4: Dylan Scott-Dawkins, Francisco Monarrez Felix, Jeffery Smith**  
**Course: AAI-530-IoT-Application-Design-Final-Project**

## 🎯 Project Overview

This project focuses on designing and implementing a machine learning-based IoT system for smart agriculture. Using a dataset collected from a greenhouse in Tikrit University, Iraq, we analyze environmental and soil sensor data to provide actionable insights for farmers. The system is tailored for **Industrial IoT** applications, aiming to optimize crop growth in inhospitable environments.

## 📁 Dataset Description

**Source**: [Kaggle Dataset](https://www.kaggle.com/datasets/wisam1985/iot-agriculture-2024/data)  
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

## 📊 Data Structure

| Column Name               | Type         | Description                           |
| ------------------------- | ------------ | ------------------------------------- |
| `date`                    | `datetime64` | Timestamp of data collection          |
| `temperature`             | `int64`      | Temperature in Celsius                |
| `humidity`                | `int64`      | Humidity percentage                   |
| `water_level`             | `int64`      | Water level as a percentage           |
| `N`                       | `int64`      | Nitrogen level in soil (0–255)        |
| `P`                       | `int64`      | Phosphorus level in soil (0–255)      |
| `K`                       | `int64`      | Potassium level in soil (0–255)       |
| `Fan_actuator_OFF`        | `float64`    | Fan actuator state (0 = off, 1 = on)  |
| `Fan_actuator_ON`         | `float64`    | Fan actuator state (0 = off, 1 = on)  |
| `Watering_plant_pump_OFF` | `float64`    | Watering pump state (0 = off, 1 = on) |
| `Watering_plant_pump_ON`  | `float64`    | Watering pump state (0 = off, 1 = on) |
| `Water_pump_actuator_OFF` | `float64`    | Water pump state (0 = off, 1 = on)    |
| `Water_pump_actuator_ON`  | `float64`    | Water pump state (0 = off, 1 = on)    |

## 🚀 Project Plan

1. **Data Exploration**: Analyze the dataset for trends, correlations, and anomalies.
2. **Feature Engineering**: Normalize and encode data for machine learning.
3. **Model Development**:
   - Predict environmental conditions (e.g., temperature, humidity).
   - Forecast soil nutrient levels.
   - Optimize actuator states for irrigation and ventilation.
4. **Tableau Dashboard**: Visualize insights (e.g., temperature trends, pump usage patterns).
5. **Documentation**: Add comments and explanations in code files.

## 📚 License

This dataset is **licensed under CC BY-ND** (Creative Commons Attribution-NonCommercial-ShareAlike). Proper attribution is required for any publication or use.

## 📌 How to Use

1. **Download**: Access the dataset from [Kaggle](https://www.kaggle.com/datasets/wisam1985/iot-agriculture-2024/data).
2. **Analysis**: Use Python (Pandas, Scikit-learn) for preprocessing and modeling.
3. **Visualization**: Export insights to Tableau Public for collaborative analysis.
4. **Attribution**: Credit the dataset to the original source (Tikrit University, 2023–2024).

## 🧑‍💻 Team Members

- **Dylan Scott-Dawkins**
- **Francisco Monarrez Felix**
- **Jeffery Smith**
