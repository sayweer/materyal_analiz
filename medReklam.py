import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

file_path = "Q1material.xlsx"
data = pd.ExcelFile(file_path)
material_data = data.parse('Sayfa1')

material_data['ZAMAN'] = pd.to_datetime(material_data['ZAMAN'])

def cap_outliers(df, columns):
    for col in columns:
        Q1 = df[col].quantile(0.25)
        Q3 = df[col].quantile(0.75)
        IQR = Q3 - Q1
        lower_bound = Q1 - 1.5 * IQR
        upper_bound = Q3 + 1.5 * IQR

        df[col] = np.where(df[col] < lower_bound, lower_bound, df[col])
        df[col] = np.where(df[col] > upper_bound, upper_bound, df[col])
    return df

material_data = material_data.ffill()

numeric_cols = ['EN', 'PARÇA GRAMAJ', 'PARÇA NEMİ (%)', 'ISI ÇEKME EN',
                'ISI ÇEKME BOY', 'GENEL EN', 'GENEL  BOY', 'ORTAM SICAKLIK', 'ORTAM NEM']

material_data = cap_outliers(material_data, numeric_cols)

material_data_sampled = material_data.iloc[::10, :].copy()

material_data_sampled.loc[:, 'GENEL EN'] = material_data_sampled['GENEL EN'].rolling(window=15).mean()

plt.figure(figsize=(12, 6))
plt.plot(material_data_sampled['ZAMAN'], material_data_sampled['GENEL EN'], label='Genel Genişlik', color='#1f77b4', linewidth=2)

plt.gca().xaxis.set_major_locator(mdates.WeekdayLocator())
plt.gcf().autofmt_xdate()

plt.title('Genel Genişlik Değişimi', fontsize=16, fontweight='bold')
plt.xlabel('Zaman', fontsize=14)
plt.ylabel('Genişlik (mm)', fontsize=14)

plt.grid(True, linestyle='--', alpha=0.8)

plt.legend(fontsize=12)

plt.tight_layout()
plt.show()

stabilization_date = material_data['ZAMAN'].iloc[int(len(material_data) * 0.75)]

plt.figure(figsize=(12, 6))
plt.plot(material_data_sampled['ZAMAN'], material_data_sampled['GENEL EN'], label='Genel Genişlik', color='#1f77b4', linewidth=2)

plt.axvline(x=stabilization_date, color='r', linestyle='--', linewidth=2, label='Stabilizasyon Noktası')

plt.gca().xaxis.set_major_locator(mdates.WeekdayLocator())
plt.gcf().autofmt_xdate()

plt.title('Malzeme Genişlik Değişimi ve Stabilizasyon Noktası', fontsize=16, fontweight='bold')
plt.xlabel('Zaman', fontsize=14)
plt.ylabel('Genişlik (mm)', fontsize=14)

plt.grid(True, linestyle='--', alpha=0.8)

plt.legend(fontsize=12)

plt.tight_layout()
plt.show()

print(f"Malzemenin stabil hale geldiği tarih: {stabilization_date}")
