#!/usr/bin/env python3
"""
Create final output file with:
1. All tickets with their classification and resolution time
2. Summary statistics by category
"""

import pandas as pd
import numpy as np
from datetime import datetime

print("="*80)
print("📊 CREATING FINAL SUMMARY OUTPUT")
print("="*80)

# Load raw data with dates
print("\n📥 Loading data...")
df_raw = pd.read_excel('raw.xlsx', header=12)

# Convert dates
df_raw['Created_Date'] = pd.to_datetime(df_raw['Created Date (UTC+0)'], format='%d.%m.%Y %H:%M', errors='coerce')
df_raw['First_Resolved_Date'] = pd.to_datetime(df_raw['First Resolved Date (UTC+0)'], format='%d.%m.%Y %H:%M', errors='coerce')

# Calculate resolution time (days)
df_raw['Resolution_Days'] = ((df_raw['First_Resolved_Date'] - df_raw['Created_Date']).dt.total_seconds() / 3600 / 24).round(2)

# Clean incident ID
df_raw['Incident_ID_Clean'] = df_raw['Incident ID'].str.strip()

print(f"✅ Raw data: {len(df_raw)} tickets")

# Load classification
df_classification = pd.read_excel('Final_Classification_Result_Improved.xlsx', sheet_name='Tất Cả Tickets (1616)')
df_classification['ID_Clean'] = df_classification['ID'].str.strip()

print(f"✅ Classification: {len(df_classification)} tickets")

# Merge data
merged = df_raw.merge(
    df_classification[['ID_Clean', 'Phân Loại']],
    left_on='Incident_ID_Clean',
    right_on='ID_Clean',
    how='left'
)

print(f"✅ Merged: {len(merged)} tickets")

# Check for unmapped tickets
unmapped = merged[merged['Phân Loại'].isna()]
if len(unmapped) > 0:
    print(f"⚠️  {len(unmapped)} tickets without classification - assigning 'Chưa Phân Loại'")
    merged['Phân Loại'] = merged['Phân Loại'].fillna('Chưa Phân Loại')

print(f"\n{'='*80}")
print("📋 CLASSIFICATION DISTRIBUTION")
print(f"{'='*80}")

# Create all tickets sheet
all_tickets = merged[[
    'Incident_ID_Clean', 
    'Summary', 
    'Created_Date',
    'First_Resolved_Date',
    'Resolution_Days',
    'Phân Loại',
    'Status',
    'Company'
]].copy()

all_tickets.columns = [
    'Ticket ID',
    'Tóm Tắt',
    'Ngày Tạo',
    'Ngày Giải Quyết',
    'Thời Gian (ngày)',
    'Phân Loại',
    'Trạng Thái',
    'Công Ty'
]

# Format dates
all_tickets['Ngày Tạo'] = all_tickets['Ngày Tạo'].dt.strftime('%d.%m.%Y %H:%M').fillna('')
all_tickets['Ngày Giải Quyết'] = all_tickets['Ngày Giải Quyết'].dt.strftime('%d.%m.%Y %H:%M').fillna('')

print(f"\n{all_tickets.shape}")

# Create summary statistics
summary_stats = []

for category in sorted(merged['Phân Loại'].unique()):
    cat_data = merged[merged['Phân Loại'] == category]
    
    # Handle NaN values in Resolution_Days
    valid_resolution_days = cat_data['Resolution_Days'].dropna()
    
    stat = {
        'Phân Loại': category,
        'Số Lượng': len(cat_data),
        'Tỷ Lệ (%)': f"{(len(cat_data) / len(merged)) * 100:.1f}%",
        'Avg Resolution (ngày)': f"{valid_resolution_days.mean():.2f}" if len(valid_resolution_days) > 0 else "N/A",
        'Median Resolution (ngày)': f"{valid_resolution_days.median():.2f}" if len(valid_resolution_days) > 0 else "N/A",
        'Min (ngày)': f"{valid_resolution_days.min():.2f}" if len(valid_resolution_days) > 0 else "N/A",
        'Max (ngày)': f"{valid_resolution_days.max():.2f}" if len(valid_resolution_days) > 0 else "N/A",
    }
    summary_stats.append(stat)
    
    print(f"\n{category:35} | {len(cat_data):4} tickets | {(len(cat_data)/len(merged))*100:5.1f}%")
    print(f"  Avg Resolution: {valid_resolution_days.mean():.2f} days | Median: {valid_resolution_days.median():.2f} days")

summary_df = pd.DataFrame(summary_stats)

print(f"\n{'='*80}")
print("💾 EXPORTING FINAL SUMMARY")
print(f"{'='*80}")

# Write to Excel
with pd.ExcelWriter('FINAL_SUMMARY_OUTPUT.xlsx', engine='openpyxl') as writer:
    # Summary statistics
    summary_df.to_excel(writer, sheet_name='Thống Kê Summary', index=False)
    
    # All tickets
    all_tickets.to_excel(writer, sheet_name='Tất Cả Tickets', index=False)
    
    print("✅ Exported: FINAL_SUMMARY_OUTPUT.xlsx")
    print(f"  - Sheet 'Thống Kê Summary': Category statistics with resolution times")
    print(f"  - Sheet 'Tất Cả Tickets': All 1,616 tickets with classifications")

# Show summary
print(f"\n{'='*80}")
print("📊 FINAL SUMMARY")
print(f"{'='*80}")
print(f"\nTotal Tickets: {len(merged)}")
print(f"Categories: {merged['Phân Loại'].nunique()}")
print(f"\nAverage Resolution Time: {merged['Resolution_Days'].mean():.2f} days")
print(f"Median Resolution Time: {merged['Resolution_Days'].median():.2f} days")
print(f"Min: {merged['Resolution_Days'].min():.2f} days | Max: {merged['Resolution_Days'].max():.2f} days")

print(f"\n✅ Output ready: FINAL_SUMMARY_OUTPUT.xlsx")

