import pandas as pd
import random
from datetime import datetime, timedelta

# 1. Define static data and parameters
products_list = ['OR-600', 'LE-600', 'CO-600', 'DC-600', 'RB-600', 'CO-2L']
product_weights = [0.05, 0.15, 0.40, 0.10, 0.15, 0.15]

operators_list = ['Mac', 'Charlie', 'Dee', 'Dennis']
min_batch_times = {
    'OR-600': 60, 'LE-600': 60, 'CO-600': 60,
    'DC-600': 60, 'RB-600': 60, 'CO-2L': 98
}

num_batches = 100000
start_date = datetime(2024, 8, 29)

productivity_data = []
downtime_data = []

current_batch_id = 422111
current_time = start_date

# 2. Generate data loop
for i in range(num_batches):
    product = random.choices(products_list, weights=product_weights, k=1)[0]
    operator = random.choice(operators_list)

    # Downtime logic
    if random.random() < 0.08:
        total_downtime = 0
        downtime_factors = {}
    else:
        downtime_factors = {}
        total_downtime = 0

        factor_probabilities = {
            1: 0.02, 2: 0.15, 3: 0.05, 4: 0.20, 5: 0.10,
            6: 0.25, 7: 0.25, 8: 0.15, 9: 0.05, 10: 0.10,
            11: 0.10, 12: 0.15
        }

        for factor_id, prob in factor_probabilities.items():
            if random.random() < prob:
                factor_downtime = random.randint(5, 60)
                downtime_factors[factor_id] = factor_downtime
                total_downtime += factor_downtime

        # Special case for CO-2L extra downtime
        if product == 'CO-2L' and total_downtime > 0:
            extra_downtime = random.randint(10, 30)
            if 7 in downtime_factors:
                downtime_factors[7] += extra_downtime
            else:
                downtime_factors[12] = extra_downtime
            total_downtime += extra_downtime

    # Time calculations
    batch_start_time = current_time
    production_time = min_batch_times[product] + total_downtime
    batch_end_time = batch_start_time + timedelta(minutes=production_time)

    # Append productivity data
    productivity_data.append({
        'Date': batch_start_time.strftime('%Y-%m-%d'),
        'Product': product,
        'Batch': current_batch_id,
        'Operator': operator,
        'Start Time': batch_start_time.strftime('%H:%M:%S'),
        'End Time': batch_end_time.strftime('%H:%M:%S')
    })

    # Append downtime data (Wide format)
    downtime_row = {'Batch': current_batch_id}
    for factor_id in range(1, 13):
        downtime_row[str(factor_id)] = downtime_factors.get(factor_id, None)

    downtime_data.append(downtime_row)

    # Update for next batch
    current_batch_id += 1
    current_time = batch_end_time + timedelta(minutes=random.randint(5, 25))

# 3. Create Dimension Tables
df_products = pd.DataFrame({
    'Product': ['OR-600', 'LE-600', 'CO-600', 'DC-600', 'RB-600', 'CO-2L'],
    'Flavor': ['Orange', 'Lemon lime', 'Cola', 'Diet Cola', 'Root Berry', 'Cola'],
    'Size': ['600 ml', '600 ml', '600 ml', '600 ml', '600 ml', '2 L'],
    'Min batch time': [60, 60, 60, 60, 60, 98]
})

df_downtime_factors = pd.DataFrame({
    'Factor': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12],
    'Description': [
        'Emergency stop', 'Batch change', 'Labeling error', 'Inventory shortage',
        'Product spill', 'Machine adjustment', 'Machine failure', 'Batch coding error',
        'Conveyor belt jam', 'Calibration error', 'Label switch', 'Other'
    ],
    'Operator Error': ['No', 'Yes', 'No', 'No', 'Yes', 'Yes', 'No', 'Yes', 'No', 'Yes', 'Yes', 'No']
})

df_productivity = pd.DataFrame(productivity_data)
df_downtime = pd.DataFrame(downtime_data)

# 4. Export to Excel with exact sheet names
file_name = "Manufacturing_Line_Productivity.xlsx"
with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
    df_productivity.to_excel(writer, sheet_name='Line productivity', index=False)
    df_downtime.to_excel(writer, sheet_name='Line downtime', index=False)
    df_products.to_excel(writer, sheet_name='Products', index=False)
    df_downtime_factors.to_excel(writer, sheet_name='Downtime factors', index=False)

print("Data generated successfully!")
print("Total batches:", len(df_productivity))