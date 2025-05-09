import pandas as pd
import copy
from concurrent.futures import ThreadPoolExecutor

input_file_path = "PythonSampleData0205250.xlsx"
output_file_path = "output.xlsx"

df = pd.read_excel(input_file_path, sheet_name="Input")
df.columns = df.columns.str.strip()

def process_customer(customer_id):
    customer_df = df[df["Customer ID"] == customer_id].sort_values("Date").reset_index(drop=True)
    
    if len(customer_df) > 15:
        customer_df = customer_df.head(15)
    
    n = len(customer_df)
    
    key_value = []
    key_value.append(customer_df.iloc[0]["Amount"])
    index_array = []
    index_array.append([0])
    cnt = 1
    
    for i in range(1, n):
        amount = customer_df.iloc[i]["Amount"]
        length = len(key_value)
        
        temp_array = copy.deepcopy(index_array)
        
        for j in range(length):
            key = key_value[j]
            if amount not in key_value:
                key_value.append(amount)
                index_array.append([i])
                cnt += 1
            if amount + key not in key_value:
                key_value.append(key + amount)
                index_array.append(temp_array[j] + [i])
                cnt += 1
            else:
                origin_index = key_value.index(key + amount)
                index = key_value.index(key)
                if index_array[origin_index][0] > index_array[index][0]:
                    index_array[origin_index] = temp_array[j] + [i]

    group_rows = []
    for i in range(cnt):
        group_key = key_value[i]
        indexs = index_array[i]
        
        for index in indexs:
            group_rows.append({
                "Customer ID": customer_id,
                "Key Amount": group_key,
                "Invoice Number": customer_df.iloc[index]["Invoice Number"],
                "Date": pd.to_datetime(customer_df.iloc[index]["Date"]).date(),
                "Amount": customer_df.iloc[index]["Amount"],
            })
    
    return group_rows

# Using ThreadPoolExecutor for multi-threading
all_group_rows = []
with ThreadPoolExecutor() as executor:
    results = list(executor.map(process_customer, df["Customer ID"].unique()))

# Flatten the list of results
for result in results:
    all_group_rows.extend(result)

# Create DataFrame from the results
output_df = pd.DataFrame(all_group_rows)

# Check columns and filter
print(output_df.columns)
output_df = output_df[["Customer ID", "Key Amount", "Invoice Number", "Date", "Amount"]]

# Write to Excel
with pd.ExcelWriter(output_file_path, engine="openpyxl") as writer:
    output_df.to_excel(writer, index=False, sheet_name="Output2")

print(f"Filtered Output2 written to {output_file_path}")
