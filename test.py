import pandas as pd
import itertools
import copy

input_file_path = "PythonSampleData0205250.xlsx"
output_file_path = "output.xlsx"

df = pd.read_excel(input_file_path, sheet_name="Input")
df.columns = df.columns.str.strip()  

group_rows = []

for customer_id in df["Customer ID"].unique():
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
        # print(f"temp array is {temp_array}")
        for j in range(length):
            key = key_value[j]
            if amount not in key_value:
                key_value.append(amount)
                index_array.append([i])
                # print(index_array)
                cnt += 1
            if amount + key not in key_value:
                key_value.append(key + amount)
                index_array.append(temp_array[j] + [i])
                # print(f"temp is {temp_array}")
                # print(index_array)
                cnt += 1
            else:
                origin_index = key_value.index(key + amount)
                index = key_value.index(key)
                if index_array[origin_index][0] > index_array[index][0]:
                    index_array[origin_index] = temp_array[j] + [i]
                # print(temp_array)
                # print(index_array)
    print(key_value)
    print(index_array)
    print(cnt)
    
    
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

output_rows = [r for r in group_rows]
output_df = pd.DataFrame(output_rows)

print(output_df.columns)
output_df = output_df[["Customer ID", "Key Amount", "Invoice Number", "Date", "Amount"]]


with pd.ExcelWriter(output_file_path, engine="openpyxl") as writer:
    output_df.to_excel(writer, index=False, sheet_name="Output2")

print(f"Filtered Output2 written to {output_file_path}")
