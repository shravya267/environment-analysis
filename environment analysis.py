import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
file_path = r"C:\Users\GHANASHYAM.Y\Desktop\New folder\Environmental_Dataset_50x10_processed.xlsx"
try:
    df = pd.read_excel(file_path, engine='openpyxl')

    df.columns = df.columns.str.strip()

    print("Dataset Loaded Successfully!\n")
    print(df.head())  

    output_path = file_path.replace(".xlsx", "_cleaned.xlsx")
    df.to_excel(output_path, index=False)
    print(f"Processed Excel file saved at: {output_path}")

    plt.figure(figsize=(12, 7))
    sns.heatmap(df.select_dtypes(include=['number']).corr(), annot=True, cmap="coolwarm", linewidths=0.5)
    plt.title("Correlation Heatmap of Environmental Factors")
    plt.show()

    target_column = "Forest Cover"
    matching_columns = [col for col in df.columns if target_column in col]

    if matching_columns:
        column_name = matching_columns[0]  
        plt.figure(figsize=(10, 6))
        sns.barplot(x=df.iloc[:, 0], y=df[column_name], palette="viridis")  
        plt.xlabel(df.columns[0])  
        plt.ylabel(column_name)
        plt.title(f"{column_name} by {df.columns[0]}")
        plt.xticks(rotation=45)
        plt.show()
    else:
        print(f"Column containing '{target_column}' not found in dataset. Please check the column names.")

except FileNotFoundError:
    print("❌ Error: File not found! Please check the path and try again.")
except Exception as e:
    print(f"❌ An error occurred: {e}")
