import pandas as pd
import calendar
from datetime import datetime

from report import generate_excel_report


# ----------------------------------------
# 1. READ EXCEL FILE
# ----------------------------------------

def read_excel_file(file_path):
    df = pd.read_excel(file_path)
    return df


file_path = "/mnt/c/Users/hi/OneDrive/Desktop/Book1.xlsx"

df = read_excel_file(file_path)

print("\nExcel file loaded successfully!\n")
print(df)


# ----------------------------------------
# 2. PREPARE DATA
# ----------------------------------------

df["Date"] = pd.to_datetime(df["Date"])


# ----------------------------------------
# 3. FINANCIAL CALCULATIONS
# ----------------------------------------

total_revenue = df["Revenue"].sum()

total_expenditure = df["Expenditure"].sum()

total_other_income = df["Other_Income"].sum()


# Current profit
# Profit = Revenue + Other Income - Expenditure

profit = (
    total_revenue
    + total_other_income
    - total_expenditure
)


# ----------------------------------------
# 4. CURRENT PROFIT MARGIN
# ----------------------------------------

profit_margin = (
    profit / total_revenue * 100
    if total_revenue != 0
    else 0
)


# ----------------------------------------
# 5. REVENUE FORECAST
# ----------------------------------------

days_of_data = df["Date"].nunique()

average_daily_revenue = (
    total_revenue / days_of_data
)


# ----------------------------------------
# 6. EXPENDITURE FORECAST
# ----------------------------------------

average_daily_expenditure = (
    total_expenditure / days_of_data
)


# ----------------------------------------
# 7. CURRENT MONTH
# ----------------------------------------

today = datetime.now()

current_year = today.year

current_month = today.month


days_in_current_month = calendar.monthrange(
    current_year,
    current_month
)[1]


# ----------------------------------------
# 8. MONTHLY FORECAST
# ----------------------------------------

forecast_revenue = (
    average_daily_revenue
    * days_in_current_month
)

forecast_expenditure = (
    average_daily_expenditure
    * days_in_current_month
)


# Forecast profit
#
# Forecast Profit =
# Forecast Revenue
# + Other Income
# - Forecast Expenditure

forecast_profit = (
    forecast_revenue
    + total_other_income
    - forecast_expenditure
)


# ----------------------------------------
# 9. FORECAST PROFIT MARGIN
# ----------------------------------------

forecast_profit_margin = (
    forecast_profit
    / forecast_revenue
    * 100
    if forecast_revenue != 0
    else 0
)


# ----------------------------------------
# 10. DISPLAY REPORT
# ----------------------------------------

print("\n========================================")
print("       COMPANY FINANCIAL REPORT")
print("========================================")


print("\nFINANCIAL SUMMARY")
print("----------------------------------------")

print(
    f"Total Revenue       : ₹{total_revenue:,.2f}"
)

print(
    f"Total Expenditure   : ₹{total_expenditure:,.2f}"
)

print(
    f"Other Income        : ₹{total_other_income:,.2f}"
)

print(
    f"Current Profit      : ₹{profit:,.2f}"
)

print(
    f"Current Profit Margin: {profit_margin:.2f}%"
)


print("\n----------------------------------------")
print("REVENUE FORECAST")
print("----------------------------------------")

print(
    f"Average Daily Revenue      : "
    f"₹{average_daily_revenue:,.2f}"
)

print(
    f"Forecast Monthly Revenue   : "
    f"₹{forecast_revenue:,.2f}"
)


print("\n----------------------------------------")
print("EXPENDITURE FORECAST")
print("----------------------------------------")

print(
    f"Average Daily Expenditure  : "
    f"₹{average_daily_expenditure:,.2f}"
)

print(
    f"Forecast Monthly Expenditure: "
    f"₹{forecast_expenditure:,.2f}"
)


print("\n----------------------------------------")
print("PROFIT FORECAST")
print("----------------------------------------")

print(
    f"Forecast Monthly Profit    : "
    f"₹{forecast_profit:,.2f}"
)

print(
    f"Forecast Profit Margin     : "
    f"{forecast_profit_margin:.2f}%"
)


print("\n----------------------------------------")
print("BUSINESS SUMMARY")
print("----------------------------------------")

print(
    "Based on the historical data available "
    "in the Excel sheet,"
)

print(
    f"the estimated monthly revenue is "
    f"₹{forecast_revenue:,.2f}."
)

print(
    f"The estimated monthly expenditure is "
    f"₹{forecast_expenditure:,.2f}."
)

print(
    f"The estimated monthly profit is "
    f"₹{forecast_profit:,.2f}."
)


print("\n========================================")


# ----------------------------------------
# 11. GENERATE EXCEL REPORT
# ----------------------------------------

generate_excel_report(
    total_revenue,
    total_expenditure,
    total_other_income,
    profit,
    average_daily_revenue,
    forecast_revenue,
    df
)