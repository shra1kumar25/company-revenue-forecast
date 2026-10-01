import pandas as pd
import calendar
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

from openpyxl.chart import BarChart, LineChart, Reference


def generate_excel_report(
    total_revenue,
    total_expenditure,
    total_other_income,
    profit,
    average_daily_revenue,
    forecast_revenue,
    df
):

    # ----------------------------------------
    # 1. FORECAST CALCULATIONS
    # ----------------------------------------

    days_of_data = df["Date"].nunique()

    average_daily_expenditure = (
        total_expenditure / days_of_data
    )

    today = datetime.now()

    days_in_current_month = calendar.monthrange(
        today.year,
        today.month
    )[1]

    forecast_expenditure = (
        average_daily_expenditure
        * days_in_current_month
    )

    forecast_profit = (
        forecast_revenue
        + total_other_income
        - forecast_expenditure
    )


    # ----------------------------------------
    # 2. PROFIT MARGINS
    # ----------------------------------------

    profit_margin = (
        profit / total_revenue * 100
        if total_revenue != 0
        else 0
    )

    forecast_profit_margin = (
        forecast_profit / forecast_revenue * 100
        if forecast_revenue != 0
        else 0
    )

    expenditure_ratio = (
        total_expenditure / total_revenue * 100
        if total_revenue != 0
        else 0
    )


    # ----------------------------------------
    # 3. FINANCIAL SUMMARY
    # ----------------------------------------

    summary_data = {
        "Metric": [
            "Total Revenue",
            "Total Expenditure",
            "Other Income",
            "Current Profit",
            "Current Profit Margin (%)",
            "Average Daily Revenue",
            "Forecast Monthly Revenue",
            "Average Daily Expenditure",
            "Forecast Monthly Expenditure",
            "Forecast Monthly Profit",
            "Forecast Profit Margin (%)"
        ],
        "Amount": [
            total_revenue,
            total_expenditure,
            total_other_income,
            profit,
            profit_margin,
            average_daily_revenue,
            forecast_revenue,
            average_daily_expenditure,
            forecast_expenditure,
            forecast_profit,
            forecast_profit_margin
        ]
    }

    summary_df = pd.DataFrame(summary_data)


    # ----------------------------------------
    # 4. PROFIT ANALYSIS
    # ----------------------------------------

    profit_data = {
        "Metric": [
            "Current Profit",
            "Current Profit Margin (%)",
            "Expenditure Ratio (%)",
            "Forecast Profit",
            "Forecast Profit Margin (%)"
        ],
        "Value": [
            profit,
            profit_margin,
            expenditure_ratio,
            forecast_profit,
            forecast_profit_margin
        ]
    }

    profit_df = pd.DataFrame(profit_data)


    # ----------------------------------------
    # 5. BUSINESS INSIGHTS
    # ----------------------------------------

    if profit > 0:
        business_status = "Profitable"
    elif profit == 0:
        business_status = "Break-even"
    else:
        business_status = "Loss"


    if forecast_profit > 0:
        forecast_status = (
            "Expected to remain profitable"
        )
    elif forecast_profit == 0:
        forecast_status = (
            "Expected to break even"
        )
    else:
        forecast_status = (
            "Expected loss"
        )


    if total_revenue > total_expenditure:
        revenue_expense_status = (
            "Revenue is higher than expenditure"
        )
    elif total_revenue == total_expenditure:
        revenue_expense_status = (
            "Revenue and expenditure are equal"
        )
    else:
        revenue_expense_status = (
            "Expenditure is higher than revenue"
        )


    if forecast_profit > 0 and forecast_profit_margin > 20:

        recommendation = (
            "The company is expected to remain "
            "profitable. Revenue performance is "
            "healthy compared with expenditure."
        )

    elif forecast_profit > 0:

        recommendation = (
            "The company is expected to remain "
            "profitable, but expenditure should "
            "be monitored closely."
        )

    else:

        recommendation = (
            "The forecast indicates a potential "
            "loss. Management should review "
            "revenue and expenditure."
        )


    insights_data = {
        "Business Metric": [
            "Current Business Status",
            "Forecast Business Status",
            "Revenue vs Expenditure",
            "Current Profit Margin",
            "Forecast Profit Margin",
            "Current Revenue",
            "Forecast Revenue",
            "Current Expenditure",
            "Forecast Expenditure",
            "Current Profit",
            "Forecast Profit",
            "Business Recommendation"
        ],
        "Details": [
            business_status,
            forecast_status,
            revenue_expense_status,
            f"{profit_margin:.2f}%",
            f"{forecast_profit_margin:.2f}%",
            f"₹{total_revenue:,.2f}",
            f"₹{forecast_revenue:,.2f}",
            f"₹{total_expenditure:,.2f}",
            f"₹{forecast_expenditure:,.2f}",
            f"₹{profit:,.2f}",
            f"₹{forecast_profit:,.2f}",
            recommendation
        ]
    }

    insights_df = pd.DataFrame(
        insights_data
    )


    # ----------------------------------------
    # 6. DAILY DATA
    # ----------------------------------------

    daily_df = df[
        ["Date", "Revenue", "Expenditure"]
    ].copy()


    # ----------------------------------------
    # 7. FORECAST SHEET
    # ----------------------------------------

    forecast_data = {
        "Forecast Metric": [
            "Historical Data Days",
            "Days in Forecast Month",
            "Average Daily Revenue",
            "Forecast Monthly Revenue",
            "Average Daily Expenditure",
            "Forecast Monthly Expenditure",
            "Forecast Monthly Profit",
            "Forecast Profit Margin (%)"
        ],
        "Value": [
            days_of_data,
            days_in_current_month,
            average_daily_revenue,
            forecast_revenue,
            average_daily_expenditure,
            forecast_expenditure,
            forecast_profit,
            forecast_profit_margin
        ]
    }

    forecast_df = pd.DataFrame(
        forecast_data
    )


    # ----------------------------------------
    # 8. CREATE EXCEL FILE
    # ----------------------------------------

    output_file = (
        "reports/company_financial_report.xlsx"
    )

    with pd.ExcelWriter(
        output_file,
        engine="openpyxl"
    ) as writer:

        summary_df.to_excel(
            writer,
            sheet_name="Financial Summary",
            index=False
        )

        profit_df.to_excel(
            writer,
            sheet_name="Profit Analysis",
            index=False
        )

        insights_df.to_excel(
            writer,
            sheet_name="Business Insights",
            index=False
        )

        daily_df.to_excel(
            writer,
            sheet_name="Daily Data",
            index=False
        )

        forecast_df.to_excel(
            writer,
            sheet_name="Forecast",
            index=False
        )


    # ----------------------------------------
    # 9. LOAD WORKBOOK
    # ----------------------------------------

    workbook = load_workbook(
        output_file
    )


    # ----------------------------------------
    # 10. FORMAT WORKSHEETS
    # ----------------------------------------

    for worksheet in workbook.worksheets:

        for cell in worksheet[1]:

            cell.font = Font(
                bold=True
            )

            cell.alignment = Alignment(
                horizontal="center"
            )


        for column in worksheet.columns:

            max_length = 0

            column_letter = get_column_letter(
                column[0].column
            )

            for cell in column:

                if cell.value is not None:

                    max_length = max(
                        max_length,
                        len(str(cell.value))
                    )

            worksheet.column_dimensions[
                column_letter
            ].width = max_length + 5


    # ----------------------------------------
    # 11. REVENUE VS EXPENDITURE CHART
    # ----------------------------------------

    daily_sheet = workbook[
        "Daily Data"
    ]

    bar_chart = BarChart()

    bar_chart.title = (
        "Daily Revenue vs Expenditure"
    )

    bar_chart.y_axis.title = "Amount"

    bar_chart.x_axis.title = "Date"


    data = Reference(
        daily_sheet,
        min_col=2,
        max_col=3,
        min_row=1,
        max_row=daily_sheet.max_row
    )

    categories = Reference(
        daily_sheet,
        min_col=1,
        min_row=2,
        max_row=daily_sheet.max_row
    )


    bar_chart.add_data(
        data,
        titles_from_data=True
    )

    bar_chart.set_categories(
        categories
    )

    bar_chart.height = 10

    bar_chart.width = 18


    daily_sheet.add_chart(
        bar_chart,
        "E2"
    )


    # ----------------------------------------
    # 12. DAILY REVENUE TREND
    # ----------------------------------------

    line_chart = LineChart()

    line_chart.title = (
        "Daily Revenue Trend"
    )

    line_chart.y_axis.title = "Revenue"

    line_chart.x_axis.title = "Date"


    revenue_data = Reference(
        daily_sheet,
        min_col=2,
        max_col=2,
        min_row=1,
        max_row=daily_sheet.max_row
    )


    line_chart.add_data(
        revenue_data,
        titles_from_data=True
    )

    line_chart.set_categories(
        categories
    )

    line_chart.height = 10

    line_chart.width = 18


    daily_sheet.add_chart(
        line_chart,
        "E20"
    )


    # ----------------------------------------
    # 13. FORECAST CHART
    # ----------------------------------------

    forecast_sheet = workbook[
        "Forecast"
    ]

    forecast_chart = BarChart()

    forecast_chart.title = (
        "Forecast Revenue vs Expenditure"
    )

    forecast_chart.y_axis.title = "Amount"

    forecast_chart.x_axis.title = "Forecast"


    # Revenue row and expenditure row

    chart_data = Reference(
        forecast_sheet,
        min_col=2,
        min_row=4,
        max_row=6
    )

    forecast_chart.add_data(
        chart_data,
        titles_from_data=False
    )

    forecast_chart.height = 10

    forecast_chart.width = 18


    forecast_sheet.add_chart(
        forecast_chart,
        "D2"
    )


    # ----------------------------------------
    # 14. SAVE REPORT
    # ----------------------------------------

    workbook.save(
        output_file
    )


    print(
        "\nExcel report generated successfully!"
    )

    print(
        f"Report location: {output_file}"
    )