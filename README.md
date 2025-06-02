# 📊 Excel Sales Dashboard Project (2014–2017)

## 🧾 Introduction
This project was developed to build an interactive sales dashboard using Microsoft Excel, aimed at analyzing sales data from 2014 to 2017 within the United States market. The dashboard provides detailed insights into sales, profit, shipping, and customer trends, leveraging advanced Excel features without relying on external tools.

### Tools and Techniques Used
- **Power Query**: For data extraction, transformation, and loading (ETL).
- **Power Pivot**: To create a data model and establish relationships between tables.
- **Measures (DAX)**: For writing advanced calculations within the data model.
- **Pivot Tables**: To summarize and display analysis results.
- **Pivot Charts**: To create interactive charts based on Pivot Tables.
- **Slicers**: To provide interactive filtering within the dashboard.

---

## 📁 Contents of the Excel File
When you open the Excel file, you will find the following worksheets (Sheet Tabs) arranged as shown:

| Sheet Name               | Description                                                                                      |
|--------------------------|--------------------------------------------------------------------------------------------------|
| **Orders**               | The raw transaction data table containing details of each sale.                                   |
| **People**               | A lookup table linking Sales Reps to their Regions.                                              |
| **Return**               | A list of returned orders with return details (Order ID, return reason, etc.).                   |
| **Shipping_Cost**        | A lookup table containing shipping cost for each U.S. state.                                     |
| **Pivot Calculations**   | Pre-built Pivot Tables used to calculate core metrics (KPIs) and intermediary summaries.         |
| **Sales Dashboard Calculation** | Measures and advanced DAX calculations within Power Pivot (e.g., Net Sales, Total Profit). |
| **Dashboard**            | The final visual dashboard that displays KPIs, charts, and interactive elements.                |

---

## 🧭 Detailed Steps to Build the Dashboard

### 1. Importing and Cleaning Data with Power Query
1. Open Excel and navigate to **Data → Get Data → From Other Sources → From Workbook** to import each worksheet (Orders, People, Return, Shipping_Cost).
2. For each table:
   - Click **Transform Data** to open the Power Query Editor.
   - Ensure proper data types:
     - **Order Date** column should be set as *Date*.
     - Numeric columns (e.g., Sales, Profit, Discount, etc.) should be set to *Decimal Number* or *Whole Number* as appropriate.
   - Remove any unnecessary columns (e.g., empty or temporary columns).
   - Filter out invalid or duplicate rows if they exist.
   - In the **Return** table, ensure the **Order ID** column is present for linking back to the original orders.
   - In the **Shipping_Cost** table, ensure there is a **State** column and a **Shipping Cost** column formatted as numeric.
3. After cleaning and transforming each table, click **Close & Load To → Only Create Connection**. This will add the tables to the Data Model without displaying them in separate sheets.

> **Note**: Power Query’s role here is to prepare, clean, and transform the data before analysis in Power Pivot. This ensures accuracy and consistency in the data model.

---

### 2. Building the Data Model with Power Pivot
1. After loading connections for all tables, go to **Power Pivot → Manage** to open the Power Pivot window.
2. In the Power Pivot window:
   - You should see the tables: **Orders**, **People**, **Return**, and **Shipping_Cost**.
   - Create relationships between tables as follows:
     - **Orders[Sales Rep]**  ⟷  **People[Sales Rep]**  
       Links each sale to its respective Sales Rep and Region.
     - **Orders[State]**  ⟷  **Shipping_Cost[State]**  
       Links each sale to the corresponding state’s shipping cost.
     - **Orders[Order ID]**  ⟷  **Return[Order ID]**  
       Connects returned orders with the original order records.
3. Ensure all join keys share the same data type (Text or Number) and formatting to avoid relationship errors.
4. Once relationships are set, switch to **Diagram View** to see a visual representation of the entity relationships.

---

### 3. Creating Measures (DAX Calculations) for KPIs
In the Power Pivot window, create Measures to calculate core performance metrics:

1. **Total Sales**  
   ```DAX
   Total Sales := SUM(Orders[Sales])
   ```

2. **Total Returns**  
   If the **Return** table has a `Return Amount` column:  
   ```DAX
   Total Returns := SUM(Return[Return Amount])
   ```  
   Otherwise, to derive returns from the original sales amount:  
   ```DAX
   Total Returns := CALCULATE(
       SUM(Orders[Sales]),
       TREATAS(VALUES(Return[Order ID]), Orders[Order ID])
   )
   ```

3. **Net Sales**  
   ```DAX
   Net Sales := [Total Sales] - [Total Returns]
   ```

4. **Total Discount**  
   ```DAX
   Total Discount := SUM(Orders[Discount])
   ```

5. **Total COGS (Cost of Goods Sold)**  
   ```DAX
   Total COGS := SUM(Orders[COGS])
   ```

6. **Total Profit**  
   If a `Profit` column exists in **Orders**:  
   ```DAX
   Total Profit := SUM(Orders[Profit])
   ```  
   Otherwise:  
   ```DAX
   Total Profit := [Total Sales] - [Total COGS] - [Total Discount]
   ```

7. **Distinct Customers**  
   ```DAX
   Distinct Customers := DISTINCTCOUNT(Orders[Customer ID])
   ```

8. **Total Orders**  
   ```DAX
   Total Orders := DISTINCTCOUNT(Orders[Order ID])
   ```

9. **Average Shipping Cost**  
   ```DAX
   Average Shipping Cost := AVERAGE(Shipping_Cost[Shipping Cost])
   ```

10. **Return Rate Percentage**  
    ```DAX
    Return Rate % := DIVIDE([Total Returns], [Total Orders], 0)
    ```

> **Note**: Additional Measures like Profit Margin (Profit / Sales), year-over-year growth, or custom calculations can be created as needed.

---

### 4. Setting Up Pivot Tables
With Measures created, build Pivot Tables to analyze data from different perspectives. It is recommended to create a base Pivot Table in the **Pivot Calculations** sheet, then copy and customize for each analysis scenario:

1. **Base Pivot Table**  
   - Go to **Pivot Calculations** sheet.
   - Click **Insert → PivotTable → Use this workbook’s Data Model**.
   - In **PivotTable Fields**, drag the following Measures into **Values**:
     - Total Sales
     - Total Returns
     - Net Sales
     - Total Discount
     - Total COGS
     - Total Profit
     - Distinct Customers
     - Total Orders
   - Leave Rows and Columns empty to display overall totals for all years and regions.

2. **Sales by Category**  
   - Copy the base Pivot Table.
   - Drag `Category` (from **Orders**) into **Rows**.
   - The table will now show each Measure broken down by Category.

3. **Sales by Sales Rep and Region**  
   - Copy the base Pivot Table.
   - Drag `Sales Rep` into **Rows**.
   - Drag `Region` (from **People**) into **Columns**.
   - Keep Measures like Total Sales, Net Sales, and Total Profit in **Values**.

4. **Sales by State and City**  
   - Insert a new Pivot Table.
   - Drag `State` into **Rows**.
   - Drag `City` (below State) into **Rows** for more detailed breakdown.
   - Include Measures such as Total Sales and Net Sales in **Values**.
   - To highlight top 10 states, use **Value Filters → Top 10** on the Total Sales field.

5. **Returns and Discounts Analysis**  
   - Create a Pivot Table focusing on returned orders.
   - Drag `Return Status` (if available in **Orders**) or fields from **Return** into **Rows**.
   - Add Measures like Total Returns and Return Rate % into **Values**.

> **Tip**: Rename each Pivot Table in the **PivotTable Name** field (e.g., “PT_ByCategory”, “PT_ByRep_Region”) for clarity when referencing them later in the dashboard.

---

### 5. Creating Pivot Charts
Once Pivot Tables are in place, generate Pivot Charts to visualize the data:

1. **Category Sales Chart**  
   - Select the “PT_ByCategory” Pivot Table.
   - Go to **Insert → PivotChart**.
   - Choose a **Column Chart** or **Pie Chart** to display sales distribution by Category.
   - Format the chart (remove gridlines, adjust axis titles, etc.).

2. **Sales by Rep and Region Chart**  
   - Select the “PT_ByRep_Region” Pivot Table.
   - Insert a **Stacked Bar Chart** or **Clustered Column Chart**.
   - This will show each Sales Rep’s contribution across Regions.

3. **Top States/Cities Chart**  
   - Use the Pivot Table for State/City (e.g., “PT_ByStateCity”).
   - Insert a **Bar Chart** or **Treemap** to highlight top 10 states or cities.

4. **Returns Analysis Chart**  
   - Select the Pivot Table focused on returns.
   - Insert a **Stacked Column Chart** to show return counts or amounts by Category or State.

> **Tip**: After adding each PivotChart, you can move or resize the underlying Pivot Table to a hidden area in the **Dashboard** sheet, keeping only the chart visible for a cleaner layout.

---

### 6. Building the Dashboard and Adding Slicers
In the **Dashboard** sheet, assemble charts and KPIs into a coherent layout:

1. **Main Title**  
   - At the top of the sheet, add a prominent title “Sales Dashboard (2014–2017)” using a large font size (e.g., 20–24 pt, Bold).

2. **KPI Tiles at the Top**  
   - Insert **Text Boxes** or use cell formatting to display key metrics prominently:
     - **Total Sales**  
     - **Total Returns**  
     - **Net Sales**  
     - **Total Profit**  
     - **Distinct Customers**  
     - **Total Orders**  
   - Each tile should show a large number with a short label (e.g., “Total Sales”).

3. **Charts Section**  
   - Arrange PivotCharts in a 2- or 3-column grid:
     - **Left Column**:  
       - Category Sales Chart  
       - Top 10 Categories by Sales (use PivotTable filter)
     - **Middle Column**:  
       - Sales by Rep and Region Chart  
       - Sales by State/City Chart (Top States/Cities)
     - **Right Column**:  
       - Returns Analysis Chart  
       - Discounts vs. Profit Margin Analysis Chart
   - Resize each chart to ensure readability and alignment.

4. **Interactive Slicers Section**  
   - Insert Slicers (**Insert → Slicer**) for the following fields:
     - `Category`  
     - `Sales Rep`  
     - `State`  
     - `Return Status` (if available)
   - Position slicers at the top or side of the Dashboard.
   - Configure each Slicer for multi-select to allow filtering on multiple items.
   - Link each Slicer to all relevant PivotTables and PivotCharts (Right-click Slicer → Report Connections).

5. **Formatting Tips**  
   - Use a consistent, subtle color palette (e.g., shades of blue and gray).
   - Apply clear titles above each chart (12–14 pt, Bold).
   - Remove gridlines from the Dashboard sheet to maintain a clean look.
   - Apply light borders around each chart and KPI tile to visually separate sections.

---

## 🔄 Refreshing the Dashboard
1. When adding new data (e.g., year 2018), ensure the new records are appended to the original **Orders** table with the same column structure.
2. Click anywhere on a PivotTable and press **Ctrl + Alt + F5** to refresh all PivotTables, PivotCharts, and Slicers at once.
3. Ensure that Power Query connections are updated if new source files or tables are used.

---

## 📝 File Naming Recommendations
- Rename the Excel file from `Book1.xlsx` to a more descriptive name, such as:  
  - `USA_Sales_Dashboard_2014-2017.xlsx`  
  - `Excel_PowerPivot_Sales_Report_USA.xlsx`  
  - `Sales_Dashboard_US.xlsx`
- When updating data for additional years, save a new version with updated years (e.g., `USA_Sales_Dashboard_2014-2018.xlsx`).

---

## 🎯 Summary of Available Metrics and Analyses
- **Total Sales**  
- **Total Returns**  
- **Net Sales**  
- **Total Profit**  
- **Total Discount**  
- **Total COGS (Cost of Goods Sold)**  
- **Distinct Customers**  
- **Total Orders**  
- **Return Rate %**  
- **Average Shipping Cost**  

### Available Analyses
- Sales distribution and profitability by **Category**.  
- Sales performance of **Sales Reps** by **Region**.  
- Top-performing **States / Cities** by Sales.  
- Analysis of **Returns** to identify high-return categories or regions.  
- Comparison of **Discounts** against **Profit Margins**.  
- Tracking **Average Shipping Cost** impact on profitability by state.

---

## 🧠 Final Notes
1. This project leverages only Excel’s built-in capabilities (Power Query, Power Pivot, Pivot Tables, Pivot Charts, Slicers, and DAX) without external software.
2. To extend for new data (future years or additional regions), maintain consistent data structure in source tables and reuse existing queries and relationships.
3. Dashboard design follows best practices for clarity, interactivity, and scalability.
4. Additional Measures or Pivot Tables can be added to address specific analysis requirements, such as Sub-Category analysis or product-level insights.

---

> **Important**: This README is ready to be uploaded alongside your Excel file. It provides a comprehensive guide for users, developers, or stakeholders to understand the structure, functionality, and update procedures for the Sales Dashboard project. If you need any further clarification or enhancements, feel free to update this document.

