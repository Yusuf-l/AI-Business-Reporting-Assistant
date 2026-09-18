# AI Business Reporting Assistant

An AI-powered business reporting and analytics dashboard that transforms Excel-based business data into interactive insights, data-quality reports, visual analytics, and AI-generated executive recommendations.

## Overview

AI Business Reporting Assistant is a Python-based business analytics application designed to help users quickly understand and report business performance from Excel datasets.

The application automatically validates and standardizes uploaded Excel files, calculates key business metrics, generates interactive visualizations, and uses Google's Gemini API to produce AI-powered business insights.

## Key Features

* Excel file upload and automatic column mapping
* Data validation and data-quality reporting
* Missing-value detection
* Duplicate-row detection
* Invalid-date detection
* Numeric data validation
* Interactive business dashboard
* Department-level sales analysis
* Product-level sales analysis
* Monthly sales and profit analysis
* Inventory analysis
* AI-generated business insights using Gemini
* Key insights, potential risks, possible explanations, and recommended actions
* Excel report export
* KPI Summary sheet
* Monthly Analysis sheet
* Inventory Analysis sheet
* AI Business Insights sheet
* Automatic Excel charts

## AI Capabilities

The application uses Gemini to analyze verified business metrics generated from the uploaded dataset.

The AI analysis is structured into:

1. Key Insights
2. Potential Risks
3. Possible Explanations
4. Recommended Actions

The AI is instructed to base its analysis strictly on the calculated business metrics and to distinguish potential explanations from verified facts.

## Data Processing Pipeline

```text
Excel Upload
      ↓
Automatic Column Mapping
      ↓
Data Validation
      ↓
Data Quality Analysis
      ↓
Business Metric Calculation
      ↓
Interactive Dashboard
      ↓
Gemini AI Analysis
      ↓
Excel Report Export
```

## Technology Stack

### Programming

* Python

### Data Analysis

* Pandas
* OpenPyXL

### Dashboard

* Streamlit
* Plotly

### Artificial Intelligence

* Google Gemini API

### Reporting

* Excel export with OpenPyXL
* Automated Excel charts

## Project Structure

```text
AI-Business-Reporting-Assistant/
│
├── app.py
├── generate_data.py
├── test_gemini.py
├── requirements.txt
├── README.md
│
├── data/
│   └── sales_data.xlsx
│
└── .gitignore
```

## Installation

Clone the repository:

```bash
git clone https://github.com/YOUR_USERNAME/AI-Business-Reporting-Assistant.git
```

Enter the project directory:

```bash
cd AI-Business-Reporting-Assistant
```

Create a virtual environment:

```bash
python -m venv .venv
```

Activate the virtual environment on Windows:

```bash
.venv\Scripts\activate
```

Install dependencies:

```bash
pip install -r requirements.txt
```

## Gemini API Configuration

Create a `.env` file in the project root:

```text
GEMINI_API_KEY=YOUR_API_KEY
```

Do not commit your `.env` file to GitHub.

The `.gitignore` file already excludes `.env` from version control.

## Running the Application

Start the Streamlit application:

```bash
streamlit run app.py
```

The application will open in your browser.

## Example Workflow

1. Upload an Excel business dataset.
2. The application automatically maps supported column names to the standard format.
3. Data quality and validation checks are performed.
4. Business KPIs and visualizations are generated.
5. Gemini analyzes the verified business metrics.
6. The user can download a structured Excel business report.

## Example Report Structure

The generated Excel report contains:

```text
KPI Summary
Dataset
Department Analysis
Product Analysis
Monthly Analysis
Inventory Analysis
AI Business Insights
```

## Future Improvements

Possible future improvements include:

* Automated PDF report generation
* More advanced forecasting
* Natural-language data querying
* Additional business KPIs
* Role-based dashboards
* Automated scheduled reporting
* Integration with enterprise data sources

## Author

Yusuf-l

Electrical & Electronics Engineering Student

Interested in embedded systems, electronics hardware design, AI-assisted analytics, and R&D.

## License

This project is intended as a personal portfolio and educational project.
