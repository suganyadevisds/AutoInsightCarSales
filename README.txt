# AutoInsight — Car Sales SmartBot

A styled Streamlit front end for the India car sales Excel dataset.

## Run on Windows
1. Extract the ZIP file.
2. Open the `SmartBot_Car_Sales_Designed` folder in VS Code.
3. Open Terminal and run:

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Keep `India_Car_Sales_150x16.xlsx` in the same folder as `app.py`.

## Features
- Dark navy sidebar and blue dashboard header
- KPI cards for sales, units, rating and discount
- Charts by manufacturer, fuel type, region and segment
- Sidebar filters
- Rule-based SmartBot questions
- Data table and CSV download

Note: The chatbot uses simple programmed rules, not a generative AI model.
