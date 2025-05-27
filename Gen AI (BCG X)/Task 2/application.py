import pandas as pd
from flask import Flask, render_template, request

application = Flask(__name__, static_folder="static")
app = application

# Read data from Excel file
df = pd.read_excel('Database.xlsx')

@app.route("/")
def home():
    return render_template('home.html')

@app.route("/chat", methods=['GET', 'POST'])
def chat():
    if request.method == "POST":
        company = request.form.get("company")
        metric = request.form.get("metric").strip()
        year = request.form.get("year")

        df.columns = df.columns.str.strip()

        # Fetch data safely
        if metric in df.columns:
            res = df[(df["Company"] == company) & (df["Fiscal Year"] == int(year))][metric]
            if res.empty:
                return render_template('chat.html', result="Oops! Unable to find data.")
            else:
                return render_template('chat.html', result=f"The {metric} of {company} in {year} is {res.values[0]} billion Dollars")
        else:
            return render_template('chat.html', result="Invalid metric selected.")

    return render_template('chat.html')

if __name__ == "__main__":
    app.run(debug=True)