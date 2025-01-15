from flask import Flask, jsonify, request, render_template_string
import pandas as pd
from datetime import datetime

app = Flask(__name__)

# Path to the Excel file
EXCEL_FILE = "links.xlsx"

def load_excel():
    """Load the Excel file as a DataFrame."""
    return pd.read_excel(EXCEL_FILE)

def save_excel(df):
    """Save the DataFrame back to the Excel file."""
    df.to_excel(EXCEL_FILE, index=False)

@app.route("/get_promo", methods=["GET"])
def get_promo():
    """Endpoint to get a unique Zomato link with status 'Available'."""
    client_ip = request.remote_addr  # Get the client's IP address
    df = load_excel()

    # Check if the IP has already accessed the endpoint
    accessed_row = df[df["IP"] == client_ip]
    if not accessed_row.empty:
        zomato_link = accessed_row.iloc[0]["Zomato Link"]
        html_template = """
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Previously Accessed Link</title>
        </head>
        <body>
            <h1>You have already accessed a Zomato link:</h1>
            <p><a href="{{ zomato_link }}" target="_blank">{{ zomato_link }}</a></p>
        </body>
        </html>
        """
        return render_template_string(html_template, zomato_link=zomato_link)

    # Get available links
    available_links = df[df["Status"] == "Available"]

    if available_links.empty:
        return "No available links at the moment.", 404

    # Pick the first available link
    link_row = available_links.iloc[0]
    zomato_link = link_row["Zomato Link"]
    link_number = link_row["Link Number"]

    # Update the status to 'Sold', set the date, and save the client's IP
    df.loc[df["Link Number"] == link_number, "Status"] = "Sold"
    df.loc[df["Link Number"] == link_number, "Date Sold"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    df.loc[df["Link Number"] == link_number, "IP"] = client_ip
    save_excel(df)

    # Render the response page
    html_template = """
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Zomato Link</title>
    </head>
    <body>
        <h1>Your Zomato Link is:</h1>
        <p><a href="{{ zomato_link }}" target="_blank">{{ zomato_link }}</a></p>
    </body>
    </html>
    """
    return render_template_string(html_template, zomato_link=zomato_link)

if __name__ == "__main__":
    app.run(debug=True, port=2000)
