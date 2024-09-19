import os
import pandas as pd
import yfinance as yf
import datetime
from collections import Counter
import re

# File where we store the results
EXCEL_FILE = 'comparable_analysis.xlsx'

# List of stopwords to exclude from keyword extraction
STOPWORDS = set([
    "the", "and", "a", "in", "of", "for", "with", "on", "at", "to", "by", "from", "as",
    "that", "which", "this", "these", "those", "it", "is", "are", "be", "has", "have",
    "was", "were", "been", "being", "its", "an", "or", "but", "also", "about", "under",
    "over", "through", "throughout", "using", "utilizing", "their", "other"
])

# Function to search for stock ticker or company name
def search_stock(stock_input):
    try:
        stock = yf.Ticker(stock_input)
        # Check for other common fields to validate stock data
        if 'shortName' in stock.info:
            print(f"Found stock: {stock.info['shortName']} ({stock_input})")
            confirmation = input(f"Is this the correct stock? (y/n): ").strip().lower()
            if confirmation == 'y':
                return stock_input
        else:
            print(f"Could not retrieve full data for {stock_input}. Please try another ticker.")
    except Exception as e:
        print(f"Error fetching data for {stock_input}: {e}")

    return None

# Extract keywords from company description, excluding stopwords and the company name
def extract_keywords(description, company_name):
    # Clean up the description by removing non-alphabetical characters
    description = re.sub(r'[^a-zA-Z\s]', '', description)
    # Split by spaces and lowercase everything
    words = description.lower().split()

    # Remove stopwords and the company's name from the keyword list
    keywords = [
        word for word in words
        if word not in STOPWORDS and word != company_name.lower() and len(word) > 3
    ]

    return keywords

# Fetch comparable metrics and description
def fetch_metrics_and_description(ticker):
    stock = yf.Ticker(ticker)
    metrics = {
        'Price': stock.info.get('currentPrice', None),
        'Earnings': stock.info.get('netIncomeToCommon', None),
        'Ebitda': stock.info.get('ebitda', None),
        'Revenue': stock.info.get('totalRevenue', None),
        'MarketCap': stock.info.get('marketCap', None),
        'SharesOutstanding': stock.info.get('sharesOutstanding', None),
        'Industry': stock.info.get('industry', None),
        'Description': stock.info.get('longBusinessSummary', None),
        'CompanyName': stock.info.get('shortName', None)
    }

    return metrics

# Function to calculate keyword match percentage
def keyword_match_percentage(target_keywords, peer_keywords):
    if not peer_keywords:  # Handle empty peer description case
        return 0
    # Find intersection of keywords and calculate match percentage
    overlap = set(target_keywords).intersection(peer_keywords)
    return len(overlap) / len(target_keywords) * 100

# Match companies based on keywords and broader market cap tolerance
def find_comparables(target_metrics, peer_companies, tolerance=0.75, keyword_min_threshold=11, keyword_max_threshold=19):
    comparables = []
    target_description = target_metrics['Description']
    target_market_cap = target_metrics['MarketCap']
    company_name = target_metrics['CompanyName']

    # Extract keywords from the target company description, excluding the company's name
    target_keywords = extract_keywords(target_description, company_name)
    print(f"Target Keywords: {target_keywords}")

    for peer in peer_companies:
        peer_metrics = fetch_metrics_and_description(peer)
        peer_description = peer_metrics.get('Description', "")
        peer_company_name = peer_metrics.get('CompanyName', "")

        # Extract keywords from the peer company description, excluding peer's company name
        peer_keywords = extract_keywords(peer_description, peer_company_name)

        # Debug output: Print peer company and its market cap
        print(f"Checking peer {peer}: MarketCap = {peer_metrics['MarketCap']}, Industry = {peer_metrics['Industry']}")

        # Calculate keyword match percentage
        match_percentage = keyword_match_percentage(target_keywords, peer_keywords)
        print(f"Keyword Match Percentage with {peer}: {match_percentage:.2f}%")

        # Match based on keyword overlap percentage and market cap tolerance
        if (peer_metrics['MarketCap'] and
            abs(peer_metrics['MarketCap'] - target_market_cap) / target_market_cap <= tolerance and
            keyword_min_threshold <= match_percentage <= keyword_max_threshold):
            comparables.append(peer)

        # Fallback: Accept if keyword match is between 11-19%, even if market cap doesn't match
        elif keyword_min_threshold <= match_percentage <= keyword_max_threshold:
            print(f"Added {peer} based on keyword match alone (between 11% and 19%).")
            comparables.append(peer)

    return comparables

# Perform Comparable Company Analysis
def comparable_company_analysis(target_ticker, comparables):
    # Fetch target company metrics
    target_metrics = fetch_metrics_and_description(target_ticker)

    # Fetch metrics for comparable companies
    comparable_metrics = {ticker: fetch_metrics_and_description(ticker) for ticker in comparables}

    # Collect relevant multiples (P/E, EV/EBITDA)
    pe_ratios = []
    ev_ebitda_ratios = []

    for ticker, metrics in comparable_metrics.items():
        # Ensure earnings are available and positive
        if metrics['Price'] and metrics['Earnings'] and metrics['Earnings'] > 0:
            pe_ratios.append(metrics['Price'] / metrics['Earnings'])

        # Ensure EBITDA and market cap are available
        if metrics['Ebitda'] and metrics['Ebitda'] > 0 and metrics['MarketCap']:
            enterprise_value = metrics['MarketCap']  # Simplified EV = Market Cap (ignores debt)
            ev_ebitda_ratios.append(enterprise_value / metrics['Ebitda'])

    # Debug output for comparables
    print(f"Comparable P/E Ratios: {pe_ratios}")
    print(f"Comparable EV/EBITDA Ratios: {ev_ebitda_ratios}")

    # Calculate median multiples
    median_pe = None
    median_ev_ebitda = None

    if pe_ratios:
        median_pe = pd.Series(pe_ratios).median()
    if ev_ebitda_ratios:
        median_ev_ebitda = pd.Series(ev_ebitda_ratios).median()

    # Debug output for calculated medians
    print(f"Median P/E: {median_pe}, Median EV/EBITDA: {median_ev_ebitda}")

    # Estimate fair values
    fair_values = {}
    if target_metrics['Ebitda'] and median_ev_ebitda and target_metrics['SharesOutstanding']:
        # Calculate fair value based on EV/EBITDA
        fair_value_ev_ebitda = median_ev_ebitda * target_metrics['Ebitda']
        fair_values['EV/EBITDA'] = fair_value_ev_ebitda / target_metrics['SharesOutstanding']  # Fair value per share

    # Handle negative earnings for P/E ratio
    if target_metrics['Earnings'] and target_metrics['Earnings'] > 0 and median_pe and target_metrics['SharesOutstanding']:
        # Calculate fair value based on P/E ratio
        fair_value_pe = median_pe * target_metrics['Earnings']
        fair_values['P/E'] = fair_value_pe / target_metrics['SharesOutstanding']  # Fair value per share
    else:
        fair_values['P/E'] = 'N/A (Negative Earnings)'

    return fair_values, target_metrics

# Function to append or update the Excel file
def append_to_excel(data, file_name):
    # Check if the Excel file already exists
    if os.path.exists(file_name):
        # Load the existing data
        df = pd.read_excel(file_name)
    else:
        # Create a new dataframe with the appropriate columns
        df = pd.DataFrame(columns=['Stock', 'Metric', 'Value'])

    # Remove empty entries before concatenating
    new_data = pd.DataFrame(data).dropna(how='all')

    # Append the new data
    df = pd.concat([df, new_data], ignore_index=True)

    # Write the updated data to the Excel file
    df.to_excel(file_name, index=False)
    print(f"Data has been appended to the Excel file.")

# Main function to run the comparable analysis
def main():
    # Target stock input
    stock_input = input("Enter the stock ticker or company name: ").strip()
    target_ticker = search_stock(stock_input)

    # Fetch target company metrics
    target_metrics = fetch_metrics_and_description(target_ticker)

    # For demonstration purposes, using a hardcoded peer companies list
    peer_companies = ['ADM', 'BG', 'INGR', 'LANC', 'HRL']  # Example companies in the food and beverages industry

    # Find comparables
    comparables = find_comparables(target_metrics, peer_companies)

    if not comparables:
        print(f"No suitable comparables found for {target_ticker}.")
        return

    if target_ticker:
        # Perform Comparable Company Analysis
        fair_values, target_metrics = comparable_company_analysis(target_ticker, comparables)

        # Print and save results
        print(f"\nTarget Company ({target_ticker}) Metrics:")
        for metric, value in target_metrics.items():
            print(f"{metric}: {value}")

        print("\nFair Values Based on Comparables:")
        for method, value in fair_values.items():
            print(f"{method}: {value}")

        # Calculate percentage difference between fair value and current price
        if fair_values.get('EV/EBITDA') and target_metrics['Price']:
            ev_ebitda_diff = ((fair_values['EV/EBITDA'] - target_metrics['Price']) / target_metrics['Price']) * 100
            print(f"EV/EBITDA Fair Value Price: {fair_values['EV/EBITDA']:.2f}, Percentage Difference: {ev_ebitda_diff:.2f}%")

        if fair_values.get('P/E') != 'N/A (Negative Earnings)' and target_metrics['Price']:
            pe_diff = ((fair_values['P/E'] - target_metrics['Price']) / target_metrics['Price']) * 100
            print(f"P/E Fair Value Price: {fair_values['P/E']:.2f}, Percentage Difference: {pe_diff:.2f}%")

        # Prepare data for saving
        data_to_save = [
            {'Stock': target_ticker, 'Metric': 'EV/EBITDA Fair Value', 'Value': fair_values.get('EV/EBITDA', 'N/A')},
            {'Stock': target_ticker, 'Metric': 'P/E Fair Value', 'Value': fair_values.get('P/E', 'N/A')},
            {'Stock': target_ticker, 'Metric': 'Current Price', 'Value': target_metrics.get('Price', 'N/A')}
        ]

        # Append to Excel
        append_to_excel(data_to_save, EXCEL_FILE)
    else:
        print("No valid stock selected.")

# Example usage
if __name__ == "__main__":
    main()
