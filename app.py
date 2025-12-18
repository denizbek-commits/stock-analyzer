import tempfile
import json
import os
import time
from flask import Flask, render_template, request, send_file, redirect, url_for, session
import yfinance as yf
import finnhub
from docx import Document
import io

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'your_secret_key_change_in_production')

# Finnhub client setup
finnhub_client = finnhub.Client(api_key=os.environ.get('FINNHUB_API_KEY', 'cs97jkhr01qoa9gbio60cs97jkhr01qoa9gbio6g'))
# Function to get stock data from yfinance
def get_stock_data(ticker):
    try:
        stock = yf.Ticker(ticker)
        info = stock.info
        return info
    except Exception as e:
        print(f"Error fetching stock data for {ticker}: {e}")
        return {}

# Function to fetch forward PE for the next year
def get_forward_PE(ticker):
    try:
        stock = yf.Ticker(ticker)
        forward_pe = stock.info.get('forwardPE', None)
        return forward_pe
    except Exception as e:
        print(f"Error fetching forward PE for {ticker}: {e}")
        return None

# Function to fetch price target data (mean, high, low)
def get_price_target_data(ticker):
    try:
        stock = yf.Ticker(ticker)
        mean_price = stock.info.get('targetMeanPrice', None)
        high_price = stock.info.get('targetHighPrice', None)
        low_price = stock.info.get('targetLowPrice', None)
        return mean_price, high_price, low_price
    except Exception as e:
        print(f"Error fetching price targets for {ticker}: {e}")
        return None, None, None

# Function to fetch Buy ratings from Finnhub
def get_analyst_ratings_finnhub(ticker):
    try:
        recommendations = finnhub_client.recommendation_trends(ticker)
        if not recommendations:
            return None, 0
        
        latest_recommendation = recommendations[0]
        total_ratings = sum([
            latest_recommendation.get('buy', 0),
            latest_recommendation.get('hold', 0),
            latest_recommendation.get('sell', 0),
            latest_recommendation.get('strongBuy', 0),
            latest_recommendation.get('strongSell', 0)
        ])
        
        buy_ratings = latest_recommendation.get('buy', 0) + latest_recommendation.get('strongBuy', 0)
        buy_percentage = (buy_ratings / total_ratings) * 100 if total_ratings > 0 else 0
        return buy_percentage, total_ratings
    except Exception as e:
        print(f"Error fetching analyst ratings for {ticker}: {e}")
        return None, 0

# Function to fetch ownership data from Yahoo Finance
def get_ownership_data(ticker):
    try:
        stock = yf.Ticker(ticker)
        insiders = stock.info.get('heldPercentInsiders', None)
        institutions = stock.info.get('heldPercentInstitutions', None)
        
        # Return None if either component is missing
        if insiders is None or institutions is None:
            insider_ownership = insiders * 100 if insiders is not None else None
            institutional_ownership = institutions * 100 if institutions is not None else None
            return None, insider_ownership, institutional_ownership
        
        # Convert ownership percentages to whole numbers
        insider_ownership = insiders * 100
        institutional_ownership = institutions * 100
        total_ownership = insider_ownership + institutional_ownership
        
        return total_ownership, insider_ownership, institutional_ownership
    except Exception as e:
        print(f"Error fetching ownership data for {ticker}: {e}")
        return None, None, None

# Function to check all conditions at once for the stock
def check_all_conditions(ticker, benchmark_forward_PE):
    info = get_stock_data(ticker)
    ticker_results = [ticker]
    details = [ticker]  # Start with ticker as first element
    conditions_met = True
    
    # 1. Analyst Ratings Check
    buy_percentage, buy_count = get_analyst_ratings_finnhub(ticker)
    if buy_percentage is None:
        ticker_results.append('❌ (Ratings Unavailable)')
        details.append(f"Buy Ratings: Data unavailable")
        conditions_met = False
    else:
        details.append(f"Buy Ratings: {buy_count} analysts, Buy Percentage: {buy_percentage:.2f}%")
        if buy_percentage < 70:
            ticker_results.append('❌')
            conditions_met = False
        else:
            ticker_results.append('✔️')
    
    # 2. Market Cap Check
    market_cap = info.get('marketCap', 0) / 1e9
    details.append(f"Market Cap: ${market_cap:.2f}B")
    if market_cap < 100:
        ticker_results.append('❌')
        conditions_met = False
    else:
        ticker_results.append('✔️')
    
    # 3. Price Target Check
    mean_price, high_price, low_price = get_price_target_data(ticker)
    current_price = info.get('currentPrice', 0)
    details.append(f"Price Targets: Current=${current_price:.2f}, Mean=${mean_price}, High=${high_price}, Low=${low_price}")
    if mean_price is None or mean_price <= current_price * 1.15 or low_price < current_price:
        ticker_results.append('❌')
        conditions_met = False
    else:
        ticker_results.append('✔️')
    
    # 4. Forward PE Check
    forward_PE_next_year = get_forward_PE(ticker)
    details.append(f"Forward PE: {forward_PE_next_year if forward_PE_next_year is not None else 'N/A'} (Benchmark: {benchmark_forward_PE})")
    if forward_PE_next_year is None:
        ticker_results.append('❌ (Missing Forward PE)')
        conditions_met = False
    elif forward_PE_next_year >= benchmark_forward_PE:
        ticker_results.append('❌')
        conditions_met = False
    else:
        ticker_results.append('✔️')
    
    # 5. Ownership Check
    total_ownership, insider_ownership, institutional_ownership = get_ownership_data(ticker)
    if total_ownership is None:
        details.append(f"Ownership: Data unavailable (Insider: {insider_ownership if insider_ownership else 'N/A'}%, Institutional: {institutional_ownership if institutional_ownership else 'N/A'}%)")
        ticker_results.append('❌ (Ownership Data Unavailable)')
        conditions_met = False
    else:
        details.append(f"Ownership: Total {total_ownership:.2f}%, Insider: {insider_ownership:.2f}%, Institutional: {institutional_ownership:.2f}%")
        if total_ownership < 70:
            ticker_results.append('❌')
            conditions_met = False
        else:
            ticker_results.append('✔️')
    
    return ticker_results, conditions_met, details

# Function to generate Word document
def generate_word_report(results, details, buy_tickers):
    doc = Document()
    doc.add_heading('Stock Analysis Report', 0)
    
    for ticker_details, ticker_results in zip(details, results):
        doc.add_heading(f"Results for {ticker_details[0]}", level=1)
        # Skip the first element (ticker) since it's in the heading
        for detail in ticker_details[1:]:
            doc.add_paragraph(detail)
        doc.add_paragraph(f"Condition Check: {' '.join(ticker_results[1:])}")
        doc.add_paragraph()
    
    doc.add_heading('Buy Tickers', level=1)
    doc.add_paragraph(', '.join(buy_tickers) if buy_tickers else "No tickers met all conditions.")
    
    word_file = io.BytesIO()
    doc.save(word_file)
    word_file.seek(0)
    return word_file

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/analyze', methods=['POST'])
def analyze():
    # Clear the session to avoid old data
    session.clear()
    
    tickers_input = request.form.get('tickers', '').strip()
    benchmark_forward_PE_input = request.form.get('benchmark_forward_PE', '').strip()
    
    if not tickers_input or not benchmark_forward_PE_input:
        return "Please provide both tickers and benchmark forward PE.", 400
    
    tickers = [t.strip().upper() for t in tickers_input.split(',') if t.strip()]
    
    try:
        benchmark_forward_PE = float(benchmark_forward_PE_input)
    except ValueError:
        return "Please enter a valid numeric benchmark forward PE.", 400
    
    buy_tickers = []
    results = []
    all_details = []
    
    for ticker in tickers:
        ticker_results, conditions_met, details = check_all_conditions(ticker, benchmark_forward_PE)
        results.append(ticker_results)
        all_details.append(details)
        if conditions_met:
            buy_tickers.append(ticker)
        
        # Rate limiting: 500ms delay between API calls
        time.sleep(0.5)
    
    # Store the results in a temporary file
    temp_file = tempfile.NamedTemporaryFile(delete=False, mode='w', suffix='.json')
    json.dump({
        'results': results,
        'buy_tickers': buy_tickers,
        'all_details': all_details
    }, temp_file)
    temp_file.close()
    
    # Store the path of the temp file in the session
    session['temp_file_path'] = temp_file.name
    
    # Redirect to the results page after storing session data
    return redirect(url_for('results'))

@app.route('/results')
def results():
    # Load the results from the temporary file
    temp_file_path = session.get('temp_file_path')
    if temp_file_path and os.path.exists(temp_file_path):
        try:
            with open(temp_file_path, 'r') as f:
                data = json.load(f)
                buy_tickers = data.get('buy_tickers', [])
                results = data.get('results', [])
                all_details = data.get('all_details', [])
        except Exception as e:
            print(f"Error reading temp file: {e}")
            buy_tickers, results, all_details = [], [], []
    else:
        buy_tickers, results, all_details = [], [], []
    
    return render_template('result.html', buy_tickers=buy_tickers, results=results, all_details=all_details)

@app.route('/download_word')
def download_word():
    # Retrieve results from the temp file
    temp_file_path = session.get('temp_file_path')
    if temp_file_path and os.path.exists(temp_file_path):
        try:
            with open(temp_file_path, 'r') as f:
                data = json.load(f)
                results = data.get('results', [])
                buy_tickers = data.get('buy_tickers', [])
                all_details = data.get('all_details', [])
            
            word_file = generate_word_report(results, all_details, buy_tickers)
            
            # Clean up temp file after generating report
            try:
                os.unlink(temp_file_path)
                session.pop('temp_file_path', None)
            except Exception as e:
                print(f"Error deleting temp file: {e}")
            
            return send_file(
                word_file,
                as_attachment=True,
                download_name='stock_analysis_report.docx',
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
        except Exception as e:
            print(f"Error generating Word report: {e}")
            return "Error generating report", 500
    else:
        return "No analysis results found. Please run analysis first.", 404

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)