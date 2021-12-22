import streamlit as st
import pandas as pd
import yfinance as yf
from datetime import datetime
from docx import Document
import base64
from download_button_function import download_button
import io
from create_word_doc import create_doc
import matplotlib.pyplot as plt
import seaborn as sns

# To run use streamlit run template.py
# to exit, use control-c
st.sidebar.subheader("""B.R.O.O.D. Beleggers Template""")
user_input_ticker = st.sidebar.text_input("Geef ticker van het aandeel op","AAPL")
author = st.sidebar.text_input("Uw naam", "Ralph Schuurman")
user_input_competitors = st.sidebar.text_input("Geef Namen/Tickers op van aandelen concurrenten, scheidt aandelen d.m.v. een komma:", "")
# Date = current date
date = datetime.today()



def main():

    ticker = yf.Ticker(user_input_ticker)
    #company name
    companyName = 'Name: ' + ticker.info['shortName']
    shortcompanyName = ticker.info['shortName']
    # Sector
    sector = 'Sector: ' + ticker.info['sector']
    # Industry
    industry = 'Industry: ' + ticker.info['industry']
    # Current price
    current_price = 'Current_price: ' + str(ticker.info['regularMarketPrice'])
    # 52 weeks high and low
    fiftyTwoWeek = "52 weken l/h: " + str(ticker.info['fiftyTwoWeekLow']) + ' - ' + str(ticker.info['fiftyTwoWeekHigh'])
    # 1 year target
    targetMeanPrice = "1 Year Target: " + str(ticker.info['targetMeanPrice'])
    # Market cap
    marketCap = "Market Cap: " + str(ticker.info['marketCap'])
    # Beta
    beta = "Beta: " + str(ticker.info['beta'])

    dividendRate = "Forward Dividend: " + str(ticker.info['trailingAnnualDividendYield']) + " (" + str(ticker.info[
        'trailingAnnualDividendRate']) + ")"

    # Company info + company logo
    companyInfo = ticker.info['longBusinessSummary']  # company info
    companyLogo = ticker.info['logo_url']  # company logo

    st.write(companyName)
    st.write(sector)
    st.write(industry)
    st.write(current_price)
    st.write(fiftyTwoWeek)
    st.write(targetMeanPrice)
    st.write(marketCap)
    st.write(beta)
    st.write(dividendRate)
    st.write(companyInfo)


    user_list = [user_input_ticker]
    competitorlist = user_input_competitors.split(",")
    comparingList = user_list + competitorlist

    if user_input_competitors == "":
        comparingList = user_list

    # create Dataframe with columns
    compare_df = pd.DataFrame(
        columns=["Company name", "D/E", "Current Ratio", " Trailing PE", "Return on Equity", "Profit Margins",
                 " Trailing Annual Dividend Yield", "enterpriseToEbitda"])

    for company in comparingList:
        ticker = yf.Ticker(company)
        dataList = [ticker.info['shortName'], ticker.info['debtToEquity'], ticker.info['currentRatio'],
                    ticker.info['trailingPE'], ticker.info['returnOnEquity'], ticker.info['profitMargins'],
                    ticker.info['trailingAnnualDividendYield'], ticker.info['enterpriseToEbitda']]
        dataList = pd.Series(dataList, index=compare_df.columns)
        compare_df = compare_df.append(dataList, ignore_index=True)


    # Graph of stock for 2 years ( year input by user)

    ## Short positions
    # Short Ratio
    shortRatio = "Short Ratio: " + str(ticker.info['shortRatio'])
    # Short % of shares outstanding
    shortPercentage = "Short % of Shares Outstanding: " + str(ticker.info['shortPercentOfFloat'])
    # Dividend history
    dividendHistory = ticker.dividends
    dividend_df = dividendHistory.to_frame()

    dividend_df = dividend_df.sort_values(by='Date', ascending=False).head(10)

    # News regarding company
    #news = ticker.news
    news = "no news"

    # company logo
    st.dataframe(compare_df)

    # Graph
    graph_data = ticker.history(period = '2y',interval = '1d' )
    st.line_chart(graph_data.Close) #for on the webpage

    #graph to use in document
    memfile = io.BytesIO()
    sns.lineplot(data = graph_data.Close)
    plt.xticks(rotation=30)
    plt.savefig(memfile)

    st.write(shortRatio)
    st.write(shortPercentage)
    st.dataframe(dividend_df)

    document = create_doc(companyName, sector, industry, current_price,
               fiftyTwoWeek, targetMeanPrice,
               marketCap, beta, dividendRate, companyInfo, companyLogo,
               shortRatio, shortPercentage, dividend_df,
               news, compare_df,memfile, author)

    file_stream = io.BytesIO()
    # Save the .docx to the buffer
    document.save(file_stream)
    # Reset the buffer's file-pointer to the beginning of the file
    file_stream.seek(0)
    # convert doc to b64
    b64 = base64.b64encode(file_stream.getvalue()).decode()
    filename = "Voorstel " + shortcompanyName + ".docx"
    download_button_str = download_button(b64, filename, f'Click here to download {filename}')
    st.markdown(download_button_str, unsafe_allow_html=True)


if st.sidebar.button('GO'):
    main()

#if st.button('test'):
    #main()

#main()

#if __name__ == "__main__":
   # main()



