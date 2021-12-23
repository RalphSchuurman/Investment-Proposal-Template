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
import requests
from numerize import numerize

# To run use streamlit run template.py
# to exit, use control-c
st.set_page_config(layout="wide")
st.sidebar.subheader("""B.R.O.O.D. Beleggers Template""")
user_input_ticker = st.sidebar.text_input("Geef ticker van het aandeel op","AAPL")
author = st.sidebar.text_input("Uw naam", "Ralph Schuurman")
user_input_competitors = st.sidebar.text_input("Geef Namen/Tickers op van aandelen concurrenten, scheidt aandelen d.m.v. een komma:", "MSFT,TSLA")
# Date = current date
date = datetime.today()


col1, col2, col3 = st.columns(3)

def main():
    col1.subheader('Bedrijfsinformatie')
    col2.subheader('Concurrentie en koersgrafiek')
    col3.subheader('Dividenduitkeringen')

    data_1_laden = col1.text('Laden...')
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
    marketCap = "Market Cap: " + str(numerize.numerize(ticker.info['marketCap']))
    # Beta
    beta = "Beta: " + str(round(ticker.info['beta'],3))

    dividendRate = "Forward Dividend: " + str(ticker.info['trailingAnnualDividendRate']) + " (" + str(round(ticker.info['trailingAnnualDividendYield']*100,2)) + "%)"

    dividendHistory = ticker.dividends

    # Company info + company logo
    companyInfo = ticker.info['longBusinessSummary']  # company info
    companyLogo = ticker.info['logo_url']  # company logo

    #save company logo to put in docx
    response = requests.get(companyLogo, stream=True)
    imagelogo = io.BytesIO(response.content)


    col1.image(companyLogo)
    col1.write(companyName)
    col1.write(sector)
    col1.write(industry)
    col1.write(current_price)
    col1.write(fiftyTwoWeek)
    col1.write(targetMeanPrice)
    col1.write(marketCap)
    col1.write(beta)
    col1.write(dividendRate)
    col1.write(companyInfo)
    data_1_laden.text('Laden... klaar')
    data_2_laden = col2.text('Laden...')
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


    # News regarding company
    #news = ticker.news
    news = "no news"

    col2.dataframe(compare_df)

    # Graph
    graph_data = ticker.history(period = '2y',interval = '1d' )
    col2.line_chart(graph_data.Close) #for on the webpage

    data_2_laden.text('Laden... klaar')

    #graph to use in document
    memfile = io.BytesIO()
    sns.lineplot(data = graph_data.Close)
    plt.xticks(rotation=30)
    plt.savefig(memfile)




    # Dividend history
    data_3_laden = col3.text('Laden...')
    dividend_df = dividendHistory.to_frame()
    dividend_df.index = dividend_df.index.strftime('%Y-%m-%d')
    dividend_df = dividend_df.sort_values(by='Date', ascending=False).head(10)
    col3.dataframe(dividend_df)

    col3.subheader("Short ratio's")
    col3.write(shortRatio)
    col3.write(shortPercentage)


    data_3_laden.text('Laden... klaar')
    document = create_doc(companyName, sector, industry, current_price,
               fiftyTwoWeek, targetMeanPrice,
               marketCap, beta, dividendRate, companyInfo, companyLogo,
               shortRatio, shortPercentage, dividend_df,
               news, compare_df,memfile, author,imagelogo)

    file_stream = io.BytesIO()
    # Save the .docx to the buffer
    document.save(file_stream)
    # Reset the buffer's file-pointer to the beginning of the file
    file_stream.seek(0)
    # convert doc to b64
    b64 = base64.b64encode(file_stream.getvalue()).decode()
    filename = "Voorstel " + shortcompanyName + ".docx"
    download_button_str = download_button(b64, filename, f'Click here to download {filename}')
    st.sidebar.markdown(download_button_str, unsafe_allow_html=True)


if st.sidebar.button('GO'):
    main()

st.sidebar.write("Als het document klaar is komt hier een download link")
#if st.button('test'):
    #main()

#main()

#if __name__ == "__main__":
   # main()



