# Yahoo Finance Fetcher

## Description

This is a simple python script that fetches the historical data of a stock from Yahoo Finance and saves it in an Excel file. The script uses yfinance library to fetch the data, and pandas library to manipulate the data. The script also uses XlsxWriter library to write the data to an Excel file. The script is written in Python 3.10.12 but it should work with any Python 3.x version.

## Instructions

### Prerequisites

What things you need to install the software and how to install them:

- Python 3.x
- pandas library
- yfinance library
- XlsxWriter library

You can install the required libraries using pip:

```bash
pip install pandas
pip install yfinance
pip install XlsxWriter
```

### Usage

To use the script, you will first need to make an adjustment to the script You will need to change the value of the variable `tickerDict` (Line 15) to the symbols and purchase price of the stock you want to fetch the data for. For example, if you want to fetch the data for Apple which was bought for $100, you will need to change the value of the variable `tickersDict` to `'AAPL': 100`. You can find the symbol of a stock on Yahoo Finance. For example, the symbol of Apple is `AAPL`, the symbol of Microsoft is `MSFT`, and the symbol of Tesla is `TSLA`.

After you have made the adjustment to the script, you can run the script using the following command:

```bash
python yfinanceScript.py
```

The script will then fetch the data for the stocks you specified and save the data in an Excel file named `spreadsheet.xlsx` in the same directory.

> Note: If you want to change the gain/loss percentage color threshold, you can change the value of the variable `gainLossMargin` (Line 17) to the percentage you want. For example, if you want the gain/loss percentage color threshold to be 10%, you will need to change the value of the variable `gainLossMargin` to `10`. The default value of the variable `gainLossMargin` is `5`.

## Authors

- **[memms](https://github.com/memms)**

## License

This project is licensed under the GPLv3.0 License - see the [LICENSE](LICENSE) file for details.