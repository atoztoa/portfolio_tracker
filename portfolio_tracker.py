import json
import glob
import urllib2
import re
from bs4 import BeautifulSoup
import codecs
from termcolor import colored

# ------- CONSTANTS --------- #
KEYWORDS = {
            "TATA ELXSI" : "NSE:TATAELXSI",
            "TATA MOTORS(TELCO)" : "NSE:TATAMOTORS",
            "TATA CONSULTANCY SE": "NSE:TCS",
            "CIPLA": "NSE:CIPLA",
            "TATA IRON & STEEL CO": "NSE:TATASTEEL",
            "KAMDHENU ISPAT LTD" : "NSE:KAMDHENU",
            "JAMNA AUTO INDUSTRIE" : "NSE:JAMNAAUTO"
        }

MISC_KEY = "--CHARGES--"

COLUMNS = ['Order No', 'Order Time', 'Trade No.', 'Trade Time', 'Security', 'Bought Qty', 'Sold Qty', 'Gross Rate', 'Gross Total', 'Brokerage', 'Net Rate', 'Service Tax', 'STT', 'Total']
COLUMNS_NEW = ['Order No', 'Order Time', 'Trade No.', 'Trade Time', 'Security', 'Buy/Sell', 'Quantity', 'Gross Rate', 'Brokerage', 'Net Rate', 'Closing Rate', 'Total', 'Remarks']

""" Get transaction data from HTML file
"""
def process_html_file(filename):
    print "Processing file: " + filename + "..."
    html = open(filename).read()

    soup = BeautifulSoup(html)

    table = soup.find("td", class_="xl27boTBL").findParents("table")[0]

    entries = []

    for row in table.findAll("tr"):
        entry = []

        for cell in row.findAll("td"):
            entry.append("".join(c for c in str(unicode(cell.string).encode('ascii', 'ignore')).strip() if c not in "*[]~"))

        # Filter
        if len(entry) > 10 and "".join(entry):
            entries.append(entry)

        # Ignore rest of the entries
        if "NET AMOUNT DUE" in "".join(entry):
            break

    return entries

""" Process transactions from Contract Notes
"""
def process_entries(entries):
    print "Processing transactions..."

    is_data = False
    is_scrip = False
    is_misc = False
    is_new_html_format = False
    item = {}
    items = []
    misc = {}

    # Prune empty entries
    entries = [entry for entry in entries if "".join(entry).strip()]

    # New format?
    head = entries[0]
    if len(head) == 13:
        is_new_html_format = True

    for entry in entries:
        scrap_entries = [ 'ISIN', 'BUY AVERAGE', 'SELL AVERAGE', 'NET AVERAGE', 'Delivery Total' ]

        if any(item in "".join(entry) for item in scrap_entries):
            continue

        # Clean
        entry = [x.strip() for x in entry]

        if entry[0] and entry[0].isdigit():
            is_data = True

        if is_data:
            # New scrip
            if not is_scrip:
                is_scrip = True

                item = {}

                if is_new_html_format:
                    item.update(dict(zip(COLUMNS_NEW, entry)))
                else:
                    item.update(dict(zip(COLUMNS, entry)))

                if is_new_html_format:
                    item['Type'] = "SELL" if item['Buy/Sell'] == 'S' else "BUY"
                else:
                    if 'Sold Qty' in item and item['Sold Qty']:
                        item['Type'] = "SELL"
                        item['Quantity'] = item.pop("Sold Qty")
                    elif 'Bought Qty' in item and item['Bought Qty']:
                        item['Type'] = "BUY"
                        item['Quantity'] = item.pop("Bought Qty")

                # Total is always positive value
                if(float(item['Total']) < 0):
                    item['Total'] = str(abs(float(item['Total'])))

                scrap_keys = COLUMNS[:4]

                if is_new_html_format:
                    scrap_keys += [ 'Buy/Sell', 'Remarks' ]

                item = { key:value for key,value in item.items() if not any(k in key for k in scrap_keys) and value }
            else:
                # Multiple entries
                if entry[0] and entry[0].isdigit():
                    next_item = dict(zip(COLUMNS_NEW, entry))

                    if is_new_html_format:
                        next_item['Type'] = "SELL" if next_item['Buy/Sell'] == 'S' else "BUY"
                    else:
                        if 'Sold Qty' in next_item:
                            next_item['Type'] = "SELL"
                            next_item['Quantity'] = next_item.pop("Sold Qty")

                        if 'Bought Qty' in next_item:
                            next_item['Type'] = "BUY"
                            next_item['Quantity'] = next_item.pop("Bought Qty")

                    # Total is always positive value
                    if(float(next_item['Total']) < 0):
                        next_item['Total'] = str(abs(float(next_item['Total'])))

                    scrap_keys = COLUMNS[:4]

                    if is_new_html_format:
                        scrap_keys += [ 'Buy/Sell', 'Remarks' ]

                    next_item = { key:value for key,value in next_item.items() if not any(k in key for k in scrap_keys) and value }

                    if "Trades" not in item:
                        item = {"Trades": [ item ]}

                    item["Trades"].append(next_item)
                    continue

                col = 11 if is_new_html_format else 12

                if entry[col]:
                    item[entry[4].strip("*").strip()] = entry[col]

                # Finish
                if "TOTAL STT" in entry[4] or not entry[col]:
                    is_scrip = False
                    is_data = False

                    # Cleanup
                    scrap_keys = COLUMNS[:4]
                    scrap_keys += [ 'STT SELL DELIVERY', 'STT BUY DELIVERY' ]

                    if is_new_html_format:
                        scrap_keys += [ 'Buy/Sell', 'Remarks' ]

                    item = { key:value for key,value in item.items() if not any(k in key for k in scrap_keys) and value }

                    if "TOTAL STT" in item:
                        item['STT'] = item.pop("TOTAL STT")

                    items.append(item)

                    if not entry[col]:
                        is_misc = True
        else:
            if items:
                is_misc = True

        if is_misc:
            col = 11 if is_new_html_format else 13
            if not entry[col - 1] or is_new_html_format:
                misc[entry[4].strip("*").strip("[]").strip("~").strip()] = entry[col]

        # Ignore rest of the entries
        if "NET AMOUNT DUE" in "".join(entry):
            break

    scrap_keys = [ 'NET AMOUNT DUE TO', 'DR. TOTAL', 'CR. TOTAL' ]

    # Misc Charges
    misc = {key:value for key,value in misc.items() if not any(k in key for k in scrap_keys)}
    misc['Total'] = sum(float(item) for key,item in misc.items())
    misc['Type'] = "MISC"
    items.append(misc)

    return items

""" Get current market price
"""
def get_quote(symbol):
    print "Getting market price: " + symbol

    base_url = 'http://finance.google.com/finance?q='

    symbol = KEYWORDS[symbol]

    try:
        response = urllib2.urlopen(base_url + symbol)
        html = response.read()
    except Exception, msg:
        print msg
        return 0.0

    soup = BeautifulSoup(html)

    try:
        price = soup.find_all("span", id=re.compile('^ref_.*_l$'))[0].string
        price = str(unicode(price).encode('ascii', 'ignore')).strip().replace(",", "")

        return price
    except Exception as e:
        print "Can't get current rate for scrip: " + symbol
        return 0.0

""" Update portfolio with new data
"""
def update_portfolio(data, portfolio):
    print "Updating portfolio..."

    for item in data:
        if "Trades" in item:
            update_portfolio(item["Trades"], portfolio)
            continue

        if item["Type"] in ["BUY", "SELL"]:
            scrip = item["Security"]

            if scrip not in portfolio:
                portfolio[scrip] = {
                            "Total Quantity": 0,
                            "Total Value": 0,
                            "Average Rate": 0,
                        }

            if item["Type"] == "BUY":
                portfolio[scrip]["Total Quantity"] += float(item["Quantity"])
                portfolio[scrip]["Total Value"] += float(item["Total"])
            else:
                portfolio[scrip]["Total Quantity"] -= float(item["Quantity"])
                portfolio[scrip]["Total Value"] -= float(item["Total"])

            portfolio[scrip]["Average Rate"] = portfolio[scrip]["Total Value"] / portfolio[scrip]["Total Quantity"]
        else:
            if MISC_KEY not in portfolio:
                portfolio[MISC_KEY] = {
                            "Total Value": 0
                        }

            portfolio[MISC_KEY]["Total Value"] += float(item["Total"])

""" Round float value to 2 decimal
"""
def round_float(value):
    return round(value, 2) if value >= 0 else - round(abs(value), 2)

""" Create and update the portfolio
"""
def generate_portfolio(transactions):
    print "Generating portfolio..."

    portfolio = {}

    for data in transactions:
        update_portfolio(data, portfolio)

    total, balance = process_portfolio(portfolio)

    percentage = balance / total * 100

    # Display results
    tabular(portfolio)

    print "=" * 50
    print " | TOTAL INVESTMENT : " + colored("{0:25}".format("Rs. {:,}".format(round_float(total))), 'white') + " |"
    print " | BROKER CHARGES   : " + colored("{0:25}".format("Rs. {:,}".format(portfolio[MISC_KEY]["Total Value"])), 'cyan') + " |"
    print " | PROFIT/LOSS      : " + colored("{0:25}".format("Rs. {:,} ( {}% )".format(round_float(balance), round_float(percentage))), "red" if balance < 0 else "green") + " |"
    print "=" * 50

""" Report from portfolio
"""
def process_portfolio(portfolio):
    print "Processing portfolio..."

    balance = 0
    total = 0

    # Final
    for key, value in portfolio.items():
        # Misc
        if key == MISC_KEY:
            balance -= portfolio[key]["Total Value"]
            portfolio[key]["Total Value"] = round(portfolio[key]["Total Value"], 2)
        else:
            if portfolio[key]["Total Quantity"] < 0:
                del portfolio[key]
                continue

            portfolio[key]["Market Rate"] = float(get_quote(key))
            portfolio[key]["Current Value"] = portfolio[key]["Total Quantity"] * portfolio[key]["Market Rate"]
            portfolio[key]["Profit/Loss"] = portfolio[key]["Current Value"] - portfolio[key]["Total Value"]

            balance += portfolio[key]["Profit/Loss"]
            total += portfolio[key]["Total Value"]

            for k,v in portfolio[key].items():
                if v >= 0:
                    portfolio[key][k] = round(v, 2)
                else:
                    portfolio[key][k] = - round(abs(v), 2)



    return (total, balance)

""" Display the portfolio in tabular form
"""
def tabular(data):
    data_table = convert_to_table(data)

    print
    print
    print_table(data_table)
    print
    print

head = [
        'Scrip',
        'Total Quantity',
        'Total Value',
        'Average Rate',
        'Market Rate',
        'Current Value',
        'Profit/Loss'
        ]

""" Convert dictionary to two-dimentional list
"""
def convert_to_table(data):
    data_table = []

    data_table.append(head)

    for key, value in data.items():
        if key == MISC_KEY:
            continue

        row = []
        row.append(key)
        for k in head:
            if k in value:
                row.append(value[k])

        data_table.append(row)

    return data_table

""" Print the two dimentional list as a table
"""
def print_table(data_table):
    for entry in data_table[0]:
        print "+ {0:-^20}".format(""),
    print "+"

    is_first = True

    for line in data_table:
        for i, entry in enumerate(line):

            if is_first:
                print "| {0:^20}".format(entry),
            else:
                if head[i] == "Profit/Loss":
                    if float(entry) < 0:
                        color = 'red'
                    else:
                        color = 'green'
                else:
                    color = 'white'

                if head[i] == "Scrip":
                    print "| " + colored("{0:20}".format(entry), color),
                else:
                    print "| " + colored("{0:>20}".format('{0:.2f}'.format(entry)), color),

        print "|"

        for entry in line:
            print "+ {0:-^20}".format(""),

        print "|"

        is_first = False

""" Main
"""
if __name__ == '__main__':
    transactions = []

    # Parse HTML files
    for filename in glob.glob('*.htm'):
        data = process_entries(process_html_file(filename))
        transactions.append(data)

    generate_portfolio(transactions)
