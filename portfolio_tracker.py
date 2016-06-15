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
            "JAMNA AUTO INDUSTRIE" : "NSE:JAMNAAUTO",
            "IDBI BANK LIMITED" : "NSE:IDBI",
            "TATA METALIKS" : "NSE:TATAMETALI",
            "KALYANI STEELS LTD": "NSE:KSL",
            "TV TODAY NETWORK LTD": "NSE:TVTODAY",
        }

MISC_KEY = "--CHARGES--"

COLUMNS = ['Order No', 'Order Time', 'Trade No.', 'Trade Time', 'Security', 'Bought Qty', 'Sold Qty', 'Gross Rate', 'Gross Total', 'Brokerage', 'Net Rate', 'Service Tax', 'STT', 'Total']
COLUMNS_NEW = ['Order No', 'Order Time', 'Trade No.', 'Trade Time', 'Security', 'Buy/Sell', 'Quantity', 'Gross Rate', 'Brokerage', 'Net Rate', 'Closing Rate', 'Total', 'Remarks']

BROKERAGE_RATE = 0.004
EXIT_LOAD_RATE = 0.004

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

    if symbol not in KEYWORDS:
        print "Scrip information not available!"
        return 0.0

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

""" Compile trades
"""
def compile_trades(data, trades):
    print "Updating trades..."

    for item in data:
        if "Trades" in item:
            compile_trades(item["Trades"], trades)
            continue

        if item["Type"] in ["BUY", "SELL"]:
            scrip = item["Security"]

            if scrip not in trades:
                trades[scrip] = {
                            "Buy": [],
                            "Sell": []
                        }

            if item["Type"] == "BUY":
                trades[scrip]["Buy"].append((float(item["Quantity"]), float(item["Total"])))
            else:
                trades[scrip]["Sell"].append((float(item["Quantity"]), float(item["Total"])))

        else:
            if MISC_KEY not in trades:
                trades[MISC_KEY] = {
                            "Total Value": 0
                        }

            trades[MISC_KEY]["Total Value"] += float(item["Total"])

""" Process Trades
"""
def process_trades(trades):
    for key in dict(trades):
        if key == MISC_KEY:
            continue

        trades[key].update({
                "Total Buy": 0,
                "Total Sell": 0,
                "Total Buy Value": 0,
                "Total Sell Value": 0,
                "Buy Rate": 0,
                "Sell Rate": 0
            })

        for trade in trades[key]["Buy"]:
            trades[key]["Total Buy"] += trade[0]
            trades[key]["Total Buy Value"] += trade[1]
            trades[key]["Buy Rate"] = trades[key]["Total Buy Value"] / trades[key]["Total Buy"]

        del(trades[key]["Buy"])

        for trade in trades[key]["Sell"]:
            trades[key]["Total Sell"] += trade[0]
            trades[key]["Total Sell Value"] += trade[1]
            trades[key]["Sell Rate"] = trades[key]["Total Sell Value"] / trades[key]["Total Sell"]

        del(trades[key]["Sell"])

        # If old trades, ignore
        if trades[key]["Total Buy"] == 0:
            del trades[key]
            continue

        trades[key]["Balance"] = trades[key]["Total Buy"] - trades[key]["Total Sell"]

        if trades[key]["Balance"] < 0:
            # Missing BUY items, only calculate for accounted trades
            trades[key]["Cleared"] = trades[key]["Sell Rate"] * trades[key]["Total Buy"] - trades[key]["Buy Rate"] * trades[key]["Total Buy"]
        elif trades[key]["Balance"] == 0:
            trades[key]["Cleared"] = trades[key]["Sell Rate"] * trades[key]["Total Sell"] - trades[key]["Buy Rate"] * trades[key]["Total Buy"]
        else:
            # Only calculate for sold quantity
            trades[key]["Cleared"] = trades[key]["Sell Rate"] * trades[key]["Total Sell"] - trades[key]["Buy Rate"] * trades[key]["Total Sell"]

""" Update portfolio with trades
"""
def update_portfolio(trades, portfolio):
    print "Updating portfolio..."

    for scrip in trades:
        if scrip == MISC_KEY:
            continue

        portfolio[scrip] = {
                    "Total Quantity": trades[scrip]["Balance"],
                    "Total Value": trades[scrip]["Buy Rate"] * trades[scrip]["Balance"],
                    "Average Rate": trades[scrip]["Buy Rate"],
                    "Cleared": trades[scrip]["Cleared"]
                }
    """

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

        else:
            if MISC_KEY not in portfolio:
                portfolio[MISC_KEY] = {
                            "Total Value": 0
                        }
    """

    portfolio[MISC_KEY] = {
                "Total Value": 0
            }
    portfolio[MISC_KEY]["Total Value"] += trades[MISC_KEY]["Total Value"]

""" Create and update the portfolio
"""
def generate_portfolio(transactions):
    print "Generating portfolio..."

    trades = {}

    for data in transactions:
        compile_trades(data, trades)

    process_trades(trades)

    #system

    portfolio = {}

    #for data in transactions:
    update_portfolio(trades, portfolio)

    report = process_portfolio(portfolio)

    percentage = report['balance'] / report['total'] * 100

    # Display results
    tabular(portfolio)

    print "=" * 50
    print " | TOTAL INVESTMENT : " + colored("{0:25}".format("Rs. {:,.2f}".format(report['total'])), 'white') + " |"
    print " | CURRENT VALUE    : " + colored("{0:25}".format("Rs. {:,.2f}".format(report['current_value'])), 'yellow') + " |"
    print " | ENTRY LOAD       : " + colored("{0:25}".format("Rs. {:,.2f}".format(report['entry_load'])), 'cyan') + " |"
    print " | EXIT LOAD        : " + colored("{0:25}".format("Rs. {:,.2f}".format(report['exit_load'])), 'cyan') + " |"
    print " | PROFIT/LOSS      : " + colored("{0:25}".format("Rs. {:,.2f} ( {:.2f}% )".format(report['balance'], percentage)), "red" if report['balance'] < 0 else "green") + " |"
    print " | CLEARED          : " + colored("{0:25}".format("Rs. {:,.2f}".format(report['cleared'])), "red" if report['cleared'] < 0 else "green") + " |"
    print "=" * 50

""" Report from portfolio
"""
def process_portfolio(portfolio):
    print "Processing portfolio..."

    balance = 0
    total = 0
    current_value = 0
    cleared = 0

    # Final
    for key in dict(portfolio):
        # Misc
        if key == MISC_KEY:
            balance -= portfolio[key]["Total Value"]
            portfolio[key]["Total Value"] = round(portfolio[key]["Total Value"], 2)
        else:
            portfolio[key]["Market Rate"] = float(get_quote(key))

            if portfolio[key]["Total Value"] > 0:
                portfolio[key]["Current Value"] = portfolio[key]["Total Quantity"] * portfolio[key]["Market Rate"]
                portfolio[key]["Profit/Loss"] = portfolio[key]["Current Value"] - portfolio[key]["Total Value"]
                portfolio[key]["ROI"] = portfolio[key]["Profit/Loss"] / portfolio[key]["Total Value"] * 100
                portfolio[key]["Average Rate"] = portfolio[key]["Total Value"] / portfolio[key]["Total Quantity"]
            else:
                portfolio[key]["Current Value"] = 0
                portfolio[key]["Profit/Loss"] = 0
                portfolio[key]["ROI"] = 0
                portfolio[key]["Average Rate"] = 0

            balance += portfolio[key]["Profit/Loss"]
            total += portfolio[key]["Total Value"]
            current_value += portfolio[key]["Current Value"]
            cleared += portfolio[key]["Cleared"]

    entry_load = portfolio[MISC_KEY]["Total Value"] + (total * BROKERAGE_RATE)
    exit_load = current_value * EXIT_LOAD_RATE
    balance -= exit_load

    return {
                "total": total,
                "current_value": current_value,
                "entry_load": entry_load,
                "exit_load": exit_load,
                "balance": balance,
                "cleared": cleared
            }

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
        'Profit/Loss',
        'ROI',
        'Cleared'
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
                if entry == 0:
                    if head[i] == 'Market Rate':
                        print "| " + colored("{0:^20}".format("_INVALID_"), 'red'),
                    else:
                        print "| " + colored("{0:^20}".format("_._"), 'grey'),
                else:
                    if head[i] == "Profit/Loss" or head[i] == "ROI" or head[i] == "Cleared":
                        if entry < 0:
                            color = 'red'
                        else:
                            color = 'green'
                    else:
                        color = 'white'

                    if head[i] == "Scrip":
                        print "| " + colored("{0:20}".format(entry), color),
                    else:
                        if head[i] == "ROI":
                            print "| " + colored("{0:>20}".format('{0:.2f}%'.format(entry)), color),
                        else:
                            print "| " + colored("{0:>20}".format('{0:.2f}'.format(entry)), color),

        print "|"

        for entry in line:
            print "+ {0:-^20}".format(""),

        print "+"

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
