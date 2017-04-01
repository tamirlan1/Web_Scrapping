# -*- coding: utf-8 -*-
import urllib #for url stuff
import xml.etree.ElementTree as ET #for XML stuff
import urllib2  # the lib that also handles the url stuff
import re #regular expressions
import warnings
import xlwt
import xlrd
from datetime import datetime # ADDED IN V2.1
warnings.filterwarnings("ignore")


# CHANGE THIS - USE RAW_INPUT OR CALENDAR!!!!!!!!
reported_date = '10-07-2016' # previous business day. Added in V2.2
s_required_min = 2000000
p_required_min = 200000

# Create Excel sheet
row = 0
col = 0

headers = ['Reported Date', 'Ticker', 'Person name', 'Person title (relationship)', 'Sold / Bought (D/A)', 'Number of shares', 
            'Value (shares * price)', 'index', 'company url', 'under 10b5', 'min_price', 'max_price', 'min_exp_date', 'max_exp_date'] ## ADDED IN V2.1

book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("Sheet 1")

for current_col in range(len(headers)):
    sheet1.write(row, current_col, headers[current_col])

row += 1
# read tickers from Excel, tickers are in column A - one ticker in one cell
excel_file = xlrd.open_workbook('tickers.xlsx')
worksheet = excel_file.sheet_by_name('Sheet1')
nrows = worksheet.nrows
tickers = []
curr_row = 0

while curr_row < nrows:
    row0 = worksheet.row_values(curr_row)
    nm = row0[0]
    tickers.append(nm)
    curr_row += 1
    
target_url = 'https://www.sec.gov/cgi-bin/current?q1=0&q2=6&q3=4'
visited_xml_url = [] # ADDED in v2.0

sec_type = ['Common Stock', 'Common Units', 'Common Shares', 'Ordinary Shares']
excel_filename = reported_date + '.xls'
count = 0

def valid_transaction(company_line, num_readings, company_url):

    success = 0 # number of successful transactions
    s_prod_sum = 0
    m_prod_sum = 0
    p_prod_sum = 0
    tot_num_s = 0
    tot_num_m = 0
    p_num_shares = 0
    match = re.findall('.xml', company_line)
    if match and len(match) > 1:
        words3 = company_line.split('"')
        xml_url = "https://www.sec.gov" + words3[3]
        if xml_url in visited_xml_url: # ADDED in v2.0
            return 0 # ADDED in v2.0
        try: # ADDED in v2.0
            uh = urllib.urlopen(xml_url) # ADDED in v2.0
        except: # ADDED in v2.0
            write_to_excel('COULD NOT ACCESS XML URL', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', xml_url) # ADDED in v2.0
            return 0 # ADDED in v2.0
        form4 = uh.read()
        tree = ET.fromstring(form4)

        ticker_name = ''

        # Check if company is in the tickers list
        for name in tree.findall('.//issuerTradingSymbol'):
            if name.text not in tickers:
                return 0
            ticker_name = name.text
        
        # Search non-deriv table
        if tree.find('nonDerivativeTable'):
            for nonderiv_sec_type in tree.findall('nonDerivativeTable/nonDerivativeTransaction/securityTitle/value'):
                if not any(word in nonderiv_sec_type.text for word in sec_type): # ADDED IN v2.1
                    return 0

            for transaction_code in tree.findall('nonDerivativeTable/nonDerivativeTransaction'):
                if transaction_code.find('transactionCoding/transactionCode').text == 'S':
                    try:
                        s_tr_shares = float(transaction_code.find('transactionAmounts/transactionShares/value').text)
                        s_tr_price = float(transaction_code.find('transactionAmounts/transactionPricePerShare/value').text)
                        s_product = s_tr_shares * s_tr_price
                        s_prod_sum += s_product
                        tot_num_s += s_tr_shares
                        success += 1
                    except:
                        pass # ADDED in v2.0
                    
                if transaction_code.find('transactionCoding/transactionCode').text == 'P':
                    try:
                        p_tr_shares = float(transaction_code.find('transactionAmounts/transactionShares/value').text)
                        p_tr_price = float(transaction_code.find('transactionAmounts/transactionPricePerShare/value').text)
                        p_num_shares += p_tr_shares
                        p_product = p_tr_shares * p_tr_price
                        p_prod_sum += p_product
                        success += 1
                    except:
                        pass # ADDED in v2.0
        
                if transaction_code.find('transactionCoding/transactionCode').text == 'M':
                    try:
                        m_tr_shares = float(transaction_code.find('transactionAmounts/transactionShares/value').text)
                        m_tr_price = float(transaction_code.find('transactionAmounts/transactionPricePerShare/value').text)
                        m_product = m_tr_shares * m_tr_price
                        m_prod_sum += m_product
                        tot_num_m += m_tr_shares
                        success += 1
                    except:
                        pass # ADDED in v2.0

            good_record = False
            diff_s_m_p = s_prod_sum - m_prod_sum - p_prod_sum
            if diff_s_m_p > s_required_min:
                value = diff_s_m_p
                good_record = True
            elif s_prod_sum == 0 and p_prod_sum > p_required_min:
                value = p_prod_sum
                good_record = True
            # elif tot_num_m > tot_num_s: # ADDDED/UPDT IN V2.2
                # if ....
                    # good_record = True
                    # value = 'exception'
                    # print 'Put info into excel as exception', tot_num_m, tot_num_s

            if good_record:
                person_name0 = tree.findall('reportingOwner/reportingOwnerId/rptOwnerName') # ADDED in v2.0
                if person_name0:
                	person_name = person_name0[0].text # ADDED in v2.0
                else:
                	person_name = "noname"
                person_title0 = tree.findall('reportingOwner/reportingOwnerRelationship/officerTitle') # ADDED in v2.0
                if person_title0:
                	person_title = person_title0[0].text # ADDED in v2.0
                else:
                	person_title = "notitle"
                sold_bought = ''
                if s_prod_sum > 0:
                    sold_bought = 'S'
                if m_prod_sum > 0:
                    sold_bought += 'M'
                if p_prod_sum > 0:
                    sold_bought += 'P'
                visited_xml_url.append(xml_url) # ADDED in v2.0

                ###### ADDED IN V2.1
                # Look for 10b5
                word_10b5 = '10b5'
                under_10b5 = 'NO'
                footnote_codes = tree.findall('footnotes/footnote')
                if footnote_codes:
                    for footnote in footnote_codes:
                        if word_10b5 in footnote.text:
                            under_10b5 = 'YES'
                            break

                if under_10b5 == 'NO':
                    remark_code = tree.findall('remarks')
                    if remark_code:
                        remark = remark_code[0].text
                        if word_10b5 == remark:
                            under_10b5 = 'YES'
                
                # Check Derivative Table
                deriv_data = check_deriv_table(tree)
                ################
                
                print 'Writing data' # ADDED in v2.0
                write_to_excel(ticker = ticker_name,
                    person_name = person_name, 
                    person_title = person_title,
                    sold_bought = sold_bought,
                    number_of_shares = p_num_shares, 
                    value = value,
                    num_readings = num_readings,
                    company_url = company_url, ## ADDED IN V2.1
                    under_10b5 = under_10b5,## ADDED IN V2.1
                    min_price = deriv_data[0],## ADDED IN V2.1
                    max_price = deriv_data[1],## ADDED IN V2.1
                    min_exp_date = deriv_data[2],## ADDED IN V2.1
                    max_exp_date = deriv_data[3]) ## ADDED IN V2.1
                global row
                row += 1
        return success
    return success


def write_to_excel(ticker, person_name, person_title, sold_bought, number_of_shares, value, num_readings, company_url, under_10b5,
                    min_price, max_price, min_exp_date, max_exp_date): ## ADDED IN V2.1
    global row
    global reported_date
    to_write = [reported_date, ticker, person_name, person_title, sold_bought, number_of_shares, value, num_readings, company_url, under_10b5,
                min_price, max_price, min_exp_date, max_exp_date] ## ADDED IN V2.1
    print to_write
    for i in range(len(to_write)):
        try:
            if not to_write[i] or to_write[i] == '':
                to_write[i] = 'NA'
            sheet1.write(row, i, to_write[i])
            book.save(excel_filename)
        except:
            continue

############ ADDED IN V2.1
# CHECK DERIVATIVE TABLE - MIN/MAX
def check_deriv_table(tree):
    min_price = None
    max_price = None
    min_exp_date = None
    max_exp_date = None
    if tree.find('derivativeTable/derivativeTransaction'):
        for der_transaction_code in tree.findall('derivativeTable/derivativeTransaction'):
            if der_transaction_code.find('transactionCoding/transactionCode').text == 'M':
                try:
                    price = der_transaction_code.find('conversionOrExercisePrice/value').text
                except:
                    price = None
                try:
                    dt = der_transaction_code.find('expirationDate/value').text
                    exp_date = datetime.strptime(dt, "%Y-%m-%d")
                except:
                    exp_date = None

                if price:
                    if not min_price:
                        min_price = price
                        max_price = price
                    else:
                        min_price = min(min_price, price)
                        max_price = max(max_price, price)

                if exp_date:
                    if not min_exp_date:
                        min_exp_date = exp_date
                        max_exp_date = exp_date
                    else:
                        min_exp_date = min(min_exp_date, exp_date)
                        max_exp_date = max(max_exp_date, exp_date)

    deriv_data = [min_price, max_price, min_exp_date, max_exp_date]
    return deriv_data
############ 

num_readings = 0

# find and open form 4 on the SEC website - in XML format
data = urllib2.urlopen(target_url)
for line in data:
    if line.startswith(reported_date) or line.startswith('<p>The total number of matches'):
        num_readings += 1
        print num_readings
        words = line.split('"')
        company_url = "https://www.sec.gov" + words[1]
        try: # ADDED in v2.0
            company_data = urllib2.urlopen(company_url) # ADDED in v2.0
        except: # ADDED in v2.0
            write_to_excel('COULD NOT ACCESS COMPANY URL', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', company_url) # ADDED in v2.0
            continue # ADDED in v2.0
        for company_line in company_data:
            text = valid_transaction(company_line, num_readings, company_url)
print 'Good companies: ', row - 1 # ADDED in v2.0
print 'Number of readings: ', num_readings