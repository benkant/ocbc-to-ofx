#!/usr/bin/env python

"""
Immigrant is a small script forked from Nordea to OFX by jgoney
(https://github.com/jgoney/Nordea-to-OFX)

It converts OCBC transaction lists (that are in some silly CSV format) to OFX,
for use with modern financial management software

"Immigrant" is a play on the name, Overseas Chinese Bank. 

Immigrant: converts OCBC transaction lists (CSV) to OFX for use with
financial management software.

This can handle both Bank Accounts and Credit Card Statements
"""

import csv
import os
import sys
import time
import re
from datetime import *
import time
import pprint
import os.path

# Pretty Print settings
pp = pprint.PrettyPrinter(indent=4)

# Here you can define the currency used with your account (e.g. EUR, SEK)
MY_CURRENCY = "SGD"

# ============= CONSTANTS ================
# Row number where the Field Names (e.g. transaction date, value date etc) are at (necessary for CSV Dictreader to know!)
# TODO: This is very fragile and breaks when it is a credit card statement
linesToSkip = 7
LINES_SKIPPED_IF_BANK_ACCOUNT = 6
LINES_SKIPPED_IF_CREDIT_CARD = 7
# ========================================

# This was taken from the Nordea OFX library
def getTransType(trans, amt):
    """
    Converts a transaction description (e.g. "Deposit") to an OFX
    standardized transaction (e.g. "DEP").
    
    @param trans: A textual description of the transaction (e.g. "Deposit")
    @type trans: String
    @param amt: The amount of a transaction, used to determine CREDIT or DEBIT.
    @type amt: String
    
    @return: The standardized transaction type
    @rtype: String
    """
    if trans == "ATM withdr/Otto." or trans == "Debit cash withdrawal":
        return "ATM"
    elif trans == "Deposit":
        return "DEP"
    elif trans == "Deposit interest":
        return "INT"
    elif trans == "Direct debit":
        return "DIRECTDEBIT"
    elif trans == "e-invoice" or trans == "e-payment":
        return "PAYMENT"
    elif trans == "ePiggy savings transfer" or trans == "Own transfer":
        return "XFER"
    elif trans == "Service fee VAT 0%":
        return "FEE"
    else:
        if amt[0] == '-':
            return "DEBIT"
        else:
            return "CREDIT"

def getTransAmount(deposits, withdrawals):
    if withdrawals: 
        return "-" + withdrawals
    else:
        return deposits

def convertFile(f, originalFilePath):
    """
    Creates new OFX file, then maps transactions from original CSV (f) to
    OFX's version of XML.
    
    @param f: A file handle (f) for the original CSV transactions list.
    @type f: File
    """

    originalParentDirectory = os.path.dirname(originalFilePath)
    originalFileName = os.path.basename(originalFilePath).split(".")[0] # TODO: Will there be a scenario where there are multiple periods?
    newOFXFileName = originalParentDirectory + "/" + originalFileName + ".ofx"
    
    print("Parsing CSV into " + newOFXFileName)

    # Open/create the .ofx file
    try:
        outFile = open(newOFXFileName, "w")
    except IOError:
        print("Output file couldn't be created. Program will now exit")
        sys.exit(2)

    # Reads the first line to ascertain if it is a bank account number or a credit card number
    # Will set parameter of whether it is a credit card or a bank account
    csvReader = csv.reader(f, dialect=csv.excel_tab)
    acctDetails = csvReader.next()[0] # TODO: verify that this is the line with the acct number
    BankAccountDetails = re.findall(r"\d{3}-\d{6}-\d{3}", acctDetails)
    CreditCardDetails = re.findall(r"\d{4}-\d{4}-\d{4}-\d{4}", acctDetails)
    acctNumber = ""
    # acctNumber = "".join(re.findall(r"\d{3}-\d{6}-\d{3}", acctDetails)) # Converted to string
    if CreditCardDetails:
        acctNumber = CreditCardDetails[0]   # Convert from List to String
        linesToSkip = LINES_SKIPPED_IF_CREDIT_CARD
    if BankAccountDetails:
        acctNumber = BankAccountDetails[0]
        linesToSkip = LINES_SKIPPED_IF_BANK_ACCOUNT

    # SKIPS THE FIRST FEW LINES
    # see http://stackoverflow.com/questions/7588426/how-to-skip-pre-header-lines-with-csv-dictreader
    # TODO: Find more antifragile way to separate metadata from transaction data
    # TODO: Find a way to do it so it can handle credit cards too
    f.seek(0)
    for i in range(0,linesToSkip-1):    # TODO: pass "5" in as a paramter (in constants) in case OCBC changes its format
        next(f)
    csv_dictReader = csv.DictReader(f)

    # Reads csv transactions into list
    transactionEntries = []
    for line in csv_dictReader:
        #print "TRANSACTION ENTRY"
        #print line
        transactionEntries.append(line)

    # Simple test (later measured against number of transactions in OFX file)
    numEntries = 0
    
    # Gets the start and end dates by iterating over all Transaction Entries
    startDate = datetime.now()
    endDate = datetime.min
    for entry in transactionEntries:
        if entry['Transaction date']:
            entryDate = datetime.strptime(entry['Transaction date'], "%d/%m/%Y")
            if entryDate < startDate: startDate = entryDate
            if entryDate > endDate: endDate = entryDate
            numEntries += 1
    startDateString = startDate.strftime('%Y%m%d')   # Arbitrary time of 000000 assigned for each start date/end date
    endDateString =  endDate.strftime('%Y%m%d')

    # Creates string from file's time stamp
    timeStamp = time.strftime(
        "%Y%m%d%H%M%S", time.localtime((os.path.getctime(f.name))))

    # Write header to file (includes timestamp)
    outFile.write(
        '''<?xml version="1.0" encoding="ANSI" standalone="no"?>
<?OFX OFXHEADER="200" VERSION="200" SECURITY="NONE" OLDFILEUID="NONE" NEWFILEUID="NONE"?>
<OFX>
        <SIGNONMSGSRSV1>
                <SONRS>
                        <STATUS>
                                <CODE>0</CODE>
                                <SEVERITY>INFO</SEVERITY>
                        </STATUS>
                        <DTSERVER>''' + timeStamp + '''</DTSERVER>
                        <LANGUAGE>ENG</LANGUAGE>
                </SONRS>
        </SIGNONMSGSRSV1>
    <BANKMSGSRSV1>
        <STMTTRNRS>
            <TRNUID>0</TRNUID>
            <STATUS>
                <CODE>0</CODE>
                <SEVERITY>INFO</SEVERITY>
            </STATUS>
            <STMTRS>
                <CURDEF>''' + MY_CURRENCY + '''</CURDEF>
                <BANKACCTFROM>
                    <BANKID>OCBC</BANKID>
                    <ACCTID>''' + acctNumber + '''</ACCTID>
                    <ACCTTYPE>CHECKING</ACCTTYPE>
                </BANKACCTFROM>
                <BANKTRANLIST>
                    <DTSTART>''' + startDateString + '''</DTSTART>
                    <DTEND>''' + endDateString + '''</DTEND>
                    ''')

    numTransactions = 0
    while len(transactionEntries):
        numTransactions += 1
        currentTransaction = transactionEntries.pop(0)
        
        # This deals with the weird formatting where descriptions are broken over two lines
        while len(transactionEntries):
            # If there is a Transaction Date, it means that there is no more description
            if transactionEntries[0]['Transaction date']: break 
            # This gets the description from the next line
            additionalDescription = transactionEntries.pop(0)['Description']
            currentTransaction['Description'] += " " + additionalDescription

        # Just adds a default time of 000000 to each of the transactions
        dateVector = currentTransaction['Transaction date'].split('/')
        if len(dateVector[1]) < 2: dateVector[1] = '0' + dateVector[1]
        if len(dateVector[0]) < 2: dateVector[0] = '0' + dateVector[0]
        entryDate = dateVector[2] + dateVector[1] + dateVector[0]
        entryAmount = getTransAmount(currentTransaction['Deposits (SGD)'], currentTransaction['Withdrawals (SGD)'])
        # pp.pprint(currentTransaction)

        # Quick and dirty trans type (needs a function table)
        outFile.write(
                '''<STMTTRN>
                        <TRNTYPE>''' + "getTransType(transaction, amount)" + '''</TRNTYPE>
                        <DTPOSTED>''' + entryDate + '''</DTPOSTED>
                        <TRNAMT>''' + entryAmount + '''</TRNAMT>
                        <FITID>''' + 'refNum' + '''</FITID>
                        <NAME>''''''</NAME>
                        <MEMO>''' + currentTransaction['Description'] + '''</MEMO>
                    </STMTTRN>
                    ''')
    
    print("========= ERROR CHECKING =========")
    print("Num Transactions (" + str(numTransactions) + ") and Num Entries (" +str(numEntries) + ") should match!")
    print("==================================")

    outFile.write(
        '''</BANKTRANLIST>
                        </STMTRS>
                </STMTTRNRS>
        </BANKMSGSRSV1>
</OFX>''')

    outFile.close()

if __name__ == '__main__':
    # Check that the args are valid
    if len(sys.argv) < 2:
        print("Error: no filenames were given.\nUsage: %s [one or more file names]" % sys.argv[0])
        sys.exit(1)

    # Open the files and put the handles in a list
    for arg in (sys.argv[1:]):
        try:
            f_in = open(arg, "rU")
            print("Opening %s" % arg)
            convertFile(f_in, arg)
        except IOError:
            print("Error: file %s couldn't be opened" % arg)
        else:
            f_in.close()
            print("%s is closed" % arg)
