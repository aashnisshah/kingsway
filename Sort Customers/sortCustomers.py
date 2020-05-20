import xlrd

# CONSTANTS
CUST_CODE = 2
EMAIL = 12
PHONE = 13


def read_data():
    workbook = xlrd.open_workbook("./accountCustomer.xls")
    sheet = workbook.sheet_by_index(0)
    data = [sheet.row_values(rowx) for rowx in range(sheet.nrows)]
    return data


def find_duplicate_customer_codes(data):
    ALL_CUSTOMERS_BY_CUST_CODE = {}
    DUPLICATE_CUSTOMER_CODES = []

    for d in data:
        if d[CUST_CODE] in ALL_CUSTOMERS_BY_CUST_CODE:
            DUPLICATE_CUSTOMER_CODES.append(d[CUST_CODE])
        else:
            ALL_CUSTOMERS_BY_CUST_CODE[d[CUST_CODE]] = d

    return DUPLICATE_CUSTOMER_CODES


def find_duplicate_emails(data):
    ALL_CUSTOMERS_BY_EMAIL = {}
    DUPLICATE_EMAIL_ADDRESSES = []
    DUPLICATE_RESULTS = {}

    for d in data:
        split_emails = d[EMAIL].split(',')
        for email in split_emails:
            if email in ALL_CUSTOMERS_BY_EMAIL:
                if email not in DUPLICATE_EMAIL_ADDRESSES:
                    DUPLICATE_EMAIL_ADDRESSES.append(email)
                ALL_CUSTOMERS_BY_EMAIL[email].append(d)
            else:
                ALL_CUSTOMERS_BY_EMAIL[email] = [d]

    for dup in DUPLICATE_EMAIL_ADDRESSES:
        curDup = ALL_CUSTOMERS_BY_EMAIL[dup]
        for customer in curDup:
            for cust2 in curDup:
                if cust2 is customer:
                    continue

                if customer[2] not in DUPLICATE_RESULTS:
                    DUPLICATE_RESULTS[customer[2]] = []

                if cust2[2] not in DUPLICATE_RESULTS[customer[2]]:
                    DUPLICATE_RESULTS[customer[2]].append(cust2[2])

    return DUPLICATE_RESULTS


def find_duplicate_phone_numbers(data):
    ALL_CUSTOMERS_BY_PHONE = {}
    DUPLICATE_PHONE_NUMBERS = []
    DUPLICATE_RESULTS = {}

    for d in data:
        split_phone_numbers = d[PHONE].split(',')
        for phone in split_phone_numbers:
            if phone in ALL_CUSTOMERS_BY_PHONE:
                if phone not in DUPLICATE_PHONE_NUMBERS:
                    DUPLICATE_PHONE_NUMBERS.append(phone)
                ALL_CUSTOMERS_BY_PHONE[phone].append(d)
            else:
                ALL_CUSTOMERS_BY_PHONE[phone] = [d]

    for dup in DUPLICATE_PHONE_NUMBERS:
        curDup = ALL_CUSTOMERS_BY_PHONE[dup]
        for customer in curDup:
            for cust2 in curDup:
                if cust2 is customer:
                    continue

                if customer[2] not in DUPLICATE_RESULTS:
                    DUPLICATE_RESULTS[customer[2]] = []

                if cust2[2] not in DUPLICATE_RESULTS[customer[2]]:
                    DUPLICATE_RESULTS[customer[2]].append(cust2[2])

    # for dup in DUPLICATE_RESULTS:
    #     print(str(dup))

    return DUPLICATE_RESULTS


def output(codes, emails, phones):
    output = ">>> Below are duplicate customers based on their customer codes, emails and phone numbers"

    output += "\n\n>>> Duplicate Customers by Customer Codes\n\n"
    for code in codes:
        output += code + "\n"

    output += "\n\n>>> Duplicate Customers by Email Address\n\n"
    for email in emails:
        output += email + ": " + str(emails[email]) + "\n"

    output += "\n\n>>> Duplicate Customers by Phone Numbers\n\n"
    for phone in phones:
        output += phone + ": " + str(phones[phone]) + "\n"

    text_file = open("duplicates.txt", "w")
    text_file.write(output)
    text_file.close()


def __main__():
    print('>>> read data')
    data = read_data()

    print('>>> find duplicate customer codes')
    duplicate_customer_by_codes = find_duplicate_customer_codes(data)
    # RESULTS FOR ALL DUPLICATE CUSTOMER CODES: ['CCP0013263', 'CKS0000112', 'CLU0003919', 'CGI0000436', 'CMT0001267', 'CAI0000001', 'CFR0000422', 'CFN0000337', 'CFN0000896']

    print('>>> find duplicate email addresses')
    duplicate_customers_by_emails = find_duplicate_emails(data)

    print('>>> find duplicate phone numbers')
    duplicate_customers_by_phone = find_duplicate_phone_numbers(data)

    output(duplicate_customer_by_codes, duplicate_customers_by_emails,
           duplicate_customers_by_phone)
    print('>>> finished')


__main__()

# ALLEMAILS = {}
# REPEATEDEMAILS = {}

# for d in data:
#     splitEmails = d[EMAIL].split(",")
#     # print(splitEmails)
#     for email in splitEmails:
#         email = email.lower()
#         if email.find('purchases@silverstone.co.ke') >= 0:
#             print(str(d[CUST_CODE]) + ' ' + str(d[EMAIL]))

#         # if email in ALLEMAILS:
#         #     # print('d[custcode]: ' + d[CUST_CODE] +
#         #     #       ' ALLEMAILS[email]: ' + ALLEMAILS[email][CUST_CODE])
#         #     if d[CUST_CODE] is not ALLEMAILS[email][CUST_CODE]:
#         #         print('cust codes dont match? ' +
#         #               d[CUST_CODE] + ' ' + ALLEMAILS[email][CUST_CODE] + ' ' + email)
#         #     #     REPEATEDEMAILS[email] = d
#         #     #     # print('repeated emails: ' + email)
#         #     else:
#         #         print('cust codes match? ' +
#         #               d[CUST_CODE] + ' ' + ALLEMAILS[email][CUST_CODE] + ' ' + email)
#         #     #     # same customer code, don't add
#         #     #     print('same cust code: ' + email)
#         # else:
#         #     ALLEMAILS[email] = d
#     # ALLEMAILS

# print(data[1][EMAIL])


# thoughts:
#  - most accounts have multiple phone numbers or multiple email addresses
#  - many accounts have multiple people from the same company - is there a
#     way to handle companies differently if there's one "bill"
#  - some customers like CFN0000896 are duplicate entries (Silverstone Tyres (K) LTD)
