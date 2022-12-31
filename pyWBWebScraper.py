import bs4, requests, re, openpyxl, sys

#wb = openpyxl.load_workbook(sys.argv[1])  # user enters the filepath to the workbook as an argument to run the script
wb = openpyxl.load_workbook('Web Scraping Project Sample.xlsx')  # open the workbook
inSheet = wb[wb.sheetnames[0]]
outSheet = wb[wb.sheetnames[1]]

search_term = str(inSheet['B1'].value)  # the term to search for, provided in cell B1 of the first sheet in the workbook
txtSearchRegex = re.compile(search_term)  # compile the regex object

prox_url = 'http://api.proxiesapi.com'  # url to ProxiesAPI
auth_key = 'REPLACE THIS TEXT WITH API KEY FROM PROXIES API'  # auth_key for ProxiesAPI

results = []  # list to store results from searches

urls = []  # list of urls to scrape and search
for i in range(2,len(inSheet['A'])):  # gather urls from workbook
    if inSheet.cell(row=i, column=1).value:
        urls.append(str(inSheet.cell(row=i, column=1).value))

def searchSoup(soup_obj, reg_obj):  # takes a beautiful soup object and regex object then performs the search, returns true on the first instance found
    for element in soup_obj.find_all():
        item = reg_obj.findall(element.get_text())
        if item:
            return True

for i in range(len(urls)):  # iterate through the list of urls provided
    url = 'https://' + urls[i]  # generate full url to request
    PARAMS = {'auth_key':auth_key, 'url':url}  # generate parameters for the api
    res = requests.get(url = prox_url, params = PARAMS)  # download the html
    soup = bs4.BeautifulSoup(res.text, 'html.parser')  # construct BeautifulSoup object
    if searchSoup(soup, txtSearchRegex):  # if found
        results.append([urls[i], res.text, 'Y'])
    else:  # if not found
        results.append([urls[i], res.text, 'N'])
        
for i in range(len(results)):  # add the results to the second sheet of the workbook
    if str(inSheet.cell(row=i+2, column=1).value) == str(results[i][0]):
        outSheet.cell(row=i+2, column=1).value = str(results[i][0])
        outSheet.cell(row=i+2, column=2).value = str(results[i][1])
        outSheet.cell(row=i+2, column=3).value = str(results[i][2])
        print(outSheet.cell(row=i+2, column=1).value + ' ' + outSheet.cell(row=2, column=3).value)

wb.save('Web Scraping Project Sample_out.xlsx')  # save the output workbook
