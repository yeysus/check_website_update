# Script to check if a webpage has changed since it was called last time.
# This script accepts as argument an excel file with a list of webpages to check,
# a column indicating if that row should be processed,
# a column with a hash of the previously accessed webpage.
# If will change the value of the hash in the excel file,
# it will also indicate in a column if there has been a changed compared to the previous hash.

# Using ideas or snips from
# https://www.geeksforgeeks.org/python-script-to-monitor-website-changes/
# https://stackoverflow.com/questions/54047212/how-to-update-a-portion-of-existing-excel-sheet-with-filtered-dataframe
# https://stackoverflow.com/questions/287871/how-to-print-colored-text-to-the-terminal

# This script uses Selenium and the Chrome driver.

# The Selenium module must be installed.
# pip3 install selenium
# For the Mac, download the Chrome driver from
# https://chromedriver.storage.googleapis.com/index.html?path=88.0.4324.96/
# as specified in:
# https://sites.google.com/a/chromium.org/chromedriver/downloads
# Download the driver, e.g. the chromedriver, in the same directory as this python script.

# For saving data to an excel file.
# sudo pip3 install openpyxl

# Importing libraries 

import sys
import pandas as pd
import hashlib 
from urllib.request import urlopen, Request 
import openpyxl
import traceback

from selenium import webdriver
# To handle cookies.
from selenium.webdriver.chrome.options import Options
from time import sleep

pathToChromedriver = '/Users/jesusdelvalle/Documents/projects/check_website_update/chromedriver'
chromeOptionsCookieFolder = "user-data-dir=selenium"

# From: https://stackoverflow.com/questions/287871/how-to-print-colored-text-to-the-terminal
class bcolors:
  HEADER = '\033[95m'
  OKBLUE = '\033[94m'
  OKCYAN = '\033[96m'
  OKGREEN = '\033[92m'
  WARNING = '\033[93m'
  FAIL = '\033[91m'
  ENDC = '\033[0m'
  BOLD = '\033[1m'
  UNDERLINE = '\033[4m'

DEBUG = True

nameOfExcelColumnWithURLs = 'URL'
# There should be a yes/no column to track if we want to check for updates to that website.
nameOfExcelColumnToDecideIfCheckForUpdates = 'Use'
# Value in that Column if it should not be checked for updates.
valueOfExcelColumnIfCheckForUpdatesNo = 'No'
# Name of column containing the last hash.
nameOfExcelColumnWithHashes = 'Hash'
# Name of column where we will store if the page (the hash) changed.
nameOfExcelColumnWithChange = 'Changed'
# Value of the column where we will store if the page (the hash) changed.
valueOfExcelColumnWithChange = 'Yes'

try:
  # Expected: Excel file with ending .xslx
  inputFilename = sys.argv[1]

  excel = pd.read_excel (inputFilename, header=0)

  workbook = openpyxl.load_workbook (inputFilename)

  # To handle cookies.
  # From: https://stackoverflow.com/questions/15058462/how-to-save-and-load-cookies-using-python-selenium-webdriver
  chrome_options = Options()
  chrome_options.add_argument(chromeOptionsCookieFolder)
  driver = webdriver.Chrome(pathToChromedriver, options=chrome_options)

  # i is the counter for the number of websites.
  i = 0
  for index, row in excel.iterrows ():


    use = row[nameOfExcelColumnToDecideIfCheckForUpdates]
    if use == valueOfExcelColumnIfCheckForUpdatesNo:
      continue

    url = row [nameOfExcelColumnWithURLs]
    oldHash = row [nameOfExcelColumnWithHashes]
    hashColumnIndex = excel.columns.get_loc (nameOfExcelColumnWithHashes)
    changeColumnIndex = excel.columns.get_loc (nameOfExcelColumnWithChange)

    # Multiple variants of the headers below were not successful for some webpages.
    # Either it took too long to get the response, or they were unreachable.
    # So I moved to use Selenium.
    """
    if (DEBUG):
      print (str (index) + " " + url + " " + str (oldHash))

    urlRequest = Request (url, headers={"Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9", 
                                        "Accept-Encoding": "gzip, deflate", 
                                        "Accept-Language": "en-GB,en-US;q=0.9,en;q=0.8", 
                                        "Dnt": "1", 
                                        "Host": "httpbin.org", 
                                        "Upgrade-Insecure-Requests": "1", 
                                        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_4) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.97 Safari/537.36"}) 
    response = urlopen (urlRequest).read () 
    """

    # To handle special situations like missing, wrong, expired domain names.
    try:
      driver.get(url)
    except Exception as error:
      print (f"{bcolors.FAIL}Error:{bcolors.ENDC}", error)
      print (traceback.format_exc ())
      continue

    # The script should be run twice.
    # The first time you run the script, you give sleep time so the webpage
    # loads and you can click on the "Accept cookies" buttons.
    # Then run the script a second time, the pages will take the previously
    # saved cookies and you don't see the "Accept Cookies" messages anymore.
    # These Screenhots are saved and will replace the Screenshots from the
    # first time you run the script.
    sleep(1)
    response = driver.page_source

    newHash = hashlib.sha224 (response.encode('utf-8')).hexdigest () 
    
    if (DEBUG):
      print ("newHash: " + newHash)
      print ("oldHash: " + str (oldHash))

    if (oldHash == newHash):
      continue
    else:
      print ("Hash changed in " + url)
      # Update cell content and save it to excel.

      # row=index+2 since index is defined by pandas dataframe, which does not count the headers,
      # and it is 0-based. row in openpyxl is 1-based.
      cellHashColumn = hashColumnIndex + 1
      cellChangeColumn = changeColumnIndex + 1
      cellRow = index + 2
      workbook['Sheet1'].cell (cellRow, cellHashColumn, value=newHash)
      workbook['Sheet1'].cell (cellRow, cellChangeColumn, value=valueOfExcelColumnWithChange)
      
      workbook.save (inputFilename)

    i = i + 1

  driver.quit ()
  print ("Number of webpages visited: ", i)

except IOError:
  print (f"{bcolors.FAIL}IO Error (Wrong filename?): {bcolors.ENDC}" + inputFilename)
  print (traceback.format_exc ())
except IndexError as error:
  print (f"{bcolors.FAIL}Index Error (Forgot filename?):{bcolors.ENDC}", error)
  print (traceback.format_exc ())
except Exception as error:
  print (f"{bcolors.FAIL}Error:{bcolors.ENDC}", error)
  print (traceback.format_exc ())

