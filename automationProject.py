from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException
import time
import os
import glob
import json
import pandas as pd
from distutils.version import StrictVersion
import re


def process(sheetName):
    # Read the Excel data without specifying columns
    excelDataDf = pd.read_excel('ドライババージョンアップ対応チェックシート(FB).xlsx', sheet_name=sheetName)
    latestVersion = {}
    driverNumber=0
    # Map the column names in English
    columnMapping = {
        'プロダクト': 'productName',
        'ドライバ': 'driverName',
        'Ver.': 'version',
        'OS': 'osVersion',
        'Unnamed: 10': 'projectName'
    }

    # Rename the columns using the mapping
    excelDataDf.rename(columns=columnMapping, inplace=True)

    # Convert DataFrame to JSON string
    jsonStr = excelDataDf.to_json(orient='records', force_ascii=False)

    # Parse the JSON string into a Python list of dictionaries
    jsonData = json.loads(jsonStr)

    printerDriver = 'プリンタードライバー'
    faxDriver = 'ファクスドライバー'
    scannerDriver = 'スキャナードライバー'
    # Define the header of the HTML table
    htmlTable = '''
    <!DOCTYPE html>
    <html>
    <head>
    <style>
    table {
    border-collapse: collapse;
    width: 100%;
    }
    th, td {
    border: 1px solid black;
    padding: 8px;
    text-align: center;
    }
    </style>
    </head>
    <body>
    <h2>Fujifilm Driver Report</h2>
    <table>
    <tr>
    <th>Number</th>
    <th>Product Name</th>
    <th>Driver Name</th>
    <th>Version</th>
    <th>OS</th>
    <th>Download Status</th>
    <th>Remark</th>
    </tr>
    '''
    # Save the JSON data into a JSON-LD file
    with open('output' + sheetName + '.jsonld', 'w', encoding='utf-8') as outfile:
        json.dump(jsonData, outfile, ensure_ascii=False, indent=2)

    with open('failed.txt', 'w') as failedToDownload, open('success.txt','w') as successfulDownload:
        failedToDownload.write(f"Failed to download the following drivers for {sheetName}:\n")
        successfulDownload.write(f"Successfuly downloaded the following drivers for {sheetName}:\n")
        for item in jsonData:
            driverName = item.get('driverName')
            version = item.get('version')
            rawProductName = item.get('productName')
            driverNumber+=1
            downloaded = False
            fullDriverName = driverName
            if rawProductName is None:
                rawProductName = ""

            # This is the name that the folder with the productName will have
            if "\n(WHQL)" in rawProductName:
                productName = rawProductName.split("\n")
            else:
                productName = rawProductName
            # Get only the first element of the list as a string
            productName = productName[0] if isinstance(productName, list) else productName

            productName = productName.strip()
            if productName == '(WHQL)':
                downloadStatus = 'N/A'
                remark = 'Skipped because there is no product name'
                htmlTable += f'''
                <tr>
                    <td>{driverNumber}</td>
                    <td>{productName}</td>
                    <td>{driverName}</td>
                    <td>{version}</td>
                    <td>{sheetName}</td>
                    <td>{downloadStatus}</td>
                    <td>{remark}</td>
                </tr>
                '''
                continue

            # Check if driverName exists and is not None
            if driverName is None:
                driverNumber-=1
                continue


            # Split driverName if it contains ' / ' or '\n'. If it doesn't, just take the full name

            driverNamesMatch = re.search(r'^(.*?)(?:\s*\/|\s*\n|$)', driverName)
            if driverNamesMatch:
                driverName = driverNamesMatch.group(1).strip()

            if 'FAX' in driverName:
                driverName = driverName.replace('FAX', '').strip()
                softwareType = faxDriver
            else:
                softwareType = printerDriver

            if driverName.startswith('FX'):
                driverName = driverName.replace('FX', '').strip()

            if 'PCL6' in driverName:
                driverName = driverName.replace('PCL6', '').strip()
            
            if 'PCL 6' in driverName:
                driverName = driverName.replace('PCL 6', '').strip()
                        
            if 'PN' in driverName:
                driverName = driverName.replace('PN', '').strip()

            if 'T2' in driverName:
                driverName = driverName.replace('T2', '').strip()

            if not latestVersion.get(driverName):
                latestVersion[driverName]={}
                latestVersion[driverName][softwareType] = version
            elif latestVersion[driverName].get(softwareType) and latestVersion[driverName][softwareType] >= version:
                downloadStatus = 'ok'
                remark = 'Skipped since there is already a newer version downloaded.'
                # Append the row data to the HTML table
                htmlTable += f'''
                <tr>
                    <td>{driverNumber}</td>
                    <td>{productName}</td>
                    <td>{driverName}</td>
                    <td>{version}</td>
                    <td>{sheetName}</td>
                    <td>{downloadStatus}</td>
                    <td>{remark}</td>
                </tr>
                '''
                continue

            # Initialize the driver variable outside the loop and set it to None
            driver = None

            try:
                if sheetName == '32bit':
                    targetFolder = os.path.join("C:\\Projects\\PSLAD3\\x86", productName, version)
                elif sheetName == '64bit':
                    targetFolder = os.path.join("C:\\Projects\\PSLAD3\\x64", productName, version)

                if os.path.exists(targetFolder):
                    successfulDownload.write(f"{productName} - {fullDriverName} - {version} has already been downloaded.\n")
                    downloadStatus = 'ok'
                    remark = ''
                    htmlTable += f'''
                <tr>
                    <td>{driverNumber}</td>
                    <td>{productName}</td>
                    <td>{driverName}</td>
                    <td>{version}</td>
                    <td>{sheetName}</td>
                    <td>{downloadStatus}</td>
                    <td>{remark}</td>
                </tr>
                '''
                    continue
                # Open 'https://www.fujifilm.com/fb/download?lnk=header'
                chromeOptions = webdriver.ChromeOptions()
                prefs = {'safebrowsing.enabled': 'false'}
                chromeOptions.add_experimental_option("prefs", prefs)
                driver = webdriver.Chrome(options=chromeOptions)
                driver.maximize_window()
                driver.get("https://www.fujifilm.com/fb/download?lnk=header")
                time.sleep(2)
                # Wait until the accept button is clickable and then click it
                wait = WebDriverWait(driver, 10)
                acceptButton = wait.until(EC.element_to_be_clickable((By.ID, "onetrust-accept-btn-handler")))
                acceptButton.click()
                
                try:
                    # Input a driver name
                    searchBox = driver.find_element(By.ID, "txtKeywd")
                    searchBox.send_keys(driverName)
                    searchBox.send_keys(Keys.ENTER)

                    # Wait until the last option is clickable and then click it
                    wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "ui-menu-item")))
                    options = driver.find_elements(By.CLASS_NAME, "ui-menu-item")
                    lastOption = options[-1]
                    lastOption.click()
                except:
                    print(f'{productName} - {driverName} was not found. Please manually download the driver.')
                    failedToDownload.write(f'{productName} - {driverName} was not found. Please manually download the driver.')

                # Wait until the accordion title is clickable and then click it
                wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "m-accordion__title")))
                button = driver.find_element(By.CLASS_NAME, "m-accordion__title")
                driver.execute_script("arguments[0].scrollIntoView();", button)
                button.click()

                # Select the OS
                wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='item1']")))
                button = Select(driver.find_element(By.XPATH, "//*[@id='item1']"))
                button.select_by_visible_text('Windows')

                # Select the OS version
                wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='item2']")))
                button = Select(driver.find_element(By.XPATH, "//*[@id='item2']"))
                if sheetName == '32bit':
                    button.select_by_visible_text('Windows 10 (32ビット) 日本語版')
                elif sheetName == '64bit':
                    button.select_by_visible_text('Windows 11 (64ビット) 日本語版')

                # Select the software type
                wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='item3']")))
                button = Select(driver.find_element(By.XPATH, "//*[@id='item3']"))
                try:
                    button.select_by_visible_text(softwareType)
                except:
                    #If there is not the 'プリンタードライバー' or the 'ファクスドライバー'option, try looking for:
                    try:
                        button.select_by_visible_text('プリンター/ファクスドライバー')
                    #If there is not the 'ファクスドライバー' option, try looking for scanner and download that:
                    except:
                        button.select_by_visible_text(scannerDriver)
                wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="disclosure-contents-0-0"]/div/div/div/div[2]/p[2]/a')))
                button = driver.find_element(By.XPATH, '//*[@id="disclosure-contents-0-0"]/div/div/div/div[2]/p[2]/a')
                driver.execute_script("arguments[0].scrollIntoView();", button)
                button.click()

                # Click on the necessary driver (fax or printer)
                if softwareType == printerDriver:
                    try:
                        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="recommend"]/p/a')))
                        button = driver.find_element(By.XPATH, '//*[@id="recommend"]/p/a')
                    except:
                        try:
                        # If the element is not found, try the alternative XPath
                            wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="content"]/div/div[2]/div[2]/div/p/a')))
                            button = driver.find_element(By.XPATH, '//*[@id="content"]/div/div[2]/div[2]/div/p/a')
                        except:
                            xpath_expression = f"//*[contains(text(), '標準ドライバー')]"
                            wait.until(EC.element_to_be_clickable((By.XPATH, xpath_expression)))
                            button = driver.find_element(By.XPATH, xpath_expression)
            
                elif softwareType == faxDriver:
                    try:
                        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="content"]/div/div[2]/div[2]/div/p/a')))
                        button = driver.find_element(By.XPATH, '//*[@id="content"]/div/div[2]/div[2]/div/p/a')
                    except:
                        xpath_expression = f"//*[contains(text(), '標準ドライバー')]"
                        wait.until(EC.element_to_be_clickable((By.XPATH, xpath_expression)))
                        button = driver.find_element(By.XPATH, xpath_expression)



                driver.execute_script("arguments[0].scrollIntoView();",button)
                button.click()
                # Download the driver
                wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "btn-m")))
                button = driver.find_element(By.CLASS_NAME, "btn-m")
                driver.execute_script("arguments[0].scrollIntoView();",button)
                button.click()

                # Wait for the download to complete
                downloadPath = "C:\\Users\\Alex\\Downloads"
                timeout = 30  # Maximum time to wait for the download in seconds
                downloadedFile = None

                for _ in range(timeout):
                    downloadedFiles = glob.glob(downloadPath + "/*")
                    if downloadedFiles:
                        downloadedFile = max(downloadedFiles, key=os.path.getctime)
                        if downloadedFile.endswith(".crdownload"):
                            time.sleep(4)
                        else:
                            break
                    time.sleep(4)

                if downloadedFile:
                    # Move it accordingly
                    if sheetName == '32bit':
                        targetFolder = os.path.join("C:\\Projects\\PSLAD3\\x86", productName, version)
                    elif sheetName == '64bit':
                        targetFolder = os.path.join("C:\\Projects\\PSLAD3\\x64", productName, version)
                    if not os.path.exists(targetFolder):
                        os.makedirs(targetFolder)
                    downloaded=True
                    newPath = os.path.join(targetFolder, os.path.basename(downloadedFile))
                    os.rename(downloadedFile, newPath)

                if downloaded:
                    successfulDownload.write(f"{productName} - {driverName} - {version} - {sheetName} \n")
                    downloadStatus = 'ok'
                    remark = ''
                else:
                    downloadStatus = 'not ok'
                    # Read the content of the failed.txt file to the 'remark' variable
                    with open('failed.txt', 'r') as failedFile:
                        remark = failedFile.read()
                        # Append the row data to the HTML table
                htmlTable += f'''
                <tr>
                    <td>{driverNumber}</td>
                    <td>{productName}</td>
                    <td>{driverName}</td>
                    <td>{version}</td>
                    <td>{sheetName}</td>
                    <td>{downloadStatus}</td>
                    <td>{remark}</td>
                </tr>
                '''




            except Exception as e:
                print(f"For driver {driverName} - {version} - {sheetName}: \n {str(e)}")       
                failedToDownload.write(f" {productName} - {fullDriverName} - {version} - {sheetName}\n")
                downloadStatus = 'not ok'
                remark='The driver was not found. Please download manually.'
                htmlTable += f'''
                <tr>
                    <td>{driverNumber}</td>
                    <td>{productName}</td>
                    <td>{driverName}</td>
                    <td>{version}</td>
                    <td>{sheetName}</td>
                    <td>{downloadStatus}</td>
                    <td>{remark}</td>
                </tr>
                '''            

            finally:
                # Close the driver properly after each iteration
                if driver is not None:
                    driver.quit()
  
        # Complete the HTML table
        htmlTable += '''
        </table>
        </body>
        </html>
        '''

        # Save the HTML table to a file
        with open('driver_report_'+sheetName+'.html', 'w', encoding='utf-8') as htmlFile:
            htmlFile.write(htmlTable)
       
        successfulDownload.write('\n************************************************************************\n')
        failedToDownload.write('\n************************************************************************\n')

process('32bit')
process('64bit')
