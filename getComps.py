import sys
import requests
import pprint
import xmltodict
import json
import pandas as pd
import openpyxl
import os

zwsidFile = open('ZWSID.text', "r")
zwsid = zwsidFile.read()

def getComps(Address, Zipcode, count):
    compsEndpoint = 'https://www.zillow.com/webservice/GetComps.htm'
    zpid = getPropertyId(Address, Zipcode)

    payload = {'zws-id': zwsid, 'zpid': zpid, 'count': count}

    compResponse = requests.post(compsEndpoint, params=payload)
    dictCompResponse = xmltodict.parse(compResponse.content)

    readableDateFrame = formatCompResponse(dictCompResponse, count)
    writeToExcel(readableDateFrame, Address, Zipcode, count)

def getPropertyId(Address, Zipcode):
    searchResultsEndpoint = 'http://www.zillow.com/webservice/GetSearchResults.htm'
    payload = {'zws-id': zwsid, 'address': Address, 'citystatezip': Zipcode}
    searchResponse = requests.post(searchResultsEndpoint, params=payload)
    dictSearchResponse = xmltodict.parse(searchResponse.content)
    zpid = dictSearchResponse['SearchResults:searchresults']['response']['results']['result']['zpid']
    return zpid

def formatCompResponse(response, count):
    listOfProperties = []
    listOfProperties.append(response['Comps:comps']['response']['properties']['principal'])
    for prop in range(int(count)):
        listOfProperties.append(response['Comps:comps']['response']['properties']['comparables']['comp'][prop])

    dataFrame = pd.io.json.json_normalize(listOfProperties) 
    dataFrame.fillna('None', inplace=True)

    dataFrame['Address'] = dataFrame[['address.street', 'address.zipcode', 'address.city', 'address.state']].agg(','.join, axis=1)
    dataFrame['Current Price'] = dataFrame[['zestimate.amount.#text', 'zestimate.amount.@currency']].agg(' '.join, axis=1)
    dataFrame['Lowest Price'] = dataFrame[['zestimate.valuationRange.low.#text','zestimate.valuationRange.low.@currency']].agg(' '.join, axis=1)
    dataFrame['Highest Price'] = dataFrame[['zestimate.valuationRange.high.#text', 'zestimate.valuationRange.high.@currency']].agg(' '.join, axis=1)

    dataFrame.drop(['zpid',
                    'address.street', 
                    'address.zipcode', 
                    'address.city',
                    'address.state',
                    'address.latitude', 
                    'address.longitude', 
                    'zestimate.amount.#text', 
                    'zestimate.amount.@currency',
                    'zestimate.oneWeekChange.@deprecated',
                    'zestimate.valueChange.@duration',
                    'zestimate.percentile',
                    'zestimate.valuationRange.low.#text',
                    'zestimate.valuationRange.low.@currency',
                    'zestimate.valueChange.#text',
                    'zestimate.valueChange.@currency',
                    'zestimate.valuationRange.high.#text', 
                    'zestimate.valuationRange.high.@currency',
                    'localRealEstate.region.@name', 
                    'localRealEstate.region.@id', 
                    'localRealEstate.region.@type',
                    'localRealEstate.region.zindexValue',
                    'localRealEstate.region.links.overview',
                    'localRealEstate.region.links.forSaleByOwner',
                    'localRealEstate.region.links.forSale'],1, inplace=True)

    dataFrame.rename(columns={'@score': 'Score',
                            'links.homedetails': 'Home Details URL',
                            'links.graphsanddata': 'Graphs URL',
                            'links.mapthishome': 'Maps URL',
                            'links.comparables': 'Comparables URL',
                            'zestimate.last-updated': 'Latest Update',
                            }, inplace=True)
    
    dataFrame = dataFrame[['Address','Score','Current Price','Lowest Price', 'Highest Price','Home Details URL','Graphs URL','Maps URL','Comparables URL','Latest Update']]                        
    
    dataFrame.transpose()
    dataFrame.index.name = 'Property'

    indexList = dataFrame.index.tolist()
    for index in indexList:
        if index == 0:
            indexList[index] = 'Main Property'
        else:
            indexList[index] = 'Comparable Property # %s' % str(index)
    
    dataFrame.index = indexList

    return dataFrame

def writeToExcel(DataFrame, Address, Zipcode, Count):
    sheetName = Address + ' - ' + Zipcode
    filePath = os.path.join(os.path.dirname(__file__),'Comparable Homes.xlsx')
    workbookExists = os.path.exists(filePath)
    writer = pd.ExcelWriter(filePath, engine='openpyxl')

    if workbookExists:
        workbook = openpyxl.load_workbook('Comparable Homes.xlsx')
        writer.book = workbook
        if sheetName in workbook.sheetnames:
            del workbook[sheetName]
    
    DataFrame.to_excel(writer, sheet_name = sheetName)
    worksheet = writer.sheets[sheetName]

    for column in worksheet.columns:
        #column is tuple with all cells for that column
        column_letter = column[0].column_letter
        cell_lengths = []
        for cell in range(int(Count)):
            cell_lengths.append(len(str(column[cell].value)))
        
        worksheet.column_dimensions[column_letter].width = max(cell_lengths)

    writer.save()
    writer.close()

    if os.name == 'posix':
        os.system("open -a 'Microsoft Excel.app' 'Comparable Homes.xlsx'")
    elif os.name == 'nt':
        os. startfile('Comparable Homes.xlsx')

if __name__ == '__main__':
    getComps(*sys.argv[1:])

