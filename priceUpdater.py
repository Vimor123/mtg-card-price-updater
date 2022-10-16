import openpyxl
import requests
from bs4 import BeautifulSoup

# Input variables
excelFileName = "album.xlsx"

# Excel Sheet Structure
startingRow = 2
nameColumn = "A"
setColumn = "B"
versionColumn = "C"
priceColumn = "D"


def getAllCards(excelFileName):
    workbook = openpyxl.load_workbook(excelFileName)
    worksheet = workbook.active

    cards = []

    readingRow = startingRow
    
    while worksheet[nameColumn + str(readingRow)].value != None:
        cardName = worksheet[nameColumn + str(readingRow)].value
        setName = worksheet[setColumn + str(readingRow)].value
        card = {}
        card["cardName"] = cardName
        card["setName"] = setName
        cards.append(card)
        readingRow += 1

    return cards


def getCardPrice(card):

    def generateCardURL(card):
        url = "https://www.cardmarket.com/en/Magic/Products/Singles/"
        setName = card["setName"]
        
        urlSet = ""
        for letter in setName:
            if letter == "'" or letter == ":":
                continue
            elif letter == " ":
                urlSet += '-'
            else:
                urlSet += letter

        url += urlSet + "/"

        cardName = card["cardName"]

        urlName = ""
        for letter in cardName:
            if letter == "'" or letter == ":" or letter == ",":
                continue
            elif letter == " ":
                urlName += '-'
            else:
                urlName += letter
        url += urlName
        
        return url

    
    url = generateCardURL(card)
    page = requests.get(url)
    soup = BeautifulSoup(page.content, 'html.parser')
    
    lists = soup.find_all('div', class_="info-list-container")
    container = lists[0]
    column = container.find_all('dd', class_="col-6 col-xl-7")
    priceTrendDD = column[5]
    priceTrendSpanList = priceTrendDD.find_all('span')
    priceTrendSpan = priceTrendSpanList[0]

    priceTrendString = priceTrendSpan.contents[0]

    priceTrendString = priceTrendString[:-2]
    
    return priceTrendString


def fetchCardPrices(cards):
    for cardIndex in range(len(cards)):
        print("({}/{}): Fetching data for: {}".format(cardIndex+1, len(cards), cards[cardIndex]["cardName"]))
        cardPrice = getCardPrice(cards[cardIndex])
        cards[cardIndex]["cardPrice"] = cardPrice


def updateExcelSpreadsheet(cards, excelFileName):
    workbook = openpyxl.load_workbook(excelFileName)
    worksheet = workbook.active

    writingRow = startingRow
    for card in cards:
        worksheet[priceColumn + str(writingRow)].value = card["cardPrice"]
        writingRow += 1

    workbook.save(excelFileName)


def main():
    cards = getAllCards(excelFileName)
    fetchCardPrices(cards)
    updateExcelSpreadsheet(cards, excelFileName)

    print("All done")


if __name__ == "__main__":
    main()
