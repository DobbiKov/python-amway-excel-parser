from modules.AmwayParser import AmwayParser



def main():
    print("Добро пожаловать в Amway парсер.")
    amwayParser = AmwayParser()
    amwayParser.get_file()
    amwayParser.get_sheet()
    amwayParser.get_letter()
    amwayParser.get_word()
    _search = amwayParser.search()

    if _search == True:
        print("Программа успешно произвела парсинг и завершила свою работу.")
    else:
        print("Программа завершила работу с ошибкой!")
    

if __name__ == "__main__":
    main()
