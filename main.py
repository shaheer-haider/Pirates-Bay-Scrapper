from selenium import webdriver
import pandas as pd
from urllib.parse import quote

category = input("Enter category in small letters,\nAvailable Cateogry is [all, video, audio, applications, games, porn, other]\n")
fileName = input("\n\nEnter Name of Input Excel File Correctly in sensitive case... \n SheeT.xlsx and sheet.xlsx is not same, please write name of extension and file both.\n")
sheetName = input("\n\nSheetName name is required. \n If you don't know it, Open sheet Excel sheet to check name of sheet.\n")
nameOfOutputSheet = input("\n\nTHIS FILE WILL CREATE AS NEW FILE AND REPLACE IF THERE'S ANOTHER FILE OF SAME NAME IN FOLDER, IT's not same as Input file\n")
input("\n\nPLEASE CLOSE EXCEL FILES OF RESULTS AND INPUTS BOTH AND DON'T OPEN IT UNTIL PROGRAM IS COMPLETE. \n\nPress any key to continue. \n")
# open browser and a page
driver = webdriver.Chrome("chromedriver.exe")
driver.get("https://thepiratebay.org")

# import movies list
moviesDF = pd.read_excel(fileName, sheet_name=sheetName)
movieList = list(moviesDF["movie_name"])
pd.DataFrame({}).to_excel(nameOfOutputSheet, index=False)
# NOW WORK IS IN parsedMovieNames
movieNumber = 1
while True:
    if(movieNumber > len(movieList)):
        break
    try:
        print(movieNumber)
        if(len(driver.window_handles) > 1):
            driver.close()
            driver.switch_to.window(driver.window_handles[0])

        driver.execute_script("window.open('');")
        driver.switch_to.window(driver.window_handles[-1])
        driver.get(f"https://thepiratebay.org/search.php?q={quote(movieList[movieNumber - 1])}&{category}=on")
        moviesLinkItems = list(driver.find_elements_by_class_name("item-title"))
        if(len(moviesLinkItems) > 4):
            moviesLinkItems = moviesLinkItems[:4]
        # new object for a movie.
        OneMovieObject = {
                "movieName": movieList[movieNumber - 1]
                }
        movieMagnetNames = list(map(lambda x: x.text, moviesLinkItems))
        movieMagnetPageLinks = list(map(lambda x: x.find_element_by_tag_name("a").get_attribute("href"), moviesLinkItems))
        for movieLinkItem in movieMagnetPageLinks:
            # add magnet page in object magnet name
            index = str(movieMagnetPageLinks.index(movieLinkItem) + 1)
            OneMovieObject["magnet_name_" + index] = movieMagnetNames[int(index) - 1]
            # go on new page of this magnet
            driver.get(movieLinkItem)
            # add magnet link, size and seeder in object
            OneMovieObject["magnet_link_" + index ] = driver.find_element_by_id("d").find_elements_by_tag_name("a")[1].get_attribute("href")
            OneMovieObject["magnet_size_" + index ] = driver.find_element_by_id("size").text
            OneMovieObject["magnet_seeders_" + index ] = driver.find_element_by_id("s").text
        # convert dict to dataframe
        currentDF = pd.DataFrame(OneMovieObject, index=[movieNumber])
        # read scrapped excel
        scrappedDF = pd.read_excel(nameOfOutputSheet)
        if(len(scrappedDF) != 0):
            pd.concat([scrappedDF, currentDF]).to_excel(
                nameOfOutputSheet,
                index=False
            )
        else:
            currentDF.to_excel(nameOfOutputSheet, index=False)
        movieNumber += 1
    except Exception as e:
        print(e)
        print("An Error Occured Trying Again, Due to Internet or Server Delays.....\nTRYING AGAIN.")
        continue

