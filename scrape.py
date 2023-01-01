from bs4 import BeautifulSoup
import requests
import openpyxl

#creates new excel files
excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = 'Top rated Movies'
print(excel.sheetnames)

# create excel column heading/names
sheet.append(['Title: ','Rank: ','Released Year: ','IMDB Rating: '])



try:
    # using requests module to access IMDB site
    URL = requests.get('https://www.imdb.com/chart/top/')
     #catch error from response object incase URL is invalid 
    URL.raise_for_status()  

    # extracts html contents and parses it using 'html.parser'
    info = BeautifulSoup(URL.text, 'html.parser')

   # finds and returns a list of all 250 movie info
    movies = info.find('tbody', class_ = "lister-list").find_all('tr')
    
    # loop to iterate through each 'tr' tag and access all 'td' tags which contain movie infos

    for movie in movies:
        title = movie.find('td',class_ = "titleColumn").a.text

        # returns just the rank value of the movies ; 'strip' removes new lines and spaces ; '.split' splits after the '.'
        rank = movie.find('td', class_ = "titleColumn" ).get_text(strip=True).split('.')[0]

        releasedYear = movie.find('td',class_ = "titleColumn").span.text.strip('()')

        IMDBrating = movie.find('td',class_ = "ratingColumn").strong.text

        print(title,rank,releasedYear,IMDBrating)

        # inserts all above info into excel sheet:

        sheet.append([title,rank,releasedYear,IMDBrating])
        

except Exception as e:
    print(e)

# save the excel file created

excel.save('IMDB Scraper Infos.xlsx')


