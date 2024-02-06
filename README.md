# Scrape-1-university-website
Scrape all emails, names, positions, and departments for these 4,336 employees from the provided URL. Each link leads to their profile, containing user information. Save the results as an Excel file.

#Code: "#Increase the max of i for Page number, 10 Profiles per a Page"
```python 
if __name__ == "__main__":
    ##Filter the option in Marketplace and copy the url
    scraper = WebScraper()
    url1 = "https://www.rmit.edu.au/search?searchtype=contacts&q=%22RMIT+University%22"
    scraper.navigate_to_url(url1)
    scraper.get_data_from_Firstpage()

    for i in range(2,20): #Increase the max of i for Page number, 10 Profiles per a Page
        url2 = "https://www.rmit.edu.au/search?q=%22RMIT+University%22&searchtype=contacts&current=" + str(i) + "&size=10"
        scraper.navigate_to_url(url2)
        scraper.get_data_from_Firstpage()
    
    # Export the data to Excel after the loop
    scraper.export_to_excel()

    scraper.close_browser()

```

# Export data to excel:
<img src="./Images\Output.PNG">
