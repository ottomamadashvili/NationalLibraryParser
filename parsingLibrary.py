from requests_html import HTMLSession
import xlwt
import requests
from xlwt import Workbook

wb = Workbook()
sheet1 = wb.add_sheet('Sheet 1')
num_of_pages = 2
counter_url_pages = 1



def func(num_of_pages,counter_url_pages):

    for _ in range(num_of_pages):
        print("started")
        url = f"https://catalog.nplg.gov.ge/search~S2*geo?/i978994130/i978994130/{counter_url_pages}%2C247%2C248%2CE/2browse/indexsort=-"
        s = HTMLSession()
        r = s.get(url)
        r.html.render(sleep=1)





        try:
            upper =  r.html.find('.browseEntryData')
            lower = r.html.find('.browseSubEntry')

            counter_one=0
            sheet1.write(counter_one, 3, 'daxasiateba')

            counter_two=0
            sheet1.write(counter_two, 4, 'weli')

            counter_three = 0
            sheet1.write(counter_three, 0, 'ISBN')

            counter_four = 0
            sheet1.write(counter_four, 1, 'saxeli')

            counter_five = 0
            sheet1.write(counter_five, 2, 'gvari')



            for x in lower:



                description = x.text.split('\n')[0]
                counter_one = counter_one + 1
                sheet1.write(counter_one, 3, description)



                year = x.text.split('\n')[-1]
                counter_two = counter_two + 1
                sheet1.write(counter_one, 4, year)


            for x in upper:
                pars = x.text.split(':')


                nomrebi = pars[0]
                counter_three = counter_three + 1
                sheet1.write(counter_three, 0, nomrebi)



                if len(pars) >1:

                    name = pars[1].split(",")[0]
                    counter_four = counter_four + 1
                    sheet1.write(counter_four, 1, name)


                    surname = pars[1].split(",")[1]
                    counter_five = counter_five + 1
                    sheet1.write(counter_five, 2, surname)

                else:
                    name = " "
                    counter_four = counter_four + 1
                    sheet1.write(counter_four, 1, name)


                    surname = " "
                    counter_five = counter_five + 1
                    sheet1.write(counter_five, 2, surname)


        except:
            pass

        filename = f'{counter_url_pages}.xls'
        wb.save(filename)
        counter_url_pages = counter_url_pages + 50
        print("ENDED", url)
